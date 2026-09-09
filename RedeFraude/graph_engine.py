import networkx as nx
import json
import sqlite3
import pandas as pd
from pathlib import Path
import streamlit as st

from utils import formatar_cpf_cnpj, formatar_tel
from database import get_db_connection, DB_PATH

# =====================================================
# 1. PROCESSAMENTO DE REDES (NETWORKX)
# =====================================================
@st.cache_data
def carregar_redes():
    if not DB_PATH.exists():
        return None, []

    conn = get_db_connection()
    try:
        df = pd.read_sql_query("SELECT cpf, telefone, placa, titular FROM assistencias", conn)
    except Exception:
        df = pd.DataFrame()
    finally:
        conn.close()

    if df.empty:
        return None, []

    G = nx.Graph()

    for _, row in df.iterrows():
        cpf, tel, placa, nome = row["cpf"], row["telefone"], row["placa"], row["titular"]
        entidades = []
        if cpf:
            nid = f"CPF_{cpf}"
            G.add_node(nid, tipo="cpf", label=f"👤 {nome or formatar_cpf_cnpj(cpf)}", valor=formatar_cpf_cnpj(cpf))
            entidades.append(nid)
        if tel:
            nid = f"TEL_{tel}"
            G.add_node(nid, tipo="telefone", label=f"🚨 {formatar_tel(tel)}", valor=formatar_tel(tel))
            entidades.append(nid)
        if placa:
            nid = f"PLACA_{placa}"
            fmt_p = f"{placa[:3]}-{placa[3:]}" if len(placa) == 7 else placa
            G.add_node(nid, tipo="placa", label=f"🚗 {fmt_p}", valor=fmt_p)
            entidades.append(nid)

        for i in range(len(entidades)):
            for j in range(i + 1, len(entidades)):
                u, v = entidades[i], entidades[j]
                if G.has_edge(u, v):
                    G[u][v]["weight"] += 1
                else:
                    G.add_edge(u, v, weight=1)

    G.remove_nodes_from(list(nx.isolates(G)))
    componentes = [c for c in nx.connected_components(G) if len(c) >= 3]
    componentes.sort(key=len, reverse=True)

    cluster_info = []
    for c_id, comp in enumerate(componentes, 1):
        sub = G.subgraph(comp)
        graus = dict(sub.degree())
        maior_hub = max(graus, key=graus.get)
        cluster_info.append({
            "id": c_id,
            "tamanho": len(comp),
            "hub_label": G.nodes[maior_hub]["label"],
            "hub_id": maior_hub,
            "nodes": list(comp)
        })

    return G, cluster_info

# =====================================================
# 2. CÁLCULO DE MÉTRICAS & SUBGRAFO FORENSE
# =====================================================
def processar_subgrafo_caso(G, cluster_nodes, filtro_tipos, hub_id, font_slider=12):
    """Filtra o subgrafo, calcula métricas de inteligência e formata nós/arestas."""
    subG = G.subgraph(cluster_nodes)
    nos_filtrados = [n for n in subG.nodes if subG.nodes[n]["tipo"] in filtro_tipos]
    subG_filtrado = subG.subgraph(nos_filtrados)

    # 1. Grau de Conexões (Degree)
    node_degrees = dict(subG_filtrado.degree())

    # 2. Intermediação de Pontes (Betweenness)
    if len(subG_filtrado) > 2:
        betweenness_scores = nx.betweenness_centrality(subG_filtrado, weight="weight")
    else:
        betweenness_scores = {n: 0.0 for n in subG_filtrado.nodes}

    # 3. Agrupamento por Comunidades (Greedy Modularity)
    try:
        comm_sets = list(nx.community.greedy_modularity_communities(subG_filtrado))
        comm_map = {}
        for c_idx, c_set in enumerate(comm_sets):
            for n in c_set:
                comm_map[n] = c_idx
    except Exception:
        comm_map = {n: 0 for n in subG_filtrado.nodes}

    cores = {
        "telefone": {"bg": "#B91C1C", "border": "#EF4444"},
        "cpf": {"bg": "#0284C7", "border": "#38BDF8"},
        "placa": {"bg": "#7C3AED", "border": "#A78BFA"}
    }

    vis_nodes = []
    for n in subG_filtrado.nodes:
        nd = subG_filtrado.nodes[n]
        cfg = cores[nd["tipo"]]
        is_hub = (n == hub_id)
        vis_nodes.append({
            "id": n,
            "label": nd["label"],
            "tipo": nd["tipo"],
            "shape": "box",
            "margin": 7,
            "borderWidth": 3 if is_hub else 1,
            "degree": node_degrees.get(n, 1),
            "betweenness": round(betweenness_scores.get(n, 0.0), 4),
            "community": comm_map.get(n, 0),
            "color": {
                "background": cfg["bg"],
                "border": "#FBBF24" if is_hub else cfg["border"],
                "highlight": {"background": cfg["border"], "border": "#FFFFFF"}
            },
            "font": {"color": "#FFFFFF", "bold": True, "size": font_slider, "face": "Segoe UI"}
        })

    vis_edges = []
    for idx_e, (u, v, data) in enumerate(subG_filtrado.edges(data=True)):
        peso = data.get("weight", 1)
        largura = min(8.0, 1.4 + (peso * 0.45))
        vis_edges.append({
            "id": f"e_{idx_e}",
            "from": u,
            "to": v,
            "title": f"Vínculo recorrente: {peso} ocorrência(s)",
            "label": str(peso) if peso >= 3 else "",
            "font": {"color": "#38BDF8", "size": 9, "strokeWidth": 0, "align": "top"},
            "color": {"color": "#64748B", "highlight": "#38BDF8"},
            "width": largura
        })

    return vis_nodes, vis_edges

# =====================================================
# 3. RENDERIZADOR HTML / VIS.JS (OFFLINE & DETERMINÍSTICO)
# =====================================================
def obter_vis_js_local():
    caminho_js = Path(__file__).parent / "vis-network.min.js"
    if caminho_js.exists():
        try:
            with open(caminho_js, "r", encoding="utf-8") as f:
                return f"<script type='text/javascript'>\n{f.read()}\n</script>"
        except Exception:
            pass
    return '<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/vis-network/9.1.9/standalone/umd/vis-network.min.js"></script>'

def gerar_html_grafo(vis_nodes_json, vis_edges_json, base_font_size=12, hub_id="", target_node_id=""):
    hub_id_safe = json.dumps(hub_id)
    target_node_id_safe = json.dumps(target_node_id)
    script_vis = obter_vis_js_local()
    
    return f"""<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta charset="utf-8" />
    {script_vis}
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; font-family: 'Segoe UI', Tahoma, sans-serif; }}
        body {{ background: #06090F; overflow: hidden; height: 100vh; }}
        #network {{ width: 100%; height: 100%; }}
        .floating-controls {{
            position: absolute; top: 10px; left: 10px; z-index: 10;
            display: flex; align-items: center; gap: 6px; background: rgba(11, 17, 30, 0.90);
            border: 1px solid #1E293B; border-radius: 6px; padding: 4px 8px;
            backdrop-filter: blur(8px);
        }}
        .floating-controls select, .floating-controls button {{
            background: #1E293B; border: 1px solid #334155; color: #F8FAFC;
            padding: 4px 8px; border-radius: 4px; font-size: 11px; cursor: pointer;
            font-weight: 600; transition: 0.15s; outline: none;
        }}
        .floating-controls select:focus, .floating-controls button:hover {{
            background: #0284C7; border-color: #38BDF8;
        }}
        .chk-adaptive {{
            display: flex; align-items: center; gap: 4px; font-size: 11px;
            color: #94A3B8; font-weight: 600; cursor: pointer; margin-left: 2px;
        }}
        .chk-adaptive input {{ accent-color: #38BDF8; cursor: pointer; }}
        .legend {{
            position: absolute; top: 10px; right: 10px; z-index: 10;
            display: flex; gap: 8px; background: rgba(11, 17, 30, 0.90);
            border: 1px solid #1E293B; border-radius: 6px; padding: 4px 10px;
            font-size: 11px; color: #E2E8F0; font-weight: 600; pointer-events: none;
            backdrop-filter: blur(8px);
        }}
        .legend span {{ display: flex; align-items: center; gap: 5px; }}
        .badge {{ width: 8px; height: 8px; border-radius: 2px; }}
    </style>
</head>
<body>
    <div class="floating-controls">
        <select id="sel-layout" onchange="changeLayout(this.value)">
            <option value="organico">🌀 Teia Fluida Orgânica</option>
            <option value="subredes">🧩 Sub-redes em Grade (Componentes Conexos)</option>
            <option value="comunidades">🏛️ Comunidades & Facções (Modularity)</option>
            <option value="pontes">🌉 Pontes & Gargalos (Betweenness Centrality)</option>
            <option value="nucleo">🎯 Núcleo vs. Periferia (Degree Centrality)</option>
            <option value="arvore">🌲 Árvore Forense (Hub-Spoke)</option>
        </select>
        <label class="chk-adaptive">
            <input type="checkbox" id="chk-adaptive" checked onchange="toggleAdaptive()"> Texto Auto
        </label>
        <button onclick="fitView()">🎯 Enquadrar</button>
        <button id="btn-freeze" onclick="toggleFreeze()">⏸️ Congelar</button>
        <button onclick="copyToClipboard()">📋 Copiar</button>
        <button onclick="exportPNG()">📷 PNG</button>
    </div>
    <div class="legend">
        <span><div class="badge" style="background:#B91C1C;"></div> Telefone</span>
        <span><div class="badge" style="background:#0284C7;"></div> Titular</span>
        <span><div class="badge" style="background:#7C3AED;"></div> Placa</span>
    </div>
    <div id="network"></div>
    <script>
        const rawNodes = {vis_nodes_json};
        const rawEdges = {vis_edges_json};
        const masterHubId = {hub_id_safe};
        const targetSearchId = {target_node_id_safe};
        
        const nodesDataSet = new vis.DataSet(rawNodes);
        const edgesDataSet = new vis.DataSet(rawEdges);
        const container = document.getElementById('network');
        const data = {{ nodes: nodesDataSet, edges: edgesDataSet }};
        
        let isFrozen = false;
        let baseFontSize = {base_font_size};
        let isAdaptive = true;
        let zoomTimeout = null;
        let currentHighlightId = targetSearchId || null;

        const options = {{
            physics: {{
                enabled: true,
                solver: 'barnesHut',
                barnesHut: {{
                    gravitationalConstant: -4200,
                    centralGravity: 0.18,
                    springLength: 130,
                    springConstant: 0.05,
                    damping: 0.88,
                    avoidOverlap: 0.35
                }},
                stabilization: {{ iterations: 120, updateInterval: 25 }}
            }},
            interaction: {{ hover: true, zoomView: true, dragNodes: true }}
        }};

        const network = new vis.Network(container, data, options);

        function fitView() {{
            network.fit({{ animation: {{ duration: 350 }}, padding: 40 }});
        }}

        function getConnectedComponents(nodes, edges) {{
            const adj = {{}};
            nodes.forEach(n => adj[n.id] = []);
            edges.forEach(e => {{
                if (adj[e.from] && adj[e.to]) {{
                    adj[e.from].push(e.to);
                    adj[e.to].push(e.from);
                }}
            }});
            const visited = new Set();
            const components = [];
            
            if (adj[masterHubId] && !visited.has(masterHubId)) {{
                const comp = [];
                const queue = [masterHubId];
                visited.add(masterHubId);
                while (queue.length > 0) {{
                    const curr = queue.shift();
                    const nObj = nodes.find(x => x.id === curr);
                    if (nObj) comp.push(nObj);
                    adj[curr].forEach(nb => {{
                        if (!visited.has(nb)) {{
                            visited.add(nb);
                            queue.push(nb);
                        }}
                    }});
                }}
                components.push(comp);
            }}

            nodes.forEach(n => {{
                if (!visited.has(n.id)) {{
                    const comp = [];
                    const queue = [n.id];
                    visited.add(n.id);
                    while (queue.length > 0) {{
                        const curr = queue.shift();
                        const nObj = nodes.find(x => x.id === curr);
                        if (nObj) comp.push(nObj);
                        adj[curr].forEach(nb => {{
                            if (!visited.has(nb)) {{
                                visited.add(nb);
                                queue.push(nb);
                            }}
                        }});
                    }}
                    components.push(comp);
                }}
            }});

            const masterComp = components.shift();
            components.sort((a, b) => b.length - a.length);
            if (masterComp) components.unshift(masterComp);

            return components;
        }}

        function getLocalHub(comp, edges) {{
            if (comp.length <= 1) return comp[0];
            const nodeIds = new Set(comp.map(n => n.id));
            const degrees = {{}};
            comp.forEach(n => degrees[n.id] = 0);
            edges.forEach(e => {{
                if (nodeIds.has(e.from) && nodeIds.has(e.to)) {{
                    degrees[e.from] = (degrees[e.from] || 0) + 1;
                    degrees[e.to] = (degrees[e.to] || 0) + 1;
                }}
            }});
            let maxDeg = -1;
            let best = comp[0];
            comp.forEach(n => {{
                if (degrees[n.id] > maxDeg) {{
                    maxDeg = degrees[n.id];
                    best = n;
                }}
            }});
            return best;
        }}

        function applyHighlight(selectedId) {{
            currentHighlightId = selectedId;
            if (selectedId && nodesDataSet.get(selectedId)) {{
                const connected = network.getConnectedNodes(selectedId);
                connected.push(selectedId);
                nodesDataSet.forEach(node => {{
                    const orig = rawNodes.find(n => n.id === node.id);
                    const isSelected = (node.id === selectedId);
                    nodesDataSet.update({{
                        id: node.id,
                        opacity: connected.includes(node.id) ? 1.0 : 0.15,
                        color: isSelected ? {{
                            background: orig.color.background, border: '#FBBF24',
                            highlight: {{ background: orig.color.background, border: '#FBBF24' }}
                        }} : orig.color,
                        borderWidth: isSelected ? 3 : 1
                    }});
                }});
            }} else {{
                nodesDataSet.forEach(node => {{
                    const orig = rawNodes.find(n => n.id === node.id);
                    nodesDataSet.update({{
                        id: node.id, opacity: 1.0, color: orig.color,
                        borderWidth: (node.id === masterHubId ? 3 : 1)
                    }});
                }});
            }}
        }}

        network.once("stabilizationIterationsDone", function() {{
            fitView();
            if (targetSearchId) applyHighlight(targetSearchId);
        }});

        network.on("click", function(params) {{
            if (params.nodes.length > 0) applyHighlight(params.nodes[0]);
            else applyHighlight(null);
        }});

        function updateAdaptiveFont() {{
            if (!isAdaptive) return;
            const scale = network.getScale();
            if (!scale || isNaN(scale) || scale <= 0.05) return;
            const targetSize = Math.max(9, Math.min(30, Math.round(baseFontSize / scale)));
            const updates = [];
            nodesDataSet.forEach(n => {{
                updates.push({{ id: n.id, font: {{ size: targetSize, color: "#FFFFFF", bold: true, face: "Segoe UI" }} }});
            }});
            nodesDataSet.update(updates);
        }}

        function toggleAdaptive() {{
            isAdaptive = document.getElementById('chk-adaptive').checked;
            const updates = [];
            nodesDataSet.forEach(n => {{
                updates.push({{ id: n.id, font: {{ size: baseFontSize, color: "#FFFFFF", bold: true, face: "Segoe UI" }} }});
            }});
            nodesDataSet.update(updates);
            if (isAdaptive) updateAdaptiveFont();
        }}

        network.on("zoom", function() {{
            if (isAdaptive) {{
                if (zoomTimeout) clearTimeout(zoomTimeout);
                zoomTimeout = setTimeout(updateAdaptiveFont, 160);
            }}
        }});

        function changeLayout(mode) {{
            if (mode === "organico") {{
                network.setOptions({{
                    layout: {{ hierarchical: false }},
                    physics: {{
                        enabled: true, solver: 'barnesHut',
                        barnesHut: {{ gravitationalConstant: -4200, centralGravity: 0.18, springLength: 130, springConstant: 0.05, damping: 0.88, avoidOverlap: 0.35 }}
                    }}
                }});
                isFrozen = false;
                document.getElementById('btn-freeze').innerText = "⏸️ Congelar";
                setTimeout(() => {{ fitView(); if (currentHighlightId) applyHighlight(currentHighlightId); }}, 350);
            }} 
            else if (mode === "arvore") {{
                network.setOptions({{
                    layout: {{ hierarchical: {{ direction: "UD", sortMethod: "hubsize", levelSeparation: 140, nodeSpacing: 150 }} }},
                    physics: {{ enabled: false }}
                }});
                isFrozen = true;
                document.getElementById('btn-freeze').innerText = "▶️ Liberar";
                setTimeout(fitView, 60);
            }}
            else {{
                network.setOptions({{ physics: {{ enabled: false }}, layout: {{ hierarchical: false }} }});
                isFrozen = true;
                document.getElementById('btn-freeze').innerText = "▶️ Liberar";

                const updates = [];

                if (mode === "subredes") {{
                    const components = getConnectedComponents(rawNodes, rawEdges);
                    const numComps = components.length;
                    const cols = numComps <= 1 ? 1 : (numComps <= 4 ? 2 : (numComps <= 9 ? 3 : 4));
                    const rows = Math.ceil(numComps / cols);

                    const compRadii = components.map(c => Math.max(90, Math.min(400, c.length * 15)));
                    const maxRadius = Math.max(...compRadii, 120);
                    const cellWidth = Math.max(700, maxRadius * 2 + 160);
                    const cellHeight = Math.max(600, maxRadius * 2 + 140);

                    components.forEach((comp, k) => {{
                        const col = k % cols;
                        const row = Math.floor(k / cols);
                        const cx = Math.round((col - (cols - 1) / 2) * cellWidth);
                        const cy = Math.round((row - (rows - 1) / 2) * cellHeight);
                        const localHub = comp.find(n => n.id === masterHubId) || getLocalHub(comp, rawEdges);

                        const others = comp.filter(n => n.id !== localHub.id);
                        const rComp = Math.max(80, Math.min(380, others.length * 16));

                        updates.push({{
                            id: localHub.id, x: cx, y: cy, physics: false,
                            borderWidth: (localHub.id === masterHubId ? 3 : 2),
                            color: {{ background: localHub.color.background, border: (localHub.id === masterHubId ? '#FBBF24' : '#38BDF8') }}
                        }});

                        others.forEach((n, idx) => {{
                            const a = (2 * Math.PI * idx) / Math.max(others.length, 1);
                            updates.push({{
                                id: n.id,
                                x: Math.round(cx + (rComp * Math.cos(a))),
                                y: Math.round(cy + (rComp * Math.sin(a))),
                                physics: false,
                                color: n.color
                            }});
                        }});
                    }});
                }}
                else if (mode === "comunidades") {{
                    const commGroups = {{}};
                    rawNodes.forEach(n => {{
                        const cid = n.community || 0;
                        if (!commGroups[cid]) commGroups[cid] = [];
                        commGroups[cid].push(n);
                    }});

                    const commKeys = Object.keys(commGroups).sort((a, b) => commGroups[b].length - commGroups[a].length);
                    const totalComms = commKeys.length;
                    const cols = totalComms <= 2 ? totalComms : (totalComms <= 4 ? 2 : 3);
                    const rows = Math.ceil(totalComms / cols);
                    const commSpacingX = 850;
                    const commSpacingY = 700;

                    commKeys.forEach((cid, idx) => {{
                        const cNodes = commGroups[cid];
                        const col = idx % cols;
                        const row = Math.floor(idx / cols);
                        const cx = Math.round((col - (cols - 1) / 2) * commSpacingX);
                        const cy = Math.round((row - (rows - 1) / 2) * commSpacingY);
                        const localHub = cNodes.find(n => n.id === masterHubId) || getLocalHub(cNodes, rawEdges);

                        const others = cNodes.filter(n => n.id !== localHub.id);
                        const rInternal = Math.max(75, Math.min(320, others.length * 15));

                        updates.push({{
                            id: localHub.id, x: cx, y: cy, physics: false,
                            borderWidth: (localHub.id === masterHubId ? 3 : 2),
                            color: {{ background: localHub.color.background, border: '#FBBF24' }}
                        }});

                        others.forEach((n, nIdx) => {{
                            const a = (2 * Math.PI * nIdx) / Math.max(others.length, 1);
                            updates.push({{
                                id: n.id,
                                x: Math.round(cx + (rInternal * Math.cos(a))),
                                y: Math.round(cy + (rInternal * Math.sin(a))),
                                physics: false,
                                color: n.color
                            }});
                        }});
                    }});
                }}
                else if (mode === "pontes") {{
                    const bridges = rawNodes.filter(n => n.betweenness > 0.001).sort((a, b) => b.betweenness - a.betweenness);
                    const nonBridges = rawNodes.filter(n => !(n.betweenness > 0.001));

                    const bridgeSpacing = Math.max(70, Math.min(130, 900 / Math.max(bridges.length, 1)));
                    const startY = Math.round(-((bridges.length - 1) * bridgeSpacing) / 2);

                    bridges.forEach((bNode, bIdx) => {{
                        updates.push({{
                            id: bNode.id,
                            x: 0,
                            y: startY + (bIdx * bridgeSpacing),
                            physics: false,
                            borderWidth: 3,
                            color: {{ background: bNode.color.background, border: '#FBBF24' }}
                        }});
                    }});

                    const leftNodes = [];
                    const rightNodes = [];
                    nonBridges.forEach((nb, i) => {{
                        if (i % 2 === 0) leftNodes.push(nb);
                        else rightNodes.push(nb);
                    }});

                    const flankSpacing = 50;
                    const leftStartY = Math.round(-((leftNodes.length - 1) * flankSpacing) / 2);
                    leftNodes.forEach((n, i) => {{
                        updates.push({{
                            id: n.id,
                            x: -520,
                            y: leftStartY + (i * flankSpacing),
                            physics: false,
                            color: n.color
                        }});
                    }});

                    const rightStartY = Math.round(-((rightNodes.length - 1) * flankSpacing) / 2);
                    rightNodes.forEach((n, i) => {{
                        updates.push({{
                            id: n.id,
                            x: 520,
                            y: rightStartY + (i * flankSpacing),
                            physics: false,
                            color: n.color
                        }});
                    }});
                }}
                else if (mode === "nucleo") {{
                    const sortedByDegree = [...rawNodes].sort((a, b) => b.degree - a.degree);
                    const total = sortedByDegree.length;
                    const coreCutoff = Math.max(1, Math.ceil(total * 0.15));
                    const midCutoff = Math.max(coreCutoff + 1, Math.ceil(total * 0.45));

                    const core = sortedByDegree.slice(0, coreCutoff);
                    const mid = sortedByDegree.slice(coreCutoff, midCutoff);
                    const outer = sortedByDegree.slice(midCutoff);

                    const rCore = Math.max(50, core.length * 20);
                    core.forEach((n, i) => {{
                        const a = (2 * Math.PI * i) / Math.max(core.length, 1);
                        updates.push({{
                            id: n.id,
                            x: Math.round(rCore * Math.cos(a)),
                            y: Math.round(rCore * Math.sin(a)),
                            physics: false,
                            borderWidth: 3,
                            color: {{ background: n.color.background, border: '#FBBF24' }}
                        }});
                    }});

                    const rMid = Math.max(260, rCore + 140);
                    mid.forEach((n, i) => {{
                        const a = (2 * Math.PI * i) / Math.max(mid.length, 1);
                        updates.push({{
                            id: n.id,
                            x: Math.round(rMid * Math.cos(a)),
                            y: Math.round(rMid * Math.sin(a)),
                            physics: false,
                            color: n.color
                        }});
                    }});

                    const rOuter = Math.max(480, rMid + 160);
                    outer.forEach((n, i) => {{
                        const a = (2 * Math.PI * i) / Math.max(outer.length, 1);
                        updates.push({{
                            id: n.id,
                            x: Math.round(rOuter * Math.cos(a)),
                            y: Math.round(rOuter * Math.sin(a)),
                            physics: false,
                            color: n.color
                        }});
                    }});
                }}

                nodesDataSet.update(updates);
                setTimeout(() => {{ fitView(); if (currentHighlightId) applyHighlight(currentHighlightId); }}, 60);
            }}
        }}

        function toggleFreeze() {{
            isFrozen = !isFrozen;
            network.setOptions({{ physics: {{ enabled: !isFrozen }} }});
            document.getElementById('btn-freeze').innerText = isFrozen ? "▶️ Liberar" : "⏸️ Congelar";
        }}

        async function copyToClipboard() {{
            try {{
                const canvas = container.querySelector('canvas');
                canvas.toBlob(async function(blob) {{
                    await navigator.clipboard.write([new ClipboardItem({{ 'image/png': blob }})]);
                    alert("✅ Imagem copiada para a área de transferência!");
                }});
            }} catch(e) {{
                alert("Use o botão 'PNG' para salvar diretamente.");
            }}
        }}

        function exportPNG() {{
            const canvas = container.querySelector('canvas');
            const a = document.createElement('a');
            a.download = 'evidencia_caso_lcfo.png';
            a.href = canvas.toDataURL('image/png');
            a.click();
        }}
    </script>
</body>
</html>
"""