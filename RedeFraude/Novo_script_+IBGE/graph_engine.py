import networkx as nx
import json
import sqlite3
import pandas as pd
from pathlib import Path
import streamlit as st
import hashlib
import re

from utils import formatar_cpf_cnpj, formatar_tel
from database import get_db_connection, DB_PATH, resolver_identidade_componente

# =====================================================
# 1. PROCESSAMENTO DE REDES OTIMIZADO COM ITERTUPLES & MEMBROS
# =====================================================
@st.cache_data
def carregar_redes():
    if not DB_PATH.exists():
        return None, []

    conn = get_db_connection()
    try:
        # Puxa apenas as colunas necessárias
        df = pd.read_sql_query("SELECT cpf, telefone, placa, titular FROM assistencias", conn)
    except Exception:
        df = pd.DataFrame()
    finally:
        conn.close()

    if df.empty:
        return None, []

    G = nx.Graph()

    # Substituição de iterrows por itertuples (Ganho de 10x a 50x em velocidade)
    for row in df.itertuples(index=False):
        cpf = str(row.cpf).strip() if row.cpf else ""
        tel = str(row.telefone).strip() if row.telefone else ""
        placa = str(row.placa).strip() if row.placa else ""
        nome = str(row.titular).strip() if row.titular else ""
        
        entidades = []
        if cpf:
            nid = f"CPF_{cpf}"
            if nid not in G:
                G.add_node(nid, tipo="cpf", label=f"👤 {nome or formatar_cpf_cnpj(cpf)}", valor=formatar_cpf_cnpj(cpf), nome_titular=nome or "N/D")
            entidades.append(nid)
        if tel:
            nid = f"TEL_{tel}"
            if nid not in G:
                G.add_node(nid, tipo="telefone", label=f"🚨 {formatar_tel(tel)}", valor=formatar_tel(tel), nome_titular="")
            entidades.append(nid)
        if placa:
            nid = f"PLACA_{placa}"
            if nid not in G:
                fmt_p = f"{placa[:3]}-{placa[3:]}" if len(placa) == 7 else placa
                G.add_node(nid, tipo="placa", label=f"🚗 {fmt_p}", valor=fmt_p, nome_titular="")
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
    for comp in componentes:
        sub = G.subgraph(comp)
        graus = dict(sub.degree())
        maior_hub = max(graus, key=graus.get)
        
        # Resolve o ID permanente via tabela caso_membros (nunca se perde no crescimento orgânico!)
        id_caso_estavel = resolver_identidade_componente(comp)

        tels_count = len([n for n in comp if n.startswith("TEL_")])
        cpfs_count = len([n for n in comp if n.startswith("CPF_")])
        placas_count = len([n for n in comp if n.startswith("PLACA_")])

        cluster_info.append({
            "id": id_caso_estavel,
            "tamanho": len(comp),
            "hub_label": G.nodes[maior_hub]["label"],
            "hub_id": maior_hub,
            "qtd_tels": tels_count,
            "qtd_cpfs": cpfs_count,
            "qtd_placas": placas_count,
            "nodes": list(comp)
        })

    return G, cluster_info

# =====================================================
# 2. SUBGRAFO FORENSE COM ROTA A (LAYOUT PRÉ-CALCULADO EM PYTHON)
# =====================================================
def processar_subgrafo_caso(G, cluster_nodes, filtro_tipos, hub_id, font_slider=12):
    subG = G.subgraph(cluster_nodes)
    nos_filtrados = [n for n in subG.nodes if subG.nodes[n]["tipo"] in filtro_tipos]
    subG_filtrado = subG.subgraph(nos_filtrados).copy()

    node_degrees = dict(subG_filtrado.degree())
    total_nos = len(subG_filtrado)

    # Inversão correta de peso para centralidade
    for u, v, data in subG_filtrado.edges(data=True):
        peso = data.get("weight", 1)
        subG_filtrado[u][v]["distance"] = 1.0 / float(peso)

    if total_nos > 2:
        betweenness_scores = nx.betweenness_centrality(subG_filtrado, weight="distance")
    else:
        betweenness_scores = {n: 0.0 for n in subG_filtrado.nodes}

    try:
        comm_sets = list(nx.community.greedy_modularity_communities(subG_filtrado))
        comm_map = {}
        for c_idx, c_set in enumerate(comm_sets):
            for n in c_set:
                comm_map[n] = c_idx
    except Exception:
        comm_map = {n: 0 for n in subG_filtrado.nodes}

    # ROTA A: Para redes densas (>150 nós), o Python pré-calcula as coordenadas (X, Y)
    # usando o spring_layout acelerado por NumPy/C, livrando o browser do cálculo de física!
    posicoes_precalculadas = {}
    if total_nos > 150:
        escala = max(800, total_nos * 18)
        pos = nx.spring_layout(subG_filtrado, k=0.15, iterations=40, seed=42, scale=escala)
        posicoes_precalculadas = {n: (int(pos[n][0]), int(pos[n][1])) for n in subG_filtrado.nodes}

    cores = {
        "telefone": {"bg": "#991B1B", "border": "#EF4444"},
        "cpf": {"bg": "#0369A1", "border": "#38BDF8"},
        "placa": {"bg": "#6D28D9", "border": "#A78BFA"}
    }

    vis_nodes = []
    for n in subG_filtrado.nodes:
        nd = subG_filtrado.nodes[n]
        cfg = cores[nd["tipo"]]
        is_hub = (n == hub_id)
        grau = node_degrees.get(n, 1)

        rotulo = nd["label"]
        if len(rotulo) > 28 and not is_hub:
            rotulo = rotulo[:26] + "…"

        if nd["tipo"] == "cpf":
            tooltip_html = f"👤 TITULAR\nNome: {nd.get('nome_titular', 'N/D')}\nCPF: {nd.get('valor', '')}\nVínculos: {grau} conexões"
        elif nd["tipo"] == "telefone":
            tooltip_html = f"🚨 TELEFONE\nNúmero: {nd.get('valor', '')}\nVínculos: {grau} conexões\n{'⭐ ÂNCORA CENTRAL' if is_hub else ''}"
        elif nd["tipo"] == "placa":
            tooltip_html = f"🚗 PLACA\nVeículo: {nd.get('valor', '')}\nVínculos: {grau} conexões"
        else:
            tooltip_html = f"{nd['label']}\nVínculos: {grau}"

        node_dict = {
            "id": n,
            "label": rotulo,
            "title": tooltip_html,
            "tipo": nd["tipo"],
            "shape": "box",
            "margin": 8,
            "borderWidth": 3 if is_hub else 1.2,
            "degree": grau,
            "betweenness": round(betweenness_scores.get(n, 0.0), 4),
            "community": comm_map.get(n, 0),
            "color": {
                "background": cfg["bg"],
                "border": "#FBBF24" if is_hub else cfg["border"],
                "highlight": {"background": cfg["border"], "border": "#FFFFFF"}
            },
            "font": {
                "color": "#FFFFFF", 
                "bold": is_hub, 
                "size": font_slider if not is_hub else font_slider + 2, 
                "face": "Segoe UI"
            }
        }

        # Se for rede densa, anexa as coordenadas fixas calculadas em Python
        if n in posicoes_precalculadas:
            node_dict["x"] = posicoes_precalculadas[n][0]
            node_dict["y"] = posicoes_precalculadas[n][1]
            node_dict["physics"] = False

        vis_nodes.append(node_dict)

    vis_edges = []
    for idx_e, (u, v, data) in enumerate(subG_filtrado.edges(data=True)):
        peso = data.get("weight", 1)
        largura = min(4.5, 1.0 + (peso * 0.3))

        vis_edges.append({
            "id": f"e_{idx_e}",
            "from": u,
            "to": v,
            "title": f"🔗 Vínculo relacional: {peso} ocorrência(s)",
            "label": "",
            "color": {
                "color": "rgba(148, 163, 184, 0.35)",
                "highlight": "#38BDF8",
                "hover": "#38BDF8"
            },
            "width": largura,
            "smooth": False
        })

    return vis_nodes, vis_edges

# =====================================================
# 3. CONSTRUTOR DINÂMICO DE GRAFO (DESCOBERTA ATIVA)
# =====================================================
def processar_grafo_dataframe(df_assistencias, alvo_principal="", font_slider=12):
    if df_assistencias.empty:
        return [], [], ""

    G_disc = nx.Graph()
    for row in df_assistencias.itertuples(index=False):
        cpf = str(getattr(row, "cpf", "")).strip() if pd.notna(getattr(row, "cpf", None)) else ""
        tel = str(getattr(row, "telefone", "")).strip() if pd.notna(getattr(row, "telefone", None)) else ""
        placa = str(getattr(row, "placa", "")).strip() if pd.notna(getattr(row, "placa", None)) else ""
        nome = str(getattr(row, "titular", "")).strip() if pd.notna(getattr(row, "titular", None)) else ""
        
        entidades = []
        if cpf:
            nid = f"CPF_{cpf}"
            G_disc.add_node(nid, tipo="cpf", label=f"👤 {nome or formatar_cpf_cnpj(cpf)}", valor=formatar_cpf_cnpj(cpf), nome_titular=nome or "N/D")
            entidades.append(nid)
        if tel:
            nid = f"TEL_{tel}"
            G_disc.add_node(nid, tipo="telefone", label=f"🚨 {formatar_tel(tel)}", valor=formatar_tel(tel), nome_titular="")
            entidades.append(nid)
        if placa:
            nid = f"PLACA_{placa}"
            fmt_p = f"{placa[:3]}-{placa[3:]}" if len(placa) == 7 else placa
            G_disc.add_node(nid, tipo="placa", label=f"🚗 {fmt_p}", valor=fmt_p, nome_titular="")
            entidades.append(nid)

        for i in range(len(entidades)):
            for j in range(i + 1, len(entidades)):
                u, v = entidades[i], entidades[j]
                if G_disc.has_edge(u, v):
                    G_disc[u][v]["weight"] += 1
                else:
                    G_disc.add_edge(u, v, weight=1)

    graus = dict(G_disc.degree())
    hub_id = max(graus, key=graus.get) if graus else (list(G_disc.nodes)[0] if G_disc.nodes else "")

    vis_nodes, vis_edges = processar_subgrafo_caso(
        G=G_disc,
        cluster_nodes=list(G_disc.nodes),
        filtro_tipos=["telefone", "cpf", "placa"],
        hub_id=hub_id,
        font_slider=font_slider
    )

    clean_alvo = re.sub(r'[^a-zA-Z0-9]', '', str(alvo_principal)).upper()
    for n in vis_nodes:
        if clean_alvo and clean_alvo in n["id"]:
            n["color"]["border"] = "#FBBF24"
            n["borderWidth"] = 3.5
            n["label"] = "🎯 " + n["label"]

    return vis_nodes, vis_edges, hub_id

# =====================================================
# 4. RENDERIZADOR HTML FORENSE VIS.JS (ADAPTATIVO)
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

def gerar_html_grafo(vis_nodes_json, vis_edges_json, base_font_size=12, hub_id="", target_node_id="", layout_ativo="organico", espacamento=260):
    hub_id_safe = json.dumps(hub_id)
    target_node_id_safe = json.dumps(target_node_id)
    layout_ativo_safe = json.dumps(layout_ativo)
    script_vis = obter_vis_js_local()
    
    return f"""<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta charset="utf-8" />
    {script_vis}
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; font-family: 'Segoe UI', Tahoma, sans-serif; }}
        html, body {{ background: #06090F; overflow: hidden; width: 100%; height: 100%; }}
        #network {{ width: 100%; height: 100%; }}
        
        .floating-controls {{
            position: absolute; top: 12px; left: 12px; z-index: 10;
            display: flex; align-items: center; gap: 6px; background: rgba(11, 17, 30, 0.92);
            border: 1px solid #1E293B; border-radius: 6px; padding: 5px 8px;
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
            position: absolute; top: 12px; right: 12px; z-index: 10;
            display: flex; gap: 10px; background: rgba(11, 17, 30, 0.90);
            border: 1px solid #1E293B; border-radius: 6px; padding: 5px 12px;
            font-size: 11px; color: #E2E8F0; font-weight: 600; pointer-events: none;
            backdrop-filter: blur(8px);
        }}
        .legend span {{ display: flex; align-items: center; gap: 6px; }}
        .badge {{ width: 8px; height: 8px; border-radius: 2px; }}
    </style>
</head>
<body>
    <div class="floating-controls">
        <select id="sel-layout" onchange="changeLayout(this.value)">
            <option value="organico">🌀 Teia Fluida Orgânica</option>
            <option value="radial">🎯 Radial Peacock (i2)</option>
            <option value="arvore">🌲 Árvore Forense</option>
            <option value="pontes">🌉 Pontes & Gargalos</option>
            <option value="subredes">🧩 Sub-redes em Grade</option>
            <option value="comunidades">🏛️ Comunidades & Facções</option>
        </select>
        <label class="chk-adaptive">
            <input type="checkbox" id="chk-adaptive" checked onchange="toggleAdaptive()"> Texto Auto
        </label>
        <button onclick="reorganizarEspacar()" title="Recalcular com alta dispersão">📐 Reorganizar</button>
        <button onclick="fitView()">🎯 Enquadrar</button>
        <button id="btn-freeze" onclick="toggleFreeze()">⏸️ Congelar</button>
        <button onclick="copyToClipboard()">📋 Copiar</button>
        <button onclick="exportPNG()">📷 PNG</button>
    </div>
    <div class="legend">
        <span><div class="badge" style="background:#EF4444;"></div> Telefone</span>
        <span><div class="badge" style="background:#38BDF8;"></div> Titular</span>
        <span><div class="badge" style="background:#A78BFA;"></div> Placa</span>
        <span><div class="badge" style="background:#FBBF24;"></div> Hub / Âncora</span>
    </div>
    <div id="network"></div>
    <script>
        const rawNodes = {vis_nodes_json};
        const rawEdges = {vis_edges_json};
        const masterHubId = {hub_id_safe};
        const targetSearchId = {target_node_id_safe};
        const initialLayout = {layout_ativo_safe};
        const springDist = {espacamento};
        
        const nodesDataSet = new vis.DataSet(rawNodes);
        const edgesDataSet = new vis.DataSet(rawEdges);
        const container = document.getElementById('network');
        const data = {{ nodes: nodesDataSet, edges: edgesDataSet }};
        
        const totalNos = rawNodes.length;
        // Para redes pequenas (1 a 6 nós), física suave; para redes densas (>150), física desligada (Rota A)
        const deveDesativarFisica = (totalNos > 150) || (initialLayout !== "organico");
        const gravidadeAdaptativa = totalNos <= 6 ? -3500 : -22000;
        const centralGravityAdaptativa = totalNos <= 6 ? 0.25 : 0.05;
        const springDistAdaptativa = totalNos <= 6 ? 160 : springDist;
        
        let isFrozen = deveDesativarFisica;
        let baseFontSize = {base_font_size};
        let isAdaptive = true;
        let zoomTimeout = null;
        let currentHighlightId = targetSearchId || null;

        const options = {{
            physics: {{
                enabled: !deveDesativarFisica,
                solver: 'barnesHut',
                barnesHut: {{
                    gravitationalConstant: gravidadeAdaptativa,
                    centralGravity: centralGravityAdaptativa,
                    springLength: springDistAdaptativa,
                    springConstant: 0.03,
                    damping: 0.92,
                    avoidOverlap: 0.95
                }},
                stabilization: {{ 
                    enabled: true,
                    iterations: (totalNos > 300 ? 50 : 180),
                    updateInterval: 25,
                    fit: true 
                }}
            }},
            interaction: {{ 
                hover: true, 
                zoomView: true, 
                dragNodes: true,
                tooltipDelay: 100
            }}
        }};

        const network = new vis.Network(container, data, options);

        function fitView() {{
            network.fit({{ animation: {{ duration: 350 }}, padding: 60 }});
        }}

        function reorganizarEspacar() {{
            network.setOptions({{
                physics: {{
                    enabled: true,
                    solver: 'barnesHut',
                    barnesHut: {{
                        gravitationalConstant: gravidadeAdaptativa - 4000,
                        centralGravity: centralGravityAdaptativa,
                        springLength: springDistAdaptativa + 40,
                        springConstant: 0.025,
                        damping: 0.94,
                        avoidOverlap: 1.0
                    }}
                }}
            }});
            isFrozen = false;
            document.getElementById('btn-freeze').innerText = "⏸️ Congelar";
            setTimeout(fitView, 400);
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
                        opacity: connected.includes(node.id) ? 1.0 : 0.12,
                        color: isSelected ? {{
                            background: orig.color.background, border: '#FBBF24',
                            highlight: {{ background: orig.color.background, border: '#FBBF24' }}
                        }} : orig.color,
                        borderWidth: isSelected ? 3.5 : 1.2
                    }});
                }});
            }} else {{
                nodesDataSet.forEach(node => {{
                    const orig = rawNodes.find(n => n.id === node.id);
                    nodesDataSet.update({{
                        id: node.id, opacity: 1.0, color: orig.color,
                        borderWidth: (node.id === masterHubId ? 3 : 1.2)
                    }});
                }});
            }}
        }}

        document.getElementById('sel-layout').value = initialLayout;

        network.once("stabilizationIterationsDone", function() {{
            fitView();
            if (targetSearchId) applyHighlight(targetSearchId);
            if (deveDesativarFisica) {{
                network.setOptions({{ physics: {{ enabled: false }} }});
                document.getElementById('btn-freeze').innerText = "▶️ Liberar";
            }}
            if (initialLayout && initialLayout !== "organico") {{
                changeLayout(initialLayout);
            }}
        }});

        if (initialLayout && initialLayout !== "organico") {{
            setTimeout(() => changeLayout(initialLayout), 50);
        }}

        network.on("click", function(params) {{
            if (params.nodes.length > 0) applyHighlight(params.nodes[0]);
            else applyHighlight(null);
        }});

        function updateAdaptiveFont() {{
            if (!isAdaptive) return;
            const scale = network.getScale();
            if (!scale || scale <= 0.05) return;
            const targetSize = Math.max(9, Math.min(20, Math.round(baseFontSize / Math.sqrt(scale))));
            const updates = [];
            nodesDataSet.forEach(n => {{
                updates.push({{ id: n.id, font: {{ size: targetSize, color: "#FFFFFF", face: "Segoe UI" }} }});
            }});
            nodesDataSet.update(updates);
        }}

        function toggleAdaptive() {{
            isAdaptive = document.getElementById('chk-adaptive').checked;
            const updates = [];
            nodesDataSet.forEach(n => {{
                updates.push({{ id: n.id, font: {{ size: baseFontSize, color: "#FFFFFF", face: "Segoe UI" }} }});
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
            document.getElementById('sel-layout').value = mode;
            if (mode === "organico") {{
                network.setOptions({{
                    layout: {{ hierarchical: false }},
                    physics: {{
                        enabled: true, solver: 'barnesHut',
                        barnesHut: {{ 
                            gravitationalConstant: gravidadeAdaptativa, 
                            centralGravity: centralGravityAdaptativa, 
                            springLength: springDistAdaptativa, 
                            springConstant: 0.03, 
                            damping: 0.92, 
                            avoidOverlap: 0.95 
                        }}
                    }}
                }});
                isFrozen = false;
                document.getElementById('btn-freeze').innerText = "⏸️ Congelar";
                setTimeout(() => {{ fitView(); if (currentHighlightId) applyHighlight(currentHighlightId); }}, 350);
            }} 
            else if (mode === "arvore") {{
                network.setOptions({{
                    layout: {{ hierarchical: {{ direction: "UD", sortMethod: "hubsize", levelSeparation: 170, nodeSpacing: 210 }} }},
                    physics: {{ enabled: false }}
                }});
                isFrozen = true;
                document.getElementById('btn-freeze').innerText = "▶️ Liberar";
                setTimeout(fitView, 60);
            }}
            else if (mode === "radial") {{
                network.setOptions({{ physics: {{ enabled: false }}, layout: {{ hierarchical: false }} }});
                isFrozen = true;
                document.getElementById('btn-freeze').innerText = "▶️ Liberar";

                const updates = [];
                const hubNode = rawNodes.find(n => n.id === masterHubId) || rawNodes[0];
                
                updates.push({{
                    id: hubNode.id, x: 0, y: 0, physics: false,
                    borderWidth: 3, color: {{ background: hubNode.color.background, border: '#FBBF24' }}
                }});

                const adj = new Set();
                rawEdges.forEach(e => {{
                    if (e.from === hubNode.id) adj.add(e.to);
                    if (e.to === hubNode.id) adj.add(e.from);
                }});

                const ring1 = rawNodes.filter(n => n.id !== hubNode.id && adj.has(n.id));
                const ring2 = rawNodes.filter(n => n.id !== hubNode.id && !adj.has(n.id));

                const r1 = Math.max(200, ring1.length * 28);
                ring1.forEach((n, i) => {{
                    const angle = (2 * Math.PI * i) / Math.max(ring1.length, 1);
                    updates.push({{
                        id: n.id,
                        x: Math.round(r1 * Math.cos(angle)),
                        y: Math.round(r1 * Math.sin(angle)),
                        physics: false,
                        color: n.color
                    }});
                }});

                const r2 = r1 + Math.max(220, ring2.length * 18);
                ring2.forEach((n, i) => {{
                    const angle = (2 * Math.PI * i) / Math.max(ring2.length, 1);
                    updates.push({{
                        id: n.id,
                        x: Math.round(r2 * Math.cos(angle)),
                        y: Math.round(r2 * Math.sin(angle)),
                        physics: false,
                        color: n.color
                    }});
                }});

                nodesDataSet.update(updates);
                setTimeout(() => {{ fitView(); if (currentHighlightId) applyHighlight(currentHighlightId); }}, 60);
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

                    const compRadii = components.map(c => Math.max(120, Math.min(500, c.length * 24)));
                    const maxRadius = Math.max(...compRadii, 160);
                    const cellWidth = Math.max(900, maxRadius * 2 + 200);
                    const cellHeight = Math.max(800, maxRadius * 2 + 180);

                    components.forEach((comp, k) => {{
                        const col = k % cols;
                        const row = Math.floor(k / cols);
                        const cx = Math.round((col - (cols - 1) / 2) * cellWidth);
                        const cy = Math.round((row - (rows - 1) / 2) * cellHeight);
                        const localHub = comp.find(n => n.id === masterHubId) || getLocalHub(comp, rawEdges);

                        const others = comp.filter(n => n.id !== localHub.id);
                        const rComp = Math.max(100, Math.min(460, others.length * 24));

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
                    const commSpacingX = 1000;
                    const commSpacingY = 850;

                    commKeys.forEach((cid, idx) => {{
                        const cNodes = commGroups[cid];
                        const col = idx % cols;
                        const row = Math.floor(idx / cols);
                        const cx = Math.round((col - (cols - 1) / 2) * commSpacingX);
                        const cy = Math.round((row - (rows - 1) / 2) * commSpacingY);
                        const localHub = cNodes.find(n => n.id === masterHubId) || getLocalHub(cNodes, rawEdges);

                        const others = cNodes.filter(n => n.id !== localHub.id);
                        const rInternal = Math.max(90, Math.min(400, others.length * 20));

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

                    const bridgeSpacing = Math.max(80, Math.min(150, 1000 / Math.max(bridges.length, 1)));
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

                    const flankSpacing = 60;
                    const leftStartY = Math.round(-((leftNodes.length - 1) * flankSpacing) / 2);
                    leftNodes.forEach((n, i) => {{
                        updates.push({{
                            id: n.id,
                            x: -620,
                            y: leftStartY + (i * flankSpacing),
                            physics: false,
                            color: n.color
                        }});
                    }});

                    const rightStartY = Math.round(-((rightNodes.length - 1) * flankSpacing) / 2);
                    rightNodes.forEach((n, i) => {{
                        updates.push({{
                            id: n.id,
                            x: 620,
                            y: rightStartY + (i * flankSpacing),
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
                    alert("✅ Grafo copiado para a área de transferência!");
                }});
            }} catch(e) {{
                alert("Use o botão 'PNG' para exportar.");
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