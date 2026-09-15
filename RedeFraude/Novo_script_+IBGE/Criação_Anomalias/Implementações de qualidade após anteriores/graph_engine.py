import networkx as nx
import json
import sqlite3
import pandas as pd
from pathlib import Path
import streamlit as st
import hashlib
import re

from utils import formatar_cpf_cnpj, formatar_tel
from database import (
    get_db_connection, DB_PATH, resolver_identidade_componente,
    carregar_layout_caso, salvar_layout_caso, obter_set_nos_ocultos
)

# =====================================================
# 1. PROCESSAMENTO DE REDES OTIMIZADO COM ITERTUPLES & MEMBROS
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
    
    conn_membros = get_db_connection()
    try:
        for comp in componentes:
            sub = G.subgraph(comp)
            graus = dict(sub.degree())
            maior_hub = max(graus, key=graus.get)
            
            id_caso_estavel = resolver_identidade_componente(comp, conn=conn_membros)

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
    finally:
        conn_membros.close()

    return G, cluster_info

# =====================================================
# 2. SUBGRAFO FORENSE COM ROTA A & FILTRO DE NÓS OCULTOS
# =====================================================
def processar_subgrafo_caso(G, cluster_nodes, filtro_tipos, hub_id, font_slider=12, id_caso=""):
    subG = G.subgraph(cluster_nodes)
    nos_ocultos = obter_set_nos_ocultos(id_caso)
    nos_filtrados = [n for n in subG.nodes if subG.nodes[n]["tipo"] in filtro_tipos and n not in nos_ocultos]
    subG_filtrado = subG.subgraph(nos_filtrados).copy()

    node_degrees = dict(subG_filtrado.degree())
    total_nos = len(subG_filtrado)

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

    posicoes_precalculadas = {}
    if total_nos > 150:
        layout_salvo = carregar_layout_caso(id_caso) if id_caso else {}
        nos_atuais = set(subG_filtrado.nodes)

        if layout_salvo and nos_atuais.issubset(set(layout_salvo.keys())):
            posicoes_precalculadas = {n: layout_salvo[n] for n in nos_atuais}
        else:
            try:
                escala = max(800, total_nos * 18)
                pos = nx.spring_layout(subG_filtrado, k=0.15, iterations=40, seed=42, scale=escala)
                posicoes_precalculadas = {n: (int(pos[n][0]), int(pos[n][1])) for n in subG_filtrado.nodes}
                if id_caso:
                    salvar_layout_caso(id_caso, posicoes_precalculadas)
            except (ImportError, ModuleNotFoundError):
                posicoes_precalculadas = {}
            except Exception:
                posicoes_precalculadas = {}

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
            "valor": nd.get("valor", ""),
            "nome_titular": nd.get("nome_titular", ""),
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
        font_slider=font_slider,
        id_caso=""
    )

    clean_alvo = re.sub(r'[^a-zA-Z0-9]', '', str(alvo_principal)).upper()
    for n in vis_nodes:
        if clean_alvo and clean_alvo in n["id"]:
            n["color"]["border"] = "#FBBF24"
            n["borderWidth"] = 3.5
            n["label"] = "🎯 " + n["label"]

    return vis_nodes, vis_edges, hub_id

# =====================================================
# 4. RENDERIZADOR HTML VIS.JS COM EXTENSÃO DE LOD (FLOWSINT)
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
        html, body {{ background: #06090F; overflow: hidden; width: 100%; height: 100%; position: relative; }}
        
        #network {{
            width: 100%; height: 100%;
            background-color: #06090F;
        }}
        
        .floating-controls {{
            position: absolute; top: 12px; left: 12px; z-index: 25;
            display: flex; align-items: center; gap: 6px; background: rgba(11, 17, 30, 0.88);
            border: 1px solid rgba(51, 65, 85, 0.6); border-radius: 6px; padding: 5px 8px;
            backdrop-filter: blur(12px);
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
        
        /* BUSCA RÁPIDA LATERAL (FLOWSINT) */
        .search-panel {{
            position: absolute; top: 54px; left: 12px; z-index: 25; width: 230px;
            background: rgba(11, 17, 30, 0.90); border: 1px solid rgba(51, 65, 85, 0.7);
            border-radius: 6px; padding: 6px; backdrop-filter: blur(12px);
            box-shadow: 0 4px 16px rgba(0,0,0,0.5);
        }}
        .search-input-wrap {{
            position: relative; display: flex; align-items: center;
        }}
        .search-panel input {{
            width: 100%; background: #1E293B; border: 1px solid #334155; color: #F8FAFC;
            padding: 6px 24px 6px 8px; border-radius: 4px; font-size: 11px; outline: none;
        }}
        .search-panel input:focus {{
            border-color: #38BDF8;
        }}
        .search-clear-btn {{
            position: absolute; right: 6px; background: transparent; border: none;
            color: #94A3B8; cursor: pointer; font-size: 14px; display: none; line-height: 1;
        }}
        .search-results {{
            overflow-y: auto; max-height: 220px; display: flex; flex-direction: column;
            gap: 4px; margin-top: 6px; padding-top: 4px; border-top: 1px solid rgba(51, 65, 85, 0.4);
        }}
        .search-result-item {{
            background: rgba(14, 23, 38, 0.9); border: 1px solid #1E293B; padding: 6px 8px;
            border-radius: 4px; font-size: 11px; color: #E2E8F0; cursor: pointer;
            white-space: nowrap; overflow: hidden; text-overflow: ellipsis; transition: 0.15s;
        }}
        .search-result-item:hover {{
            border-color: #38BDF8; background: #1E293B; color: #38BDF8;
        }}
        
        .legend {{
            position: absolute; top: 12px; right: 12px; z-index: 20;
            display: flex; gap: 10px; background: rgba(11, 17, 30, 0.88);
            border: 1px solid rgba(51, 65, 85, 0.6); border-radius: 6px; padding: 5px 12px;
            font-size: 11px; color: #E2E8F0; font-weight: 600; pointer-events: none;
            backdrop-filter: blur(12px);
        }}
        .legend span {{ display: flex; align-items: center; gap: 6px; }}
        .badge {{ width: 8px; height: 8px; border-radius: 2px; }}

        /* MINIMAPA NO CANTO INFERIOR DIREITO */
        .minimap-container {{
            position: absolute; bottom: 14px; right: 14px; z-index: 20;
            background: rgba(11, 17, 30, 0.85); border: 1px solid rgba(245, 158, 11, 0.5);
            border-radius: 6px; padding: 4px; backdrop-filter: blur(10px);
            box-shadow: 0 4px 16px rgba(0,0,0,0.6); pointer-events: none;
        }}
        #minimap-canvas {{
            display: block; border-radius: 3px; background: rgba(6, 9, 15, 0.8);
        }}

        /* MENU DE CONTEXTO */
        .context-menu {{
            position: absolute; display: none; z-index: 100;
            background: rgba(11, 17, 30, 0.96); border: 1px solid #334155; border-radius: 6px;
            padding: 4px; min-width: 180px; box-shadow: 0 8px 24px rgba(0,0,0,0.7);
            backdrop-filter: blur(16px);
        }}
        .context-menu-header {{
            padding: 6px 10px 4px 10px; font-size: 10px; font-weight: 700; color: #94A3B8;
            text-transform: uppercase; border-bottom: 1px solid rgba(51, 65, 85, 0.5);
            margin-bottom: 4px; text-overflow: ellipsis; overflow: hidden; white-space: nowrap;
        }}
        .context-menu-item {{
            padding: 7px 10px; font-size: 11px; color: #F8FAFC; cursor: pointer;
            border-radius: 4px; transition: 0.12s; font-weight: 600;
        }}
        .context-menu-item:hover {{
            background: #0284C7; color: #FFFFFF;
        }}

        /* INSPECTOR DRAWER REDIMENSIONÁVEL EM 2D */
        .drawer {{
            position: absolute; top: 0; right: 0; width: 360px;
            height: calc(100% - 140px);
            background: rgba(11, 17, 30, 0.88); border-left: 1px solid rgba(51, 65, 85, 0.6);
            border-bottom: 1px solid rgba(51, 65, 85, 0.6); border-bottom-left-radius: 8px;
            box-shadow: -10px 0 30px rgba(0, 0, 0, 0.7); backdrop-filter: blur(16px);
            z-index: 50; display: flex; flex-direction: column;
            transform: translateX(100%);
            transition: transform 0.28s cubic-bezier(0.4, 0, 0.2, 1);
            color: #F8FAFC;
        }}
        .drawer.open {{
            transform: translateX(0);
        }}
        .drawer-resize-handle {{
            position: absolute; left: -4px; top: 0; width: 8px; height: 100%;
            cursor: ew-resize; z-index: 60; background: transparent; transition: background 0.15s;
        }}
        .drawer-resize-handle:hover, .drawer-resize-handle.resizing {{
            background: rgba(56, 189, 248, 0.4);
        }}
        .drawer-resize-handle-v {{
            position: absolute; left: 0; bottom: -4px; width: 100%; height: 8px;
            cursor: ns-resize; z-index: 60; background: transparent; transition: background 0.15s;
        }}
        .drawer-resize-handle-v:hover, .drawer-resize-handle-v.resizing {{
            background: rgba(56, 189, 248, 0.4);
        }}
        .drawer-header {{
            padding: 14px 16px; display: flex; align-items: center; justify-content: space-between;
            border-bottom: 1px solid rgba(51, 65, 85, 0.6); background: rgba(14, 23, 38, 0.7);
        }}
        .drawer-close-btn {{
            background: transparent; border: none; color: #94A3B8; font-size: 20px;
            cursor: pointer; line-height: 1; padding: 2px 6px; border-radius: 4px; transition: 0.15s;
        }}
        .drawer-close-btn:hover {{
            color: #F8FAFC; background: #1E293B;
        }}
        .drawer-badge {{
            font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 4px; text-transform: uppercase;
        }}
        .drawer-body {{
            padding: 14px 16px; overflow-y: auto; flex: 1;
        }}
        .drawer-section {{
            background: rgba(11, 17, 30, 0.7); border: 1px solid rgba(51, 65, 85, 0.5);
            border-radius: 6px; padding: 10px 12px;
        }}
        .drawer-field {{
            display: flex; justify-content: space-between; align-items: center;
            padding: 6px 0; border-bottom: 1px solid rgba(30, 41, 59, 0.5); font-size: 12px;
        }}
        .drawer-field:last-child {{
            border-bottom: none;
        }}
        .field-label {{
            color: #94A3B8;
        }}
        .field-value {{
            color: #F8FAFC; font-weight: 600; text-align: right; max-width: 180px;
            overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
        }}
        .field-value.highlight {{
            color: #38BDF8;
        }}
        .drawer-btn-row {{
            display: flex; gap: 6px; margin-top: 10px;
        }}
        .drawer-action-btn {{
            flex: 1; background: #1E293B; border: 1px solid #334155; color: #F8FAFC;
            padding: 6px 8px; border-radius: 4px; font-size: 11px; cursor: pointer;
            font-weight: 600; transition: 0.15s;
        }}
        .drawer-action-btn:hover {{
            background: #0284C7; border-color: #38BDF8;
        }}
        .neighbors-container {{
            margin-top: 14px;
        }}
        .neighbors-header {{
            font-size: 11px; font-weight: 700; color: #94A3B8; text-transform: uppercase;
            margin-bottom: 8px; display: flex; justify-content: space-between;
        }}
        .neighbors-preview-wrap {{
            display: flex; justify-content: center; margin: 8px 0 6px 0;
            background: rgba(6, 9, 15, 0.6); border: 1px solid rgba(51, 65, 85, 0.4);
            border-radius: 6px; padding: 6px;
        }}
        #drawer-neighbors-preview {{
            display: block; border-radius: 4px;
        }}
        .neighbors-list {{
            display: flex; flex-direction: column; gap: 6px; max-height: 220px; overflow-y: auto;
        }}
        .neighbor-item {{
            display: flex; align-items: center; justify-content: space-between;
            background: rgba(14, 23, 38, 0.8); border: 1px solid #1E293B; padding: 8px 10px;
            border-radius: 5px; cursor: pointer; transition: 0.15s; font-size: 11px;
        }}
        .neighbor-item:hover {{
            border-color: #38BDF8; background: #1E293B;
        }}
        .neighbor-item-label {{
            color: #F8FAFC; font-weight: 600; max-width: 180px;
            overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
        }}
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
            <input type="checkbox" id="chk-adaptive" checked onchange="toggleAdaptive()"> LOD Adaptativo
        </label>
        <button onclick="reorganizarEspacar()" title="Recalcular com alta dispersão">📐 Reorganizar</button>
        <button onclick="fitView()">🎯 Enquadrar</button>
        <button id="btn-freeze" onclick="toggleFreeze()">⏸️ Congelar</button>
        <button onclick="copyToClipboard()">📋 Copiar</button>
        <button onclick="exportPNG()">📷 PNG</button>
    </div>

    <!-- BUSCA RÁPIDA LATERAL (FLOWSINT) -->
    <div id="search-panel" class="search-panel">
        <div class="search-input-wrap">
            <input type="text" id="search-input" placeholder="🔍 Filtrar no grafo..." oninput="filtrarListaBusca(this.value)">
            <button id="search-clear-btn" class="search-clear-btn" onclick="limparBuscaGrafo()">&times;</button>
        </div>
        <div id="search-results" class="search-results" style="display:none;"></div>
    </div>

    <div class="legend">
        <span><div class="badge" style="background:#EF4444;"></div> Telefone</span>
        <span><div class="badge" style="background:#38BDF8;"></div> Titular</span>
        <span><div class="badge" style="background:#A78BFA;"></div> Placa</span>
        <span><div class="badge" style="background:#FBBF24;"></div> Hub / Âncora</span>
    </div>

    <!-- MINIMAPA NO CANTO INFERIOR DIREITO -->
    <div id="minimap-container" class="minimap-container">
        <canvas id="minimap-canvas" width="160" height="110"></canvas>
    </div>

    <!-- MENU DE CONTEXTO -->
    <div id="context-menu" class="context-menu">
        <div class="context-menu-header" id="ctx-node-title">Entidade</div>
        <div class="context-menu-item" onclick="ctxCopiarDado()">📋 Copiar Dado</div>
        <div class="context-menu-item" onclick="ctxCentralizar()">🎯 Centralizar no Grafo</div>
        <div class="context-menu-item" onclick="ctxIsolar()">👁️ Isolar Vizinhança</div>
        <div class="context-menu-item" onclick="ctxAbrirDrawer()">📂 Inspecionar Detalhes</div>
    </div>

    <div id="network"></div>

    <!-- INSPECTOR DRAWER REDIMENSIONÁVEL EM 2D -->
    <div id="inspector-drawer" class="drawer">
        <div id="drawer-resize-handle" class="drawer-resize-handle" title="Arraste para ajustar largura"></div>
        <div id="drawer-resize-handle-v" class="drawer-resize-handle-v" title="Arraste para ajustar altura"></div>
        <div class="drawer-header">
            <div style="display:flex; align-items:center; gap:8px;">
                <span id="drawer-badge" class="drawer-badge"></span>
                <span id="drawer-title" style="font-weight:700; font-size:13px; color:#F8FAFC; text-overflow:ellipsis; overflow:hidden; white-space:nowrap; max-width:200px;"></span>
            </div>
            <button class="drawer-close-btn" onclick="closeDrawer()">&times;</button>
        </div>
        <div class="drawer-body">
            <div class="drawer-section">
                <div class="drawer-field">
                    <span class="field-label">Identificador / Dado:</span>
                    <span id="drawer-val" class="field-value highlight"></span>
                </div>
                <div id="drawer-row-nome" class="drawer-field" style="display:none;">
                    <span class="field-label">Nome do Titular:</span>
                    <span id="drawer-nome" class="field-value"></span>
                </div>
                <div class="drawer-field">
                    <span class="field-label">Conexões Diretas:</span>
                    <span id="drawer-degree" class="field-value"></span>
                </div>
                <div class="drawer-field">
                    <span class="field-label">Intermediação (Betweenness):</span>
                    <span id="drawer-betweenness" class="field-value"></span>
                </div>
                <div class="drawer-field">
                    <span class="field-label">Comunidade Modular:</span>
                    <span id="drawer-community" class="field-value"></span>
                </div>
                <div id="drawer-row-hub" class="drawer-field" style="display:none;">
                    <span class="field-label">Status Especial:</span>
                    <span style="color:#FBBF24; font-weight:700;">⭐ Âncora Central (Hub)</span>
                </div>
            </div>

            <div class="drawer-btn-row">
                <button class="drawer-action-btn" onclick="focusCurrentNode()">🎯 Centralizar Nó</button>
                <button class="drawer-action-btn" onclick="isolateCurrentNode()">👁️ Isolar Vizinhança</button>
            </div>

            <div class="neighbors-container">
                <div class="neighbors-header">
                    <span>Entidades Vinculadas (Neighbors)</span>
                    <span id="drawer-neighbors-count" style="color:#38BDF8;">0</span>
                </div>
                <div class="neighbors-preview-wrap">
                    <canvas id="drawer-neighbors-preview" width="260" height="130"></canvas>
                </div>
                <div id="drawer-neighbors-list" class="neighbors-list"></div>
            </div>
        </div>
    </div>

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
        const deveDesativarFisica = (totalNos > 150) || (initialLayout !== "organico");
        const gravidadeAdaptativa = totalNos <= 6 ? -3500 : -22000;
        const centralGravityAdaptativa = totalNos <= 6 ? 0.25 : 0.05;
        const springDistAdaptativa = totalNos <= 6 ? 160 : springDist;
        
        let isFrozen = deveDesativarFisica;
        let baseFontSize = {base_font_size};
        let isAdaptive = true;
        let zoomTimeout = null;
        let currentHighlightId = targetSearchId || null;
        let currentInspectedNodeId = null;
        let contextMenuNodeId = null;
        let isLodSimplified = false;

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
                tooltipDelay: 100,
                hideEdgesOnDrag: (totalNos > 250) // Otimização para redes densas
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

        /* =====================================================
           GRID PONTILHADO DINÂMICO NO CANVAS
        ===================================================== */
        function obterLimitesVisiveis() {{
            const canvasEl = container.querySelector('canvas');
            if (!canvasEl) return null;
            const topLeft = network.DOMtoCanvas({{ x: 0, y: 0 }});
            const bottomRight = network.DOMtoCanvas({{ x: canvasEl.clientWidth, y: canvasEl.clientHeight }});
            return {{ minX: topLeft.x, minY: topLeft.y, maxX: bottomRight.x, maxY: bottomRight.y }};
        }}

        network.on("beforeDrawing", function(ctx) {{
            const limites = obterLimitesVisiveis();
            if (!limites) return;
            const espacamentoGrid = 42;
            const qtdColunas = (limites.maxX - limites.minX) / espacamentoGrid;
            const qtdLinhas = (limites.maxY - limites.minY) / espacamentoGrid;
            if (qtdColunas * qtdLinhas > 6000) return;

            const startX = Math.floor(limites.minX / espacamentoGrid) * espacamentoGrid;
            const startY = Math.floor(limites.minY / espacamentoGrid) * espacamentoGrid;
            ctx.save();
            ctx.fillStyle = "rgba(148, 163, 184, 0.16)";
            for (let x = startX; x <= limites.maxX; x += espacamentoGrid) {{
                for (let y = startY; y <= limites.maxY; y += espacamentoGrid) {{
                    ctx.beginPath();
                    ctx.arc(x, y, 1.3, 0, 2 * Math.PI);
                    ctx.fill();
                }}
            }}
            ctx.restore();
        }});

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

        /* =====================================================
           MINIMAPA 2D EM TEMPO REAL COM SUPORTE A REDES GRANDES
        ===================================================== */
        function desenharMinimapa() {{
            const canvas = document.getElementById('minimap-canvas');
            if (!canvas) return;
            const ctx = canvas.getContext('2d');
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            const positions = network.getPositions();
            const ids = Object.keys(positions);
            if (ids.length === 0) return;

            let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
            ids.forEach(id => {{
                const p = positions[id];
                if (p) {{
                    minX = Math.min(minX, p.x); maxX = Math.max(maxX, p.x);
                    minY = Math.min(minY, p.y); maxY = Math.max(maxY, p.y);
                }}
            }});

            const rangeX = Math.max(maxX - minX, 1);
            const rangeY = Math.max(maxY - minY, 1);
            const pad = 8;

            function toMini(x, y) {{
                return [
                    pad + ((x - minX) / rangeX) * (canvas.width - 2 * pad),
                    pad + ((y - minY) / rangeY) * (canvas.height - 2 * pad)
                ];
            }}

            ids.forEach(id => {{
                const p = positions[id];
                if (!p) return;
                const [mx, my] = toMini(p.x, p.y);
                const nodeObj = rawNodes.find(n => n.id === id);
                ctx.fillStyle = (id === masterHubId) ? "#FBBF24" : (nodeObj && nodeObj.color ? nodeObj.color.background : "#38BDF8");
                ctx.beginPath();
                ctx.arc(mx, my, (id === masterHubId ? 2.5 : 1.5), 0, 2 * Math.PI);
                ctx.fill();
            }});

            try {{
                const viewPos = network.getViewPosition();
                const scale = network.getScale();
                const canvasEl = container.querySelector('canvas');
                if (canvasEl && scale > 0) {{
                    const viewW = canvasEl.clientWidth / scale;
                    const viewH = canvasEl.clientHeight / scale;
                    const [vx1, vy1] = toMini(viewPos.x - viewW / 2, viewPos.y - viewH / 2);
                    const [vx2, vy2] = toMini(viewPos.x + viewW / 2, viewPos.y + viewH / 2);
                    ctx.strokeStyle = "#F59E0B";
                    ctx.lineWidth = 1.4;
                    ctx.strokeRect(vx1, vy1, Math.max(vx2 - vx1, 4), Math.max(vy2 - vy1, 4));
                }}
            }} catch (e) {{}}
        }}

        network.on("afterDrawing", function() {{
            desenharMinimapa();
        }});

        /* =====================================================
           BUSCA LATERAL FLUTUANTE EM MEMÓRIA
        ===================================================== */
        function filtrarListaBusca(termo) {{
            const resultsEl = document.getElementById('search-results');
            const clearBtn = document.getElementById('search-clear-btn');
            if (!termo || termo.trim().length < 2) {{
                resultsEl.style.display = "none";
                resultsEl.innerHTML = "";
                clearBtn.style.display = "none";
                return;
            }}
            clearBtn.style.display = "block";
            const termoClean = termo.replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
            
            const encontrados = rawNodes.filter(n => {{
                const vClean = (n.valor || "").replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
                const lClean = (n.label || "").replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
                const nClean = (n.nome_titular || "").replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
                return vClean.includes(termoClean) || lClean.includes(termoClean) || nClean.includes(termoClean);
            }}).slice(0, 25);

            resultsEl.innerHTML = "";
            if (encontrados.length === 0) {{
                resultsEl.innerHTML = "<div style='color:#94A3B8; font-size:11px; padding:4px;'>Nenhum nó localizado.</div>";
            }} else {{
                encontrados.forEach(n => {{
                    const item = document.createElement('div');
                    item.className = 'search-result-item';
                    let icone = "🔗";
                    if (n.tipo === "cpf") icone = "👤";
                    else if (n.tipo === "telefone") icone = "🚨";
                    else if (n.tipo === "placa") icone = "🚗";
                    item.innerHTML = `<span style='margin-right:4px;'>${{icone}}</span><span>${{n.label}}</span>`;
                    item.onclick = function() {{
                        selectInspectedNode(n.id);
                    }};
                    resultsEl.appendChild(item);
                }});
            }}
            resultsEl.style.display = "flex";
        }}

        function limparBuscaGrafo() {{
            const inp = document.getElementById('search-input');
            inp.value = "";
            filtrarListaBusca("");
        }}

        /* =====================================================
           MENU DE CONTEXTO
        ===================================================== */
        network.on("oncontext", function(params) {{
            params.event.preventDefault();
            const nodeId = network.getNodeAt(params.pointer.DOM);
            const menu = document.getElementById('context-menu');
            if (!nodeId) {{
                menu.style.display = "none";
                return;
            }}
            contextMenuNodeId = nodeId;
            const nodeObj = rawNodes.find(n => n.id === nodeId);
            document.getElementById('ctx-node-title').innerText = nodeObj ? nodeObj.label : "Entidade";
            
            let posX = params.pointer.DOM.x + 8;
            let posY = params.pointer.DOM.y + 8;
            if (posX + 190 > container.clientWidth) posX = params.pointer.DOM.x - 190;
            if (posY + 160 > container.clientHeight) posY = params.pointer.DOM.y - 150;
            
            menu.style.left = posX + "px";
            menu.style.top = posY + "px";
            menu.style.display = "block";
        }});

        function hideContextMenu() {{
            const menu = document.getElementById('context-menu');
            if (menu) menu.style.display = "none";
        }}

        document.addEventListener("click", hideContextMenu);
        network.on("dragStart", hideContextMenu);
        network.on("zoom", hideContextMenu);

        function ctxCopiarDado() {{
            if (!contextMenuNodeId) return;
            const nodeObj = rawNodes.find(n => n.id === contextMenuNodeId);
            if (nodeObj) {{
                navigator.clipboard.writeText(nodeObj.valor || nodeObj.id);
            }}
            hideContextMenu();
        }}

        function ctxCentralizar() {{
            if (contextMenuNodeId) focusNode(contextMenuNodeId);
            hideContextMenu();
        }}

        function ctxIsolar() {{
            if (contextMenuNodeId) applyHighlight(contextMenuNodeId);
            hideContextMenu();
        }}

        function ctxAbrirDrawer() {{
            if (contextMenuNodeId) {{
                openInspector(contextMenuNodeId);
                applyHighlight(contextMenuNodeId);
            }}
            hideContextMenu();
        }}

        /* =====================================================
           DIAGRAMA ORBITAL/RADIAL DE NEIGHBORS
        ===================================================== */
        function desenharNeighborsPreview(nodeId, connectedIds) {{
            const canvas = document.getElementById('drawer-neighbors-preview');
            if (!canvas) return;
            const ctx = canvas.getContext('2d');
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            if (!connectedIds || connectedIds.length === 0) return;

            const cx = canvas.width / 2, cy = canvas.height / 2;
            const raio = Math.min(50, 24 + connectedIds.length * 1.1);
            const posicoes = [];

            ctx.strokeStyle = "rgba(148, 163, 184, 0.25)";
            ctx.lineWidth = 1;
            connectedIds.forEach((id, i) => {{
                const ang = (2 * Math.PI * i) / connectedIds.length;
                const px = cx + raio * Math.cos(ang);
                const py = cy + raio * Math.sin(ang);
                posicoes.push([px, py]);
                ctx.beginPath();
                ctx.moveTo(cx, cy);
                ctx.lineTo(px, py);
                ctx.stroke();
            }});

            connectedIds.forEach((id, i) => {{
                const nObj = rawNodes.find(n => n.id === id);
                const [px, py] = posicoes[i];
                ctx.beginPath();
                ctx.arc(px, py, 4.5, 0, 2 * Math.PI);
                ctx.fillStyle = nObj && nObj.color ? nObj.color.background : "#38BDF8";
                ctx.fill();
            }});

            const centerObj = rawNodes.find(n => n.id === nodeId);
            ctx.beginPath();
            ctx.arc(cx, cy, 7.5, 0, 2 * Math.PI);
            ctx.fillStyle = centerObj && centerObj.color ? centerObj.color.background : "#38BDF8";
            ctx.fill();
            ctx.strokeStyle = "#FBBF24";
            ctx.lineWidth = 1.8;
            ctx.stroke();
        }}

        /* =====================================================
           INSPECTOR DRAWER TRANSLÚCIDO
        ===================================================== */
        function openInspector(nodeId) {{
            const nodeObj = rawNodes.find(n => n.id === nodeId);
            if (!nodeObj) return;

            currentInspectedNodeId = nodeId;
            const drawer = document.getElementById('inspector-drawer');
            const badge = document.getElementById('drawer-badge');
            const rowNome = document.getElementById('drawer-row-nome');
            const rowHub = document.getElementById('drawer-row-hub');

            if (nodeObj.tipo === "cpf") {{
                badge.style.background = "#0369A1";
                badge.style.color = "#E0F2FE";
                badge.innerText = "TITULAR";
                rowNome.style.display = "flex";
                document.getElementById('drawer-nome').innerText = nodeObj.nome_titular || "N/D";
            }} else if (nodeObj.tipo === "telefone") {{
                badge.style.background = "#991B1B";
                badge.style.color = "#FEE2E2";
                badge.innerText = "TELEFONE";
                rowNome.style.display = "none";
            }} else if (nodeObj.tipo === "placa") {{
                badge.style.background = "#6D28D9";
                badge.style.color = "#EDE9FE";
                badge.innerText = "PLACA";
                rowNome.style.display = "none";
            }} else {{
                badge.style.background = "#334155";
                badge.style.color = "#F8FAFC";
                badge.innerText = "ENTIDADE";
                rowNome.style.display = "none";
            }}

            document.getElementById('drawer-title').innerText = nodeObj.label;
            document.getElementById('drawer-val').innerText = nodeObj.valor || nodeObj.id;
            document.getElementById('drawer-degree').innerText = `${{nodeObj.degree}} conexão(ões)`;
            document.getElementById('drawer-betweenness').innerText = nodeObj.betweenness !== undefined ? `${{nodeObj.betweenness}}` : '0.0';
            document.getElementById('drawer-community').innerText = nodeObj.community !== undefined ? `Grupo #${{nodeObj.community}}` : 'Principal';
            
            rowHub.style.display = (nodeId === masterHubId) ? "flex" : "none";

            const connectedNodeIds = network.getConnectedNodes(nodeId);
            document.getElementById('drawer-neighbors-count').innerText = connectedNodeIds.length;
            
            desenharNeighborsPreview(nodeId, connectedNodeIds);

            const listEl = document.getElementById('drawer-neighbors-list');
            listEl.innerHTML = "";

            connectedNodeIds.forEach(nbId => {{
                const nbObj = rawNodes.find(n => n.id === nbId);
                if (!nbObj) return;

                const edgeObj = rawEdges.find(e => (e.from === nodeId && e.to === nbId) || (e.from === nbId && e.to === nodeId));
                const edgeTitle = edgeObj && edgeObj.title ? edgeObj.title.replace('🔗 Vínculo relacional: ', '') : 'Conexão';

                const item = document.createElement('div');
                item.className = 'neighbor-item';
                item.onclick = function() {{
                    selectInspectedNode(nbId);
                }};

                let icone = "🔗";
                if (nbObj.tipo === "cpf") icone = "👤";
                else if (nbObj.tipo === "telefone") icone = "🚨";
                else if (nbObj.tipo === "placa") icone = "🚗";

                item.innerHTML = `
                    <div style="display:flex; align-items:center; gap:6px; overflow:hidden;">
                        <span>${{icone}}</span>
                        <span class="neighbor-item-label">${{nbObj.label}}</span>
                    </div>
                    <span style="background:rgba(56, 189, 248, 0.15); color:#38BDF8; padding:2px 6px; border-radius:3px; font-size:9px; font-weight:700; white-space:nowrap;">
                        ${{edgeTitle}}
                    </span>
                `;
                listEl.appendChild(item);
            }});

            drawer.classList.add('open');
        }}

        function closeDrawer() {{
            const drawer = document.getElementById('inspector-drawer');
            drawer.classList.remove('open');
            currentInspectedNodeId = null;
        }}

        function selectInspectedNode(nodeId) {{
            openInspector(nodeId);
            applyHighlight(nodeId);
            focusNode(nodeId);
        }}

        function focusNode(nodeId) {{
            network.focus(nodeId, {{
                scale: 1.15,
                animation: {{ duration: 350, easingFunction: 'easeInOutQuad' }}
            }});
        }}

        function focusCurrentNode() {{
            if (currentInspectedNodeId) {{
                focusNode(currentInspectedNodeId);
            }}
        }}

        function isolateCurrentNode() {{
            if (currentInspectedNodeId) {{
                applyHighlight(currentInspectedNodeId);
            }}
        }}

        /* REDIMENSIONAMENTO INTERATIVO DO DRAWER (LARGURA E ALTURA) */
        (function initDrawerResize() {{
            const handleH = document.getElementById('drawer-resize-handle');
            const handleV = document.getElementById('drawer-resize-handle-v');
            const drawer = document.getElementById('inspector-drawer');
            let resizingH = false;
            let resizingV = false;

            handleH.addEventListener('mousedown', function(e) {{
                resizingH = true;
                handleH.classList.add('resizing');
                document.body.style.userSelect = 'none';
                e.preventDefault();
            }});

            handleV.addEventListener('mousedown', function(e) {{
                resizingV = true;
                handleV.classList.add('resizing');
                document.body.style.userSelect = 'none';
                e.preventDefault();
            }});

            document.addEventListener('mousemove', function(e) {{
                if (resizingH) {{
                    const novaLargura = Math.min(640, Math.max(300, window.innerWidth - e.clientX));
                    drawer.style.width = novaLargura + "px";
                }}
                if (resizingV) {{
                    const novaAltura = Math.min(window.innerHeight, Math.max(220, e.clientY));
                    drawer.style.height = novaAltura + "px";
                }}
            }});

            document.addEventListener('mouseup', function() {{
                if (resizingH) {{
                    resizingH = false;
                    handleH.classList.remove('resizing');
                }}
                if (resizingV) {{
                    resizingV = false;
                    handleV.classList.remove('resizing');
                }}
                document.body.style.userSelect = '';
            }});
        }})();

        document.getElementById('sel-layout').value = initialLayout;

        network.once("stabilizationIterationsDone", function() {{
            fitView();
            if (targetSearchId) {{
                applyHighlight(targetSearchId);
                openInspector(targetSearchId);
            }}
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
            if (params.nodes.length > 0) {{
                const clickedId = params.nodes[0];
                applyHighlight(clickedId);
                openInspector(clickedId);
            }} else {{
                applyHighlight(null);
                closeDrawer();
            }}
        }});

        /* =====================================================
           LOD (LEVEL OF DETAIL) ADAPTATIVO POR NÍVEL DE ZOOM
        ===================================================== */
        function updateAdaptiveLOD() {{
            if (!isAdaptive) return;
            const scale = network.getScale();
            if (!scale || scale <= 0.05) return;

            // Transição LOD: Se o zoom estiver muito distante (< 0.22), vira micro-pontos para performance máxima
            if (scale < 0.22 && !isLodSimplified) {{
                isLodSimplified = true;
                const updates = [];
                nodesDataSet.forEach(n => {{
                    updates.push({{
                        id: n.id,
                        shape: "dot",
                        size: (n.id === masterHubId ? 5.5 : 3.5),
                        label: "" // Corta desenho de texto da CPU
                    }});
                }});
                nodesDataSet.update(updates);
                return;
            }} 
            else if (scale >= 0.22 && isLodSimplified) {{
                isLodSimplified = false;
                const updates = [];
                nodesDataSet.forEach(n => {{
                    const orig = rawNodes.find(x => x.id === n.id);
                    updates.push({{
                        id: n.id,
                        shape: "box",
                        label: orig ? orig.label : ""
                    }});
                }});
                nodesDataSet.update(updates);
            }}

            // Cálculo dinâmico do tamanho da fonte para manter legibilidade quando em visão de cartões
            if (!isLodSimplified) {{
                const targetSize = Math.max(9, Math.min(20, Math.round(baseFontSize / Math.sqrt(scale))));
                const updates = [];
                nodesDataSet.forEach(n => {{
                    updates.push({{ id: n.id, font: {{ size: targetSize, color: "#FFFFFF", face: "Segoe UI" }} }});
                }});
                nodesDataSet.update(updates);
            }}
        }}

        function toggleAdaptive() {{
            isAdaptive = document.getElementById('chk-adaptive').checked;
            if (!isAdaptive && isLodSimplified) {{
                isLodSimplified = false;
                const updates = [];
                nodesDataSet.forEach(n => {{
                    const orig = rawNodes.find(x => x.id === n.id);
                    updates.push({{
                        id: n.id,
                        shape: "box",
                        label: orig ? orig.label : "",
                        font: {{ size: baseFontSize, color: "#FFFFFF", face: "Segoe UI" }}
                    }});
                }});
                nodesDataSet.update(updates);
            }} else if (isAdaptive) {{
                updateAdaptiveLOD();
            }}
        }}

        network.on("zoom", function() {{
            if (isAdaptive) {{
                if (zoomTimeout) clearTimeout(zoomTimeout);
                zoomTimeout = setTimeout(updateAdaptiveLOD, 120); // Debounce leve para estabilidade da CPU
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