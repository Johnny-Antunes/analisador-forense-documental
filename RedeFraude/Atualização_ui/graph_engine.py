"""
Módulo: graph_engine.py
Objetivo: Construtor e renderizador gráfico Vis.js com estética inspirada no Flowsint.
          Nós circulares com ícones vetoriais SVG e texto limpo inferior (sem prefixos colchetes).
          Dispersão espacial calibrada e consistente entre os 4 layouts técnicos, física
          sempre desligada (posições sempre pré-calculadas em Python, nunca simuladas ao vivo).
"""

from typing import Tuple, List, Dict, Any, Optional
import networkx as nx
import json
import sqlite3
import pandas as pd
from pathlib import Path
import streamlit as st
import hashlib
import re
import urllib.parse
import math

from utils import formatar_cpf_cnpj, formatar_tel
from database import (
    get_db_connection, DB_PATH, resolver_identidade_componente,
    carregar_layout_caso, salvar_layout_caso, obter_set_nos_ocultos
)


def _criar_svg_nodo_flowsint(tipo: str, cor_bg: str, cor_borda: str, is_hub: bool = False) -> str:
    """Gera um ícone vetorial SVG circular com pictograma branco no centro, estilo Flowsint."""
    stroke_w = "3.5" if is_hub else "1.8"
    borda_final = "#FBBF24" if is_hub else cor_borda

    if tipo == "cpf":
        path_icon = '<path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" fill="none" stroke="#FFFFFF" stroke-width="2.2"/><circle cx="12" cy="7" r="4" fill="none" stroke="#FFFFFF" stroke-width="2.2"/>'
    elif tipo == "telefone":
        path_icon = '<path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z" fill="none" stroke="#FFFFFF" stroke-width="2.2"/>'
    elif tipo == "placa":
        path_icon = '<path d="M19 17h2c.6 0 1-.4 1-1v-3c0-.9-.7-1.7-1.5-1.9C18.7 10.6 16 10 16 10s-1.3-1.4-2.2-2.3c-.5-.4-1.1-.7-1.8-.7H5c-.6 0-1.1.4-1.4.9l-1.5 2.8C2 10.9 2 11.2 2 11.5V16c0 .6.4 1 1 1h2" fill="none" stroke="#FFFFFF" stroke-width="2"/><circle cx="7" cy="17" r="2" fill="none" stroke="#FFFFFF" stroke-width="2"/><circle cx="17" cy="17" r="2" fill="none" stroke="#FFFFFF" stroke-width="2"/>'
    else:
        path_icon = '<rect x="4" y="2" width="16" height="20" rx="2" fill="none" stroke="#FFFFFF" stroke-width="2"/><line x1="9" y1="6" x2="9.01" y2="6" stroke="#FFFFFF" stroke-width="2.5"/><line x1="15" y1="6" x2="15.01" y2="6" stroke="#FFFFFF" stroke-width="2.5"/><line x1="9" y1="10" x2="9.01" y2="10" stroke="#FFFFFF" stroke-width="2.5"/><line x1="15" y1="10" x2="15.01" y2="10" stroke="#FFFFFF" stroke-width="2.5"/><line x1="9" y1="14" x2="9.01" y2="14" stroke="#FFFFFF" stroke-width="2.5"/><line x1="15" y1="14" x2="15.01" y2="14" stroke="#FFFFFF" stroke-width="2.5"/>'

    svg_raw = (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 48 48">'
        f'<circle cx="24" cy="24" r="21" fill="{cor_bg}" stroke="{borda_final}" stroke-width="{stroke_w}"/>'
        f'<g transform="translate(12, 12)">{path_icon}</g>'
        f'</svg>'
    )
    return f"data:image/svg+xml;utf8,{urllib.parse.quote(svg_raw)}"


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
                G.add_node(nid, tipo="cpf", label=nome or formatar_cpf_cnpj(cpf), valor=formatar_cpf_cnpj(cpf), nome_titular=nome or "N/D")
            entidades.append(nid)
        if tel:
            nid = f"TEL_{tel}"
            if nid not in G:
                G.add_node(nid, tipo="telefone", label=formatar_tel(tel), valor=formatar_tel(tel), nome_titular="")
            entidades.append(nid)
        if placa:
            nid = f"PLACA_{placa}"
            if nid not in G:
                fmt_p = f"{placa[:3]}-{placa[3:]}" if len(placa) == 7 else placa
                G.add_node(nid, tipo="placa", label=fmt_p, valor=fmt_p, nome_titular="")
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

            cluster_info.append({
                "id": id_caso_estavel,
                "tamanho": len(comp),
                "hub_label": G.nodes[maior_hub]["label"],
                "hub_id": maior_hub,
                "qtd_tels": len([n for n in comp if n.startswith("TEL_")]),
                "qtd_cpfs": len([n for n in comp if n.startswith("CPF_")]),
                "qtd_placas": len([n for n in comp if n.startswith("PLACA_")]),
                "nodes": list(comp)
            })
    finally:
        conn_membros.close()

    return G, cluster_info


def processar_subgrafo_caso(G, cluster_nodes, filtro_tipos, hub_id, font_slider=11, id_caso="", espacamento=280):
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

    # Posições SEMPRE pré-calculadas e persistidas -- física ao vivo nunca roda,
    # em nenhum tamanho de célula, em nenhum layout.
    posicoes_precalculadas = {}
    layout_salvo = carregar_layout_caso(id_caso) if id_caso else {}
    nos_atuais = set(subG_filtrado.nodes)

    if layout_salvo and nos_atuais.issubset(set(layout_salvo.keys())):
        posicoes_precalculadas = {n: layout_salvo[n] for n in nos_atuais}
    else:
        try:
            # Teto em 30.000px para não deixar o texto ilegível em casos de 1000+ nós.
            escala = min(30000, max(2400, total_nos * 100))
            k_calc = max(0.28, 3.2 / math.sqrt(max(total_nos, 1)))
            pos = nx.spring_layout(subG_filtrado, k=k_calc, iterations=70, seed=42, scale=escala)
            posicoes_precalculadas = {n: (int(pos[n][0]), int(pos[n][1])) for n in subG_filtrado.nodes}
            if id_caso:
                salvar_layout_caso(id_caso, posicoes_precalculadas)
        except Exception:
            posicoes_precalculadas = {}

    cores = {
        "telefone": {"bg": "#991B1B", "border": "#EF4444"},
        "cpf": {"bg": "#0369A1", "border": "#38BDF8"},
        "placa": {"bg": "#6D28D9", "border": "#A78BFA"},
        "empresa": {"bg": "#1E293B", "border": "#64748B"}
    }

    vis_nodes = []
    for n in subG_filtrado.nodes:
        nd = subG_filtrado.nodes[n]
        cfg = cores.get(nd["tipo"], {"bg": "#334155", "border": "#94A3B8"})
        is_hub = (n == hub_id)
        grau = node_degrees.get(n, 1)

        rotulo_limpo = nd["label"]
        if len(rotulo_limpo) > 24 and not is_hub:
            rotulo_limpo = rotulo_limpo[:22] + "…"

        svg_image = _criar_svg_nodo_flowsint(nd["tipo"], cfg["bg"], cfg["border"], is_hub=is_hub)

        node_dict = {
            "id": n,
            "label": rotulo_limpo,
            "title": f"{nd['tipo'].upper()}: {nd.get('valor', '')}\nConexões: {grau}",
            "tipo": nd["tipo"],
            "shape": "image",
            "image": svg_image,
            "size": 26 if is_hub else 20,
            "degree": grau,
            "betweenness": round(betweenness_scores.get(n, 0.0), 4),
            "community": comm_map.get(n, 0),
            "valor": nd.get("valor", ""),
            "nome_titular": nd.get("nome_titular", ""),
            "color": {
                "background": cfg["bg"],
                "border": "#FBBF24" if is_hub else cfg["border"]
            },
            "font": {
                "color": "#F8FAFC",
                "size": font_slider,
                "face": "Segoe UI",
                "strokeWidth": 2.5,
                "strokeColor": "#06090F"
            }
        }

        if n in posicoes_precalculadas:
            node_dict["x"] = posicoes_precalculadas[n][0]
            node_dict["y"] = posicoes_precalculadas[n][1]

        vis_nodes.append(node_dict)

    vis_edges = []
    for idx_e, (u, v, data) in enumerate(subG_filtrado.edges(data=True)):
        peso = data.get("weight", 1)
        largura = min(4.0, 1.0 + (peso * 0.3))

        vis_edges.append({
            "id": f"e_{idx_e}",
            "from": u,
            "to": v,
            "title": f"Vínculo relacional: {peso} ocorrência(s)",
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


def processar_grafo_dataframe(df_assistencias, alvo_principal="", font_slider=11, espacamento=280):
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
            G_disc.add_node(nid, tipo="cpf", label=nome or formatar_cpf_cnpj(cpf), valor=formatar_cpf_cnpj(cpf), nome_titular=nome or "N/D")
            entidades.append(nid)
        if tel:
            nid = f"TEL_{tel}"
            G_disc.add_node(nid, tipo="telefone", label=formatar_tel(tel), valor=formatar_tel(tel), nome_titular="")
            entidades.append(nid)
        if placa:
            nid = f"PLACA_{placa}"
            fmt_p = f"{placa[:3]}-{placa[3:]}" if len(placa) == 7 else placa
            G_disc.add_node(nid, tipo="placa", label=fmt_p, valor=fmt_p, nome_titular="")
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
        id_caso="",
        espacamento=espacamento
    )

    clean_alvo = re.sub(r'[^a-zA-Z0-9]', '', str(alvo_principal)).upper()
    for n in vis_nodes:
        if clean_alvo and clean_alvo in n["id"]:
            n["size"] = 28
            n["label"] = "★ " + n["label"]

    return vis_nodes, vis_edges, hub_id


def obter_vis_js_local() -> str:
    caminho_js = Path(__file__).parent / "vis-network.min.js"
    if caminho_js.exists():
        try:
            with open(caminho_js, "r", encoding="utf-8") as f:
                return f"<script type='text/javascript'>\n{f.read()}\n</script>"
        except Exception:
            pass
    return '<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/vis-network/9.1.9/standalone/umd/vis-network.min.js"></script>'


def gerar_html_grafo(vis_nodes_json, vis_edges_json, base_font_size=11, hub_id="", target_node_id="", layout_ativo="organico", espacamento=280):
    """
    Fallback usado quando TEM_COMPONENTE_BIDIRECIONAL é False. Restaurado com paridade
    total de funcionalidade em relação ao componente principal: grid pontilhado, busca,
    menu de contexto, drawer redimensionável nos dois eixos, LOD adaptativo, Find Path
    e os mesmos 4 layouts -- adaptado para os nós circulares com ícone SVG. A única
    diferença real é a ausência do canal de volta para o Python (sem Enrich, sem
    sincronização Grafo->Tabela), porque components.html não tem esse retorno.
    """
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
        html, body {{ background: #06090F; overflow: hidden; width: 100%; height: 100%; position: relative; }}
        #network {{ width: 100%; height: 100%; background-color: #06090F; }}

        .left-rail {{
            position: absolute; top: 14px; left: 14px; z-index: 30;
            display: flex; flex-direction: column; gap: 4px; background: rgba(11, 17, 30, 0.92);
            border: 1px solid rgba(51, 65, 85, 0.6); border-radius: 8px; padding: 6px 4px;
            backdrop-filter: blur(14px); box-shadow: 0 4px 18px rgba(0,0,0,0.6);
        }}
        .rail-btn {{
            background: transparent; border: none; color: #94A3B8; width: 34px; height: 32px;
            border-radius: 6px; display: flex; align-items: center; justify-content: center;
            cursor: pointer; transition: 0.15s; outline: none;
        }}
        .rail-btn:hover {{ background: #1E293B; color: #38BDF8; }}
        .rail-btn.active {{ background: #0F2A3D; color: #38BDF8; border: 1px solid rgba(56, 189, 248, 0.4); }}
        .rail-separator {{ width: 22px; height: 1px; background: rgba(51, 65, 85, 0.5); margin: 3px auto; }}

        .search-panel {{
            position: absolute; top: 14px; left: 60px; z-index: 25; width: 220px;
            background: rgba(11, 17, 30, 0.90); border: 1px solid rgba(51, 65, 85, 0.7);
            border-radius: 6px; padding: 6px; backdrop-filter: blur(12px);
            box-shadow: 0 4px 16px rgba(0,0,0,0.5);
        }}
        .search-input-wrap {{ position: relative; display: flex; align-items: center; }}
        .search-panel input {{
            width: 100%; background: #1E293B; border: 1px solid #334155; color: #F8FAFC;
            padding: 6px 24px 6px 8px; border-radius: 4px; font-size: 11px; outline: none;
        }}
        .search-panel input:focus {{ border-color: #38BDF8; }}
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
        .search-result-item:hover {{ border-color: #38BDF8; background: #1E293B; color: #38BDF8; }}

        .minimap-container {{
            position: absolute; bottom: 14px; right: 14px; z-index: 20;
            background: rgba(11, 17, 30, 0.85); border: 1px solid rgba(245, 158, 11, 0.5);
            border-radius: 6px; padding: 4px; backdrop-filter: blur(10px);
            box-shadow: 0 4px 16px rgba(0,0,0,0.6); pointer-events: none;
        }}
        #minimap-canvas {{ display: block; border-radius: 3px; background: rgba(6, 9, 15, 0.8); }}

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
        .context-menu-item:hover {{ background: #0284C7; color: #FFFFFF; }}

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
        .drawer.open {{ transform: translateX(0); }}
        .drawer-resize-handle {{
            position: absolute; left: -4px; top: 0; width: 8px; height: 100%;
            cursor: ew-resize; z-index: 60; background: transparent; transition: background 0.15s;
        }}
        .drawer-resize-handle:hover, .drawer-resize-handle.resizing {{ background: rgba(56, 189, 248, 0.4); }}
        .drawer-resize-handle-v {{
            position: absolute; left: 0; bottom: -4px; width: 100%; height: 8px;
            cursor: ns-resize; z-index: 60; background: transparent; transition: background 0.15s;
        }}
        .drawer-resize-handle-v:hover, .drawer-resize-handle-v.resizing {{ background: rgba(56, 189, 248, 0.4); }}
        .drawer-header {{
            padding: 14px 16px; display: flex; align-items: center; justify-content: space-between;
            border-bottom: 1px solid rgba(51, 65, 85, 0.6); background: rgba(14, 23, 38, 0.7);
        }}
        .drawer-close-btn {{
            background: transparent; border: none; color: #94A3B8; font-size: 20px;
            cursor: pointer; line-height: 1; padding: 2px 6px; border-radius: 4px; transition: 0.15s;
        }}
        .drawer-close-btn:hover {{ color: #F8FAFC; background: #1E293B; }}
        .drawer-badge {{
            font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 4px; text-transform: uppercase;
        }}
        .drawer-body {{ padding: 14px 16px; overflow-y: auto; flex: 1; }}
        .drawer-section {{
            background: rgba(11, 17, 30, 0.7); border: 1px solid rgba(51, 65, 85, 0.5);
            border-radius: 6px; padding: 10px 12px;
        }}
        .drawer-field {{
            display: flex; justify-content: space-between; align-items: center;
            padding: 6px 0; border-bottom: 1px solid rgba(30, 41, 59, 0.5); font-size: 12px;
        }}
        .drawer-field:last-child {{ border-bottom: none; }}
        .field-label {{ color: #94A3B8; }}
        .field-value {{
            color: #F8FAFC; font-weight: 600; text-align: right; max-width: 180px;
            overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
        }}
        .neighbors-container {{ margin-top: 14px; }}
        .neighbors-header {{
            font-size: 11px; font-weight: 700; color: #94A3B8; text-transform: uppercase;
            margin-bottom: 8px; display: flex; justify-content: space-between;
        }}
        .neighbors-preview-wrap {{
            display: flex; justify-content: center; margin: 8px 0 6px 0;
            background: rgba(6, 9, 15, 0.6); border: 1px solid rgba(51, 65, 85, 0.4);
            border-radius: 6px; padding: 6px;
        }}
        #drawer-neighbors-preview {{ display: block; border-radius: 4px; }}
        .neighbors-list {{
            display: flex; flex-direction: column; gap: 6px; max-height: 220px; overflow-y: auto;
        }}
        .neighbor-item {{
            display: flex; align-items: center; justify-content: space-between;
            background: rgba(14, 23, 38, 0.8); border: 1px solid #1E293B; padding: 8px 10px;
            border-radius: 5px; cursor: pointer; transition: 0.15s; font-size: 11px;
        }}
        .neighbor-item:hover {{ border-color: #38BDF8; background: #1E293B; }}
        .neighbor-item-label {{
            color: #F8FAFC; font-weight: 600; max-width: 180px;
            overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
        }}
    </style>
</head>
<body>
    <div class="left-rail">
        <button class="rail-btn" onclick="fitView()" title="Enquadrar">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M8 3H5a2 2 0 0 0-2 2v3m18 0V5a2 2 0 0 0-2-2h-3m0 18h3a2 2 0 0 0 2-2v-3M3 16v3a2 2 0 0 0 2 2h3"/></svg>
        </button>
        <button class="rail-btn" onclick="applyHighlight(null)" title="Limpar Destaque">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>
        </button>
        <button class="rail-btn active" id="rail-lod" onclick="toggleAdaptive()" title="LOD Adaptativo">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
        </button>
        <button class="rail-btn" onclick="exportPNG()" title="Exportar PNG">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
        </button>
        <div class="rail-separator"></div>
        <button class="rail-btn" onclick="executarFindPath()" title="Rastrear Elo entre 2 Nós (Find Path)">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="6" cy="19" r="3"/><circle cx="18" cy="5" r="3"/><path d="M6 16v-3a3 3 0 0 1 3-3h6a3 3 0 0 1 3 3v2"/></svg>
        </button>
        <div class="rail-separator"></div>
        <button class="rail-btn" id="rail-organico" onclick="changeLayout('organico')" title="Força Orgânica">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="3"/><circle cx="19" cy="5" r="2"/><circle cx="5" cy="19" r="2"/><path d="M12 9V5m0 14v-4m-3-3H5m14 0h-4"/></svg>
        </button>
        <button class="rail-btn" id="rail-arvore" onclick="changeLayout('arvore')" title="Hierarquia (Árvore)">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="3" width="6" height="4" rx="1"/><rect x="4" y="17" width="6" height="4" rx="1"/><rect x="14" y="17" width="6" height="4" rx="1"/><path d="M12 7v5m-5 5v-2a3 3 0 0 1 3-3h4a3 3 0 0 1 3 3v2"/></svg>
        </button>
        <button class="rail-btn" id="rail-pontes" onclick="changeLayout('pontes')" title="Pontes & Gargalos">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="4" y1="12" x2="20" y2="12"/><circle cx="12" cy="12" r="4"/><circle cx="4" cy="12" r="2"/><circle cx="20" cy="12" r="2"/></svg>
        </button>
        <button class="rail-btn" id="rail-comunidades" onclick="changeLayout('comunidades')" title="Comunidades Modulares">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>
        </button>
    </div>

    <div id="search-panel" class="search-panel">
        <div class="search-input-wrap">
            <input type="text" id="search-input" placeholder="Filtrar no grafo..." oninput="filtrarListaBusca(this.value)">
            <button id="search-clear-btn" class="search-clear-btn" onclick="limparBuscaGrafo()">&times;</button>
        </div>
        <div id="search-results" class="search-results" style="display:none;"></div>
    </div>

    <div id="minimap-container" class="minimap-container">
        <canvas id="minimap-canvas" width="160" height="110"></canvas>
    </div>

    <div id="context-menu" class="context-menu">
        <div class="context-menu-header" id="ctx-node-title">Entidade</div>
        <div class="context-menu-item" onclick="ctxCopiarDado()">Copiar Identificador</div>
        <div class="context-menu-item" onclick="ctxCentralizar()">Focar no Canvas</div>
        <div class="context-menu-item" onclick="ctxIsolar()">Isolar Vizinhança</div>
        <div class="context-menu-item" onclick="ctxAbrirDrawer()">Inspecionar Atributos</div>
    </div>

    <div id="network"></div>

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
                    <span class="field-label">Identificador:</span>
                    <span id="drawer-val" class="field-value"></span>
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
            </div>
            <div class="neighbors-container">
                <div class="neighbors-header">
                    <span>Entidades Vinculadas</span>
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
        const baseFontSize = {base_font_size};

        const nodesDataSet = new vis.DataSet(rawNodes);
        const edgesDataSet = new vis.DataSet(rawEdges);
        const container = document.getElementById('network');

        const nodeIndexMap = new Map(rawNodes.map(n => [n.id, n]));
        const adjacencyMap = new Map();
        const adjacencySimples = new Map();
        rawNodes.forEach(n => {{ adjacencyMap.set(n.id, []); adjacencySimples.set(n.id, []); }});
        rawEdges.forEach(e => {{
            const title = e.title ? e.title.replace('Vínculo relacional: ', '') : 'Conexão';
            if (adjacencyMap.has(e.from)) adjacencyMap.get(e.from).push({{ id: e.to, title }});
            if (adjacencyMap.has(e.to)) adjacencyMap.get(e.to).push({{ id: e.from, title }});
            if (adjacencySimples.has(e.from)) adjacencySimples.get(e.from).push(e.to);
            if (adjacencySimples.has(e.to)) adjacencySimples.get(e.to).push(e.from);
        }});

        let isAdaptive = true;
        let zoomTimeout = null;
        let currentHighlightId = targetSearchId || null;
        let currentInspectedNodeId = null;
        let contextMenuNodeId = null;
        let selectedNodesQueue = [];

        const network = new vis.Network(container, {{ nodes: nodesDataSet, edges: edgesDataSet }}, {{
            physics: {{ enabled: false }},
            interaction: {{ hover: true, zoomView: true, dragNodes: true, tooltipDelay: 100 }}
        }});

        function fitView() {{ network.fit({{ animation: {{ duration: 350 }}, padding: 60 }}); }}

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
            if (((limites.maxX - limites.minX) / espacamentoGrid) * ((limites.maxY - limites.minY) / espacamentoGrid) > 6000) return;
            const startX = Math.floor(limites.minX / espacamentoGrid) * espacamentoGrid;
            const startY = Math.floor(limites.minY / espacamentoGrid) * espacamentoGrid;
            ctx.save();
            ctx.fillStyle = "rgba(148, 163, 184, 0.16)";
            for (let x = startX; x <= limites.maxX; x += espacamentoGrid) {{
                for (let y = startY; y <= limites.maxY; y += espacamentoGrid) {{ ctx.beginPath(); ctx.arc(x, y, 1.3, 0, 2 * Math.PI); ctx.fill(); }}
            }}
            ctx.restore();
        }});

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
            let maxDeg = -1, best = comp[0];
            comp.forEach(n => {{ if (degrees[n.id] > maxDeg) {{ maxDeg = degrees[n.id]; best = n; }} }});
            return best;
        }}

        function applyHighlight(selectedId) {{
            currentHighlightId = selectedId;
            if (selectedId && nodesDataSet.get(selectedId)) {{
                const connected = network.getConnectedNodes(selectedId);
                connected.push(selectedId);
                const updates = [];
                nodesDataSet.forEach(node => updates.push({{ id: node.id, opacity: connected.includes(node.id) ? 1.0 : 0.12, borderWidth: (node.id === selectedId ? 3.5 : 1.2) }}));
                nodesDataSet.update(updates);
            }} else {{
                const updates = [];
                nodesDataSet.forEach(node => updates.push({{ id: node.id, opacity: 1.0, borderWidth: (node.id === masterHubId ? 3 : 1.2) }}));
                nodesDataSet.update(updates);
            }}
        }}

        function executarFindPath() {{
            if (selectedNodesQueue.length < 2) {{ alert("Clique em 2 nós no grafo (nessa ordem) e depois use este botão para rastrear o elo relacional entre eles."); return; }}
            const startNode = selectedNodesQueue[selectedNodesQueue.length - 2];
            const endNode = selectedNodesQueue[selectedNodesQueue.length - 1];
            const queue = [[startNode]];
            const visited = new Set([startNode]);
            let foundPath = null;
            while (queue.length > 0) {{
                const path = queue.shift();
                const node = path[path.length - 1];
                if (node === endNode) {{ foundPath = path; break; }}
                (adjacencySimples.get(node) || []).forEach(nextNode => {{ if (!visited.has(nextNode)) {{ visited.add(nextNode); queue.push([...path, nextNode]); }} }});
            }}
            if (!foundPath) {{ alert("Nenhum caminho relacional localizado conectando as duas entidades selecionadas na teia atual."); return; }}
            const pathSet = new Set(foundPath);
            const updates = [];
            nodesDataSet.forEach(node => {{ const isInPath = pathSet.has(node.id); updates.push({{ id: node.id, opacity: isInPath ? 1.0 : 0.10, borderWidth: isInPath ? 3.5 : 1.0 }}); }});
            nodesDataSet.update(updates);
        }}

        function desenharMinimapa() {{
            const canvas = document.getElementById('minimap-canvas');
            if (!canvas) return;
            const ctx = canvas.getContext('2d');
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            const positions = network.getPositions();
            const ids = Object.keys(positions);
            if (ids.length === 0) return;
            let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
            ids.forEach(id => {{ const p = positions[id]; if (p) {{ minX = Math.min(minX, p.x); maxX = Math.max(maxX, p.x); minY = Math.min(minY, p.y); maxY = Math.max(maxY, p.y); }} }});
            const rangeX = Math.max(maxX - minX, 1), rangeY = Math.max(maxY - minY, 1), pad = 8;
            function toMini(x, y) {{ return [pad + ((x - minX) / rangeX) * (canvas.width - 2 * pad), pad + ((y - minY) / rangeY) * (canvas.height - 2 * pad)]; }}
            ids.forEach(id => {{
                const p = positions[id];
                if (!p) return;
                const [mx, my] = toMini(p.x, p.y);
                const nodeObj = nodeIndexMap.get(id);
                ctx.fillStyle = (id === masterHubId) ? "#FBBF24" : (nodeObj && nodeObj.color ? nodeObj.color.background : "#38BDF8");
                ctx.beginPath(); ctx.arc(mx, my, (id === masterHubId ? 2.5 : 1.5), 0, 2 * Math.PI); ctx.fill();
            }});
        }}
        network.on("afterDrawing", desenharMinimapa);

        function filtrarListaBusca(termo) {{
            const resultsEl = document.getElementById('search-results');
            const clearBtn = document.getElementById('search-clear-btn');
            if (!termo || termo.trim().length < 2) {{ resultsEl.style.display = "none"; resultsEl.innerHTML = ""; clearBtn.style.display = "none"; return; }}
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
                resultsEl.innerHTML = "<div style='color:#94A3B8; font-size:11px; padding:4px;'>Nenhum registro localizado.</div>";
            }} else {{
                encontrados.forEach(n => {{
                    const item = document.createElement('div');
                    item.className = 'search-result-item';
                    item.innerText = n.label;
                    item.onclick = function() {{ selectInspectedNode(n.id); }};
                    resultsEl.appendChild(item);
                }});
            }}
            resultsEl.style.display = "flex";
        }}
        function limparBuscaGrafo() {{ document.getElementById('search-input').value = ""; filtrarListaBusca(""); }}

        network.on("oncontext", function(params) {{
            params.event.preventDefault();
            const nodeId = network.getNodeAt(params.pointer.DOM);
            const menu = document.getElementById('context-menu');
            if (!nodeId) {{ menu.style.display = "none"; return; }}
            contextMenuNodeId = nodeId;
            const nodeObj = nodeIndexMap.get(nodeId);
            document.getElementById('ctx-node-title').innerText = nodeObj ? nodeObj.label : "Entidade";
            let posX = params.pointer.DOM.x + 8, posY = params.pointer.DOM.y + 8;
            if (posX + 190 > container.clientWidth) posX = params.pointer.DOM.x - 190;
            if (posY + 160 > container.clientHeight) posY = params.pointer.DOM.y - 150;
            menu.style.left = posX + "px"; menu.style.top = posY + "px"; menu.style.display = "block";
        }});
        function hideContextMenu() {{ const menu = document.getElementById('context-menu'); if (menu) menu.style.display = "none"; }}
        document.addEventListener("click", hideContextMenu);
        network.on("dragStart", hideContextMenu);

        function ctxCopiarDado() {{ if (!contextMenuNodeId) return; const nodeObj = nodeIndexMap.get(contextMenuNodeId); if (nodeObj) navigator.clipboard.writeText(nodeObj.valor || nodeObj.id); hideContextMenu(); }}
        function ctxCentralizar() {{ if (contextMenuNodeId) focusNode(contextMenuNodeId); hideContextMenu(); }}
        function ctxIsolar() {{ if (contextMenuNodeId) applyHighlight(contextMenuNodeId); hideContextMenu(); }}
        function ctxAbrirDrawer() {{ if (contextMenuNodeId) {{ openInspector(contextMenuNodeId); applyHighlight(contextMenuNodeId); }} hideContextMenu(); }}

        function desenharNeighborsPreview(nodeId, connectedIds) {{
            const canvas = document.getElementById('drawer-neighbors-preview');
            if (!canvas) return;
            const ctx = canvas.getContext('2d');
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            if (!connectedIds || connectedIds.length === 0) return;
            const cx = canvas.width / 2, cy = canvas.height / 2;
            const exibicaoIds = connectedIds.slice(0, 36);
            const raio = Math.min(50, 24 + exibicaoIds.length * 1.1);
            const posicoes = [];
            ctx.strokeStyle = "rgba(148, 163, 184, 0.25)"; ctx.lineWidth = 1;
            exibicaoIds.forEach((id, i) => {{
                const ang = (2 * Math.PI * i) / exibicaoIds.length;
                const px = cx + raio * Math.cos(ang), py = cy + raio * Math.sin(ang);
                posicoes.push([px, py]);
                ctx.beginPath(); ctx.moveTo(cx, cy); ctx.lineTo(px, py); ctx.stroke();
            }});
            exibicaoIds.forEach((id, i) => {{
                const nObj = nodeIndexMap.get(id);
                const [px, py] = posicoes[i];
                ctx.beginPath(); ctx.arc(px, py, 4.5, 0, 2 * Math.PI);
                ctx.fillStyle = nObj && nObj.color ? nObj.color.background : "#38BDF8"; ctx.fill();
            }});
            const centerObj = nodeIndexMap.get(nodeId);
            ctx.beginPath(); ctx.arc(cx, cy, 7.5, 0, 2 * Math.PI);
            ctx.fillStyle = centerObj && centerObj.color ? centerObj.color.background : "#38BDF8"; ctx.fill();
            ctx.strokeStyle = "#FBBF24"; ctx.lineWidth = 1.8; ctx.stroke();
        }}

        function openInspector(nodeId) {{
            const nodeObj = nodeIndexMap.get(nodeId);
            if (!nodeObj) return;
            currentInspectedNodeId = nodeId;
            const drawer = document.getElementById('inspector-drawer');
            const badge = document.getElementById('drawer-badge');
            const rowNome = document.getElementById('drawer-row-nome');
            badge.innerText = (nodeObj.tipo || "ENTIDADE").toUpperCase();
            badge.style.background = nodeObj.color ? nodeObj.color.background : "#0369A1";
            badge.style.color = "#FFFFFF";
            document.getElementById('drawer-title').innerText = nodeObj.label;
            document.getElementById('drawer-val').innerText = nodeObj.valor || nodeObj.id;
            document.getElementById('drawer-degree').innerText = `${{nodeObj.degree}} conexão(ões)`;
            document.getElementById('drawer-betweenness').innerText = nodeObj.betweenness !== undefined ? `${{nodeObj.betweenness}}` : '0.0';
            document.getElementById('drawer-community').innerText = nodeObj.community !== undefined ? `Grupo #${{nodeObj.community}}` : 'Principal';
            if (nodeObj.tipo === "cpf" && nodeObj.nome_titular) {{ rowNome.style.display = "flex"; document.getElementById('drawer-nome').innerText = nodeObj.nome_titular; }}
            else {{ rowNome.style.display = "none"; }}
            const vizinhos = adjacencyMap.get(nodeId) || [];
            document.getElementById('drawer-neighbors-count').innerText = vizinhos.length;
            desenharNeighborsPreview(nodeId, vizinhos.map(v => v.id));
            const listEl = document.getElementById('drawer-neighbors-list');
            listEl.innerHTML = "";
            const limiteExibicao = 35;
            vizinhos.slice(0, limiteExibicao).forEach(({{ id: nbId, title: edgeTitle }}) => {{
                const nbObj = nodeIndexMap.get(nbId);
                if (!nbObj) return;
                const item = document.createElement('div');
                item.className = 'neighbor-item';
                item.onclick = function() {{ selectInspectedNode(nbId); }};
                item.innerHTML = `<span class="neighbor-item-label">${{nbObj.label}}</span><span style="background:rgba(56, 189, 248, 0.15); color:#38BDF8; padding:2px 6px; border-radius:3px; font-size:9px; font-weight:700; white-space:nowrap;">${{edgeTitle}}</span>`;
                listEl.appendChild(item);
            }});
            if (vizinhos.length > limiteExibicao) {{
                const alerta = document.createElement('div');
                alerta.style.cssText = "color:#94A3B8; font-size:10px; text-align:center; padding:6px;";
                alerta.innerText = `Exibindo 35 de ${{vizinhos.length}} vínculos no painel.`;
                listEl.appendChild(alerta);
            }}
            drawer.classList.add('open');
        }}
        function closeDrawer() {{ document.getElementById('inspector-drawer').classList.remove('open'); currentInspectedNodeId = null; }}
        function selectInspectedNode(nodeId) {{ openInspector(nodeId); applyHighlight(nodeId); focusNode(nodeId); }}
        function focusNode(nodeId) {{ network.focus(nodeId, {{ scale: 1.15, animation: {{ duration: 400, easingFunction: 'easeInOutQuad' }} }}); }}

        (function initDrawerResize() {{
            const handleH = document.getElementById('drawer-resize-handle');
            const handleV = document.getElementById('drawer-resize-handle-v');
            const drawer = document.getElementById('inspector-drawer');
            let resizingH = false, resizingV = false;
            handleH.addEventListener('mousedown', function(e) {{ resizingH = true; handleH.classList.add('resizing'); document.body.style.userSelect = 'none'; e.preventDefault(); }});
            handleV.addEventListener('mousedown', function(e) {{ resizingV = true; handleV.classList.add('resizing'); document.body.style.userSelect = 'none'; e.preventDefault(); }});
            document.addEventListener('mousemove', function(e) {{
                if (resizingH) drawer.style.width = Math.min(640, Math.max(300, window.innerWidth - e.clientX)) + "px";
                if (resizingV) drawer.style.height = Math.min(window.innerHeight, Math.max(220, e.clientY)) + "px";
            }});
            document.addEventListener('mouseup', function() {{
                if (resizingH) {{ resizingH = false; handleH.classList.remove('resizing'); }}
                if (resizingV) {{ resizingV = false; handleV.classList.remove('resizing'); }}
                document.body.style.userSelect = '';
            }});
        }})();

        function updateAdaptiveLOD() {{
            if (!isAdaptive) return;
            const scale = network.getScale();
            if (!scale || scale <= 0.05) return;
            const targetSize = Math.max(9, Math.min(20, Math.round(baseFontSize / Math.sqrt(scale))));
            const updates = [];
            nodesDataSet.forEach(n => updates.push({{ id: n.id, font: {{ size: targetSize, color: "#F8FAFC", face: "Segoe UI", strokeWidth: 2.5, strokeColor: "#06090F" }} }}));
            nodesDataSet.update(updates);
        }}
        function toggleAdaptive() {{ isAdaptive = !isAdaptive; document.getElementById('rail-lod').classList.toggle('active', isAdaptive); if (isAdaptive) updateAdaptiveLOD(); }}
        network.on("zoom", function() {{
            hideContextMenu();
            if (isAdaptive) {{ if (zoomTimeout) clearTimeout(zoomTimeout); zoomTimeout = setTimeout(updateAdaptiveLOD, 120); }}
        }});

        function exportPNG() {{
            const canvas = container.querySelector('canvas');
            const a = document.createElement('a');
            a.download = 'evidencia_caso_lcfo.png';
            a.href = canvas.toDataURL('image/png');
            a.click();
        }}

        function marcarLayoutAtivo(mode) {{
            ['organico', 'arvore', 'pontes', 'comunidades'].forEach(m => {{
                const el = document.getElementById(`rail-${{m}}`);
                if (el) el.classList.toggle('active', m === mode);
            }});
        }}

        function changeLayout(mode) {{
            marcarLayoutAtivo(mode);
            if (mode === "organico") {{
                const updates = rawNodes.filter(n => n.x !== undefined && n.y !== undefined).map(n => ({{ id: n.id, x: n.x, y: n.y }}));
                nodesDataSet.update(updates);
                setTimeout(() => {{ fitView(); if (currentHighlightId) applyHighlight(currentHighlightId); }}, 60);
            }} else if (mode === "arvore") {{
                network.setOptions({{ layout: {{ hierarchical: {{ direction: "UD", sortMethod: "hubsize", levelSeparation: 170, nodeSpacing: 210 }} }} }});
                setTimeout(() => {{ network.setOptions({{ layout: {{ hierarchical: false }} }}); fitView(); if (currentHighlightId) applyHighlight(currentHighlightId); }}, 350);
            }} else if (mode === "pontes") {{
                const updates = [];
                const bridges = rawNodes.filter(n => n.betweenness > 0.001).sort((a, b) => b.betweenness - a.betweenness);
                const nonBridges = rawNodes.filter(n => !(n.betweenness > 0.001));
                const bridgeSpacing = Math.max(110, Math.min(200, 1300 / Math.max(bridges.length, 1)));
                const startY = Math.round(-((bridges.length - 1) * bridgeSpacing) / 2);
                bridges.forEach((bNode, bIdx) => updates.push({{ id: bNode.id, x: 0, y: startY + (bIdx * bridgeSpacing) }}));
                const leftNodes = [], rightNodes = [];
                nonBridges.forEach((nb, i) => (i % 2 === 0 ? leftNodes : rightNodes).push(nb));
                const flankSpacing = 80;
                leftNodes.forEach((n, i) => updates.push({{ id: n.id, x: -820, y: Math.round(-((leftNodes.length - 1) * flankSpacing) / 2) + (i * flankSpacing) }}));
                rightNodes.forEach((n, i) => updates.push({{ id: n.id, x: 820, y: Math.round(-((rightNodes.length - 1) * flankSpacing) / 2) + (i * flankSpacing) }}));
                nodesDataSet.update(updates);
                setTimeout(() => {{ fitView(); if (currentHighlightId) applyHighlight(currentHighlightId); }}, 60);
            }} else if (mode === "comunidades") {{
                const updates = [];
                const commGroups = {{}};
                rawNodes.forEach(n => {{ const cid = n.community || 0; if (!commGroups[cid]) commGroups[cid] = []; commGroups[cid].push(n); }});
                const commKeys = Object.keys(commGroups).sort((a, b) => commGroups[b].length - commGroups[a].length);
                const cols = commKeys.length <= 2 ? Math.max(commKeys.length, 1) : 3;
                const rows = Math.ceil(commKeys.length / cols);
                const commSpacingX = 1300, commSpacingY = 1100;
                commKeys.forEach((cid, idx) => {{
                    const cNodes = commGroups[cid];
                    const col = idx % cols, row = Math.floor(idx / cols);
                    const cx = Math.round((col - (cols - 1) / 2) * commSpacingX);
                    const cy = Math.round((row - (rows - 1) / 2) * commSpacingY);
                    const localHub = cNodes.find(n => n.id === masterHubId) || getLocalHub(cNodes, rawEdges);
                    const others = cNodes.filter(n => n.id !== localHub.id);
                    const rInternal = Math.max(120, Math.min(520, others.length * 26));
                    updates.push({{ id: localHub.id, x: cx, y: cy }});
                    others.forEach((n, nIdx) => {{
                        const a = (2 * Math.PI * nIdx) / Math.max(others.length, 1);
                        updates.push({{ id: n.id, x: Math.round(cx + (rInternal * Math.cos(a))), y: Math.round(cy + (rInternal * Math.sin(a))) }});
                    }});
                }});
                nodesDataSet.update(updates);
                setTimeout(() => {{ fitView(); if (currentHighlightId) applyHighlight(currentHighlightId); }}, 60);
            }}
        }}

        network.on("click", function(params) {{
            if (params.nodes.length > 0) {{
                const clickedId = params.nodes[0];
                selectedNodesQueue.push(clickedId);
                if (selectedNodesQueue.length > 6) selectedNodesQueue.shift();
                applyHighlight(clickedId);
                openInspector(clickedId);
            }} else {{ applyHighlight(null); closeDrawer(); }}
        }});

        setTimeout(function() {{ fitView(); if (targetSearchId) selectInspectedNode(targetSearchId); if (isAdaptive) updateAdaptiveLOD(); }}, 80);
    </script>
</body>
</html>
"""