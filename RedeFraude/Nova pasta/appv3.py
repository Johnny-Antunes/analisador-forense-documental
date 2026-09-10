import streamlit as st
import pandas as pd
import json
import re
from pathlib import Path
import streamlit.components.v1 as components
import pydeck as pdk

from database import (
    CAMINHO_REDE_OFICIAL, PASTA_LOCAL_BLACKLIST,
    CAMINHO_REDE_CRIACAO, PASTA_LOCAL_CRIACAO,
    carregar_arquivos_para_sqlite, carregar_criacoes_diarias_para_sqlite,
    consultar_detalhes_caso, cruzar_com_criacoes_diarias, obter_radar_expansoes
)
from graph_engine import carregar_redes, processar_subgrafo_caso, gerar_html_grafo

# =====================================================
# 1. CONFIGURAÇÃO DE PÁGINA & CSS FORENSE LIMPO
# =====================================================
st.set_page_config(
    page_title="Mesa de Inteligência de Fraudes - LCFO",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .stDeployButton, .stAppDeployButton, #MainMenu, footer, [data-testid="stDecoration"] {
        display: none !important;
        visibility: hidden !important;
    }

    header[data-testid="stHeader"] {
        background: transparent !important;
    }

    [data-testid="stSidebarCollapseButton"] button,
    [data-testid="collapsedControl"] button,
    [data-testid="stSidebarCollapsedControl"] button {
        background: #0B111E !important;
        border: 1px solid #1E293B !important;
        color: #38BDF8 !important;
        border-radius: 6px !important;
    }

    .block-container {
        padding-top: 2.2rem !important;
        padding-bottom: 0.8rem !important;
        padding-left: 1.2rem !important;
        padding-right: 1.2rem !important;
        max-width: 100% !important;
    }
    
    .hud-bar {
        background: #0B111E;
        border: 1px solid #1E293B;
        border-radius: 8px;
        padding: 8px 14px;
        margin-top: 4px;
        margin-bottom: 10px;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    .intel-card {
        background: #0E1726;
        border: 1px solid #1E293B;
        border-radius: 6px;
        padding: 10px 14px;
        text-align: center;
    }
    .kpi-card {
        background: #0B111E;
        border: 1px solid #1E293B;
        border-radius: 8px;
        padding: 14px 18px;
        text-align: left;
    }
</style>
""", unsafe_allow_html=True)

# =====================================================
# 2. DEFINIÇÕES GLOBAIS E ESTADO PERSISTENTE DESACOPLADO
# =====================================================
LAYOUT_MAP = {
    "organico": "🌀 Teia Fluida Orgânica",
    "radial": "🎯 Radial Peacock (i2)",
    "arvore": "🌲 Árvore Forense",
    "pontes": "🌉 Pontes & Gargalos",
    "subredes": "🧩 Sub-redes em Grade",
    "comunidades": "🏛️ Comunidades & Facções"
}

OPCOES_PAINEL = [
    "📍 Radar Territorial (Mapa)",
    "🕸️ Grafo de Vínculos",
    "📋 Tabela de Ocorrências",
    "🕒 Evolução Temporal",
    "🔍 Expansão com Criações"
]

PROPORCOES_MAP = {
    "50% | 50%": [1.0, 1.0],
    "60% | 40%": [1.5, 1.0],
    "40% | 60%": [1.0, 1.5],
    "70% | 30%": [2.3, 1.0],
    "30% | 70%": [1.0, 2.3]
}

if "celula_ativa_id" not in st.session_state:
    st.session_state["celula_ativa_id"] = None

if "cfg_layout_ativo" not in st.session_state:
    st.session_state["cfg_layout_ativo"] = "organico"

if "cfg_modo_exibicao" not in st.session_state:
    st.session_state["cfg_modo_exibicao"] = "📑 Abas Clássicas"

if "cfg_painel_esquerdo" not in st.session_state:
    st.session_state["cfg_painel_esquerdo"] = "📍 Radar Territorial (Mapa)"

if "cfg_painel_direito" not in st.session_state:
    st.session_state["cfg_painel_direito"] = "🕸️ Grafo de Vínculos"

if "cfg_cockpit_proporcao" not in st.session_state:
    st.session_state["cfg_cockpit_proporcao"] = "50% | 50%"

def cb_atualizar_modo():
    st.session_state["cfg_modo_exibicao"] = st.session_state["w_modo_exibicao"]

def cb_atualizar_layout():
    st.session_state["cfg_layout_ativo"] = st.session_state["w_layout_ativo"]

def cb_atualizar_pesq():
    st.session_state["cfg_painel_esquerdo"] = st.session_state["w_painel_esq"]

def cb_atualizar_pdir():
    st.session_state["cfg_painel_direito"] = st.session_state["w_painel_dir"]

def cb_atualizar_prop():
    st.session_state["cfg_cockpit_proporcao"] = st.session_state["w_proporcao"]

# =====================================================
# 3. GESTÃO DE BASES (SIDEBAR)
# =====================================================
with st.sidebar:
    st.markdown("### ⚙️ Gestão de Bases de Dados")
    
    with st.expander("📁 1. Base Blacklist (Duplicidades)", expanded=True):
        pasta_padrao_bl = CAMINHO_REDE_OFICIAL if Path(CAMINHO_REDE_OFICIAL).exists() else str(PASTA_LOCAL_BLACKLIST)
        caminho_input = st.text_input("Pasta Blacklist (.xlsx):", value=pasta_padrao_bl)
        
        qtd_detectada = 0
        p_check = Path(caminho_input.strip())
        if p_check.exists():
            qtd_detectada = len(list(p_check.glob("*.xlsx"))) + len(list(p_check.glob("*.csv")))
            st.caption(f"🟢 **{qtd_detectada}** planilhas detectadas.")
        else:
            st.caption("🔴 Caminho não acessível.")

        if st.button("🔄 Ingerir Base Blacklist", use_container_width=True):
            with st.spinner("Atualizando registros de Blacklist..."):
                qtd_arq, qtd_reg, msg_erro, erros_bl = carregar_arquivos_para_sqlite(caminho_input)
                st.cache_data.clear()
                if qtd_arq > 0:
                    st.success(f"✅ {qtd_arq - len(erros_bl)}/{qtd_arq} planilhas ({qtd_reg:,} assistências).")
                    if erros_bl:
                        with st.expander(f"⚠️ {len(erros_bl)} arquivos com aviso"):
                            for e in erros_bl:
                                st.caption(f"• {e}")
                    st.rerun()
                else:
                    st.warning(msg_erro or "Nenhuma planilha encontrada.")

    with st.expander("🌐 2. Criações Diárias (Base Geral)", expanded=False):
        pasta_padrao_cr = CAMINHO_REDE_CRIACAO if Path(CAMINHO_REDE_CRIACAO).exists() else str(PASTA_LOCAL_CRIACAO)
        caminho_criacao = st.text_input("Pasta Criação (.xlsx/.csv):", value=pasta_padrao_cr)
        
        qtd_criacao = 0
        p_criacao = Path(caminho_criacao.strip())
        if p_criacao.exists():
            qtd_criacao = len(list(p_criacao.glob("*.xlsx"))) + len(list(p_criacao.glob("*.csv")))
            st.caption(f"🟢 **{qtd_criacao}** arquivos detectados na pasta.")
        else:
            st.caption("🔴 Caminho não acessível.")

        if st.button("🔄 Ingerir Criações no SQLite", use_container_width=True):
            with st.spinner("Atualizando registros de Criações Diárias..."):
                qtd_arq_c, qtd_reg_c, msg_c, erros_c = carregar_criacoes_diarias_para_sqlite(caminho_criacao)
                if qtd_arq_c > 0:
                    st.success(f"✅ {qtd_arq_c - len(erros_c)}/{qtd_arq_c} arquivos ({qtd_reg_c:,} assistências ingeridas).")
                    if erros_c:
                        with st.expander(f"⚠️ {len(erros_c)} arquivos com inconsistências"):
                            for e in erros_c:
                                st.caption(f"• {e}")
                    st.rerun()
                else:
                    st.warning(msg_c or "Nenhum arquivo encontrado.")

# =====================================================
# 4. CARREGAMENTO DAS REDES E VALIDAÇÃO DA BASE
# =====================================================
G, cluster_info = carregar_redes()

if not G or not cluster_info:
    st.info("👈 Nenhum dado processado no banco. Verifique o caminho da pasta na barra lateral e clique em **'🔄 Ingerir Base Blacklist'**.")
    st.stop()

radar_alertas = obter_radar_expansoes(cluster_info)

opcoes_celulas = {}
for c in cluster_info[:150]:
    alerta_badge = f" 🚨 (+{radar_alertas[c['id']]} criações)" if c["id"] in radar_alertas else ""
    rotulo = f"Célula #{c['id']} ({c['tamanho']} nós) - {c['hub_label']}{alerta_badge}"
    opcoes_celulas[rotulo] = c['id']

mapa_id_para_label = {v: k for k, v in opcoes_celulas.items()}

# =====================================================
# 5. CONTROLES NA BARRA LATERAL
# =====================================================
with st.sidebar:
    st.markdown("---")
    
    if st.session_state["celula_ativa_id"] is not None:
        if st.button("⬅️ Voltar à Triagem Geral", use_container_width=True):
            st.session_state["celula_ativa_id"] = None
            st.rerun()
        
        st.markdown("### 📂 Caso em Análise")
        
        label_atual = mapa_id_para_label.get(st.session_state["celula_ativa_id"], list(opcoes_celulas.keys())[0])
        idx_caso = list(opcoes_celulas.keys()).index(label_atual) if label_atual in opcoes_celulas else 0
        novo_caso_selecionado = st.selectbox("Alternar Caso:", list(opcoes_celulas.keys()), index=idx_caso)
        
        if opcoes_celulas[novo_caso_selecionado] != st.session_state["celula_ativa_id"]:
            st.session_state["celula_ativa_id"] = opcoes_celulas[novo_caso_selecionado]
            st.rerun()

        cluster_selecionado = next((c for c in cluster_info if c["id"] == st.session_state["celula_ativa_id"]), cluster_info[0])

        st.markdown("---")
        st.markdown("### 📐 Layout do Grafo")
        
        idx_layout = list(LAYOUT_MAP.keys()).index(st.session_state["cfg_layout_ativo"]) if st.session_state["cfg_layout_ativo"] in LAYOUT_MAP else 0
        st.selectbox(
            "Estrutura Topológica:",
            options=list(LAYOUT_MAP.keys()),
            format_func=lambda k: LAYOUT_MAP[k],
            index=idx_layout,
            key="w_layout_ativo",
            on_change=cb_atualizar_layout
        )

        st.markdown("---")
        st.markdown("### 🎯 Filtros de Entidades")
        mostrar_tel = st.checkbox(f"🚨 Telefones ({cluster_selecionado['qtd_tels']})", value=True)
        mostrar_cpf = st.checkbox(f"👤 Titulares ({cluster_selecionado['qtd_cpfs']})", value=True)
        mostrar_placa = st.checkbox(f"🚗 Placas ({cluster_selecionado['qtd_placas']})", value=True)

        st.markdown("---")
        st.markdown("### 🔤 Ajuste Visual")
        font_slider = st.slider("Fonte Base do Grafo (px):", min_value=10, max_value=18, value=12)
        espacamento_slider = st.slider("Dispersão / Distância (px):", min_value=180, max_value=400, value=260, step=20)

    else:
        st.markdown("### 🔍 Busca Rápida de Entidade")
        termo_busca_global = st.text_input("CPF, Placa ou Telefone:", placeholder="Digite para rastrear...").strip()
        
        if termo_busca_global:
            termo_clean = re.sub(r'[^a-zA-Z0-9]', '', termo_busca_global).upper()
            alvos = [n for n in G.nodes if termo_clean in n]
            if alvos:
                alvo = alvos[0]
                for c in cluster_info:
                    if alvo in c["nodes"]:
                        st.success(f"🎯 Localizado na Célula #{c['id']}")
                        if st.button(f"Abrir Célula #{c['id']}", use_container_width=True):
                            st.session_state["celula_ativa_id"] = c["id"]
                            st.rerun()
                        break
            else:
                st.warning("Nenhum vínculo localizado na Blacklist.")

# ====================================================================
# 6. TELA 1: PAINEL GERAL & TRIAGEM DE CASOS
# ====================================================================
if st.session_state["celula_ativa_id"] is None:
    st.markdown("<h3 style='margin:0; color:#38BDF8;'>🛡️ Mesa de Inteligência Forense - LCFO</h3>", unsafe_allow_html=True)
    st.caption("Painel Executivo de Triagem, Radar de Expansões e Monitoramento Contínuo")
    
    st.markdown("<br/>", unsafe_allow_html=True)
    
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(f"""
        <div class="kpi-card">
            <span style="color:#94A3B8; font-size:12px; font-weight:600;">CÉLULAS MAPEADAS</span><br/>
            <b style="font-size:26px; color:#F8FAFC;">{len(cluster_info):,}</b>
        </div>
        """, unsafe_allow_html=True)
    with k2:
        st.markdown(f"""
        <div class="kpi-card">
            <span style="color:#94A3B8; font-size:12px; font-weight:600;">ENTIDADES NA BLACKLIST</span><br/>
            <b style="font-size:26px; color:#38BDF8;">{len(G.nodes):,}</b>
        </div>
        """, unsafe_allow_html=True)
    with k3:
        st.markdown(f"""
        <div class="kpi-card">
            <span style="color:#94A3B8; font-size:12px; font-weight:600;">CONEXÕES RELACIONAIS</span><br/>
            <b style="font-size:26px; color:#A78BFA;">{len(G.edges):,}</b>
        </div>
        """, unsafe_allow_html=True)
    with k4:
        qtd_com_alerta = len(radar_alertas)
        cor_alerta = "#EF4444" if qtd_com_alerta > 0 else "#10B981"
        st.markdown(f"""
        <div class="kpi-card">
            <span style="color:#94A3B8; font-size:12px; font-weight:600;">CÉLULAS REINCIDENTES</span><br/>
            <b style="font-size:26px; color:{cor_alerta};">{qtd_com_alerta} no Radar</b>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br/>", unsafe_allow_html=True)

    if radar_alertas:
        st.markdown(f"##### 🚨 Radar de Alertas: {len(radar_alertas)} Células Reincidentes nas Criações Diárias")
        cols_radar = st.columns(min(4, len(radar_alertas)))
        for i, (cid, total_hits) in enumerate(list(radar_alertas.items())[:4]):
            with cols_radar[i]:
                c_obj = next((c for c in cluster_info if c["id"] == cid), None)
                if c_obj:
                    if st.button(f"🚨 Célula #{cid}: +{total_hits} ocorrências\nÂncora: {c_obj['hub_label'][:22]}", use_container_width=True):
                        st.session_state["celula_ativa_id"] = cid
                        st.rerun()

    st.markdown("---")
    st.markdown("#### 📋 Fila de Investigação Operacional")
    
    col_sel1, col_sel2 = st.columns([3.5, 1.5])
    with col_sel1:
        caso_escolhido = st.selectbox("Selecione um caso para abrir a mesa de análise:", list(opcoes_celulas.keys()))
    with col_sel2:
        st.write("")
        st.write("")
        if st.button("🔬 Abrir Mesa Forense", use_container_width=True, type="primary"):
            st.session_state["celula_ativa_id"] = opcoes_celulas[caso_escolhido]
            st.rerun()

    tabela_casos = []
    for c in cluster_info[:100]:
        status_radar = f"🚨 +{radar_alertas[c['id']]} novas" if c["id"] in radar_alertas else "Estável"
        tabela_casos.append({
            "Célula": f"#{c['id']}",
            "Âncora Central (Hub)": c["hub_label"],
            "Tamanho": c["tamanho"],
            "Telefones": c["qtd_tels"],
            "Titulares": c["qtd_cpfs"],
            "Placas": c["qtd_placas"],
            "Status no Radar": status_radar
        })
    df_resumo_casos = pd.DataFrame(tabela_casos)
    st.dataframe(df_resumo_casos, use_container_width=True, hide_index=True)

# ====================================================================
# 7. TELA 2: MESA DE INVESTIGAÇÃO (ABAS CLÁSSICAS + COCKPIT DIVIDIDO)
# ====================================================================
else:
    cluster_selecionado = next((c for c in cluster_info if c["id"] == st.session_state["celula_ativa_id"]), cluster_info[0])

    lista_cpfs = [n.replace("CPF_", "") for n in cluster_selecionado["nodes"] if n.startswith("CPF_")]
    lista_tels = [n.replace("TEL_", "") for n in cluster_selecionado["nodes"] if n.startswith("TEL_")]
    lista_placas = [n.replace("PLACA_", "") for n in cluster_selecionado["nodes"] if n.startswith("PLACA_")]

    df_detalhes = consultar_detalhes_caso(lista_cpfs, lista_tels, lista_placas)
    total_assistencias = len(df_detalhes)

    periodo_str = "Sem datas"
    if not df_detalhes.empty and "data" in df_detalhes.columns:
        df_temp = df_detalhes[df_detalhes["data"].str.strip() != ""].copy()
        if not df_temp.empty:
            dts_validas = pd.to_datetime(df_temp["data"], errors="coerce").dropna()
            if not dts_validas.empty:
                periodo_str = f"{dts_validas.min().strftime('%d/%m/%Y')} → {dts_validas.max().strftime('%d/%m/%Y')}"

    # Barra Superior com Navegação e Modo de Exibição
    col_nav1, col_nav2, col_nav3 = st.columns([3.2, 1.8, 1.0])
    with col_nav1:
        st.markdown(f"<div style='padding-top:4px;'><span style='color:#94A3B8; font-size:12px;'>Mesa de Triagem / </span> <b style='color:#38BDF8; font-size:14px;'>Célula #{cluster_selecionado['id']}</b></div>", unsafe_allow_html=True)
    with col_nav2:
        idx_modo = 0 if st.session_state["cfg_modo_exibicao"] == "📑 Abas Clássicas" else 1
        st.radio(
            "Modo de Visualização:", 
            ["📑 Abas Clássicas", "🖥️ Cockpit Dividido"], 
            index=idx_modo,
            horizontal=True, 
            key="w_modo_exibicao",
            on_change=cb_atualizar_modo,
            label_visibility="collapsed"
        )
    with col_nav3:
        if st.button("⬅️ Sair da Célula", use_container_width=True):
            st.session_state["celula_ativa_id"] = None
            st.rerun()

    # HUD Tático
    st.markdown(f"""
    <div class="hud-bar">
        <div style="display:flex; align-items:center; gap:20px;">
            <span style="font-size:0.95rem; font-weight:700; color:#F8FAFC;">
                🚨 CÉLULA #{cluster_selecionado['id']}
            </span>
            <span style="font-size:0.82rem; color:#94A3B8;">
                Âncora Central: <b style="color:#FBBF24;">{cluster_selecionado['hub_label']}</b>
            </span>
            <span style="font-size:0.82rem; color:#94A3B8;">
                Total de Assistências: <b style="color:#38BDF8;">{total_assistencias} ocorrências</b>
            </span>
            <span style="font-size:0.82rem; color:#94A3B8;">
                Período Ativo: <b style="color:#E2E8F0;">{periodo_str}</b>
            </span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Subgrafo
    filtro_tipos = []
    if mostrar_tel: filtro_tipos.append("telefone")
    if mostrar_cpf: filtro_tipos.append("cpf")
    if mostrar_placa: filtro_tipos.append("placa")

    vis_nodes, vis_edges = processar_subgrafo_caso(
        G=G,
        cluster_nodes=cluster_selecionado["nodes"],
        filtro_tipos=filtro_tipos,
        hub_id=cluster_selecionado["hub_id"],
        font_slider=font_slider
    )

    html_codigo = gerar_html_grafo(
        vis_nodes_json=json.dumps(vis_nodes),
        vis_edges_json=json.dumps(vis_edges),
        base_font_size=font_slider,
        hub_id=cluster_selecionado["hub_id"],
        target_node_id="",
        layout_ativo=st.session_state["cfg_layout_ativo"],
        espacamento=espacamento_slider
    )

    # Função Renderizadora das Visões
    def render_visao(nome_visao, altura_pixel=720):
        if "Grafo" in nome_visao:
            components.html(html_codigo, height=altura_pixel)
            c_ginfo, c_gbtn = st.columns([3.5, 1.5])
            with c_ginfo:
                st.caption(f"Cluster Forense com {len(vis_nodes)} entidades ativas e {len(vis_edges)} conexões ponderadas.")
            with c_gbtn:
                st.download_button(
                    label="💾 Baixar Dossiê HTML deste Caso",
                    data=html_codigo,
                    file_name=f"Dossie_Forense_Caso_{cluster_selecionado['id']}.html",
                    mime="text/html",
                    use_container_width=True
                )

        elif "Radar Territorial" in nome_visao or "Mapa" in nome_visao:
            if not df_detalhes.empty and "latitude" in df_detalhes.columns:
                df_mapa = df_detalhes.dropna(subset=["latitude", "longitude"]).copy()
                if not df_mapa.empty:
                    lat_centro = float(df_mapa["latitude"].mean())
                    lon_centro = float(df_mapa["longitude"].mean())

                    camada_pontos = pdk.Layer(
                        "ScatterplotLayer",
                        data=df_mapa,
                        get_position="[longitude, latitude]",
                        get_color="[56, 189, 248, 180]",
                        get_line_color="[255, 255, 255, 220]",
                        line_width_min_pixels=1,
                        stroked=True,
                        get_radius=3200,
                        radius_min_pixels=6,
                        radius_max_pixels=20,
                        pickable=True,
                        auto_highlight=True,
                    )

                    view_state = pdk.ViewState(
                        latitude=lat_centro,
                        longitude=lon_centro,
                        zoom=7,
                        pitch=0
                    )

                    tooltip = {
                        "html": """
                            <div style="font-family:'Segoe UI', Tahoma, sans-serif; padding:4px;">
                                <b style="color:#38BDF8;">🚨 Assistência: {id_assistencia}</b><br/>
                                <b>Serviço:</b> {servico}<br/>
                                <b>Titular:</b> {titular}<br/>
                                <b>CPF:</b> {cpf}<br/>
                                <b>Placa:</b> {placa}<br/>
                                <b>Local:</b> {bairro}, {cidade}/{uf}<br/>
                                <b>Data:</b> {data}
                            </div>
                        """,
                        "style": {
                            "backgroundColor": "rgba(11, 17, 30, 0.95)",
                            "color": "#F8FAFC",
                            "border": "1px solid #1E293B",
                            "borderRadius": "6px",
                            "fontSize": "11px",
                            "zIndex": "1000"
                        }
                    }

                    deck = pdk.Deck(
                        layers=[camada_pontos],
                        initial_view_state=view_state,
                        tooltip=tooltip,
                        map_style="dark"
                    )

                    st.pydeck_chart(deck, use_container_width=True)
                    resumo_cidades = df_detalhes.groupby(["cidade", "uf"]).size().reset_index(name="Total_Assistencias").sort_values(by="Total_Assistencias", ascending=False)
                    st.dataframe(resumo_cidades, use_container_width=True, hide_index=True)
                else:
                    st.info("Sem coordenadas válidas para plotagem.")
            else:
                st.info("Sem dados de localização.")

        elif "Tabela" in nome_visao:
            st.dataframe(
                df_detalhes[["data", "id_assistencia", "servico", "titular", "cpf", "telefone", "placa", "bairro", "cidade", "uf"]],
                use_container_width=True,
                hide_index=True
            )
            csv_dados = df_detalhes.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Baixar Ocorrências em CSV",
                data=csv_dados,
                file_name=f"Assistencias_Caso_{cluster_selecionado['id']}.csv",
                mime="text/csv"
            )

        elif "Temporal" in nome_visao:
            if not df_detalhes.empty and "data" in df_detalhes.columns:
                df_temporal = df_detalhes.copy()
                df_temporal["dt_parsed"] = pd.to_datetime(df_temporal["data"], errors="coerce")
                df_temporal = df_temporal.dropna(subset=["dt_parsed"])
                
                if not df_temporal.empty:
                    df_temporal["Periodo"] = df_temporal["dt_parsed"].dt.strftime("%Y-%m")
                    serie_temporal = df_temporal.groupby("Periodo").size().reset_index(name="Quantidade de Assistências")
                    serie_temporal = serie_temporal.sort_values(by="Periodo")
                    
                    st.bar_chart(
                        data=serie_temporal.set_index("Periodo"),
                        color="#38BDF8",
                        use_container_width=True
                    )
                else:
                    st.info("Datas sem padrão cronológico reconhecido.")
            else:
                st.info("Sem datas disponíveis.")

        elif "Expansão" in nome_visao:
            st.markdown(f"#### 🔍 Cruzamento da Célula #{cluster_selecionado['id']} contra as Criações Diárias")
            st.caption("Verifica acionamentos no universo total da companhia e extrai vínculos inéditos.")

            resumo_intel, df_matches_criacao, df_novos_suspeitos = cruzar_com_criacoes_diarias(lista_cpfs, lista_tels, lista_placas)

            if resumo_intel is None:
                st.info("💡 Criações Diárias ainda não ingeridas.")
            elif resumo_intel["total_assistencias"] == 0:
                st.success("✅ Sem acionamentos adicionais nas Criações Diárias.")
            else:
                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    st.markdown(f"""<div class="intel-card"><span style="color:#94A3B8; font-size:10px;">ACIONAMENTOS</span><br/><b style="font-size:18px; color:#38BDF8;">{resumo_intel['total_assistencias']}</b></div>""", unsafe_allow_html=True)
                with c2:
                    st.markdown(f"""<div class="intel-card"><span style="color:#94A3B8; font-size:10px;">NOVOS TELS</span><br/><b style="font-size:18px; color:#EF4444;">+{resumo_intel['novos_tels']}</b></div>""", unsafe_allow_html=True)
                with c3:
                    st.markdown(f"""<div class="intel-card"><span style="color:#94A3B8; font-size:10px;">NOVAS PLACAS</span><br/><b style="font-size:18px; color:#A78BFA;">+{resumo_intel['novas_placas']}</b></div>""", unsafe_allow_html=True)
                with c4:
                    st.markdown(f"""<div class="intel-card"><span style="color:#94A3B8; font-size:10px;">NOVOS CPFS</span><br/><b style="font-size:18px; color:#FBBF24;">+{resumo_intel['novos_cpfs']}</b></div>""", unsafe_allow_html=True)

                st.markdown("<br/>", unsafe_allow_html=True)
                if not df_novos_suspeitos.empty:
                    st.markdown(f"**🚨 {len(df_novos_suspeitos)} Novos Suspeitos Detectados:**")
                    st.dataframe(df_novos_suspeitos, use_container_width=True, hide_index=True)
                    csv_suspeitos = df_novos_suspeitos.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Exportar Suspeitos (.csv)",
                        data=csv_suspeitos,
                        file_name=f"Novos_Suspeitos_Caso_{cluster_selecionado['id']}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
                else:
                    st.info("Acionamentos pertencem apenas a entidades já catalogadas.")

    # =====================================================
    # EXIBIÇÃO: MODO 1 (ABAS CLÁSSICAS - PADRÃO)
    # =====================================================
    if st.session_state["cfg_modo_exibicao"] == "📑 Abas Clássicas":
        tab_g, tab_e, tab_t, tab_m, tab_d = st.tabs([
            "🕸️ Grafo de Vínculos Forense",
            "🔍 Expansão com Criações Diárias",
            "🕒 Evolução Temporal",
            "📍 Radar Territorial",
            "📋 Tabela de Ocorrências"
        ])
        with tab_g:
            render_visao("Grafo", altura_pixel=740)
        with tab_e:
            render_visao("Expansão")
        with tab_t:
            render_visao("Temporal")
        with tab_m:
            render_visao("Radar Territorial")
        with tab_d:
            render_visao("Tabela")

    # =====================================================
    # EXIBIÇÃO: MODO 2 (COCKPIT DIVIDIDO LADO A LADO)
    # =====================================================
    else:
        c_pesq, c_prop, c_pdir = st.columns([2.5, 1.4, 2.5])
        with c_pesq:
            idx_pesq = OPCOES_PAINEL.index(st.session_state["cfg_painel_esquerdo"]) if st.session_state["cfg_painel_esquerdo"] in OPCOES_PAINEL else 0
            st.selectbox(
                "Painel Esquerdo:", 
                OPCOES_PAINEL, 
                index=idx_pesq,
                key="w_painel_esq",
                on_change=cb_atualizar_pesq
            )
        with c_prop:
            idx_prop = list(PROPORCOES_MAP.keys()).index(st.session_state["cfg_cockpit_proporcao"]) if st.session_state["cfg_cockpit_proporcao"] in PROPORCOES_MAP else 0
            st.selectbox(
                "Proporção:", 
                list(PROPORCOES_MAP.keys()), 
                index=idx_prop,
                key="w_proporcao",
                on_change=cb_atualizar_prop
            )
        with c_pdir:
            idx_pdir = OPCOES_PAINEL.index(st.session_state["cfg_painel_direito"]) if st.session_state["cfg_painel_direito"] in OPCOES_PAINEL else 1
            st.selectbox(
                "Painel Direito:", 
                OPCOES_PAINEL, 
                index=idx_pdir,
                key="w_painel_dir",
                on_change=cb_atualizar_pdir
            )

        col_left, col_right = st.columns(PROPORCOES_MAP[st.session_state["cfg_cockpit_proporcao"]])
        with col_left:
            st.markdown(f"<div style='margin-bottom:4px;'><b style='color:#38BDF8;'>{st.session_state['cfg_painel_esquerdo']}</b></div>", unsafe_allow_html=True)
            render_visao(st.session_state["cfg_painel_esquerdo"], altura_pixel=680)

        with col_right:
            st.markdown(f"<div style='margin-bottom:4px;'><b style='color:#38BDF8;'>{st.session_state['cfg_painel_direito']}</b></div>", unsafe_allow_html=True)
            render_visao(st.session_state["cfg_painel_direito"], altura_pixel=680)