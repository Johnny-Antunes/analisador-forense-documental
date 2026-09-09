import streamlit as st
import pandas as pd
import json
from pathlib import Path
import streamlit.components.v1 as components
import pydeck as pdk

from database import (
    CAMINHO_REDE_OFICIAL, PASTA_LOCAL,
    carregar_arquivos_para_sqlite, consultar_detalhes_caso
)
from graph_engine import carregar_redes, processar_subgrafo_caso, gerar_html_grafo

# =====================================================
# 1. CONFIGURAÇÃO DE PÁGINA & CSS FORENSE
# =====================================================
st.set_page_config(
    page_title="Mesa de Inteligência de Fraudes - LCFO",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .stDeployButton, [data-testid="stToolbarActions"], [data-testid="stToolbar"], #MainMenu, footer {
        display: none !important;
        visibility: hidden !important;
    }
    [data-testid="collapsedControl"], [data-testid="stSidebarCollapseButton"] {
        display: flex !important;
        visibility: visible !important;
        pointer-events: auto !important;
        z-index: 999999 !important;
    }
    .block-container {
        padding-top: 2.2rem !important;
        padding-bottom: 0.5rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        max-width: 100% !important;
    }
    .hud-bar {
        background: #0B111E;
        border: 1px solid #1E293B;
        border-radius: 8px;
        padding: 8px 14px;
        margin-top: 4px;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
</style>
""", unsafe_allow_html=True)

# =====================================================
# 2. GESTÃO DA BASE (SEMPRE VISÍVEL NA BARRA LATERAL)
# =====================================================
with st.sidebar:
    st.markdown("### ⚙️ Gestão da Base Operacional")
    pasta_sugerida = CAMINHO_REDE_OFICIAL if Path(CAMINHO_REDE_OFICIAL).exists() else str(PASTA_LOCAL)
    caminho_input = st.text_input("📁 Caminho da Pasta (.xlsx):", value=pasta_sugerida, help="Pode ser o drive T: ou uma pasta local.")
    
    qtd_detectada = 0
    p_check = Path(caminho_input.strip())
    if p_check.exists():
        qtd_detectada = len(list(p_check.glob("*.xlsx")))
        st.caption(f"🟢 **{qtd_detectada}** planilhas detectadas na pasta.")
    else:
        st.caption("🔴 Caminho não acessível ou inexistente.")

    if st.button("🔄 Ingerir Planilhas no SQLite", use_container_width=True):
        with st.spinner("Atualizando registros..."):
            qtd_arq, qtd_reg, msg_erro = carregar_arquivos_para_sqlite(caminho_input)
            st.cache_data.clear()
            if qtd_arq > 0:
                st.success(f"✅ {qtd_arq} planilhas ingeridas ({qtd_reg:,} assistências consolidadas).")
                st.rerun()
            else:
                st.warning(msg_erro or "Nenhuma planilha encontrada para carregar.")

# =====================================================
# 3. CARREGAMENTO DAS REDES E VALIDAÇÃO DA BASE
# =====================================================
G, cluster_info = carregar_redes()

if not G or not cluster_info:
    st.info("👈 Nenhum dado processado no banco. Verifique o caminho da pasta na barra lateral e clique em **'🔄 Ingerir Planilhas no SQLite'**.")
    st.stop()

# =====================================================
# 4. CONTROLES DO CASO NA BARRA LATERAL
# =====================================================
with st.sidebar:
    st.markdown("---")
    st.markdown("### 📂 Seleção da Célula")
    
    termo_busca = st.text_input("Buscar por CPF, Placa ou Telefone:", placeholder="Digite para filtrar...").strip()
    
    cluster_selecionado = None
    node_alvo_destaque = ""
    
    if termo_busca:
        termo_clean = re.sub(r'[^a-zA-Z0-9]', '', termo_busca).upper()
        alvos = [n for n in G.nodes if termo_clean in n]
        if alvos:
            alvo = alvos[0]
            node_alvo_destaque = alvo
            for c in cluster_info:
                if alvo in c["nodes"]:
                    cluster_selecionado = c
                    st.success(f"🎯 Localizado na Célula #{c['id']}")
                    break
        else:
            st.warning("Nenhum vínculo com essa chave.")

    opcoes = {f"Célula #{c['id']} ({c['tamanho']} nós) - {c['hub_label']}": c['id'] for c in cluster_info[:120]}
    
    if not cluster_selecionado:
        sel_box = st.selectbox("Selecione o Caso Ativo:", list(opcoes.keys()))
        cluster_selecionado = next(c for c in cluster_info if c["id"] == opcoes[sel_box])

    nos_celula = cluster_selecionado["nodes"]
    qtd_tels_caso = len([n for n in nos_celula if n.startswith("TEL_")])
    qtd_cpfs_caso = len([n for n in nos_celula if n.startswith("CPF_")])
    qtd_placas_caso = len([n for n in nos_celula if n.startswith("PLACA_")])

    st.markdown("---")
    st.markdown("### 🎯 Filtros de Entidades")
    mostrar_tel = st.checkbox(f"🚨 Telefones ({qtd_tels_caso})", value=True)
    mostrar_cpf = st.checkbox(f"👤 Titulares ({qtd_cpfs_caso})", value=True)
    mostrar_placa = st.checkbox(f"🚗 Placas ({qtd_placas_caso})", value=True)

    st.markdown("---")
    st.markdown("### 🔤 Ajuste Visual")
    font_slider = st.slider("Tamanho da Fonte (px):", min_value=10, max_value=20, value=12)

# =====================================================
# 5. CONSULTA DE DETALHES DO CASO NO SQLITE
# =====================================================
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

# =====================================================
# 6. CABEÇALHO COMPACTO & HUD DO CASO ATIVO
# =====================================================
col_header1, col_header2 = st.columns([3.0, 2.0])
with col_header1:
    st.markdown("<h4 style='margin:0; color:#38BDF8;'>🛡️ Mesa de Inteligência Forense - LCFO</h4>", unsafe_allow_html=True)
with col_header2:
    st.caption(f"Base Consolidada: {len(cluster_info):,} Células | {len(G.nodes):,} Entidades | {len(G.edges):,} Conexões")

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

# =====================================================
# 7. SUBGRAFO FORENSE & CÁLCULOS (CHAMA GRAPH_ENGINE)
# =====================================================
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

# =====================================================
# 8. ABAS DE VISUALIZAÇÃO
# =====================================================
tab_grafo, tab_tempo, tab_mapa, tab_dados = st.tabs([
    "🕸️ Grafo de Vínculos Forense",
    "🕒 Evolução Temporal",
    "📍 Radar Territorial",
    "📋 Tabela de Ocorrências"
])

html_codigo = gerar_html_grafo(
    json.dumps(vis_nodes),
    json.dumps(vis_edges),
    font_slider,
    cluster_selecionado["hub_id"],
    node_alvo_destaque
)

with tab_grafo:
    c_info, c_btn = st.columns([3.5, 1.5])
    with c_info:
        st.caption(f"Cluster Forense com {len(vis_nodes)} entidades ativas e {len(vis_edges)} conexões ponderadas.")
    with c_btn:
        st.download_button(
            label="💾 Baixar Dossiê HTML deste Caso",
            data=html_codigo,
            file_name=f"Dossie_Forense_Caso_{cluster_selecionado['id']}.html",
            mime="text/html",
            use_container_width=True,
            help="Gera arquivo interativo autônomo para envio por e-mail ou relatório jurídico."
        )
    
    components.html(html_codigo, height=780)

with tab_tempo:
    if not df_detalhes.empty and "data" in df_detalhes.columns:
        st.markdown(f"**Dispersão Cronológica das Assistências da Célula #{cluster_selecionado['id']}** ({len(df_detalhes)} registros)")
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
            st.info("Datas em formato não reconhecível para plotagem gráfica.")
    else:
        st.info("Sem dados de data disponíveis para esta seleção.")

with tab_mapa:
    if not df_detalhes.empty and "latitude" in df_detalhes.columns:
        st.markdown(f"**Concentração Geográfica da Célula #{cluster_selecionado['id']}** ({len(df_detalhes)} assistências)")
        df_mapa = df_detalhes.dropna(subset=["latitude", "longitude"]).copy()
        if not df_mapa.empty:
            lat_centro = float(df_mapa["latitude"].mean())
            lon_centro = float(df_mapa["longitude"].mean())
            
            camada_pontos = pdk.Layer(
                "ScatterplotLayer",
                data=df_mapa,
                get_position="[longitude, latitude]",
                get_color="[56, 189, 248, 190]",
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
                pitch=20
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
                map_style=None
            )

            st.pydeck_chart(deck, use_container_width=True)

        resumo_cidades = df_detalhes.groupby(["cidade", "uf"]).size().reset_index(name="Total_Assistencias").sort_values(by="Total_Assistencias", ascending=False)
        st.dataframe(resumo_cidades, use_container_width=True, hide_index=True)
    else:
        st.info("Sem dados geográficos para esta seleção.")

with tab_dados:
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