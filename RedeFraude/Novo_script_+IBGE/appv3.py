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
    consultar_detalhes_caso, cruzar_com_criacoes_diarias, obter_radar_expansoes,
    salvar_dados_caso, carregar_dados_caso, carregar_todos_casos_cadastrados,
    cadastrar_entidade_suspeita, listar_entidades_suspeitas, remover_entidade_suspeita,
    consultar_todas_ocorrencias_entidade, semear_base_mestra_da_blacklist,
    investigar_alvo_em_criacoes_diarias, promover_descoberta_para_caso,
    vincular_caso_e_entidades, anexar_descoberta_a_caso_existente,
    carregar_historico_pareceres
)
from graph_engine import (
    carregar_redes, processar_subgrafo_caso, gerar_html_grafo, processar_grafo_dataframe
)

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

    [data-testid="stSidebar"] {
        overflow-y: auto !important;
    }
    [data-testid="stSidebarContent"] {
        overflow-y: auto !important;
        max-height: 100vh !important;
        padding-bottom: 2rem !important;
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
# 2. DEFINIÇÕES GLOBAIS E ESTADO PERSISTENTE
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

if "modo_descoberta_ativo" not in st.session_state:
    st.session_state["modo_descoberta_ativo"] = False
    st.session_state["dados_descoberta"] = None

if "aba_principal_triagem" not in st.session_state:
    st.session_state["aba_principal_triagem"] = "🛡️ Células & Redes da Blacklist"

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
# 3. GESTÃO DE BASES (SIDEBAR COM EXPANDERS FECHADOS)
# =====================================================
with st.sidebar:
    st.markdown("### ⚙️ Gestão de Bases de Dados")
    
    with st.expander("📁 1. Base Blacklist (Duplicidades)", expanded=False):
        pasta_padrao_bl = CAMINHO_REDE_OFICIAL if Path(CAMINHO_REDE_OFICIAL).exists() else str(PASTA_LOCAL_BLACKLIST)
        caminho_input = st.text_input("Pasta Blacklist (.xlsx/.csv):", value=pasta_padrao_bl)
        
        qtd_detectada = 0
        p_check = Path(caminho_input.strip())
        if p_check.exists():
            qtd_detectada = len(list(p_check.glob("*.xlsx"))) + len(list(p_check.glob("*.csv")))
            st.caption(f"🟢 **{qtd_detectada}** planilhas detectadas.")
        else:
            st.caption("🔴 Caminho não acessível.")

        if st.button("🔄 Ingerir Blacklist (Novos)", use_container_width=True):
            with st.spinner("Ingerindo arquivos novos da Blacklist..."):
                qtd_arq, qtd_reg, msg_info, erros_bl = carregar_arquivos_para_sqlite(caminho_input, forcar_releitura=False)
                st.cache_data.clear()
                if qtd_reg > 0:
                    st.success(f"✅ {qtd_reg:,} novas assistências inseridas e Base Mestra semeada!")
                else:
                    st.info(msg_info or "Base já atualizada.")
                if erros_bl:
                    with st.expander(f"⚠️ {len(erros_bl)} arquivos com aviso"):
                        for e in erros_bl: st.caption(f"• {e}")
                st.rerun()

    with st.expander("🌐 2. Criações Diárias (Base Geral)", expanded=False):
        pasta_padrao_cr = CAMINHO_REDE_CRIACAO if Path(CAMINHO_REDE_CRIACAO).exists() else str(PASTA_LOCAL_CRIACAO)
        caminho_criacao = st.text_input("Pasta Criação (.xlsx/.csv):", value=pasta_padrao_cr)
        
        qtd_criacao = 0
        p_criacao = Path(caminho_criacao.strip())
        if p_criacao.exists():
            qtd_criacao = len(list(p_criacao.glob("*.xlsx"))) + len(list(p_criacao.glob("*.csv")))
            st.caption(f"🟢 **{qtd_criacao}** arquivos detectados.")
        else:
            st.caption("🔴 Caminho não acessível.")

        if st.button("🔄 Ingerir Criações no SQLite (Novos)", use_container_width=True):
            with st.spinner("Ingerindo arquivos novos de Criações..."):
                qtd_arq_c, qtd_reg_c, msg_c, erros_c = carregar_criacoes_diarias_para_sqlite(caminho_criacao, forcar_releitura=False)
                if qtd_reg_c > 0:
                    st.success(f"✅ {qtd_reg_c:,} novas assistências ingeridas!")
                else:
                    st.info(msg_c or "Nenhum arquivo novo para ingerir.")
                if erros_c:
                    with st.expander(f"⚠️ {len(erros_c)} arquivos com aviso"):
                        for e in erros_c: st.caption(f"• {e}")
                st.rerun()

# =====================================================
# 4. CARREGAMENTO DAS REDES E METADADOS
# =====================================================
G, cluster_info = carregar_redes()

if not G or not cluster_info:
    st.info("👈 Nenhum dado processado no banco. Verifique o caminho da pasta na barra lateral e clique em **'🔄 Ingerir Blacklist'**.")
    st.stop()

radar_alertas = obter_radar_expansoes(cluster_info)
casos_cadastrados = carregar_todos_casos_cadastrados()

opcoes_celulas = {}
for c in cluster_info[:150]:
    info_caso = casos_cadastrados.get(c["id"], {})
    nome_custom = info_caso.get("nome", "")
    tag_nome = f" - 📁 {nome_custom.upper()}" if nome_custom else ""
    alerta_badge = f" 🚨 (+{radar_alertas[c['id']]} novas)" if c["id"] in radar_alertas else ""
    rotulo = f"[{c['id']}]{tag_nome} ({c['tamanho']} nós) - {c['hub_label']}{alerta_badge}"
    opcoes_celulas[rotulo] = c['id']

mapa_id_para_label = {v: k for k, v in opcoes_celulas.items()}

# =====================================================
# 5. CONTROLES NA BARRA LATERAL & BUSCA INTELIGENTE
# =====================================================
with st.sidebar:
    st.markdown("---")
    
    if st.session_state["celula_ativa_id"] is not None or st.session_state["modo_descoberta_ativo"]:
        if st.button("⬅️ Voltar à Triagem Geral", use_container_width=True):
            st.session_state["celula_ativa_id"] = None
            st.session_state["modo_descoberta_ativo"] = False
            st.session_state["dados_descoberta"] = None
            st.rerun()
        
        if st.session_state["celula_ativa_id"] is not None:
            st.markdown("### 📂 Caso em Análise")
            label_atual = mapa_id_para_label.get(st.session_state["celula_ativa_id"], list(opcoes_celulas.keys())[0])
            idx_caso = list(opcoes_celulas.keys()).index(label_atual) if label_atual in opcoes_celulas else 0
            novo_caso_selecionado = st.selectbox("Alternar Caso:", list(opcoes_celulas.keys()), index=idx_caso)
            
            if opcoes_celulas[novo_caso_selecionado] != st.session_state["celula_ativa_id"]:
                st.session_state["celula_ativa_id"] = opcoes_celulas[novo_caso_selecionado]
                st.rerun()

            cluster_selecionado = next((c for c in cluster_info if c["id"] == st.session_state["celula_ativa_id"]), cluster_info[0])

            st.markdown("---")
            st.markdown("### 🎯 Filtros de Entidades")
            mostrar_tel = st.checkbox(f"🚨 Telefones ({cluster_selecionado['qtd_tels']})", value=True)
            mostrar_cpf = st.checkbox(f"👤 Titulares ({cluster_selecionado['qtd_cpfs']})", value=True)
            mostrar_placa = st.checkbox(f"🚗 Placas ({cluster_selecionado['qtd_placas']})", value=True)

        st.markdown("---")
        st.markdown("### 📐 Layout do Grafo")
        idx_layout = list(LAYOUT_MAP.keys()).index(st.session_state["cfg_layout_ativo"]) if st.session_state["cfg_layout_ativo"] in LAYOUT_MAP else 0
        st.selectbox("Estrutura Topológica:", options=list(LAYOUT_MAP.keys()), format_func=lambda k: LAYOUT_MAP[k], index=idx_layout, key="w_layout_ativo", on_change=cb_atualizar_layout)

        st.markdown("---")
        st.markdown("### 🔤 Ajuste Visual")
        font_slider = st.slider("Fonte Base do Grafo (px):", min_value=10, max_value=18, value=12)
        espacamento_slider = st.slider("Dispersão / Distância (px):", min_value=180, max_value=450, value=280, step=20)

    else:
        st.markdown("### 🔍 Busca Unificada de Inteligência")
        st.caption("Pesquise na Blacklist ou investigue alvos inéditos nas Criações Diárias.")
        termo_busca_global = st.text_input("CPF, Placa ou Telefone:", placeholder="Digite para rastrear...").strip()
        
        if termo_busca_global:
            termo_clean = re.sub(r'[^a-zA-Z0-9]', '', termo_busca_global).upper()
            alvos_bl = [n for n in G.nodes if termo_clean in n]
            
            # 1. Alvo já existe na Blacklist
            if alvos_bl:
                alvo = alvos_bl[0]
                for c in cluster_info:
                    if alvo in c["nodes"]:
                        nome_caso_bl = casos_cadastrados.get(c["id"], {}).get("nome", c["hub_label"])
                        st.success(f"🎯 Localizado no Caso Oficial:\n**{nome_caso_bl}** ({c['id']})")
                        if st.button(f"Abrir Caso {c['id']}", use_container_width=True, type="primary"):
                            st.session_state["celula_ativa_id"] = c["id"]
                            st.rerun()
                        break
            else:
                # 2. Alvo inédito: Busca nas Criações Diárias
                resumo_disc, df_disc, ents_disc = investigar_alvo_em_criacoes_diarias(termo_busca_global)
                if resumo_disc and resumo_disc["total_assistencias"] > 0:
                    st.warning(f"🚨 **ALVO INÉDITO ENCONTRADO!**\n\n• {resumo_disc['total_assistencias']} sinistros em Criações Diárias\n• Rede: {resumo_disc['qtd_cpfs']} CPFs, {resumo_disc['qtd_tels']} telefones, {resumo_disc['qtd_placas']} placas.")
                    if st.button("🔬 Abrir Mesa de Investigação Ativa", use_container_width=True, type="primary"):
                        st.session_state["modo_descoberta_ativo"] = True
                        st.session_state["dados_descoberta"] = {
                            "resumo": resumo_disc,
                            "df": df_disc,
                            "entidades": ents_disc
                        }
                        st.rerun()
                else:
                    st.info("⚪ Nenhum acionamento localizado para este dado em nenhuma das bases.")

# ====================================================================
# 6. TELA 3: MESA DE INVESTIGAÇÃO ATIVA (DESCOBERTA DE CASOS INÉDITOS)
# ====================================================================
if st.session_state["modo_descoberta_ativo"] and st.session_state["dados_descoberta"]:
    dados_d = st.session_state["dados_descoberta"]
    resumo_d = dados_d["resumo"]
    df_d = dados_d["df"]

    col_nav1, col_nav2 = st.columns([4.0, 1.0])
    with col_nav1:
        st.markdown(f"<div style='padding-top:4px;'><span style='color:#94A3B8; font-size:12px;'>Mesa de Triagem / </span> <b style='color:#FBBF24; font-size:14px;'>Investigação Ativa de Descoberta: {resumo_d['alvo_buscado']}</b></div>", unsafe_allow_html=True)
    with col_nav2:
        if st.button("⬅️ Fechar Investigação", use_container_width=True):
            st.session_state["modo_descoberta_ativo"] = False
            st.session_state["dados_descoberta"] = None
            st.rerun()

    st.markdown(f"""
    <div class="hud-bar" style="border-color: #FBBF24;">
        <div style="display:flex; align-items:center; gap:20px;">
            <span style="font-size:0.95rem; font-weight:700; color:#FBBF24;">
                🚨 ALVO INÉDITO DETECTADO
            </span>
            <span style="font-size:0.82rem; color:#94A3B8;">
                Acionamentos Diretos: <b style="color:#F8FAFC;">{resumo_d['assistencias_diretas']}</b>
            </span>
            <span style="font-size:0.82rem; color:#94A3B8;">
                Rede Total Descoberta: <b style="color:#38BDF8;">{resumo_d['total_assistencias']} sinistros</b>
            </span>
            <span style="font-size:0.82rem; color:#94A3B8;">
                Entidades: <b style="color:#A78BFA;">{resumo_d['qtd_cpfs']} CPFs | {resumo_d['qtd_tels']} Tels | {resumo_d['qtd_placas']} Placas</b>
            </span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    with st.expander("📌 Ações de Promoção e Vínculo", expanded=False):
        tipo_acao = st.radio("Escolha o destino desta rede descoberta:", ["🆕 Criar Novo Caso Oficial", "🔗 Anexar a uma Quadrilha Existente"], horizontal=True)

        if tipo_acao == "🆕 Criar Novo Caso Oficial":
            c_prom1, c_prom2, c_prom3 = st.columns([2.5, 1.5, 1.5])
            with c_prom1:
                nome_prom = st.text_input("Nome da Nova Quadrilha / Operação:", placeholder="Ex: Esquema Guinchos Santo André")
            with c_prom2:
                st_prom = st.selectbox("Status Operacional:", ["Confirmado Fraude", "Em Investigação", "Monitoramento Contínuo"])
            with c_prom3:
                analista_prom = st.text_input("Analista:", placeholder="Seu nome")
            
            parecer_prom = st.text_area("Parecer da Descoberta:", placeholder="Descreva os vínculos observados...")

            if st.button("💾 Promover a Novo Caso Oficial", type="primary", use_container_width=True):
                if nome_prom.strip():
                    nos_promovidos = promover_descoberta_para_caso(df_d)
                    st.cache_data.clear()
                    _, cluster_info_novo = carregar_redes()
                    
                    cluster_promovido = next((c for c in cluster_info_novo if any(n in c["nodes"] for n in nos_promovidos)), None)
                    
                    if cluster_promovido:
                        id_real = cluster_promovido["id"]
                        vincular_caso_e_entidades(id_real, nos_promovidos, nome_prom, st_prom, analista_prom, parecer_prom)
                        st.session_state["modo_descoberta_ativo"] = False
                        st.session_state["dados_descoberta"] = None
                        st.session_state["celula_ativa_id"] = id_real
                        st.success(f"✅ Rede promovida com sucesso! Caso vinculado: {id_real}")
                        st.rerun()
                    else:
                        st.error("Erro ao localizar o componente recalculado. Tente reabrir pela triagem.")
                else:
                    st.error("Informe um nome para a nova quadrilha antes de promover.")

        else:
            c_anex1, c_anex2 = st.columns([3.0, 2.0])
            with c_anex1:
                caso_destino_label = st.selectbox("Selecione a Quadrilha de Destino:", list(opcoes_celulas.keys()))
                id_destino = opcoes_celulas[caso_destino_label]
            with c_anex2:
                analista_anex = st.text_input("Analista Responsável:", placeholder="Seu nome", key="anex_analista")
            
            obs_anexo = st.text_input("Observação / Motivo do Vínculo:", placeholder="Ex: Telefone identificado acionando para os mesmos alvos.")

            if st.button(f"🔗 Anexar Ocorrências ao Caso {id_destino}", type="primary", use_container_width=True):
                anexar_descoberta_a_caso_existente(df_d, id_destino, analista_anex, obs_anexo)
                st.cache_data.clear()
                st.session_state["modo_descoberta_ativo"] = False
                st.session_state["dados_descoberta"] = None
                st.session_state["celula_ativa_id"] = id_destino
                st.success(f"✅ Dados anexados com sucesso ao Caso {id_destino}!")
                st.rerun()

    vis_nodes_d, vis_edges_d, hub_id_d = processar_grafo_dataframe(df_d, alvo_principal=resumo_d["termo_limpo"], font_slider=12)
    html_codigo_d = gerar_html_grafo(
        vis_nodes_json=json.dumps(vis_nodes_d),
        vis_edges_json=json.dumps(vis_edges_d),
        base_font_size=12,
        hub_id=hub_id_d,
        target_node_id="",
        layout_ativo=st.session_state["cfg_layout_ativo"],
        espacamento=280
    )

    t_g, t_m, t_t, t_d = st.tabs(["🕸️ Grafo da Rede Descoberta", "📍 Radar Territorial", "🕒 Evolução Temporal", "📋 Ocorrências Encontradas"])
    with t_g:
        components.html(html_codigo_d, height=720)
    with t_m:
        df_mapa_d = df_d.dropna(subset=["latitude", "longitude"]).copy()
        if not df_mapa_d.empty:
            camada_pontos = pdk.Layer(
                "ScatterplotLayer", data=df_mapa_d, get_position="[longitude, latitude]",
                get_color="[251, 191, 36, 180]", get_line_color="[255, 255, 255, 220]",
                line_width_min_pixels=1, stroked=True, get_radius=3200, radius_min_pixels=6, radius_max_pixels=20, pickable=True
            )
            view_state = pdk.ViewState(latitude=float(df_mapa_d["latitude"].mean()), longitude=float(df_mapa_d["longitude"].mean()), zoom=7, pitch=0)
            deck = pdk.Deck(layers=[camada_pontos], initial_view_state=view_state, map_style="dark")
            st.pydeck_chart(deck, use_container_width=True)
            st.dataframe(df_mapa_d.groupby(["cidade", "uf"]).size().reset_index(name="Total").sort_values(by="Total", ascending=False), use_container_width=True, hide_index=True)
        else:
            st.info("Sem coordenadas para plotagem territorial.")
    with t_t:
        df_temp_d = df_d.copy()
        df_temp_d["Periodo"] = pd.to_datetime(df_temp_d["data"], errors="coerce").dt.strftime("%Y-%m")
        df_temp_d = df_temp_d.dropna(subset=["Periodo"]).groupby("Periodo").size().reset_index(name="Assistências").sort_values(by="Periodo")
        st.bar_chart(data=df_temp_d.set_index("Periodo"), color="#FBBF24", use_container_width=True)
    with t_d:
        st.dataframe(df_d[["data", "id_assistencia", "servico", "titular", "cpf", "telefone", "placa", "bairro", "cidade", "uf"]], use_container_width=True, hide_index=True)

# ====================================================================
# 7. TELA 1: PAINEL GERAL (CÉLULAS vs. BASE MESTRA)
# ====================================================================
elif st.session_state["celula_ativa_id"] is None:
    st.markdown("<h3 style='margin:0; color:#38BDF8;'>🛡️ Mesa de Inteligência Forense - LCFO</h3>", unsafe_allow_html=True)
    st.caption("Painel Executivo de Triagem, Gestão de Quadrilhas e Base Mestra de Suspeitos")
    
    st.markdown("<br/>", unsafe_allow_html=True)
    
    aba_triagem = st.radio(
        "Navegação da Mesa:",
        ["🛡️ Células & Redes da Blacklist", "🎯 Base Mestra de Entidades Suspeitas"],
        horizontal=True,
        key="aba_principal_triagem"
    )

    st.markdown("<br/>", unsafe_allow_html=True)

    if aba_triagem == "🛡️ Células & Redes da Blacklist":
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(f"""<div class="kpi-card"><span style="color:#94A3B8; font-size:12px; font-weight:600;">CÉLULAS MAPEADAS</span><br/><b style="font-size:26px; color:#F8FAFC;">{len(cluster_info):,}</b></div>""", unsafe_allow_html=True)
        with k2:
            st.markdown(f"""<div class="kpi-card"><span style="color:#94A3B8; font-size:12px; font-weight:600;">ENTIDADES NA BLACKLIST</span><br/><b style="font-size:26px; color:#38BDF8;">{len(G.nodes):,}</b></div>""", unsafe_allow_html=True)
        with k3:
            st.markdown(f"""<div class="kpi-card"><span style="color:#94A3B8; font-size:12px; font-weight:600;">QUADRILHAS NOMEADAS</span><br/><b style="font-size:26px; color:#10B981;">{len([c for c in casos_cadastrados.values() if c['nome']]):,}</b></div>""", unsafe_allow_html=True)
        with k4:
            qtd_com_alerta = len(radar_alertas)
            cor_alerta = "#EF4444" if qtd_com_alerta > 0 else "#10B981"
            st.markdown(f"""<div class="kpi-card"><span style="color:#94A3B8; font-size:12px; font-weight:600;">CÉLULAS REINCIDENTES</span><br/><b style="font-size:26px; color:{cor_alerta};">{qtd_com_alerta} no Radar</b></div>""", unsafe_allow_html=True)

        st.markdown("<br/>", unsafe_allow_html=True)

        if radar_alertas:
            st.markdown(f"##### 🚨 Radar de Alertas: {len(radar_alertas)} Células Reincidentes nas Criações Diárias")
            cols_radar = st.columns(min(4, len(radar_alertas)))
            for i, (cid, total_hits) in enumerate(list(radar_alertas.items())[:4]):
                with cols_radar[i]:
                    c_obj = next((c for c in cluster_info if c["id"] == cid), None)
                    if c_obj:
                        nome_q = casos_cadastrados.get(cid, {}).get("nome", c_obj['hub_label'][:20])
                        if st.button(f"🚨 {nome_q}\n+{total_hits} ocorrências ({c_obj['id']})", use_container_width=True):
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
            info_c = casos_cadastrados.get(c["id"], {})
            status_radar = f"🚨 +{radar_alertas[c['id']]} novas" if c["id"] in radar_alertas else "Estável"
            tabela_casos.append({
                "Código": c["id"],
                "Nome da Quadrilha": info_c.get("nome", "Não Batizado"),
                "Status": info_c.get("status", "Em Investigação"),
                "Âncora Central (Hub)": c["hub_label"],
                "Tamanho": c["tamanho"],
                "Telefones": c["qtd_tels"],
                "Titulares": c["qtd_cpfs"],
                "Placas": c["qtd_placas"],
                "Status Radar": status_radar
            })
        df_resumo_casos = pd.DataFrame(tabela_casos)
        st.dataframe(df_resumo_casos, use_container_width=True, hide_index=True)

    else:
        c_m_top1, c_m_top2 = st.columns([3.5, 1.5])
        with c_m_top1:
            st.markdown("#### 🎯 Base Mestra de Entidades Monitoradas")
            st.caption("Repositório permanente de suspeitos. Varre retrospectivamente todo o universo de Criações Diárias.")
        with c_m_top2:
            if st.button("🔄 Sincronizar com a Blacklist", use_container_width=True):
                with st.spinner("Importando entidades únicas da Blacklist..."):
                    novos_sem = semear_base_mestra_da_blacklist()
                    st.success(f"✅ {novos_sem} novas entidades catalogadas na Base Mestra!")
                    st.rerun()

        with st.expander("➕ Cadastrar Nova Entidade Suspeita", expanded=False):
            with st.form("form_novo_suspeito", clear_on_submit=True):
                c_t1, c_t2, c_t3 = st.columns([1.5, 2.5, 2.5])
                with c_t1:
                    novo_tipo = st.selectbox("Tipo de Entidade:", ["TELEFONE", "PLACA", "CPF"])
                with c_t2:
                    novo_val = st.text_input("Dado / Número / Placa:", placeholder="Ex: (11) 97000-1122 ou ABC-1234")
                with c_t3:
                    novo_nome = st.text_input("Nome / Apelido / Titular:", placeholder="Ex: Marcos V. (Laranja)")

                c_m1, c_m2, c_m3 = st.columns([2.0, 2.0, 2.0])
                with c_m1:
                    novo_caso = st.text_input("Quadrilha / Operação Associada:", placeholder="Ex: Quadrilha Sapeaçu")
                with c_m2:
                    novo_st = st.selectbox("Status:", ["Ativo", "Confirmado Fraude", "Em Monitoramento", "Inativo"])
                with c_m3:
                    novo_analista = st.text_input("Analista Responsável:", placeholder="Seu nome")

                novo_motivo = st.text_area("Motivo da Inclusão / Modus Operandi:", placeholder="Ex: Solicitou guinchos sequenciais para o mesmo destino...")

                btn_salvar_suspeito = st.form_submit_button("💾 Salvar na Base Mestra & Varrer Histórico", type="primary", use_container_width=True)
                if btn_salvar_suspeito:
                    ok, msg, total_hits = cadastrar_entidade_suspeita(novo_tipo, novo_val, novo_nome, novo_motivo, novo_caso, status=novo_st, analista=novo_analista)
                    if ok:
                        st.success(f"{msg} 🔍 **Varredura retrospectiva:** {total_hits} ocorrências identificadas!")
                        st.rerun()
                    else:
                        st.error(msg)

        st.markdown("---")
        f_c1, f_c2, f_c3 = st.columns([1.5, 1.5, 3.0])
        with f_c1:
            f_tipo = st.selectbox("Filtrar Tipo:", ["TODOS", "TELEFONE", "PLACA", "CPF"])
        with f_c2:
            f_status = st.selectbox("Filtrar Status:", ["TODOS", "Ativo", "Confirmado Fraude", "Em Monitoramento", "Inativo"])
        with f_c3:
            f_busca = st.text_input("Buscar na Base Mestra:", placeholder="Digite dado, nome ou quadrilha...")

        suspeitos = listar_entidades_suspeitas(f_tipo, f_status, f_busca)

        if suspeitos:
            st.caption(f"Exibindo **{len(suspeitos)}** entidades suspeitas catalogadas.")
            tabela_mestra = []
            for s in suspeitos:
                tabela_mestra.append({
                    "ID": s["id"],
                    "Tipo": s["tipo"],
                    "Dado Monitorado": s["valor_formatado"],
                    "Titular / Apelido": s["nome_referencia"] or "-",
                    "Quadrilha / Caso": s["quadrilha_exibicao"],
                    "Status": s["status"],
                    "Histórico Total": f"🚨 {s['total_historico']} sinistros" if s['total_historico'] > 0 else "Nenhum",
                    "Data Cadastro": s["data_cadastro"],
                    "Analista": s["analista"] or "-"
                })
            df_mestra = pd.DataFrame(tabela_mestra)
            st.dataframe(df_mestra, use_container_width=True, hide_index=True)

            with st.expander("🔎 Ver Histórico Detalhado de uma Entidade Específica"):
                opcoes_ent = {f"[{s['tipo']}] {s['valor_formatado']} - {s['nome_referencia']}": s for s in suspeitos}
                sel_ent_label = st.selectbox("Selecione para auditar todas as assistências passadas:", list(opcoes_ent.keys()))
                ent_alvo = opcoes_ent[sel_ent_label]

                col_btn1, col_btn2 = st.columns([4.0, 1.0])
                with col_btn2:
                    if st.button(f"🗑️ Excluir #{ent_alvo['id']}", use_container_width=True):
                        remover_entidade_suspeita(ent_alvo["id"])
                        st.success("Entidade removida da Base Mestra.")
                        st.rerun()

                df_ocorr = consultar_todas_ocorrencias_entidade(ent_alvo["tipo"], ent_alvo["valor"])
                if not df_ocorr.empty:
                    st.markdown(f"**{len(df_ocorr)} Ocorrências Encontradas (Blacklist + Criações Diárias):**")
                    st.dataframe(df_ocorr, use_container_width=True, hide_index=True)
                    csv_ent = df_ocorr.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label=f"📥 Baixar Ocorrências de {ent_alvo['valor_formatado']} (.csv)",
                        data=csv_ent,
                        file_name=f"Historico_{ent_alvo['valor']}.csv",
                        mime="text/csv"
                    )
                else:
                    st.info("Nenhuma ocorrência registrada nos arquivos ingeridos para esta entidade.")
        else:
            st.info("Nenhuma entidade cadastrada. Clique no botão acima para sincronizar com a Blacklist.")

# =====================================================
# 8. TELA 2: MESA DE INVESTIGAÇÃO DE CASO OFICIAL
# =====================================================
else:
    cluster_selecionado = next((c for c in cluster_info if c["id"] == st.session_state["celula_ativa_id"]), cluster_info[0])
    dados_caso = carregar_dados_caso(cluster_selecionado["id"])
    nome_quadrilha_display = dados_caso["nome_personalizado"] or f"Caso {cluster_selecionado['id']}"

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

    col_nav1, col_nav2, col_nav3 = st.columns([3.2, 1.8, 1.0])
    with col_nav1:
        st.markdown(f"<div style='padding-top:4px;'><span style='color:#94A3B8; font-size:12px;'>Mesa de Triagem / </span> <b style='color:#38BDF8; font-size:14px;'>{nome_quadrilha_display}</b> <span style='font-size:11px; color:#A78BFA;'>({cluster_selecionado['id']})</span></div>", unsafe_allow_html=True)
    with col_nav2:
        idx_modo = 0 if st.session_state["cfg_modo_exibicao"] == "📑 Abas Clássicas" else 1
        st.radio("Modo de Visualização:", ["📑 Abas Clássicas", "🖥️ Cockpit Dividido"], index=idx_modo, horizontal=True, key="w_modo_exibicao", on_change=cb_atualizar_modo, label_visibility="collapsed")
    with col_nav3:
        if st.button("⬅️ Sair da Célula", use_container_width=True):
            st.session_state["celula_ativa_id"] = None
            st.rerun()

    st.markdown(f"""
    <div class="hud-bar">
        <div style="display:flex; align-items:center; gap:20px;">
            <span style="font-size:0.95rem; font-weight:700; color:#F8FAFC;">
                🚨 {nome_quadrilha_display.upper()}
            </span>
            <span style="font-size:0.82rem; color:#94A3B8;">
                Status: <b style="color:#10B981;">{dados_caso['status']}</b>
            </span>
            <span style="font-size:0.82rem; color:#94A3B8;">
                Âncora: <b style="color:#FBBF24;">{cluster_selecionado['hub_label']}</b>
            </span>
            <span style="font-size:0.82rem; color:#94A3B8;">
                Ocorrências: <b style="color:#38BDF8;">{total_assistencias} sinistros</b>
            </span>
            <span style="font-size:0.82rem; color:#94A3B8;">
                Período: <b style="color:#E2E8F0;">{periodo_str}</b>
            </span>
        </div>
    </div>
    """, unsafe_allow_html=True)

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

    def render_visao(nome_visao, altura_pixel=720):
        if "Grafo" in nome_visao:
            components.html(html_codigo, height=altura_pixel)
            c_ginfo, c_gbtn = st.columns([3.5, 1.5])
            with c_ginfo:
                st.caption(f"Cluster Forense com {len(vis_nodes)} entidades ativas e {len(vis_edges)} conexões ponderadas.")
            with c_gbtn:
                st.download_button(label="💾 Baixar Dossiê HTML", data=html_codigo, file_name=f"Dossie_Forense_{cluster_selecionado['id']}.html", mime="text/html", use_container_width=True)

        elif "Radar Territorial" in nome_visao or "Mapa" in nome_visao:
            if not df_detalhes.empty and "latitude" in df_detalhes.columns:
                df_mapa = df_detalhes.dropna(subset=["latitude", "longitude"]).copy()
                if not df_mapa.empty:
                    lat_centro = float(df_mapa["latitude"].mean())
                    lon_centro = float(df_mapa["longitude"].mean())

                    camada_pontos = pdk.Layer(
                        "ScatterplotLayer", data=df_mapa, get_position="[longitude, latitude]",
                        get_color="[56, 189, 248, 180]", get_line_color="[255, 255, 255, 220]",
                        line_width_min_pixels=1, stroked=True, get_radius=3200, radius_min_pixels=6, radius_max_pixels=20, pickable=True, auto_highlight=True
                    )

                    view_state = pdk.ViewState(latitude=lat_centro, longitude=lon_centro, zoom=7, pitch=0)
                    tooltip = {
                        "html": "<div style='font-family:sans-serif; padding:4px;'><b style='color:#38BDF8;'>🚨 Assistência: {id_assistencia}</b><br/><b>Serviço:</b> {servico}<br/><b>Titular:</b> {titular}<br/><b>Local:</b> {cidade}/{uf}<br/><b>Data:</b> {data}</div>",
                        "style": {"backgroundColor": "rgba(11, 17, 30, 0.95)", "color": "#F8FAFC", "border": "1px solid #1E293B", "borderRadius": "6px", "fontSize": "11px", "zIndex": "1000"}
                    }
                    deck = pdk.Deck(layers=[camada_pontos], initial_view_state=view_state, tooltip=tooltip, map_style="dark")
                    st.pydeck_chart(deck, use_container_width=True)
                    st.dataframe(df_detalhes.groupby(["cidade", "uf"]).size().reset_index(name="Total").sort_values(by="Total", ascending=False), use_container_width=True, hide_index=True)
                else:
                    st.info("Sem coordenadas para plotar.")
            else:
                st.info("Sem dados de localização.")

        elif "Tabela" in nome_visao:
            st.dataframe(df_detalhes[["data", "id_assistencia", "servico", "titular", "cpf", "telefone", "placa", "bairro", "cidade", "uf"]], use_container_width=True, hide_index=True)
            csv_dados = df_detalhes.to_csv(index=False).encode('utf-8')
            st.download_button(label="📥 Baixar Ocorrências em CSV", data=csv_dados, file_name=f"Assistencias_{cluster_selecionado['id']}.csv", mime="text/csv")

        elif "Temporal" in nome_visao:
            if not df_detalhes.empty and "data" in df_detalhes.columns:
                df_temporal = df_detalhes.copy()
                df_temporal["Periodo"] = pd.to_datetime(df_temporal["data"], errors="coerce").dt.strftime("%Y-%m")
                df_temporal = df_temporal.dropna(subset=["Periodo"]).groupby("Periodo").size().reset_index(name="Assistências").sort_values(by="Periodo")
                st.bar_chart(data=df_temporal.set_index("Periodo"), color="#38BDF8", use_container_width=True)
            else:
                st.info("Sem datas disponíveis.")

        elif "Expansão" in nome_visao:
            st.markdown(f"#### 🔍 Cruzamento da Célula contra as Criações Diárias")
            resumo_intel, df_matches_criacao, df_novos_suspeitos = cruzar_com_criacoes_diarias(lista_cpfs, lista_tels, lista_placas)

            if resumo_intel is None:
                st.info("💡 Criações Diárias ainda não ingeridas.")
            elif resumo_intel["total_assistencias"] == 0:
                st.success("✅ Sem acionamentos adicionais nas Criações Diárias.")
            else:
                c1, c2, c3, c4 = st.columns(4)
                with c1: st.markdown(f"""<div class="intel-card"><span style="color:#94A3B8; font-size:10px;">ACIONAMENTOS</span><br/><b style="font-size:18px; color:#38BDF8;">{resumo_intel['total_assistencias']}</b></div>""", unsafe_allow_html=True)
                with c2: st.markdown(f"""<div class="intel-card"><span style="color:#94A3B8; font-size:10px;">NOVOS TELS</span><br/><b style="font-size:18px; color:#EF4444;">+{resumo_intel['novos_tels']}</b></div>""", unsafe_allow_html=True)
                with c3: st.markdown(f"""<div class="intel-card"><span style="color:#94A3B8; font-size:10px;">NOVAS PLACAS</span><br/><b style="font-size:18px; color:#A78BFA;">+{resumo_intel['novas_placas']}</b></div>""", unsafe_allow_html=True)
                with c4: st.markdown(f"""<div class="intel-card"><span style="color:#94A3B8; font-size:10px;">NOVOS CPFS</span><br/><b style="font-size:18px; color:#FBBF24;">+{resumo_intel['novos_cpfs']}</b></div>""", unsafe_allow_html=True)

                st.markdown("<br/>", unsafe_allow_html=True)
                if not df_novos_suspeitos.empty:
                    st.markdown(f"**🚨 {len(df_novos_suspeitos)} Novos Suspeitos Detectados:**")
                    st.dataframe(df_novos_suspeitos, use_container_width=True, hide_index=True)
                    
                    c_inc1, c_inc2 = st.columns([2.5, 2.5])
                    with c_inc1:
                        if st.button("📥 Incorporar Novos Suspeitos à Base Mestra Desta Quadrilha", type="primary", use_container_width=True):
                            tot_inc = 0
                            for _, r_s in df_novos_suspeitos.iterrows():
                                tp = "CPF" if "CPF" in r_s["Tipo"] else ("TELEFONE" if "Telefone" in r_s["Tipo"] else "PLACA")
                                cadastrar_entidade_suspeita(
                                    tipo=tp, valor_bruto=r_s["Dado Suspeito"], nome_ref=r_s.get("Titular", ""),
                                    motivo="Pescado na Expansão com Criações", id_caso=cluster_selecionado["id"],
                                    quadrilha=nome_quadrilha_display, status="Ativo", analista="INVESTIGAÇÃO"
                                )
                                tot_inc += 1
                            st.success(f"✅ {tot_inc} entidades vinculadas ao Caso {cluster_selecionado['id']} na Base Mestra!")
                            st.rerun()

                    with c_inc2:
                        csv_suspeitos = df_novos_suspeitos.to_csv(index=False).encode('utf-8')
                        st.download_button(label="📥 Exportar Suspeitos (.csv)", data=csv_suspeitos, file_name=f"Novos_Suspeitos_{cluster_selecionado['id']}.csv", mime="text/csv", use_container_width=True)
                else:
                    st.info("Acionamentos pertencem apenas a entidades já catalogadas.")

    if st.session_state["cfg_modo_exibicao"] == "📑 Abas Clássicas":
        tab_g, tab_e, tab_t, tab_m, tab_d, tab_caso = st.tabs([
            "🕸️ Grafo de Vínculos Forense", "🔍 Expansão com Criações Diárias", "🕒 Evolução Temporal",
            "📍 Radar Territorial", "📋 Tabela de Ocorrências", "📁 Gestão do Caso & Dossiê"
        ])
        with tab_g: render_visao("Grafo", altura_pixel=740)
        with tab_e: render_visao("Expansão")
        with tab_t: render_visao("Temporal")
        with tab_m: render_visao("Radar Territorial")
        with tab_d: render_visao("Tabela")
        with tab_caso:
            st.markdown(f"#### 📁 Identificação & Parecer da Investigação ({cluster_selecionado['id']})")
            c_f1, c_f2, c_f3 = st.columns([2.5, 1.5, 1.5])
            with c_f1: novo_nome = st.text_input("Nome da Quadrilha:", value=dados_caso["nome_personalizado"], placeholder="Ex: Quadrilha Sapeaçu")
            with c_f2: novo_status = st.selectbox("Status:", ["Em Investigação", "Confirmado Fraude", "Monitoramento Contínuo", "Falso Positivo", "Arquivado"], index=["Em Investigação", "Confirmado Fraude", "Monitoramento Contínuo", "Falso Positivo", "Arquivado"].index(dados_caso["status"]) if dados_caso["status"] in ["Em Investigação", "Confirmado Fraude", "Monitoramento Contínuo", "Falso Positivo", "Arquivado"] else 0)
            with c_f3: novo_analista = st.text_input("Analista:", value=dados_caso["analista_responsavel"])
            novo_parecer = st.text_area("Parecer Técnico:", value=dados_caso["parecer"], height=140)
            if st.button("💾 Salvar Metadados do Caso no Banco", type="primary", use_container_width=True):
                salvar_dados_caso(cluster_selecionado["id"], novo_nome, novo_status, novo_analista, novo_parecer)
                st.success("✅ Metadados gravados com sucesso!")
                st.rerun()

            # Trilha de Auditoria e Histórico Append-Only de Pareceres
            st.markdown("---")
            st.markdown("##### 📜 Trilha de Auditoria & Histórico de Pareceres Deste Caso")
            historico_notas = carregar_historico_pareceres(cluster_selecionado['id'])
            if historico_notas:
                for h in historico_notas:
                    st.markdown(f"""
                    <div style="background:#0E1726; border:1px solid #1E293B; border-radius:6px; padding:10px 14px; margin-bottom:8px;">
                        <span style="color:#38BDF8; font-size:12px; font-weight:600;">{h['data_registro']}</span> • 
                        <span style="color:#FBBF24; font-size:12px;">Analista: <b>{h['analista'] or 'Sistema'}</b></span> • 
                        <span style="color:#10B981; font-size:12px;">Status: <b>{h['status']}</b></span>
                        <div style="color:#E2E8F0; font-size:13px; margin-top:6px; white-space: pre-wrap;">{h['parecer']}</div>
                    </div>
                    """, unsafe_allow_html=True)
            else:
                st.caption("Nenhum parecer técnico anterior registrado no histórico para este caso.")

    else:
        c_pesq, c_prop, c_pdir = st.columns([2.5, 1.4, 2.5])
        with c_pesq:
            idx_pesq = OPCOES_PAINEL.index(st.session_state["cfg_painel_esquerdo"]) if st.session_state["cfg_painel_esquerdo"] in OPCOES_PAINEL else 0
            st.selectbox("Painel Esquerdo:", OPCOES_PAINEL, index=idx_pesq, key="w_painel_esq", on_change=cb_atualizar_pesq)
        with c_prop:
            idx_prop = list(PROPORCOES_MAP.keys()).index(st.session_state["cfg_cockpit_proporcao"]) if st.session_state["cfg_cockpit_proporcao"] in PROPORCOES_MAP else 0
            st.selectbox("Proporção:", list(PROPORCOES_MAP.keys()), index=idx_prop, key="w_proporcao", on_change=cb_atualizar_prop)
        with c_pdir:
            idx_pdir = OPCOES_PAINEL.index(st.session_state["cfg_painel_direito"]) if st.session_state["cfg_painel_direito"] in OPCOES_PAINEL else 1
            st.selectbox("Painel Direito:", OPCOES_PAINEL, index=idx_pdir, key="w_painel_dir", on_change=cb_atualizar_pdir)

        col_left, col_right = st.columns(PROPORCOES_MAP[st.session_state["cfg_cockpit_proporcao"]])
        with col_left:
            st.markdown(f"<div style='margin-bottom:4px;'><b style='color:#38BDF8;'>{st.session_state['cfg_painel_esquerdo']}</b></div>", unsafe_allow_html=True)
            render_visao(st.session_state["cfg_painel_esquerdo"], altura_pixel=680)
        with col_right:
            st.markdown(f"<div style='margin-bottom:4px;'><b style='color:#38BDF8;'>{st.session_state['cfg_painel_direito']}</b></div>", unsafe_allow_html=True)
            render_visao(st.session_state["cfg_painel_direito"], altura_pixel=680)