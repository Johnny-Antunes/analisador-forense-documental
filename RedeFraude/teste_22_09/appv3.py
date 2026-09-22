"""
Módulo: appv3.py
Objetivo: Interface executiva da Mesa de Inteligência Forense de Fraudes em Assistências (LCFO).
          Suporta investigação unificada (Casos Oficiais e Descobertas Ativas),
          sincronização triangular no 1º clique (Tabela -> Mapa -> Grafo e Grafo -> Tabela)
          e Funil da Tela 4.
"""

from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple
import json
import re
import streamlit as st
import pandas as pd
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
    carregar_historico_pareceres, resetar_layout_caso,
    cadastrar_cidade_risco, listar_cidades_risco, remover_cidade_risco,
    reparar_mojibake_historico,
    ocultar_no_do_caso, restaurar_no_do_caso, listar_nos_ocultos_do_caso
)
from graph_engine import (
    carregar_redes, processar_subgrafo_caso, gerar_html_grafo, processar_grafo_dataframe
)
from correlation_engine import (
    calcular_score_caso, obter_scores_triagem
)
from anomalias_engine import (
    obter_radar_anomalias_macro, extrair_top_infratores_municipio
)
from utils import formatar_mencoes_forenses

# Importação defensiva do componente bidirecional
try:
    from grafo_component import renderizar_grafo_bidirecional
    TEM_COMPONENTE_BIDIRECIONAL = True
except ImportError:
    TEM_COMPONENTE_BIDIRECIONAL = False

# =====================================================
# 1. CONFIGURAÇÃO DE PÁGINA & CSS FORENSE SÓBRIO
# =====================================================
st.set_page_config(
    page_title="Mesa de Inteligência Forense - LCFO",
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
    "organico": "Teia Fluida Orgânica",
    "radial": "Radial Peacock (i2)",
    "arvore": "Árvore Forense",
    "pontes": "Pontes & Gargalos",
    "subredes": "Sub-redes em Grade",
    "comunidades": "Comunidades & Facções"
}

OPCOES_PAINEL_OFICIAL = [
    "Radar Territorial (Mapa)",
    "Grafo de Vínculos",
    "Tabela de Ocorrências",
    "Evolução Temporal",
    "Expansão com Criações"
]

OPCOES_PAINEL_DESCOBERTA = [
    "Radar Territorial (Mapa)",
    "Grafo de Vínculos",
    "Tabela de Ocorrências",
    "Evolução Temporal"
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
    st.session_state["aba_principal_triagem"] = "Células e Redes da Blacklist"

if "cfg_layout_ativo" not in st.session_state:
    st.session_state["cfg_layout_ativo"] = "organico"

if "cfg_modo_exibicao" not in st.session_state:
    st.session_state["cfg_modo_exibicao"] = "Abas Clássicas"

if "cfg_painel_esquerdo" not in st.session_state:
    st.session_state["cfg_painel_esquerdo"] = "Radar Territorial (Mapa)"

if "cfg_painel_direito" not in st.session_state:
    st.session_state["cfg_painel_direito"] = "Grafo de Vínculos"

if "cfg_cockpit_proporcao" not in st.session_state:
    st.session_state["cfg_cockpit_proporcao"] = "50% | 50%"

if "coordenada_foco" not in st.session_state:
    st.session_state["coordenada_foco"] = None

if "entidade_foco_grafo" not in st.session_state:
    st.session_state["entidade_foco_grafo"] = None

if "municipio_foco_funil" not in st.session_state:
    st.session_state["municipio_foco_funil"] = None

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


def extrair_linhas_selecionadas(evento: Any) -> List[int]:
    """Extrai com segurança os índices de linhas selecionadas no st.dataframe."""
    if not evento:
        return []
    try:
        if hasattr(evento, "selection"):
            sel = evento.selection
            if hasattr(sel, "rows"):
                return list(sel.rows)
            elif isinstance(sel, dict) and "rows" in sel:
                return list(sel["rows"])
        elif isinstance(evento, dict):
            sel = evento.get("selection", {})
            if isinstance(sel, dict) and "rows" in sel:
                return list(sel["rows"])
            elif hasattr(sel, "rows"):
                return list(sel.rows)
    except Exception:
        return []
    return []


def montar_hud_bar(cor: str, nome_exibicao: str, score: int, nivel: str,
                   campos_extra_html: str, total_registros: int, periodo_str: str) -> str:
    """Constrói a HUD Bar como string linear sem quebras de linha que possam induzir bloco de código."""
    return (
        f'<div class="hud-bar" style="border-color: {cor};">'
        f'<div style="display:flex; align-items:center; gap:18px; flex-wrap:wrap;">'
        f'<span style="font-size:0.95rem; font-weight:700; color:#F8FAFC;">{nome_exibicao.upper()}</span>'
        f'<span style="font-size:0.82rem; color:#94A3B8;">Fraud Score: <b style="color:{cor};">{score}/100 [{nivel}]</b></span>'
        f'{campos_extra_html}'
        f'<span style="font-size:0.82rem; color:#38BDF8;">Acionamentos: <b>{total_registros} registros</b></span>'
        f'<span style="font-size:0.82rem; color:#94A3B8;">Período: <b style="color:#E2E8F0;">{periodo_str}</b></span>'
        f'</div>'
        f'</div>'
    )


def montar_caixa_fatores(cor: str, fatores: List[str]) -> str:
    """Caixa de indicadores forenses em linha única para proteção contra quebra de Markdown."""
    texto_fatores = ' • '.join(fatores) if fatores else "Comportamento relacional estável dentro da normalidade"
    return (
        f'<div style="background: rgba(11, 17, 30, 0.95); border-left: 4px solid {cor}; border-radius: 6px; '
        f'padding: 8px 14px; margin-bottom: 12px; border-top: 1px solid #1E293B; border-right: 1px solid #1E293B; '
        f'border-bottom: 1px solid #1E293B;">'
        f'<div style="display: flex; align-items: center; justify-content: space-between;">'
        f'<span style="color: {cor}; font-weight: 700; font-size: 13px;">INDICADORES FORENSES IDENTIFICADOS</span>'
        f'<span style="color: #94A3B8; font-size: 11px;">Avaliação Comportamental & Topológica</span>'
        f'</div>'
        f'<div style="color: #E2E8F0; font-size: 12px; margin-top: 4px;">{texto_fatores}</div>'
        f'</div>'
    )

# =====================================================
# 3. GESTÃO DE BASES (SIDEBAR COM EXPANDERS)
# =====================================================
with st.sidebar:
    st.markdown("### Gestão de Bases de Dados")
    
    with st.expander("Base Blacklist (Duplicidades)", expanded=False):
        pasta_padrao_bl = CAMINHO_REDE_OFICIAL if Path(CAMINHO_REDE_OFICIAL).exists() else str(PASTA_LOCAL_BLACKLIST)
        caminho_input = st.text_input("Pasta Blacklist (.xlsx/.csv):", value=pasta_padrao_bl)
        
        qtd_detectada = 0
        p_check = Path(caminho_input.strip())
        if p_check.exists():
            qtd_detectada = len(list(p_check.glob("*.xlsx"))) + len(list(p_check.glob("*.csv")))
            st.caption(f"Status: **{qtd_detectada}** planilhas detectadas.")
        else:
            st.caption("Status: Caminho não acessível.")

        if st.button("Ingerir Blacklist (Novos)", use_container_width=True):
            with st.spinner("Ingerindo arquivos novos da Blacklist..."):
                qtd_arq, qtd_reg, msg_info, erros_bl = carregar_arquivos_para_sqlite(caminho_input, forcar_releitura=False)
                st.cache_data.clear()
                if qtd_reg > 0:
                    st.success(f"{qtd_reg:,} novas assistências inseridas e Base Mestra semeada.")
                else:
                    st.info(msg_info or "Base já atualizada.")
                if erros_bl:
                    with st.expander(f"{len(erros_bl)} arquivos com aviso"):
                        for e in erros_bl: st.caption(f"• {e}")
                st.rerun()

    with st.expander("Criações Diárias (Base Geral)", expanded=False):
        pasta_padrao_cr = CAMINHO_REDE_CRIACAO if Path(CAMINHO_REDE_CRIACAO).exists() else str(PASTA_LOCAL_CRIACAO)
        caminho_criacao = st.text_input("Pasta Criação (.xlsx/.csv):", value=pasta_padrao_cr)
        
        qtd_criacao = 0
        p_criacao = Path(caminho_criacao.strip())
        if p_criacao.exists():
            qtd_criacao = len(list(p_criacao.glob("*.xlsx"))) + len(list(p_criacao.glob("*.csv")))
            st.caption(f"Status: **{qtd_criacao}** arquivos detectados.")
        else:
            st.caption("Status: Caminho não acessível.")

        if st.button("Ingerir Criações no SQLite (Novos)", use_container_width=True):
            with st.spinner("Ingerindo arquivos novos de Criações..."):
                qtd_arq_c, qtd_reg_c, msg_c, erros_c = carregar_criacoes_diarias_para_sqlite(caminho_criacao, forcar_releitura=False)
                st.cache_data.clear()
                if qtd_reg_c > 0:
                    st.success(f"{qtd_reg_c:,} novas assistências ingeridas.")
                else:
                    st.info(msg_c or "Nenhum arquivo novo para ingerir.")
                if erros_c:
                    with st.expander(f"{len(erros_c)} arquivos com aviso"):
                        for e in erros_c: st.caption(f"• {e}")
                st.rerun()

    with st.expander("Manutenção e Integridade", expanded=False):
        st.caption("Corrige nomes de cidade/bairro/titular com codificação corrompida (mojibake).")
        if st.button("Reparar Codificação Histórica", use_container_width=True):
            with st.spinner("Varrendo e corrigindo textos no banco..."):
                resultado = reparar_mojibake_historico()
                st.cache_data.clear()
                total_geral = sum(resultado.values())
                if total_geral > 0:
                    detalhe = " | ".join(f"{k}: {v}" for k, v in resultado.items() if v > 0)
                    st.success(f"{total_geral} registros corrigidos. ({detalhe})")
                else:
                    st.info("Nenhuma corrupção de codificação encontrada nos dados atuais.")
                st.rerun()

# =====================================================
# 4. CARREGAMENTO DAS REDES E METADADOS
# =====================================================
G, cluster_info = carregar_redes()

if not G or not cluster_info:
    st.info("Nenhum dado processado no banco. Verifique o caminho da pasta na barra lateral e clique em 'Ingerir Blacklist'.")
    st.stop()

radar_alertas = obter_radar_expansoes(cluster_info)
casos_cadastrados = carregar_todos_casos_cadastrados()

opcoes_celulas = {}
for c in cluster_info[:150]:
    info_caso = casos_cadastrados.get(c["id"], {})
    nome_custom = info_caso.get("nome", "")
    tag_nome = f" - {nome_custom.upper()}" if nome_custom else ""
    alerta_badge = f" (+{radar_alertas[c['id']]} novas)" if c["id"] in radar_alertas else ""
    rotulo = f"[{c['id']}]{tag_nome} ({c['tamanho']} nós) - {c['hub_label']}{alerta_badge}"
    opcoes_celulas[rotulo] = c['id']

mapa_id_para_label = {v: k for k, v in opcoes_celulas.items()}

# =====================================================
# 5. CONTROLES NA BARRA LATERAL & BUSCA UNIFICADA
# =====================================================
with st.sidebar:
    st.markdown("---")
    
    if st.session_state["celula_ativa_id"] is not None or st.session_state["modo_descoberta_ativo"]:
        if st.button("Voltar à Triagem Geral", use_container_width=True):
            st.session_state["celula_ativa_id"] = None
            st.session_state["modo_descoberta_ativo"] = False
            st.session_state["dados_descoberta"] = None
            st.session_state["coordenada_foco"] = None
            st.session_state["entidade_foco_grafo"] = None
            st.rerun()
        
        if st.session_state["celula_ativa_id"] is not None:
            st.markdown("### Caso em Análise")
            label_atual = mapa_id_para_label.get(st.session_state["celula_ativa_id"], list(opcoes_celulas.keys())[0])
            idx_caso = list(opcoes_celulas.keys()).index(label_atual) if label_atual in opcoes_celulas else 0
            novo_caso_selecionado = st.selectbox("Alternar Caso:", list(opcoes_celulas.keys()), index=idx_caso)
            
            if opcoes_celulas[novo_caso_selecionado] != st.session_state["celula_ativa_id"]:
                st.session_state["celula_ativa_id"] = opcoes_celulas[novo_caso_selecionado]
                st.session_state["coordenada_foco"] = None
                st.session_state["entidade_foco_grafo"] = None
                st.rerun()

            cluster_selecionado = next((c for c in cluster_info if c["id"] == st.session_state["celula_ativa_id"]), cluster_info[0])

            st.markdown("---")
            st.markdown("### Filtros de Entidades")
            mostrar_tel = st.checkbox(f"[TEL] Telefones ({cluster_selecionado['qtd_tels']})", value=True)
            mostrar_cpf = st.checkbox(f"[CPF] Titulares ({cluster_selecionado['qtd_cpfs']})", value=True)
            mostrar_placa = st.checkbox(f"[PLACA] Placas ({cluster_selecionado['qtd_placas']})", value=True)

        st.markdown("---")
        st.markdown("### Layout Topológico")
        idx_layout = list(LAYOUT_MAP.keys()).index(st.session_state["cfg_layout_ativo"]) if st.session_state["cfg_layout_ativo"] in LAYOUT_MAP else 0
        st.selectbox("Estrutura:", options=list(LAYOUT_MAP.keys()), format_func=lambda k: LAYOUT_MAP[k], index=idx_layout, key="w_layout_ativo", on_change=cb_atualizar_layout)

        st.markdown("---")
        st.markdown("### Ajustes Visuais")
        font_slider = st.slider("Fonte Base do Grafo (px):", min_value=10, max_value=18, value=12)
        espacamento_slider = st.slider("Dispersão / Distância (px):", min_value=180, max_value=450, value=280, step=20)

    else:
        st.markdown("### Busca Unificada de Inteligência")
        st.caption("Pesquise na Blacklist ou investigue alvos inéditos nas Criações Diárias.")
        termo_busca_global = st.text_input("CPF, Placa ou Telefone:", placeholder="Digite para rastrear...").strip()
        
        if termo_busca_global:
            termo_clean = re.sub(r'[^a-zA-Z0-9]', '', termo_busca_global).upper()
            alvos_bl = [n for n in G.nodes if termo_clean in n]
            
            if alvos_bl:
                alvo = alvos_bl[0]
                for c in cluster_info:
                    if alvo in c["nodes"]:
                        nome_caso_bl = casos_cadastrados.get(c["id"], {}).get("nome", c["hub_label"])
                        st.success(f"Localizado no Caso Oficial: {nome_caso_bl} ({c['id']})")
                        if st.button(f"Abrir Caso {c['id']}", use_container_width=True, type="primary"):
                            st.session_state["celula_ativa_id"] = c["id"]
                            st.session_state["coordenada_foco"] = None
                            st.session_state["entidade_foco_grafo"] = None
                            st.rerun()
                        break
            else:
                resumo_disc, df_disc, ents_disc = investigar_alvo_em_criacoes_diarias(termo_busca_global)
                if resumo_disc and resumo_disc["total_assistencias"] > 0:
                    st.warning(f"Dados localizados nas Criações Diárias:\n• {resumo_disc['total_assistencias']} assistências identificadas\n• Rede: {resumo_disc['qtd_cpfs']} CPFs, {resumo_disc['qtd_tels']} telefones, {resumo_disc['qtd_placas']} placas.")
                    if st.button("Abrir Mesa de Investigação Ativa", use_container_width=True, type="primary"):
                        st.session_state["modo_descoberta_ativo"] = True
                        st.session_state["dados_descoberta"] = {
                            "resumo": resumo_disc,
                            "df": df_disc,
                            "entidades": ents_disc
                        }
                        st.session_state["coordenada_foco"] = None
                        st.session_state["entidade_foco_grafo"] = None
                        st.rerun()
                else:
                    st.info("Nenhum acionamento localizado para este dado em nenhuma das bases.")

# ====================================================================
# 6. MOTOR UNIFICADO DE INVESTIGAÇÃO (MESA FORENSE)
# ====================================================================
def renderizar_mesa_investigacao(
    df_dados: pd.DataFrame,
    identificador_caso: str,
    nome_exibicao: str,
    hub_id: str,
    eh_oficial: bool,
    dados_caso_oficial: Optional[Dict[str, Any]] = None,
    resumo_descoberta: Optional[Dict[str, Any]] = None,
    cluster_obj: Optional[Dict[str, Any]] = None
):
    # =================================================================
    # RESOLUÇÃO ANTECIPADA NO 1º CLIQUE: Tabela -> Mapa & Grafo
    #
    # Só verifica a(s) chave(s) de tabela do modo de exibição REALMENTE
    # ativo agora. Antes, o loop varria ["aba", "esq", "dir"] sempre,
    # e como o session_state de um widget com key não é limpo quando o
    # widget para de ser renderizado (ex: ao trocar de "Abas Clássicas"
    # para "Cockpit Dividido"), uma seleção antiga em outro modo podia
    # "assombrar" a tela atual e sobrescrever o foco a cada rerun.
    # =================================================================
    sufixos_ativos = ["aba"] if st.session_state["cfg_modo_exibicao"] == "Abas Clássicas" else ["esq", "dir"]
    for sufixo_teste in sufixos_ativos:
        chave_tab = f"tab_ocorr_{identificador_caso}_{sufixo_teste}"
        if chave_tab in st.session_state:
            linhas_clicadas = extrair_linhas_selecionadas(st.session_state[chave_tab])
            if linhas_clicadas and 0 <= linhas_clicadas[0] < len(df_dados):
                reg_sel = df_dados.iloc[linhas_clicadas[0]]
                if pd.notna(reg_sel.get("latitude")) and pd.notna(reg_sel.get("longitude")):
                    st.session_state["coordenada_foco"] = (float(reg_sel["latitude"]), float(reg_sel["longitude"]))

                p_cand = str(reg_sel.get("placa", "")).strip()
                t_cand = str(reg_sel.get("telefone", "")).strip()
                c_cand = str(reg_sel.get("cpf", "")).strip()

                if p_cand: st.session_state["entidade_foco_grafo"] = f"PLACA_{p_cand}"
                elif t_cand: st.session_state["entidade_foco_grafo"] = f"TEL_{t_cand}"
                elif c_cand: st.session_state["entidade_foco_grafo"] = f"CPF_{c_cand}"
                break

    col_nav1, col_nav2, col_nav3 = st.columns([3.2, 1.8, 1.0])
    with col_nav1:
        prefixo_nav = "Mesa de Triagem / Caso Oficial: " if eh_oficial else "Mesa de Triagem / Descoberta Ativa: "
        st.markdown(f"<div style='padding-top:4px;'><span style='color:#94A3B8; font-size:12px;'>{prefixo_nav}</span> <b style='color:#38BDF8; font-size:14px;'>{nome_exibicao}</b></div>", unsafe_allow_html=True)
    with col_nav2:
        idx_modo = 0 if st.session_state["cfg_modo_exibicao"] == "Abas Clássicas" else 1
        st.radio("Visualização:", ["Abas Clássicas", "Cockpit Dividido"], index=idx_modo, horizontal=True, key="w_modo_exibicao", on_change=cb_atualizar_modo, label_visibility="collapsed")
    with col_nav3:
        btn_label = "Sair do Caso" if eh_oficial else "Fechar"
        if st.button(btn_label, use_container_width=True):
            st.session_state["celula_ativa_id"] = None
            st.session_state["modo_descoberta_ativo"] = False
            st.session_state["dados_descoberta"] = None
            st.session_state["coordenada_foco"] = None
            st.session_state["entidade_foco_grafo"] = None
            st.rerun()

    # Processamento relacional do subgrafo
    if eh_oficial:
        filtro_tipos = []
        if mostrar_tel: filtro_tipos.append("telefone")
        if mostrar_cpf: filtro_tipos.append("cpf")
        if mostrar_placa: filtro_tipos.append("placa")
        vis_nodes, vis_edges = processar_subgrafo_caso(
            G=G,
            cluster_nodes=cluster_obj["nodes"] if cluster_obj else [],
            filtro_tipos=filtro_tipos,
            hub_id=hub_id,
            font_slider=font_slider,
            id_caso=identificador_caso
        )
        max_bet = max((n.get("betweenness", 0.0) for n in vis_nodes), default=0.0)
        subG_ativo = G.subgraph(cluster_obj["nodes"]) if cluster_obj else None
        score, nivel, cor, fatores, _ = calcular_score_caso(
            cluster_obj=cluster_obj,
            df_assistencias=df_dados,
            subgrafo=subG_ativo,
            betweenness_precalculado=max_bet
        )
    else:
        vis_nodes, vis_edges, hub_id_calc = processar_grafo_dataframe(
            df_dados, alvo_principal=resumo_descoberta.get("termo_limpo", "") if resumo_descoberta else "", font_slider=font_slider
        )
        hub_id = hub_id_calc
        score, nivel, cor, fatores, _ = calcular_score_caso(
            cluster_obj=None,
            df_assistencias=df_dados,
            subgrafo=None
        )

    periodo_str = "Sem datas disponíveis"
    if not df_dados.empty and "data" in df_dados.columns:
        df_temp = df_dados[df_dados["data"].astype(str).str.strip() != ""].copy()
        if not df_temp.empty:
            dts_validas = pd.to_datetime(df_temp["data"], errors="coerce").dropna()
            if not dts_validas.empty:
                periodo_str = f"{dts_validas.min().strftime('%d/%m/%Y')} → {dts_validas.max().strftime('%d/%m/%Y')}"

    # Metadados complementares da HUD Bar
    campos_extra = []
    if eh_oficial and dados_caso_oficial:
        st_val = dados_caso_oficial.get("status", "Em Investigação")
        campos_extra.append(f'<span style="font-size:0.82rem; color:#94A3B8;">Status: <b style="color:#10B981;">{st_val}</b></span>')
        if hub_id:
            campos_extra.append(f'<span style="font-size:0.82rem; color:#FBBF24;">Âncora: <b>{hub_id}</b></span>')
    elif not eh_oficial and resumo_descoberta:
        campos_extra.append(
            f'<span style="font-size:0.82rem; color:#A78BFA;">Entidades: <b>'
            f'{resumo_descoberta.get("qtd_cpfs", 0)} CPFs | {resumo_descoberta.get("qtd_tels", 0)} Tels | '
            f'{resumo_descoberta.get("qtd_placas", 0)} Placas</b></span>'
        )
    campos_extra_html = "".join(campos_extra)

    # Renderização blindada em string única
    st.markdown(
        montar_hud_bar(cor, nome_exibicao, score, nivel, campos_extra_html, len(df_dados), periodo_str),
        unsafe_allow_html=True
    )
    st.markdown(montar_caixa_fatores(cor, fatores), unsafe_allow_html=True)

    # Barra de controle de foco ativo
    if st.session_state["coordenada_foco"] or st.session_state["entidade_foco_grafo"]:
        c_foco1, c_foco2 = st.columns([4.0, 1.0])
        with c_foco1:
            ent_nome = st.session_state["entidade_foco_grafo"] or "Ocorrência selecionada"
            st.info(f"Filtro ativo na assistência: **{ent_nome}** — Mapa e grafo sincronizados neste ponto.")
        with c_foco2:
            if st.button("Limpar Foco", use_container_width=True):
                st.session_state["coordenada_foco"] = None
                st.session_state["entidade_foco_grafo"] = None
                st.rerun()

    # Painel de Promoção/Anexação para Descoberta Ativa
    if not eh_oficial:
        with st.expander("Ações de Promoção e Vínculo Desta Rede", expanded=False):
            tipo_acao = st.radio("Destino da rede:", ["Criar Novo Caso Oficial", "Anexar a Caso Existente"], horizontal=True)
            if tipo_acao == "Criar Novo Caso Oficial":
                c_pr1, c_pr2, c_pr3 = st.columns([2.5, 1.5, 1.5])
                with c_pr1: nome_prom = st.text_input("Nome da Operação / Quadrilha:", placeholder="Ex: Operação Rota Sul")
                with c_pr2: st_prom = st.selectbox("Status:", ["Confirmado Fraude", "Em Investigação", "Monitoramento Contínuo", "Sem Irregularidade Identificada"])
                with c_pr3: analista_prom = st.text_input("Analista:", placeholder="Seu nome")
                parecer_prom = st.text_area("Parecer da Descoberta:", placeholder="Descreva os vínculos observados...")

                if st.button("Promover a Novo Caso Oficial", type="primary", use_container_width=True):
                    if nome_prom.strip():
                        nos_promovidos = promover_descoberta_para_caso(df_dados)
                        st.cache_data.clear()
                        _, cluster_info_novo = carregar_redes()
                        cluster_promovido = next((c for c in cluster_info_novo if any(n in c["nodes"] for n in nos_promovidos)), None)
                        if cluster_promovido:
                            vincular_caso_e_entidades(cluster_promovido["id"], nos_promovidos, nome_prom, st_prom, analista_prom, parecer_prom)
                            st.session_state["modo_descoberta_ativo"] = False
                            st.session_state["dados_descoberta"] = None
                            st.session_state["celula_ativa_id"] = cluster_promovido["id"]
                            st.success(f"Rede promovida e vinculada ao Caso {cluster_promovido['id']}.")
                            st.rerun()
                    else:
                        st.error("Informe um nome para o caso antes de promover.")
            else:
                c_an1, c_an2 = st.columns([3.0, 2.0])
                with c_an1:
                    caso_destino_label = st.selectbox("Selecione o Caso de Destino:", list(opcoes_celulas.keys()))
                    id_destino = opcoes_celulas[caso_destino_label]
                with c_an2:
                    analista_anex = st.text_input("Analista Responsável:", key="anex_analista")
                obs_anexo = st.text_input("Motivo do Vínculo:", placeholder="Ex: Telefone comum identificado acionando para os mesmos alvos.")

                if st.button(f"Anexar Ocorrências ao Caso {id_destino}", type="primary", use_container_width=True):
                    anexar_descoberta_a_caso_existente(df_dados, id_destino, analista_anex, obs_anexo)
                    st.cache_data.clear()
                    st.session_state["modo_descoberta_ativo"] = False
                    st.session_state["dados_descoberta"] = None
                    st.session_state["celula_ativa_id"] = id_destino
                    st.success(f"Dados anexados com sucesso ao Caso {id_destino}.")
                    st.rerun()

    # Validação do target_node_id no conjunto de nós exibidos
    alvo_grafo_id = st.session_state.get("entidade_foco_grafo") or ""
    if alvo_grafo_id and not any(n["id"] == alvo_grafo_id for n in vis_nodes):
        alvo_grafo_id = ""

    html_codigo = gerar_html_grafo(
        vis_nodes_json=json.dumps(vis_nodes),
        vis_edges_json=json.dumps(vis_edges),
        base_font_size=font_slider,
        hub_id=hub_id,
        target_node_id=alvo_grafo_id,
        layout_ativo=st.session_state["cfg_layout_ativo"],
        espacamento=espacamento_slider
    )

    def render_modulo(nome_modulo, altura=720, sufixo_key=""):
        if "Grafo" in nome_modulo:
            evento_grafo = None
            if TEM_COMPONENTE_BIDIRECIONAL:
                evento_grafo = renderizar_grafo_bidirecional(
                    nodes=vis_nodes,
                    edges=vis_edges,
                    hub_id=hub_id,
                    target_node_id=alvo_grafo_id,
                    layout_ativo=st.session_state["cfg_layout_ativo"],
                    base_font_size=font_slider,
                    espacamento=espacamento_slider,
                    height=altura,
                    key=f"comp_grafo_{identificador_caso}_{sufixo_key}"
                )
            else:
                components.html(html_codigo, height=altura)

            # Sincronização Grafo -> Mapa & Tabela
            if evento_grafo and isinstance(evento_grafo, dict):
                tipo_ev = evento_grafo.get("tipo")
                node_ev = evento_grafo.get("nodeId")
                if tipo_ev == "node_click" and node_ev and node_ev != st.session_state.get("entidade_foco_grafo"):
                    st.session_state["entidade_foco_grafo"] = node_ev
                    val_puro = node_ev.split("_", 1)[1] if "_" in node_ev else node_ev
                    sub_match = df_dados[
                        (df_dados["cpf"].astype(str) == val_puro) |
                        (df_dados["telefone"].astype(str) == val_puro) |
                        (df_dados["placa"].astype(str) == val_puro)
                    ].dropna(subset=["latitude", "longitude"])
                    if not sub_match.empty:
                        st.session_state["coordenada_foco"] = (float(sub_match.iloc[0]["latitude"]), float(sub_match.iloc[0]["longitude"]))
                    st.rerun()

            if eh_oficial:
                c_gi, c_gb, c_gr = st.columns([2.6, 1.3, 1.3])
                with c_gi: st.caption(f"Cluster Forense com {len(vis_nodes)} entidades ativas e {len(vis_edges)} conexões ponderadas.")
                with c_gb:
                    st.download_button("Baixar Dossiê HTML", html_codigo, file_name=f"Dossie_{identificador_caso}.html", mime="text/html", use_container_width=True, key=f"dl_dos_{identificador_caso}_{sufixo_key}")
                with c_gr:
                    if st.button("Recalcular Layout", use_container_width=True, key=f"rst_lay_{identificador_caso}_{sufixo_key}"):
                        resetar_layout_caso(identificador_caso)
                        st.rerun()

        elif "Radar" in nome_modulo or "Mapa" in nome_modulo:
            df_mapa = df_dados.dropna(subset=["latitude", "longitude"]).copy()
            if not df_mapa.empty:
                foco = st.session_state["coordenada_foco"]
                lat_c = foco[0] if foco else float(df_mapa["latitude"].mean())
                lon_c = foco[1] if foco else float(df_mapa["longitude"].mean())
                zoom_c = 11 if foco else 7

                camada_geral = pdk.Layer(
                    "ScatterplotLayer", data=df_mapa, get_position="[longitude, latitude]",
                    get_color="[56, 189, 248, 180]", get_line_color="[255, 255, 255, 220]",
                    line_width_min_pixels=1, stroked=True, get_radius=3200, radius_min_pixels=6, radius_max_pixels=20, pickable=True
                )
                camadas = [camada_geral]

                if foco:
                    df_foco = pd.DataFrame([{"latitude": foco[0], "longitude": foco[1]}])
                    camada_foco = pdk.Layer(
                        "ScatterplotLayer", data=df_foco, get_position="[longitude, latitude]",
                        get_color="[251, 191, 36, 230]", get_line_color="[255, 255, 255, 255]",
                        line_width_min_pixels=2, stroked=True, get_radius=5500, radius_min_pixels=10, radius_max_pixels=28, pickable=False
                    )
                    camadas.append(camada_foco)

                deck = pdk.Deck(layers=camadas, initial_view_state=pdk.ViewState(latitude=lat_c, longitude=lon_c, zoom=zoom_c), map_style="dark")
                st.pydeck_chart(deck, use_container_width=True)
                st.dataframe(df_dados.groupby(["cidade", "uf"]).size().reset_index(name="Total").sort_values(by="Total", ascending=False), use_container_width=True, hide_index=True)
            else:
                st.info("Sem coordenadas geográficas para plotagem.")

        elif "Tabela" in nome_modulo:
            cols_exib = [c for c in ["data", "id_assistencia", "servico", "titular", "cpf", "telefone", "placa", "bairro", "cidade", "uf", "empresa_cliente"] if c in df_dados.columns]
            st.caption("Selecione uma linha para centralizar o acionamento no Radar Territorial e destacar a entidade no Grafo.")
            
            try:
                st.dataframe(
                    df_dados[cols_exib],
                    use_container_width=True,
                    hide_index=True,
                    on_select="rerun",
                    selection_mode="single-row",
                    key=f"tab_ocorr_{identificador_caso}_{sufixo_key}"
                )
            except Exception:
                st.dataframe(df_dados[cols_exib], use_container_width=True, hide_index=True)

            csv_exp = df_dados.to_csv(index=False).encode('utf-8')
            st.download_button("Baixar Ocorrências em CSV", csv_exp, file_name=f"Assistencias_{identificador_caso}.csv", mime="text/csv", key=f"dl_csv_{identificador_caso}_{sufixo_key}")

        elif "Temporal" in nome_modulo:
            if not df_dados.empty and "data" in df_dados.columns:
                df_temp = df_dados.copy()
                df_temp["Periodo"] = pd.to_datetime(df_temp["data"], errors="coerce").dt.strftime("%Y-%m")
                df_temp = df_temp.dropna(subset=["Periodo"]).groupby("Periodo").size().reset_index(name="Assistências").sort_values(by="Periodo")
                st.bar_chart(data=df_temp.set_index("Periodo"), color="#38BDF8", use_container_width=True)
            else:
                st.info("Sem datas disponíveis para evolução temporal.")

        elif "Expansão" in nome_modulo:
            if not eh_oficial or not cluster_obj:
                st.info("A Expansão com Criações Diárias está disponível apenas para casos oficiais já catalogados na Blacklist.")
            else:
                st.markdown("#### Cruzamento da Célula contra as Criações Diárias")
                cpfs_alvo = [n.replace("CPF_", "") for n in cluster_obj["nodes"] if n.startswith("CPF_")]
                tels_alvo = [n.replace("TEL_", "") for n in cluster_obj["nodes"] if n.startswith("TEL_")]
                placas_alvo = [n.replace("PLACA_", "") for n in cluster_obj["nodes"] if n.startswith("PLACA_")]
                
                resumo_intel, df_matches_criacao, df_novos_suspeitos = cruzar_com_criacoes_diarias(cpfs_alvo, tels_alvo, placas_alvo)

                if resumo_intel is None:
                    st.info("Criações Diárias ainda não ingeridas.")
                elif resumo_intel["total_assistencias"] == 0:
                    st.success("Sem acionamentos adicionais nas Criações Diárias.")
                else:
                    c1, c2, c3, c4 = st.columns(4)
                    with c1: st.markdown(f"""<div class="intel-card"><span style="color:#94A3B8; font-size:10px;">ACIONAMENTOS</span><br/><b style="font-size:18px; color:#38BDF8;">{resumo_intel['total_assistencias']}</b></div>""", unsafe_allow_html=True)
                    with c2: st.markdown(f"""<div class="intel-card"><span style="color:#94A3B8; font-size:10px;">NOVOS TELS</span><br/><b style="font-size:18px; color:#EF4444;">+{resumo_intel['novos_tels']}</b></div>""", unsafe_allow_html=True)
                    with c3: st.markdown(f"""<div class="intel-card"><span style="color:#94A3B8; font-size:10px;">NOVAS PLACAS</span><br/><b style="font-size:18px; color:#A78BFA;">+{resumo_intel['novas_placas']}</b></div>""", unsafe_allow_html=True)
                    with c4: st.markdown(f"""<div class="intel-card"><span style="color:#94A3B8; font-size:10px;">NOVOS CPFS</span><br/><b style="font-size:18px; color:#FBBF24;">+{resumo_intel['novos_cpfs']}</b></div>""", unsafe_allow_html=True)

                st.markdown("<br/>", unsafe_allow_html=True)
                if not df_novos_suspeitos.empty:
                    st.markdown(f"**Novos Suspeitos Detectados na Rede:**")
                    st.dataframe(df_novos_suspeitos, use_container_width=True, hide_index=True)
                    
                    c_inc1, c_inc2 = st.columns([2.5, 2.5])
                    with c_inc1:
                        if st.button("Incorporar Novos Suspeitos e Anexar ao Caso", type="primary", use_container_width=True, key=f"btn_inc_{identificador_caso}_{sufixo_key}"):
                            anexar_descoberta_a_caso_existente(
                                df_matches_criacao,
                                identificador_caso,
                                analista=dados_caso_oficial.get("analista_responsavel", "") if dados_caso_oficial else "",
                                observacao="Incorporado via Expansão com Criações Diárias"
                            )
                            st.cache_data.clear()
                            st.success(f"Entidades e assistências vinculadas com sucesso ao Caso {identificador_caso}.")
                            st.rerun()

                    with c_inc2:
                        csv_susp = df_novos_suspeitos.to_csv(index=False).encode('utf-8')
                        st.download_button("Exportar Suspeitos (.csv)", csv_susp, file_name=f"Suspeitos_{identificador_caso}.csv", mime="text/csv", use_container_width=True, key=f"btn_dl_susp_{identificador_caso}_{sufixo_key}")
                else:
                    st.info("Acionamentos pertencem apenas a entidades já catalogadas.")

    if st.session_state["cfg_modo_exibicao"] == "Abas Clássicas":
        abas_titulos = ["Grafo de Vínculos", "Tabela de Ocorrências", "Radar Territorial", "Evolução Temporal"]
        if eh_oficial:
            abas_titulos.extend(["Expansão com Criações", "Gestão do Caso & Dossiê"])

        abas = st.tabs(abas_titulos)
        with abas[0]: render_modulo("Grafo", altura=740, sufixo_key="aba")
        with abas[1]: render_modulo("Tabela", sufixo_key="aba")
        with abas[2]: render_modulo("Radar", sufixo_key="aba")
        with abas[3]: render_modulo("Temporal", sufixo_key="aba")
        
        if eh_oficial:
            with abas[4]: render_modulo("Expansão", sufixo_key="aba")
            with abas[5]:
                st.markdown(f"#### Identificação e Parecer da Investigação ({identificador_caso})")
                c_f1, c_f2, c_f3 = st.columns([2.5, 1.5, 1.5])
                with c_f1: novo_nome = st.text_input("Nome da Quadrilha / Caso:", value=dados_caso_oficial["nome_personalizado"], placeholder="Ex: Quadrilha Santo André")
                with c_f2:
                    lista_status_caso = ["Em Investigação", "Confirmado Fraude", "Monitoramento Contínuo", "Sem Irregularidade Identificada", "Falso Positivo", "Arquivado"]
                    idx_status_atual = lista_status_caso.index(dados_caso_oficial["status"]) if dados_caso_oficial["status"] in lista_status_caso else 0
                    novo_status = st.selectbox("Status:", lista_status_caso, index=idx_status_atual)
                with c_f3: novo_analista = st.text_input("Analista:", value=dados_caso_oficial["analista_responsavel"])

                st.caption("Use referências como @CPF_..., @TEL_... ou @PLACA_... no parecer para destacá-las na Trilha de Auditoria.")
                novo_parecer = st.text_area("Parecer Técnico:", value=dados_caso_oficial["parecer"], height=140)
                if st.button("Salvar Metadados do Caso", type="primary", use_container_width=True):
                    salvar_dados_caso(identificador_caso, novo_nome, novo_status, novo_analista, novo_parecer)
                    st.success("Metadados gravados com sucesso.")
                    st.rerun()

                st.markdown("---")
                st.markdown("##### Trilha de Auditoria e Histórico de Pareceres")
                historico_notas = carregar_historico_pareceres(identificador_caso)
                if historico_notas:
                    for h in historico_notas:
                        st.markdown(f"""
                        <div style="background:#0E1726; border:1px solid #1E293B; border-radius:6px; padding:10px 14px; margin-bottom:8px;">
                            <span style="color:#38BDF8; font-size:12px; font-weight:600;">{h['data_registro']}</span> • 
                            <span style="color:#FBBF24; font-size:12px;">Analista: <b>{h['analista'] or 'Sistema'}</b></span> • 
                            <span style="color:#10B981; font-size:12px;">Status: <b>{h['status']}</b></span>
                            <div style="color:#E2E8F0; font-size:13px; margin-top:6px; white-space: pre-wrap;">{formatar_mencoes_forenses(h['parecer'])}</div>
                        </div>
                        """, unsafe_allow_html=True)
                else:
                    st.caption("Nenhum parecer técnico anterior registrado no histórico para este caso.")

                st.markdown("---")
                st.markdown("##### Isolamento de Terceiros (Ocultar do Grafo Deste Caso)")
                st.caption("Remove a entidade da visualização deste caso específico, preservando todos os dados brutos no banco.")

                col_oc1, col_oc2, col_oc3 = st.columns([2.5, 2.5, 1.5])
                with col_oc1: no_para_ocultar = st.selectbox("Entidade do caso:", cluster_obj["nodes"] if cluster_obj else [], key="sel_no_ocultar")
                with col_oc2: motivo_oc = st.text_input("Motivo:", placeholder="Ex: Terceiro sem relação com o conluio", key="motivo_ocultar")
                with col_oc3:
                    st.write("")
                    if st.button("Isolar Terceiro", use_container_width=True):
                        ocultar_no_do_caso(identificador_caso, no_para_ocultar, motivo_oc, dados_caso_oficial.get("analista_responsavel", ""))
                        st.cache_data.clear()
                        st.success(f"Entidade {no_para_ocultar} ocultada deste caso.")
                        st.rerun()

                ocultos = listar_nos_ocultos_do_caso(identificador_caso)
                if ocultos:
                    st.markdown("**Nós ocultados nesta célula:**")
                    for oc in ocultos:
                        c_r1, c_r2 = st.columns([4.0, 1.0])
                        with c_r1: st.caption(f"• **{oc['node_id']}** — {oc['motivo'] or 'Sem justificativa'} (Ocultado em: {oc['data_ocultacao']})")
                        with c_r2:
                            if st.button("Reativar Vínculo", key=f"restaurar_{oc['node_id']}", use_container_width=True):
                                restaurar_no_do_caso(identificador_caso, oc["node_id"])
                                st.cache_data.clear()
                                st.success("Nó restaurado ao grafo do caso.")
                                st.rerun()
    else:
        opcoes_painel_atual = OPCOES_PAINEL_OFICIAL if eh_oficial else OPCOES_PAINEL_DESCOBERTA

        c_pesq, c_prop, c_pdir = st.columns([2.5, 1.4, 2.5])
        with c_pesq:
            idx_pesq = opcoes_painel_atual.index(st.session_state["cfg_painel_esquerdo"]) if st.session_state["cfg_painel_esquerdo"] in opcoes_painel_atual else 0
            st.selectbox("Painel Esquerdo:", opcoes_painel_atual, index=idx_pesq, key="w_painel_esq", on_change=cb_atualizar_pesq)
        with c_prop:
            idx_prop = list(PROPORCOES_MAP.keys()).index(st.session_state["cfg_cockpit_proporcao"]) if st.session_state["cfg_cockpit_proporcao"] in PROPORCOES_MAP else 0
            st.selectbox("Proporção:", list(PROPORCOES_MAP.keys()), index=idx_prop, key="w_proporcao", on_change=cb_atualizar_prop)
        with c_pdir:
            idx_pdir = opcoes_painel_atual.index(st.session_state["cfg_painel_direito"]) if st.session_state["cfg_painel_direito"] in opcoes_painel_atual else (1 if len(opcoes_painel_atual) > 1 else 0)
            st.selectbox("Painel Direito:", opcoes_painel_atual, index=idx_pdir, key="w_painel_dir", on_change=cb_atualizar_pdir)

        col_left, col_right = st.columns(PROPORCOES_MAP[st.session_state["cfg_cockpit_proporcao"]])
        with col_left:
            st.markdown(f"<div style='margin-bottom:4px;'><b style='color:#38BDF8;'>{st.session_state['cfg_painel_esquerdo']}</b></div>", unsafe_allow_html=True)
            render_modulo(st.session_state["cfg_painel_esquerdo"], altura=680, sufixo_key="esq")
        with col_right:
            st.markdown(f"<div style='margin-bottom:4px;'><b style='color:#38BDF8;'>{st.session_state['cfg_painel_direito']}</b></div>", unsafe_allow_html=True)
            render_modulo(st.session_state["cfg_painel_direito"], altura=680, sufixo_key="dir")


# ====================================================================
# 7. ROTEAMENTO DE TELAS (DESCOBERTA vs. CASO OFICIAL vs. TRIAGEM GERAL)
# ====================================================================
if st.session_state["modo_descoberta_ativo"] and st.session_state["dados_descoberta"]:
    dados_d = st.session_state["dados_descoberta"]
    renderizar_mesa_investigacao(
        df_dados=dados_d["df"],
        identificador_caso="DESCOBERTA",
        nome_exibicao=dados_d["resumo"]["alvo_buscado"],
        hub_id="",
        eh_oficial=False,
        resumo_descoberta=dados_d["resumo"],
        cluster_obj=None
    )

elif st.session_state["celula_ativa_id"] is not None:
    cluster_selecionado = next((c for c in cluster_info if c["id"] == st.session_state["celula_ativa_id"]), cluster_info[0])
    dados_caso = carregar_dados_caso(cluster_selecionado["id"])
    nome_exib = dados_caso["nome_personalizado"] or f"Caso {cluster_selecionado['id']}"

    lista_cpfs = [n.replace("CPF_", "") for n in cluster_selecionado["nodes"] if n.startswith("CPF_")]
    lista_tels = [n.replace("TEL_", "") for n in cluster_selecionado["nodes"] if n.startswith("TEL_")]
    lista_placas = [n.replace("PLACA_", "") for n in cluster_selecionado["nodes"] if n.startswith("PLACA_")]
    df_caso = consultar_detalhes_caso(lista_cpfs, lista_tels, lista_placas)

    renderizar_mesa_investigacao(
        df_dados=df_caso,
        identificador_caso=cluster_selecionado["id"],
        nome_exibicao=nome_exib,
        hub_id=cluster_selecionado["hub_id"],
        eh_oficial=True,
        dados_caso_oficial=dados_caso,
        cluster_obj=cluster_selecionado
    )

# ====================================================================
# 8. TELA 1: PAINEL GERAL DE TRIAGEM
# ====================================================================
else:
    st.markdown("<h3 style='margin:0; color:#38BDF8;'>Mesa de Inteligência Forense - LCFO</h3>", unsafe_allow_html=True)
    st.caption("Painel Executivo de Triagem, Gestão de Quadrilhas e Catálogo Estratégico")
    st.markdown("<br/>", unsafe_allow_html=True)

    aba_triagem = st.radio(
        "Módulos Operacionais:",
        ["Células e Redes da Blacklist", "Base Mestra de Entidades Monitoradas", "Watchlist de Municípios de Risco", "Centro de Comando e Radar Territorial (Tela 4)"],
        horizontal=True,
        key="aba_principal_triagem"
    )

    st.markdown("<br/>", unsafe_allow_html=True)

    # SUB-ABA 1: CÉLULAS DA BLACKLIST
    if aba_triagem == "Células e Redes da Blacklist":
        k1, k2, k3, k4 = st.columns(4)
        with k1: st.markdown(f"""<div class="kpi-card"><span style="color:#94A3B8; font-size:12px; font-weight:600;">CÉLULAS MAPEADAS</span><br/><b style="font-size:26px; color:#F8FAFC;">{len(cluster_info):,}</b></div>""", unsafe_allow_html=True)
        with k2: st.markdown(f"""<div class="kpi-card"><span style="color:#94A3B8; font-size:12px; font-weight:600;">ENTIDADES NA BLACKLIST</span><br/><b style="font-size:26px; color:#38BDF8;">{len(G.nodes):,}</b></div>""", unsafe_allow_html=True)
        with k3: st.markdown(f"""<div class="kpi-card"><span style="color:#94A3B8; font-size:12px; font-weight:600;">QUADRILHAS NOMEADAS</span><br/><b style="font-size:26px; color:#10B981;">{len([c for c in casos_cadastrados.values() if c['nome']]):,}</b></div>""", unsafe_allow_html=True)
        with k4:
            qtd_com_alerta = len(radar_alertas)
            cor_alerta = "#EF4444" if qtd_com_alerta > 0 else "#10B981"
            st.markdown(f"""<div class="kpi-card"><span style="color:#94A3B8; font-size:12px; font-weight:600;">CÉLULAS REINCIDENTES</span><br/><b style="font-size:26px; color:{cor_alerta};">{qtd_com_alerta} no Radar</b></div>""", unsafe_allow_html=True)

        st.markdown("<br/>", unsafe_allow_html=True)

        if radar_alertas:
            st.markdown(f"##### Radar de Alertas: {len(radar_alertas)} Células Reincidentes nas Criações Diárias")
            cols_radar = st.columns(min(4, len(radar_alertas)))
            for i, (cid, total_hits) in enumerate(list(radar_alertas.items())[:4]):
                with cols_radar[i]:
                    c_obj = next((c for c in cluster_info if c["id"] == cid), None)
                    if c_obj:
                        nome_q = casos_cadastrados.get(cid, {}).get("nome", c_obj['hub_label'][:20])
                        if st.button(f"{nome_q}\n+{total_hits} ocorrências ({c_obj['id']})", use_container_width=True):
                            st.session_state["celula_ativa_id"] = cid
                            st.session_state["coordenada_foco"] = None
                            st.session_state["entidade_foco_grafo"] = None
                            st.rerun()

        st.markdown("---")
        st.markdown("#### Fila de Investigação Operacional Priorizada por Risco")
        
        col_sel1, col_sel2 = st.columns([3.5, 1.5])
        with col_sel1: caso_escolhido = st.selectbox("Selecione um caso para abrir a mesa de análise:", list(opcoes_celulas.keys()))
        with col_sel2:
            st.write("")
            st.write("")
            if st.button("Abrir Mesa Forense", use_container_width=True, type="primary"):
                st.session_state["celula_ativa_id"] = opcoes_celulas[caso_escolhido]
                st.session_state["coordenada_foco"] = None
                st.session_state["entidade_foco_grafo"] = None
                st.rerun()

        mapa_scores = obter_scores_triagem(cluster_info)
        tabela_casos = []
        for c in cluster_info[:100]:
            info_c = casos_cadastrados.get(c["id"], {})
            status_radar = f"+{radar_alertas[c['id']]} novas" if c["id"] in radar_alertas else "Estável"
            dados_score = mapa_scores.get(c["id"], {"score": 0, "nivel": "BAIXO", "cor": "#10B981"})

            tabela_casos.append({
                "Código": c["id"],
                "Score": dados_score["score"],
                "Risco": dados_score["nivel"],
                "Nome da Quadrilha": info_c.get("nome", "Não Batizado"),
                "Status": info_c.get("status", "Em Investigação"),
                "Âncora Central (Hub)": c["hub_label"],
                "Tamanho": c["tamanho"],
                "Telefones": c["qtd_tels"],
                "Titulares": c["qtd_cpfs"],
                "Placas": c["qtd_placas"],
                "Status Radar": status_radar
            })

        df_resumo_casos = pd.DataFrame(tabela_casos).sort_values(by=["Score", "Tamanho"], ascending=[False, False])
        st.dataframe(df_resumo_casos, use_container_width=True, hide_index=True)
        st.caption("Score preliminar sem topologia profunda de subgrafo — abra o caso para o Fraud Score definitivo.")

    # SUB-ABA 2: BASE MESTRA DE ENTIDADES
    elif aba_triagem == "Base Mestra de Entidades Monitoradas":
        c_m_top1, c_m_top2 = st.columns([3.5, 1.5])
        with c_m_top1:
            st.markdown("#### Base Mestra de Entidades Monitoradas")
            st.caption("Repositório permanente de suspeitos. Varre retrospectivamente todo o acervo de Criações Diárias.")
        with c_m_top2:
            if st.button("Sincronizar com a Blacklist", use_container_width=True):
                with st.spinner("Importando entidades da Blacklist..."):
                    novos_sem = semear_base_mestra_da_blacklist()
                    st.success(f"{novos_sem} novas entidades catalogadas na Base Mestra.")
                    st.rerun()

        with st.expander("Cadastrar Nova Entidade Suspeita", expanded=False):
            with st.form("form_novo_suspeito", clear_on_submit=True):
                c_t1, c_t2, c_t3 = st.columns([1.5, 2.5, 2.5])
                with c_t1: novo_tipo = st.selectbox("Tipo de Entidade:", ["TELEFONE", "PLACA", "CPF"])
                with c_t2: novo_val = st.text_input("Dado / Número / Placa:", placeholder="Ex: (11) 97000-1122 ou ABC-1234")
                with c_t3: novo_nome = st.text_input("Nome / Apelido / Titular:", placeholder="Ex: Marcos V. (Laranja)")

                c_m1, c_m2, c_m3 = st.columns([2.0, 2.0, 2.0])
                with c_m1: novo_caso = st.text_input("Quadrilha / Operação Associada:", placeholder="Ex: Quadrilha Santo André")
                with c_m2: novo_st = st.selectbox("Status:", ["Ativo", "Confirmado Fraude", "Em Monitoramento", "Inativo"])
                with c_m3: novo_analista = st.text_input("Analista Responsável:", placeholder="Seu nome")

                novo_motivo = st.text_area("Motivo da Inclusão / Modus Operandi:", placeholder="Ex: Solicitou guinchos sequenciais para o mesmo destino...")

                btn_salvar_suspeito = st.form_submit_button("Salvar na Base Mestra & Varrer Histórico", type="primary", use_container_width=True)
                if btn_salvar_suspeito:
                    ok, msg, total_hits = cadastrar_entidade_suspeita(novo_tipo, novo_val, novo_nome, novo_motivo, novo_caso, status=novo_st, analista=novo_analista)
                    if ok:
                        st.success(f"{msg} Varredura retrospectiva: {total_hits} assistências identificadas.")
                        st.rerun()
                    else:
                        st.error(msg)

        st.markdown("---")
        f_c1, f_c2, f_c3 = st.columns([1.5, 1.5, 3.0])
        with f_c1: f_tipo = st.selectbox("Filtrar Tipo:", ["TODOS", "TELEFONE", "PLACA", "CPF"])
        with f_c2: f_status = st.selectbox("Filtrar Status:", ["TODOS", "Ativo", "Confirmado Fraude", "Em Monitoramento", "Inativo", "Removido"])
        with f_c3: f_busca = st.text_input("Buscar na Base Mestra:", placeholder="Digite dado, nome ou quadrilha...")

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
                    "Histórico Total": f"{s['total_historico']} assistências" if s.get('total_historico', 0) > 0 else "Nenhum",
                    "Data Cadastro": s["data_cadastro"],
                    "Analista": s["analista"] or "-"
                })
            df_mestra = pd.DataFrame(tabela_mestra)
            st.dataframe(df_mestra, use_container_width=True, hide_index=True)

            with st.expander("Auditar / Remover Entidade Específica"):
                opcoes_ent = {f"[{s['tipo']}] {s['valor_formatado']} - {s['nome_referencia']} ({s['status']})": s for s in suspeitos}
                sel_ent_label = st.selectbox("Selecione a entidade para auditoria:", list(opcoes_ent.keys()))
                ent_alvo = opcoes_ent[sel_ent_label]

                col_btn1, col_btn2, col_btn3 = st.columns([3.0, 2.0, 1.0])
                with col_btn2: motivo_rem = st.text_input("Motivo da remoção / arquivamento:", key=f"motivo_rem_{ent_alvo['id']}", placeholder="Ex: Terceiro inocente comprovado")
                with col_btn3:
                    st.write("")
                    if st.button(f"Excluir #{ent_alvo['id']}", use_container_width=True):
                        remover_entidade_suspeita(ent_alvo["id"], motivo_rem)
                        st.success("Entidade marcada como 'Removido' (cadeia de custódia preservada).")
                        st.rerun()

                df_ocorr = consultar_todas_ocorrencias_entidade(ent_alvo["tipo"], ent_alvo["valor"])
                if not df_ocorr.empty:
                    st.markdown(f"**{len(df_ocorr)} Assistências Encontradas (Blacklist + Criações Diárias):**")
                    st.dataframe(df_ocorr, use_container_width=True, hide_index=True)
                    csv_ent = df_ocorr.to_csv(index=False).encode('utf-8')
                    st.download_button(f"Baixar Ocorrências de {ent_alvo['valor_formatado']} (.csv)", csv_ent, file_name=f"Historico_{ent_alvo['valor']}.csv", mime="text/csv")
                else:
                    st.info("Nenhuma assistência registrada nos arquivos ingeridos para esta entidade.")
        else:
            st.info("Nenhuma entidade cadastrada. Sincronize com a Blacklist ou cadastre manualmente acima.")

    # SUB-ABA 3: WATCHLIST DE MUNICÍPIOS DE RISCO
    elif aba_triagem == "Watchlist de Municípios de Risco":
        c_w_top1, c_w_top2 = st.columns([3.5, 1.5])
        with c_w_top1:
            st.markdown("#### Watchlist de Municípios de Alto Risco")
            st.caption("Cadastro de praças com predominância de fraude. Alimenta multiplicadores do Fraud Score.")

        with st.expander("Cadastrar Novo Município na Watchlist", expanded=False):
            with st.form("form_nova_cidade_risco", clear_on_submit=True):
                col_c1, col_c2 = st.columns([3.0, 1.2])
                with col_c1: nova_cidade = st.text_input("Nome do Município:", placeholder="Ex: Santa Quitéria")
                with col_c2: novo_uf = st.selectbox("UF:", ["AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO"], index=5)
                
                col_m1, col_m2 = st.columns([3.0, 2.0])
                with col_m1: novo_motivo_cid = st.text_input("Motivo / Padrão de Fraude Mapeado:", placeholder="Ex: Frequência atípica de guinchos e colusão regional")
                with col_m2: novo_analista_cid = st.text_input("Analista Responsável:", placeholder="Seu nome", key="analista_cidade_input")
                
                btn_salvar_cidade = st.form_submit_button("Salvar Município na Watchlist", type="primary", use_container_width=True)
                if btn_salvar_cidade:
                    ok, msg = cadastrar_cidade_risco(nova_cidade, novo_uf, novo_motivo_cid, novo_analista_cid)
                    if ok:
                        st.cache_data.clear()
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)
        
        st.markdown("---")
        cidades_cadastradas = listar_cidades_risco()
        if cidades_cadastradas:
            st.caption(f"Exibindo **{len(cidades_cadastradas)}** municípios monitorados na Watchlist.")
            tabela_cidades = []
            for c_item in cidades_cadastradas:
                tabela_cidades.append({
                    "ID": c_item["id"],
                    "Município": c_item["cidade"],
                    "UF": c_item["uf"],
                    "Motivo / Modus Operandi": c_item["motivo"] or "Sem observação",
                    "Analista": c_item["analista"] or "Sistema",
                    "Data Cadastro": c_item["data_cadastro"]
                })
            df_cidades_view = pd.DataFrame(tabela_cidades)
            st.dataframe(df_cidades_view, use_container_width=True, hide_index=True)

            with st.expander("Excluir Município da Watchlist"):
                opcoes_del_cid = {f"#{c['id']} - {c['cidade']}/{c['uf']} ({c['motivo'][:35]}...)": c['id'] for c in cidades_cadastradas}
                cid_sel_label = st.selectbox("Selecione o município para remover:", list(opcoes_del_cid.keys()))
                id_cid_del = opcoes_del_cid[cid_sel_label]
                if st.button("Remover Município Selecionado", use_container_width=True):
                    remover_cidade_risco(id_cid_del)
                    st.cache_data.clear()
                    st.success("Município removido da Watchlist de Risco.")
                    st.rerun()
        else:
            st.info("Nenhum município cadastrado na Watchlist de Risco até o momento.")

    # SUB-ABA 4: CENTRO DE COMANDO PROATIVO & FUNIL DA TELA 4
    else:
        st.markdown("#### Centro de Comando Proativo e Radar de Anomalias Territoriais")
        st.caption("Detecção estatística de desvios volumétricos municipais, anomalias pet e abusos residenciais.")

        with st.spinner("Calculando desvios volumétricos e baselines históricos..."):
            df_anomalias, kpis_anomalias, mes_referencia = obter_radar_anomalias_macro()

        if df_anomalias.empty:
            st.info("É necessário ingerir arquivos em Criações Diárias para calcular os baselines territoriais.")
        else:
            c_ano1, c_ano2, c_ano3, c_ano4 = st.columns(4)
            with c_ano1: st.markdown(f"""<div class="kpi-card"><span style="color:#94A3B8; font-size:12px; font-weight:600;">MUNICÍPIOS COM ANOMALIA</span><br/><b style="font-size:24px; color:#EF4444;">{kpis_anomalias.get('total_anomalias', 0)}</b><br/><span style="font-size:10px; color:#94A3B8;">Mês ref: {mes_referencia}</span></div>""", unsafe_allow_html=True)
            with c_ano2: st.markdown(f"""<div class="kpi-card"><span style="color:#94A3B8; font-size:12px; font-weight:600;">ALERTAS PET ATIVOS</span><br/><b style="font-size:24px; color:#FBBF24;">{kpis_anomalias.get('alertas_pet', 0)}</b><br/><span style="font-size:10px; color:#94A3B8;">Saltos atípicos em clínicas</span></div>""", unsafe_allow_html=True)
            with c_ano3: st.markdown(f"""<div class="kpi-card"><span style="color:#94A3B8; font-size:12px; font-weight:600;">SALTOS RESIDENCIAIS</span><br/><b style="font-size:24px; color:#38BDF8;">{kpis_anomalias.get('alertas_res', 0)}</b><br/><span style="font-size:10px; color:#94A3B8;">Eletricista / Encanador</span></div>""", unsafe_allow_html=True)
            with c_ano4: st.markdown(f"""<div class="kpi-card"><span style="color:#94A3B8; font-size:12px; font-weight:600;">PRAÇAS NA WATCHLIST</span><br/><b style="font-size:24px; color:#A78BFA;">{kpis_anomalias.get('alertas_watchlist', 0)}</b><br/><span style="font-size:10px; color:#94A3B8;">Municípios monitorados</span></div>""", unsafe_allow_html=True)

            st.markdown("<br/>", unsafe_allow_html=True)
            tab_radar_mapa, tab_radar_tabela = st.tabs(["Radar Territorial de Anomalias", "Fila de Anomalias por Município"])

            with tab_radar_mapa:
                df_mapa_ano = df_anomalias.dropna(subset=["latitude", "longitude"]).copy()
                if not df_mapa_ano.empty:
                    df_mapa_ano["raio_calc"] = df_mapa_ano["volume_atual"] * 400
                    camada_anomalias = pdk.Layer(
                        "ScatterplotLayer",
                        data=df_mapa_ano,
                        get_position="[longitude, latitude]",
                        get_color="[239, 68, 68, 190]",
                        get_line_color="[255, 255, 255, 230]",
                        line_width_min_pixels=1.5,
                        stroked=True,
                        get_radius="raio_calc",
                        radius_min_pixels=7,
                        radius_max_pixels=32,
                        pickable=True,
                        auto_highlight=True
                    )
                    lat_c = float(df_mapa_ano["latitude"].mean())
                    lon_c = float(df_mapa_ano["longitude"].mean())
                    view_state_ano = pdk.ViewState(latitude=lat_c, longitude=lon_c, zoom=4.2, pitch=0)

                    tooltip_ano = {
                        "html": "<div style='font-family:sans-serif; padding:5px;'><b style='color:#EF4444; font-size:13px;'>{cidade}/{uf}</b><br/><b>Volume Atual:</b> {volume_atual} acionamentos<br/><b>Média Histórica:</b> {media_hist}<br/><b>Desvio:</b> {desvio_macro}<br/><b>Padrões:</b> <span style='color:#FBBF24;'>{alertas_str}</span></div>",
                        "style": {"backgroundColor": "rgba(11, 17, 30, 0.95)", "color": "#F8FAFC", "border": "1px solid #1E293B", "borderRadius": "6px", "fontSize": "11px", "zIndex": "1000"}
                    }
                    deck_ano = pdk.Deck(layers=[camada_anomalias], initial_view_state=view_state_ano, tooltip=tooltip_ano, map_style="dark")
                    st.pydeck_chart(deck_ano, use_container_width=True)
                else:
                    st.info("Sem coordenadas disponíveis para plotagem das anomalias.")

            with tab_radar_tabela:
                col_f1, col_f2 = st.columns([2.5, 3.5])
                with col_f1:
                    filtro_tipo_ano = st.selectbox("Filtrar Categoria de Alerta:", ["TODOS", "Explosão de Volume Macro", "Anomalia em Serviços Pet", "Salto em Serviços Residenciais", "Município em Watchlist de Risco", "Pico sem Histórico Prévio"])
                with col_f2:
                    busca_cid_ano = st.text_input("Buscar Município ou UF:", placeholder="Digite cidade ou estado...")

                df_view_ano = df_anomalias.copy()
                if filtro_tipo_ano != "TODOS":
                    df_view_ano = df_view_ano[df_view_ano["alertas"].apply(lambda lista: filtro_tipo_ano in lista)]
                if busca_cid_ano:
                    b_norm = busca_cid_ano.upper().strip()
                    df_view_ano = df_view_ano[df_view_ano["cidade"].str.upper().str.contains(b_norm) | df_view_ano["uf"].str.upper().str.contains(b_norm)]

                tabela_exibicao = []
                for r_ano in df_view_ano.itertuples():
                    tabela_exibicao.append({
                        "Município": r_ano.cidade,
                        "UF": r_ano.uf,
                        "Mês Ref.": r_ano.mes_ref,
                        "Volume Atual": f"{r_ano.volume_atual} acionamentos",
                        "Média Histórica": f"{r_ano.media_hist}",
                        "Desvio Relativo": r_ano.desvio_macro,
                        "Meses Ativos": r_ano.meses_ativos_hist,
                        "Pet (Mês)": r_ano.pet_atual,
                        "Residencial (Mês)": r_ano.res_atual,
                        "Padrões Identificados": r_ano.alertas_str,
                        "Watchlist": r_ano.watchlist
                    })
                df_tabela_final = pd.DataFrame(tabela_exibicao)

                linhas_ano_sel = []
                try:
                    evento_tab_ano = st.dataframe(
                        df_tabela_final,
                        use_container_width=True,
                        hide_index=True,
                        on_select="rerun",
                        selection_mode="single-row",
                        key="tab_anomalias_macro_view"
                    )
                    linhas_ano_sel = extrair_linhas_selecionadas(evento_tab_ano)
                except Exception:
                    st.dataframe(df_tabela_final, use_container_width=True, hide_index=True)

                if linhas_ano_sel:
                    idx_ano = linhas_ano_sel[0]
                    if 0 <= idx_ano < len(df_view_ano):
                        cid_sel_tabela = df_view_ano.iloc[idx_ano]["cidade"]
                        if st.session_state["municipio_foco_funil"] != cid_sel_tabela:
                            st.session_state["municipio_foco_funil"] = cid_sel_tabela

            # FUNIL DA TELA 4: EXTRAÇÃO CIRÚRGICA DE INFRATORES
            st.markdown("---")
            st.markdown("##### Desdobramento Investigativo da Anomalia Municipal")
            st.caption("Isole um município alertado para consolidar os principais operadores e autuar um caso de conluio regional.")

            cidades_opcoes = {f"{r['cidade']}/{r['uf']} ({r['volume_atual']} acionamentos - {r['desvio_macro']})": r for _, r in df_anomalias.iterrows()}
            if cidades_opcoes:
                lista_labels_cidades = list(cidades_opcoes.keys())
                idx_default_cidade = 0
                if st.session_state["municipio_foco_funil"]:
                    for i_lbl, lbl in enumerate(lista_labels_cidades):
                        if lbl.startswith(st.session_state["municipio_foco_funil"] + "/"):
                            idx_default_cidade = i_lbl
                            break

                c_sel_label = st.selectbox("Selecione a Praça em Alerta:", lista_labels_cidades, index=idx_default_cidade)
                anomalia_alvo = cidades_opcoes[c_sel_label]

                dados_infratores = extrair_top_infratores_municipio(
                    cidade=anomalia_alvo["cidade"],
                    uf=anomalia_alvo["uf"],
                    cidade_banco=anomalia_alvo.get("cidade_banco", "")
                )

                if dados_infratores["total_ocorrencias"] > 0:
                    c_inf1, c_inf2, c_inf3 = st.columns(3)
                    with c_inf1:
                        st.markdown(f"**Top Titulares ([CPF]):**")
                        st.dataframe(dados_infratores["top_cpfs"], use_container_width=True, hide_index=True)
                    with c_inf2:
                        st.markdown(f"**Top Telefones de Contato ([TEL]):**")
                        st.dataframe(dados_infratores["top_tels"], use_container_width=True, hide_index=True)
                    with c_inf3:
                        st.markdown(f"**Top Veículos Envolvidos ([PLACA]):**")
                        st.dataframe(dados_infratores["top_placas"], use_container_width=True, hide_index=True)

                    if st.button("Autuar Caso a partir Desta Anomalia Territorial", type="primary", use_container_width=True):
                        st.session_state["modo_descoberta_ativo"] = True
                        st.session_state["dados_descoberta"] = {
                            "resumo": {
                                "alvo_buscado": f"Conluio Regional - {anomalia_alvo['cidade']}/{anomalia_alvo['uf']}",
                                "termo_limpo": anomalia_alvo["cidade"],
                                "total_assistencias": dados_infratores["total_ocorrencias"],
                                "assistencias_diretas": dados_infratores["total_ocorrencias"],
                                "qtd_cpfs": len(dados_infratores["top_cpfs"]),
                                "qtd_tels": len(dados_infratores["top_tels"]),
                                "qtd_placas": len(dados_infratores["top_placas"])
                            },
                            "df": dados_infratores["df_completo"],
                            "entidades": {
                                "cpfs": list(dados_infratores["top_cpfs"]["cpf"]) if not dados_infratores["top_cpfs"].empty else [],
                                "tels": list(dados_infratores["top_tels"]["telefone"]) if not dados_infratores["top_tels"].empty else [],
                                "placas": list(dados_infratores["top_placas"]["placa"]) if not dados_infratores["top_placas"].empty else []
                            }
                        }
                        st.session_state["coordenada_foco"] = None
                        st.session_state["entidade_foco_grafo"] = None
                        st.rerun()
                else:
                    st.info("Nenhuma ocorrência detalhada localizada para os parâmetros deste município.")