"""
Módulo: appv3.py
Objetivo: Interface executiva da Mesa de Inteligência Forense de Fraudes em Assistências (LCFO).
          Navegação em três telas (Dashboard de Casos -> Overview do Caso -> Mesa de
          Investigação), barra lateral global com ícones, painel de Entidades com busca
          e adição manual (via Enrich), barra fixa compacta na Mesa (voltar/alternar caso,
          busca de entidade, toggle de Notas/Análises), paleta neutra e dashboard compacto
          com filtros em chips (st.pills) e cards clicáveis pelo título.
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
    carregar_redes, processar_subgrafo_caso, gerar_html_grafo, processar_grafo_dataframe,
    _criar_svg_nodo_flowsint
)
from correlation_engine import (
    calcular_score_caso, obter_scores_triagem
)
from anomalias_engine import (
    obter_radar_anomalias_macro, extrair_top_infratores_municipio
)
from utils import formatar_mencoes_forenses, formatar_cpf_cnpj, formatar_tel

try:
    from grafo_component import renderizar_grafo_bidirecional
    TEM_COMPONENTE_BIDIRECIONAL = True
except ImportError:
    TEM_COMPONENTE_BIDIRECIONAL = False

try:
    from enrich_engine import enriquecer_entidade_local, CORES_POR_TIPO, PREFIXO_POR_TIPO
    TEM_ENRICH_ENGINE = True
except ImportError:
    TEM_ENRICH_ENGINE = False

try:
    from exif_engine import extrair_metadados_foto, validar_coerencia_geografica_foto
    TEM_EXIF_ENGINE = True
except ImportError:
    TEM_EXIF_ENGINE = False

# =====================================================
# 1. CONFIGURAÇÃO DE PÁGINA & CSS (PALETA NEUTRA)
# =====================================================
st.set_page_config(
    page_title="LCFO",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .stDeployButton, .stAppDeployButton, #MainMenu, footer, [data-testid="stDecoration"] {
        display: none !important;
        visibility: hidden !important;
    }
    header[data-testid="stHeader"] { background: transparent !important; }
    [data-testid="stSidebar"] { overflow-y: auto !important; }
    [data-testid="stSidebarContent"] {
        overflow-y: auto !important; max-height: 100vh !important; padding-bottom: 2rem !important;
    }
    [data-testid="stSidebarCollapseButton"] button,
    [data-testid="collapsedControl"] button,
    [data-testid="stSidebarCollapsedControl"] button {
        background: #12161F !important; border: 1px solid #262B36 !important;
        color: #94A3B8 !important; border-radius: 6px !important;
    }
    .block-container {
        padding-top: 1.6rem !important; padding-bottom: 0.8rem !important;
        padding-left: 1.2rem !important; padding-right: 1.2rem !important; max-width: 100% !important;
    }

    .hud-bar {
        background: #10141C; border: 1px solid #262B36; border-radius: 8px;
        padding: 8px 14px; margin-top: 4px; margin-bottom: 10px;
        display: flex; align-items: center; justify-content: space-between;
    }
    .intel-card {
        background: #12161F; border: 1px solid #262B36; border-radius: 6px;
        padding: 10px 14px; text-align: center;
    }
    .kpi-card-mini {
        background: #10141C; border: 1px solid #262B36; border-radius: 8px;
        padding: 6px 10px; text-align: left;
    }
    .case-card {
        background: #12161F; border: 1px solid #262B36; border-radius: 10px;
        padding: 10px 14px; height: 100%;
    }
    .case-card-top { display: flex; align-items: center; gap: 6px; margin-bottom: 6px; flex-wrap: wrap; }
    .case-card-sub {
        font-size: 11px; color: #7C8798; margin-bottom: 8px;
        overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
    }
    .case-card-stats { font-size: 11px; color: #A7B0BE; display: flex; gap: 10px; flex-wrap: wrap; }
    .badge-pill {
        font-size: 10px; font-weight: 600; padding: 2px 8px; border-radius: 999px;
        text-transform: uppercase; white-space: nowrap;
    }
    .overview-header-card {
        background: #12161F; border: 1px solid #262B36; border-radius: 10px;
        padding: 16px 18px; margin-bottom: 14px;
    }
    .parecer-card {
        background: #12161F; border: 1px solid #262B36; border-radius: 8px;
        padding: 10px 12px; margin-bottom: 8px;
    }
    .entidade-item-caption { font-size: 10px; color: #7C8798; }

    /* Botão de título dos cards do dashboard -- vira o "card clicável" */
    div[class*="st-key-casecard_"] button {
        background: transparent !important; border: none !important; text-align: left !important;
        font-size: 14px !important; font-weight: 700 !important; color: #E5E9F0 !important;
        padding: 0 !important; justify-content: flex-start !important; box-shadow: none !important;
    }
    div[class*="st-key-casecard_"] button:hover {
        color: #8FB4CE !important; background: transparent !important;
    }
    div[class*="st-key-casecard_"] button:focus:not(:active) {
        color: #8FB4CE !important; box-shadow: none !important;
    }
</style>
""", unsafe_allow_html=True)

# =====================================================
# 2. DEFINIÇÕES GLOBAIS E ESTADO PERSISTENTE
# =====================================================
OPCOES_TOOLBAR_OFICIAL = [
    "Grafo de Vínculos", "Tabela de Ocorrências", "Radar Territorial (Mapa)",
    "Evolução Temporal", "Expansão com Criações", "Ferramentas"
]
OPCOES_TOOLBAR_DESCOBERTA = [
    "Grafo de Vínculos", "Tabela de Ocorrências", "Radar Territorial (Mapa)", "Evolução Temporal"
]
ICONE_TOOLBAR = {
    "Grafo de Vínculos": ":material/hub: Grafo",
    "Tabela de Ocorrências": ":material/table_chart: Tabela",
    "Radar Territorial (Mapa)": ":material/map: Mapa",
    "Evolução Temporal": ":material/timeline: Temporal",
    "Expansão com Criações": ":material/share: Relação",
    "Ferramentas": ":material/settings: Ferramentas"
}

OPCOES_PAINEL_OFICIAL = [
    "Radar Territorial (Mapa)", "Grafo de Vínculos", "Tabela de Ocorrências",
    "Evolução Temporal", "Expansão com Criações"
]
OPCOES_PAINEL_DESCOBERTA = [
    "Radar Territorial (Mapa)", "Grafo de Vínculos", "Tabela de Ocorrências", "Evolução Temporal"
]

PROPORCOES_MAP = {
    "50% | 50%": [1.0, 1.0], "60% | 40%": [1.5, 1.0], "40% | 60%": [1.0, 1.5],
    "70% | 30%": [2.3, 1.0], "30% | 70%": [1.0, 2.3]
}

FONT_PADRAO = 11
ESPACAMENTO_PADRAO = 280
LAYOUT_INICIAL_PADRAO = "organico"

LIMITE_ENTIDADES_ENRIQUECIDAS_POR_CASO = 400
TOLERANCIA_KM_EXIF = 60.0
LIMITE_ENTIDADES_LATERAL_SEM_BUSCA = 60
LIMITE_ENTIDADES_LATERAL_COM_BUSCA = 100

STATUS_ATIVOS = {"Em Investigação", "Confirmado Fraude", "Monitoramento Contínuo"}
STATUS_ENCERRADOS = {"Sem Irregularidade Identificada", "Falso Positivo", "Arquivado"}

# Paleta neutra -- tons dessaturados, mais próximos do Flowsint que dos azuis/vermelhos vivos anteriores.
CORES_STATUS = {
    "Em Investigação": "#6C93B0",
    "Confirmado Fraude": "#C0625F",
    "Monitoramento Contínuo": "#8F86B5",
    "Sem Irregularidade Identificada": "#6FA98A",
    "Falso Positivo": "#8B95A5",
    "Arquivado": "#5B6472",
}
CORES_RISCO_NEUTRAS = {
    "BAIXO": "#6FA98A", "MÉDIO": "#C9A66B", "ALTO": "#C98756", "CRÍTICO": "#C0625F",
}

def obter_cor_risco(nivel: str, cor_original: str) -> str:
    """
    O nível/score vêm de correlation_engine.py, mas a cor exibida é normalizada
    aqui para a paleta neutra -- sem precisar tocar no motor de cálculo.
    """
    return CORES_RISCO_NEUTRAS.get(nivel, cor_original)


NAV_PRINCIPAL = [
    (":material/folder_open: Casos", "Células e Redes da Blacklist"),
    (":material/hub: Base Mestra", "Base Mestra de Entidades Monitoradas"),
    (":material/warning: Watchlist", "Watchlist de Municípios de Risco"),
    (":material/dashboard: Centro de Comando", "Centro de Comando e Radar Territorial (Tela 4)"),
]

TIPO_LABEL_PARA_INTERNO = {"Titular (CPF)": "cpf", "Telefone": "telefone", "Placa": "placa"}
ICONE_POR_TIPO = {"cpf": ":material/person:", "telefone": ":material/call:", "placa": ":material/directions_car:", "empresa": ":material/apartment:"}

if "celula_ativa_id" not in st.session_state:
    st.session_state["celula_ativa_id"] = None
if "modo_descoberta_ativo" not in st.session_state:
    st.session_state["modo_descoberta_ativo"] = False
    st.session_state["dados_descoberta"] = None
if "subtela_caso" not in st.session_state:
    st.session_state["subtela_caso"] = "overview"
if "aba_principal_triagem" not in st.session_state:
    st.session_state["aba_principal_triagem"] = "Células e Redes da Blacklist"
if "cfg_modo_exibicao" not in st.session_state:
    st.session_state["cfg_modo_exibicao"] = "Abas Clássicas"
if "cfg_painel_unico" not in st.session_state:
    st.session_state["cfg_painel_unico"] = "Grafo de Vínculos"
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
if "nos_enriquecidos_por_caso" not in st.session_state:
    st.session_state["nos_enriquecidos_por_caso"] = {}
if "filtro_status_dash" not in st.session_state:
    st.session_state["filtro_status_dash"] = "Todos"
if "filtro_risco_dash" not in st.session_state:
    st.session_state["filtro_risco_dash"] = "Todos"
if "modo_visualizacao_casos" not in st.session_state:
    st.session_state["modo_visualizacao_casos"] = "Grade"
if "mostrar_form_nova_analise" not in st.session_state:
    st.session_state["mostrar_form_nova_analise"] = False
if "mostrar_notas_grafo" not in st.session_state:
    st.session_state["mostrar_notas_grafo"] = False
if "aba_painel_entidades" not in st.session_state:
    st.session_state["aba_painel_entidades"] = "Entities"


def cb_atualizar_modo():
    st.session_state["cfg_modo_exibicao"] = st.session_state["w_modo_exibicao"]

def cb_atualizar_pesq():
    st.session_state["cfg_painel_esquerdo"] = st.session_state["w_painel_esq"]

def cb_atualizar_pdir():
    st.session_state["cfg_painel_direito"] = st.session_state["w_painel_dir"]

def cb_atualizar_prop():
    st.session_state["cfg_cockpit_proporcao"] = st.session_state["w_proporcao"]


def abrir_caso_em_overview(id_caso: str):
    st.session_state["celula_ativa_id"] = id_caso
    st.session_state["modo_descoberta_ativo"] = False
    st.session_state["dados_descoberta"] = None
    st.session_state["subtela_caso"] = "overview"
    st.session_state["coordenada_foco"] = None
    st.session_state["entidade_foco_grafo"] = None
    st.session_state["mostrar_form_nova_analise"] = False
    st.rerun()


def abrir_descoberta_ativa(resumo: Dict[str, Any], df: pd.DataFrame, ents: Dict[str, Any]):
    st.session_state["modo_descoberta_ativo"] = True
    st.session_state["dados_descoberta"] = {"resumo": resumo, "df": df, "entidades": ents}
    st.session_state["celula_ativa_id"] = None
    st.session_state["coordenada_foco"] = None
    st.session_state["entidade_foco_grafo"] = None
    st.rerun()


def voltar_para_dashboard():
    st.session_state["celula_ativa_id"] = None
    st.session_state["modo_descoberta_ativo"] = False
    st.session_state["dados_descoberta"] = None
    st.session_state["coordenada_foco"] = None
    st.session_state["entidade_foco_grafo"] = None
    st.session_state["mostrar_form_nova_analise"] = False
    st.rerun()


def extrair_linhas_selecionadas(evento: Any) -> List[int]:
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
    return (
        f'<div class="hud-bar" style="border-color: {cor};">'
        f'<div style="display:flex; align-items:center; gap:18px; flex-wrap:wrap;">'
        f'<span style="font-size:0.95rem; font-weight:700; color:#E5E9F0;">{nome_exibicao.upper()}</span>'
        f'<span style="font-size:0.82rem; color:#7C8798;">Fraud Score: <b style="color:{cor};">{score}/100 [{nivel}]</b></span>'
        f'{campos_extra_html}'
        f'<span style="font-size:0.82rem; color:#6C93B0;">Acionamentos: <b>{total_registros} registros</b></span>'
        f'<span style="font-size:0.82rem; color:#7C8798;">Período: <b style="color:#C7CDD6;">{periodo_str}</b></span>'
        f'</div>'
        f'</div>'
    )


def montar_caixa_fatores(cor: str, fatores: List[str]) -> str:
    texto_fatores = ' • '.join(fatores) if fatores else "Comportamento relacional estável dentro da normalidade"
    return (
        f'<div style="background: #10141C; border-left: 4px solid {cor}; border-radius: 6px; '
        f'padding: 8px 14px; margin-bottom: 12px; border-top: 1px solid #262B36; border-right: 1px solid #262B36; '
        f'border-bottom: 1px solid #262B36;">'
        f'<div style="display: flex; align-items: center; justify-content: space-between;">'
        f'<span style="color: {cor}; font-weight: 700; font-size: 13px;">INDICADORES FORENSES IDENTIFICADOS</span>'
        f'<span style="color: #7C8798; font-size: 11px;">Avaliação Comportamental & Topológica</span>'
        f'</div>'
        f'<div style="color: #C7CDD6; font-size: 12px; margin-top: 4px;">{texto_fatores}</div>'
        f'</div>'
    )


def obter_enrich_do_caso(identificador_caso: str) -> Dict[str, Dict[str, Any]]:
    baldes = st.session_state["nos_enriquecidos_por_caso"]
    if identificador_caso not in baldes:
        baldes[identificador_caso] = {"nodes": {}, "edges": {}}
    return baldes[identificador_caso]


def mesclar_enrich_no_grafo(identificador_caso: str, vis_nodes: List[Dict[str, Any]], vis_edges: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    balde = obter_enrich_do_caso(identificador_caso)
    if not balde["nodes"] and not balde["edges"]:
        return vis_nodes, vis_edges
    ids_ja_presentes = {n["id"] for n in vis_nodes}
    nodes_finais = list(vis_nodes) + [n for nid, n in balde["nodes"].items() if nid not in ids_ja_presentes]
    ids_arestas_presentes = {e.get("id") for e in vis_edges if e.get("id")}
    edges_finais = list(vis_edges) + [e for eid, e in balde["edges"].items() if eid not in ids_arestas_presentes]
    return nodes_finais, edges_finais


def aplicar_resultado_enrich_ao_balde(
    identificador_caso: str,
    resultado_enrich: Dict[str, Any],
    vis_nodes_atuais: List[Dict[str, Any]],
    no_origem_extra: Optional[Dict[str, Any]] = None
) -> int:
    """
    Aplica os novos_nos/novas_arestas de um resultado de Enrich (vindo de um clique
    no grafo OU de uma busca manual na aba "Add") ao balde acumulado do caso,
    respeitando o limite e sem duplicar nada já presente.
    """
    balde = obter_enrich_do_caso(identificador_caso)
    ids_ja_no_subgrafo = {n["id"] for n in vis_nodes_atuais}
    novos_efetivos = 0

    candidatos = list(resultado_enrich.get("novos_nos", []))
    if no_origem_extra is not None:
        candidatos = [no_origem_extra] + candidatos

    for n_novo in candidatos:
        if len(balde["nodes"]) + len(vis_nodes_atuais) >= LIMITE_ENTIDADES_ENRIQUECIDAS_POR_CASO:
            break
        if n_novo["id"] not in ids_ja_no_subgrafo and n_novo["id"] not in balde["nodes"]:
            balde["nodes"][n_novo["id"]] = n_novo
            novos_efetivos += 1

    for e_novo in resultado_enrich.get("novas_arestas", []):
        balde["edges"][e_novo["id"]] = e_novo

    return novos_efetivos


def construir_no_manual(tipo_norm: str, valor_limpo: str, orbit_ao_redor: str) -> Dict[str, Any]:
    """Constrói o nó do próprio valor buscado manualmente na aba Add (não vem do enrich_engine,
    que só retorna vizinhos -- o valor pesquisado precisa virar nó por conta própria)."""
    cor = CORES_POR_TIPO.get(tipo_norm, {"bg": "#334155", "border": "#94A3B8"})
    prefixo = PREFIXO_POR_TIPO.get(tipo_norm, "ENT")
    if tipo_norm == "cpf":
        valor_fmt = formatar_cpf_cnpj(valor_limpo)
    elif tipo_norm == "telefone":
        valor_fmt = formatar_tel(valor_limpo)
    elif tipo_norm == "placa":
        valor_fmt = f"{valor_limpo[:3]}-{valor_limpo[3:]}" if len(valor_limpo) == 7 else valor_limpo
    else:
        valor_fmt = valor_limpo.upper()
    svg_image = _criar_svg_nodo_flowsint(tipo_norm, cor["bg"], cor["border"], is_hub=False)
    return {
        "id": f"{prefixo}_{valor_limpo}",
        "label": valor_fmt,
        "title": f"{tipo_norm.upper()}: {valor_fmt}\nAdicionado manualmente",
        "tipo": tipo_norm,
        "shape": "image",
        "image": svg_image,
        "size": 20,
        "degree": 1,
        "betweenness": 0.0,
        "community": 0,
        "valor": valor_fmt,
        "nome_titular": "",
        "orbit_ao_redor": orbit_ao_redor,
        "color": {"background": cor["bg"], "border": cor["border"]},
        "font": {"color": "#F8FAFC", "size": FONT_PADRAO, "face": "Segoe UI", "strokeWidth": 2.5, "strokeColor": "#06090F"}
    }


def badge_status_html(status: str) -> str:
    cor = CORES_STATUS.get(status, "#8B95A5")
    return f'<span class="badge-pill" style="background:{cor}1F; color:{cor}; border:1px solid {cor}44;">{status}</span>'


def badge_risco_html(nivel: str, score: int) -> str:
    cor = CORES_RISCO_NEUTRAS.get(nivel, "#8B95A5")
    return f'<span class="badge-pill" style="background:{cor}1F; color:{cor}; border:1px solid {cor}44;">{nivel} · {score}/100</span>'


def renderizar_painel_analises(identificador_caso: str, dados_caso_oficial: Dict[str, Any], key_sufixo: str = ""):
    """Compartilhado entre o Overview do caso e o toggle de Notas dentro da Mesa."""
    col_an1, col_an2 = st.columns([4.0, 1.6])
    with col_an1:
        st.markdown("##### Análises (Pareceres Técnicos)")
    with col_an2:
        if st.button("+ Nova Análise", use_container_width=True, key=f"btn_nova_analise_{identificador_caso}_{key_sufixo}"):
            st.session_state["mostrar_form_nova_analise"] = not st.session_state["mostrar_form_nova_analise"]

    if st.session_state["mostrar_form_nova_analise"]:
        st.caption("Use referências como @CPF_..., @TEL_... ou @PLACA_... para destacá-las na trilha de auditoria.")
        novo_parecer = st.text_area("Parecer Técnico:", height=120, key=f"novo_parecer_{identificador_caso}_{key_sufixo}")
        if st.button("Salvar Análise", type="primary", use_container_width=True, key=f"salvar_parecer_{identificador_caso}_{key_sufixo}"):
            if novo_parecer.strip():
                salvar_dados_caso(
                    identificador_caso,
                    dados_caso_oficial["nome_personalizado"],
                    dados_caso_oficial["status"],
                    dados_caso_oficial["analista_responsavel"],
                    novo_parecer
                )
                st.session_state["mostrar_form_nova_analise"] = False
                st.success("Análise registrada.")
                st.rerun()
            else:
                st.error("Escreva um parecer antes de salvar.")

    historico_notas = carregar_historico_pareceres(identificador_caso)
    if historico_notas:
        for h in historico_notas:
            st.markdown(f"""
            <div class="parecer-card">
                <span style="color:#6C93B0; font-size:12px; font-weight:600;">{h['data_registro']}</span> • 
                <span style="color:#C9A66B; font-size:12px;">Analista: <b>{h['analista'] or 'Sistema'}</b></span> • 
                <span style="color:#6FA98A; font-size:12px;">Status: <b>{h['status']}</b></span>
                <div style="color:#C7CDD6; font-size:13px; margin-top:6px; white-space: pre-wrap;">{formatar_mencoes_forenses(h['parecer'])}</div>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("Nenhuma análise registrada ainda. Use \"+ Nova Análise\" para escrever o primeiro parecer.")


def renderizar_painel_entidades(vis_nodes: List[Dict[str, Any]], identificador_caso: str, df_dados: pd.DataFrame):
    """Coluna lateral estilo Flowsint: lista de entidades do grafo com busca e foco,
    mais aba de adição manual (busca por valor cruzando o acervo local via Enrich)."""
    balde = obter_enrich_do_caso(identificador_caso)
    total_entidades = len(vis_nodes) + len(balde["nodes"])

    c_tab1, c_tab2 = st.columns(2)
    with c_tab1:
        if st.button(":material/group: Entities", use_container_width=True,
                     type=("primary" if st.session_state["aba_painel_entidades"] == "Entities" else "secondary")):
            st.session_state["aba_painel_entidades"] = "Entities"
            st.rerun()
    with c_tab2:
        if st.button(":material/person_add: Add", use_container_width=True,
                     type=("primary" if st.session_state["aba_painel_entidades"] == "Add" else "secondary")):
            st.session_state["aba_painel_entidades"] = "Add"
            st.rerun()

    if st.session_state["aba_painel_entidades"] == "Entities":
        st.caption(f"{total_entidades} nós")
        termo_lateral = st.text_input("Buscar tipo...", key=f"busca_lateral_{identificador_caso}", label_visibility="collapsed", placeholder="Buscar...")

        lista_todas = list(vis_nodes) + list(balde["nodes"].values())
        if termo_lateral.strip():
            termo_norm = termo_lateral.strip().upper()
            lista_exibicao = [
                n for n in lista_todas
                if termo_norm in (n.get("label", "") or "").upper()
                or termo_norm in (n.get("valor", "") or "").upper()
                or termo_norm in (n.get("nome_titular", "") or "").upper()
            ][:LIMITE_ENTIDADES_LATERAL_COM_BUSCA]
        else:
            lista_exibicao = sorted(lista_todas, key=lambda n: n.get("degree", 0), reverse=True)[:LIMITE_ENTIDADES_LATERAL_SEM_BUSCA]
            if len(lista_todas) > LIMITE_ENTIDADES_LATERAL_SEM_BUSCA:
                st.caption(f"Mostrando as {LIMITE_ENTIDADES_LATERAL_SEM_BUSCA} de maior grau. Digite para buscar entre as demais {len(lista_todas) - LIMITE_ENTIDADES_LATERAL_SEM_BUSCA}.")

        with st.container(height=420):
            for n in lista_exibicao:
                icone = ICONE_POR_TIPO.get(n.get("tipo", ""), ":material/help:")
                if st.button(f"{icone} {n.get('label', n['id'])}", use_container_width=True, key=f"lat_ent_{identificador_caso}_{n['id']}"):
                    st.session_state["entidade_foco_grafo"] = n["id"]
                    val_puro = n["id"].split("_", 1)[1] if "_" in n["id"] else n["id"]
                    sub_match = df_dados[
                        (df_dados["cpf"].astype(str) == val_puro) |
                        (df_dados["telefone"].astype(str) == val_puro) |
                        (df_dados["placa"].astype(str) == val_puro)
                    ].dropna(subset=["latitude", "longitude"]) if not df_dados.empty else pd.DataFrame()
                    if not sub_match.empty:
                        st.session_state["coordenada_foco"] = (float(sub_match.iloc[0]["latitude"]), float(sub_match.iloc[0]["longitude"]))
                    st.rerun()
    else:
        st.caption("Busque um CPF, telefone ou placa no acervo local e adicione ao grafo deste caso, junto com os vínculos diretos encontrados.")
        if not TEM_ENRICH_ENGINE:
            st.info("Módulo enrich_engine.py não localizado -- adição manual indisponível.")
        else:
            tipo_label_sel = st.selectbox("Tipo:", list(TIPO_LABEL_PARA_INTERNO.keys()), key=f"add_tipo_{identificador_caso}")
            valor_sel = st.text_input("Valor:", placeholder="Ex: 12345678900, (11) 97000-1122, ABC1234", key=f"add_valor_{identificador_caso}")

            if st.button(":material/search: Buscar e Adicionar", type="primary", use_container_width=True, key=f"add_btn_{identificador_caso}"):
                tipo_norm = TIPO_LABEL_PARA_INTERNO[tipo_label_sel]
                if tipo_norm == "placa":
                    valor_limpo = re.sub(r'[^a-zA-Z0-9]', '', valor_sel).upper()
                else:
                    valor_limpo = re.sub(r'[^0-9]', '', valor_sel)

                if not valor_limpo:
                    st.error("Informe um valor válido antes de buscar.")
                elif total_entidades >= LIMITE_ENTIDADES_ENRIQUECIDAS_POR_CASO:
                    st.warning(f"Limite de {LIMITE_ENTIDADES_ENRIQUECIDAS_POR_CASO} entidades por caso atingido.")
                else:
                    node_pai_id = f"{PREFIXO_POR_TIPO[tipo_norm]}_{valor_limpo}"
                    resultado_add = enriquecer_entidade_local(tipo_norm, valor_limpo)

                    no_origem = None
                    ids_existentes = {n["id"] for n in vis_nodes} | set(balde["nodes"].keys())
                    if node_pai_id not in ids_existentes:
                        no_origem = construir_no_manual(tipo_norm, valor_limpo, orbit_ao_redor="")

                    if resultado_add["total_encontrado"] == 0 and no_origem is None:
                        st.info("Nenhum vínculo adicional encontrado para este valor no acervo local.")
                    elif resultado_add["total_encontrado"] == 0 and no_origem is not None:
                        st.info("A entidade já existe no caso ou não possui vínculos cruzados no acervo local para adicionar.")
                    else:
                        novos_efetivos = aplicar_resultado_enrich_ao_balde(identificador_caso, resultado_add, vis_nodes, no_origem_extra=no_origem)
                        if novos_efetivos > 0:
                            st.success(f"{novos_efetivos} entidade(s) adicionada(s) ao grafo.")
                            st.rerun()
                        else:
                            st.info("As entidades relacionadas já estão presentes no grafo atual.")

# =====================================================
# 3. SIDEBAR GLOBAL (NAVEGAÇÃO FIXA + GESTÃO DE BASES)
# =====================================================
with st.sidebar:
    st.markdown("### LCFO")
    st.caption("Mesa de Inteligência Forense")

    for rotulo_nav, aba_alvo in NAV_PRINCIPAL:
        esta_no_dashboard = st.session_state["celula_ativa_id"] is None and not st.session_state["modo_descoberta_ativo"]
        ativo_nav = esta_no_dashboard and st.session_state["aba_principal_triagem"] == aba_alvo
        if st.button(rotulo_nav, use_container_width=True, type=("primary" if ativo_nav else "secondary"), key=f"nav_{aba_alvo}"):
            st.session_state["celula_ativa_id"] = None
            st.session_state["modo_descoberta_ativo"] = False
            st.session_state["dados_descoberta"] = None
            st.session_state["aba_principal_triagem"] = aba_alvo
            st.rerun()

    st.markdown("---")

    with st.popover(":material/database: Gestão de Bases", use_container_width=True):
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
    st.info("Nenhum dado processado no banco. Abra 'Gestão de Bases' na lateral e ingira a Blacklist.")
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

with st.sidebar:
    if st.session_state["celula_ativa_id"] is not None or st.session_state["modo_descoberta_ativo"]:
        st.markdown("---")
        if st.button(":material/arrow_back: Voltar à Triagem Geral", use_container_width=True):
            voltar_para_dashboard()

# ====================================================================
# 5. MOTOR UNIFICADO DE INVESTIGAÇÃO (MESA FORENSE)
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
    sufixos_ativos = ["unico", "esq", "dir"]
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

    # =================================================================
    # BARRA FIXA COMPACTA: voltar/alternar (esquerda) | busca (centro) | notas (direita)
    # =================================================================
    col_bar1, col_bar2, col_bar3 = st.columns([1.4, 2.6, 1.0])
    with col_bar1:
        with st.popover(f":material/menu: {nome_exibicao[:22]}", use_container_width=True):
            if eh_oficial:
                if st.button(":material/arrow_back: Voltar ao Overview", use_container_width=True):
                    st.session_state["subtela_caso"] = "overview"
                    st.session_state["coordenada_foco"] = None
                    st.session_state["entidade_foco_grafo"] = None
                    st.rerun()
                st.markdown("---")
                label_atual = mapa_id_para_label.get(identificador_caso, list(opcoes_celulas.keys())[0])
                idx_caso = list(opcoes_celulas.keys()).index(label_atual) if label_atual in opcoes_celulas else 0
                novo_caso_sel = st.selectbox("Alternar Caso:", list(opcoes_celulas.keys()), index=idx_caso, key=f"alternar_caso_{identificador_caso}")
                if opcoes_celulas[novo_caso_sel] != identificador_caso:
                    abrir_caso_em_overview(opcoes_celulas[novo_caso_sel])
            else:
                if st.button(":material/close: Fechar Descoberta", use_container_width=True):
                    voltar_para_dashboard()

    with col_bar2:
        with st.popover(":material/search: Buscar entidade neste caso", use_container_width=True):
            termo_busca_topo = st.text_input("Valor:", key=f"busca_topo_{identificador_caso}", placeholder="CPF, telefone, placa ou nome...")
            if termo_busca_topo.strip():
                termo_norm = termo_busca_topo.strip().upper()
                achados = [
                    n for n in st.session_state.get("_vis_nodes_atual_cache", [])
                    if termo_norm in (n.get("label", "") or "").upper() or termo_norm in (n.get("valor", "") or "").upper()
                ][:15]
                for n in achados:
                    icone = ICONE_POR_TIPO.get(n.get("tipo", ""), ":material/help:")
                    if st.button(f"{icone} {n.get('label', n['id'])}", use_container_width=True, key=f"busca_topo_res_{identificador_caso}_{n['id']}"):
                        st.session_state["entidade_foco_grafo"] = n["id"]
                        st.rerun()

    with col_bar3:
        notas_ativas = st.session_state["mostrar_notas_grafo"]
        if st.button(":material/edit_note: Notas", use_container_width=True, type=("primary" if notas_ativas else "secondary")):
            st.session_state["mostrar_notas_grafo"] = not notas_ativas
            st.rerun()

    col_nav2, col_nav3, col_nav4 = st.columns([1.8, 1.0, 1.0])
    with col_nav2:
        idx_modo = 0 if st.session_state["cfg_modo_exibicao"] == "Abas Clássicas" else 1
        st.radio("Visualização:", ["Abas Clássicas", "Cockpit Dividido"], index=idx_modo, horizontal=True, key="w_modo_exibicao", on_change=cb_atualizar_modo, label_visibility="collapsed")

    mostrar_tel, mostrar_cpf, mostrar_placa = True, True, True
    with col_nav3:
        if eh_oficial:
            with st.popover(":material/filter_alt: Filtros", use_container_width=True):
                st.caption("Tipos de entidade exibidos no grafo deste caso.")
                mostrar_tel = st.checkbox(f"Telefones ({cluster_obj['qtd_tels'] if cluster_obj else 0})", value=True, key=f"flt_tel_{identificador_caso}")
                mostrar_cpf = st.checkbox(f"Titulares ({cluster_obj['qtd_cpfs'] if cluster_obj else 0})", value=True, key=f"flt_cpf_{identificador_caso}")
                mostrar_placa = st.checkbox(f"Placas ({cluster_obj['qtd_placas'] if cluster_obj else 0})", value=True, key=f"flt_placa_{identificador_caso}")
    with col_nav4:
        if not eh_oficial:
            if st.button(":material/close: Fechar", use_container_width=True):
                voltar_para_dashboard()

    if eh_oficial:
        filtro_tipos = []
        if mostrar_tel: filtro_tipos.append("telefone")
        if mostrar_cpf: filtro_tipos.append("cpf")
        if mostrar_placa: filtro_tipos.append("placa")
        vis_nodes, vis_edges = processar_subgrafo_caso(
            G=G, cluster_nodes=cluster_obj["nodes"] if cluster_obj else [],
            filtro_tipos=filtro_tipos, hub_id=hub_id, font_slider=FONT_PADRAO,
            id_caso=identificador_caso, espacamento=ESPACAMENTO_PADRAO
        )
        max_bet = max((n.get("betweenness", 0.0) for n in vis_nodes), default=0.0)
        subG_ativo = G.subgraph(cluster_obj["nodes"]) if cluster_obj else None
        score, nivel, cor, fatores, _ = calcular_score_caso(
            cluster_obj=cluster_obj, df_assistencias=df_dados, subgrafo=subG_ativo, betweenness_precalculado=max_bet
        )
    else:
        vis_nodes, vis_edges, hub_id_calc = processar_grafo_dataframe(
            df_dados, alvo_principal=resumo_descoberta.get("termo_limpo", "") if resumo_descoberta else "",
            font_slider=FONT_PADRAO, espacamento=ESPACAMENTO_PADRAO
        )
        hub_id = hub_id_calc
        score, nivel, cor, fatores, _ = calcular_score_caso(cluster_obj=None, df_assistencias=df_dados, subgrafo=None)

    cor = obter_cor_risco(nivel, cor)
    vis_nodes, vis_edges = mesclar_enrich_no_grafo(identificador_caso, vis_nodes, vis_edges)
    st.session_state["_vis_nodes_atual_cache"] = vis_nodes

    periodo_str = "Sem datas disponíveis"
    if not df_dados.empty and "data" in df_dados.columns:
        df_temp = df_dados[df_dados["data"].astype(str).str.strip() != ""].copy()
        if not df_temp.empty:
            dts_validas = pd.to_datetime(df_temp["data"], errors="coerce").dropna()
            if not dts_validas.empty:
                periodo_str = f"{dts_validas.min().strftime('%d/%m/%Y')} → {dts_validas.max().strftime('%d/%m/%Y')}"

    campos_extra = []
    if eh_oficial and dados_caso_oficial:
        st_val = dados_caso_oficial.get("status", "Em Investigação")
        campos_extra.append(f'<span style="font-size:0.82rem; color:#7C8798;">Status: <b style="color:#6FA98A;">{st_val}</b></span>')
        if hub_id:
            campos_extra.append(f'<span style="font-size:0.82rem; color:#C9A66B;">Âncora: <b>{hub_id}</b></span>')
    elif not eh_oficial and resumo_descoberta:
        campos_extra.append(
            f'<span style="font-size:0.82rem; color:#8F86B5;">Entidades: <b>'
            f'{resumo_descoberta.get("qtd_cpfs", 0)} CPFs | {resumo_descoberta.get("qtd_tels", 0)} Tels | '
            f'{resumo_descoberta.get("qtd_placas", 0)} Placas</b></span>'
        )
    campos_extra_html = "".join(campos_extra)

    st.markdown(montar_hud_bar(cor, nome_exibicao, score, nivel, campos_extra_html, len(df_dados), periodo_str), unsafe_allow_html=True)
    st.markdown(montar_caixa_fatores(cor, fatores), unsafe_allow_html=True)

    balde_enrich = obter_enrich_do_caso(identificador_caso)
    qtd_enrich = len(balde_enrich["nodes"])
    if st.session_state["coordenada_foco"] or st.session_state["entidade_foco_grafo"] or qtd_enrich > 0:
        c_foco1, c_foco2, c_foco3 = st.columns([3.4, 1.0, 1.0])
        with c_foco1:
            ent_nome = st.session_state["entidade_foco_grafo"] or "Ocorrência selecionada"
            if qtd_enrich > 0:
                st.info(f"Filtro ativo: **{ent_nome}** — Mapa e grafo sincronizados. **{qtd_enrich}** entidade(s) via Enrich/Add.")
            else:
                st.info(f"Filtro ativo: **{ent_nome}** — Mapa e grafo sincronizados neste ponto.")
        with c_foco2:
            if st.button("Limpar Foco", use_container_width=True):
                st.session_state["coordenada_foco"] = None
                st.session_state["entidade_foco_grafo"] = None
                for sfx in ["unico", "esq", "dir"]:
                    k_tab = f"tab_ocorr_{identificador_caso}_{sfx}"
                    if k_tab in st.session_state:
                        del st.session_state[k_tab]
                st.rerun()
        with c_foco3:
            if qtd_enrich > 0 and st.button("Limpar Enrich", use_container_width=True):
                st.session_state["nos_enriquecidos_por_caso"][identificador_caso] = {"nodes": {}, "edges": {}}
                st.rerun()

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
                            abrir_caso_em_overview(cluster_promovido["id"])
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
                    st.success(f"Dados anexados com sucesso ao Caso {id_destino}.")
                    abrir_caso_em_overview(id_destino)

    alvo_grafo_id = st.session_state.get("entidade_foco_grafo") or ""
    if alvo_grafo_id and not any(n["id"] == alvo_grafo_id for n in vis_nodes):
        alvo_grafo_id = ""

    html_codigo = gerar_html_grafo(
        vis_nodes_json=json.dumps(vis_nodes), vis_edges_json=json.dumps(vis_edges),
        base_font_size=FONT_PADRAO, hub_id=hub_id, target_node_id=alvo_grafo_id,
        layout_ativo=LAYOUT_INICIAL_PADRAO, espacamento=ESPACAMENTO_PADRAO
    )

    def render_modulo(nome_modulo, altura=720, sufixo_key=""):
        if "Grafo" in nome_modulo:
            evento_grafo = None
            if TEM_COMPONENTE_BIDIRECIONAL:
                evento_grafo = renderizar_grafo_bidirecional(
                    nodes=vis_nodes, edges=vis_edges, hub_id=hub_id, target_node_id=alvo_grafo_id,
                    layout_ativo=LAYOUT_INICIAL_PADRAO, base_font_size=FONT_PADRAO,
                    espacamento=ESPACAMENTO_PADRAO, height=altura, key=f"comp_grafo_{identificador_caso}_{sufixo_key}"
                )
            else:
                components.html(html_codigo, height=altura)

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

                elif tipo_ev == "reset_layout_request":
                    resetar_layout_caso(identificador_caso)
                    st.cache_data.clear()
                    st.success("Layout recalculado a partir do canvas.")
                    st.rerun()

                elif tipo_ev == "enrich_request" and TEM_ENRICH_ENGINE:
                    tipo_entidade_ev = evento_grafo.get("tipoEntidade", "")
                    valor_ev = evento_grafo.get("valor", "")
                    node_id_origem = node_ev
                    balde = obter_enrich_do_caso(identificador_caso)
                    total_atual = len(vis_nodes) + len(balde["nodes"])

                    if total_atual >= LIMITE_ENTIDADES_ENRIQUECIDAS_POR_CASO:
                        st.warning(f"Limite de {LIMITE_ENTIDADES_ENRIQUECIDAS_POR_CASO} entidades por caso atingido.")
                    elif node_id_origem:
                        valor_bruto = node_id_origem.split("_", 1)[1] if "_" in node_id_origem else valor_ev
                        resultado_enrich = enriquecer_entidade_local(tipo_entidade_ev, valor_bruto)
                        if resultado_enrich["total_encontrado"] == 0:
                            st.info("Nenhum vínculo adicional localizado para esta entidade no acervo local.")
                        else:
                            novos_efetivos = aplicar_resultado_enrich_ao_balde(identificador_caso, resultado_enrich, vis_nodes)
                            if novos_efetivos > 0:
                                st.success(f"Enrich: {novos_efetivos} nova(s) entidade(s) incorporada(s) ao grafo.")
                            else:
                                st.info("As entidades relacionadas já estão presentes no grafo atual.")
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
                    get_color="[108, 147, 176, 180]", get_line_color="[255, 255, 255, 200]",
                    line_width_min_pixels=1, stroked=True, get_radius=3200, radius_min_pixels=6, radius_max_pixels=20, pickable=True
                )
                camadas = [camada_geral]
                if foco:
                    df_foco = pd.DataFrame([{"latitude": foco[0], "longitude": foco[1]}])
                    camada_foco = pdk.Layer(
                        "ScatterplotLayer", data=df_foco, get_position="[longitude, latitude]",
                        get_color="[201, 166, 107, 220]", get_line_color="[255, 255, 255, 240]",
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
                st.dataframe(df_dados[cols_exib], use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row", key=f"tab_ocorr_{identificador_caso}_{sufixo_key}")
            except Exception:
                st.dataframe(df_dados[cols_exib], use_container_width=True, hide_index=True)
            csv_exp = df_dados.to_csv(index=False).encode('utf-8')
            st.download_button("Baixar Ocorrências em CSV", csv_exp, file_name=f"Assistencias_{identificador_caso}.csv", mime="text/csv", key=f"dl_csv_{identificador_caso}_{sufixo_key}")

        elif "Temporal" in nome_modulo:
            if not df_dados.empty and "data" in df_dados.columns:
                df_temp = df_dados.copy()
                df_temp["Periodo"] = pd.to_datetime(df_temp["data"], errors="coerce").dt.strftime("%Y-%m")
                df_temp = df_temp.dropna(subset=["Periodo"]).groupby("Periodo").size().reset_index(name="Assistências").sort_values(by="Periodo")
                st.bar_chart(data=df_temp.set_index("Periodo"), color="#6C93B0", use_container_width=True)
            else:
                st.info("Sem datas disponíveis para evolução temporal.")

        elif "Expansão" in nome_modulo:
            if not eh_oficial or not cluster_obj:
                st.info("A Expansão com Criações Diárias está disponível apenas para casos oficiais.")
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
                    with c1: st.markdown(f"""<div class="intel-card"><span style="color:#7C8798; font-size:10px;">ACIONAMENTOS</span><br/><b style="font-size:18px; color:#6C93B0;">{resumo_intel['total_assistencias']}</b></div>""", unsafe_allow_html=True)
                    with c2: st.markdown(f"""<div class="intel-card"><span style="color:#7C8798; font-size:10px;">NOVOS TELS</span><br/><b style="font-size:18px; color:#C0625F;">+{resumo_intel['novos_tels']}</b></div>""", unsafe_allow_html=True)
                    with c3: st.markdown(f"""<div class="intel-card"><span style="color:#7C8798; font-size:10px;">NOVAS PLACAS</span><br/><b style="font-size:18px; color:#8F86B5;">+{resumo_intel['novas_placas']}</b></div>""", unsafe_allow_html=True)
                    with c4: st.markdown(f"""<div class="intel-card"><span style="color:#7C8798; font-size:10px;">NOVOS CPFS</span><br/><b style="font-size:18px; color:#C9A66B;">+{resumo_intel['novos_cpfs']}</b></div>""", unsafe_allow_html=True)

                st.markdown("<br/>", unsafe_allow_html=True)
                if not df_novos_suspeitos.empty:
                    st.markdown(f"**Novos Suspeitos Detectados na Rede:**")
                    st.dataframe(df_novos_suspeitos, use_container_width=True, hide_index=True)
                    c_inc1, c_inc2 = st.columns([2.5, 2.5])
                    with c_inc1:
                        if st.button("Incorporar Novos Suspeitos e Anexar ao Caso", type="primary", use_container_width=True, key=f"btn_inc_{identificador_caso}_{sufixo_key}"):
                            anexar_descoberta_a_caso_existente(
                                df_matches_criacao, identificador_caso,
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

        elif "Ferramentas" in nome_modulo:
            if not eh_oficial or not cluster_obj or not dados_caso_oficial:
                st.info("As ferramentas de investigação estão disponíveis apenas para casos oficiais.")
                return

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

            if TEM_EXIF_ENGINE:
                st.markdown("---")
                with st.expander("Perícia Forense de Imagens (EXIF de Fotos de Reembolso)", expanded=False):
                    st.caption("Confronte o GPS da câmera contra a praça declarada NA ASSISTÊNCIA ESPECÍFICA selecionada abaixo.")
                    if df_dados.empty or "id_assistencia" not in df_dados.columns:
                        st.info("Não há assistências com identificador disponível neste caso.")
                    else:
                        opcoes_assist = {}
                        for _, r_assist in df_dados.iterrows():
                            rot = f"{r_assist.get('id_assistencia', '?')} — {r_assist.get('cidade', '?')}/{r_assist.get('uf', '?')} ({r_assist.get('data', '?')})"
                            opcoes_assist[rot] = (str(r_assist.get("cidade", "")), str(r_assist.get("uf", "")))
                        rotulo_assist_sel = st.selectbox("Assistência à qual esta foto se refere:", list(opcoes_assist.keys()), key=f"exif_assist_sel_{identificador_caso}")
                        cidade_sel, uf_sel = opcoes_assist[rotulo_assist_sel]
                        foto_submetida = st.file_uploader("Arquivo de Imagem (JPEG/PNG):", type=["jpg", "jpeg", "png"], key=f"exif_upload_{identificador_caso}")
                        if foto_submetida:
                            bytes_img = foto_submetida.read()
                            metadados_extraidos = extrair_metadados_foto(bytes_img)
                            if not metadados_extraidos["possui_exif"]:
                                st.warning("Metadados EXIF ausentes na imagem.")
                            elif metadados_extraidos.get("erro"):
                                st.error(f"Não foi possível processar a imagem: {metadados_extraidos['erro']}")
                            else:
                                c_ex1, c_ex2 = st.columns(2)
                                with c_ex1:
                                    st.markdown("**Metadados da Câmera:**")
                                    st.caption(f"• Dispositivo: {metadados_extraidos.get('fabricante') or 'N/D'} {metadados_extraidos.get('modelo') or ''}")
                                    st.caption(f"• Data do Registro: {metadados_extraidos.get('data_captura') or 'N/D'}")
                                    lat_exif = metadados_extraidos.get("latitude")
                                    lon_exif = metadados_extraidos.get("longitude")
                                    st.caption(f"• Coordenadas GPS: [{lat_exif}, {lon_exif}]" if lat_exif is not None else "• Coordenadas GPS: não presentes.")
                                with c_ex2:
                                    st.markdown(f"**Confronto Territorial ({cidade_sel}/{uf_sel}):**")
                                    resultado_confronto = validar_coerencia_geografica_foto(
                                        metadados_foto=metadados_extraidos, cidade_declarada=cidade_sel,
                                        uf_declarada=uf_sel, tolerancia_km=TOLERANCIA_KM_EXIF
                                    )
                                    if not resultado_confronto["apto_para_confronto"]:
                                        st.info(resultado_confronto["parecer_exif"])
                                    elif resultado_confronto["divergencia_critica"]:
                                        st.error(resultado_confronto["parecer_exif"])
                                    else:
                                        st.success(resultado_confronto["parecer_exif"])

    if st.session_state["cfg_modo_exibicao"] == "Abas Clássicas":
        col_entidades, col_conteudo = st.columns([1.3, 4.2])
        with col_entidades:
            renderizar_painel_entidades(vis_nodes, identificador_caso, df_dados)
        with col_conteudo:
            opcoes_toolbar = OPCOES_TOOLBAR_OFICIAL if eh_oficial else OPCOES_TOOLBAR_DESCOBERTA
            if st.session_state["cfg_painel_unico"] not in opcoes_toolbar:
                st.session_state["cfg_painel_unico"] = opcoes_toolbar[0]

            cols_toolbar = st.columns(len(opcoes_toolbar))
            for i, opc in enumerate(opcoes_toolbar):
                with cols_toolbar[i]:
                    ativo = st.session_state["cfg_painel_unico"] == opc
                    if st.button(ICONE_TOOLBAR[opc], use_container_width=True, type=("primary" if ativo else "secondary"), key=f"toolbar_{identificador_caso}_{i}"):
                        st.session_state["cfg_painel_unico"] = opc
                        st.rerun()

            st.markdown("<br/>", unsafe_allow_html=True)
            render_modulo(st.session_state["cfg_painel_unico"], altura=680, sufixo_key="unico")

        if eh_oficial and st.session_state["mostrar_notas_grafo"] and dados_caso_oficial:
            st.markdown("---")
            renderizar_painel_analises(identificador_caso, dados_caso_oficial, key_sufixo="mesa_toggle")
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
            st.markdown(f"<div style='margin-bottom:4px;'><b style='color:#6C93B0;'>{st.session_state['cfg_painel_esquerdo']}</b></div>", unsafe_allow_html=True)
            render_modulo(st.session_state["cfg_painel_esquerdo"], altura=680, sufixo_key="esq")
        with col_right:
            st.markdown(f"<div style='margin-bottom:4px;'><b style='color:#6C93B0;'>{st.session_state['cfg_painel_direito']}</b></div>", unsafe_allow_html=True)
            render_modulo(st.session_state["cfg_painel_direito"], altura=680, sufixo_key="dir")


# ====================================================================
# 6. TELA DE OVERVIEW DO CASO
# ====================================================================
def renderizar_overview_caso(cluster_obj: Dict[str, Any], dados_caso_oficial: Dict[str, Any], identificador_caso: str):
    nome_exib = dados_caso_oficial["nome_personalizado"] or f"Caso {identificador_caso}"

    if st.button(":material/arrow_back: Casos", key="breadcrumb_voltar"):
        voltar_para_dashboard()
    st.markdown(
        f"<span style='color:#7C8798; font-size:12px;'>Casos / </span>"
        f"<span style='color:#E5E9F0; font-size:12px; font-weight:600;'>{nome_exib}</span>",
        unsafe_allow_html=True
    )

    mapa_scores = obter_scores_triagem(cluster_info)
    dados_score = mapa_scores.get(identificador_caso, {"score": 0, "nivel": "BAIXO", "cor": "#6FA98A"})
    status_atual = dados_caso_oficial.get("status", "Em Investigação")

    st.markdown(f"""
    <div class="overview-header-card">
        <div style="font-size:20px; font-weight:700; color:#E5E9F0; margin-bottom:8px;">{nome_exib}</div>
        <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin-bottom:6px;">
            {badge_status_html(status_atual)}
            {badge_risco_html(dados_score['nivel'], dados_score['score'])}
        </div>
        <div style="font-size:12px; color:#7C8798;">
            Atualizado: <b style="color:#C7CDD6;">{dados_caso_oficial.get('data_atualizacao') or '-'}</b>
            &nbsp;·&nbsp; Analista: <b style="color:#C7CDD6;">{dados_caso_oficial.get('analista_responsavel') or 'Não atribuído'}</b>
            &nbsp;·&nbsp; Âncora Central: <b style="color:#C7CDD6;">{cluster_obj['hub_label']}</b>
        </div>
    </div>
    """, unsafe_allow_html=True)

    c_meta1, c_meta2, c_meta3 = st.columns([2.5, 1.5, 1.5])
    with c_meta1: novo_nome = st.text_input("Nome da Quadrilha / Caso:", value=dados_caso_oficial["nome_personalizado"], placeholder="Ex: Quadrilha Santo André")
    with c_meta2:
        lista_status_caso = ["Em Investigação", "Confirmado Fraude", "Monitoramento Contínuo", "Sem Irregularidade Identificada", "Falso Positivo", "Arquivado"]
        idx_status_atual = lista_status_caso.index(status_atual) if status_atual in lista_status_caso else 0
        novo_status = st.selectbox("Status:", lista_status_caso, index=idx_status_atual, key=f"overview_status_{identificador_caso}")
    with c_meta3: novo_analista = st.text_input("Analista:", value=dados_caso_oficial["analista_responsavel"], key=f"overview_analista_{identificador_caso}")

    if st.button("Salvar Metadados do Caso", type="primary", use_container_width=True, key=f"overview_salvar_meta_{identificador_caso}"):
        salvar_dados_caso(identificador_caso, novo_nome, novo_status, novo_analista, dados_caso_oficial["parecer"])
        st.success("Metadados gravados com sucesso.")
        st.rerun()

    st.markdown("---")
    st.markdown("#### Grafo de Vínculos")
    c_sk1, c_sk2 = st.columns([3.0, 1.0])
    with c_sk1:
        st.markdown(f"""
        <div class="case-card">
            <div class="case-card-stats">
                <span>{cluster_obj['tamanho']} entidades</span><span>·</span>
                <span>{cluster_obj['qtd_cpfs']} titulares</span><span>·</span>
                <span>{cluster_obj['qtd_tels']} telefones</span><span>·</span>
                <span>{cluster_obj['qtd_placas']} placas</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    with c_sk2:
        st.write("")
        if st.button("Abrir Mesa Completa", type="primary", use_container_width=True):
            st.session_state["subtela_caso"] = "mesa"
            st.rerun()

    st.markdown("---")
    renderizar_painel_analises(identificador_caso, dados_caso_oficial, key_sufixo="overview")


# ====================================================================
# 7. ROTEAMENTO DE TELAS
# ====================================================================
if st.session_state["modo_descoberta_ativo"] and st.session_state["dados_descoberta"]:
    dados_d = st.session_state["dados_descoberta"]
    renderizar_mesa_investigacao(
        df_dados=dados_d["df"], identificador_caso="DESCOBERTA",
        nome_exibicao=dados_d["resumo"]["alvo_buscado"], hub_id="", eh_oficial=False,
        resumo_descoberta=dados_d["resumo"], cluster_obj=None
    )

elif st.session_state["celula_ativa_id"] is not None:
    cluster_selecionado = next((c for c in cluster_info if c["id"] == st.session_state["celula_ativa_id"]), cluster_info[0])
    dados_caso = carregar_dados_caso(cluster_selecionado["id"])
    nome_exib = dados_caso["nome_personalizado"] or f"Caso {cluster_selecionado['id']}"

    if st.session_state.get("subtela_caso", "overview") == "mesa":
        lista_cpfs = [n.replace("CPF_", "") for n in cluster_selecionado["nodes"] if n.startswith("CPF_")]
        lista_tels = [n.replace("TEL_", "") for n in cluster_selecionado["nodes"] if n.startswith("TEL_")]
        lista_placas = [n.replace("PLACA_", "") for n in cluster_selecionado["nodes"] if n.startswith("PLACA_")]
        df_caso = consultar_detalhes_caso(lista_cpfs, lista_tels, lista_placas)
        renderizar_mesa_investigacao(
            df_dados=df_caso, identificador_caso=cluster_selecionado["id"], nome_exibicao=nome_exib,
            hub_id=cluster_selecionado["hub_id"], eh_oficial=True, dados_caso_oficial=dados_caso, cluster_obj=cluster_selecionado
        )
    else:
        renderizar_overview_caso(cluster_obj=cluster_selecionado, dados_caso_oficial=dados_caso, identificador_caso=cluster_selecionado["id"])

# ====================================================================
# 8. DASHBOARD (TELA 1)
# ====================================================================
else:
    aba_triagem = st.session_state["aba_principal_triagem"]

    if aba_triagem == "Células e Redes da Blacklist":
        st.markdown("### Casos")
        st.caption("Gerencie e acompanhe as investigações de fraude detectadas na Blacklist.")

        termo_busca_global = st.text_input(
            "Buscar:", placeholder="Busque por CPF, placa, telefone, ou investigue um alvo inédito nas Criações Diárias...",
            key="busca_dashboard", label_visibility="collapsed"
        ).strip()

        if termo_busca_global:
            termo_clean = re.sub(r'[^a-zA-Z0-9]', '', termo_busca_global).upper()
            alvos_bl = [n for n in G.nodes if termo_clean in n]
            if alvos_bl:
                alvo = alvos_bl[0]
                for c in cluster_info:
                    if alvo in c["nodes"]:
                        nome_caso_bl = casos_cadastrados.get(c["id"], {}).get("nome", c["hub_label"])
                        st.success(f"Localizado no Caso Oficial: {nome_caso_bl} ({c['id']})")
                        if st.button(f"Abrir Caso {c['id']}", type="primary"):
                            abrir_caso_em_overview(c["id"])
                        break
            else:
                resumo_disc, df_disc, ents_disc = investigar_alvo_em_criacoes_diarias(termo_busca_global)
                if resumo_disc and resumo_disc["total_assistencias"] > 0:
                    st.warning(f"Dados localizados nas Criações Diárias:\n• {resumo_disc['total_assistencias']} assistências\n• Rede: {resumo_disc['qtd_cpfs']} CPFs, {resumo_disc['qtd_tels']} telefones, {resumo_disc['qtd_placas']} placas.")
                    if st.button("Abrir Mesa de Investigação Ativa", type="primary"):
                        abrir_descoberta_ativa(resumo_disc, df_disc, ents_disc)
                else:
                    st.info("Nenhum acionamento localizado para este dado em nenhuma das bases.")

        st.markdown("<br/>", unsafe_allow_html=True)

        k1, k2, k3, k4 = st.columns(4)
        with k1: st.markdown(f"""<div class="kpi-card-mini"><span style="color:#7C8798; font-size:10px; font-weight:600;">CÉLULAS</span> <b style="font-size:15px; color:#E5E9F0;">{len(cluster_info):,}</b></div>""", unsafe_allow_html=True)
        with k2: st.markdown(f"""<div class="kpi-card-mini"><span style="color:#7C8798; font-size:10px; font-weight:600;">ENTIDADES</span> <b style="font-size:15px; color:#6C93B0;">{len(G.nodes):,}</b></div>""", unsafe_allow_html=True)
        with k3: st.markdown(f"""<div class="kpi-card-mini"><span style="color:#7C8798; font-size:10px; font-weight:600;">QUADRILHAS</span> <b style="font-size:15px; color:#6FA98A;">{len([c for c in casos_cadastrados.values() if c['nome']]):,}</b></div>""", unsafe_allow_html=True)
        with k4:
            qtd_com_alerta = len(radar_alertas)
            cor_alerta = "#C0625F" if qtd_com_alerta > 0 else "#6FA98A"
            st.markdown(f"""<div class="kpi-card-mini"><span style="color:#7C8798; font-size:10px; font-weight:600;">NO RADAR</span> <b style="font-size:15px; color:{cor_alerta};">{qtd_com_alerta}</b></div>""", unsafe_allow_html=True)

        st.markdown("<br/>", unsafe_allow_html=True)

        if radar_alertas:
            with st.expander(f"{len(radar_alertas)} células reincidentes precisam de atenção", expanded=False):
                cols_radar = st.columns(min(4, len(radar_alertas)))
                for i, (cid, total_hits) in enumerate(list(radar_alertas.items())[:4]):
                    with cols_radar[i]:
                        c_obj = next((c for c in cluster_info if c["id"] == cid), None)
                        if c_obj:
                            nome_q = casos_cadastrados.get(cid, {}).get("nome", c_obj['hub_label'][:20])
                            if st.button(f"{nome_q} (+{total_hits})", use_container_width=True):
                                abrir_caso_em_overview(cid)

        mapa_scores = obter_scores_triagem(cluster_info)
        casos_processados = []
        for c in cluster_info[:150]:
            info_c = casos_cadastrados.get(c["id"], {})
            status_c = info_c.get("status", "Em Investigação")
            dados_score = mapa_scores.get(c["id"], {"score": 0, "nivel": "BAIXO", "cor": "#6FA98A"})
            casos_processados.append({
                "id": c["id"], "cluster": c, "info": info_c, "status": status_c,
                "score": dados_score["score"], "nivel": dados_score["nivel"],
                "em_radar": c["id"] in radar_alertas
            })

        qtd_todos = len(casos_processados)
        qtd_ativos = len([cp for cp in casos_processados if cp["status"] in STATUS_ATIVOS])
        qtd_encerrados = len([cp for cp in casos_processados if cp["status"] in STATUS_ENCERRADOS])

        opcoes_status_pills = [f"Todos ({qtd_todos})", f"Ativos ({qtd_ativos})", f"Encerrados ({qtd_encerrados})"]
        mapa_status_reverso = {opcoes_status_pills[0]: "Todos", opcoes_status_pills[1]: "Ativos", opcoes_status_pills[2]: "Encerrados"}
        idx_status_atual_pill = {"Todos": 0, "Ativos": 1, "Encerrados": 2}.get(st.session_state["filtro_status_dash"], 0)

        c_filtro1, c_filtro2 = st.columns([2.5, 3.5])
        with c_filtro1:
            sel_status_pill = st.pills("Status", opcoes_status_pills, selection_mode="single", default=opcoes_status_pills[idx_status_atual_pill], key="pills_status_dash")
            if sel_status_pill:
                st.session_state["filtro_status_dash"] = mapa_status_reverso[sel_status_pill]

        niveis_risco = ["Todos", "BAIXO", "MÉDIO", "ALTO", "CRÍTICO"]
        contagem_risco = {niv: len([cp for cp in casos_processados if cp["nivel"] == niv]) for niv in niveis_risco[1:]}
        contagem_risco["Todos"] = qtd_todos
        opcoes_risco_pills = [f"{niv} ({contagem_risco[niv]})" for niv in niveis_risco]
        mapa_risco_reverso = {f"{niv} ({contagem_risco[niv]})": niv for niv in niveis_risco}
        idx_risco_atual = niveis_risco.index(st.session_state["filtro_risco_dash"]) if st.session_state["filtro_risco_dash"] in niveis_risco else 0

        with c_filtro2:
            sel_risco_pill = st.pills("Risco", opcoes_risco_pills, selection_mode="single", default=opcoes_risco_pills[idx_risco_atual], key="pills_risco_dash")
            if sel_risco_pill:
                st.session_state["filtro_risco_dash"] = mapa_risco_reverso[sel_risco_pill]

        col_tit_lista, col_view = st.columns([4.0, 1.6])
        with col_tit_lista:
            st.markdown("#### Fila de Investigação Priorizada por Risco")
        with col_view:
            sel_view = st.pills("Modo", [":material/grid_view: Grade", ":material/view_list: Lista"], selection_mode="single",
                                 default=(":material/grid_view: Grade" if st.session_state["modo_visualizacao_casos"] == "Grade" else ":material/view_list: Lista"),
                                 key="pills_view_dash", label_visibility="collapsed")
            if sel_view:
                st.session_state["modo_visualizacao_casos"] = "Grade" if "Grade" in sel_view else "Lista"

        casos_filtrados = casos_processados
        if st.session_state["filtro_status_dash"] == "Ativos":
            casos_filtrados = [cp for cp in casos_filtrados if cp["status"] in STATUS_ATIVOS]
        elif st.session_state["filtro_status_dash"] == "Encerrados":
            casos_filtrados = [cp for cp in casos_filtrados if cp["status"] in STATUS_ENCERRADOS]
        if st.session_state["filtro_risco_dash"] != "Todos":
            casos_filtrados = [cp for cp in casos_filtrados if cp["nivel"] == st.session_state["filtro_risco_dash"]]

        casos_filtrados.sort(key=lambda cp: (cp["score"], cp["cluster"]["tamanho"]), reverse=True)

        if not casos_filtrados:
            st.info("Nenhum caso corresponde aos filtros selecionados.")
        elif st.session_state["modo_visualizacao_casos"] == "Grade":
            n_cols = 3
            for linha_inicio in range(0, len(casos_filtrados), n_cols):
                cols_grade = st.columns(n_cols)
                for j, cp in enumerate(casos_filtrados[linha_inicio:linha_inicio + n_cols]):
                    with cols_grade[j]:
                        c = cp["cluster"]
                        nome_exib_card = cp["info"].get("nome", "Não Batizado")
                        radar_txt = f"<div style='margin-top:4px; font-size:10px; color:#C0625F;'>+{radar_alertas[cp['id']]} novas</div>" if cp["em_radar"] else ""
                        with st.container(border=True, key=f"casecard_{cp['id']}"):
                            if st.button(nome_exib_card, key=f"abrir_grade_{cp['id']}", use_container_width=True):
                                abrir_caso_em_overview(cp["id"])
                            st.markdown(f"""
                            <div class="case-card-sub">{c['hub_label']}</div>
                            <div class="case-card-top">{badge_status_html(cp['status'])}{badge_risco_html(cp['nivel'], cp['score'])}</div>
                            <div class="case-card-stats">
                                <span>{c['tamanho']} nós</span><span>·</span>
                                <span>{c['qtd_tels']} tel</span><span>·</span>
                                <span>{c['qtd_cpfs']} tit</span><span>·</span>
                                <span>{c['qtd_placas']} placas</span>
                            </div>
                            {radar_txt}
                            """, unsafe_allow_html=True)
                st.markdown("<br/>", unsafe_allow_html=True)
        else:
            for cp in casos_filtrados:
                c = cp["cluster"]
                nome_exib_card = cp["info"].get("nome", "Não Batizado")
                with st.container(key=f"casecard_{cp['id']}"):
                    c_l1, c_l2, c_l3 = st.columns([3.0, 2.5, 2.0])
                    with c_l1:
                        if st.button(nome_exib_card, key=f"abrir_lista_{cp['id']}", use_container_width=True):
                            abrir_caso_em_overview(cp["id"])
                        st.markdown(f"<span style='color:#7C8798; font-size:11px;'>{c['hub_label']}</span>", unsafe_allow_html=True)
                    with c_l2:
                        st.markdown(badge_status_html(cp["status"]) + " " + badge_risco_html(cp["nivel"], cp["score"]), unsafe_allow_html=True)
                    with c_l3:
                        st.markdown(f"<span style='color:#A7B0BE; font-size:12px;'>{c['tamanho']} nós · {c['qtd_tels']} tel · {c['qtd_cpfs']} tit · {c['qtd_placas']} placas</span>", unsafe_allow_html=True)
                st.markdown("<hr style='margin:4px 0; border-color:#262B36;'/>", unsafe_allow_html=True)

    elif aba_triagem == "Base Mestra de Entidades Monitoradas":
        c_m_top1, c_m_top2 = st.columns([3.5, 1.5])
        with c_m_top1:
            st.markdown("### Base Mestra de Entidades Monitoradas")
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
                    "ID": s["id"], "Tipo": s["tipo"], "Dado Monitorado": s["valor_formatado"],
                    "Titular / Apelido": s["nome_referencia"] or "-", "Quadrilha / Caso": s["quadrilha_exibicao"],
                    "Status": s["status"], "Histórico Total": f"{s['total_historico']} assistências" if s.get('total_historico', 0) > 0 else "Nenhum",
                    "Data Cadastro": s["data_cadastro"], "Analista": s["analista"] or "-"
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

    elif aba_triagem == "Watchlist de Municípios de Risco":
        c_w_top1, c_w_top2 = st.columns([3.5, 1.5])
        with c_w_top1:
            st.markdown("### Watchlist de Municípios de Alto Risco")
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
                    "ID": c_item["id"], "Município": c_item["cidade"], "UF": c_item["uf"],
                    "Motivo / Modus Operandi": c_item["motivo"] or "Sem observação",
                    "Analista": c_item["analista"] or "Sistema", "Data Cadastro": c_item["data_cadastro"]
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

    else:
        st.markdown("### Centro de Comando Proativo e Radar de Anomalias Territoriais")
        st.caption("Detecção estatística de desvios volumétricos municipais, anomalias pet e abusos residenciais.")

        with st.spinner("Calculando desvios volumétricos e baselines históricos..."):
            df_anomalias, kpis_anomalias, mes_referencia = obter_radar_anomalias_macro()

        if df_anomalias.empty:
            st.info("É necessário ingerir arquivos em Criações Diárias para calcular os baselines territoriais.")
        else:
            c_ano1, c_ano2, c_ano3, c_ano4 = st.columns(4)
            with c_ano1: st.markdown(f"""<div class="kpi-card-mini"><span style="color:#7C8798; font-size:10px; font-weight:600;">MUNICÍPIOS</span> <b style="font-size:15px; color:#C0625F;">{kpis_anomalias.get('total_anomalias', 0)}</b></div>""", unsafe_allow_html=True)
            with c_ano2: st.markdown(f"""<div class="kpi-card-mini"><span style="color:#7C8798; font-size:10px; font-weight:600;">PET</span> <b style="font-size:15px; color:#C9A66B;">{kpis_anomalias.get('alertas_pet', 0)}</b></div>""", unsafe_allow_html=True)
            with c_ano3: st.markdown(f"""<div class="kpi-card-mini"><span style="color:#7C8798; font-size:10px; font-weight:600;">RESIDENCIAL</span> <b style="font-size:15px; color:#6C93B0;">{kpis_anomalias.get('alertas_res', 0)}</b></div>""", unsafe_allow_html=True)
            with c_ano4: st.markdown(f"""<div class="kpi-card-mini"><span style="color:#7C8798; font-size:10px; font-weight:600;">WATCHLIST</span> <b style="font-size:15px; color:#8F86B5;">{kpis_anomalias.get('alertas_watchlist', 0)}</b></div>""", unsafe_allow_html=True)

            st.markdown("<br/>", unsafe_allow_html=True)
            tab_radar_mapa, tab_radar_tabela = st.tabs(["Radar Territorial de Anomalias", "Fila de Anomalias por Município"])

            with tab_radar_mapa:
                df_mapa_ano = df_anomalias.dropna(subset=["latitude", "longitude"]).copy()
                if not df_mapa_ano.empty:
                    df_mapa_ano["raio_calc"] = df_mapa_ano["volume_atual"] * 400
                    camada_anomalias = pdk.Layer(
                        "ScatterplotLayer", data=df_mapa_ano, get_position="[longitude, latitude]",
                        get_color="[192, 98, 95, 190]", get_line_color="[255, 255, 255, 220]",
                        line_width_min_pixels=1.5, stroked=True, get_radius="raio_calc",
                        radius_min_pixels=7, radius_max_pixels=32, pickable=True, auto_highlight=True
                    )
                    lat_c = float(df_mapa_ano["latitude"].mean())
                    lon_c = float(df_mapa_ano["longitude"].mean())
                    view_state_ano = pdk.ViewState(latitude=lat_c, longitude=lon_c, zoom=4.2, pitch=0)
                    tooltip_ano = {
                        "html": "<div style='font-family:sans-serif; padding:5px;'><b style='color:#C0625F; font-size:13px;'>{cidade}/{uf}</b><br/><b>Volume Atual:</b> {volume_atual} acionamentos<br/><b>Média Histórica:</b> {media_hist}<br/><b>Desvio:</b> {desvio_macro}<br/><b>Padrões:</b> <span style='color:#C9A66B;'>{alertas_str}</span></div>",
                        "style": {"backgroundColor": "#10141C", "color": "#E5E9F0", "border": "1px solid #262B36", "borderRadius": "6px", "fontSize": "11px", "zIndex": "1000"}
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
                        "Município": r_ano.cidade, "UF": r_ano.uf, "Mês Ref.": r_ano.mes_ref,
                        "Volume Atual": f"{r_ano.volume_atual} acionamentos", "Média Histórica": f"{r_ano.media_hist}",
                        "Desvio Relativo": r_ano.desvio_macro, "Meses Ativos": r_ano.meses_ativos_hist,
                        "Pet (Mês)": r_ano.pet_atual, "Residencial (Mês)": r_ano.res_atual,
                        "Padrões Identificados": r_ano.alertas_str, "Watchlist": r_ano.watchlist
                    })
                df_tabela_final = pd.DataFrame(tabela_exibicao)

                linhas_ano_sel = []
                try:
                    evento_tab_ano = st.dataframe(df_tabela_final, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row", key="tab_anomalias_macro_view")
                    linhas_ano_sel = extrair_linhas_selecionadas(evento_tab_ano)
                except Exception:
                    st.dataframe(df_tabela_final, use_container_width=True, hide_index=True)

                if linhas_ano_sel:
                    idx_ano = linhas_ano_sel[0]
                    if 0 <= idx_ano < len(df_view_ano):
                        cid_sel_tabela = df_view_ano.iloc[idx_ano]["cidade"]
                        if st.session_state["municipio_foco_funil"] != cid_sel_tabela:
                            st.session_state["municipio_foco_funil"] = cid_sel_tabela

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
                dados_infratores = extrair_top_infratores_municipio(cidade=anomalia_alvo["cidade"], uf=anomalia_alvo["uf"], cidade_banco=anomalia_alvo.get("cidade_banco", ""))

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
                        resumo_funil = {
                            "alvo_buscado": f"Conluio Regional - {anomalia_alvo['cidade']}/{anomalia_alvo['uf']}",
                            "termo_limpo": anomalia_alvo["cidade"], "total_assistencias": dados_infratores["total_ocorrencias"],
                            "assistencias_diretas": dados_infratores["total_ocorrencias"],
                            "qtd_cpfs": len(dados_infratores["top_cpfs"]), "qtd_tels": len(dados_infratores["top_tels"]),
                            "qtd_placas": len(dados_infratores["top_placas"])
                        }
                        ents_funil = {
                            "cpfs": list(dados_infratores["top_cpfs"]["cpf"]) if not dados_infratores["top_cpfs"].empty else [],
                            "tels": list(dados_infratores["top_tels"]["telefone"]) if not dados_infratores["top_tels"].empty else [],
                            "placas": list(dados_infratores["top_placas"]["placa"]) if not dados_infratores["top_placas"].empty else []
                        }
                        abrir_descoberta_ativa(resumo_funil, dados_infratores["df_completo"], ents_funil)
                else:
                    st.info("Nenhuma ocorrência detalhada localizada para os parâmetros deste município.")