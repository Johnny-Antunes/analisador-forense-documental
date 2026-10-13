"""
Constantes, paleta forense e módulos opcionais (enrich / EXIF).
"""

from __future__ import annotations



from lcfo_dominio import tipos

# Cores e prefixos por tipo: registro único (lcfo_dominio/tipos.py); aliases para o código antigo.
CORES_POR_TIPO = {k: t.cores for k, t in tipos.REGISTRO.items()}
PREFIXO_POR_TIPO = {k: t.prefixo for k, t in tipos.REGISTRO.items()}

try:
    from enrich_engine import enriquecer_entidade_local
    TEM_ENRICH_ENGINE = True
except ImportError:
    TEM_ENRICH_ENGINE = False
    enriquecer_entidade_local = None

try:
    from exif_engine import extrair_metadados_foto, validar_coerencia_geografica_foto
    TEM_EXIF_ENGINE = True
except ImportError:
    TEM_EXIF_ENGINE = False
    extrair_metadados_foto = None
    validar_coerencia_geografica_foto = None

# =========================================================================
# CONSTANTES & PALETA FORENSE
# =========================================================================
FONT_PADRAO = 11
ESPACAMENTO_PADRAO = 280
TOLERANCIA_KM_EXIF = 60.0

CORES_STATUS = {
    "Em Investigação": "#6C93B0", "Confirmado Fraude": "#C0625F", "Monitoramento Contínuo": "#8F86B5",
    "Sem Irregularidade Identificada": "#6FA98A", "Falso Positivo": "#8B95A5", "Arquivado": "#5B6472",
}
CORES_RISCO = {"BAIXO": "#6FA98A", "MÉDIO": "#C9A66B", "ALTO": "#C98756", "CRÍTICO": "#C0625F"}

STATUS_ATIVOS = {"Em Investigação", "Confirmado Fraude", "Monitoramento Contínuo"}
STATUS_ENCERRADOS = {"Sem Irregularidade Identificada", "Falso Positivo", "Arquivado"}

UFS_BRASIL = ["AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG",
              "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO"]

OPCOES_TOOLBAR = [
    "Grafo de Vínculos", "Tabela de Ocorrências", "Expansão com Criações",
    "Radar Territorial (Mapa)", "Evolução Temporal", "Ferramentas"
]
ICONE_TOOLBAR = {
    "Grafo de Vínculos": "hub", "Tabela de Ocorrências": "table_chart",
    "Expansão com Criações": "swap_horiz", "Radar Territorial (Mapa)": "location_on",
    "Evolução Temporal": "timeline", "Ferramentas": "tune"
}

# Altura (px) ocupada pelo HUD flutuante da mesa, medida a partir do topo da área de trabalho.
# Os painéis de tela cheia (grafo, mapa) ficam ATRÁS do HUD, como o grafo sempre fez; este valor
# diz ao mapa onde termina a área livre, para o enquadramento e os controles não ficarem por baixo.
ALTURA_HUD_MESA = 112          # HUD + dock de painéis (um painel só)
ALTURA_HUD_COCKPIT = 60        # HUD apenas (tela dividida: o dock fica oculto)
LARGURA_PAINEL_MAPA = 280      # painel flutuante "Municípios do caso"
ALTURA_PROMOVER = 52          # barra "Promover a Novo Caso..." (só na mesa de uma descoberta)
ALTURA_CHIP_SPLIT = 40        # chip de seleção de painel de cada metade da tela dividida
PAINEIS_TELA_CHEIA = {"Grafo de Vínculos", "Radar Territorial (Mapa)"}   # passam por trás do HUD

LABELS_FILTRO_TIPO = {"TODOS": "Todos", **{k: t.rotulo for k, t in tipos.REGISTRO.items()}}
CLASSE_MENU_PADRAO = 'bg-[#18181A]/95 backdrop-blur-lg border border-[#2B2B2F] p-2 rounded-xl shadow-2xl'
