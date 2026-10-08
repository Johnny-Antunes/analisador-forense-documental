"""
Constantes, paleta forense e módulos opcionais (enrich / EXIF).
"""

from __future__ import annotations



try:
    from enrich_engine import enriquecer_entidade_local, CORES_POR_TIPO, PREFIXO_POR_TIPO
    TEM_ENRICH_ENGINE = True
except ImportError:
    TEM_ENRICH_ENGINE = False
    enriquecer_entidade_local = None
    CORES_POR_TIPO = {
        "cpf": {"bg": "#3C6FA8", "border": "#5A94D6"},
        "telefone": {"bg": "#B94A3C", "border": "#E0684F"},
        "placa": {"bg": "#6D4FA8", "border": "#9273D6"},
        "prestador": {"bg": "#4A7A5A", "border": "#6BA57D"},
        "empresa": {"bg": "#4A4A4A", "border": "#767676"}
    }
    PREFIXO_POR_TIPO = {"cpf": "CPF", "telefone": "TEL", "placa": "PLACA", "prestador": "PREST", "empresa": "EMP"}

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

LABELS_FILTRO_TIPO = {"TODOS": "Todos", "CPF": "CPF", "TELEFONE": "Telefone", "PLACA": "Placa"}
CLASSE_MENU_PADRAO = 'bg-[#18181A]/95 backdrop-blur-lg border border-[#2B2B2F] p-2 rounded-xl shadow-2xl'
