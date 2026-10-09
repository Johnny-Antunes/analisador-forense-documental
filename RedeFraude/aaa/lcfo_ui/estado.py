"""
Estado de sessão (um por aba/cliente do navegador).
"""

from __future__ import annotations

from typing import Any, Dict, Optional, Set


# =========================================================================
# ESTADO DE SESSÃO
# =========================================================================
class SessionState:
    def __init__(self):
        self.tela_ativa = "Casos"
        self.caso_ativo_id: Optional[str] = None
        self.sketch_ativo_id: Optional[str] = None
        self.subtela_caso = "overview"
        self.mostrar_notas = False
        self.mostrar_form_nova_analise = False
        self.sidebar_visivel = True
        self.dados_descoberta: Optional[Dict[str, Any]] = None
        self.entidade_foco: Optional[str] = None
        self.layout_por_caso: Dict[str, str] = {}
        self.nos_enriquecidos: Dict[str, Dict[str, Any]] = {}
        self.selecao_entidades: Dict[str, Set[str]] = {}
        self.contexto_grafo_ativo: Dict[str, Dict[str, Any]] = {}

        self.aba_painel_lateral = "Entities"
        self.filtro_tipo_entidade = "TODOS"
        self.painel_unico = "Grafo de Vínculos"

        self.cockpit_ativo = False
        self.cockpit_proporcao = 50.0
        self.painel_esquerdo = "Grafo de Vínculos"
        self.painel_direito = "Radar Territorial (Mapa)"

        self.filtro_status_dash = "Todos"
        self.filtro_batismo_dash = "Todos"
        self.filtro_risco_dash = "Todos"
        self.modo_visualizacao_casos = "Grade"
        self.termo_busca_dash = ""
        self.municipio_foco_funil: Optional[str] = None
        self.mapa_lista_aberta = True                 # painel flutuante "Municípios do caso" do mapa da mesa
        # Sketches/descobertas cuja mesa já foi preparada nesta sessão (caches quentes).
        self.mesas_aquecidas: Set[Any] = set()

        # Centro de Comando (cockpit territorial)
        self.radar_mes: Optional[int] = None          # índice do mês na linha do tempo
        self.radar_uf: Optional[str] = None
        self.radar_mun: Optional[str] = None          # código IBGE do município aberto
        self.radar_visao = "mapa"                     # "mapa" | "tabela"
        self.radar_ordem = "criticos"                 # "criticos" | "volume" | "az"
        self.radar_filtro = 0                         # bit do alerta (0 = todos)
        self.radar_metrica = "alerta"
        self.radar_unidade = "municipios"
        self.radar_esq_aberto = True                  # painel de vidro da evolução (esquerda)
        self.radar_dir_aberto = True                  # painel de vidro do ranking/praça (direita)
