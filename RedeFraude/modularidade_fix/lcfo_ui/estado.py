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
