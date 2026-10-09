"""
Dados já calculados da mesa (sketch ou descoberta) repassados a cada painel.

A mesa calcula tudo uma vez (df, payload do grafo, score...) e os painéis
apenas leem daqui — para adicionar um painel novo, crie um módulo em
telas/paineis/ e registre-o em telas/paineis/__init__.py.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any, Dict, List, Optional

import pandas as pd

if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina


@dataclass
class ContextoMesa:
    ctx: "Pagina"
    identificador_caso: str
    dossie_persistido: bool
    sketch_ativo: Optional[Dict[str, Any]]
    cluster_obj: Optional[Dict[str, Any]]
    dados_caso: Dict[str, Any]
    nome_exib: str
    df_dados: pd.DataFrame
    resumo_descoberta: Optional[Dict[str, Any]]
    cpfs: List[str]
    tels: List[str]
    placas: List[str]
    chave_layout: str
    layout_ativo: str
    payload: Dict[str, Any]
    vis_nodes: List[Dict[str, Any]]
    vis_edges: List[Dict[str, Any]]
    hub_id: Optional[str]
    # Altura (px) do topo coberta pelo HUD flutuante (+ dock / barra Promover / chip da tela dividida).
    # Painéis de tela cheia passam por trás e só posicionam câmera/controles abaixo disso.
    topo_livre: int = 0
