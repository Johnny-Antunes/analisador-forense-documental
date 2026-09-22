"""
Módulo: grafo_component.py
Objetivo: Interface Python para o Custom Component de Grafo Bidirecional.
          Transmite nós e arestas para o Vis.js e recebe eventos de clique
          no nó de volta para o Streamlit sem necessidade de build externo.
"""

from pathlib import Path
from typing import Dict, Any, List, Optional
import streamlit as st
import streamlit.components.v1 as components

# Diretório onde reside o index.html vendorizado
_DIRETORIO_COMPONENTE = Path(__file__).parent / "componente_grafo"

if _DIRETORIO_COMPONENTE.exists():
    _componente_func = components.declare_component(
        "lcfo_grafo_bidirecional",
        path=str(_DIRETORIO_COMPONENTE)
    )
else:
    _componente_func = None


def renderizar_grafo_bidirecional(
    nodes: List[Dict[str, Any]],
    edges: List[Dict[str, Any]],
    hub_id: str = "",
    target_node_id: str = "",
    layout_ativo: str = "organico",
    base_font_size: int = 12,
    espacamento: int = 280,
    height: int = 740,
    key: Optional[str] = None
) -> Optional[Dict[str, Any]]:
    """
    Renderiza o grafo interativo e captura eventos de seleção do nó para o Python.
    Retorna um dicionário {'tipo': 'node_click', 'nodeId': '...'} ou None.
    """
    if _componente_func is None:
        return None

    return _componente_func(
        nodes=nodes,
        edges=edges,
        hub_id=hub_id,
        target_node_id=target_node_id,
        layout_ativo=layout_ativo,
        base_font_size=base_font_size,
        espacamento=espacamento,
        height=height,
        key=key,
        default=None
    )