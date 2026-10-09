"""
Contexto de uma página aberta (uma aba do navegador).

Antes, tudo vivia como closures dentro de main(); agora cada módulo recebe
`ctx` e acessa por ele o estado (ctx.st), os refreshables (ctx.workspace,
ctx.sidebar_container, ...) e as funções de navegação. Os atributos são
preenchidos pelas fábricas em lcfo_ui/pagina.py, na mesma ordem do main()
original.
"""

from __future__ import annotations

from typing import Any, Callable, Dict, Optional

from lcfo_ui.estado import SessionState


class Pagina:
    def __init__(self) -> None:
        self.st = SessionState()

        # Refreshables (criados por sidebar.py, navbar.py, workspace.py)
        self.workspace: Any = None
        self.sidebar_container: Any = None
        self.navbar_breadcrumbs: Any = None
        self.navbar_acoes: Any = None
        self.rail: Any = None                 # botões das telas no rail esquerdo (rail.py)

        # Diálogos globais (dialogos.py)
        self.dlg_palette: Any = None
        self.dlg_ingestao: Any = None

        # Navegação (navegacao.py)
        self.abrir_caso_overview: Optional[Callable[[str], None]] = None
        self.abrir_mesa_sketch: Optional[Callable[[str], None]] = None
        self.voltar_ao_dashboard: Optional[Callable[[], None]] = None
        self.abrir_descoberta: Optional[Callable[..., None]] = None
        self.alternar_notas: Optional[Callable[[], None]] = None
        self.alternar_sidebar: Optional[Callable[[], None]] = None

        # Componentes compartilhados entre telas
        self.renderizar_painel_analises: Optional[Callable[..., None]] = None
        self.abrir_dialog_novo_sketch: Optional[Callable[[str], None]] = None
        self.abrir_confirmacao_remocao_sketch: Optional[Callable[..., None]] = None

        # Receptores de mensagens de iframes (chave -> callback(valor)), além do grafo.
        self.receptores_iframe: Dict[str, Callable[[Dict[str, Any]], None]] = {}
