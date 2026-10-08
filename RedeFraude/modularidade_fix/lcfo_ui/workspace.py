"""
Workspace principal: decide qual tela desenhar a partir do estado de sessão.
Cada tela vive no seu módulo em lcfo_ui/telas/.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from nicegui import ui

from lcfo_ui.telas import base_mestra, casos, centro_comando, mesa, overview, watchlist

if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina


def criar_workspace(ctx: "Pagina") -> None:
    """Publica ctx.workspace (refreshable)."""
    st = ctx.st

    @ui.refreshable
    def workspace():
        if (st.sketch_ativo_id or st.dados_descoberta) and st.subtela_caso == "mesa":
            mesa.render(ctx)
        elif st.caso_ativo_id and st.subtela_caso == "overview":
            overview.render(ctx)
        elif st.tela_ativa == "Casos":
            casos.render(ctx)
        elif st.tela_ativa == "Base Mestra":
            base_mestra.render(ctx)
        elif st.tela_ativa == "Watchlist":
            watchlist.render(ctx)
        else:
            centro_comando.render(ctx)

    ctx.workspace = workspace
