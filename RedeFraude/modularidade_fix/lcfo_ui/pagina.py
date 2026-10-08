"""
Página principal ('/'). Monta a página na mesma ordem do antigo main()
monolítico, delegando cada parte ao seu módulo.
"""

from __future__ import annotations

from nicegui import ui

from lcfo_ui.analises import criar_painel_analises
from lcfo_ui.contexto import Pagina
from lcfo_ui.dialogos import criar_dialogos_globais, criar_dialogos_sketch
from lcfo_ui.js_globais import (
    instalar_bloqueio_atalhos_navegador, instalar_ponte_grafo_global,
    instalar_resize_notas_global, instalar_resize_sidebar_global,
)
from lcfo_ui.navbar import criar_navbar
from lcfo_ui.navegacao import criar_navegacao
from lcfo_ui.ponte_grafo import instalar_receptor_grafo
from lcfo_ui.rail import criar_rail
from lcfo_ui.sidebar import criar_sidebar
from lcfo_ui.workspace import criar_workspace


@ui.page('/')
def main():
    ctx = Pagina()
    instalar_ponte_grafo_global()
    instalar_resize_notas_global()
    instalar_resize_sidebar_global()
    instalar_bloqueio_atalhos_navegador()

    instalar_receptor_grafo(ctx)
    criar_navegacao(ctx)
    criar_dialogos_globais(ctx)

    # -------------------------------------------------------------------
    # ATALHOS DE TECLADO
    # -------------------------------------------------------------------
    def _tratar_atalho_teclado(e):
        if not e.action.keydown:
            return
        if not (e.modifiers.ctrl or e.modifiers.meta):
            return
        if e.key == 'j':
            ctx.dlg_palette.open()
        elif e.key == 'b':
            ctx.alternar_sidebar()
        elif e.key == 'l':
            ctx.alternar_notas()

    ui.keyboard(on_key=_tratar_atalho_teclado, ignore=[])

    criar_navbar(ctx)
    criar_rail(ctx)
    criar_painel_analises(ctx)
    criar_sidebar(ctx)
    criar_dialogos_sketch(ctx)
    criar_workspace(ctx)

    # -------------------------------------------------------------------
    # MONTAGEM FINAL: GAVETA LATERAL MAIOR + WORKSPACE
    # -------------------------------------------------------------------
    with ui.row().classes('w-full h-full flex-nowrap gap-0'):
        ctx.sidebar_container()
        with ui.column().classes('flex-1 h-full overflow-hidden'):
            ctx.workspace()
