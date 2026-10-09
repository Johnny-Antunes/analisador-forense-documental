"""
Top navbar (44px): breadcrumbs, busca e ações da mesa.
"""

from __future__ import annotations


from nicegui import ui

from lcfo_ui.config import CLASSE_MENU_PADRAO
from lcfo_ui.servicos import listar_sketches_do_caso, carregar_sketch
from lcfo_ui.dados import DADOS

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina


def criar_navbar(ctx: "Pagina") -> None:
    """Monta o header e publica os refreshables em ctx."""
    st = ctx.st

    # -------------------------------------------------------------------
    # TOP NAVBAR (44px)
    # -------------------------------------------------------------------
    with ui.header().classes('h-[44px] leading-none bg-[#121214]/98 backdrop-blur-md border-b border-[#2B2B2F] px-4 flex items-center z-50'):
        with ui.row().classes('items-center gap-1.5 flex-1 justify-start h-full no-wrap'):
            ui.icon('hub', size='16px').classes('text-[#FF7300]')
            ui.label('LCFO').classes('font-black text-sm text-[#FF7300] cursor-pointer tracking-wider leading-none').on('click', ctx.voltar_ao_dashboard)
            ui.label('/').classes('text-[#71717A] text-xs font-light leading-none')

            @ui.refreshable
            def navbar_breadcrumbs():
                if st.caso_ativo_id:
                    inf = DADOS.casos_cadastrados.get(st.caso_ativo_id, {})
                    nm = inf.get('nome') or st.caso_ativo_id
                    with ui.button(f"{nm[:14]} ⌵", color=None).props('flat dense no-caps').classes('text-xs text-gray-300 hover:text-white leading-none h-7 px-2'):
                        with ui.menu().classes(f'{CLASSE_MENU_PADRAO} max-h-64 overflow-y-auto custom-scroll min-w-[220px]'):
                            for c in DADOS.cluster_info[:50]:
                                n_lbl = DADOS.casos_cadastrados.get(c['id'], {}).get('nome') or c['hub_label'][:20]
                                with ui.menu_item(on_click=lambda cid=c['id']: ctx.abrir_caso_overview(cid)).classes(
                                    'text-xs rounded-md hover:bg-[#222225] flex items-center gap-2'
                                ):
                                    ui.icon('folder', size='14px').classes('text-[#71717A]')
                                    ui.label(n_lbl).classes('truncate flex-1 text-gray-200')
                    ui.label('/').classes('text-[#71717A] text-xs font-light leading-none')

                    rotulo_secundario = "Overview"
                    if st.sketch_ativo_id:
                        sk_cur = carregar_sketch(st.sketch_ativo_id)
                        if sk_cur:
                            rotulo_secundario = sk_cur["nome_sketch"]

                    with ui.button(f"{rotulo_secundario[:16]} ⌵", color=None).props('flat dense no-caps').classes('text-xs text-[#FF7300] leading-none h-7 px-2'):
                        with ui.menu().classes(f'{CLASSE_MENU_PADRAO} min-w-[220px]'):
                            with ui.menu_item(on_click=lambda: (
                                setattr(st, 'subtela_caso', 'overview'),
                                setattr(st, 'sketch_ativo_id', None),
                                ctx.sidebar_container.refresh(),
                                ctx.workspace.refresh()
                            )).classes('text-xs rounded-md hover:bg-[#222225] flex items-center gap-2'):
                                ui.icon('home', size='15px').classes('text-gray-300')
                                ui.label('Voltar ao Overview').classes('text-gray-200 font-medium')

                            sketches_caso = listar_sketches_do_caso(st.caso_ativo_id)
                            if sketches_caso:
                                ui.separator().classes('bg-[#2B2B2F] my-1')
                                ui.label('Sketches').classes('text-[9px] font-semibold text-[#71717A] px-2.5 py-1 tracking-wider uppercase')
                                for sk_item in sketches_caso:
                                    ativo_sk = (st.sketch_ativo_id == sk_item["id_sketch"])
                                    with ui.menu_item(on_click=lambda sid=sk_item['id_sketch']: ctx.abrir_mesa_sketch(sid)).classes(
                                        f"text-xs rounded-md hover:bg-[#222225] flex items-center gap-2 {'text-[#FF7300] font-semibold' if ativo_sk else 'text-gray-300'}"
                                    ):
                                        ui.icon('hub', size='13px').classes('text-[#71717A]')
                                        ui.label(sk_item["nome_sketch"]).classes('truncate flex-1')
                else:
                    ui.label(st.tela_ativa).classes('text-xs font-semibold text-gray-300 leading-none')
            navbar_breadcrumbs()

        with ui.row().classes('items-center justify-center flex-1 h-full no-wrap'):
            ui.button('Buscar no LCFO (Ctrl+J)', icon='search', color=None, on_click=ctx.dlg_palette.open).props('dense flat no-caps').classes(
                'top-search-btn'
            )

        with ui.row().classes('items-center justify-end gap-3 flex-1 h-full no-wrap'):
            @ui.refreshable
            def navbar_acoes():
                # Mesa de sketch ou de descoberta: a tela dividida vale para as duas; as notas
                # (pareceres) só existem para casos já autuados.
                if (st.caso_ativo_id or st.dados_descoberta) and st.subtela_caso == "mesa":
                    with ui.row().classes('items-center gap-2 h-7 my-auto no-wrap'):
                        if st.caso_ativo_id:
                            ui.switch('Notas (Ctrl+L)', value=st.mostrar_notas, on_change=lambda e: (
                                setattr(st, 'mostrar_notas', e.value), ctx.workspace.refresh()
                            )).props('dense color=orange').classes('text-xs text-gray-300 my-0')
                            ui.separator().props('vertical').classes('bg-[#2B2B2F] mx-1 h-4 self-center')
                        ui.button(icon='vertical_split', color=None, on_click=lambda: (
                            setattr(st, 'cockpit_ativo', not st.cockpit_ativo), ctx.workspace.refresh()
                        )).props('flat dense round').classes(
                            'text-[#FF7300]' if st.cockpit_ativo else 'text-gray-400 hover:text-white'
                        ).tooltip('Dividir tela (Cockpit)')
            navbar_acoes()

    ctx.navbar_breadcrumbs = navbar_breadcrumbs
    ctx.navbar_acoes = navbar_acoes
