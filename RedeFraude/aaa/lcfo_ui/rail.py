"""
Left rail fixo (56px) com as telas principais.
"""

from __future__ import annotations


from nicegui import ui

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina


def criar_rail(ctx: "Pagina") -> None:
    """Monta o rail esquerdo."""
    st = ctx.st

    # -------------------------------------------------------------------
    # LEFT RAIL FIXO (56px)
    # -------------------------------------------------------------------
    with ui.left_drawer(value=True, fixed=True).props(':width="56"').classes(
        'w-[56px] bg-[#0F0F10] border-r border-[#2B2B2F] p-2 flex flex-col items-center justify-between z-50'
    ):
        # Refreshable: antes o rail era desenhado uma vez só e o destaque ficava preso na tela
        # inicial ("Casos"); agora a navegação (navegacao.py) o redesenha a cada troca de tela.
        @ui.refreshable
        def botoes_telas():
            with ui.column().classes('items-center gap-2.5 w-full mt-1'):
                for icone, nome, tip in [
                    ("folder_open", "Casos", "Células e redes da Blacklist"),
                    ("hub", "Base Mestra", "Base Mestra de entidades"),
                    ("warning", "Watchlist", "Watchlist de municípios de risco"),
                    ("dashboard", "Centro de Comando", "Centro de Comando e Radar Territorial"),
                ]:
                    # Dentro de um caso/descoberta, a seção de origem continua destacada.
                    ativo = st.tela_ativa == nome
                    ui.button(icon=icone, color=None, on_click=lambda n=nome: (setattr(st, 'tela_ativa', n), ctx.voltar_ao_dashboard())).props(
                        'flat dense round'
                    ).classes(
                        f"w-9 h-9 rounded-lg {'bg-[#FF7300]/15 text-[#FF7300]' if ativo else 'text-[#71717A] hover:text-white hover:bg-[#18181A]'}"
                    ).tooltip(tip)

        botoes_telas()
        ctx.rail = botoes_telas

        with ui.column().classes('items-center gap-2 w-full pb-2'):
            ui.button(icon='storage', color=None, on_click=ctx.dlg_ingestao.open).props(
                'flat dense round'
            ).classes('w-9 h-9 rounded-lg text-[#71717A] hover:text-white hover:bg-[#18181A]').tooltip('Gestão de bases & ingestão')

            ui.button(icon='menu_open', color=None, on_click=ctx.alternar_sidebar).props(
                'flat dense round'
            ).classes('w-9 h-9 rounded-lg text-[#71717A] hover:text-[#FF7300] hover:bg-[#18181A]').tooltip('Alternar painel lateral (Ctrl+B)')
