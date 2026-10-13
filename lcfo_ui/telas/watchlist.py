"""
Watchlist de municípios de alto risco.
"""

from __future__ import annotations


import pandas as pd
from nicegui import ui

from lcfo_ui.config import UFS_BRASIL
from lcfo_ui.servicos import cadastrar_cidade_risco, listar_cidades_risco, remover_cidade_risco
from lcfo_ui.dados import DADOS
from lcfo_ui.util_ui import sanitizar_df_para_tabela

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina


def render(ctx: "Pagina") -> None:
    with ui.column().classes('w-full h-full overflow-y-auto custom-scroll'):
        with ui.column().classes('w-full max-w-7xl mx-auto p-8 gap-4'):
            ui.label('Watchlist de Municípios de Alto Risco').classes('text-xl font-bold text-white tracking-tight')
            ui.label('Cadastro de praças com predominância de fraude. Alimenta multiplicadores do Fraud Score.').classes('text-xs text-[#71717A]')

            with ui.expansion('Cadastrar Novo Município na Watchlist', icon='add_location').classes('w-full bg-[#141416] border border-[#2B2B2F] rounded-lg'):
                with ui.row().classes('w-full gap-3 mt-2'):
                    ipt_cidade_nova = ui.input(label='Nome do Município:', placeholder='Ex: Santa Quitéria').props('dense outlined').classes('flex-1 text-xs')
                    sel_uf_nova = ui.select(UFS_BRASIL, value='CE', label='UF:').props('dense outlined').classes('w-28 text-xs')
                with ui.row().classes('w-full gap-3'):
                    ipt_motivo_cid = ui.input(label='Motivo / Padrão Mapeado:', placeholder='Ex: Frequência atípica de guinchos e colusão regional').props('dense outlined').classes('flex-1 text-xs')
                    ipt_analista_cid = ui.input(label='Analista Responsável:', placeholder='Seu nome').props('dense outlined').classes('w-56 text-xs')

                def _salvar_cidade():
                    ok, msg = cadastrar_cidade_risco(ipt_cidade_nova.value, sel_uf_nova.value, ipt_motivo_cid.value, ipt_analista_cid.value)
                    if ok:
                        DADOS.recarregar()
                        ui.notify(msg)
                        ctx.workspace.refresh()
                    else:
                        ui.notify(msg, type="negative")

                ui.button('Salvar Município na Watchlist', on_click=_salvar_cidade).classes('w-full bg-[#FF7300] hover:bg-[#E0670A] text-white text-xs font-semibold h-8 rounded-md mt-3')

            cidades = listar_cidades_risco()
            if cidades:
                ui.label(f"Exibindo {len(cidades)} municípios monitorados.").classes('text-xs text-[#71717A]')
                df_c = pd.DataFrame(cidades)[['id', 'cidade', 'uf', 'motivo', 'analista', 'data_cadastro']]
                df_c.columns = ['ID', 'Município', 'UF', 'Motivo/Modus Operandi', 'Analista', 'Data Cadastro']
                with ui.element('div').classes('w-full').style('max-height: calc(100vh - 460px); overflow: auto;'):
                    ui.table.from_pandas(sanitizar_df_para_tabela(df_c), pagination=15).classes('w-full bg-[#18181A] border border-[#2B2B2F] text-xs rounded-xl')

                with ui.expansion('Excluir Município da Watchlist', icon='delete').classes('w-full bg-[#141416] border border-[#2B2B2F] rounded-lg'):
                    opcoes_del = {c['id']: f"#{c['id']} - {c['cidade']}/{c['uf']}" for c in cidades}
                    sel_del_cid = ui.select(opcoes_del, label='Selecione o município para remover:').props('dense outlined').classes('w-full text-xs mt-2')

                    def _remover_cidade():
                        if not sel_del_cid.value: return
                        remover_cidade_risco(sel_del_cid.value)
                        DADOS.recarregar()
                        ui.notify("Município removido.")
                        ctx.workspace.refresh()

                    ui.button('Remover Município Selecionado', on_click=_remover_cidade).classes('bg-red-900/60 hover:bg-red-800 text-white text-xs h-8 rounded-md px-3 border border-red-700/50 mt-2')
            else:
                ui.label('Nenhum município cadastrado na Watchlist de Risco.').classes('text-[#71717A]')
