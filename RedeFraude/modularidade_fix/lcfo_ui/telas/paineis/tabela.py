"""
Painel da mesa: Tabela de Ocorrências.
"""

from __future__ import annotations


from nicegui import ui

from lcfo_ui.util_ui import sanitizar_df_para_tabela

from lcfo_ui.telas.contexto_mesa import ContextoMesa


def render(m: ContextoMesa, lado: str) -> None:
    df_dados = m.df_dados
    with ui.column().classes('w-full h-full pt-14 px-8 overflow-y-auto custom-scroll'):
        cols = [c for c in ["data", "id_assistencia", "titular", "cpf", "telefone", "placa", "servico", "cidade", "uf", "empresa_cliente"] if c in df_dados.columns]
        if cols and not df_dados.empty:
            with ui.element('div').classes('w-full').style('max-height: calc(100vh - 220px); overflow: auto;'):
                ui.table.from_pandas(sanitizar_df_para_tabela(df_dados[cols]), pagination=15).classes(
                    'w-full bg-[#18181A] border border-[#2B2B2F] text-xs rounded-xl'
                )
        else:
            ui.label('Sem ocorrências para exibir.').classes('text-[#71717A]')
