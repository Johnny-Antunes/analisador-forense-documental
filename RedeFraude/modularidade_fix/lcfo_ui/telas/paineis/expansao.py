"""
Painel da mesa: Expansão com Criações.
"""

from __future__ import annotations


from nicegui import ui

from lcfo_ui.servicos import cruzar_com_criacoes_diarias, anexar_descoberta_a_caso_existente
from lcfo_ui.dados import DADOS
from lcfo_ui.util_ui import sanitizar_df_para_tabela

from lcfo_ui.telas.contexto_mesa import ContextoMesa


def render(m: ContextoMesa, lado: str) -> None:
    ctx = m.ctx
    dossie_persistido = m.dossie_persistido
    cluster_obj = m.cluster_obj
    dados_caso = m.dados_caso
    nome_exib = m.nome_exib
    cpfs = m.cpfs
    tels = m.tels
    placas = m.placas
    with ui.column().classes('w-full h-full pt-14 px-8 overflow-y-auto custom-scroll'):
        if not dossie_persistido:
            ui.label('Expansão habilitada em sketches oficiais autuados. Use o painel superior para promover esta descoberta.').classes('text-xs text-[#71717A]')
        elif not cluster_obj:
            ui.label('Este sketch ainda não tem uma célula oficial vinculada — não é possível cruzar com Criações Diárias.').classes('text-xs text-[#71717A]')
        else:
            res_int, df_m, df_susp = cruzar_com_criacoes_diarias(cpfs, tels, placas)
            if res_int is None:
                ui.label('Criações Diárias ainda não ingeridas.').classes('text-[#71717A]')
            elif df_susp.empty:
                ui.label('Nenhum suspeito inédito detectado nas Criações Diárias.').classes('text-[#71717A]')
            else:
                ui.label('Novos Suspeitos Detectados').classes('text-sm font-bold text-[#FF7300] mb-2')
                ui.table.from_pandas(sanitizar_df_para_tabela(df_susp)).classes('w-full bg-[#18181A] border border-[#2B2B2F] text-xs mb-3 rounded-xl')
                nome_dossie_atual = (dados_caso.get("nome_personalizado") if isinstance(dados_caso, dict) else None) or nome_exib
                ui.button('Incorporar Suspeitos ao Sketch', on_click=lambda: (
                    anexar_descoberta_a_caso_existente(df_m, cluster_obj["id"], nome_quadrilha_override=nome_dossie_atual),
                    ui.notify('Incorporado!'),
                    DADOS.recarregar(),
                    ctx.workspace.refresh(),
                )).classes('bg-[#FF7300] hover:bg-[#E0670A] text-white text-xs font-semibold h-8 rounded-md')
