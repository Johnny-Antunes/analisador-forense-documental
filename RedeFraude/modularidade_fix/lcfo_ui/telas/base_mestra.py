"""
Base Mestra de entidades monitoradas.
"""

from __future__ import annotations


import pandas as pd
from nicegui import ui

from lcfo_ui.servicos import (
    cadastrar_entidade_suspeita, listar_entidades_suspeitas, remover_entidade_suspeita, consultar_todas_ocorrencias_entidade, semear_base_mestra_da_blacklist,
)
from lcfo_ui.util_ui import sanitizar_df_para_tabela

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina


def render(ctx: "Pagina") -> None:
    with ui.column().classes('w-full h-full overflow-y-auto custom-scroll'):
        with ui.column().classes('w-full max-w-7xl mx-auto p-8 gap-4'):
            with ui.row().classes('w-full items-center justify-between'):
                with ui.column().classes('gap-1'):
                    ui.label('Base Mestra de Entidades Monitoradas').classes('text-xl font-bold text-white tracking-tight')
                    ui.label('Repositório permanente de suspeitos. Varre retrospectivamente todo o acervo de Criações Diárias.').classes('text-xs text-[#71717A]')
                ui.button('Sincronizar com a Blacklist', icon='sync', on_click=lambda: (semear_base_mestra_da_blacklist(), ui.notify('Sincronizado!'), ctx.workspace.refresh())).classes('bg-[#222225] hover:bg-[#2B2B2F] text-gray-200 text-xs px-3 h-8 rounded-md border border-[#2B2B2F]')

            with ui.expansion('Cadastrar Nova Entidade Suspeita', icon='person_add').classes('w-full bg-[#141416] border border-[#2B2B2F] rounded-lg'):
                with ui.row().classes('w-full gap-3 mt-2'):
                    sel_novo_tipo = ui.select(["TELEFONE", "PLACA", "CPF"], value="TELEFONE", label='Tipo de Entidade:').props('dense outlined').classes('flex-1 text-xs')
                    ipt_novo_val = ui.input(label='Dado / Número / Placa:', placeholder='Ex: (11) 97000-1122 ou ABC-1234').props('dense outlined').classes('flex-1 text-xs')
                    ipt_novo_nome = ui.input(label='Nome / Apelido / Titular:', placeholder='Ex: Marcos V. (Laranja)').props('dense outlined').classes('flex-1 text-xs')
                with ui.row().classes('w-full gap-3'):
                    ipt_novo_caso = ui.input(label='Quadrilha / Operação Associada:', placeholder='Ex: Quadrilha Santo André').props('dense outlined').classes('flex-1 text-xs')
                    sel_novo_status = ui.select(["Ativo", "Confirmado Fraude", "Em Monitoramento", "Inativo"], value="Ativo", label='Status:').props('dense outlined').classes('flex-1 text-xs')
                    ipt_novo_analista = ui.input(label='Analista Responsável:', placeholder='Seu nome').props('dense outlined').classes('flex-1 text-xs')
                txt_novo_motivo = ui.textarea(label='Motivo da Inclusão / Modus Operandi:', placeholder='Ex: Solicitou guinchos sequenciais para o mesmo destino...').props('dense outlined').classes('w-full text-xs mt-2')

                def _salvar_suspeito():
                    ok, msg, total_hits = cadastrar_entidade_suspeita(
                        sel_novo_tipo.value, ipt_novo_val.value, ipt_novo_nome.value,
                        txt_novo_motivo.value, ipt_novo_caso.value,
                        status=sel_novo_status.value, analista=ipt_novo_analista.value
                    )
                    if ok:
                        ui.notify(f"{msg} Varredura retrospectiva: {total_hits} assistências identificadas.")
                        ctx.workspace.refresh()
                    else:
                        ui.notify(msg, type="negative")

                ui.button('Salvar na Base Mestra & Varrer Histórico', on_click=_salvar_suspeito).classes('w-full bg-[#FF7300] hover:bg-[#E0670A] text-white text-xs font-semibold h-8 rounded-md mt-3')

            with ui.row().classes('w-full gap-3 items-end'):
                sel_f_tipo = ui.select(["TODOS", "TELEFONE", "PLACA", "CPF"], value="TODOS", label='Filtrar Tipo:').props('dense outlined').classes('w-40 text-xs')
                sel_f_status = ui.select(["TODOS", "Ativo", "Confirmado Fraude", "Em Monitoramento", "Inativo", "Removido"], value="TODOS", label='Filtrar Status:').props('dense outlined').classes('w-52 text-xs')
                ipt_f_busca = ui.input(label='Buscar na Base Mestra:', placeholder='Digite dado, nome ou quadrilha...').props('dense outlined').classes('flex-1 text-xs')

            box_mestra = ui.column().classes('w-full gap-2')

            def _atualizar_base_mestra():
                box_mestra.clear()
                suspeitos = listar_entidades_suspeitas(sel_f_tipo.value, sel_f_status.value, ipt_f_busca.value)
                with box_mestra:
                    if suspeitos:
                        ui.label(f"Exibindo {len(suspeitos)} entidades suspeitas catalogadas.").classes('text-xs text-[#71717A]')
                        df_m = pd.DataFrame(suspeitos)[['id', 'tipo', 'valor_formatado', 'nome_referencia', 'quadrilha_exibicao', 'status', 'total_historico', 'data_cadastro', 'analista']]
                        df_m.columns = ['ID', 'Tipo', 'Dado Monitorado', 'Titular/Apelido', 'Quadrilha/Caso', 'Status', 'Histórico', 'Data Cadastro', 'Analista']
                        with ui.element('div').classes('w-full').style('max-height: calc(100vh - 340px); overflow: auto;'):
                            ui.table.from_pandas(sanitizar_df_para_tabela(df_m), pagination=15).classes('w-full bg-[#18181A] border border-[#2B2B2F] text-xs rounded-xl')

                        with ui.expansion('Auditar / Remover Entidade Específica', icon='search').classes('w-full bg-[#141416] border border-[#2B2B2F] rounded-lg'):
                            opcoes_ent = {s['id']: f"[{s['tipo']}] {s['valor_formatado']} - {s['nome_referencia']} ({s['status']})" for s in suspeitos}
                            sel_ent_audit = ui.select(opcoes_ent, label='Selecione a entidade para auditoria:').props('dense outlined').classes('w-full text-xs mt-2')
                            ipt_motivo_rem = ui.input(label='Motivo da remoção/arquivamento:', placeholder='Ex: Terceiro inocente comprovado').props('dense outlined').classes('w-full text-xs mt-2')

                            resultado_audit_box = ui.column().classes('w-full gap-2 mt-3')

                            with ui.row().classes('gap-2 mt-2'):
                                def _executar_remocao():
                                    if not sel_ent_audit.value:
                                        return
                                    remover_entidade_suspeita(sel_ent_audit.value, ipt_motivo_rem.value)
                                    ui.notify("Entidade marcada como 'Removido' (cadeia de custódia preservada).")
                                    _atualizar_base_mestra()

                                ui.button('Excluir Entidade', on_click=_executar_remocao).classes('bg-red-900/60 hover:bg-red-800 text-white text-xs h-8 rounded-md px-3 border border-red-700/50')

                                def _carregar_ocorrencias():
                                    resultado_audit_box.clear()
                                    if not sel_ent_audit.value:
                                        return
                                    ent_alvo = next((s for s in suspeitos if s['id'] == sel_ent_audit.value), None)
                                    if not ent_alvo:
                                        return
                                    df_ocorr = consultar_todas_ocorrencias_entidade(ent_alvo['tipo'], ent_alvo['valor'])
                                    with resultado_audit_box:
                                        if not df_ocorr.empty:
                                            ui.label(f"{len(df_ocorr)} Assistências Encontradas (Blacklist + Criações Diárias):").classes('text-xs font-bold text-gray-300')
                                            ui.table.from_pandas(sanitizar_df_para_tabela(df_ocorr), pagination=10).classes('w-full bg-[#18181A] border border-[#2B2B2F] text-xs rounded-xl')
                                            ui.button('Baixar Ocorrências (.csv)', on_click=lambda: ui.download(
                                                df_ocorr.to_csv(index=False).encode('utf-8'), f"Historico_{ent_alvo['valor']}.csv"
                                            )).classes('bg-[#222225] hover:bg-[#2B2B2F] text-gray-200 text-xs px-3 h-8 rounded-md border border-[#2B2B2F] mt-2')
                                        else:
                                            ui.label('Nenhuma ocorrência encontrada.').classes('text-xs text-[#71717A]')

                                ui.button('Ver Ocorrências', on_click=_carregar_ocorrencias).classes('bg-[#222225] hover:bg-[#2B2B2F] text-gray-200 text-xs h-8 rounded-md px-3 border border-[#2B2B2F]')
                    else:
                        ui.label('Nenhuma entidade cadastrada. Sincronize com a Blacklist ou cadastre manualmente acima.').classes('text-[#71717A]')

            ipt_f_busca.on('keydown.enter', lambda e: _atualizar_base_mestra())
            ui.button('Filtrar', on_click=_atualizar_base_mestra).props('dense').classes('bg-[#222225] hover:bg-[#2B2B2F] text-gray-200 text-xs px-3 h-8 rounded-md border border-[#2B2B2F]')
            _atualizar_base_mestra()
