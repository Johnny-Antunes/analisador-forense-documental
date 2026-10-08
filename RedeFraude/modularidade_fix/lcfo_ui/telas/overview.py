"""
Overview do dossiê: cabeçalho, metadados, grid de sketches e análises.
"""

from __future__ import annotations


from nicegui import ui

from lcfo_ui.config import CORES_STATUS, CORES_RISCO, CLASSE_MENU_PADRAO
from lcfo_ui.servicos import (
    atualizar_metadados_caso, carregar_dados_caso, criar_sketch, listar_sketches_do_caso, carregar_nos_extras_sketch, obter_scores_triagem,
)
from lcfo_ui.dados import DADOS
from lcfo_ui.grafo import gerar_svg_miniatura_cluster

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina


def render(ctx: "Pagina") -> None:
    st = ctx.st
    id_dossie = st.caso_ativo_id
    dados_caso = carregar_dados_caso(id_dossie)
    nome_exib = dados_caso["nome_personalizado"] or f"Dossiê {id_dossie}"

    sketches = listar_sketches_do_caso(id_dossie)
    if not sketches:
        c_match = next((c for c in DADOS.cluster_info if c["id"] == id_dossie), None)
        if c_match:
            criar_sketch(id_dossie, "Cluster Relacional Principal", tipo_origem="CLUSTER_OFICIAL", cluster_origem_id=id_dossie)
            sketches = listar_sketches_do_caso(id_dossie)

    mapa_scores_ov = obter_scores_triagem(DADOS.cluster_info)
    score_ov = mapa_scores_ov.get(id_dossie, {"score": 0, "nivel": "BAIXO"})
    cor_risco_ov = CORES_RISCO.get(score_ov['nivel'], '#8B95A5')
    cor_status_ov = CORES_STATUS.get(dados_caso['status'], '#8B95A5')

    with ui.column().classes('w-full h-full overflow-y-auto custom-scroll'):
        with ui.column().classes('w-full max-w-6xl mx-auto p-8 gap-4'):
            with ui.row().classes('w-full items-start justify-between border-b border-[#2B2B2F] pb-4'):
                with ui.column().classes('gap-1'):
                    ui.label(f"Cases / {nome_exib}").classes('text-xs text-[#71717A] font-medium')
                    ui.label(nome_exib).classes('text-2xl font-bold text-white tracking-tight')

                    with ui.row().classes('items-center gap-2 mt-1 text-xs no-wrap'):
                        ui.label(f"Status: {dados_caso['status']}").classes('text-xs').style(f"color:{cor_status_ov}; font-weight:600;")
                        ui.label('•').classes('text-[#71717A]')
                        ui.label(f"Risco {score_ov['nivel']} · {score_ov['score']}/100").style(f"color:{cor_risco_ov}; font-weight:700;")
                        ui.label('•').classes('text-[#71717A]')
                        ui.label(f"Updated {dados_caso.get('data_atualizacao') or 'Hoje'}").classes('text-[#71717A]')

                    with ui.row().classes('items-center gap-1 text-xs mt-0.5'):
                        ui.label('Investigador:').classes('text-[#71717A]')
                        ui.label(dados_caso.get('analista_responsavel') or 'Sistema').classes('text-gray-200 font-medium')

                with ui.button(icon='settings', color=None).props('flat dense round').classes(
                    'w-8 h-8 rounded-lg text-[#71717A] hover:text-white hover:bg-[#18181A]'
                ).tooltip('Editar metadados'):
                    with ui.menu().classes(f'{CLASSE_MENU_PADRAO} w-80'):
                        ui.label('Nome da Quadrilha:').classes('text-xs text-gray-300 font-medium mb-1')
                        ipt_edit_nome = ui.input(value=dados_caso['nome_personalizado']).props('dense outlined').classes('w-full text-xs mb-3')

                        ui.label('Status:').classes('text-xs text-gray-300 font-medium mb-1')
                        lista_status_edit = ["Em Investigação", "Confirmado Fraude", "Monitoramento Contínuo", "Sem Irregularidade Identificada", "Falso Positivo", "Arquivado"]
                        sel_edit_status = ui.select(
                            lista_status_edit,
                            value=dados_caso['status'] if dados_caso['status'] in lista_status_edit else lista_status_edit[0]
                        ).props('dense outlined').classes('w-full text-xs mb-3')

                        ui.label('Analista:').classes('text-xs text-gray-300 font-medium mb-1')
                        ipt_edit_analista = ui.input(value=dados_caso['analista_responsavel']).props('dense outlined').classes('w-full text-xs mb-4')

                        def _salvar_edicao_metadados():
                            novo_nome = (ipt_edit_nome.value or "").strip()
                            novo_status = sel_edit_status.value
                            novo_analista = (ipt_edit_analista.value or "").strip()

                            atualizar_metadados_caso(id_dossie, novo_nome, novo_status, novo_analista)
                            DADOS.recarregar()
                            ui.notify("Metadados atualizados com sucesso.")
                            ctx.navbar_breadcrumbs.refresh()
                            ctx.sidebar_container.refresh()
                            ctx.workspace.refresh()

                        ui.button('Salvar Alterações', on_click=_salvar_edicao_metadados).classes(
                            'w-full bg-[#FF7300] hover:bg-[#E0670A] text-white text-xs font-semibold h-8 rounded-md'
                        )

            with ui.row().classes('w-full items-center justify-between mt-2'):
                ui.label(f'Sketches ({len(sketches)})').classes('text-sm font-bold text-gray-200 tracking-tight')
                ui.button('Novo Sketch', icon='add', on_click=lambda: ctx.abrir_dialog_novo_sketch(id_dossie)).classes(
                    'bg-[#FF7300] hover:bg-[#E0670A] text-white text-xs font-semibold px-4 h-8 rounded-md'
                )

            if not sketches:
                ui.label('Nenhum sketch neste dossiê ainda.').classes('text-[#71717A] mt-2')
            else:
                with ui.grid().classes('grid-cols-3 gap-4 w-full mt-2'):
                    for sk in sketches:
                        cluster_sk = next((c for c in DADOS.cluster_info if c["id"] == sk["cluster_origem_id"]), None) if sk["cluster_origem_id"] else None
                        nos_extra_sk, _ = carregar_nos_extras_sketch(sk["id_sketch"])
                        total_nos_sk = (cluster_sk["tamanho"] if cluster_sk else 0) + len(nos_extra_sk)
                        qtd_tel_sk = (cluster_sk["qtd_tels"] if cluster_sk else 0)
                        qtd_cpf_sk = (cluster_sk["qtd_cpfs"] if cluster_sk else 0)
                        qtd_placa_sk = (cluster_sk.get("qtd_placas", 0) if cluster_sk else 0)

                        with ui.element('div').classes('lcfo-sketch-card flex flex-col'):
                            with ui.row().classes('w-full items-center justify-between px-3 pt-3'):
                                ui.label(sk['nome_sketch']).classes('text-xs font-bold text-white truncate flex-1')
                                ui.button(icon='delete_outline', color=None,
                                          on_click=lambda sk=sk: ctx.abrir_confirmacao_remocao_sketch(sk)).props(
                                    'flat dense round size=xs'
                                ).classes('text-[#71717A] hover:text-red-400 shrink-0').tooltip('Remover sketch')

                            with ui.element('div').classes('cursor-pointer').on(
                                'click', lambda sid=sk['id_sketch']: ctx.abrir_mesa_sketch(sid)
                            ):
                                with ui.element('div').classes('w-full h-20 bg-[#141416] mx-0 mt-1 flex items-center justify-center px-2'):
                                    ui.html(gerar_svg_miniatura_cluster(cluster_sk)).classes('w-full h-full')

                                with ui.column().classes('gap-0.5 w-full border-t border-[#2B2B2F] p-3'):
                                    if sk.get('descricao'):
                                        ui.label(sk['descricao']).classes('text-[11px] text-[#71717A] truncate')
                                    if cluster_sk:
                                        ui.label(f"{total_nos_sk} nós · {qtd_tel_sk} tels · {qtd_cpf_sk} tit · {qtd_placa_sk} placas").classes('text-[11px] text-[#71717A]')
                                    else:
                                        ui.label(f"Sem célula vinculada · {len(nos_extra_sk)} entidade(s) manual(is)").classes('text-[11px] text-[#71717A]')

            with ui.card().classes('bg-transparent border-none p-0 w-full mt-4'):
                ctx.renderizar_painel_analises(id_dossie, dados_caso, ctx.workspace.refresh)
