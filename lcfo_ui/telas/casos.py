"""
Dashboard: fila de investigações priorizada.
"""

from __future__ import annotations

import re

from nicegui import ui

from lcfo_ui.config import CORES_STATUS, CORES_RISCO, STATUS_ATIVOS, STATUS_ENCERRADOS
from lcfo_ui.servicos import investigar_alvo_em_criacoes_diarias, carregar_historico_pareceres, listar_sketches_do_caso, obter_scores_triagem
from lcfo_ui.dados import DADOS

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina


def render(ctx: "Pagina") -> None:
    st = ctx.st
    with ui.column().classes('w-full h-full overflow-y-auto custom-scroll'):
        with ui.column().classes('w-full max-w-7xl mx-auto p-8 gap-4'):
            with ui.row().classes('w-full items-center justify-between'):
                with ui.column().classes('gap-1'):
                    ui.label('Investigations').classes('text-xl font-bold text-white tracking-tight')
                    ui.label('Manage and track your insurance fraud investigations').classes('text-xs text-[#71717A]')
                with ui.row().classes('gap-3'):
                    qtd_quadrilhas = len([c for c in DADOS.casos_cadastrados.values() if c.get('nome')])
                    for titulo, valor, cor in [
                        ('TOTAL CÉLULAS', f"{len(DADOS.cluster_info):,}", 'text-white'),
                        ('ENTIDADES', f"{DADOS.G.number_of_nodes():,}" if DADOS.G else "0", 'text-gray-300'),
                        ('QUADRILHAS', f"{qtd_quadrilhas:,}", 'text-green-400'),
                        ('NO RADAR', f"{len(DADOS.radar_alertas):,}", 'text-red-400'),
                    ]:
                        with ui.card().classes('bg-[#18181A] border border-[#2B2B2F] px-3.5 py-1.5 rounded-lg'):
                            ui.label(titulo).classes('text-[9px] text-[#71717A] font-bold')
                            ui.label(valor).classes(f'text-sm font-bold {cor}')

            if not DADOS.G or not DADOS.cluster_info:
                ui.label("Nenhum dado ingerido ainda. Use o ícone de banco de dados no rail esquerdo.").classes('text-[#71717A] mt-4')
                return

            def _on_busca_dash_change(e):
                st.termo_busca_dash = e.value or ""
                secao_investigacoes.refresh()

            ipt_busca_dash = ui.input(value=st.termo_busca_dash, placeholder='Search investigations, CPFs, phones, plates...').props('dense outlined').classes('w-full bg-[#18181A] text-xs')
            ipt_busca_dash.on_value_change(_on_busca_dash_change)

            def _buscar_no_enter():
                termo = (ipt_busca_dash.value or "").strip()
                if len(termo) < 3:
                    return
                termo_clean = re.sub(r'[^a-zA-Z0-9]', '', termo).upper()
                alvos = [n for n in DADOS.G.nodes if termo_clean in n]
                if alvos:
                    for c in DADOS.cluster_info:
                        if alvos[0] in c["nodes"]:
                            ctx.abrir_caso_overview(c["id"])
                            return
                res, df_d, ents = investigar_alvo_em_criacoes_diarias(termo)
                if res and res["total_assistencias"] > 0:
                    ui.notify(f"Alvo localizado em Criações! {res['total_assistencias']} assistências.")
                    ctx.abrir_descoberta(res, df_d, ents)
                else:
                    ui.notify("Nenhum acionamento localizado.")

            ipt_busca_dash.on('keydown.enter', lambda e: _buscar_no_enter())

            if DADOS.radar_alertas:
                with ui.expansion(f"{len(DADOS.radar_alertas)} células reincidentes precisam de atenção", icon='warning').classes('w-full bg-[#18181A] border border-[#2B2B2F] rounded-lg'):
                    with ui.row().classes('gap-2 flex-wrap'):
                        for cid, total_hits in list(DADOS.radar_alertas.items())[:4]:
                            c_obj_r = next((c for c in DADOS.cluster_info if c['id'] == cid), None)
                            if c_obj_r:
                                nome_q = DADOS.casos_cadastrados.get(cid, {}).get('nome') or c_obj_r['hub_label'][:20]
                                ui.button(f"{nome_q} (+{total_hits})", on_click=lambda cid=cid: ctx.abrir_caso_overview(cid)).classes(
                                    'bg-[#222225] hover:bg-[#2B2B2F] text-gray-200 text-xs px-3 h-8 rounded-md border border-[#2B2B2F]'
                                )

            @ui.refreshable
            def secao_investigacoes():
                mapa_scores = obter_scores_triagem(DADOS.cluster_info)
                casos_processados = []
                for c in DADOS.cluster_info[:150]:
                    info_c = DADOS.casos_cadastrados.get(c['id'], {})
                    status_c = info_c.get('status', 'Em Investigação')
                    ds = mapa_scores.get(c['id'], {"score": 0, "nivel": "BAIXO"})
                    casos_processados.append({
                        "id": c['id'], "cluster": c, "info": info_c, "status": status_c,
                        "score": ds["score"], "nivel": ds["nivel"], "em_radar": c['id'] in DADOS.radar_alertas,
                        "nomeado": bool(info_c.get('nome'))
                    })

                qtd_todos = len(casos_processados)
                qtd_ativos = len([cp for cp in casos_processados if cp["status"] in STATUS_ATIVOS])
                qtd_encerrados = len([cp for cp in casos_processados if cp["status"] in STATUS_ENCERRADOS])
                qtd_nomeados = len([cp for cp in casos_processados if cp["nomeado"]])
                qtd_nao_nomeados = len([cp for cp in casos_processados if not cp["nomeado"]])

                with ui.row().classes('w-full items-end gap-3 flex-wrap'):
                    with ui.column().classes('gap-1'):
                        ui.label('Status').classes('text-[9px] text-[#71717A] uppercase tracking-wider px-1')
                        with ui.row().classes('gap-1.5'):
                            for label, key, count in [("Todos", "Todos", qtd_todos), ("Ativos", "Ativos", qtd_ativos), ("Encerrados", "Encerrados", qtd_encerrados)]:
                                ativo_f = st.filtro_status_dash == key
                                ui.button(f"{label} ({count})", on_click=lambda k=key: (setattr(st, 'filtro_status_dash', k), secao_investigacoes.refresh())).props('flat dense no-caps').classes(
                                    f"px-3 h-7 rounded-full text-xs border {'bg-[#FF7300]/15 text-[#FF7300] border-[#FF7300]/40' if ativo_f else 'bg-[#18181A] text-[#71717A] border-[#2B2B2F] hover:text-white'}"
                                )

                    ui.separator().props('vertical').classes('bg-[#2B2B2F] h-8')

                    with ui.column().classes('gap-1'):
                        ui.label('Batismo').classes('text-[9px] text-[#71717A] uppercase tracking-wider px-1')
                        with ui.row().classes('gap-1.5'):
                            for label, key, count in [("Todos", "Todos", qtd_todos), ("Nomeados", "Nomeados", qtd_nomeados), ("Não Batizados", "Não Batizados", qtd_nao_nomeados)]:
                                ativo_b = st.filtro_batismo_dash == key
                                ui.button(f"{label} ({count})", on_click=lambda k=key: (setattr(st, 'filtro_batismo_dash', k), secao_investigacoes.refresh())).props('flat dense no-caps').classes(
                                    f"px-3 h-7 rounded-full text-xs border {'bg-[#FF7300]/15 text-[#FF7300] border-[#FF7300]/40' if ativo_b else 'bg-[#18181A] text-[#71717A] border-[#2B2B2F] hover:text-white'}"
                                )

                    ui.separator().props('vertical').classes('bg-[#2B2B2F] h-8')

                    with ui.column().classes('gap-1'):
                        ui.label('Risco').classes('text-[9px] text-[#71717A] uppercase tracking-wider px-1')
                        with ui.row().classes('gap-1.5'):
                            niveis_risco = ["Todos", "BAIXO", "MÉDIO", "ALTO", "CRÍTICO"]
                            contagem_risco = {niv: len([cp for cp in casos_processados if cp["nivel"] == niv]) for niv in niveis_risco[1:]}
                            contagem_risco["Todos"] = qtd_todos
                            for niv in niveis_risco:
                                ativo_r = st.filtro_risco_dash == niv
                                ui.button(f"{niv} ({contagem_risco[niv]})", on_click=lambda n=niv: (setattr(st, 'filtro_risco_dash', n), secao_investigacoes.refresh())).props('flat dense no-caps').classes(
                                    f"px-3 h-7 rounded-full text-xs border {'bg-[#FF7300]/15 text-[#FF7300] border-[#FF7300]/40' if ativo_r else 'bg-[#18181A] text-[#71717A] border-[#2B2B2F] hover:text-white'}"
                                )

                    with ui.row().classes('gap-1 ml-auto items-end'):
                        for icn, modo in [('grid_view', 'Grade'), ('view_list', 'Lista')]:
                            ativo_v = st.modo_visualizacao_casos == modo
                            ui.button(icon=icn, color=None, on_click=lambda m=modo: (setattr(st, 'modo_visualizacao_casos', m), secao_investigacoes.refresh())).props('flat dense round').classes(
                                f"w-8 h-8 rounded-md {'bg-[#FF7300]/15 text-[#FF7300]' if ativo_v else 'text-[#71717A] hover:text-white'}"
                            ).tooltip(modo)

                ui.label('Fila de Investigação Priorizada').classes('text-sm font-bold text-gray-300 mt-1')

                casos_filtrados = casos_processados
                if st.termo_busca_dash.strip():
                    termo_norm = st.termo_busca_dash.strip().upper()
                    casos_filtrados = [
                        cp for cp in casos_filtrados
                        if termo_norm in (cp['info'].get('nome') or '').upper()
                        or termo_norm in cp['cluster']['hub_label'].upper()
                        or termo_norm in cp['id'].upper()
                    ]
                if st.filtro_status_dash == "Ativos":
                    casos_filtrados = [cp for cp in casos_filtrados if cp["status"] in STATUS_ATIVOS]
                elif st.filtro_status_dash == "Encerrados":
                    casos_filtrados = [cp for cp in casos_filtrados if cp["status"] in STATUS_ENCERRADOS]

                if st.filtro_batismo_dash == "Nomeados":
                    casos_filtrados = [cp for cp in casos_filtrados if cp["nomeado"]]
                elif st.filtro_batismo_dash == "Não Batizados":
                    casos_filtrados = [cp for cp in casos_filtrados if not cp["nomeado"]]

                if st.filtro_risco_dash != "Todos":
                    casos_filtrados = [cp for cp in casos_filtrados if cp["nivel"] == st.filtro_risco_dash]
                casos_filtrados.sort(key=lambda cp: (cp["score"], cp["cluster"]["tamanho"]), reverse=True)

                if not casos_filtrados:
                    ui.label('Nenhum caso corresponde aos filtros selecionados.').classes('text-[#71717A]')
                elif st.modo_visualizacao_casos == "Grade":
                    with ui.grid().classes('grid-cols-3 gap-4 w-full mt-2'):
                        for cp in casos_filtrados[:45]:
                            c = cp["cluster"]
                            nome = cp["info"].get('nome') or 'Não Batizado'
                            cor_st = CORES_STATUS.get(cp["status"], '#8B95A5')
                            cor_r = CORES_RISCO.get(cp["nivel"], '#8B95A5')

                            qtd_sketches_card = len(listar_sketches_do_caso(cp['id'])) or 1
                            qtd_pareceres_card = len(carregar_historico_pareceres(cp['id']))

                            with ui.card().classes('bg-[#18181A] border border-[#2B2B2F] hover:border-[#FF7300]/60 p-4 rounded-xl flex flex-col justify-between h-36 cursor-pointer shadow-xs transition-all').on('click', lambda cid=cp['id']: ctx.abrir_caso_overview(cid)):
                                with ui.row().classes('w-full items-center justify-between'):
                                    with ui.row().classes('items-center gap-1.5'):
                                        ui.element('div').classes('w-2 h-2 rounded-full').style(f'background-color:{cor_st};')
                                        ui.label(cp["status"]).classes('text-[11px] font-semibold').style(f'color:{cor_st};')
                                    ui.badge(f"{cp['nivel']} · {cp['score']}").style(f'background-color:{cor_r}22;color:{cor_r};border:1px solid {cor_r}55;font-size:10px;padding:2px 6px;border-radius:4px;')

                                with ui.column().classes('gap-0.5 w-full'):
                                    ui.label(nome).classes('text-sm font-semibold text-white tracking-tight truncate w-full')
                                    ui.label(c['hub_label']).classes('text-[11px] text-[#71717A] truncate w-full')

                                with ui.row().classes('w-full items-center gap-4 text-xs text-[#71717A] mt-auto pt-1'):
                                    with ui.row().classes('items-center gap-1'):
                                        ui.icon('hub', size='14px').classes('text-[#71717A]')
                                        ui.label(str(qtd_sketches_card)).classes('text-xs text-gray-300 font-medium')
                                    with ui.row().classes('items-center gap-1'):
                                        ui.icon('description', size='14px').classes('text-[#71717A]')
                                        ui.label(str(qtd_pareceres_card)).classes('text-xs text-gray-300 font-medium')
                else:
                    for cp in casos_filtrados[:45]:
                        c = cp["cluster"]
                        cor_st = CORES_STATUS.get(cp["status"], '#8B95A5')
                        cor_r = CORES_RISCO.get(cp["nivel"], '#8B95A5')
                        qtd_sketches_card = len(listar_sketches_do_caso(cp['id'])) or 1
                        qtd_pareceres_card = len(carregar_historico_pareceres(cp['id']))

                        with ui.row().classes('w-full items-center justify-between p-3 rounded-lg bg-[#18181A] border border-[#2B2B2F] hover:border-[#FF7300]/40 cursor-pointer transition-colors mb-2').on('click', lambda cid=cp['id']: ctx.abrir_caso_overview(cid)):
                            with ui.column().classes('gap-0.5'):
                                ui.label(cp["info"].get('nome') or 'Não Batizado').classes('text-sm font-semibold text-white')
                                ui.label(c['hub_label']).classes('text-[11px] text-[#71717A]')
                            with ui.row().classes('items-center gap-3'):
                                with ui.row().classes('items-center gap-1'):
                                    ui.icon('hub', size='14px').classes('text-[#71717A]')
                                    ui.label(str(qtd_sketches_card)).classes('text-xs text-gray-300 font-medium')
                                with ui.row().classes('items-center gap-1'):
                                    ui.icon('description', size='14px').classes('text-[#71717A]')
                                    ui.label(str(qtd_pareceres_card)).classes('text-xs text-gray-300 font-medium')
                                ui.element('div').classes('w-2 h-2 rounded-full').style(f'background-color:{cor_st};')
                                ui.label(cp["status"]).classes('text-[11px]').style(f'color:{cor_st};')
                                ui.badge(f"{cp['nivel']} · {cp['score']}").style(f'background-color:{cor_r}22;color:{cor_r};border:1px solid {cor_r}55;font-size:10px;padding:2px 6px;border-radius:4px;')

            secao_investigacoes()
