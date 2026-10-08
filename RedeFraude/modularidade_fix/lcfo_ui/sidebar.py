"""
Gaveta lateral maior: Entities / Add (na mesa) ou lista de Cases.
"""

from __future__ import annotations

import json
from typing import Any, Dict

from nicegui import ui

from lcfo_ui.config import (
    TEM_ENRICH_ENGINE, CORES_POR_TIPO, enriquecer_entidade_local, CORES_STATUS, LABELS_FILTRO_TIPO, CLASSE_MENU_PADRAO,
)
from lcfo_ui.servicos import carregar_sketch, salvar_nos_extras_sketch
from lcfo_ui.dados import DADOS

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina


def criar_sidebar(ctx: "Pagina") -> None:
    """Publica ctx.sidebar_container (refreshable)."""
    st = ctx.st

    # -------------------------------------------------------------------
    # GAVETA LATERAL MAIOR (Entities / Cases)
    # -------------------------------------------------------------------
    @ui.refreshable
    def sidebar_container():
        if st.sidebar_visivel:
            with ui.column().props('id=lcfo-entities-drawer').classes(
                'h-full shrink-0 bg-[#18181A] border-r border-[#2B2B2F] p-3 overflow-y-auto custom-scroll relative'
            ).style('width:280px;'):

                if (st.sketch_ativo_id or st.dados_descoberta) and st.subtela_caso == "mesa":
                    identificador_caso = st.sketch_ativo_id or "DESCOBERTA"
                    sketch_ativo = carregar_sketch(st.sketch_ativo_id) if st.sketch_ativo_id else None
                    c_obj = next((c for c in DADOS.cluster_info
                                  if sketch_ativo and c["id"] == sketch_ativo.get("cluster_origem_id")), None)
                    nos_base = c_obj["nodes"] if c_obj else []
                    balde = st.nos_enriquecidos.get(identificador_caso, {"nodes": {}, "edges": {}})
                    selecao_atual = st.selecao_entidades.setdefault(identificador_caso, set())
                    total_nos = len(set(nos_base) | set(balde["nodes"].keys()))

                    def _tipo_de_no(node_id: str) -> str:
                        if node_id in balde["nodes"]: return balde["nodes"][node_id].get("tipo", "cpf")
                        if node_id.startswith("CPF_"): return "cpf"
                        if node_id.startswith("TEL_"): return "telefone"
                        if node_id.startswith("PLACA_"): return "placa"
                        return "cpf"

                    def _rotulo_de_no(node_id: str) -> str:
                        if node_id in balde["nodes"]: return balde["nodes"][node_id].get("label", node_id)
                        return node_id.split('_', 1)[-1]

                    def _dict_minimo(node_id: str) -> Dict[str, Any]:
                        if node_id in balde["nodes"]: return balde["nodes"][node_id]
                        return {"id": node_id, "tipo": _tipo_de_no(node_id), "valor": _rotulo_de_no(node_id), "label": _rotulo_de_no(node_id)}

                    qtd_enrich = len(balde["nodes"])
                    if st.entidade_foco or qtd_enrich > 0:
                        ent_nome = st.entidade_foco or "Ocorrência selecionada"
                        texto_foco = f"Foco: <b>{ent_nome}</b>"
                        if qtd_enrich > 0:
                            texto_foco += f"<br/><span style='color:#71717A;'>{qtd_enrich} entidade(s) via Enrich/Add</span>"
                        with ui.element('div').classes('foco-lateral-card'):
                            ui.html(texto_foco)
                        with ui.row().classes('w-full gap-2 mb-2'):
                            ui.button('Limpar Foco', on_click=lambda: (
                                setattr(st, 'entidade_foco', None), ctx.workspace.refresh(), ctx.sidebar_container.refresh()
                            )).props('dense').classes('flex-1 bg-[#18181A] hover:bg-[#222225] text-gray-300 text-xs border border-[#2B2B2F] rounded-md')
                            if qtd_enrich > 0:
                                def _limpar_enrich():
                                    st.nos_enriquecidos.pop(identificador_caso, None)
                                    if st.sketch_ativo_id:
                                        salvar_nos_extras_sketch(st.sketch_ativo_id, {}, {})
                                    ctx.workspace.refresh()
                                    ctx.sidebar_container.refresh()
                                ui.button('Limpar Enrich', on_click=_limpar_enrich).props('dense').classes('flex-1 bg-[#18181A] hover:bg-[#222225] text-gray-300 text-xs border border-[#2B2B2F] rounded-md')

                    if selecao_atual:
                        with ui.element('div').classes('selecao-lateral-card'):
                            ui.html(f"<b>{len(selecao_atual)} selecionado(s)</b>")

                        def _exportar_selecao():
                            dados_export = [_dict_minimo(nid) for nid in selecao_atual]
                            ui.download(
                                json.dumps(dados_export, ensure_ascii=False, indent=2).encode('utf-8'),
                                f'selecao_{identificador_caso}.json'
                            )

                        with ui.row().classes('w-full gap-2 mb-2'):
                            ui.button('Exportar JSON', on_click=_exportar_selecao).props('dense').classes('flex-1 bg-[#18181A] hover:bg-[#222225] text-gray-300 text-xs border border-[#2B2B2F] rounded-md')

                            def _on_limpar_selecao():
                                selecao_atual.clear()
                                for ctx in st.contexto_grafo_ativo.values():
                                    ctx["grafo"].sincronizar_selecao(selecao_atual)
                                ctx.sidebar_container.refresh()

                            ui.button('Limpar Seleção', on_click=_on_limpar_selecao).props('dense').classes('flex-1 bg-[#18181A] hover:bg-[#222225] text-gray-300 text-xs border border-[#2B2B2F] rounded-md')

                    with ui.row().classes('w-full bg-[#141416] p-1 rounded-lg border border-[#2B2B2F] gap-1 mb-2.5'):
                        ativo_ent = (st.aba_painel_lateral == "Entities")
                        ui.button(
                            f"Entities ({total_nos})",
                            icon='group',
                            color=None,
                            on_click=lambda: (setattr(st, 'aba_painel_lateral', 'Entities'), ctx.sidebar_container.refresh())
                        ).props('flat dense no-caps').classes(
                            f'flex-1 text-xs py-1 rounded-md font-medium transition-all '
                            f'{"bg-[#26262B] text-white border border-[#3A3A40] shadow-xs" if ativo_ent else "text-[#71717A] hover:text-white"}'
                        )
                        ui.button(
                            "Add",
                            icon='person_add',
                            color=None,
                            on_click=lambda: (setattr(st, 'aba_painel_lateral', 'Add'), ctx.sidebar_container.refresh())
                        ).props('flat dense no-caps').classes(
                            f'flex-1 text-xs py-1 rounded-md font-medium transition-all '
                            f'{"bg-[#26262B] text-white border border-[#3A3A40] shadow-xs" if not ativo_ent else "text-[#71717A] hover:text-white"}'
                        )

                    if st.aba_painel_lateral == "Entities":
                        todos_ids = list(dict.fromkeys(list(nos_base) + list(balde["nodes"].keys())))
                        with ui.row().classes('w-full items-center gap-1.5 mb-2 px-0.5'):
                            def _on_toggle_selecionar_todos(e):
                                if e.value:
                                    selecao_atual.update(todos_ids)
                                else:
                                    selecao_atual.clear()
                                for ctx in st.contexto_grafo_ativo.values():
                                    ctx["grafo"].sincronizar_selecao(selecao_atual)
                                ctx.sidebar_container.refresh()

                            ui.checkbox(
                                value=(len(selecao_atual) == len(todos_ids) and len(todos_ids) > 0),
                                on_change=_on_toggle_selecionar_todos
                            ).props('dense').classes('shrink-0 scale-90')

                            ui.badge(f"{len(todos_ids)} nodes").classes(
                                'bg-[#222225] border border-[#2B2B2F] text-[10px] text-gray-300 font-medium px-2 py-0.5 rounded shrink-0'
                            )

                            ipt_busca_ent = ui.input(placeholder='Search...').props('dense outlined').classes(
                                'flex-1 bg-[#141416] text-xs h-7 min-h-0 [&_.q-field__control]:h-7 [&_.q-field__control]:min-h-0 [&_.q-field__marginal]:h-7'
                            )

                            with ui.button(icon='filter_list', color=None).props('flat dense round size=sm').classes(
                                f"shrink-0 {'text-[#FF7300]' if st.filtro_tipo_entidade != 'TODOS' else 'text-[#71717A] hover:text-white'}"
                            ).tooltip('Filtrar por tipo'):
                                with ui.menu().classes(CLASSE_MENU_PADRAO):
                                    for t_opt in ["TODOS", "CPF", "TELEFONE", "PLACA"]:
                                        ui.menu_item(LABELS_FILTRO_TIPO[t_opt], on_click=lambda t=t_opt: (
                                            setattr(st, 'filtro_tipo_entidade', t),
                                            atualizar_lista_entidades(ipt_busca_ent.value)
                                        )).classes('text-xs rounded hover:bg-[#222225]')

                        box_ent_list = ui.column().classes('w-full gap-1 h-[calc(100vh-215px)] overflow-y-auto custom-scroll pr-1')

                        def atualizar_lista_entidades(filtro: str = ""):
                            box_ent_list.clear()
                            termo = (filtro or "").strip().upper()
                            with box_ent_list:
                                for n in todos_ids[:150]:
                                    tipo = _tipo_de_no(n)
                                    if st.filtro_tipo_entidade != "TODOS" and tipo.upper() != st.filtro_tipo_entidade:
                                        continue
                                    if termo and termo not in n.upper():
                                        continue
                                    val = _rotulo_de_no(n)

                                    cfg_cor = CORES_POR_TIPO.get(tipo, {"bg": "#23232A", "border": "#3A3A40"})
                                    bg_b = cfg_cor["bg"]
                                    bd_b = cfg_cor["border"]

                                    with ui.row().classes('w-full items-center gap-1.5 p-1.5 rounded-md bg-[#141416]/70 hover:bg-[#222225] border border-transparent hover:border-[#2B2B2F] transition-all'):
                                        def _on_toggle_selecao(e, nid=n):
                                            if e.value:
                                                selecao_atual.add(nid)
                                            else:
                                                selecao_atual.discard(nid)
                                            for ctx in st.contexto_grafo_ativo.values():
                                                ctx["grafo"].sincronizar_selecao(selecao_atual)
                                            ctx.sidebar_container.refresh()

                                        ui.checkbox(
                                            value=(n in selecao_atual),
                                            on_change=_on_toggle_selecao
                                        ).props('dense').classes('shrink-0 scale-90')

                                        def _focar_no(nid=n):
                                            st.entidade_foco = nid
                                            for ctx in st.contexto_grafo_ativo.values():
                                                layout_atual = st.layout_por_caso.get(ctx["identificador_caso"], "organico")
                                                novo_payload = ctx["montar_payload"](layout_atual)
                                                ctx["grafo"].enviar_atualizacao(novo_payload)
                                            ctx.sidebar_container.refresh()

                                        ui.label(val).classes('text-xs font-medium cursor-pointer text-gray-200 hover:text-[#FF7300] truncate flex-1').on('click', _focar_no)
                                        ui.html(
                                            f'<span style="background-color:{bg_b}33 !important; color:{bd_b} !important; border:1px solid {bd_b}88 !important; font-size:9px; padding:1px 6px; border-radius:4px; font-weight:600; text-transform:uppercase;">{tipo}</span>'
                                        ).classes('shrink-0')

                        ipt_busca_ent.on_value_change(lambda e: atualizar_lista_entidades(e.value))
                        atualizar_lista_entidades()

                    else:
                        with ui.column().classes('w-full gap-2 pt-1'):
                            ui.label('Adicionar Entidade').classes('text-[10px] font-semibold text-[#71717A] tracking-wider uppercase mb-1')
                            if not TEM_ENRICH_ENGINE:
                                ui.label('Módulo enrich_engine.py não localizado.').classes('text-xs text-[#71717A]')
                            else:
                                sel_tipo = ui.select(['Titular (CPF)', 'Telefone', 'Placa'], value='Titular (CPF)').props('dense outlined').classes('w-full text-xs')
                                ipt_val = ui.input(placeholder='Digite o número/placa...').props('dense outlined').classes('w-full text-xs')

                                def executar_add_entidade():
                                    v = (ipt_val.value or "").strip()
                                    if not v: return
                                    tp = "cpf" if "CPF" in sel_tipo.value else ("telefone" if "Telefone" in sel_tipo.value else "placa")
                                    res = enriquecer_entidade_local(tp, v)
                                    balde_local = st.nos_enriquecidos.setdefault(identificador_caso, {"nodes": {}, "edges": {}})
                                    for n_novo in res.get("novos_nos", []):
                                        balde_local["nodes"].setdefault(n_novo["id"], n_novo)
                                    for e_novo in res.get("novas_arestas", []):
                                        balde_local["edges"].setdefault(e_novo["id"], e_novo)
                                    if st.sketch_ativo_id:
                                        salvar_nos_extras_sketch(st.sketch_ativo_id, balde_local["nodes"], balde_local["edges"])
                                    ui.notify(f"{res.get('total_encontrado', 0)} vínculo(s) processado(s).")
                                    ctx.workspace.refresh()

                                ui.button('Buscar e Adicionar', icon='search', on_click=executar_add_entidade).classes('w-full bg-[#FF7300] hover:bg-[#E0670A] text-white font-semibold text-xs h-8 rounded-md mt-1')

                else:
                    ui.label('Cases').classes('text-xs font-bold text-[#71717A] uppercase tracking-wider mb-2 px-1')
                    with ui.column().classes('w-full gap-1 h-[calc(100vh-130px)] overflow-y-auto custom-scroll pr-1'):
                        if not DADOS.cluster_info:
                            ui.label('Nenhum dado ingerido ainda.').classes('text-xs text-[#71717A] px-1')
                        for c in DADOS.cluster_info[:50]:
                            inf = DADOS.casos_cadastrados.get(c['id'], {})
                            nome = inf.get('nome') or c['hub_label'][:20]
                            cor_st = CORES_STATUS.get(inf.get('status', 'Em Investigação'), '#8B95A5')
                            with ui.row().classes('w-full items-center justify-between p-2 rounded-md bg-[#141416]/50 hover:bg-[#222225] cursor-pointer transition-colors border border-transparent hover:border-[#2B2B2F]').on('click', lambda cid=c['id']: ctx.abrir_caso_overview(cid)):
                                with ui.row().classes('items-center gap-2 truncate'):
                                    ui.icon('folder', size='15px').classes('text-[#71717A]')
                                    ui.label(nome).classes('text-xs font-medium text-gray-200 truncate')
                                ui.element('div').classes('w-1.5 h-1.5 rounded-full shrink-0').style(f'background-color:{cor_st};')

            ui.element('div').props('id=lcfo-sidebar-handle').classes(
                'h-full w-[5px] shrink-0 cursor-ew-resize bg-[#2B2B2F] hover:bg-[#FF7300]/60 transition-colors'
            )
            ui.run_javascript("window.lcfoRegistrarSidebarResize('lcfo-entities-drawer', 'lcfo-sidebar-handle');")

    ctx.sidebar_container = sidebar_container
