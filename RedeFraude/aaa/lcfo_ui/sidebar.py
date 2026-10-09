"""
Gaveta lateral maior: Entities / Add (na mesa) ou lista de Cases.
"""

from __future__ import annotations

import json
import re
import unicodedata
from typing import Any, Dict

from nicegui import ui

from lcfo_ui.config import (
    TEM_ENRICH_ENGINE, CORES_POR_TIPO, enriquecer_entidade_local, CORES_STATUS, LABELS_FILTRO_TIPO, CLASSE_MENU_PADRAO,
)
from lcfo_ui.servicos import carregar_sketch, salvar_nos_extras_sketch
from lcfo_ui.dados import DADOS
from utils import formatar_cpf_cnpj, formatar_tel

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina


def _normalizar_busca(texto: str) -> str:
    """Maiúsculas, sem acento e só letras/dígitos: "651.981.140-20" e "65198114020" casam, "José" e "JOSE" também."""
    sem = "".join(c for c in unicodedata.normalize("NFKD", str(texto)) if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]", "", sem.upper())


def criar_sidebar(ctx: "Pagina") -> None:
    """Publica ctx.sidebar_container (refreshable)."""
    st = ctx.st

    # -------------------------------------------------------------------
    # GAVETA LATERAL MAIOR (Entities / Cases)
    # -------------------------------------------------------------------
    @ui.refreshable
    def sidebar_container():
        # Centro de Comando: o mapa ocupa a tela toda e o painel do radar faz o papel da gaveta.
        if st.tela_ativa == "Centro de Comando" and not st.caso_ativo_id and not st.dados_descoberta:
            return
        if st.sidebar_visivel:
            with ui.column().props('id=lcfo-entities-drawer').classes(
                'h-full shrink-0 bg-[#18181A] p-3 overflow-hidden relative no-wrap gap-0'
            ).style('width:280px;'):

                if (st.sketch_ativo_id or st.dados_descoberta) and st.subtela_caso == "mesa":
                    identificador_caso = st.sketch_ativo_id or "DESCOBERTA"
                    sketch_ativo = carregar_sketch(st.sketch_ativo_id) if st.sketch_ativo_id else None
                    c_obj = next((c for c in DADOS.cluster_info
                                  if sketch_ativo and c["id"] == sketch_ativo.get("cluster_origem_id")), None)
                    nos_base = list(c_obj["nodes"]) if c_obj else []
                    balde = st.nos_enriquecidos.get(identificador_caso, {"nodes": {}, "edges": {}})
                    # Descoberta (sem célula salva): as entidades vêm do próprio df — antes a lista ficava vazia.
                    nomes_desc: Dict[str, str] = {}
                    if not c_obj and st.dados_descoberta is not None:
                        df_d = st.dados_descoberta.get("df")
                        if df_d is not None and not df_d.empty:
                            for col, pref in (("cpf", "CPF_"), ("telefone", "TEL_"), ("placa", "PLACA_")):
                                if col in df_d.columns:
                                    for v in df_d[col].dropna().astype(str).str.strip().unique():
                                        if v and v.lower() not in ("nan", "none"):
                                            nos_base.append(pref + v)
                            if {"cpf", "titular"} <= set(df_d.columns):
                                for c_, n_ in zip(df_d["cpf"].astype(str), df_d["titular"].astype(str)):
                                    if n_ and n_.lower() not in ("nan", "none"):
                                        nomes_desc.setdefault("CPF_" + c_.strip(), n_.strip())
                    selecao_atual = st.selecao_entidades.setdefault(identificador_caso, set())
                    total_nos = len(set(nos_base) | set(balde["nodes"].keys()))

                    def _tipo_de_no(node_id: str) -> str:
                        if node_id in balde["nodes"]: return balde["nodes"][node_id].get("tipo", "cpf")
                        if node_id.startswith("CPF_"): return "cpf"
                        if node_id.startswith("TEL_"): return "telefone"
                        if node_id.startswith("PLACA_"): return "placa"
                        return "cpf"

                    def _rotulo_de_no(node_id: str) -> str:
                        return _info_no(node_id)["principal"]

                    def _info_no(node_id: str) -> Dict[str, str]:
                        """Nome (CPF) e valor formatado — do Enrich, do grafo ou da descoberta."""
                        cru = node_id.split('_', 1)[-1]
                        tipo = _tipo_de_no(node_id)
                        nome, valor = "", ""
                        if node_id in balde["nodes"]:
                            b_ = balde["nodes"][node_id]
                            nome, valor = b_.get("nome_titular") or "", b_.get("valor") or b_.get("label") or ""
                        elif DADOS.G is not None and node_id in DADOS.G:
                            a_ = DADOS.G.nodes[node_id]
                            nome, valor = a_.get("nome_titular") or "", a_.get("valor") or ""
                        nome = nome if nome and nome != "N/D" else nomes_desc.get(node_id, "")
                        if not valor:
                            valor = formatar_cpf_cnpj(cru) if tipo == "cpf" else formatar_tel(cru) if tipo == "telefone" else (
                                f"{cru[:3]}-{cru[3:]}" if len(cru) == 7 else cru)
                        return {"principal": nome or valor, "secundario": valor if nome else "",
                                "busca": _normalizar_busca(f"{nome} {valor} {cru}")}

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
                                for cg in st.contexto_grafo_ativo.values():
                                    cg["grafo"].sincronizar_selecao(selecao_atual)
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
                                for cg in st.contexto_grafo_ativo.values():
                                    cg["grafo"].sincronizar_selecao(selecao_atual)
                                ctx.sidebar_container.refresh()

                            ui.checkbox(
                                value=(len(selecao_atual) == len(todos_ids) and len(todos_ids) > 0),
                                on_change=_on_toggle_selecionar_todos
                            ).props('dense').classes('shrink-0 scale-90')

                            # (a contagem 'N de M entidades' fica na linha abaixo da busca)

                            ipt_busca_ent = ui.input(placeholder='Nome, CPF, telefone, placa...').props('dense outlined clearable').classes(
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

                        lbl_contagem = ui.label('').classes('text-[10px] text-[#71717A] px-0.5 mb-1')
                        # A lista é a única área com rolagem; -mr-3 leva a barra até a borda da gaveta
                        # (antes ela ficava entre os itens e os selos de tipo, com outra rolagem por fora).
                        box_ent_list = ui.column().classes('w-full gap-1 flex-1 min-h-0 overflow-y-auto custom-scroll -mr-3 pr-2 no-wrap')
                        MAX_LISTA = 300

                        def atualizar_lista_entidades(filtro: str = ""):
                            box_ent_list.clear()
                            termo = _normalizar_busca(filtro or "")
                            casam = []
                            for n in todos_ids:
                                tipo = _tipo_de_no(n)
                                if st.filtro_tipo_entidade != "TODOS" and tipo.upper() != st.filtro_tipo_entidade:
                                    continue
                                info = _info_no(n)
                                # Como a busca do grafo: sem pontuação nem acento, por nome, valor formatado ou dígitos.
                                if termo and termo not in info["busca"]:
                                    continue
                                casam.append((n, tipo, info))
                            casam.sort(key=lambda x: (not x[2]["secundario"], x[2]["principal"]))
                            lbl_contagem.set_text(
                                f"{len(casam)} de {len(todos_ids)} entidades"
                                + (f" · mostrando {MAX_LISTA} — refine a busca" if len(casam) > MAX_LISTA else ""))
                            with box_ent_list:
                                if not casam:
                                    ui.label('Nenhuma entidade encontrada.').classes('text-xs text-[#71717A] px-1 py-2')
                                for n, tipo, info in casam[:MAX_LISTA]:

                                    cfg_cor = CORES_POR_TIPO.get(tipo, {"bg": "#23232A", "border": "#3A3A40"})
                                    bg_b = cfg_cor["bg"]
                                    bd_b = cfg_cor["border"]

                                    with ui.row().classes('w-full items-center gap-1.5 p-1.5 rounded-md bg-[#141416]/70 hover:bg-[#222225] border border-transparent hover:border-[#2B2B2F] transition-all'):
                                        def _on_toggle_selecao(e, nid=n):
                                            if e.value:
                                                selecao_atual.add(nid)
                                            else:
                                                selecao_atual.discard(nid)
                                            for cg in st.contexto_grafo_ativo.values():
                                                cg["grafo"].sincronizar_selecao(selecao_atual)
                                            ctx.sidebar_container.refresh()

                                        ui.checkbox(
                                            value=(n in selecao_atual),
                                            on_change=_on_toggle_selecao
                                        ).props('dense').classes('shrink-0 scale-90')

                                        def _focar_no(nid=n):
                                            st.entidade_foco = nid
                                            for cg in st.contexto_grafo_ativo.values():
                                                layout_atual = st.layout_por_caso.get(cg["identificador_caso"], "organico")
                                                novo_payload = cg["montar_payload"](layout_atual)
                                                cg["grafo"].enviar_atualizacao(novo_payload)
                                            ctx.sidebar_container.refresh()

                                        with ui.column().classes('gap-0 flex-1 min-w-0 cursor-pointer').on('click', _focar_no):
                                            ui.label(info["principal"]).classes('text-xs font-medium text-gray-200 hover:text-[#FF7300] truncate w-full')
                                            if info["secundario"]:
                                                ui.label(info["secundario"]).classes('text-[10px] text-[#71717A] truncate w-full')
                                        ui.html(
                                            f'<span style="background-color:{bg_b}33 !important; color:{bd_b} !important; border:1px solid {bd_b}88 !important; font-size:9px; padding:1px 6px; border-radius:4px; font-weight:600; text-transform:uppercase;">{tipo}</span>'
                                        ).classes('shrink-0')

                        ipt_busca_ent.on_value_change(lambda e: atualizar_lista_entidades(e.value or ""))
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
                    with ui.column().classes('w-full gap-1 flex-1 min-h-0 overflow-y-auto custom-scroll -mr-3 pr-2 no-wrap'):
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

            # Alça no mesmo estilo do divisor da tela dividida: linha de 1px com a "pega" pontilhada no meio.
            ui.element('div').props('id=lcfo-sidebar-handle title="Arraste para ajustar a largura"').classes(
                'lcfo-alca-lateral h-full shrink-0'
            )
            ui.run_javascript("window.lcfoRegistrarSidebarResize('lcfo-entities-drawer', 'lcfo-sidebar-handle');")

    ctx.sidebar_container = sidebar_container
