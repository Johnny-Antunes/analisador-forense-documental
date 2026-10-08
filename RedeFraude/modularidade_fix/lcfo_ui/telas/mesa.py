"""
Mesa de análise do sketch / descoberta: HUD, promoção, toolbar, cockpit e notas.
"""

from __future__ import annotations

from typing import Dict

import pandas as pd
from nicegui import ui

from lcfo_ui.config import OPCOES_TOOLBAR, ICONE_TOOLBAR, CLASSE_MENU_PADRAO
from lcfo_ui.servicos import (
    consultar_detalhes_caso, carregar_dados_caso, promover_descoberta_para_caso, vincular_caso_e_entidades, anexar_descoberta_a_caso_existente, criar_sketch, listar_sketches_do_caso, carregar_sketch, calcular_score_caso,
)
from lcfo_ui.dados import DADOS
from lcfo_ui.grafo import montar_payload_grafo

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina

from lcfo_ui.telas.contexto_mesa import ContextoMesa
from lcfo_ui.telas.paineis import renderizar_painel


def render(ctx: "Pagina") -> None:
    st = ctx.st
    dossie_persistido = st.sketch_ativo_id is not None
    identificador_caso = st.sketch_ativo_id or "DESCOBERTA"

    st.contexto_grafo_ativo = {}

    cluster_obj = None
    sketch_ativo = None
    if dossie_persistido:
        sketch_ativo = carregar_sketch(st.sketch_ativo_id)
        if not sketch_ativo:
            ui.label('Sketch não encontrado. Volte ao Dashboard.').classes('text-red-400 p-8')
            return
        cluster_obj = next((c for c in DADOS.cluster_info if c["id"] == sketch_ativo.get("cluster_origem_id")), None)
        dados_caso = carregar_dados_caso(sketch_ativo["id_caso"])
        nome_dossie = dados_caso["nome_personalizado"] or f"Dossiê {sketch_ativo['id_caso']}"
        nome_exib = f"{nome_dossie} · {sketch_ativo['nome_sketch']}"

        cpfs, tels, placas = [], [], []
        if cluster_obj:
            cpfs += [n.replace("CPF_", "") for n in cluster_obj["nodes"] if n.startswith("CPF_")]
            tels += [n.replace("TEL_", "") for n in cluster_obj["nodes"] if n.startswith("TEL_")]
            placas += [n.replace("PLACA_", "") for n in cluster_obj["nodes"] if n.startswith("PLACA_")]

        balde_atual = st.nos_enriquecidos.get(identificador_caso, {"nodes": {}, "edges": {}})
        for ninfo in balde_atual["nodes"].values():
            tipo_n, val_n = ninfo.get("tipo"), ninfo.get("valor", "")
            if tipo_n == "cpf" and val_n: cpfs.append(val_n)
            elif tipo_n == "telefone" and val_n: tels.append(val_n)
            elif tipo_n == "placa" and val_n: placas.append(val_n)

        df_dados = consultar_detalhes_caso(cpfs, tels, placas) if (cpfs or tels or placas) else pd.DataFrame()
        resumo_descoberta = None
    else:
        if not st.dados_descoberta:
            ui.label('Nenhuma descoberta ativa.').classes('text-gray-400 p-8')
            return
        d = st.dados_descoberta
        df_dados = d["df"]
        resumo_descoberta = d["resumo"]
        nome_exib = d["resumo"]["alvo_buscado"]
        dados_caso = {"status": "Descoberta Ativa", "analista_responsavel": "", "data_atualizacao": "Agora"}
        cpfs = list(df_dados["cpf"].dropna().unique()) if "cpf" in df_dados.columns else []
        tels = list(df_dados["telefone"].dropna().unique()) if "telefone" in df_dados.columns else []
        placas = list(df_dados["placa"].dropna().unique()) if "placa" in df_dados.columns else []

    layout_ativo = st.layout_por_caso.get(identificador_caso, "organico")
    chave_layout = (sketch_ativo.get("cluster_origem_id") if sketch_ativo else None) or identificador_caso

    payload, vis_nodes, vis_edges, hub_id = montar_payload_grafo(st, 
        identificador_caso, cluster_obj is not None, cluster_obj, df_dados, resumo_descoberta, layout_ativo,
        chave_layout=chave_layout
    )

    if cluster_obj:
        max_bet = max((n.get("betweenness", 0.0) for n in vis_nodes), default=0.0)
        subG = DADOS.G.subgraph(cluster_obj["nodes"])
        score, nivel, cor, fatores, _ = calcular_score_caso(cluster_obj, df_dados, subG, betweenness_precalculado=max_bet)
    else:
        score, nivel, cor, fatores, _ = calcular_score_caso(None, df_dados, None)

    periodo = "Sem datas disponíveis"
    if not df_dados.empty and "data" in df_dados.columns:
        dts = pd.to_datetime(df_dados["data"], errors="coerce").dropna()
        if not dts.empty:
            periodo = f"{dts.min().strftime('%d/%m/%Y')} → {dts.max().strftime('%d/%m/%Y')}"

    @ui.refreshable
    def painel_notas():
        ctx.renderizar_painel_analises(sketch_ativo["id_caso"] if sketch_ativo else identificador_caso, dados_caso, painel_notas.refresh)

    m = ContextoMesa(
        ctx=ctx,
        identificador_caso=identificador_caso,
        dossie_persistido=dossie_persistido,
        sketch_ativo=sketch_ativo,
        cluster_obj=cluster_obj,
        dados_caso=dados_caso,
        nome_exib=nome_exib,
        df_dados=df_dados,
        resumo_descoberta=resumo_descoberta,
        cpfs=cpfs,
        tels=tels,
        placas=placas,
        chave_layout=chave_layout,
        layout_ativo=layout_ativo,
        payload=payload,
        vis_nodes=vis_nodes,
        vis_edges=vis_edges,
        hub_id=hub_id,
    )

    def renderizar_painel_interno(nome_painel: str, lado: str = "unico"):
        renderizar_painel(m, nome_painel, lado)

    with ui.row().classes('w-full h-full flex-nowrap gap-0'):
        with ui.element('div').classes('relative flex-1 h-full overflow-hidden bg-[#121214]'):

            with ui.column().classes('absolute top-3 left-1/2 -translate-x-1/2 z-40 items-stretch gap-2 w-[92%] max-w-[1120px] pointer-events-none'):
                with ui.row().classes('w-full flex-nowrap items-stretch gap-2 pointer-events-auto'):
                    with ui.row().classes('lcfo-panel hud-compact-row flex-1 min-w-0'):
                        ui.html(f"""
                            <span style="font-size:13px; font-weight:700; color:#EDEDED;">{nome_exib.upper()}</span>
                            <span class="badge-pill" style="background:{cor}22; color:{cor}; border:1px solid {cor}55;">FRAUD SCORE: {score}/100 [{nivel}]</span>
                            <span style="font-size:11px; color:#71717A;">Status: <b style="color:#6FA98A;">{dados_caso['status']}</b></span>
                            <span style="font-size:11px; color:#71717A;">Âncora: <b style="color:#C9A66B;">{hub_id or 'N/D'}</b></span>
                            <span style="font-size:11px; color:#71717A;">Acionamentos: <b style="color:#EDEDED;">{len(df_dados)} reg.</b></span>
                            <span style="font-size:11px; color:#71717A;">Período: <b style="color:#CCCCCC;">{periodo}</b></span>
                        """)
                    with ui.button(icon='analytics', color=None).props('flat dense round').classes(
                        'lcfo-panel lcfo-dossie-btn shrink-0 self-stretch !rounded-[10px]'
                    ).tooltip('Dossiê Topológico'):
                        with ui.menu().classes(f'{CLASSE_MENU_PADRAO} w-[760px] max-w-[95vw]').style('display: flex; gap: 12px; flex-direction: row;'):
                            with ui.column().classes('flex-1 bg-[#18181A] border border-[#2B2B2F] p-3.5 rounded-lg shadow-sm'):
                                ui.label('INDICADORES IDENTIFICADOS').style('font-size: 11px; font-weight: 700; color: #FF7300; letter-spacing: 0.05em; margin-bottom: 6px;')
                                ui.html('<br/>'.join([f"• {f}" for f in fatores[:5]]) if fatores else "Comportamento relacional estável dentro da normalidade.", tag='div').style('font-size: 12px; line-height: 1.6; color: #CCCCCC;')
                            with ui.column().classes('flex-1 bg-[#18181A] border border-[#2B2B2F] p-3.5 rounded-lg shadow-sm'):
                                ui.label('MÉTRICAS DA CÉLULA').style('font-size: 11px; font-weight: 700; color: #FF7300; letter-spacing: 0.05em; margin-bottom: 6px;')
                                ui.html(f"• Nós: {len(vis_nodes)}<br/>• Conexões: {len(vis_edges)}<br/>• Acionamentos: {len(df_dados)}", tag='div').style('font-size: 12px; line-height: 1.6; color: #CCCCCC;')
                            with ui.column().classes('flex-1 bg-[#18181A] border border-[#2B2B2F] p-3.5 rounded-lg shadow-sm'):
                                ui.label('STATUS DA INVESTIGAÇÃO').style('font-size: 11px; font-weight: 700; color: #FF7300; letter-spacing: 0.05em; margin-bottom: 6px;')
                                ui.html(f"• Analista: {dados_caso.get('analista_responsavel') or 'Não atribuído'}<br/>• Atualizado: {dados_caso.get('data_atualizacao') or 'Hoje'}", tag='div').style('font-size: 12px; line-height: 1.6; color: #CCCCCC;')

                if not dossie_persistido:
                    opcoes_sketches_todos: Dict[str, str] = {}
                    for cid_dossie, info_dossie in DADOS.casos_cadastrados.items():
                        for sk in listar_sketches_do_caso(cid_dossie):
                            if sk["cluster_origem_id"]:
                                nome_dossie_opt = info_dossie.get("nome") or cid_dossie
                                opcoes_sketches_todos[sk["id_sketch"]] = f"{nome_dossie_opt} · {sk['nome_sketch']}"
                    opcoes_dossies_todos = {cid: (info.get('nome') or cid) for cid, info in DADOS.casos_cadastrados.items()}

                    with ui.expansion('⚡ Promover a Novo Caso, Novo Sketch ou Fundir', icon='auto_awesome').classes('lcfo-panel pointer-events-auto w-full text-xs'):
                        with ui.row().classes('w-full gap-3 mt-2 flex-wrap'):
                            with ui.card().classes('bg-[#18181A] border border-[#2B2B2F] flex-1 min-w-[240px] p-3'):
                                ui.label('Criar Nova Operação').classes('text-xs font-bold text-[#FF7300]')
                                ipt_p_nome = ui.input(placeholder='Nome da Investigação...').props('dense outlined').classes('w-full text-xs mt-1')
                                sel_p_st = ui.select(["Em Investigação", "Confirmado Fraude", "Monitoramento Contínuo"], value="Em Investigação").props('dense outlined').classes('w-full text-xs mt-1')
                                ipt_p_an = ui.input(placeholder='Analista Responsável').props('dense outlined').classes('w-full text-xs mt-1')

                                def _exec_promover():
                                    nm = (ipt_p_nome.value or "").strip()
                                    if not nm:
                                        ui.notify("Informe um nome para o caso.", type="warning")
                                        return
                                    nos_p = promover_descoberta_para_caso(df_dados)
                                    DADOS.recarregar()
                                    cl_p = next((c for c in DADOS.cluster_info if any(n in c["nodes"] for n in nos_p)), None)
                                    if cl_p:
                                        vincular_caso_e_entidades(cl_p["id"], nos_p, nm, sel_p_st.value, ipt_p_an.value, "")
                                        DADOS.recarregar()
                                        criar_sketch(cl_p["id"], "Sketch Principal", tipo_origem="DESCOBERTA_DIARIA", cluster_origem_id=cl_p["id"])
                                        ui.notify("Rede autuada com sucesso!")
                                        ctx.abrir_caso_overview(cl_p["id"])
                                    else:
                                        ui.notify("Erro na consolidação da malha.", type="negative")

                                ui.button('Autuar Nova Operação', on_click=_exec_promover).classes('w-full bg-[#FF7300] hover:bg-[#E0670A] text-white text-xs h-7 rounded-md mt-2')

                            with ui.card().classes('bg-[#18181A] border border-[#2B2B2F] flex-1 min-w-[240px] p-3'):
                                ui.label('Novo Sketch em Dossiê Existente').classes('text-xs font-bold text-[#FF7300]')
                                sel_ns_dossie = ui.select(opcoes_dossies_todos, label='Dossiê:').props('dense outlined').classes('w-full text-xs mt-1')
                                ipt_ns_nome = ui.input(placeholder='Nome deste sketch...').props('dense outlined').classes('w-full text-xs mt-1')

                                def _exec_novo_sketch():
                                    if not sel_ns_dossie.value or not (ipt_ns_nome.value or "").strip():
                                        ui.notify("Selecione o dossiê e informe um nome.", type="warning")
                                        return
                                    nos_p = promover_descoberta_para_caso(df_dados)
                                    DADOS.recarregar()
                                    cl_p = next((c for c in DADOS.cluster_info if any(n in c["nodes"] for n in nos_p)), None)
                                    id_sk = criar_sketch(
                                        sel_ns_dossie.value, ipt_ns_nome.value.strip(),
                                        tipo_origem="DESCOBERTA_DIARIA",
                                        cluster_origem_id=(cl_p["id"] if cl_p else "")
                                    )
                                    ui.notify("Novo sketch criado no dossiê selecionado!")
                                    ctx.abrir_mesa_sketch(id_sk)

                                ui.button('Criar Sketch no Dossiê', on_click=_exec_novo_sketch).classes('w-full bg-[#222225] hover:bg-[#2B2B2F] text-white text-xs h-7 rounded-md mt-2')

                            with ui.card().classes('bg-[#18181A] border border-[#2B2B2F] flex-1 min-w-[240px] p-3'):
                                ui.label('Fundir a Sketch Existente').classes('text-xs font-bold text-[#FF7300]')
                                sel_ns_alvo = ui.select(opcoes_sketches_todos, label='Sketch de destino:').props('dense outlined').classes('w-full text-xs mt-1')
                                ipt_ns_motivo = ui.input(placeholder='Motivo do vínculo...').props('dense outlined').classes('w-full text-xs mt-1')

                                def _exec_fundir():
                                    if not sel_ns_alvo.value:
                                        ui.notify("Selecione o sketch de destino.", type="warning")
                                        return
                                    sketch_alvo = carregar_sketch(sel_ns_alvo.value)
                                    if not sketch_alvo or not sketch_alvo["cluster_origem_id"]:
                                        ui.notify("Sketch selecionado ainda não tem célula oficial vinculada.", type="negative")
                                        return
                                    nome_dossie_ativo = DADOS.casos_cadastrados.get(sketch_alvo["id_caso"], {}).get("nome") or f"Dossiê {sketch_alvo['id_caso']}"
                                    anexar_descoberta_a_caso_existente(
                                        df_dados, sketch_alvo["cluster_origem_id"], "", ipt_ns_motivo.value,
                                        nome_quadrilha_override=nome_dossie_ativo
                                    )
                                    DADOS.recarregar()
                                    ui.notify("Ocorrências fundidas ao sketch selecionado.")
                                    ctx.abrir_mesa_sketch(sketch_alvo["id_sketch"])

                                ui.button('Fundir Ocorrências', on_click=_exec_fundir).classes('w-full bg-[#222225] hover:bg-[#2B2B2F] text-white text-xs h-7 rounded-md mt-2')

                if not st.cockpit_ativo:
                    with ui.row().classes('lcfo-panel gap-1 px-2 py-1 self-center pointer-events-auto'):
                        for opc in OPCOES_TOOLBAR:
                            icn = ICONE_TOOLBAR.get(opc, "circle")
                            ativo_opc = (st.painel_unico == opc)
                            ui.button(icon=icn, color=None, on_click=lambda o=opc: (setattr(st, 'painel_unico', o), ctx.workspace.refresh())).props('flat dense round').classes(
                                f"lcfo-dock-btn {'lcfo-dock-btn-active' if ativo_opc else ''}"
                            ).tooltip(opc)

            if not st.cockpit_ativo:
                with ui.element('div').classes('w-full h-full mt-16'):
                    renderizar_painel_interno(st.painel_unico, lado="unico")
            else:
                with ui.splitter(value=st.cockpit_proporcao, on_change=lambda e: setattr(st, 'cockpit_proporcao', e.value)).classes('w-full h-full mt-[68px]') as splitter:
                    with splitter.before:
                        with ui.row().classes('viewport-header'):
                            with ui.button(f"{st.painel_esquerdo[:15]} ⌵", color=None).props('flat dense no-caps').classes('text-xs font-semibold text-[#FF7300]'):
                                with ui.menu().classes(CLASSE_MENU_PADRAO):
                                    for opc in OPCOES_TOOLBAR:
                                        ui.menu_item(opc, on_click=lambda o=opc: (setattr(st, 'painel_esquerdo', o), ctx.workspace.refresh())).classes('text-xs rounded-md hover:bg-[#222225]')
                            ui.button(icon='close', color=None, on_click=lambda: (setattr(st, 'cockpit_ativo', False), ctx.workspace.refresh())).props('flat dense round size=xs').classes('text-[#71717A] hover:text-white').tooltip('Fechar painel')
                        with ui.element('div').classes('w-full h-[calc(100%-32px)]'):
                            renderizar_painel_interno(st.painel_esquerdo, lado="esq")

                    with splitter.after:
                        with ui.row().classes('viewport-header'):
                            with ui.button(f"{st.painel_direito[:15]} ⌵", color=None).props('flat dense no-caps').classes('text-xs font-semibold text-[#FF7300]'):
                                with ui.menu().classes(CLASSE_MENU_PADRAO):
                                    for opc in OPCOES_TOOLBAR:
                                        ui.menu_item(opc, on_click=lambda o=opc: (setattr(st, 'painel_direito', o), ctx.workspace.refresh())).classes('text-xs rounded-md hover:bg-[#222225]')
                            ui.button(icon='close', color=None, on_click=lambda: (setattr(st, 'cockpit_ativo', False), ctx.workspace.refresh())).props('flat dense round size=xs').classes('text-[#71717A] hover:text-white').tooltip('Fechar painel')
                        with ui.element('div').classes('w-full h-[calc(100%-32px)]'):
                            renderizar_painel_interno(st.painel_direito, lado="dir")

        if st.mostrar_notas and dossie_persistido:
            ui.element('div').props('id=lcfo-notes-handle').classes(
                'h-full w-[5px] shrink-0 cursor-ew-resize bg-[#2B2B2F] hover:bg-[#FF7300]/60 transition-colors'
            )
            with ui.column().props('id=lcfo-notes-panel').classes(
                'h-full shrink-0 bg-[#18181A] border-l border-[#2B2B2F] overflow-y-auto custom-scroll p-4 gap-2'
            ).style('width:380px;'):
                painel_notas()
            ui.run_javascript("window.lcfoRegistrarNotasResize('lcfo-notes-panel', 'lcfo-notes-handle');")
