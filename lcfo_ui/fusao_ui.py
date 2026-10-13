"""
Interface da fusão de entidades (item 6): diálogo "Fundir…" (prioritária, tabela campo a campo,
prévia dos vínculos), "Desfazer fusão" e a lista "Possíveis duplicatas" da gaveta Entities.
As regras ficam em lcfo_dominio/fusao.py.
"""

from __future__ import annotations

from typing import Any, Dict, List

from nicegui import ui

from lcfo_dominio import fusao, sketch_repo, tipos
from lcfo_ui import overlay
from lcfo_ui.dados import DADOS

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina

CLASSE_DIALOGO = 'bg-[#18181A] border border-[#2B2B2F] rounded-xl p-4 gap-3'
SEGUNDOS_DESFAZER = 8


def _grafos(ctx: "Pagina", identificador: str) -> List[Any]:
    return [cg["grafo"] for cg in ctx.st.contexto_grafo_ativo.values() if cg.get("identificador_caso") == identificador]


def _contexto_do_sketch(ctx: "Pagina", identificador: str):
    return next((cg for cg in ctx.st.contexto_grafo_ativo.values() if cg.get("identificador_caso") == identificador), None)


def _campos_oficiais(node_id: str) -> Dict[str, Any]:
    if DADOS.G is not None and node_id in DADOS.G:
        nome = DADOS.G.nodes[node_id].get("nome_titular")
        return {"nome_titular": nome} if nome and nome != "N/D" else {}
    return {}


def _rotulo(ctx: "Pagina", identificador: str, node_id: str) -> str:
    no = overlay.carregar(ctx.st, identificador)["nodes"].get(node_id)
    if no and no.get("label"):
        return f"{no['label']} · {tipos.formatar_id(node_id)}" if no["label"] != tipos.formatar_id(node_id) else no["label"]
    nome = _campos_oficiais(node_id).get("nome_titular")
    return f"{nome} · {tipos.formatar_id(node_id)}" if nome else tipos.formatar_id(node_id)


def _n_vinculos(node_id: str, ov: Dict[str, Any]) -> int:
    n = sum(1 for e in ov["edges"].values() if node_id in (e["from"], e["to"]))
    if DADOS.G is not None and node_id in DADOS.G:
        n += DADOS.G.degree(node_id)
    return n


# ------------------------------------------------------------------ diálogo
def abrir_dialogo_fusao(ctx: "Pagina", contexto: Dict[str, Any], ids: List[str]) -> None:
    identificador = contexto["identificador_caso"]
    ids = list(dict.fromkeys(ids))
    if identificador == overlay.DESCOBERTA:
        ui.notify("A mesa de descoberta ainda não é um sketch: promova a caso para fundir.", type="warning")
        return
    if len(ids) < 2:
        ui.notify("Selecione duas ou mais entidades do mesmo tipo.", type="warning")
        return
    tps = {(tipos.tipo_do_id(n).chave if tipos.tipo_do_id(n) else "?") for n in ids}
    if len(tps) > 1:
        ui.notify("Só entidades do mesmo tipo podem ser fundidas.", type="warning")
        return
    oficiais = contexto.get("ids_oficiais", set())
    ofi_sel = [n for n in ids if n in oficiais]
    if len(ofi_sel) > 1:
        _oferecer_mesma_entidade(ctx, contexto, ofi_sel)
        return

    ov = overlay.carregar(ctx.st, identificador)
    campos_sk = sketch_repo.carregar_campos(identificador, ids)
    info = [{"node_id": n, "oficial": n in oficiais,
             "campos": campos_sk.get(n) or (_campos_oficiais(n) if n in oficiais else {}),
             "n_vinculos": _n_vinculos(n, ov)} for n in ids]
    t = tipos.tipo_do_id(ids[0])
    nomes_campos = list(dict.fromkeys([c.nome for c in (t.campos if t else ()) if not c.primario] +
                                      [k for i in info for k in i["campos"]]))
    rotulos_campo = {c.nome: c.rotulo for c in (t.campos if t else ())}

    with ui.dialog() as dlg, ui.card().classes(CLASSE_DIALOGO + ' w-[600px] max-w-[94vw]'):
        ui.label(f'Fundir {len(ids)} entidades · {t.rotulo if t else ""}').classes('text-sm font-bold text-gray-100')
        ui.label('A prioritária fica (o id dela não muda); as outras vão para a lixeira e seus vínculos passam para ela.').classes(
            'text-[11px] text-[#71717A]')
        ui.label('Prioritária').classes('text-[10px] font-semibold text-[#71717A] uppercase tracking-wider')
        opcoes = {i["node_id"]: _rotulo(ctx, identificador, i["node_id"]) + (" · base oficial" if i["oficial"] else "")
                  + f" · {i['n_vinculos']} vínculo(s)" for i in info}
        radio = ui.radio(opcoes, value=fusao.escolher_prioritaria(info)).props('dense').classes('text-xs text-gray-200')

        escolhas: Dict[str, Any] = {}
        tabela = ui.column().classes('w-full gap-1')
        previa_lbl = ui.label('').classes('text-[11px] text-gray-300')

        def _montar():
            tabela.clear()
            prior = radio.value
            with tabela:
                distintos = [c for c in nomes_campos if len({str(i["campos"].get(c)) for i in info if i["campos"].get(c)}) > 0]
                if distintos:
                    ui.label('Campos (padrão: o da prioritária; vazios dela são preenchidos)').classes(
                        'text-[10px] font-semibold text-[#71717A] uppercase tracking-wider mt-1')
                for c in distintos:
                    valores = list(dict.fromkeys(str(i["campos"][c]) for i in info if i["campos"].get(c)))
                    padrao = next((str(i["campos"][c]) for i in info if i["node_id"] == prior and i["campos"].get(c)), valores[0])
                    with ui.row().classes('w-full items-center gap-2 no-wrap'):
                        ui.label(rotulos_campo.get(c, c)).classes('text-xs text-[#A1A1AA] w-32 shrink-0')
                        if len(valores) == 1:
                            ui.label(valores[0]).classes('text-xs text-gray-200')
                            escolhas[c] = valores[0]
                        else:
                            escolhas[c] = padrao
                            ui.radio(valores, value=padrao, on_change=lambda e, c=c: escolhas.__setitem__(c, e.value)).props(
                                'dense inline').classes('text-xs text-gray-200')
            p = fusao.previa(identificador, prior, [n for n in ids if n != prior])
            previa_lbl.set_text(f"{p['redirecionados']} vínculo(s) redirecionado(s) · {p['unidos']} par(es) igual(is) unido(s) · "
                                f"{p['descartados']} autolaço(s) descartado(s)")
        radio.on_value_change(lambda _: _montar())
        _montar()

        with ui.row().classes('w-full justify-end gap-2'):
            ui.button('Cancelar', color=None, on_click=dlg.close).props('flat dense no-caps').classes('text-xs text-gray-400')

            def _confirmar():
                prior = radio.value
                absorvidas = [n for n in ids if n != prior]
                if any(a in oficiais for a in absorvidas):
                    ui.notify("Entidade da base oficial não pode ser absorvida: escolha-a como prioritária.", type="warning")
                    return
                try:
                    r = fusao.fundir(identificador, prior, absorvidas, oficiais=oficiais, campos_escolhidos=dict(escolhas),
                                     campos_oficiais={n: _campos_oficiais(n) for n in ids if n in oficiais})
                except fusao.ErroFusao as e:
                    ui.notify(str(e), type="warning")
                    return
                dlg.close()
                _aplicar_fusao(ctx, identificador, r)
            ui.button('Fundir', color=None, on_click=_confirmar).props('dense no-caps').classes(
                'text-xs bg-[#FF7300] text-white rounded-md px-3')
    dlg.open()


def _aplicar_fusao(ctx: "Pagina", identificador: str, r: Dict[str, Any]) -> None:
    ov = overlay.carregar(ctx.st, identificador, forcar=True)
    novas = [ov["edges"][v] for v in r["vinculos_criados"] if v in ov["edges"]]
    patch = {"remove": {"nodes": r["absorvidas"], "edges": r["vinculos_removidos"]}, "add": {"nodes": [], "edges": novas},
             "focus": r["prioritaria"]}
    if r["prioritaria"] in ov["nodes"]:
        patch["update"] = {"nodes": [ov["nodes"][r["prioritaria"]]], "edges": []}
    for g in _grafos(ctx, identificador):
        g.enviar_patch(patch)
    for sel in ctx.st.selecao_entidades.values():
        sel.difference_update(r["absorvidas"])
    ctx.sidebar_container.refresh()
    _aviso_desfazer(ctx, identificador, r)


def _aviso_desfazer(ctx: "Pagina", identificador: str, r: Dict[str, Any]) -> None:
    with ui.dialog().props('seamless position=bottom') as aviso, ui.card().classes(
            'bg-[#222225] border border-[#3A3A40] rounded-lg px-4 py-2 mb-6 shadow-2xl'):
        with ui.row().classes('items-center gap-4 no-wrap'):
            ui.icon('call_merge', size='18px').classes('text-[#FF7300]')
            ui.label(f"{len(r['absorvidas']) + 1} entidades fundidas · {r['redirecionados']} vínculo(s) redirecionado(s)").classes(
                'text-xs text-gray-200')
            ui.button('Desfazer fusão', color=None, on_click=lambda: (aviso.close(), desfazer(ctx, identificador, r["lote"]))).props(
                'flat dense no-caps').classes('text-xs font-bold text-[#FF7300]')
    aviso.open()
    ui.timer(SEGUNDOS_DESFAZER, aviso.close, once=True)


def desfazer(ctx: "Pagina", identificador: str, lote: str) -> None:
    try:
        d = fusao.desfazer_fusao(identificador, lote)
    except fusao.ErroFusao as e:
        ui.notify(str(e), type="warning")
        return
    ov = overlay.carregar(ctx.st, identificador, forcar=True)
    patch = {"remove": {"nodes": [], "edges": d["vinculos_removidos"]},
             "add": {"nodes": [ov["nodes"][n] for n in d["ids_nos"] if n in ov["nodes"]],
                     "edges": [ov["edges"][v] for v in d["ids_vinculos"] if v in ov["edges"]]}}
    if d["prioritaria"] in ov["nodes"]:
        patch["update"] = {"nodes": [ov["nodes"][d["prioritaria"]]], "edges": []}
    for g in _grafos(ctx, identificador):
        g.enviar_patch(patch)
    ctx.sidebar_container.refresh()
    ui.notify("Fusão desfeita: o grafo voltou ao estado anterior.", type="positive")


def _oferecer_mesma_entidade(ctx: "Pagina", contexto: Dict[str, Any], ofi: List[str]) -> None:
    identificador = contexto["identificador_caso"]
    with ui.dialog() as dlg, ui.card().classes(CLASSE_DIALOGO + ' w-[420px]'):
        ui.label('Duas entidades da base oficial').classes('text-sm font-bold text-gray-100')
        ui.label('Dado oficial não é fundido. Dá para registrar a suspeita como vínculo hipótese "MESMA_ENTIDADE", '
                 'preservando a evidência de cada uma.').classes('text-xs text-gray-300')
        with ui.row().classes('w-full justify-end gap-2'):
            ui.button('Cancelar', color=None, on_click=dlg.close).props('flat dense no-caps').classes('text-xs text-gray-400')

            def _vincular():
                dlg.close()
                if len(ofi) != 2:
                    ui.notify("Selecione exatamente duas entidades oficiais para vincular.", type="warning")
                    return
                vid = fusao.vincular_mesma_entidade(identificador, ofi[0], ofi[1])
                overlay.carregar(ctx.st, identificador, forcar=True)
                contexto["grafo"].enviar_patch({"add": {"nodes": [], "edges": [sketch_repo.carregar_vinculo(identificador, vid)]}})
                ui.notify("Vínculo hipótese MESMA_ENTIDADE gravado.", type="positive")
            ui.button('Vincular como mesma entidade', color=None, on_click=_vincular).props('dense no-caps').classes(
                'text-xs bg-[#FF7300] text-white rounded-md px-3')
    dlg.open()


# ------------------------------------------------------------------ possíveis duplicatas (gaveta)
def render_duplicatas(ctx: "Pagina", identificador: str, ids: List[str]) -> None:
    if identificador == overlay.DESCOBERTA:
        return
    campos = sketch_repo.carregar_campos(identificador, ids)
    sug = fusao.sugerir_duplicatas(identificador, ids, campos=campos, descartados=fusao.pares_descartados(identificador))
    if not sug:
        return
    with ui.expansion(f'Possíveis duplicatas ({len(sug)})', icon='content_copy').props('dense').classes(
            'w-full text-xs text-[#C9A66B] bg-[#141416]/70 rounded-md mb-2'):
        ui.label('Sugestões para revisar: a fusão é sempre decisão sua.').classes('text-[10px] text-[#71717A] px-1')
        with ui.column().classes('w-full gap-1 max-h-60 overflow-y-auto custom-scroll'):
            for s in sug[:50]:
                with ui.column().classes('w-full gap-0.5 p-1.5 rounded-md border border-[#2B2B2F]'):
                    ui.label(f"{tipos.formatar_id(s['a'])}  ×  {tipos.formatar_id(s['b'])}").classes('text-xs text-gray-200')
                    ui.label(s["motivo"]).classes('text-[10px] text-[#71717A]')
                    with ui.row().classes('gap-1'):
                        def _fundir(s=s):
                            cg = _contexto_do_sketch(ctx, identificador)
                            if cg:
                                abrir_dialogo_fusao(ctx, cg, [s["a"], s["b"]])

                        def _vizinhos(s=s):
                            vid = sketch_repo.criar_vinculo(identificador, s["a"], s["b"], rotulo="vizinhos", direcao="nenhuma")
                            fusao.descartar_sugestao(identificador, s["a"], s["b"])
                            overlay.carregar(ctx.st, identificador, forcar=True)
                            for g in _grafos(ctx, identificador):
                                g.enviar_patch({"add": {"nodes": [], "edges": [sketch_repo.carregar_vinculo(identificador, vid)]}})
                            ctx.sidebar_container.refresh()

                        def _descartar(s=s):
                            fusao.descartar_sugestao(identificador, s["a"], s["b"])
                            ctx.sidebar_container.refresh()
                        for rot, fn in (("Fundir", _fundir), ("Vincular como vizinhos", _vizinhos), ("Descartar", _descartar)):
                            ui.button(rot, color=None, on_click=fn).props('flat dense no-caps size=sm').classes(
                                'text-[10px] text-[#FF7300]' if rot == "Fundir" else 'text-[10px] text-gray-400')
