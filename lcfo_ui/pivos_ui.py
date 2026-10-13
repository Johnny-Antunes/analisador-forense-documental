"""
Interface dos pivôs (item 7): folha de parâmetros, execução fora do event loop com progresso
e cancelamento, painel Atividade (log, Cancelar, Desfazer, "Mostrar só o que este pivô
trouxe") e receitas. As regras ficam em lcfo_dominio/pivos.py.
"""

from __future__ import annotations

import json
from datetime import datetime
from typing import Any, Dict, List, Optional

from nicegui import background_tasks, run, ui

from lcfo_dominio import pivos, sketch_repo, tipos
from lcfo_ui import overlay

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina

CLASSE_DIALOGO = 'bg-[#18181A] border border-[#2B2B2F] rounded-xl p-4 gap-3'
ESTADOS = {"pendente": ("Pendente", "#8B95A5"), "executando": ("Executando", "#FF7300"), "concluida": ("Concluída", "#6FA98A"),
           "falha": ("Falha", "#C0625F"), "cancelada": ("Cancelada", "#8B95A5")}

# Execuções em andamento neste processo: progresso e pedido de cancelamento (cooperativo).
EM_ANDAMENTO: Dict[str, Dict[str, Any]] = {}


def _grafos(ctx: "Pagina", identificador: str) -> List[Any]:
    return [cg["grafo"] for cg in ctx.st.contexto_grafo_ativo.values() if cg.get("identificador_caso") == identificador]


def _entradas(identificador: str, ids: List[str], aceitos) -> List[dict]:
    campos = sketch_repo.carregar_campos(identificador, ids)
    out = []
    for n in dict.fromkeys(ids):
        t = tipos.tipo_do_id(n)
        if t and t.chave in aceitos:
            out.append({"node_id": n, "tipo": t.chave, "campos": campos.get(n, {})})
    return out


# ------------------------------------------------------------------ folha de parâmetros
def abrir_folha(ctx: "Pagina", contexto: Dict[str, Any], chave_pivo: str, node_ids: List[str]) -> None:
    identificador = contexto["identificador_caso"]
    if identificador == overlay.DESCOBERTA:
        ui.notify("A mesa de descoberta ainda não é um sketch: promova a caso para rodar pivôs com log.", type="warning")
        return
    pv = pivos.CATALOGO.get(chave_pivo)
    if not pv:
        return
    entradas = _entradas(identificador, node_ids, pv.entrada)
    if not entradas:
        ui.notify(f"{pv.rotulo} aceita: {', '.join(tipos.obter(t).rotulo for t in pv.entrada)}.", type="warning")
        return
    padrao = pivos.normalizar_parametros(pv, {})
    with ui.dialog() as dlg, ui.card().classes(CLASSE_DIALOGO + ' w-[440px] max-w-[94vw]'):
        ui.label(pv.rotulo).classes('text-sm font-bold text-gray-100')
        ui.label(f"{pv.descricao} · {len(entradas)} entrada(s) · custo {pv.custo}").classes('text-[11px] text-[#71717A]')
        campos: Dict[str, Any] = {}
        for p in pv.parametros:
            if p.nome == "expandir_hubs":
                continue
            if p.tipo == "int":
                campos[p.nome] = ui.number(p.rotulo, value=padrao[p.nome], min=p.minimo, max=p.maximo, step=1).props(
                    'dense outlined').classes('w-full text-xs')
            elif p.tipo == "tipos":
                opcoes = {t: tipos.obter(t).rotulo for t in pv.saidas}
                campos[p.nome] = ui.select(opcoes, value=padrao[p.nome], multiple=True, label=p.rotulo).props(
                    'dense outlined use-chips').classes('w-full text-xs')
            elif p.tipo == "data":
                campos[p.nome] = ui.input(p.rotulo, placeholder='aaaa-mm-dd').props('dense outlined type=date').classes('w-full text-xs')
        estimativa = ui.label('').classes('text-[11px] text-gray-300')
        hub_box = ui.column().classes('w-full gap-1')
        hub_ok = {"cb": None}

        def _params() -> dict:
            out = {k: (int(c.value) if isinstance(c, ui.number) and c.value is not None else c.value) for k, c in campos.items()}
            out["expandir_hubs"] = bool(hub_ok["cb"] and hub_ok["cb"].value)
            return out

        est = {"r": None}

        def _estimar():
            try:
                r = pivos.estimar(chave_pivo, entradas, _params())
            except Exception as e:
                estimativa.set_text(f"Estimativa indisponível: {e}")
                return
            est["r"] = r
            estimativa.set_text(f"Estimativa: {'' if r['exato'] else '~'}{r['itens']} item(ns)"
                                + (" — acima do teto: vai pedir confirmação" if r["acima_do_teto"] else ""))
            hub_box.clear()
            if r["hubs"]:
                with hub_box:
                    ui.label("Hub: " + ", ".join(f"{tipos.formatar_id(h['node_id'])} ({h['municipios']} municípios)" for h in r["hubs"][:3])
                             + " — não será expandido.").classes('text-[11px] text-[#C9A66B]')
                    hub_ok["cb"] = ui.checkbox('Expandir mesmo assim').props('dense').classes('text-xs text-gray-300')
        for c in campos.values():
            c.on_value_change(lambda _: _estimar())
        _estimar()

        with ui.row().classes('w-full justify-end gap-2'):
            ui.button('Cancelar', color=None, on_click=dlg.close).props('flat dense no-caps').classes('text-xs text-gray-400')
            btn = ui.button('Rodar', color=None).props('dense no-caps').classes('text-xs bg-[#FF7300] text-white rounded-md px-3')
            confirmado = {"v": False}

            def _rodar():
                r = est["r"]
                if r and r["acima_do_teto"] and not confirmado["v"]:
                    confirmado["v"] = True
                    btn.set_text(f"Rodar mesmo assim (~{r['itens']})")
                    return
                dlg.close()
                rodar(ctx, contexto, chave_pivo, entradas, _params())
            btn.on_click(_rodar)
    dlg.open()


# ------------------------------------------------------------------ execução
def rodar(ctx: "Pagina", contexto: Dict[str, Any], chave_pivo: str, entradas: List[dict], params: Optional[dict] = None) -> str:
    """Dispara a execução numa thread; a tela segue respondendo e a Atividade mostra o progresso."""
    identificador = contexto["identificador_caso"]
    id_exec = pivos.iniciar_execucao(identificador, chave_pivo, entradas, params)
    reg = EM_ANDAMENTO[id_exec] = {"progresso": (0, 1, "na fila"), "cancelar": False, "id_sketch": identificador}
    container = ui.context.slot.parent
    ctx.st.aba_painel_lateral = "Atividade"
    ctx.sidebar_container.refresh()

    def _progresso(a, t, msg):
        reg["progresso"] = (a, t, msg)

    async def _tarefa():
        try:
            res = await run.io_bound(pivos.executar, identificador, chave_pivo, entradas, params, progresso=_progresso,
                                     cancelado=lambda: reg["cancelar"], id_exec=id_exec)
        except Exception as e:
            res = {"estado": "falha", "erro": str(e)}
        finally:
            EM_ANDAMENTO.pop(id_exec, None)
        if res is None:                                       # app encerrando
            return
        with container:
            _concluir(ctx, identificador, chave_pivo, res)
    background_tasks.create(_tarefa(), name=f"pivo-{id_exec}")
    return id_exec


def _concluir(ctx: "Pagina", identificador: str, chave_pivo: str, res: dict) -> None:
    rot = pivos.CATALOGO[chave_pivo].rotulo
    if res.get("estado") == "falha":
        ui.notify(f"{rot}: falhou — {res.get('erro', '')}", type="negative")
    elif res.get("estado") == "cancelada":
        ui.notify(f"{rot}: cancelado (nada foi gravado).")
    else:
        ov = overlay.carregar(ctx.st, identificador, forcar=True)
        novos = [ov["nodes"][n] for n in res["nos_novos"] + res["reativados"] if n in ov["nodes"]]
        arestas = [ov["edges"][v] for v in res["vinculos_novos"] + res["reativados"] if v in ov["edges"]]
        for g in _grafos(ctx, identificador):
            g.enviar_patch({"add": {"nodes": novos, "edges": arestas}})
        partes = [f"{len(res['nos_novos'])} nó(s) novo(s)", f"{len(res['vinculos_novos'])} vínculo(s)"]
        if res["existentes"]:
            partes.append(f"{len(res['existentes'])} já existia(m)")
        ui.notify(f"{rot}: " + " · ".join(partes) + (" (cache)" if res.get("cache_hit") else ""), type="positive")
        if res.get("hubs"):
            ui.notify("Hub não expandido: " + ", ".join(tipos.formatar_id(h["node_id"]) for h in res["hubs"][:3])
                      + ". Rode de novo marcando 'Expandir mesmo assim'.", type="warning")
        for a in res.get("avisos") or []:
            ui.notify(a, type="warning")
    ctx.sidebar_container.refresh()


def cancelar(id_exec: str) -> None:
    if id_exec in EM_ANDAMENTO:
        EM_ANDAMENTO[id_exec]["cancelar"] = True


def desfazer(ctx: "Pagina", identificador: str, id_exec: str) -> None:
    r = pivos.desfazer_execucao(identificador, id_exec)
    overlay.carregar(ctx.st, identificador, forcar=True)
    for g in _grafos(ctx, identificador):
        g.enviar_patch({"remove": {"nodes": r["ids_nos"], "edges": r["ids_vinculos"]}})
    ui.notify(f"Execução desfeita: {r['nos']} nó(s) e {r['vinculos']} vínculo(s) foram para a lixeira.")
    ctx.sidebar_container.refresh()


def mostrar_so(ctx: "Pagina", identificador: str, ex: dict) -> None:
    rot = pivos.CATALOGO[ex["pivo"]].rotulo if ex["pivo"] in pivos.CATALOGO else ex["pivo"]
    for g in _grafos(ctx, identificador):
        g.enviar_mensagem({"type": "lcfo:filtro_execucao", "id": ex["id"], "rotulo": rot,
                           "entradas": json.loads(ex.get("entradas_json") or "[]")})


# ------------------------------------------------------------------ aba Atividade
def _fmt_hora(iso: Optional[str]) -> str:
    try:
        return datetime.fromisoformat(iso).strftime("%d/%m %H:%M")
    except Exception:
        return iso or ""


def render_aba_atividade(ctx: "Pagina", identificador: str) -> None:
    if identificador == overlay.DESCOBERTA:
        ui.label('A mesa de descoberta não tem log de execuções: promova a caso primeiro.').classes('text-xs text-[#71717A] px-1 py-2')
        return

    @ui.refreshable
    def lista():
        exs = pivos.listar_execucoes(identificador, limite=50)
        if not exs:
            ui.label('Nenhum pivô rodado neste sketch. Clique com o botão direito num nó para escolher um.').classes(
                'text-xs text-[#71717A] px-1 py-2')
        for ex in exs:
            rot = pivos.CATALOGO[ex["pivo"]].rotulo if ex["pivo"] in pivos.CATALOGO else ex["pivo"]
            nome_est, cor = ESTADOS.get(ex["estado"], (ex["estado"], "#8B95A5"))
            andamento = EM_ANDAMENTO.get(ex["id"])
            with ui.column().classes('w-full gap-0.5 p-2 rounded-md bg-[#141416]/70 border border-[#2B2B2F]'):
                with ui.row().classes('w-full items-center justify-between no-wrap'):
                    ui.label(rot + (" · receita" if ex.get("id_receita_exec") else "")).classes('text-xs font-semibold text-gray-200 truncate')
                    ui.label(nome_est).classes('text-[10px] font-bold shrink-0').style(f'color:{cor}')
                if andamento:
                    a, t, msg = andamento["progresso"]
                    ui.linear_progress(value=(a / t) if t else 0, show_value=False).props('color=orange rounded size=4px').classes('w-full')
                    ui.label(msg).classes('text-[10px] text-[#71717A]')
                elif ex["estado"] == "concluida":
                    partes = [f"{ex['n_nos_novos'] or 0} nó(s) novo(s)", f"{ex['n_vinculos_novos'] or 0} vínculo(s)"]
                    if ex["n_existentes"]:
                        partes.append(f"{ex['n_existentes']} já existia(m)")
                    partes.append(f"{ex['ms'] or 0} ms" + (" · cache" if ex["cache_hit"] else ""))
                    ui.label(" · ".join(partes)).classes('text-[11px] text-gray-400')
                    avisos = json.loads(ex.get("avisos_json") or "{}")
                    if avisos.get("hubs"):
                        ui.label("Hub não expandido: " + ", ".join(tipos.formatar_id(h["node_id"]) for h in avisos["hubs"][:2])).classes(
                            'text-[10px] text-[#C9A66B]')
                elif ex["estado"] == "falha":
                    ui.label(ex.get("erro") or "").classes('text-[10px] text-[#C0625F]')
                ui.label(f"{ex.get('usuario') or '—'} · {_fmt_hora(ex.get('iniciado_em'))}").classes('text-[10px] text-[#71717A]')
                with ui.row().classes('gap-1 self-end'):
                    if andamento:
                        ui.button('Cancelar', color=None, on_click=lambda i=ex["id"]: cancelar(i)).props('flat dense no-caps size=sm').classes(
                            'text-[11px] text-[#C0625F]')
                    elif ex["estado"] == "concluida":
                        ui.button('Mostrar só', color=None, on_click=lambda e=ex: mostrar_so(ctx, identificador, e)).props(
                            'flat dense no-caps size=sm').classes('text-[11px] text-gray-300').tooltip('Mostrar só o que este pivô trouxe')
                        if ex["n_nos_novos"] or ex["n_vinculos_novos"]:
                            ui.button('Desfazer', color=None, on_click=lambda i=ex["id"]: desfazer(ctx, identificador, i)).props(
                                'flat dense no-caps size=sm').classes('text-[11px] text-[#FF7300]')

    with ui.column().classes('w-full gap-1.5 flex-1 min-h-0 overflow-y-auto custom-scroll -mr-3 pr-2 no-wrap'):
        _render_receitas(ctx, identificador)
        ui.label('Execuções').classes('text-[10px] font-semibold text-[#71717A] uppercase tracking-wider mt-1')
        lista()
    # enquanto houver execução deste sketch rodando, a lista se atualiza sozinha
    ui.timer(0.5, lambda: lista.refresh() if any(r["id_sketch"] == identificador for r in EM_ANDAMENTO.values()) else None)


# ------------------------------------------------------------------ receitas
def _selecao(ctx: "Pagina", identificador: str) -> List[str]:
    sel = list(ctx.st.selecao_entidades.get(identificador) or [])
    if not sel and ctx.st.entidade_foco:
        sel = [ctx.st.entidade_foco]
    return sel


def _render_receitas(ctx: "Pagina", identificador: str) -> None:
    with ui.expansion('Receitas', icon='playlist_play').props('dense').classes('w-full text-xs text-gray-300 bg-[#141416]/70 rounded-md'):
        ui.label('Rodam sobre a seleção da gaveta (ou o nó em foco).').classes('text-[10px] text-[#71717A] px-1')
        for rec in pivos.listar_receitas():
            passos = " → ".join(pivos.CATALOGO[p["pivo"]].rotulo + (f" ({', '.join(p.get('params', {}).get('tipos_saida', []))})"
                                                                      if p.get("params", {}).get("tipos_saida") else "")
                                for p in rec["passos"] if p["pivo"] in pivos.CATALOGO)
            with ui.column().classes('w-full gap-0 p-1.5 rounded-md border border-[#2B2B2F]'):
                ui.label(f"{rec['nome']} · entrada: {tipos.obter(rec['entrada']).rotulo}").classes('text-xs text-gray-200')
                ui.label(passos).classes('text-[10px] text-[#71717A]')
                saida = ui.label('').classes('text-[10px] text-gray-400')
                with ui.row().classes('gap-1 self-end'):
                    def _calcular(rec=rec, saida=saida):
                        ents = _entradas(identificador, _selecao(ctx, identificador), {rec["entrada"]})
                        if not ents:
                            saida.set_text(f"Selecione entidades do tipo {tipos.obter(rec['entrada']).rotulo}.")
                            return
                        calc = pivos.calcular_receita(rec, ents)
                        saida.set_text(" · ".join(f"{c['passo']}. {c['pivo']}: {c['itens']}" for c in calc))

                    def _executar(rec=rec, saida=saida):
                        ents = _entradas(identificador, _selecao(ctx, identificador), {rec["entrada"]})
                        if not ents:
                            saida.set_text(f"Selecione entidades do tipo {tipos.obter(rec['entrada']).rotulo}.")
                            return
                        rodar_receita(ctx, identificador, rec, ents)
                    ui.button('Calcular', color=None, on_click=_calcular).props('flat dense no-caps size=sm').classes('text-[11px] text-gray-300')
                    ui.button('Executar', color=None, on_click=_executar).props('flat dense no-caps size=sm').classes('text-[11px] text-[#FF7300]')
        ui.button('Nova receita', icon='add', color=None, on_click=lambda: abrir_editor_receita(ctx)).props(
            'flat dense no-caps size=sm').classes('text-[11px] text-[#FF7300]')


def rodar_receita(ctx: "Pagina", identificador: str, rec: dict, entradas: List[dict]) -> None:
    container = ui.context.slot.parent
    chave = f"receita_{rec['id']}"
    reg = EM_ANDAMENTO[chave] = {"progresso": (0, 1, "na fila"), "cancelar": False, "id_sketch": identificador}

    async def _tarefa():
        try:
            out = await run.io_bound(pivos.executar_receita, identificador, rec, entradas,
                                     progresso=lambda a, t, m: reg.__setitem__("progresso", (a, t, m)), cancelado=lambda: reg["cancelar"])
        finally:
            EM_ANDAMENTO.pop(chave, None)
        if out is None:
            return
        with container:
            for ps in out["passos"]:
                _concluir(ctx, identificador, ps["pivo"], ps)
    background_tasks.create(_tarefa(), name=chave)
    ui.notify(f"Receita '{rec['nome']}' em execução.")


def abrir_editor_receita(ctx: "Pagina") -> None:
    passos: List[dict] = []
    with ui.dialog() as dlg, ui.card().classes(CLASSE_DIALOGO + ' w-[480px] max-w-[94vw]'):
        ui.label('Nova receita').classes('text-sm font-bold text-gray-100')
        nome = ui.input('Nome').props('dense outlined').classes('w-full text-xs')
        entrada = ui.select({t.chave: t.rotulo for t in tipos.REGISTRO.values() if pivos.pivos_para(t.chave)}, value='cpf',
                            label='Tipo de entrada').props('dense outlined').classes('w-full text-xs')
        lista = ui.column().classes('w-full gap-1')
        erro = ui.label('').classes('text-[11px] text-[#C9A66B]')

        def _disponiveis() -> set:
            disp = {entrada.value}
            for ps in passos:
                pv = pivos.CATALOGO[ps["pivo"]]
                disp = set(pivos.normalizar_parametros(pv, ps["params"]).get("tipos_saida") or pv.saidas)
            return disp

        def _redesenhar():
            lista.clear()
            with lista:
                for i, ps in enumerate(passos, 1):
                    pv = pivos.CATALOGO[ps["pivo"]]
                    ts = ps["params"].get("tipos_saida")
                    ui.label(f"{i}. {pv.rotulo}" + (f" → {', '.join(tipos.obter(t).rotulo for t in ts)}" if ts else "")).classes('text-xs text-gray-200')
                disp = _disponiveis()
                compat = {p.chave: p.rotulo for p in pivos.CATALOGO.values() if disp & set(p.entrada)}
                if compat:
                    with ui.row().classes('w-full items-center gap-2 no-wrap'):
                        sel = ui.select(compat, label='Adicionar passo').props('dense outlined').classes('flex-1 text-xs')
                        saidas = ui.select({}, multiple=True, label='Trazer').props('dense outlined use-chips').classes('flex-1 text-xs')

                        def _ao_escolher(e, saidas=saidas):
                            pv = pivos.CATALOGO.get(e.value)
                            saidas.options = {t: tipos.obter(t).rotulo for t in (pv.saidas if pv else ())}
                            saidas.value = list(pv.saidas) if pv else []
                            saidas.update()
                        sel.on_value_change(_ao_escolher)

                        def _add(sel=sel, saidas=saidas):
                            if not sel.value:
                                return
                            params = {"tipos_saida": list(saidas.value)} if saidas.value else {}
                            passos.append({"pivo": sel.value, "params": params})
                            _redesenhar()
                        ui.button(icon='add', color=None, on_click=_add).props('flat dense round size=sm').classes('text-[#FF7300]')
        entrada.on_value_change(lambda _: (passos.clear(), _redesenhar()))
        _redesenhar()
        with ui.row().classes('w-full justify-end gap-2'):
            ui.button('Cancelar', color=None, on_click=dlg.close).props('flat dense no-caps').classes('text-xs text-gray-400')

            def _salvar():
                try:
                    pivos.salvar_receita(nome.value or "Receita", entrada.value, passos)
                except ValueError as e:
                    erro.set_text(str(e))
                    return
                dlg.close()
                ctx.sidebar_container.refresh()
                ui.notify("Receita salva.", type="positive")
            ui.button('Salvar', color=None, on_click=_salvar).props('dense no-caps').classes('text-xs bg-[#FF7300] text-white rounded-md px-3')
    dlg.open()
