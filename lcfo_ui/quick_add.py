"""
Quick-add no canvas e colar com detecção de tipo (item 2).

O iframe abre o campo (duplo clique no fundo ou tecla N), pede a detecção a cada pausa
(quick_add_detect) e, ao confirmar, desenha um nó provisório tracejado (tmp_...) e manda
quick_add_commit. Aqui se valida, grava (sketch_repo) e responde com um patch que troca o
provisório pelo definitivo — ou o remove, com aviso. A regra de detecção existe só no Python
(lcfo_dominio.tipos.detectar), não duplicada em JavaScript.

Colar (Ctrl+V no canvas) manda as linhas; aqui abre a prévia: tipo detectado por linha,
editável, ambíguas em destaque e exigindo escolha, e o botão "Adicionar N".
"""

from __future__ import annotations

import math
from typing import Any, Dict, List

from nicegui import ui

from lcfo_dominio import sketch_repo, tipos
from lcfo_ui import overlay
from lcfo_ui.servicos import obter_set_nos_ocultos, restaurar_no_do_caso

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina

MAX_LINHAS_COLAR = 500
ESPACO_GRADE = 110
CLASSE_DIALOGO = 'bg-[#18181A] border border-[#2B2B2F] rounded-xl p-4 gap-3'


def _candidatos_json(texto: str) -> List[Dict[str, Any]]:
    return [{"tipo": c.tipo, "rotulo": tipos.obter(c.tipo).rotulo, "confianca": c.confianca, "valor": c.valor,
             "valor_fmt": tipos.obter(c.tipo).formatar(c.valor)} for c in tipos.detectar(texto)]


def ambiguo(cands: List[Dict[str, Any]]) -> bool:
    """Dois ou mais tipos com confiança próxima (ex.: CPF × celular de 11 dígitos): o analista escolhe."""
    return len(cands) > 1 and cands[0]["confianca"] - cands[1]["confianca"] <= 10


def campos_obrigatorios_extras(tipo: str) -> List[tipos.Campo]:
    t = tipos.obter(tipo)
    return [c for c in (t.campos if t else ()) if c.obrigatorio and not c.primario]


# ------------------------------------------------------------------ detecção
def tratar_detect(ctx: "Pagina", contexto: Dict[str, Any], valor: Dict[str, Any]) -> None:
    texto = str(valor.get("texto") or "")[:300]
    cands = _candidatos_json(texto)
    contexto["grafo"].enviar_mensagem({
        "type": "lcfo:quick_add_result", "req_id": valor.get("req_id"), "candidatos": cands,
        "ambiguo": ambiguo(cands),
        # sem candidato: a interface oferece a lista de tipos (todos do registro, não só os presentes)
        "tipos": [] if cands else [{"chave": t.chave, "rotulo": t.rotulo, "cor": t.cores} for t in tipos.REGISTRO.values()],
    })


# ------------------------------------------------------------------ gravação
def _no_do_overlay(ctx: "Pagina", identificador: str, node_id: str) -> Dict[str, Any]:
    return overlay.carregar(ctx.st, identificador, forcar=True)["nodes"].get(node_id)


def _responder(contexto: Dict[str, Any], tmp_id: str, no: Dict[str, Any] = None, foco: str = None) -> None:
    patch: Dict[str, Any] = {"remove": {"nodes": [tmp_id] if tmp_id else [], "edges": []}}
    if no:
        patch["add"] = {"nodes": [no], "edges": []}
    if foco:
        patch["focus"] = foco
    contexto["grafo"].enviar_patch(patch)


def tratar_commit(ctx: "Pagina", contexto: Dict[str, Any], valor: Dict[str, Any]) -> None:
    identificador = contexto["identificador_caso"]
    tmp_id, tipo, texto = valor.get("tmp_id") or "", valor.get("tipo_entidade") or "", str(valor.get("valor") or "")
    pos = {"x": float(valor.get("x") or 0), "y": float(valor.get("y") or 0)}
    if identificador == overlay.DESCOBERTA:
        _responder(contexto, tmp_id)
        ui.notify("A mesa de descoberta ainda não é um sketch: promova a caso para criar entidades.", type="warning")
        return
    try:
        t, canon, node_id = sketch_repo.canonico(tipo, texto)
    except ValueError as e:
        _responder(contexto, tmp_id)
        ui.notify(str(e), type="warning")
        return

    # Já existe no grafo (oficial ou do analista): não duplica, foca o existente.
    if node_id in contexto.get("ids_oficiais", set()):
        alvo_ocultos = contexto.get("chave_layout") or identificador
        if node_id in obter_set_nos_ocultos(alvo_ocultos):
            _responder(contexto, tmp_id)
            _oferecer_reativar(ctx, contexto, alvo_ocultos, node_id)
            return
        _responder(contexto, tmp_id, foco=node_id)
        ui.notify(f"{t.rotulo} {t.formatar(canon)} já está no sketch (base oficial).")
        return
    if node_id in overlay.carregar(ctx.st, identificador)["nodes"]:
        _responder(contexto, tmp_id, foco=node_id)
        ui.notify(f"{t.rotulo} {t.formatar(canon)} já está no sketch.")
        return

    extras = campos_obrigatorios_extras(t.chave)
    if extras and not valor.get("campos"):
        _responder(contexto, tmp_id)
        abrir_formulario(ctx, contexto, t.chave, texto, pos)
        return
    _criar_e_mostrar(ctx, contexto, t.chave, texto, valor.get("campos") or {}, pos, tmp_id)


def _criar_e_mostrar(ctx, contexto, tipo, texto, campos, pos, tmp_id="") -> None:
    identificador = contexto["identificador_caso"]
    try:
        r = sketch_repo.criar_entidades(identificador, [{"tipo": tipo, "valor": texto, "campos": campos, "posicao": pos}])[0]
    except Exception as e:                                     # falha de gravação: o provisório some
        _responder(contexto, tmp_id)
        ui.notify(f"Não foi possível gravar: {e}", type="negative")
        return
    no = _no_do_overlay(ctx, identificador, r["node_id"])
    _responder(contexto, tmp_id, no, foco=r["node_id"])
    if r["estado"] == "reativado":
        ui.notify("Estava na lixeira do sketch: restaurado.", type="positive")
    ctx.sidebar_container.refresh()


def _oferecer_reativar(ctx, contexto, alvo_ocultos: str, node_id: str) -> None:
    with ui.dialog() as dlg, ui.card().classes(CLASSE_DIALOGO + ' w-[380px]'):
        ui.label('Entidade isolada').classes('text-sm font-bold text-gray-100')
        ui.label(f"{tipos.formatar_id(node_id)} é da base oficial e foi isolada deste sketch.").classes('text-xs text-gray-300')
        with ui.row().classes('w-full justify-end gap-2'):
            ui.button('Cancelar', color=None, on_click=dlg.close).props('flat dense no-caps').classes('text-xs text-gray-400')

            def _reativar():
                restaurar_no_do_caso(alvo_ocultos, node_id)
                dlg.close()
                ctx.st.entidade_foco = node_id
                layout_atual = ctx.st.layout_por_caso.get(contexto["identificador_caso"], "organico")
                contexto["grafo"].enviar_atualizacao(contexto["montar_payload"](layout_atual))
                ui.notify("Reativada.", type="positive")
            ui.button('Reativar', color=None, on_click=_reativar).props('dense no-caps').classes(
                'text-xs bg-[#FF7300] text-white rounded-md px-3')
    dlg.open()


# ------------------------------------------------------------------ formulário (tipos com campos obrigatórios)
def abrir_formulario(ctx: "Pagina", contexto: Dict[str, Any], tipo: str, texto: str, pos: Dict[str, float]) -> None:
    """Endereço pede cidade e UF; clínica pede a razão social... Em vez de criar incompleto, pergunta."""
    t = tipos.obter(tipo)
    with ui.dialog() as dlg, ui.card().classes(CLASSE_DIALOGO + ' w-[420px] max-w-[92vw]'):
        ui.label(f'Nova entidade · {t.rotulo}').classes('text-sm font-bold text-gray-100')
        prim = next((c for c in t.campos if c.primario), None)
        ui.label(f"{prim.rotulo if prim else 'Valor'}: {t.formatar(t.normalizar(texto))}").classes('text-xs text-gray-300')
        entradas = {}
        for c in t.campos:
            if c.primario:
                continue
            entradas[c.nome] = ui.input(c.rotulo + (' *' if c.obrigatorio else '')).props('dense outlined').classes('w-full text-xs')
        faltando = ui.label('').classes('text-[11px] text-[#C9A66B]')
        with ui.row().classes('w-full justify-end gap-2 mt-1'):
            ui.button('Cancelar', color=None, on_click=dlg.close).props('flat dense no-caps').classes('text-xs text-gray-400')

            def _criar(essencial: bool):
                campos = {k: (e.value or "").strip() for k, e in entradas.items() if (e.value or "").strip()}
                falta = [c.rotulo for c in campos_obrigatorios_extras(tipo) if c.nome not in campos]
                if falta and not essencial:
                    faltando.set_text("Faltam: " + ", ".join(falta))
                    return
                dlg.close()
                _criar_e_mostrar(ctx, contexto, tipo, texto, campos, pos)
            ui.button('Criar só com o essencial', color=None, on_click=lambda: _criar(True)).props('flat dense no-caps').classes(
                'text-xs text-gray-300')
            ui.button('Criar', color=None, on_click=lambda: _criar(False)).props('dense no-caps').classes(
                'text-xs bg-[#FF7300] text-white rounded-md px-3')
    dlg.open()


# ------------------------------------------------------------------ colar
def classificar_linhas(linhas: List[str]) -> List[Dict[str, Any]]:
    """Uma entrada por linha: candidatos, tipo sugerido (None se ambígua ou desconhecida)."""
    out = []
    for texto in linhas[:MAX_LINHAS_COLAR]:
        cands = _candidatos_json(texto)
        amb = ambiguo(cands)
        out.append({"texto": texto, "candidatos": cands, "ambiguo": amb,
                    "tipo": None if (amb or not cands) else cands[0]["tipo"], "incluir": bool(cands)})
    return out


def posicoes_em_grade(n: int, cx: float, cy: float) -> List[Dict[str, float]]:
    cols = max(1, math.ceil(math.sqrt(n)))
    linhas = math.ceil(n / cols)
    return [{"x": cx + (i % cols - (cols - 1) / 2) * ESPACO_GRADE, "y": cy + (i // cols - (linhas - 1) / 2) * ESPACO_GRADE}
            for i in range(n)]


def tratar_colar(ctx: "Pagina", contexto: Dict[str, Any], valor: Dict[str, Any]) -> None:
    identificador = contexto["identificador_caso"]
    if identificador == overlay.DESCOBERTA:
        ui.notify("A mesa de descoberta ainda não é um sketch: promova a caso para criar entidades.", type="warning")
        return
    linhas = [str(l).strip() for l in (valor.get("linhas") or []) if str(l).strip()]
    if not linhas:
        return
    itens = classificar_linhas(linhas)
    cx, cy = float(valor.get("x") or 0), float(valor.get("y") or 0)
    opcoes_tipo = {t.chave: t.rotulo for t in tipos.REGISTRO.values()}

    with ui.dialog() as dlg, ui.card().classes(CLASSE_DIALOGO + ' w-[640px] max-w-[94vw]'):
        ui.label(f'Colar {len(itens)} linha(s) no sketch').classes('text-sm font-bold text-gray-100')
        if len(linhas) > MAX_LINHAS_COLAR:
            ui.label(f'Só as primeiras {MAX_LINHAS_COLAR} linhas.').classes('text-[11px] text-[#C9A66B]')
        ui.label('Confira o tipo de cada linha. As ambíguas (destacadas) precisam de escolha.').classes('text-[11px] text-[#71717A]')
        with ui.column().classes('w-full gap-1 max-h-[55vh] overflow-y-auto custom-scroll pr-1'):
            for it in itens:
                destaque = 'border-[#C9A66B]/60 bg-[#C9A66B]/5' if it["ambiguo"] else 'border-[#2B2B2F] bg-[#141416]/70'
                with ui.row().classes(f'w-full items-center gap-2 px-2 py-1 rounded-md border no-wrap {destaque}'):
                    cb = ui.checkbox(value=it["incluir"]).props('dense')
                    ui.label(it["texto"]).classes('text-xs text-gray-200 flex-1 truncate')
                    sugeridos = {c["tipo"]: f'{c["rotulo"]} · {c["confianca"]}' for c in it["candidatos"]}
                    sel = ui.select({**sugeridos, **{k: v for k, v in opcoes_tipo.items() if k not in sugeridos}},
                                    value=it["tipo"], label='Escolha o tipo' if it["ambiguo"] else None).props(
                        'dense outlined options-dense').classes('w-48 text-xs')
                    it["cb"], it["sel"] = cb, sel
        aviso = ui.label('').classes('text-[11px] text-[#C9A66B]')
        with ui.row().classes('w-full justify-end gap-2'):
            ui.button('Cancelar', color=None, on_click=dlg.close).props('flat dense no-caps').classes('text-xs text-gray-400')
            btn = ui.button('', color=None).props('dense no-caps').classes('text-xs bg-[#FF7300] text-white rounded-md px-3')

        def _escolhidos():
            return [it for it in itens if it["cb"].value]

        def _atualizar(*_):
            esc = _escolhidos()
            sem_tipo = [it for it in esc if not it["sel"].value]
            btn.set_text(f'Adicionar {len(esc)}')
            aviso.set_text(f'{len(sem_tipo)} linha(s) marcada(s) sem tipo escolhido.' if sem_tipo else '')
            btn.enable() if esc and not sem_tipo else btn.disable()
        for it in itens:
            it["cb"].on_value_change(_atualizar)
            it["sel"].on_value_change(_atualizar)
        _atualizar()

        def _adicionar():
            esc = _escolhidos()
            invalidos, validos, ja_oficiais = [], [], 0
            oficiais = contexto.get("ids_oficiais", set())
            for it in esc:
                try:
                    _, _, nid = sketch_repo.canonico(it["sel"].value, it["texto"])
                except ValueError:
                    invalidos.append(it["texto"])
                    continue
                if nid in oficiais:
                    ja_oficiais += 1                 # já é nó da base oficial: não vira linha duplicada
                else:
                    validos.append(it)
            if invalidos:
                aviso.set_text(f'Valor inválido para o tipo escolhido: {", ".join(invalidos[:3])}' + (' …' if len(invalidos) > 3 else ''))
                return
            grade = posicoes_em_grade(len(validos), cx, cy)
            res = [] if not validos else sketch_repo.criar_entidades(identificador, [
                {"tipo": it["sel"].value, "valor": it["texto"], "posicao": p} for it, p in zip(validos, grade)])
            dlg.close()
            ov = overlay.carregar(ctx.st, identificador, forcar=True)
            novos = [ov["nodes"][r["node_id"]] for r in res if r["estado"] != "existente" and r["node_id"] in ov["nodes"]]
            contexto["grafo"].enviar_patch({"add": {"nodes": novos, "edges": []}})
            ja = len(res) - len(novos) + ja_oficiais
            ui.notify(f"{len(novos)} entidade(s) adicionada(s)" + (f" · {ja} já estava(m) no sketch" if ja else "") + ".",
                      type="positive")
            ctx.sidebar_container.refresh()
        btn.on_click(_adicionar)
    dlg.open()
