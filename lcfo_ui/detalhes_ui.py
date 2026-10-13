"""
Painel de detalhes editável (item 9): Propriedades, Notas, Aparência e Origem.

O drawer do iframe continua como resumo de leitura; o botão "Editar" pede este painel ao
Python (node_select). O iframe nunca grava: as edições voltam a ele como patch do nó.
Dado oficial nunca é sobrescrito — nota, aparência e flag de um nó do cluster vão numa
anotação por cima (sketch_repo.anotar).
"""

from __future__ import annotations

import time
from datetime import datetime
from typing import Any, Dict, Optional

from nicegui import ui

import database
from lcfo_dominio import pivos, sketch_repo, tipos
from lcfo_ui import overlay
from lcfo_ui.dados import DADOS

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina

SEGUNDOS_AUTOSALVAR = 0.8
NOMES_CORES = {"azul": "Azul", "vermelho": "Vermelho", "roxo": "Roxo", "verde": "Verde", "ambar": "Âmbar", "rosa": "Rosa",
               "ciano": "Ciano", "cinza": "Cinza"}
NOMES_FORMAS = {"circulo": "Círculo", "quadrado": "Quadrado", "losango": "Losango", "hexagono": "Hexágono", "triangulo": "Triângulo"}
NOMES_FLAGS = {"red": "Vermelha", "orange": "Laranja", "blue": "Azul", "green": "Verde", "yellow": "Amarela"}


def _grafos(ctx: "Pagina", identificador: str):
    return [cg["grafo"] for cg in ctx.st.contexto_grafo_ativo.values() if cg.get("identificador_caso") == identificador]


def _fmt_quando(iso: Optional[str]) -> str:
    try:
        return datetime.fromisoformat(iso).strftime("%d/%m/%Y às %H:%M")
    except Exception:
        return iso or ""


def enviar_estilo(ctx: "Pagina", identificador: str, node_id: str, hub_id: str = "") -> None:
    """Patch do nó com o estilo final (aparência, flag, indicador de nota) depois de uma gravação."""
    d = sketch_repo.detalhes(identificador, node_id)
    t = tipos.tipo_do_id(node_id)
    estilo = sketch_repo.estilo_no(t.chave if t else "empresa", d["aparencia"], is_hub=(node_id == hub_id), tem_nota=bool(d["nota"]))
    no = {"id": node_id, **estilo}
    if not d["oficial"]:
        ov = overlay.carregar(ctx.st, identificador, forcar=True)["nodes"].get(node_id)
        if ov:
            no["label"] = ov.get("label")
            no["nome_titular"] = ov.get("nome_titular", "")
    else:
        overlay.carregar(ctx.st, identificador, forcar=True)
    for g in _grafos(ctx, identificador):
        g.enviar_patch({"update": {"nodes": [no], "edges": []}})


def tratar_atualizacao(ctx: "Pagina", contexto: Dict[str, Any], valor: Dict[str, Any]) -> None:
    """entity_update vindo do iframe (ex.: flag pelo menu do nó)."""
    identificador = contexto["identificador_caso"]
    node_id = valor.get("node_id") or ""
    if identificador == overlay.DESCOBERTA:
        ui.notify("Na descoberta flags e notas não são salvas: promova a caso.", type="warning")
        return
    ap_nova = valor.get("aparencia")
    try:
        if ap_nova is not None:
            atual = sketch_repo.detalhes(identificador, node_id)["aparencia"]
            sketch_repo.anotar(identificador, node_id, aparencia={**atual, **ap_nova} if "flag" not in ap_nova or ap_nova["flag"]
                               else {k: v for k, v in atual.items() if k != "flag"})
        if "nota" in valor:
            sketch_repo.anotar(identificador, node_id, nota=valor["nota"])
    except Exception as e:
        ui.notify(f"Não foi possível gravar: {e}", type="negative")
    enviar_estilo(ctx, identificador, node_id, contexto.get("hub_id", ""))


def _dados_oficiais(node_id: str) -> Dict[str, Any]:
    """Derivados de um nó do cluster (só leitura): grau, ocorrências, primeira/última data, arquivos."""
    out: Dict[str, Any] = {}
    if DADOS.G is not None and node_id in DADOS.G:
        out["grau"] = DADOS.G.degree(node_id)
        out["nome"] = DADOS.G.nodes[node_id].get("nome_titular") or ""
    t = tipos.tipo_do_id(node_id)
    if t and t.coluna_banco in ("cpf", "telefone", "placa"):
        conn = database.get_db_connection()
        try:
            r = conn.execute(f"""SELECT COUNT(*), MIN(data), MAX(data), COUNT(DISTINCT arquivo) FROM assistencias
                                 WHERE {t.coluna_banco} = ?""", (tipos.valor_do_id(node_id),)).fetchone()
            arqs = [a[0] for a in conn.execute(f"SELECT DISTINCT arquivo FROM assistencias WHERE {t.coluna_banco} = ? LIMIT 5",
                                               (tipos.valor_do_id(node_id),))]
        finally:
            conn.close()
        out.update({"ocorrencias": r[0], "primeira": r[1], "ultima": r[2], "n_arquivos": r[3], "arquivos": arqs})
    return out


def abrir_painel(ctx: "Pagina", contexto: Dict[str, Any], node_id: str) -> None:
    identificador = contexto["identificador_caso"]
    if identificador == overlay.DESCOBERTA:
        ui.notify("A mesa de descoberta ainda não é um sketch: promova a caso para editar.", type="warning")
        return
    t = tipos.tipo_do_id(node_id)
    d = sketch_repo.detalhes(identificador, node_id)
    hub = contexto.get("hub_id", "")
    with ui.dialog().props('position=right full-height') as dlg, ui.card().classes(
            'bg-[#18181A] border-l border-[#2B2B2F] h-full w-[440px] max-w-[96vw] p-0 gap-0 rounded-none'):
        with ui.row().classes('w-full items-center justify-between px-4 pt-3 pb-2 no-wrap'):
            with ui.column().classes('gap-0'):
                ui.label(tipos.formatar_id(node_id)).classes('text-sm font-bold text-gray-100')
                ui.label((t.rotulo if t else "Entidade") + (" · base oficial (anotação por cima)" if d["oficial"] else "")).classes(
                    'text-[11px] text-[#71717A]')
            ui.button(icon='close', color=None, on_click=dlg.close).props('flat dense round size=sm').classes('text-gray-400')
        with ui.tabs().props('dense no-caps active-color=orange indicator-color=orange').classes('w-full text-gray-400 text-xs') as abas:
            a_prop, a_nota, a_ap, a_orig = ui.tab('Propriedades'), ui.tab('Notas'), ui.tab('Aparência'), ui.tab('Origem')
        with ui.tab_panels(abas, value=a_prop).classes('w-full flex-1 bg-transparent text-gray-200'):
            with ui.tab_panel(a_prop).classes('gap-2'):
                _aba_propriedades(ctx, identificador, node_id, t, d, hub)
            with ui.tab_panel(a_nota).classes('gap-2'):
                _aba_notas(ctx, identificador, node_id, d, hub)
            with ui.tab_panel(a_ap).classes('gap-3'):
                _aba_aparencia(ctx, identificador, node_id, d, hub)
            with ui.tab_panel(a_orig).classes('gap-2'):
                _aba_origem(node_id, d)
    dlg.open()


def _aba_propriedades(ctx, identificador, node_id, t, d, hub) -> None:
    prim = next((c for c in (t.campos if t else ()) if c.primario), None)
    with ui.row().classes('w-full items-center gap-2'):
        ui.label(prim.rotulo if prim else 'Identificador').classes('text-[11px] text-[#71717A] w-32')
        ui.label(tipos.formatar_id(node_id)).classes('text-xs text-gray-200').tooltip('A identidade não muda: para isso, funda ou recrie a entidade')
    if d["oficial"]:
        of = _dados_oficiais(node_id)
        for rot, val in (("Nome", of.get("nome")), ("Conexões", of.get("grau")), ("Ocorrências", of.get("ocorrencias")),
                         ("Primeira data", of.get("primeira")), ("Última data", of.get("ultima"))):
            if val not in (None, ""):
                with ui.row().classes('w-full items-center gap-2'):
                    ui.label(rot).classes('text-[11px] text-[#71717A] w-32')
                    ui.label(str(val)).classes('text-xs text-gray-300')
        ui.label('Dados da base oficial: só leitura. Nota, aparência e flag ficam numa anotação por cima.').classes(
            'text-[10px] text-[#71717A] mt-1')
        return
    entradas = {}
    for c in (t.campos if t else ()):
        if c.primario:
            continue
        entradas[c.nome] = ui.input(c.rotulo + (' *' if c.obrigatorio else ''), value=str(d["campos"].get(c.nome) or "")).props(
            'dense outlined').classes('w-full text-xs')
    status = ui.label('').classes('text-[10px] text-[#71717A]')

    def _salvar():
        campos = {k: (e.value or "").strip() for k, e in entradas.items()}
        try:
            sketch_repo.anotar(identificador, node_id, campos=campos)
        except Exception as ex:
            status.set_text(f"Erro: {ex}")
            return
        status.set_text("salvo")
        enviar_estilo(ctx, identificador, node_id, hub)
    ui.button('Salvar campos', color=None, on_click=_salvar).props('dense no-caps').classes(
        'self-end text-xs bg-[#FF7300] text-white rounded-md px-3')


def _aba_notas(ctx, identificador, node_id, d, hub) -> None:
    ed = ui.editor(value=d["nota"], placeholder='Observações sobre esta entidade…').props(
        'dense min-height=12rem toolbar-bg=grey-10').classes('w-full text-xs')
    status = ui.label('salvo' if d["nota"] else '').classes('text-[10px] text-[#71717A]')
    if d.get("nota_por"):
        ui.label(f"Última edição: {d['nota_por']['quem'] or '—'} em {_fmt_quando(d['nota_por']['quando'])}").classes('text-[10px] text-[#71717A]')
    estado = {"sujo": False, "t": 0.0, "salvo": d["nota"]}

    def _mudou(e):
        if e.value == estado["salvo"]:
            return
        estado["sujo"], estado["t"] = True, time.monotonic()
        status.set_text("editando…")

    def _talvez_salvar():
        # salva sozinho 800 ms depois da última tecla; em erro, o texto continua no editor
        if not estado["sujo"] or time.monotonic() - estado["t"] < SEGUNDOS_AUTOSALVAR:
            return
        estado["sujo"] = False
        texto = ed.value or ""
        try:
            sketch_repo.anotar(identificador, node_id, nota=texto)
        except Exception as ex:
            estado["sujo"] = True
            status.set_text(f"erro ao salvar — o texto foi preservado ({ex})")
            return
        estado["salvo"] = texto
        status.set_text("salvo")
        enviar_estilo(ctx, identificador, node_id, hub)
    ed.on_value_change(_mudou)
    ui.timer(0.25, _talvez_salvar)


def _aba_aparencia(ctx, identificador, node_id, d, hub) -> None:
    ap = dict(d["aparencia"])

    def _aplicar(**mud):
        for k, v in mud.items():
            if v is None:
                ap.pop(k, None)
            else:
                ap[k] = v
        try:
            sketch_repo.anotar(identificador, node_id, aparencia=ap)
        except Exception as ex:
            ui.notify(f"Não foi possível gravar: {ex}", type="negative")
            return
        enviar_estilo(ctx, identificador, node_id, hub)
        corpo.refresh()

    @ui.refreshable
    def corpo():
        ui.label('Cor').classes('text-[10px] font-semibold text-[#71717A] uppercase tracking-wider')
        with ui.row().classes('gap-1.5'):
            for k, (bg, borda) in sketch_repo.CORES_APARENCIA.items():
                ui.element('div').classes('w-6 h-6 rounded-full cursor-pointer').style(
                    f'background:{bg}; border:2px solid {"#FFFFFF" if ap.get("cor") == k else borda}').tooltip(NOMES_CORES[k]).on(
                    'click', lambda k=k: _aplicar(cor=None if ap.get("cor") == k else k))
        ui.label('Forma').classes('text-[10px] font-semibold text-[#71717A] uppercase tracking-wider mt-1')
        with ui.row().classes('gap-1 flex-wrap'):
            for k, rot in NOMES_FORMAS.items():
                ativo = ap.get("forma", "circulo") == k
                ui.button(rot, color=None, on_click=lambda k=k: _aplicar(forma=None if k == "circulo" else k)).props(
                    'flat dense no-caps size=sm').classes(f'text-[11px] rounded border {"border-[#FF7300] text-[#FF7300]" if ativo else "border-[#2B2B2F] text-gray-300"}')
        ui.label('Tamanho').classes('text-[10px] font-semibold text-[#71717A] uppercase tracking-wider mt-1')
        with ui.row().classes('gap-1'):
            for k in sketch_repo.TAMANHOS_APARENCIA:
                ativo = ap.get("tamanho", "M") == k
                ui.button(k, color=None, on_click=lambda k=k: _aplicar(tamanho=None if k == "M" else k)).props(
                    'flat dense no-caps size=sm').classes(f'text-[11px] rounded border w-9 {"border-[#FF7300] text-[#FF7300]" if ativo else "border-[#2B2B2F] text-gray-300"}')
        ui.label('Flag').classes('text-[10px] font-semibold text-[#71717A] uppercase tracking-wider mt-1')
        rotulos = sketch_repo.rotulos_flags(identificador)
        with ui.row().classes('gap-1.5 items-center'):
            for k, cor in sketch_repo.FLAGS.items():
                ui.element('div').classes('w-5 h-5 rounded-full cursor-pointer').style(
                    f'background:{cor}; border:2px solid {"#FFFFFF" if ap.get("flag") == k else "transparent"}').tooltip(
                    rotulos.get(k) or NOMES_FLAGS[k]).on('click', lambda k=k: _aplicar(flag=None if ap.get("flag") == k else k))
        with ui.expansion('O que cada flag significa neste sketch').props('dense').classes('w-full text-[11px] text-gray-400'):
            campos = {}
            for k, cor in sketch_repo.FLAGS.items():
                with ui.row().classes('w-full items-center gap-2 no-wrap'):
                    ui.element('div').classes('w-3 h-3 rounded-full shrink-0').style(f'background:{cor}')
                    campos[k] = ui.input(placeholder=NOMES_FLAGS[k], value=rotulos.get(k, "")).props('dense outlined').classes('flex-1 text-xs')

            def _salvar_rotulos():
                sketch_repo.salvar_rotulos_flags(identificador, {k: c.value or "" for k, c in campos.items()})
                novos = sketch_repo.rotulos_flags(identificador)
                for g in _grafos(ctx, identificador):
                    g.enviar_mensagem({"type": "lcfo:rotulos_flags", "rotulos": novos})
                ui.notify("Legenda das flags salva.", type="positive")
            ui.button('Salvar legenda', color=None, on_click=_salvar_rotulos).props('flat dense no-caps size=sm').classes('text-[11px] text-[#FF7300]')
        if ap:
            ui.button('Voltar ao padrão do tipo', color=None, on_click=lambda: (ap.clear(), _aplicar())).props(
                'flat dense no-caps size=sm').classes('text-[11px] text-gray-400 mt-1')
    corpo()


def _retrolinks(node_id: str) -> None:
    from lcfo_dominio import mencoes
    cit = mencoes.citado_em(node_id)
    if not cit:
        return
    ui.label(f"Citado em {len(cit)} parecer(es)").classes('text-[10px] font-semibold text-[#71717A] uppercase tracking-wider mt-2')
    for p in cit[:10]:
        ui.label(f"• {p['data_registro']} — {p['titulo'] or 'sem título'} ({p['id_caso']})").classes('text-[11px] text-gray-300')


def _aba_origem(node_id: str, d: Dict[str, Any]) -> None:
    _retrolinks(node_id)
    if d["oficial"]:
        of = _dados_oficiais(node_id)
        ui.label(f"Origem: blacklist · {of.get('ocorrencias', 0)} assistência(s)"
                 + (f" em {of.get('n_arquivos')} arquivo(s)" if of.get("n_arquivos") else "")).classes('text-xs text-gray-200')
        for a in of.get("arquivos") or []:
            ui.label(f"• {a}").classes('text-[11px] text-gray-400')
        if d["origem"] == "anotacao_oficial":
            ui.label(f"Anotação criada por {d['criado_por'] or '—'} em {_fmt_quando(d['criado_em'])}").classes('text-[11px] text-[#71717A] mt-1')
        return
    if d["execucao"]:
        ex = d["execucao"]
        rot = pivos.CATALOGO[ex["pivo"]].rotulo if ex["pivo"] in pivos.CATALOGO else ex["pivo"]
        ui.label(f"Criado pelo pivô {rot} em {_fmt_quando(ex['iniciado_em'])} ({ex['usuario'] or '—'})").classes('text-xs text-gray-200')
        ui.label(f"Execução {d['id_execucao']}").classes('text-[11px] text-[#71717A]')
        ui.label(f"Parâmetros: {ex['parametros_json']}").classes('text-[10px] text-[#71717A] break-all')
    else:
        ui.label(f"Criado por {d['criado_por'] or '—'} em {_fmt_quando(d['criado_em'])}, {d['origem'] or 'manual'}").classes('text-xs text-gray-200')
