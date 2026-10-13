"""
Remoção lógica com motivo, "Desfazer" e a lixeira do sketch (item 4).

Três ações que não se confundem:
  - Isolar (nó oficial do cluster): oculta do sketch, mantém a evidência; "Reativar" em Ferramentas.
  - Remover do sketch (nó/vínculo do analista: manual, pivô, importação): vai para a lixeira, com
    quem, quando, motivo e lote; Restaurar devolve o lote inteiro.
  - Excluir definitivamente: não existe (cadeia de custódia).

O iframe só PEDE (remove_request / isolate_request); aqui se confirma, grava pelo sketch_repo e se
responde ao grafo com um patch (sem recarregar o payload inteiro).
"""

from __future__ import annotations

from datetime import datetime
from typing import Any, Dict, List

from nicegui import ui

from lcfo_dominio import fusao, sketch_repo
from lcfo_ui import overlay
from lcfo_ui.servicos import ocultar_no_do_caso

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina

MOTIVOS_SUGERIDOS = ["Erro de digitação", "Homônimo", "Fora do escopo", "Duplicado"]
SEGUNDOS_DESFAZER = 8
CLASSE_DIALOGO = 'bg-[#18181A] border border-[#2B2B2F] rounded-xl p-4 w-[420px] max-w-[92vw] gap-3'
CLASSE_PILULA = 'text-xs rounded-full px-2.5 py-0.5 border'


def _grafos_do_sketch(ctx: "Pagina", identificador: str) -> List[Any]:
    """Todos os grafos abertos deste sketch (a tela dividida pode mostrar o mesmo duas vezes)."""
    return [cg["grafo"] for cg in ctx.st.contexto_grafo_ativo.values() if cg.get("identificador_caso") == identificador]


def _plural(n: int, um: str, varios: str) -> str:
    return f"{n} {um if n == 1 else varios}"


def _seletor_motivo(container) -> Dict[str, Any]:
    """Pílulas com motivos sugeridos + detalhe livre. Devolve {'motivo': fn} que monta o texto final."""
    escolha = {"v": ""}
    with container:
        ui.label('Motivo').classes('text-[10px] font-semibold text-[#71717A] uppercase tracking-wider')
        with ui.row().classes('gap-1 flex-wrap'):
            botoes = {}

            def _marcar(m):
                escolha["v"] = "" if escolha["v"] == m else m
                for k, b in botoes.items():
                    ativo = k == escolha["v"]
                    b.classes(remove='bg-[#FF7300]/15 text-[#FF7300] border-[#FF7300]/50 text-gray-300 border-[#2B2B2F]')
                    b.classes(add='bg-[#FF7300]/15 text-[#FF7300] border-[#FF7300]/50' if ativo else 'text-gray-300 border-[#2B2B2F]')
                atualizar()

            for m in MOTIVOS_SUGERIDOS:
                botoes[m] = ui.button(m, color=None, on_click=lambda m=m: _marcar(m)).props('flat dense no-caps').classes(
                    f'{CLASSE_PILULA} text-gray-300 border-[#2B2B2F]')
        detalhe = ui.input(placeholder='Detalhe (opcional)').props('dense outlined').classes('w-full text-xs')

    def texto() -> str:
        partes = [p for p in (escolha["v"], (detalhe.value or "").strip()) if p]
        return " — ".join(partes)

    estado = {"texto": texto, "ao_mudar": None}

    def atualizar():
        if estado["ao_mudar"]:
            estado["ao_mudar"]()
    detalhe.on_value_change(lambda _: atualizar())
    return estado


# ------------------------------------------------------------------ remover
def abrir_dialogo_remocao(ctx: "Pagina", contexto: Dict[str, Any], node_ids: List[str], edge_ids: List[str]) -> None:
    identificador = contexto["identificador_caso"]
    if identificador == overlay.DESCOBERTA:
        ui.notify("A mesa de descoberta ainda não é um sketch: promova a caso para editar e ter lixeira.", type="warning")
        return
    prev = sketch_repo.previa_remocao(identificador, node_ids, edge_ids)
    n, v = len(prev["nos"]), len(prev["vinculos"])
    if not n and not v:
        if prev["oficiais_vinculos"] and not prev["oficiais_nos"]:
            ui.notify("Vínculo derivado da base: não pode ser removido. Use Isolar nó.", type="warning")
        else:
            ui.notify("Entidade da base oficial: não pode ser removida. Use Isolar nó.", type="warning")
        return

    with ui.dialog() as dlg, ui.card().classes(CLASSE_DIALOGO):
        ui.label('Remover do sketch').classes('text-sm font-bold text-gray-100')
        partes = [p for p in (_plural(n, "nó", "nós") if n else "", _plural(v, "vínculo", "vínculos") if v else "") if p]
        ui.label(f"{' e '.join(partes)} {'será removido' if n + v == 1 else 'serão removidos'}.").classes('text-xs text-gray-300')
        if prev["oficiais_nos"] or prev["oficiais_vinculos"]:
            k = len(prev["oficiais_nos"]) + len(prev["oficiais_vinculos"])
            ui.label(f"{_plural(k, 'item', 'itens')} da base oficial fica{'' if k == 1 else 'm'} (dado oficial só pode ser isolado).").classes(
                'text-[11px] text-[#C9A66B]')
        ui.label('Vai para a lixeira do sketch: nada é apagado e dá para restaurar.').classes('text-[11px] text-[#71717A]')
        area = ui.column().classes('w-full gap-1.5')
        motivo = _seletor_motivo(area)
        with ui.row().classes('w-full justify-end gap-2 mt-1'):
            ui.button('Cancelar', color=None, on_click=dlg.close).props('flat dense no-caps').classes('text-xs text-gray-400')
            btn = ui.button('Remover', color=None).props('dense no-caps').classes(
                'text-xs bg-[#B94A3C] hover:bg-[#A33F33] text-white rounded-md px-3')
        btn.disable()
        motivo["ao_mudar"] = lambda: btn.enable() if motivo["texto"]() else btn.disable()

        def _confirmar():
            r = sketch_repo.remover(identificador, prev["nos"], prev["vinculos"], motivo=motivo["texto"]())
            dlg.close()
            overlay.carregar(ctx.st, identificador, forcar=True)
            for g in _grafos_do_sketch(ctx, identificador):
                g.enviar_patch({"remove": {"nodes": r["ids_nos"], "edges": r["ids_vinculos"]}})
            for nid in r["ids_nos"]:
                for sel in ctx.st.selecao_entidades.values():
                    sel.discard(nid)
            if ctx.st.entidade_foco in r["ids_nos"]:
                ctx.st.entidade_foco = None
            ctx.sidebar_container.refresh()
            mostrar_desfazer(ctx, identificador, r)
        btn.on_click(_confirmar)
    dlg.open()


def mostrar_desfazer(ctx: "Pagina", identificador: str, r: Dict[str, Any]) -> None:
    """Aviso "Removido · Desfazer" por alguns segundos (não bloqueia a tela)."""
    if not r.get("lote"):
        return
    partes = [p for p in (_plural(r["nos"], "nó", "nós") if r["nos"] else "",
                          _plural(r["vinculos"], "vínculo", "vínculos") if r["vinculos"] else "") if p]
    with ui.dialog().props('seamless position=bottom') as aviso, ui.card().classes(
            'bg-[#222225] border border-[#3A3A40] rounded-lg px-4 py-2 mb-6 shadow-2xl'):
        with ui.row().classes('items-center gap-4 no-wrap'):
            ui.icon('delete_outline', size='18px').classes('text-[#E0684F]')
            ui.label(f"Removido · {' e '.join(partes)}").classes('text-xs text-gray-200')

            def _desfazer():
                aviso.close()
                restaurar_lote(ctx, identificador, r["lote"])
            ui.button('Desfazer', color=None, on_click=_desfazer).props('flat dense no-caps').classes(
                'text-xs font-bold text-[#FF7300]')
    aviso.open()
    ui.timer(SEGUNDOS_DESFAZER, aviso.close, once=True)


def restaurar_lote(ctx: "Pagina", identificador: str, lote: str) -> Dict[str, Any]:
    rr = sketch_repo.restaurar(identificador, lote)
    ov = overlay.carregar(ctx.st, identificador, forcar=True)
    patch = {"add": {"nodes": [ov["nodes"][i] for i in rr["ids_nos"] if i in ov["nodes"]],
                     "edges": [ov["edges"][i] for i in rr["ids_vinculos"] if i in ov["edges"]]}}
    for g in _grafos_do_sketch(ctx, identificador):
        g.enviar_patch(patch)
    ctx.sidebar_container.refresh()
    if rr["nos"] or rr["vinculos"]:
        ui.notify(f"Restaurado: {_plural(rr['nos'], 'nó', 'nós')} e {_plural(rr['vinculos'], 'vínculo', 'vínculos')}.", type="positive")
    return rr


# ------------------------------------------------------------------ isolar (nó oficial)
def abrir_dialogo_isolar(ctx: "Pagina", contexto: Dict[str, Any], node_id: str) -> None:
    if not contexto.get("eh_oficial"):
        ui.notify("Isolar está disponível só em sketches salvos.", type="warning")
        return
    alvo = contexto.get("chave_layout") or contexto["identificador_caso"]
    with ui.dialog() as dlg, ui.card().classes(CLASSE_DIALOGO):
        ui.label('Isolar nó').classes('text-sm font-bold text-gray-100')
        ui.label(f"{node_id} sai deste sketch, mas a evidência continua na base. "
                 "Para trazê-lo de volta: Ferramentas › Reativar.").classes('text-xs text-gray-300')
        area = ui.column().classes('w-full gap-1.5')
        motivo = _seletor_motivo(area)
        with ui.row().classes('w-full justify-end gap-2 mt-1'):
            ui.button('Cancelar', color=None, on_click=dlg.close).props('flat dense no-caps').classes('text-xs text-gray-400')
            btn = ui.button('Isolar', color=None).props('dense no-caps').classes(
                'text-xs bg-[#3A3A40] hover:bg-[#4A4A50] text-white rounded-md px-3')
        btn.disable()
        motivo["ao_mudar"] = lambda: btn.enable() if motivo["texto"]() else btn.disable()

        def _confirmar():
            ocultar_no_do_caso(alvo, node_id, motivo["texto"](), sketch_repo.usuario_atual())
            dlg.close()
            contexto["grafo"].enviar_patch({"remove": {"nodes": [node_id], "edges": []}})
            ui.notify("Nó isolado deste sketch (Ferramentas › Reativar para desfazer).")
        btn.on_click(_confirmar)
    dlg.open()


# ------------------------------------------------------------------ aba Lixeira (sidebar)
def _fmt_quando(iso: str) -> str:
    try:
        return datetime.fromisoformat(iso).strftime("%d/%m %H:%M")
    except Exception:
        return iso or ""


def contar_lotes(identificador: str) -> int:
    if identificador == overlay.DESCOBERTA:
        return 0
    return len(sketch_repo.listar_lixeira(identificador, limite=100))


def render_aba_lixeira(ctx: "Pagina", identificador: str) -> None:
    if identificador == overlay.DESCOBERTA:
        ui.label('A mesa de descoberta não tem lixeira: promova a caso primeiro.').classes('text-xs text-[#71717A] px-1 py-2')
        return
    lotes = sketch_repo.listar_lixeira(identificador, limite=100)
    ui.label('Removidos deste sketch (100 mais recentes)').classes('text-[10px] font-semibold text-[#71717A] uppercase tracking-wider mb-1')
    with ui.column().classes('w-full gap-1.5 flex-1 min-h-0 overflow-y-auto custom-scroll -mr-3 pr-2 no-wrap'):
        if not lotes:
            ui.label('Nada na lixeira.').classes('text-xs text-[#71717A] px-1 py-2')
        for l in lotes:
            amostra = [a for a in (l.get("amostra") or "").split("|") if a]
            titulo = ", ".join(amostra[:2]) + (f" +{len(amostra) - 2}" if len(amostra) > 2 else "") if amostra else "Vínculos"
            with ui.column().classes('w-full gap-0.5 p-2 rounded-md bg-[#141416]/70 border border-[#2B2B2F]'):
                ui.label(titulo).classes('text-xs font-semibold text-gray-200 truncate w-full')
                ui.label(f"{_plural(l['n_nos'] or 0, 'nó', 'nós')} · {_plural(l['n_vinculos'] or 0, 'vínculo', 'vínculos')}").classes(
                    'text-[11px] text-gray-400')
                ui.label(f"{l.get('motivo') or 'sem motivo'} · {l.get('quem') or '—'} · {_fmt_quando(l.get('quando'))}").classes(
                    'text-[10px] text-[#71717A]')
                if fusao.eh_lote_de_fusao(l.get("motivo") or ""):
                    # lote de fusão: restaurar só as absorvidas deixaria os vínculos duplicados — desfaz a fusão inteira
                    from lcfo_ui import fusao_ui
                    ui.button('Desfazer fusão', icon='call_split', color=None,
                              on_click=lambda lote=l["lote"]: fusao_ui.desfazer(ctx, identificador, lote)).props(
                        'flat dense no-caps size=sm').classes('self-end text-[11px] text-[#FF7300]')
                else:
                    ui.button('Restaurar', icon='restore', color=None,
                              on_click=lambda lote=l["lote"]: restaurar_lote(ctx, identificador, lote)).props(
                        'flat dense no-caps size=sm').classes('self-end text-[11px] text-[#FF7300]')
