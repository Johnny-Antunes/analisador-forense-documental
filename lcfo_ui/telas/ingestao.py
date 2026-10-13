"""
Tela de ingestão com prévia (item 10): Selecionar → Analisar (sem gravar) → Revisar → Confirmar,
mais o histórico das ingestões (rejeitos e Reverter). As regras ficam em lcfo_dominio/ingestao.py.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List

from nicegui import background_tasks, run, ui

from lcfo_dominio import ingestao, sketch_repo
from lcfo_ui.dados import DADOS
from lcfo_ui.servicos import CAMINHO_REDE_CRIACAO, CAMINHO_REDE_OFICIAL, PASTA_LOCAL_BLACKLIST, PASTA_LOCAL_CRIACAO

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina

CARD = 'w-full bg-[#141416] border border-[#2B2B2F] rounded-lg p-3 gap-2'
TITULO = 'text-[10px] font-semibold text-[#71717A] uppercase tracking-wider'
MAX_REJEITOS_TELA = 200


def _pasta_padrao(tipo: str) -> str:
    if tipo == "BLACKLIST":
        return str(CAMINHO_REDE_OFICIAL) if Path(CAMINHO_REDE_OFICIAL).exists() else str(PASTA_LOCAL_BLACKLIST)
    return str(CAMINHO_REDE_CRIACAO) if Path(CAMINHO_REDE_CRIACAO).exists() else str(PASTA_LOCAL_CRIACAO)


def _tam(n: int) -> str:
    return f"{n / 1e6:.1f} MB" if n >= 1e6 else f"{n / 1e3:.0f} KB"


def render(ctx: "Pagina") -> None:
    # Estado da tela (sobrevive aos refresh das etapas, não à troca de tela).
    est: Dict[str, Any] = {"tipo": "CRIACAO", "pasta": _pasta_padrao("CRIACAO"), "arquivos": [], "selecionados": set(),
                           "relatorios": [], "corrigidos": {}, "descartados": {}, "ignorar_duplicatas": False,
                           "progresso": None, "cancelar": False, "resultado": None}

    with ui.column().classes('w-full h-full overflow-y-auto custom-scroll px-8 py-6 gap-4 max-w-6xl'):
        with ui.row().classes('w-full items-center justify-between'):
            with ui.column().classes('gap-0'):
                ui.label('Ingestão com prévia').classes('text-lg font-bold text-gray-100')
                ui.label('Analise sem gravar, revise os rejeitos e só então grave — com backup e histórico reversível.').classes(
                    'text-xs text-[#71717A]')
        with ui.stepper().props('vertical flat dark active-color=orange done-color=green').classes('w-full bg-transparent') as passos:
            with ui.step('Selecionar', icon='folder_open'):
                _etapa_selecionar(ctx, est, passos)
            with ui.step('Analisar', icon='query_stats'):
                etapa_analisar = ui.column().classes('w-full gap-3')
                est["render_analisar"] = lambda: _etapa_analisar(ctx, est, passos, etapa_analisar)
            with ui.step('Revisar', icon='rule'):
                etapa_revisar = ui.column().classes('w-full gap-3')
                est["render_revisar"] = lambda: _etapa_revisar(ctx, est, passos, etapa_revisar)
            with ui.step('Confirmar', icon='save'):
                etapa_confirmar = ui.column().classes('w-full gap-3')
                est["render_confirmar"] = lambda: _etapa_confirmar(ctx, est, passos, etapa_confirmar)
        ui.separator().classes('bg-[#2B2B2F]')
        est["historico_refresh"] = _historico(ctx)


# ------------------------------------------------------------------ 1. selecionar
def _etapa_selecionar(ctx, est, passos) -> None:
    with ui.row().classes('w-full items-end gap-2'):
        tipo = ui.select({"CRIACAO": "Criações Diárias", "BLACKLIST": "Blacklist"}, value=est["tipo"], label='Base').props(
            'dense outlined').classes('w-48 text-xs')
        pasta = ui.input('Pasta (local ou de rede)', value=est["pasta"]).props('dense outlined').classes('flex-1 text-xs')
    lista = ui.column().classes('w-full gap-1')

    def _listar():
        est["tipo"], est["pasta"] = tipo.value, pasta.value
        est["arquivos"] = ingestao.listar_arquivos(pasta.value, tipo.value)
        est["selecionados"] = {a["caminho"] for a in est["arquivos"] if not a["ja_ingerido"]}
        lista.clear()
        with lista:
            if not est["arquivos"]:
                ui.label('Nenhuma planilha (.xlsx/.csv) nesta pasta.').classes('text-xs text-[#71717A]')
            for a in est["arquivos"]:
                with ui.row().classes('w-full items-center gap-2 px-2 py-1 rounded-md bg-[#141416]/70 no-wrap'):
                    ui.checkbox(value=a["caminho"] in est["selecionados"],
                                on_change=lambda e, c=a["caminho"]: est["selecionados"].add(c) if e.value else est["selecionados"].discard(c)).props('dense')
                    ui.label(a["nome"]).classes('text-xs text-gray-200 flex-1 truncate')
                    ui.label(_tam(a["tamanho"])).classes('text-[11px] text-[#71717A] w-16 text-right')
                    ui.label(a["modificado"].replace("T", " ")[:16]).classes('text-[11px] text-[#71717A] w-32')
                    ui.label('já ingerido' if a["ja_ingerido"] else 'novo').classes(
                        f'text-[10px] w-20 text-right {"text-[#71717A]" if a["ja_ingerido"] else "text-[#6FA98A]"}')
    tipo.on_value_change(lambda e: (setattr(pasta, 'value', _pasta_padrao(e.value)), _listar()))
    with ui.row().classes('gap-2'):
        ui.button('Listar arquivos', icon='refresh', color=None, on_click=_listar).props('flat dense no-caps').classes('text-xs text-gray-300')

        async def _analisar():
            sel = [a for a in est["arquivos"] if a["caminho"] in est["selecionados"]]
            if not sel:
                ui.notify("Selecione pelo menos um arquivo.", type="warning")
                return
            n = ui.notification('Analisando (nada é gravado)…', spinner=True, timeout=None)
            rels, erros = [], []
            for a in sel:
                try:
                    r = await run.io_bound(ingestao.analisar_arquivo, a["caminho"], est["tipo"])
                    if r is None:
                        return
                    rels.append(r)
                except Exception as e:
                    erros.append(f"{a['nome']}: {e}")
            n.dismiss()
            for e in erros:
                ui.notify(e, type="negative")
            est["relatorios"], est["corrigidos"], est["descartados"] = rels, {}, {}
            est["render_analisar"]()
            passos.next()
        ui.button('Analisar (sem gravar)', icon='query_stats', color=None, on_click=_analisar).props('dense no-caps').classes(
            'text-xs bg-[#FF7300] text-white rounded-md px-3')
    _listar()


# ------------------------------------------------------------------ 2. analisar
def _etapa_analisar(ctx, est, passos, alvo) -> None:
    alvo.clear()
    with alvo:
        for i, r in enumerate(est["relatorios"]):
            with ui.column().classes(CARD):
                with ui.row().classes('w-full items-center justify-between'):
                    ui.label(r["arquivo"]).classes('text-sm font-semibold text-gray-100')
                    ui.label(f"{r['n_linhas']} linhas · {_tam(r['tamanho'])}" + (" · layout: " + r["especificacao"] if r["especificacao"] else "")).classes(
                        'text-[11px] text-[#71717A]')
                ui.label('Mapeamento das colunas (sugerido; editável)').classes(TITULO)
                opcoes = {c: c for c in r["colunas"]}
                selects = {}
                with ui.grid(columns=3).classes('w-full gap-2'):
                    for campo, (rot, _, _) in ingestao.CAMPOS.items():
                        selects[campo] = ui.select({None: '— não mapear —', **opcoes}, value=r["mapeamento"].get(campo), label=rot).props(
                            'dense outlined options-dense').classes('text-xs')
                if r["nao_mapeadas"]:
                    ui.label("Não mapeadas (guardadas em dados extras): " + ", ".join(r["nao_mapeadas"][:12])).classes('text-[11px] text-[#71717A]')
                ui.label(f"Formato de data: {r['datas']['formato']}").classes('text-xs text-gray-300')

                async def _reanalisar(i=i, selects=selects):
                    mapa = {c: s.value for c, s in selects.items()}
                    novo = await run.io_bound(ingestao.analisar_arquivo, est["relatorios"][i]["caminho"], est["tipo"], mapa)
                    if novo is not None:
                        est["relatorios"][i] = novo
                        est["render_analisar"]()

                def _salvar_layout(i=i, selects=selects):
                    rr = est["relatorios"][i]
                    ingestao.salvar_especificacao(f"Layout {rr['arquivo']}", rr["colunas"], {c: s.value for c, s in selects.items()},
                                                  rr["datas"]["formato"])
                    ui.notify("Layout salvo: planilhas com as mesmas colunas usarão este mapeamento.", type="positive")
                with ui.row().classes('gap-2'):
                    ui.button('Reanalisar com este mapeamento', color=None, on_click=_reanalisar).props('flat dense no-caps').classes('text-xs text-gray-300')
                    ui.button('Salvar como layout', color=None, on_click=_salvar_layout).props('flat dense no-caps').classes('text-xs text-gray-300')
                if r["amostra"]:
                    ui.label('Amostra (20 linhas válidas, já normalizadas)').classes(TITULO)
                    cols = [{"name": k, "label": k, "field": k, "align": "left"} for k in r["amostra"][0].keys() if k != "arquivo"]
                    ui.table(columns=cols, rows=r["amostra"]).props('dense flat dark').classes('w-full text-[11px] bg-transparent')
        with ui.row().classes('gap-2'):
            ui.button('Voltar', color=None, on_click=passos.previous).props('flat dense no-caps').classes('text-xs text-gray-400')
            ui.button('Revisar', color=None, on_click=lambda: (est["render_revisar"](), passos.next())).props('dense no-caps').classes(
                'text-xs bg-[#FF7300] text-white rounded-md px-3')


# ------------------------------------------------------------------ 3. revisar
def _etapa_revisar(ctx, est, passos, alvo) -> None:
    alvo.clear()
    with alvo:
        for r in est["relatorios"]:
            corr = est["corrigidos"].setdefault(r["arquivo"], {})
            desc = est["descartados"].setdefault(r["arquivo"], set())
            with ui.column().classes(CARD):
                ui.label(r["arquivo"]).classes('text-sm font-semibold text-gray-100')
                with ui.row().classes('gap-6 flex-wrap'):
                    for rot, val, cor in (("válidas", r["n_validas"], "#6FA98A"), ("rejeitadas", r["n_rejeitadas"], "#C0625F"),
                                          ("datas ambíguas", r["datas"]["ambiguas"], "#C9A66B"), ("duplicatas na base", r["duplicatas"], "#8B95A5"),
                                          ("sem data (usa a do arquivo)", r["n_data_do_arquivo"], "#8B95A5")):
                        with ui.column().classes('gap-0'):
                            ui.label(f"{val:,}".replace(",", ".")).classes('text-lg font-bold').style(f'color:{cor}')
                            ui.label(rot).classes('text-[10px] text-[#71717A]')
                if r["datas"]["mudariam_de_mes"]:
                    ui.label(f"{r['datas']['mudariam_de_mes']} linha(s) têm dia e mês até 12: lidas ao contrário, cairiam em outro mês. "
                             "A leitura usada é dia/mês/ano.").classes('text-[11px] text-[#C9A66B]')
                if r["sem_ibge"]:
                    ui.label("Cidades sem código IBGE (não aparecerão no mapa): " + ", ".join(
                        f"{x['cidade']}/{x['uf']} ({x['n']})" for x in r["sem_ibge"][:8])).classes('text-[11px] text-[#C9A66B]')
                if r["rejeitos_por_motivo"]:
                    ui.label('Rejeitos por motivo').classes(TITULO)
                    ui.label(" · ".join(f"{m}: {n}" for m, n in r["rejeitos_por_motivo"].items())).classes('text-xs text-gray-300')
                    ui.label(f"Corrija o valor e reinclua, ou exclua (primeiras {MAX_REJEITOS_TELA}).").classes('text-[11px] text-[#71717A]')
                    with ui.column().classes('w-full gap-1 max-h-72 overflow-y-auto custom-scroll pr-1'):
                        for rej in r["rejeitos"][:MAX_REJEITOS_TELA]:
                            with ui.row().classes('w-full items-center gap-2 no-wrap'):
                                ui.label(f"linha {rej['linha']}").classes('text-[11px] text-[#71717A] w-16 shrink-0')
                                ui.label(rej["motivo"]).classes('text-[11px] text-gray-300 w-56 shrink-0 truncate')
                                campo = rej["campo"] or "cpf"
                                ent = ui.input(f"{ingestao.CAMPOS.get(campo, (campo,))[0]} corrigido", value=rej["valor"] or "").props(
                                    'dense outlined').classes('flex-1 text-xs')

                                def _reincluir(rej=rej, ent=ent, campo=campo, corr=corr, desc=desc):
                                    corr[rej["linha"]] = {campo: ent.value or ""}
                                    desc.discard(rej["linha"])
                                    ui.notify(f"Linha {rej['linha']} será reincluída com a correção.")

                                def _excluir(rej=rej, corr=corr, desc=desc):
                                    desc.add(rej["linha"])
                                    corr.pop(rej["linha"], None)
                                    ui.notify(f"Linha {rej['linha']} será descartada (fica registrada como descartada).")
                                ui.button('Reincluir', color=None, on_click=_reincluir).props('flat dense no-caps size=sm').classes('text-[11px] text-[#FF7300]')
                                ui.button('Excluir', color=None, on_click=_excluir).props('flat dense no-caps size=sm').classes('text-[11px] text-gray-400')
        ui.checkbox('Não inserir assistências que já existem na base', value=est["ignorar_duplicatas"],
                    on_change=lambda e: est.__setitem__("ignorar_duplicatas", e.value)).props('dense').classes('text-xs text-gray-300')
        with ui.row().classes('gap-2'):
            ui.button('Voltar', color=None, on_click=passos.previous).props('flat dense no-caps').classes('text-xs text-gray-400')
            ui.button('Confirmar', color=None, on_click=lambda: (est["render_confirmar"](), passos.next())).props('dense no-caps').classes(
                'text-xs bg-[#FF7300] text-white rounded-md px-3')


# ------------------------------------------------------------------ 4. confirmar
def _etapa_confirmar(ctx, est, passos, alvo) -> None:
    alvo.clear()
    with alvo:
        total = sum(r["n_validas"] for r in est["relatorios"]) + sum(len(c) for c in est["corrigidos"].values())
        ui.label(f"Vai gravar até {total:,} linha(s) de {len(est['relatorios'])} arquivo(s) em "
                 f"{'Criações Diárias' if est['tipo'] == 'CRIACAO' else 'Blacklist'}. Antes, um backup completo do banco vai para backups/."
                 .replace(",", ".")).classes('text-xs text-gray-300')
        barra = ui.linear_progress(value=0, show_value=False).props('color=orange rounded size=6px').classes('w-full')
        rotulo = ui.label('').classes('text-[11px] text-[#71717A]')
        resumo = ui.column().classes('w-full gap-1')
        botoes = ui.row().classes('gap-2')

        def _tick():
            p = est.get("progresso")
            if p:
                arq, feito, tot = p
                barra.value = (feito / tot) if tot else 1
                rotulo.set_text(f"{arq}: {feito:,} de {tot:,} linhas".replace(",", "."))
        ui.timer(0.3, _tick)

        async def _gravar():
            est["cancelar"] = False
            btn_gravar.disable()
            btn_cancelar.enable()
            decisoes = {"corrigidos": est["corrigidos"], "descartados": {k: sorted(v) for k, v in est["descartados"].items()},
                        "ignorar_duplicatas": est["ignorar_duplicatas"]}
            try:
                res = await run.io_bound(ingestao.aplicar, est["relatorios"], decisoes, usuario=sketch_repo.usuario_atual(),
                                         progresso=lambda a, f, t: est.__setitem__("progresso", (a, f, t)),
                                         cancelado=lambda: est["cancelar"])
            except Exception as e:
                ui.notify(f"Não gravado: {e}", type="negative")
                btn_gravar.enable()
                btn_cancelar.disable()
                return
            if res is None:
                return
            btn_cancelar.disable()
            DADOS.recarregar()
            resumo.clear()
            with resumo:
                ui.label(f"{'Concluída' if res['estado'] == 'concluida' else 'Cancelada entre lotes'}: "
                         f"{res.get('inseridas', 0):,} inserida(s), {res.get('duplicadas', 0):,} duplicada(s), "
                         f"{res.get('rejeitadas', 0):,} rejeitada(s).".replace(",", ".")).classes('text-sm text-gray-100')
                if res.get("backup"):
                    ui.label(f"Backup: {res['backup']}").classes('text-[11px] text-[#71717A] break-all')
                ui.label('O lote aparece no histórico abaixo, com os rejeitos e a opção de reverter.').classes('text-[11px] text-[#71717A]')
            if est.get("historico_refresh"):
                est["historico_refresh"]()

        with botoes:
            ui.button('Voltar', color=None, on_click=passos.previous).props('flat dense no-caps').classes('text-xs text-gray-400')
            btn_gravar = ui.button('Gravar', icon='save', color=None, on_click=_gravar).props('dense no-caps').classes(
                'text-xs bg-[#FF7300] text-white rounded-md px-3')
            btn_cancelar = ui.button('Cancelar gravação', color=None, on_click=lambda: est.__setitem__("cancelar", True)).props(
                'flat dense no-caps').classes('text-xs text-[#C0625F]')
            btn_cancelar.disable()


# ------------------------------------------------------------------ histórico
def _historico(ctx) -> None:
    @ui.refreshable
    def lista():
        ui.label('Histórico de ingestões').classes('text-sm font-bold text-gray-100')
        exs = ingestao.listar_execucoes(50)
        if not exs:
            ui.label('Nenhuma ingestão feita por esta tela ainda.').classes('text-xs text-[#71717A]')
        for ex in exs:
            with ui.row().classes('w-full items-center gap-3 px-3 py-2 rounded-md bg-[#141416] border border-[#2B2B2F] no-wrap'):
                with ui.column().classes('gap-0 flex-1 min-w-0'):
                    ui.label(f"{(ex['iniciado_em'] or '').replace('T', ' ')[:16]} · {'Criações' if ex['tipo_base'] == 'CRIACAO' else 'Blacklist'}"
                             f" · {ex['usuario'] or '—'}").classes('text-xs text-gray-200')
                    ui.label(f"{ex['n_lidas'] or 0} lidas · {ex['n_inseridas'] or 0} inseridas · {ex['n_duplicadas'] or 0} duplicadas · "
                             f"{ex['n_rejeitadas'] or 0} rejeitadas · {ex['estado']}").classes('text-[11px] text-[#71717A] truncate')
                ui.button('Rejeitos', color=None, on_click=lambda e=ex: _dialogo_rejeitos(e)).props('flat dense no-caps size=sm').classes(
                    'text-[11px] text-gray-300')
                if not str(ex["estado"]).startswith("revertida"):
                    ui.button('Reverter', color=None, on_click=lambda e=ex: _confirmar_reverter(e, lista.refresh)).props(
                        'flat dense no-caps size=sm').classes('text-[11px] text-[#C0625F]')
    lista()
    return lista.refresh


def _dialogo_rejeitos(ex) -> None:
    rej = ingestao.rejeitos(ex["id"])
    with ui.dialog() as dlg, ui.card().classes('bg-[#18181A] border border-[#2B2B2F] rounded-xl p-4 w-[720px] max-w-[94vw] gap-2'):
        ui.label(f"Rejeitos da ingestão {ex['id']}").classes('text-sm font-bold text-gray-100')
        if not rej:
            ui.label('Nenhum rejeito.').classes('text-xs text-[#71717A]')
        else:
            cols = [{"name": k, "label": k, "field": k, "align": "left"} for k in ("arquivo", "linha", "campo", "valor", "motivo", "estado")]
            ui.table(columns=cols, rows=rej, pagination=15).props('dense flat dark').classes('w-full text-[11px] bg-transparent')
        ui.button('Fechar', color=None, on_click=dlg.close).props('flat dense no-caps').classes('self-end text-xs text-gray-400')
    dlg.open()


def _confirmar_reverter(ex, depois) -> None:
    deps = ingestao.dependencias_de_casos(ex["id"])
    with ui.dialog() as dlg, ui.card().classes('bg-[#18181A] border border-[#2B2B2F] rounded-xl p-4 w-[440px] gap-2'):
        ui.label('Reverter ingestão').classes('text-sm font-bold text-gray-100')
        if deps:
            ui.label('Bloqueado: linhas deste lote já sustentam casos (promovidas/anexadas): ' + ", ".join(deps)).classes('text-xs text-[#C0625F]')
            ui.button('Fechar', color=None, on_click=dlg.close).props('flat dense no-caps').classes('self-end text-xs text-gray-400')
        else:
            ui.label(f"Apaga exatamente as {ex['n_inseridas'] or 0} linha(s) gravadas por este lote (com backup antes). "
                     "Linhas promovidas ou anexadas a casos nunca são tocadas.").classes('text-xs text-gray-300')
            with ui.row().classes('w-full justify-end gap-2'):
                ui.button('Cancelar', color=None, on_click=dlg.close).props('flat dense no-caps').classes('text-xs text-gray-400')

                async def _reverter():
                    try:
                        r = await run.io_bound(ingestao.reverter, ex["id"], sketch_repo.usuario_atual())
                    except Exception as e:
                        ui.notify(str(e), type="negative")
                        return
                    if r is None:
                        return
                    dlg.close()
                    DADOS.recarregar()
                    ui.notify(f"Revertida: {r['removidas']} linha(s) removida(s).", type="positive")
                    depois()
                ui.button('Reverter', color=None, on_click=_reverter).props('dense no-caps').classes('text-xs bg-[#B94A3C] text-white rounded-md px-3')
    dlg.open()
