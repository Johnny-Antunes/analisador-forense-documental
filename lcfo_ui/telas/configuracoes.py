"""
Tela de configurações (item 13): tipos de entidade de fábrica (só leitura) e tipos criados pelo
analista (criar, editar, desativar, excluir). As regras ficam em lcfo_dominio/tipos_custom.py.
"""

from __future__ import annotations

import re
import unicodedata
from typing import Any, Dict, List, Optional

from nicegui import ui

from lcfo_dominio import sketch_repo, tipos, tipos_custom
from lcfo_dominio.icones import PICTOGRAMAS, svg_no

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina

CARD = 'bg-[#141416] border border-[#2B2B2F] rounded-lg p-3 gap-1'


def _sugerir_chave(rotulo: str) -> str:
    s = "".join(c for c in unicodedata.normalize("NFKD", rotulo) if not unicodedata.combining(c)).lower()
    s = re.sub(r"[^a-z0-9]+", "_", s).strip("_")
    return (s if s and s[0].isalpha() else f"t_{s}")[:30]


def render(ctx: "Pagina") -> None:
    with ui.column().classes('w-full h-full overflow-y-auto custom-scroll px-8 py-6 gap-4 max-w-6xl'):
        ui.label('Configurações').classes('text-lg font-bold text-gray-100')
        ui.label('Tipos de entidade: os de fábrica e os criados aqui (quick-add, colar, filtros, legenda e fusão passam a usá-los '
                 'sem mudança de código).').classes('text-xs text-[#71717A]')

        @ui.refreshable
        def lista():
            ui.label('Tipos criados pelo analista').classes('text-sm font-bold text-gray-200 mt-2')
            customs = tipos_custom.listar()
            if not customs:
                ui.label('Nenhum ainda.').classes('text-xs text-[#71717A]')
            with ui.grid(columns=3).classes('w-full gap-3'):
                for d in customs:
                    n = tipos_custom.em_uso(d["chave"])
                    with ui.column().classes(CARD + ('' if d["ativo"] else ' opacity-50')):
                        with ui.row().classes('items-center gap-2 no-wrap'):
                            ui.image(svg_no(d.get("icone") or "documento", d["cor_fundo"], d["cor_borda"])).classes('w-7 h-7')
                            ui.label(d["rotulo"]).classes('text-sm font-semibold text-gray-100')
                        ui.label(f"{d['prefixo']}_… · chave {d['chave']} · {n} entidade(s) em sketches"
                                 + ("" if d["ativo"] else " · desativado")).classes('text-[11px] text-[#71717A]')
                        with ui.row().classes('gap-1'):
                            ui.button('Editar', color=None, on_click=lambda d=d: abrir_editor(ctx, d, lista.refresh)).props(
                                'flat dense no-caps size=sm').classes('text-[11px] text-[#FF7300]')
                            ui.button('Ativar' if not d["ativo"] else 'Desativar', color=None,
                                      on_click=lambda d=d: (tipos_custom.definir_ativo(d["chave"], not d["ativo"]), lista.refresh())).props(
                                'flat dense no-caps size=sm').classes('text-[11px] text-gray-300')
                            if not n:
                                def _excluir(d=d):
                                    try:
                                        tipos_custom.excluir(d["chave"])
                                    except ValueError as e:
                                        ui.notify(str(e), type="warning")
                                    lista.refresh()
                                ui.button('Excluir', color=None, on_click=_excluir).props('flat dense no-caps size=sm').classes('text-[11px] text-[#C0625F]')
            ui.button('Novo tipo', icon='add', color=None, on_click=lambda: abrir_editor(ctx, None, lista.refresh)).props(
                'dense no-caps').classes('text-xs bg-[#FF7300] text-white rounded-md px-3 mt-1')

            ui.label('Tipos de fábrica').classes('text-sm font-bold text-gray-200 mt-4')
            with ui.grid(columns=3).classes('w-full gap-3'):
                for k in sorted(tipos_custom.FABRICA):
                    t = tipos.REGISTRO[k]
                    with ui.column().classes(CARD):
                        with ui.row().classes('items-center gap-2 no-wrap'):
                            ui.image(t.icone_svg()).classes('w-7 h-7')
                            ui.label(t.rotulo).classes('text-sm font-semibold text-gray-100')
                        ui.label(f"{t.prefixo}_… · " + ", ".join(c.rotulo for c in t.campos)).classes('text-[11px] text-[#71717A]')
        lista()


def abrir_editor(ctx: "Pagina", d: Optional[Dict[str, Any]], depois) -> None:
    novo = d is None
    d = dict(d or {"cor_fundo": "#5B6472", "cor_borda": "#8B95A5", "icone": "documento", "normalizacao": "texto", "confianca": 60,
                   "campos": [{"nome": "identificacao", "rotulo": "Identificação", "obrigatorio": True, "primario": True}]})
    usado = (not novo) and tipos_custom.em_uso(d["chave"]) > 0
    campos: List[Dict[str, Any]] = [dict(c) for c in d.get("campos") or []]
    with ui.dialog() as dlg, ui.card().classes('bg-[#18181A] border border-[#2B2B2F] rounded-xl p-4 w-[640px] max-w-[94vw] gap-2'):
        ui.label('Novo tipo de entidade' if novo else f"Editar {d['rotulo']}").classes('text-sm font-bold text-gray-100')
        if usado:
            ui.label('Este tipo já tem entidades em sketches: chave, prefixo e normalização ficam travados (mudariam os ids).').classes(
                'text-[11px] text-[#C9A66B]')
        with ui.row().classes('w-full gap-2 no-wrap'):
            rot = ui.input('Rótulo', value=d.get("rotulo", "")).props('dense outlined').classes('flex-1 text-xs')
            chave = ui.input('Chave', value=d.get("chave", "")).props('dense outlined').classes('w-40 text-xs')
            pref = ui.input('Prefixo do id', value=d.get("prefixo", "")).props('dense outlined').classes('w-32 text-xs')
        if usado:
            chave.disable(); pref.disable()
        elif novo:
            def _auto(e):
                k = _sugerir_chave(e.value or "")
                chave.value = k
                pref.value = re.sub(r"[^A-Z0-9]", "", k.upper())[:4] or "TIP"
            rot.on_value_change(_auto)
        with ui.row().classes('w-full gap-2 items-center no-wrap'):
            cor_f = ui.color_input('Fundo', value=d["cor_fundo"]).props('dense outlined').classes('w-36 text-xs')
            cor_b = ui.color_input('Borda', value=d["cor_borda"]).props('dense outlined').classes('w-36 text-xs')
            icone = ui.select({k: k for k in PICTOGRAMAS}, value=d.get("icone") or "documento", label='Ícone').props('dense outlined').classes('w-36 text-xs')
            previa = ui.image(svg_no(icone.value, cor_f.value, cor_b.value)).classes('w-9 h-9')

            def _previa(_=None):
                try:
                    previa.set_source(svg_no(icone.value, cor_f.value or "#5B6472", cor_b.value or "#8B95A5"))
                except Exception:
                    pass
            for w in (cor_f, cor_b, icone):
                w.on_value_change(_previa)
        with ui.row().classes('w-full gap-2 no-wrap'):
            norm = ui.select(tipos_custom.NORMALIZACOES, value=d.get("normalizacao") or "texto", label='Valor canônico (identidade)').props(
                'dense outlined').classes('flex-1 text-xs')
            if usado:
                norm.disable()
            rel = ui.input('Vínculos sugeridos (vírgulas)', value=d.get("relacoes") or "").props('dense outlined').classes('flex-1 text-xs')
        with ui.row().classes('w-full gap-2 no-wrap'):
            rv = ui.input('Validação (expressão regular, opcional)', value=d.get("regex_validacao") or "").props('dense outlined').classes('flex-1 text-xs')
            rd = ui.input('Detecção no quick-add (regex, opcional)', value=d.get("regex_deteccao") or "").props('dense outlined').classes('flex-1 text-xs')
            conf = ui.number('Confiança', value=d.get("confianca") or 60, min=50, max=100).props('dense outlined').classes('w-24 text-xs')

        ui.label('Campos').classes('text-[10px] font-semibold text-[#71717A] uppercase tracking-wider mt-1')
        caixa = ui.column().classes('w-full gap-1')

        def _desenhar_campos():
            caixa.clear()
            with caixa:
                for i, c in enumerate(campos):
                    with ui.row().classes('w-full items-center gap-2 no-wrap'):
                        ui.input('Nome', value=c.get("nome", ""), on_change=lambda e, c=c: c.__setitem__("nome", _sugerir_chave(e.value or "") or "campo")).props(
                            'dense outlined').classes('w-36 text-xs')
                        ui.input('Rótulo', value=c.get("rotulo", ""), on_change=lambda e, c=c: c.__setitem__("rotulo", e.value)).props(
                            'dense outlined').classes('flex-1 text-xs')
                        ui.checkbox('obrigatório', value=bool(c.get("obrigatorio")), on_change=lambda e, c=c: c.__setitem__("obrigatorio", e.value)).props(
                            'dense').classes('text-[11px] text-gray-300')

                        def _primario(i=i):
                            for j, cc in enumerate(campos):
                                cc["primario"] = (j == i)
                            _desenhar_campos()
                        ui.button('identidade' if c.get("primario") else 'tornar identidade', color=None, on_click=_primario).props(
                            'flat dense no-caps size=sm').classes('text-[11px] ' + ('text-[#FF7300]' if c.get("primario") else 'text-gray-400'))
                        if not c.get("primario"):
                            ui.button(icon='close', color=None, on_click=lambda i=i: (campos.pop(i), _desenhar_campos())).props(
                                'flat dense round size=xs').classes('text-gray-500')
                ui.button('+ campo', color=None, on_click=lambda: (campos.append({"nome": f"campo_{len(campos)}", "rotulo": "", "obrigatorio": False}),
                                                                    _desenhar_campos())).props('flat dense no-caps size=sm').classes('text-[11px] text-[#FF7300]')
        _desenhar_campos()
        erro = ui.label('').classes('text-[11px] text-[#C0625F]')
        with ui.row().classes('w-full justify-end gap-2'):
            ui.button('Cancelar', color=None, on_click=dlg.close).props('flat dense no-caps').classes('text-xs text-gray-400')

            def _salvar():
                defin = {"chave": chave.value, "prefixo": (pref.value or "").upper(), "rotulo": rot.value, "cor_fundo": cor_f.value,
                         "cor_borda": cor_b.value, "icone": icone.value, "campos": campos, "normalizacao": norm.value,
                         "regex_validacao": rv.value, "regex_deteccao": rd.value, "confianca": int(conf.value or 60), "relacoes": rel.value}
                try:
                    tipos_custom.salvar(defin, sketch_repo.usuario_atual())
                except ValueError as e:
                    erro.set_text(str(e))
                    return
                dlg.close()
                ui.notify(f"Tipo '{rot.value}' salvo: já disponível no quick-add, nos filtros e na legenda.", type="positive")
                depois()
            ui.button('Salvar', color=None, on_click=_salvar).props('dense no-caps').classes('text-xs bg-[#FF7300] text-white rounded-md px-3')
    dlg.open()
