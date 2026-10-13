"""
Vínculo manual com rascunho (item 3).

O rascunho vive só no navegador (Alt+arrastar de um nó a outro, ou "Criar vínculo…" com dois
nós selecionados); Esc o descarta sem gravar nada. Ao confirmar, o iframe manda edge_create e
aqui se grava (sketch_repo) e se responde com um patch que troca o rascunho pela aresta real.

Vínculo do cluster (base oficial) é imutável: o menu dele oferece "Anotar vínculo", que cria
uma linha origem='anotacao' com ref_oficial apontando para a aresta — o rótulo do analista
aparece ao lado e a aresta oficial segue intacta.

Edição (edge_update) é otimista no iframe: se a gravação falhar, o patch devolve os valores
antigos e um aviso aparece.
"""

from __future__ import annotations

from typing import Any, Dict

from nicegui import ui

from lcfo_dominio import sketch_repo
from lcfo_ui import overlay

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina

CAMPOS = ("rotulo", "direcao", "situacao", "data_ref", "confianca", "observacao")


def _limpo(valor: Dict[str, Any]) -> Dict[str, Any]:
    """Campos do menu do vínculo, normalizados (vazio -> None; confiança 1..5)."""
    out: Dict[str, Any] = {}
    for c in CAMPOS:
        if c not in valor:
            continue
        v = valor.get(c)
        v = v.strip() if isinstance(v, str) else v
        if c == "confianca":
            try:
                v = int(v) if v not in (None, "") else None
            except (TypeError, ValueError):
                v = None
            v = v if v is None or 1 <= v <= 5 else None
        elif c in ("data_ref", "observacao", "rotulo"):
            v = v or (None if c == "data_ref" else "")
        out[c] = v
    return out


def _persistivel(ctx: "Pagina", contexto: Dict[str, Any]) -> bool:
    if contexto["identificador_caso"] == overlay.DESCOBERTA:
        ui.notify("A mesa de descoberta ainda não é um sketch: promova a caso para criar vínculos.", type="warning")
        return False
    return True


def tratar_criar(ctx: "Pagina", contexto: Dict[str, Any], valor: Dict[str, Any]) -> None:
    identificador = contexto["identificador_caso"]
    tmp_id = valor.get("tmp_id") or ""
    de, para = valor.get("de") or "", valor.get("para") or ""
    remover_rascunho = {"remove": {"nodes": [], "edges": [tmp_id] if tmp_id else []}}
    if not _persistivel(ctx, contexto) or not de or not para or de == para:
        contexto["grafo"].enviar_patch(remover_rascunho)
        return
    campos = _limpo(valor)
    ref_oficial = valor.get("ref_oficial") or None
    try:
        vid = sketch_repo.criar_vinculo(
            identificador, de, para, rotulo=campos.get("rotulo", ""), direcao=campos.get("direcao") or "de_para",
            situacao=campos.get("situacao") or "confirmado", data_ref=campos.get("data_ref"),
            confianca=campos.get("confianca"), observacao=campos.get("observacao", ""),
            origem="anotacao" if ref_oficial else "manual", ref_oficial=ref_oficial)
    except Exception as e:
        contexto["grafo"].enviar_patch(remover_rascunho)
        ui.notify(f"Vínculo não gravado: {e}", type="negative")
        return
    overlay.carregar(ctx.st, identificador, forcar=True)
    aresta = sketch_repo.carregar_vinculo(identificador, vid)
    contexto["grafo"].enviar_patch({"remove": remover_rascunho["remove"], "add": {"nodes": [], "edges": [aresta]}})
    ui.notify("Anotação gravada sobre o vínculo da base." if ref_oficial else
              ("Vínculo gravado como hipótese." if aresta["situacao"] == "hipotese" else "Vínculo gravado."), type="positive")


def tratar_atualizar(ctx: "Pagina", contexto: Dict[str, Any], valor: Dict[str, Any]) -> None:
    identificador = contexto["identificador_caso"]
    vid = valor.get("id") or ""
    antes = sketch_repo.carregar_vinculo(identificador, vid) if identificador != overlay.DESCOBERTA else None
    if antes is None:
        ui.notify("Vínculo derivado da base: não é editável. Use 'Anotar vínculo'.", type="warning")
        return
    try:
        sketch_repo.atualizar_vinculo(identificador, vid, **_limpo(valor))
        final = sketch_repo.carregar_vinculo(identificador, vid)
    except Exception as e:
        final = antes                                         # reverte a edição otimista do iframe
        ui.notify(f"Edição não gravada (revertida): {e}", type="negative")
    overlay.carregar(ctx.st, identificador, forcar=True)
    contexto["grafo"].enviar_patch({"update": {"nodes": [], "edges": [final]}})
