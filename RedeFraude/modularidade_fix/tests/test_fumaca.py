"""
Teste de fumaça da interface: simula um usuário (sem navegador) percorrendo
todas as telas e painéis, contra o banco real (banco_fraudes.db).

Objetivo: garantir que nenhuma refatoração ou evolução quebre uma tela já
existente. Falha se qualquer tela lançar exceção (o plugin do NiceGUI
reprova o teste se houver log de ERROR).

Rodar:  python -m pytest -q
"""

from __future__ import annotations

import asyncio

import pytest
from nicegui import ui
from nicegui.testing import User
from nicegui.testing.user_interaction import UserInteraction

from lcfo_ui.config import ICONE_TOOLBAR
from lcfo_ui.dados import DADOS
from lcfo_ui.telas.paineis import PAINEIS

pytestmark = pytest.mark.skipif(not DADOS.cluster_info, reason="banco sem dados ingeridos")


async def _clicar(user: User, elementos) -> None:
    """Clica e devolve o controle ao event loop: no NiceGUI 3 o refresh() dos
    refreshables roda na próxima iteração do loop, não dentro do clique."""
    assert elementos, "nenhum elemento encontrado para clicar"
    UserInteraction(user, set(elementos), None).click()
    for _ in range(5):
        await asyncio.sleep(0.01)


def _botoes_com_icone(user: User, icone: str):
    return [b for b in user.find(ui.button).elements if b.props.get("icon") == icone]


def _elementos_clicaveis_com_classe(user: User, tipo, classe: str):
    """Elementos do tipo exato (sem subclasses — ex.: ui.label também é ui.element) com a classe e um handler de click."""
    return sorted(
        (e for e in user.find(tipo).elements
         if type(e) is tipo and classe in e.classes and "click" in [l.type for l in e._event_listeners.values()]),
        key=lambda e: e.id,
    )


async def _abrir_primeira_mesa(user: User) -> None:
    await user.open("/")
    await user.should_see("Investigations")
    # Primeiro card da fila -> overview do dossiê
    await _clicar(user, _elementos_clicaveis_com_classe(user, ui.card, "cursor-pointer")[:1])
    await user.should_see("Sketches (")
    # Primeiro sketch -> mesa
    await _clicar(user, _elementos_clicaveis_com_classe(user, ui.element, "cursor-pointer")[:1])
    await user.should_see("Toggle notes (Ctrl+L)")


async def test_dashboard(user: User) -> None:
    await user.open("/")
    await user.should_see("Investigations")
    await user.should_see("Fila de Investigação Priorizada")


@pytest.mark.parametrize("icone, texto", [
    ("hub", "Base Mestra de Entidades Monitoradas"),
    ("warning", "Watchlist de Municípios de Alto Risco"),
    ("dashboard", "Centro de Comando Proativo"),
    ("folder_open", "Investigations"),
])
async def test_telas_do_rail(user: User, icone: str, texto: str) -> None:
    await user.open("/")
    await _clicar(user, _botoes_com_icone(user, icone)[:1])
    await user.should_see(texto)


async def test_overview_e_todos_os_paineis_da_mesa(user: User, monkeypatch: pytest.MonkeyPatch) -> None:
    renderizados = []
    for nome, render in list(PAINEIS.items()):
        monkeypatch.setitem(PAINEIS, nome, lambda m, lado, _n=nome, _r=render: (renderizados.append(_n), _r(m, lado)))

    await _abrir_primeira_mesa(user)
    for painel, icone in ICONE_TOOLBAR.items():
        await _clicar(user, _botoes_com_icone(user, icone)[-1:])
    assert set(renderizados) == set(PAINEIS), f"painéis não renderizados: {set(PAINEIS) - set(renderizados)}"
    # Cockpit (tela dividida) e painel de notas
    await _clicar(user, _botoes_com_icone(user, "vertical_split"))
    await user.should_see("Toggle notes (Ctrl+L)")


async def test_atalhos_de_painel_lateral(user: User) -> None:
    await _abrir_primeira_mesa(user)
    await _clicar(user, _botoes_com_icone(user, "menu_open"))   # esconde a sidebar
    await _clicar(user, _botoes_com_icone(user, "menu_open"))   # mostra de novo
    await user.should_see("Toggle notes (Ctrl+L)")
