"""
Busca global / paleta de comandos (Ctrl+J).

Procura em quatro fontes, agrupadas: comandos da ferramenta, investigações
(casos), entidades (CPF, telefone, placa — abre o caso que as contém) e
municípios (abre o Centro de Comando focado no município). Teclado no mesmo
padrão da busca do mapa: ↑/↓ movem a opção ativa (com volta), Enter executa,
o mouse também move a opção ativa, Esc fecha.

`buscar_global` é uma função pura (sem UI) — testável e reaproveitável.
"""

from __future__ import annotations

import re
import unicodedata
from typing import TYPE_CHECKING, Any, Dict, List

from nicegui import ui

from lcfo_ui.dados import DADOS
from lcfo_ui.servicos import carregar_municipios_ibge

if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina

MAX_POR_GRUPO = {"Comandos": 6, "Investigações": 8, "Entidades": 6, "Municípios": 8}
MIN_ENTIDADE = 4   # dígitos/caracteres mínimos para varrer as entidades do grafo

COMANDOS = [
    {"titulo": "Investigações", "detalhe": "Fila de casos priorizada", "icone": "folder_open", "acao": ("tela", "Casos")},
    {"titulo": "Centro de Comando", "detalhe": "Radar territorial e mapa", "icone": "dashboard", "acao": ("tela", "Centro de Comando")},
    {"titulo": "Base Mestra", "detalhe": "Entidades monitoradas", "icone": "hub", "acao": ("tela", "Base Mestra")},
    {"titulo": "Watchlist", "detalhe": "Municípios de alto risco", "icone": "warning", "acao": ("tela", "Watchlist")},
    {"titulo": "Gestão de bases & ingestão", "detalhe": "Ingerir planilhas, corrigir datas", "icone": "storage", "acao": ("ingestao", None)},
    {"titulo": "Alternar painel lateral", "detalhe": "Ctrl+B", "icone": "menu_open", "acao": ("sidebar", None)},
    {"titulo": "Alternar notas da mesa", "detalhe": "Ctrl+L", "icone": "notes", "acao": ("notas", None)},
]


def _norm(texto: str) -> str:
    sem = "".join(c for c in unicodedata.normalize("NFKD", str(texto)) if not unicodedata.combining(c))
    return " ".join(sem.upper().split())


def buscar_global(termo: str) -> List[Dict[str, Any]]:
    """Lista plana de resultados {grupo, titulo, detalhe, icone, acao}, na ordem de exibição."""
    t = _norm(termo)
    if not t:
        return [{**c, "grupo": "Comandos"} for c in COMANDOS]

    res: List[Dict[str, Any]] = []
    comandos = [c for c in COMANDOS if t in _norm(c["titulo"]) or t in _norm(c["detalhe"])]
    res += [{**c, "grupo": "Comandos"} for c in comandos[:MAX_POR_GRUPO["Comandos"]]]

    casos = []
    for c in DADOS.cluster_info:
        nome = DADOS.casos_cadastrados.get(c["id"], {}).get("nome") or ""
        alvo = _norm(f"{nome} {c['hub_label']} {c['id']}")
        if t in alvo:
            casos.append({"grupo": "Investigações", "titulo": nome or c["hub_label"], "icone": "folder",
                          "detalhe": f"{c['id']} · {c['tamanho']} entidades", "acao": ("caso", c["id"])})
            if len(casos) >= MAX_POR_GRUPO["Investigações"]:
                break
    res += casos

    limpo = re.sub(r"[^A-Z0-9]", "", t)
    if len(limpo) >= MIN_ENTIDADE and DADOS.G is not None:
        entidades = []
        for n in DADOS.G.nodes:
            if limpo in n.split("_", 1)[-1]:
                tipo, valor = n.split("_", 1)
                entidades.append({"grupo": "Entidades", "titulo": DADOS.G.nodes[n].get("valor") or valor,
                                  "detalhe": f"{tipo} · {DADOS.G.nodes[n].get('nome_titular') or ''}".rstrip(" ·"),
                                  "icone": {"CPF": "badge", "TEL": "call", "PLACA": "directions_car"}.get(tipo, "tag"),
                                  "acao": ("entidade", n)})
                if len(entidades) >= MAX_POR_GRUPO["Entidades"]:
                    break
        res += entidades

    if len(t) >= 3:
        muns = []
        for cod, m in carregar_municipios_ibge().items():
            pos = _norm(m["nome"]).find(t)
            if pos >= 0:
                muns.append((pos != 0, -(m["pop"] or 0), cod, m))
        muns.sort()
        res += [{"grupo": "Municípios", "titulo": m["nome"], "detalhe": f"{m['uf']} · {m['pop']:,} hab.".replace(",", "."),
                 "icone": "location_on", "acao": ("municipio", cod)} for _, _, cod, m in muns[:MAX_POR_GRUPO["Municípios"]]]
    return res


def criar_busca_global(ctx: "Pagina") -> None:
    """Monta o diálogo da paleta e publica ctx.dlg_palette."""
    st = ctx.st
    estado: Dict[str, Any] = {"itens": [], "ativo": 0}

    def executar(item: Dict[str, Any]) -> None:
        tipo, arg = item["acao"]
        dlg.close()
        if tipo == "tela":
            st.tela_ativa = arg
            ctx.voltar_ao_dashboard()
        elif tipo == "ingestao":
            ctx.dlg_ingestao.open()
        elif tipo == "sidebar":
            ctx.alternar_sidebar()
        elif tipo == "notas":
            ctx.alternar_notas()
        elif tipo == "caso":
            ctx.abrir_caso_overview(arg)
        elif tipo == "entidade":
            caso = next((c["id"] for c in DADOS.cluster_info if arg in c["nodes"]), None)
            if caso:
                ctx.abrir_caso_overview(caso)
            else:
                ui.notify("Entidade fora das células catalogadas (componente com menos de 3 entidades).")
        elif tipo == "municipio":
            st.tela_ativa = "Centro de Comando"
            st.radar_uf, st.radar_mun = carregar_municipios_ibge()[arg]["uf"], arg
            ctx.voltar_ao_dashboard()

    @ui.refreshable
    def lista():
        itens = estado["itens"]
        if not itens:
            ui.label(f"Nada encontrado para “{campo.value.strip()}”.").classes('text-xs text-[#71717A] px-3 py-3')
            return
        grupo_atual = None
        for i, it in enumerate(itens):
            if it["grupo"] != grupo_atual:
                grupo_atual = it["grupo"]
                ui.label(grupo_atual).classes('text-[10px] font-semibold text-[#71717A] tracking-wider px-2.5 pt-2 pb-1 uppercase')
            ativo = i == estado["ativo"]
            linha = ui.row().classes(
                'lcfo-opcao-paleta w-full items-center gap-2.5 px-2.5 py-1.5 rounded-md cursor-pointer no-wrap '
                + ('bg-[#FF7300]/15 shadow-[inset_2px_0_0_#FF7300]' if ativo else 'hover:bg-[#222225]')
            ).props(f'role=option aria-selected={str(ativo).lower()}')
            linha.on('click', lambda _, it=it: executar(it))
            linha.on('mouseenter', lambda _, i=i: _ativar(i))
            with linha:
                ui.icon(it["icone"], size='15px').classes('text-[#FF7300]' if ativo else 'text-[#71717A]')
                with ui.column().classes('gap-0 min-w-0 flex-1'):
                    ui.label(it["titulo"]).classes('text-xs text-gray-100 font-medium truncate')
                    ui.label(it["detalhe"]).classes('text-[10px] text-[#71717A] truncate')
                if ativo:
                    ui.label('Enter').classes('text-[9px] text-[#71717A] border border-[#3A3A40] rounded px-1 shrink-0')

    def _ativar(i: int) -> None:
        if i != estado["ativo"]:
            estado["ativo"] = i
            lista.refresh()

    def _mover(passo: int) -> None:
        n = len(estado["itens"])
        if n:
            estado["ativo"] = (estado["ativo"] + passo) % n
            lista.refresh()
            ui.run_javascript("document.querySelector('.lcfo-opcao-paleta[aria-selected=true]')?.scrollIntoView({block: 'nearest'})")

    def _filtrar(termo: str) -> None:
        estado["itens"], estado["ativo"] = buscar_global(termo or ""), 0
        lista.refresh()

    def _enter() -> None:
        if estado["itens"]:
            executar(estado["itens"][estado["ativo"]])

    with ui.dialog() as dlg, ui.card().classes('w-[560px] bg-[#18181A] border border-[#2B2B2F] p-0 rounded-xl shadow-2xl overflow-hidden gap-0'):
        with ui.row().classes('h-11 border-b border-[#2B2B2F] px-3 items-center gap-2 w-full no-wrap'):
            ui.icon('search', size='15px').classes('text-[#71717A]')
            campo = ui.input(placeholder='Buscar casos, CPF, telefone, placa, municípios ou comandos...').props(
                'dense borderless autofocus spellcheck=false').classes('flex-1 text-white text-xs bg-transparent')
        with ui.column().classes('w-full gap-0.5 p-1.5 max-h-80 overflow-y-auto custom-scroll').props('role=listbox'):
            lista()
        with ui.row().classes('w-full gap-4 px-3 py-1.5 border-t border-[#2B2B2F] text-[10px] text-[#71717A]'):
            ui.label('↑ ↓ navegar')
            ui.label('Enter abrir')
            ui.label('Esc fechar')

    campo.on_value_change(lambda e: _filtrar(e.value))
    campo.on('keydown.down.prevent', lambda: _mover(1))
    campo.on('keydown.up.prevent', lambda: _mover(-1))
    campo.on('keydown.enter.prevent', _enter)

    def _ao_abrir():
        campo.value = ""
        _filtrar("")

    dlg.on('show', _ao_abrir)
    ctx.dlg_palette = dlg
    _filtrar("")
