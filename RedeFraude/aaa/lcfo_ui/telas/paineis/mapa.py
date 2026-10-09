"""
Painel da mesa: Radar Territorial (Mapa) do caso.

Mesmo componente em Canvas do Centro de Comando, em modo "caso": colore os
municípios pelos acionamentos do caso e enquadra a câmera neles.

Layout (mesmo conceito do grafo): o mapa ocupa a área toda, ATRÁS do HUD e do
dock de painéis da mesa; a lista "Municípios do caso" é um painel de vidro
flutuante à direita (recolhível). Como o HUD e o painel cobrem parte do canvas,
o Python informa essas áreas ao componente (payload["layout"] e enviar_layout):
a câmera enquadra o caso na área livre e os controles do mapa (modos, zoom,
legenda) ficam fora do que está coberto.

Os pontos de cada assistência não são plotados: hoje a coordenada de cada uma é
o centro da cidade, então o município é a menor unidade fiel.
"""

from __future__ import annotations

from typing import Any, Dict

from nicegui import ui

from lcfo_ui.config import ALTURA_HUD_MESA, LARGURA_PAINEL_MAPA
from lcfo_ui.mapa import MapaTerritorialNiceGUI
from lcfo_ui.servicos import carregar_municipios_ibge, montar_payload_mapa_caso
from lcfo_ui.telas.contexto_mesa import ContextoMesa
from territorio_engine import resolver_codigo_ibge

MARGEM_PAINEL = 12          # px entre o painel flutuante e a borda
LARGURA_CHIP = 40           # botão que reabre o painel recolhido


def render(m: ContextoMesa, lado: str) -> None:
    ctx, df_dados = m.ctx, m.df_dados
    st = ctx.st
    payload = montar_payload_mapa_caso(df_dados)
    if not payload["municipios"]:
        with ui.column().classes('w-full h-full pt-2 px-8'):
            ui.label('Sem municípios identificáveis nas ocorrências deste caso.').classes('text-[#71717A]')
        return

    ibge = carregar_municipios_ibge()
    total = sum(d["v"][-1] for d in payload["municipios"].values()) or 1   # v = acumulado; o último é o total

    # Painel único: o mapa passa por trás do HUD e do dock (topo coberto). Tela dividida: o painel
    # já começa abaixo do cabeçalho de cada metade, nada cobre o topo.
    unico = lado == "unico"
    estado: Dict[str, Any] = {"uf": None, "mun": None, "mapa": None, "aberta": st.mapa_lista_aberta if unico else False}   # metade da tela: lista começa recolhida
    topo = m.topo_livre or (ALTURA_HUD_MESA if unico else 0)

    def inset_direito() -> int:
        return (LARGURA_PAINEL_MAPA + MARGEM_PAINEL) if estado["aberta"] else (LARGURA_CHIP + MARGEM_PAINEL)

    payload["layout"] = {"topo": topo, "base": 0, "esq": 0, "dir": inset_direito()}

    # Código IBGE de cada ocorrência, resolvido uma vez (antes era refeito, linha a linha, a cada clique).
    tem_bairro = {"cidade", "uf", "bairro"} <= set(df_dados.columns)
    cods_linhas = [resolver_codigo_ibge(c, u) for c, u in zip(df_dados["cidade"], df_dados["uf"])] if tem_bairro else []

    def ao_mudar_mapa(valor: Dict[str, Any]):
        if valor.get("tipo") != "mapa_estado":
            return
        estado["uf"], estado["mun"] = valor.get("uf"), valor.get("municipio")
        painel.refresh()

    def abrir(cod):
        estado["uf"], estado["mun"] = ibge[cod]["uf"], cod
        estado["mapa"].enviar_estado(estado["uf"], cod, 0)
        painel.refresh()

    def alternar_painel():
        estado["aberta"] = not estado["aberta"]
        st.mapa_lista_aberta = estado["aberta"]
        painel.refresh()
        estado["mapa"].enviar_layout(topo, 0, 0, inset_direito())

    def conteudo():
        for cod, d in sorted(payload["municipios"].items(), key=lambda x: -x[1]["v"][-1]):
            ativo = cod == estado["mun"]
            with ui.row().classes(
                f"w-full items-center gap-2 no-wrap px-2 py-1 rounded cursor-pointer hover:bg-[#222225] "
                f"{'bg-[#FF7300]/10 border border-[#FF7300]/40' if ativo else 'border border-transparent'}"
            ).on('click', lambda c=cod: abrir(c)):
                ui.label(ibge[cod]["uf"]).classes('text-[10px] font-bold text-[#71717A] w-6 shrink-0')
                ui.label(ibge[cod]["nome"]).classes('text-xs text-gray-200 truncate flex-1')
                ui.label(str(d["v"][-1])).classes('text-xs font-bold text-gray-300 shrink-0')
                ui.label(f"{d['v'][-1] / total * 100:.0f}%").classes('text-[10px] text-[#71717A] w-8 text-right shrink-0')
        desl = payload.get("deslocamentos") or []
        if desl:
            crit = payload.get("criterio_deslocamento", {})
            ui.label(f"Deslocamentos impossíveis ({len(desl)})").classes(
                'text-[10px] font-semibold text-[#C0625F] uppercase tracking-wider mt-3').tooltip(
                f"Mesma placa ou telefone em municípios a ≥ {crit.get('km_mesmo_dia', 300)} km no mesmo dia, "
                f"ou ≥ {crit.get('km_janela', 600)} km em até {crit.get('janela_dias', 1)} dia(s). A base só tem datas, sem hora.")
            for par in desl[:30]:
                de, para = ibge.get(par["de"], {}), ibge.get(par["para"], {})
                with ui.column().classes(
                    'w-full gap-0 px-2 py-1 rounded cursor-pointer hover:bg-[#222225] border border-transparent hover:border-[#C0625F]/40'
                ).on('click', lambda c=par["de"]: abrir(c)):
                    ui.label(f"{de.get('nome', par['de'])} ↔ {para.get('nome', par['para'])} · {par['km']:,} km".replace(",", ".")).classes(
                        'text-xs text-gray-200 truncate')
                    for o in par["ocorrencias"][:3]:
                        quando = "mesmo dia" if o["dias"] == 0 else f"{o['dias']} dia depois"
                        ui.label(f"{o['tipo']} {o['valor']} · {o['data_de'][8:10]}/{o['data_de'][5:7]} → "
                                 f"{o['data_para'][8:10]}/{o['data_para'][5:7]} ({quando})").classes('text-[10px] text-[#A1A1AA] truncate')
                    if par["n"] > 3:
                        ui.label(f"+ {par['n'] - 3} ocorrência(s)").classes('text-[10px] text-[#71717A]')
        if payload["auditoria"]:
            ui.label('Sem código IBGE').classes('text-[10px] font-semibold text-[#C9A66B] uppercase tracking-wider mt-2')
            for a in payload["auditoria"]:
                ui.label(f"{a['cidade']} / {a['uf']} — {a['total']}").classes('text-[11px] text-gray-400')

        if estado["mun"] and tem_bairro:
            do_mun = df_dados[[c == estado["mun"] for c in cods_linhas]]
            bairros = do_mun["bairro"].fillna("").astype(str).str.strip()
            top_b = bairros[bairros != ""].value_counts().head(10)
            ui.label(f"Bairros · {ibge[estado['mun']]['nome']}").classes(
                'text-[10px] font-semibold text-[#71717A] uppercase tracking-wider mt-3')
            if top_b.empty:
                ui.label('sem bairro informado').classes('text-[11px] text-[#52525B]')
            for b, n in top_b.items():
                with ui.row().classes('w-full justify-between no-wrap px-2'):
                    ui.label(b).classes('text-[11px] text-gray-300 truncate')
                    ui.label(str(n)).classes('text-[11px] text-gray-400 shrink-0')

    @ui.refreshable
    def painel():
        topo_css = f'top:{topo + MARGEM_PAINEL}px;'
        if not estado["aberta"]:
            ui.button(icon='list', color=None, on_click=alternar_painel).props('flat dense round').classes(
                'absolute right-3 z-30 lcfo-panel text-gray-300 hover:text-white'
            ).style(topo_css).tooltip('Municípios do caso')
            return
        with ui.column().classes(
            f'absolute right-3 bottom-3 z-30 w-[{LARGURA_PAINEL_MAPA}px] lcfo-panel p-3 gap-1 no-wrap min-h-0'
        ).style(topo_css):
            with ui.row().classes('w-full items-center justify-between no-wrap'):
                ui.label('Municípios do caso').classes('text-[10px] font-semibold text-[#71717A] uppercase tracking-wider')
                ui.button(icon='chevron_right', color=None, on_click=alternar_painel).props('flat dense round size=sm').classes(
                    'text-[#71717A] hover:text-white').tooltip('Recolher')
            with ui.column().classes('w-full gap-1 flex-1 min-h-0 overflow-y-auto custom-scroll pr-1'):
                conteudo()

    # O contêiner é posicionado: o iframe do mapa (absolute inset-0) e o painel ficam relativos a ele.
    # Painel único: absolute inset-0 relativo à área da mesa = tela cheia atrás do HUD, como o grafo.
    with ui.element('div').classes('absolute inset-0'):
        estado["mapa"] = MapaTerritorialNiceGUI(ctx, f"mapa_caso_{m.identificador_caso}_{lado}", ao_mudar_mapa).render(payload)
        painel()
