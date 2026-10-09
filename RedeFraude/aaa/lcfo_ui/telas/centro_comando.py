"""
Centro de Comando: cockpit territorial em tela única (sem rolagem de página).

  ┌ cabeçalho: mês · KPIs do mês · % mapeado · Mapa | Tabela ───────────────┐
  │ esquerda (280px)  │ centro: mapa em Canvas (ou tabela) │ direita (340px) │
  │ evolução mensal,  │ drill Brasil > UF > município,      │ ranking em      │
  │ novos / saíram,   │ linha do tempo mensal               │ destaque, ou o  │
  │ composição        │                                     │ funil do município│
  └──────────────────────────────────────────────────────────────────────────┘

Os dados (radar mês a mês) são calculados fora do event loop (run.io_bound),
assim como o funil de cada município. O Python guarda o estado confirmado
(mês, UF, município); o iframe do mapa responde na hora e avisa o Python.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Dict

import pandas as pd
from nicegui import run, ui

from lcfo_ui.mapa import MapaTerritorialNiceGUI
from lcfo_ui.servicos import (
    carregar_municipios_ibge, extrair_top_infratores_municipio, marcar_entidades_conhecidas,
    montar_dados_mapa_radar, resumo_do_mes,
)
from lcfo_ui.util_ui import sanitizar_df_para_tabela

if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina

MESES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez']
ALERTA_CURTO = {1: ("Explosão", "#C0625F"), 2: ("Pet", "#C9A66B"), 4: ("Residencial", "#6C93B0"),
                8: ("Watchlist", "#8F86B5"), 16: ("Sem histórico", "#FF7300")}
FILTROS = [(0, "Todos"), (1, "Explosão"), (2, "Pet"), (4, "Resid."), (8, "Watchlist"), (16, "Sem hist.")]
ORDENS = [("criticos", "Críticos"), ("volume", "Volume"), ("az", "A-Z")]
CLASSE_PILULA = "px-2.5 h-6 rounded-full text-[11px] border"
PILULA_ATIVA = "bg-[#FF7300]/15 text-[#FF7300] border-[#FF7300]/40"
PILULA_INATIVA = "bg-[#18181A] text-[#71717A] border-[#2B2B2F] hover:text-white"

# Colunas laterais como painéis de vidro flutuantes sobre o mapa (mesma linguagem do mapa do
# caso e do drawer do grafo): cantos arredondados, blur, recolhíveis. O mapa recebe a área
# coberta (enviar_layout) e mantém câmera e controles no vão livre entre eles.
LARGURA_ESQ, LARGURA_DIR, MARGEM, LARGURA_RECOLHIDO = 280, 340, 12, 44
VIDRO = 'bg-[#18181A]/90 backdrop-blur-md border border-[#2B2B2F] rounded-xl shadow-2xl'


def _fmt_mes(m: str) -> str:
    return f"{MESES[int(m[5:7]) - 1]}/{m[:4]}" if m and len(m) >= 7 else str(m)


def _fmt_int(n) -> str:
    return "—" if n is None else f"{int(round(n)):,}".replace(",", ".")


def _fmt_razao(r, media=1) -> str:
    """Desvio sobre a média. Sem média histórica (município novo) a razão do radar
    divide por 0,4 e vira um número sem sentido (ex.: 100×) — mostra "novo"."""
    if r is None:
        return "—"
    if not media:
        return "novo"
    return f"{r:.1f}×".replace(".", ",")


def _selos(bits: int) -> str:
    return "".join(
        f'<span style="font-size:9px; font-weight:600; padding:1px 6px; border-radius:999px; margin-right:3px;'
        f'background:{cor}22; color:{cor}; border:1px solid {cor}55;">{rot}</span>'
        for bit, (rot, cor) in ALERTA_CURTO.items() if bits & bit
    )


def _gravidade(linha: Dict[str, Any]) -> tuple:
    bits = linha["bits"]
    pontos = bin(bits).count("1") + (1 if bits & 1 else 0)
    return (pontos, linha["razao"] or 0, linha["volume"])


def _indicador_carregando(texto: str) -> None:
    with ui.row().classes('items-center gap-3 py-6 px-2'):
        ui.spinner(size='md', color='orange')
        ui.label(texto).classes('text-xs text-[#71717A] max-w-xl')


def render(ctx: "Pagina") -> None:
    # A gaveta "Cases" não é desenhada nesta tela (ver sidebar.py): o cockpit já tem as próprias colunas.
    area = ui.column().classes('w-full h-full gap-0 overflow-hidden')
    with area:
        with ui.column().classes('p-6 gap-1'):
            ui.label('Centro de Comando · Radar Territorial').classes('text-lg font-bold text-white tracking-tight')
            _indicador_carregando(
                'Calculando o radar territorial... Na primeira abertura após atualizar a ferramenta '
                '(ou após uma releitura completa) pode levar alguns minutos; depois disso fica salvo e '
                'só as linhas novas de cada ingestão são processadas.'
            )

    async def _carregar():
        dados = await run.io_bound(montar_dados_mapa_radar)
        if area.is_deleted:
            return
        area.clear()
        with area:
            _montar_cockpit(ctx, dados)

    ui.timer(0.01, _carregar, once=True)


def _montar_cockpit(ctx: "Pagina", dados: Dict[str, Any]) -> None:
    st = ctx.st
    p = dados.get("payload")
    if not p or len(p["meses"]) < 2:
        with ui.column().classes('p-8 gap-2'):
            ui.label('Centro de Comando · Radar Territorial').classes('text-lg font-bold text-white')
            ui.label('É necessário ingerir arquivos em Criações Diárias para calcular os baselines territoriais.').classes('text-[#71717A]')
        return

    n_meses = len(p["meses"])
    if st.radar_mes is None or not (0 <= st.radar_mes < n_meses):
        st.radar_mes = n_meses - 1
    if st.radar_mun and st.radar_mun not in dados["indice"]:
        st.radar_mun = None
    ibge = carregar_municipios_ibge()
    estado: Dict[str, Any] = {"resumo": resumo_do_mes(dados, st.radar_mes), "mapa": None, "seq_funil": 0}

    # ------------------------------------------------------------ área coberta pelos painéis
    def _layout() -> Dict[str, int]:
        esq = (LARGURA_ESQ if st.radar_esq_aberto else LARGURA_RECOLHIDO) + 2 * MARGEM
        dir_ = (LARGURA_DIR if st.radar_dir_aberto else LARGURA_RECOLHIDO) + 2 * MARGEM
        return {"topo": 0, "base": 0, "esq": esq, "dir": dir_}

    def _alternar(lado: str):
        if lado == "esq":
            st.radar_esq_aberto = not st.radar_esq_aberto
            painel_esq.refresh()
        else:
            st.radar_dir_aberto = not st.radar_dir_aberto
            painel_dir.refresh()
        if estado["mapa"] and st.radar_visao == "mapa":
            estado["mapa"].enviar_layout(**_layout())
        else:
            centro.refresh()

    # ------------------------------------------------------------ sincronização
    def _recalcular():
        estado["resumo"] = resumo_do_mes(dados, st.radar_mes)

    def selecionar(uf, mun, mes=None, avisar_mapa=True):
        mudou_mes = mes is not None and mes != st.radar_mes
        st.radar_uf, st.radar_mun = uf, mun
        if mudou_mes:
            st.radar_mes = mes
            _recalcular()
            cabecalho.refresh()
            if st.radar_visao == "tabela":
                centro.refresh()
        # A coluna esquerda mostra a série do recorte aberto (Brasil / UF / município).
        coluna_esquerda.refresh()
        coluna_direita.refresh()
        if avisar_mapa and estado["mapa"] and st.radar_visao == "mapa":
            estado["mapa"].enviar_estado(st.radar_uf, st.radar_mun, st.radar_mes)

    def ao_mudar_mapa(valor: Dict[str, Any]):
        tipo = valor.get("tipo")
        if tipo == "mapa_pronto":
            estado["mapa"].enviar_estado(st.radar_uf, st.radar_mun, st.radar_mes, st.radar_metrica, st.radar_unidade)
            return
        if tipo != "mapa_estado":
            return
        st.radar_metrica = valor.get("metrica") or st.radar_metrica
        st.radar_unidade = valor.get("unidade") or st.radar_unidade
        mes = valor.get("mes")
        mudou = (valor.get("uf"), valor.get("municipio")) != (st.radar_uf, st.radar_mun) or (mes is not None and mes != st.radar_mes)
        if mudou:
            selecionar(valor.get("uf"), valor.get("municipio"), mes, avisar_mapa=False)

    # ------------------------------------------------------------ cabeçalho
    @ui.refreshable
    def cabecalho():
        k = estado["resumo"]["kpis"]
        with ui.row().classes('w-full h-11 shrink-0 items-center gap-4 px-4 border-b border-[#2B2B2F] bg-[#121214] no-wrap'):
            ui.icon('radar', size='18px').classes('text-[#FF7300]')
            ui.label('Radar Territorial').classes('text-sm font-bold text-white whitespace-nowrap')
            ui.label(f"Mês: {_fmt_mes(estado['resumo']['mes'])}").classes(
                'text-xs text-gray-300 px-2 py-0.5 rounded bg-[#18181A] border border-[#2B2B2F] whitespace-nowrap')
            pct_mapeado = f"{p['pct_mapeado'] * 100:.1f}".replace(".", ",")
            ui.html(
                f"<span style='font-size:11px; color:#A1A1AA; white-space:nowrap;'>"
                f"<b style='color:#C0625F'>{k['municipios_em_alerta']}</b> municípios em alerta · "
                f"<b style='color:#EDEDED'>{_fmt_int(k['acionamentos'])}</b> acionamentos "
                f"({k['pct_em_alerta'] * 100:.0f}% em alerta) · "
                f"<b style='color:#8F86B5'>{k['watchlist']}</b> watchlist · "
                f"{pct_mapeado}% mapeado</span>"
            ).classes('flex-1 min-w-0 overflow-hidden')
            if dados["auditoria"]:
                with ui.button(f"{len(dados['auditoria'])} sem código IBGE", icon='rule', color=None).props(
                        'flat dense no-caps').classes('text-[11px] text-[#C9A66B]'):
                    with ui.menu().classes('bg-[#18181A] border border-[#2B2B2F] p-2 max-h-80 overflow-y-auto custom-scroll'):
                        ui.label('Municípios do banco que não casaram com a malha do IBGE').classes('text-[11px] text-[#71717A] mb-1')
                        for a in dados["auditoria"][:60]:
                            ui.label(f"{a['cidade']} / {a['uf']} — {_fmt_int(a['total'])} acionamentos").classes('text-xs text-gray-300')
            with ui.row().classes('gap-1 shrink-0'):
                for chave, rot, icone in [("mapa", "Mapa", "map"), ("tabela", "Tabela", "table_chart")]:
                    ativo = st.radar_visao == chave
                    ui.button(rot, icon=icone, color=None, on_click=lambda c=chave: _trocar_visao(c)).props('flat dense no-caps').classes(
                        f"{CLASSE_PILULA} {PILULA_ATIVA if ativo else PILULA_INATIVA}")

    def _trocar_visao(v):
        if st.radar_visao != v:
            st.radar_visao = v
            cabecalho.refresh()
            centro.refresh()

    # ------------------------------------------------------------ coluna esquerda
    @ui.refreshable
    def coluna_esquerda():
        r = estado["resumo"]
        if st.radar_mun:
            serie = [x or 0 for x in p["municipios"][st.radar_mun]["v"]]
            escopo = f"{ibge[st.radar_mun]['nome']} / {ibge[st.radar_mun]['uf']}"
        elif st.radar_uf:
            serie = [sum((d["v"][i] or 0) for c, d in p["municipios"].items() if ibge.get(c, {}).get("uf") == st.radar_uf)
                     for i in range(n_meses)]
            escopo = st.radar_uf
        else:
            serie, escopo = r["serie_nacional"], "Brasil"
        ui.label('Evolução mensal').classes('text-[10px] font-semibold text-[#71717A] uppercase tracking-wider')
        ui.label(f"{escopo} · clique num mês para levar o mapa até ele").classes('text-[10px] text-[#52525B] -mt-1')
        ui.echart({
            "backgroundColor": "transparent",
            "grid": {"left": 34, "right": 8, "top": 10, "bottom": 22},
            "tooltip": {"trigger": "axis", "backgroundColor": "rgba(24,24,26,.95)", "borderColor": "#2B2B2F",
                        "textStyle": {"color": "#EDEDED", "fontSize": 11}},
            "xAxis": {"type": "category", "data": [_fmt_mes(m) for m in p["meses"]],
                      "axisLabel": {"color": "#71717A", "fontSize": 9}, "axisLine": {"lineStyle": {"color": "#2B2B2F"}}},
            "yAxis": {"type": "value", "axisLabel": {"color": "#71717A", "fontSize": 9},
                      "splitLine": {"lineStyle": {"color": "#1F1F23"}}},
            "series": [{"type": "line", "data": serie, "smooth": True, "symbolSize": 6,
                        "lineStyle": {"color": "#FF7300", "width": 2}, "itemStyle": {"color": "#FF7300"},
                        "areaStyle": {"color": "rgba(255,115,0,.08)"},
                        "markLine": {"symbol": "none", "silent": True, "lineStyle": {"color": "#EDEDED", "type": "dashed", "opacity": .5},
                                     "data": [{"xAxis": st.radar_mes}], "label": {"show": False}}}],
        }, on_point_click=lambda e: selecionar(st.radar_uf, st.radar_mun, int(e.data_index))).classes('w-full h-40')

        def lista(titulo, itens, cor):
            ui.label(f"{titulo} ({len(itens)})").classes('text-[10px] font-semibold uppercase tracking-wider mt-2').style(f'color:{cor}')
            if not itens:
                ui.label('nenhum').classes('text-[11px] text-[#52525B]')
            for cod in itens[:12]:
                m = ibge.get(cod, {})
                ui.label(f"{m.get('nome', cod)} / {m.get('uf', '')}").classes(
                    'text-xs text-gray-300 hover:text-[#FF7300] cursor-pointer truncate').on(
                    'click', lambda c=cod: selecionar(ibge[c]["uf"], c))

        lista('Novos alertas', r["novos"], '#C0625F')
        lista('Saíram do alerta', [s["codigo"] for s in r["sairam"]], '#6FA98A')

        k = r["kpis"]
        tot = max(k["acionamentos"], 1)
        ui.label('Composição do mês').classes('text-[10px] font-semibold text-[#71717A] uppercase tracking-wider mt-3')
        for rot, val, cor in [("Pet", k["pet"], "#C9A66B"), ("Residencial", k["residencial"], "#6C93B0"),
                              ("Outros", k["acionamentos"] - k["pet"] - k["residencial"], "#3F3F46")]:
            with ui.row().classes('w-full items-center gap-2 no-wrap'):
                ui.label(rot).classes('text-[11px] text-gray-400 w-20 shrink-0')
                with ui.element('div').classes('flex-1 h-2 rounded bg-[#1F1F23] overflow-hidden'):
                    ui.element('div').classes('h-full rounded').style(f'width:{val / tot * 100:.1f}%; background:{cor}')
                ui.label(f"{val / tot * 100:.0f}%").classes('text-[11px] text-gray-300 w-9 text-right shrink-0')

    # ------------------------------------------------------------ centro
    @ui.refreshable
    def centro():
        if st.radar_visao == "mapa":
            with ui.element('div').classes('absolute inset-0'):
                # Cópia rasa: o payload vem de cache compartilhado; o layout é desta sessão.
                estado["mapa"] = MapaTerritorialNiceGUI(ctx, "mapa_radar", ao_mudar_mapa).render({**p, "layout": _layout()})
            return
        estado["mapa"] = None
        linhas = estado["resumo"]["em_alerta"]
        lay = _layout()
        with ui.column().classes('absolute inset-0 py-3 overflow-y-auto custom-scroll').style(
                f'padding-left:{lay["esq"]}px; padding-right:{lay["dir"]}px'):
            if not linhas:
                ui.label('Nenhum município em alerta neste mês.').classes('text-[#71717A]')
                return
            df = pd.DataFrame([{
                "Município": l["nome"], "UF": l["uf"], "Volume": l["volume"],
                "Média Histórica": round(l["media"] or 0, 1), "Desvio": _fmt_razao(l["razao"], l["media"]),
                "Padrões": " • ".join(l["alertas"]),
            } for l in sorted(linhas, key=_gravidade, reverse=True)])
            ui.table.from_pandas(sanitizar_df_para_tabela(df), pagination=20).classes(
                'w-full bg-[#18181A] border border-[#2B2B2F] text-xs rounded-xl')

    # ------------------------------------------------------------ coluna direita
    @ui.refreshable
    def coluna_direita():
        if st.radar_mun:
            _painel_municipio()
        else:
            _ranking()

    def _ranking():
        linhas = estado["resumo"]["em_alerta"]
        if st.radar_uf:
            linhas = [l for l in linhas if l["uf"] == st.radar_uf]
        if st.radar_filtro:
            linhas = [l for l in linhas if l["bits"] & st.radar_filtro]
        if st.radar_ordem == "volume":
            linhas = sorted(linhas, key=lambda l: l["volume"], reverse=True)
        elif st.radar_ordem == "az":
            linhas = sorted(linhas, key=lambda l: (l["nome"], l["uf"]))
        else:
            linhas = sorted(linhas, key=_gravidade, reverse=True)

        with ui.row().classes('w-full items-center justify-between'):
            ui.label('Municípios em destaque' + (f' · {st.radar_uf}' if st.radar_uf else '')).classes('text-sm font-bold text-gray-200')
            if st.radar_uf:
                ui.button('Brasil', icon='arrow_upward', color=None, on_click=lambda: selecionar(None, None)).props(
                    'flat dense no-caps').classes('text-[11px] text-[#71717A]')
        with ui.row().classes('w-full gap-1 flex-wrap'):
            for bit, rot in FILTROS:
                ui.button(rot, color=None, on_click=lambda b=bit: (setattr(st, 'radar_filtro', b), coluna_direita.refresh())).props(
                    'flat dense no-caps').classes(f"{CLASSE_PILULA} {PILULA_ATIVA if st.radar_filtro == bit else PILULA_INATIVA}")
        with ui.row().classes('w-full gap-1'):
            for chave, rot in ORDENS:
                ui.button(rot, color=None, on_click=lambda c=chave: (setattr(st, 'radar_ordem', c), coluna_direita.refresh())).props(
                    'flat dense no-caps').classes(f"{CLASSE_PILULA} {PILULA_ATIVA if st.radar_ordem == chave else PILULA_INATIVA}")

        with ui.column().classes('w-full gap-1 flex-1 min-h-0 overflow-y-auto custom-scroll pr-1'):
            if not linhas:
                ui.label('Nenhum município em alerta com este filtro.').classes('text-xs text-[#71717A] mt-2')
            for l in linhas[:200]:
                with ui.element('div').classes(
                    'lcfo-linha-ranking w-full p-2 rounded-md bg-[#141416] hover:bg-[#222225] border border-transparent hover:border-[#2B2B2F] cursor-pointer'
                ).on('click', lambda c=l["codigo"], u=l["uf"]: selecionar(u, c)):
                    with ui.row().classes('w-full items-center gap-2 no-wrap'):
                        ui.label(l["uf"]).classes('text-[10px] font-bold text-[#71717A] w-6 shrink-0')
                        ui.label(l["nome"]).classes('text-xs font-semibold text-gray-200 truncate flex-1')
                        ui.label(_fmt_razao(l["razao"], l["media"])).classes('text-xs font-bold text-[#FF7300] shrink-0')
                        ui.label(_fmt_int(l["volume"])).classes('text-[11px] text-gray-400 w-10 text-right shrink-0')
                    ui.html(_selos(l["bits"])).classes('mt-1 ml-8')

            if not st.radar_uf:
                por_uf: Dict[str, int] = {}
                for l in estado["resumo"]["em_alerta"]:
                    por_uf[l["uf"]] = por_uf.get(l["uf"], 0) + 1
                if por_uf:
                    ui.label('Estados').classes('text-[10px] font-semibold text-[#71717A] uppercase tracking-wider mt-3')
                    for uf, n in sorted(por_uf.items(), key=lambda x: -x[1]):
                        with ui.row().classes('w-full items-center justify-between px-2 py-1 rounded hover:bg-[#222225] cursor-pointer').on(
                                'click', lambda u=uf: selecionar(u, None)):
                            ui.label(uf).classes('text-xs text-gray-300')
                            ui.label(f"{n} em alerta").classes('text-[11px] text-[#71717A]')

    def _painel_municipio():
        cod = st.radar_mun
        info, m, d = dados["indice"][cod], ibge[cod], p["municipios"][cod]
        i = st.radar_mes
        bits = d["a"][i]
        with ui.row().classes('w-full items-center justify-between'):
            ui.button('Ranking', icon='arrow_back', color=None, on_click=lambda: selecionar(st.radar_uf, None)).props(
                'flat dense no-caps').classes('text-[11px] text-[#71717A]')
            ui.label(f"IBGE {cod}").classes('text-[10px] text-[#52525B] mr-7')   # espaço para o botão de recolher
        ui.label(f"{m['nome']} / {m['uf']}").classes('text-base font-bold text-white leading-tight')
        ui.label(f"{_fmt_int(m['pop'])} hab. (Censo 2022) · {_fmt_mes(p['meses'][i])}").classes('text-[11px] text-[#71717A]')
        ui.html(
            f"<span style='font-size:11px; color:#A1A1AA'>Volume no mês <b style='color:#EDEDED'>{_fmt_int(d['v'][i])}</b> · "
            f"média histórica <b style='color:#EDEDED'>{_fmt_int(d['m'][i])}</b> · desvio "
            f"<b style='color:#FF7300'>{_fmt_razao(d['r'][i], d['m'][i])}</b></span>"
        )
        if bits:
            ui.html(_selos(bits))
        ui.echart({
            "backgroundColor": "transparent", "grid": {"left": 28, "right": 6, "top": 6, "bottom": 16},
            "xAxis": {"type": "category", "data": [MESES[int(x[5:7]) - 1] for x in p["meses"]],
                      "axisLabel": {"color": "#52525B", "fontSize": 8}, "axisLine": {"lineStyle": {"color": "#2B2B2F"}}},
            "yAxis": {"type": "value", "axisLabel": {"color": "#52525B", "fontSize": 8}, "splitLine": {"show": False}},
            "series": [
                {"type": "bar", "data": [{"value": v or 0, "itemStyle": {"color": "#C0625F" if a else "#3F3F46"}}
                                         for v, a in zip(d["v"], d["a"])], "barWidth": "60%"},
                {"type": "line", "data": d["m"], "symbol": "none", "lineStyle": {"color": "#C9A66B", "type": "dashed", "width": 1}},
            ],
        }).classes('w-full h-24')

        box = ui.column().classes('w-full gap-2 flex-1 min-h-0 overflow-y-auto custom-scroll pr-1')
        with box:
            _indicador_carregando(f"Consolidando operadores de {m['nome']}...")

        async def _carregar_funil():
            estado["seq_funil"] += 1
            minha_vez = estado["seq_funil"]
            dados_inf = await run.io_bound(
                extrair_top_infratores_municipio, cidade=info["cidade_norm"], uf=info["uf"],
                cidade_banco=info["cidade_banco"], grafias=info["grafias"])
            conhecidas = {"blacklist": set(), "base_mestra": set()}
            if dados_inf["total_ocorrencias"]:
                conhecidas = await run.io_bound(
                    marcar_entidades_conhecidas,
                    list(dados_inf["top_cpfs"]["cpf"]) if not dados_inf["top_cpfs"].empty else [],
                    list(dados_inf["top_tels"]["telefone"]) if not dados_inf["top_tels"].empty else [],
                    list(dados_inf["top_placas"]["placa"]) if not dados_inf["top_placas"].empty else [])
            if box.is_deleted or minha_vez != estado["seq_funil"]:
                return
            box.clear()
            with box:
                _montar_funil(m, dados_inf, conhecidas)

        ui.timer(0.01, _carregar_funil, once=True)

    def _montar_funil(m, dados_inf, conhecidas):
        if not dados_inf["total_ocorrencias"]:
            ui.label('Nenhuma ocorrência detalhada para este município.').classes('text-xs text-[#71717A]')
            return
        ui.label(f"Funil da praça · {_fmt_int(dados_inf['total_ocorrencias'])} ocorrências (histórico completo)").classes(
            'text-[10px] font-semibold text-[#71717A] uppercase tracking-wider')

        def marca(chave):
            out = []
            if chave in conhecidas["blacklist"]:
                out.append('<span style="font-size:9px;color:#C0625F;border:1px solid #C0625F55;border-radius:4px;padding:0 4px;margin-left:4px">Blacklist</span>')
            if chave in conhecidas["base_mestra"]:
                out.append('<span style="font-size:9px;color:#8F86B5;border:1px solid #8F86B555;border-radius:4px;padding:0 4px;margin-left:4px">Base Mestra</span>')
            return "".join(out)

        with ui.tabs().classes('w-full text-xs text-[#71717A]').props('dense') as abas:
            a_cpf = ui.tab('CPF')
            a_tel = ui.tab('TEL')
            a_placa = ui.tab('PLACA')
        with ui.tab_panels(abas, value=a_cpf).classes('w-full bg-transparent'):
            for aba, df, col, pref, extra in [(a_cpf, dados_inf["top_cpfs"], "cpf", "CPF", "titular"),
                                              (a_tel, dados_inf["top_tels"], "telefone", "TEL", None),
                                              (a_placa, dados_inf["top_placas"], "placa", "PLACA", None)]:
                with ui.tab_panel(aba).classes('p-0 gap-1'):
                    if df.empty:
                        ui.label('—').classes('text-xs text-[#52525B]')
                    for _, row in df.head(5).iterrows():
                        with ui.row().classes('w-full items-center justify-between no-wrap py-0.5'):
                            rot = f"{row[col]}" + (f" · {str(row[extra])[:18]}" if extra else "")
                            ui.html(f"<span style='font-size:11px;color:#D4D4D8'>{rot}</span>{marca(f'{pref}_{row[col]}')}").classes('min-w-0 truncate')
                            ui.label(str(int(row["total"]))).classes('text-xs font-bold text-gray-300 shrink-0')

        df_c = dados_inf["df_completo"]
        if "bairro" in df_c.columns:
            bairros = df_c["bairro"].fillna("").astype(str).str.strip()
            top_b = bairros[bairros != ""].value_counts().head(6)
            if not top_b.empty:
                ui.label('Bairros mais frequentes').classes('text-[10px] font-semibold text-[#71717A] uppercase tracking-wider mt-1')
                for b, n in top_b.items():
                    with ui.row().classes('w-full justify-between no-wrap'):
                        ui.label(b).classes('text-[11px] text-gray-300 truncate')
                        ui.label(str(n)).classes('text-[11px] text-gray-400 shrink-0')

        def _autuar_caso():
            resumo_funil = {
                "alvo_buscado": f"Conluio Regional - {m['nome']}/{m['uf']}",
                "termo_limpo": m["nome"],
                "total_assistencias": dados_inf["total_ocorrencias"],
                "qtd_cpfs": len(dados_inf["top_cpfs"]), "qtd_tels": len(dados_inf["top_tels"]),
                "qtd_placas": len(dados_inf["top_placas"]),
            }
            ents_funil = {
                "cpfs": list(dados_inf["top_cpfs"]["cpf"]) if not dados_inf["top_cpfs"].empty else [],
                "tels": list(dados_inf["top_tels"]["telefone"]) if not dados_inf["top_tels"].empty else [],
                "placas": list(dados_inf["top_placas"]["placa"]) if not dados_inf["top_placas"].empty else [],
            }
            ctx.abrir_descoberta(resumo_funil, dados_inf["df_completo"], ents_funil)

        ui.button('Autuar Conluio Regional', icon='gavel', on_click=_autuar_caso).classes(
            'w-full bg-[#FF7300] hover:bg-[#E0670A] text-white text-xs font-semibold h-8 rounded-md mt-2')

    # ------------------------------------------------------------ painéis de vidro
    def _botao_recolhido(lado: str, icone: str, dica: str):
        pos = 'left-3' if lado == "esq" else 'right-3'
        seta = 'chevron_right' if lado == "esq" else 'chevron_left'
        with ui.button(color=None, on_click=lambda: _alternar(lado)).props('flat dense').classes(
                f'absolute top-3 {pos} z-20 {VIDRO} w-11 h-11').tooltip(dica):
            with ui.column().classes('items-center gap-0'):
                ui.icon(icone, size='17px').classes('text-[#FF7300]')
                ui.icon(seta, size='13px').classes('text-[#71717A]')

    @ui.refreshable
    def painel_esq():
        if not st.radar_esq_aberto:
            _botao_recolhido("esq", 'insights', 'Abrir evolução e movimento do mês')
            return
        with ui.column().classes(f'absolute top-3 left-3 bottom-3 z-20 {VIDRO} p-3 gap-1 overflow-hidden no-wrap').style(
                f'width:{LARGURA_ESQ}px'):
            ui.button(icon='chevron_left', color=None, on_click=lambda: _alternar("esq")).props(
                'flat dense round size=sm').classes('absolute top-2 right-2 z-10 text-[#71717A] hover:text-white').tooltip('Recolher (mais mapa)')
            with ui.column().classes('w-full flex-1 min-h-0 overflow-y-auto custom-scroll gap-1 -mr-2 pr-2 no-wrap'):
                coluna_esquerda()

    @ui.refreshable
    def painel_dir():
        if not st.radar_dir_aberto:
            nome = carregar_municipios_ibge().get(st.radar_mun, {}).get("nome") if st.radar_mun else None
            _botao_recolhido("dir", 'location_on' if nome else 'format_list_numbered',
                             f"Abrir {nome}" if nome else 'Abrir municípios em destaque')
            return
        with ui.column().classes(f'absolute top-3 right-3 bottom-3 z-20 {VIDRO} p-3 gap-2 overflow-hidden no-wrap').style(
                f'width:{LARGURA_DIR}px'):
            ui.button(icon='chevron_right', color=None, on_click=lambda: _alternar("dir")).props(
                'flat dense round size=sm').classes('absolute top-2 right-2 z-10 text-[#71717A] hover:text-white').tooltip('Recolher (mais mapa)')
            coluna_direita()

    # ------------------------------------------------------------ montagem
    with ui.column().classes('w-full h-full gap-0 overflow-hidden'):
        cabecalho()
        with ui.element('div').classes('relative w-full flex-1 min-h-0 overflow-hidden'):
            centro()
            painel_esq()
            painel_dir()
