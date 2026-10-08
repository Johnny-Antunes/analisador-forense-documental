"""
Centro de Comando: radar de anomalias territoriais.

O radar e o funil por município são calculados fora do event loop
(run.io_bound) com um indicador de carregamento: antes, o cálculo síncrono
congelava a aplicação inteira e, passando do reconnect_timeout, derrubava a
conexão ("Connection lost").
"""

from __future__ import annotations

import html

import pydeck as pdk
from nicegui import run, ui

from lcfo_ui.servicos import obter_radar_anomalias_macro, extrair_top_infratores_municipio
from lcfo_ui.util_ui import sanitizar_df_para_tabela

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina


def _indicador_carregando(texto: str) -> None:
    with ui.row().classes('items-center gap-3 py-6'):
        ui.spinner(size='md', color='orange')
        ui.label(texto).classes('text-xs text-[#71717A]')


def render(ctx: "Pagina") -> None:
    with ui.column().classes('w-full h-full overflow-y-auto custom-scroll'):
        with ui.column().classes('w-full max-w-7xl mx-auto p-8 gap-4'):
            ui.label('Centro de Comando Proativo e Radar de Anomalias Territoriais').classes('text-xl font-bold text-white tracking-tight')
            ui.label('Detecção estatística de desvios volumétricos municipais, anomalias pet e abusos residenciais.').classes('text-xs text-[#71717A]')

            area_radar = ui.column().classes('w-full gap-4')
            with area_radar:
                _indicador_carregando(
                    'Calculando o radar territorial... Na primeira abertura após atualizar '
                    'a ferramenta (ou após uma releitura completa) pode levar alguns minutos; '
                    'depois disso fica salvo e só as linhas novas de cada ingestão são processadas.'
                )

            async def _carregar_radar():
                df_ano, kpis_ano, _mes_ref = await run.io_bound(obter_radar_anomalias_macro)
                if area_radar.is_deleted:
                    return
                area_radar.clear()
                with area_radar:
                    _montar_radar(ctx, df_ano, kpis_ano)

            ui.timer(0.01, _carregar_radar, once=True)


def _montar_radar(ctx: "Pagina", df_ano, kpis_ano) -> None:
    if df_ano.empty:
        ui.label('É necessário ingerir arquivos em Criações Diárias para calcular os baselines territoriais.').classes('text-[#71717A]')
    else:
        with ui.row().classes('gap-3'):
            for titulo, valor, cor in [
                ('MUNICÍPIOS', kpis_ano.get('total_anomalias', 0), 'text-red-400'),
                ('PET', kpis_ano.get('alertas_pet', 0), 'text-[#C9A66B]'),
                ('RESIDENCIAL', kpis_ano.get('alertas_res', 0), 'text-[#6C93B0]'),
                ('WATCHLIST', kpis_ano.get('alertas_watchlist', 0), 'text-[#8F86B5]'),
            ]:
                with ui.card().classes('bg-[#18181A] border border-[#2B2B2F] px-3.5 py-1.5 rounded-lg'):
                    ui.label(titulo).classes('text-[9px] text-[#71717A] font-bold')
                    ui.label(str(valor)).classes(f'text-sm font-bold {cor}')

        with ui.tabs().classes('w-full text-xs text-[#71717A] mt-2') as tabs_ano:
            tab_mapa_ano = ui.tab('Radar Territorial de Anomalias', icon='public')
            tab_tabela_ano = ui.tab('Fila de Anomalias por Município', icon='list')

        with ui.tab_panels(tabs_ano, value=tab_mapa_ano).classes('w-full bg-transparent p-0 mt-2'):
            with ui.tab_panel(tab_mapa_ano).classes('p-0'):
                df_mapa_ano = df_ano.dropna(subset=["latitude", "longitude"]).copy()
                if not df_mapa_ano.empty:
                    df_mapa_ano["raio_calc"] = df_mapa_ano["volume_atual"] * 400
                    camada_anomalias = pdk.Layer(
                        "ScatterplotLayer", data=df_mapa_ano, get_position="[longitude, latitude]",
                        get_color="[192, 98, 95, 190]", get_line_color="[255, 255, 255, 220]",
                        line_width_min_pixels=1.5, stroked=True, get_radius="raio_calc",
                        radius_min_pixels=7, radius_max_pixels=32, pickable=True, auto_highlight=True
                    )
                    view_state_ano = pdk.ViewState(
                        latitude=float(df_mapa_ano["latitude"].mean()),
                        longitude=float(df_mapa_ano["longitude"].mean()),
                        zoom=4.2, pitch=0
                    )
                    tooltip_ano = {
                        "html": "<div style='font-family:sans-serif; padding:6px;'><b style='color:#FF9955; font-size:12px;'>{cidade}/{uf}</b><br/><b>Volume Atual:</b> {volume_atual}<br/><b>Média Histórica:</b> {media_hist}<br/><b>Desvio:</b> {desvio_macro}<br/><b>Padrões:</b> <span style='color:#C9A66B;'>{alertas_str}</span></div>",
                        "style": {"backgroundColor": "#18181A", "color": "#EDEDED", "border": "1px solid #2B2B2F", "borderRadius": "8px", "fontSize": "11px"}
                    }
                    deck_ano = pdk.Deck(layers=[camada_anomalias], initial_view_state=view_state_ano, tooltip=tooltip_ano, map_style="dark")
                    ui.html(
                        f'<iframe srcdoc="{html.escape(deck_ano.to_html(as_string=True))}" style="width:100%;height:560px;border:none;border-radius:12px;"></iframe>',
                        sanitize=False
                    ).classes('w-full')
                else:
                    ui.label('Sem coordenadas disponíveis para plotagem das anomalias.').classes('text-[#71717A]')

            with ui.tab_panel(tab_tabela_ano).classes('p-0'):
                with ui.row().classes('w-full gap-3 mb-2 items-end'):
                    sel_filtro_alerta = ui.select(
                        ["TODOS", "Explosão de Volume Macro", "Anomalia em Serviços Pet", "Salto em Serviços Residenciais", "Município em Watchlist de Risco", "Pico sem Histórico Prévio"],
                        value="TODOS", label='Filtrar Categoria de Alerta:'
                    ).props('dense outlined').classes('w-72 text-xs')
                    ipt_busca_cid_ano = ui.input(label='Buscar Município ou UF:', placeholder='Digite cidade ou estado...').props('dense outlined').classes('flex-1 text-xs')

                box_tabela_ano = ui.column().classes('w-full')

                def _atualizar_tabela_ano():
                    box_tabela_ano.clear()
                    df_view = df_ano.copy()
                    if sel_filtro_alerta.value != "TODOS":
                        df_view = df_view[df_view["alertas"].apply(lambda lst: sel_filtro_alerta.value in lst)]
                    if ipt_busca_cid_ano.value:
                        b_norm = ipt_busca_cid_ano.value.upper().strip()
                        df_view = df_view[df_view["cidade"].str.upper().str.contains(b_norm) | df_view["uf"].str.upper().str.contains(b_norm)]
                    with box_tabela_ano:
                        if df_view.empty:
                            ui.label('Nenhum município corresponde ao filtro.').classes('text-[#71717A]')
                        else:
                            df_exib = df_view[['cidade', 'uf', 'mes_ref', 'volume_atual', 'media_hist', 'desvio_macro', 'meses_ativos_hist', 'pet_atual', 'res_atual', 'alertas_str', 'watchlist']].copy()
                            df_exib.columns = ['Município', 'UF', 'Mês Ref.', 'Volume Atual', 'Média Histórica', 'Desvio Relativo', 'Meses Ativos', 'Pet (Mês)', 'Residencial (Mês)', 'Padrões', 'Watchlist']
                            with ui.element('div').classes('w-full').style('max-height: calc(100vh - 460px); overflow: auto;'):
                                ui.table.from_pandas(sanitizar_df_para_tabela(df_exib), pagination=15).classes('w-full bg-[#18181A] border border-[#2B2B2F] text-xs rounded-xl')

                ipt_busca_cid_ano.on('keydown.enter', lambda e: _atualizar_tabela_ano())
                ui.button('Filtrar', on_click=_atualizar_tabela_ano).props('dense').classes('bg-[#222225] hover:bg-[#2B2B2F] text-gray-200 text-xs px-3 h-8 rounded-md border border-[#2B2B2F] mb-2')
                _atualizar_tabela_ano()

        ui.label('Desdobramento Investigativo da Anomalia Municipal').classes('text-sm font-bold text-gray-300 mt-4')
        ui.label('Isole um município alertado para consolidar os principais operadores e autuar um caso de conluio regional.').classes('text-xs text-[#71717A] mb-2')

        cidades_opcoes = {}
        for _, r in df_ano.iterrows():
            rot = f"{r['cidade']}/{r['uf']} ({r['volume_atual']} acionamentos - {r['desvio_macro']})"
            cidades_opcoes[rot] = r

        if cidades_opcoes:
            sel_cidade_alerta = ui.select(list(cidades_opcoes.keys()), value=list(cidades_opcoes.keys())[0], label='Selecione a Praça em Alerta:').props('dense outlined').classes('w-full text-xs')

            box_infratores = ui.column().classes('w-full gap-2 mt-3')

            estado_inf = {"seq": 0}

            async def _carregar_infratores():
                # Consulta fora do event loop; "seq" descarta respostas de
                # cliques anteriores que terminem depois do clique mais novo.
                estado_inf["seq"] += 1
                minha_vez = estado_inf["seq"]
                anomalia_alvo = cidades_opcoes[sel_cidade_alerta.value]
                box_infratores.clear()
                with box_infratores:
                    _indicador_carregando(f"Consolidando operadores de {anomalia_alvo['cidade']}/{anomalia_alvo['uf']}...")
                dados_inf = await run.io_bound(
                    extrair_top_infratores_municipio,
                    cidade=anomalia_alvo["cidade"], uf=anomalia_alvo["uf"],
                    cidade_banco=anomalia_alvo.get("cidade_banco", "")
                )
                if box_infratores.is_deleted or minha_vez != estado_inf["seq"]:
                    return
                box_infratores.clear()
                with box_infratores:
                    if dados_inf["total_ocorrencias"] > 0:
                        with ui.row().classes('w-full gap-4'):
                            with ui.column().classes('flex-1 gap-1'):
                                ui.label('Top Titulares (CPF):').classes('text-xs font-bold text-gray-300')
                                ui.table.from_pandas(sanitizar_df_para_tabela(dados_inf["top_cpfs"])).classes('w-full bg-[#18181A] border border-[#2B2B2F] text-xs rounded-lg')
                            with ui.column().classes('flex-1 gap-1'):
                                ui.label('Top Telefones (TEL):').classes('text-xs font-bold text-gray-300')
                                ui.table.from_pandas(sanitizar_df_para_tabela(dados_inf["top_tels"])).classes('w-full bg-[#18181A] border border-[#2B2B2F] text-xs rounded-lg')
                            with ui.column().classes('flex-1 gap-1'):
                                ui.label('Top Veículos (PLACA):').classes('text-xs font-bold text-gray-300')
                                ui.table.from_pandas(sanitizar_df_para_tabela(dados_inf["top_placas"])).classes('w-full bg-[#18181A] border border-[#2B2B2F] text-xs rounded-lg')

                        def _autuar_caso():
                            resumo_funil = {
                                "alvo_buscado": f"Conluio Regional - {anomalia_alvo['cidade']}/{anomalia_alvo['uf']}",
                                "termo_limpo": anomalia_alvo["cidade"],
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

                        ui.button('Autuar Caso a partir Desta Anomalia Territorial', on_click=_autuar_caso).classes('w-full bg-[#FF7300] hover:bg-[#E0670A] text-white text-xs font-semibold h-8 rounded-md mt-2')
                    else:
                        ui.label('Nenhuma ocorrência detalhada localizada para os parâmetros deste município.').classes('text-[#71717A]')

            ui.button('Investigar Município Selecionado', on_click=_carregar_infratores).classes('bg-[#222225] hover:bg-[#2B2B2F] text-gray-200 text-xs px-3 h-8 rounded-md border border-[#2B2B2F] mt-2')
            ui.timer(0.01, _carregar_infratores, once=True)
