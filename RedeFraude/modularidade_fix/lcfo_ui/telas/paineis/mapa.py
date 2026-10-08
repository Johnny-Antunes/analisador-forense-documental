"""
Painel da mesa: Radar Territorial (Mapa).
"""

from __future__ import annotations

import html

import pandas as pd
import pydeck as pdk
from nicegui import ui

from lcfo_ui.util_ui import sanitizar_df_para_tabela

from lcfo_ui.telas.contexto_mesa import ContextoMesa


def render(m: ContextoMesa, lado: str) -> None:
    df_dados = m.df_dados
    with ui.column().classes('w-full h-full pt-14 px-8 pb-8 overflow-y-auto custom-scroll gap-4'):
        df_map = df_dados.dropna(subset=["latitude", "longitude"]) if "latitude" in df_dados.columns else pd.DataFrame()
        if not df_map.empty:
            campos_possiveis = ["titular", "cpf", "telefone", "placa", "servico", "cidade", "uf", "data", "empresa_cliente"]
            campos_tooltip = "<br/>".join(
                f"<b>{c.replace('_', ' ').title()}:</b> {{{c}}}" for c in campos_possiveis if c in df_map.columns
            )
            tooltip_caso_map = {
                "html": f"<div style='font-family:sans-serif; padding:8px; line-height:1.5;'>{campos_tooltip}</div>",
                "style": {"backgroundColor": "#18181A", "color": "#EDEDED", "border": "1px solid #2B2B2F", "borderRadius": "8px", "fontSize": "11px"}
            }
            deck = pdk.Deck(
                layers=[pdk.Layer("ScatterplotLayer", data=df_map, get_position="[longitude, latitude]", get_color="[255, 115, 0, 200]", get_radius=3500, pickable=True)],
                initial_view_state=pdk.ViewState(latitude=float(df_map["latitude"].mean()), longitude=float(df_map["longitude"].mean()), zoom=6),
                tooltip=tooltip_caso_map,
                map_style="dark",
            )
            ui.html(
                f'<iframe srcdoc="{html.escape(deck.to_html(as_string=True))}" style="width:100%;height:600px;border:none;border-radius:12px;"></iframe>',
                sanitize=False
            ).classes('w-full')

            if "cidade" in df_dados.columns and "uf" in df_dados.columns:
                df_resumo_mapa = (
                    df_dados.groupby(["cidade", "uf"]).size()
                    .reset_index(name="Total")
                    .sort_values(by="Total", ascending=False)
                )
                ui.table.from_pandas(sanitizar_df_para_tabela(df_resumo_mapa), pagination=15).classes('w-full bg-[#18181A] border border-[#2B2B2F] text-xs rounded-xl')
        else:
            ui.label('Sem coordenadas geográficas para plotagem.').classes('text-[#71717A]')
