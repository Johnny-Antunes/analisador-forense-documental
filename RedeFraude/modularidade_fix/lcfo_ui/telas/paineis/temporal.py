"""
Painel da mesa: Evolução Temporal.
"""

from __future__ import annotations


import pandas as pd
from nicegui import ui

from lcfo_ui.telas.contexto_mesa import ContextoMesa


def render(m: ContextoMesa, lado: str) -> None:
    df_dados = m.df_dados
    with ui.column().classes('w-full h-full pt-14 px-8'):
        df_t = df_dados.copy()
        if "data" in df_t.columns and not df_t.empty:
            df_t["mes"] = pd.to_datetime(df_t["data"], errors="coerce").dt.strftime("%Y-%m")
            df_t = df_t.dropna(subset=["mes"]).groupby("mes").size().reset_index(name="total")
            ui.echart({
                "backgroundColor": "transparent",
                "tooltip": {
                    "trigger": "axis",
                    "backgroundColor": "rgba(24, 24, 26, 0.95)",
                    "borderColor": "#2B2B2F",
                    "borderWidth": 1,
                    "padding": [8, 12],
                    "textStyle": {"color": "#EDEDED", "fontSize": 11},
                    "axisPointer": {
                        "type": "line",
                        "lineStyle": {"color": "#FF7300", "opacity": 0.25, "width": 14}
                    },
                    "extraCssText": "backdrop-filter: blur(8px); border-radius: 8px; box-shadow: 0 8px 24px rgba(0,0,0,0.5);"
                },
                "xAxis": {
                    "type": "category", "data": df_t["mes"].tolist(),
                    "axisLine": {"lineStyle": {"color": "#2B2B2F"}},
                    "axisLabel": {"color": "#71717A", "fontSize": 11}
                },
                "yAxis": {
                    "type": "value",
                    "splitLine": {"lineStyle": {"color": "#1F1F23", "opacity": 0.6}},
                    "axisLabel": {"color": "#71717A", "fontSize": 11}
                },
                "series": [{"data": df_t["total"].tolist(), "type": "bar", "itemStyle": {"color": "#FF7300", "borderRadius": [4, 4, 0, 0]}}],
            }).classes('w-full h-[520px]')
        else:
            ui.label('Sem datas disponíveis.').classes('text-[#71717A]')
