"""
Módulo: anomalias_engine.py
Objetivo: Detecção de anomalias macro-territoriais, conluios em serviços pet e
          abusos residenciais a partir da base histórica de criações diárias.
"""

from typing import Tuple, Dict, Any
import sqlite3
import pandas as pd
import streamlit as st

from database import (
    get_db_connection,
    obter_versao_criacoes,
    obter_versao_cidades_risco,
    obter_cidades_risco_set
)
from utils import normalizar_cidade, obter_coordenadas, normalizar_uf_segura


@st.cache_data
def calcular_anomalias_macro_cached(
    versao_dados: Any,
    versao_cidades: Any,
    versao_schema_fix: str = "v5"
) -> Tuple[pd.DataFrame, Dict[str, int], str]:
    """
    Varre o histórico agregado de criações diárias calculando o baseline territorial
    e resolvendo as coordenadas diretamente pelo dicionário oficial do IBGE.
    
    Aplica filtros de saneamento para expurgar registros sem localização real e
    assistências puramente informativas que distorcem as capitais no radar.
    """
    conn = get_db_connection()
    try:
        query = """
            SELECT 
                cidade, 
                uf, 
                strftime('%Y-%m', data) as mes, 
                COUNT(*) as total_mes,
                SUM(CASE WHEN UPPER(servico) LIKE '%PET%' OR UPPER(servico) LIKE '%VETERIN%' THEN 1 ELSE 0 END) as pet_mes,
                SUM(CASE WHEN UPPER(servico) LIKE '%ELETRIC%' OR UPPER(servico) LIKE '%ENCANAD%' OR UPPER(servico) LIKE '%CHAVEIR%' OR UPPER(servico) LIKE '%DESENTUP%' OR UPPER(servico) LIKE '%HIDRAUL%' THEN 1 ELSE 0 END) as res_mes
            FROM criacoes_diarias
            WHERE data IS NOT NULL AND data != '' 
              AND cidade IS NOT NULL AND cidade != ''
              AND UPPER(cidade) NOT LIKE '%NAO%LOCALIZ%'
              AND UPPER(cidade) NOT LIKE '%INFORMATIV%'
              AND UPPER(servico) NOT LIKE '%INFORMATIV%'
              AND UPPER(servico) NOT LIKE '%CONSULTA%'
            GROUP BY cidade, uf, mes
            ORDER BY mes ASC
        """
        df = pd.read_sql_query(query, conn)
    except Exception:
        df = pd.DataFrame()
    finally:
        conn.close()

    if df.empty:
        return pd.DataFrame(), {}, ""

    df["cidade_norm"] = df["cidade"].astype(str).map(normalizar_cidade)
    df["uf_norm"] = df["uf"].map(normalizar_uf_segura)

    meses = sorted(df["mes"].unique())
    if len(meses) < 2:
        return pd.DataFrame(), {}, meses[-1] if meses else ""

    mes_atual = meses[-1]
    meses_hist = meses[:-1]

    df_hist = df[df["mes"].isin(meses_hist)].groupby(["cidade_norm", "uf_norm"]).agg(
        media_total=("total_mes", "mean"),
        media_pet=("pet_mes", "mean"),
        media_res=("res_mes", "mean"),
        meses_com_acionamento=("mes", "nunique")
    ).reset_index()

    df_atual = df[df["mes"] == mes_atual].copy()
    df_merged = pd.merge(df_atual, df_hist, on=["cidade_norm", "uf_norm"], how="left").fillna(0)

    # Cálculo das razões de desvio estatístico
    df_merged["razao_macro"] = (df_merged["total_mes"] / df_merged["media_total"].replace(0, 0.4)).round(1)
    df_merged["razao_pet"] = (df_merged["pet_mes"] / df_merged["media_pet"].replace(0, 0.4)).round(1)
    df_merged["razao_res"] = (df_merged["res_mes"] / df_merged["media_res"].replace(0, 0.4)).round(1)

    cidades_risco_set = obter_cidades_risco_set()
    anomalias = []

    for row in df_merged.itertuples():
        alertas = []

        # Categorias em português, sem emoji — legíveis por analistas e para exportação em relatórios
        if row.total_mes >= 15:
            if row.meses_com_acionamento >= 3:
                if row.razao_macro >= 3.0 or (row.media_total <= 3.0 and row.total_mes >= 15):
                    alertas.append("Explosão Macro")
            else:
                alertas.append("Pico sem Histórico")

        if row.pet_mes >= 6 and (row.razao_pet >= 3.0 or row.media_pet <= 1.0):
            alertas.append("Anomalia Pet")

        if row.res_mes >= 12 and (row.razao_res >= 2.5 or row.media_res <= 2.0):
            alertas.append("Salto Residencial")

        eh_watchlist = (row.cidade_norm, row.uf_norm) in cidades_risco_set
        if eh_watchlist and row.total_mes >= 5:
            alertas.append("Watchlist de Risco")

        if alertas:
            lat, lon = obter_coordenadas(row.cidade, row.uf_norm)
            anomalias.append({
                "cidade": row.cidade_norm,
                "uf": row.uf_norm,
                "mes_ref": mes_atual,
                "volume_atual": int(row.total_mes),
                "media_hist": round(float(row.media_total), 1),
                "desvio_macro": f"+{row.razao_macro}x" if row.razao_macro > 1 else f"{row.razao_macro}x",
                "meses_ativos_hist": int(row.meses_com_acionamento),
                "pet_atual": int(row.pet_mes),
                "media_pet": round(float(row.media_pet), 1),
                "res_atual": int(row.res_mes),
                "media_res": round(float(row.media_res), 1),
                "alertas": alertas,
                "alertas_str": " • ".join(alertas),
                "watchlist": "Sim" if eh_watchlist else "Não",
                "latitude": float(lat) if lat is not None else None,
                "longitude": float(lon) if lon is not None else None
            })

    df_res = pd.DataFrame(anomalias)
    if not df_res.empty:
        df_res = df_res.sort_values(by=["volume_atual"], ascending=False)

    kpis = {
        "total_anomalias": len(df_res),
        "alertas_pet": len([r for r in anomalias if "Anomalia Pet" in r["alertas"]]),
        "alertas_res": len([r for r in anomalias if "Salto Residencial" in r["alertas"]]),
        "alertas_watchlist": len([r for r in anomalias if "Watchlist de Risco" in r["alertas"]]),
        "alertas_sem_hist": len([r for r in anomalias if "Pico sem Histórico" in r["alertas"]])
    }

    return df_res, kpis, mes_atual


def obter_radar_anomalias_macro() -> Tuple[pd.DataFrame, Dict[str, int], str]:
    """Encapsulador para cálculo do radar de anomalias com verificação de versões ativas."""
    versao_dados = obter_versao_criacoes()
    versao_cidades = obter_versao_cidades_risco()
    return calcular_anomalias_macro_cached(versao_dados, versao_cidades, versao_schema_fix="v5")