import sqlite3
import pandas as pd
import streamlit as st
from database import get_db_connection, obter_versao_criacoes, obter_versao_cidades_risco, obter_cidades_risco_set
from utils import normalizar_cidade, obter_coordenadas

@st.cache_data
def calcular_anomalias_macro_cached(versao_dados, versao_cidades):
    """
    Varre o histórico agregado de criações diárias (21 meses) e identifica:
    1. Explosões macro de volume em praças com histórico consolidado (>= 3 meses ativos).
    2. Picos isolados em praças com baixo histórico prévio (< 3 meses ativos).
    3. Anomalias de volume de Pet sem histórico prévio.
    4. Concentração atípica de serviços residenciais.
    5. Acionamentos em praças da Watchlist de Risco.
    Retorna (df_anomalias, kpis_dict, mes_referencia).
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
            WHERE data IS NOT NULL AND data != '' AND cidade IS NOT NULL AND cidade != ''
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
    df["uf_norm"] = df["uf"].astype(str).str.strip().str.upper()

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

        # 1. Distinção estatística: Explosão consolidada vs. Pico sem histórico prévio
        if row.total_mes >= 15:
            if row.meses_com_acionamento >= 3:
                if row.razao_macro >= 3.0 or (row.media_total <= 3.0 and row.total_mes >= 15):
                    alertas.append("🚨 Explosão Macro")
            else:
                alertas.append("⚠️ Pico sem Histórico")

        # 2. Anomalia Pet (Pico de consultas em cidade pequena sem histórico)
        if row.pet_mes >= 6 and (row.razao_pet >= 3.0 or row.media_pet <= 1.0):
            alertas.append("🐶 Anomalia Pet")

        # 3. Anomalia Residencial (Reincidência em massa de eletricista/encanador)
        if row.res_mes >= 12 and (row.razao_res >= 2.5 or row.media_res <= 2.0):
            alertas.append("🏠 Salto Residencial")

        # 4. Cidade da Watchlist de Risco com atividade
        eh_watchlist = (row.cidade_norm, row.uf_norm) in cidades_risco_set
        if eh_watchlist and row.total_mes >= 5:
            alertas.append("📍 Watchlist de Risco")

        if alertas:
            lat, lon = obter_coordenadas(row.cidade, row.uf)
            anomalias.append({
                "cidade": row.cidade,
                "uf": row.uf,
                "mes_ref": mes_atual,
                "volume_atual": row.total_mes,
                "media_hist": round(row.media_total, 1),
                "desvio_macro": f"+{row.razao_macro}x" if row.razao_macro > 1 else f"{row.razao_macro}x",
                "meses_ativos_hist": int(row.meses_com_acionamento),
                "pet_atual": row.pet_mes,
                "media_pet": round(row.media_pet, 1),
                "res_atual": row.res_mes,
                "media_res": round(row.media_res, 1),
                "alertas": alertas,
                "alertas_str": " • ".join(alertas),
                "watchlist": "Sim" if eh_watchlist else "Não",
                "latitude": lat,
                "longitude": lon
            })

    df_res = pd.DataFrame(anomalias)
    if not df_res.empty:
        df_res = df_res.sort_values(by=["volume_atual"], ascending=False)

    kpis = {
        "total_anomalias": len(df_res),
        "alertas_pet": len([r for r in anomalias if "🐶 Anomalia Pet" in r["alertas"]]),
        "alertas_res": len([r for r in anomalias if "🏠 Salto Residencial" in r["alertas"]]),
        "alertas_watchlist": len([r for r in anomalias if "📍 Watchlist de Risco" in r["alertas"]]),
        "alertas_sem_hist": len([r for r in anomalias if "⚠️ Pico sem Histórico" in r["alertas"]])
    }

    return df_res, kpis, mes_atual

def obter_radar_anomalias_macro():
    versao_dados = obter_versao_criacoes()
    versao_cidades = obter_versao_cidades_risco()
    return calcular_anomalias_macro_cached(versao_dados, versao_cidades)