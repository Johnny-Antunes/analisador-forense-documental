"""
Módulo: anomalias_engine.py
Objetivo: Detecção de anomalias macro-territoriais e extração de entidades ofensivas
          (Funil da Tela 4) a partir da base histórica de criações diárias.

AJUSTES DESTA RODADA:
  - Removida a dependência de Streamlit nesta camada (import streamlit as st).
    calcular_anomalias_macro_cached usava @st.cache_data apenas para computar
    uma vez por combinação de (versao_dados, versao_cidades, versao_schema_fix)
    e guardar em memória — comportamento reproduzido de forma idêntica com
    functools.lru_cache, sem depender de haver uma sessão Streamlit ativa (o
    que gerava os warnings "No runtime found" / "missing ScriptRunContext" no
    terminal ao rodar via NiceGUI puro). Os parâmetros já são tuplas/strings
    hasháveis, então nenhuma assinatura precisou mudar. Nenhuma outra função
    deste módulo usava Streamlit.
"""

from typing import Tuple, Dict, Any, List
import sqlite3
import json
import threading
import pandas as pd
from functools import lru_cache

from database import (
    get_db_connection,
    obter_versao_criacoes,
    obter_versao_cidades_risco,
    obter_cidades_risco_set
)
from utils import normalizar_cidade, obter_coordenadas, normalizar_uf_segura, canonizar_municipio


# Mesmos filtros e somas da consulta original. UPPER() foi retirado porque o
# LIKE do SQLite já ignora maiúsculas/minúsculas em ASCII (e o UPPER() do
# SQLite também só converte ASCII) — resultado idêntico, sem chamar uma
# função por linha. "id > ?" permite agregar só as linhas novas.
_SQL_AGREGADO_MENSAL = """
    SELECT
        cidade,
        uf,
        strftime('%Y-%m', data) as mes,
        COUNT(*) as total_mes,
        SUM(CASE WHEN servico LIKE '%PET%' OR servico LIKE '%VETERIN%' THEN 1 ELSE 0 END) as pet_mes,
        SUM(CASE WHEN servico LIKE '%ELETRIC%' OR servico LIKE '%ENCANAD%' OR servico LIKE '%CHAVEIR%' OR servico LIKE '%DESENTUP%' OR servico LIKE '%HIDRAUL%' THEN 1 ELSE 0 END) as res_mes
    FROM criacoes_diarias
    WHERE id > ?
      AND data IS NOT NULL AND data != ''
      AND cidade IS NOT NULL AND cidade != ''
      AND cidade NOT LIKE '%NAO%LOCALIZ%'
      AND cidade NOT LIKE '%INFORMATIV%'
      AND servico NOT LIKE '%INFORMATIV%'
      AND servico NOT LIKE '%CONSULTA%'
    GROUP BY cidade, uf, mes
"""
_ESQUEMA_AGREGADO = "v1"
_LOCK_AGREGADO = threading.Lock()


def _ler_agregado_mensal(conn) -> pd.DataFrame:
    """
    Totais mensais por município, equivalentes ao GROUP BY sobre toda a
    criacoes_diarias, mas mantidos numa tabela persistida e atualizados só
    com as linhas novas (id > último processado). A varredura completa
    (lenta em bases grandes) acontece uma única vez — não a cada reinício
    do app nem a cada ingestão.

    Refaz do zero quando o conjunto deixa de ser "só acréscimo": linhas
    apagadas do início (MIN(id) mudou), tabela encolheu (MAX(id) menor) ou
    invalidação explícita (releitura forçada / reparo de codificação apagam
    a meta — ver database.invalidar_agregados_criacoes).
    """
    with _LOCK_AGREGADO:
        cur = conn.cursor()
        cur.execute("BEGIN IMMEDIATE")
        try:
            # MIN e MAX em subconsultas separadas: juntos no mesmo SELECT o SQLite
            # perde a otimização pela chave primária e varre a tabela inteira.
            cur.execute("SELECT (SELECT MAX(id) FROM criacoes_diarias) AS max_id, (SELECT MIN(id) FROM criacoes_diarias) AS min_id")
            r = cur.fetchone()
            max_id, min_id = (r["max_id"] or 0), (r["min_id"] or 0)

            cur.execute("SELECT esquema, ultimo_id, min_id FROM cache_agregado_criacoes_meta WHERE id = 1")
            meta = cur.fetchone()
            reconstruir = (
                meta is None
                or meta["esquema"] != _ESQUEMA_AGREGADO
                or meta["min_id"] != min_id
                or max_id < (meta["ultimo_id"] or 0)
            )
            ultimo_id = 0 if reconstruir else (meta["ultimo_id"] or 0)
            if reconstruir:
                cur.execute("DELETE FROM cache_agregado_criacoes")

            if max_id > ultimo_id or reconstruir:
                cur.execute(
                    "INSERT INTO cache_agregado_criacoes (cidade, uf, mes, total_mes, pet_mes, res_mes) "
                    + _SQL_AGREGADO_MENSAL, (ultimo_id,)
                )
                cur.execute(
                    "INSERT OR REPLACE INTO cache_agregado_criacoes_meta (id, esquema, ultimo_id, min_id) VALUES (1, ?, ?, ?)",
                    (_ESQUEMA_AGREGADO, max_id, min_id),
                )
            conn.commit()
        except Exception:
            conn.rollback()
            raise

    # Cada ingestão acrescenta grupos parciais; somar de novo por
    # (cidade, uf, mes) dá exatamente o total da consulta original
    # (o GROUP BY trata NULL como um grupo só, como antes).
    return pd.read_sql_query("""
        SELECT cidade, uf, mes,
               SUM(total_mes) AS total_mes, SUM(pet_mes) AS pet_mes, SUM(res_mes) AS res_mes
        FROM cache_agregado_criacoes
        GROUP BY cidade, uf, mes
        ORDER BY mes ASC
    """, conn)


@lru_cache(maxsize=4)
def _agregado_consolidado_cached(versao_dados: Any) -> pd.DataFrame:
    conn = get_db_connection()
    try:
        df = _ler_agregado_mensal(conn)
    except Exception:
        df = pd.DataFrame()
    finally:
        conn.close()
    if df.empty:
        return df
    df["uf_norm"] = df["uf"].map(normalizar_uf_segura)
    df["cidade_norm"] = [canonizar_municipio(normalizar_cidade(str(c)), u) for c, u in zip(df["cidade"], df["uf_norm"])]
    return _consolidar_grafias(df)


def _agregado_consolidado(versao_dados: Any = None) -> pd.DataFrame:
    """
    Totais mensais por município (todas as grafias somadas, nome oficial),
    base comum do radar e do mapa. Devolve uma cópia (o cache não é alterado).
    """
    return _agregado_consolidado_cached(versao_dados if versao_dados is not None else obter_versao_criacoes()).copy()


def _consolidar_grafias(df: pd.DataFrame) -> pd.DataFrame:
    """
    Junta as grafias de um mesmo município ("São Paulo", "SAO PAULO", "sp"...)
    numa linha por (cidade_norm, uf_norm, mes), somando os totais. Antes cada
    grafia virava um município à parte: o volume do mês se dividia entre elas
    e a média histórica era tirada sobre grafias, não sobre meses.

    Mantém `cidade`/`uf` com a grafia mais frequente (exibição, coordenadas)
    e `grafias_banco` com todas as grafias originais (cidade, uf) — o funil
    usa essa lista para buscar as ocorrências de todas elas.
    """
    chave = ["cidade_norm", "uf_norm"]
    peso = (df.groupby(chave + ["cidade", "uf"], dropna=False)["total_mes"].sum().reset_index()
              .sort_values(["total_mes", "cidade", "uf"], ascending=[False, True, True], kind="mergesort"))
    canonica = peso.drop_duplicates(chave)[chave + ["cidade", "uf"]]
    grafias = (peso.groupby(chave, dropna=False)
                   .apply(lambda g: tuple(sorted({(str(c), str(u)) for c, u in zip(g["cidade"], g["uf"])})),
                          include_groups=False)
                   .rename("grafias_banco").reset_index())
    somado = df.groupby(chave + ["mes"], dropna=False)[["total_mes", "pet_mes", "res_mes"]].sum().reset_index()
    return (somado.merge(canonica, on=chave, how="left").merge(grafias, on=chave, how="left")
                  .sort_values("mes", kind="mergesort").reset_index(drop=True))


@lru_cache(maxsize=8)
def calcular_anomalias_macro_cached(
    versao_dados: Any,
    versao_cidades: Any,
    versao_schema_fix: str = "v5"
) -> Tuple[pd.DataFrame, Dict[str, int], str]:
    """
    Varre o histórico agregado de criações diárias calculando o baseline territorial
    e resolvendo as coordenadas diretamente pelo dicionário oficial do IBGE.
    
    Aplica filtros de saneamento para expurgar registros sem localização real e
    assistências puramente informativas que distorcem o radar.
    """
    df = _agregado_consolidado(versao_dados)
    if df.empty:
        return pd.DataFrame(), {}, ""

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

    df_merged["razao_macro"] = (df_merged["total_mes"] / df_merged["media_total"].replace(0, 0.4)).round(1)
    df_merged["razao_pet"] = (df_merged["pet_mes"] / df_merged["media_pet"].replace(0, 0.4)).round(1)
    df_merged["razao_res"] = (df_merged["res_mes"] / df_merged["media_res"].replace(0, 0.4)).round(1)

    cidades_risco_set = obter_cidades_risco_set()
    anomalias = []

    for row in df_merged.itertuples():
        alertas = []

        if row.total_mes >= 15:
            if row.meses_com_acionamento >= 3:
                if row.razao_macro >= 3.0 or (row.media_total <= 3.0 and row.total_mes >= 15):
                    alertas.append("Explosão de Volume Macro")
            else:
                alertas.append("Pico sem Histórico Prévio")

        if row.pet_mes >= 6 and (row.razao_pet >= 3.0 or row.media_pet <= 1.0):
            alertas.append("Anomalia em Serviços Pet")

        if row.res_mes >= 12 and (row.razao_res >= 2.5 or row.media_res <= 2.0):
            alertas.append("Salto em Serviços Residenciais")

        eh_watchlist = (row.cidade_norm, row.uf_norm) in cidades_risco_set
        if eh_watchlist and row.total_mes >= 5:
            alertas.append("Município em Watchlist de Risco")

        if alertas:
            lat, lon = obter_coordenadas(row.cidade, row.uf_norm)
            anomalias.append({
                "cidade": row.cidade_norm,
                "cidade_banco": str(row.cidade).strip(),
                "grafias_banco": row.grafias_banco,
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
        "alertas_pet": len([r for r in anomalias if "Anomalia em Serviços Pet" in r["alertas"]]),
        "alertas_res": len([r for r in anomalias if "Salto em Serviços Residenciais" in r["alertas"]]),
        "alertas_watchlist": len([r for r in anomalias if "Município em Watchlist de Risco" in r["alertas"]]),
        "alertas_sem_hist": len([r for r in anomalias if "Pico sem Histórico Prévio" in r["alertas"]])
    }

    return df_res, kpis, mes_atual


def obter_radar_anomalias_macro() -> Tuple[pd.DataFrame, Dict[str, int], str]:
    """Encapsulador para cálculo do radar de anomalias com verificação de versões ativas."""
    versao_dados = obter_versao_criacoes()
    versao_cidades = obter_versao_cidades_risco()
    return calcular_anomalias_macro_cached(versao_dados, versao_cidades, versao_schema_fix="v7")


def _ler_ocorrencias_por_grafias(conn, grafias) -> pd.DataFrame:
    """Ocorrências (fora informativos/consultas) de todas as grafias (cidade, uf) de um município."""
    cidades_g = sorted({c for c, _ in grafias})
    ufs_g = sorted({u for _, u in grafias})
    return pd.read_sql_query(f"""
        SELECT 
            data, id_assistencia, titular, cpf, telefone, placa, 
            servico, bairro, cidade, uf, latitude, longitude, empresa_cliente
        FROM criacoes_diarias
        WHERE UPPER(uf) IN (SELECT UPPER(value) FROM json_each(?))
          AND UPPER(cidade) IN (SELECT UPPER(value) FROM json_each(?))
          AND uf IN ({",".join("?" * len(ufs_g))})
          AND cidade IN ({",".join("?" * len(cidades_g))})
          AND UPPER(servico) NOT LIKE '%INFORMATIV%'
          AND UPPER(servico) NOT LIKE '%CONSULTA%'
        ORDER BY data DESC
    """, conn, params=[json.dumps(ufs_g), json.dumps(cidades_g), *ufs_g, *cidades_g],
        dtype={"latitude": "float64", "longitude": "float64"})


def extrair_top_infratores_municipio(
    cidade: str,
    uf: str,
    cidade_banco: str = "",
    limite: int = 10,
    grafias: Any = None,
) -> Dict[str, Any]:
    """
    `grafias`: pares (cidade, uf) exatamente como gravados no banco — vindos de
    `grafias_banco` do radar. Quando informados, o funil reúne as ocorrências
    de todas as grafias do município (o mesmo conjunto que o radar somou).

    Versão com cache por (município, versão de criacoes_diarias): reabrir o
    Centro de Comando ou alternar entre praças já vistas não refaz a consulta.
    Devolve cópias dos DataFrames para que quem chama possa alterá-los sem
    contaminar o cache.
    """
    grafias_t = tuple(sorted({(str(c), str(u)) for c, u in grafias})) if grafias else ()
    res = _extrair_top_infratores_cached(cidade, uf, cidade_banco, limite, obter_versao_criacoes(), grafias_t)
    return {k: (v.copy() if isinstance(v, pd.DataFrame) else (list(v) if isinstance(v, list) else v))
            for k, v in res.items()}


@lru_cache(maxsize=32)
def _extrair_top_infratores_cached(
    cidade: str,
    uf: str,
    cidade_banco: str,
    limite: int,
    versao_dados: Any,
    grafias: tuple = (),
) -> Dict[str, Any]:
    """
    Funil da Tela 4: Localiza todas as assistências de um município alertado
    e identifica os principais operadores (CPFs, telefones e placas) causadores do pico.
    
    Aplica tratamento duplo de grafia para assegurar correspondência exata no banco
    mesmo diante de variações de acentuação (ex: 'SÃO PAULO' vs 'SAO PAULO').
    """
    conn = get_db_connection()
    cidade_norm_busca = normalizar_cidade(cidade)
    cidade_banco_busca = cidade_banco.strip().upper() if cidade_banco else cidade_norm_busca
    uf_busca = normalizar_uf_segura(uf)

    try:
        if grafias:
            df = _ler_ocorrencias_por_grafias(conn, grafias)
        else:
            query = """
                SELECT 
                    data, id_assistencia, titular, cpf, telefone, placa, 
                    servico, bairro, cidade, uf, latitude, longitude, empresa_cliente
                FROM criacoes_diarias
                WHERE (UPPER(cidade) = ? OR UPPER(cidade) = ?)
                  AND UPPER(uf) = ?
                  AND UPPER(servico) NOT LIKE '%INFORMATIV%'
                  AND UPPER(servico) NOT LIKE '%CONSULTA%'
                ORDER BY data DESC
            """
            df = pd.read_sql_query(query, conn, params=[cidade_banco_busca, cidade_norm_busca, uf_busca])
        
            # Fallback defensivo: se não encontrar por correspondência direta de string no SQL,
            # filtra a UF e normaliza via Python para garantir que nenhuma praça acentuada seja perdida
            # Mesma regra de antes (normalizar_cidade(cidade) == cidade buscada),
            # mas aplicada só aos nomes DISTINTOS de cidade da UF — antes a UF
            # inteira era carregada no pandas para normalizar linha a linha.
            if df.empty:
                nomes_uf = [r[0] for r in conn.execute(
                    "SELECT DISTINCT cidade FROM criacoes_diarias WHERE UPPER(uf) = ?", [uf_busca]
                ).fetchall()]
                grafias_uf = [c for c in nomes_uf if c is not None and normalizar_cidade(str(c)) == cidade_norm_busca]
                if grafias_uf:
                    marcadores = ",".join("?" * len(grafias_uf))
                    query_fallback = f"""
                        SELECT
                            data, id_assistencia, titular, cpf, telefone, placa,
                            servico, bairro, cidade, uf, latitude, longitude, empresa_cliente
                        FROM criacoes_diarias
                        WHERE UPPER(uf) = ?
                          AND UPPER(cidade) IN (SELECT UPPER(value) FROM json_each(?))
                          AND cidade IN ({marcadores})
                          AND UPPER(servico) NOT LIKE '%INFORMATIV%'
                          AND UPPER(servico) NOT LIKE '%CONSULTA%'
                    """
                    # dtype fixo: o fallback antigo lia a UF inteira, onde sempre há coordenadas,
                    # então latitude/longitude vinham como float (NaN quando ausentes).
                    df = pd.read_sql_query(query_fallback, conn, params=[uf_busca, json.dumps(grafias_uf), *grafias_uf],
                                           dtype={"latitude": "float64", "longitude": "float64"})
    except Exception:
        df = pd.DataFrame()
    finally:
        conn.close()

    if df.empty:
        return {
            "total_ocorrencias": 0,
            "df_completo": pd.DataFrame(),
            "top_cpfs": pd.DataFrame(),
            "top_tels": pd.DataFrame(),
            "top_placas": pd.DataFrame(),
            "nos_sugeridos": []
        }

    # Agrupamento dos maiores ofensores por frequência
    df_cpfs = (
        df[df["cpf"].str.strip() != ""]
        .groupby(["cpf", "titular"])
        .size()
        .reset_index(name="total")
        .sort_values("total", ascending=False)
        .head(limite)
    )
    df_tels = (
        df[df["telefone"].str.strip() != ""]
        .groupby("telefone")
        .size()
        .reset_index(name="total")
        .sort_values("total", ascending=False)
        .head(limite)
    )
    df_placas = (
        df[df["placa"].str.strip() != ""]
        .groupby("placa")
        .size()
        .reset_index(name="total")
        .sort_values("total", ascending=False)
        .head(limite)
    )

    nos_alvo = set()
    for c in df_cpfs["cpf"]:
        nos_alvo.add(f"CPF_{c}")
    for t in df_tels["telefone"]:
        nos_alvo.add(f"TEL_{t}")
    for p in df_placas["placa"]:
        nos_alvo.add(f"PLACA_{p}")

    return {
        "total_ocorrencias": len(df),
        "df_completo": df,
        "top_cpfs": df_cpfs,
        "top_tels": df_tels,
        "top_placas": df_placas,
        "nos_sugeridos": list(nos_alvo)
    }