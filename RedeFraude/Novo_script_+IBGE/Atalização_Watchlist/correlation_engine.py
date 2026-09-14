import math
import sqlite3
import pandas as pd
import networkx as nx
import streamlit as st
from datetime import datetime
from database import (
    get_db_connection, obter_versao_criacoes, 
    obter_cidades_risco_set, obter_versao_cidades_risco
)
from utils import normalizar_cidade

# =====================================================
# CÁLCULO DE DISTÂNCIA GEODÉSICA (HAVERSINE PURO OFFLINE)
# =====================================================
def calcular_distancia_km(lat1, lon1, lat2, lon2):
    """Calcula a distância em linha reta entre duas coordenadas em km."""
    if any(v is None for v in [lat1, lon1, lat2, lon2]):
        return 0.0
    try:
        r = 6371.0  # Raio médio da Terra em km
        dlat = math.radians(lat2 - lat1)
        dlon = math.radians(lon2 - lon1)
        a = math.sin(dlat / 2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2)**2
        c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
        return round(r * c, 2)
    except Exception:
        return 0.0


# =====================================================
# TRANSFORMS / ANALISADORES MODULARES INDEPENDENTES
# =====================================================
def _analisar_densidade_relacional(df):
    """Transform 1: Avalia razões estruturais entre titulares, telefones e placas."""
    pontos = 0.0
    fatores = []
    
    qtd_cpfs = max(1, df["cpf"].replace("", None).nunique())
    qtd_tels = max(1, df["telefone"].replace("", None).nunique())
    qtd_placas = max(1, df["placa"].replace("", None).nunique())

    razao_placas_cpf = qtd_placas / qtd_cpfs
    if razao_placas_cpf >= 3.0:
        pontos += 12.0
        fatores.append(f"Alta concentração veicular: {razao_placas_cpf:.1f} placas/CPF")
    elif razao_placas_cpf >= 1.8:
        pontos += 6.0

    razao_cpfs_tel = qtd_cpfs / qtd_tels
    if razao_cpfs_tel >= 2.5:
        pontos += 13.0
        fatores.append(f"Telefone compartilhado por múltiplos titulares ({razao_cpfs_tel:.1f} CPFs/contato)")
    elif razao_cpfs_tel >= 1.5:
        pontos += 7.0

    return min(25.0, pontos), fatores, {
        "razao_placas_cpf": round(razao_placas_cpf, 2),
        "razao_cpfs_tel": round(razao_cpfs_tel, 2),
        "qtd_cpfs": qtd_cpfs,
        "qtd_tels": qtd_tels,
        "qtd_placas": qtd_placas
    }


def _analisar_velocidade_temporal(df):
    """Transform 2: Avalia intensidade, rajadas operacionais e recência de acionamentos."""
    pontos = 0.0
    fatores = []
    total_assistencias = len(df)
    
    datas_validas = pd.to_datetime(df["data"], errors="coerce").dropna().sort_values()
    dias_ativos = 1
    assistencias_por_dia = 0.0
    recencia_dias = 999
    rajada_detectada = False

    if not datas_validas.empty:
        dt_min = datas_validas.min()
        dt_max = datas_validas.max()
        dias_ativos = max(1, (dt_max - dt_min).days + 1)
        assistencias_por_dia = total_assistencias / dias_ativos
        recencia_dias = (datetime.now() - dt_max).days

        if total_assistencias >= 4 and dias_ativos <= 7:
            pontos += 15.0
            rajada_detectada = True
            fatores.append(f"Ataque concentrado: {total_assistencias} acionamentos em {dias_ativos} dias")
        elif assistencias_por_dia >= 0.5:
            pontos += 10.0
            fatores.append(f"Alta frequência operacional: {assistencias_por_dia:.2f} assistências/dia")
        elif assistencias_por_dia >= 0.2:
            pontos += 5.0

        if recencia_dias <= 15:
            pontos += 10.0
            fatores.append("Atividade recente: acionamento nos últimos 15 dias")
        elif recencia_dias <= 45:
            pontos += 5.0
    else:
        pontos += 2.0

    return min(25.0, pontos), fatores, {
        "dias_ativos": dias_ativos,
        "assistencias_por_dia": round(assistencias_por_dia, 3),
        "recencia_dias": recencia_dias if recencia_dias != 999 else -1,
        "rajada_7dias": 1 if rajada_detectada else 0
    }


def _analisar_reincidencia_residencial(df):
    """
    Transform 1b: Detecta solicitação repetida de serviços residenciais 
    (chaveiro, encanador, eletricista, etc.) pelo mesmo titular em curta janela,
    típico de pressão para forçar reembolso particular ou esgotamento de apólice.
    """
    pontos = 0.0
    fatores = []
    padrao_res = r"ELETRIC|ENCANAD|CHAVEIR|DESENTUP|TELHAD|HIDRAUL|LINHA BRANCA|VIDRAC|RESIDENC"

    if "servico" not in df.columns or "cpf" not in df.columns or "data" not in df.columns:
        return 0.0, [], {"reincidencia_residencial": 0, "max_reincidencia_cpf_servico": 0}

    df_res = df[df["servico"].astype(str).str.upper().str.contains(padrao_res, na=False)].copy()
    if df_res.empty:
        return 0.0, [], {"reincidencia_residencial": 0, "max_reincidencia_cpf_servico": 0}

    df_res["data_dt"] = pd.to_datetime(df_res["data"], errors="coerce")
    df_res = df_res.dropna(subset=["data_dt"]).sort_values("data_dt")

    max_reincidencias = 0
    alerta_reincidencia = False

    for (cpf, serv), grupo in df_res.groupby(["cpf", "servico"]):
        if not cpf or len(grupo) < 3:
            continue
        datas = grupo["data_dt"].tolist()
        for i in range(len(datas) - 2):
            delta = (datas[i+2] - datas[i]).days
            if delta <= 60:
                alerta_reincidencia = True
                qtd = len(grupo)
                max_reincidencias = max(max_reincidencias, qtd)
                cpf_masc = str(cpf)[-4:].rjust(11, '*')
                fatores.append(f"Reincidência residencial abusiva: CPF {cpf_masc} solicitou '{serv}' {len(datas)}x (3x em {delta} dias)")
                break
        if alerta_reincidencia:
            break

    if alerta_reincidencia:
        pontos = 15.0 if max_reincidencias >= 4 else 10.0

    return min(20.0, pontos), fatores, {
        "reincidencia_residencial": 1 if alerta_reincidencia else 0,
        "max_reincidencia_cpf_servico": max_reincidencias
    }


def _analisar_dispersao_geografica(df, cidades_risco_set=None):
    """
    Transform 3: Avalia dispersão territorial, incompatibilidade física
    e presença em municípios da Watchlist de Risco (1a).
    """
    pontos = 0.0
    fatores = []
    
    cidades_unicas = df["cidade"].replace("", None).nunique()
    ufs_unicas = df["uf"].replace("", None).nunique() if "uf" in df.columns else 1

    if cidades_unicas >= 5:
        pontos += 10.0
        fatores.append(f"Operação multi-territorial em {cidades_unicas} cidades distintas")
    elif cidades_unicas >= 3:
        pontos += 5.0

    # 1a: Cruzamento com a Watchlist de Cidades de Risco
    if cidades_risco_set is None:
        try:
            cidades_risco_set = obter_cidades_risco_set()
        except Exception:
            cidades_risco_set = set()

    cidades_watchlist_count = 0
    if cidades_risco_set and "cidade" in df.columns and "uf" in df.columns:
        cidades_no_caso = set(zip(
            df["cidade"].astype(str).map(normalizar_cidade),
            df["uf"].astype(str).str.strip().str.upper()
        ))
        matches_risco = cidades_no_caso.intersection(cidades_risco_set)
        if matches_risco:
            cidades_watchlist_count = len(matches_risco)
            pontos += 12.0
            nomes_fmt = [f"{c}/{u}" for c, u in sorted(list(matches_risco))]
            fatores.append(f"Acionamento em município da Watchlist de Risco: {', '.join(nomes_fmt)}")

    # Deslocamento impossível por entidade
    incompatibilidade_detectada = False
    max_velocidade_estimada = 0.0
    df_geo = df.dropna(subset=["latitude", "longitude", "data"]).copy()

    if len(df_geo) >= 2:
        df_geo["dt_obj"] = pd.to_datetime(df_geo["data"], errors="coerce")
        df_geo = df_geo.sort_values(by="dt_obj")

        for ent_col in ["placa", "telefone"]:
            for ent_val, grupo in df_geo.groupby(ent_col):
                if not ent_val or len(grupo) < 2:
                    continue
                registros = grupo.to_dict("records")
                for i in range(len(registros) - 1):
                    r1 = registros[i]
                    r2 = registros[i + 1]
                    delta_horas = abs((r2["dt_obj"] - r1["dt_obj"]).total_seconds()) / 3600.0
                    
                    if 0 < delta_horas <= 24.0:
                        dist_km = calcular_distancia_km(r1["latitude"], r1["longitude"], r2["latitude"], r2["longitude"])
                        vel = dist_km / max(0.5, delta_horas)
                        if vel > max_velocidade_estimada:
                            max_velocidade_estimada = vel
                        if dist_km >= 350.0 or vel > 120.0:
                            incompatibilidade_detectada = True
                            rotulo_ent = "Placa" if ent_col == "placa" else "Telefone"
                            fatores.append(f"Incompatibilidade física: {rotulo_ent} {ent_val} em {r1['cidade']} e {r2['cidade']} ({dist_km:.0f} km em {delta_horas:.1f}h)")
                            break

    if incompatibilidade_detectada:
        pontos += 15.0

    return min(25.0, pontos), fatores, {
        "cidades_unicas": cidades_unicas,
        "ufs_unicas": ufs_unicas,
        "cidades_watchlist_count": cidades_watchlist_count,
        "incompatibilidade_fisica": 1 if incompatibilidade_detectada else 0,
        "velocidade_max_estimada_kmh": round(max_velocidade_estimada, 1)
    }


def _analisar_topologia_grafo(subgrafo, betweenness_precalculado=None):
    """Transform 4: Avalia estrutura de rede, âncora central e pontes de intermediação."""
    pontos = 0.0
    fatores = []
    max_bet = 0.0
    densidade_g = 0.0
    proporcao_hub = 0.0
    grau_hub = 0

    if subgrafo is not None and len(subgrafo) > 2:
        try:
            graus = dict(subgrafo.degree())
            hub_id = max(graus, key=graus.get)
            grau_hub = graus[hub_id]
            densidade_g = round(nx.density(subgrafo), 4)

            proporcao_hub = grau_hub / max(1, len(subgrafo))
            if proporcao_hub >= 0.5:
                pontos += 12.0
                fatores.append(f"Topologia em estrela: âncora central concentra {proporcao_hub*100:.0f}% dos vínculos")
            elif proporcao_hub >= 0.3:
                pontos += 6.0

            if betweenness_precalculado is not None:
                max_bet = float(betweenness_precalculado)
            elif len(subgrafo) > 150:
                bet_approx = nx.betweenness_centrality(subgrafo, k=25, seed=42)
                max_bet = max(bet_approx.values()) if bet_approx else 0.0
            else:
                bet_full = nx.betweenness_centrality(subgrafo)
                max_bet = max(bet_full.values()) if bet_full else 0.0

            if max_bet >= 0.35:
                pontos += 13.0
                fatores.append(f"Estrutura com pontes críticas (Betweenness de {max_bet:.2f})")
            elif max_bet >= 0.15:
                pontos += 7.0
        except Exception:
            pontos += 5.0
    else:
        pontos += 5.0

    return min(25.0, pontos), fatores, {
        "densidade_grafo": densidade_g,
        "grau_maximo_hub": grau_hub,
        "proporcao_hub": round(proporcao_hub, 2),
        "betweenness_maximo": round(max_bet, 4)
    }


# =====================================================
# ORQUESTRADOR PRINCIPAL DO FRAUD SCORE (0 - 100)
# =====================================================
def calcular_score_caso(cluster_obj, df_assistencias, subgrafo=None, betweenness_precalculado=None, cidades_risco_set=None):
    """
    Executa os Transforms modulares e consolida o Fraud Score e as features forenses.
    """
    if df_assistencias.empty:
        return 0, "BAIXO", "#10B981", ["Sem assistências registradas"], {}

    pts_rel, fat_rel, feat_rel = _analisar_densidade_relacional(df_assistencias)
    pts_vel, fat_vel, feat_vel = _analisar_velocidade_temporal(df_assistencias)
    pts_res, fat_res, feat_res = _analisar_reincidencia_residencial(df_assistencias)
    pts_geo, fat_geo, feat_geo = _analisar_dispersao_geografica(df_assistencias, cidades_risco_set)
    pts_top, fat_top, feat_top = _analisar_topologia_grafo(subgrafo, betweenness_precalculado)

    score_final = int(round(min(100.0, pts_rel + pts_vel + pts_res + pts_geo + pts_top)))

    if score_final >= 75:
        nivel = "CRÍTICO"
        cor = "#EF4444"
    elif score_final >= 50:
        nivel = "ALTO"
        cor = "#F97316"
    elif score_final >= 25:
        nivel = "MÉDIO"
        cor = "#FBBF24"
    else:
        nivel = "BAIXO"
        cor = "#10B981"

    total_assistencias = len(df_assistencias)
    taxa_guincho = 0.0
    if "servico" in df_assistencias.columns and total_assistencias > 0:
        guinchos = df_assistencias["servico"].astype(str).str.upper().str.contains("GUINCHO|REBOQUE").sum()
        taxa_guincho = round(guinchos / total_assistencias, 2)

    metricas_ml = {
        "score_total": score_final,
        "pilar_relacional": round(pts_rel, 1),
        "pilar_velocidade": round(pts_vel, 1),
        "pilar_residencial": round(pts_res, 1),
        "pilar_geografico": round(pts_geo, 1),
        "pilar_topologico": round(pts_top, 1),
        "total_assistencias": total_assistencias,
        "taxa_servico_guincho": taxa_guincho,
        "razao_assistencias_cpf": round(total_assistencias / feat_rel["qtd_cpfs"], 2),
        "razao_assistencias_placa": round(total_assistencias / feat_rel["qtd_placas"], 2),
        **feat_rel,
        **feat_vel,
        **feat_res,
        **feat_geo,
        **feat_top
    }

    todos_fatores = fat_rel + fat_vel + fat_res + fat_geo + fat_top
    if not todos_fatores:
        todos_fatores.append("Comportamento relacional estável dentro da normalidade")

    return score_final, nivel, cor, todos_fatores, metricas_ml


# =====================================================
# BATCH CACHEADO PARA A FILA DE TRIAGEM (TELA 1)
# =====================================================
@st.cache_data
def gerar_scores_triagem_cached(ids_tupla, nos_serializados, versao_dados, versao_cidades, dia_atual):
    """
    Calcula os scores em lote para a triagem.
    'versao_cidades' garante recálculo autônomo se a Watchlist mudar.
    'dia_atual' força recálculo diário automático da recência.
    """
    scores_map = {}
    if not ids_tupla:
        return scores_map

    cidades_risco_set = obter_cidades_risco_set()
    conn = get_db_connection()
    try:
        for cid, nodes in zip(ids_tupla, nos_serializados):
            cpfs = [n.replace("CPF_", "") for n in nodes if n.startswith("CPF_")]
            tels = [n.replace("TEL_", "") for n in nodes if n.startswith("TEL_")]
            placas = [n.replace("PLACA_", "") for n in nodes if n.startswith("PLACA_")]

            clausulas, params = [], []
            if cpfs:
                clausulas.append(f"cpf IN ({','.join(['?']*len(cpfs))})")
                params.extend(cpfs)
            if tels:
                clausulas.append(f"telefone IN ({','.join(['?']*len(tels))})")
                params.extend(tels)
            if placas:
                clausulas.append(f"placa IN ({','.join(['?']*len(placas))})")
                params.extend(placas)

            if clausulas:
                query = f"""
                    SELECT data, titular, cpf, telefone, placa, servico, cidade, uf, latitude, longitude
                    FROM assistencias
                    WHERE {' OR '.join(clausulas)}
                """
                df_c = pd.read_sql_query(query, conn, params=params)
            else:
                df_c = pd.DataFrame()

            score, nivel, cor, fatores, _ = calcular_score_caso(None, df_c, subgrafo=None, cidades_risco_set=cidades_risco_set)
            scores_map[cid] = {
                "score": score,
                "nivel": nivel,
                "cor": cor,
                "fatores": fatores
            }
        return scores_map
    finally:
        conn.close()


def obter_scores_triagem(cluster_info):
    """Encapsulador que alimenta o cache com tipos imutáveis, versões do banco e data de hoje."""
    if not cluster_info:
        return {}
    top_100 = cluster_info[:100]
    ids_tupla = tuple(c["id"] for c in top_100)
    nos_serializados = tuple(tuple(sorted(c["nodes"])) for c in top_100)
    versao = obter_versao_criacoes()
    versao_cidades = obter_versao_cidades_risco()
    dia_atual = datetime.now().strftime("%Y-%m-%d")
    return gerar_scores_triagem_cached(ids_tupla, nos_serializados, versao, versao_cidades, dia_atual)