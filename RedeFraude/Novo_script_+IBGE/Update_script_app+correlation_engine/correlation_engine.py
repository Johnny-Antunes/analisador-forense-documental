import math
import sqlite3
import pandas as pd
import networkx as nx
import streamlit as st
from datetime import datetime
from database import get_db_connection, obter_versao_criacoes

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
    """Transform 2: Avalia intensidade, rajadas operacionais e recência de sinistros."""
    pontos = 0.0
    fatores = []
    total_sinistros = len(df)
    
    datas_validas = pd.to_datetime(df["data"], errors="coerce").dropna().sort_values()
    dias_ativos = 1
    sinistros_por_dia = 0.0
    recencia_dias = 999
    rajada_detectada = False

    if not datas_validas.empty:
        dt_min = datas_validas.min()
        dt_max = datas_validas.max()
        dias_ativos = max(1, (dt_max - dt_min).days + 1)
        sinistros_por_dia = total_sinistros / dias_ativos
        recencia_dias = (datetime.now() - dt_max).days

        if total_sinistros >= 4 and dias_ativos <= 7:
            pontos += 15.0
            rajada_detectada = True
            fatores.append(f"Ataque concentrado: {total_sinistros} acionamentos em {dias_ativos} dias")
        elif sinistros_por_dia >= 0.5:
            pontos += 10.0
            fatores.append(f"Alta frequência operacional: {sinistros_por_dia:.2f} sinistros/dia")
        elif sinistros_por_dia >= 0.2:
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
        "sinistros_por_dia": round(sinistros_por_dia, 3),
        "recencia_dias": recencia_dias if recencia_dias != 999 else -1,
        "rajada_7dias": 1 if rajada_detectada else 0
    }


def _analisar_dispersao_geografica(df):
    """
    Transform 3: Avalia dispersão e incompatibilidade física (deslocamento impossível).
    CORREÇÃO B APLICADA: Varre placas e telefones sem quebrar o loop prematuramente,
    garantindo que todas as evidências entrem no laudo.
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
                            break  # Para este valor específico, já achou salto impossível; continua checando as demais entidades!

    if incompatibilidade_detectada:
        pontos += 15.0

    return min(25.0, pontos), fatores, {
        "cidades_unicas": cidades_unicas,
        "ufs_unicas": ufs_unicas,
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

            # Otimização: se já veio pré-calculado do graph_engine, usa direto!
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
def calcular_score_caso(cluster_obj, df_assistencias, subgrafo=None, betweenness_precalculado=None):
    """
    Executa os Transforms independentes e consolida o Fraud Score e as features forenses.
    """
    if df_assistencias.empty:
        return 0, "BAIXO", "#10B981", ["Sem assistências registradas"], {}

    pts_rel, fat_rel, feat_rel = _analisar_densidade_relacional(df_assistencias)
    pts_vel, fat_vel, feat_vel = _analisar_velocidade_temporal(df_assistencias)
    pts_geo, fat_geo, feat_geo = _analisar_dispersao_geografica(df_assistencias)
    pts_top, fat_top, feat_top = _analisar_topologia_grafo(subgrafo, betweenness_precalculado)

    score_final = int(round(min(100.0, pts_rel + pts_vel + pts_geo + pts_top)))

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

    total_sinistros = len(df_assistencias)
    taxa_guincho = 0.0
    if "servico" in df_assistencias.columns and total_sinistros > 0:
        guinchos = df_assistencias["servico"].astype(str).str.upper().str.contains("GUINCHO|REBOQUE").sum()
        taxa_guincho = round(guinchos / total_sinistros, 2)

    metricas_ml = {
        "score_total": score_final,
        "pilar_relacional": round(pts_rel, 1),
        "pilar_velocidade": round(pts_vel, 1),
        "pilar_geografico": round(pts_geo, 1),
        "pilar_topologico": round(pts_top, 1),
        "total_sinistros": total_sinistros,
        "taxa_servico_guincho": taxa_guincho,
        "razao_assistencias_cpf": round(total_sinistros / feat_rel["qtd_cpfs"], 2),
        "razao_assistencias_placa": round(total_sinistros / feat_rel["qtd_placas"], 2),
        **feat_rel,
        **feat_vel,
        **feat_geo,
        **feat_top
    }

    todos_fatores = fat_rel + fat_vel + fat_geo + fat_top
    if not todos_fatores:
        todos_fatores.append("Comportamento relacional estável dentro da normalidade")

    return score_final, nivel, cor, todos_fatores, metricas_ml


# =====================================================
# BATCH CACHEADO PARA A FILA DE TRIAGEM (TELA 1)
# =====================================================
@st.cache_data
def gerar_scores_triagem_cached(ids_tupla, nos_serializados, versao_dados, dia_atual):
    """
    Calcula os scores em lote para a triagem.
    CORREÇÃO A APLICADA: 'dia_atual' faz parte da chave de cache,
    forçando recálculo diário automático da recência mesmo sem arquivos novos.
    """
    scores_map = {}
    if not ids_tupla:
        return scores_map

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

            score, nivel, cor, fatores, _ = calcular_score_caso(None, df_c, subgrafo=None)
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
    """Encapsulador que alimenta o cache com tipos imutáveis e a data de hoje."""
    if not cluster_info:
        return {}
    top_100 = cluster_info[:100]
    ids_tupla = tuple(c["id"] for c in top_100)
    nos_serializados = tuple(tuple(sorted(c["nodes"])) for c in top_100)
    versao = obter_versao_criacoes()
    dia_atual = datetime.now().strftime("%Y-%m-%d")
    return gerar_scores_triagem_cached(ids_tupla, nos_serializados, versao, dia_atual)