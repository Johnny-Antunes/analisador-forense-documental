"""
Módulo: territorio_engine.py
Objetivo: Base do mapa territorial — junção município -> código IBGE, alertas do
          radar calculados mês a mês (linha do tempo) e payloads do mapa em
          Canvas (componente_mapa/), tanto do Centro de Comando quanto da mesa.

As regras de alerta são exatamente as de anomalias_engine (o último mês é
idêntico ao radar — coberto por teste); aqui elas são aplicadas de forma
vetorizada a cada mês, usando como histórico só os meses anteriores a ele.
"""

from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

import numpy as np
import pandas as pd

from anomalias_engine import _agregado_consolidado
from database import obter_cidades_risco_set, obter_versao_criacoes
from utils import canonizar_municipio, normalizar_cidade, normalizar_uf_segura

CAMINHO_MALHA = Path(__file__).parent / "componente_mapa" / "data" / "brasil.topo.json"

# Bits dos alertas (payload do mapa) e rótulos idênticos aos do radar.
ALERTAS: List[Tuple[int, str]] = [
    (1, "Explosão de Volume Macro"),
    (2, "Anomalia em Serviços Pet"),
    (4, "Salto em Serviços Residenciais"),
    (8, "Município em Watchlist de Risco"),
    (16, "Pico sem Histórico Prévio"),
]
FONTE_MALHA = "Malha e população: IBGE (Censo 2022)"


# ---------------------------------------------------------------- junção com o IBGE
@lru_cache(maxsize=1)
def carregar_municipios_ibge() -> Dict[str, Dict[str, Any]]:
    """{código IBGE: {"nome", "uf", "pop"}} lido das propriedades da malha."""
    if not CAMINHO_MALHA.exists():
        return {}
    topo = json.loads(CAMINHO_MALHA.read_text(encoding="utf-8"))
    return {g["properties"]["id"]: {"nome": g["properties"]["n"], "uf": g["properties"]["uf"], "pop": g["properties"]["p"]}
            for g in topo["objects"]["municipios"]["geometries"]}


@lru_cache(maxsize=1)
def _indice_nome_uf() -> Dict[Tuple[str, str], str]:
    return {(normalizar_cidade(m["nome"]), m["uf"]): cod for cod, m in carregar_municipios_ibge().items()}


def resolver_codigo_ibge(cidade: str, uf: str) -> Optional[str]:
    """Código IBGE de um município escrito de qualquer forma ('São Paulo', 'SAO PAULO', 'Açu'/'RN'...)."""
    uf_n = normalizar_uf_segura(uf)
    return _indice_nome_uf().get((canonizar_municipio(normalizar_cidade(str(cidade or "")), uf_n), uf_n))


# ---------------------------------------------------------------- alertas mês a mês
def calcular_alertas_por_mes(df: pd.DataFrame, cidades_risco: Set[Tuple[str, str]]) -> Dict[str, Any]:
    """
    A partir do agregado consolidado (cidade_norm, uf_norm, mes, totais),
    devolve matrizes município x mês: volume, pet, residencial, média
    histórica, razão sobre a média e bits de alerta. Para o mês i o histórico
    são os meses 0..i-1 — no último mês o resultado coincide com o radar.
    """
    meses = sorted(df["mes"].dropna().unique())
    chaves = df[["cidade_norm", "uf_norm"]].drop_duplicates().sort_values(["uf_norm", "cidade_norm"])
    idx = pd.MultiIndex.from_frame(chaves)

    def matriz(col):
        return (df.pivot_table(index=["cidade_norm", "uf_norm"], columns="mes", values=col, aggfunc="sum")
                  .reindex(index=idx, columns=meses).to_numpy(dtype=float))

    tot, pet, res = matriz("total_mes"), matriz("pet_mes"), matriz("res_mes")
    n_mun, n_mes = tot.shape
    bits = np.zeros((n_mun, n_mes), dtype=np.int64)
    media = np.full((n_mun, n_mes), np.nan)
    razao = np.full((n_mun, n_mes), np.nan)
    watch = np.array([(c, u) in cidades_risco for c, u in idx], dtype=bool)

    with np.errstate(invalid="ignore", divide="ignore"):
        for i in range(1, n_mes):
            presente_h = ~np.isnan(tot[:, :i])
            meses_h = presente_h.sum(axis=1)
            def media_h(m):
                s = np.where(presente_h, m[:, :i], 0).sum(axis=1)
                return np.where(meses_h > 0, s / np.maximum(meses_h, 1), 0.0)   # fillna(0) do radar
            m_tot, m_pet, m_res = media_h(tot), media_h(pet), media_h(res)
            t, p, r = tot[:, i], pet[:, i], res[:, i]
            agora = ~np.isnan(t)
            t0, p0, r0 = np.nan_to_num(t), np.nan_to_num(p), np.nan_to_num(r)
            rz = np.round(t0 / np.where(m_tot == 0, 0.4, m_tot), 1)
            rz_p = np.round(p0 / np.where(m_pet == 0, 0.4, m_pet), 1)
            rz_r = np.round(r0 / np.where(m_res == 0, 0.4, m_res), 1)

            b = np.zeros(n_mun, dtype=np.int64)
            vol = t0 >= 15
            b |= np.where(vol & (meses_h >= 3) & ((rz >= 3.0) | ((m_tot <= 3.0) & (t0 >= 15))), 1, 0)
            b |= np.where(vol & (meses_h < 3), 16, 0)
            b |= np.where((p0 >= 6) & ((rz_p >= 3.0) | (m_pet <= 1.0)), 2, 0)
            b |= np.where((r0 >= 12) & ((rz_r >= 2.5) | (m_res <= 2.0)), 4, 0)
            b |= np.where(watch & (t0 >= 5), 8, 0)
            bits[:, i] = np.where(agora, b, 0)
            media[:, i] = np.where(agora, m_tot, np.nan)
            razao[:, i] = np.where(agora, rz, np.nan)

    return {"meses": meses, "chaves": list(idx), "volume": tot, "pet": pet, "res": res,
            "media": media, "razao": razao, "bits": bits}


def rotulos_alerta(bits: int) -> List[str]:
    return [rotulo for bit, rotulo in ALERTAS if bits & bit]


def _arr(v):
    """Lista JSON compacta: NaN -> None, inteiros sem casas decimais."""
    out = []
    for x in v:
        if x is None or (isinstance(x, float) and np.isnan(x)):
            out.append(None)
        elif float(x).is_integer():
            out.append(int(x))
        else:
            out.append(round(float(x), 2))
    return out


# ---------------------------------------------------------------- payloads do mapa
@lru_cache(maxsize=4)
def _dados_mapa_radar_cached(versao_dados: Any, versao_cidades_set: frozenset) -> Dict[str, Any]:
    df = _agregado_consolidado(versao_dados)
    if df.empty:
        return {"payload": None, "indice": {}, "auditoria": [], "por_mes": {}}
    calc = calcular_alertas_por_mes(df, set(versao_cidades_set))
    meses = calc["meses"]

    # Grafias e nome de exibição por chave normalizada.
    info = (df.drop_duplicates(["cidade_norm", "uf_norm"])
              .set_index(["cidade_norm", "uf_norm"])[["cidade", "uf", "grafias_banco"]].to_dict("index"))

    municipios, indice, auditoria = {}, {}, []
    for k, (cid, uf) in enumerate(calc["chaves"]):
        cod = _indice_nome_uf().get((cid, uf))
        vol = calc["volume"][k]
        if cod is None:
            auditoria.append({"cidade": info[(cid, uf)]["cidade"], "uf": uf, "total": int(np.nansum(vol)),
                              "grafias": list(info[(cid, uf)]["grafias_banco"])})
            continue
        municipios[cod] = {"v": _arr(vol), "a": [int(x) for x in calc["bits"][k]],
                           "m": _arr(calc["media"][k]), "r": _arr(calc["razao"][k]),
                           "pet": _arr(calc["pet"][k]), "res": _arr(calc["res"][k])}
        indice[cod] = {"cidade_norm": cid, "uf": uf, "cidade_banco": info[(cid, uf)]["cidade"],
                       "grafias": info[(cid, uf)]["grafias_banco"]}

    total = float(np.nansum(calc["volume"]))
    nao_mapeado = sum(a["total"] for a in auditoria)
    auditoria.sort(key=lambda a: -a["total"])
    primeiro_com_historico = min(3, max(len(meses) - 1, 0))
    payload = {
        "modo": "radar",
        "meses": meses,
        "mes_inicial": primeiro_com_historico,
        "municipios": municipios,
        "alertas": [{"bit": b, "rotulo": r} for b, r in ALERTAS],
        "pct_mapeado": (1 - nao_mapeado / total) if total else 1.0,
        "fonte": FONTE_MALHA,
    }
    return {"payload": payload, "indice": indice, "auditoria": auditoria}


def montar_dados_mapa_radar() -> Dict[str, Any]:
    """
    {"payload": dict para o iframe do mapa, "indice": {código: grafias/nome no
    banco} para o funil, "auditoria": municípios do banco sem código IBGE}.
    Cacheado por versão das Criações e da watchlist.
    """
    return _dados_mapa_radar_cached(obter_versao_criacoes(), frozenset(obter_cidades_risco_set()))


def resumo_do_mes(dados: Dict[str, Any], i_mes: int) -> Dict[str, Any]:
    """KPIs, ranking de alertas, novos/saíram e composição de um mês (colunas laterais)."""
    payload = dados.get("payload") or {}
    mun, ibge = payload.get("municipios", {}), carregar_municipios_ibge()
    linhas, novos, sairam = [], [], []
    vol_total = pet_total = res_total = vol_alerta = 0
    for cod, d in mun.items():
        v = d["v"][i_mes] or 0
        vol_total += v
        pet_total += d["pet"][i_mes] or 0
        res_total += d["res"][i_mes] or 0
        bits, antes = d["a"][i_mes], (d["a"][i_mes - 1] if i_mes > 0 else 0)
        if bits:
            vol_alerta += v
            linhas.append({"codigo": cod, "nome": ibge[cod]["nome"], "uf": ibge[cod]["uf"], "volume": v,
                           "media": d["m"][i_mes], "razao": d["r"][i_mes], "bits": bits, "alertas": rotulos_alerta(bits)})
            if not antes:
                novos.append(cod)
        elif antes:
            sairam.append({"codigo": cod, "nome": ibge[cod]["nome"], "uf": ibge[cod]["uf"]})
    return {
        "mes": payload["meses"][i_mes] if payload else "",
        "em_alerta": linhas, "novos": novos, "sairam": sairam,
        "kpis": {"municipios_em_alerta": len(linhas), "acionamentos": int(vol_total),
                 "pct_em_alerta": (vol_alerta / vol_total) if vol_total else 0.0,
                 "watchlist": sum(1 for l in linhas if l["bits"] & 8),
                 "pet": int(pet_total), "residencial": int(res_total)},
        "serie_nacional": [int(sum((d["v"][k] or 0) for d in mun.values())) for k in range(len(payload.get("meses", [])))],
    }


# ---------------------------------------------------------------- mapa do caso
# Deslocamento impossível: a base só tem a DATA de cada assistência (sem hora), então a
# janela é em dias. Mesmo dia em cidades a >= KM_MESMO_DIA, ou até JANELA_DIAS depois a
# >= KM_JANELA, é fisicamente implausível para a mesma placa / o mesmo telefone.
JANELA_DIAS = 1
KM_MESMO_DIA = 300
KM_JANELA = 600


def _haversine_km(a, b) -> float:
    import math
    (la1, lo1), (la2, lo2) = a, b
    p1, p2 = math.radians(la1), math.radians(la2)
    dp, dl = p2 - p1, math.radians(lo2 - lo1)
    h = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * 6371.0 * math.asin(math.sqrt(h))


@lru_cache(maxsize=8192)
def _coord_municipio(cod: str):
    m = carregar_municipios_ibge().get(cod)
    if not m:
        return None
    from utils import obter_coordenadas
    lat, lon = obter_coordenadas(m["nome"], m["uf"])
    return (lat, lon) if lat is not None and lon is not None else None


def detectar_deslocamentos_impossiveis(df: pd.DataFrame, codigos: List[Optional[str]], janela_dias: int = JANELA_DIAS,
                                       km_mesmo_dia: float = KM_MESMO_DIA, km_janela: float = KM_JANELA) -> List[Dict[str, Any]]:
    """
    Mesma placa ou mesmo telefone acionado em municípios distantes num intervalo curto.
    `codigos`: código IBGE de cada linha do df (mesma ordem). Devolve uma lista de pares
    de municípios {de, para, km, n, ocorrencias: [{tipo, valor, data_de, data_para}]},
    mais distantes primeiro. As datas inválidas são ignoradas.
    """
    if df is None or df.empty or "data" not in df.columns:
        return []
    datas = pd.to_datetime(df["data"], errors="coerce")
    pares: Dict[tuple, Dict[str, Any]] = {}
    for col, tipo in (("placa", "PLACA"), ("telefone", "TEL")):
        if col not in df.columns:
            continue
        por_entidade: Dict[str, set] = {}
        for valor, dt, cod in zip(df[col], datas, codigos):
            v = str(valor or "").strip()
            if not v or v.lower() in ("nan", "none") or cod is None or pd.isna(dt):
                continue
            por_entidade.setdefault(v, set()).add((dt.normalize(), cod))
        for valor, ocorr in por_entidade.items():
            if len({c for _, c in ocorr}) < 2:
                continue
            lista = sorted(ocorr)
            for i, (d1, c1) in enumerate(lista):
                for d2, c2 in lista[i + 1:]:
                    dias = (d2 - d1).days
                    if dias > janela_dias:
                        break
                    if c1 == c2:
                        continue
                    a, b = _coord_municipio(c1), _coord_municipio(c2)
                    if not a or not b:
                        continue
                    km = _haversine_km(a, b)
                    if km < (km_mesmo_dia if dias == 0 else km_janela):
                        continue
                    chave = tuple(sorted((c1, c2)))
                    par = pares.setdefault(chave, {"de": c1, "para": c2, "km": round(km), "ocorrencias": [], "_vistos": set()})
                    marca = (tipo, valor, d1, d2)
                    if marca in par["_vistos"]:
                        continue
                    par["_vistos"].add(marca)
                    par["ocorrencias"].append({"tipo": tipo, "valor": valor, "data_de": d1.strftime("%Y-%m-%d"),
                                               "data_para": d2.strftime("%Y-%m-%d"), "dias": dias})
    saida = []
    for par in pares.values():
        par.pop("_vistos")
        par["ocorrencias"].sort(key=lambda o: (o["data_para"], o["tipo"], o["valor"]))
        par["n"] = len(par["ocorrencias"])
        saida.append(par)
    saida.sort(key=lambda p: (-p["n"], -p["km"]))
    return saida


def montar_payload_mapa_caso(df_dados: pd.DataFrame) -> Dict[str, Any]:
    """
    Mapa da mesa: acionamentos do caso por município, mês a mês.
      v[i] = acumulado até o mês i (o mapa "cresce" ao reproduzir a linha do tempo);
      n[i] = acionamentos no mês i (o mapa destaca quem esteve ativo no mês).
    Sem datas válidas, um único "mês" (total). Inclui os deslocamentos impossíveis
    (mesma placa/telefone em municípios distantes em pouco tempo), com o mês em que ocorrem.
    """
    municipios: Dict[str, Dict[str, Any]] = {}
    auditoria: List[Dict[str, Any]] = []
    deslocamentos: List[Dict[str, Any]] = []
    meses = ["caso"]
    if df_dados is not None and not df_dados.empty and {"cidade", "uf"} <= set(df_dados.columns):
        codigos = [resolver_codigo_ibge(c, u) for c, u in zip(df_dados["cidade"], df_dados["uf"])]
        datas = pd.to_datetime(df_dados["data"], errors="coerce") if "data" in df_dados.columns else pd.Series([pd.NaT] * len(df_dados))
        mes_txt = datas.dt.strftime("%Y-%m")
        validos = sorted(m for m in mes_txt.dropna().unique())
        meses = validos or ["caso"]
        idx_mes = {m: i for i, m in enumerate(meses)}
        n_mes = len(meses)
        sem_cod: Dict[tuple, int] = {}
        for cod, mes, cid, uf in zip(codigos, mes_txt, df_dados["cidade"], df_dados["uf"]):
            if cod is None:
                sem_cod[(str(cid), str(uf))] = sem_cod.get((str(cid), str(uf)), 0) + 1
                continue
            reg = municipios.setdefault(cod, {"n": [0] * n_mes})
            # Linha sem data válida entra no último mês (conta no total, não na animação).
            reg["n"][idx_mes.get(mes, n_mes - 1) if isinstance(mes, str) else n_mes - 1] += 1
        for reg in municipios.values():
            acc, v = 0, []
            for x in reg["n"]:
                acc += x
                v.append(acc)
            reg.update({"v": v, "a": [0] * n_mes, "m": [None] * n_mes, "r": [None] * n_mes})
        auditoria = [{"cidade": c, "uf": u, "total": n} for (c, u), n in sorted(sem_cod.items(), key=lambda x: -x[1])]
        for par in detectar_deslocamentos_impossiveis(df_dados, codigos):
            ult = max(o["data_para"][:7] for o in par["ocorrencias"])
            primeiro = min(o["data_para"][:7] for o in par["ocorrencias"])
            par["mes"] = idx_mes.get(primeiro, 0)
            par["mes_ultimo"] = idx_mes.get(ult, 0)
            deslocamentos.append(par)
    total = sum(a["total"] for a in auditoria) + sum(d["v"][-1] for d in municipios.values())
    return {"modo": "caso", "meses": meses, "mes_inicial": 0, "municipios": municipios,
            "alertas": [{"bit": b, "rotulo": r} for b, r in ALERTAS], "auditoria": auditoria,
            "deslocamentos": deslocamentos,
            "criterio_deslocamento": {"janela_dias": JANELA_DIAS, "km_mesmo_dia": KM_MESMO_DIA, "km_janela": KM_JANELA},
            "pct_mapeado": (1 - sum(a["total"] for a in auditoria) / total) if total else 1.0,
            "fonte": FONTE_MALHA}
