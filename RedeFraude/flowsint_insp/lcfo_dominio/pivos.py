"""
Pivôs de enrich, execução com log e proveniência, cache e receitas (item 7).

Um pivô tem contrato tipado (tipos de entrada -> tipos que pode criar), parâmetros e uma
estimativa. O executor grava cada execução em sketch_execucoes (estado, tempo, contagens,
cache_hit, a "impressão" da base usada) e cada nó/vínculo criado leva o id_execucao — é o
que permite "mostrar só o que este pivô trouxe" e "desfazer esta execução".

Sem NiceGUI: a interface roda `executar` numa thread (run.io_bound) passando `progresso` e
`cancelado`; o pivô checa o cancelamento entre lotes (a resposta fica abaixo de 2 s).

Proteção de escala: estimativa antes de rodar, profundidade máxima 3, limite de nós
(padrão 500) e entidades "hub" (um telefone/placa em muitos municípios) sinalizadas e não
expandidas sem confirmação.
"""

from __future__ import annotations

import json
import statistics
import time
import uuid
from collections import OrderedDict
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, Iterable, List, Optional, Tuple

import pandas as pd

import database
from lcfo_dominio import sketch_repo as repo
from lcfo_dominio import tipos

PROFUNDIDADE_MAX = 3
LIMITE_NOS_PADRAO = 500
TETO_CONFIRMACAO = 500
HUB_MIN_MUNICIPIOS = 15          # um telefone/placa acionado em 15+ municípios é "hub"
HUB_FATOR_MEDIANA = 5
TIPOS_CODIGO = ("cpf", "telefone", "placa")
COLUNAS = "data, id_assistencia, titular, cpf, telefone, placa, servico, cidade, uf, empresa_cliente"
LOTE_SQL = 400

from lcfo_dominio.schema import garantir_schema  # noqa: E402,F401  (DDL em schema.py)


class Cancelado(Exception):
    pass


# ------------------------------------------------------------------ contrato
@dataclass(frozen=True)
class Param:
    nome: str
    rotulo: str
    tipo: str                      # int | tipos | data | bool
    padrao: Any = None
    minimo: Optional[int] = None
    maximo: Optional[int] = None


@dataclass
class Contexto:
    progresso: Callable[[int, int, str], None] = lambda a, t, m: None
    cancelado: Callable[[], bool] = lambda: False

    def checar(self):
        if self.cancelado():
            raise Cancelado()


@dataclass(frozen=True)
class Pivo:
    chave: str
    rotulo: str
    entrada: Tuple[str, ...]
    saidas: Tuple[str, ...]
    parametros: Tuple[Param, ...]
    custo: str                                              # leve | medio | pesado
    executar: Callable[[Contexto, List[dict], dict], dict]
    estimar: Callable[[List[dict], dict], dict]
    descricao: str = ""


CATALOGO: "OrderedDict[str, Pivo]" = OrderedDict()


def registrar(p: Pivo) -> Pivo:
    CATALOGO[p.chave] = p
    return p


def pivos_para(tipo: str) -> List[Pivo]:
    return [p for p in CATALOGO.values() if tipo in p.entrada]


def normalizar_parametros(pivo: Pivo, params: Optional[dict]) -> dict:
    """Padrões + limites (profundidade 1..3 etc.); chaves ordenadas para o cache."""
    params = dict(params or {})
    out = {}
    for p in pivo.parametros:
        v = params.get(p.nome, p.padrao)
        if p.tipo == "int":
            try:
                v = int(v)
            except (TypeError, ValueError):
                v = p.padrao
            if p.minimo is not None:
                v = max(p.minimo, v)
            if p.maximo is not None:
                v = min(p.maximo, v)
        elif p.tipo == "tipos":
            v = sorted(t for t in (v or p.padrao or []) if t in pivo.saidas) or list(p.padrao or pivo.saidas)
        elif p.tipo == "bool":
            v = bool(v)
        elif p.tipo == "data":
            v = (str(v).strip()[:10] or None) if v else None
        out[p.nome] = v
    return dict(sorted(out.items()))


def payload_catalogo() -> List[dict]:
    """Para o menu do nó no iframe: pivôs com entrada por tipo."""
    return [{"chave": p.chave, "rotulo": p.rotulo, "entrada": list(p.entrada), "saidas": list(p.saidas), "custo": p.custo,
             "tem_parametros": bool(p.parametros), "descricao": p.descricao} for p in CATALOGO.values()]


# ------------------------------------------------------------------ leitura da base
def versao_base() -> str:
    """Impressão (contagem, maior id) das tabelas que os pivôs leem: torna a execução reprodutível."""
    conn = database.get_db_connection()
    try:
        out = []
        for t in ("assistencias", "criacoes_diarias"):
            try:
                r = conn.execute(f"SELECT (SELECT MAX(id) FROM {t}), (SELECT MIN(id) FROM {t})").fetchone()
                out.append(f"{t}:{r[0] or 0}:{r[1] or 0}")
            except Exception:
                out.append(f"{t}:-")
        return "|".join(out)
    finally:
        conn.close()


def _ocorrencias(ctx: Contexto, coluna: str, valores: List[str], data_de=None, data_ate=None) -> pd.DataFrame:
    """Linhas de Blacklist + Criações em que `coluna` está em `valores` (em lotes, cancelável)."""
    if coluna not in ("cpf", "telefone", "placa", "id_assistencia"):
        raise ValueError(coluna)
    partes = []
    valores = list(dict.fromkeys(v for v in valores if v))
    conn = database.get_db_connection()
    try:
        for i in range(0, len(valores), LOTE_SQL):
            ctx.checar()
            lote = valores[i:i + LOTE_SQL]
            marc = ",".join("?" * len(lote))
            filtro, args = "", []
            if data_de:
                filtro += " AND data >= ?"; args.append(data_de)
            if data_ate:
                filtro += " AND data <= ?"; args.append(data_ate)
            sql = f"""SELECT {COLUNAS}, '{coluna}' AS col_match FROM criacoes_diarias WHERE {coluna} IN ({marc}){filtro}
                      UNION ALL
                      SELECT {COLUNAS}, '{coluna}' AS col_match FROM assistencias WHERE {coluna} IN ({marc}){filtro}"""
            partes.append(pd.read_sql_query(sql, conn, params=[*lote, *args, *lote, *args]))
    finally:
        conn.close()
    if not partes:
        return pd.DataFrame(columns=COLUNAS.replace(" ", "").split(",") + ["col_match"])
    df = pd.concat(partes, ignore_index=True).fillna("")
    for c in ("cpf", "telefone", "placa", "id_assistencia", "titular", "empresa_cliente", "cidade", "uf", "servico", "data"):
        df[c] = df[c].astype(str).str.strip().replace({"nan": "", "None": ""})
    # a mesma assistência pode estar nas duas tabelas
    return df.drop_duplicates(subset=["id_assistencia", "data", "cpf", "telefone", "placa"]).reset_index(drop=True)


def _municipios_por_valor(df: pd.DataFrame, coluna: str) -> Dict[str, int]:
    if df.empty:
        return {}
    return (df.assign(_m=df["cidade"].str.upper() + "/" + df["uf"].str.upper())
              .groupby(coluna)["_m"].nunique().to_dict())


def _eh_hub(n_mun: int, mediana: float) -> bool:
    return n_mun >= max(HUB_MIN_MUNICIPIOS, HUB_FATOR_MEDIANA * (mediana or 0))


# ------------------------------------------------------------------ construção de nós/arestas vis
def _no_codigo(tipo: str, valor: str, grau: int, pai: str, nome: str = "") -> dict:
    from enrich_engine import _montar_no_entidade           # mesmo nó do Enrich (cor, ícone, rótulo)
    return _montar_no_entidade(tipo, valor, grau, pai, nome=nome)


def _no_generico(tipo: str, valor: str, label: str, title: str, pai: str, extras: Optional[dict] = None) -> dict:
    t = tipos.obter(tipo)
    no = {"id": t.node_id(valor), "label": label, "title": title, "tipo": t.chave, "shape": t.forma, "image": t.icone_svg(),
          "size": 16, "degree": 1, "betweenness": 0.0, "community": 0, "valor": t.formatar(valor), "nome_titular": "",
          "orbit_ao_redor": pai, "color": {"background": t.cor[0], "border": t.cor[1]}, "font": dict(tipos.FONTE_NO)}
    no.update(extras or {})
    return no


def _aresta(de: str, para: str, titulo: str, n: int = 1, rotulo: str = "", situacao: str = "confirmado", estilo: str = "pivo") -> dict:
    e = {"id": f"enr_{de}__{para}", "from": de, "to": para, "title": titulo, "ocorrencias": n,
         "width": min(3.5, 1.0 + 0.3 * n), "color": {"color": "rgba(140,140,140,0.55)", "highlight": "#FF7300"}}
    if rotulo:
        e["label"] = rotulo
        e["font"] = {"color": "#EDEDED", "size": 10, "strokeWidth": 3, "strokeColor": "#121214"}
    if situacao == "hipotese":
        e["dashes"] = [6, 4]
        e["situacao"] = "hipotese"
    return e


# ------------------------------------------------------------------ pivô: vínculos diretos
def _exec_vinculos_diretos(ctx: Contexto, entradas: List[dict], p: dict) -> dict:
    tipos_saida = set(p["tipos_saida"])
    limite, prof = p["limite_nos"], p["profundidade"]
    fronteira = {(e["tipo"], tipos.valor_do_id(e["node_id"])) for e in entradas if e["tipo"] in TIPOS_CODIGO}
    visitados = {e["node_id"] for e in entradas}
    nos: Dict[str, dict] = {}
    arestas: Dict[str, dict] = {}
    hubs, truncado = [], False
    nomes: Dict[str, str] = {}
    for nivel in range(1, prof + 1):
        ctx.progresso(nivel - 1, prof, f"nível {nivel}: {len(fronteira)} entidade(s)")
        por_col: Dict[str, List[str]] = {}
        for tp, v in fronteira:
            por_col.setdefault(tp, []).append(v)
        dfs = [_ocorrencias(ctx, col, vals, p.get("data_de"), p.get("data_ate")) for col, vals in por_col.items()]
        df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
        if df.empty:
            break
        for r in df.itertuples(index=False):
            if r.cpf and r.titular:
                nomes.setdefault(r.cpf, r.titular)
        # hubs da fronteira: sinalizados; não expandem (a não ser que o analista confirme)
        bloqueados = set()
        for col in por_col:
            nmun = _municipios_por_valor(df[df["col_match"] == col], col)
            med = statistics.median(nmun.values()) if nmun else 0
            for v, n in nmun.items():
                nid = tipos.obter(col).node_id(v)
                if _eh_hub(n, med if len(nmun) > 2 else 0) and not p.get("expandir_hubs"):
                    hubs.append({"node_id": nid, "municipios": int(n)})
                    bloqueados.add((col, v))
        proxima = set()
        contagem: Dict[Tuple[str, str], Dict[str, Any]] = {}
        for r in df.itertuples(index=False):
            origem = (r.col_match, getattr(r, r.col_match))              # quem casou com a linha
            if origem in bloqueados:
                continue
            pai = tipos.obter(origem[0]).node_id(origem[1])
            for tp in TIPOS_CODIGO:
                v = getattr(r, tp)
                if not v or tp not in tipos_saida or (tp, v) == origem:
                    continue
                nid = tipos.obter(tp).node_id(v)
                c = contagem.setdefault((pai, nid), {"n": 0, "data": r.data, "ass": r.id_assistencia, "emp": set(), "tp": tp, "v": v})
                c["n"] += 1
                if r.empresa_cliente:
                    c["emp"].add(r.empresa_cliente.upper())
        for (pai, nid), c in contagem.items():
            if c["n"] < p["min_ocorrencias"]:
                continue
            if nid not in visitados and nid not in nos:
                if len(nos) >= limite:
                    truncado = True
                    continue
                nos[nid] = _no_codigo(c["tp"], c["v"], c["n"], pai, nome=nomes.get(c["v"], "") if c["tp"] == "cpf" else "")
                proxima.add((c["tp"], c["v"]))
            if nid in nos or nid in visitados:
                emp = ", ".join(sorted(c["emp"])[:3])
                arestas[f"enr_{pai}__{nid}"] = _aresta(pai, nid, f"Vínculo via pivô: {c['n']} ocorrência(s) em comum "
                                                       f"(última: {c['ass']} em {c['data']})" + (f"\nSeguradora(s): {emp}" if emp else ""), c["n"])
        visitados |= set(nos)
        fronteira = proxima
        if not fronteira or truncado:
            break
    ctx.progresso(prof, prof, "concluído")
    return {"nos": list(nos.values()), "arestas": list(arestas.values()), "hubs": hubs, "truncado": truncado}


def _estimar_vinculos_diretos(entradas: List[dict], p: dict) -> dict:
    ctx = Contexto()
    por_col: Dict[str, List[str]] = {}
    for e in entradas:
        if e["tipo"] in TIPOS_CODIGO:
            por_col.setdefault(e["tipo"], []).append(tipos.valor_do_id(e["node_id"]))
    vizinhos, hubs = set(), []
    for col, vals in por_col.items():
        df = _ocorrencias(ctx, col, vals, p.get("data_de"), p.get("data_ate"))
        for v, n in _municipios_por_valor(df, col).items():
            if _eh_hub(n, 0):
                hubs.append({"node_id": tipos.obter(col).node_id(v), "municipios": int(n)})
        for tp in TIPOS_CODIGO:
            if tp in p["tipos_saida"]:
                vizinhos |= {(tp, v) for v in df[tp] if v} - {(col, v) for v in vals}
    n1 = len(vizinhos)
    return {"itens": n1 if p["profundidade"] == 1 else n1 * (4 ** (p["profundidade"] - 1)), "exato": p["profundidade"] == 1, "hubs": hubs}


registrar(Pivo(
    chave="vinculos_diretos", rotulo="Vínculos diretos", entrada=("cpf", "telefone", "placa"), saidas=("cpf", "telefone", "placa"),
    parametros=(Param("profundidade", "Profundidade", "int", 1, 1, PROFUNDIDADE_MAX),
                Param("tipos_saida", "Trazer", "tipos", ["cpf", "telefone", "placa"]),
                Param("data_de", "De", "data"), Param("data_ate", "Até", "data"),
                Param("min_ocorrencias", "Mínimo de ocorrências", "int", 1, 1, 1000),
                Param("limite_nos", "Limite de nós", "int", LIMITE_NOS_PADRAO, 1, 5000),
                Param("expandir_hubs", "Expandir hubs", "bool", False)),
    custo="medio", executar=_exec_vinculos_diretos, estimar=_estimar_vinculos_diretos,
    descricao="CPFs, telefones e placas que aparecem nas mesmas assistências (o Enrich), até 3 saltos."))


# ------------------------------------------------------------------ pivô: acionamentos
def _no_assistencia(r, pai: str) -> dict:
    rot = f"{r.servico or 'Assistência'} · {r.data}"
    title = f"Assistência {r.id_assistencia}\n{r.data} · {r.servico}\n{r.cidade}/{r.uf}"
    return _no_generico("assistencia", r.id_assistencia, rot[:28], title, pai,
                        {"campos": {"data": r.data, "servico": r.servico, "cidade": r.cidade, "uf": r.uf}})


def _exec_acionamentos(ctx: Contexto, entradas: List[dict], p: dict) -> dict:
    nos, arestas, truncado = {}, {}, False
    for i, e in enumerate(entradas):
        ctx.progresso(i, len(entradas), f"acionamentos de {tipos.formatar_id(e['node_id'])}")
        df = _ocorrencias(ctx, e["tipo"], [tipos.valor_do_id(e["node_id"])], p.get("data_de"), p.get("data_ate"))
        for r in df.sort_values("data").itertuples(index=False):
            if not r.id_assistencia:
                continue
            nid = tipos.obter("assistencia").node_id(r.id_assistencia)
            if nid not in nos:
                if len(nos) >= p["limite_nos"]:
                    truncado = True
                    break
                nos[nid] = _no_assistencia(r, e["node_id"])
            arestas[f"enr_{e['node_id']}__{nid}"] = _aresta(e["node_id"], nid, f"Acionou em {r.data} ({r.cidade}/{r.uf})")
    ctx.progresso(len(entradas), len(entradas), "concluído")
    return {"nos": list(nos.values()), "arestas": list(arestas.values()), "hubs": [], "truncado": truncado}


def _estimar_acionamentos(entradas: List[dict], p: dict) -> dict:
    n = 0
    for e in entradas:
        n += len(_ocorrencias(Contexto(), e["tipo"], [tipos.valor_do_id(e["node_id"])], p.get("data_de"), p.get("data_ate")))
    return {"itens": n, "exato": True, "hubs": []}


registrar(Pivo(
    chave="acionamentos", rotulo="Acionamentos", entrada=("cpf", "telefone", "placa"), saidas=("assistencia",),
    parametros=(Param("data_de", "De", "data"), Param("data_ate", "Até", "data"),
                Param("limite_nos", "Limite de nós", "int", LIMITE_NOS_PADRAO, 1, 5000)),
    custo="leve", executar=_exec_acionamentos, estimar=_estimar_acionamentos,
    descricao="Cada assistência da entidade vira um nó (data, serviço, cidade)."))


# ------------------------------------------------------------------ pivô: entidades da praça
def _sql_praca(cidade: str, uf: str, limite: int, data_de=None, data_ate=None) -> pd.DataFrame:
    conn = database.get_db_connection()
    try:
        filtro, args = "", []
        if data_de:
            filtro += " AND data >= ?"; args.append(data_de)
        if data_ate:
            filtro += " AND data <= ?"; args.append(data_ate)
        sql = f"""SELECT {COLUNAS} FROM criacoes_diarias WHERE UPPER(cidade) = ? AND UPPER(uf) = ?{filtro}
                  UNION ALL SELECT {COLUNAS} FROM assistencias WHERE UPPER(cidade) = ? AND UPPER(uf) = ?{filtro} LIMIT ?"""
        return pd.read_sql_query(sql, conn, params=[cidade, uf, *args, cidade, uf, *args, limite * 4]).fillna("")
    finally:
        conn.close()


def _exec_praca(ctx: Contexto, entradas: List[dict], p: dict) -> dict:
    nos, arestas, avisos = {}, {}, []
    for i, e in enumerate(entradas):
        ctx.checar()
        campos = e.get("campos") or {}
        cidade, uf = (campos.get("cidade") or "").strip().upper(), (campos.get("uf") or "").strip().upper()
        if not cidade or not uf:
            avisos.append(f"{tipos.formatar_id(e['node_id'])}: sem cidade/UF — preencha os campos do endereço.")
            continue
        ctx.progresso(i, len(entradas), f"praça {cidade}/{uf}")
        df = _sql_praca(cidade, uf, p["limite_nos"], p.get("data_de"), p.get("data_ate"))
        cont: Dict[str, Dict[str, Any]] = {}
        for r in df.itertuples(index=False):
            for tp in TIPOS_CODIGO:
                v = str(getattr(r, tp) or "").strip()
                if v and tp in p["tipos_saida"]:
                    c = cont.setdefault(tipos.obter(tp).node_id(v), {"tp": tp, "v": v, "n": 0, "nome": ""})
                    c["n"] += 1
                    if tp == "cpf" and r.titular:
                        c["nome"] = c["nome"] or str(r.titular).strip()
        for nid, c in sorted(cont.items(), key=lambda x: -x[1]["n"])[:p["limite_nos"]]:
            if c["n"] < p["min_ocorrencias"]:
                continue
            nos.setdefault(nid, _no_codigo(c["tp"], c["v"], c["n"], e["node_id"], nome=c["nome"]))
            arestas[f"enr_{e['node_id']}__{nid}"] = _aresta(e["node_id"], nid, f"Mesma praça ({cidade}/{uf}): {c['n']} acionamento(s)", c["n"])
    return {"nos": list(nos.values()), "arestas": list(arestas.values()), "hubs": [], "truncado": False, "avisos": avisos}


def _estimar_praca(entradas: List[dict], p: dict) -> dict:
    return {"itens": min(p["limite_nos"], LIMITE_NOS_PADRAO) * len(entradas), "exato": False, "hubs": []}


registrar(Pivo(
    chave="entidades_da_praca", rotulo="Entidades da praça", entrada=("endereco",), saidas=("cpf", "telefone", "placa"),
    parametros=(Param("tipos_saida", "Trazer", "tipos", ["cpf", "telefone", "placa"]),
                Param("min_ocorrencias", "Mínimo de acionamentos", "int", 2, 1, 1000),
                Param("data_de", "De", "data"), Param("data_ate", "Até", "data"),
                Param("limite_nos", "Limite de nós", "int", 100, 1, 2000)),
    custo="pesado", executar=_exec_praca, estimar=_estimar_praca,
    descricao="Quem mais foi acionado na mesma cidade do endereço (precisa de cidade e UF)."))


# ------------------------------------------------------------------ pivô: deslocamentos
def _exec_deslocamentos(ctx: Contexto, entradas: List[dict], p: dict) -> dict:
    import territorio_engine as te
    nos, arestas = {}, {}
    for i, e in enumerate(entradas):
        ctx.progresso(i, len(entradas), f"deslocamentos de {tipos.formatar_id(e['node_id'])}")
        df = _ocorrencias(ctx, e["tipo"], [tipos.valor_do_id(e["node_id"])], p.get("data_de"), p.get("data_ate"))
        if df.empty:
            continue
        cods = [te.resolver_codigo_ibge(c, u) for c, u in zip(df["cidade"], df["uf"])]
        df = df.assign(cod_ibge=cods)
        pares = te.detectar_deslocamentos_impossiveis(df, cods)
        for par in pares:
            for oc in par["ocorrencias"]:
                ctx.checar()
                a = df[(df["cod_ibge"] == par["de"]) & (df["data"].str[:10] == oc["data_de"])].head(1)
                b = df[(df["cod_ibge"] == par["para"]) & (df["data"].str[:10] == oc["data_para"])].head(1)
                if a.empty or b.empty:
                    a = df[(df["cod_ibge"] == par["para"]) & (df["data"].str[:10] == oc["data_de"])].head(1)
                    b = df[(df["cod_ibge"] == par["de"]) & (df["data"].str[:10] == oc["data_para"])].head(1)
                if a.empty or b.empty:
                    continue
                ra, rb = next(a.itertuples(index=False)), next(b.itertuples(index=False))
                na, nb = _no_assistencia(ra, e["node_id"]), _no_assistencia(rb, e["node_id"])
                nos.setdefault(na["id"], na)
                nos.setdefault(nb["id"], nb)
                arestas.setdefault(f"enr_{e['node_id']}__{na['id']}", _aresta(e["node_id"], na["id"], f"Acionou em {ra.data}"))
                arestas.setdefault(f"enr_{e['node_id']}__{nb['id']}", _aresta(e["node_id"], nb["id"], f"Acionou em {rb.data}"))
                arestas[f"enr_{na['id']}__{nb['id']}"] = _aresta(
                    na["id"], nb["id"], f"Deslocamento impossível: {par['km']} km em {oc['dias']} dia(s)\n"
                    f"{ra.cidade}/{ra.uf} → {rb.cidade}/{rb.uf}", 1, rotulo=f"{par['km']} km", situacao="hipotese")
    return {"nos": list(nos.values()), "arestas": list(arestas.values()), "hubs": [], "truncado": False}


def _estimar_deslocamentos(entradas: List[dict], p: dict) -> dict:
    return _estimar_acionamentos(entradas, p)


registrar(Pivo(
    chave="deslocamentos", rotulo="Deslocamentos impossíveis", entrada=("placa", "telefone"), saidas=("assistencia",),
    parametros=(Param("data_de", "De", "data"), Param("data_ate", "Até", "data")),
    custo="medio", executar=_exec_deslocamentos, estimar=_estimar_deslocamentos,
    descricao="Mesma placa/telefone em municípios distantes em pouco tempo: vínculos hipótese entre as assistências."))


# ------------------------------------------------------------------ executor, log e cache
_CACHE: "OrderedDict[str, dict]" = OrderedDict()
_CACHE_MAX = 64


def _chave_cache(pivo: str, params: dict, entradas: List[dict], vb: str) -> str:
    return json.dumps([pivo, params, sorted(e["node_id"] for e in entradas), vb], sort_keys=True, ensure_ascii=False)


def limpar_cache() -> None:
    _CACHE.clear()


def estimar(pivo: str, entradas: List[dict], params: Optional[dict] = None) -> dict:
    pv = CATALOGO[pivo]
    p = normalizar_parametros(pv, params)
    r = pv.estimar(entradas, p)
    r["acima_do_teto"] = r["itens"] > TETO_CONFIRMACAO
    return r


def calcular(pivo: str, entradas: List[dict], params: Optional[dict] = None, ctx: Optional[Contexto] = None) -> Tuple[dict, bool]:
    """Roda o pivô sem gravar (com cache). Devolve (resultado, cache_hit)."""
    pv = CATALOGO[pivo]
    p = normalizar_parametros(pv, params)
    chave = _chave_cache(pivo, p, entradas, versao_base())
    if chave in _CACHE:
        _CACHE.move_to_end(chave)
        return _CACHE[chave], True
    res = pv.executar(ctx or Contexto(), entradas, p)
    _CACHE[chave] = res
    while len(_CACHE) > _CACHE_MAX:
        _CACHE.popitem(last=False)
    return res, False


def _atualizar_execucao(id_exec: str, **campos) -> None:
    conn = database.get_db_connection()
    try:
        conn.execute(f"UPDATE sketch_execucoes SET {', '.join(f'{k} = ?' for k in campos)} WHERE id = ?", (*campos.values(), id_exec))
        conn.commit()
    finally:
        conn.close()


def iniciar_execucao(id_sketch: str, pivo: str, entradas: List[dict], params: Optional[dict] = None,
                     id_receita_exec: Optional[str] = None, quem: Optional[str] = None) -> str:
    pv = CATALOGO[pivo]
    id_exec = f"ex_{uuid.uuid4().hex[:12]}"
    conn = database.get_db_connection()
    try:
        conn.execute("""INSERT INTO sketch_execucoes (id, id_sketch, pivo, parametros_json, entradas_json, estado, iniciado_em,
                        versao_base, id_receita_exec, usuario) VALUES (?,?,?,?,?,'pendente',?,?,?,?)""",
                     (id_exec, id_sketch, pivo, json.dumps(normalizar_parametros(pv, params), ensure_ascii=False),
                      json.dumps([e["node_id"] for e in entradas], ensure_ascii=False), repo.agora(), versao_base(),
                      id_receita_exec, quem if quem is not None else repo.usuario_atual()))
        conn.commit()
    finally:
        conn.close()
    return id_exec


def executar(id_sketch: str, pivo: str, entradas: List[dict], params: Optional[dict] = None, *,
             progresso: Optional[Callable[[int, int, str], None]] = None, cancelado: Optional[Callable[[], bool]] = None,
             quem: Optional[str] = None, id_receita_exec: Optional[str] = None, id_exec: Optional[str] = None) -> dict:
    """
    Roda o pivô e grava o resultado no sketch com proveniência. Reexecutar com os mesmos
    parâmetros não duplica nada (a identidade é o node_id): dá zero nós novos e cache_hit.
    """
    id_exec = id_exec or iniciar_execucao(id_sketch, pivo, entradas, params, id_receita_exec, quem)
    t0 = time.perf_counter()
    _atualizar_execucao(id_exec, estado="executando")
    ctx = Contexto(progresso=progresso or (lambda a, t, m: None), cancelado=cancelado or (lambda: False))
    try:
        res, hit = calcular(pivo, entradas, params, ctx)
        ctx.checar()
        r = repo.adicionar(id_sketch, res["nos"], res["arestas"], origem="pivo", id_execucao=id_exec, quem=quem, acao="pivo")
        ms = int((time.perf_counter() - t0) * 1000)
        avisos = list(res.get("avisos") or [])
        if res.get("truncado"):
            avisos.append("Limite de nós atingido: o resultado foi cortado.")
        _atualizar_execucao(id_exec, estado="concluida", concluido_em=repo.agora(), ms=ms, n_nos_novos=len(r["nos_novos"]),
                            n_vinculos_novos=len(r["vinculos_novos"]), n_existentes=len(r["existentes"]), cache_hit=int(hit),
                            avisos_json=json.dumps({"avisos": avisos, "hubs": res.get("hubs") or []}, ensure_ascii=False))
        return {"id": id_exec, "pivo": pivo, "estado": "concluida", "nos_novos": r["nos_novos"], "vinculos_novos": r["vinculos_novos"],
                "reativados": r["reativados"], "existentes": r["existentes"], "cache_hit": hit, "ms": ms,
                "hubs": res.get("hubs") or [], "avisos": avisos, "saida": [n["id"] for n in res["nos"]]}
    except Cancelado:
        _atualizar_execucao(id_exec, estado="cancelada", concluido_em=repo.agora(), ms=int((time.perf_counter() - t0) * 1000))
        return {"id": id_exec, "pivo": pivo, "estado": "cancelada", "nos_novos": [], "vinculos_novos": [], "reativados": [], "existentes": [],
                "cache_hit": False, "hubs": [], "avisos": [], "saida": []}
    except Exception as e:
        _atualizar_execucao(id_exec, estado="falha", concluido_em=repo.agora(), erro=str(e)[:500],
                            ms=int((time.perf_counter() - t0) * 1000))
        raise


def listar_execucoes(id_sketch: str, limite: int = 50) -> List[dict]:
    conn = database.get_db_connection()
    try:
        return [dict(r) for r in conn.execute(
            "SELECT * FROM sketch_execucoes WHERE id_sketch = ? ORDER BY iniciado_em DESC, rowid DESC LIMIT ?", (id_sketch, limite))]
    finally:
        conn.close()


def obter_execucao(id_exec: str) -> Optional[dict]:
    conn = database.get_db_connection()
    try:
        r = conn.execute("SELECT * FROM sketch_execucoes WHERE id = ?", (id_exec,)).fetchone()
        return dict(r) if r else None
    finally:
        conn.close()


def desfazer_execucao(id_sketch: str, id_exec: str, quem: Optional[str] = None) -> dict:
    """Remove (para a lixeira) só o que esta execução criou."""
    ex = obter_execucao(id_exec)
    rot = CATALOGO[ex["pivo"]].rotulo if ex and ex["pivo"] in CATALOGO else id_exec
    return repo.remover(id_sketch, motivo=f"Desfazer execução: {rot}", quem=quem, id_execucao=id_exec)


# ------------------------------------------------------------------ receitas
RECEITAS_PADRAO = [
    {"id": "rec_rede_titular", "nome": "Rede do titular", "entrada": "cpf", "padrao": True,
     "passos": [{"pivo": "vinculos_diretos", "params": {"tipos_saida": ["telefone"], "profundidade": 1}},
                {"pivo": "vinculos_diretos", "params": {"tipos_saida": ["placa"], "profundidade": 1}}]},
    {"id": "rec_nucleo_telefone", "nome": "Núcleo do telefone", "entrada": "telefone", "padrao": True,
     "passos": [{"pivo": "vinculos_diretos", "params": {"tipos_saida": ["cpf"], "profundidade": 1}},
                {"pivo": "vinculos_diretos", "params": {"tipos_saida": ["placa"], "profundidade": 1}}]},
]


def listar_receitas() -> List[dict]:
    conn = database.get_db_connection()
    try:
        salvas = [{"id": r["id"], "nome": r["nome"], "entrada": r["entrada"], "passos": json.loads(r["passos_json"]), "padrao": False}
                  for r in conn.execute("SELECT * FROM receitas ORDER BY criado_em")]
    finally:
        conn.close()
    return [dict(r) for r in RECEITAS_PADRAO] + salvas


def obter_receita(id_receita: str) -> Optional[dict]:
    return next((r for r in listar_receitas() if r["id"] == id_receita), None)


def validar_receita(entrada: str, passos: List[dict]) -> List[str]:
    """Cada passo precisa aceitar algum tipo produzido pelo anterior (o 1º, o tipo de entrada)."""
    erros, disponiveis = [], {entrada}
    for i, ps in enumerate(passos, 1):
        pv = CATALOGO.get(ps.get("pivo"))
        if not pv:
            erros.append(f"passo {i}: pivô desconhecido")
            continue
        if not disponiveis & set(pv.entrada):
            erros.append(f"passo {i}: {pv.rotulo} não aceita {', '.join(sorted(disponiveis))}")
        p = normalizar_parametros(pv, ps.get("params"))
        disponiveis = set(p.get("tipos_saida") or pv.saidas)
    return erros


def salvar_receita(nome: str, entrada: str, passos: List[dict], quem: Optional[str] = None) -> str:
    erros = validar_receita(entrada, passos)
    if erros:
        raise ValueError("; ".join(erros))
    rid = f"rec_{uuid.uuid4().hex[:10]}"
    conn = database.get_db_connection()
    try:
        conn.execute("INSERT INTO receitas (id, nome, entrada, passos_json, criado_em, criado_por) VALUES (?,?,?,?,?,?)",
                     (rid, nome.strip() or "Receita", entrada, json.dumps(passos, ensure_ascii=False), repo.agora(),
                      quem if quem is not None else repo.usuario_atual()))
        conn.commit()
    finally:
        conn.close()
    return rid


def _entradas_do_resultado(res: dict, aceitos: Iterable[str]) -> List[dict]:
    aceitos = set(aceitos)
    return [{"node_id": n["id"], "tipo": n["tipo"], "campos": n.get("campos") or {}} for n in res["nos"] if n.get("tipo") in aceitos]


def calcular_receita(receita: dict, entradas: List[dict]) -> List[dict]:
    """Ordem e estimativa por passo, SEM gravar (a saída de cada passo alimenta o seguinte)."""
    out, atuais = [], [e for e in entradas if e["tipo"] == receita["entrada"]]
    for i, ps in enumerate(receita["passos"], 1):
        pv = CATALOGO[ps["pivo"]]
        atuais = [e for e in atuais if e["tipo"] in pv.entrada]
        if not atuais:
            out.append({"passo": i, "pivo": pv.rotulo, "entradas": 0, "itens": 0})
            continue
        res, _ = calcular(ps["pivo"], atuais, ps.get("params"))
        out.append({"passo": i, "pivo": pv.rotulo, "entradas": len(atuais), "itens": len(res["nos"]), "hubs": len(res.get("hubs") or [])})
        atuais = _entradas_do_resultado(res, set().union(*[set(p.entrada) for p in CATALOGO.values()]))
    return out


def executar_receita(id_sketch: str, receita: dict, entradas: List[dict], *, progresso=None, cancelado=None,
                     quem: Optional[str] = None) -> dict:
    """Uma linha em sketch_execucoes por passo, ligadas pelo mesmo id_receita_exec."""
    id_rec = f"rx_{uuid.uuid4().hex[:10]}"
    passos, atuais = [], [e for e in entradas if e["tipo"] == receita["entrada"]]
    for i, ps in enumerate(receita["passos"], 1):
        pv = CATALOGO[ps["pivo"]]
        atuais = [e for e in atuais if e["tipo"] in pv.entrada]
        if not atuais or (cancelado and cancelado()):
            break
        r = executar(id_sketch, ps["pivo"], atuais, ps.get("params"), progresso=progresso, cancelado=cancelado,
                     quem=quem, id_receita_exec=id_rec)
        passos.append(r)
        if r["estado"] != "concluida":
            break
        res, _ = calcular(ps["pivo"], atuais, ps.get("params"))       # do cache: a saída completa do passo
        atuais = _entradas_do_resultado(res, set().union(*[set(p.entrada) for p in CATALOGO.values()]))
    return {"id_receita_exec": id_rec, "passos": passos}
