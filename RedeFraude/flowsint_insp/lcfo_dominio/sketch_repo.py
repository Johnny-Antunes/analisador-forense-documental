"""
Persistência do sketch em tabelas (item 5): o que o analista cria ou altera num sketch
vive em linhas próprias, com origem, autor, data e remoção lógica.

Antes: um blob JSON (caso_sketches.nos_extras_json) regravado por inteiro a cada mudança,
sem dizer quem criou nem quando, e a memória da sessão (st.nos_enriquecidos) como fonte
da verdade. Agora a memória é só cache do overlay lido daqui.

Regras:
  - toda escrita abre uma transação (BEGIN IMMEDIATE, WAL, busy_timeout), grava um evento
    em sketch_eventos NA MESMA transação e incrementa caso_sketches.versao;
  - nada é apagado fisicamente: remover grava removido_em/por/motivo/lote e restaurar
    devolve o lote inteiro (cadeia de custódia);
  - dado oficial (cluster da blacklist) não entra aqui como entidade editável: só como
    anotação por cima (origem 'anotacao_oficial', itens 3 e 9).

Sem NiceGUI: testável com pytest (banco temporário via fixture banco_temporario).
"""

from __future__ import annotations

import getpass
import json
import sqlite3
import uuid
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

import database
from lcfo_dominio import tipos

ORIGENS_ENTIDADE = ("manual", "pivo", "importacao", "anotacao_oficial")
ORIGENS_VINCULO = ("manual", "pivo", "anotacao", "fusao")
CAMPOS_VINCULO_EDITAVEIS = ("rotulo", "direcao", "situacao", "data_ref", "confianca", "observacao")
# Chaves do nó vis que são reconstruídas a cada leitura (não vão para extra_json).
_CHAVES_DERIVADAS = ("id", "image")

# DDL e garantir_schema moram em lcfo_dominio/schema.py (sem dependências): o database chama
# garantir_schema no próprio import, e importar este módulo de lá criava um import circular.
from lcfo_dominio.schema import garantir_schema  # noqa: E402,F401


# ------------------------------------------------------------------ infraestrutura
def agora() -> str:
    return datetime.now().isoformat(timespec="seconds")


def usuario_atual() -> str:
    try:
        return getpass.getuser()
    except Exception:
        return ""


@contextmanager
def transacao():
    """BEGIN IMMEDIATE: a escrita reserva o banco já no início (sem 'database is locked' no meio)."""
    conn = database.get_db_connection()
    conn.isolation_level = None
    conn.execute("PRAGMA busy_timeout = 10000")
    conn.execute("BEGIN IMMEDIATE")
    try:
        yield conn
        conn.execute("COMMIT")
    except BaseException:
        conn.execute("ROLLBACK")
        raise
    finally:
        conn.close()


def registrar_evento(conn, id_sketch: str, acao: str, alvo_tipo: str = "", alvo_id: str = "",
                     detalhe: Optional[dict] = None, quem: Optional[str] = None) -> None:
    conn.execute(
        "INSERT INTO sketch_eventos (id_sketch, quando, quem, acao, alvo_tipo, alvo_id, detalhe_json) VALUES (?,?,?,?,?,?,?)",
        (id_sketch, agora(), quem if quem is not None else usuario_atual(), acao, alvo_tipo, alvo_id,
         json.dumps(detalhe or {}, ensure_ascii=False)))
    conn.execute("UPDATE caso_sketches SET versao = COALESCE(versao, 0) + 1 WHERE id_sketch = ?", (id_sketch,))


def versao(id_sketch: str) -> int:
    conn = database.get_db_connection()
    try:
        r = conn.execute("SELECT versao FROM caso_sketches WHERE id_sketch = ?", (id_sketch,)).fetchone()
        return int(r[0] or 0) if r else 0
    finally:
        conn.close()


def _lista_curta(ids: List[str], n: int = 200) -> List[str]:
    return list(ids)[:n]


# ------------------------------------------------------------------ entidades e vínculos
def _linha_de_no_vis(no: Dict[str, Any]) -> Dict[str, Any]:
    """Nó vis (do enrich ou do JSON antigo) -> colunas de sketch_entidades."""
    node_id = str(no["id"])
    t = tipos.obter(no.get("tipo")) or tipos.tipo_do_id(node_id)
    chave = t.chave if t else str(no.get("tipo") or "empresa")
    campos = {}
    if no.get("nome_titular"):
        campos["nome_titular"] = no["nome_titular"]
    extra = {k: v for k, v in no.items() if k not in _CHAVES_DERIVADAS}
    return {"node_id": node_id, "tipo": chave, "valor": tipos.valor_do_id(node_id),
            "rotulo": no.get("label") or "", "campos_json": json.dumps(campos, ensure_ascii=False),
            "extra_json": json.dumps(extra, ensure_ascii=False)}


def _linha_de_aresta_vis(e: Dict[str, Any]) -> Dict[str, Any]:
    extra = {k: v for k, v in e.items() if k not in ("id", "from", "to")}
    n = e.get("ocorrencias") or e.get("n_ocorrencias") or 1
    return {"id": str(e.get("id") or f"vin_{uuid.uuid4().hex[:12]}"), "de_id": str(e["from"]), "para_id": str(e["to"]),
            "rotulo": e.get("label") or "", "n_ocorrencias": int(n) if str(n).isdigit() else 1,
            "situacao": "hipotese" if e.get("situacao") == "hipotese" else "confirmado",
            "extra_json": json.dumps(extra, ensure_ascii=False)}


def adicionar(id_sketch: str, nos: Iterable[Dict[str, Any]] = (), arestas: Iterable[Dict[str, Any]] = (), *,
              origem: str = "pivo", origem_por_no: Optional[Dict[str, str]] = None, id_execucao: Optional[str] = None,
              quem: Optional[str] = None, acao: str = "adicionar", _conn=None) -> Dict[str, List[str]]:
    """
    Grava nós e vínculos vindos do Enrich/Add/pivô. Quem já está ativo no sketch não
    duplica (a identidade é o node_id); quem estava na lixeira volta (reativado).
    """
    origem_por_no = origem_por_no or {}
    res = {"nos_novos": [], "vinculos_novos": [], "reativados": [], "existentes": []}
    quem = quem if quem is not None else usuario_atual()

    def _executar(conn):
        ts = agora()
        for no in nos:
            ln = _linha_de_no_vis(no)
            atual = conn.execute("SELECT removido_em FROM sketch_entidades WHERE id_sketch = ? AND node_id = ?",
                                 (id_sketch, ln["node_id"])).fetchone()
            if atual is None:
                conn.execute(
                    """INSERT INTO sketch_entidades (id_sketch, node_id, tipo, valor, rotulo, campos_json, origem,
                       id_execucao, extra_json, criado_em, criado_por) VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                    (id_sketch, ln["node_id"], ln["tipo"], ln["valor"], ln["rotulo"], ln["campos_json"],
                     origem_por_no.get(ln["node_id"], origem), id_execucao, ln["extra_json"], ts, quem))
                res["nos_novos"].append(ln["node_id"])
            elif atual[0]:
                conn.execute("""UPDATE sketch_entidades SET removido_em = NULL, removido_por = NULL, motivo_remocao = NULL,
                                lote_remocao = NULL WHERE id_sketch = ? AND node_id = ?""", (id_sketch, ln["node_id"]))
                res["reativados"].append(ln["node_id"])
            else:
                res["existentes"].append(ln["node_id"])
        for e in arestas:
            la = _linha_de_aresta_vis(e)
            atual = conn.execute("SELECT removido_em FROM sketch_vinculos WHERE id = ?", (la["id"],)).fetchone()
            if atual is None:
                conn.execute(
                    """INSERT INTO sketch_vinculos (id, id_sketch, de_id, para_id, rotulo, situacao, n_ocorrencias, origem,
                       id_execucao, extra_json, criado_em, criado_por) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (la["id"], id_sketch, la["de_id"], la["para_id"], la["rotulo"], la["situacao"], la["n_ocorrencias"],
                     origem if origem in ORIGENS_VINCULO and origem != "manual" else "pivo", id_execucao, la["extra_json"], ts, quem))
                res["vinculos_novos"].append(la["id"])
            elif atual[0]:
                conn.execute("""UPDATE sketch_vinculos SET removido_em = NULL, removido_por = NULL, motivo_remocao = NULL,
                                lote_remocao = NULL WHERE id = ?""", (la["id"],))
                res["reativados"].append(la["id"])
        if res["nos_novos"] or res["vinculos_novos"] or res["reativados"]:
            registrar_evento(conn, id_sketch, acao, "lote", id_execucao or "", {
                "nos": _lista_curta(res["nos_novos"]), "vinculos": len(res["vinculos_novos"]),
                "reativados": _lista_curta(res["reativados"]), "origem": origem}, quem)

    if _conn is not None:
        _executar(_conn)
    else:
        with transacao() as conn:
            _executar(conn)
    return res


def canonico(tipo: str, valor: str):
    """(TipoEntidade, valor canônico, node_id) ou ValueError — a mesma regra do quick-add e do colar."""
    t = tipos.obter(tipo)
    if not t:
        raise ValueError(f"tipo desconhecido: {tipo}")
    canon = t.normalizar(valor)
    if not canon or not t.validar(canon):
        raise ValueError(f"valor inválido para {t.rotulo}: {valor!r}")
    return t, canon, t.node_id(canon)


def _criar(conn, id_sketch: str, tipo: str, valor: str, campos: Optional[dict], rotulo: str, origem: str,
           id_execucao: Optional[str], quem: str) -> Dict[str, str]:
    """Uma entidade, dentro de uma transação já aberta. estado: novo | existente | reativado."""
    t, canon, node_id = canonico(tipo, valor)
    atual = conn.execute("SELECT removido_em FROM sketch_entidades WHERE id_sketch = ? AND node_id = ?",
                         (id_sketch, node_id)).fetchone()
    if atual is not None and not atual[0]:
        return {"node_id": node_id, "estado": "existente"}         # já está no sketch: não duplica
    if atual is not None:                                           # estava na lixeira: volta
        conn.execute("""UPDATE sketch_entidades SET removido_em = NULL, removido_por = NULL, motivo_remocao = NULL,
                        lote_remocao = NULL WHERE id_sketch = ? AND node_id = ?""", (id_sketch, node_id))
        return {"node_id": node_id, "estado": "reativado"}
    conn.execute(
        """INSERT INTO sketch_entidades (id_sketch, node_id, tipo, valor, rotulo, campos_json, origem, id_execucao,
           criado_em, criado_por) VALUES (?,?,?,?,?,?,?,?,?,?)""",
        (id_sketch, node_id, t.chave, canon, rotulo or "", json.dumps(campos or {}, ensure_ascii=False),
         origem, id_execucao, agora(), quem))
    return {"node_id": node_id, "estado": "novo"}


def criar_entidade(id_sketch: str, tipo: str, valor: str, *, campos: Optional[dict] = None, rotulo: str = "",
                   origem: str = "manual", id_execucao: Optional[str] = None, quem: Optional[str] = None,
                   posicao: Optional[Dict[str, float]] = None, layout: str = "organico") -> str:
    """Entidade criada pelo analista (quick-add, formulário). Devolve o node_id canônico."""
    return criar_entidades(id_sketch, [{"tipo": tipo, "valor": valor, "campos": campos, "rotulo": rotulo, "posicao": posicao}],
                           origem=origem, id_execucao=id_execucao, quem=quem, layout=layout)[0]["node_id"]


def criar_entidades(id_sketch: str, itens: List[Dict[str, Any]], *, origem: str = "manual", id_execucao: Optional[str] = None,
                    quem: Optional[str] = None, layout: str = "organico") -> List[Dict[str, str]]:
    """
    Várias entidades numa transação só (colar N linhas = um evento). Cada item: tipo, valor e,
    opcionais, campos, rotulo, posicao {x, y}. Valor inválido levanta ValueError antes de gravar
    qualquer coisa. A posição de quem é novo vai para sketch_posicoes (sobrevive à recarga).
    """
    if origem not in ORIGENS_ENTIDADE:
        raise ValueError(f"origem inválida: {origem}")
    for it in itens:
        canonico(it["tipo"], it["valor"])                         # valida tudo antes de abrir a transação
    quem = quem if quem is not None else usuario_atual()
    out = []
    with transacao() as conn:
        ts = agora()
        for it in itens:
            r = _criar(conn, id_sketch, it["tipo"], it["valor"], it.get("campos"), it.get("rotulo") or "", origem, id_execucao, quem)
            pos = it.get("posicao")
            if pos and r["estado"] != "existente":
                conn.execute(
                    """INSERT INTO sketch_posicoes (id_sketch, node_id, layout, x, y, fixado, atualizado_em) VALUES (?,?,?,?,?,0,?)
                       ON CONFLICT(id_sketch, node_id, layout) DO UPDATE SET x = excluded.x, y = excluded.y,
                       atualizado_em = excluded.atualizado_em""", (id_sketch, r["node_id"], layout, float(pos["x"]), float(pos["y"]), ts))
            out.append(r)
        mudou = [r for r in out if r["estado"] != "existente"]
        if mudou:
            acao = "criar_entidade" if len(itens) == 1 and mudou[0]["estado"] == "novo" else (
                "reativar" if len(itens) == 1 else "criar_entidades")
            registrar_evento(conn, id_sketch, acao, "entidade" if len(itens) == 1 else "lote",
                             mudou[0]["node_id"] if len(itens) == 1 else (id_execucao or ""),
                             {"nos": _lista_curta([r["node_id"] for r in mudou]), "origem": origem,
                              "reativados": [r["node_id"] for r in mudou if r["estado"] == "reativado"],
                              "campos": sorted(set().union(*[set(it.get("campos") or {}) for it in itens]))}, quem)
    return out


def criar_vinculo(id_sketch: str, de_id: str, para_id: str, *, rotulo: str = "", direcao: str = "de_para",
                  situacao: str = "confirmado", data_ref: Optional[str] = None, confianca: Optional[int] = None,
                  observacao: str = "", origem: str = "manual", ref_oficial: Optional[str] = None,
                  id_execucao: Optional[str] = None, lote_fusao: Optional[str] = None, quem: Optional[str] = None,
                  id_vinculo: Optional[str] = None) -> str:
    if origem not in ORIGENS_VINCULO:
        raise ValueError(f"origem inválida: {origem}")
    if direcao not in ("de_para", "para_de", "nenhuma") or situacao not in ("confirmado", "hipotese"):
        raise ValueError("direção ou situação inválida")
    vid = id_vinculo or f"vm_{uuid.uuid4().hex[:12]}"
    quem = quem if quem is not None else usuario_atual()
    with transacao() as conn:
        conn.execute(
            """INSERT INTO sketch_vinculos (id, id_sketch, de_id, para_id, rotulo, direcao, situacao, data_ref, confianca,
               observacao, origem, ref_oficial, id_execucao, lote_fusao, criado_em, criado_por)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (vid, id_sketch, de_id, para_id, rotulo, direcao, situacao, data_ref, confianca, observacao, origem,
             ref_oficial, id_execucao, lote_fusao, agora(), quem))
        registrar_evento(conn, id_sketch, "criar_vinculo", "vinculo", vid,
                         {"de": de_id, "para": para_id, "rotulo": rotulo, "situacao": situacao, "origem": origem}, quem)
    return vid


def atualizar_entidade(id_sketch: str, node_id: str, *, campos: Optional[dict] = None, rotulo: Optional[str] = None,
                       nota: Optional[str] = None, aparencia: Optional[dict] = None, quem: Optional[str] = None) -> None:
    sets, vals, detalhe = [], [], {}
    with transacao() as conn:
        r = conn.execute("SELECT campos_json FROM sketch_entidades WHERE id_sketch = ? AND node_id = ? AND removido_em IS NULL",
                         (id_sketch, node_id)).fetchone()
        if r is None:
            raise KeyError(f"{node_id} não está ativo no sketch {id_sketch}")
        if campos is not None:
            t = tipos.tipo_do_id(node_id)
            primarios = {c.nome for c in (t.campos if t else ()) if c.primario}
            if primarios & set(campos):
                raise ValueError("o campo primário é a identidade: para mudá-lo, funda ou recrie a entidade")
            novos = {**json.loads(r[0] or "{}"), **campos}
            sets.append("campos_json = ?"); vals.append(json.dumps(novos, ensure_ascii=False)); detalhe["campos"] = sorted(campos)
        if rotulo is not None:
            sets.append("rotulo = ?"); vals.append(rotulo); detalhe["rotulo"] = rotulo
        if nota is not None:
            sets.append("nota = ?"); vals.append(nota); detalhe["nota"] = True
        if aparencia is not None:
            sets.append("aparencia_json = ?"); vals.append(json.dumps(aparencia, ensure_ascii=False)); detalhe["aparencia"] = aparencia
        if not sets:
            return
        conn.execute(f"UPDATE sketch_entidades SET {', '.join(sets)} WHERE id_sketch = ? AND node_id = ?", (*vals, id_sketch, node_id))
        registrar_evento(conn, id_sketch, "atualizar_entidade", "entidade", node_id, detalhe, quem)


def atualizar_vinculo(id_sketch: str, id_vinculo: str, quem: Optional[str] = None, **campos) -> None:
    invalidos = set(campos) - set(CAMPOS_VINCULO_EDITAVEIS)
    if invalidos:
        raise ValueError(f"campos não editáveis: {sorted(invalidos)}")
    if not campos:
        return
    with transacao() as conn:
        n = conn.execute(f"UPDATE sketch_vinculos SET {', '.join(f'{k} = ?' for k in campos)} WHERE id = ? AND id_sketch = ? AND removido_em IS NULL",
                         (*campos.values(), id_vinculo, id_sketch)).rowcount
        if not n:
            raise KeyError(f"vínculo {id_vinculo} não está ativo no sketch {id_sketch}")
        registrar_evento(conn, id_sketch, "atualizar_vinculo", "vinculo", id_vinculo, campos, quem)


# ------------------------------------------------------------------ remoção lógica e lixeira (item 4)
def remover(id_sketch: str, node_ids: Iterable[str] = (), vinculo_ids: Iterable[str] = (), *, motivo: str = "",
            quem: Optional[str] = None, origens: Optional[Iterable[str]] = None, id_execucao: Optional[str] = None) -> Dict[str, Any]:
    """
    Remove do sketch (sem apagar linhas). Os vínculos ativos ligados aos nós saem no mesmo
    lote. `origens` (com node_ids vazio) remove todas as entidades ativas dessas origens —
    é o "Limpar Enrich", agora reversível pela lixeira.
    """
    lote = uuid.uuid4().hex
    quem = quem if quem is not None else usuario_atual()
    node_ids, vinculo_ids = list(node_ids), list(vinculo_ids)
    with transacao() as conn:
        if id_execucao:                          # "Desfazer esta execução": só o que ela criou
            node_ids += [r[0] for r in conn.execute(
                "SELECT node_id FROM sketch_entidades WHERE id_sketch = ? AND removido_em IS NULL AND id_execucao = ?",
                (id_sketch, id_execucao))]
            vinculo_ids += [r[0] for r in conn.execute(
                "SELECT id FROM sketch_vinculos WHERE id_sketch = ? AND removido_em IS NULL AND id_execucao = ?",
                (id_sketch, id_execucao))]
        if origens is not None and not node_ids:
            marc = ",".join("?" * len(list(origens)))
            node_ids = [r[0] for r in conn.execute(
                f"SELECT node_id FROM sketch_entidades WHERE id_sketch = ? AND removido_em IS NULL AND origem IN ({marc})",
                (id_sketch, *origens))]
        ts = agora()
        prev = _previa(conn, id_sketch, node_ids, vinculo_ids)
        for coluna, tabela, ids in (("node_id", "sketch_entidades", prev["nos"]), ("id", "sketch_vinculos", prev["vinculos"])):
            for i in range(0, len(ids), 500):
                parte = ids[i:i + 500]
                conn.execute(
                    f"""UPDATE {tabela} SET removido_em = ?, removido_por = ?, motivo_remocao = ?, lote_remocao = ?
                        WHERE id_sketch = ? AND removido_em IS NULL AND {coluna} IN ({",".join("?" * len(parte))})""",
                    (ts, quem, motivo, lote, id_sketch, *parte))
        nos, vins = len(prev["nos"]), len(prev["vinculos"])
        if nos or vins:
            registrar_evento(conn, id_sketch, "remover", "lote", lote,
                             {"nos": _lista_curta(prev["nos"]), "n_nos": nos, "n_vinculos": vins, "motivo": motivo}, quem)
    return {"lote": lote if (nos or vins) else None, "nos": nos, "vinculos": vins,
            "ids_nos": prev["nos"], "ids_vinculos": prev["vinculos"]}


def _previa(conn, id_sketch: str, node_ids: List[str], vinculo_ids: List[str]) -> Dict[str, List[str]]:
    """O que sairia: entidades ATIVAS pedidas + vínculos pedidos + vínculos ligados a elas (cascata)."""
    nos: List[str] = []
    vins: List[str] = []
    for i in range(0, len(node_ids), 500):
        parte = node_ids[i:i + 500]
        marc = ",".join("?" * len(parte))
        nos += [r[0] for r in conn.execute(
            f"SELECT node_id FROM sketch_entidades WHERE id_sketch = ? AND removido_em IS NULL AND node_id IN ({marc})",
            (id_sketch, *parte))]
    for i in range(0, len(nos), 500):
        parte = nos[i:i + 500]
        marc = ",".join("?" * len(parte))
        vins += [r[0] for r in conn.execute(
            f"""SELECT id FROM sketch_vinculos WHERE id_sketch = ? AND removido_em IS NULL
                AND (de_id IN ({marc}) OR para_id IN ({marc}))""", (id_sketch, *parte, *parte))]
    for i in range(0, len(vinculo_ids), 500):
        parte = vinculo_ids[i:i + 500]
        marc = ",".join("?" * len(parte))
        vins += [r[0] for r in conn.execute(
            f"SELECT id FROM sketch_vinculos WHERE id_sketch = ? AND removido_em IS NULL AND id IN ({marc})",
            (id_sketch, *parte))]
    return {"nos": nos, "vinculos": list(dict.fromkeys(vins))}


def previa_remocao(id_sketch: str, node_ids: Iterable[str] = (), vinculo_ids: Iterable[str] = ()) -> Dict[str, List[str]]:
    """
    Para o diálogo ("1 nó e 3 vínculos serão removidos"). `oficiais_*` = o que foi pedido mas não
    é do sketch (cluster da blacklist): não pode ser removido, só isolado/anotado.
    """
    node_ids, vinculo_ids = list(node_ids), list(vinculo_ids)
    conn = database.get_db_connection()
    try:
        prev = _previa(conn, id_sketch, node_ids, vinculo_ids)
    finally:
        conn.close()
    ativos_n, ativos_v = set(prev["nos"]), set(prev["vinculos"])
    prev["oficiais_nos"] = [n for n in node_ids if n not in ativos_n]
    prev["oficiais_vinculos"] = [v for v in vinculo_ids if v not in ativos_v]
    return prev


def restaurar(id_sketch: str, lote: str, quem: Optional[str] = None) -> Dict[str, int]:
    """Devolve o lote inteiro (nós, vínculos; as posições nunca saíram de sketch_posicoes)."""
    with transacao() as conn:
        ids_nos = [r[0] for r in conn.execute(
            "SELECT node_id FROM sketch_entidades WHERE id_sketch = ? AND lote_remocao = ?", (id_sketch, lote))]
        ids_vins = [r[0] for r in conn.execute(
            "SELECT id FROM sketch_vinculos WHERE id_sketch = ? AND lote_remocao = ?", (id_sketch, lote))]
        limpar = "removido_em = NULL, removido_por = NULL, motivo_remocao = NULL, lote_remocao = NULL"
        nos = conn.execute(f"UPDATE sketch_entidades SET {limpar} WHERE id_sketch = ? AND lote_remocao = ?", (id_sketch, lote)).rowcount
        vins = conn.execute(f"UPDATE sketch_vinculos SET {limpar} WHERE id_sketch = ? AND lote_remocao = ?", (id_sketch, lote)).rowcount
        if nos or vins:
            registrar_evento(conn, id_sketch, "restaurar", "lote", lote, {"n_nos": nos, "n_vinculos": vins}, quem)
    return {"nos": nos, "vinculos": vins, "ids_nos": ids_nos, "ids_vinculos": ids_vins}


def listar_lixeira(id_sketch: str, limite: int = 100) -> List[Dict[str, Any]]:
    """Lotes removidos, do mais recente ao mais antigo, com contagens, quem, quando e motivo."""
    conn = database.get_db_connection()
    try:
        rows = conn.execute("""
            SELECT lote_remocao AS lote, MAX(removido_em) AS quando, MAX(removido_por) AS quem, MAX(motivo_remocao) AS motivo,
                   SUM(n_nos) AS n_nos, SUM(n_vinculos) AS n_vinculos, GROUP_CONCAT(amostra, '|') AS amostra FROM (
                SELECT lote_remocao, removido_em, removido_por, motivo_remocao, 1 AS n_nos, 0 AS n_vinculos,
                       COALESCE(NULLIF(rotulo, ''), node_id) AS amostra
                  FROM sketch_entidades WHERE id_sketch = ? AND removido_em IS NOT NULL AND lote_remocao NOT LIKE 'desfeita_%'
                UNION ALL
                SELECT lote_remocao, removido_em, removido_por, motivo_remocao, 0, 1, NULL
                  FROM sketch_vinculos WHERE id_sketch = ? AND removido_em IS NOT NULL AND lote_remocao NOT LIKE 'desfeita_%')
            GROUP BY lote_remocao ORDER BY quando DESC LIMIT ?""", (id_sketch, id_sketch, limite)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


# ------------------------------------------------------------------ posições (item 8)
def salvar_posicoes(id_sketch: str, layout: str, posicoes: Dict[str, Dict[str, float]], quem: Optional[str] = None) -> int:
    if not posicoes:
        return 0
    ts = agora()
    with transacao() as conn:
        conn.executemany(
            """INSERT INTO sketch_posicoes (id_sketch, node_id, layout, x, y, fixado, atualizado_em) VALUES (?,?,?,?,?,?,?)
               ON CONFLICT(id_sketch, node_id, layout) DO UPDATE SET x = excluded.x, y = excluded.y,
               fixado = COALESCE(excluded.fixado, sketch_posicoes.fixado), atualizado_em = excluded.atualizado_em""",
            [(id_sketch, nid, layout, float(p["x"]), float(p["y"]), int(bool(p.get("fixado", 0))), ts) for nid, p in posicoes.items()])
        registrar_evento(conn, id_sketch, "posicoes", "layout", layout, {"n": len(posicoes)}, quem)
    return len(posicoes)


def fixar(id_sketch: str, posicoes: Dict[str, Dict[str, float]], layout: str, fixado: bool, quem: Optional[str] = None) -> int:
    """Nó fixado é ignorado pelo relayout (grava a posição atual junto)."""
    for p in posicoes.values():
        p["fixado"] = bool(fixado)
    ts = agora()
    with transacao() as conn:
        conn.executemany(
            """INSERT INTO sketch_posicoes (id_sketch, node_id, layout, x, y, fixado, atualizado_em) VALUES (?,?,?,?,?,?,?)
               ON CONFLICT(id_sketch, node_id, layout) DO UPDATE SET x = excluded.x, y = excluded.y, fixado = excluded.fixado,
               atualizado_em = excluded.atualizado_em""",
            [(id_sketch, nid, layout, float(p["x"]), float(p["y"]), int(bool(fixado)), ts) for nid, p in posicoes.items()])
        registrar_evento(conn, id_sketch, "fixar" if fixado else "soltar", "layout", layout, {"nos": _lista_curta(list(posicoes))}, quem)
    return len(posicoes)


def limpar_posicoes(id_sketch: str, layout: str, manter_fixados: bool = True, quem: Optional[str] = None) -> int:
    """'Refazer tudo': apaga as posições salvas do layout (as fixadas ficam)."""
    with transacao() as conn:
        n = conn.execute(f"DELETE FROM sketch_posicoes WHERE id_sketch = ? AND layout = ?{' AND fixado = 0' if manter_fixados else ''}",
                         (id_sketch, layout)).rowcount
        registrar_evento(conn, id_sketch, "refazer_layout", "layout", layout, {"n": n}, quem)
    return n


def carregar_posicoes(id_sketch: str, layout: str) -> Dict[str, Dict[str, float]]:
    conn = database.get_db_connection()
    try:
        rows = conn.execute("SELECT node_id, x, y, fixado FROM sketch_posicoes WHERE id_sketch = ? AND layout = ?",
                            (id_sketch, layout)).fetchall()
        return {r["node_id"]: {"x": r["x"], "y": r["y"], "fixado": bool(r["fixado"])} for r in rows}
    finally:
        conn.close()


# ------------------------------------------------------------------ aparência, notas e anotações (item 9)
CORES_APARENCIA = {"azul": ("#3C6FA8", "#5A94D6"), "vermelho": ("#B94A3C", "#E0684F"), "roxo": ("#6D4FA8", "#9273D6"),
                   "verde": ("#4A7A5A", "#6BA57D"), "ambar": ("#A8843C", "#D6AA5A"), "rosa": ("#A84F7A", "#D6739F"),
                   "ciano": ("#3F7F86", "#5FA8B0"), "cinza": ("#4A4A4A", "#767676")}
FORMAS_APARENCIA = ("circulo", "quadrado", "losango", "hexagono", "triangulo")
TAMANHOS_APARENCIA = {"P": 14, "M": 20, "G": 28}
FLAGS = {"red": "#EF4444", "orange": "#F97316", "blue": "#3B82F6", "green": "#22C55E", "yellow": "#EAB308"}


def validar_aparencia(ap: Optional[dict]) -> dict:
    ap = dict(ap or {})
    out = {}
    if ap.get("cor") in CORES_APARENCIA:
        out["cor"] = ap["cor"]
    if ap.get("forma") in FORMAS_APARENCIA:
        out["forma"] = ap["forma"]
    if ap.get("tamanho") in TAMANHOS_APARENCIA:
        out["tamanho"] = ap["tamanho"]
    if ap.get("flag") in FLAGS:
        out["flag"] = ap["flag"]
    return out


def estilo_no(tipo: str, aparencia: Optional[dict], is_hub: bool = False, tem_nota: bool = False) -> Dict[str, Any]:
    """Propriedades visuais do vis para um nó com essa aparência (vazia = o padrão do tipo)."""
    t = tipos.obter(tipo)
    ap = validar_aparencia(aparencia)
    bg, borda = CORES_APARENCIA.get(ap.get("cor"), t.cor if t else ("#4A4A4A", "#767676"))
    from lcfo_dominio.icones import svg_no
    flag = ap.get("flag")
    cor_borda = FLAGS[flag] if flag else ("#FF7300" if is_hub else borda)
    return {"image": svg_no(t.icone if t else "empresa", bg, borda, is_hub, ap.get("forma", "circulo"), FLAGS.get(flag, "")),
            "color": {"background": bg, "border": cor_borda}, "size": TAMANHOS_APARENCIA.get(ap.get("tamanho"), 26 if is_hub else 20),
            "borderWidth": 4 if flag else (3.5 if is_hub else 1.2), "flag": flag, "aparencia": ap, "tem_nota": bool(tem_nota)}


def anotar(id_sketch: str, node_id: str, *, nota: Optional[str] = None, aparencia: Optional[dict] = None,
           campos: Optional[dict] = None, quem: Optional[str] = None) -> str:
    """
    Nota/aparência/campos de uma entidade. Do analista: atualiza a própria linha. Oficial (do
    cluster): cria/atualiza uma anotação por cima (origem 'anotacao_oficial') — o dado oficial
    nunca é sobrescrito. Devolve 'entidade' ou 'anotacao'.
    """
    quem = quem if quem is not None else usuario_atual()
    ap = validar_aparencia(aparencia) if aparencia is not None else None
    with transacao() as conn:
        r = conn.execute("SELECT origem, campos_json FROM sketch_entidades WHERE id_sketch = ? AND node_id = ? AND removido_em IS NULL",
                         (id_sketch, node_id)).fetchone()
        alvo = "entidade" if r is not None and r["origem"] != "anotacao_oficial" else "anotacao"
        if r is None:
            t = tipos.tipo_do_id(node_id)
            conn.execute(
                """INSERT INTO sketch_entidades (id_sketch, node_id, tipo, valor, origem, criado_em, criado_por)
                   VALUES (?,?,?,?,'anotacao_oficial',?,?)
                   ON CONFLICT(id_sketch, node_id) DO UPDATE SET removido_em = NULL, removido_por = NULL,
                   motivo_remocao = NULL, lote_remocao = NULL""",
                (id_sketch, node_id, t.chave if t else "empresa", tipos.valor_do_id(node_id), agora(), quem))
        sets, vals, det = [], [], {}
        if nota is not None:
            sets.append("nota = ?"); vals.append(nota); det["nota"] = True
        if ap is not None:
            sets.append("aparencia_json = ?"); vals.append(json.dumps(ap, ensure_ascii=False)); det["aparencia"] = ap
        if campos is not None:
            t = tipos.tipo_do_id(node_id)
            if {c.nome for c in (t.campos if t else ()) if c.primario} & set(campos):
                raise ValueError("o campo primário é a identidade: para mudá-lo, funda ou recrie a entidade")
            atuais = json.loads((r["campos_json"] if r is not None else None) or "{}")
            sets.append("campos_json = ?"); vals.append(json.dumps({**atuais, **campos}, ensure_ascii=False)); det["campos"] = sorted(campos)
        if sets:
            conn.execute(f"UPDATE sketch_entidades SET {', '.join(sets)} WHERE id_sketch = ? AND node_id = ?", (*vals, id_sketch, node_id))
        registrar_evento(conn, id_sketch, "anotar" if alvo == "anotacao" else "atualizar_entidade", "entidade", node_id, det, quem)
    return alvo


def detalhes(id_sketch: str, node_id: str) -> Dict[str, Any]:
    """Para o painel de edição: linha do sketch (ou anotação), campos, nota, aparência e origem."""
    conn = database.get_db_connection()
    try:
        r = conn.execute("SELECT * FROM sketch_entidades WHERE id_sketch = ? AND node_id = ? AND removido_em IS NULL",
                         (id_sketch, node_id)).fetchone()
        ex = None
        if r is not None and r["id_execucao"]:
            ex = conn.execute("SELECT pivo, parametros_json, iniciado_em, usuario FROM sketch_execucoes WHERE id = ?",
                              (r["id_execucao"],)).fetchone()
        ultima_nota = conn.execute("""SELECT quem, quando FROM sketch_eventos WHERE id_sketch = ? AND alvo_id = ?
                                      AND detalhe_json LIKE '%"nota": true%' ORDER BY id DESC LIMIT 1""", (id_sketch, node_id)).fetchone()
    finally:
        conn.close()
    oficial = r is None or r["origem"] == "anotacao_oficial"
    return {"node_id": node_id, "oficial": oficial, "origem": None if r is None else r["origem"],
            "campos": json.loads(r["campos_json"] or "{}") if r is not None else {},
            "nota": (r["nota"] or "") if r is not None else "",
            "aparencia": json.loads(r["aparencia_json"] or "{}") if r is not None and r["aparencia_json"] else {},
            "criado_em": None if r is None else r["criado_em"], "criado_por": None if r is None else r["criado_por"],
            "id_execucao": None if r is None else r["id_execucao"],
            "execucao": dict(ex) if ex else None,
            "nota_por": dict(ultima_nota) if ultima_nota else None}


# ------------------------------------------------------------------ filtros por regra (item 12)
OPERADORES = {"texto": ("é", "não é", "contém", "começa com", "vazio", "preenchido"),
              "numero": ("=", "≠", ">", "≥", "<", "≤", "entre"),
              "data": ("antes", "depois", "entre"),
              "lista": ("é um de",)}


def validar_filtro(filtro: dict) -> dict:
    """{combinacao: todas|qualquer, modo: ocultar|destacar, regras: [{campo, operador, valor}]}."""
    regras = []
    for r in (filtro or {}).get("regras") or []:
        if isinstance(r, dict) and r.get("campo") and r.get("operador") in {o for ops in OPERADORES.values() for o in ops}:
            regras.append({"campo": str(r["campo"])[:40], "operador": r["operador"], "valor": r.get("valor")})
    return {"combinacao": "qualquer" if (filtro or {}).get("combinacao") == "qualquer" else "todas",
            "modo": "destacar" if (filtro or {}).get("modo") == "destacar" else "ocultar", "regras": regras}


def salvar_filtro(id_sketch: str, nome: str, filtro: dict, quem: Optional[str] = None) -> str:
    f = validar_filtro(filtro)
    if not f["regras"]:
        raise ValueError("o filtro não tem nenhuma regra válida")
    fid = f"flt_{uuid.uuid4().hex[:10]}"
    quem = quem if quem is not None else usuario_atual()
    with transacao() as conn:
        conn.execute("INSERT INTO sketch_filtros (id, id_sketch, nome, regras_json, criado_em, criado_por) VALUES (?,?,?,?,?,?)",
                     (fid, id_sketch, (nome or "Filtro").strip()[:60], json.dumps(f, ensure_ascii=False), agora(), quem))
        registrar_evento(conn, id_sketch, "salvar_filtro", "filtro", fid, {"nome": nome}, quem)
    return fid


def listar_filtros(id_sketch: str) -> List[Dict[str, Any]]:
    conn = database.get_db_connection()
    try:
        return [{"id": r["id"], "nome": r["nome"], "filtro": json.loads(r["regras_json"])} for r in conn.execute(
            "SELECT * FROM sketch_filtros WHERE id_sketch = ? ORDER BY criado_em", (id_sketch,))]
    finally:
        conn.close()


def remover_filtro(id_sketch: str, id_filtro: str, quem: Optional[str] = None) -> None:
    """Filtro salvo é só uma preferência de visualização: sai de verdade (fica o evento)."""
    with transacao() as conn:
        conn.execute("DELETE FROM sketch_filtros WHERE id = ? AND id_sketch = ?", (id_filtro, id_sketch))
        registrar_evento(conn, id_sketch, "remover_filtro", "filtro", id_filtro, {}, quem)


def rotulos_flags(id_sketch: str) -> Dict[str, str]:
    conn = database.get_db_connection()
    try:
        r = conn.execute("SELECT flags_json FROM caso_sketches WHERE id_sketch = ?", (id_sketch,)).fetchone()
        return json.loads(r[0]) if r and r[0] else {}
    finally:
        conn.close()


def salvar_rotulos_flags(id_sketch: str, rotulos: Dict[str, str], quem: Optional[str] = None) -> None:
    rotulos = {k: str(v).strip()[:60] for k, v in rotulos.items() if k in FLAGS and str(v).strip()}
    with transacao() as conn:
        conn.execute("UPDATE caso_sketches SET flags_json = ? WHERE id_sketch = ?", (json.dumps(rotulos, ensure_ascii=False), id_sketch))
        registrar_evento(conn, id_sketch, "rotulos_flags", "sketch", id_sketch, rotulos, quem)


# ------------------------------------------------------------------ leitura (overlay do payload)
def _no_vis(r: sqlite3.Row) -> Dict[str, Any]:
    extra = json.loads(r["extra_json"] or "{}")
    campos = json.loads(r["campos_json"] or "{}")
    t = tipos.obter(r["tipo"])
    no = dict(extra)
    if not no:                                   # criado pelo analista: monta pelo registro de tipos
        nome = campos.get("nome_titular", "")
        valor_fmt = t.formatar(r["valor"]) if t else r["valor"]
        cor = t.cor if t else ("#4A4A4A", "#767676")
        no = {"label": t.rotular({"valor": r["valor"], "nome_titular": nome}) if t else r["valor"],
              "title": f"{(t.rotulo if t else r['tipo'])}: {valor_fmt}" + (f"\n{nome}" if nome else ""),
              "shape": t.forma if t else "image", "size": 20, "degree": 1, "betweenness": 0.0, "community": 0,
              "valor": valor_fmt, "nome_titular": nome, "flag": None,
              "color": {"background": cor[0], "border": cor[1]}, "font": dict(tipos.FONTE_NO)}
    no["id"] = r["node_id"]
    no["tipo"] = r["tipo"]
    no["image"] = t.icone_svg() if t else tipos.REGISTRO["empresa"].icone_svg()
    if r["rotulo"]:
        no["label"] = r["rotulo"]
    no["origem"] = r["origem"]
    if r["id_execucao"]:
        no["id_execucao"] = r["id_execucao"]
    if r["aparencia_json"] or r["nota"]:
        ap = json.loads(r["aparencia_json"] or "{}")
        if ap:
            no.update(estilo_no(r["tipo"], ap, tem_nota=bool(r["nota"])))
        no["tem_nota"] = bool(r["nota"])
    return no


def _aresta_vis(r: sqlite3.Row) -> Dict[str, Any]:
    e = json.loads(r["extra_json"] or "{}")
    if not e:                                    # vínculo manual/anotação: estilo pela situação e direção
        hip = r["situacao"] == "hipotese"
        detalhes = [x for x in (
            "hipótese" if hip else "confirmado", f"data {r['data_ref']}" if r["data_ref"] else "",
            f"confiança {r['confianca']}/5" if r["confianca"] else "", r["observacao"] or "",
            "anotação sobre vínculo da base" if r["origem"] == "anotacao" else "") if x]
        e = {"title": (r["rotulo"] or "vínculo") + "\n" + " · ".join(detalhes),
             "label": r["rotulo"] or "", "width": 1.8, "dashes": [6, 4] if hip else False,
             "font": {"color": "#EDEDED", "size": 10, "strokeWidth": 3, "strokeColor": "#121214", "align": "top"},
             "manter_setas": True,       # o iframe não troca a direção escolhida pelo analista
             "arrows": {"to": {"enabled": r["direcao"] == "de_para", "scaleFactor": 0.5, "type": "arrow"},
                        "from": {"enabled": r["direcao"] == "para_de", "scaleFactor": 0.5, "type": "arrow"}}}
    elif r["rotulo"]:
        e["label"] = r["rotulo"]
    if r["data_ref"]:
        e["datas"] = [r["data_ref"]]                 # o caminho cronológico do Find Path usa a data de referência
    e.setdefault("ocorrencias", r["n_ocorrencias"] or 1)
    e.update({"id": r["id"], "from": r["de_id"], "to": r["para_id"], "origem": r["origem"], "situacao": r["situacao"],
              "rotulo": r["rotulo"] or "", "direcao": r["direcao"], "data_ref": r["data_ref"] or "",
              "confianca": r["confianca"], "observacao": r["observacao"] or "", "ref_oficial": r["ref_oficial"] or ""})
    return e


def carregar_vinculo(id_sketch: str, id_vinculo: str) -> Optional[Dict[str, Any]]:
    """Um vínculo ativo no formato do vis (resposta das escritas: patch de uma aresta só)."""
    conn = database.get_db_connection()
    try:
        r = conn.execute("SELECT * FROM sketch_vinculos WHERE id = ? AND id_sketch = ? AND removido_em IS NULL",
                         (id_vinculo, id_sketch)).fetchone()
        return _aresta_vis(r) if r else None
    finally:
        conn.close()


def carregar_overlay(id_sketch: str, layout: str = "organico") -> Dict[str, Dict[str, Dict[str, Any]]]:
    """Nós e vínculos ATIVOS do sketch, no formato do vis (o mesmo do antigo nos_extras_json),
    com a posição salva no layout (quem não tem posição entra em órbita no iframe)."""
    conn = database.get_db_connection()
    try:
        nos, anotacoes = {}, {}
        for r in conn.execute(
                "SELECT * FROM sketch_entidades WHERE id_sketch = ? AND removido_em IS NULL ORDER BY criado_em, rowid", (id_sketch,)):
            # Anotação sobre nó oficial (nota, aparência, campos de fusão): o nó continua vindo do cluster.
            (anotacoes if r["origem"] == "anotacao_oficial" else nos)[r["node_id"]] = _no_vis(r)
        aliases = {r[0]: r[1] for r in conn.execute(
            "SELECT node_id_antigo, node_id_destino FROM sketch_aliases WHERE id_sketch = ?", (id_sketch,))}
        for r in conn.execute("SELECT node_id, x, y FROM sketch_posicoes WHERE id_sketch = ? AND layout = ?", (id_sketch, layout)):
            alvo = r["node_id"] if r["node_id"] in nos else aliases.get(r["node_id"])   # posição de absorvida: da prioritária
            if alvo in nos and (alvo == r["node_id"] or "x" not in nos[alvo]):
                nos[alvo]["x"], nos[alvo]["y"] = r["x"], r["y"]
        arestas = {r["id"]: _aresta_vis(r) for r in conn.execute(
            "SELECT * FROM sketch_vinculos WHERE id_sketch = ? AND removido_em IS NULL ORDER BY criado_em, rowid", (id_sketch,))}
        return {"nodes": nos, "edges": arestas, "anotacoes": anotacoes}
    finally:
        conn.close()


def carregar_campos(id_sketch: str, node_ids: Iterable[str]) -> Dict[str, dict]:
    """Campos (schema do tipo) das entidades ativas pedidas — inclui anotações sobre nós oficiais."""
    ids = list(node_ids)
    out: Dict[str, dict] = {}
    conn = database.get_db_connection()
    try:
        for i in range(0, len(ids), 500):
            parte = ids[i:i + 500]
            for r in conn.execute(
                    f"""SELECT node_id, campos_json FROM sketch_entidades WHERE id_sketch = ? AND removido_em IS NULL
                        AND node_id IN ({",".join("?" * len(parte))})""", (id_sketch, *parte)):
                out[r[0]] = json.loads(r[1] or "{}")
        return out
    finally:
        conn.close()


def contar_entidades(id_sketch: str) -> int:
    conn = database.get_db_connection()
    try:
        return conn.execute("""SELECT COUNT(*) FROM sketch_entidades WHERE id_sketch = ? AND removido_em IS NULL
                               AND origem <> 'anotacao_oficial'""", (id_sketch,)).fetchone()[0]
    finally:
        conn.close()


def listar_eventos(id_sketch: str, limite: int = 200) -> List[Dict[str, Any]]:
    conn = database.get_db_connection()
    try:
        return [dict(r) for r in conn.execute(
            "SELECT * FROM sketch_eventos WHERE id_sketch = ? ORDER BY id DESC LIMIT ?", (id_sketch, limite))]
    finally:
        conn.close()


# ------------------------------------------------------------------ migração do blob JSON
def _fazer_backup(motivo: str) -> Path:
    pasta = Path(database.DB_PATH).parent / "backups"
    pasta.mkdir(exist_ok=True)
    destino = pasta / f"banco_fraudes_antes_{motivo}_{datetime.now():%Y%m%d_%H%M%S}.db"
    origem = database.get_db_connection()
    try:
        bkp = sqlite3.connect(destino)
        try:
            origem.backup(bkp)
        finally:
            bkp.close()
    finally:
        origem.close()
    if not destino.exists() or destino.stat().st_size == 0:
        raise RuntimeError(f"backup não foi gravado em {destino}")
    return destino


def migrar_nos_extras(fazer_backup: bool = True) -> Dict[str, Any]:
    """
    Copia caso_sketches.nos_extras_json para as tabelas (idempotente: cada sketch é marcado
    extras_migrado = 1). Itens vindos do Enrich -> origem 'pivo', ids e arestas preservados.
    Formatos inesperados vão para o relatório, nunca descartados em silêncio. Backup antes;
    sem backup, não migra. A coluna JSON fica intacta (só leitura) por duas versões.
    """
    rel: Dict[str, Any] = {"sketches": 0, "nos": 0, "vinculos": 0, "ignorados": [], "backup": None}
    conn = database.get_db_connection()
    try:
        pend = conn.execute("""SELECT id_sketch, nos_extras_json FROM caso_sketches
                               WHERE COALESCE(extras_migrado, 0) = 0""").fetchall()
    finally:
        conn.close()
    if not pend:
        return rel
    lidos = []
    for r in pend:
        bruto = r["nos_extras_json"] or ""
        try:
            d = json.loads(bruto) if bruto.strip() else {}
        except Exception as e:
            rel["ignorados"].append({"id_sketch": r["id_sketch"], "motivo": f"JSON inválido: {e}"})
            lidos.append((r["id_sketch"], [], []))
            continue
        if not isinstance(d, dict):
            rel["ignorados"].append({"id_sketch": r["id_sketch"], "motivo": f"raiz {type(d).__name__}, esperado objeto"})
            lidos.append((r["id_sketch"], [], []))
            continue
        nos, arestas = [], []
        for k, v in (d.get("nodes") or {}).items() if isinstance(d.get("nodes"), dict) else []:
            if isinstance(v, dict) and (v.get("id") or k):
                nos.append({**v, "id": v.get("id") or k})
            else:
                rel["ignorados"].append({"id_sketch": r["id_sketch"], "item": str(k), "motivo": "nó sem formato de objeto"})
        for k, v in (d.get("edges") or {}).items() if isinstance(d.get("edges"), dict) else []:
            if isinstance(v, dict) and v.get("from") and v.get("to"):
                arestas.append({**v, "id": v.get("id") or k})
            else:
                rel["ignorados"].append({"id_sketch": r["id_sketch"], "item": str(k), "motivo": "aresta sem from/to"})
        for chave in ("nodes", "edges"):
            if chave in d and not isinstance(d[chave], dict):
                rel["ignorados"].append({"id_sketch": r["id_sketch"], "motivo": f"'{chave}' não é objeto"})
        lidos.append((r["id_sketch"], nos, arestas))

    if fazer_backup and any(nos or arestas for _, nos, arestas in lidos):
        rel["backup"] = str(_fazer_backup("migracao_sketch"))      # se falhar, levanta e não migra nada

    with transacao() as conn:
        for id_sketch, nos, arestas in lidos:
            if nos or arestas:
                r = adicionar(id_sketch, nos, arestas, origem="pivo", quem="migracao", acao="migrar_json", _conn=conn)
                rel["nos"] += len(r["nos_novos"])
                rel["vinculos"] += len(r["vinculos_novos"])
                rel["sketches"] += 1
            conn.execute("UPDATE caso_sketches SET extras_migrado = 1 WHERE id_sketch = ?", (id_sketch,))
    return rel
