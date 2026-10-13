"""
Fusão de entidades (item 6): o analista funde duas ou mais entidades do mesmo tipo
escolhendo qual prevalece. Nunca automática (sugerir_duplicatas só sugere).

Regras:
  - mesmo tipo; dado oficial nunca é absorvido (pode ser a prioritária); duas oficiais não
    fundem — use vincular_mesma_entidade (vínculo hipótese MESMA_ENTIDADE);
  - vínculos das absorvidas são recriados na prioritária com o mesmo rótulo, direção,
    situação e data; só pares REALMENTE iguais (mesmo par, rótulo e direção) se unem:
    soma n_ocorrencias, data mais antiga, observações concatenadas; rótulos diferentes
    ficam (em curva); vínculo entre as próprias fundidas vira autolaço e é descartado
    (listado no relatório). Vínculo oficial do cluster não é tocado;
  - nada é apagado: absorvidas e vínculos substituídos vão para a lixeira no lote da fusão
    (fundida_em aponta a prioritária); os vínculos novos levam lote_fusao; os ids das
    absorvidas ficam em sketch_aliases (posições e menções continuam resolvendo);
  - desfazer_fusao devolve o grafo exatamente ao estado anterior.
"""

from __future__ import annotations

import difflib
import json
import uuid
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple

import database
from lcfo_dominio import sketch_repo as repo
from lcfo_dominio import tipos

PREFIXO_MOTIVO = "Fusão em "
MOTIVO_DESFEITA = "fusão desfeita"
ROTULO_MESMA_ENTIDADE = "MESMA_ENTIDADE"


class ErroFusao(ValueError):
    """Fusão recusada; `sugestao` = 'mesma_entidade' quando são duas oficiais."""

    def __init__(self, msg: str, sugestao: str = ""):
        super().__init__(msg)
        self.sugestao = sugestao


# ------------------------------------------------------------------ leitura
def _entidade(conn, id_sketch: str, node_id: str):
    return conn.execute("SELECT * FROM sketch_entidades WHERE id_sketch = ? AND node_id = ? AND removido_em IS NULL",
                        (id_sketch, node_id)).fetchone()


def _vinculos_ativos(conn, id_sketch: str, ids: Iterable[str]) -> List[Dict[str, Any]]:
    ids = list(ids)
    if not ids:
        return []
    marc = ",".join("?" * len(ids))
    return [dict(r) for r in conn.execute(
        f"""SELECT * FROM sketch_vinculos WHERE id_sketch = ? AND removido_em IS NULL
            AND (de_id IN ({marc}) OR para_id IN ({marc})) ORDER BY criado_em, rowid""", (id_sketch, *ids, *ids))]


def _chave_vinculo(de: str, para: str, rotulo: str, direcao: str) -> Tuple:
    """Iguais = mesmo par, rótulo e direção (sem direção, a ordem do par não importa)."""
    if direcao == "nenhuma":
        de, para = sorted((de, para))
    elif direcao == "para_de":
        de, para, direcao = para, de, "de_para"
    return (de, para, (rotulo or "").strip().lower(), direcao)


def _planejar(conn, id_sketch: str, prioritaria: str, absorvidas: List[str]) -> Dict[str, Any]:
    grupo = set(absorvidas) | {prioritaria}
    mapa = {a: prioritaria for a in absorvidas}
    tocados = _vinculos_ativos(conn, id_sketch, absorvidas)
    da_prior = [v for v in _vinculos_ativos(conn, id_sketch, [prioritaria]) if v["id"] not in {t["id"] for t in tocados}]
    descartados, grupos = [], {}
    for v in tocados:
        de, para = mapa.get(v["de_id"], v["de_id"]), mapa.get(v["para_id"], v["para_id"])
        if de == para:
            descartados.append(v["id"])                  # vínculo entre as próprias fundidas: autolaço
            continue
        grupos.setdefault(_chave_vinculo(de, para, v["rotulo"], v["direcao"]), {"novo": (de, para), "membros": []})["membros"].append(v)
    for v in da_prior:                                   # vínculos que a prioritária já tinha e ficam iguais
        k = _chave_vinculo(v["de_id"], v["para_id"], v["rotulo"], v["direcao"])
        if k in grupos:
            grupos[k]["membros"].insert(0, v)
    unidos = sum(1 for g in grupos.values() if len(g["membros"]) > 1)
    return {"grupo": grupo, "tocados": tocados, "grupos": grupos, "descartados": descartados,
            "redirecionados": len(tocados) - len(descartados), "unidos": unidos}


def previa(id_sketch: str, prioritaria: str, absorvidas: List[str]) -> Dict[str, int]:
    conn = database.get_db_connection()
    try:
        p = _planejar(conn, id_sketch, prioritaria, list(absorvidas))
    finally:
        conn.close()
    return {"redirecionados": p["redirecionados"], "unidos": p["unidos"], "descartados": len(p["descartados"])}


def escolher_prioritaria(candidatas: List[Dict[str, Any]]) -> str:
    """A mais completa (campos preenchidos), depois a com mais vínculos. Oficial tem precedência."""
    def chave(c):
        return (bool(c.get("oficial")), sum(1 for v in (c.get("campos") or {}).values() if v), c.get("n_vinculos", 0))
    return max(candidatas, key=chave)["node_id"]


# ------------------------------------------------------------------ fundir
def fundir(id_sketch: str, prioritaria: str, absorvidas: Iterable[str], *, oficiais: Set[str] = frozenset(),
           campos_escolhidos: Optional[Dict[str, Any]] = None, campos_oficiais: Optional[Dict[str, dict]] = None,
           quem: Optional[str] = None) -> Dict[str, Any]:
    """
    Funde `absorvidas` em `prioritaria`. `oficiais` = ids do cluster (não absorvíveis);
    `campos_escolhidos` = {campo: valor} que o analista escolheu na tabela campo a campo;
    `campos_oficiais` = campos conhecidos das oficiais (ex.: nome do titular, vindo do grafo).
    """
    absorvidas = [a for a in dict.fromkeys(absorvidas) if a != prioritaria]
    if not absorvidas:
        raise ErroFusao("Escolha pelo menos duas entidades.")
    todos = [prioritaria] + absorvidas
    tps = {(tipos.tipo_do_id(n).chave if tipos.tipo_do_id(n) else n.split("_", 1)[0]) for n in todos}
    if len(tps) > 1:
        raise ErroFusao("Só entidades do mesmo tipo podem ser fundidas.")
    oficiais_sel = [n for n in todos if n in oficiais]
    if len(oficiais_sel) > 1:
        raise ErroFusao("Duas entidades da base oficial não se fundem: vincule como mesma entidade.", sugestao="mesma_entidade")
    if any(a in oficiais for a in absorvidas):
        raise ErroFusao("Entidade da base oficial não pode ser absorvida: escolha-a como prioritária.")
    quem = quem if quem is not None else repo.usuario_atual()
    lote = uuid.uuid4().hex
    campos_oficiais = campos_oficiais or {}

    with repo.transacao() as conn:
        for a in absorvidas:
            if _entidade(conn, id_sketch, a) is None:
                raise ErroFusao(f"{a} não está ativa no sketch.")
        p = _planejar(conn, id_sketch, prioritaria, absorvidas)
        ts = repo.agora()

        # 1. campos: prioritária vence; vazios dela são preenchidos; escolhas do analista por cima
        linha_prior = _entidade(conn, id_sketch, prioritaria)
        campos_prior = json.loads(linha_prior["campos_json"] or "{}") if linha_prior else dict(campos_oficiais.get(prioritaria, {}))
        campos_antigos = json.loads(linha_prior["campos_json"] or "{}") if linha_prior else None
        novos_campos = dict(campos_prior)
        for a in absorvidas:
            for k, v in json.loads(_entidade(conn, id_sketch, a)["campos_json"] or "{}").items():
                if v and not novos_campos.get(k):
                    novos_campos[k] = v
        novos_campos.update({k: v for k, v in (campos_escolhidos or {}).items() if v is not None})
        anotacao_criada = False
        if linha_prior is not None:
            conn.execute("UPDATE sketch_entidades SET campos_json = ? WHERE id_sketch = ? AND node_id = ?",
                         (json.dumps(novos_campos, ensure_ascii=False), id_sketch, prioritaria))
        elif novos_campos != campos_prior:
            # prioritária oficial: o dado oficial não muda; os campos vão numa anotação por cima
            t = tipos.tipo_do_id(prioritaria)
            conn.execute(
                """INSERT INTO sketch_entidades (id_sketch, node_id, tipo, valor, campos_json, origem, criado_em, criado_por)
                   VALUES (?,?,?,?,?,'anotacao_oficial',?,?)
                   ON CONFLICT(id_sketch, node_id) DO UPDATE SET campos_json = excluded.campos_json""",
                (id_sketch, prioritaria, t.chave if t else "empresa", tipos.valor_do_id(prioritaria),
                 json.dumps(novos_campos, ensure_ascii=False), ts, quem))
            anotacao_criada = True

        # 2. vínculos: recriados na prioritária (lote_fusao); os substituídos vão para a lixeira no lote
        motivo = f"{PREFIXO_MOTIVO}{prioritaria}"
        criados, substituidos = [], []
        mapa = {a: prioritaria for a in absorvidas}
        for g in p["grupos"].values():
            membros = g["membros"]
            base = membros[0]
            de, para = mapa.get(base["de_id"], base["de_id"]), mapa.get(base["para_id"], base["para_id"])
            obs = " | ".join(dict.fromkeys(m["observacao"] for m in membros if m["observacao"]))
            datas = [m["data_ref"] for m in membros if m["data_ref"]]
            vid = f"vf_{uuid.uuid4().hex[:12]}"
            conn.execute(
                """INSERT INTO sketch_vinculos (id, id_sketch, de_id, para_id, rotulo, direcao, situacao, data_ref, confianca,
                   observacao, n_ocorrencias, origem, ref_oficial, id_execucao, lote_fusao, extra_json, criado_em, criado_por)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (vid, id_sketch, de, para, base["rotulo"], base["direcao"], base["situacao"], min(datas) if datas else None, base["confianca"],
                 obs, sum(int(m["n_ocorrencias"] or 1) for m in membros), base["origem"], base["ref_oficial"],
                 base["id_execucao"], lote, _extra_redirecionado(base["extra_json"]), ts, quem))
            criados.append(vid)
            substituidos += [m["id"] for m in membros]
        substituidos += p["descartados"]
        for i in range(0, len(substituidos), 500):
            parte = substituidos[i:i + 500]
            conn.execute(f"""UPDATE sketch_vinculos SET removido_em = ?, removido_por = ?, motivo_remocao = ?, lote_remocao = ?
                             WHERE id_sketch = ? AND id IN ({",".join("?" * len(parte))})""", (ts, quem, motivo, lote, id_sketch, *parte))

        # 3. absorvidas: lixeira (fundida_em) + aliases (posições e menções seguem resolvendo)
        for a in absorvidas:
            conn.execute("""UPDATE sketch_entidades SET removido_em = ?, removido_por = ?, motivo_remocao = ?, lote_remocao = ?,
                            fundida_em = ? WHERE id_sketch = ? AND node_id = ?""", (ts, quem, motivo, lote, prioritaria, id_sketch, a))
            conn.execute("UPDATE sketch_aliases SET node_id_destino = ? WHERE id_sketch = ? AND node_id_destino = ?",
                         (prioritaria, id_sketch, a))
            conn.execute("INSERT OR REPLACE INTO sketch_aliases (id_sketch, node_id_antigo, node_id_destino) VALUES (?,?,?)",
                         (id_sketch, a, prioritaria))

        detalhe = {"lote": lote, "prioritaria": prioritaria, "absorvidas": absorvidas, "campos": novos_campos,
                   "campos_antigos": campos_antigos, "anotacao_criada": anotacao_criada,
                   "vinculos_substituidos": substituidos, "vinculos_criados": criados, "descartados": p["descartados"]}
        repo.registrar_evento(conn, id_sketch, "fusao", "entidade", prioritaria, detalhe, quem)
    return {"lote": lote, "prioritaria": prioritaria, "absorvidas": absorvidas, "vinculos_criados": criados,
            "vinculos_removidos": substituidos, "descartados": p["descartados"], "redirecionados": p["redirecionados"],
            "unidos": p["unidos"], "campos": novos_campos}


def _extra_redirecionado(extra_json: str) -> str:
    """Estilo herdado da aresta (enrich), sem o título que citava as pontas antigas."""
    e = json.loads(extra_json or "{}")
    if e:
        e.pop("smooth", None)
    return json.dumps(e, ensure_ascii=False)


def eh_lote_de_fusao(motivo: str) -> bool:
    return bool(motivo) and motivo.startswith(PREFIXO_MOTIVO)


def desfazer_fusao(id_sketch: str, lote: str, quem: Optional[str] = None) -> Dict[str, Any]:
    """Absorvidas e vínculos originais voltam; os criados pela fusão saem; campos e aliases revertem."""
    quem = quem if quem is not None else repo.usuario_atual()
    with repo.transacao() as conn:
        ev = conn.execute("""SELECT detalhe_json FROM sketch_eventos WHERE id_sketch = ? AND acao = 'fusao' AND alvo_tipo = 'entidade'
                             AND detalhe_json LIKE ? ORDER BY id DESC LIMIT 1""", (id_sketch, f'%"lote": "{lote}"%')).fetchone()
        if ev is None:
            raise ErroFusao("Fusão não encontrada.")
        d = json.loads(ev[0])
        ts = repo.agora()
        criados = [r[0] for r in conn.execute("SELECT id FROM sketch_vinculos WHERE id_sketch = ? AND lote_fusao = ? AND removido_em IS NULL",
                                              (id_sketch, lote))]
        # os vínculos da fusão saem (não vão para a lixeira: não são escolha do analista, são o lote desfeito)
        conn.execute(f"""UPDATE sketch_vinculos SET removido_em = ?, removido_por = ?, motivo_remocao = ?, lote_remocao = ?
                         WHERE id_sketch = ? AND lote_fusao = ? AND removido_em IS NULL""",
                     (ts, quem, MOTIVO_DESFEITA, f"desfeita_{lote}", id_sketch, lote))
        limpar = "removido_em = NULL, removido_por = NULL, motivo_remocao = NULL, lote_remocao = NULL"
        nos = [r[0] for r in conn.execute("SELECT node_id FROM sketch_entidades WHERE id_sketch = ? AND lote_remocao = ?", (id_sketch, lote))]
        vins = [r[0] for r in conn.execute("SELECT id FROM sketch_vinculos WHERE id_sketch = ? AND lote_remocao = ?", (id_sketch, lote))]
        conn.execute(f"UPDATE sketch_entidades SET {limpar}, fundida_em = NULL WHERE id_sketch = ? AND lote_remocao = ?", (id_sketch, lote))
        conn.execute(f"UPDATE sketch_vinculos SET {limpar} WHERE id_sketch = ? AND lote_remocao = ?", (id_sketch, lote))
        if d.get("campos_antigos") is not None:
            conn.execute("UPDATE sketch_entidades SET campos_json = ? WHERE id_sketch = ? AND node_id = ?",
                         (json.dumps(d["campos_antigos"], ensure_ascii=False), id_sketch, d["prioritaria"]))
        if d.get("anotacao_criada"):
            conn.execute("""UPDATE sketch_entidades SET removido_em = ?, removido_por = ?, motivo_remocao = ?, lote_remocao = ?
                            WHERE id_sketch = ? AND node_id = ? AND origem = 'anotacao_oficial'""",
                         (ts, quem, MOTIVO_DESFEITA, f"desfeita_{lote}", id_sketch, d["prioritaria"]))
        for a in d["absorvidas"]:
            conn.execute("DELETE FROM sketch_aliases WHERE id_sketch = ? AND node_id_antigo = ?", (id_sketch, a))
        repo.registrar_evento(conn, id_sketch, "desfazer_fusao", "entidade", d["prioritaria"],
                              {"lote": lote, "restaurados": nos, "vinculos_restaurados": vins, "vinculos_removidos": criados}, quem)
    return {"ids_nos": nos, "ids_vinculos": vins, "vinculos_removidos": criados, "prioritaria": d["prioritaria"]}


def vincular_mesma_entidade(id_sketch: str, a: str, b: str, quem: Optional[str] = None) -> str:
    """Duas oficiais que parecem a mesma: vínculo hipótese, cada uma mantém sua evidência."""
    return repo.criar_vinculo(id_sketch, a, b, rotulo=ROTULO_MESMA_ENTIDADE, direcao="nenhuma", situacao="hipotese",
                              observacao="possível mesma entidade (fusão de duas oficiais recusada)", quem=quem)


def resolver_id(id_sketch: str, node_id: str) -> str:
    """Id atual de uma entidade (segue os aliases deixados por fusões)."""
    conn = database.get_db_connection()
    try:
        r = conn.execute("SELECT node_id_destino FROM sketch_aliases WHERE id_sketch = ? AND node_id_antigo = ?",
                         (id_sketch, node_id)).fetchone()
        return r[0] if r else node_id
    finally:
        conn.close()


# ------------------------------------------------------------------ sugestões de duplicatas
MAX_NOMES_FUZZY = 500          # comparação de nomes é quadrática: só em listas pequenas


def _pares_codigos(lista: List[str]) -> Dict[Tuple[str, str], str]:
    """Pares por um caractere trocado ou dois adjacentes transpostos, em O(n·L) (chave com coringa)."""
    pares: Dict[Tuple[str, str], str] = {}
    por_valor = {tipos.valor_do_id(n): n for n in lista}
    baldes: Dict[str, List[str]] = {}
    for v in por_valor:
        for i in range(len(v)):
            baldes.setdefault(f"{len(v)}:{i}:{v[:i]}*{v[i + 1:]}", []).append(v)
    for grupo in baldes.values():
        for i, a in enumerate(grupo):
            for b in grupo[i + 1:]:
                pares.setdefault(tuple(sorted((por_valor[a], por_valor[b]))), "um caractere trocado")
    for v in por_valor:
        for i in range(len(v) - 1):
            w = v[:i] + v[i + 1] + v[i] + v[i + 2:]
            if w != v and w in por_valor:
                pares.setdefault(tuple(sorted((por_valor[v], por_valor[w]))), "dois caracteres transpostos")
    return pares


def sugerir_duplicatas(id_sketch: str, ids: Iterable[str], tipo: Optional[str] = None, limiar: float = 0.88,
                       campos: Optional[Dict[str, dict]] = None, descartados: Iterable[Tuple[str, str]] = ()) -> List[Dict[str, str]]:
    """
    Pares para REVISÃO (nunca fundidos sozinhos): placas/telefones/CPFs que diferem por um
    caractere ou uma transposição; endereços e clínicas com nome parecido no mesmo município.
    Dois telefones transpostos são, na perícia, linhas distintas: o par só é sugerido.
    """
    campos = campos or {}
    vistos = {tuple(sorted(p)) for p in descartados}
    por_tipo: Dict[str, List[str]] = {}
    for n in dict.fromkeys(ids):
        t = tipos.tipo_do_id(n)
        if t and (tipo is None or t.chave == tipo):
            por_tipo.setdefault(t.chave, []).append(n)
    out = []
    for chave, lista in por_tipo.items():
        if chave in ("placa", "telefone", "cpf"):
            for (a, b), motivo in sorted(_pares_codigos(lista).items()):
                if (a, b) not in vistos:
                    out.append({"a": a, "b": b, "tipo": chave, "motivo": motivo})
        elif chave in ("endereco", "clinica", "pet", "empresa", "prestador") and len(lista) <= MAX_NOMES_FUZZY:
            for i, a in enumerate(lista):
                for b in lista[i + 1:]:
                    if tuple(sorted((a, b))) in vistos:
                        continue
                    ca, cb = campos.get(a, {}), campos.get(b, {})
                    if (ca.get("cidade") or "").upper() != (cb.get("cidade") or "").upper():
                        continue
                    r = difflib.SequenceMatcher(None, tipos.valor_do_id(a), tipos.valor_do_id(b)).ratio()
                    if r >= limiar:
                        x, y = sorted((a, b))
                        out.append({"a": x, "b": y, "tipo": chave,
                                    "motivo": f"nome parecido ({r:.0%})" + (f" em {ca.get('cidade')}" if ca.get("cidade") else "")})
    return out


def descartar_sugestao(id_sketch: str, a: str, b: str, quem: Optional[str] = None) -> None:
    with repo.transacao() as conn:
        repo.registrar_evento(conn, id_sketch, "descartar_duplicata", "par", f"{a}|{b}", {"a": a, "b": b}, quem)


def pares_descartados(id_sketch: str) -> List[Tuple[str, str]]:
    conn = database.get_db_connection()
    try:
        return [tuple(r[0].split("|", 1)) for r in conn.execute(
            "SELECT alvo_id FROM sketch_eventos WHERE id_sketch = ? AND acao = 'descartar_duplicata'", (id_sketch,))]
    finally:
        conn.close()
