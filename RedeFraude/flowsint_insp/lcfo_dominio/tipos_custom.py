"""
Tipos de entidade customizados (item 13): o analista cria um tipo (ex.: "Oficina", "Clínica
veterinária") sem mexer em código. Fica numa tabela e é registrado no registro único
(lcfo_dominio/tipos.py) na subida do app e a cada gravação — daí em diante quick-add, colar,
filtros, legenda, fusão e pivôs o tratam como qualquer outro tipo (critério do item 1:
tipo novo = uma entrada no registro, nenhuma edição em index.html, sidebar ou config).

Travas: chave e prefixo não colidem com os tipos de fábrica nem entre si; depois que o tipo
tem entidades em algum sketch, a identidade (chave, prefixo, normalização) não muda e o
tipo não pode ser excluído — só desativado.
"""

from __future__ import annotations

import json
import re
from datetime import datetime
from typing import Any, Dict, List, Optional

import database
from lcfo_dominio import tipos
from lcfo_dominio.icones import PICTOGRAMAS
from lcfo_dominio.tipos import Campo, TipoEntidade

NORMALIZACOES = {"texto": "Texto (maiúsculas, sem acento)", "digitos": "Só dígitos", "livre": "Como digitado (sem espaços nas pontas)"}
FABRICA = set(tipos.REGISTRO)          # os tipos de fábrica (registrados no import de tipos.py)
_RE_CHAVE = re.compile(r"^[a-z][a-z0-9_]{1,30}$")
_RE_PREFIXO = re.compile(r"^[A-Z][A-Z0-9]{1,7}$")


def _normalizador(modo: str):
    if modo == "digitos":
        return lambda v: re.sub(r"\D", "", str(v))
    if modo == "livre":
        return lambda v: str(v).strip()
    return tipos._texto_canonico


def _validador(regex: str):
    if not regex:
        return lambda v: bool(v)
    r = re.compile(regex)
    return lambda v: bool(v) and bool(r.fullmatch(v))


def _detector(regex: str, confianca: int):
    if not regex:
        return lambda t: 0
    r = re.compile(regex, re.I)
    return lambda t: confianca if r.fullmatch(str(t).strip()) else 0


def montar_tipo(d: Dict[str, Any]) -> TipoEntidade:
    campos = tuple(Campo(c["nome"], c.get("rotulo") or c["nome"], bool(c.get("obrigatorio")), c.get("tipo") or "texto",
                         bool(c.get("primario"))) for c in d.get("campos") or [])
    if not any(c.primario for c in campos):
        campos = (Campo("identificacao", "Identificação", True, primario=True),) + campos
    return TipoEntidade(
        chave=d["chave"], prefixo=d["prefixo"], rotulo=d["rotulo"], cor=(d["cor_fundo"], d["cor_borda"]),
        icone=d.get("icone") if d.get("icone") in PICTOGRAMAS else "documento", campos=campos,
        normalizar=_normalizador(d.get("normalizacao") or "texto"), validar=_validador(d.get("regex_validacao") or ""),
        detectar=_detector(d.get("regex_deteccao") or "", int(d.get("confianca") or 60)),
        relacoes=tuple(x.strip() for x in (d.get("relacoes") or "").split(",") if x.strip()))


def validar_definicao(d: Dict[str, Any], existente: Optional[Dict[str, Any]] = None) -> List[str]:
    erros = []
    if not _RE_CHAVE.match(d.get("chave") or ""):
        erros.append("chave: minúsculas, dígitos e _, começando por letra (ex.: oficina)")
    if not _RE_PREFIXO.match(d.get("prefixo") or ""):
        erros.append("prefixo: 2 a 8 letras/dígitos maiúsculos (ex.: OFC)")
    if not (d.get("rotulo") or "").strip():
        erros.append("rótulo obrigatório")
    for k in ("cor_fundo", "cor_borda"):
        if not re.fullmatch(r"#[0-9A-Fa-f]{6}", d.get(k) or ""):
            erros.append(f"{k}: cor no formato #RRGGBB")
    if d.get("chave") in FABRICA:
        erros.append("chave já usada por um tipo de fábrica")
    prefixos_fabrica = {tipos.REGISTRO[k].prefixo for k in FABRICA if k in tipos.REGISTRO}
    if d.get("prefixo") in prefixos_fabrica:
        erros.append("prefixo já usado por um tipo de fábrica")
    for outro in listar():
        if existente and outro["chave"] == existente["chave"]:
            continue
        if outro["chave"] == d.get("chave"):
            erros.append("já existe um tipo customizado com essa chave")
        if outro["prefixo"] == d.get("prefixo"):
            erros.append("já existe um tipo customizado com esse prefixo")
    for k in ("regex_validacao", "regex_deteccao"):
        if d.get(k):
            try:
                re.compile(d[k])
            except re.error as e:
                erros.append(f"{k}: expressão inválida ({e})")
    if existente and em_uso(existente["chave"]):
        for k in ("chave", "prefixo", "normalizacao"):
            if (d.get(k) or "") != (existente.get(k) or ""):
                erros.append(f"{k} não pode mudar: o tipo já tem entidades em sketches (mudaria os ids)")
    return erros


# ------------------------------------------------------------------ persistência
def listar(incluir_inativos: bool = True) -> List[Dict[str, Any]]:
    conn = database.get_db_connection()
    try:
        rows = conn.execute("SELECT * FROM tipos_customizados ORDER BY criado_em").fetchall()
    finally:
        conn.close()
    out = [{**json.loads(r["definicao_json"]), "ativo": bool(r["ativo"])} for r in rows]
    return out if incluir_inativos else [d for d in out if d["ativo"]]


def em_uso(chave: str) -> int:
    conn = database.get_db_connection()
    try:
        return conn.execute("SELECT COUNT(*) FROM sketch_entidades WHERE tipo = ?", (chave,)).fetchone()[0]
    finally:
        conn.close()


def salvar(d: Dict[str, Any], usuario: str = "") -> None:
    existente = next((x for x in listar() if x["chave"] == d.get("chave")), None)
    erros = validar_definicao(d, existente)
    if erros:
        raise ValueError("; ".join(erros))
    d = {k: d.get(k) for k in ("chave", "prefixo", "rotulo", "cor_fundo", "cor_borda", "icone", "campos", "normalizacao",
                               "regex_validacao", "regex_deteccao", "confianca", "relacoes")}
    agora = datetime.now().isoformat(timespec="seconds")
    conn = database.get_db_connection()
    try:
        conn.execute("""INSERT INTO tipos_customizados (chave, definicao_json, ativo, criado_em, criado_por, atualizado_em)
                        VALUES (?,?,1,?,?,?) ON CONFLICT(chave) DO UPDATE SET definicao_json = excluded.definicao_json,
                        atualizado_em = excluded.atualizado_em""",
                     (d["chave"], json.dumps(d, ensure_ascii=False), agora, usuario, agora))
        conn.commit()
    finally:
        conn.close()
    registrar_todos()


def definir_ativo(chave: str, ativo: bool) -> None:
    conn = database.get_db_connection()
    try:
        conn.execute("UPDATE tipos_customizados SET ativo = ? WHERE chave = ?", (int(ativo), chave))
        conn.commit()
    finally:
        conn.close()
    registrar_todos()


def excluir(chave: str) -> None:
    n = em_uso(chave)
    if n:
        raise ValueError(f"o tipo tem {n} entidade(s) em sketches: desative em vez de excluir")
    conn = database.get_db_connection()
    try:
        conn.execute("DELETE FROM tipos_customizados WHERE chave = ?", (chave,))
        conn.commit()
    finally:
        conn.close()
    registrar_todos()


def registrar_todos() -> int:
    """(Re)registra os ativos no registro e tira os inativos/excluídos (os de fábrica nunca saem)."""
    ativos = {d["chave"]: d for d in listar(incluir_inativos=False)}
    for chave in [k for k in list(tipos.REGISTRO) if k not in FABRICA and k not in ativos]:
        t = tipos.REGISTRO.pop(chave)
        tipos._POR_PREFIXO.pop(t.prefixo, None)
    for d in ativos.values():
        antigo = tipos.REGISTRO.get(d["chave"])
        if antigo:
            tipos._POR_PREFIXO.pop(antigo.prefixo, None)
        tipos.registrar(montar_tipo(d))
    return len(ativos)
