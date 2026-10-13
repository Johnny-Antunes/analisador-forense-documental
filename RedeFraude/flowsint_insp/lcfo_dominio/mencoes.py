"""
Menções clicáveis nos pareceres (item 11).

Formato no HTML do parecer:
    <span class="lcfo-mencao" data-ref="TEL_11987654321" data-tipo="telefone" contenteditable="false">@(11) 98765-4321</span>
data-ref = node_id (registro de tipos); municípios MUN_<código IBGE>; sketches SKT_<id>.

  - converter_legado: o texto antigo "@CPF_..." vira menção NA LEITURA (o banco não é reescrito);
  - limpar_html: spans de menção só com class/data-ref/data-tipo/contenteditable/title (o resto do
    HTML segue a regra de sempre: DOMPurify no navegador);
  - retrolinks em parecer_mencoes (reconstruída a cada gravação), resolvendo aliases de fusão;
  - buscar: índice em memória (entidades do dossiê, municípios do IBGE, sketches), 8 por grupo.
"""

from __future__ import annotations

import html as _html
import re
import time
import unicodedata
from typing import Any, Dict, Iterable, List, Optional, Tuple

import database
from lcfo_dominio import tipos

MAX_POR_GRUPO = 8
_PREFIXOS_ESPECIAIS = {"MUN": "municipio", "SKT": "sketch"}
_SPAN = re.compile(r'<span\b[^>]*class="[^"]*\blcfo-mencao\b[^"]*"[^>]*>(.*?)</span>', re.I | re.S)
_ATTR = re.compile(r'\b(data-ref|data-tipo)="([^"]*)"', re.I)


def _norm(t: str) -> str:
    s = "".join(c for c in unicodedata.normalize("NFKD", str(t)) if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]", "", s.upper())


def tipo_do_ref(ref: str) -> str:
    prefixo = str(ref).split("_", 1)[0]
    if prefixo in _PREFIXOS_ESPECIAIS:
        return _PREFIXOS_ESPECIAIS[prefixo]
    t = tipos.tipo_do_id(ref)
    return t.chave if t else ""


def rotulo_padrao(ref: str) -> str:
    if ref.startswith("MUN_"):
        try:
            import territorio_engine as te
            m = te.carregar_municipios_ibge().get(ref[4:])
            return f"{m['nome']}/{m['uf']}" if m else ref
        except Exception:
            return ref
    if ref.startswith("SKT_"):
        return ref[4:]
    return tipos.formatar_id(ref)


def mencao_html(ref: str, rotulo: Optional[str] = None, tipo: Optional[str] = None) -> str:
    tipo = tipo or tipo_do_ref(ref)
    rot = rotulo or rotulo_padrao(ref)
    t = tipos.obter(tipo)
    dica = f"{t.rotulo if t else tipo.capitalize()}: {rot}"
    return (f'<span class="lcfo-mencao" data-ref="{_html.escape(ref, quote=True)}" data-tipo="{_html.escape(tipo, quote=True)}" '
            f'contenteditable="false" title="{_html.escape(dica, quote=True)}">@{_html.escape(rot)}</span>')


def converter_legado(texto: str) -> str:
    """@CPF_123..., @TEL_..., @PLACA_... (texto antigo) -> menção; o que já é span fica como está."""
    if not texto:
        return texto or ""
    prefixos = "|".join(sorted({t.prefixo for t in tipos.REGISTRO.values()} | set(_PREFIXOS_ESPECIAIS), key=len, reverse=True))
    padrao = re.compile(r'(?<![="\w])@((?:' + prefixos + r')_[A-Za-z0-9-]+)')
    partes = re.split(r'(<span\b[^>]*lcfo-mencao[^>]*>.*?</span>)', texto, flags=re.S | re.I)
    return "".join(p if p.lower().startswith("<span") else padrao.sub(lambda m: mencao_html(m.group(1)), p) for p in partes)


def limpar_html(texto: str) -> str:
    """Reconstrói cada span de menção só com os atributos permitidos (valores escapados)."""
    def _refazer(m):
        attrs = dict((k.lower(), _html.unescape(v)) for k, v in _ATTR.findall(m.group(0)))
        ref = attrs.get("data-ref", "")
        if not re.fullmatch(r"[A-Z]+_[A-Za-z0-9 ,.\-/]+", ref or ""):
            return _html.escape(re.sub(r"<[^>]+>", "", m.group(1)))      # menção inválida vira texto
        rot = _html.unescape(re.sub(r"<[^>]+>", "", m.group(1))).lstrip("@")
        return mencao_html(ref, rot, attrs.get("data-tipo") or None)
    return _SPAN.sub(_refazer, texto or "")


def preparar_para_exibir(texto: str) -> str:
    return limpar_html(converter_legado(texto or ""))


def extrair(texto: str) -> List[Tuple[str, str]]:
    vistos, out = set(), []
    for m in _SPAN.finditer(converter_legado(texto or "")):
        attrs = dict((k.lower(), _html.unescape(v)) for k, v in _ATTR.findall(m.group(0)))
        ref = attrs.get("data-ref")
        if ref and ref not in vistos:
            vistos.add(ref)
            out.append((ref, attrs.get("data-tipo") or tipo_do_ref(ref)))
    return out


# ------------------------------------------------------------------ retrolinks
def reconstruir_retrolinks(id_parecer: int, texto: str) -> int:
    refs = extrair(texto)
    conn = database.get_db_connection()
    try:
        conn.execute("DELETE FROM parecer_mencoes WHERE id_parecer = ?", (id_parecer,))
        conn.executemany("INSERT INTO parecer_mencoes (id_parecer, ref, tipo) VALUES (?,?,?)",
                         [(id_parecer, r, t) for r, t in refs])
        conn.commit()
    finally:
        conn.close()
    return len(refs)


def reindexar_todos() -> int:
    """Uma vez: indexa os pareceres antigos (sem reescrever o texto deles)."""
    conn = database.get_db_connection()
    try:
        rows = conn.execute("SELECT id, parecer FROM casos_historico_pareceres").fetchall()
    finally:
        conn.close()
    return sum(reconstruir_retrolinks(r[0], r[1] or "") for r in rows)


def resolver(ref: str) -> str:
    """Segue aliases de fusão (item 6): a menção a uma absorvida leva à prioritária."""
    conn = database.get_db_connection()
    try:
        for _ in range(5):
            r = conn.execute("SELECT node_id_destino FROM sketch_aliases WHERE node_id_antigo = ? LIMIT 1", (ref,)).fetchone()
            if not r or r[0] == ref:
                break
            ref = r[0]
        return ref
    finally:
        conn.close()


def citado_em(ref: str) -> List[Dict[str, Any]]:
    """Pareceres que citam a entidade (ou uma absorvida por ela numa fusão)."""
    conn = database.get_db_connection()
    try:
        apelidos = [r[0] for r in conn.execute("SELECT DISTINCT node_id_antigo FROM sketch_aliases WHERE node_id_destino = ?", (ref,))]
        refs = [ref] + apelidos
        marc = ",".join("?" * len(refs))
        return [dict(r) for r in conn.execute(
            f"""SELECT DISTINCT p.id, p.id_caso, p.titulo, p.data_registro FROM parecer_mencoes m
                JOIN casos_historico_pareceres p ON p.id = m.id_parecer WHERE m.ref IN ({marc}) ORDER BY p.id DESC""", refs)]
    finally:
        conn.close()


# ------------------------------------------------------------------ busca do autocomplete
_CACHE_MUN: Optional[List[Tuple[str, str, str]]] = None


def _indice_municipios() -> List[Tuple[str, str, str]]:
    global _CACHE_MUN
    if _CACHE_MUN is None:
        import territorio_engine as te
        _CACHE_MUN = [(_norm(m["nome"]) + cod, f"MUN_{cod}", f"{m['nome']}/{m['uf']}") for cod, m in te.carregar_municipios_ibge().items()]
    return _CACHE_MUN


def buscar(termo: str, entidades: Iterable[Tuple[str, str]], sketches: Iterable[Tuple[str, str]] = ()) -> Dict[str, List[Dict[str, str]]]:
    """
    entidades: (node_id, rótulo) do dossiê; sketches: (id, nome). Casa por trecho do rótulo, do
    valor formatado ou dos dígitos ("3548" acha o código IBGE e entidades com o trecho).
    """
    q = _norm(termo)
    if not q:
        return {"Entidades": [], "Municípios": [], "Sketches": []}
    ents = []
    for nid, rot in entidades:
        chave = _norm(f"{rot} {tipos.formatar_id(nid)} {tipos.valor_do_id(nid)}")
        if q in chave:
            ents.append({"ref": nid, "rotulo": rot or tipos.formatar_id(nid), "tipo": tipo_do_ref(nid),
                         "detalhe": tipos.formatar_id(nid) if rot and rot != tipos.formatar_id(nid) else ""})
            if len(ents) >= MAX_POR_GRUPO:
                break
    muns = [{"ref": ref, "rotulo": rot, "tipo": "municipio", "detalhe": ref[4:]}
            for chave, ref, rot in _indice_municipios() if q in chave][:MAX_POR_GRUPO]
    skts = [{"ref": f"SKT_{sid}", "rotulo": nome, "tipo": "sketch", "detalhe": sid} for sid, nome in sketches
            if q in _norm(f"{nome} {sid}")][:MAX_POR_GRUPO]
    return {"Entidades": ents, "Municípios": muns, "Sketches": skts}
