"""
Overlay do sketch na sessão: nós e vínculos que o analista trouxe (Enrich, Add...).

A fonte da verdade são as tabelas do sketch (lcfo_dominio/sketch_repo.py);
st.nos_enriquecidos é só o cache do overlay lido de lá, recarregado após cada escrita.
A mesa de uma DESCOBERTA ainda não é um sketch gravado: o overlay dela vive só na memória
até a promoção a caso.
"""

from __future__ import annotations

from typing import Any, Dict, Optional

from lcfo_dominio import sketch_repo
from lcfo_ui.estado import SessionState

DESCOBERTA = "DESCOBERTA"


def _persistido(identificador: str) -> bool:
    return bool(identificador) and identificador != DESCOBERTA


def carregar(st: SessionState, identificador: str, forcar: bool = False) -> Dict[str, Any]:
    """Overlay do sketch (do banco, com cache na sessão)."""
    if _persistido(identificador) and (forcar or identificador not in st.nos_enriquecidos):
        st.nos_enriquecidos[identificador] = sketch_repo.carregar_overlay(identificador)
    return st.nos_enriquecidos.setdefault(identificador, {"nodes": {}, "edges": {}})


def adicionar_resultado_enrich(st: SessionState, identificador: str, res: Dict[str, Any],
                               manual: Optional[str] = None) -> int:
    """
    Incorpora o resultado de enriquecer_entidade_local. `manual` = node_id que o analista
    pediu (Add): esse entra com origem 'manual'; os vizinhos, 'pivo'. Devolve quantas
    entidades eram novas (ativas que já estavam não contam).
    """
    nos, arestas = res.get("novos_nos", []), res.get("novas_arestas", [])
    if not _persistido(identificador):
        balde = st.nos_enriquecidos.setdefault(identificador, {"nodes": {}, "edges": {}})
        novos = 0
        for n in nos:
            if n["id"] not in balde["nodes"]:
                balde["nodes"][n["id"]] = n
                novos += 1
        for e in arestas:
            balde["edges"].setdefault(e["id"], e)
        return novos
    r = sketch_repo.adicionar(identificador, nos, arestas, origem="pivo",
                              origem_por_no={manual: "manual"} if manual else None)
    carregar(st, identificador, forcar=True)
    return len(r["nos_novos"]) + len([x for x in r["reativados"] if not str(x).startswith("enr_")])


def limpar(st: SessionState, identificador: str, motivo: str = "Limpar Enrich") -> Dict[str, Any]:
    """Tira do sketch tudo que veio de Enrich/Add — para a lixeira (reversível), não mais apagado."""
    if not _persistido(identificador):
        balde = st.nos_enriquecidos.pop(identificador, None) or {"nodes": {}}
        return {"lote": None, "nos": len(balde["nodes"]), "vinculos": 0}
    r = sketch_repo.remover(identificador, origens=("pivo", "manual", "importacao"), motivo=motivo)
    carregar(st, identificador, forcar=True)
    return r


def contar(identificador: str) -> int:
    return sketch_repo.contar_entidades(identificador) if _persistido(identificador) else 0
