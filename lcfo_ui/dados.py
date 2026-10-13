"""
DataStore global (grafo + clusters + casos) com recarregamento integrado.
"""

from __future__ import annotations

from typing import Any, Dict, List

from lcfo_ui.servicos import obter_radar_expansoes, carregar_todos_casos_cadastrados, carregar_redes

import database


# =========================================================================
# DATASTORE COM RECARREGAMENTO INTEGRADO
# =========================================================================
class DataStore:
    def __init__(self):
        self.G = None
        self.cluster_info: List[Dict[str, Any]] = []
        self.casos_cadastrados: Dict[str, Any] = {}
        self.radar_alertas: Dict[str, int] = {}
        self.recarregar()

    def recarregar(self):
        self.G, cluster_info = carregar_redes()
        self.cluster_info = cluster_info or []
        self.radar_alertas = obter_radar_expansoes(self.cluster_info) if self.cluster_info else {}
        self.casos_cadastrados = carregar_todos_casos_cadastrados()

DADOS = DataStore()


def resolver_dossie_do_cluster(cluster_id: str) -> str:
    if cluster_id in DADOS.casos_cadastrados:
        return cluster_id
    conn = database.get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id_caso FROM caso_sketches WHERE cluster_origem_id = ? LIMIT 1", (cluster_id,))
        row = cursor.fetchone()
        return row["id_caso"] if row else cluster_id
    finally:
        conn.close()
