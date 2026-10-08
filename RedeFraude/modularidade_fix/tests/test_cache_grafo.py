"""
O grafo deve ser reaproveitado enquanto `assistencias` não muda, e
recalculado quando muda (ingestão / promoção / fusão). Não grava no banco:
simula a mudança de versão.
"""

from __future__ import annotations

import graph_engine


def test_mesma_versao_reaproveita_o_grafo():
    g1, _ = graph_engine.carregar_redes()
    g2, _ = graph_engine.carregar_redes()
    assert g1 is g2


def test_versao_nova_recalcula_o_grafo(monkeypatch):
    g1, _ = graph_engine.carregar_redes()
    versao = graph_engine.obter_versao_assistencias()
    monkeypatch.setattr(graph_engine, "obter_versao_assistencias", lambda: (versao[0] + 1, versao[1] + 1))
    g2, _ = graph_engine.carregar_redes()
    assert g1 is not g2
    assert g1.number_of_nodes() == g2.number_of_nodes()  # mesmos dados -> mesmo grafo, só que recalculado
