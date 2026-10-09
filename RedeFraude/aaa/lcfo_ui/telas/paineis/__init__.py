"""
Registro dos painéis da mesa. A chave é o nome exibido na toolbar
(config.OPCOES_TOOLBAR); o valor é a função render(m, lado) do módulo.
"""

from __future__ import annotations

from lcfo_ui.telas.contexto_mesa import ContextoMesa
from lcfo_ui.telas.paineis import expansao, ferramentas, grafo, mapa, tabela, temporal

PAINEIS = {
    "Grafo de Vínculos": grafo.render,
    "Tabela de Ocorrências": tabela.render,
    "Radar Territorial (Mapa)": mapa.render,
    "Evolução Temporal": temporal.render,
    "Expansão com Criações": expansao.render,
    "Ferramentas": ferramentas.render,
}


def renderizar_painel(m: ContextoMesa, nome_painel: str, lado: str = "unico") -> None:
    render = PAINEIS.get(nome_painel)
    if render:
        render(m, lado)
