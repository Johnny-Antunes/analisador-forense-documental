"""
Painel da mesa: Grafo de Vínculos.
"""

from __future__ import annotations


from lcfo_ui.grafo import GrafoBidirecionalNiceGUI, montar_payload_grafo

from lcfo_ui.telas.contexto_mesa import ContextoMesa


def render(m: ContextoMesa, lado: str) -> None:
    ctx = m.ctx
    st = m.ctx.st
    identificador_caso = m.identificador_caso
    dossie_persistido = m.dossie_persistido
    cluster_obj = m.cluster_obj
    df_dados = m.df_dados
    resumo_descoberta = m.resumo_descoberta
    chave_layout = m.chave_layout
    payload = m.payload
    chave_grafo = f"{identificador_caso}_{lado}"
    grafo = GrafoBidirecionalNiceGUI(chave=chave_grafo)
    grafo.render(payload)

    def _montar_payload_atual(layout: str):
        p, _, _, _ = montar_payload_grafo(st, 
            identificador_caso, cluster_obj is not None, cluster_obj, df_dados, resumo_descoberta, layout,
            chave_layout=chave_layout
        )
        return p

    st.contexto_grafo_ativo[lado] = {
        "identificador_caso": identificador_caso,
        "chave_layout": chave_layout,
        "chave_grafo": chave_grafo,
        "eh_oficial": dossie_persistido,
        "montar_payload": _montar_payload_atual,
        "sidebar_refresh": ctx.sidebar_container.refresh,
        "grafo": grafo,
    }
