"""
Roteamento entre telas (dashboard, overview do dossiê, mesa do sketch, descoberta).
"""

from __future__ import annotations


from nicegui import ui

from lcfo_ui.servicos import carregar_sketch, carregar_nos_extras_sketch
from lcfo_ui.dados import resolver_dossie_do_cluster

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina


def criar_navegacao(ctx: "Pagina") -> None:
    """Cria as funções de navegação e as publica em ctx."""
    st = ctx.st

    # -------------------------------------------------------------------
    # ROTEAMENTO
    # -------------------------------------------------------------------
    def abrir_caso_overview(cid: str):
        cid_resolvido = resolver_dossie_do_cluster(cid)
        st.caso_ativo_id = cid_resolvido
        st.sketch_ativo_id = None
        st.subtela_caso = "overview"
        st.dados_descoberta = None
        st.entidade_foco = None
        st.cockpit_ativo = False
        ctx.dlg_palette.close()
        ctx.navbar_breadcrumbs.refresh()
        ctx.navbar_acoes.refresh()
        ctx.sidebar_container.refresh()
        ctx.workspace.refresh()

    def abrir_mesa_sketch(id_sketch: str):
        sketch = carregar_sketch(id_sketch)
        if not sketch:
            ui.notify("Sketch não encontrado.", type="negative")
            return
        if id_sketch not in st.nos_enriquecidos:
            nos_extra, arestas_extra = carregar_nos_extras_sketch(id_sketch)
            st.nos_enriquecidos[id_sketch] = {"nodes": nos_extra, "edges": arestas_extra}

        st.caso_ativo_id = sketch["id_caso"]
        st.sketch_ativo_id = id_sketch
        st.subtela_caso = "mesa"
        st.painel_esquerdo = "Grafo de Vínculos"
        st.painel_unico = "Grafo de Vínculos"
        st.entidade_foco = None
        ctx.navbar_breadcrumbs.refresh()
        ctx.navbar_acoes.refresh()
        ctx.sidebar_container.refresh()
        ctx.workspace.refresh()

    def voltar_ao_dashboard():
        st.caso_ativo_id = None
        st.sketch_ativo_id = None
        st.dados_descoberta = None
        st.contexto_grafo_ativo = {}
        st.cockpit_ativo = False
        ctx.navbar_breadcrumbs.refresh()
        ctx.navbar_acoes.refresh()
        ctx.sidebar_container.refresh()
        ctx.workspace.refresh()

    def abrir_descoberta(resumo, df, ents):
        st.dados_descoberta = {"resumo": resumo, "df": df, "entidades": ents}
        st.caso_ativo_id = None
        st.sketch_ativo_id = None
        st.subtela_caso = "mesa"
        st.painel_esquerdo = "Grafo de Vínculos"
        st.painel_unico = "Grafo de Vínculos"
        st.cockpit_ativo = False
        ctx.navbar_breadcrumbs.refresh()
        ctx.navbar_acoes.refresh()
        ctx.sidebar_container.refresh()
        ctx.workspace.refresh()

    def alternar_notas():
        st.mostrar_notas = not st.mostrar_notas
        ctx.navbar_acoes.refresh()
        ctx.workspace.refresh()

    def alternar_sidebar():
        st.sidebar_visivel = not st.sidebar_visivel
        ctx.sidebar_container.refresh()

    ctx.abrir_caso_overview = abrir_caso_overview
    ctx.abrir_mesa_sketch = abrir_mesa_sketch
    ctx.voltar_ao_dashboard = voltar_ao_dashboard
    ctx.abrir_descoberta = abrir_descoberta
    ctx.alternar_notas = alternar_notas
    ctx.alternar_sidebar = alternar_sidebar
