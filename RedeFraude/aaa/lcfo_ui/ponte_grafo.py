"""
Recepção de mensagens JS -> Python vindas do iframe do grafo.
"""

from __future__ import annotations

import json
from typing import Any, Dict

from nicegui import ui

from lcfo_ui.config import TEM_ENRICH_ENGINE, enriquecer_entidade_local
from lcfo_ui.servicos import resetar_layout_caso, salvar_posicoes_layout_caso, salvar_nos_extras_sketch

from typing import TYPE_CHECKING
if TYPE_CHECKING:
    from lcfo_ui.contexto import Pagina


def instalar_receptor_grafo(ctx: "Pagina") -> None:
    """Dispatcher das mensagens do iframe do grafo (clique, layout, enrich, seleção)."""
    st = ctx.st

    # -------------------------------------------------------------------
    # DISPATCHER DE MENSAGENS VINDAS DO IFRAME DO GRAFO (JS -> Python)
    # -------------------------------------------------------------------
    def tratar_reset_layout(valor: Dict[str, Any], contexto: Dict[str, Any]):
        if not contexto["eh_oficial"]:
            return
        alvo_reset = contexto.get("chave_layout") or contexto["identificador_caso"]
        resetar_layout_caso(alvo_reset)
        st.layout_por_caso[contexto["identificador_caso"]] = "organico"
        novo_payload = contexto["montar_payload"]("organico")
        contexto["grafo"].enviar_atualizacao(novo_payload)
        ui.notify("Layout recalculado (Orgânico).")

    def tratar_enrich(valor: Dict[str, Any], contexto: Dict[str, Any]):
        if not TEM_ENRICH_ENGINE:
            ui.notify("Módulo enrich_engine.py não localizado.")
            return
        tipo_entidade = valor.get("tipoEntidade", "")
        node_id = valor.get("nodeId", "")
        valor_bruto = node_id.split("_", 1)[1] if "_" in node_id else valor.get("valor", "")
        resultado = enriquecer_entidade_local(tipo_entidade, valor_bruto)
        if resultado.get("total_encontrado", 0) == 0:
            ui.notify("Nenhum vínculo adicional localizado no acervo local.")
            return

        balde = st.nos_enriquecidos.setdefault(contexto["identificador_caso"], {"nodes": {}, "edges": {}})
        novos_efetivos = 0
        for n in resultado.get("novos_nos", []):
            if n["id"] not in balde["nodes"]:
                balde["nodes"][n["id"]] = n
                novos_efetivos += 1
        for e_novo in resultado.get("novas_arestas", []):
            balde["edges"][e_novo["id"]] = e_novo

        if novos_efetivos > 0:
            if contexto["identificador_caso"] != "DESCOBERTA":
                salvar_nos_extras_sketch(contexto["identificador_caso"], balde["nodes"], balde["edges"])
            ui.notify(f"Enrich: {novos_efetivos} nova(s) entidade(s) incorporada(s).")
            layout_atual = st.layout_por_caso.get(contexto["identificador_caso"], "organico")
            novo_payload = contexto["montar_payload"](layout_atual)
            contexto["grafo"].enviar_atualizacao(novo_payload)
            contexto["sidebar_refresh"]()
        else:
            ui.notify("As entidades relacionadas já estão presentes.")

    def _ao_receber_mensagem_grafo(e):
        raw = e.args
        if isinstance(raw, list) and raw:
            raw = raw[0]
        if isinstance(raw, dict) and "detail" in raw:
            raw = raw["detail"]
        try:
            dados = json.loads(raw) if isinstance(raw, str) else raw
        except Exception:
            return
        if not isinstance(dados, dict):
            return

        chave = dados.get("key")
        valor = dados.get("value") or {}

        # Outros iframes que usam a mesma ponte (ex.: o mapa territorial) registram um receptor.
        receptor = ctx.receptores_iframe.get(chave)
        if receptor:
            receptor(valor)
            return

        contexto = None
        for cg in st.contexto_grafo_ativo.values():
            if cg.get("chave_grafo") == chave:
                contexto = cg
                break
        if not contexto:
            return

        tipo_ev = valor.get("tipo")
        if tipo_ev == "node_click":
            st.entidade_foco = valor.get("nodeId")
            contexto["sidebar_refresh"]()
        elif tipo_ev == "save_layout_positions":
            layout = valor.get("layout", "organico")
            st.layout_por_caso[contexto["identificador_caso"]] = layout
            posicoes = valor.get("positions")
            if posicoes and layout == "arvore":
                alvo_salvar = contexto.get("chave_layout") or contexto["identificador_caso"]
                salvar_posicoes_layout_caso(alvo_salvar, posicoes, layout_tipo="arvore")
        elif tipo_ev == "reset_layout_request":
            tratar_reset_layout(valor, contexto)
        elif tipo_ev == "enrich_request":
            tratar_enrich(valor, contexto)
        elif tipo_ev == "clear_selection":
            # Qualquer clique no canvas (nó ou espaço vazio) avisa o Python
            # para esvaziar a seleção por checkbox da sidebar — mesmo
            # comportamento de "clicar fora limpa" que já existia para o
            # foco de entidade única, agora também para a seleção múltipla.
            sel = st.selecao_entidades.get(contexto["identificador_caso"])
            if sel:
                sel.clear()
            contexto["sidebar_refresh"]()
        elif tipo_ev == "sync_selection_from_canvas":
            # Seleção feita direto no canvas (retângulo, laço, Ctrl+A ou
            # shift+click) — sincroniza de volta para a sidebar: o card
            # "X selecionado(s)" e o Set st.selecao_entidades passam a
            # refletir o que foi marcado no grafo, não só os checkboxes da
            # lista Entities.
            ids_recebidos = valor.get("ids") or []
            sel = st.selecao_entidades.setdefault(contexto["identificador_caso"], set())
            sel.clear()
            sel.update(ids_recebidos)
            contexto["sidebar_refresh"]()

    bridge = ui.element('div').props('id=lcfo-graph-bridge').style('display:none')
    bridge.on('graph-message', _ao_receber_mensagem_grafo, args=['detail'])
