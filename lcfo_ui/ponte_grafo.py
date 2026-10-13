"""
Recepção de mensagens JS -> Python vindas do iframe do grafo.
"""

from __future__ import annotations

import json
from typing import Any, Dict

from nicegui import ui

from lcfo_ui.config import TEM_ENRICH_ENGINE, enriquecer_entidade_local
from lcfo_ui.servicos import resetar_layout_caso, salvar_posicoes_layout_caso
from lcfo_dominio import pivos, sketch_repo
from lcfo_ui import detalhes_ui, fusao_ui, lixeira, overlay, pivos_ui, quick_add, vinculos

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
        # "Refazer tudo" no orgânico: as posições arrumadas à mão também saem (as fixadas ficam)
        sketch_repo.limpar_posicoes(contexto["identificador_caso"], "organico")
        st.layout_por_caso[contexto["identificador_caso"]] = "organico"
        novo_payload = contexto["montar_payload"]("organico")
        contexto["grafo"].enviar_atualizacao(novo_payload)
        ui.notify("Layout recalculado (Orgânico).")

    def tratar_salvar_posicoes(valor: Dict[str, Any], contexto: Dict[str, Any]):
        """Lote de posições (só o que mudou), 1,2 s depois do último arraste. Responde ao indicador."""
        identificador = contexto["identificador_caso"]
        rev = valor.get("rev")
        if identificador == overlay.DESCOBERTA:
            contexto["grafo"].enviar_mensagem({"type": "lcfo:positions_saved", "rev": rev, "ignorado": True})
            return
        posicoes = {str(k): {"x": float(v["x"]), "y": float(v["y"])} for k, v in (valor.get("posicoes") or {}).items()
                    if not str(k).startswith(("tmp_", "draft_"))}
        try:
            sketch_repo.salvar_posicoes(identificador, valor.get("layout") or "organico", posicoes)
            contexto["grafo"].enviar_mensagem({"type": "lcfo:positions_saved", "rev": rev})
        except Exception as e:
            contexto["grafo"].enviar_mensagem({"type": "lcfo:positions_saved", "rev": rev, "erro": str(e)[:200]})

    def tratar_fixar(valor: Dict[str, Any], contexto: Dict[str, Any]):
        identificador = contexto["identificador_caso"]
        if identificador == overlay.DESCOBERTA:
            ui.notify("Na descoberta as posições não são salvas: promova a caso.", type="warning")
            return
        posicoes = {str(k): {"x": float(v["x"]), "y": float(v["y"])} for k, v in (valor.get("posicoes") or {}).items()}
        sketch_repo.fixar(identificador, posicoes, valor.get("layout") or "organico", bool(valor.get("fixado")))

    def tratar_filtros(valor: Dict[str, Any], contexto: Dict[str, Any]):
        """Filtros por regra salvos por sketch (a avaliação é toda no navegador)."""
        identificador = contexto["identificador_caso"]
        if identificador == overlay.DESCOBERTA:
            ui.notify("Na descoberta os filtros não são salvos: promova a caso.", type="warning")
            return
        try:
            if valor.get("tipo") == "filter_save":
                sketch_repo.salvar_filtro(identificador, valor.get("nome") or "Filtro", valor.get("filtro") or {})
                ui.notify(f"Filtro '{valor.get('nome') or 'Filtro'}' salvo neste sketch.", type="positive")
            else:
                sketch_repo.remover_filtro(identificador, valor.get("id") or "")
        except ValueError as e:
            ui.notify(str(e), type="warning")
        contexto["grafo"].enviar_mensagem({"type": "lcfo:filtros_salvos", "filtros": sketch_repo.listar_filtros(identificador)})

    def tratar_enrich(valor: Dict[str, Any], contexto: Dict[str, Any]):
        if not TEM_ENRICH_ENGINE:
            ui.notify("Módulo enrich_engine.py não localizado.")
            return
        tipo_entidade = valor.get("tipoEntidade", "")
        node_id = valor.get("nodeId", "")
        if contexto["identificador_caso"] != overlay.DESCOBERTA and tipo_entidade in pivos.CATALOGO["vinculos_diretos"].entrada:
            # Sketch salvo: o Enrich é o pivô "Vínculos diretos" com os padrões — fora do event loop, com log e proveniência.
            pivos_ui.rodar(ctx, contexto, "vinculos_diretos", [{"node_id": node_id, "tipo": tipo_entidade, "campos": {}}], {})
            return
        valor_bruto = node_id.split("_", 1)[1] if "_" in node_id else valor.get("valor", "")
        resultado = enriquecer_entidade_local(tipo_entidade, valor_bruto)
        if resultado.get("total_encontrado", 0) == 0:
            ui.notify("Nenhum vínculo adicional localizado no acervo local.")
            return

        # Grava no sketch (tabelas, com origem e evento); na descoberta, só na memória.
        novos_efetivos = overlay.adicionar_resultado_enrich(st, contexto["identificador_caso"], resultado)

        if novos_efetivos > 0:
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
            # Só o cartão de foco (antes: a gaveta inteira, que perdia busca e rolagem a cada clique).
            (ctx.sidebar_parcial.get("foco") or contexto["sidebar_refresh"])()
        elif tipo_ev == "save_layout_positions":
            layout = valor.get("layout", "organico")
            st.layout_por_caso[contexto["identificador_caso"]] = layout
            posicoes = valor.get("positions")
            if posicoes and layout == "arvore":
                alvo_salvar = contexto.get("chave_layout") or contexto["identificador_caso"]
                salvar_posicoes_layout_caso(alvo_salvar, posicoes, layout_tipo="arvore")
        elif tipo_ev == "save_positions":
            tratar_salvar_posicoes(valor, contexto)
        elif tipo_ev == "pin_nodes":
            tratar_fixar(valor, contexto)
        elif tipo_ev == "clear_layout_positions" and contexto["identificador_caso"] != overlay.DESCOBERTA:
            sketch_repo.limpar_posicoes(contexto["identificador_caso"], valor.get("layout") or "arvore")
        elif tipo_ev == "reset_layout_request":
            tratar_reset_layout(valor, contexto)
        elif tipo_ev == "enrich_request":
            tratar_enrich(valor, contexto)
        elif tipo_ev == "quick_add_detect":
            quick_add.tratar_detect(ctx, contexto, valor)
        elif tipo_ev == "quick_add_commit":
            quick_add.tratar_commit(ctx, contexto, valor)
        elif tipo_ev == "paste_request":
            quick_add.tratar_colar(ctx, contexto, valor)
        elif tipo_ev == "edge_create":
            vinculos.tratar_criar(ctx, contexto, valor)
        elif tipo_ev == "edge_update":
            vinculos.tratar_atualizar(ctx, contexto, valor)
        elif tipo_ev == "pivot_request" and valor.get("pivo") in pivos.CATALOGO:
            pivos_ui.abrir_folha(ctx, contexto, valor["pivo"], list(valor.get("nodeIds") or []))
        elif tipo_ev == "node_select" and valor.get("node_id"):
            detalhes_ui.abrir_painel(ctx, contexto, valor["node_id"])
        elif tipo_ev == "entity_update" and valor.get("node_id"):
            detalhes_ui.tratar_atualizacao(ctx, contexto, valor)
        elif tipo_ev in ("filter_save", "filter_delete"):
            tratar_filtros(valor, contexto)
        elif tipo_ev == "merge_request":
            fusao_ui.abrir_dialogo_fusao(ctx, contexto, list(valor.get("ids") or []))
        elif tipo_ev == "remove_request":
            lixeira.abrir_dialogo_remocao(ctx, contexto, list(valor.get("nodeIds") or []), list(valor.get("edgeIds") or []))
        elif tipo_ev == "isolate_request" and valor.get("nodeId"):
            lixeira.abrir_dialogo_isolar(ctx, contexto, valor["nodeId"])
        elif tipo_ev == "clear_selection":
            # Qualquer clique no canvas (nó ou espaço vazio) avisa o Python
            # para esvaziar a seleção por checkbox da sidebar — mesmo
            # comportamento de "clicar fora limpa" que já existia para o
            # foco de entidade única, agora também para a seleção múltipla.
            sel = st.selecao_entidades.get(contexto["identificador_caso"])
            if sel:
                sel.clear()
                (ctx.sidebar_parcial.get("selecao") or contexto["sidebar_refresh"])()
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
            (ctx.sidebar_parcial.get("selecao") or contexto["sidebar_refresh"])()

    bridge = ui.element('div').props('id=lcfo-graph-bridge').style('display:none')
    bridge.on('graph-message', _ao_receber_mensagem_grafo, args=['detail'])
