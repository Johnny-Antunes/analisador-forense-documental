"""
Componente de grafo (iframe vis.js + transporte de payload via arquivo) e montagem do payload.
"""

from __future__ import annotations

import json
import re
import uuid
from pathlib import Path
from typing import Any, Dict, Optional

from nicegui import app, ui

from lcfo_ui.config import FONT_PADRAO, ESPACAMENTO_PADRAO
from lcfo_ui.servicos import carregar_posicoes_layout_caso, processar_subgrafo_caso, processar_grafo_dataframe_em_cache
from lcfo_ui.dados import DADOS

from lcfo_ui.estado import SessionState


_DIR_COMPONENTE_GRAFO = Path(__file__).resolve().parent.parent / "componente_grafo"
_DIR_PAYLOADS_GRAFO = _DIR_COMPONENTE_GRAFO / "_payloads"
if _DIR_COMPONENTE_GRAFO.exists():
    app.add_static_files('/lcfo_grafo_assets', str(_DIR_COMPONENTE_GRAFO))
    # Pasta onde os payloads do grafo (nós/arestas/posições) são gravados como
    # arquivos estáticos, para o iframe buscar via fetch() em vez de recebê-los
    # embutidos numa mensagem WebSocket — ver GrafoBidirecionalNiceGUI abaixo.
    _DIR_PAYLOADS_GRAFO.mkdir(exist_ok=True)


# =========================================================================
# MINIATURA DE GRAFO EM SVG (PREVIEW PARA SKETCHES)
# =========================================================================
def gerar_svg_miniatura_cluster(c_obj: Optional[Dict[str, Any]]) -> str:
    if not c_obj or not c_obj.get("nodes"):
        return '<svg width="100%" height="76" viewBox="0 0 300 76" xmlns="http://www.w3.org/2000/svg"></svg>'

    pontos = [
        (65, 52, "#FFB800", 6),
        (130, 26, "#EDEDED", 5),
        (185, 48, "#EDEDED", 5),
        (235, 30, "#3C6FA8", 4.5),
        (95, 30, "#B94A3C", 4),
        (155, 60, "#6D4FA8", 4),
    ]
    conexoes = [(0, 1), (1, 2), (2, 3), (0, 4), (1, 5)]

    svg_linhas = "".join(
        f'<line x1="{pontos[i][0]}" y1="{pontos[i][1]}" x2="{pontos[j][0]}" y2="{pontos[j][1]}" stroke="#3A3A40" stroke-width="1.5" />'
        for i, j in conexoes
    )
    svg_pontos = "".join(
        f'<circle cx="{p[0]}" cy="{p[1]}" r="{p[3]}" fill="{p[2]}" stroke="#18181A" stroke-width="1.5" />'
        for p in pontos
    )
    return f'''
    <svg width="100%" height="76" viewBox="0 0 300 76" xmlns="http://www.w3.org/2000/svg" style="display:block; overflow:visible;">
        {svg_linhas}
        {svg_pontos}
    </svg>
    '''


class GrafoBidirecionalNiceGUI:
    def __init__(self, chave: str):
        self.chave = chave
        self.sufixo_seguro = re.sub(r'[^a-zA-Z0-9_-]', '_', chave)
        self.iframe_id = f"grafo-iframe-{self.sufixo_seguro}-{uuid.uuid4().hex[:6]}"

    def _gravar_payload_e_obter_url(self, payload: Dict[str, Any]) -> str:
        """
        Grava o payload do grafo (nós/arestas/posições) em um arquivo estático
        em componente_grafo/_payloads/ em vez de embuti-lo no comando JS
        enviado via ui.run_javascript (WebSocket). Isso evita o teto de
        tamanho de mensagem do Socket.IO (~1 MB por padrão), que descartava a
        mensagem em silêncio em grafos grandes (~1000 nós) — o grafo carregava
        com o fundo vazio, sem erro nenhum. Também é mais rápido de processar
        no navegador: o conteúdo chega como dado (fetch + JSON.parse nativo)
        em vez de precisar ser interpretado como texto-fonte de um script
        gigante.

        O nome do arquivo é fixo por chave de grafo (self.sufixo_seguro) —
        cada atualização sobrescreve o mesmo arquivo, sem acumular lixo em
        disco a cada refresh; o cache-busting (?v=) garante que o fetch()
        sempre pegue a versão recém-gravada, não uma cópia em cache do
        navegador.

        Limitação conhecida, aceitável neste uso (ferramenta de analista
        rodando localmente): se dois processos abrirem o mesmo caso ao mesmo
        tempo, o arquivo é compartilhado pela mesma chave — a última gravação
        prevalece até o próximo fetch de cada um.
        """
        caminho_arquivo = _DIR_PAYLOADS_GRAFO / f"{self.sufixo_seguro}.json"
        with open(caminho_arquivo, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
        cache_bust = uuid.uuid4().hex[:8]
        return f"/lcfo_grafo_assets/_payloads/{self.sufixo_seguro}.json?v={cache_bust}"

    def render(self, payload: Dict[str, Any]):
        cache_bust = uuid.uuid4().hex[:8]
        ui.html(
            f'<iframe id="{self.iframe_id}" src="/lcfo_grafo_assets/index.html?v={cache_bust}" '
            f'style="position:absolute; inset:0; width:100%; height:100%; border:none; display:block;"></iframe>',
            sanitize=False
        ).classes('absolute inset-0 w-full h-full')
        payload_url = self._gravar_payload_e_obter_url(payload)
        ui.run_javascript(
            f"window.lcfoRegistrarGrafoIframe('{self.chave}', '{self.iframe_id}', {json.dumps(payload_url)});"
        )
        return self

    def enviar_atualizacao(self, payload: Dict[str, Any]):
        payload_url = self._gravar_payload_e_obter_url(payload)
        payload_url_json = json.dumps(payload_url)
        ui.run_javascript(f"""
            (function() {{
                const el = document.getElementById('{self.iframe_id}');
                if (el && el.contentWindow) {{
                    el.contentWindow.postMessage({{ type: 'streamlit:render_url', url: {payload_url_json} }}, '*');
                }}
            }})();
        """)

    def sincronizar_selecao(self, ids):
        """
        Envia ao iframe a seleção feita na gaveta lateral (Entities), sem
        disparar um re-render completo do grafo — usa um tipo de mensagem
        próprio ('streamlit:selection') que o index.html trata separadamente
        do fluxo normal de payload. É isso que faz o botão Find Path do
        componente reconhecer 2 entidades marcadas na sidebar, e não só 2 nós
        clicados diretamente no canvas ou via retângulo/laço. Mensagem
        sempre pequena (poucos IDs) — não passa pelo mecanismo de arquivo,
        pois nunca teve risco de esbarrar no limite do WebSocket.
        """
        ids_seguros = json.dumps(list(ids)[-6:])
        ui.run_javascript(f"""
            (function() {{
                const el = document.getElementById('{self.iframe_id}');
                if (el && el.contentWindow) {{
                    el.contentWindow.postMessage({{ type: 'streamlit:selection', ids: {ids_seguros} }}, '*');
                }}
            }})();
        """)


# =========================================================================
# MONTAGEM DO PAYLOAD DO GRAFO
# =========================================================================
def montar_payload_grafo(st: SessionState, identificador_caso, tem_cluster, cluster_obj, df_dados, resumo_descoberta, layout_ativo, chave_layout: Optional[str] = None):
    chave_layout_efetiva = chave_layout or identificador_caso
    if tem_cluster:
        vis_nodes, vis_edges = processar_subgrafo_caso(
            G=DADOS.G, cluster_nodes=cluster_obj["nodes"], filtro_tipos=["telefone", "cpf", "placa"],
            hub_id=cluster_obj["hub_id"], font_slider=FONT_PADRAO, id_caso=chave_layout_efetiva,
            espacamento=ESPACAMENTO_PADRAO
        )
        hub_id = cluster_obj["hub_id"]
        if layout_ativo == "arvore":
            pos_banco = carregar_posicoes_layout_caso(chave_layout_efetiva, layout_tipo="arvore")
            if pos_banco:
                for n in vis_nodes:
                    if n["id"] in pos_banco:
                        n["x"] = pos_banco[n["id"]]["x"]
                        n["y"] = pos_banco[n["id"]]["y"]
    else:
        vis_nodes, vis_edges, hub_id = processar_grafo_dataframe_em_cache(
            df_dados, alvo_principal=(resumo_descoberta or {}).get("termo_limpo", ""),
            font_slider=FONT_PADRAO, espacamento=ESPACAMENTO_PADRAO
        )

    balde = st.nos_enriquecidos.get(identificador_caso, {"nodes": {}, "edges": {}})
    ids_presentes = {n["id"] for n in vis_nodes}
    vis_nodes = vis_nodes + [n for nid, n in balde["nodes"].items() if nid not in ids_presentes]
    ids_arestas = {e.get("id") for e in vis_edges if e.get("id")}
    vis_edges = vis_edges + [e for eid, e in balde["edges"].items() if eid not in ids_arestas]

    payload = {
        "nodes": vis_nodes, "edges": vis_edges, "hub_id": hub_id,
        "target_node_id": st.entidade_foco or "", "layout_ativo": layout_ativo,
        "base_font_size": FONT_PADRAO,
    }
    return payload, vis_nodes, vis_edges, hub_id
