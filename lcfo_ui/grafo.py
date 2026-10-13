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

from lcfo_dominio import pivos, sketch_repo, tipos
from lcfo_ui import overlay
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

    def enviar_mensagem(self, mensagem: Dict[str, Any]):
        """Mensagem pequena Python -> iframe (respostas do quick-add, patches)."""
        ui.run_javascript(f"""
            (function() {{
                const el = document.getElementById('{self.iframe_id}');
                if (el && el.contentWindow) el.contentWindow.postMessage({json.dumps(mensagem, ensure_ascii=False)}, '*');
            }})();
        """)

    def enviar_patch(self, patch: Dict[str, Any]):
        """
        Alteração incremental no grafo aberto ({add, update, remove} de nós e arestas), sem
        reenviar o payload inteiro nem refazer a vista. É a resposta do Python às escritas do
        analista (remover, restaurar, criar...): o iframe nunca grava, só pede.
        """
        self.enviar_mensagem({"type": "lcfo:graph_patch", "patch": patch})

    def sincronizar_selecao(self, ids):
        """
        Envia ao iframe a seleção feita na gaveta lateral (Entities), sem
        disparar um re-render completo do grafo — usa um tipo de mensagem
        próprio ('streamlit:selection') que o index.html trata separadamente
        do fluxo normal de payload. É isso que faz o botão Find Path do
        componente reconhecer 2 entidades marcadas na sidebar, e não só 2 nós
        clicados diretamente no canvas ou via retângulo/laço.

        Envia TODOS os ids marcados (antes: list(set)[-6:], ou seja, 6 ids
        aleatórios — "selecionar todos" destacava só 6 nós no grafo). O
        iframe destaca todos e usa os últimos para o Find Path. Mesmo com
        ~1.000 ids a mensagem fica na casa de dezenas de KB.
        """
        ids_seguros = json.dumps(sorted(ids))
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
def campos_filtraveis(vis_nodes) -> list:
    """Campos que as regras de filtro podem usar; as listas trazem só os valores presentes."""
    from lcfo_dominio import pivos as _pivos
    presentes = {n.get("tipo") for n in vis_nodes}
    origens = sorted({n.get("origem") or "oficial" for n in vis_nodes})
    execs = sorted({n["id_execucao"] for n in vis_nodes if n.get("id_execucao")})
    return [
        {"campo": "tipo", "rotulo": "Tipo", "tipo": "lista",
         "opcoes": [{"valor": t.chave, "rotulo": t.rotulo} for t in tipos.REGISTRO.values() if t.chave in presentes]},
        {"campo": "rotulo", "rotulo": "Rótulo", "tipo": "texto"},
        {"campo": "origem", "rotulo": "Origem", "tipo": "lista",
         "opcoes": [{"valor": o, "rotulo": {"oficial": "Base oficial", "manual": "Manual", "pivo": "Pivô/Enrich", "importacao": "Importação"}.get(o, o)} for o in origens]},
        {"campo": "flag", "rotulo": "Flag", "tipo": "lista",
         "opcoes": [{"valor": k, "rotulo": k} for k in sketch_repo.FLAGS]},
        {"campo": "grau", "rotulo": "Grau (conexões)", "tipo": "numero"},
        {"campo": "vinculos", "rotulo": "Total de vínculos", "tipo": "numero"},
        {"campo": "intermediacao", "rotulo": "Intermediação", "tipo": "numero"},
        {"campo": "comunidade", "rotulo": "Comunidade", "tipo": "numero"},
        {"campo": "ocorrencias", "rotulo": "Nº de ocorrências", "tipo": "numero"},
        {"campo": "primeira_data", "rotulo": "Primeira data", "tipo": "data"},
        {"campo": "ultima_data", "rotulo": "Última data", "tipo": "data"},
        {"campo": "execucao", "rotulo": "Execução de origem", "tipo": "lista",
         "opcoes": [{"valor": e, "rotulo": e} for e in execs]},
    ]


def montar_payload_grafo(st: SessionState, identificador_caso, tem_cluster, cluster_obj, df_dados, resumo_descoberta, layout_ativo, chave_layout: Optional[str] = None):
    chave_layout_efetiva = chave_layout or identificador_caso
    if tem_cluster:
        vis_nodes, vis_edges = processar_subgrafo_caso(
            G=DADOS.G, cluster_nodes=cluster_obj["nodes"], filtro_tipos=list(tipos.REGISTRO),
            hub_id=cluster_obj["hub_id"], font_slider=FONT_PADRAO, id_caso=chave_layout_efetiva,
            espacamento=ESPACAMENTO_PADRAO
        )
        hub_id = cluster_obj["hub_id"]
        if layout_ativo == "arvore":
            pos_banco = carregar_posicoes_layout_caso(chave_layout_efetiva, layout_tipo="arvore")    # legado (antes do item 8)
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

    balde = overlay.carregar(st, identificador_caso)      # tabelas do sketch (cache na sessão)
    ids_presentes = {n["id"] for n in vis_nodes}
    vis_nodes = vis_nodes + [n for nid, n in balde["nodes"].items() if nid not in ids_presentes]
    ids_arestas = {e.get("id") for e in vis_edges if e.get("id")}
    vis_edges = vis_edges + [e for eid, e in balde["edges"].items() if eid not in ids_arestas]

    # Anotações sobre nós oficiais (nota, aparência, flag): o nó vem do cluster, o estilo vem da anotação.
    for n in vis_nodes:
        an = balde.get("anotacoes", {}).get(n["id"])
        if an and (an.get("aparencia") or an.get("tem_nota")):
            n.update(sketch_repo.estilo_no(n.get("tipo", ""), an.get("aparencia"), is_hub=(n["id"] == hub_id),
                                           tem_nota=an.get("tem_nota", False)))

    # Posições arrumadas à mão (todos os layouts) ficam em sketch_posicoes e vencem as calculadas.
    posicoes_por_layout: Dict[str, Dict[str, Any]] = {}
    if identificador_caso and identificador_caso != "DESCOBERTA":
        for lay in ("organico", "arvore", "leque"):
            pos = sketch_repo.carregar_posicoes(identificador_caso, lay)
            if pos:
                posicoes_por_layout[lay] = pos
        for n in vis_nodes:
            p = posicoes_por_layout.get(layout_ativo, {}).get(n["id"])
            if p:
                n["x"], n["y"] = p["x"], p["y"]
                if p.get("fixado"):
                    n["fixado"] = True

    payload = {
        "nodes": vis_nodes, "edges": vis_edges, "hub_id": hub_id,
        "posicoes_por_layout": posicoes_por_layout,
        "rotulos_flags": sketch_repo.rotulos_flags(identificador_caso) if identificador_caso != "DESCOBERTA" else {},
        # filtros por regra (item 12): campos gerados a partir do registro de tipos e os filtros salvos
        "campos_filtraveis": campos_filtraveis(vis_nodes),
        "operadores_filtro": {k: list(v) for k, v in sketch_repo.OPERADORES.items()},
        "filtros_salvos": sketch_repo.listar_filtros(identificador_caso) if identificador_caso != "DESCOBERTA" else [],
        "target_node_id": st.entidade_foco or "", "layout_ativo": layout_ativo,
        "base_font_size": FONT_PADRAO,
        # Registro de tipos (chave, rótulo, cor, ícone): o iframe monta filtros e legenda a partir dele.
        "tipos": tipos.payload_tipos({n.get("tipo") for n in vis_nodes}),
        "sugestoes_vinculo": tipos.payload_sugestoes(),
        "pivos": pivos.payload_catalogo(),                      # menu do nó: pivôs por tipo (item 7)        # rótulos sugeridos por par de tipos (item 3)
    }
    return payload, vis_nodes, vis_edges, hub_id
