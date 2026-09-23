"""
Módulo: enrich_engine.py
Objetivo: Implementação do protocolo de Enrichers locais.
          Permite a expansão sob demanda de novos vínculos a partir de uma entidade,
          gerando nós circulares com ícones SVG sem prefixos colchetes e preservando o pai orbital.
"""

from typing import Dict, Any, List, Tuple
import pandas as pd
import urllib.parse

from database import get_db_connection
from utils import formatar_cpf_cnpj, formatar_tel
from graph_engine import _criar_svg_nodo_flowsint

PREFIXO_POR_TIPO = {
    "cpf": "CPF",
    "telefone": "TEL",
    "placa": "PLACA",
}

COLUNA_POR_TIPO = {
    "cpf": "cpf",
    "telefone": "telefone",
    "placa": "placa",
}

CORES_POR_TIPO = {
    "cpf": {"bg": "#0369A1", "border": "#38BDF8"},
    "telefone": {"bg": "#991B1B", "border": "#EF4444"},
    "placa": {"bg": "#6D28D9", "border": "#A78BFA"},
    "empresa": {"bg": "#1E293B", "border": "#64748B"},
}


def _normalizar_tipo(tipo_bruto: str) -> str:
    return str(tipo_bruto).strip().lower().replace("[", "").replace("]", "")


def _montar_no_empresa(empresa: str, grau: int, node_pai_id: str) -> Dict[str, Any]:
    cor = CORES_POR_TIPO["empresa"]
    svg_icon = _criar_svg_nodo_flowsint("empresa", cor["bg"], cor["border"], is_hub=False)
    return {
        "id": f"EMPRESA_{empresa.upper()}",
        "label": empresa.upper(),
        "title": f"SEGURADORA CLIENTE: {empresa.upper()}\nDescoberto via Enrich",
        "tipo": "empresa",
        "shape": "image",
        "image": svg_icon,
        "size": 20,
        "degree": grau,
        "betweenness": 0.0,
        "community": 0,
        "valor": empresa.upper(),
        "nome_titular": "",
        "orbit_ao_redor": node_pai_id,
        "color": {
            "background": cor["bg"], "border": cor["border"]
        },
        "font": {"color": "#F8FAFC", "size": 11, "face": "Segoe UI", "strokeWidth": 2.5, "strokeColor": "#06090F"}
    }


def _montar_no_entidade(tipo: str, valor: str, grau: int, node_pai_id: str) -> Dict[str, Any]:
    cor = CORES_POR_TIPO[tipo]
    prefixo = PREFIXO_POR_TIPO[tipo]

    if tipo == "cpf":
        valor_fmt = formatar_cpf_cnpj(valor)
    elif tipo == "telefone":
        valor_fmt = formatar_tel(valor)
    else:
        valor_fmt = f"{valor[:3]}-{valor[3:]}" if len(valor) == 7 else valor

    svg_icon = _criar_svg_nodo_flowsint(tipo, cor["bg"], cor["border"], is_hub=False)

    return {
        "id": f"{prefixo}_{valor}",
        "label": valor_fmt,
        "title": f"{tipo.upper()}: {valor_fmt}\nDescoberto via Enrich",
        "tipo": tipo,
        "shape": "image",
        "image": svg_icon,
        "size": 20,
        "degree": grau,
        "betweenness": 0.0,
        "community": 0,
        "valor": valor_fmt,
        "nome_titular": "",
        "orbit_ao_redor": node_pai_id,
        "color": {
            "background": cor["bg"], "border": cor["border"]
        },
        "font": {"color": "#F8FAFC", "size": 11, "face": "Segoe UI", "strokeWidth": 2.5, "strokeColor": "#06090F"}
    }


def enriquecer_entidade_local(tipo_entidade: str, valor_identificador: str) -> Dict[str, Any]:
    """Expande os vínculos de uma entidade buscando novos CPFs, Tels, Placas e Seguradoras associados."""
    tipo_norm = _normalizar_tipo(tipo_entidade)
    val_limpo = str(valor_identificador).strip()

    if tipo_norm not in COLUNA_POR_TIPO or not val_limpo:
        return {"novos_nos": [], "novas_arestas": [], "total_encontrado": 0}

    coluna_busca = COLUNA_POR_TIPO[tipo_norm]
    node_pai_id = f"{PREFIXO_POR_TIPO[tipo_norm]}_{val_limpo}"

    conn = get_db_connection()
    try:
        query = f"""
            SELECT data, id_assistencia, cpf, telefone, placa, empresa_cliente
            FROM criacoes_diarias
            WHERE {coluna_busca} = ?
            UNION ALL
            SELECT data, id_assistencia, cpf, telefone, placa, empresa_cliente
            FROM assistencias
            WHERE {coluna_busca} = ?
        """
        df_hits = pd.read_sql_query(query, conn, params=[val_limpo, val_limpo])
    except Exception:
        df_hits = pd.DataFrame()
    finally:
        conn.close()

    if df_hits.empty:
        return {"novos_nos": [], "novas_arestas": [], "total_encontrado": 0}

    contagem_nos: Dict[str, int] = {}
    arestas_por_destino: Dict[str, Dict[str, Any]] = {}
    infos_no: Dict[str, Tuple[str, str]] = {}

    for r in df_hits.itertuples(index=False):
        candidatos = []
        cpf = str(r.cpf).strip() if pd.notna(r.cpf) else ""
        tel = str(r.telefone).strip() if pd.notna(r.telefone) else ""
        placa = str(r.placa).strip() if pd.notna(r.placa) else ""
        empresa = str(r.empresa_cliente).strip() if pd.notna(r.empresa_cliente) else ""

        if cpf and tipo_norm != "cpf":
            candidatos.append(("cpf", cpf, f"{PREFIXO_POR_TIPO['cpf']}_{cpf}"))
        if tel and tipo_norm != "telefone":
            candidatos.append(("telefone", tel, f"{PREFIXO_POR_TIPO['telefone']}_{tel}"))
        if placa and tipo_norm != "placa":
            candidatos.append(("placa", placa, f"{PREFIXO_POR_TIPO['placa']}_{placa}"))
        if empresa:
            candidatos.append(("empresa", empresa, f"EMPRESA_{empresa.upper()}"))

        for tipo_cand, valor_cand, id_cand in candidatos:
            contagem_nos[id_cand] = contagem_nos.get(id_cand, 0) + 1
            infos_no[id_cand] = (tipo_cand, valor_cand)
            if id_cand not in arestas_por_destino:
                arestas_por_destino[id_cand] = {
                    "from": node_pai_id, "to": id_cand,
                    "ocorrencias": 0, "ultima_data": str(r.data), "ultima_assistencia": str(r.id_assistencia)
                }
            arestas_por_destino[id_cand]["ocorrencias"] += 1

    novos_nos: List[Dict[str, Any]] = []
    for id_no, (tipo_cand, valor_cand) in infos_no.items():
        grau = contagem_nos[id_no]
        if tipo_cand == "empresa":
            novos_nos.append(_montar_no_empresa(valor_cand, grau, node_pai_id))
        else:
            novos_nos.append(_montar_no_entidade(tipo_cand, valor_cand, grau, node_pai_id))

    novas_arestas: List[Dict[str, Any]] = []
    for idx, (id_no, dados_aresta) in enumerate(arestas_por_destino.items()):
        rotulo_qtd = f"{dados_aresta['ocorrencias']} ocorrência(s) em comum"
        novas_arestas.append({
            "id": f"enr_{node_pai_id}_{idx}",
            "from": dados_aresta["from"],
            "to": dados_aresta["to"],
            "title": f"Vínculo via Enrich: {rotulo_qtd} (última: {dados_aresta['ultima_assistencia']} em {dados_aresta['ultima_data']})",
            "label": "",
            "color": {"color": "rgba(148, 163, 184, 0.4)", "highlight": "#38BDF8", "hover": "#38BDF8"},
            "width": min(4.0, 1.0 + dados_aresta["ocorrencias"] * 0.3),
            "smooth": False
        })

    return {
        "novos_nos": novos_nos,
        "novas_arestas": novas_arestas,
        "total_encontrado": len(df_hits)
    }