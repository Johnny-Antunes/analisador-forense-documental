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
from lcfo_dominio import tipos

# Tudo por tipo vem do registro único (lcfo_dominio/tipos.py). Os nomes antigos ficam como
# aliases para quem ainda importa daqui (lcfo_ui/config.py, scripts).
# Normalização = a mesma da ingestão: "123.456.789-01" -> "12345678901", "abc-1d23" -> "ABC1D23".
LIMPEZA_POR_TIPO = {k: t.normalizar for k, t in tipos.REGISTRO.items() if t.vizinho_no_enrich}
PREFIXO_POR_TIPO = {k: t.prefixo for k, t in tipos.REGISTRO.items()}
COLUNA_POR_TIPO = {t.chave: t.coluna_banco for t in tipos.tipos_enriqueciveis()}
CORES_POR_TIPO = {k: t.cores for k, t in tipos.REGISTRO.items()}
FONTE_NO = tipos.FONTE_NO


def _normalizar_tipo(tipo_bruto: str) -> str:
    t = tipos.obter(tipo_bruto)
    return t.chave if t else str(tipo_bruto).strip().lower().replace("[", "").replace("]", "")


def _montar_no_entidade(tipo: str, valor: str, grau: int, node_pai_id: str, nome: str = "") -> Dict[str, Any]:
    t = tipos.obter(tipo)
    cor = t.cores
    valor_fmt = t.formatar(valor)

    return {
        "id": t.node_id(valor),
        # CPF: o rótulo é o nome do titular (como no grafo); sem nome, o CPF formatado.
        "label": t.rotular({"valor": valor, "nome_titular": nome}),
        "title": f"{tipo.upper()}: {valor_fmt}" + (f"\n{nome}" if nome else "") + "\nDescoberto via Enrich",
        "tipo": tipo,
        "shape": "image",
        "image": t.icone_svg(),
        "size": 20,
        "degree": grau,
        "betweenness": 0.0,
        "community": 0,
        "valor": valor_fmt,
        "nome_titular": nome,
        "orbit_ao_redor": node_pai_id,
        "color": {
            "background": cor["bg"], "border": cor["border"]
        },
        "font": dict(FONTE_NO),
    }


def enriquecer_entidade_local(tipo_entidade: str, valor_identificador: str) -> Dict[str, Any]:
    """
    Expande os vínculos de uma entidade: CPFs, telefones e placas que aparecem nas mesmas
    assistências (Blacklist + Criações Diárias).

    Devolve também o nó da PRÓPRIA entidade (`no_origem`, incluído em `novos_nos`): no Add ela
    ainda não está no grafo, e antes só chegavam os vizinhos — com arestas apontando para um nó
    inexistente (o CPF adicionado não aparecia; só telefone e empresa soltos). Quem consome
    descarta os ids que já existem no grafo.

    A seguradora cliente (empresa_cliente) NÃO vira nó: milhares de assistências compartilham a
    mesma seguradora, e um nó dela ligava pessoas sem relação entre si. Ela fica na descrição
    do vínculo.
    """
    tipo_norm = _normalizar_tipo(tipo_entidade)
    bruto = str(valor_identificador).strip()
    limpar = LIMPEZA_POR_TIPO.get(tipo_norm)
    val_limpo = (limpar(bruto) if limpar else "") or bruto

    if tipo_norm not in COLUNA_POR_TIPO or not val_limpo:
        return {"novos_nos": [], "novas_arestas": [], "total_encontrado": 0}

    coluna_busca = COLUNA_POR_TIPO[tipo_norm]
    node_pai_id = f"{PREFIXO_POR_TIPO[tipo_norm]}_{val_limpo}"

    conn = get_db_connection()
    try:
        query = f"""
            SELECT data, id_assistencia, cpf, telefone, placa, empresa_cliente, titular
            FROM criacoes_diarias
            WHERE {coluna_busca} = ?
            UNION ALL
            SELECT data, id_assistencia, cpf, telefone, placa, empresa_cliente, titular
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
    nomes_cpf: Dict[str, str] = {}

    def _txt(v) -> str:
        return str(v).strip() if pd.notna(v) and str(v).strip().lower() not in ("nan", "none") else ""

    for r in df_hits.itertuples(index=False):
        candidatos = []
        cpf, tel, placa = _txt(r.cpf), _txt(r.telefone), _txt(r.placa)
        empresa, titular = _txt(r.empresa_cliente), _txt(r.titular)
        if cpf and titular:
            nomes_cpf.setdefault(cpf, titular)

        if cpf and tipo_norm != "cpf":
            candidatos.append(("cpf", cpf, f"{PREFIXO_POR_TIPO['cpf']}_{cpf}"))
        if tel and tipo_norm != "telefone":
            candidatos.append(("telefone", tel, f"{PREFIXO_POR_TIPO['telefone']}_{tel}"))
        if placa and tipo_norm != "placa":
            candidatos.append(("placa", placa, f"{PREFIXO_POR_TIPO['placa']}_{placa}"))

        for tipo_cand, valor_cand, id_cand in candidatos:
            contagem_nos[id_cand] = contagem_nos.get(id_cand, 0) + 1
            infos_no[id_cand] = (tipo_cand, valor_cand)
            if id_cand not in arestas_por_destino:
                arestas_por_destino[id_cand] = {
                    "from": node_pai_id, "to": id_cand, "empresas": set(),
                    "ocorrencias": 0, "ultima_data": str(r.data), "ultima_assistencia": str(r.id_assistencia)
                }
            arestas_por_destino[id_cand]["ocorrencias"] += 1
            if empresa:
                arestas_por_destino[id_cand]["empresas"].add(empresa.upper())

    # O próprio nó de origem primeiro (no Add ele ainda não está no grafo).
    no_origem = _montar_no_entidade(tipo_norm, val_limpo, len(df_hits), node_pai_id,
                                    nome=nomes_cpf.get(val_limpo, "") if tipo_norm == "cpf" else "")
    no_origem.pop("orbit_ao_redor", None)
    novos_nos: List[Dict[str, Any]] = [no_origem]
    for id_no, (tipo_cand, valor_cand) in infos_no.items():
        novos_nos.append(_montar_no_entidade(tipo_cand, valor_cand, contagem_nos[id_no], node_pai_id,
                                             nome=nomes_cpf.get(valor_cand, "") if tipo_cand == "cpf" else ""))

    novas_arestas: List[Dict[str, Any]] = []
    for idx, (id_no, dados_aresta) in enumerate(arestas_por_destino.items()):
        rotulo_qtd = f"{dados_aresta['ocorrencias']} ocorrência(s) em comum"
        empresas = ", ".join(sorted(dados_aresta["empresas"])[:3])
        novas_arestas.append({
            # id estável por par (antes era por posição: enriquecer de novo podia trocar arestas)
            "id": f"enr_{node_pai_id}__{id_no}",
            "from": dados_aresta["from"],
            "to": dados_aresta["to"],
            "title": f"Vínculo via Enrich: {rotulo_qtd} (última: {dados_aresta['ultima_assistencia']} em {dados_aresta['ultima_data']})"
                     + (f"\nSeguradora(s): {empresas}" if empresas else ""),
            "label": "",
            "color": {"color": "rgba(148, 163, 184, 0.4)", "highlight": "#38BDF8", "hover": "#38BDF8"},
            "width": min(4.0, 1.0 + dados_aresta["ocorrencias"] * 0.3),
            "smooth": False
        })

    return {
        "novos_nos": novos_nos,
        "novas_arestas": novas_arestas,
        "no_origem": no_origem["id"],
        "total_encontrado": len(df_hits)
    }