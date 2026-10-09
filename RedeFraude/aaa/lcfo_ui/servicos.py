"""
Fachada única entre a interface (lcfo_ui) e os motores de dados
(database, graph_engine, correlation_engine, anomalias_engine, utils).

Toda tela importa daqui — nunca direto dos motores. Isso dá um ponto único
para medir tempo (perf.medir), e futuramente para cachear ou mover
chamadas pesadas para fora do event loop (run.io_bound) sem caçar
chamadas espalhadas pelas telas.
"""

from __future__ import annotations

from database import (
    CAMINHO_REDE_OFICIAL, PASTA_LOCAL_BLACKLIST,
    CAMINHO_REDE_CRIACAO, PASTA_LOCAL_CRIACAO,
    carregar_arquivos_para_sqlite, carregar_criacoes_diarias_para_sqlite,
    consultar_detalhes_caso, cruzar_com_criacoes_diarias, obter_radar_expansoes,
    atualizar_metadados_caso, carregar_dados_caso, carregar_todos_casos_cadastrados,
    cadastrar_entidade_suspeita, listar_entidades_suspeitas, remover_entidade_suspeita,
    consultar_todas_ocorrencias_entidade, semear_base_mestra_da_blacklist,
    investigar_alvo_em_criacoes_diarias, promover_descoberta_para_caso,
    vincular_caso_e_entidades, anexar_descoberta_a_caso_existente,
    carregar_historico_pareceres, resetar_layout_caso,
    cadastrar_cidade_risco, listar_cidades_risco, remover_cidade_risco,
    reparar_mojibake_historico,
    ocultar_no_do_caso, restaurar_no_do_caso, listar_nos_ocultos_do_caso,
    salvar_posicoes_layout_caso, carregar_posicoes_layout_caso,
    criar_sketch, listar_sketches_do_caso, carregar_sketch,
    remover_sketch, salvar_nos_extras_sketch, carregar_nos_extras_sketch,
    vincular_cluster_como_sketch, registrar_parecer, atualizar_parecer_historico,
    corrigir_datas_pela_origem,
)
from graph_engine import (
    carregar_redes, processar_subgrafo_caso, processar_grafo_dataframe,
)
from correlation_engine import calcular_score_caso, obter_scores_triagem
from anomalias_engine import obter_radar_anomalias_macro, extrair_top_infratores_municipio
from utils import formatar_mencoes_forenses
from database import marcar_entidades_conhecidas
from territorio_engine import (
    montar_dados_mapa_radar, resumo_do_mes, montar_payload_mapa_caso, carregar_municipios_ibge,
)

from lcfo_ui.perf import medir

# Funções potencialmente pesadas — cronometradas (ver lcfo_ui/perf.py).
carregar_redes = medir(carregar_redes)
processar_subgrafo_caso = medir(processar_subgrafo_caso)
processar_grafo_dataframe = medir(processar_grafo_dataframe)
calcular_score_caso = medir(calcular_score_caso)
obter_scores_triagem = medir(obter_scores_triagem)
obter_radar_anomalias_macro = medir(obter_radar_anomalias_macro)
extrair_top_infratores_municipio = medir(extrair_top_infratores_municipio)
consultar_detalhes_caso = medir(consultar_detalhes_caso)
cruzar_com_criacoes_diarias = medir(cruzar_com_criacoes_diarias)
obter_radar_expansoes = medir(obter_radar_expansoes)
carregar_arquivos_para_sqlite = medir(carregar_arquivos_para_sqlite)
carregar_criacoes_diarias_para_sqlite = medir(carregar_criacoes_diarias_para_sqlite)
montar_dados_mapa_radar = medir(montar_dados_mapa_radar)

__all__ = [
    "CAMINHO_REDE_OFICIAL",
    "PASTA_LOCAL_BLACKLIST",
    "CAMINHO_REDE_CRIACAO",
    "PASTA_LOCAL_CRIACAO",
    "carregar_arquivos_para_sqlite",
    "carregar_criacoes_diarias_para_sqlite",
    "consultar_detalhes_caso",
    "cruzar_com_criacoes_diarias",
    "obter_radar_expansoes",
    "atualizar_metadados_caso",
    "carregar_dados_caso",
    "carregar_todos_casos_cadastrados",
    "cadastrar_entidade_suspeita",
    "listar_entidades_suspeitas",
    "remover_entidade_suspeita",
    "consultar_todas_ocorrencias_entidade",
    "semear_base_mestra_da_blacklist",
    "investigar_alvo_em_criacoes_diarias",
    "promover_descoberta_para_caso",
    "vincular_caso_e_entidades",
    "anexar_descoberta_a_caso_existente",
    "carregar_historico_pareceres",
    "resetar_layout_caso",
    "cadastrar_cidade_risco",
    "listar_cidades_risco",
    "remover_cidade_risco",
    "reparar_mojibake_historico",
    "ocultar_no_do_caso",
    "restaurar_no_do_caso",
    "listar_nos_ocultos_do_caso",
    "salvar_posicoes_layout_caso",
    "carregar_posicoes_layout_caso",
    "criar_sketch",
    "listar_sketches_do_caso",
    "carregar_sketch",
    "remover_sketch",
    "salvar_nos_extras_sketch",
    "carregar_nos_extras_sketch",
    "vincular_cluster_como_sketch",
    "registrar_parecer",
    "atualizar_parecer_historico",
    "corrigir_datas_pela_origem",
    "carregar_redes",
    "processar_subgrafo_caso",
    "processar_grafo_dataframe",
    "calcular_score_caso",
    "obter_scores_triagem",
    "obter_radar_anomalias_macro",
    "extrair_top_infratores_municipio",
    "formatar_mencoes_forenses",
    "marcar_entidades_conhecidas",
    "montar_dados_mapa_radar",
    "resumo_do_mes",
    "montar_payload_mapa_caso",
    "carregar_municipios_ibge",
]


# ---------------------------------------------------------------------------
# Cache do score do caso (0,3s por redesenho da mesa em células grandes).
# A chave cobre tudo de que o cálculo depende: conteúdo do df (hash), nós da
# célula, betweenness, versões de assistencias / watchlist e o dia (o score
# usa recência relativa a hoje).
# ---------------------------------------------------------------------------
import copy as _copy
import threading as _threading
from collections import OrderedDict as _OrderedDict
from datetime import date as _date

import pandas as _pd

from database import obter_versao_cidades_risco as _versao_cidades
from database import obter_versao_criacoes as _versao_criacoes  # inclui a geração (correção de datas)
from graph_engine import obter_versao_assistencias as _versao_assistencias

_CACHE_SCORE: "_OrderedDict" = _OrderedDict()
_CACHE_GRAFO_DF: "_OrderedDict" = _OrderedDict()
_LOCK_CACHES = _threading.Lock()
_MAX_CACHE = 32


def _hash_df(df) -> tuple:
    if df is None or df.empty:
        return (0, 0)
    return (len(df), tuple(df.columns), int(_pd.util.hash_pandas_object(df, index=False).sum()))


def _memo(cache, chave, calcular):
    with _LOCK_CACHES:
        if chave in cache:
            cache.move_to_end(chave)
            return _copy.deepcopy(cache[chave])
    valor = calcular()
    with _LOCK_CACHES:
        cache[chave] = valor
        while len(cache) > _MAX_CACHE:
            cache.popitem(last=False)
    return _copy.deepcopy(valor)


def calcular_score_caso_em_cache(cluster_obj, df_dados, subgrafo=None, betweenness_precalculado=None):
    chave = (
        cluster_obj["id"] if cluster_obj else None,
        tuple(cluster_obj["nodes"]) if cluster_obj else None,
        _hash_df(df_dados), betweenness_precalculado,
        _versao_assistencias(), _versao_cidades(), _versao_criacoes(), _date.today().isoformat(),
    )
    return _memo(_CACHE_SCORE, chave, lambda: calcular_score_caso(
        cluster_obj, df_dados, subgrafo, betweenness_precalculado=betweenness_precalculado))


def processar_grafo_dataframe_em_cache(df, alvo_principal="", font_slider=11, espacamento=280):
    """Grafo de uma descoberta (df avulso) — mesmo resultado, sem recalcular a cada redesenho."""
    chave = (_hash_df(df), alvo_principal, font_slider, espacamento)
    return _memo(_CACHE_GRAFO_DF, chave, lambda: processar_grafo_dataframe(
        df, alvo_principal=alvo_principal, font_slider=font_slider, espacamento=espacamento))


__all__ += ["calcular_score_caso_em_cache", "processar_grafo_dataframe_em_cache"]
