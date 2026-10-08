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
)
from graph_engine import (
    carregar_redes, processar_subgrafo_caso, processar_grafo_dataframe,
)
from correlation_engine import calcular_score_caso, obter_scores_triagem
from anomalias_engine import obter_radar_anomalias_macro, extrair_top_infratores_municipio
from utils import formatar_mencoes_forenses

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
    "carregar_redes",
    "processar_subgrafo_caso",
    "processar_grafo_dataframe",
    "calcular_score_caso",
    "obter_scores_triagem",
    "obter_radar_anomalias_macro",
    "extrair_top_infratores_municipio",
    "formatar_mencoes_forenses",
]
