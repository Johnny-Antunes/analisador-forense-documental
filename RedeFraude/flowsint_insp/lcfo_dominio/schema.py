"""
Esquema das tabelas do domínio (sketch e pivôs). Sem dependências de propósito: o database
chama garantir_schema() no próprio import (garantir_schema_db), e se este código morasse em
sketch_repo/pivos — que importam database — importar um deles primeiro criava um import
circular (o app quebrava ao subir com `python app_nicegui.py`).
"""

from __future__ import annotations

import sqlite3

DDL_SKETCH = [
    """CREATE TABLE IF NOT EXISTS sketch_entidades (
        id_sketch TEXT NOT NULL,
        node_id TEXT NOT NULL,                  -- PREFIXO_valorCanônico
        tipo TEXT NOT NULL, valor TEXT NOT NULL, rotulo TEXT,
        campos_json TEXT NOT NULL DEFAULT '{}', -- campos do schema do tipo (item 1)
        origem TEXT NOT NULL,                   -- manual | pivo | importacao | anotacao_oficial
        id_execucao TEXT,                       -- item 7
        nota TEXT, aparencia_json TEXT,         -- item 9
        fundida_em TEXT,                        -- item 6
        extra_json TEXT NOT NULL DEFAULT '{}',  -- exibição herdada (grau, órbita, título do vis)
        criado_em TEXT NOT NULL, criado_por TEXT,
        removido_em TEXT, removido_por TEXT, motivo_remocao TEXT, lote_remocao TEXT,
        PRIMARY KEY (id_sketch, node_id))""",
    """CREATE TABLE IF NOT EXISTS sketch_vinculos (
        id TEXT PRIMARY KEY, id_sketch TEXT NOT NULL,
        de_id TEXT NOT NULL, para_id TEXT NOT NULL,
        rotulo TEXT, direcao TEXT DEFAULT 'de_para',    -- de_para | para_de | nenhuma
        situacao TEXT DEFAULT 'confirmado',             -- confirmado | hipotese
        data_ref TEXT, confianca INTEGER, observacao TEXT,
        n_ocorrencias INTEGER DEFAULT 1,
        origem TEXT NOT NULL,                           -- manual | pivo | anotacao | fusao
        ref_oficial TEXT, id_execucao TEXT, lote_fusao TEXT,
        extra_json TEXT NOT NULL DEFAULT '{}',          -- estilo vis herdado (cor, largura, título)
        criado_em TEXT NOT NULL, criado_por TEXT,
        removido_em TEXT, removido_por TEXT, motivo_remocao TEXT, lote_remocao TEXT)""",
    """CREATE TABLE IF NOT EXISTS sketch_posicoes (
        id_sketch TEXT, node_id TEXT, layout TEXT,      -- organico | arvore | leque
        x REAL NOT NULL, y REAL NOT NULL, fixado INTEGER DEFAULT 0,
        atualizado_em TEXT, PRIMARY KEY (id_sketch, node_id, layout))""",
    """CREATE TABLE IF NOT EXISTS sketch_eventos (
        id INTEGER PRIMARY KEY AUTOINCREMENT, id_sketch TEXT NOT NULL,
        quando TEXT NOT NULL, quem TEXT, acao TEXT NOT NULL,
        alvo_tipo TEXT, alvo_id TEXT, detalhe_json TEXT)""",
    """CREATE TABLE IF NOT EXISTS sketch_aliases (
        id_sketch TEXT, node_id_antigo TEXT, node_id_destino TEXT,
        PRIMARY KEY (id_sketch, node_id_antigo))""",
    "CREATE INDEX IF NOT EXISTS idx_skent_ativas ON sketch_entidades(id_sketch, removido_em)",
    "CREATE INDEX IF NOT EXISTS idx_skvin_ativos ON sketch_vinculos(id_sketch, removido_em)",
    "CREATE INDEX IF NOT EXISTS idx_skvin_pontas ON sketch_vinculos(id_sketch, de_id, para_id)",
    "CREATE INDEX IF NOT EXISTS idx_skent_lote ON sketch_entidades(id_sketch, lote_remocao)",
    "CREATE INDEX IF NOT EXISTS idx_skvin_lote ON sketch_vinculos(id_sketch, lote_remocao)",
    "CREATE INDEX IF NOT EXISTS idx_skevt_quando ON sketch_eventos(id_sketch, quando)",
]

DDL_PIVOS = [
    """CREATE TABLE IF NOT EXISTS sketch_execucoes (
        id TEXT PRIMARY KEY, id_sketch TEXT NOT NULL,
        pivo TEXT NOT NULL, parametros_json TEXT, entradas_json TEXT,
        estado TEXT NOT NULL,                 -- pendente | executando | concluida | falha | cancelada
        iniciado_em TEXT, concluido_em TEXT, ms INTEGER,
        n_nos_novos INTEGER, n_vinculos_novos INTEGER, n_existentes INTEGER,
        cache_hit INTEGER DEFAULT 0, erro TEXT,
        versao_base TEXT,                     -- impressão da base: (contagem, maior id) das tabelas lidas
        id_receita_exec TEXT, usuario TEXT, avisos_json TEXT)""",
    "CREATE INDEX IF NOT EXISTS idx_skexec ON sketch_execucoes(id_sketch, iniciado_em)",
    """CREATE TABLE IF NOT EXISTS receitas (
        id TEXT PRIMARY KEY, nome TEXT NOT NULL, entrada TEXT NOT NULL,
        passos_json TEXT NOT NULL, criado_em TEXT, criado_por TEXT)""",
]


DDL_INGESTAO = [
    """CREATE TABLE IF NOT EXISTS ingestao_execucoes (
        id TEXT PRIMARY KEY, tipo_base TEXT, iniciado_em TEXT, concluido_em TEXT, usuario TEXT,
        arquivos_json TEXT, n_lidas INTEGER, n_inseridas INTEGER, n_duplicadas INTEGER, n_rejeitadas INTEGER,
        estado TEXT, backup_path TEXT)""",
    """CREATE TABLE IF NOT EXISTS ingestao_rejeitos (
        id_execucao TEXT, arquivo TEXT, linha INTEGER, campo TEXT, valor TEXT,
        motivo TEXT, corrigido_valor TEXT, estado TEXT)""",          # pendente | corrigido | descartado
    "CREATE INDEX IF NOT EXISTS idx_ingrej ON ingestao_rejeitos(id_execucao)",
    """CREATE TABLE IF NOT EXISTS ingestao_especificacoes (
        id TEXT PRIMARY KEY, nome TEXT, assinatura_colunas TEXT,
        mapeamento_json TEXT, formato_data TEXT, criado_em TEXT)""",
]


DDL_FILTROS = [
    """CREATE TABLE IF NOT EXISTS sketch_filtros (
        id TEXT PRIMARY KEY, id_sketch TEXT NOT NULL, nome TEXT NOT NULL, regras_json TEXT NOT NULL,
        criado_em TEXT, criado_por TEXT)""",                                  # filtros por regra salvos (item 12)
    "CREATE INDEX IF NOT EXISTS idx_skfiltro ON sketch_filtros(id_sketch)",
]

DDL_TIPOS = [
    """CREATE TABLE IF NOT EXISTS tipos_customizados (
        chave TEXT PRIMARY KEY, definicao_json TEXT NOT NULL, ativo INTEGER DEFAULT 1,
        criado_em TEXT, criado_por TEXT, atualizado_em TEXT)""",             # tipos criados pelo analista (item 13)
]

DDL_MENCOES = [
    "CREATE TABLE IF NOT EXISTS parecer_mencoes (id_parecer INTEGER, ref TEXT, tipo TEXT)",        # retrolinks (item 11)
    "CREATE INDEX IF NOT EXISTS idx_pmenc_ref ON parecer_mencoes(ref)",
    "CREATE INDEX IF NOT EXISTS idx_pmenc_par ON parecer_mencoes(id_parecer)",
]


def garantir_schema(cursor: sqlite3.Cursor) -> None:
    """Chamado por database.garantir_schema_db (mesma conexão, mesmo commit)."""
    for ddl in DDL_SKETCH + DDL_PIVOS + DDL_INGESTAO + DDL_MENCOES + DDL_FILTROS + DDL_TIPOS:
        cursor.execute(ddl)
    # lote de cada linha (reverter uma ingestão) e colunas não mapeadas (item 10). ADD COLUMN não reescreve a tabela.
    for tabela in ("assistencias", "criacoes_diarias"):
        cols = {r[1] for r in cursor.execute(f"PRAGMA table_info({tabela})").fetchall()}
        if cols and "id_ingestao" not in cols:
            cursor.execute(f"ALTER TABLE {tabela} ADD COLUMN id_ingestao TEXT")
        if cols and "dados_extra_json" not in cols:
            cursor.execute(f"ALTER TABLE {tabela} ADD COLUMN dados_extra_json TEXT")
    colunas = {r[1] for r in cursor.execute("PRAGMA table_info(caso_sketches)").fetchall()}
    if colunas and "versao" not in colunas:
        cursor.execute("ALTER TABLE caso_sketches ADD COLUMN versao INTEGER DEFAULT 0")
    if colunas and "extras_migrado" not in colunas:
        cursor.execute("ALTER TABLE caso_sketches ADD COLUMN extras_migrado INTEGER DEFAULT 0")
    if colunas and "flags_json" not in colunas:                   # significado de cada cor de flag, por sketch (item 9)
        cursor.execute("ALTER TABLE caso_sketches ADD COLUMN flags_json TEXT DEFAULT '{}'")
