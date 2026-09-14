import sqlite3
import pandas as pd
from pathlib import Path
import streamlit as st
import csv
import re
import hashlib
from datetime import datetime

from utils import (
    mapear_coluna, limpar_cpf_cnpj, limpar_tel, limpar_placa,
    normalizar_texto, tratar_data_flexivel, obter_coordenadas,
    extrair_data_arquivo, formatar_cpf_cnpj, formatar_tel
)

DB_PATH = Path(__file__).parent / "banco_fraudes.db"

# =====================================================
# CONFIGURAÇÃO DE CAMINHOS CORPORATIVOS E LOCAIS
# =====================================================
CAMINHO_REDE_OFICIAL = r"T:\Corporativo\OPERAÇÕES\CPO\QUERIES\Verificação Blacklist - Data Criação ME\Duplicidade de Placa"
PASTA_LOCAL_BLACKLIST = Path(__file__).parent / "casos_duplicidade"
PASTA_LOCAL_BLACKLIST.mkdir(exist_ok=True)

CAMINHO_REDE_CRIACAO = r"T:\Corporativo\OPERAÇOES\CPO\QUERIES\Atalizaçoes_Diarias\Criação"
PASTA_LOCAL_CRIACAO = Path(__file__).parent / "base_criacao"
PASTA_LOCAL_CRIACAO.mkdir(exist_ok=True)


def get_db_connection():
    conn = sqlite3.connect(DB_PATH, timeout=60.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode = WAL;")
    conn.execute("PRAGMA synchronous = NORMAL;")
    conn.execute("PRAGMA cache_size = -64000;")
    return conn


def garantir_schema_db():
    conn = get_db_connection()
    try:
        cursor = conn.cursor()

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS arquivos_processados (
                nome_arquivo TEXT PRIMARY KEY,
                tipo_base TEXT,
                data_processamento TEXT,
                registros_inseridos INTEGER
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS assistencias (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                arquivo TEXT,
                id_assistencia TEXT,
                data TEXT,
                titular TEXT,
                cpf TEXT,
                telefone TEXT,
                placa TEXT,
                servico TEXT,
                bairro TEXT,
                cidade TEXT,
                uf TEXT,
                latitude REAL,
                longitude REAL,
                UNIQUE (arquivo, id_assistencia, cpf, telefone, placa, servico)
            )
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_cpf ON assistencias(cpf)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_tel ON assistencias(telefone)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_placa ON assistencias(placa)")

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS criacoes_diarias (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                arquivo TEXT,
                id_assistencia TEXT,
                data TEXT,
                titular TEXT,
                cpf TEXT,
                telefone TEXT,
                placa TEXT,
                servico TEXT,
                bairro TEXT,
                cidade TEXT,
                uf TEXT,
                latitude REAL,
                longitude REAL,
                UNIQUE (arquivo, id_assistencia, cpf, telefone, placa, servico)
            )
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_criacao_cpf ON criacoes_diarias(cpf)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_criacao_tel ON criacoes_diarias(telefone)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_criacao_placa ON criacoes_diarias(placa)")

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS casos_investigacao (
                id_caso TEXT PRIMARY KEY,
                nome_personalizado TEXT,
                status TEXT DEFAULT 'Em Investigação',
                analista_responsavel TEXT DEFAULT '',
                parecer TEXT DEFAULT '',
                data_atualizacao TEXT
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS caso_membros (
                id_caso TEXT,
                node_id TEXT,
                data_associacao TEXT,
                PRIMARY KEY (id_caso, node_id)
            )
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_membros_node ON caso_membros(node_id)")

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS casos_historico_pareceres (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                id_caso TEXT,
                data_registro TEXT,
                analista TEXT,
                status TEXT,
                parecer TEXT
            )
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_hist_caso ON casos_historico_pareceres(id_caso)")

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS entidades_suspeitas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tipo TEXT NOT NULL,
                valor TEXT NOT NULL,
                valor_formatado TEXT,
                nome_referencia TEXT,
                motivo TEXT,
                id_caso TEXT DEFAULT '',
                quadrilha_caso TEXT,
                status TEXT DEFAULT 'Ativo',
                analista TEXT,
                data_cadastro TEXT,
                UNIQUE(tipo, valor)
            )
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_suspeito_busca ON entidades_suspeitas(tipo, valor)")

        cursor.execute("PRAGMA table_info(entidades_suspeitas)")
        colunas_suspeitas = [r["name"] for r in cursor.fetchall()]
        if colunas_suspeitas and "id_caso" not in colunas_suspeitas:
            cursor.execute("ALTER TABLE entidades_suspeitas ADD COLUMN id_caso TEXT DEFAULT ''")

        # =========================================================================
        # PASSO 3: Tabela de persistência de coordenadas de layout
        # =========================================================================
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS caso_layout_nos (
                id_caso TEXT,
                node_id TEXT,
                x INTEGER,
                y INTEGER,
                PRIMARY KEY (id_caso, node_id)
            )
        """)

        conn.commit()
    finally:
        conn.close()

garantir_schema_db()


# =====================================================
# RESOLUÇÃO DE IDENTIDADE PERSISTENTE DE CASOS
# =====================================================
def resolver_identidade_componente(comp_nodes, conn=None):
    fechar_ao_fim = False
    if conn is None:
        conn = get_db_connection()
        fechar_ao_fim = True

    try:
        cursor = conn.cursor()
        placeholders = ','.join(['?'] * len(comp_nodes))
        cursor.execute(f"SELECT DISTINCT id_caso FROM caso_membros WHERE node_id IN ({placeholders})", list(comp_nodes))
        casos_encontrados = [r["id_caso"] for r in cursor.fetchall()]

        dt_agora = datetime.now().strftime("%d/%m/%Y %H:%M")

        if len(casos_encontrados) == 1:
            id_caso_final = casos_encontrados[0]
        elif len(casos_encontrados) > 1:
            cursor.execute(f"""
                SELECT id_caso, COUNT(*) as total 
                FROM caso_membros 
                WHERE id_caso IN ({','.join(['?']*len(casos_encontrados))})
                GROUP BY id_caso 
                ORDER BY total DESC 
                LIMIT 1
            """, casos_encontrados)
            id_caso_final = cursor.fetchone()["id_caso"]
        else:
            assinatura = "-".join(sorted(list(comp_nodes)))
            id_caso_final = "CEL-" + hashlib.md5(assinatura.encode("utf-8")).hexdigest()[:6].upper()

        lote_membros = [(id_caso_final, n, dt_agora) for n in comp_nodes]
        cursor.executemany("""
            INSERT OR IGNORE INTO caso_membros (id_caso, node_id, data_associacao)
            VALUES (?, ?, ?)
        """, lote_membros)
        conn.commit()

        return id_caso_final
    finally:
        if fechar_ao_fim:
            conn.close()


# =====================================================
# GESTÃO DE CASOS E HISTÓRICO APPEND-ONLY
# =====================================================
def salvar_dados_caso(id_caso, nome_personalizado, status, analista, parecer):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        dt_agora = datetime.now().strftime("%d/%m/%Y %H:%M")

        cursor.execute("""
            INSERT INTO casos_investigacao (id_caso, nome_personalizado, status, analista_responsavel, parecer, data_atualizacao)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(id_caso) DO UPDATE SET
                nome_personalizado = excluded.nome_personalizado,
                status = excluded.status,
                analista_responsavel = excluded.analista_responsavel,
                parecer = excluded.parecer,
                data_atualizacao = excluded.data_atualizacao
        """, (id_caso, nome_personalizado.strip().upper(), status, analista.strip(), parecer.strip(), dt_agora))

        if parecer.strip():
            cursor.execute("""
                INSERT INTO casos_historico_pareceres (id_caso, data_registro, analista, status, parecer)
                VALUES (?, ?, ?, ?, ?)
            """, (id_caso, dt_agora, analista.strip(), status, parecer.strip()))

        cursor.execute("""
            UPDATE entidades_suspeitas 
            SET quadrilha_caso = ? 
            WHERE id_caso = ?
        """, (nome_personalizado.strip().upper(), id_caso))

        conn.commit()
    finally:
        conn.close()


def carregar_dados_caso(id_caso):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM casos_investigacao WHERE id_caso = ?", (id_caso,))
        row = cursor.fetchone()
        if row:
            return dict(row)
        return {
            "id_caso": id_caso, "nome_personalizado": "", "status": "Em Investigação",
            "analista_responsavel": "", "parecer": "", "data_atualizacao": ""
        }
    finally:
        conn.close()


def carregar_historico_pareceres(id_caso):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM casos_historico_pareceres WHERE id_caso = ? ORDER BY id DESC", (id_caso,))
        return [dict(r) for r in cursor.fetchall()]
    finally:
        conn.close()


def carregar_todos_casos_cadastrados():
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT id_caso, nome_personalizado, status FROM casos_investigacao")
        rows = cursor.fetchall()
        return {r["id_caso"]: {"nome": r["nome_personalizado"], "status": r["status"]} for r in rows}
    finally:
        conn.close()


# =====================================================
# PERSISTÊNCIA DE COORDENADAS DE LAYOUT (PASSO 3)
# =====================================================
def carregar_layout_caso(id_caso):
    """Retorna dict {node_id: (x, y)} com o layout salvo do caso, ou vazio se nunca calculado."""
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT node_id, x, y FROM caso_layout_nos WHERE id_caso = ?", (id_caso,))
        return {r["node_id"]: (r["x"], r["y"]) for r in cursor.fetchall()}
    finally:
        conn.close()


def salvar_layout_caso(id_caso, posicoes):
    """Sobrescreve por completo o layout salvo do caso com as novas posições."""
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM caso_layout_nos WHERE id_caso = ?", (id_caso,))
        lote = [(id_caso, node_id, int(x), int(y)) for node_id, (x, y) in posicoes.items()]
        if lote:
            cursor.executemany(
                "INSERT INTO caso_layout_nos (id_caso, node_id, x, y) VALUES (?, ?, ?, ?)", lote
            )
        conn.commit()
    finally:
        conn.close()


def resetar_layout_caso(id_caso):
    """Apaga o layout salvo, forçando recálculo na próxima abertura. Útil para reorganizar."""
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM caso_layout_nos WHERE id_caso = ?", (id_caso,))
        conn.commit()
    finally:
        conn.close()


# =====================================================
# SEMEADURA AUTOMÁTICA DA BASE MESTRA
# =====================================================
def semear_base_mestra_da_blacklist():
    conn = get_db_connection()
    total_novos = 0
    try:
        cursor = conn.cursor()
        dt_agora = datetime.now().strftime("%d/%m/%Y %H:%M")

        cursor.execute("SELECT DISTINCT cpf, titular FROM assistencias WHERE cpf IS NOT NULL AND cpf != ''")
        for r in cursor.fetchall():
            cpf = r["cpf"]
            nome = r["titular"] or ""
            cursor.execute("""
                INSERT OR IGNORE INTO entidades_suspeitas 
                (tipo, valor, valor_formatado, nome_referencia, motivo, id_caso, quadrilha_caso, status, analista, data_cadastro)
                VALUES ('CPF', ?, ?, ?, 'Histórico Blacklist', '', '', 'Ativo', 'SISTEMA', ?)
            """, (cpf, formatar_cpf_cnpj(cpf), nome, dt_agora))
            if cursor.rowcount > 0: total_novos += 1

        cursor.execute("SELECT DISTINCT telefone FROM assistencias WHERE telefone IS NOT NULL AND telefone != ''")
        for r in cursor.fetchall():
            tel = r["telefone"]
            cursor.execute("""
                INSERT OR IGNORE INTO entidades_suspeitas 
                (tipo, valor, valor_formatado, nome_referencia, motivo, id_caso, quadrilha_caso, status, analista, data_cadastro)
                VALUES ('TELEFONE', ?, ?, '', 'Histórico Blacklist', '', '', 'Ativo', 'SISTEMA', ?)
            """, (tel, formatar_tel(tel), dt_agora))
            if cursor.rowcount > 0: total_novos += 1

        cursor.execute("SELECT DISTINCT placa FROM assistencias WHERE placa IS NOT NULL AND placa != ''")
        for r in cursor.fetchall():
            p = r["placa"]
            fmt_p = f"{p[:3]}-{p[3:]}" if len(p) == 7 else p
            cursor.execute("""
                INSERT OR IGNORE INTO entidades_suspeitas 
                (tipo, valor, valor_formatado, nome_referencia, motivo, id_caso, quadrilha_caso, status, analista, data_cadastro)
                VALUES ('PLACA', ?, ?, '', 'Histórico Blacklist', '', '', 'Ativo', 'SISTEMA', ?)
            """, (p, fmt_p, dt_agora))
            if cursor.rowcount > 0: total_novos += 1

        conn.commit()
        return total_novos
    finally:
        conn.close()


# =====================================================
# MOTOR DE DESCOBERTA ATIVA EM CRIAÇÕES (3 GRAUS)
# =====================================================
def investigar_alvo_em_criacoes_diarias(termo_busca):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='criacoes_diarias'")
        if not cursor.fetchone():
            return None, pd.DataFrame(), {}

        term_clean = re.sub(r'[^a-zA-Z0-9]', '', str(termo_busca)).upper()
        if not term_clean:
            return None, pd.DataFrame(), {}

        query_1g = "SELECT * FROM criacoes_diarias WHERE cpf = ? OR telefone = ? OR placa = ?"
        df_1g = pd.read_sql_query(query_1g, conn, params=[term_clean, term_clean, term_clean])

        if df_1g.empty:
            query_1g_like = "SELECT * FROM criacoes_diarias WHERE cpf LIKE ? OR telefone LIKE ? OR placa LIKE ?"
            df_1g = pd.read_sql_query(query_1g_like, conn, params=[f"%{term_clean}%", f"%{term_clean}%", f"%{term_clean}%"])
            if df_1g.empty:
                return None, pd.DataFrame(), {}

        cpfs_2g = set(df_1g["cpf"].dropna().unique()) - {""}
        tels_2g = set(df_1g["telefone"].dropna().unique()) - {""}
        placas_2g = set(df_1g["placa"].dropna().unique()) - {""}

        clausulas, params = [], []
        if cpfs_2g:
            clausulas.append(f"cpf IN ({','.join(['?']*len(cpfs_2g))})")
            params.extend(list(cpfs_2g))
        if tels_2g:
            clausulas.append(f"telefone IN ({','.join(['?']*len(tels_2g))})")
            params.extend(list(tels_2g))
        if placas_2g:
            clausulas.append(f"placa IN ({','.join(['?']*len(placas_2g))})")
            params.extend(list(placas_2g))

        query_total = f"""
            SELECT DISTINCT data, id_assistencia, titular, cpf, telefone, placa, servico, bairro, cidade, uf, latitude, longitude
            FROM criacoes_diarias
            WHERE {' OR '.join(clausulas)}
            ORDER BY data ASC
        """
        df_rede = pd.read_sql_query(query_total, conn, params=params)

        resumo = {
            "alvo_buscado": termo_busca,
            "termo_limpo": term_clean,
            "total_assistencias": len(df_rede),
            "assistencias_diretas": len(df_1g),
            "qtd_cpfs": df_rede["cpf"].replace("", None).nunique(),
            "qtd_tels": df_rede["telefone"].replace("", None).nunique(),
            "qtd_placas": df_rede["placa"].replace("", None).nunique(),
            "periodo_min": df_rede["data"].min() if not df_rede.empty else "",
            "periodo_max": df_rede["data"].max() if not df_rede.empty else ""
        }

        entidades_dict = {
            "cpfs": list(df_rede["cpf"].replace("", None).dropna().unique()),
            "tels": list(df_rede["telefone"].replace("", None).dropna().unique()),
            "placas": list(df_rede["placa"].replace("", None).dropna().unique())
        }

        return resumo, df_rede, entidades_dict
    finally:
        conn.close()


def promover_descoberta_para_caso(df_descoberta):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        lote_assistencias = []
        nos_promovidos = set()

        for _, r in df_descoberta.iterrows():
            cpf = str(r.get("cpf", "")).strip() if pd.notna(r.get("cpf")) else ""
            tel = str(r.get("telefone", "")).strip() if pd.notna(r.get("telefone")) else ""
            placa = str(r.get("placa", "")).strip() if pd.notna(r.get("placa")) else ""

            if cpf: nos_promovidos.add(f"CPF_{cpf}")
            if tel: nos_promovidos.add(f"TEL_{tel}")
            if placa: nos_promovidos.add(f"PLACA_{placa}")

            lote_assistencias.append((
                "PROMOVIDO_CRIACAO", str(r["id_assistencia"]), str(r["data"]), str(r["titular"]),
                cpf, tel, placa, str(r["servico"]),
                str(r["bairro"]), str(r["cidade"]), str(r["uf"]), r.get("latitude"), r.get("longitude")
            ))

        cursor.executemany("""
            INSERT OR IGNORE INTO assistencias 
            (arquivo, id_assistencia, data, titular, cpf, telefone, placa, servico, bairro, cidade, uf, latitude, longitude)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, lote_assistencias)
        conn.commit()

        return nos_promovidos
    finally:
        conn.close()


def vincular_caso_e_entidades(id_caso_real, nos_promovidos, nome_quadrilha, status, analista, parecer):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        dt_agora = datetime.now().strftime("%d/%m/%Y %H:%M")

        cursor.execute("""
            INSERT INTO casos_investigacao (id_caso, nome_personalizado, status, analista_responsavel, parecer, data_atualizacao)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(id_caso) DO UPDATE SET
                nome_personalizado = excluded.nome_personalizado,
                status = excluded.status,
                analista_responsavel = excluded.analista_responsavel,
                parecer = excluded.parecer,
                data_atualizacao = excluded.data_atualizacao
        """, (id_caso_real, nome_quadrilha.strip().upper(), status, analista.strip(), parecer.strip(), dt_agora))

        if parecer.strip():
            cursor.execute("""
                INSERT INTO casos_historico_pareceres (id_caso, data_registro, analista, status, parecer)
                VALUES (?, ?, ?, ?, ?)
            """, (id_caso_real, dt_agora, analista.strip(), status, parecer.strip()))

        for n in nos_promovidos:
            if n.startswith("CPF_"):
                val = n.replace("CPF_", "")
                cursor.execute("""
                    INSERT OR IGNORE INTO entidades_suspeitas 
                    (tipo, valor, valor_formatado, nome_referencia, motivo, id_caso, quadrilha_caso, status, analista, data_cadastro)
                    VALUES ('CPF', ?, ?, '', 'Promovido de Descoberta', ?, ?, ?, ?, ?)
                """, (val, formatar_cpf_cnpj(val), id_caso_real, nome_quadrilha.upper(), status, analista, dt_agora))
            elif n.startswith("TEL_"):
                val = n.replace("TEL_", "")
                cursor.execute("""
                    INSERT OR IGNORE INTO entidades_suspeitas 
                    (tipo, valor, valor_formatado, nome_referencia, motivo, id_caso, quadrilha_caso, status, analista, data_cadastro)
                    VALUES ('TELEFONE', ?, ?, '', 'Promovido de Descoberta', ?, ?, ?, ?, ?)
                """, (val, formatar_tel(val), id_caso_real, nome_quadrilha.upper(), status, analista, dt_agora))
            elif n.startswith("PLACA_"):
                val = n.replace("PLACA_", "")
                fmt_p = f"{val[:3]}-{val[3:]}" if len(val) == 7 else val
                cursor.execute("""
                    INSERT OR IGNORE INTO entidades_suspeitas 
                    (tipo, valor, valor_formatado, nome_referencia, motivo, id_caso, quadrilha_caso, status, analista, data_cadastro)
                    VALUES ('PLACA', ?, ?, '', 'Promovido de Descoberta', ?, ?, ?, ?, ?)
                """, (val, fmt_p, id_caso_real, nome_quadrilha.upper(), status, analista, dt_agora))

        conn.commit()
    finally:
        conn.close()


def anexar_descoberta_a_caso_existente(df_descoberta, id_caso_destino, analista="", observacao=""):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT nome_personalizado FROM casos_investigacao WHERE id_caso = ?", (id_caso_destino,))
        row = cursor.fetchone()
        nome_quadrilha = row["nome_personalizado"] if row and row["nome_personalizado"] else f"Caso {id_caso_destino}"

        lote_assistencias = []
        nos_anexados = set()
        for _, r in df_descoberta.iterrows():
            cpf = str(r.get("cpf", "")).strip() if pd.notna(r.get("cpf")) else ""
            tel = str(r.get("telefone", "")).strip() if pd.notna(r.get("telefone")) else ""
            placa = str(r.get("placa", "")).strip() if pd.notna(r.get("placa")) else ""

            if cpf: nos_anexados.add(f"CPF_{cpf}")
            if tel: nos_anexados.add(f"TEL_{tel}")
            if placa: nos_anexados.add(f"PLACA_{placa}")

            lote_assistencias.append((
                f"ANEXADO_{id_caso_destino}", str(r["id_assistencia"]), str(r["data"]), str(r["titular"]),
                cpf, tel, placa, str(r["servico"]),
                str(r["bairro"]), str(r["cidade"]), str(r["uf"]), r.get("latitude"), r.get("longitude")
            ))

        cursor.executemany("""
            INSERT OR IGNORE INTO assistencias 
            (arquivo, id_assistencia, data, titular, cpf, telefone, placa, servico, bairro, cidade, uf, latitude, longitude)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, lote_assistencias)

        dt_agora = datetime.now().strftime("%d/%m/%Y %H:%M")

        lote_membros = [(id_caso_destino, n, dt_agora) for n in nos_anexados]
        cursor.executemany("""
            INSERT OR IGNORE INTO caso_membros (id_caso, node_id, data_associacao)
            VALUES (?, ?, ?)
        """, lote_membros)

        for cpf in df_descoberta["cpf"].dropna().unique():
            if cpf:
                cursor.execute("""
                    INSERT OR IGNORE INTO entidades_suspeitas (tipo, valor, valor_formatado, nome_referencia, motivo, id_caso, quadrilha_caso, status, analista, data_cadastro)
                    VALUES ('CPF', ?, ?, '', ?, ?, ?, 'Ativo', ?, ?)
                """, (cpf, formatar_cpf_cnpj(cpf), f"Vinculado ao Caso {id_caso_destino}: {observacao}", id_caso_destino, nome_quadrilha, analista, dt_agora))

        for tel in df_descoberta["telefone"].dropna().unique():
            if tel:
                cursor.execute("""
                    INSERT OR IGNORE INTO entidades_suspeitas (tipo, valor, valor_formatado, nome_referencia, motivo, id_caso, quadrilha_caso, status, analista, data_cadastro)
                    VALUES ('TELEFONE', ?, ?, '', ?, ?, ?, 'Ativo', ?, ?)
                """, (tel, formatar_tel(tel), f"Vinculado ao Caso {id_caso_destino}: {observacao}", id_caso_destino, nome_quadrilha, analista, dt_agora))

        for p in df_descoberta["placa"].dropna().unique():
            if p:
                fmt_p = f"{p[:3]}-{p[3:]}" if len(p) == 7 else p
                cursor.execute("""
                    INSERT OR IGNORE INTO entidades_suspeitas (tipo, valor, valor_formatado, nome_referencia, motivo, id_caso, quadrilha_caso, status, analista, data_cadastro)
                    VALUES ('PLACA', ?, ?, '', ?, ?, ?, 'Ativo', ?, ?)
                """, (p, fmt_p, f"Vinculado ao Caso {id_caso_destino}: {observacao}", id_caso_destino, nome_quadrilha, analista, dt_agora))

        conn.commit()
        return True
    finally:
        conn.close()


# =====================================================
# GESTÃO DA BASE MESTRA DE ENTIDADES SUSPEITAS
# =====================================================
def cadastrar_entidade_suspeita(tipo, valor_bruto, nome_ref="", motivo="", quadrilha="", id_caso="", status="Ativo", analista=""):
    tipo = tipo.upper().strip()
    if tipo == "CPF":
        val_limpo = limpar_cpf_cnpj(valor_bruto)
        val_fmt = formatar_cpf_cnpj(val_limpo)
    elif tipo in ["TELEFONE", "TEL"]:
        val_limpo = limpar_tel(valor_bruto)
        val_fmt = formatar_tel(val_limpo)
        tipo = "TELEFONE"
    elif tipo == "PLACA":
        val_limpo = limpar_placa(valor_bruto)
        val_fmt = f"{val_limpo[:3]}-{val_limpo[3:]}" if len(val_limpo) == 7 else val_limpo
    else:
        return False, "Tipo inválido.", 0

    if not val_limpo:
        return False, "Dado inválido.", 0

    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        dt_agora = datetime.now().strftime("%d/%m/%Y %H:%M")
        cursor.execute("""
            INSERT INTO entidades_suspeitas (tipo, valor, valor_formatado, nome_referencia, motivo, id_caso, quadrilha_caso, status, analista, data_cadastro)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(tipo, valor) DO UPDATE SET
                nome_referencia = CASE WHEN excluded.nome_referencia != '' THEN excluded.nome_referencia ELSE entidades_suspeitas.nome_referencia END,
                motivo = CASE WHEN excluded.motivo != '' THEN excluded.motivo ELSE entidades_suspeitas.motivo END,
                id_caso = CASE WHEN excluded.id_caso != '' THEN excluded.id_caso ELSE entidades_suspeitas.id_caso END,
                quadrilha_caso = CASE WHEN excluded.quadrilha_caso != '' THEN excluded.quadrilha_caso ELSE entidades_suspeitas.quadrilha_caso END,
                status = excluded.status,
                analista = excluded.analista,
                data_cadastro = excluded.data_cadastro
        """, (tipo, val_limpo, val_fmt, nome_ref.strip().upper(), motivo.strip(), id_caso.strip(), quadrilha.strip().upper(), status, analista.strip(), dt_agora))
        conn.commit()

        col = "cpf" if tipo == "CPF" else ("telefone" if tipo == "TELEFONE" else "placa")
        cursor.execute(f"""
            SELECT 
                (SELECT COUNT(*) FROM assistencias WHERE {col} = ?) +
                (SELECT COUNT(*) FROM criacoes_diarias WHERE {col} = ?) as total
        """, (val_limpo, val_limpo))
        total_historico = cursor.fetchone()["total"]

        return True, f"✅ {tipo} {val_fmt} registrado na Base Mestra!", total_historico
    except Exception as e:
        return False, f"Erro: {str(e)}", 0
    finally:
        conn.close()


def listar_entidades_suspeitas(filtro_tipo="TODOS", filtro_status="TODOS", termo_busca=""):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        query = """
            SELECT e.*, c.nome_personalizado as nome_quadrilha_resolvido
            FROM entidades_suspeitas e
            LEFT JOIN casos_investigacao c ON e.id_caso = c.id_caso
            WHERE 1=1
        """
        params = []

        if filtro_tipo and filtro_tipo != "TODOS":
            query += " AND e.tipo = ?"
            params.append(filtro_tipo)
        if filtro_status and filtro_status != "TODOS":
            query += " AND e.status = ?"
            params.append(filtro_status)
        if termo_busca:
            clean_b = re.sub(r'[^a-zA-Z0-9]', '', termo_busca).upper()
            query += " AND (e.valor LIKE ? OR e.nome_referencia LIKE ? OR e.quadrilha_caso LIKE ? OR c.nome_personalizado LIKE ?)"
            params.extend([f"%{clean_b}%", f"%{termo_busca.upper()}%", f"%{termo_busca.upper()}%", f"%{termo_busca.upper()}%"])

        query += " ORDER BY e.id DESC LIMIT 300"
        cursor.execute(query, params)
        rows = [dict(r) for r in cursor.fetchall()]

        for r in rows:
            r["quadrilha_exibicao"] = r["nome_quadrilha_resolvido"] or r["quadrilha_caso"] or r["id_caso"] or "-"

        return rows
    finally:
        conn.close()


def remover_entidade_suspeita(id_entidade):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM entidades_suspeitas WHERE id = ?", (id_entidade,))
        conn.commit()
    finally:
        conn.close()


def consultar_todas_ocorrencias_entidade(tipo, valor):
    conn = get_db_connection()
    try:
        col = "cpf" if tipo == "CPF" else ("telefone" if tipo == "TELEFONE" else "placa")
        query = f"""
            SELECT 'Blacklist' as origem, data, id_assistencia, servico, titular, cpf, telefone, placa, bairro, cidade, uf
            FROM assistencias WHERE {col} = ?
            UNION ALL
            SELECT 'Criações Diárias' as origem, data, id_assistencia, servico, titular, cpf, telefone, placa, bairro, cidade, uf
            FROM criacoes_diarias WHERE {col} = ?
            ORDER BY data DESC
        """
        return pd.read_sql_query(query, conn, params=[valor, valor])
    finally:
        conn.close()


# =====================================================
# LEITURA E INGESTÃO INCREMENTAL VETORIZADA
# =====================================================
def detectar_encoding_e_separador(arq):
    encodings = ["utf-8-sig", "latin1", "cp1252", "utf-8"]
    separadores = [";", "\t", "|", ","]
    for encoding in encodings:
        try:
            with open(arq, "r", encoding=encoding, errors="strict") as f:
                amostra = f.read(150_000)
            linhas = [linha for linha in amostra.splitlines() if linha.strip()]
            if not linhas:
                raise ValueError("Arquivo vazio.")
            cabecalho = linhas[0]
            contagens = {sep: cabecalho.count(sep) for sep in separadores}
            separador = max(contagens, key=contagens.get)
            if contagens[separador] == 0:
                separador = ";"
            return encoding, separador
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Não foi possível identificar o encoding do arquivo {arq.name}")


def ler_arquivo_seguro(arq):
    ext = arq.suffix.lower()
    if ext == ".xlsx":
        return pd.read_excel(arq, dtype=str, engine="openpyxl")
    elif ext == ".csv":
        encoding, separador = detectar_encoding_e_separador(arq)
        parametros = {
            "filepath_or_buffer": arq, "sep": separador, "encoding": encoding,
            "dtype": str, "keep_default_na": False, "na_filter": False,
            "quotechar": '"', "low_memory": False
        }
        try:
            return pd.read_csv(engine="c", **parametros)
        except Exception:
            parametros.pop("low_memory", None)
            return pd.read_csv(engine="python", **parametros)
    else:
        raise ValueError(f"Formato não suportado: {ext}")


def carregar_arquivos_para_sqlite(caminho_pasta_str, forcar_releitura=False):
    pasta = Path(caminho_pasta_str.strip())
    if not pasta.exists():
        return 0, 0, "Pasta da Blacklist não encontrada.", []

    arquivos = list(pasta.glob("*.xlsx")) + list(pasta.glob("*.csv"))
    if not arquivos:
        return 0, 0, f"Nenhum arquivo encontrado em '{pasta}'.", []

    conn = get_db_connection()
    cursor = conn.cursor()
    
    if forcar_releitura:
        cursor.execute("DELETE FROM assistencias")
        cursor.execute("DELETE FROM arquivos_processados WHERE tipo_base = 'BLACKLIST'")
        conn.commit()

    cursor.execute("SELECT nome_arquivo FROM arquivos_processados WHERE tipo_base = 'BLACKLIST'")
    ja_processados = set(r["nome_arquivo"] for r in cursor.fetchall())
    conn.close()

    arquivos_para_ler = [a for a in arquivos if a.name not in ja_processados]
    if not arquivos_para_ler:
        return len(arquivos), 0, "Todos os arquivos já foram processados.", []

    total_inseridos = 0
    erros_arquivos = []
    progresso_barra = st.progress(0, text="Processando Blacklist...")

    for idx_arq, arq in enumerate(arquivos_para_ler):
        progresso_barra.progress((idx_arq + 1) / len(arquivos_para_ler), text=f"Lendo ({idx_arq + 1}/{len(arquivos_para_ler)}): {arq.name}")

        conn = get_db_connection()
        cursor = conn.cursor()
        try:
            df = ler_arquivo_seguro(arq)
            if df.empty:
                continue

            c_id = mapear_coluna(df, ["nro_assistencia", "numero_da_assistencia", "numero_assistencia", "assistencia"])
            c_tel = mapear_coluna(df, ["telefone_titular", "telefone_do_titular", "telefone", "celular", "contato"])
            c_cpf = mapear_coluna(df, ["cpf_cnpj_usuario", "cpf_cnpj", "cpf", "cnpj", "documento"], excluir_se_conter=["prestador", "cliente"])
            c_placa = mapear_coluna(df, ["placa", "veiculo"])
            c_nome = mapear_coluna(df, ["titular", "nome_titular", "segurado"], excluir_se_conter=["cliente", "prestador", "garantia"])
            c_serv = mapear_coluna(df, ["servico", "descricao_do_servico", "tipo_servico"])
            c_data = mapear_coluna(df, ["data_do_expediente", "data_inclusao_da_assistencia", "data_do_servico", "data_abertura", "data"])
            c_cid = mapear_coluna(df, ["cidade_ocorrencia", "cidade_de_ocorrencia", "cidade_origem", "cidade_de_origem", "cidade"])
            c_uf = mapear_coluna(df, ["estado_ocorrencia", "estado_de_ocorrencia", "estado_origem", "estado_de_origem", "uf", "estado"])
            c_bairro = mapear_coluna(df, ["bairro_ocorrencia", "bairro_de_ocorrencia", "bairro_origem", "bairro_de_origem", "bairro"])

            n_rows = len(df)
            s_arquivo = pd.Series([arq.name] * n_rows)
            s_id = df[c_id].astype(str).str.strip() if c_id else pd.Series([""] * n_rows)
            s_cpf = df[c_cpf].map(limpar_cpf_cnpj) if c_cpf else pd.Series([""] * n_rows)
            s_tel = df[c_tel].map(limpar_tel) if c_tel else pd.Series([""] * n_rows)
            s_placa = df[c_placa].map(limpar_placa) if c_placa else pd.Series([""] * n_rows)
            s_nome = df[c_nome].map(normalizar_texto) if c_nome else pd.Series([""] * n_rows)
            s_cid = df[c_cid].map(normalizar_texto) if c_cid else pd.Series(["SAO PAULO"] * n_rows)
            s_uf = df[c_uf].map(normalizar_texto) if c_uf else pd.Series(["SP"] * n_rows)
            s_bairro = df[c_bairro].map(normalizar_texto) if c_bairro else pd.Series([""] * n_rows)
            s_serv = df[c_serv].map(normalizar_texto) if c_serv else pd.Series(["ASSISTENCIA"] * n_rows)
            s_data = df[c_data].map(lambda v: tratar_data_flexivel(v, arq.name)) if c_data else pd.Series([extrair_data_arquivo(arq.name)] * n_rows)

            coords = [obter_coordenadas(c, u) for c, u in zip(s_cid, s_uf)]
            s_lat = pd.Series([c[0] for c in coords])
            s_lon = pd.Series([c[1] for c in coords])

            df_saneado = pd.DataFrame({
                "arquivo": s_arquivo, "id_assistencia": s_id, "data": s_data, "titular": s_nome,
                "cpf": s_cpf, "telefone": s_tel, "placa": s_placa, "servico": s_serv,
                "bairro": s_bairro, "cidade": s_cid, "uf": s_uf, "latitude": s_lat, "longitude": s_lon
            })

            mask_valido = (df_saneado["id_assistencia"] != "") & (
                (df_saneado["cpf"] != "") | (df_saneado["telefone"] != "") | (df_saneado["placa"] != "")
            )
            df_final = df_saneado[mask_valido]

            inseridos_neste_arq = 0
            if not df_final.empty:
                lote = list(df_final.itertuples(index=False, name=None))
                antes = conn.total_changes
                cursor.executemany("""
                    INSERT OR IGNORE INTO assistencias (arquivo, id_assistencia, data, titular, cpf, telefone, placa, servico, bairro, cidade, uf, latitude, longitude)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, lote)
                inseridos_neste_arq = conn.total_changes - antes
                total_inseridos += inseridos_neste_arq

            cursor.execute("""
                INSERT OR REPLACE INTO arquivos_processados (nome_arquivo, tipo_base, data_processamento, registros_inseridos)
                VALUES (?, 'BLACKLIST', datetime('now'), ?)
            """, (arq.name, inseridos_neste_arq))
            conn.commit()

        except Exception as e:
            conn.rollback()
            erros_arquivos.append(f"{arq.name}: {str(e)[:90]}")
        finally:
            cursor.close()
            conn.close()

    semear_base_mestra_da_blacklist()

    progresso_barra.empty()
    return len(arquivos), total_inseridos, "", erros_arquivos


def carregar_criacoes_diarias_para_sqlite(caminho_pasta_str, forcar_releitura=False):
    pasta = Path(caminho_pasta_str.strip())
    if not pasta.exists():
        return 0, 0, "Pasta de Criações Diárias não encontrada.", []

    arquivos = list(pasta.glob("*.xlsx")) + list(pasta.glob("*.csv"))
    if not arquivos:
        return 0, 0, f"Nenhum arquivo encontrado em '{pasta}'.", []

    conn = get_db_connection()
    cursor = conn.cursor()

    if forcar_releitura:
        cursor.execute("DELETE FROM criacoes_diarias")
        cursor.execute("DELETE FROM arquivos_processados WHERE tipo_base = 'CRIACAO'")
        conn.commit()

    cursor.execute("SELECT nome_arquivo FROM arquivos_processados WHERE tipo_base = 'CRIACAO'")
    ja_processados = set(r["nome_arquivo"] for r in cursor.fetchall())
    conn.close()

    arquivos_para_ler = [a for a in arquivos if a.name not in ja_processados]
    if not arquivos_para_ler:
        return len(arquivos), 0, "Todas as planilhas já foram processadas anteriormente.", []

    total_inseridos = 0
    erros_arquivos = []
    barra = st.progress(0, text="Processando Criações...")

    for idx_arq, arq in enumerate(arquivos_para_ler):
        barra.progress((idx_arq + 1) / len(arquivos_para_ler), text=f"Ingerindo Criações ({idx_arq + 1}/{len(arquivos_para_ler)}): {arq.name}")

        conn = get_db_connection()
        cursor = conn.cursor()
        try:
            df = ler_arquivo_seguro(arq)
            if df.empty:
                continue

            c_id = mapear_coluna(df, ["nro_assistencia", "numero_da_assistencia", "numero_assistencia", "assistencia"])
            c_tel = mapear_coluna(df, ["telefone_titular", "telefone_do_titular", "telefone", "celular", "contato"])
            c_cpf = mapear_coluna(df, ["cpf_cnpj_usuario", "cpf_cnpj", "cpf", "cnpj", "documento"], excluir_se_conter=["prestador", "cliente"])
            c_placa = mapear_coluna(df, ["placa", "veiculo"])
            c_nome = mapear_coluna(df, ["titular", "nome_titular", "segurado"], excluir_se_conter=["cliente", "prestador", "garantia"])
            c_serv = mapear_coluna(df, ["servico", "descricao_do_servico", "tipo_servico"])
            c_data = mapear_coluna(df, ["data_do_expediente", "data_inclusao_da_assistencia", "data_do_servico", "data_abertura", "data"])
            c_cid = mapear_coluna(df, ["cidade_ocorrencia", "cidade_de_ocorrencia", "cidade_origem", "cidade_de_origem", "cidade"])
            c_uf = mapear_coluna(df, ["estado_ocorrencia", "estado_de_ocorrencia", "estado_origem", "estado_de_origem", "uf", "estado"])
            c_bairro = mapear_coluna(df, ["bairro_ocorrencia", "bairro_de_ocorrencia", "bairro_origem", "bairro_de_origem", "bairro"])

            n_rows = len(df)
            s_arquivo = pd.Series([arq.name] * n_rows)
            s_id = df[c_id].astype(str).str.strip() if c_id else pd.Series([""] * n_rows)
            s_cpf = df[c_cpf].map(limpar_cpf_cnpj) if c_cpf else pd.Series([""] * n_rows)
            s_tel = df[c_tel].map(limpar_tel) if c_tel else pd.Series([""] * n_rows)
            s_placa = df[c_placa].map(limpar_placa) if c_placa else pd.Series([""] * n_rows)
            s_nome = df[c_nome].map(normalizar_texto) if c_nome else pd.Series([""] * n_rows)
            s_cid = df[c_cid].map(normalizar_texto) if c_cid else pd.Series(["SAO PAULO"] * n_rows)
            s_uf = df[c_uf].map(normalizar_texto) if c_uf else pd.Series(["SP"] * n_rows)
            s_bairro = df[c_bairro].map(normalizar_texto) if c_bairro else pd.Series([""] * n_rows)
            s_serv = df[c_serv].map(normalizar_texto) if c_serv else pd.Series(["ASSISTENCIA"] * n_rows)
            s_data = df[c_data].map(lambda v: tratar_data_flexivel(v, arq.name)) if c_data else pd.Series([extrair_data_arquivo(arq.name)] * n_rows)

            coords = [obter_coordenadas(c, u) for c, u in zip(s_cid, s_uf)]
            s_lat = pd.Series([c[0] for c in coords])
            s_lon = pd.Series([c[1] for c in coords])

            df_saneado = pd.DataFrame({
                "arquivo": s_arquivo, "id_assistencia": s_id, "data": s_data, "titular": s_nome,
                "cpf": s_cpf, "telefone": s_tel, "placa": s_placa, "servico": s_serv,
                "bairro": s_bairro, "cidade": s_cid, "uf": s_uf, "latitude": s_lat, "longitude": s_lon
            })

            mask_valido = (df_saneado["id_assistencia"] != "") & (
                (df_saneado["cpf"] != "") | (df_saneado["telefone"] != "") | (df_saneado["placa"] != "")
            )
            df_final = df_saneado[mask_valido]

            inseridos_neste_arq = 0
            if not df_final.empty:
                lote = list(df_final.itertuples(index=False, name=None))
                antes = conn.total_changes
                cursor.executemany("""
                    INSERT OR IGNORE INTO criacoes_diarias (arquivo, id_assistencia, data, titular, cpf, telefone, placa, servico, bairro, cidade, uf, latitude, longitude)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, lote)
                inseridos_neste_arq = conn.total_changes - antes
                total_inseridos += inseridos_neste_arq

            cursor.execute("""
                INSERT OR REPLACE INTO arquivos_processados (nome_arquivo, tipo_base, data_processamento, registros_inseridos)
                VALUES (?, 'CRIACAO', datetime('now'), ?)
            """, (arq.name, inseridos_neste_arq))
            conn.commit()

        except Exception as e:
            conn.rollback()
            erros_arquivos.append(f"{arq.name}: {str(e)[:90]}")
        finally:
            cursor.close()
            conn.close()

    barra.empty()
    return len(arquivos), total_inseridos, "", erros_arquivos


def consultar_detalhes_caso(cpfs, tels, placas):
    conn = get_db_connection()
    try:
        clausulas, params = [], []
        if cpfs:
            clausulas.append(f"cpf IN ({','.join(['?']*len(cpfs))})")
            params.extend(cpfs)
        if tels:
            clausulas.append(f"telefone IN ({','.join(['?']*len(tels))})")
            params.extend(tels)
        if placas:
            clausulas.append(f"placa IN ({','.join(['?']*len(placas))})")
            params.extend(placas)

        if not clausulas:
            return pd.DataFrame()

        query = f"""
            SELECT data, id_assistencia, titular, cpf, telefone, placa, servico, bairro, cidade, uf, latitude, longitude
            FROM assistencias 
            WHERE {' OR '.join(clausulas)}
        """
        return pd.read_sql_query(query, conn, params=params)
    finally:
        conn.close()


def cruzar_com_criacoes_diarias(cpfs_conhecidos, tels_conhecidos, placas_conhecidas):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='criacoes_diarias'")
        if not cursor.fetchone():
            return None, pd.DataFrame(), pd.DataFrame()

        clausulas, params = [], []
        if cpfs_conhecidos:
            clausulas.append(f"cpf IN ({','.join(['?']*len(cpfs_conhecidos))})")
            params.extend(cpfs_conhecidos)
        if tels_conhecidos:
            clausulas.append(f"telefone IN ({','.join(['?']*len(tels_conhecidos))})")
            params.extend(tels_conhecidos)
        if placas_conhecidas:
            clausulas.append(f"placa IN ({','.join(['?']*len(placas_conhecidas))})")
            params.extend(placas_conhecidas)

        if not clausulas:
            return None, pd.DataFrame(), pd.DataFrame()

        query = f"""
            SELECT data, id_assistencia, titular, cpf, telefone, placa, servico, bairro, cidade, uf
            FROM criacoes_diarias
            WHERE {' OR '.join(clausulas)}
        """
        df_matches = pd.read_sql_query(query, conn, params=params)
    finally:
        conn.close()

    if df_matches.empty:
        return {"total_assistencias": 0, "novos_tels": 0, "novos_cpfs": 0, "novas_placas": 0}, df_matches, pd.DataFrame()

    set_cpfs = set(cpfs_conhecidos)
    set_tels = set(tels_conhecidos)
    set_placas = set(placas_conhecidas)

    novos_suspeitos = []
    for _, r in df_matches.iterrows():
        r_cpf = str(r["cpf"]).strip() if pd.notna(r["cpf"]) else ""
        r_tel = str(r["telefone"]).strip() if pd.notna(r["telefone"]) else ""
        r_placa = str(r["placa"]).strip() if pd.notna(r["placa"]) else ""
        r_nome = str(r["titular"]).strip() if pd.notna(r["titular"]) else ""
        r_id = str(r["id_assistencia"]).strip() if pd.notna(r["id_assistencia"]) else ""

        elos = []
        if r_cpf in set_cpfs: elos.append(f"CPF: {r_cpf}")
        if r_tel in set_tels: elos.append(f"Tel: {r_tel}")
        if r_placa in set_placas: elos.append(f"Placa: {r_placa}")
        elo_str = " | ".join(elos)

        if r_cpf and r_cpf not in set_cpfs:
            novos_suspeitos.append({
                "Tipo": "👤 CPF Inédito", "Dado Suspeito": r_cpf, "Titular": r_nome,
                "Elo com a Blacklist": elo_str, "Nº Assistência": r_id, "Data": r["data"], "Local": f"{r['cidade']}/{r['uf']}"
            })
        if r_tel and r_tel not in set_tels:
            novos_suspeitos.append({
                "Tipo": "🚨 Telefone Inédito", "Dado Suspeito": r_tel, "Titular": r_nome,
                "Elo com a Blacklist": elo_str, "Nº Assistência": r_id, "Data": r["data"], "Local": f"{r['cidade']}/{r['uf']}"
            })
        if r_placa and r_placa not in set_placas:
            novos_suspeitos.append({
                "Tipo": "🚗 Placa Inédita", "Dado Suspeito": r_placa, "Titular": r_nome,
                "Elo com a Blacklist": elo_str, "Nº Assistência": r_id, "Data": r["data"], "Local": f"{r['cidade']}/{r['uf']}"
            })

    df_novos = pd.DataFrame(novos_suspeitos).drop_duplicates(subset=["Tipo", "Dado Suspeito"])

    resumo = {
        "total_assistencias": len(df_matches),
        "novos_tels": len(df_novos[df_novos["Tipo"] == "🚨 Telefone Inédito"]) if not df_novos.empty else 0,
        "novos_cpfs": len(df_novos[df_novos["Tipo"] == "👤 CPF Inédito"]) if not df_novos.empty else 0,
        "novas_placas": len(df_novos[df_novos["Tipo"] == "🚗 Placa Inédita"]) if not df_novos.empty else 0,
    }

    return resumo, df_matches, df_novos


# =====================================================
# RADAR DE EXPANSÕES BLINDADO COM IMPRESSÃO DIGITAL DE DADOS
# =====================================================
def obter_versao_criacoes():
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) as total, MAX(id) as max_id FROM criacoes_diarias")
        row = cursor.fetchone()
        if row and row["total"]:
            return (row["total"], row["max_id"] or 0)
        return (0, 0)
    except Exception:
        return (0, 0)
    finally:
        conn.close()


@st.cache_data
def obter_radar_expansoes_cached(ids_tupla, nos_serializados, versao_dados):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='criacoes_diarias'")
        if not cursor.fetchone():
            return {}

        radar = {}
        for cid, nodes in zip(ids_tupla, nos_serializados):
            cpfs = [n.replace("CPF_", "") for n in nodes if n.startswith("CPF_")]
            tels = [n.replace("TEL_", "") for n in nodes if n.startswith("TEL_")]
            placas = [n.replace("PLACA_", "") for n in nodes if n.startswith("PLACA_")]

            clausulas, params = [], []
            if cpfs:
                clausulas.append(f"cpf IN ({','.join(['?']*len(cpfs))})")
                params.extend(cpfs)
            if tels:
                clausulas.append(f"telefone IN ({','.join(['?']*len(tels))})")
                params.extend(tels)
            if placas:
                clausulas.append(f"placa IN ({','.join(['?']*len(placas))})")
                params.extend(placas)

            if clausulas:
                query = f"""SELECT COUNT(*) as total FROM criacoes_diarias WHERE {' OR '.join(clausulas)}"""
                cursor.execute(query, params)
                qtd = cursor.fetchone()["total"]
                if qtd > 0:
                    radar[cid] = qtd
        return radar
    finally:
        conn.close()


def obter_radar_expansoes(cluster_info):
    if not cluster_info:
        return {}
    top_100 = cluster_info[:100]
    ids_tupla = tuple(c["id"] for c in top_100)
    nos_serializados = tuple(tuple(sorted(c["nodes"])) for c in top_100)
    versao_atual = obter_versao_criacoes()
    return obter_radar_expansoes_cached(ids_tupla, nos_serializados, versao_atual)