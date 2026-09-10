import sqlite3
import pandas as pd
from pathlib import Path
import streamlit as st
import csv
import re

from utils import (
    mapear_coluna, limpar_cpf_cnpj, limpar_tel, limpar_placa,
    normalizar_texto, tratar_data_flexivel, obter_coordenadas
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
    """Cria conexão otimizada com o SQLite utilizando WAL Mode para alta velocidade de I/O."""
    conn = sqlite3.connect(DB_PATH, timeout=30.0)
    conn.row_factory = sqlite3.Row
    # Otimizações de performance para grandes volumes corporativos
    conn.execute("PRAGMA journal_mode = WAL;")
    conn.execute("PRAGMA synchronous = NORMAL;")
    conn.execute("PRAGMA cache_size = -64000;")  # Cache de 64MB em RAM para consultas rápidas
    return conn


def garantir_schema_db():
    if not DB_PATH.exists():
        return
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(assistencias)")
        cols = [r["name"] for r in cursor.fetchall()]
        if cols:
            if "servico" not in cols:
                cursor.execute("ALTER TABLE assistencias ADD COLUMN servico TEXT DEFAULT 'ASSISTENCIA'")
            if "bairro" not in cols:
                cursor.execute("ALTER TABLE assistencias ADD COLUMN bairro TEXT DEFAULT ''")

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
        conn.commit()
    except Exception:
        pass
    finally:
        conn.close()

garantir_schema_db()


def detectar_encoding_e_separador(arq):
    encodings = ["utf-8-sig", "latin1", "cp1252", "utf-8"]
    separadores = [";", "\t", "|", ","]
    ultimo_erro = None

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
        except UnicodeDecodeError as e:
            ultimo_erro = e
            continue

    raise ValueError(f"Não foi possível identificar o encoding do arquivo {arq.name}: {ultimo_erro}")


def ler_arquivo_seguro(arq):
    ext = arq.suffix.lower()
    if ext == ".xlsx":
        return pd.read_excel(arq, dtype=str, engine="openpyxl")
    elif ext == ".csv":
        encoding, separador = detectar_encoding_e_separador(arq)
        
        parametros = {
            "filepath_or_buffer": arq,
            "sep": separador,
            "encoding": encoding,
            "dtype": str,
            "keep_default_na": False,
            "na_filter": False,
            "quotechar": '"',
            "low_memory": False
        }
        try:
            return pd.read_csv(engine="c", **parametros)
        except Exception:
            parametros.pop("low_memory", None)
            return pd.read_csv(engine="python", **parametros)
    else:
        raise ValueError(f"Formato não suportado: {ext}")


def carregar_arquivos_para_sqlite(caminho_pasta_str):
    pasta = Path(caminho_pasta_str.strip())
    if not pasta.exists():
        return 0, 0, "Pasta da Blacklist não encontrada.", []

    arquivos = list(pasta.glob("*.xlsx")) + list(pasta.glob("*.csv"))
    if not arquivos:
        return 0, 0, f"Nenhum arquivo encontrado em '{pasta}'.", []

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DROP TABLE IF EXISTS assistencias")
        cursor.execute("""
            CREATE TABLE assistencias (
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
        cursor.execute("CREATE INDEX idx_cpf ON assistencias(cpf)")
        cursor.execute("CREATE INDEX idx_tel ON assistencias(telefone)")
        cursor.execute("CREATE INDEX idx_placa ON assistencias(placa)")
        conn.commit()
    finally:
        cursor.close()
        conn.close()

    total_inseridos = 0
    erros_arquivos = []
    progresso_barra = st.progress(0, text="Processando planilhas da Blacklist...")

    for idx_arq, arq in enumerate(arquivos):
        progresso_barra.progress((idx_arq + 1) / len(arquivos), text=f"Lendo Blacklist ({idx_arq + 1}/{len(arquivos)}): {arq.name}")
        
        conn = get_db_connection()
        cursor = conn.cursor()
        try:
            df = ler_arquivo_seguro(arq)
            if df.empty:
                continue

            col_id = mapear_coluna(df, ["nro_assistencia", "numero_da_assistencia", "numero_assistencia", "assistencia"])
            col_tel = mapear_coluna(df, ["telefone_titular", "telefone_do_titular", "telefone", "celular", "contato"])
            col_cpf = mapear_coluna(df, ["cpf_cnpj_usuario", "cpf_cnpj", "cpf", "cnpj", "documento"], excluir_se_conter=["prestador", "cliente"])
            col_placa = mapear_coluna(df, ["placa", "veiculo"])
            col_nome = mapear_coluna(df, ["titular", "nome_titular", "segurado"], excluir_se_conter=["cliente", "prestador", "garantia"])
            col_servico = mapear_coluna(df, ["servico", "descricao_do_servico", "tipo_servico"])
            col_data = mapear_coluna(df, ["data_do_expediente", "data_inclusao_da_assistencia", "data_do_servico", "data_abertura", "data"])
            
            col_cid = mapear_coluna(df, ["cidade_ocorrencia", "cidade_de_ocorrencia", "cidade_origem", "cidade_de_origem", "cidade"])
            col_uf = mapear_coluna(df, ["estado_ocorrencia", "estado_de_ocorrencia", "estado_origem", "estado_de_origem", "uf", "estado"])
            col_bairro = mapear_coluna(df, ["bairro_ocorrencia", "bairro_de_ocorrencia", "bairro_origem", "bairro_de_origem", "bairro"])

            lote = []
            for _, r in df.iterrows():
                ida = str(r[col_id]).strip() if col_id and pd.notna(r[col_id]) else ""
                cpf = limpar_cpf_cnpj(r[col_cpf]) if col_cpf else ""
                tel = limpar_tel(r[col_tel]) if col_tel else ""
                placa = limpar_placa(r[col_placa]) if col_placa else ""
                nome = normalizar_texto(r[col_nome]) if col_nome else ""
                cid = normalizar_texto(r[col_cid]) if col_cid else "SAO PAULO"
                uf = normalizar_texto(r[col_uf]) if col_uf else "SP"
                bairro = normalizar_texto(r[col_bairro]) if col_bairro else ""
                servico = normalizar_texto(r[col_servico]) if col_servico else "ASSISTENCIA"
                
                raw_dt = r[col_data] if col_data and pd.notna(r[col_data]) else None
                dt = tratar_data_flexivel(raw_dt, arq.name)
                lat, lon = obter_coordenadas(cid, uf)

                if ida and (cpf or tel or placa):
                    lote.append((arq.name, ida, dt, nome, cpf, tel, placa, servico, bairro, cid, uf, lat, lon))

            if lote:
                antes = conn.total_changes
                cursor.executemany("""
                    INSERT OR IGNORE INTO assistencias (arquivo, id_assistencia, data, titular, cpf, telefone, placa, servico, bairro, cidade, uf, latitude, longitude)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, lote)
                total_inseridos += (conn.total_changes - antes)
                conn.commit()

        except Exception as e:
            conn.rollback()
            erros_arquivos.append(f"{arq.name}: {str(e)[:90]}")
        finally:
            cursor.close()
            conn.close()

    progresso_barra.empty()
    return len(arquivos), total_inseridos, "", erros_arquivos


def carregar_criacoes_diarias_para_sqlite(caminho_pasta_str):
    pasta = Path(caminho_pasta_str.strip())
    if not pasta.exists():
        return 0, 0, "Pasta de Criações Diárias não encontrada.", []

    arquivos = list(pasta.glob("*.xlsx")) + list(pasta.glob("*.csv"))
    if not arquivos:
        return 0, 0, f"Nenhum arquivo encontrado em '{pasta}'.", []

    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("DROP TABLE IF EXISTS criacoes_diarias")
        cursor.execute("""
            CREATE TABLE criacoes_diarias (
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
        cursor.execute("CREATE INDEX idx_criacao_cpf ON criacoes_diarias(cpf)")
        cursor.execute("CREATE INDEX idx_criacao_tel ON criacoes_diarias(telefone)")
        cursor.execute("CREATE INDEX idx_criacao_placa ON criacoes_diarias(placa)")
        conn.commit()
    finally:
        cursor.close()
        conn.close()

    total_inseridos = 0
    erros_arquivos = []
    barra = st.progress(0, text="Processando Criações Diárias...")

    for idx_arq, arq in enumerate(arquivos):
        barra.progress((idx_arq + 1) / len(arquivos), text=f"Ingerindo Criações ({idx_arq + 1}/{len(arquivos)}): {arq.name}")
        
        conn = get_db_connection()
        cursor = conn.cursor()
        try:
            df = ler_arquivo_seguro(arq)
            if df.empty:
                continue

            col_id = mapear_coluna(df, ["nro_assistencia", "numero_da_assistencia", "numero_assistencia", "assistencia"])
            col_tel = mapear_coluna(df, ["telefone_titular", "telefone_do_titular", "telefone", "celular", "contato"])
            col_cpf = mapear_coluna(df, ["cpf_cnpj_usuario", "cpf_cnpj", "cpf", "cnpj", "documento"], excluir_se_conter=["prestador", "cliente"])
            col_placa = mapear_coluna(df, ["placa", "veiculo"])
            col_nome = mapear_coluna(df, ["titular", "nome_titular", "segurado"], excluir_se_conter=["cliente", "prestador", "garantia"])
            col_servico = mapear_coluna(df, ["servico", "descricao_do_servico", "tipo_servico"])
            col_data = mapear_coluna(df, ["data_do_expediente", "data_inclusao_da_assistencia", "data_do_servico", "data_abertura", "data"])
            
            col_cid = mapear_coluna(df, ["cidade_ocorrencia", "cidade_de_ocorrencia", "cidade_origem", "cidade_de_origem", "cidade"])
            col_uf = mapear_coluna(df, ["estado_ocorrencia", "estado_de_ocorrencia", "estado_origem", "estado_de_origem", "uf", "estado"])
            col_bairro = mapear_coluna(df, ["bairro_ocorrencia", "bairro_de_ocorrencia", "bairro_origem", "bairro_de_origem", "bairro"])

            lote = []
            for _, r in df.iterrows():
                ida = str(r[col_id]).strip() if col_id and pd.notna(r[col_id]) else ""
                cpf = limpar_cpf_cnpj(r[col_cpf]) if col_cpf else ""
                tel = limpar_tel(r[col_tel]) if col_tel else ""
                placa = limpar_placa(r[col_placa]) if col_placa else ""
                nome = normalizar_texto(r[col_nome]) if col_nome else ""
                cid = normalizar_texto(r[col_cid]) if col_cid else "SAO PAULO"
                uf = normalizar_texto(r[col_uf]) if col_uf else "SP"
                bairro = normalizar_texto(r[col_bairro]) if col_bairro else ""
                servico = normalizar_texto(r[col_servico]) if col_servico else "ASSISTENCIA"
                
                raw_dt = r[col_data] if col_data and pd.notna(r[col_data]) else None
                dt = tratar_data_flexivel(raw_dt, arq.name)
                lat, lon = obter_coordenadas(cid, uf)

                if ida and (cpf or tel or placa):
                    lote.append((arq.name, ida, dt, nome, cpf, tel, placa, servico, bairro, cid, uf, lat, lon))

            if lote:
                antes = conn.total_changes
                cursor.executemany("""
                    INSERT OR IGNORE INTO criacoes_diarias (arquivo, id_assistencia, data, titular, cpf, telefone, placa, servico, bairro, cidade, uf, latitude, longitude)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, lote)
                total_inseridos += (conn.total_changes - antes)
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
        clausulas = []
        params = []
        
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

        clausulas = []
        params = []
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
                "Tipo": "👤 CPF Inédito",
                "Dado Suspeito": r_cpf,
                "Titular": r_nome,
                "Elo com a Blacklist": elo_str,
                "Nº Assistência": r_id,
                "Data": r["data"],
                "Local": f"{r['cidade']}/{r['uf']}"
            })
        if r_tel and r_tel not in set_tels:
            novos_suspeitos.append({
                "Tipo": "🚨 Telefone Inédito",
                "Dado Suspeito": r_tel,
                "Titular": r_nome,
                "Elo com a Blacklist": elo_str,
                "Nº Assistência": r_id,
                "Data": r["data"],
                "Local": f"{r['cidade']}/{r['uf']}"
            })
        if r_placa and r_placa not in set_placas:
            novos_suspeitos.append({
                "Tipo": "🚗 Placa Inédita",
                "Dado Suspeito": r_placa,
                "Titular": r_nome,
                "Elo com a Blacklist": elo_str,
                "Nº Assistência": r_id,
                "Data": r["data"],
                "Local": f"{r['cidade']}/{r['uf']}"
            })

    df_novos = pd.DataFrame(novos_suspeitos).drop_duplicates(subset=["Tipo", "Dado Suspeito"])

    resumo = {
        "total_assistencias": len(df_matches),
        "novos_tels": len(df_novos[df_novos["Tipo"] == "🚨 Telefone Inédito"]) if not df_novos.empty else 0,
        "novos_cpfs": len(df_novos[df_novos["Tipo"] == "👤 CPF Inédito"]) if not df_novos.empty else 0,
        "novas_placas": len(df_novos[df_novos["Tipo"] == "🚗 Placa Inédita"]) if not df_novos.empty else 0,
    }

    return resumo, df_matches, df_novos


def obter_radar_expansoes(cluster_info):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='criacoes_diarias'")
        if not cursor.fetchone():
            return {}

        radar = {}
        for c in cluster_info[:50]:
            cpfs = [n.replace("CPF_", "") for n in c["nodes"] if n.startswith("CPF_")]
            tels = [n.replace("TEL_", "") for n in c["nodes"] if n.startswith("TEL_")]
            placas = [n.replace("PLACA_", "") for n in c["nodes"] if n.startswith("PLACA_")]

            clausulas = []
            params = []
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
                query = f"SELECT COUNT(*) as total FROM criacoes_diarias WHERE {' OR '.join(clausulas)}"
                cursor.execute(query, params)
                qtd = cursor.fetchone()["total"]
                if qtd > 0:
                    radar[c["id"]] = qtd
        return radar
    finally:
        conn.close()