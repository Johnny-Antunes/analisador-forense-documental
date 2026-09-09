import sqlite3
import pandas as pd
from pathlib import Path
import streamlit as st
from utils import (
    mapear_coluna, limpar_cpf_cnpj, limpar_tel, limpar_placa,
    normalizar_texto, tratar_data_flexivel, obter_coordenadas
)

DB_PATH = Path(__file__).parent / "banco_fraudes.db"

# Caminho corporativo oficial com fallback para pasta local
CAMINHO_REDE_OFICIAL = r"T:\Corporativo\OPERAÇÕES\CPO\QUERIES\Verificação Blacklist - Data Criação ME\Duplicidade de Placa"
PASTA_LOCAL = Path(__file__).parent / "casos_duplicidade"
PASTA_LOCAL.mkdir(exist_ok=True)

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def garantir_schema_db():
    if not DB_PATH.exists():
        return
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("PRAGMA table_info(assistencias)")
        cols = [r["name"] for r in cursor.fetchall()]
        if cols:
            if "servico" not in cols:
                cursor.execute("ALTER TABLE assistencias ADD COLUMN servico TEXT DEFAULT 'ASSISTENCIA'")
            if "bairro" not in cols:
                cursor.execute("ALTER TABLE assistencias ADD COLUMN bairro TEXT DEFAULT ''")
            conn.commit()
    except Exception:
        pass
    finally:
        conn.close()

# Executa a validação de schema assim que o módulo é importado
garantir_schema_db()

def carregar_arquivos_para_sqlite(caminho_pasta_str):
    pasta = Path(caminho_pasta_str.strip())
    if not pasta.exists():
        return 0, 0, "Pasta não encontrada no caminho informado."

    arquivos = list(pasta.glob("*.xlsx"))
    if not arquivos:
        return 0, 0, f"Nenhuma planilha .xlsx encontrada em '{pasta}'."

    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DROP TABLE IF EXISTS assistencias")
    cursor.execute("""
        CREATE TABLE assistencias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            arquivo TEXT,
            id_assistencia TEXT UNIQUE,
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
            longitude REAL
        )
    """)
    cursor.execute("CREATE INDEX idx_cpf ON assistencias(cpf)")
    cursor.execute("CREATE INDEX idx_tel ON assistencias(telefone)")
    cursor.execute("CREATE INDEX idx_placa ON assistencias(placa)")

    total_inseridos = 0
    progresso_barra = st.progress(0, text="Processando planilhas operacionais...")

    for idx_arq, arq in enumerate(arquivos):
        progresso_barra.progress((idx_arq + 1) / len(arquivos), text=f"Lendo ({idx_arq + 1}/{len(arquivos)}): {arq.name}")
        try:
            df = pd.read_excel(arq)
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
                cursor.executemany("""
                    INSERT OR IGNORE INTO assistencias (arquivo, id_assistencia, data, titular, cpf, telefone, placa, servico, bairro, cidade, uf, latitude, longitude)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, lote)
                total_inseridos += len(lote)

        except Exception as e:
            st.error(f"Erro ao processar {arq.name}: {e}")

    conn.commit()
    conn.close()
    progresso_barra.empty()
    return len(arquivos), total_inseridos, ""

def consultar_detalhes_caso(cpfs, tels, placas):
    """Executa a busca dos registros da célula ativa no banco SQLite."""
    conn = get_db_connection()
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
        conn.close()
        return pd.DataFrame()

    where_sql = " OR ".join(clausulas)
    query = f"""
        SELECT data, id_assistencia, titular, cpf, telefone, placa, servico, bairro, cidade, uf, latitude, longitude
        FROM assistencias 
        WHERE {where_sql}
    """
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()
    return df