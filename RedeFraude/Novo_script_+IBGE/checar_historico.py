import sqlite3
from pathlib import Path
import pandas as pd

DB_PATH = Path(__file__).parent / "banco_fraudes.db"

def verificar_meses_historico():
    if not DB_PATH.exists():
        print(f"❌ Banco não encontrado: {DB_PATH}")
        return

    conn = sqlite3.connect(DB_PATH)
    query = """
        SELECT 
            strftime('%Y-%m', data) as mes, 
            COUNT(*) as total_acionamentos,
            COUNT(DISTINCT cidade) as cidades_atendidas
        FROM criacoes_diarias 
        WHERE data IS NOT NULL AND data != ''
        GROUP BY mes 
        ORDER BY mes ASC;
    """
    df = pd.read_sql_query(query, conn)
    conn.close()

    print("\n" + "="*50)
    print("📊 PROFUNDIDADE HISTÓRICA - CRIAÇÕES DIÁRIAS")
    print("="*50)
    if df.empty:
        print("Nenhum acionamento com data válida encontrado.")
    else:
        print(df.to_string(index=False))
        print("="*50)
        total_meses = len(df)
        print(f"Total de meses com registro: {total_meses}")
        
        if total_meses >= 10:
            print("🟢 Base pronta para cálculo de Baseline Macro (12 meses).")
        else:
            print(f"🟡 Atenção: Apenas {total_meses} meses mapeados. Baseline estatístico instável no momento.")
    print("="*50 + "\n")

if __name__ == "__main__":
    verificar_meses_historico()
