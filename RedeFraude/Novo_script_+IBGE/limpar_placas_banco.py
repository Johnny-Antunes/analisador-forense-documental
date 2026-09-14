import sqlite3
from pathlib import Path

DB_PATH = Path(__file__).parent / "banco_fraudes.db"

def executar_limpeza_placas_mascaradas():
    if not DB_PATH.exists():
        print(f"❌ Banco de dados não encontrado em: {DB_PATH}")
        return

    print("Iniciando varredura e limpeza de placas mascaradas...")
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    try:
        # 1. Zera a placa nas assistências da Blacklist (apenas onde tiver XXXXX)
        cursor.execute("""
            UPDATE assistencias 
            SET placa = '' 
            WHERE placa LIKE '%XXXX%' OR placa GLOB '[A-Z][A-Z]XXXXX'
        """)
        afetados_bl = cursor.rowcount

        # 2. Zera a placa nas Criações Diárias (apenas onde tiver XXXXX)
        cursor.execute("""
            UPDATE criacoes_diarias 
            SET placa = '' 
            WHERE placa LIKE '%XXXX%' OR placa GLOB '[A-Z][A-Z]XXXXX'
        """)
        afetados_cr = cursor.rowcount

        # 3. Remove o nó fantasma da tabela de membros (para não prender o hash antigo)
        cursor.execute("""
            DELETE FROM caso_membros 
            WHERE node_id LIKE 'PLACA_%XXXX%' OR node_id GLOB 'PLACA_[A-Z][A-Z]XXXXX'
        """)
        afetados_membros = cursor.rowcount

        # 4. Remove da Base Mestra se tiver sido cadastrado acidentalmente
        cursor.execute("""
            DELETE FROM entidades_suspeitas 
            WHERE tipo = 'PLACA' AND (valor LIKE '%XXXX%' OR valor GLOB '[A-Z][A-Z]XXXXX')
        """)
        afetados_mestra = cursor.rowcount

        conn.commit()

        print("\n✅ Limpeza concluída com absoluto sucesso!")
        print(f"• Tabela 'assistencias' (Blacklist): {afetados_bl} placas mascaradas zeradas.")
        print(f"• Tabela 'criacoes_diarias': {afetados_cr} placas mascaradas zeradas.")
        print(f"• Nós fantasmas removidos de 'caso_membros': {afetados_membros}")
        print(f"• Entidades fictícias removidas da Base Mestra: {afetados_mestra}")
        print("\nOs demais dados (CPFs, telefones, cidades, serviços e datas) continuam 100% intactos.")

    except Exception as e:
        conn.rollback()
        print(f"❌ Erro ao executar limpeza: {e}")
    finally:
        conn.close()

if __name__ == "__main__":
    executar_limpeza_placas_mascaradas()
