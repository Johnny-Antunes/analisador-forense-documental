import whisper
import hashlib
import datetime
import sys
import os
from fpdf import FPDF

def calcular_hash(caminho_arquivo):
    """Calcula o Hash SHA-256 para garantir a integridade forense do áudio."""
    sha256_hash = hashlib.sha256()
    try:
        with open(caminho_arquivo, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    except Exception as e:
        return f"Erro ao calcular Hash: {e}"

def gerar_relatorio_pdf(caminho_audio, resultado, hash_digital, nome_saida):
    """Gera um PDF formatado com os metadados e a transcrição."""
    pdf = FPDF()
    pdf.add_page()
    
    # Cabeçalho
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(190, 10, "Relatorio de Transcricao de Audio - Analise de Fraude", ln=True, align='C')
    pdf.ln(10)
    
    # Metadados do Arquivo
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(190, 7, "METADADOS DE INTEGRIDADE:", ln=True)
    pdf.set_font("Arial", size=9)
    pdf.cell(190, 7, f"Arquivo: {os.path.basename(caminho_audio)}", ln=True)
    pdf.cell(190, 7, f"Data da Analise: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", ln=True)
    pdf.multi_cell(190, 7, f"Hash SHA-256 (Identidade Digital): {hash_digital}")
    pdf.ln(5)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(10)

    # Transcrição
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(190, 10, "TRANSCRICAO COM TIMESTAMPS:", ln=True)
    pdf.ln(5)
    
    pdf.set_font("Arial", size=10)
    for segment in resultado['segments']:
        # Formata o tempo: [MM:SS]
        inicio = int(segment['start'])
        minutos = inicio // 60
        segundos = inicio % 60
        tempo_formatado = f"[{minutos:02d}:{segundos:02d}]"
        
        texto = segment['text'].strip()
        
        # Escreve o tempo em negrito e o texto normal
        pdf.set_font("Arial", 'B', 10)
        pdf.write(7, f"{tempo_formatado} ")
        pdf.set_font("Arial", size=10)
        pdf.write(7, f"{texto}\n\n")

    pdf.output(nome_saida)

def executar_transcricao():
    # Verifica se o arquivo foi arrastado para o terminal
    if len(sys.argv) < 2:
        print("\n[!] Uso: python transcritor_forense.py [ARRASTE O AUDIO AQUI]")
        return

    caminho_audio = sys.argv[1].strip('"') # Limpa aspas do Windows

    if not os.path.exists(caminho_audio):
        print(f"\n[X] Erro: Arquivo nao encontrado: {caminho_audio}")
        return

    print(f"\n[1/3] Calculando Hash SHA-256...")
    hash_digital = calcular_hash(caminho_audio)

    print(f"[2/3] Carregando IA e Transcrevendo... (Aguarde)")
    # 'small' e otimo para o seu PC (Dell/Lenovo). Se quiser mais velocidade use 'base'
    model = whisper.load_model("small")
    resultado = model.transcribe(caminho_audio, language="pt")

    print(f"[3/3] Gerando Relatorio PDF...")
    nome_pdf = os.path.splitext(caminho_audio)[0] + "_Transcricao.pdf"
    gerar_relatorio_pdf(caminho_audio, resultado, hash_digital, nome_pdf)

    print(f"\n[V] CONCLUIDO!")
    print(f"Relatorio salvo em: {nome_pdf}")

if __name__ == "__main__":
    executar_transcricao()
