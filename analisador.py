import os
import sys
import exifread
import folium
import webbrowser
import zipfile
import base64
from lxml import etree
from datetime import datetime
from PyPDF2 import PdfReader
from docx import Document
from pptx import Presentation
from io import BytesIO

# --- FUNÇÕES DE APOIO ---

def extrair_thumbnail_base64(tags):
    try:
        jpeg_thumbnail_data = tags.get('JPEGThumbnail')
        if not jpeg_thumbnail_data:
            return None
        base64_data = base64.b64encode(jpeg_thumbnail_data).decode('ascii')
        return f"data:image/jpeg;base64,{base64_data}"
    except:
        return None

def extrair_tempo_edicao_xml(caminho):
    try:
        with zipfile.ZipFile(caminho, 'r') as z:
            with z.open('docProps/app.xml') as f:
                tree = etree.parse(f)
                ns = {'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'}
                total_time = tree.xpath('//ep:TotalTime', namespaces=ns)
                if total_time:
                    minutos = int(total_time[0].text)
                    horas = minutos // 60
                    restante = minutos % 60
                    return f"{horas}h {restante}min"
    except:
        return "N/A"
    return "N/A"

def extrair_dados(caminho):
    dados = {
        "tags": {}, "lat": None, "lon": None, "tipo": "Desconhecido",
        "data_captura": "N/A", "data_edicao": "N/A",
        "data_modificacao": "N/A", "data_criacao": "N/A",
        "tem_gps": False, "thumbnail_b64": None
    }
    
    stats = os.stat(caminho)
    dados["data_criacao"] = datetime.fromtimestamp(stats.st_ctime).strftime('%d/%m/%Y %H:%M:%S')
    dados["data_modificacao"] = datetime.fromtimestamp(stats.st_mtime).strftime('%d/%m/%Y %H:%M:%S')

    ext = caminho.lower()

    if ext.endswith(('.jpg', '.jpeg', '.png')):
        dados["tipo"] = "Imagem"
        with open(caminho, 'rb') as f:
            tags = exifread.process_file(f)
            for tag, val in tags.items(): dados["tags"][tag] = str(val)
            dados["thumbnail_b64"] = extrair_thumbnail_base64(tags)
            
            dt_orig = tags.get('EXIF DateTimeOriginal')
            if dt_orig:
                try:
                    dados["data_captura"] = datetime.strptime(str(dt_orig), '%Y:%m:%d %H:%M:%S').strftime('%d/%m/%Y %H:%M:%S')
                except: dados["data_captura"] = str(dt_orig)

            dt_img = tags.get('Image DateTime')
            if dt_img:
                try:
                    dados["data_edicao"] = datetime.strptime(str(dt_img), '%Y:%m:%d %H:%M:%S').strftime('%d/%m/%Y %H:%M:%S')
                except: dados["data_edicao"] = str(dt_img)

            def conv(v):
                try:
                    d = float(v.values[0].num) / float(v.values[0].den)
                    m = float(v.values[1].num) / float(v.values[1].den)
                    s = float(v.values[2].num) / float(v.values[2].den)
                    return d + (m/60.0) + (s/3600.0)
                except: return None
            
            lat, lon = conv(tags.get('GPS GPSLatitude')), conv(tags.get('GPS GPSLongitude'))
            if lat and lon:
                if str(tags.get('GPS GPSLatitudeRef')) != 'N': lat = -lat
                if str(tags.get('GPS GPSLongitudeRef')) != 'E': lon = -lon
                dados["lat"], dados["lon"] = lat, lon
                dados["tem_gps"] = True

    elif ext.endswith('.pdf'):
        dados["tipo"] = "PDF"
        reader = PdfReader(caminho)
        for key, val in reader.metadata.items():
            dados["tags"][key.replace('/', '')] = str(val)
        dados["data_captura"] = dados["tags"].get('CreationDate', 'N/A')
        dados["data_edicao"] = dados["tags"].get('ModDate', 'N/A')

    elif ext.endswith(('.docx', '.pptx')):
        if ext.endswith('.docx'):
            dados["tipo"] = "Word (DOCX)"
            prop = Document(caminho).core_properties
        else:
            dados["tipo"] = "PowerPoint (PPTX)"
            prop = Presentation(caminho).core_properties
        dados["data_captura"] = prop.created.strftime('%d/%m/%Y %H:%M:%S') if prop.created else "N/A"
        dados["data_edicao"] = prop.modified.strftime('%d/%m/%Y %H:%M:%S') if prop.modified else "N/A"
        dados["tags"] = {"Autor Original": prop.author or "N/A", "Ultima Modificacao por": prop.last_modified_by or "N/A", "Tempo Total de Edicao": extrair_tempo_edicao_xml(caminho)}

    return dados

def gerar_relatorio_html(dados, caminho_original):
    mapa_coluna = ""
    col_resumo = "col-lg-12"
    if dados["tem_gps"]:
        m = folium.Map(location=[dados["lat"], dados["lon"]], zoom_start=16)
        folium.Marker([dados["lat"], dados["lon"]]).add_to(m)
        mapa_coluna = f"""
            <div class="col-lg-7">
                <div class="card p-0" style="height: 600px; position: relative; border: 1px solid #dee2e6;">
                    <div style="height: 100%; width: 100%;">{m._repr_html_()}</div>
                    <a href="https://www.google.com/maps?q={dados['lat']},{dados['lon']}" target="_blank" class="btn-maps-floating">📍 Abrir no Google Maps</a>
                </div>
            </div>"""
        col_resumo = "col-lg-5"

    img_header = f'<img src="{dados["thumbnail_b64"]}" style="height: 80px; width: 80px; object-fit: contain; border-radius: 4px; border: 1px solid #ddd; background: white; margin-right: 20px;">' if dados["thumbnail_b64"] else ""

    html_content = f"""
    <html>
    <head>
        <meta charset="utf-8">
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>
            body {{ background-color: #f4f6f9; color: #333; font-family: 'Segoe UI', sans-serif; }}
            .card {{ border-radius: 8px; border: 1px solid #dee2e6; background: #fff; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }}
            .header-info {{ background: #ffffff; border-bottom: 3px solid #0056b3; padding: 25px; display: flex; align-items: center; }}
            .timeline-box {{ border-left: 4px solid #0056b3; padding: 12px; background: #f8f9fa; margin-bottom: 12px; }}
            .timeline-alert {{ border-left: 4px solid #e03131; padding: 12px; background: #fff5f5; margin-bottom: 12px; }}
            .scroll-table {{ max-height: 400px; overflow-y: auto; font-size: 0.85rem; }}
            .btn-maps-floating {{ position: absolute; bottom: 20px; right: 20px; background: white; padding: 10px 15px; border: 1px solid #0056b3; border-radius: 8px; text-decoration: none; color: #0056b3; font-weight: bold; z-index: 1000; box-shadow: 0 4px 12px rgba(0,0,0,0.15); }}
            .btn-maps-floating:hover {{ background: #0056b3; color: white; }}
        </style>
    </head>
    <body>
        <div class="header-info mb-4">
            {img_header}
            <div>
                <h2 class="mb-1 text-primary" style="font-weight: 600;">Informações Técnicas do Arquivo</h2>
                <p class="text-muted mb-0">Arquivo: {os.path.basename(caminho_original)}</p>
            </div>
        </div>
        <div class="container-fluid px-4">
            <div class="row">
                <div class="{col_resumo}">
                    <div class="card p-4">
                        <h5 class="mb-4 text-secondary" style="border-bottom: 2px solid #eee; padding-bottom: 10px;">📅 Timeline do Arquivo</h5>
                        <div class="timeline-box"><small class="text-muted d-block">1. Criação Interna (Captura/Origem):</small><strong>{dados['data_captura']}</strong></div>
                        <div class="timeline-alert"><small class="text-danger d-block">2. Edição/Processamento (Software):</small><strong>{dados['data_edicao']}</strong></div>
                        <div class="timeline-box"><small class="text-muted d-block">3. Última Modificação (Sistema):</small><strong>{dados['data_modificacao']}</strong></div>
                        <div class="timeline-box"><small class="text-muted d-block">4. Entrada na Máquina (Criação):</small><strong>{dados['data_criacao']}</strong></div>
                    </div>
                    <div class="card p-4">
                        <h5 class="mb-3 text-secondary" style="border-bottom: 2px solid #eee; padding-bottom: 10px;">🛠️ Informações de Origem</h5>
                        <p class="mb-1"><strong>Tipo de Arquivo:</strong> {dados['tipo']}</p>
                        <p class="mb-1"><strong>Equipamento/Autor:</strong> {dados['tags'].get('Autor Original', dados['tags'].get('Image Model', 'N/A'))}</p>
                        <p class="mb-1"><strong>Software Gerador:</strong> {dados['tags'].get('Image Software', dados['tags'].get('Producer', 'N/A'))}</p>
                    </div>
                </div>
                {mapa_coluna}
                <div class="col-12 mt-2">
                    <div class="card p-4">
                        <h5 class="text-secondary mb-3">📋 Metadados Técnicos Detalhados</h5>
                        <div class="scroll-table">
                            <table class="table table-sm table-hover">
                                <thead class="table-light"><tr><th>Tag</th><th>Valor</th></tr></thead>
                                <tbody>
                                    {"".join([f"<tr><td>{k}</td><td>{v}</td></tr>" for k, v in dados['tags'].items()])}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    nome_relatorio = f"analise_{os.path.splitext(os.path.basename(caminho_original))[0]}.html"
    with open(nome_relatorio, "w", encoding="utf-8") as f: f.write(html_content)
    webbrowser.open(f"file://{os.path.abspath(nome_relatorio)}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        caminho = " ".join(sys.argv[1:]).replace('"', '').strip()
        if os.path.exists(caminho):
            res = extrair_dados(caminho)
            gerar_relatorio_html(res, caminho)