"""
Extrator de metadados forense (imagem / PDF / DOCX / PPTX) -> relatório HTML autocontido.

Uso:
    python extrator_metadados.py arquivo1 [arquivo2 ...] [--offline] [--json] [--nao-abrir]

    --offline    não envia coordenadas a serviços externos (sem endereço e sem mapa embutido)
    --json       grava também um .json com os dados extraídos (para integrar com outras ferramentas)
    --nao-abrir  não abre o navegador ao final

Dependências: exifread lxml pypdf python-docx python-pptx requests pillow pypdfium2 (opcional: truststore)
"""
import os
import sys
import re
import io
import json
import html
import base64
import hashlib
import zipfile
import argparse
import webbrowser
import warnings
from datetime import datetime

import exifread
from lxml import etree
from docx import Document
from pptx import Presentation

try:
    from pypdf import PdfReader          # sucessor mantido do PyPDF2
except ImportError:
    from PyPDF2 import PdfReader

try:
    import requests
except ImportError:
    requests = None

try:
    from PIL import Image
except ImportError:
    Image = None

# Em rede corporativa com inspeção SSL, o Python não confia no certificado do proxy.
# O truststore faz o Python usar o repositório de certificados do Windows (o mesmo do navegador).
try:
    import truststore
    truststore.inject_into_ssl()
except Exception:
    pass

warnings.filterwarnings("ignore", category=DeprecationWarning)

FMT = "%d/%m/%Y %H:%M:%S"
EXT_IMAGEM = (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".heic", ".webp")
SOFTWARES_EDICAO = ("photoshop", "gimp", "lightroom", "snapseed", "picsart", "canva", "pixlr",
                    "facetune", "remini", "photopea", "paint.net", "affinity", "fotor", "polarr")
PRODUTORES_SUSPEITOS = ("ilovepdf", "smallpdf", "sejda", "pdfescape", "pdf candy", "pdf24",
                        "foxit phantom", "pdffiller", "dochub", "sodapdf", "pdfelement", "canva")


def esc(v):
    """Escapa qualquer valor antes de ir para o HTML (metadados podem conter <script>)."""
    return html.escape(str(v), quote=True)


# --- APOIO ----------------------------------------------------------------------------

def hashes(caminho):
    sha256, md5 = hashlib.sha256(), hashlib.md5()
    with open(caminho, "rb") as f:
        for bloco in iter(lambda: f.read(1 << 20), b""):
            sha256.update(bloco)
            md5.update(bloco)
    return sha256.hexdigest(), md5.hexdigest()


def datas_sistema(caminho):
    st = os.stat(caminho)
    nascimento = getattr(st, "st_birthtime", None)
    if nascimento is None and os.name == "nt":
        nascimento = st.st_ctime            # no Windows, ctime = criação
    return {
        "criacao": datetime.fromtimestamp(nascimento).strftime(FMT) if nascimento else "N/A (indisponível neste SO)",
        "modificacao": datetime.fromtimestamp(st.st_mtime).strftime(FMT),
    }


def data_exif(valor):
    if not valor:
        return "N/A"
    try:
        return datetime.strptime(str(valor).strip(), "%Y:%m:%d %H:%M:%S").strftime(FMT)
    except ValueError:
        return str(valor)


def data_pdf(valor):
    """D:20240115123000-03'00' -> 15/01/2024 12:30:00 (-03'00')"""
    if not valor or valor == "N/A":
        return "N/A"
    m = re.match(r"D?:?(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?(.*)", str(valor))
    if not m:
        return str(valor)
    a, me, d, h, mi, s, tz = m.groups()
    try:
        dt = datetime(int(a), int(me or 1), int(d or 1), int(h or 0), int(mi or 0), int(s or 0))
        tz = tz.strip().replace("'", "")
        return dt.strftime(FMT) + (f" (UTC{tz[:3]}:{tz[3:5]})" if tz and tz[0] in "+-" else " (Z/UTC)" if tz.startswith("Z") else "")
    except ValueError:
        return str(valor)


def racional(r):
    if hasattr(r, "num"):
        return float(r.num) / float(r.den) if r.den else 0.0
    return float(r)


def gps_decimal(tag, ref):
    try:
        g, m, s = (racional(x) for x in tag.values[:3])
        dec = g + m / 60 + s / 3600
        return -dec if str(ref).strip().upper() in ("S", "W") else dec
    except Exception:
        return None


def preview_b64(caminho, lado=360):
    if Image is None:
        return None
    try:
        with Image.open(caminho) as im:
            im.thumbnail((lado, lado))
            buf = io.BytesIO()
            im.convert("RGB").save(buf, "JPEG", quality=80)
            return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()
    except Exception:
        return None


def _pil_b64(im, lado=220):
    im.thumbnail((lado, lado))
    buf = io.BytesIO()
    im.convert("RGB").save(buf, "JPEG", quality=85)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()


def miniatura_identificacao(caminho, d):
    """Miniatura do CONTEÚDO ATUAL do arquivo, para o cabeçalho (identifica o documento no print).
    Não usa a miniatura EXIF: ela pode mostrar uma versão anterior à edição."""
    ext = caminho.lower()
    try:
        if ext.endswith(EXT_IMAGEM):
            return d.get("preview"), "Imagem analisada"
        if ext.endswith(".pdf"):
            try:
                import pypdfium2 as pdfium
                pg = pdfium.PdfDocument(caminho)[0]
                return _pil_b64(pg.render(scale=0.6).to_pil()), "PDF · página 1"
            except ImportError:
                import fitz  # PyMuPDF, alternativa
                pix = fitz.open(caminho)[0].get_pixmap(dpi=50)
                return _pil_b64(Image.open(io.BytesIO(pix.tobytes("png")))), "PDF · página 1"
        if ext.endswith((".docx", ".pptx")) and Image:
            with zipfile.ZipFile(caminho) as z:
                nomes = [n for n in z.namelist() if n.lower().startswith("docprops/thumbnail")]
                if nomes:
                    return _pil_b64(Image.open(io.BytesIO(z.read(nomes[0])))), "Miniatura do Office (pode estar desatualizada)"
    except Exception:
        pass
    return None, None


def endereco_reverso(lat, lon):
    """Geocodificação reversa via Nominatim (OpenStreetMap). Limite: 1 requisição/s."""
    if requests is None:
        return {"erro": "biblioteca requests não instalada"}
    try:
        r = requests.get(
            "https://nominatim.openstreetmap.org/reverse",
            params={"lat": lat, "lon": lon, "format": "jsonv2", "addressdetails": 1,
                    "zoom": 18, "accept-language": "pt-BR"},
            headers={"User-Agent": "AnaliseMetadadosForense/2.0 (uso interno)"},
            timeout=10,
        )
        r.raise_for_status()
        j = r.json()
        a = j.get("address", {})
        uf = a.get("ISO3166-2-lvl4", "").replace("BR-", "") or a.get("state", "")
        rua = ", ".join(x for x in (a.get("road"), a.get("house_number")) if x)
        bairro = a.get("suburb") or a.get("neighbourhood") or a.get("quarter") or ""
        cidade = a.get("city") or a.get("town") or a.get("village") or a.get("municipality") or ""
        linha = " - ".join(x for x in (rua, bairro) if x)
        linha += (", " if linha else "") + "/".join(x for x in (cidade, uf) if x)
        if a.get("postcode"):
            linha += f", CEP {a['postcode']}"
        return {"formatado": linha or j.get("display_name", "N/A"),
                "completo": j.get("display_name", "N/A"),
                "tipo_local": j.get("type", "")}
    except Exception as ex:
        return {"erro": f"{type(ex).__name__}: {str(ex)[:180]}"}


# --- EXTRAÇÃO -------------------------------------------------------------------------

def extrair_imagem(caminho, d):
    d["tipo"] = "Imagem"
    with open(caminho, "rb") as f:
        tags = exifread.process_file(f)
    thumb = tags.pop("JPEGThumbnail", None)
    if thumb:
        d["thumbnail_exif"] = "data:image/jpeg;base64," + base64.b64encode(thumb).decode()
    d["tags"] = {k: str(v) for k, v in tags.items() if k != "EXIF MakerNote"}
    d["preview"] = preview_b64(caminho)

    d["data_original"] = data_exif(tags.get("EXIF DateTimeOriginal"))
    d["data_edicao"] = data_exif(tags.get("Image DateTime"))
    fuso = tags.get("EXIF OffsetTimeOriginal")
    if fuso and d["data_original"] != "N/A":
        d["data_original"] += f" (UTC{fuso})"
    d["software"] = str(tags.get("Image Software", "N/A"))
    d["equipamento"] = " ".join(str(tags.get(t, "")) for t in ("Image Make", "Image Model")).strip() or "N/A"

    lat = gps_decimal(tags["GPS GPSLatitude"], tags.get("GPS GPSLatitudeRef", "N")) if "GPS GPSLatitude" in tags else None
    lon = gps_decimal(tags["GPS GPSLongitude"], tags.get("GPS GPSLongitudeRef", "E")) if "GPS GPSLongitude" in tags else None
    if lat is not None and lon is not None and not (lat == 0 and lon == 0):
        d["gps"] = {"lat": round(lat, 7), "lon": round(lon, 7)}
        if "GPS GPSAltitude" in tags:
            d["gps"]["altitude_m"] = round(racional(tags["GPS GPSAltitude"].values[0]), 1)
        if "GPS GPSDate" in tags:
            d["gps"]["data_hora_utc"] = f"{tags['GPS GPSDate']} {tags.get('GPS GPSTimeStamp', '')}"

    # Indicadores
    if not tags:
        d["alertas"].append("Arquivo sem EXIF: comum em imagens reenviadas por WhatsApp/redes sociais, prints "
                            "ou limpeza deliberada. Não comprova a origem nem a data da imagem.")
    if any(s in d["software"].lower() for s in SOFTWARES_EDICAO):
        d["alertas"].append(f"Software de edição registrado: {d['software']}")
    if d["data_original"] != "N/A" and d["data_edicao"] != "N/A" and d["data_original"][:19] != d["data_edicao"][:19]:
        d["alertas"].append("Data de modificação (DateTime) diferente da data de captura (DateTimeOriginal).")
    if thumb and d["preview"]:
        d["alertas_info"].append("Compare a miniatura EXIF com a imagem real: se a imagem foi editada, a miniatura "
                                 "pode ainda mostrar o conteúdo original.")


def extrair_pdf(caminho, d):
    d["tipo"] = "PDF"
    d.update(pdf_versao="N/A", pdf_paginas="N/A", pdf_conteudo="N/A", pdf_revisoes="N/A")
    with open(caminho, "rb") as f:
        bruto = f.read()
    reader = PdfReader(io.BytesIO(bruto))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception:
            d["alertas"].append("PDF criptografado: metadados podem estar incompletos.")

    d["pdf_versao"] = bruto[:8].decode("latin-1", "ignore").strip()
    d["pdf_paginas"] = len(reader.pages)
    try:
        texto = "".join((reader.pages[i].extract_text() or "") for i in range(min(len(reader.pages), 3)))
        d["pdf_conteudo"] = "Texto digital (pesquisável)" if len(texto.strip()) > 20 else "Escaneado / somente imagem"
    except Exception:
        d["pdf_conteudo"] = "Indeterminado"

    meta = reader.metadata or {}
    d["tags"] = {k.lstrip("/"): str(v) for k, v in meta.items()}
    d["data_original"] = data_pdf(d["tags"].get("CreationDate"))
    d["data_edicao"] = data_pdf(d["tags"].get("ModDate"))
    criador, produtor = d["tags"].get("Creator", "N/A"), d["tags"].get("Producer", "N/A")
    d["software"] = f"{criador}  →  {produtor}" if criador != produtor else produtor
    d["equipamento"] = d["tags"].get("Author", "N/A")

    try:
        xmp = reader.xmp_metadata
        if xmp:
            for rotulo, attr in (("XMP Produtor", "pdf_producer"), ("XMP Ferramenta", "xmp_creator_tool"),
                                 ("XMP Criação", "xmp_create_date"), ("XMP Modificação", "xmp_modify_date"),
                                 ("XMP Metadata Date", "xmp_metadata_date"), ("XMP DocumentID", "xmpmm_document_id"),
                                 ("XMP InstanceID", "xmpmm_instance_id")):
                v = getattr(xmp, attr, None)
                if v:
                    d["tags"][rotulo] = str(v)
    except Exception:
        pass

    # Atualizações incrementais: cada salvamento "por cima" acrescenta um %%EOF
    eofs = bruto.count(b"%%EOF")
    linearizado = b"/Linearized" in bruto[:2048]
    revisoes = max(eofs - (2 if linearizado else 1), 0)
    d["pdf_revisoes"] = revisoes
    if revisoes:
        d["alertas"].append(f"{revisoes} atualização(ões) incremental(is) após a geração: o PDF foi alterado/salvo "
                            "novamente (edição, anotação ou assinatura). Vale extrair as versões anteriores.")
    if d["data_original"] != "N/A" and d["data_edicao"] != "N/A" and d["data_original"][:19] != d["data_edicao"][:19]:
        d["alertas"].append("Data de modificação diferente da data de criação.")
    if any(s in f"{criador} {produtor}".lower() for s in PRODUTORES_SUSPEITOS):
        d["alertas"].append(f"Gerado/alterado por editor de PDF online ou de terceiros: {d['software']}")
    if d["pdf_conteudo"].startswith("Escaneado") and "scan" not in f"{criador} {produtor}".lower():
        d["alertas_info"].append("Documento sem texto: confirme se o produtor condiz com um scanner/app de digitalização.")


def extrair_office(caminho, d):
    docx = caminho.lower().endswith(".docx")
    d["tipo"] = "Word (DOCX)" if docx else "PowerPoint (PPTX)"
    p = (Document(caminho) if docx else Presentation(caminho)).core_properties
    f = lambda dt: dt.strftime(FMT) + " (UTC)" if dt else "N/A"
    d["data_original"], d["data_edicao"] = f(p.created), f(p.modified)
    d["tags"] = {"Autor original": p.author or "N/A", "Última modificação por": p.last_modified_by or "N/A",
                 "Revisão": p.revision, "Última impressão": f(p.last_printed), "Título": p.title or "N/A"}
    try:
        with zipfile.ZipFile(caminho) as z, z.open("docProps/app.xml") as x:
            raiz = etree.parse(x).getroot()
        for el in raiz:
            nome = etree.QName(el).localname
            if nome in ("Application", "AppVersion", "Company", "Template", "TotalTime") and el.text:
                d["tags"][nome] = el.text
        if "TotalTime" in d["tags"]:
            m = int(d["tags"]["TotalTime"])
            d["tags"]["TotalTime"] = f"{m // 60}h {m % 60}min"
    except Exception:
        pass
    d["software"] = f"{d['tags'].get('Application', 'N/A')} {d['tags'].get('AppVersion', '')}".strip()
    d["equipamento"] = d["tags"]["Autor original"]
    if p.author and p.last_modified_by and p.author != p.last_modified_by:
        d["alertas"].append(f"Autor ({p.author}) diferente de quem modificou por último ({p.last_modified_by}).")


def extrair_dados(caminho, offline=False):
    sha256, md5 = hashes(caminho)
    d = {"arquivo": os.path.basename(caminho), "caminho": os.path.abspath(caminho),
         "tamanho_bytes": os.path.getsize(caminho), "sha256": sha256, "md5": md5,
         "analisado_em": datetime.now().strftime(FMT), "sistema": datas_sistema(caminho),
         "tipo": "Desconhecido", "data_original": "N/A", "data_edicao": "N/A",
         "software": "N/A", "equipamento": "N/A", "tags": {}, "gps": None, "endereco": None,
         "alertas": [], "alertas_info": []}
    ext = caminho.lower()
    try:
        if ext.endswith(EXT_IMAGEM):
            extrair_imagem(caminho, d)
        elif ext.endswith(".pdf"):
            extrair_pdf(caminho, d)
        elif ext.endswith((".docx", ".pptx")):
            extrair_office(caminho, d)
    except Exception as ex:
        d["alertas"].append(f"Falha na extração ({type(ex).__name__}): {ex}")

    d["miniatura"], d["miniatura_legenda"] = miniatura_identificacao(caminho, d)
    if d["gps"] and not offline:
        d["endereco"] = endereco_reverso(d["gps"]["lat"], d["gps"]["lon"])
    return d


# --- RELATÓRIO ------------------------------------------------------------------------

CSS = """
body{margin:0;background:#f4f6f9;color:#222;font:14px 'Segoe UI',Arial,sans-serif}
header{background:#fff;border-bottom:3px solid #0056b3;padding:14px 24px;display:flex;justify-content:space-between;gap:20px;align-items:center;flex-wrap:wrap}
header h1{margin:0 0 2px;font-size:20px;color:#0056b3} .muted{color:#6c757d} .mono{font-family:Consolas,monospace;font-size:11.5px;word-break:break-all}
main{padding:16px 24px;display:flex;flex-direction:column;gap:16px}
.topo{display:grid;grid-template-columns:5fr 7fr;gap:16px;align-items:stretch} .topo.sem-gps{grid-template-columns:1fr}
.baixo{display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:16px}
.card{background:#fff;border:1px solid #dee2e6;border-radius:8px;padding:16px;box-shadow:0 2px 4px rgba(0,0,0,.05)}
.full{grid-column:1/-1} h2{font-size:15px;margin:0 0 10px;color:#495057;border-bottom:1px solid #eee;padding-bottom:6px}
.tl{border-left:4px solid #0056b3;background:#f8f9fa;padding:6px 10px;margin-bottom:7px}
.tl.red{border-color:#e03131;background:#fff5f5} .tl small{display:block;color:#6c757d}
.tec p{margin:3px 0}
.alerta{background:#fff5f5;border-left:4px solid #e03131;padding:8px 10px;margin-bottom:6px}
.info{background:#e7f1ff;border-left:4px solid #0056b3;padding:8px 10px;margin-bottom:6px}
.ok{background:#ebfbee;border-left:4px solid #2f9e44;padding:8px 10px}
table{width:100%;border-collapse:collapse;font-size:12.5px} td,th{border-bottom:1px solid #eee;padding:4px 6px;text-align:left;vertical-align:top}
th{background:#f1f3f5} .scroll{max-height:420px;overflow:auto}
.mapa{display:flex;flex-direction:column} .mapa iframe{flex:1;width:100%;min-height:380px;border:0;border-radius:6px;margin-top:8px}
.end{font-size:15px;font-weight:600} .links{margin-top:8px}
.btn{display:inline-block;margin:0 6px 0 0;padding:5px 11px;border:1px solid #0056b3;border-radius:6px;color:#0056b3;text-decoration:none;font-size:13px}
.btn:hover{background:#0056b3;color:#fff} .imgs{display:flex;gap:14px;flex-wrap:wrap} .imgs img{max-height:260px;border:1px solid #ddd;border-radius:4px}
.id{display:flex;gap:16px;align-items:center}
.thumb{text-align:center} .thumb img,.thumb .sem{height:110px;max-width:160px;object-fit:contain;border:1px solid #ccc;border-radius:4px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.15)}
.thumb .sem{width:76px;display:flex;align-items:center;justify-content:center;font-weight:700;color:#0056b3;background:#e7f1ff}
.thumb small{display:block;font-size:10.5px;color:#6c757d;margin-top:2px}
@media (max-width:1000px){.topo{grid-template-columns:1fr}}
"""


def bloco_mapa(d, offline):
    g = d["gps"]
    if not g:
        return ""
    lat, lon = g["lat"], g["lon"]
    links = (f'<a class="btn" target="_blank" href="https://www.google.com/maps?q={lat},{lon}">Google Maps</a>'
             f'<a class="btn" target="_blank" href="https://www.google.com/maps/@?api=1&map_action=pano&viewpoint={lat},{lon}">Street View</a>'
             f'<a class="btn" target="_blank" href="https://www.openstreetmap.org/?mlat={lat}&mlon={lon}#map=18/{lat}/{lon}">OpenStreetMap</a>')
    end = d.get("endereco") or {}
    if offline:
        txt_end = '<div class="info">Modo offline: endereço e mapa não consultados (coordenadas não saíram da máquina).</div>'
    elif "erro" in end:
        txt_end = (f'<div class="alerta" title="{esc(end["erro"])}">Endereço não obtido: falha ao consultar o serviço de '
                   f'endereços (passe o mouse para ver o detalhe). Use os links abaixo.</div>')
    else:
        txt_end = (f'<div class="end">📍 {esc(end.get("formatado", "N/A"))}</div>'
                   f'<div class="muted" style="font-size:12px">{esc(end.get("completo", ""))}</div>')
    extras = " · ".join(f"{esc(k)}: {esc(v)}" for k, v in g.items() if k not in ("lat", "lon"))
    mapa = "" if offline else (f'<iframe loading="lazy" referrerpolicy="no-referrer" '
                               f'src="https://maps.google.com/maps?q={lat},{lon}&z=17&output=embed"></iframe>')
    return f"""<section class="card mapa"><h2>🌍 Geolocalização</h2>
        {txt_end}<div class="mono" style="margin-top:4px">Coordenadas: {lat}, {lon}{' · ' + extras if extras else ''}</div>
        {mapa}<div class="links">{links}</div></section>"""


def gerar_relatorio_html(d, offline=False, abrir=True):
    alertas = "".join(f'<div class="alerta">⚠️ {esc(a)}</div>' for a in d["alertas"])
    alertas += "".join(f'<div class="info">ℹ️ {esc(a)}</div>' for a in d["alertas_info"])
    alertas = alertas or '<div class="ok">Nenhum indicador automático encontrado (não substitui a análise do analista).</div>'

    imgs = ""
    if d.get("preview") or d.get("thumbnail_exif"):
        imgs = '<section class="card"><h2>🖼️ Imagem x miniatura EXIF</h2><div class="imgs">'
        if d.get("preview"):
            imgs += f'<div><small class="muted">Imagem atual</small><br><img src="{d["preview"]}"></div>'
        if d.get("thumbnail_exif"):
            imgs += f'<div><small class="muted">Miniatura gravada no EXIF</small><br><img src="{d["thumbnail_exif"]}"></div>'
        imgs += "</div></section>"

    pdf = ""
    if d["tipo"] == "PDF":
        pdf = (f"<p><b>Versão:</b> {esc(d['pdf_versao'])} · <b>Páginas:</b> {d['pdf_paginas']}</p>"
               f"<p><b>Conteúdo:</b> {esc(d['pdf_conteudo'])} · <b>Revisões incrementais:</b> {d['pdf_revisoes']}</p>")

    if d.get("miniatura"):
        thumb = (f'<div class="thumb"><img src="{d["miniatura"]}" alt="miniatura">'
                 f'<small>{esc(d["miniatura_legenda"])}</small></div>')
    else:
        ext = os.path.splitext(d["arquivo"])[1].lstrip(".").upper() or "?"
        thumb = f'<div class="thumb"><div class="sem">{esc(ext)}</div><small>sem pré-visualização</small></div>'

    linhas = "".join(f"<tr><td>{esc(k)}</td><td>{esc(v)}</td></tr>" for k, v in d["tags"].items())

    doc = f"""<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src data:; frame-src https://maps.google.com https://www.google.com">
<title>Metadados - {esc(d['arquivo'])}</title><style>{CSS}</style></head><body>
<header>
  <div class="id">{thumb}
  <div><h1>Informações Técnicas do Arquivo</h1>
  <div><b>Arquivo:</b> {esc(d['arquivo'])} <span class="muted">({d['tamanho_bytes']:,} bytes · {esc(d['tipo'])})</span></div></div></div>
  <div class="mono" style="text-align:right">SHA-256: {d['sha256']}<br>MD5: {d['md5']}<br><span class="muted">Analisado em {d['analisado_em']}</span></div>
</header>
<main>
<div class="topo{'' if d['gps'] else ' sem-gps'}">
  <section class="card"><h2>📅 Cronologia de Metadados</h2>
    <div class="tl"><small>Criação / captura (metadado interno)</small><b>{esc(d['data_original'])}</b></div>
    <div class="tl red"><small>Modificação (software)</small><b>{esc(d['data_edicao'])}</b></div>
    <div class="tl"><small>Modificação no sistema de arquivos</small><b>{esc(d['sistema']['modificacao'])}</b></div>
    <div class="tl"><small>Criação no sistema de arquivos (muda ao copiar/baixar)</small><b>{esc(d['sistema']['criacao'])}</b></div>
    <h2 style="margin-top:14px">🛠️ Informações Técnicas</h2>
    <div class="tec"><p><b>Tipo:</b> {esc(d['tipo'])}</p>
    <p><b>Software:</b> {esc(d['software'])}</p>
    <p><b>Equipamento/Autor:</b> {esc(d['equipamento'])}</p>{pdf}</div>
  </section>
  {bloco_mapa(d, offline)}
</div>
<div class="baixo">
  <section class="card"><h2>🚩 Indicadores</h2>{alertas}</section>
  {imgs}
  <section class="card full"><h2>📋 Metadados detalhados ({len(d['tags'])})</h2>
  <div class="scroll"><table><tr><th>Tag</th><th>Valor</th></tr>{linhas}</table></div></section>
</div>
</main></body></html>"""

    nome = f"relatorio_{d['arquivo'].replace('.', '_')}.html"
    with open(nome, "w", encoding="utf-8") as f:
        f.write(doc)
    print(f"✅ Relatório: {nome}")
    if abrir:
        webbrowser.open("file://" + os.path.abspath(nome))
    return nome


# --- CLI ------------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Extrator de metadados forense")
    ap.add_argument("arquivos", nargs="+")
    ap.add_argument("--offline", action="store_true", help="não consulta endereço nem carrega mapa")
    ap.add_argument("--json", action="store_true", help="salva também os dados em JSON")
    ap.add_argument("--nao-abrir", action="store_true", help="não abre o navegador")
    a = ap.parse_args()

    # Compatível com caminho com espaços passado sem aspas
    caminhos = [p.strip('"') for p in a.arquivos]
    junto = " ".join(caminhos)
    if not all(os.path.exists(p) for p in caminhos) and os.path.exists(junto):
        caminhos = [junto]

    for c in caminhos:
        if not os.path.isfile(c):
            print(f"❌ Arquivo não encontrado: {c}")
            continue
        d = extrair_dados(c, offline=a.offline)
        gerar_relatorio_html(d, offline=a.offline, abrir=not a.nao_abrir)
        if a.json:
            dj = {k: v for k, v in d.items() if k not in ("preview", "thumbnail_exif")}
            nome = f"relatorio_{d['arquivo'].replace('.', '_')}.json"
            with open(nome, "w", encoding="utf-8") as f:
                json.dump(dj, f, ensure_ascii=False, indent=2, default=str)
            print(f"   JSON: {nome}")


if __name__ == "__main__":
    main()
