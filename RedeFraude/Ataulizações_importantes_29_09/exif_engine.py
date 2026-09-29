"""
Módulo: exif_engine.py
Objetivo: Extração forense de metadados EXIF (coordenadas GPS e timestamp original)
          de fotografias de reembolso para identificação de fraude territorial.
          Usa exclusivamente a API pública do Pillow (Image.getexif / get_ifd),
          compatível com JPEG e blocos eXIf de PNG moderno.
"""

from typing import Dict, Any, Optional
from PIL import Image
from PIL.ExifTags import TAGS, GPSTAGS
from datetime import datetime
import io

from utils import obter_coordenadas
from correlation_engine import calcular_distancia_km

TAG_GPS_INFO = 0x8825
TAG_EXIF_SUBIFD = 0x8769


def _converter_para_graus_decimais(coordenada_dms, referencia: str) -> Optional[float]:
    """Converte coordenadas em Graus, Minutos e Segundos (DMS) do EXIF para Graus Decimais."""
    try:
        graus = float(coordenada_dms[0])
        minutos = float(coordenada_dms[1]) / 60.0
        segundos = float(coordenada_dms[2]) / 3600.0
        decimais = graus + minutos + segundos
        if referencia in ['S', 'W']:
            decimais = -decimais
        return round(decimais, 6)
    except Exception:
        return None


def extrair_metadados_foto(arquivo_bytes: bytes) -> Dict[str, Any]:
    """
    Inspeciona o fluxo de bytes de uma imagem e recupera metadados forenses de fabricação,
    data de captura da câmera e coordenadas geográficas gravadas pelo GPS do dispositivo.
    """
    resultado: Dict[str, Any] = {
        "possui_exif": False,
        "data_captura": None,
        "fabricante": None,
        "modelo": None,
        "latitude": None,
        "longitude": None,
        "erro": None
    }

    try:
        imagem = Image.open(io.BytesIO(arquivo_bytes))
        exif_obj = imagem.getexif()

        if not exif_obj:
            return resultado

        resultado["possui_exif"] = True

        for tag_id, valor in exif_obj.items():
            nome_tag = TAGS.get(tag_id, tag_id)
            if nome_tag == "DateTime":
                try:
                    resultado["data_captura"] = datetime.strptime(str(valor), "%Y:%m:%d %H:%M:%S").strftime("%d/%m/%Y %H:%M:%S")
                except Exception:
                    resultado["data_captura"] = str(valor)
            elif nome_tag == "Make":
                resultado["fabricante"] = str(valor).strip()
            elif nome_tag == "Model":
                resultado["modelo"] = str(valor).strip()

        if hasattr(exif_obj, "get_ifd"):
            try:
                sub_ifd = exif_obj.get_ifd(TAG_EXIF_SUBIFD)
                for tag_id, valor in (sub_ifd or {}).items():
                    if TAGS.get(tag_id, tag_id) == "DateTimeOriginal":
                        try:
                            resultado["data_captura"] = datetime.strptime(str(valor), "%Y:%m:%d %H:%M:%S").strftime("%d/%m/%Y %H:%M:%S")
                        except Exception:
                            resultado["data_captura"] = str(valor)
                        break
            except Exception:
                pass

            dados_gps: Dict[str, Any] = {}
            try:
                gps_dict = exif_obj.get_ifd(TAG_GPS_INFO)
                for chave_gps, valor_gps in (gps_dict or {}).items():
                    subtag = GPSTAGS.get(chave_gps, chave_gps)
                    dados_gps[subtag] = valor_gps
            except Exception:
                pass

            if "GPSLatitude" in dados_gps and "GPSLatitudeRef" in dados_gps:
                resultado["latitude"] = _converter_para_graus_decimais(dados_gps["GPSLatitude"], dados_gps["GPSLatitudeRef"])
            if "GPSLongitude" in dados_gps and "GPSLongitudeRef" in dados_gps:
                resultado["longitude"] = _converter_para_graus_decimais(dados_gps["GPSLongitude"], dados_gps["GPSLongitudeRef"])

    except Exception as e:
        resultado["erro"] = f"Falha na leitura forense: {str(e)[:80]}"

    return resultado


def validar_coerencia_geografica_foto(
    metadados_foto: Dict[str, Any],
    cidade_declarada: str,
    uf_declarada: str,
    tolerancia_km: float = 60.0
) -> Dict[str, Any]:
    """
    Confronta a localização física real em que a fotografia foi tirada contra
    o município declarado na solicitação da assistência.
    """
    analise = {
        "apto_para_confronto": False,
        "distancia_km": 0.0,
        "divergencia_critica": False,
        "parecer_exif": "Metadados de GPS não presentes na imagem (arquivo sem EXIF ou comprimido por app de mensagens)."
    }

    lat_foto = metadados_foto.get("latitude")
    lon_foto = metadados_foto.get("longitude")

    if lat_foto is None or lon_foto is None:
        return analise

    lat_declarada, lon_declarada = obter_coordenadas(cidade_declarada, uf_declarada)
    if lat_declarada is None or lon_declarada is None:
        analise["parecer_exif"] = f"Município declarado ({cidade_declarada}/{uf_declarada}) não localizado no dicionário IBGE."
        return analise

    dist_km = calcular_distancia_km(lat_foto, lon_foto, lat_declarada, lon_declarada)
    analise["apto_para_confronto"] = True
    analise["distancia_km"] = dist_km

    if dist_km > tolerancia_km:
        analise["divergencia_critica"] = True
        analise["parecer_exif"] = (
            f"DIVERGÊNCIA MATERIAL DE LOCALIZAÇÃO: A fotografia foi registrada a {dist_km:.1f} km "
            f"do município declarado na assistência ({cidade_declarada}/{uf_declarada}). "
            f"Coordenadas EXIF reais: [{lat_foto:.5f}, {lon_foto:.5f}]."
        )
    else:
        analise["divergencia_critica"] = False
        analise["parecer_exif"] = (
            f"Compatibilidade geográfica confirmada: A fotografia foi registrada a {dist_km:.1f} km "
            f"da praça declarada, dentro do raio operacional de atendimento."
        )

    return analise