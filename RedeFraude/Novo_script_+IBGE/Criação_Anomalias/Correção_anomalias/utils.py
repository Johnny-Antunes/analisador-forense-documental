import pandas as pd
import unicodedata
import re

# =====================================================
# BASE ESTÁTICA LOCAL DO IBGE (5.570 MUNICÍPIOS)
# =====================================================
try:
    from ibge_dados import COORDENADAS_MUNICIPIOS
except ImportError:
    COORDENADAS_MUNICIPIOS = {}

# =====================================================
# COORDENADAS OFICIAIS DAS 27 CAPITAIS (FALLBACK)
# =====================================================
CAPITAIS_BRASIL = {
    'AC': (-9.9753, -67.8249), 'AL': (-9.6658, -35.7351), 'AP': (0.0356, -51.0705),
    'AM': (-3.1190, -60.0217), 'BA': (-12.9777, -38.5016), 'CE': (-3.7172, -38.5433),
    'DF': (-15.7975, -47.8919), 'ES': (-20.3155, -40.3128), 'GO': (-16.6869, -49.2648),
    'MA': (-2.5307, -44.3068), 'MT': (-15.6010, -56.0974), 'MS': (-20.4697, -54.6201),
    'MG': (-19.9217, -43.9386), 'PA': (-1.4558, -48.4902), 'PB': (-7.1195, -34.8450),
    'PR': (-25.4284, -49.2733), 'PE': (-8.0578, -34.8829), 'PI': (-5.0920, -42.8038),
    'RJ': (-22.9068, -43.1729), 'RN': (-5.7945, -35.2110), 'RS': (-30.0346, -51.2177),
    'RO': (-8.7619, -63.9039), 'RR': (2.8235, -60.6758),  'SC': (-27.5954, -48.5480),
    'SP': (-23.5505, -46.6333), 'SE': (-10.9472, -37.0731), 'TO': (-10.2491, -48.3243)
}


def corrigir_mojibake(texto):
    """
    Reverte corrupção de codificação dupla (UTF-8 decodificado como Windows-1252).
    Ex: "SÃƒO PAULO" -> "SÃO PAULO".
    Se o texto já estiver correto, o round-trip falha e o original é mantido com segurança.
    """
    if not texto or not isinstance(texto, str):
        return texto
    try:
        reparado = texto.encode('cp1252').decode('utf-8')
        if reparado and reparado != texto and '\ufffd' not in reparado:
            return reparado
    except (UnicodeDecodeError, UnicodeEncodeError):
        pass
    return texto


def normalizar_cidade(texto):
    if not texto or not isinstance(texto, str):
        return ""
    texto = corrigir_mojibake(texto)
    nfkd = unicodedata.normalize('NFKD', str(texto))
    sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    limpo = re.sub(r'[^a-zA-Z0-9\s]', ' ', sem_acento)
    return " ".join(limpo.upper().split())


def obter_coordenadas(cidade, uf):
    cid_norm = normalizar_cidade(cidade)
    uf_norm = (uf or "").strip().upper()
    
    if (cid_norm, uf_norm) in COORDENADAS_MUNICIPIOS:
        return COORDENADAS_MUNICIPIOS[(cid_norm, uf_norm)]
    if uf_norm in CAPITAIS_BRASIL:
        return CAPITAIS_BRASIL[uf_norm]
    return None, None


# =====================================================
# UTILITÁRIOS FORENSES DE LIMPEZA
# =====================================================
def mapear_coluna(df, nomes_possiveis, excluir_se_conter=None):
    colunas_df = {normalizar_cidade(c): c for c in df.columns}
    for nome in nomes_possiveis:
        nome_norm = normalizar_cidade(nome)
        for c_norm, c_orig in colunas_df.items():
            if nome_norm in c_norm:
                if excluir_se_conter:
                    if any(normalizar_cidade(exc) in c_norm for exc in excluir_se_conter):
                        continue
                return c_orig
    return None


def limpar_cpf_cnpj(val):
    if pd.isna(val): return ""
    v = str(val).strip()
    
    if re.search(r"[eE][+-]?\d+", v):
        try:
            v = f"{int(float(v.replace(',', '.')))}"
        except Exception:
            pass
    elif re.match(r"^\d+\.\d+$", v):
        v = v.split('.')[0]
    
    v = re.sub(r'\D', '', v)
    if len(v) in [11, 14] and len(set(v)) > 1:
        return v
    if 0 < len(v) < 11:
        v = v.zfill(11)
        return v if len(set(v)) > 1 else ""
    return ""


def limpar_tel(val):
    if pd.isna(val): return ""
    v = str(val).strip()
    
    if re.search(r"[eE][+-]?\d+", v):
        try:
            v = f"{int(float(v.replace(',', '.')))}"
        except Exception:
            pass
    elif re.match(r"^\d+\.\d+$", v):
        v = v.split('.')[0]
        
    v = re.sub(r'\D', '', v)
    if v.startswith("55") and len(v) in [12, 13]:
        v = v[2:]
    if len(v) in [10, 11]:
        return v
    return ""


def limpar_placa(val):
    if pd.isna(val): return ""
    v = re.sub(r'[^a-zA-Z0-9]', '', str(val)).upper()
    if len(v) != 7:
        return ""
    padrao_brasil = r'^[A-Z]{3}[0-9][0-9A-Z][0-9]{2}$'
    if re.match(padrao_brasil, v):
        return v
    return ""


def normalizar_texto(val):
    if pd.isna(val): return ""
    texto = corrigir_mojibake(str(val).strip())
    return texto.upper()


def extrair_data_arquivo(nome_arquivo):
    m = re.search(r'(\d{2})(\d{2})(\d{4})', str(nome_arquivo))
    if m:
        dia, mes, ano = m.groups()
        return f"{ano}-{mes}-{dia}"
    m_iso = re.search(r'(\d{4})-(\d{2})-(\d{2})', str(nome_arquivo))
    if m_iso:
        return m_iso.group(0)
    return ""


def tratar_data_flexivel(val, nome_arquivo):
    if pd.isna(val): 
        return extrair_data_arquivo(nome_arquivo)
    v = str(val).strip()
    if not v or v.lower() in ["nan", "nat", "none", "null", "-"]:
        return extrair_data_arquivo(nome_arquivo)
    if re.match(r"^00\d{2}", v):
        v = "20" + v[2:]
    try:
        dt = pd.to_datetime(v, errors="coerce", dayfirst=True)
        if pd.notna(dt):
            return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
    return extrair_data_arquivo(nome_arquivo)


def formatar_cpf_cnpj(val):
    v = re.sub(r'\D', '', str(val))
    if len(v) == 11:
        return f"{v[:3]}.{v[3:6]}.{v[6:9]}-{v[9:]}"
    elif len(v) == 14:
        return f"{v[:2]}.{v[2:5]}.{v[5:8]}/{v[8:12]}-{v[12:]}"
    return val


def formatar_tel(val):
    v = re.sub(r'\D', '', str(val))
    if len(v) == 11:
        return f"({v[:2]}) {v[2:7]}-{v[7:]}"
    elif len(v) == 10:
        return f"({v[:2]}) {v[2:6]}-{v[6:]}"
    return val