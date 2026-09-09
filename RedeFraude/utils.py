import re
import pandas as pd

# Dicionário de coordenadas base para geolocalização automática
COORDENADAS_CIDADES = {
    "SAO PAULO": (-23.5505, -46.6333), "CAMPINAS": (-22.9056, -47.0608),
    "SANTOS": (-23.9608, -46.3336), "SAO BERNARDO DO CAMPO": (-23.6944, -46.5654),
    "SANTO ANDRE": (-23.6639, -46.5383), "OSASCO": (-23.5329, -46.7920),
    "SAO JOSE DOS CAMPOS": (-23.1791, -45.8872), "RIBEIRAO PRETO": (-21.1767, -47.8208),
    "SOROCABA": (-23.5015, -47.4526), "GUARULHOS": (-23.4542, -46.5341),
    "PRAIA GRANDE": (-24.0058, -46.4028), "BAURU": (-22.3145, -49.0587),
    "JUNDIAI": (-23.1857, -46.8892), "PIRACICABA": (-22.7253, -47.6492),
    "CARAPICUIBA": (-23.5222, -46.8356), "DIADEMA": (-23.6865, -46.6234),
    "MOGI DAS CRUZES": (-23.5206, -46.1854), "TAUBATE": (-23.0264, -45.5553),
    "FRANCA": (-20.5386, -47.4008), "BARUERI": (-23.5105, -46.8761),
    "RIO DE JANEIRO": (-22.9068, -43.1729), "CURITIBA": (-25.4284, -49.2733),
    "BELO HORIZONTE": (-19.9167, -43.9345), "SALVADOR": (-12.9714, -38.5014),
    "FORTALEZA": (-3.7172, -38.5434), "RECIFE": (-8.0476, -34.8770),
    "PORTO ALEGRE": (-30.0346, -51.2177), "BRASILIA": (-15.7975, -47.8919)
}

def sanitizar_nome_coluna(col):
    if not col: return ""
    c = str(col).strip().lower()
    c = re.sub(r"[àáâãä]", "a", c)
    c = re.sub(r"[èéêë]", "e", c)
    c = re.sub(r"[ìíîï]", "i", c)
    c = re.sub(r"[òóôõö]", "o", c)
    c = re.sub(r"[ùúûü]", "u", c)
    c = re.sub(r"[ç]", "c", c)
    return re.sub(r"[^a-z0-9]", "", c)

def mapear_coluna(df, lista_sinonimos, excluir_se_conter=None):
    if excluir_se_conter is None:
        excluir_se_conter = []
    
    colunas_mapa = {sanitizar_nome_coluna(c): c for c in df.columns}
    
    for sin in lista_sinonimos:
        sin_clean = sanitizar_nome_coluna(sin)
        if sin_clean in colunas_mapa:
            orig = colunas_mapa[sin_clean]
            if not any(sanitizar_nome_coluna(exc) in sin_clean for exc in excluir_se_conter):
                return orig
                
    for sin in lista_sinonimos:
        sin_clean = sanitizar_nome_coluna(sin)
        for col_clean, col_orig in colunas_mapa.items():
            if sin_clean in col_clean:
                if not any(sanitizar_nome_coluna(exc) in col_clean for exc in excluir_se_conter):
                    return col_orig
    return None

def normalizar_texto(val):
    if pd.isna(val): return ""
    v = str(val).strip().upper()
    return "" if v in ["NAN", "NONE", "NULL", "-", "NAT", ""] else v

def limpar_placa(val):
    if pd.isna(val): return ""
    v = str(val).strip().upper()
    placa = re.sub(r"[^A-Z0-9]", "", v)
    return placa if len(placa) == 7 else ""

def limpar_cpf_cnpj(val):
    if pd.isna(val): return ""
    v = f"{int(val)}" if isinstance(val, (float, int)) else str(val).strip()
    if re.search(r"[eE][+-]?\d+", v):
        try: v = f"{int(float(v.replace(',', '.')))}"
        except Exception: pass
    if re.match(r"^\d+\.0+$", v):
        v = v.split(".")[0]
    nums = re.sub(r"\D", "", v)
    if not nums or nums.endswith("0000000"): return ""
    if len(nums) <= 11:
        nums = nums.zfill(11)
        return "" if len(set(nums)) == 1 else nums
    elif len(nums) <= 14:
        return nums.zfill(14)
    return nums

def formatar_cpf_cnpj(digits):
    if len(digits) == 11: return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}"
    elif len(digits) == 14: return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"
    return digits

def limpar_tel(val):
    if pd.isna(val): return ""
    v = f"{int(val)}" if isinstance(val, (float, int)) else str(val).strip()
    if re.search(r"[eE][+-]?\d+", v):
        try: v = f"{int(float(v.replace(',', '.')))}"
        except Exception: pass
    nums = re.sub(r"\D", "", v)
    if not nums or nums.endswith("000000"): return ""
    if len(nums) in [12, 13] and nums.startswith("55"): nums = nums[2:]
    if len(nums) in [11, 12] and nums.startswith("0"): nums = nums[1:]
    return nums if len(nums) in [10, 11] else ""

def formatar_tel(digits):
    if len(digits) == 11: return f"({digits[:2]}) {digits[2:7]}-{digits[7:]}"
    elif len(digits) == 10: return f"({digits[:2]}) {digits[2:6]}-{digits[6:]}"
    return digits

def extrair_data_arquivo(nome_arquivo):
    match = re.search(r"(\d{2})(\d{2})(\d{4})", nome_arquivo)
    if match:
        dia, mes, ano = match.groups()
        return f"{ano}-{mes}-{dia}"
    return ""

def tratar_data_flexivel(val, nome_arquivo):
    if pd.isna(val): return extrair_data_arquivo(nome_arquivo)
    v = str(val).strip()
    if not v or v.lower() in ["nan", "nat", "none", "null", "-"]:
        return extrair_data_arquivo(nome_arquivo)
    if re.match(r"^00\d{2}", v):
        v = "20" + v[2:]
    try:
        dt = pd.to_datetime(v, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
    return extrair_data_arquivo(nome_arquivo)

def obter_coordenadas(cidade, uf):
    cid_clean = re.sub(r"[^A-Z ]", "", normalizar_texto(cidade))
    if cid_clean in COORDENADAS_CIDADES:
        return COORDENADAS_CIDADES[cid_clean]
    return (-23.5505, -46.6333)