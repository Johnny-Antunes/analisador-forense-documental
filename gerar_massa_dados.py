import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta
from pathlib import Path

PASTA_DESTINO = Path(__file__).parent / "casos_duplicidade"
PASTA_DESTINO.mkdir(exist_ok=True)

print("⚡ Gerando super base forense assimétrica (Redes Complexas & Pontes)...")

# Coordenadas geográficas reais para alimentar o PyDeck
LOCAIS = {
    "SANTOS": {"cidade": "Santos", "uf": "SP", "lat": -23.9608, "lon": -46.3336},
    "PRAIA_GRANDE": {"cidade": "Praia Grande", "uf": "SP", "lat": -24.0058, "lon": -46.4028},
    "SAO_PAULO": {"cidade": "São Paulo", "uf": "SP", "lat": -23.5505, "lon": -46.6333},
    "SJC": {"cidade": "São José dos Campos", "uf": "SP", "lat": -23.1791, "lon": -45.8872},
    "CAMPINAS": {"cidade": "Campinas", "uf": "SP", "lat": -22.9056, "lon": -47.0608},
    "SOROCABA": {"cidade": "Sorocaba", "uf": "SP", "lat": -23.5015, "lon": -47.4526}
}

NOMES = ["Silva", "Santos", "Oliveira", "Souza", "Rodrigues", "Ferreira", "Alves", "Pereira", "Lima", "Gomes", "Costa", "Martins"]
PRENOMES = ["Carlos", "Marcos", "Lucas", "Mariana", "Ana", "Juliana", "Roberto", "Clodoaldo", "Larissa", "Eduardo", "Fernanda", "Bruno", "Ricardo", "Vanessa"]

def gerar_cpf():
    return f"{random.randint(100, 999)}{random.randint(100, 999)}{random.randint(100, 999)}{random.randint(10, 99)}"

def gerar_placa():
    letras = "".join(random.choices("ABCDEFGHIJKLMNOPQRSTUVWXYZ", k=3))
    nums = f"{random.randint(0,9)}{random.choice('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789')}{random.randint(10,99)}"
    return f"{letras}{nums}"

registros = []
sinistro_seq = 100000

# ==============================================================================
# CÉLULA #1: OPERAÇÃO HIDRA (REDE MULTICLUSTER COM PONTES E MULTIHUB)
# ==============================================================================

# Núcleo A: Baixada Santista
tels_nucleo_a = ["13991112233", "13997778899"]
cpfs_nucleo_a = [gerar_cpf() for _ in range(16)]
placas_nucleo_a = [gerar_placa() for _ in range(8)]

# Núcleo B: Vale do Paraíba (SJC)
tels_nucleo_b = ["12981223344", "12988334455"]
cpfs_nucleo_b = [gerar_cpf() for _ in range(18)]
placas_nucleo_b = [gerar_placa() for _ in range(9)]

# A Ponte (Corretor / Despachante de São Paulo Capital que opera ambos)
tel_ponte = "11970001122"
cpfs_ponte = [gerar_cpf(), gerar_cpf()]
placas_migratorias = [gerar_placa(), gerar_placa()] # Carros usados em ambos os núcleos

# 1. Gerando acionamentos do Núcleo A (Janeiro e Fevereiro - Litoral)
for _ in range(110):
    sinistro_seq += 1
    dt = datetime(2026, random.choice([1, 2]), random.randint(1, 28), random.randint(7, 22), random.randint(0, 59))
    cid = random.choice([LOCAIS["SANTOS"], LOCAIS["PRAIA_GRANDE"]])
    
    # Simula forte recorrência em certas duplas (linhas grossas)
    cpf_escolhido = cpfs_nucleo_a[0] if random.random() < 0.25 else random.choice(cpfs_nucleo_a)
    placa_escolhida = placas_nucleo_a[0] if random.random() < 0.30 else random.choice(placas_nucleo_a)

    registros.append({
        "DATA_ABERTURA": dt.strftime("%Y-%m-%d %H:%M:%S"),
        "ID_ASSISTENCIA": f"ME-{sinistro_seq}",
        "TITULAR": f"{random.choice(PRENOMES)} {random.choice(NOMES)} (Litoral)",
        "CPF_CNPJ_USUARIO": cpf_escolhido,
        "TELEFONE_TITULAR": random.choice(tels_nucleo_a),
        "PLACA": placa_escolhida,
        "CIDADE": cid["cidade"],
        "UF": cid["uf"],
        "LAT": cid["lat"] + random.uniform(-0.015, 0.015),
        "LON": cid["lon"] + random.uniform(-0.015, 0.015)
    })

# 2. Gerando acionamentos do Núcleo B (Abril e Maio - Vale do Paraíba)
for _ in range(130):
    sinistro_seq += 1
    dt = datetime(2026, random.choice([4, 5]), random.randint(1, 28), random.randint(6, 23), random.randint(0, 59))
    cid = LOCAIS["SJC"]

    cpf_escolhido = cpfs_nucleo_b[0] if random.random() < 0.20 else random.choice(cpfs_nucleo_b)
    placa_escolhida = placas_nucleo_b[0] if random.random() < 0.25 else random.choice(placas_nucleo_b)

    registros.append({
        "DATA_ABERTURA": dt.strftime("%Y-%m-%d %H:%M:%S"),
        "ID_ASSISTENCIA": f"ME-{sinistro_seq}",
        "TITULAR": f"{random.choice(PRENOMES)} {random.choice(NOMES)} (Vale)",
        "CPF_CNPJ_USUARIO": cpf_escolhido,
        "TELEFONE_TITULAR": random.choice(tels_nucleo_b),
        "PLACA": placa_escolhida,
        "CIDADE": cid["cidade"],
        "UF": cid["uf"],
        "LAT": cid["lat"] + random.uniform(-0.02, 0.02),
        "LON": cid["lon"] + random.uniform(-0.02, 0.02)
    })

# 3. Gerando os Acionamentos da PONTE (Explosão em Março - Conecta A com B)
# O telefone do corretor e as placas migratórias transitam entre Santos, Capital e SJC
for _ in range(65):
    sinistro_seq += 1
    # Explosão em Março (Burst Fraud)
    dt = datetime(2026, 3, random.randint(1, 31), random.randint(8, 20), random.randint(0, 59))
    
    # Alterna entre conectar com elementos de A e de B
    if random.random() < 0.5:
        cpf_usado = random.choice(cpfs_nucleo_a)
        cid = LOCAIS["SAO_PAULO"]
    else:
        cpf_usado = random.choice(cpfs_nucleo_b)
        cid = LOCAIS["SJC"]

    registros.append({
        "DATA_ABERTURA": dt.strftime("%Y-%m-%d %H:%M:%S"),
        "ID_ASSISTENCIA": f"ME-{sinistro_seq}",
        "TITULAR": f"{random.choice(PRENOMES)} {random.choice(NOMES)} (Ponte)",
        "CPF_CNPJ_USUARIO": cpf_usado,
        "TELEFONE_TITULAR": tel_ponte, # Conector mestre
        "PLACA": random.choice(placas_migratorias),
        "CIDADE": cid["cidade"],
        "UF": cid["uf"],
        "LAT": cid["lat"] + random.uniform(-0.025, 0.025),
        "LON": cid["lon"] + random.uniform(-0.025, 0.025)
    })

# ==============================================================================
# CÉLULA #2: GOLPE FAMILIAR TRIANGULAR (PEQUENA, DENSA E FECHADA)
# ==============================================================================
tel_familia = "19982229900"
cpfs_familia = [gerar_cpf() for _ in range(5)]
placas_familia = [gerar_placa(), gerar_placa()]
cid_fam = LOCAIS["CAMPINAS"]

for _ in range(35):
    sinistro_seq += 1
    dt = datetime(2026, 2, random.randint(1, 28))
    registros.append({
        "DATA_ABERTURA": dt.strftime("%Y-%m-%d %H:%M:%S"),
        "ID_ASSISTENCIA": f"ME-{sinistro_seq}",
        "TITULAR": f"Família {random.choice(NOMES)}",
        "CPF_CNPJ_USUARIO": random.choice(cpfs_familia),
        "TELEFONE_TITULAR": tel_familia,
        "PLACA": random.choice(placas_familia),
        "CIDADE": cid_fam["cidade"],
        "UF": cid_fam["uf"],
        "LAT": cid_fam["lat"] + random.uniform(-0.01, 0.01),
        "LON": cid_fam["lon"] + random.uniform(-0.01, 0.01)
    })

# ==============================================================================
# 14.500 SINISTROS LEGÍTIMOS (RUÍDO REAL DE MERCADO)
# ==============================================================================
cidades_lista = list(LOCAIS.values())
for _ in range(14500):
    sinistro_seq += 1
    cid = random.choice(cidades_lista)
    dt = datetime(2026, 1, 1) + timedelta(days=random.randint(0, 170), hours=random.randint(0, 23))
    
    tel = f"119{random.randint(1000000, 9999999)}"
    if random.random() < 0.03:
        tel = "1,19E+10" # Ruído de notação científica comum em planilhas

    registros.append({
        "DATA_ABERTURA": dt.strftime("%Y-%m-%d %H:%M:%S"),
        "ID_ASSISTENCIA": f"ME-{sinistro_seq}",
        "TITULAR": f"{random.choice(PRENOMES)} {random.choice(NOMES)}",
        "CPF_CNPJ_USUARIO": gerar_cpf(),
        "TELEFONE_TITULAR": tel,
        "PLACA": gerar_placa(),
        "CIDADE": cid["cidade"],
        "UF": cid["uf"],
        "LAT": cid["lat"] + random.uniform(-0.03, 0.03),
        "LON": cid["lon"] + random.uniform(-0.03, 0.03)
    })

df = pd.DataFrame(registros)
caminho_saida = PASTA_DESTINO / "super_massa_forense.xlsx"
df.to_excel(caminho_saida, index=False)

print(f"✅ Base gerada com sucesso: {caminho_saida.name} ({len(df)} linhas).")
