"""
Gera a malha municipal própria do mapa territorial a partir das APIs
oficiais do IBGE (dados públicos — citar "Fonte: IBGE"):

  - Malha: API de malhas v3, TopoJSON dos 5.570 municípios (qualidade mínima),
    com o código IBGE de 7 dígitos em cada polígono.
  - Nomes e UF: API de localidades v1.
  - População: Censo 2022 (SIDRA, tabela 4709, variável 93 — população residente).

Saída: componente_mapa/data/brasil.topo.json, no formato que o desenho em
Canvas espera (o mesmo do projeto de referência open-apuracao-brazil):
  - coordenadas projetadas (Albers cônica equivalente, em metros), y para o norte;
  - objeto "municipios", propriedades: id (código IBGE), n (nome), uf, p (população).

Rodar uma vez numa máquina com internet e versionar o arquivo gerado:
    python tools/gerar_malha.py
    python tools/gerar_malha.py --cache pasta_com_downloads   # reaproveita/guarda os downloads
"""

from __future__ import annotations

import argparse
import json
import math
import urllib.request
from datetime import date
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
SAIDA = RAIZ / "componente_mapa" / "data" / "brasil.topo.json"

URL_MALHA = ("https://servicodados.ibge.gov.br/api/v3/malhas/paises/BR"
             "?formato=application/json&qualidade=minima&intrarregiao=municipio")
URL_MUNICIPIOS = "https://servicodados.ibge.gov.br/api/v1/localidades/municipios"
URL_POPULACAO = "https://apisidra.ibge.gov.br/values/t/4709/n6/all/v/93/p/2022"

# Albers cônica equivalente usada pelo IBGE para o Brasil.
R = 6_371_000.0
LAT0, LON0, PAR1, PAR2 = math.radians(-12), math.radians(-54), math.radians(-2), math.radians(-22)
_N = (math.sin(PAR1) + math.sin(PAR2)) / 2
_C = math.cos(PAR1) ** 2 + 2 * _N * math.sin(PAR1)
_RHO0 = R * math.sqrt(_C - 2 * _N * math.sin(LAT0)) / _N

# Precisão da grade da malha de saída (metros). 50 m é muito abaixo do que a
# simplificação "mínima" do IBGE preserva e mantém o arquivo pequeno.
GRADE_METROS = 50.0


def projetar(lon: float, lat: float) -> tuple:
    rho = R * math.sqrt(_C - 2 * _N * math.sin(math.radians(lat))) / _N
    theta = _N * (math.radians(lon) - LON0)
    return rho * math.sin(theta), _RHO0 - rho * math.cos(theta)


def baixar_json(url: str, cache: Path | None, nome: str):
    if cache and (cache / nome).exists():
        return json.loads((cache / nome).read_text(encoding="utf-8"))
    req = urllib.request.Request(url, headers={"User-Agent": "LCFO-gerar-malha"})
    with urllib.request.urlopen(req, timeout=180) as r:
        texto = r.read().decode("utf-8")
    if cache:
        cache.mkdir(parents=True, exist_ok=True)
        (cache / nome).write_text(texto, encoding="utf-8")
    return json.loads(texto)


def uf_do_municipio(m: dict) -> str:
    # Alguns municípios recentes vêm sem microrregião; a região imediata sempre traz a UF.
    try:
        return m["microrregiao"]["mesorregiao"]["UF"]["sigla"]
    except (TypeError, KeyError):
        return m["regiao-imediata"]["regiao-intermediaria"]["UF"]["sigla"]


def converter(malha: dict, municipios: list, populacao: list) -> dict:
    sc, tr = malha["transform"]["scale"], malha["transform"]["translate"]

    # 1) Arcos: decodifica (delta + quantização) -> lon/lat -> Albers (m).
    arcos_m = []
    for arco in malha["arcs"]:
        x = y = 0
        pts = []
        for dx, dy in arco:
            x += dx
            y += dy
            pts.append(projetar(x * sc[0] + tr[0], y * sc[1] + tr[1]))
        arcos_m.append(pts)

    min_x = min(p[0] for a in arcos_m for p in a)
    min_y = min(p[1] for a in arcos_m for p in a)

    # 2) Requantiza numa grade de GRADE_METROS e reaplica o delta.
    arcos_q = []
    for pts in arcos_m:
        saida, px, py = [], 0, 0
        for i, (xm, ym) in enumerate(pts):
            qx, qy = round((xm - min_x) / GRADE_METROS), round((ym - min_y) / GRADE_METROS)
            if i and qx == px and qy == py and i != len(pts) - 1:
                continue
            saida.append([qx - px, qy - py] if saida else [qx, qy])
            px, py = qx, qy
        if len(saida) == 1:                     # arco degenerado: mantém 2 pontos
            saida.append([0, 0])
        arcos_q.append(saida)

    # 3) Propriedades: código, nome, UF, população.
    info = {str(m["id"]): (m["nome"], uf_do_municipio(m)) for m in municipios}
    pop = {}
    for linha in populacao[1:]:
        try:
            pop[str(linha["D1C"])] = int(linha["V"])
        except (ValueError, KeyError):
            pass

    geometrias, sem_nome = [], []
    objeto = next(iter(malha["objects"].values()))
    for g in objeto["geometries"]:
        cod = str(g["properties"]["codarea"])
        nome, uf = info.get(cod, (None, None))
        if nome is None:
            sem_nome.append(cod)
            continue
        geometrias.append({"type": g["type"], "arcs": g["arcs"],
                           "properties": {"id": cod, "n": nome, "uf": uf, "p": pop.get(cod, 0)}})

    return {
        "type": "Topology",
        "transform": {"scale": [GRADE_METROS, GRADE_METROS], "translate": [min_x, min_y]},
        "objects": {"municipios": {"type": "GeometryCollection", "geometries": geometrias}},
        "arcs": arcos_q,
        "metadata": {
            "fonte": "IBGE — API de malhas v3 (qualidade mínima), localidades v1, SIDRA 4709 (Censo 2022)",
            "projecao": "Albers cônica equivalente (lat0 -12, lon0 -54, paralelos -2/-22), metros",
            "populacao_ano": 2022,
            "gerado_em": date.today().isoformat(),
            "municipios": len(geometrias),
            "sem_nome": sem_nome,
        },
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--cache", type=Path, default=None, help="pasta para guardar/reaproveitar os downloads")
    ap.add_argument("--saida", type=Path, default=SAIDA)
    a = ap.parse_args()

    print("Baixando malha, municípios e população do IBGE...")
    malha = baixar_json(URL_MALHA, a.cache, "malha_municipios_minima.json")
    municipios = baixar_json(URL_MUNICIPIOS, a.cache, "localidades_municipios.json")
    populacao = baixar_json(URL_POPULACAO, a.cache, "sidra_4709_censo2022.json")

    topo = converter(malha, municipios, populacao)
    a.saida.parent.mkdir(parents=True, exist_ok=True)
    a.saida.write_text(json.dumps(topo, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    md = topo["metadata"]
    print(f"OK: {md['municipios']} municípios -> {a.saida} ({a.saida.stat().st_size / 1e6:.2f} MB)")
    if md["sem_nome"]:
        print(f"Atenção: {len(md['sem_nome'])} polígono(s) sem nome no cadastro de localidades: {md['sem_nome']}")


if __name__ == "__main__":
    main()
