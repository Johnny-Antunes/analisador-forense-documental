"""
Gera planilhas SINTÉTICAS de Criações Diárias para testar o radar de
anomalias, o funil por município e (futuramente) o mapa territorial.

Os arquivos saem no mesmo formato da exportação real (CSV ";" com as
colunas NRO_ASSISTENCIA, DATA_DO_EXPEDIENTE, TITULAR, ...) e são ingeridos
pela tela normal de ingestão. Todos os nomes começam com TESTE_CRIACOES_,
o que permite removê-los depois com --remover.

Cenários plantados no ÚLTIMO mês (o radar só avalia o mês mais recente):
  - Santo André/SP ........ Explosão de Volume Macro (quadrilha de CPFs/telefones/placas)
  - Santa Quitéria/CE ..... Anomalia em Serviços Pet
  - Juazeiro do Norte/CE .. Salto em Serviços Residenciais
  - Itapuí/SP ............. Pico sem Histórico Prévio (grafado com acento)
  - Caucaia/CE ............ Município em Watchlist (só se cadastrado; ver --cadastrar-watchlist)
E em meses anteriores (para a linha do tempo do mapa):
  - Feira de Santana/BA ... pico 5 meses antes do último
  - Mossoró/RN ............ pico no penúltimo mês (some no último)

Uso (na pasta do projeto):
    python tools/gerar_criacoes_teste.py                       # gera em base_criacao/
    python tools/gerar_criacoes_teste.py --linhas-por-mes 20000
    python tools/gerar_criacoes_teste.py --cadastrar-watchlist # + Caucaia/CE na watchlist
    python tools/gerar_criacoes_teste.py --remover             # apaga os dados de teste do banco

Depois de gerar: ícone de banco (rail esquerdo) > Criações Diárias > Ingerir.
"""

from __future__ import annotations

import argparse
import random
import sqlite3
import sys
from datetime import date
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

from ibge_dados import COORDENADAS_MUNICIPIOS  # noqa: E402

PREFIXO_ARQUIVO = "TESTE_CRIACOES_"
MOTIVO_WATCHLIST_TESTE = "[TESTE] gerado por tools/gerar_criacoes_teste.py"

CAPITAIS = [
    ("RIO BRANCO", "AC"), ("MACEIO", "AL"), ("MACAPA", "AP"), ("MANAUS", "AM"), ("SALVADOR", "BA"),
    ("FORTALEZA", "CE"), ("BRASILIA", "DF"), ("VITORIA", "ES"), ("GOIANIA", "GO"), ("SAO LUIS", "MA"),
    ("CUIABA", "MT"), ("CAMPO GRANDE", "MS"), ("BELO HORIZONTE", "MG"), ("BELEM", "PA"), ("JOAO PESSOA", "PB"),
    ("CURITIBA", "PR"), ("RECIFE", "PE"), ("TERESINA", "PI"), ("RIO DE JANEIRO", "RJ"), ("NATAL", "RN"),
    ("PORTO ALEGRE", "RS"), ("PORTO VELHO", "RO"), ("BOA VISTA", "RR"), ("FLORIANOPOLIS", "SC"),
    ("SAO PAULO", "SP"), ("ARACAJU", "SE"), ("PALMAS", "TO"),
]

# Grafias como viriam da exportação (acentos, caixa variada) — a ingestão e o
# radar precisam juntar tudo pelo nome normalizado.
GRAFIAS = {
    ("SAO PAULO", "SP"): ["São Paulo", "SAO PAULO", "São Paulo"],
    ("SANTO ANDRE", "SP"): ["Santo André"],
    ("SANTA QUITERIA", "CE"): ["Santa Quitéria"],
    ("JUAZEIRO DO NORTE", "CE"): ["Juazeiro do Norte"],
    ("ITAPUI", "SP"): ["Itapuí"],
    ("CAUCAIA", "CE"): ["Caucaia"],
    ("FEIRA DE SANTANA", "BA"): ["Feira de Santana"],
    ("MOSSORO", "RN"): ["Mossoró"],
    ("BELEM", "PA"): ["Belém"],
    ("GOIANIA", "GO"): ["Goiânia"],
    ("MACEIO", "AL"): ["Maceió"],
}

DDD_POR_UF = {
    "AC": 68, "AL": 82, "AP": 96, "AM": 92, "BA": 71, "CE": 85, "DF": 61, "ES": 27, "GO": 62, "MA": 98,
    "MT": 65, "MS": 67, "MG": 31, "PA": 91, "PB": 83, "PR": 41, "PE": 81, "PI": 86, "RJ": 21, "RN": 84,
    "RS": 51, "RO": 69, "RR": 95, "SC": 48, "SP": 11, "SE": 79, "TO": 63,
}

# Atenção às regras do radar (busca por trecho no nome do serviço):
#   pet  = PET | VETERIN      residencial = ELETRIC | ENCANAD | CHAVEIR | DESENTUP | HIDRAUL
#   fora do radar = INFORMATIV | CONSULTA
SERV_AUTO = ["REBOQUE", "TROCA DE PNEU", "CARGA DE BATERIA", "SOCORRO MECANICO", "TAXI", "GUINCHO LEVE", "AUXILIO COMBUSTIVEL"]
SERV_RESID = ["ELETRICISTA", "ENCANADOR", "CHAVEIRO RESIDENCIAL", "DESENTUPIMENTO", "REPARO HIDRAULICO"]
SERV_PET = ["ATENDIMENTO VETERINARIO", "TRANSPORTE PET", "BANHO E TOSA PET"]
SERV_FORA = ["INFORMATIVO", "CONSULTA DE COBERTURA"]

PRENOMES = ["ANA", "BRUNO", "CARLOS", "DANIELA", "EDUARDO", "FERNANDA", "GABRIEL", "HELENA", "IGOR", "JULIANA",
            "LUCAS", "MARIANA", "NATALIA", "OTAVIO", "PAULA", "RAFAEL", "SANDRA", "TIAGO", "VANESSA", "WAGNER"]
SOBRENOMES = ["SILVA", "SANTOS", "OLIVEIRA", "SOUZA", "RODRIGUES", "FERREIRA", "ALVES", "PEREIRA", "LIMA",
              "GOMES", "COSTA", "RIBEIRO", "MARTINS", "CARVALHO", "ARAUJO", "MELO", "BARBOSA", "ROCHA"]
BAIRROS = ["CENTRO", "JARDIM AMERICA", "VILA NOVA", "SAO JOSE", "BOA VISTA", "PLANALTO", "INDUSTRIAL",
           "SANTA CRUZ", "PARQUE DAS FLORES", "NOVA ESPERANCA", "ALTO DA BOA VISTA", "VILA RICA"]
CLIENTES = ["SEGURADORA FICTICIA ALFA", "ASSISTENCIA FICTICIA BETA", "PROTECAO VEICULAR FICTICIA GAMA"]

COLUNAS = ["NRO_ASSISTENCIA", "DATA_DO_EXPEDIENTE", "TITULAR", "CPF_CNPJ_USUARIO", "TELEFONE_TITULAR", "PLACA",
           "SERVICO", "BAIRRO_OCORRENCIA", "CIDADE_OCORRENCIA", "ESTADO_OCORRENCIA", "NOME_DO_CLIENTE"]


# ------------------------------------------------------------------ geradores de entidades
def cpf(rng):
    while True:
        v = "".join(str(rng.randint(0, 9)) for _ in range(11))
        if len(set(v)) > 1:
            return v


def telefone(rng, uf):
    return f"{DDD_POR_UF[uf]}9{rng.randint(10_000_000, 99_999_999)}"


def placa(rng):
    L = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    meio = rng.choice(L) if rng.random() < 0.6 else str(rng.randint(0, 9))  # Mercosul ou antiga
    return f"{''.join(rng.choice(L) for _ in range(3))}{rng.randint(0, 9)}{meio}{rng.randint(10, 99)}"


def nome(rng):
    return f"{rng.choice(PRENOMES)} {rng.choice(SOBRENOMES)} {rng.choice(SOBRENOMES)}"


def grafia(rng, cid, uf):
    opcoes = GRAFIAS.get((cid, uf))
    return rng.choice(opcoes) if opcoes else cid


def servico_normal(rng):
    r = rng.random()
    if r < 0.82:
        return rng.choice(SERV_AUTO)
    if r < 0.92:
        return rng.choice(SERV_RESID)
    if r < 0.95:
        return rng.choice(SERV_PET)
    return rng.choice(SERV_FORA)


class Quadrilha:
    """Conjunto pequeno de CPFs/telefones/placas que se repetem — vira célula no grafo."""

    def __init__(self, rng, uf, n_cpfs=3, n_tels=2, n_placas=3, extras_tel=(), extras_cpf=()):
        self.titulares = [(cpf(rng), nome(rng)) for _ in range(n_cpfs)] + [(c, "TITULAR DA BLACKLIST") for c in extras_cpf]
        self.tels = [telefone(rng, uf) for _ in range(n_tels)] + list(extras_tel)
        self.placas = [placa(rng) for _ in range(n_placas)]

    def entidades(self, rng):
        c, n = rng.choice(self.titulares)
        return c, n, rng.choice(self.tels), rng.choice(self.placas)


# ------------------------------------------------------------------ montagem
def meses_ate(ate: str, n: int):
    ano, mes = map(int, ate.split("-"))
    saida = []
    for _ in range(n):
        saida.append((ano, mes))
        mes -= 1
        if mes == 0:
            ano, mes = ano - 1, 12
    return list(reversed(saida))


def dias_no_mes(ano, mes):
    prox = date(ano + (mes == 12), mes % 12 + 1, 1)
    return (prox - date(ano, mes, 1)).days


def entidades_da_blacklist(caminho_banco: Path, limite=3):
    """Telefones/CPFs mais frequentes da Blacklist (leitura apenas) — cruzam com as criações."""
    if not caminho_banco.exists():
        return [], []
    try:
        conn = sqlite3.connect(f"file:{caminho_banco}?mode=ro", uri=True)
        tels = [r[0] for r in conn.execute(
            "SELECT telefone FROM assistencias WHERE telefone != '' GROUP BY telefone ORDER BY COUNT(*) DESC LIMIT ?", (limite,))]
        cpfs = [r[0] for r in conn.execute(
            "SELECT cpf FROM assistencias WHERE cpf != '' GROUP BY cpf ORDER BY COUNT(*) DESC LIMIT ?", (limite,))]
        conn.close()
        return tels, cpfs
    except sqlite3.Error:
        return [], []


def gerar(saida: Path, ate: str, n_meses: int, linhas_por_mes: int, semente: int, caminho_banco: Path):
    rng = random.Random(semente)
    for chave in CAPITAIS + list(GRAFIAS):
        assert chave in COORDENADAS_MUNICIPIOS, f"município ausente do IBGE: {chave}"

    meses = meses_ate(ate, n_meses)
    ultimo = len(meses) - 1

    plantados = {("SANTO ANDRE", "SP"), ("SANTA QUITERIA", "CE"), ("JUAZEIRO DO NORTE", "CE"), ("ITAPUI", "SP"),
                 ("CAUCAIA", "CE"), ("FEIRA DE SANTANA", "BA"), ("MOSSORO", "RN")}
    outros = [k for k in COORDENADAS_MUNICIPIOS if k not in plantados and k not in CAPITAIS]
    interior = rng.sample(outros, min(700, len(outros)))

    # Peso de cada município na linha de base (capitais concentram ~45%).
    pesos = {}
    for k in CAPITAIS:
        pesos[k] = 0.45 / len(CAPITAIS) * rng.uniform(0.4, 2.2)
    brutos = {k: rng.lognormvariate(0, 1.0) for k in interior}
    soma = sum(brutos.values())
    for k, b in brutos.items():
        pesos[k] = 0.55 * b / soma

    tels_bl, cpfs_bl = entidades_da_blacklist(caminho_banco)
    q_santo_andre = Quadrilha(rng, "SP", extras_tel=tels_bl[:1], extras_cpf=cpfs_bl[:1])
    q_pet = Quadrilha(rng, "CE", n_cpfs=2, n_tels=2, n_placas=1)
    q_resid = Quadrilha(rng, "CE", n_cpfs=3, n_tels=3, n_placas=1)
    q_itapui = Quadrilha(rng, "SP", n_cpfs=4, n_tels=2, n_placas=4)
    q_feira = Quadrilha(rng, "BA")
    q_mossoro = Quadrilha(rng, "RN")

    saida.mkdir(parents=True, exist_ok=True)
    resumo = []
    seq = 0
    for i, (ano, mes) in enumerate(meses):
        linhas = []
        nd = dias_no_mes(ano, mes)

        def add(cid, uf, serv, ent=None, bairro=None):
            nonlocal seq
            seq += 1
            if ent:
                c, n, t, p = ent
            else:
                c, n, t, p = cpf(rng), nome(rng), telefone(rng, uf), placa(rng)
                if rng.random() < 0.15:
                    p = ""              # nem toda assistência tem veículo
            linhas.append([
                f"TST-{ano}{mes:02d}-{seq:07d}", f"{rng.randint(1, nd):02d}/{mes:02d}/{ano} {rng.randint(0, 23):02d}:{rng.randint(0, 59):02d}",
                n, c, t, p, serv, bairro or rng.choice(BAIRROS), grafia(rng, cid, uf), uf, rng.choice(CLIENTES),
            ])

        # --- linha de base nacional
        fator_mes = 1 + 0.08 * ((mes % 12) in (0, 1, 7))   # leve sazonalidade (jan/dez/jul)
        for (cid, uf), w in pesos.items():
            esperado = w * linhas_por_mes * fator_mes
            n = int(esperado) + (rng.random() < esperado - int(esperado))
            for _ in range(n):
                add(cid, uf, servico_normal(rng))

        # --- cenários plantados (linha de base própria + picos)
        for _ in range(rng.randint(7, 10)):
            add("SANTO ANDRE", "SP", servico_normal(rng))
        for _ in range(rng.randint(4, 6)):
            add("SANTA QUITERIA", "CE", rng.choice(SERV_AUTO))
        for _ in range(rng.randint(17, 22)):
            add("JUAZEIRO DO NORTE", "CE", rng.choice(SERV_AUTO))
        for _ in range(2):
            add("JUAZEIRO DO NORTE", "CE", rng.choice(SERV_RESID))
        for _ in range(rng.randint(9, 12)):
            add("CAUCAIA", "CE", servico_normal(rng))
        for _ in range(rng.randint(10, 14)):
            add("FEIRA DE SANTANA", "BA", servico_normal(rng))
        for _ in range(rng.randint(6, 9)):
            add("MOSSORO", "RN", servico_normal(rng))

        if i == ultimo:
            for _ in range(48):
                add("SANTO ANDRE", "SP", rng.choice(["REBOQUE", "SOCORRO MECANICO", "GUINCHO LEVE"]),
                    ent=q_santo_andre.entidades(rng), bairro="VILA ASSUNCAO")
            for _ in range(28):
                add("SANTA QUITERIA", "CE", rng.choice(SERV_PET), ent=q_pet.entidades(rng), bairro="CENTRO")
            for _ in range(30):
                add("JUAZEIRO DO NORTE", "CE", rng.choice(SERV_RESID), ent=q_resid.entidades(rng), bairro="PIRAJA")
            for _ in range(40):
                add("ITAPUI", "SP", rng.choice(SERV_AUTO), ent=q_itapui.entidades(rng))
            # sujeira proposital: deve ser descartada ou filtrada
            linhas.append(["", f"05/{mes:02d}/{ano}", "SEM NUMERO", cpf(rng), telefone(rng, "SP"), placa(rng), "REBOQUE", "CENTRO", "Campinas", "SP", CLIENTES[0]])
            linhas.append([f"TST-SUJO-{ano}{mes:02d}-1", f"06/{mes:02d}/{ano}", "SEM CONTATO", "00000000000", "", "XX", "REBOQUE", "CENTRO", "Campinas", "SP", CLIENTES[0]])
            for _ in range(3):
                add("CIDADE NAO LOCALIZADA", "SP", "REBOQUE")
            add("SANTO ANDRE.", "SP", "REBOQUE")            # grafia que não casa com o IBGE (lista de auditoria do mapa)
        if i == ultimo - 5:
            for _ in range(70):
                add("FEIRA DE SANTANA", "BA", rng.choice(SERV_AUTO), ent=q_feira.entidades(rng), bairro="CAMPO LIMPO")
        if i == ultimo - 1:
            for _ in range(45):
                add("MOSSORO", "RN", rng.choice(SERV_AUTO), ent=q_mossoro.entidades(rng), bairro="ALTO DE SAO MANOEL")

        rng.shuffle(linhas)
        arq = saida / f"{PREFIXO_ARQUIVO}{ano}-{mes:02d}.csv"
        with open(arq, "w", encoding="utf-8-sig", newline="") as f:
            f.write(";".join(COLUNAS) + "\n")
            for ln in linhas:
                f.write(";".join(str(x).replace(";", ",") for x in ln) + "\n")
        resumo.append((arq.name, len(linhas)))

    return resumo, meses, (tels_bl[:1], cpfs_bl[:1])


# ------------------------------------------------------------------ banco (opcional)
def cadastrar_watchlist():
    import database
    ok, msg = database.cadastrar_cidade_risco("CAUCAIA", "CE", MOTIVO_WATCHLIST_TESTE, "teste")
    print(f"Watchlist: {msg}")


def remover_do_banco():
    import database
    conn = database.get_db_connection()
    try:
        cur = conn.cursor()
        cur.execute("DELETE FROM criacoes_diarias WHERE arquivo LIKE ?", (PREFIXO_ARQUIVO + "%",))
        n = cur.rowcount
        cur.execute("DELETE FROM arquivos_processados WHERE tipo_base = 'CRIACAO' AND nome_arquivo LIKE ?", (PREFIXO_ARQUIVO + "%",))
        cur.execute("DELETE FROM cidades_risco WHERE motivo = ?", (MOTIVO_WATCHLIST_TESTE,))
        database.invalidar_agregados_criacoes(cur)
        conn.commit()
    finally:
        conn.close()
    print(f"Removidas {n:,} linhas de teste de criacoes_diarias (e o registro dos arquivos {PREFIXO_ARQUIVO}*).")
    print("Reinicie o app para limpar os caches em memória.")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--saida", type=Path, default=RAIZ / "base_criacao", help="pasta de destino (padrão: base_criacao/)")
    ap.add_argument("--ate", default="2026-09", help="último mês gerado, AAAA-MM (padrão: 2026-09)")
    ap.add_argument("--meses", type=int, default=13, help="quantidade de meses (padrão: 13)")
    ap.add_argument("--linhas-por-mes", type=int, default=6000, help="volume aproximado da linha de base por mês")
    ap.add_argument("--semente", type=int, default=42, help="semente aleatória (mesma semente = mesmos arquivos)")
    ap.add_argument("--cadastrar-watchlist", action="store_true", help="cadastra Caucaia/CE na watchlist (marcado como teste)")
    ap.add_argument("--remover", action="store_true", help="remove do banco tudo o que este script gerou e sai")
    a = ap.parse_args()

    if a.remover:
        remover_do_banco()
        return

    resumo, meses, (tel_bl, cpf_bl) = gerar(a.saida, a.ate, a.meses, a.linhas_por_mes, a.semente, RAIZ / "banco_fraudes.db")
    total = sum(n for _, n in resumo)
    print(f"Gerados {len(resumo)} arquivos ({total:,} linhas) em {a.saida}")
    for nome_arq, n in resumo:
        print(f"  {nome_arq}: {n:,} linhas")
    print(f"\nÚltimo mês (avaliado pelo radar): {meses[-1][0]}-{meses[-1][1]:02d}")
    print("Esperado no Centro de Comando:")
    print("  SANTO ANDRE/SP ........ Explosão de Volume Macro (quadrilha no bairro VILA ASSUNCAO)")
    print("  SANTA QUITERIA/CE ..... Anomalia em Serviços Pet (+ Explosão)")
    print("  JUAZEIRO DO NORTE/CE .. Salto em Serviços Residenciais")
    print("  ITAPUI/SP ............. Pico sem Histórico Prévio")
    print("  CAUCAIA/CE ............ Município em Watchlist" + ("" if a.cadastrar_watchlist else " (só com --cadastrar-watchlist)"))
    if tel_bl or cpf_bl:
        print(f"  Quadrilha de Santo André reutiliza entidades da Blacklist (tel {tel_bl}, cpf {cpf_bl})"
              " -> aparece em 'Expansão com Criações' do caso correspondente.")
    if a.cadastrar_watchlist:
        cadastrar_watchlist()
    print("\nPróximo passo: ícone de banco (rail esquerdo) > Criações Diárias > Ingerir Criações Diárias.")


if __name__ == "__main__":
    main()
