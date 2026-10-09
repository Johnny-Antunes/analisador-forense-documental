"""
Diagnóstico SOMENTE-LEITURA de datas (não altera banco nem planilhas).

Uso:
    python diagnostico_datas.py "T:\\pasta\\das\\planilhas" "banco_fraudes.db"
    (o 2º argumento é opcional; sem ele só analisa as planilhas)

Gera diagnostico_datas.txt com contagens, nomes de arquivo e datas.
Não imprime CPF, telefone, placa nem nome.
"""
import re
import sys
import sqlite3
import unicodedata
from collections import Counter
from pathlib import Path

import pandas as pd

MAX_ARQUIVOS = 40  # os mais recentes da pasta
SAIDA = []


def p(*partes):
    linha = " ".join(str(x) for x in partes)
    print(linha)
    SAIDA.append(linha)


def norm(texto):
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "_", texto.lower()).strip("_")


CANDIDATAS = ["data_do_expediente", "data_inclusao_da_assistencia",
              "data_do_servico", "data_abertura", "data"]
RE_DATA = re.compile(r"^(\d{4}-\d{2}-\d{2}|\d{1,2}/\d{1,2}/\d{4})")


def achar_coluna(df):
    colunas = {norm(c): c for c in df.columns}
    for cand in CANDIDATAS:
        for n, original in colunas.items():
            if cand in n:
                return original
    return None


def ler(arq):
    if arq.suffix.lower() == ".xlsx":
        return pd.read_excel(arq, dtype=str, engine="openpyxl")
    for enc in ("utf-8-sig", "latin1"):
        try:
            with open(arq, "r", encoding=enc) as f:
                cab = f.readline()
            sep = max([";", "\t", "|", ","], key=cab.count)
            return pd.read_csv(arq, sep=sep, encoding=enc, dtype=str,
                               keep_default_na=False, engine="python")
        except UnicodeDecodeError:
            continue
    return None


def antigo(chave):
    # lógica antiga: dayfirst=True mesmo em data ISO
    return pd.to_datetime(chave, errors="coerce", dayfirst=True)


def correto(chave):
    if chave[4:5] == "-":  # ISO: ano primeiro, nunca dia primeiro
        return pd.to_datetime(chave, format="%Y-%m-%d", errors="coerce")
    return pd.to_datetime(chave, errors="coerce", dayfirst=True)


def formato(chave):
    return "ISO (aaaa-mm-dd)" if chave[4:5] == "-" else "BR (dd/mm/aaaa)"


def analisar_planilha(arq):
    df = ler(arq)
    if df is None or df.empty:
        p(f"- {arq.name}: não foi possível ler / vazio")
        return None
    col = achar_coluna(df)
    if not col:
        p(f"- {arq.name}: coluna de data não encontrada")
        return None
    chaves, sem_data = Counter(), 0
    for v in df[col].astype(str):
        m = RE_DATA.match(v.strip())
        if m:
            chaves[m.group(1)] += 1
        else:
            sem_data += 1
    total = sum(chaves.values())
    if not total:
        p(f"- {arq.name}: nenhuma data reconhecível ({sem_data} valores)")
        return None
    formatos, divergentes, exemplos = Counter(), 0, []
    meses_antigo, meses_correto = set(), set()
    for chave, qtd in chaves.items():
        formatos[formato(chave)] += qtd
        a, c = antigo(chave), correto(chave)
        if pd.notna(a):
            meses_antigo.add(a.strftime("%Y-%m"))
        if pd.notna(c):
            meses_correto.add(c.strftime("%Y-%m"))
        if pd.isna(a) != pd.isna(c) or (pd.notna(a) and pd.notna(c) and a != c):
            divergentes += qtd
            if len(exemplos) < 2:
                da = a.date() if pd.notna(a) else "NaT"
                dc = c.date() if pd.notna(c) else "NaT"
                exemplos.append(f"{chave} -> antigo {da} | correto {dc}")
    pct = 100 * divergentes / total
    p(f"- {arq.name}: {total} linhas com data | {formatos.most_common(1)[0][0]} | "
      f"divergentes {divergentes} ({pct:.0f}%) | meses antigo={len(meses_antigo)} "
      f"correto={len(meses_correto)}")
    for e in exemplos:
        p(f"    ex.: {e}")
    return divergentes, total


def analisar_pasta(pasta):
    pasta = Path(pasta)
    arquivos = list(pasta.glob("*.xlsx")) + list(pasta.glob("*.csv"))
    arquivos = sorted(arquivos, key=lambda a: a.stat().st_mtime, reverse=True)[:MAX_ARQUIVOS]
    p(f"== PLANILHAS em '{pasta.name}' ({len(arquivos)} arquivos) | pandas {pd.__version__} ==")
    tot_div = tot = 0
    for arq in arquivos:
        r = analisar_planilha(arq)
        if r:
            tot_div += r[0]
            tot += r[1]
    if tot:
        p(f"\nRESUMO: {tot_div} de {tot} linhas ({100 * tot_div / tot:.0f}%) mudariam de data")


def analisar_banco(caminho):
    uri = Path(caminho).resolve().as_uri() + "?mode=ro"
    con = sqlite3.connect(uri, uri=True)
    for tabela in ("assistencias", "criacoes_diarias"):
        p(f"\n== BANCO: {tabela} (15 arquivos com mais meses distintos) ==")
        try:
            linhas = con.execute(f"""
                SELECT arquivo, COUNT(*), COUNT(DISTINCT substr(data,1,7)),
                       MIN(data), MAX(data)
                FROM {tabela}
                WHERE data IS NOT NULL AND data != ''
                GROUP BY arquivo ORDER BY 3 DESC, 2 DESC LIMIT 15""").fetchall()
        except sqlite3.Error as e:
            p("erro:", e)
            continue
        p("arquivo | linhas | meses_distintos | data_min | data_max")
        for arq, n, meses, dmin, dmax in linhas:
            p(f"{arq} | {n} | {meses} | {dmin} | {dmax}")
    con.close()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    analisar_pasta(sys.argv[1])
    if len(sys.argv) > 2:
        analisar_banco(sys.argv[2])
    Path("diagnostico_datas.txt").write_text("\n".join(SAIDA), encoding="utf-8")
    print("\nSalvo em diagnostico_datas.txt")