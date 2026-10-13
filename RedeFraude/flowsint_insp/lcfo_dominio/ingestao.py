"""
Prévia de ingestão, revisão antes de gravar e log reversível (item 10).

Divide a ingestão em duas partes:
  - analisar_arquivo(caminho, tipo_base, mapeamento) -> relatório, PURA (não escreve nada):
    colunas detectadas e mapeamento para os campos canônicos (sugerido, editável, reutilizável
    por layout de planilha), formato de data, amostra, linhas válidas, rejeitos por linha com
    motivo, datas ambíguas (quantas mudariam de mês), duplicatas contra a base e cidades sem
    código IBGE (as que não aparecem no mapa);
  - aplicar(relatorios, decisoes, ...) -> grava: backup antes, confere o hash de cada arquivo
    (aborta se mudou desde a análise), lotes com progresso e cancelamento, rejeitos e histórico
    em tabelas, id_ingestao em cada linha (para reverter exatamente aquele lote) e as colunas
    não mapeadas em dados_extra_json, em vez de descartá-las.

A sanitização (sanear) é a mesma das funções antigas de database.py, que agora a usam também.
"""

from __future__ import annotations

import hashlib
import json
import re
import uuid
from collections import Counter, OrderedDict
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Tuple

import pandas as pd

from utils import (extrair_data_arquivo, limpar_cpf_cnpj, limpar_placa, limpar_tel, mapear_coluna,
                   normalizar_cidade, normalizar_texto, obter_coordenadas, tratar_data_flexivel)

TABELA = {"BLACKLIST": "assistencias", "CRIACAO": "criacoes_diarias"}
LOTE_GRAVACAO = 2000
MAX_REJEITOS_RELATORIO = 2000

# campo canônico -> (rótulo, sinônimos de coluna, excluir se o nome contiver)
CAMPOS: "OrderedDict[str, Tuple[str, List[str], List[str]]]" = OrderedDict([
    ("id_assistencia", ("Nº da assistência", ["nro_assistencia", "numero_da_assistencia", "numero_assistencia", "assistencia"], [])),
    ("telefone", ("Telefone", ["telefone_titular", "telefone_do_titular", "telefone", "celular", "contato"], [])),
    ("cpf", ("CPF/CNPJ", ["cpf_cnpj_usuario", "cpf_cnpj", "cpf", "cnpj", "documento"], ["prestador", "cliente"])),
    ("placa", ("Placa", ["placa", "veiculo"], [])),
    ("titular", ("Titular", ["titular", "nome_titular", "segurado"], ["cliente", "prestador", "garantia"])),
    # "data" excluído: antes "Data do Servico" casava com Serviço e a data ia parar na coluna servico
    ("servico", ("Serviço", ["servico", "descricao_do_servico", "tipo_servico"], ["data"])),
    ("data", ("Data", ["data_do_expediente", "data_inclusao_da_assistencia", "data_do_servico", "data_abertura", "data"], [])),
    ("cidade", ("Cidade", ["cidade_ocorrencia", "cidade_de_ocorrencia", "cidade_origem", "cidade_de_origem", "cidade"], [])),
    ("uf", ("UF", ["estado_ocorrencia", "estado_de_ocorrencia", "estado_origem", "estado_de_origem", "uf", "estado"], [])),
    ("bairro", ("Bairro", ["bairro_ocorrencia", "bairro_de_ocorrencia", "bairro_origem", "bairro_de_origem", "bairro"], [])),
    ("empresa_cliente", ("Seguradora/cliente", ["nome_do_cliente", "nome_cliente", "cliente", "empresa_cliente", "seguradora",
                                                "seguradora_contratante", "companhia_cliente", "cia_cliente"], ["titular", "prestador", "garantia"])),
])
COLUNAS_TABELA = ["arquivo", "id_assistencia", "data", "titular", "cpf", "telefone", "placa", "servico", "bairro", "cidade",
                  "uf", "latitude", "longitude", "empresa_cliente"]


# ------------------------------------------------------------------ leitura e mapeamento
def ler(caminho: Path) -> pd.DataFrame:
    from database import ler_arquivo_seguro
    return ler_arquivo_seguro(Path(caminho))


def hash_arquivo(caminho: Path) -> str:
    h = hashlib.sha1()
    with open(caminho, "rb") as f:
        for bloco in iter(lambda: f.read(1 << 20), b""):
            h.update(bloco)
    return h.hexdigest()


def assinatura_colunas(colunas: Iterable[str]) -> str:
    """Identifica o layout da planilha (as mesmas colunas, em qualquer ordem)."""
    return hashlib.sha1("|".join(sorted(normalizar_cidade(str(c)) for c in colunas)).encode("utf-8")).hexdigest()[:16]


def sugerir_mapeamento(colunas: List[str]) -> Dict[str, Optional[str]]:
    vazio = pd.DataFrame(columns=list(colunas))
    return {campo: mapear_coluna(vazio, sinonimos, excluir_se_conter=excluir or None) for campo, (_, sinonimos, excluir) in CAMPOS.items()}


# ------------------------------------------------------------------ sanitização
def _serie(df: pd.DataFrame, coluna: Optional[str], fn, padrao):
    n = len(df)
    if coluna and coluna in df.columns:
        return df[coluna].map(fn).reset_index(drop=True)
    return pd.Series([padrao] * n)


def sanear(df: pd.DataFrame, arquivo: str, mapa: Dict[str, Optional[str]], com_coordenadas: bool = True) -> pd.DataFrame:
    """Mesma regra da ingestão de sempre. Devolve o df canônico + colunas de controle
    (_linha, _valido, _motivo, _campo, _valor, _data_do_arquivo, _extra)."""
    df = df.reset_index(drop=True)
    n = len(df)
    s_id = df[mapa["id_assistencia"]].astype(str).str.strip().reset_index(drop=True) if mapa.get("id_assistencia") else pd.Series([""] * n)
    out = pd.DataFrame({
        "arquivo": [arquivo] * n,
        "id_assistencia": s_id.replace({"nan": "", "None": ""}),
        "data": _serie(df, mapa.get("data"), lambda v: tratar_data_flexivel(v, arquivo), extrair_data_arquivo(arquivo)),
        "titular": _serie(df, mapa.get("titular"), normalizar_texto, ""),
        "cpf": _serie(df, mapa.get("cpf"), limpar_cpf_cnpj, ""),
        "telefone": _serie(df, mapa.get("telefone"), limpar_tel, ""),
        "placa": _serie(df, mapa.get("placa"), limpar_placa, ""),
        "servico": _serie(df, mapa.get("servico"), normalizar_texto, "ASSISTENCIA"),
        "bairro": _serie(df, mapa.get("bairro"), normalizar_texto, ""),
        "cidade": _serie(df, mapa.get("cidade"), normalizar_texto, "SAO PAULO"),
        "uf": _serie(df, mapa.get("uf"), normalizar_texto, "SP"),
        "empresa_cliente": _serie(df, mapa.get("empresa_cliente"), normalizar_texto, ""),
    })
    if com_coordenadas:
        coords = [obter_coordenadas(c, u) for c, u in zip(out["cidade"], out["uf"])]
        out["latitude"] = [c[0] for c in coords]
        out["longitude"] = [c[1] for c in coords]
    else:
        out["latitude"], out["longitude"] = None, None
    # rejeição explicada por linha (antes as linhas sumiam sem relatório)
    brutos = {c: (df[mapa[c]].astype(str).str.strip().reset_index(drop=True) if mapa.get(c) and mapa[c] in df.columns else pd.Series([""] * n))
              for c in ("cpf", "telefone", "placa", "data")}
    motivos, campos, valores = [], [], []
    for i in range(n):
        if not out.at[i, "id_assistencia"]:
            motivos.append("sem número da assistência"); campos.append("id_assistencia"); valores.append("")
            continue
        if out.at[i, "cpf"] or out.at[i, "telefone"] or out.at[i, "placa"]:
            motivos.append(""); campos.append(""); valores.append("")
            continue
        invalidos = [(c, brutos[c][i]) for c in ("cpf", "telefone", "placa") if brutos[c][i] and brutos[c][i].lower() not in ("nan", "none")]
        if invalidos:
            c, v = invalidos[0]
            motivos.append(f"{CAMPOS[c][0]} inválido" + (f" (e mais {len(invalidos) - 1})" if len(invalidos) > 1 else ""))
            campos.append(c); valores.append(v)
        else:
            motivos.append("sem CPF, telefone ou placa"); campos.append("cpf"); valores.append("")
    out["ctl_linha"] = range(2, n + 2)                    # linha na planilha (1 = cabeçalho)
    out["ctl_motivo"] = motivos
    out["ctl_campo"] = campos
    out["ctl_valor"] = valores
    out["ctl_valido"] = out["ctl_motivo"] == ""
    datas_brutas = brutos["data"]
    out["ctl_data_do_arquivo"] = [(not d or d.lower() in ("nan", "nat", "none", "null", "-")) or
                               (tratar_data_flexivel(d, "") == "") for d in datas_brutas] if mapa.get("data") else [True] * n
    usadas = {v for v in mapa.values() if v}
    extras = [c for c in df.columns if c not in usadas]
    if extras:
        sub = df[extras].astype(str)
        out["ctl_extra"] = [json.dumps({k: v for k, v in zip(extras, linha) if v and v.lower() not in ("nan", "none")}, ensure_ascii=False)
                         for linha in sub.itertuples(index=False, name=None)]
    else:
        out["ctl_extra"] = "{}"
    return out


def detectar_formato_data(valores: Iterable[str]) -> Dict[str, Any]:
    """ISO (ano na frente) é sempre ano-mês-dia; dd/mm/aaaa com dia e mês até 12 é ambíguo."""
    cont = Counter()
    ambiguas = mudariam = 0
    for v in valores:
        s = str(v).strip()
        if re.match(r"^\d{4}-\d{1,2}-\d{1,2}", s):
            cont["iso"] += 1
        elif m := re.match(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})", s):
            a, b = int(m.group(1)), int(m.group(2))
            cont["dia/mês/ano"] += 1
            if a <= 12 and b <= 12 and a != b:
                ambiguas += 1
                mudariam += 1                     # lida ao contrário, a data cai em outro mês
        elif s and s.lower() not in ("nan", "nat", "none", "-"):
            cont["outro"] += 1
    formato = cont.most_common(1)[0][0] if cont else "ausente"
    if len([k for k, n in cont.items() if n]) > 1:
        formato = "misto (" + ", ".join(f"{k}: {n}" for k, n in cont.most_common()) + ")"
    return {"formato": formato, "ambiguas": ambiguas, "mudariam_de_mes": mudariam}


# ------------------------------------------------------------------ análise (pura)
def _ids_existentes(tabela: str, ids: List[str]) -> set:
    import database
    out = set()
    conn = database.get_db_connection()
    try:
        for i in range(0, len(ids), 500):
            parte = ids[i:i + 500]
            out |= {r[0] for r in conn.execute(f"SELECT DISTINCT id_assistencia FROM {tabela} WHERE id_assistencia IN ({','.join('?' * len(parte))})", parte)}
    finally:
        conn.close()
    return out


def _ja_ingerido(nome: str, tipo_base: str) -> bool:
    import database
    conn = database.get_db_connection()
    try:
        return conn.execute("SELECT 1 FROM arquivos_processados WHERE nome_arquivo = ? AND tipo_base = ?", (nome, tipo_base)).fetchone() is not None
    finally:
        conn.close()


def listar_arquivos(pasta: str, tipo_base: str) -> List[Dict[str, Any]]:
    p = Path(str(pasta).strip())
    if not p.exists():
        return []
    out = []
    for a in sorted(list(p.glob("*.xlsx")) + list(p.glob("*.csv")), key=lambda x: x.name):
        st = a.stat()
        out.append({"caminho": str(a), "nome": a.name, "tamanho": st.st_size,
                    "modificado": datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"),
                    "ja_ingerido": _ja_ingerido(a.name, tipo_base)})
    return out


def analisar_arquivo(caminho: str, tipo_base: str, mapeamento: Optional[Dict[str, Optional[str]]] = None,
                     amostra: int = 20) -> Dict[str, Any]:
    """Relatório do arquivo SEM tocar no banco (nenhuma escrita)."""
    import territorio_engine as te
    arq = Path(caminho)
    df = ler(arq)
    colunas = [str(c) for c in df.columns]
    assinatura = assinatura_colunas(colunas)
    espec = especificacao_por_assinatura(assinatura) if mapeamento is None else None
    mapa = mapeamento or (espec["mapeamento"] if espec else sugerir_mapeamento(colunas))
    mapa = {c: (mapa.get(c) if mapa.get(c) in colunas else None) for c in CAMPOS}
    san = sanear(df, arq.name, mapa, com_coordenadas=False)
    validas = san[san["ctl_valido"]]
    rejeitadas = san[~san["ctl_valido"]]
    datas = detectar_formato_data(df[mapa["data"]].astype(str)) if mapa.get("data") else {"formato": "ausente (usa a data do nome do arquivo)", "ambiguas": 0, "mudariam_de_mes": 0}
    ids = sorted({i for i in validas["id_assistencia"] if i})
    existentes = _ids_existentes(TABELA[tipo_base], ids) if ids else set()
    sem_ibge = Counter()
    for (cid, uf), n in validas.groupby(["cidade", "uf"]).size().items():
        if te.resolver_codigo_ibge(cid, uf) is None:
            sem_ibge[(cid, uf)] += int(n)
    rejeitos = [{"linha": int(r.ctl_linha), "campo": r.ctl_campo, "valor": r.ctl_valor, "motivo": r.ctl_motivo}
                for r in rejeitadas[["ctl_linha", "ctl_campo", "ctl_valor", "ctl_motivo"]].itertuples(index=False)]
    return {
        "arquivo": arq.name, "caminho": str(arq), "tipo_base": tipo_base, "hash": hash_arquivo(arq), "tamanho": arq.stat().st_size,
        "ja_ingerido": _ja_ingerido(arq.name, tipo_base), "colunas": colunas, "assinatura": assinatura,
        "especificacao": espec["nome"] if espec else None, "mapeamento": mapa,
        "nao_mapeadas": [c for c in colunas if c not in {v for v in mapa.values() if v}],
        "n_linhas": len(df), "n_validas": len(validas), "n_rejeitadas": len(rejeitadas),
        "rejeitos_por_motivo": dict(Counter(rejeitadas["ctl_motivo"]).most_common()),
        "rejeitos": rejeitos[:MAX_REJEITOS_RELATORIO],
        "datas": datas, "n_data_do_arquivo": int(san["ctl_data_do_arquivo"].sum()) if mapa.get("data") else len(df),
        "duplicatas": int(validas["id_assistencia"].isin(existentes).sum()),
        "sem_ibge": [{"cidade": c, "uf": u, "n": n} for (c, u), n in sem_ibge.most_common(50)],
        "amostra": validas.drop(columns=[c for c in validas.columns if c.startswith("ctl_") or c in ("latitude", "longitude")])
                          .head(amostra).to_dict("records"),
    }


# ------------------------------------------------------------------ especificações de importação
def especificacao_por_assinatura(assinatura: str) -> Optional[Dict[str, Any]]:
    import database
    conn = database.get_db_connection()
    try:
        r = conn.execute("SELECT * FROM ingestao_especificacoes WHERE assinatura_colunas = ? ORDER BY criado_em DESC LIMIT 1",
                         (assinatura,)).fetchone()
        return {"id": r["id"], "nome": r["nome"], "mapeamento": json.loads(r["mapeamento_json"]), "formato_data": r["formato_data"]} if r else None
    finally:
        conn.close()


def salvar_especificacao(nome: str, colunas: List[str], mapeamento: Dict[str, Optional[str]], formato_data: str = "") -> str:
    import database
    eid = f"esp_{uuid.uuid4().hex[:10]}"
    conn = database.get_db_connection()
    try:
        conn.execute("INSERT INTO ingestao_especificacoes (id, nome, assinatura_colunas, mapeamento_json, formato_data, criado_em) VALUES (?,?,?,?,?,?)",
                     (eid, nome.strip() or "Layout", assinatura_colunas(colunas), json.dumps(mapeamento, ensure_ascii=False), formato_data,
                      datetime.now().isoformat(timespec="seconds")))
        conn.commit()
    finally:
        conn.close()
    return eid


# ------------------------------------------------------------------ aplicação (grava)
class Cancelado(Exception):
    pass


class ArquivoMudou(Exception):
    pass


def _aplicar_correcoes(san: pd.DataFrame, corrigidos: Dict[int, Dict[str, str]], descartados: set) -> pd.DataFrame:
    limpeza = {"cpf": limpar_cpf_cnpj, "telefone": limpar_tel, "placa": limpar_placa, "id_assistencia": lambda v: str(v).strip()}
    for linha, campos in corrigidos.items():
        idx = san.index[san["ctl_linha"] == int(linha)]
        for i in idx:
            for c, v in campos.items():
                if c in limpeza:
                    san.at[i, c] = limpeza[c](v)
            ok = bool(san.at[i, "id_assistencia"]) and bool(san.at[i, "cpf"] or san.at[i, "telefone"] or san.at[i, "placa"])
            san.at[i, "ctl_valido"] = ok
            san.at[i, "ctl_motivo"] = "" if ok else (san.at[i, "ctl_motivo"] or "correção ainda inválida")
            san.at[i, "ctl_corrigido"] = True
    if descartados:
        san.loc[san["ctl_linha"].isin(descartados), "ctl_descartado"] = True
    return san


def aplicar(relatorios: List[Dict[str, Any]], decisoes: Optional[Dict[str, Any]] = None, *, usuario: str = "",
            progresso: Optional[Callable[[str, int, int], None]] = None, cancelado: Optional[Callable[[], bool]] = None,
            fazer_backup: bool = True) -> Dict[str, Any]:
    """
    Grava os arquivos analisados. decisoes = {"corrigidos": {arquivo: {linha: {campo: valor}}},
    "descartados": {arquivo: [linhas]}, "ignorar_duplicatas": bool}. Aborta ANTES de escrever se
    algum arquivo mudou desde a análise (hash diferente).
    """
    import database
    from lcfo_dominio import sketch_repo
    decisoes = decisoes or {}
    progresso = progresso or (lambda a, f, t: None)
    cancelado = cancelado or (lambda: False)
    if not relatorios:
        raise ValueError("nenhum arquivo para gravar")
    tipo_base = relatorios[0]["tipo_base"]
    tabela = TABELA[tipo_base]
    for r in relatorios:                                   # a pasta de rede pode ter mudado entre análise e confirmação
        if hash_arquivo(Path(r["caminho"])) != r["hash"]:
            raise ArquivoMudou(f"{r['arquivo']} mudou desde a análise: analise de novo antes de gravar.")
    backup = str(sketch_repo._fazer_backup("ingestao")) if fazer_backup else None
    id_exec = f"ing_{uuid.uuid4().hex[:10]}"
    agora = lambda: datetime.now().isoformat(timespec="seconds")
    conn = database.get_db_connection()
    try:
        conn.execute("""INSERT INTO ingestao_execucoes (id, tipo_base, iniciado_em, usuario, arquivos_json, estado, backup_path)
                        VALUES (?,?,?,?,?,'gravando',?)""",
                     (id_exec, tipo_base, agora(), usuario, json.dumps([r["arquivo"] for r in relatorios], ensure_ascii=False), backup))
        conn.commit()
    finally:
        conn.close()
    tot = Counter()
    estado = "concluida"
    try:
        for r in relatorios:
            df = ler(Path(r["caminho"]))
            san = sanear(df, r["arquivo"], r["mapeamento"], com_coordenadas=True)
            san["ctl_corrigido"], san["ctl_descartado"] = False, False
            corr = {int(k): v for k, v in (decisoes.get("corrigidos", {}).get(r["arquivo"]) or {}).items()}
            desc = {int(x) for x in (decisoes.get("descartados", {}).get(r["arquivo"]) or [])}
            san = _aplicar_correcoes(san, corr, desc)
            gravar = san[san["ctl_valido"] & ~san["ctl_descartado"].astype(bool)]
            if decisoes.get("ignorar_duplicatas"):
                existentes = _ids_existentes(tabela, sorted(set(gravar["id_assistencia"])))
                tot["duplicadas"] += int(gravar["id_assistencia"].isin(existentes).sum())
                gravar = gravar[~gravar["id_assistencia"].isin(existentes)]
            tot["lidas"] += len(san)
            linhas = [tuple(x) + (id_exec, ex) for x, ex in zip(gravar[COLUNAS_TABELA].itertuples(index=False, name=None), gravar["ctl_extra"])]
            inseridas_arq = 0
            for i in range(0, len(linhas), LOTE_GRAVACAO):
                if cancelado():
                    raise Cancelado()
                conn = database.get_db_connection()
                try:
                    antes = conn.total_changes
                    conn.executemany(f"""INSERT OR IGNORE INTO {tabela} ({', '.join(COLUNAS_TABELA)}, id_ingestao, dados_extra_json)
                                         VALUES ({', '.join('?' * (len(COLUNAS_TABELA) + 2))})""", linhas[i:i + LOTE_GRAVACAO])
                    inseridas_arq += conn.total_changes - antes
                    conn.commit()
                finally:
                    conn.close()
                progresso(r["arquivo"], min(i + LOTE_GRAVACAO, len(linhas)), len(linhas))
            tot["inseridas"] += inseridas_arq
            tot["duplicadas"] += len(linhas) - inseridas_arq
            # rejeitos: pendentes (para corrigir depois), corrigidos e descartados ficam registrados
            rej = san[(~san["ctl_valido"]) | san["ctl_descartado"].astype(bool) | san["ctl_corrigido"].astype(bool)]
            conn = database.get_db_connection()
            try:
                conn.executemany("""INSERT INTO ingestao_rejeitos (id_execucao, arquivo, linha, campo, valor, motivo, corrigido_valor, estado)
                                    VALUES (?,?,?,?,?,?,?,?)""",
                                 [(id_exec, r["arquivo"], int(x.ctl_linha), x.ctl_campo, x.ctl_valor, x.ctl_motivo or "corrigido",
                                   json.dumps(corr.get(int(x.ctl_linha)), ensure_ascii=False) if int(x.ctl_linha) in corr else None,
                                   "descartado" if x.ctl_descartado else ("corrigido" if x.ctl_corrigido and x.ctl_valido else "pendente"))
                                  for x in rej.itertuples(index=False)])
                tot["rejeitadas"] += int((~san["ctl_valido"]).sum() + san["ctl_descartado"].astype(bool).sum())
                conn.execute("""INSERT OR REPLACE INTO arquivos_processados (nome_arquivo, tipo_base, data_processamento, registros_inseridos)
                                VALUES (?, ?, datetime('now'), ?)""", (r["arquivo"], tipo_base, inseridas_arq))
                conn.commit()
            finally:
                conn.close()
    except Cancelado:
        estado = "cancelada"
    except Exception as e:
        estado = "falha"
        _fechar_execucao(id_exec, estado, tot, erro=str(e))
        raise
    finally:
        if estado != "falha":
            _fechar_execucao(id_exec, estado, tot)
        _pos_gravacao(tipo_base)
    return {"id": id_exec, "estado": estado, "backup": backup, **dict(tot)}


def _fechar_execucao(id_exec: str, estado: str, tot: Counter, erro: str = "") -> None:
    import database
    conn = database.get_db_connection()
    try:
        conn.execute("""UPDATE ingestao_execucoes SET concluido_em = ?, estado = ?, n_lidas = ?, n_inseridas = ?, n_duplicadas = ?,
                        n_rejeitadas = ? WHERE id = ?""",
                     (datetime.now().isoformat(timespec="seconds"), estado + (f": {erro[:200]}" if erro else ""), tot["lidas"],
                      tot["inseridas"], tot["duplicadas"], tot["rejeitadas"], id_exec))
        conn.commit()
    finally:
        conn.close()


def _pos_gravacao(tipo_base: str) -> None:
    import database
    if tipo_base == "BLACKLIST":
        database.semear_base_mestra_da_blacklist()
    else:
        conn = database.get_db_connection()
        try:
            database.invalidar_agregados_criacoes(conn.cursor())
            conn.commit()
        finally:
            conn.close()


# ------------------------------------------------------------------ histórico e reversão
def listar_execucoes(limite: int = 50) -> List[Dict[str, Any]]:
    import database
    conn = database.get_db_connection()
    try:
        return [dict(r) for r in conn.execute("SELECT * FROM ingestao_execucoes ORDER BY iniciado_em DESC LIMIT ?", (limite,))]
    finally:
        conn.close()


def rejeitos(id_execucao: str, limite: int = 2000) -> List[Dict[str, Any]]:
    import database
    conn = database.get_db_connection()
    try:
        return [dict(r) for r in conn.execute("SELECT * FROM ingestao_rejeitos WHERE id_execucao = ? ORDER BY arquivo, linha LIMIT ?",
                                              (id_execucao, limite))]
    finally:
        conn.close()


def dependencias_de_casos(id_execucao: str) -> List[str]:
    """Casos que dependem de linhas deste lote (promovidas/anexadas a partir das Criações)."""
    import database
    conn = database.get_db_connection()
    try:
        ex = conn.execute("SELECT tipo_base FROM ingestao_execucoes WHERE id = ?", (id_execucao,)).fetchone()
        if not ex or ex["tipo_base"] != "CRIACAO":
            return []
        rows = conn.execute("""SELECT DISTINCT a.arquivo FROM assistencias a
                               JOIN criacoes_diarias c ON c.id_assistencia = a.id_assistencia
                               WHERE c.id_ingestao = ? AND (a.arquivo = 'PROMOVIDO_CRIACAO' OR a.arquivo LIKE 'ANEXADO_%')""",
                            (id_execucao,)).fetchall()
        return sorted(r[0].replace("ANEXADO_", "") if r[0].startswith("ANEXADO_") else "caso promovido de descoberta" for r in rows)
    finally:
        conn.close()


def reverter(id_execucao: str, usuario: str = "", fazer_backup: bool = True) -> Dict[str, Any]:
    """Apaga EXATAMENTE as linhas do lote. Bloqueado se alguma já sustenta um caso."""
    import database
    from lcfo_dominio import sketch_repo
    deps = dependencias_de_casos(id_execucao)
    if deps:
        raise ValueError("Linhas deste lote já foram promovidas/anexadas: " + ", ".join(deps))
    conn = database.get_db_connection()
    try:
        ex = conn.execute("SELECT * FROM ingestao_execucoes WHERE id = ?", (id_execucao,)).fetchone()
    finally:
        conn.close()
    if not ex or str(ex["estado"]).startswith("revertida"):
        raise ValueError("ingestão inexistente ou já revertida")
    backup = str(sketch_repo._fazer_backup("reverter_ingestao")) if fazer_backup else None
    tabela = TABELA[ex["tipo_base"]]
    conn = database.get_db_connection()
    try:
        n = conn.execute(f"DELETE FROM {tabela} WHERE id_ingestao = ?", (id_execucao,)).rowcount
        for a in json.loads(ex["arquivos_json"] or "[]"):
            conn.execute("DELETE FROM arquivos_processados WHERE nome_arquivo = ? AND tipo_base = ?", (a, ex["tipo_base"]))
        conn.execute("UPDATE ingestao_execucoes SET estado = ? WHERE id = ?",
                     (f"revertida em {datetime.now().isoformat(timespec='seconds')} por {usuario or '—'}", id_execucao))
        conn.commit()
    finally:
        conn.close()
    _pos_gravacao(ex["tipo_base"])
    return {"removidas": n, "backup": backup}
