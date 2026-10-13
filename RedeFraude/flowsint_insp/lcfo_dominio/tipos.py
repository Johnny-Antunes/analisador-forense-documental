"""
Registro único dos tipos de entidade (prefixo, cor, campos, normalização, validação,
detecção e rótulo).

Antes essas regras estavam espalhadas: CORES/PREFIXO/COLUNA em enrich_engine.py, uma
segunda paleta em graph_engine.py, ícones num if/elif, filtros fixos no index.html,
_tipo_de_no na sidebar e LABELS_FILTRO_TIPO em config.py. Tipo novo = uma chamada a
registrar() aqui.

Identidade: o id do nó continua PREFIXO_valorCanônico (CPF_, TEL_, PLACA_...). O valor
canônico sai de `normalizar`, que para CPF, telefone e placa é EXATAMENTE a limpeza da
ingestão (utils.limpar_*): mudar isso quebraria ids já gravados em sketches e pareceres
(tests/test_tipos.py tem o teste "golden" contra a base).
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from typing import Callable, Dict, List, NamedTuple, Optional, Tuple

from utils import formatar_cpf_cnpj, formatar_tel, limpar_cpf_cnpj, limpar_placa, limpar_tel

from lcfo_dominio.icones import svg_no


# ------------------------------------------------------------------ estruturas
@dataclass(frozen=True)
class Campo:
    nome: str
    rotulo: str
    obrigatorio: bool = False
    tipo: str = "texto"            # texto | numero | data | uf
    primario: bool = False         # o campo que é a identidade (somente leitura depois de criado)


@dataclass(frozen=True)
class TipoEntidade:
    chave: str                      # "cpf", "telefone", "placa", "endereco", "clinica"...
    prefixo: str                    # "CPF", "TEL", "PLACA", "END", "CLI"
    rotulo: str                     # "Titular (CPF)"
    cor: Tuple[str, str]            # (fundo, borda)
    icone: str                      # chave em icones.PICTOGRAMAS
    forma: str = "image"            # forma no vis-network
    campos: Tuple[Campo, ...] = ()
    normalizar: Callable[[str], str] = str.strip       # valor canônico (vira o id)
    validar: Callable[[str], bool] = lambda v: bool(v)
    detectar: Callable[[str], int] = lambda t: 0        # confiança de 0 a 100 (sobre o texto digitado)
    formatar: Callable[[str], str] = lambda v: v        # valor canônico -> exibição
    rotulo_no: Optional[Callable[[dict], str]] = None   # dados do nó -> rótulo (padrão: valor formatado)
    coluna_banco: Optional[str] = None                  # None = tipo só manual (sem enrich por coluna)
    vizinho_no_enrich: bool = False                     # vira nó quando aparece na mesma assistência
    relacoes: Tuple[str, ...] = ()                      # rótulos sugeridos de vínculo (item 3)

    @property
    def cores(self) -> Dict[str, str]:
        """Formato {"bg", "border"} usado pelo vis-network e pela sidebar."""
        return {"bg": self.cor[0], "border": self.cor[1]}

    def node_id(self, valor_canonico: str) -> str:
        return f"{self.prefixo}_{valor_canonico}"

    def rotular(self, dados: dict) -> str:
        if self.rotulo_no:
            r = self.rotulo_no(dados)
            if r:
                return r
        return self.formatar(dados.get("valor", ""))

    def icone_svg(self, is_hub: bool = False) -> str:
        return svg_no(self.icone, self.cor[0], self.cor[1], is_hub)


class Candidato(NamedTuple):
    tipo: str          # chave do tipo
    confianca: int     # 0..100
    valor: str         # valor canônico


# Fonte dos rótulos dos nós no vis-network (a mesma do grafo do cluster, graph_engine).
FONTE_NO = {"color": "#EDEDED", "size": 11, "face": "Segoe UI, -apple-system, sans-serif",
            "background": "rgba(20, 20, 22, 0.94)", "strokeWidth": 0}

REGISTRO: Dict[str, TipoEntidade] = {}
_POR_PREFIXO: Dict[str, TipoEntidade] = {}
PISO_DETECCAO = 50


def registrar(t: TipoEntidade) -> TipoEntidade:
    if t.prefixo in _POR_PREFIXO and _POR_PREFIXO[t.prefixo].chave != t.chave:
        raise ValueError(f"prefixo {t.prefixo} já usado por {_POR_PREFIXO[t.prefixo].chave}")
    REGISTRO[t.chave] = t
    _POR_PREFIXO[t.prefixo] = t
    return t


def obter(chave: str) -> Optional[TipoEntidade]:
    return REGISTRO.get(str(chave or "").strip().lower().strip("[]"))


def tipo_do_id(node_id: str) -> Optional[TipoEntidade]:
    """Tipo pelo prefixo do id (CPF_123 -> cpf). None se o prefixo não é registrado."""
    prefixo = str(node_id).split("_", 1)[0]
    return _POR_PREFIXO.get(prefixo)


def valor_do_id(node_id: str) -> str:
    return str(node_id).split("_", 1)[-1]


def formatar_id(node_id: str) -> str:
    """Valor de exibição a partir do id ("CPF_12345678901" -> "123.456.789-01")."""
    t = tipo_do_id(node_id)
    v = valor_do_id(node_id)
    return t.formatar(v) if t else v


def detectar(texto: str, piso: int = PISO_DETECCAO) -> List[Candidato]:
    """
    Candidatos de tipo para um texto digitado/colado, do mais provável ao menos.

    Quando o texto serve para mais de um tipo (11 dígitos que fecham como CPF e também
    parecem celular), devolve os dois: o analista escolhe, o sistema não decide sozinho.
    Nenhum candidato acima do piso -> lista vazia (a interface pede o tipo).
    """
    texto = str(texto or "").strip()
    if not texto:
        return []
    out = []
    for t in REGISTRO.values():
        conf = t.detectar(texto)
        if conf >= piso:
            valor = t.normalizar(texto)
            if valor and t.validar(valor):
                out.append(Candidato(t.chave, conf, valor))
    return sorted(out, key=lambda c: -c.confianca)


def tipos_enriqueciveis() -> List[TipoEntidade]:
    """Tipos com coluna no banco: podem ser buscados no acervo (aba Add, Enrich)."""
    return [t for t in REGISTRO.values() if t.coluna_banco and t.vizinho_no_enrich]


def payload_tipos(chaves: Optional[set] = None) -> List[dict]:
    """O que o iframe do grafo precisa para montar filtros, legenda e menus por tipo."""
    return [{"chave": t.chave, "rotulo": t.rotulo, "cor": t.cores, "forma": t.forma, "icone": t.icone_svg()}
            for t in REGISTRO.values() if chaves is None or t.chave in chaves]


# ------------------------------------------------------------------ regras locais
def _digitos(t: str) -> str:
    return re.sub(r"\D", "", t)


def _sem_acento(t: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", t) if not unicodedata.combining(c))


def _texto_canonico(t: str) -> str:
    """Maiúsculas, sem acento, espaços simples (chave fraca de nomes)."""
    return re.sub(r"\s+", " ", _sem_acento(str(t)).upper()).strip()


def cpf_valido(d: str) -> bool:
    if len(d) != 11 or len(set(d)) == 1 or not d.isdigit():
        return False
    for n in (9, 10):
        s = sum(int(d[i]) * (n + 1 - i) for i in range(n))
        if (s * 10) % 11 % 10 != int(d[n]):
            return False
    return True


def cnpj_valido(d: str) -> bool:
    if len(d) != 14 or len(set(d)) == 1 or not d.isdigit():
        return False
    for n in (12, 13):
        pesos = list(range(n - 7, 1, -1)) + list(range(9, 1, -1))
        s = sum(int(d[i]) * pesos[i] for i in range(n))
        dv = 11 - s % 11
        if (0 if dv >= 10 else dv) != int(d[n]):
            return False
    return True


DDDS = {11, 12, 13, 14, 15, 16, 17, 18, 19, 21, 22, 24, 27, 28, 31, 32, 33, 34, 35, 37, 38,
        41, 42, 43, 44, 45, 46, 47, 48, 49, 51, 53, 54, 55, 61, 62, 63, 64, 65, 66, 67, 68, 69,
        71, 73, 74, 75, 77, 79, 81, 82, 83, 84, 85, 86, 87, 88, 89, 91, 92, 93, 94, 95, 96, 97, 98, 99}

_SO_DOCUMENTO = re.compile(r"^[\d\s.\-/]+$")
_SO_TELEFONE = re.compile(r"^\+?[\d\s().\-]+$")
_PLACA = re.compile(r"^[A-Z]{3}-?\d[A-Z0-9]\d{2}$")
_LOGRADOURO = re.compile(
    r"^(RUA|R|AVENIDA|AV|ALAMEDA|AL|TRAVESSA|TRAV|TV|RODOVIA|ROD|ESTRADA|ESTR|PRACA|PCA|LARGO)\b\.?\s+\S.*?\d+", re.I)
_ABREV_LOGRADOURO = {"R": "RUA", "AV": "AVENIDA", "AL": "ALAMEDA", "TRAV": "TRAVESSA", "TV": "TRAVESSA",
                     "ROD": "RODOVIA", "ESTR": "ESTRADA", "PCA": "PRACA"}


def _detectar_documento(t: str) -> int:
    """CPF ou CNPJ com dígitos verificadores válidos. Os dois vivem no tipo "cpf": a base grava
    CNPJ na coluna cpf e os nós já existentes são CPF_<cnpj>; um tipo à parte duplicaria a identidade."""
    d = _digitos(t)
    return 95 if _SO_DOCUMENTO.match(t) and (cpf_valido(d) or cnpj_valido(d)) else 0


def _detectar_telefone(t: str) -> int:
    if not _SO_TELEFONE.match(t):
        return 0
    d = _digitos(t)
    if d.startswith("55") and len(d) in (12, 13):
        d = d[2:]
    if len(d) not in (10, 11) or int(d[:2]) not in DDDS:
        return 0
    if len(d) == 11:
        return 90 if d[2] == "9" else 0          # celular começa com 9
    return 70 if d[2] in "2345" else 0           # fixo: 2 a 5 depois do DDD


def _detectar_placa(t: str) -> int:
    return 90 if _PLACA.match(t.strip().upper().replace(" ", "")) else 0


def _detectar_endereco(t: str) -> int:
    return 60 if _LOGRADOURO.match(t.strip()) else 0


def normalizar_endereco(t: str) -> str:
    """
    Logradouro + número, sem bairro (as bases divergem no bairro) e com o tipo de
    logradouro por extenso: "Av. Paulista, 1000" -> "AVENIDA PAULISTA, 1000".
    """
    s = _texto_canonico(t).replace(".", " ")
    s = re.sub(r"\s+", " ", s).strip()
    partes = s.split(" ", 1)
    if partes and partes[0] in _ABREV_LOGRADOURO:
        s = _ABREV_LOGRADOURO[partes[0]] + (" " + partes[1] if len(partes) > 1 else "")
    s = re.sub(r"\s*,\s*", ", ", s)
    if "," not in s:                       # "RUA 25 DE MARCO 100" -> "RUA 25 DE MARCO, 100"
        s = re.sub(r"\s+(\d+[A-Z]?)$", r", \1", s)
    return s


def normalizar_clinica(t: str) -> str:
    d = _digitos(t)
    return d if cnpj_valido(d) else _texto_canonico(t)


def formatar_placa(v: str) -> str:
    return f"{v[:3]}-{v[3:]}" if len(v) == 7 else v


def _rotulo_cpf(d: dict) -> str:
    nome = (d.get("nome_titular") or "").strip()
    return nome if nome and nome != "N/D" else formatar_cpf_cnpj(d.get("valor", ""))


# ------------------------------------------------------------------ tipos iniciais
# Paleta única = a do grafo (graph_engine). Tipos novos ficam fora do laranja (#FF7300),
# reservado a hub, seleção e destaque.
registrar(TipoEntidade(
    chave="cpf", prefixo="CPF", rotulo="Titular (CPF/CNPJ)", cor=("#3C6FA8", "#5A94D6"), icone="pessoa",
    campos=(Campo("cpf", "CPF", True, primario=True), Campo("nome_titular", "Nome do titular")),
    normalizar=limpar_cpf_cnpj, validar=lambda v: len(v) in (11, 14), detectar=_detectar_documento,
    formatar=formatar_cpf_cnpj, rotulo_no=_rotulo_cpf, coluna_banco="cpf", vizinho_no_enrich=True,
    relacoes=("pai", "mãe", "cônjuge", "irmão", "sócio", "laranja")))
registrar(TipoEntidade(
    chave="telefone", prefixo="TEL", rotulo="Telefone", cor=("#B94A3C", "#E0684F"), icone="telefone",
    campos=(Campo("telefone", "Número", True, primario=True), Campo("operadora", "Operadora")),
    normalizar=limpar_tel, validar=lambda v: len(v) in (10, 11), detectar=_detectar_telefone,
    formatar=formatar_tel, coluna_banco="telefone", vizinho_no_enrich=True))
registrar(TipoEntidade(
    chave="placa", prefixo="PLACA", rotulo="Placa", cor=("#6D4FA8", "#9273D6"), icone="carro",
    campos=(Campo("placa", "Placa", True, primario=True), Campo("modelo", "Modelo")),
    normalizar=limpar_placa, validar=lambda v: len(v) == 7, detectar=_detectar_placa,
    formatar=formatar_placa, coluna_banco="placa", vizinho_no_enrich=True))
registrar(TipoEntidade(
    chave="prestador", prefixo="PREST", rotulo="Prestador", cor=("#4A7A5A", "#6BA57D"), icone="predio",
    campos=(Campo("nome", "Nome", True, primario=True),), normalizar=_texto_canonico))
registrar(TipoEntidade(
    chave="empresa", prefixo="EMP", rotulo="Empresa", cor=("#4A4A4A", "#767676"), icone="empresa",
    campos=(Campo("nome", "Nome", True, primario=True),), normalizar=_texto_canonico,
    detectar=lambda t: 0, coluna_banco="empresa_cliente"))   # seguradora: fica na aresta, não vira nó no enrich
registrar(TipoEntidade(
    chave="endereco", prefixo="END", rotulo="Endereço", cor=("#3F7F86", "#5FA8B0"), icone="endereco",
    campos=(Campo("logradouro", "Logradouro e número", True, primario=True), Campo("complemento", "Complemento"),
            Campo("cidade", "Cidade", True), Campo("uf", "UF", True, tipo="uf"), Campo("cep", "CEP"),
            Campo("lat", "Latitude", tipo="numero"), Campo("lon", "Longitude", tipo="numero")),
    normalizar=normalizar_endereco, detectar=_detectar_endereco,
    relacoes=("mesmo endereço", "vizinhos")))
registrar(TipoEntidade(
    chave="clinica", prefixo="CLI", rotulo="Clínica", cor=("#8A6D3B", "#B8935A"), icone="clinica",
    campos=(Campo("identificacao", "CNPJ, ou nome + município", True, primario=True), Campo("razao_social", "Razão social", True),
            Campo("telefone", "Telefone"), Campo("endereco_ref", "Referência de endereço")),
    normalizar=normalizar_clinica, relacoes=("cliente", "localizada em")))
registrar(TipoEntidade(
    chave="pet", prefixo="PET", rotulo="Pet", cor=("#A84F7A", "#D6739F"), icone="pet",
    campos=(Campo("identificacao", "CPF do tutor + nome do pet", True, primario=True), Campo("especie", "Espécie"),
            Campo("raca", "Raça")),
    normalizar=_texto_canonico, relacoes=("tutor",)))
registrar(TipoEntidade(
    chave="assistencia", prefixo="ASS", rotulo="Assistência", cor=("#5B6472", "#8B95A5"), icone="documento",
    campos=(Campo("id_assistencia", "Id da assistência", True, primario=True), Campo("data", "Data", tipo="data"),
            Campo("servico", "Serviço"), Campo("cidade", "Cidade"), Campo("uf", "UF", tipo="uf")),
    normalizar=lambda t: str(t).strip(), coluna_banco="id_assistencia"))


# ------------------------------------------------------------------ sugestões de vínculo (item 3)
_RELACOES_POR_PAR = {
    frozenset({"cpf"}): ("pai", "mãe", "cônjuge", "irmão", "sócio", "laranja"),
    frozenset({"cpf", "placa"}): ("proprietário", "condutor"),
    frozenset({"cpf", "telefone"}): ("titular da linha", "usuário"),
    frozenset({"cpf", "clinica"}): ("cliente",),
    frozenset({"cpf", "pet"}): ("tutor",),
    frozenset({"cpf", "endereco"}): ("reside em",),
    frozenset({"clinica", "endereco"}): ("localizada em",),
}
RELACOES_GERAIS = ("vizinhos", "mesmo endereço")


def sugestoes_vinculo(tipo_a: str, tipo_b: str) -> Tuple[str, ...]:
    """Rótulos sugeridos para um vínculo manual entre dois tipos (texto livre continua aceito)."""
    return _RELACOES_POR_PAR.get(frozenset({tipo_a, tipo_b}), ()) + RELACOES_GERAIS


def payload_sugestoes() -> Dict[str, List[str]]:
    """Para o iframe: "tipoA|tipoB" (ordem alfabética) -> rótulos sugeridos, para todo par do registro."""
    chaves = sorted(REGISTRO)
    return {f"{a}|{b}": list(sugestoes_vinculo(a, b)) for i, a in enumerate(chaves) for b in chaves[i:]}
