"""Form definitions and derived values for the municipal templates."""

import re
from dataclasses import dataclass
from datetime import date
from decimal import Decimal

from .docx import replace_tags

ASSETS = {
    "avaliacao": "Avaliação de imóvel",
    "desdobro": "Autorização de desdobro",
    "retificacao": "Autorização de retificação",
}

FIELDS = {
    "Emissão": {
        "NumeroDocumento": "Nº do documento", "Exercicio": "Exercício",
        "DataEmissao": "Data de emissão", "DataAssinatura": "Data da assinatura",
        "DataExtenso": "Data por extenso",
    },
    "Contribuinte": {
        "NomeProprietario": "Nome do proprietário / razão social",
        "CpfCnpj": "CPF / CNPJ", "Rg": "RG", "IdContribuinte": "Código do contribuinte",
        "NomeProprietario2": "Segundo proprietário (opcional)", "CpfCnpj2": "CPF / CNPJ do segundo",
    },
    "Endereço residencial": {
        "LogradouroResidencial": "Logradouro", "NumeroResidencial": "Número",
        "ComplementoResidencial": "Complemento", "BairroResidencial": "Bairro",
        "MunicipioResidencial": "Município", "UfResidencial": "UF", "CepResidencial": "CEP",
    },
    "Imóvel": {
        "CimImovel": "CIM do imóvel", "InscricaoImobiliaria": "Inscrição imobiliária",
        "Matricula": "Matrícula", "Livro": "Livro", "Setor": "Setor",
        "Quadra": "Quadra", "Lote": "Lote", "LogradouroImovel": "Logradouro",
        "NumeroImovel": "Número", "ComplementoImovel": "Complemento", "BairroImovel": "Bairro",
        "MunicipioImovel": "Município", "UfImovel": "UF", "CepImovel": "CEP",
        "AreaTerreno": "Área do terreno (m²)", "Testada": "Testada (m)",
        "AreaConstruida": "Área construída (m²)", "Pavimentos": "Pavimentos",
        "AnoEdificacao": "Ano da edificação", "AreaMensurada": "Área mensurada (m²)",
    },
    "Avaliação": {"ValorAvaliacao": "Valor da avaliação (R$)"},
}
REQUIRED = {
    "avaliacao": {"NomeProprietario", "NumeroDocumento", "Exercicio", "DataEmissao",
                 "ValorAvaliacao"},
    "desdobro": {"Lote", "Quadra", "Matricula", "NomeProprietario", "CpfCnpj", "CimImovel", "DataExtenso"},
    "retificacao": {"Matricula", "Livro", "CimImovel", "Quadra", "Lote", "AreaMensurada",
                    "LogradouroImovel", "NumeroImovel", "NomeProprietario", "CpfCnpj", "DataExtenso"},
}
DERIVED = {"Fracoes", "FracoesEmLinha", "QuantidadeFracoes", "DescricaoFracoes", "SegundoProprietario", "ValorExtenso"}
ADDRESS_PARTS = ("Logradouro", "Numero", "Complemento", "Bairro", "Municipio", "Uf", "Cep")
FRACTION_FIELDS = {"lote": "Lote / identificação", "area": "Área total (m²)",
                   "endereco": "Logradouro completo", "numero": "Número"}
MAX_FRACTIONS = 99
TAX_ID_FIELDS = {"CpfCnpj", "CpfCnpj2"}


@dataclass(frozen=True)
class Model:
    id: str
    title: str
    kind: str
    document: bytes

    def keys(self):
        return replace_tags(self.document)

    def groups(self):
        keys = self.keys()
        visible = set(keys) - DERIVED
        if "ValorExtenso" in keys:
            visible.add("ValorAvaliacao")
        if "SegundoProprietario" in keys:
            visible.update(("NomeProprietario2", "CpfCnpj2"))
        groups = {
            title: {key: label for key, label in fields.items() if key in visible}
            for title, fields in FIELDS.items()
        }
        known = {key for fields in FIELDS.values() for key in fields}
        groups["Campos adicionais"] = {
            key: re.sub(r"(?<=[a-z])(?=[A-Z])", " ", key).replace("_", " ")
            for key in keys if key in visible and key not in known
        }
        return {title: fields for title, fields in groups.items() if fields}

    def has_fractions(self):
        return bool({"Fracoes", "FracoesEmLinha", "QuantidadeFracoes", "DescricaoFracoes"} & set(self.keys()))


def defaults():
    today = date.today()
    months = "janeiro fevereiro março abril maio junho julho agosto setembro outubro novembro dezembro".split()
    return {
        "Exercicio": str(today.year),
        "DataEmissao": today.strftime("%d/%m/%Y"),
        "DataAssinatura": today.strftime("%d/%m/%Y"),
        "DataExtenso": f"{today.day} de {months[today.month - 1]} de {today.year}",
    }


def number_br(value, label, *, positive=False):
    value = value.strip()
    if not re.fullmatch(r"(?:[0-9]+|[0-9]{1,3}(?:\.[0-9]{3})+)(?:,[0-9]{1,2})?", value):
        raise ValueError(f"Informe {label} em formato brasileiro, por exemplo 150,00.")
    amount = Decimal(value.replace(".", "").replace(",", "."))
    if positive and amount <= 0:
        raise ValueError(f"Informe {label} maior que zero.")
    return format(amount, ",.2f").translate(str.maketrans(",.", ".,"))


def format_tax_id(value, *, partial=False):
    if not re.fullmatch(r"[0-9.\-/\s]*", value):
        return value
    digits = re.sub(r"[^0-9]", "", value)
    if len(digits) > 14 or (not partial and len(digits) not in (11, 14)):
        return value
    is_cnpj = len(digits) > 11 or (partial and "/" in value)
    sizes, separators = ((2, 3, 3, 4, 2), (".", ".", "/", "-")) if is_cnpj else ((3, 3, 3, 2), (".", ".", "-"))
    groups = []
    for size in sizes:
        if not digits:
            break
        groups.append(digits[:size])
        digits = digits[size:]
    if not groups:
        return ""
    return groups[0] + "".join(separator + group for separator, group in zip(separators, groups[1:]))


def normalize_fields(fields):
    values = {key: value.strip() for key, value in fields.items()}
    for key in ("NomeProprietario", "NomeProprietario2"):
        if key in values:
            values[key] = values[key].upper()
    for key in TAX_ID_FIELDS & values.keys():
        values[key] = format_tax_id(values[key])
    return values


def reais_por_extenso(amount):
    from num2words import num2words

    whole, cents = map(int, amount.replace(".", "").split(","))
    parts = []
    try:
        if whole or not cents:
            currency = "real" if whole == 1 else "reais"
            if whole >= 1_000_000 and whole % 1_000_000 == 0:
                currency = "de reais"
            parts.append(f"{num2words(whole, lang='pt_BR')} {currency}")
        if cents:
            unit = "centavo" if cents == 1 else "centavos"
            parts.append(f"{num2words(cents, lang='pt_BR')} {unit}")
    except OverflowError as exc:
        raise ValueError("O valor informado excede o limite da conversão por extenso.") from exc
    result = " e ".join(parts)
    return result[0].upper() + result[1:] + "."


def prepare(model, fields, fractions):
    keys = set(model.keys())
    values = normalize_fields(fields)
    for group in model.groups().values():
        for key, label in group.items():
            if key in REQUIRED.get(model.kind, ()) and not values.get(key):
                raise ValueError(f"Preencha o campo «{label}».")
    for key, label, unit in (
        ("AreaMensurada", "a área mensurada", ""),
        ("AreaTerreno", "a área do terreno", " m²"),
        ("AreaConstruida", "a área construída", " m²"),
        ("Testada", "a testada", " m"),
    ):
        if key in keys and values.get(key):
            values[key] = number_br(values[key], label, positive=(key == "AreaMensurada")) + unit
    if "ValorExtenso" in keys and not values.get("ValorAvaliacao"):
        raise ValueError("Preencha o valor da avaliação para gerar o texto por extenso.")
    if {"ValorAvaliacao", "ValorExtenso"} & keys and values.get("ValorAvaliacao"):
        values["ValorAvaliacao"] = number_br(values["ValorAvaliacao"], "o valor da avaliação")
        if "ValorExtenso" in keys:
            values["ValorExtenso"] = reais_por_extenso(values["ValorAvaliacao"])
    if model.kind == "retificacao" and values.get("LogradouroImovel"):
        values["LogradouroImovel"] = re.sub(r"^rua\s+", "", values["LogradouroImovel"], flags=re.I)
    if "SegundoProprietario" in keys:
        name, cpf = (values.get(key, "") for key in ("NomeProprietario2", "CpfCnpj2"))
        if bool(name) != bool(cpf):
            raise ValueError("Preencha o nome e o CPF/CNPJ do segundo proprietário, ou deixe ambos vazios.")
        values["SegundoProprietario"] = f" e de {name}, CPF: {cpf}" if name else ""
    if model.has_fractions():
        if not 1 <= len(fractions) <= MAX_FRACTIONS:
            raise ValueError(f"Informe entre 1 e {MAX_FRACTIONS} frações.")
        values["QuantidadeFracoes"] = f"{len(fractions):02d}"
        values["DescricaoFracoes"] = "fração assim denominada" if len(fractions) == 1 else "frações assim denominadas"
        lines = []
        inline = []
        for index, fraction in enumerate(fractions, 1):
            row = {key: fraction.get(key, "").strip() for key in FRACTION_FIELDS}
            for key, label in FRACTION_FIELDS.items():
                if key != "numero" and not row[key]:
                    raise ValueError(f"Fração {index} · preencha «{label}».")
            area = number_br(row["area"], f"a área da fração {index}", positive=True)
            lot = row["lote"]
            if not lot.casefold().startswith("lote "):
                lot = f"Lote {lot}"
            number = row["numero"]
            lines.append(f"{lot}, com uma área total de {area} m², situado em {row['endereco']}"
                         + (f", nº {number}" if number else ""))
            if "FracoesEmLinha" in keys:
                street = row["endereco"]
                if not re.match(r"(?i)^(rua|avenida|av\.|praça|travessa|estrada|rodovia|alameda|beco)\s", street):
                    street = "rua " + street
                address = f", {number}" if number.upper() in {"S/N", "SN"} else f", nº {number}" if number else ""
                if inline:
                    inline.append((" e " if index == len(fractions) else ", ", False))
                inline.extend([(lot, True), (f", com uma área total de {area}m², situado à {street}{address}", False)])
        values["Fracoes"] = ";\n".join(lines)
        values["FracoesEmLinha"] = inline
    return replace_tags(model.document, {key: values.get(key, "") for key in keys})
