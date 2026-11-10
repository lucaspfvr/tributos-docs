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
    "Avaliação": {"ValorAvaliacao": "Valor da avaliação (R$)", "ValorExtenso": "Valor por extenso"},
}
REQUIRED = {
    "avaliacao": {"NomeProprietario", "NumeroDocumento", "Exercicio", "DataEmissao",
                 "ValorAvaliacao", "ValorExtenso"},
    "desdobro": {"Lote", "Quadra", "Matricula", "NomeProprietario", "CpfCnpj", "CimImovel", "DataExtenso"},
    "retificacao": {"Matricula", "Livro", "CimImovel", "Quadra", "Lote", "AreaMensurada",
                    "LogradouroImovel", "NumeroImovel", "NomeProprietario", "CpfCnpj", "DataExtenso"},
}
DERIVED = {"Fracoes", "QuantidadeFracoes", "DescricaoFracoes", "SegundoProprietario"}
FRACTION_FIELDS = {"lote": "Lote / identificação", "area": "Área total (m²)",
                   "endereco": "Logradouro completo", "numero": "Número"}
MAX_FRACTIONS = 99


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
        return bool({"Fracoes", "QuantidadeFracoes", "DescricaoFracoes"} & set(self.keys()))


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


def prepare(model, fields, fractions):
    keys = set(model.keys())
    values = {key: value.strip() for key, value in fields.items()}
    for group in model.groups().values():
        for key, label in group.items():
            if key in REQUIRED.get(model.kind, ()) and not values.get(key):
                raise ValueError(f"Preencha o campo «{label}».")
    for key, label, unit in (
        ("AreaMensurada", "a área mensurada", ""),
        ("AreaTerreno", "a área do terreno", " m²"),
        ("AreaConstruida", "a área construída", " m²"),
        ("Testada", "a testada", " m"),
        ("ValorAvaliacao", "o valor da avaliação", ""),
    ):
        if key in keys and values.get(key):
            values[key] = number_br(values[key], label, positive=(key == "AreaMensurada")) + unit
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
        for index, fraction in enumerate(fractions, 1):
            row = {key: fraction.get(key, "").strip() for key in FRACTION_FIELDS}
            for key, label in FRACTION_FIELDS.items():
                if not row[key]:
                    raise ValueError(f"Fração {index} · preencha «{label}».")
            area = number_br(row["area"], f"a área da fração {index}", positive=True)
            lot = row["lote"]
            if not lot.casefold().startswith("lote "):
                lot = f"Lote {lot}"
            lines.append(f"{lot}, com uma área total de {area} m², situado em {row['endereco']}, nº {row['numero']}")
        values["Fracoes"] = ";\n".join(lines)
    return replace_tags(model.document, {key: values.get(key, "") for key in keys})
