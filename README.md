# Tributos Docs

Aplicação desktop para automação de processos documentais do setor de tributos. O projeto nasceu da rotina de Avaliação, Desdobro e Retificação em Biquinhas/MG.

O operador preenche os dados e confere a prévia. O programa aplica os campos ao modelo Word, gera o PDF, sugere o número do documento e registra a emissão no histórico local.

## Fluxo de trabalho

1. Abra um formulário de Avaliação, Desdobro ou Retificação.
2. Preencha os campos e confira a prévia em **Revisar documento**.
3. Use **Salvar PDF** para arquivar. O arquivo aparece em **Histórico**, com as opções **Abrir PDF**, **Detalhes** e **Editar**.
4. Use **Configurações → Modelos de documentos** para editar ou cadastrar um modelo. Uma chave como `<CimImovel>` vira um campo do formulário.

Cada tipo tem sua sequência local. Uma emissão salva avança a sequência do seu tipo. Uma prévia fechada deixa o número disponível. Editar pelo histórico abre uma revisão com o número original. No Desdobro, cada fração começa em uma linha própria.

## Rodar pelo código no Windows

Requisitos de desenvolvimento

- Python 3.12 com o inicializador `py`.
- Microsoft Word desktop ou LibreOffice Writer para a conversão em PDF.

Na pasta que contém `app.py`, abra o PowerShell.

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe app.py
```

Para trabalhar com uma base de teste separada, defina as duas variáveis abaixo antes de abrir o aplicativo.

```powershell
$env:TRIBUTOS_DOCS_DATA = Join-Path $env:TEMP "tributos-docs-demo-dados"
$env:TRIBUTOS_DOCS_ARCHIVE = Join-Path $env:TEMP "tributos-docs-demo-pdfs"
.\.venv\Scripts\python.exe app.py
```

## Organização do código

| Caminho | Responsabilidade |
|---|---|
| `app.py` | Ponto de entrada |
| `tributos/ui.py` | Telas, formulários, prévia e interação |
| `tributos/models.py` | Campos, validação e regras de preenchimento |
| `tributos/docx.py` | Edição do modelo Word e exportação por Word ou Writer |
| `tributos/storage.py` | SQLite, numeração, revisões, arquivo e backups |
| `tributos/assets/` | Modelos Word de demonstração e ícones |
| `windows/` | Inno Setup e código do empacotador offline para Windows |
| `tests/` | Testes de preenchimento, arquivo, exportação e interface |

## Testes

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

A suíte inclui testes gráficos quando há uma sessão Tk disponível e testes de conversão quando o Writer está no PATH. A distribuição para Windows deve ser validada com o conversor e a impressora da estação de destino.

## Distribuição no Windows

A v0.13 foi empacotada para Windows x64. O código do instalador offline está em `windows/offline/`, junto com o empacotador e os avisos de terceiros. O instalador pode oferecer a Aptos pelo pacote oficial da Microsoft. O repositório de código usa modelos demonstrativos e deve ser distribuído separadamente da instalação que contém os modelos de trabalho.

## Dados e histórico do projeto

Os modelos incluídos neste repositório são demonstrativos e sem validade. Use dados fictícios nas capturas e nos testes. Documentos emitidos, bancos SQLite e backups ficam fora do Git. A numeração fica na base local de cada estação.

O histórico deste repositório é uma reconstrução incremental do projeto concluído. As primeiras etapas isolam a janela, o preenchimento do Word e a geração de PDF. Os marcos seguintes acompanham as versões do aplicativo até a v0.13.
