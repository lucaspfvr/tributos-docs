param([string]$InnoCompiler = "")
$ErrorActionPreference = "Stop"
Set-Location (Split-Path $PSScriptRoot)
if (-not (Test-Path ".venv\Scripts\python.exe")) {
    py -3.12 -m venv .venv
    if ($LASTEXITCODE -ne 0) { throw "Instale o Python 3.12 para Windows e o inicializador py." }
}
$Python = (Resolve-Path ".venv\Scripts\python.exe").Path
& $Python -m pip install -r requirements.txt "pyinstaller>=6.13,<7"
if ($LASTEXITCODE -ne 0) { throw "Falha ao instalar as dependências de compilação." }
$PyInstallerArgs = @(
    "--noconfirm", "--clean", "--windowed", "--onedir", "--name", "TributosDocs",
    "--add-data", "tributos/assets;tributos/assets", "--collect-all", "pypdfium2",
    "--collect-all", "pypdfium2_raw", "app.py"
)
& $Python -m PyInstaller @PyInstallerArgs
if ($LASTEXITCODE -ne 0) { throw "Falha ao empacotar o aplicativo." }
if (-not $InnoCompiler) {
    $InnoCompiler = @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
}
if (-not $InnoCompiler) {
    Write-Host "Aplicativo em dist\TributosDocs\TributosDocs.exe."
    throw "Instale o Inno Setup 6 ou informe o caminho de ISCC.exe com -InnoCompiler."
}
& $InnoCompiler "windows\installer.iss"
if ($LASTEXITCODE -ne 0) { throw "Falha ao gerar o instalador." }
