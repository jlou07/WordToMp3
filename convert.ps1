$ErrorActionPreference = 'Stop'
$python = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $python)) {
    throw 'Missing environment. Run: python -m venv .venv ; .\.venv\Scripts\python.exe -m pip install -r requirements.txt'
}
& $python (Join-Path $PSScriptRoot 'convert.py') @args
exit $LASTEXITCODE
