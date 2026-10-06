param(
  [int]$Port = 8878,
  [switch]$NoInstall
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$venv = Join-Path $root '.venv-pdf-excel'
$venvPython = Join-Path $venv 'Scripts\python.exe'

if (-not (Test-Path -LiteralPath $venvPython)) {
  if ($NoInstall) {
    throw 'The PDF conversion environment is not installed.'
  }
  $python = if ($env:KUNHWA_PYTHON) { $env:KUNHWA_PYTHON } else { 'python' }
  & $python -m venv $venv
  if ($LASTEXITCODE -ne 0) { throw 'Could not create the local PDF environment.' }
}

& $venvPython -c "import importlib.metadata as m, sys; packages={'PyMuPDF':'1.26.7','python-docx':'1.2.0','pdfplumber':'0.11.9','openpyxl':'3.1.5'}; installed={d.metadata['Name'].lower():d.version for d in m.distributions()}; sys.exit(0 if all(installed.get(k.lower())==v for k,v in packages.items()) else 1)"
if ($LASTEXITCODE -ne 0) {
  if ($NoInstall) { throw 'PDF dependencies need updating. Run without -NoInstall.' }
  & $venvPython -m pip install -r (Join-Path $root 'requirements-pdf-excel.txt')
  if ($LASTEXITCODE -ne 0) { throw 'Could not install the local PDF dependencies.' }
}

Set-Location -LiteralPath $root
& $venvPython (Join-Path $root 'server.py') --port $Port
