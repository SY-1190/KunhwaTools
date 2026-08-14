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
  & $venvPython -m pip install --upgrade pip
  & $venvPython -m pip install -r (Join-Path $root 'requirements-pdf-excel.txt')
}

Set-Location -LiteralPath $root
& $venvPython (Join-Path $root 'server.py') --port $Port
