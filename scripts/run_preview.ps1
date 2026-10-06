$Root = Split-Path -Parent $PSScriptRoot
$Python = Join-Path $Root ".venv\Scripts\python.exe"
if (-not (Test-Path $Python)) {
    Write-Error "Missing venv. Run scripts/install.ps1 first."
}
$env:PYTHONPATH = Join-Path $Root "src"
& $Python -m wpif
