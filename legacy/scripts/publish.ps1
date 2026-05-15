# Ejemplo de pipeline diario (PC en LAN con biométrico + internet).
# Ajustá rutas y el destino de deploy (SFTP, rsync, cPanel, etc.).

param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$RemainingArgs
)

$ErrorActionPreference = "Stop"
# Este script está en legacy/scripts/ → subir dos niveles a la raíz del repo
Set-Location (Join-Path $PSScriptRoot "..\..")
$RepoRoot = (Get-Location).Path
$env:PYTHONPATH = "$RepoRoot;$RepoRoot\legacy"

Write-Host "== extract_biometrico (Excel) ==" -ForegroundColor Cyan
if (-not $RemainingArgs -or $RemainingArgs.Count -eq 0) {
    python biometrico/extract_biometrico.py --mes-actual
} else {
    python biometrico/extract_biometrico.py @RemainingArgs
}

Write-Host "== build HTML ==" -ForegroundColor Cyan
python legacy/doc/reporte_web.py

if ($env:MOTHERDUCK_TOKEN) {
    Write-Host "== upload MotherDuck (opcional) ==" -ForegroundColor Cyan
    python legacy/alexa/upload_motherduck.py @RemainingArgs
}

# Write-Host "== deploy (ejemplo) ==" -ForegroundColor Cyan
# scp legacy/reports/biometrico/reporte_biometrico.html user@host:/var/www/talento/

Write-Host "Listo." -ForegroundColor Green
