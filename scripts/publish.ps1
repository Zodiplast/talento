# Ejemplo de pipeline diario (PC en LAN con biométrico + internet).
# Ajustá rutas y el destino de deploy (SFTP, rsync, cPanel, etc.).

param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$RemainingArgs
)

$ErrorActionPreference = "Stop"
Set-Location (Join-Path $PSScriptRoot "..")

Write-Host "== sync_device ==" -ForegroundColor Cyan
python biometrico/sync_device.py @RemainingArgs

Write-Host "== build HTML ==" -ForegroundColor Cyan
python doc/reporte_web.py

if ($env:MOTHERDUCK_TOKEN) {
    Write-Host "== upload MotherDuck (opcional) ==" -ForegroundColor Cyan
    python biometrico/upload_motherduck.py @RemainingArgs
}

# Write-Host "== deploy (ejemplo) ==" -ForegroundColor Cyan
# scp reports/biometrico/reporte_biometrico.html user@host:/var/www/talento/

Write-Host "Listo." -ForegroundColor Green
