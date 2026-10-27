$ErrorActionPreference = 'Stop'

# ============================================================
# Limpiar archivo de exclusión de sync (si existe)
# ============================================================
$syncOptOutPath = Join-Path $env:LOCALAPPDATA "ConversorPDF\Data\.sync_optout"
if (Test-Path $syncOptOutPath) {
    Remove-Item $syncOptOutPath -Force
    Write-Host "✅ Archivo de exclusión de sync eliminado: $syncOptOutPath"
}
$env:Configuration = 'Release'
$env:Platform = 'x64'
& .\build_project.ps1

if ($LASTEXITCODE -eq 0) {
    Write-Host "Iniciando compilación del Instalador (WiX)..."
    dotnet build ".\Installer\InstaladorConversorPDF.wixproj" -c Release -p:Platform=x64
}