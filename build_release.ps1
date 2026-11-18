param(
    [switch]$ResetSyncOptOut
)

$ErrorActionPreference = 'Stop'

function Invoke-Native([string]$exe, [string[]]$argv) {
    & $exe @argv
    if ($LASTEXITCODE -ne 0) { throw "FALLO_NATIVO: $exe $($argv -join ' ') exit=$LASTEXITCODE" }
}

# Limpiar archivo de exclusión de sync (solo si se solicita explícitamente)
if ($ResetSyncOptOut) {
    $syncOptOutPath = Join-Path $env:LOCALAPPDATA "ConversorPDF\Data\.sync_optout"
    if (Test-Path $syncOptOutPath) {
        Remove-Item $syncOptOutPath -Force
        Write-Host "✅ Archivo de exclusión de sync eliminado: $syncOptOutPath"
    }
}

$env:Configuration = 'Release'
$env:Platform = 'x64'

Invoke-Native 'powershell' @('-ExecutionPolicy', 'Bypass', '-File', '.\build_project.ps1', '-Configuration', 'Release', '-Platform', 'x64')

Write-Host "Iniciando compilación del Instalador (WiX)..."
Invoke-Native 'dotnet' @('build', '.\installer\InstaladorConversorPDF.wixproj', '-c', 'Release', '-p:Platform=x64')

exit 0
