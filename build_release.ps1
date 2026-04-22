$ErrorActionPreference = 'Stop'
$env:Configuration = 'Release'
$env:Platform = 'x64'
& .\build_project.ps1

if ($LASTEXITCODE -eq 0) {
    Write-Host "Iniciando compilación del Instalador (WiX)..."
    dotnet build ".\Installer\InstaladorConversorPDF.wixproj" -c Release -p:Platform=x64
}