# Script para construir la documentación técnica con DocFX
# Uso: .\scripts\build_docs.ps1

$ErrorActionPreference = "Stop"
$root = Resolve-Path "$PSScriptRoot\.."
$docsDir = Join-Path $root "docs"
$siteDir = Join-Path $docsDir "_site"
$docfxJson = Join-Path $docsDir "docfx.json"
$csproj = Join-Path $root "ConversorPDF.XLL\ConversorPDF.XLL.csproj"

Write-Host "--- Iniciando Generación de Documentación ---" -ForegroundColor Cyan

# 1. Verificar si docfx está instalado
Write-Host "1. Verificando herramientas..." -ForegroundColor Yellow
$docfxInstalled = dotnet tool list -g | Select-String "docfx"
if (-not $docfxInstalled) {
    Write-Host "Instalando DocFX globalmente..." -ForegroundColor Green
    dotnet tool install -g docfx
}

# 2. Compilar el proyecto para generar el XML de documentación
Write-Host "2. Preparando binarios y XML..." -ForegroundColor Yellow
dotnet build $csproj -c Release
if ($LASTEXITCODE -ne 0) { throw "El build del proyecto falló." }

# 3. Generar Metadatos y Construir Sitio
Write-Host "3. Ejecutando DocFX Build..." -ForegroundColor Yellow
Set-Location $docsDir
& docfx build $docfxJson
if ($LASTEXITCODE -ne 0) { throw "DocFX falló al construir el sitio." }

# 4. Configurar para GitHub Pages
Write-Host "4. Preparando para GitHub Pages..." -ForegroundColor Yellow
$noJekyll = Join-Path $siteDir ".nojekyll"
if (-not (Test-Path $noJekyll)) {
    New-Item -Path $noJekyll -ItemType File -Force
}

# 5. Finalizar
Write-Host "--- Documentación generada en: $siteDir ---" -ForegroundColor Cyan
Write-Host "Para previsualizar localmente, ejecuta: docfx serve `"$siteDir`"" -ForegroundColor Gray

Set-Location $root
