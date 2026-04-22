[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]$NewVersion
)

$ErrorActionPreference = "Stop"
$root = Resolve-Path "$PSScriptRoot\..\..\..\.."
$appCs = Join-Path $root "ConversorPDF.XLL\Core\App.cs"
$csproj = Join-Path $root "ConversorPDF.XLL\ConversorPDF.XLL.csproj"
$wxs = Join-Path $root "installer\InstaladorConversorPDF.wxs"
$wixproj = Join-Path $root "installer\InstaladorConversorPDF.wixproj"
$installerDir = Join-Path $root "installer"

Write-Host "--- Iniciando Proceso de Compilación Pro v$NewVersion ---" -ForegroundColor Cyan

# 1. Cleaning
Write-Host "1. Limpiando archivos MSI antiguos..." -ForegroundColor Yellow
Get-ChildItem -Path $installerDir -Filter "*.msi" | Remove-Item -Force -ErrorAction SilentlyContinue

# 2. Version Synchronization
Write-Host "2. Sincronizando números de versión..." -ForegroundColor Yellow

# App.cs
(Get-Content $appCs) -replace 'APP_VERSION = ".*"', "APP_VERSION = `"$NewVersion`"" | Set-Content $appCs

# CSPROJ
(Get-Content $csproj) `
    -replace '<Version>.*</Version>', "<Version>$NewVersion</Version>" `
    -replace '<AssemblyVersion>.*</AssemblyVersion>', "<AssemblyVersion>$NewVersion</AssemblyVersion>" `
    -replace '<FileVersion>.*</FileVersion>', "<FileVersion>$NewVersion</FileVersion>" | Set-Content $csproj

# WXS (Version y MajorUpgrade Schedule)
$content = Get-Content $wxs
$content = $content -replace '(<Package[^>]*\sVersion=")[^"]+(")', "`${1}$NewVersion`${2}"
$content = $content -replace '(<MajorUpgrade[^>]*\sSchedule=")[^"]+(")', "`${1}afterInstallValidate`${2}"
$content | Set-Content $wxs

# 3. Validating Tests
Write-Host "3. Validando pruebas unitarias..." -ForegroundColor Yellow
$testProj = Join-Path $root "ConversorPDF.Tests\ConversorPDF.Tests.csproj"
dotnet test $testProj -c Release
if ($LASTEXITCODE -ne 0) { throw "Las pruebas unitarias fallaron. Abortando proceso." }
 Write-Host "4. Compilando XLL en modo Release..." -ForegroundColor Yellow
dotnet build $csproj -c Release
if ($LASTEXITCODE -ne 0) { throw "Fallo al compilar el proyecto XLL." }

# 5. Build MSI
Write-Host "5. Generando instalador MSI..." -ForegroundColor Yellow
dotnet build $wixproj -c Release
if ($LASTEXITCODE -ne 0) { throw "Fallo al compilar el instalador MSI." }

# 6. Deployment & Rename
Write-Host "6. Desplegando y renombrando instalador..." -ForegroundColor Yellow
$msiSource = Join-Path $root "installer\bin\Release\InstaladorConversorPDF.msi"
$msiTarget = Join-Path $installerDir "InstaladorConversorPDFv$NewVersion.msi"

if (Test-Path $msiSource) {
    Move-Item -Path $msiSource -Destination $msiTarget -Force
    Write-Host "Éxito: $msiTarget generado." -ForegroundColor Green
} else {
    Write-Error "Error: No se encontró el MSI generado en $msiSource"
}

# 7. Verification
Write-Host "7. Verificando versión del ensamblado..." -ForegroundColor Yellow
$dllPath = Join-Path $root "ConversorPDF.XLL\bin\x64\Release\ConversorPDF.XLL.dll"
if (Test-Path $dllPath) {
    $vi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($dllPath)
    if ($vi.FileVersion -eq $NewVersion) {
        Write-Host "Verificación OK: Ensamblado v$($vi.FileVersion)" -ForegroundColor Green
    } else {
        Write-Warning "Advertencia: Versión del ensamblado ($($vi.FileVersion)) no coincide con $NewVersion"
    }
}

Write-Host "--- Proceso Finalizado Exitosamente ---" -ForegroundColor Cyan
