param(
    [string]$Configuration = 'Release',
    [string]$Platform = 'x64'
)

$ErrorActionPreference = 'Stop'

function Invoke-Native([string]$exe, [string[]]$argv) {
    & $exe @argv
    if ($LASTEXITCODE -ne 0) { throw "FALLO_NATIVO: $exe $($argv -join ' ') exit=$LASTEXITCODE" }
}

$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (Test-Path $vswhere) {
    # Preferir MSBuild de 64 bits para evitar errores de packing (System.BadImageFormatException)
    $msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\amd64\MSBuild.exe | Select-Object -First 1
    if (-not $msbuild) {
        $msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe | Select-Object -First 1
    }
}
if (-not $msbuild) {
    Write-Host "Trying fallback paths..."
    $candidates = @(
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\amd64\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\amd64\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\amd64\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\amd64\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\amd64\MSBuild.exe",
        "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe"
    )
    foreach ($p in $candidates) { if (Test-Path $p) { $msbuild = $p; break } }
}

if (-not $msbuild) {
    throw "MSBUILD_NO_ENCONTRADO"
}

Write-Host "MSBuild found: $msbuild"
$config = if ($env:Configuration) { $env:Configuration } else { $Configuration }
$plat = if ($env:Platform) { $env:Platform } else { $Platform }

Write-Host "Restoring ConversorPDF.XLL..."
Invoke-Native $msbuild @(".\ConversorPDF.XLL\ConversorPDF.XLL.csproj", "-t:Restore")

# Asegurar que System.Reflection.Metadata.dll esté en la carpeta de salida para el empaquetador
$outputPath = ".\ConversorPDF.XLL\bin\$plat\$config"
if (!(Test-Path $outputPath)) { New-Item -ItemType Directory -Force -Path $outputPath }

$metadataDll = "$env:USERPROFILE\.nuget\packages\system.reflection.metadata\6.0.0\lib\net461\System.Reflection.Metadata.dll"
if (Test-Path $metadataDll) { Copy-Item $metadataDll -Destination "$outputPath\" -Force }

$immutableDll = "$env:USERPROFILE\.nuget\packages\system.collections.immutable\6.0.0\lib\net461\System.Collections.Immutable.dll"
if (Test-Path $immutableDll) { Copy-Item $immutableDll -Destination "$outputPath\" -Force }

$functionalDll = "$env:USERPROFILE\.nuget\packages\csharpfunctionalextensions\2.40.3\lib\net461\CSharpFunctionalExtensions.dll"
if (Test-Path $functionalDll) { Copy-Item $functionalDll -Destination "$outputPath\" -Force }

# T3: clon profundo del escalador; debe existir junto al add-in para resolver en tiempo de ejecución
$newtonsoftDll = "$env:USERPROFILE\.nuget\packages\newtonsoft.json\13.0.1\lib\net45\Newtonsoft.Json.dll"
if (Test-Path $newtonsoftDll) { Copy-Item $newtonsoftDll -Destination "$outputPath\" -Force }

Write-Host "Iniciando compilación de ConversorPDF.XLL..."
Invoke-Native $msbuild @(".\ConversorPDF.XLL\ConversorPDF.XLL.csproj", "/p:Configuration=$config", "/p:Platform=$plat", "/m", "/v:m")

exit 0
