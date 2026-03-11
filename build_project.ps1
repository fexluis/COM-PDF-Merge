$ErrorActionPreference = 'Stop'
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
if (Test-Path $vswhere) {
    $msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe | Select-Object -First 1
}
if (-not $msbuild) {
    Write-Host "Trying fallback paths..."
    $candidates = @(
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    )
    foreach ($p in $candidates) { if (Test-Path $p) { $msbuild = $p; break } }
}

if ($msbuild) {
    Write-Host "MSBuild found: $msbuild"
    & $msbuild ".\ConversorPDF\ConversorPDF.csproj" -t:Restore
    if ($LASTEXITCODE -eq 0) {
        $config = if ($env:Configuration) { $env:Configuration } else { "Debug" }
        $plat = if ($env:Platform) { $env:Platform } else { "x64" }
        & $msbuild ".\ConversorPDF\ConversorPDF.csproj" /p:Configuration=$config /p:Platform=$plat /m /v:m
    }
} else {
    Write-Error "MSBuild not found."
}