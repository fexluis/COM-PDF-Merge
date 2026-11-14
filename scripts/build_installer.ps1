param(
    [string]$NewVersion,
    [switch]$SkipTests
)

$ErrorActionPreference = 'Stop'

# Fuente de verdad: App.cs
$APP_CS = "ConversorPDF.XLL\Core\Helpers\App.cs"

# 1. Actualizar versión si se proporciona el parámetro
if (-not [string]::IsNullOrWhiteSpace($NewVersion)) {
    Write-Host "--- Actualizando versión a $NewVersion ---" -ForegroundColor Yellow
    
    # 1.1 Actualizar App.cs
    $content = Get-Content $APP_CS -Raw
    $newContent = $content -replace '(public const string APP_VERSION = )"[^"]+"', ("`$1`"$NewVersion`"")
    Set-Content -Path $APP_CS -Value $newContent -NoNewline

    # 1.2 Actualizar .csproj
    $CSPROJ = "ConversorPDF.XLL\ConversorPDF.XLL.csproj"
    if (Test-Path $CSPROJ) {
        $csprojContent = Get-Content $CSPROJ -Raw
        $csprojContent = $csprojContent -replace '<Version>[^<]+</Version>', "<Version>$NewVersion</Version>"
        $csprojContent = $csprojContent -replace '<AssemblyVersion>[^<]+</AssemblyVersion>', "<AssemblyVersion>$NewVersion</AssemblyVersion>"
        $csprojContent = $csprojContent -replace '<FileVersion>[^<]+</FileVersion>', "<FileVersion>$NewVersion</FileVersion>"
        Set-Content -Path $CSPROJ -Value $csprojContent -NoNewline
    }

    # 1.3 Actualizar .wxs
    $WXS = "installer\InstaladorConversorPDF.wxs"
    if (Test-Path $WXS) {
        $wxsContent = Get-Content $WXS -Raw
        $wxsContent = $wxsContent -replace 'Version="[^"]+"', "Version=`"$NewVersion`""
        $wxsContent = $wxsContent -replace 'DefaultVersion="[^"]+"', "DefaultVersion=`"$NewVersion`""
        Set-Content -Path $WXS -Value $wxsContent -NoNewline
    }
}

# Leer versión final
$match = Select-String -Path $APP_CS -Pattern 'APP_VERSION\s*=\s*"([^"]+)"' | Select-Object -First 1
if (-not $match) { throw "No se pudo leer APP_VERSION desde $APP_CS" }
$VERSION = $match.Matches[0].Groups[1].Value

$INSTALLER_DIR = "installer"
$OUTPUT_MSI = "$INSTALLER_DIR\InstaladorConversorPDFv$VERSION.msi"

Write-Host "--- Iniciando Proceso de Compilación Pro v$VERSION ---" -ForegroundColor Cyan

# 1. Limpieza
Write-Host "1. Limpiando instaladores y procesos residuales..."
Get-Process | Where-Object { $_.ProcessName -match "^(testhost|vstest)" } | ForEach-Object {
    try { Stop-Process -Id $_.Id -Force -ErrorAction Stop }
    catch { }
}
if (Test-Path $INSTALLER_DIR) {
    Get-ChildItem -Path $INSTALLER_DIR -Filter "*.msi" | ForEach-Object {
        try { Remove-Item $_.FullName -Force -ErrorAction Stop }
        catch { Write-Warning "No se pudo eliminar $($_.Name) por estar en uso en otro proceso. Se omite." }
    }
}

# 2. Verificación de Tests
if (-not $SkipTests) {
    Write-Host "2. Ejecutando pruebas unitarias..."
    dotnet test ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Release
    if ($LASTEXITCODE -ne 0) { throw "Las pruebas unitarias fallaron. Abortando build." }
} else {
    Write-Host "2. Pruebas unitarias omitidas (-SkipTests)..." -ForegroundColor Yellow
}

# 3. Compilación XLL (Release)
Write-Host "3. Compilando proyecto XLL (Release|x64)..."
dotnet restore ConversorPDF.XLL\ConversorPDF.XLL.csproj -p:Platform=x64

# Asegurar dependencias para ExcelDnaPack
$outputPath = "ConversorPDF.XLL\bin\x64\Release"
if (!(Test-Path $outputPath)) { New-Item -ItemType Directory -Force -Path $outputPath }

$packages = "$env:USERPROFILE\.nuget\packages"
Copy-Item "$packages\system.reflection.metadata\6.0.0\lib\net461\System.Reflection.Metadata.dll" -Destination "$outputPath\" -Force
Copy-Item "$packages\system.collections.immutable\6.0.0\lib\net461\System.Collections.Immutable.dll" -Destination "$outputPath\" -Force
Copy-Item "$packages\csharpfunctionalextensions\2.40.3\lib\net461\CSharpFunctionalExtensions.dll" -Destination "$outputPath\" -Force

dotnet build ConversorPDF.XLL\ConversorPDF.XLL.csproj -c Release -p:Platform=x64 --no-restore
if ($LASTEXITCODE -ne 0) { throw "Error al compilar el proyecto XLL." }

# 4. Compilación MSI (Release)
Write-Host "4. Compilando Instalador MSI (Release)..."
dotnet build "$INSTALLER_DIR\InstaladorConversorPDF.wixproj" -c Release -p:Platform=x64
if ($LASTEXITCODE -ne 0) { throw "Error al generar el instalador MSI." }

# 5. Renombrado y Finalización
$GENERATED_MSI = "$INSTALLER_DIR\bin\x64\Release\InstaladorConversorPDF.msi"
if (Test-Path $GENERATED_MSI) {
    Move-Item -Path $GENERATED_MSI -Destination $OUTPUT_MSI -Force
    Write-Host "--- Proceso Completado Exitosamente ---" -ForegroundColor Green
    Write-Host "Artefacto generado: $OUTPUT_MSI" -ForegroundColor Yellow
} else {
    throw "No se encontró el MSI generado en $GENERATED_MSI"
}
