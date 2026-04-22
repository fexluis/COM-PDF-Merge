param(
    [string]$NewVersion
)

$ErrorActionPreference = 'Stop'

# Fuente de verdad: App.cs
$APP_CS = "ConversorPDF.XLL\Core\Helpers\App.cs"

# 1. Actualizar versión si se proporciona el parámetro
if (-not [string]::IsNullOrWhiteSpace($NewVersion)) {
    Write-Host "--- Actualizando versión a $NewVersion ---" -ForegroundColor Yellow
    $content = Get-Content $APP_CS -Raw
    $newContent = $content -replace '(public const string APP_VERSION = )"[^"]+"', ("`$1`"$NewVersion`"")
    Set-Content -Path $APP_CS -Value $newContent -NoNewline
}

# Leer versión final
$match = Select-String -Path $APP_CS -Pattern 'APP_VERSION\s*=\s*"([^"]+)"' | Select-Object -First 1
if (-not $match) { throw "No se pudo leer APP_VERSION desde $APP_CS" }
$VERSION = $match.Matches[0].Groups[1].Value

$INSTALLER_DIR = "installer"
$OUTPUT_MSI = "$INSTALLER_DIR\InstaladorConversorPDFv$VERSION.msi"

Write-Host "--- Iniciando Proceso de Compilación Pro v$VERSION ---" -ForegroundColor Cyan

# 1. Limpieza
Write-Host "1. Limpiando instaladores antiguos..."
if (Test-Path $INSTALLER_DIR) {
    Get-ChildItem -Path $INSTALLER_DIR -Filter "*.msi" | Remove-Item -Force
}

# 2. Verificación de Tests
Write-Host "2. Ejecutando pruebas unitarias..."
dotnet test ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Release
if ($LASTEXITCODE -ne 0) { throw "Las pruebas unitarias fallaron. Abortando build." }

# 3. Compilación XLL (Release)
Write-Host "3. Compilando proyecto XLL (Release|x64)..."
dotnet build ConversorPDF.XLL\ConversorPDF.XLL.csproj -c Release -p:Platform=x64
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
