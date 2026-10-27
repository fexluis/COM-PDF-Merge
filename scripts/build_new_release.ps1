<#
.SYNOPSIS
    Script maestro de build/release para ConversorPDF — diseñado para ser llamado desde Cline.

.DESCRIPTION
    Automatiza el flujo completo de release:
    1. Backup de archivos de versión
    2. Sincronización de versión en App.cs, .csproj, .wxs
    3. Validación de referencias .dna
    4. Ejecución de tests unitarios
    5. Compilación XLL (Release|x64)
    6. Validación del XLL empaquetado
    7. Compilación del instalador MSI (WiX 4)
    8. Verificación del artefacto final
    9. Rollback automático si falla cualquier paso

.PARAMETER NewVersion
    Nueva versión en formato Major.Minor.Patch.Build (ej: "1.2.53.20")

.PARAMETER SkipTests
    Omite la ejecución de tests unitarios (para builds rápidas de debug)

.PARAMETER Verbose
    Muestra información detallada de cada paso

.EXAMPLE
    .\scripts\build_new_release.ps1 -NewVersion "1.2.53.20"

.EXAMPLE
    .\scripts\build_new_release.ps1 -NewVersion "1.2.53.20" -SkipTests -Verbose
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$NewVersion,

    [switch]$SkipTests,
    [switch]$ShowVerbose
)

$ErrorActionPreference = 'Stop'
if ($ShowVerbose) { $VerbosePreference = 'Continue' }

# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIGURACIÓN
# ═══════════════════════════════════════════════════════════════════════════════

$APP_CS = "ConversorPDF.XLL\Core\Helpers\App.cs"
$CSPROJ = "ConversorPDF.XLL\ConversorPDF.XLL.csproj"
$WXS = "installer\InstaladorConversorPDF.wxs"
$DNA = "ConversorPDF.XLL\ConversorPDF.XLL-AddIn64.dna"
$TEST_PROJ = "ConversorPDF.Tests\ConversorPDF.Tests.csproj"
$XLL_PROJ = "ConversorPDF.XLL\ConversorPDF.XLL.csproj"
$WIX_PROJ = "installer\InstaladorConversorPDF.wixproj"
$INSTALLER_DIR = "installer"
$BACKUP_DIR = ".build-backup"
$OUTPUT_PATH = "ConversorPDF.XLL\bin\x64\Release"

# ═══════════════════════════════════════════════════════════════════════════════
#  FUNCIONES AUXILIARES
# ═══════════════════════════════════════════════════════════════════════════════

function Write-Step {
    param([string]$Message)
    Write-Host "`n=== $Message ===" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host "✅ $Message" -ForegroundColor Green
}

function Write-Warning {
    param([string]$Message)
    Write-Host "⚠️  $Message" -ForegroundColor Yellow
}

function Write-Error {
    param([string]$Message)
    Write-Host "❌ $Message" -ForegroundColor Red
}

function Backup-VersionFiles {
    Write-Step "FASE 0: Backup de archivos de versión"
    if (!(Test-Path $BACKUP_DIR)) {
        New-Item -ItemType Directory -Force -Path $BACKUP_DIR | Out-Null
    }
    $files = @($APP_CS, $CSPROJ, $WXS)
    foreach ($f in $files) {
        if (Test-Path $f) {
            $dest = "$BACKUP_DIR\$(Split-Path $f -Leaf).bak"
            Copy-Item $f $dest -Force
            Write-Host "   Backup: $f → $dest"
        }
    }
    Write-Success "Backup completado"
}

function Restore-VersionFiles {
    Write-Step "ROLLBACK: Restaurando archivos de versión"
    $files = @($APP_CS, $CSPROJ, $WXS)
    foreach ($f in $files) {
        $bak = "$BACKUP_DIR\$(Split-Path $f -Leaf).bak"
        if (Test-Path $bak) {
            Copy-Item $bak $f -Force
            Write-Host "   Restaurado: $f"
        }
    }
    Write-Success "Rollback completado"
}

function Update-Version {
    param([string]$Version)
    Write-Step "FASE 1: Sincronizando versión a $Version"

    # 1.1 App.cs
    $content = Get-Content $APP_CS -Raw
    $newContent = $content -replace '(public const string APP_VERSION = )"[^"]+"', ("`$1`"$Version`"")
    Set-Content -Path $APP_CS -Value $newContent -NoNewline
    Write-Host "   ✅ App.cs actualizado"

    # 1.2 .csproj
    if (Test-Path $CSPROJ) {
        $csprojContent = Get-Content $CSPROJ -Raw
        $csprojContent = $csprojContent -replace '<Version>[^<]+</Version>', "<Version>$Version</Version>"
        $csprojContent = $csprojContent -replace '<AssemblyVersion>[^<]+</AssemblyVersion>', "<AssemblyVersion>$Version</AssemblyVersion>"
        $csprojContent = $csprojContent -replace '<FileVersion>[^<]+</FileVersion>', "<FileVersion>$Version</FileVersion>"
        Set-Content -Path $CSPROJ -Value $csprojContent -NoNewline
        Write-Host "   ✅ .csproj actualizado"
    }

    # 1.3 .wxs
    if (Test-Path $WXS) {
        $wxsContent = Get-Content $WXS -Raw
        $wxsContent = $wxsContent -replace 'Version="[^"]+"', "Version=`"$Version`""
        Set-Content -Path $WXS -Value $wxsContent -NoNewline
        Write-Host "   ✅ .wxs actualizado"
    }

    Write-Success "Versión sincronizada en todos los archivos"
}

function Test-DnaReferences {
    Write-Step "FASE 2: Validando referencias .dna"
    if (!(Test-Path $DNA)) {
        throw "Archivo .dna no encontrado: $DNA"
    }

    # Asegurar que el directorio de salida exista (puede que no haya compilado aún)
    if (!(Test-Path $OUTPUT_PATH)) {
        Write-Warning "Directorio $OUTPUT_PATH no existe aún (se validará post-build)"
        return
    }

    [xml]$dnaXml = Get-Content $DNA
    $refs = $dnaXml.DnaLibrary.Reference | Where-Object { $_.Pack -eq "true" }
    $missing = @()

    foreach ($ref in $refs) {
        $dllName = $ref.Path
        $fullPath = Join-Path $OUTPUT_PATH $dllName
        if (!(Test-Path $fullPath)) {
            $missing += $dllName
        }
    }

    if ($missing.Count -gt 0) {
        $msg = "REFERENCIAS .dna INVÁLIDAS:`n" +
        ($missing | ForEach-Object { "   - $_`n" }) +
        "No existen en $OUTPUT_PATH.`n" +
        "Verificar Lección Aprendida #1 (Microsoft.Bcl.HashCode.dll)."
        throw $msg
    }

    Write-Success "Todas las referencias .dna son válidas ($($refs.Count) archivos verificados)"
}

function Invoke-Tests {
    Write-Step "FASE 3: Ejecutando pruebas unitarias"
    dotnet test $TEST_PROJ -c Release --no-build
    if ($LASTEXITCODE -ne 0) {
        throw "Las pruebas unitarias fallaron. Use -SkipTests para omitir."
    }
    Write-Success "Todas las pruebas pasaron"
}

function Build-XLL {
    Write-Step "FASE 4: Compilando proyecto XLL (Release|x64)"

    # Restore
    dotnet restore $XLL_PROJ -p:Platform=x64
    if ($LASTEXITCODE -ne 0) { throw "Error en dotnet restore del XLL" }

    # Asegurar directorio de salida
    if (!(Test-Path $OUTPUT_PATH)) {
        New-Item -ItemType Directory -Force -Path $OUTPUT_PATH | Out-Null
    }

    # Copiar dependencias necesarias para ExcelDnaPack
    $packages = "$env:USERPROFILE\.nuget\packages"
    $deps = @(
        @{ Source = "$packages\system.reflection.metadata\6.0.0\lib\net461\System.Reflection.Metadata.dll"; Dest = "$OUTPUT_PATH\" },
        @{ Source = "$packages\system.collections.immutable\6.0.0\lib\net461\System.Collections.Immutable.dll"; Dest = "$OUTPUT_PATH\" },
        @{ Source = "$packages\csharpfunctionalextensions\2.40.3\lib\net461\CSharpFunctionalExtensions.dll"; Dest = "$OUTPUT_PATH\" }
    )

    foreach ($dep in $deps) {
        if (Test-Path $dep.Source) {
            Copy-Item $dep.Source -Destination $dep.Dest -Force
            Write-Host "   Copiado: $(Split-Path $dep.Source -Leaf)"
        }
        else {
            Write-Warning "Dependencia no encontrada: $($dep.Source)"
        }
    }

    # Build
    dotnet build $XLL_PROJ -c Release -p:Platform=x64 --no-restore
    if ($LASTEXITCODE -ne 0) { throw "Error al compilar el proyecto XLL" }

    Write-Success "XLL compilado exitosamente"
}

function Validate-PackedXLL {
    Write-Step "FASE 5: Validando XLL empaquetado"
    $packedXll = "$OUTPUT_PATH\ConversorPDF.XLL-AddIn64-packed.xll"

    if (!(Test-Path $packedXll)) {
        throw "El XLL empaquetado NO se generó: $packedXll"
    }

    $size = (Get-Item $packedXll).Length
    if ($size -lt 1024) {
        throw "El XLL empaquetado parece corrupto (tamaño: $size bytes)"
    }

    # Validar SQLite.Interop.dll en subcarpeta x64
    $sqliteInterop = "$OUTPUT_PATH\x64\SQLite.Interop.dll"
    if (!(Test-Path $sqliteInterop)) {
        throw "SQLite.Interop.dll no encontrado en $OUTPUT_PATH\x64\"
    }

    Write-Success "XLL empaquetado válido ($size bytes)"
}

function Build-MSI {
    param([string]$Version)
    Write-Step "FASE 6: Compilando Instalador MSI (WiX 4)"

    dotnet build $WIX_PROJ -c Release -p:Platform=x64
    if ($LASTEXITCODE -ne 0) { throw "Error al generar el instalador MSI" }

    $generated = "$INSTALLER_DIR\bin\x64\Release\InstaladorConversorPDF.msi"
    $outputMsi = "$INSTALLER_DIR\InstaladorConversorPDFv$Version.msi"

    if (Test-Path $generated) {
        Move-Item -Path $generated -Destination $outputMsi -Force
        Write-Success "MSI generado: $outputMsi"
    }
    else {
        throw "No se encontró el MSI generado en $generated"
    }

    return $outputMsi
}

function Verify-MSI {
    param([string]$MsiPath)
    Write-Step "FASE 7: Verificación final del artefacto"

    if (!(Test-Path $MsiPath)) {
        throw "MSI no encontrado: $MsiPath"
    }

    $fileInfo = Get-Item $MsiPath
    $sizeMB = [math]::Round($fileInfo.Length / 1MB, 2)

    if ($fileInfo.Length -lt 1MB) {
        throw "MSI sospechosamente pequeño ($sizeMB MB). Posible build incompleto."
    }

    Write-Success "Artefacto verificado: $($fileInfo.Name) ($sizeMB MB)"
}

# ═══════════════════════════════════════════════════════════════════════════════
#  FLUJO PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════

try {
    Write-Host "`n═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "   ConversorPDF - Build & Release v$NewVersion" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan

    # ─── FASE 0: Backup ───
    Backup-VersionFiles

    # ─── FASE 1: Versionado ───
    Update-Version -Version $NewVersion

    # ─── FASE 2: Validar .dna (pre-build) ───
    Test-DnaReferences

    # ─── FASE 3: Tests ───
    if (!$SkipTests) {
        Invoke-Tests
    }
    else {
        Write-Warning "Tests omitidos (-SkipTests)"
    }

    # ─── FASE 4: Build XLL ───
    Build-XLL

    # ─── FASE 5: Validar XLL ───
    Validate-PackedXLL

    # ─── FASE 6: Build MSI ───
    $finalMsi = Build-MSI -Version $NewVersion

    # ─── FASE 7: Verificar ───
    Verify-MSI -MsiPath $finalMsi

    # ─── ÉXITO ───
    Write-Host "`n═══════════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "   ✅ BUILD COMPLETADO EXITOSAMENTE" -ForegroundColor Green
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "   Versión:    $NewVersion" -ForegroundColor White
    Write-Host "   XLL:        $OUTPUT_PATH\ConversorPDF.XLL-AddIn64-packed.xll" -ForegroundColor White
    Write-Host "   MSI:        $finalMsi" -ForegroundColor White
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Green

    exit 0
}
catch {
    Write-Error "ERROR CRÍTICO: $_"
    Write-Host "`nIniciando rollback automático..." -ForegroundColor Yellow
    Restore-VersionFiles
    Write-Host "`nBuild abortado. Version restaurada." -ForegroundColor Red
    exit 1
}
