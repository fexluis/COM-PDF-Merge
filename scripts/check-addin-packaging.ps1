param(
    [string]$Configuration = 'Release',
    [ValidateSet('x64')][string]$Platform = 'x64'
)

# Contrato de empaquetado del add-in: las cuatro fuentes de verdad deben coincidir.
#   1. Referencias declaradas en el .dna  -> deben existir en bin (o Excel-DNA fallara al cargarlas).
#   2. Archivos del instalador WiX        -> su Source debe existir y todo DLL de bin debe estar cubierto.
#   3. Salida de compilacion (bin)        -> es lo unico que realmente se puede desplegar.
# Falla (exit 3) ante cualquier desajuste; exit 2 si aun no hay binario compilado.
# Uso: powershell -ExecutionPolicy Bypass -File scripts/check-addin-packaging.ps1 [-Configuration Release]

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$bin = Join-Path $root "ConversorPDF.XLL\bin\$Platform\$Configuration"
$dnaPath = Join-Path $root 'ConversorPDF.XLL\ConversorPDF.XLL-AddIn64.dna'
$wxsPath = Join-Path $root 'installer\InstaladorConversorPDF.wxs'

foreach ($required in @($dnaPath, $wxsPath)) {
    if (-not (Test-Path $required)) {
        Write-Host "EMPACADO_ERROR: no existe $required" -ForegroundColor Red
        exit 2
    }
}
if (-not (Test-Path $bin)) {
    Write-Host "EMPACADO_ERROR: no hay binario compilado en $bin (compile primero)" -ForegroundColor Red
    exit 2
}

$binDlls = @(Get-ChildItem $bin -File -Filter *.dll | Select-Object -ExpandProperty Name | Sort-Object -Unique)
$dnaText = Get-Content $dnaPath -Raw
$wxsText = Get-Content $wxsPath -Raw

$dnaRefs = @([regex]::Matches($dnaText, 'Path="([^"]+\.dll)"') |
    ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
$wxsNames = @([regex]::Matches($wxsText, 'Name="([^"]+\.dll)"') |
    ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
$wxsSources = @([regex]::Matches($wxsText, 'Source="([^"]+)"') |
    ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)

$errors = New-Object System.Collections.Generic.List[string]

# 1. Toda referencia del .dna debe existir en bin.
foreach ($ref in $dnaRefs) {
    if (-not (Test-Path (Join-Path $bin $ref))) {
        $errors.Add("DNA_REFERENCIA_INEXISTENTE: '$ref' declarada en el .dna y ausente en $Configuration|$Platform")
    }
}

# 2. Todo archivo del instalador debe tener Source valido (relativo a installer\).
foreach ($src in $wxsSources) {
    $resolved = Join-Path (Join-Path $root 'installer') $src
    if (-not (Test-Path $resolved)) {
        $errors.Add("WXS_SOURCE_INEXISTENTE: '$src' no existe")
    }
}

# 3. Todo DLL compilado debe declararse en el .dna o instalarse por WiX.
foreach ($dll in $binDlls) {
    $enDna = $dnaRefs -contains $dll
    $enWxs = $wxsNames -contains $dll
    if (-not $enDna -and -not $enWxs) {
        $errors.Add("DLL_NO_DESPLEGADA: '$dll' esta en bin pero no en el .dna ni en el instalador")
    }
}

# 4. Interop nativo: el preloader lo busca en x64\ y el instalador debe copiarlo.
if ((Test-Path (Join-Path $bin 'x64\SQLite.Interop.dll')) -and -not ($wxsNames -contains 'SQLite.Interop.dll')) {
    $errors.Add("INTEROP_NO_INSTALADO: 'x64\SQLite.Interop.dll' existe en bin y el instalador no lo declara")
}

Write-Host ''
Write-Host "Empaquetado del add-in ($Configuration|$Platform)" -ForegroundColor Cyan
Write-Host ("  bin: {0} DLL(s) | .dna: {1} referencia(s) | WiX: {2} DLL(s)" -f $binDlls.Count, $dnaRefs.Count, $wxsNames.Count)
Write-Host ''

if ($errors.Count -eq 0) {
    Write-Host 'EMPACADO_OK: .dna, instalador y binario coinciden.' -ForegroundColor Green
    exit 0
}

foreach ($e in $errors) { Write-Host "  $e" -ForegroundColor Red }
Write-Host ''
Write-Host ("EMPACADO_ERROR: {0} desajuste(s). El add-in instalado fallaria en tiempo de ejecucion." -f $errors.Count) -ForegroundColor Red
exit 3
