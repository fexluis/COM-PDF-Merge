[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$Slice,
    [string]$OutputDir = '.agents/context',
    [int]$MaxLines = 400
)

# ---------------------------------------------------------------------------
# context-pack.ps1 - Genera un pack de contexto REAL de un slice.
# Localiza archivos por segmento de ruta y por namespace declarado, no por
# nombre de archivo (los servicios se llaman Ag*, Calculadora*, etc.).
# ---------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'
$repo     = Split-Path -Parent $PSScriptRoot
$academic = Join-Path $repo 'ConversorPDF.XLL\Academic'
$testsRoot = Join-Path $repo 'ConversorPDF.Tests'

if (-not (Test-Path $academic)) {
    Write-Host ('ERROR: no existe ' + $academic) -ForegroundColor Red
    exit 1
}

$outDir = Join-Path $repo $OutputDir
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
$outFile = Join-Path $outDir ('academic-' + $Slice.ToLower() + '.md')
$sliceRx = [regex]::Escape($Slice)

# --- 1. Archivos de produccion ---------------------------------------------
$allCs = @(Get-ChildItem -Path $academic -Recurse -File -Filter '*.cs' |
    Where-Object { $_.FullName -notmatch '\\(bin|obj)\\' })

$byPath = @($allCs | Where-Object { $_.FullName -match ('\\' + $sliceRx + '\\') })

$byNamespace = @()
if ($allCs.Count -gt 0) {
    $paths = @($allCs | ForEach-Object { $_.FullName })
    $hits = @(Select-String -Path $paths -Pattern ('namespace\s+[^\r\n]*\.' + $sliceRx + '\b') -List -ErrorAction SilentlyContinue)
    if ($hits.Count -gt 0) {
        $hitPaths = @($hits | ForEach-Object { $_.Path })
        $byNamespace = @($allCs | Where-Object { $hitPaths -contains $_.FullName })
    }
}

$prodFiles = @(@($byPath) + @($byNamespace) | Sort-Object FullName -Unique)

# --- 2. Tests ---------------------------------------------------------------
$allTests = @(Get-ChildItem -Path $testsRoot -Recurse -File -Filter '*.cs' -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -notmatch '\\(bin|obj|TestResults)\\' })
$testFiles = @($allTests | Where-Object {
    $_.FullName -match ('\\' + $sliceRx + '\\') -or $_.BaseName -match $sliceRx
} | Sort-Object FullName)

# --- 3. Contratos -----------------------------------------------------------
$interfaces = @()
if ($prodFiles.Count -gt 0) {
    $prodPaths = @($prodFiles | ForEach-Object { $_.FullName })
    $interfaces = @(Select-String -Path $prodPaths -Pattern 'public\s+interface\s+(\w+)' -AllMatches |
        ForEach-Object { $_.Matches } | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
}

$domainInterfaces = @()
$ifDir = Join-Path $academic 'Domain\Interfaces'
if ((Test-Path $ifDir) -and $prodFiles.Count -gt 0) {
    $corpus = ($prodFiles | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) -join "`n"
    $domainInterfaces = @(Get-ChildItem -Path $ifDir -File -Filter 'I*.cs' |
        Where-Object { $corpus -match ('\b' + [regex]::Escape($_.BaseName) + '\b') } |
        ForEach-Object { $_.BaseName } | Sort-Object -Unique)
}

# --- 4. Persistencia --------------------------------------------------------
$dbFiles = @()
if ($prodFiles.Count -gt 0) {
    $dbFiles = @($prodFiles | Where-Object {
        (Get-Content -LiteralPath $_.FullName -Raw) -match 'SQLite|Dapper|EncryptedSQLiteContext'
    })
}
$migrationsDir = Join-Path $academic 'Infrastructure\Persistence\Migrations'
$lastMigration = 0
if (Test-Path $migrationsDir) {
    $versions = @(Get-ChildItem -Path $migrationsDir -File -Filter 'MigrationV*.cs' |
        ForEach-Object { if ($_.BaseName -match 'MigrationV(\d+)') { [int]$Matches[1] } } |
        Sort-Object)
    if ($versions.Count -gt 0) { $lastMigration = $versions[-1] }
}

# --- 5. UI y entry points ---------------------------------------------------
$uiFiles = @($prodFiles | Where-Object { $_.FullName -match '\\UI\\' -or $_.FullName -match '\\Forms\\' })
$ribbonRefs = @()
$ribbon = Join-Path $repo 'ConversorPDF.XLL\Ribbon\RibbonController.cs'
if ((Test-Path $ribbon) -and $prodFiles.Count -gt 0) {
    $ribbonRefs = @($prodFiles | Where-Object {
        Select-String -LiteralPath $ribbon -Pattern ('\b' + [regex]::Escape($_.BaseName) + '\b') -Quiet
    } | ForEach-Object { $_.BaseName })
}

# --- 6. Gotchas detectadas (hechos, no prosa) -------------------------------
function Count-Hits($files, [string]$pattern) {
    if (-not $files -or $files.Count -eq 0) { return 0 }
    $paths = @($files | ForEach-Object { $_.FullName })
    $hits = @(Select-String -Path $paths -Pattern $pattern -AllMatches -ErrorAction SilentlyContinue)
    $sum = 0
    foreach ($h in $hits) { $sum += $h.Matches.Count }
    return $sum
}

$interopHits = Count-Hits $prodFiles 'Microsoft\.Office\.Interop|Excel\.Range|Excel\.Worksheet'
$winFormsHits = Count-Hits $prodFiles 'System\.Windows\.Forms|DataGridView'
$csharp12Hits = Count-Hits $prodFiles '=\s*\[\]|new\s*\[\]\s*\{'
$dateTimeNowHits = Count-Hits $prodFiles 'DateTime\.Now'
$blockingHits = Count-Hits $prodFiles '\.Result\b|\.Wait\(\)'

# --- 7. Render --------------------------------------------------------------
$lines = New-Object System.Collections.Generic.List[string]
$lines.Add('# Context Pack: Academic / ' + $Slice)
$lines.Add('')
$lines.Add('- Generado: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm'))
$lines.Add('- Cabecera obligatoria: ver `references/plan-contract.md` (Execution envelope).')
$lines.Add('- Presupuesto: <= ' + $MaxLines + ' lineas.')
$lines.Add('')
$lines.Add('## 1. Inventario del slice')
$lines.Add('- Archivos de produccion: ' + $prodFiles.Count)
$lines.Add('- Archivos de test: ' + $testFiles.Count)
$lines.Add('- Ultima migracion detectada: V' + $lastMigration)
$lines.Add('')
$lines.Add('## 2. Contratos')
if ($interfaces.Count -gt 0) {
    $lines.Add('Interfaces declaradas en el slice:')
    foreach ($i in $interfaces) { $lines.Add('- ' + $i) }
} else {
    $lines.Add('- Sin declaraciones `public interface` dentro del slice.')
}
if ($domainInterfaces.Count -gt 0) {
    $lines.Add('')
    $lines.Add('Interfaces de `Academic/Domain/Interfaces` referenciadas por el slice:')
    foreach ($i in $domainInterfaces) { $lines.Add('- ' + $i + ' (Academic/Domain/Interfaces/' + $i + '.cs)') }
}
$lines.Add('')
$lines.Add('## 3. Persistencia')
$lines.Add('- Archivos con SQLite/Dapper/EncryptedSQLiteContext: ' + $dbFiles.Count)
foreach ($d in ($dbFiles | Select-Object -First 10)) { $lines.Add('  - ' + $d.Name) }
$lines.Add('- Version de migracion vigente (maxima en disco): V' + $lastMigration)
$lines.Add('')
$lines.Add('## 4. UI y entry points')
$lines.Add('- Formularios/UI en el slice: ' + $uiFiles.Count)
foreach ($u in ($uiFiles | Select-Object -First 10)) { $lines.Add('  - ' + $u.Name) }
$lines.Add('- Referenciados desde RibbonController: ' + $ribbonRefs.Count)
foreach ($r in ($ribbonRefs | Select-Object -First 10)) { $lines.Add('  - ' + $r) }
$lines.Add('')
$lines.Add('## 5. Gotchas detectadas (conteo de patrones prohibidos)')
$lines.Add('- Interop Office en el slice: ' + $interopHits + ' (debe ser 0 fuera de XLL/Infrastructure)')
$lines.Add('- WinForms/DataGridView en el slice: ' + $winFormsHits + ' (0 esperado en Application/Domain)')
$lines.Add('- Sintaxis C# 12 (`[]`): ' + $csharp12Hits + ' (debe ser 0 en net48)')
$lines.Add('- DateTime.Now: ' + $dateTimeNowHits + ' (usar IDateTimeProvider)')
$lines.Add('- .Result/.Wait(): ' + $blockingHits + ' (prohibido salvo COM STA justificado)')
$lines.Add('')
$lines.Add('## 6. Archivos')
$lines.Add('### Produccion (' + $prodFiles.Count + ')')
foreach ($f in $prodFiles) { $lines.Add('- ' + $f.FullName.Replace($repo + '\', '')) }
$lines.Add('')
$lines.Add('### Tests (' + $testFiles.Count + ')')
foreach ($t in $testFiles) { $lines.Add('- ' + $t.FullName.Replace($repo + '\', '')) }

if ($lines.Count -gt $MaxLines) {
    $trimmed = $lines.GetRange(0, $MaxLines - 1)
    $trimmed.Add('... [truncado: ' + ($lines.Count - $MaxLines) + ' lineas omitidas]')
    $lines = $trimmed
}

$lines | Set-Content -LiteralPath $outFile -Encoding ASCII
Write-Host ('Context pack: ' + $outFile) -ForegroundColor Green
Write-Host ('Produccion: ' + $prodFiles.Count + ' | Tests: ' + $testFiles.Count + ' | Lineas: ' + $lines.Count) -ForegroundColor Green

if ($prodFiles.Count -eq 0) {
    Write-Host ('AVISO: el slice "' + $Slice + '" no tiene archivos. Verifica el nombre (segmento de ruta o namespace).') -ForegroundColor Yellow
    exit 2
}
exit 0
