[CmdletBinding()]
param(
    [ValidateSet('Build','Fast','Full')][string]$Mode = 'Fast',
    [string]$Filter,
    [string]$Configuration = 'Debug',
    [string]$Platform = 'x64',
    [switch]$NoRestore,
    [string]$EvidenceRoot = '_evidence'
)

# ---------------------------------------------------------------------------
# verify.ps1 - Unico gate de verificacion (Build | Fast | Full)
# Exit codes: 0 OK | 1 Build fallo | 2 Tests fallaron / infraestructura
#             3 Regresion vs baseline / suite reducida / sin baseline
#             4 Tooling o workspace invalido
# ---------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'

$repo     = Split-Path -Parent $PSScriptRoot
$xllProj  = Join-Path $repo 'ConversorPDF.XLL\ConversorPDF.XLL.csproj'
$testProj = Join-Path $repo 'ConversorPDF.Tests\ConversorPDF.Tests.csproj'
$baseline = Join-Path $repo 'tests\baseline\test_baseline.txt'

$runStart = Get-Date
$stamp    = $runStart.ToString('yyyyMMdd-HHmmss')
$evDir    = Join-Path $repo (Join-Path $EvidenceRoot $stamp)
New-Item -ItemType Directory -Force -Path $evDir | Out-Null
$buildLog = Join-Path $evDir 'build.log'
$testLog  = Join-Path $evDir 'test.log'
$trxPath  = Join-Path $evDir 'test.trx'
$summary  = Join-Path $evDir 'summary.md'

function Write-Step([string]$text, [string]$color = 'Cyan') {
    Write-Host $text -ForegroundColor $color
}

function Get-RepoHead([string]$root) {
    $gitCmd = Get-Command git -ErrorAction SilentlyContinue
    if ($gitCmd) {
        $sha = (& git -C $root rev-parse --short HEAD 2>$null)
        if ($LASTEXITCODE -eq 0 -and $sha) { return ([string]$sha).Trim() }
    }
    $gitDir   = Join-Path $root '.git'
    $headFile = Join-Path $gitDir 'HEAD'
    if (-not (Test-Path $headFile)) { return 'no-git' }
    $head = (Get-Content -LiteralPath $headFile -Raw).Trim()
    if ($head -notmatch '^ref:\s*(.+)$') { return $head }
    $ref     = $Matches[1]
    $refFile = Join-Path $gitDir ($ref -replace '/', '\')
    if (Test-Path $refFile) {
        $sha = (Get-Content -LiteralPath $refFile -Raw).Trim()
        return ('{0} ({1})' -f $sha, $ref)
    }
    $packed = Join-Path $gitDir 'packed-refs'
    if (Test-Path $packed) {
        $hit = Select-String -LiteralPath $packed -Pattern ([regex]::Escape($ref) + '$') | Select-Object -First 1
        if ($hit) { return ('{0} ({1} packed)' -f $hit.Line.Split(' ')[0], $ref) }
    }
    return ($ref + ' (unresolved)')
}

function Get-RepoDirty([string]$root) {
    $gitCmd = Get-Command git -ErrorAction SilentlyContinue
    if (-not $gitCmd) { return 'unknown' }
    $lines = @(& git -C $root status --porcelain 2>$null)
    if ($LASTEXITCODE -ne 0) { return 'unknown' }
    return $lines.Count
}

function Read-Baseline([string]$path) {
    $result = [pscustomobject]@{ Total = 0; Failures = @() }
    if (-not (Test-Path $path)) { return $result }
    foreach ($line in @(Get-Content -LiteralPath $path)) {
        if ($line -match '^\s*TOTAL\s*:\s*(\d+)') { $result.Total = [int]$Matches[1] }
        if ($line -like 'FAILED *') { $result.Failures += $line.Substring(7).Trim() }
    }
    return $result
}

# --- Precondiciones ---------------------------------------------------------
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    Write-Step 'ERROR: dotnet no encontrado en PATH' 'Red'
    exit 4
}
if (-not (Test-Path $xllProj) -or -not (Test-Path $testProj)) {
    Write-Step 'ERROR: workspace invalido (faltan ConversorPDF.XLL.csproj o ConversorPDF.Tests.csproj)' 'Red'
    exit 4
}
if ($repo -match 'OneDrive') {
    Write-Step 'AVISO: repo dentro de OneDrive. Riesgo de locks con MSBuild/VSTest.' 'Yellow'
}

$head       = Get-RepoHead $repo
$dirtyCount = Get-RepoDirty $repo

# --- 1. Build XLL -----------------------------------------------------------
Write-Step "[1/3] Build XLL ($Configuration/$Platform)" 'Cyan'
$buildXll = @('build', $xllProj, '-c', $Configuration, "-p:Platform=$Platform", '-v:m')
if ($NoRestore) { $buildXll += '--no-restore' }

$ErrorActionPreference = 'Continue'
$outXll = & dotnet @buildXll 2>&1
$xllExit = $LASTEXITCODE
$ErrorActionPreference = 'Stop'
$outXll | Out-File -FilePath $buildLog -Encoding ASCII
$outXll | ForEach-Object { Write-Host $_ }

if ($xllExit -ne 0) {
    Write-Step 'BUILD_FAILED: la compilacion del XLL fallo (ver build.log)' 'Red'
    exit 1
}

# --- 2. Build tests ---------------------------------------------------------
Write-Step "[2/3] Build tests ($Configuration/$Platform)" 'Cyan'
$buildTests = @('build', $testProj, '-c', $Configuration, "-p:Platform=$Platform", '--no-dependencies', '-v:m')
if ($NoRestore) { $buildTests += '--no-restore' }

$ErrorActionPreference = 'Continue'
$outTests = & dotnet @buildTests 2>&1
$testsExit = $LASTEXITCODE
$ErrorActionPreference = 'Stop'
$outTests | Out-File -FilePath $buildLog -Append -Encoding ASCII
$outTests | ForEach-Object { Write-Host $_ }

if ($testsExit -ne 0) {
    Write-Step 'BUILD_FAILED: la compilacion del proyecto de tests fallo (ver build.log)' 'Red'
    exit 1
}

if ($Mode -eq 'Build') {
    $buildSummary = @(
        '# Verificacion Build - ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),
        '',
        '- Configuracion: ' + $Configuration + '/' + $Platform,
        '- HEAD: ' + $head,
        '- Working tree: ' + $dirtyCount + ' archivo(s) con cambios',
        '- Build XLL: OK',
        '- Build tests: OK',
        '- Veredicto: PASS',
        '- Evidencia: ' + $evDir
    )
    $buildSummary | Set-Content -LiteralPath $summary -Encoding ASCII
    Write-Step ('Build OK. Evidencia: ' + $evDir) 'Green'
    exit 0
}

# --- 3. Tests ---------------------------------------------------------------
Write-Step "[3/3] Tests ($Mode)" 'Cyan'
$testArgs = @(
    'test', $testProj,
    '-c', $Configuration,
    "-p:Platform=$Platform",
    '--no-build',
    '--logger', 'trx;LogFileName=test.trx',
    '--results-directory', $evDir
)
if ($Filter) { $testArgs += @('--filter', $Filter) }

$ErrorActionPreference = 'Continue'
$outTest = & dotnet @testArgs 2>&1
$testExit = $LASTEXITCODE
$ErrorActionPreference = 'Stop'
$outTest | Out-File -FilePath $testLog -Encoding ASCII
$outTest | ForEach-Object { Write-Host $_ }

if (-not (Test-Path $trxPath)) {
    Write-Step ('INFRA_ERROR: no se genero TRX en ' + $trxPath) 'Red'
    exit 2
}
if ((Get-Item -LiteralPath $trxPath).LastWriteTime -lt $runStart.AddSeconds(-2)) {
    Write-Step 'INFRA_ERROR: el TRX es anterior a esta ejecucion (stale)' 'Red'
    exit 2
}

[xml]$doc = Get-Content -LiteralPath $trxPath
$results = @($doc.SelectNodes("//*[local-name()='UnitTestResult']"))
$failedNames = @($results | Where-Object { $_.outcome -eq 'Failed' } | ForEach-Object { [string]$_.testName })
$passedCount = @($results | Where-Object { $_.outcome -eq 'Passed' }).Count
$skippedCount = @($results | Where-Object { $_.outcome -eq 'NotExecuted' }).Count
$total = $results.Count

$base = Read-Baseline $baseline
$regressions = @($failedNames | Where-Object { $base.Failures -notcontains $_ })
$suiteShrunk = ($base.Total -gt 0 -and $total -lt $base.Total)

# --- Veredicto --------------------------------------------------------------
$verdict = 'PASS'
$exitCode = 0
$notes = @()

if ($total -eq 0) {
    $verdict = 'INFRA_ERROR'
    $exitCode = 2
    $notes += 'El TRX no contiene resultados ejecutados: el test host no corrio la suite.'
}
elseif ($Mode -eq 'Full' -and -not (Test-Path $baseline)) {
    $verdict = 'NO_BASELINE'
    $exitCode = 3
    $notes += 'Full exige baseline: sin tests/baseline/test_baseline.txt no se puede certificar nada.'
}
elseif ($Mode -eq 'Full' -and $suiteShrunk) {
    $verdict = 'SUITE_SHRUNK'
    $exitCode = 3
    $notes += ('La suite se redujo: {0} pruebas ejecutadas vs {1} del baseline. Ejecucion no certificable.' -f $total, $base.Total)
}
elseif ($regressions.Count -gt 0) {
    $verdict = 'REGRESSION'
    $exitCode = 3
}
elseif ($testExit -ne 0 -and $failedNames.Count -eq 0) {
    $verdict = 'INFRA_ERROR'
    $exitCode = 2
    $notes += 'dotnet test fallo sin pruebas marcadas como Failed: fallo de infraestructura.'
}
elseif ($Mode -eq 'Fast' -and $testExit -ne 0) {
    $verdict = 'RED'
    $exitCode = 2
}

# --- Resumen ----------------------------------------------------------------
$summaryLines = New-Object System.Collections.Generic.List[string]
$summaryLines.Add('# Verificacion ' + $Mode + ' - ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
$summaryLines.Add('')
$summaryLines.Add('- Modo: ' + $Mode)
$summaryLines.Add('- Configuracion: ' + $Configuration + '/' + $Platform)
$summaryLines.Add('- Filtro: ' + $(if ($Filter) { $Filter } else { '<todos>' }))
$summaryLines.Add('- HEAD: ' + $head)
$summaryLines.Add('- Working tree: ' + $dirtyCount + ' archivo(s) con cambios')
$summaryLines.Add('- Total: ' + $total + ' | Passed: ' + $passedCount + ' | Failed: ' + $failedNames.Count + ' | Skipped: ' + $skippedCount)
$summaryLines.Add('- Baseline: total ' + $base.Total + ' | fallos conocidos ' + $base.Failures.Count)
$summaryLines.Add('- Suite reducida: ' + $(if ($suiteShrunk) { 'SI' } else { 'NO' }))
$summaryLines.Add('- Regresiones: ' + $regressions.Count)
$summaryLines.Add('- Veredicto: ' + $verdict)
$summaryLines.Add('- Evidencia: ' + $evDir)
$summaryLines.Add('')
$summaryLines.Add('## Fallos (' + $failedNames.Count + ')')
if ($failedNames.Count -gt 0) { foreach ($f in $failedNames) { $summaryLines.Add('- ' + $f) } }
else { $summaryLines.Add('- ninguno') }
$summaryLines.Add('')
$summaryLines.Add('## Regresiones no presentes en baseline (' + $regressions.Count + ')')
if ($regressions.Count -gt 0) { foreach ($r in $regressions) { $summaryLines.Add('- ' + $r) } }
else { $summaryLines.Add('- ninguna') }
if ($notes.Count -gt 0) {
    $summaryLines.Add('')
    $summaryLines.Add('## Notas')
    foreach ($n in $notes) { $summaryLines.Add('- ' + $n) }
}

$summaryLines | Set-Content -LiteralPath $summary -Encoding ASCII

$color = switch ($verdict) {
    'PASS'  { 'Green' }
    'RED'   { 'Yellow' }
    default { 'Red' }
}
Write-Step ('VEREDICTO: ' + $verdict) $color
Write-Step ('Evidencia: ' + $evDir) 'Green'

exit $exitCode
