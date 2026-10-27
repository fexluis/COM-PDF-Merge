[CmdletBinding()]
param(
    [string]$Workspace = '',
    [string]$Plan,
    [int]$MinXllCs = 300,
    [int]$MinTestsCs = 100,
    [switch]$AllowOneDrive
)

# ---------------------------------------------------------------------------
# preflight.ps1 - Valida que el workspace sea una copia REAL del repositorio
# antes de planificar o ejecutar un swarm.
# Exit codes: 0 OK | 1 workspace incompleto | 2 prerequisitos faltantes
#             3 workspace en OneDrive (bloqueado)
# ---------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'
if (-not $Workspace) { $Workspace = Split-Path -Parent $PSScriptRoot }
$Workspace = (Resolve-Path -LiteralPath $Workspace).Path

$checks = New-Object System.Collections.Generic.List[object]
function Add-Check([string]$category, [string]$name, [string]$status, [string]$detail) {
    $checks.Add([pscustomobject]@{ Category = $category; Check = $name; Status = $status; Detail = $detail })
}

function Count-Source([string]$root, [string]$sub) {
    $path = Join-Path $root $sub
    if (-not (Test-Path $path)) { return -1 }
    $items = @(Get-ChildItem -Path $path -Recurse -File -Filter '*.cs' |
        Where-Object { $_.FullName -notmatch '\\(bin|obj|TestResults)\\' })
    return $items.Count
}

$xllCs   = Count-Source $Workspace 'ConversorPDF.XLL'
$testsCs = Count-Source $Workspace 'ConversorPDF.Tests'

# --- Workspace: tamano real del arbol ---------------------------------------
if ($xllCs -lt 0) {
    Add-Check 'WORKSPACE' 'ConversorPDF.XLL presente' 'FAIL' 'No existe la carpeta ConversorPDF.XLL'
}
elseif ($xllCs -lt $MinXllCs) {
    Add-Check 'WORKSPACE' 'Archivos .cs en XLL' 'FAIL' ("{0} archivos (minimo {1}): el workspace no es una copia del arbol real" -f $xllCs, $MinXllCs)
}
else {
    Add-Check 'WORKSPACE' 'Archivos .cs en XLL' 'OK' ("{0} archivos" -f $xllCs)
}

if ($testsCs -lt 0) {
    Add-Check 'WORKSPACE' 'ConversorPDF.Tests presente' 'FAIL' 'No existe la carpeta ConversorPDF.Tests'
}
elseif ($testsCs -lt $MinTestsCs) {
    Add-Check 'WORKSPACE' 'Archivos .cs en Tests' 'FAIL' ("{0} archivos (minimo {1}): suite incompleta, baseline no representativo" -f $testsCs, $MinTestsCs)
}
else {
    Add-Check 'WORKSPACE' 'Archivos .cs en Tests' 'OK' ("{0} archivos" -f $testsCs)
}

$academic = Join-Path $Workspace 'ConversorPDF.XLL\Academic'
if (Test-Path $academic) {
    Add-Check 'WORKSPACE' 'Slice Academic presente' 'OK' $academic
}
else {
    Add-Check 'WORKSPACE' 'Slice Academic presente' 'FAIL' 'Falta ConversorPDF.XLL\Academic (snapshot previo a la modularizacion)'
}

# --- OneDrive ---------------------------------------------------------------
if ($Workspace -match 'OneDrive') {
    if ($AllowOneDrive) {
        Add-Check 'ENV' 'Ubicacion del workspace' 'WARN' 'Dentro de OneDrive: riesgo de locks (permitido por -AllowOneDrive)'
    }
    else {
        Add-Check 'ENV' 'Ubicacion del workspace' 'FAIL' 'Dentro de OneDrive: usa una copia local fuera de OneDrive'
    }
}
else {
    Add-Check 'ENV' 'Ubicacion del workspace' 'OK' $Workspace
}

# --- Prerequisitos del gate y de los agentes --------------------------------
$requiredFiles = @(
    @{ Path = 'scripts\verify.ps1';                    Label = 'scripts/verify.ps1' },
    @{ Path = 'tests\baseline\test_baseline.txt';      Label = 'tests/baseline/test_baseline.txt' },
    @{ Path = '.agents\rules\parallel-work.md';        Label = '.agents/rules/parallel-work.md' },
    @{ Path = '.agents\rules\verification-commands.md'; Label = '.agents/rules/verification-commands.md' },
    @{ Path = '.agents\agents\plan-reviewer.md';       Label = '.agents/agents/plan-reviewer.md' },
    @{ Path = '.agents\agents\tdd-red-writer.md';      Label = '.agents/agents/tdd-red-writer.md' },
    @{ Path = '.agents\agents\implementer.md';         Label = '.agents/agents/implementer.md' },
    @{ Path = '.agents\agents\integrator.md';          Label = '.agents/agents/integrator.md' },
    @{ Path = '.agents\agents\task-reviewer.md';       Label = '.agents/agents/task-reviewer.md' },
    @{ Path = '.agents\agents\build-verifier.md';      Label = '.agents/agents/build-verifier.md' }
)
$missing = @()
foreach ($item in $requiredFiles) {
    if (-not (Test-Path (Join-Path $Workspace $item.Path))) { $missing += $item.Label }
}
if ($missing.Count -gt 0) {
    Add-Check 'PREREQ' 'Artefactos obligatorios' 'FAIL' ('Faltan: ' + ($missing -join ', '))
}
else {
    Add-Check 'PREREQ' 'Artefactos obligatorios' 'OK' 'verify.ps1, baseline, reglas y 6 agentes presentes'
}

$baselinePath = Join-Path $Workspace 'tests\baseline\test_baseline.txt'
$baselineTotal = 0
if (Test-Path $baselinePath) {
    foreach ($line in @(Get-Content -LiteralPath $baselinePath)) {
        if ($line -match '^\s*TOTAL\s*:\s*(\d+)') { $baselineTotal = [int]$Matches[1] }
    }
    if ($baselineTotal -eq 0) {
        Add-Check 'PREREQ' 'Baseline con TOTAL' 'FAIL' 'test_baseline.txt no declara la linea TOTAL: no se puede detectar suite reducida'
    }
    elseif ($baselineTotal -lt $MinTestsCs) {
        Add-Check 'PREREQ' 'Coherencia baseline vs arbol' 'FAIL' ("baseline TOTAL={0} pero el arbol tiene {1} archivos de test: recalibrar" -f $baselineTotal, $testsCs)
    }
    else {
        Add-Check 'PREREQ' 'Baseline con TOTAL' 'OK' ("TOTAL={0}" -f $baselineTotal)
    }
}

# --- Plan (opcional) --------------------------------------------------------
if ($Plan) {
    $planPath = if (Test-Path $Plan) { (Resolve-Path -LiteralPath $Plan).Path } else { Join-Path $Workspace $Plan }
    if (-not (Test-Path $planPath)) {
        Add-Check 'PLAN' 'Plan indicado' 'FAIL' ('No existe: ' + $Plan)
    }
    else {
        $text = Get-Content -LiteralPath $planPath -Raw
        $planIssues = @()
        if ($text -notmatch '(?m)^##\s*Execution envelope') { $planIssues += 'sin Execution envelope' }
        if ($text -notmatch '(?m)^\s*-\s*\*\*files_lock\*\*') { $planIssues += 'tareas sin files_lock' }
        if ($text -notmatch '(?m)^\s*-\s*no_push:\s*true') { $planIssues += 'sin no_push: true' }
        if ($text -match 'total:\s*51\b') { $planIssues += 'baseline hardcodeado (51) sin recalibrar' }
        if ($planIssues.Count -gt 0) {
            Add-Check 'PLAN' 'Contrato del plan' 'FAIL' ($planIssues -join '; ')
        }
        else {
            Add-Check 'PLAN' 'Contrato del plan' 'OK' (Split-Path -Leaf $planPath)
        }
    }
}

# --- Salida -----------------------------------------------------------------
Write-Host ''
Write-Host ('PREFLIGHT WORKSPACE: ' + $Workspace) -ForegroundColor Cyan
$checks | Format-Table -AutoSize -Wrap Category, Status, Check, Detail | Out-String -Width 200 | Write-Host

$oneDriveFail = @($checks | Where-Object { $_.Category -eq 'ENV' -and $_.Status -eq 'FAIL' }).Count
$workspaceFail = @($checks | Where-Object { $_.Category -eq 'WORKSPACE' -and $_.Status -eq 'FAIL' }).Count
$prereqFail = @($checks | Where-Object { $_.Category -in @('PREREQ','PLAN') -and $_.Status -eq 'FAIL' }).Count

if ($oneDriveFail -gt 0) {
    Write-Host 'PREFLIGHT_FAILED: workspace en OneDrive (usa copia local).' -ForegroundColor Red
    exit 3
}
if ($workspaceFail -gt 0) {
    Write-Host 'PREFLIGHT_FAILED: workspace incompleto (no es una copia del arbol real).' -ForegroundColor Red
    exit 1
}
if ($prereqFail -gt 0) {
    Write-Host 'PREFLIGHT_FAILED: faltan prerequisitos del gate.' -ForegroundColor Red
    exit 2
}

Write-Host 'PREFLIGHT_OK' -ForegroundColor Green
exit 0
