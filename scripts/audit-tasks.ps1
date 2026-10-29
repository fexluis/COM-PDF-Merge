[CmdletBinding()]
param(
    [string]$Plan,
    [string[]]$Tasks,
    [string[]]$Paths,
    [string]$EvidenceRoot = '_evidence'
)

# ---------------------------------------------------------------------------
# audit-tasks.ps1 - Auditoria de codigo por tarea (Ronda.md: Fase <-> Task T<n>)
# Complementa verify.ps1 (compila + tests) con detectores deterministas de
# errores de ejecucion: bloqueos STA, Interop en capas limpias, SQL destructivo
# en tablas CRDT, migraciones duplicadas, tests debilitados, etc.
# Exit codes: 0 sin hallazgos Critical | 3 hay hallazgos Critical (FIX_PLAN)
#             4 plan/entrada invalida (nada que auditar)
# ---------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot

function Rel([string]$full) {
    $r = $full.Replace($repo + '\', '')
    return ($r -replace '\\', '/')
}

function Normalize-Path([string]$p) {
    $t = $p.Trim()
    $t = $t -replace '`', ''
    if ($t -match '^(.+?)\s+\(') { $t = $Matches[1] }
    $t = $t.Trim().Trim(',', ';', '"', "'")
    $t = $t -replace '/', '\'
    return $t.Trim()
}

# --- 1. Localizar el plan ------------------------------------------------
$plansDir = Join-Path $repo '.agents\plans'
if (-not $Plan) {
    $candidates = @(Get-ChildItem -LiteralPath $plansDir -Filter '*-plan.md' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending)
    if ($candidates.Count -eq 0) {
        Write-Host 'AUDIT-ERROR: no hay planes en .agents/plans' -ForegroundColor Red
        exit 4
    }
    $Plan = $candidates[0].FullName
}
elseif (-not (Test-Path -LiteralPath $Plan)) {
    $alt = Join-Path $repo $Plan
    if (Test-Path -LiteralPath $alt) { $Plan = $alt }
    else {
        Write-Host "AUDIT-ERROR: plan no encontrado: $Plan" -ForegroundColor Red
        exit 4
    }
}
$PlanFull = (Resolve-Path -LiteralPath $Plan).Path

# --- 2. Parsear tareas del plan (### T<n>: <nombre>) ---------------------
$lines = Get-Content -LiteralPath $PlanFull
$parsed = New-Object System.Collections.Generic.List[object]
$current = $null
$mode = ''

function Add-Files([object]$task, [string]$text) {
    if (-not $text) { return }
    $candidates = @()
    if ($text -match '`') {
        $candidates = @([regex]::Matches($text, '`([^`]+)`') | ForEach-Object { $_.Groups[1].Value })
    }
    else {
        $inner = $text
        if ($inner -match '\[(.*)\]') { $inner = $Matches[1] }
        $candidates = @($inner -split ',')
    }
    foreach ($c in $candidates) {
        $n = Normalize-Path $c
        if (-not $n) { continue }
        if ($n -notmatch '(?i)\.(cs|csproj|sql|md|config|json|xaml|ps1|txt|resx)$') { continue }
        [void]$task.Files.Add($n)
    }
}

for ($i = 0; $i -lt $lines.Count; $i++) {
    $line = $lines[$i]

    if ($line -match '^\s{0,3}#{2,4}\s*(T\d+)\s*[:.\-]?\s*(.*)$') {
        if ($current) { $parsed.Add($current) }
        $current = [pscustomobject]@{
            Id    = $Matches[1]
            Title = ($Matches[2] -replace '^\s*', '')
            Fase  = ''
            Files = New-Object System.Collections.Generic.List[string]
        }
        $mode = ''
        continue
    }
    if (-not $current) { continue }

    if ($line -match '^\s*-\s+\*\*(files_lock|files edited/created|fase)\*\*\s*:\s*(.*)$') {
        $field = $Matches[1].ToLowerInvariant()
        $rest = $Matches[2].Trim()
        if ($field -eq 'fase') {
            $current.Fase = ($rest -replace '[^0-9]', '')
            $mode = ''
        }
        else {
            $mode = 'files'
            if ($rest) { Add-Files $current $rest }
        }
        continue
    }
    if ($line -match '^\s*-\s+\*\*[^*]+\*\*\s*:') { $mode = ''; continue }
    if ($mode -eq 'files' -and $line -match '^\s{2,}-\s+(.+)$') {
        Add-Files $current $Matches[1]
        continue
    }
    if ($line -match '^\s*$') { $mode = '' }
}
if ($current) { $parsed.Add($current) }

$selected = @($parsed | Where-Object { $_.Files.Count -gt 0 })
if ($Tasks -and $Tasks.Count -gt 0) {
    $wanted = @($Tasks | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim().ToUpperInvariant() } | Where-Object { $_ })
    $selected = @($selected | Where-Object { $wanted -contains $_.Id.ToUpperInvariant() })
}

if ($Paths -and $Paths.Count -gt 0) {
    $unit = [pscustomobject]@{
        Id    = 'ADHOC'
        Title = 'Rutas suministradas por -Paths'
        Fase  = ''
        Files = New-Object System.Collections.Generic.List[string]
    }
    foreach ($p in $Paths) {
        foreach ($tok in ($p -split ',')) {
            $n = Normalize-Path $tok
            if ($n) { [void]$unit.Files.Add($n) }
        }
    }
    $selected = @($unit)
}

if ($selected.Count -eq 0) {
    Write-Host 'AUDIT-ERROR: el plan no declara tareas con files_lock ni se paso -Paths' -ForegroundColor Red
    exit 4
}

# --- 3. Catalogo de detectores ------------------------------------------
# Kind: cs (todo .cs de la tarea) | clean (Application/Domain) |
#       tests (ConversorPDF.Tests) | any (cualquier archivo de texto)
$detectors = @(
    @{ Id = 'AR-01'; Sev = 'Compilation'; Kind = 'clean'; Pattern = 'Excel\.|Microsoft\.Office|DataGridView|System\.Windows\.Forms|System\.Data|PdfSharp|System\.Drawing'; Msg = 'Interop/UI/Infra en capa limpia: exponer interfaz y mover al XLL' }
    @{ Id = 'CA-01'; Sev = 'Compilation'; Kind = 'cs'; Pattern = '=\s*\[\s*\]\s*;'; Msg = 'Collection expression C# 12: prohibida en net48' }
    @{ Id = 'CA-02'; Sev = 'Warning'; Kind = 'cs'; Pattern = '^\s*(public|internal)\s+(sealed\s+)?class\s+\w+\s*\([^)]*\)\s*$'; Msg = 'Posible primary constructor (C# 12): revisar' }
    @{ Id = 'RT-01'; Sev = 'Runtime'; Kind = 'cs'; Pattern = '(?<!Result)\.Result\b(?!\s*\.)'; Msg = 'Bloqueo sincrono sobre Task: deadlock en hilo STA' }
    @{ Id = 'RT-02'; Sev = 'Runtime'; Kind = 'cs'; Pattern = '\.Wait\(\s*\)'; Msg = 'Bloqueo sincrono sobre Task: deadlock en hilo STA' }
    @{ Id = 'RT-03'; Sev = 'Runtime'; Kind = 'cs'; Pattern = '\basync\s+void\b'; Msg = 'async void: excepcion no observable, riesgo de cuelgue' }
    @{ Id = 'RT-04'; Sev = 'Runtime'; Kind = 'cs'; Pattern = 'catch\s*(\(\s*\w*(Exception)?\s*\w*\s*\))?\s*\{\s*\}'; Msg = 'catch vacio: excepcion tragada, estado inconsistente' }
    @{ Id = 'RT-05'; Sev = 'Warning'; Kind = 'cs'; Pattern = '\bnew\s+SqliteConnection\b'; Msg = 'SqliteConnection: verificar using/Dispose (lock WAL)' }
    @{ Id = 'RT-06'; Sev = 'Runtime'; Kind = 'cs'; Pattern = 'Application\.DoEvents|Thread\.Sleep\s*\('; Msg = 'Bombeo de mensajes/espera bloqueante: reentrancy' }
    @{ Id = 'RT-07'; Sev = 'Warning'; Kind = 'cs'; Pattern = 'ScreenUpdating\s*=\s*true|DisplayAlerts\s*=\s*true|EnableEvents\s*=\s*true'; Msg = 'Restauracion de estado COM: confirmar que vive en finally' }
    @{ Id = 'RT-08'; Sev = 'Warning'; Kind = 'cs'; Pattern = 'QueryMultiple\s*\('; Msg = 'Dapper QueryMultiple: consumir/dispose completo el GridReader' }
    @{ Id = 'LG-01'; Sev = 'Logic'; Kind = 'cs'; Pattern = 'DateTime\.(Now|Today)\b'; Msg = 'DateTime.Now/Today: usar IDateTimeProvider' }
    @{ Id = 'LG-02'; Sev = 'Warning'; Kind = 'cs'; Pattern = 'DateTime\.UtcNow\b'; Msg = 'DateTime.UtcNow directo: revisar si debe pasar por IDateTimeProvider' }
    @{ Id = 'LG-03'; Sev = 'Logic'; Kind = 'cs'; Pattern = 'SELECT\s+\*\s+FROM'; Msg = 'SELECT *: desincronizacion posicional al anadir columnas' }
    @{ Id = 'LG-04'; Sev = 'Logic'; Kind = 'cs'; Pattern = 'INSERT\s+INTO\s+\w+\s+VALUES'; Msg = 'INSERT sin lista de columnas: orden posicional fragil' }
    @{ Id = 'LG-05'; Sev = 'Runtime'; Kind = 'cs'; Pattern = '\.(First|Single)\(\s*\)'; Msg = 'First()/Single() lanza con secuencia vacia: usar FirstOrDefault + guarda' }
    @{ Id = 'LN-01'; Sev = 'Runtime'; Kind = 'clean'; Pattern = 'MessageBox\.Show'; Msg = 'MessageBox en capa limpia: acopla WinForms y bloquea tests' }
    @{ Id = 'TH-01'; Sev = 'Logic'; Kind = 'tests'; Pattern = '(\[Ignore\b|\bIgnore\s*[,\]])|\[Skip\b|Skip\s*=|Assert\.Inconclusive'; Msg = 'Test debilitado/saltado: prohibido por verification-commands' }
    @{ Id = 'TH-02'; Sev = 'Warning'; Kind = 'tests'; Pattern = 'Assert\.(True|IsTrue)\(\s*true\s*\)'; Msg = 'Assert tautologico: la prueba no es falsable' }
    @{ Id = 'WR-01'; Sev = 'Warning'; Kind = 'cs'; Pattern = 'Console\.Write'; Msg = 'Console.Write en add-in: usar el logger de App' }
)

function Test-Kind([string]$kind, [string]$rel) {
    switch ($kind) {
        'cs' { return ($rel -match '(?i)\.cs$') }
        'clean' {
            return (($rel -match '(?i)\.cs$') -and ($rel -match '(?i)/(Application|Domain)/') -and
                    ($rel -notmatch '(?i)ConversorPDF\.Tests'))
        }
        'tests' { return (($rel -match '(?i)\.cs$') -and ($rel -match '(?i)ConversorPDF\.Tests')) }
        default { return $true }
    }
}

# --- 4. Ejecucion de detectores -----------------------------------------
$stamp = (Get-Date).ToString('yyyyMMdd-HHmmss-fff')
$evDir = Join-Path $repo (Join-Path $EvidenceRoot ($stamp + '-audit'))
New-Item -ItemType Directory -Force -Path $evDir | Out-Null
$reportPath = Join-Path $evDir 'audit.md'

$findings = New-Object System.Collections.Generic.List[object]
$audited = New-Object System.Collections.Generic.List[string]
$blocking = @('Compilation', 'Runtime', 'Logic')

foreach ($task in $selected) {
    $seen = @{}
    foreach ($raw in $task.Files) {
        $rel = ($raw -replace '\\', '/')
        if ($seen.ContainsKey($rel)) { continue }
        $seen[$rel] = $true
        $full = Join-Path $repo ($rel -replace '/', '\')

        if (-not (Test-Path -LiteralPath $full)) {
            $findings.Add([pscustomobject]@{
                    Task = $task.Id; Fase = $task.Fase; Sev = 'Compilation'; Detector = 'FS-01'
                    File = $rel; Line = 0; Snippet = ''; Msg = 'Archivo declarado en el plan no existe en disco'
                })
            continue
        }
        if ($rel -notmatch '(?i)\.(cs|sql)$') { continue }
        $audited.Add($rel)

        $content = Get-Content -LiteralPath $full -ErrorAction SilentlyContinue
        for ($ln = 0; $ln -lt $content.Count; $ln++) {
            $text = [string]$content[$ln]
            $trim = $text.Trim()
            if ($trim -eq '' -or $trim.StartsWith('//') -or $trim.StartsWith('///') -or
                $trim.StartsWith('*') -or $trim.StartsWith('/*')) { continue }
            foreach ($d in $detectors) {
                if (-not (Test-Kind $d.Kind $rel)) { continue }
                if ([regex]::IsMatch($text, $d.Pattern)) {
                    $snippet = $trim
                    if ($snippet.Length -gt 120) { $snippet = $snippet.Substring(0, 120) + '...' }
                    $findings.Add([pscustomobject]@{
                            Task = $task.Id; Fase = $task.Fase; Sev = $d.Sev; Detector = $d.Id
                            File = $rel; Line = ($ln + 1); Snippet = $snippet; Msg = $d.Msg
                        })
                }
            }
        }
    }
}

# --- 5. Detectores globales ---------------------------------------------
# MIG-01: versiones de migracion duplicadas (correlativo V<n> ambiguo).
$migDir = Join-Path $repo 'ConversorPDF.XLL'
$migFiles = @(Get-ChildItem -LiteralPath $migDir -Recurse -Filter 'MigrationV*.cs' -File -ErrorAction SilentlyContinue)
$versions = @{}
foreach ($m in $migFiles) {
    if ($m.Name -match 'MigrationV(\d+)') {
        $v = [int]$Matches[1]
        if (-not $versions.ContainsKey($v)) { $versions[$v] = New-Object System.Collections.Generic.List[string] }
        $versions[$v].Add((Rel $m.FullName))
    }
}
foreach ($v in ($versions.Keys | Sort-Object)) {
    if ($versions[$v].Count -gt 1) {
        $findings.Add([pscustomobject]@{
                Task = 'GLOBAL'; Fase = ''; Sev = 'Compilation'; Detector = 'MIG-01'
                File = ($versions[$v] -join ' | '); Line = 0; Snippet = "V$v"
                Msg = "Version de migracion V$v duplicada: __SchemaVersion queda ambiguo"
            })
    }
}

# LG-06: DELETE sobre tablas replicadas por CRDT (genera tombstone propagable).
$syncInit = Join-Path $migDir 'Core\Sync\SyncInitializer.cs'
if (Test-Path -LiteralPath $syncInit) {
    $crdtTables = @()
    $syncText = Get-Content -LiteralPath $syncInit -Raw
    $m = [regex]::Match($syncText, '(?s)CrdtTables\s*=\s*(?:new\s+string\s*\[\s*\]\s*)?\{([^}]*)\}')
    if ($m.Success) { $crdtTables = @([regex]::Matches($m.Groups[1].Value, '"([^"]+)"') | ForEach-Object { $_.Groups[1].Value }) }
    if ($crdtTables.Count -gt 0) {
        foreach ($rel in $audited) {
            $full = Join-Path $repo ($rel -replace '/', '\')
            $content = Get-Content -LiteralPath $full -ErrorAction SilentlyContinue
            for ($ln = 0; $ln -lt $content.Count; $ln++) {
                $text = [string]$content[$ln]
                foreach ($t in $crdtTables) {
                    if ([regex]::IsMatch($text, ('(?i)DELETE\s+FROM\s+' + [regex]::Escape($t) + '\b'))) {
                        $findings.Add([pscustomobject]@{
                                Task = 'GLOBAL'; Fase = ''; Sev = 'Logic'; Detector = 'LG-06'
                                File = $rel; Line = ($ln + 1); Snippet = $text.Trim()
                                Msg = "DELETE en tabla CRDT '$t': el tombstone borra la fila en el vault destino"
                            })
                    }
                }
            }
        }
    }
}

# --- 6. Informe de auditoria --------------------------------------------
$blockers = @($findings | Where-Object { $blocking -contains $_.Sev })
$warns = @($findings | Where-Object { $blocking -notcontains $_.Sev })
$taskIds = ($selected | ForEach-Object { $_.Id }) -join ', '
$faseIds = (($selected | Where-Object { $_.Fase } | ForEach-Object { $_.Fase })) -join ', '
if (-not $faseIds) { $faseIds = '(sin field fase: usar naming de Ronda.md al reportar)' }

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('# Auditoria de tareas (audit-tasks.ps1)')
[void]$sb.AppendLine('')
[void]$sb.AppendLine("- **Plan**: $(Rel $PlanFull)")
[void]$sb.AppendLine("- **Fecha**: $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))")
[void]$sb.AppendLine("- **Tareas auditadas**: $taskIds")
[void]$sb.AppendLine("- **Fases (Ronda.md)**: $faseIds")
[void]$sb.AppendLine("- **Archivos auditados**: $(@($audited | Sort-Object -Unique).Count)")
[void]$sb.AppendLine("- **Hallazgos**: $($findings.Count) (bloqueantes: $($blockers.Count) / warnings: $($warns.Count))")
[void]$sb.AppendLine('')
[void]$sb.AppendLine('## Hallazgos')
if ($findings.Count -eq 0) {
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Sin hallazgos: ningun detector del catalogo coincide en los archivos de las tareas.')
}
else {
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('| Fase | Task | Severidad | Archivo | Lineas | Detector | Hallazgo |')
    [void]$sb.AppendLine('|---|---|---|---|---|---|---|')
    $ordered = @($findings | Sort-Object @{ Expression = { $blocking -notcontains $_.Sev } }, Task, File, Line)
    foreach ($f in $ordered) {
        $fase = if ($f.Fase) { $f.Fase } else { '-' }
        $loc = if ($f.Line -gt 0) { "$($f.File):$($f.Line)" } else { $f.File }
        $msg = ($f.Msg -replace '\|', '/')
        [void]$sb.AppendLine("| $fase | $($f.Task) | $($f.Sev) | $loc | $($f.Line) | $($f.Detector) | $msg |")
    }
}
[void]$sb.AppendLine('')
[void]$sb.AppendLine('## Resumen por tarea')
[void]$sb.AppendLine('')
[void]$sb.AppendLine('| Task | Fase | Archivos | Bloqueantes | Warnings |')
[void]$sb.AppendLine('|---|---|---|---|---|')
foreach ($t in $selected) {
    $tb = @($findings | Where-Object { $_.Task -eq $t.Id -and ($blocking -contains $_.Sev) }).Count
    $tw = @($findings | Where-Object { $_.Task -eq $t.Id -and ($blocking -notcontains $_.Sev) }).Count
    $tf = if ($t.Fase) { $t.Fase } else { '-' }
    [void]$sb.AppendLine("| $($t.Id) | $tf | $($t.Files.Count) | $tb | $tw |")
}
[void]$sb.AppendLine('')
[void]$sb.AppendLine('## Archivos auditados')
foreach ($a in ($audited | Sort-Object -Unique)) { [void]$sb.AppendLine("- ``$a``") }
[void]$sb.AppendLine('')
[void]$sb.AppendLine('## Veredicto')
if ($blockers.Count -gt 0) {
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("``AUDIT_FINDINGS`` - $($blockers.Count) hallazgo(s) bloqueante(s): convertir cada uno en ``Fase {n}`` (Ronda.md) y plan ``T-fix``.")
}
else {
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('``AUDIT_PASS`` - sin hallazgos bloqueantes. Los warnings no impiden sellar, pero se listan en el acta.')
}

Set-Content -LiteralPath $reportPath -Value $sb.ToString() -Encoding UTF8

Write-Host ''
Write-Host "AUDIT report: $(Rel $reportPath)" -ForegroundColor Cyan
Write-Host ("Tareas: {0} | Archivos: {1} | Hallazgos: {2} | Bloqueantes: {3} | Warnings: {4}" -f `
        $taskIds, $audited.Count, $findings.Count, $blockers.Count, $warns.Count) -ForegroundColor Cyan

foreach ($f in ($blockers | Sort-Object Task, File, Line)) {
    Write-Host ("[{0}] {1} {2}:{3} {4} - {5}" -f $f.Sev, $f.Task, $f.File, $f.Line, $f.Detector, $f.Msg) -ForegroundColor Red
}

if ($blockers.Count -gt 0) {
    Write-Host "AUDIT_FINDINGS ($($blockers.Count) bloqueantes) -> FIX_PLAN" -ForegroundColor Yellow
    exit 3
}

Write-Host 'AUDIT_PASS' -ForegroundColor Green
exit 0
