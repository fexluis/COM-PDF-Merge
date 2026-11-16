#Requires -Version 5.1
<#
.SYNOPSIS
  Escinde implementation_plan.md / walkthrough.md en trozos por rango de fases y deja un indice en el archivo vivo.
.DESCRIPTION
  - Corte jerarquico por encabezados H2..H4, ignorando encabezados dentro de bloques de codigo.
  - No reordena ni edita contenido: cada trozo copia exacta de un rango de lineas.
  - El archivo vivo conserva indice + ultimas N rondas.
  - Valida que trozos + vivo == original antes de escribir.
  - Genera docs/plans/history/LINE_MAP.md para citas antiguas "<archivo>.md:NNNN".
.PARAMETER Base
  both | implementation_plan | walkthrough
.PARAMETER Target
  Lineas objetivo por trozo (default 1000).
.PARAMETER LiveRondas
  Rondas que quedan en el archivo vivo (default 5).
.PARAMETER Apply
  Sin este switch solo hace DryRun.
#>
[CmdletBinding()]
param(
    [ValidateSet('both', 'implementation_plan', 'walkthrough')]
    [string]$Base = 'both',
    [int]$Target = 1000,
    [int]$LiveRondas = 5,
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if (-not (Test-Path (Join-Path $root 'implementation_plan.md'))) {
    throw "Raiz del repositorio no detectada desde '$PSScriptRoot'."
}

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Read-Doc([string]$path) {
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    $text = [System.Text.Encoding]::UTF8.GetString($bytes)
    if ($hasBom) { $text = $text.Substring(1) }
    $eol = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }
    $norm = $text.Replace("`r`n", "`n")
    if ($norm.EndsWith("`n")) { $norm = $norm.Substring(0, $norm.Length - 1) }
    $lines = if ($norm.Length -eq 0) { @() } else { @($norm -split "`n") }
    return [pscustomobject]@{ Lines = $lines; Eol = $eol; Bom = $hasBom; Original = $text }
}

function Get-Heads($lines) {
    $heads = New-Object System.Collections.ArrayList
    $inFence = $false
    $mark = ''
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $trim = $lines[$i].TrimStart()
        if ($trim -match '^(`{3,}|~{3,})') {
            if (-not $inFence) { $inFence = $true; $mark = $Matches[1].Substring(0, 3) }
            elseif ($trim.StartsWith($mark)) { $inFence = $false; $mark = '' }
            continue
        }
        if ($inFence) { continue }
        if ($lines[$i] -match '^(#{1,6})\s+(\S.*)$') {
            [void]$heads.Add([pscustomobject]@{ Index = $i; Level = $Matches[1].Length; Text = $Matches[2] })
        }
    }
    return $heads
}

function Get-Phases($lines, [int]$from, [int]$to) {
    $found = New-Object System.Collections.ArrayList
    $inFence = $false
    $mark = ''
    for ($i = $from; $i -lt $to; $i++) {
        $trim = $lines[$i].TrimStart()
        if ($trim -match '^(`{3,}|~{3,})') {
            if (-not $inFence) { $inFence = $true; $mark = $Matches[1].Substring(0, 3) }
            elseif ($trim.StartsWith($mark)) { $inFence = $false; $mark = '' }
            continue
        }
        if ($inFence) { continue }
        foreach ($m in [regex]::Matches($lines[$i], 'Fases?\s*(\d{3,4})(?:\s*(?:a|A|/|-|\u2013|\u2014)\s*(\d{3,4}))?')) {
            [void]$found.Add([int]$m.Groups[1].Value)
            if ($m.Groups[2].Success) { [void]$found.Add([int]$m.Groups[2].Value) }
        }
    }
    return $found
}

function Get-PhasesFromHeads($heads, [int]$from, [int]$to) {
    $found = New-Object System.Collections.ArrayList
    foreach ($h in $heads) {
        if ($h.Index -lt $from -or $h.Index -ge $to) { continue }
        foreach ($m in [regex]::Matches($h.Text, 'Fases?\s*(\d{3,4})(?:\s*(?:a|A|/|-|\u2013|\u2014)\s*(\d{3,4}))?')) {
            [void]$found.Add([int]$m.Groups[1].Value)
            if ($m.Groups[2].Success) { [void]$found.Add([int]$m.Groups[2].Value) }
        }
    }
    return $found
}

function Get-Rondas($lines, [int]$from, [int]$to) {
    $labels = New-Object System.Collections.ArrayList
    for ($i = $from; $i -lt $to; $i++) {
        if ($lines[$i] -match '^##\s+Ronda\s+([0-9]+)') { [void]$labels.Add([int]$Matches[1]) }
    }
    return $labels
}

function Get-RondasHeredadas($heads, [int]$from, [int]$to, $own) {
    $labels = New-Object System.Collections.ArrayList
    if ($own.Count -gt 0) { foreach ($n in $own) { [void]$labels.Add($n) } ; return $labels }
    $last = -1
    foreach ($h in $heads) {
        if ($h.Index -ge $to) { break }
        if ($h.Text -match '^Ronda\s+([0-9]+)') {
            if ($h.Index -lt $from) { $last = [int]$Matches[1] }
            else { [void]$labels.Add([int]$Matches[1]) }
        }
    }
    if ($labels.Count -eq 0 -and $last -ge 0) { [void]$labels.Add($last) }
    return $labels
}

function Format-Range($list) {
    if ($null -eq $list -or $list.Count -eq 0) { return '-' }
    $arr = @($list | Sort-Object -Unique)
    if ($arr.Count -le 6) { return ($arr -join ', ') }
    return ($arr[0].ToString() + '..' + $arr[$arr.Count - 1].ToString())
}

function New-Chunks($lines, $heads, [int]$target, [int]$liveRondas) {
    $rondaHeads = @($heads | Where-Object { $_.Level -eq 2 -and $_.Text -match '^Ronda\s+[0-9]+' })
    $liveStart = $lines.Count
    if ($rondaHeads.Count -gt $liveRondas) {
        $candidate = $rondaHeads[$rondaHeads.Count - $liveRondas].Index
        if (($lines.Count - $candidate) -le 400) { $liveStart = $candidate }
    }
    $cutMap = @{}
    foreach ($h in $heads) {
        if ($h.Index -gt 0 -and $h.Index -lt $liveStart) { $cutMap[$h.Index] = $h.Level }
    }
    $cuts = @($cutMap.Keys | Sort-Object)

    $chunks = New-Object System.Collections.ArrayList
    $start = 0
    while ($start -lt $liveStart) {
        if (($liveStart - $start) -le [int]($target * 1.3)) {
            [void]$chunks.Add([pscustomobject]@{ Start = $start; End = $liveStart })
            break
        }
        $low = $start + [int]($target * 0.80)
        $high = $start + [int]($target * 1.25)
        $pick = -1
        foreach ($c in $cuts) {
            if ($c -le $low) { continue }
            if ($c -gt $high) { break }
            if ($cutMap[$c] -le 3) { $pick = $c }
        }
        if ($pick -lt 0) {
            foreach ($c in $cuts) {
                if ($c -le $low) { continue }
                if ($c -gt $high) { break }
                $pick = $c
            }
        }
        if ($pick -lt 0) {
            foreach ($c in $cuts) {
                if ($c -gt $start + [int]($target * 0.5)) { $pick = $c; break }
            }
        }
        if ($pick -le $start) { $pick = $liveStart }
        [void]$chunks.Add([pscustomobject]@{ Start = $start; End = $pick })
        $start = $pick
    }
    return [pscustomobject]@{ Chunks = $chunks; LiveStart = $liveStart }
}

function Invoke-Base([string]$name, [int]$target, [int]$liveRondas, [bool]$apply, [string]$sha, [string]$stamp) {
    $path = Join-Path $root "$name.md"
    $doc = Read-Doc $path
    if ($doc.Lines.Count -gt 0 -and $doc.Lines[0] -match '^#\s+Indice de Historial') {
        throw "$name.md ya esta escindido (indice detectado). Restaurar el monolitico antes de volver a escindir."
    }
    $heads = Get-Heads $doc.Lines
    $plan = New-Chunks $doc.Lines $heads $target $liveRondas
    $historyDir = Join-Path $root ("docs\plans\history\" + $name)
    if ($apply -and -not (Test-Path $historyDir)) { [void](New-Item -ItemType Directory -Path $historyDir -Force) }

    $used = @{}
    $items = New-Object System.Collections.ArrayList
    foreach ($c in $plan.Chunks) {
        $phases = Get-PhasesFromHeads $heads $c.Start $c.End
        if ($phases.Count -eq 0) { $phases = Get-Phases $doc.Lines $c.Start $c.End }
        $min = if ($phases.Count -gt 0) { ($phases | Measure-Object -Minimum).Minimum } else { 0 }
        $max = if ($phases.Count -gt 0) { ($phases | Measure-Object -Maximum).Maximum } else { 0 }
        $file = ('{0}_F{1:0000}_F{2:0000}.md' -f $name, $min, $max)
        if ($used.ContainsKey($file)) {
            $used[$file] = $used[$file] + 1
            $file = ('{0}_F{1:0000}_F{2:0000}_{3}.md' -f $name, $min, $max, [char](96 + $used[$file]))
        } else { $used[$file] = 1 }
        [void]$items.Add([pscustomobject]@{
                File = $file; Start = $c.Start; End = $c.End; Lines = $c.End - $c.Start
                Min = $min; Max = $max
                Rondas = (Get-RondasHeredadas $heads $c.Start $c.End @(Get-Rondas $doc.Lines $c.Start $c.End))
            })
    }

    $header = @(
        "<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit $sha - no editar -->",
        "# $name - Fases $($items[0].Min)-$($items[$items.Count - 1].Max)",
        '',
        "> Indice: [../../../$name.md](../../../$name.md) - Escision: $stamp",
        '', '---'
    )

    $rebuilt = New-Object System.Collections.ArrayList
    foreach ($it in $items) {
        if ($it.End -gt $it.Start) { foreach ($l in $doc.Lines[$it.Start..($it.End - 1)]) { [void]$rebuilt.Add($l) } }
    }
    if ($plan.LiveStart -lt $doc.Lines.Count) {
        foreach ($l in $doc.Lines[$plan.LiveStart..($doc.Lines.Count - 1)]) { [void]$rebuilt.Add($l) }
    }
    $rebuiltText = if ($rebuilt.Count -gt 0) { ($rebuilt -join $doc.Eol) + $doc.Eol } else { '' }
    if ($rebuiltText -ne $doc.Original) {
        throw "Validacion fallida en $name.md: trozos + vivo no reproducen el original. No se escribio nada."
    }
    return [pscustomobject]@{
        Name = $name; Doc = $doc; Items = $items; LiveStart = $plan.LiveStart
        Header = $header; Dir = $historyDir; LiveLines = $doc.Lines.Count - $plan.LiveStart
    }
}

# ---------------- ejecucion ----------------
$sha = 'desconocido'
try {
    $shaOut = & git -C $root rev-parse --short HEAD 2>$null
    if ($LASTEXITCODE -eq 0 -and $shaOut) { $sha = $shaOut.Trim() }
} catch { }
$stamp = (Get-Date).ToString('yyyy-MM-dd')
$bases = if ($Base -eq 'both') { @('implementation_plan', 'walkthrough') } else { @($Base) }
$results = New-Object System.Collections.ArrayList
foreach ($name in $bases) {
    [void]$results.Add((Invoke-Base $name $Target $LiveRondas ([bool]$Apply) $sha $stamp))
}

foreach ($r in $results) {
    Write-Host ''
    Write-Host ("== {0}.md ==  trozos={1}  lineas_vivo={2}" -f $r.Name, $r.Items.Count, $r.LiveLines) -ForegroundColor Cyan
    foreach ($it in $r.Items) {
        Write-Host ("  {0,-46} L{1,6}-{2,-6} lineas={3,5} fases={4}-{5}" -f $it.File, ($it.Start + 1), $it.End, $it.Lines, $it.Min, $it.Max)
    }
    Write-Host ("  {0,-46} L{1,6}-{2,-6} lineas={3,5} (vivo)" -f "$($r.Name).md", ($r.LiveStart + 1), $r.Doc.Lines.Count, $r.LiveLines)
}

if (-not $Apply) {
    Write-Host ''
    Write-Host 'DRY-RUN: no se escribio nada. Ejecutar con -Apply para materializar.' -ForegroundColor Yellow
    return
}

$lineMaps = @{}
foreach ($r in $results) {
    $name = $r.Name
    $doc = $r.Doc
    $index = New-Object System.Collections.ArrayList
    [void]$index.Add("# Indice de Historial - $name.md")
    [void]$index.Add('')
    [void]$index.Add("Escision: $stamp - commit ``$sha`` - trozo objetivo: $Target lineas")
    [void]$index.Add('')
    [void]$index.Add("Contenido historico congelado en ``docs/plans/history/$name/``. Orden append-only: no se reordena ni se edita.")
    [void]$index.Add('')
    [void]$index.Add('| Archivo | Fases | Rondas | Lineas |')
    [void]$index.Add('|---|---|---|---|')
    foreach ($it in $r.Items) {
        $rel = "docs/plans/history/$name/" + $it.File
        $ph = if ($it.Min -gt 0) { "$($it.Min)-$($it.Max)" } else { 'sin fase' }
        [void]$index.Add("| [$($it.File)]($rel) | $ph | $(Format-Range $it.Rondas) | $($it.Lines) |")
    }
    [void]$index.Add("| [``$name.md``]($name.md) (vivo) | posteriores al corte | ultimas $LiveRondas | $($r.LiveLines) |")
    [void]$index.Add('')
    [void]$index.Add('## Busqueda por fase')
    [void]$index.Add('Toda `Fase N` se localiza en el primer archivo cuyo rango la contenga; ante solape de rangos, usar la columna Rondas.')
    [void]$index.Add('')
    [void]$index.Add('## Citas historicas')
    [void]$index.Add("Las citas ``$name.md:NNNN`` anteriores al corte apuntan al archivo monolitico; su ubicacion actual esta en [LINE_MAP.md](docs/plans/history/LINE_MAP.md).")
    [void]$index.Add('')
    [void]$index.Add('---')
    [void]$index.Add('')

    $map = @{}
    foreach ($it in $r.Items) {
        $body = if ($it.End -gt $it.Start) { $doc.Lines[$it.Start..($it.End - 1)] } else { @() }
        $ph = if ($it.Min -gt 0) { "Fases $($it.Min)-$($it.Max)" } else { 'sin fase' }
        $head = @(
            ($r.Header[0] -replace 'commit .* - no editar', "commit $sha - no editar"),
            "# $($r.Name) - $ph (Rondas $(Format-Range $it.Rondas))"
        ) + $r.Header[2..5]
        $content = (($head + $body) -join $doc.Eol) + $doc.Eol
        [System.IO.File]::WriteAllText((Join-Path $r.Dir $it.File), $content, $utf8NoBom)
        for ($i = $it.Start; $i -lt $it.End; $i++) { $map[$i + 1] = @($it.File, ($head.Count + ($i - $it.Start) + 1)) }
    }

    $live = New-Object System.Collections.ArrayList
    foreach ($l in $index) { [void]$live.Add($l) }
    if ($r.LiveStart -lt $doc.Lines.Count) {
        foreach ($l in $doc.Lines[$r.LiveStart..($doc.Lines.Count - 1)]) { [void]$live.Add($l) }
    }
    [System.IO.File]::WriteAllText((Join-Path $root "$($r.Name).md"), (($live -join $doc.Eol) + $doc.Eol), $utf8NoBom)
    for ($i = $r.LiveStart; $i -lt $doc.Lines.Count; $i++) { $map[$i + 1] = @("$($r.Name).md", ($index.Count + ($i - $r.LiveStart) + 1)) }
    $lineMaps[$r.Name] = $map
}

# ---------------- LINE_MAP ----------------
$rows = New-Object System.Collections.ArrayList
$pattern = '(implementation_plan|walkthrough)\.md:(\d+)'
$scan = Get-ChildItem -Path $root -Recurse -File -Filter '*.md' -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -notmatch '\\_evidence\\' -and $_.FullName -notmatch '\\\.git\\' -and $_.Name -ne 'LINE_MAP.md' }
foreach ($f in $scan) {
    $text = [System.IO.File]::ReadAllText($f.FullName)
    foreach ($m in [regex]::Matches($text, $pattern)) {
        $docName = $m.Groups[1].Value
        $old = [int]$m.Groups[2].Value
        if (-not $lineMaps.ContainsKey($docName)) { continue }
        $map = $lineMaps[$docName]
        if ($map.ContainsKey($old)) {
            $t = $map[$old]
            [void]$rows.Add("| ``$docName.md:$old`` | ``$($t[0]):$($t[1])`` |")
        }
    }
}
$mlines = New-Object System.Collections.ArrayList
[void]$mlines.Add('# LINE_MAP - citas historicas')
[void]$mlines.Add('')
[void]$mlines.Add("Escision: $stamp - commit ``$sha``")
[void]$mlines.Add('')
[void]$mlines.Add('Citas `<archivo>.md:NNNN` previas a la escision, con su ubicacion actual. El texto historico no se reescribe.')
[void]$mlines.Add('')
[void]$mlines.Add('| Cita original | Ubicacion actual |')
[void]$mlines.Add('|---|---|')
foreach ($row in ($rows | Sort-Object -Unique)) { [void]$mlines.Add($row) }
[System.IO.File]::WriteAllText((Join-Path $root 'docs\plans\history\LINE_MAP.md'), (($mlines -join "`r`n") + "`r`n"), $utf8NoBom)

Write-Host ''
Write-Host ("LINE_MAP: {0} citas mapeadas" -f ($rows | Sort-Object -Unique).Count) -ForegroundColor Green
Write-Host 'Aplicado.' -ForegroundColor Green

