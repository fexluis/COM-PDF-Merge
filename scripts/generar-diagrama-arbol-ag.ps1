# Genera el diagrama Sede -> Semestre del perfil AG 2027-01 (1920x1080, 16:9)
# Salida: docs\AG\arbol_AG_2027-01.svg
param(
    [string]$OutDir = 'docs\AG',
    [string]$Name = 'arbol_AG_2027-01'
)

$ErrorActionPreference = 'Stop'

$A   = [char]193   # A con acento agudo
$O   = [char]211   # O con acento agudo
$DOT = [char]183   # punto medio
$EM  = [char]8212  # guion largo
$ARR = [char]8594  # flecha derecha

$sedes = @(
    @{ n = 'CALI'; c = @('I-A|53','I-B|53','II-A|53','II-B|53','III-A|53','III-B|53','IV-A|53','IV-B|53','V-A|53','V-B|53','VI-A|53','VI-B|53','VII-A|53','VII-B|53','VIII-A|52','IX-A|52','X-A|52') },
    @{ n = 'CARTAGO'; c = @('IV-A|53','VII-A|53','X-A|53') },
    @{ n = ('TULU' + $A); c = @('I-A|53','VI-A|53','X-A|52') },
    @{ n = 'FLORIDA'; c = @('III-A|53','IX-A|52') },
    @{ n = 'BUENAVENTURA'; c = @('V-A|53','IX-A|52') },
    @{ n = 'PRADERA'; c = @('IV-A|53','VI-A|53') },
    @{ n = 'GUADALAJARA DE BUGA'; c = @('V-A|53','X-A|52') },
    @{ n = 'CAICEDONIA'; c = @('VIII-A|52') },
    @{ n = ('EL ' + $A + 'GUILA'); c = @('III-A|53') },
    @{ n = ('LA UNI' + $O + 'N'); c = @('IX-A|52') },
    @{ n = 'SEVILLA'; c = @('IX-A|52') }
)

$L = New-Object System.Collections.Generic.List[string]
$L.Add('<?xml version="1.0" encoding="UTF-8"?>')
$L.Add('<svg xmlns="http://www.w3.org/2000/svg" width="1920" height="1080" viewBox="0 0 1920 1080" font-family="Segoe UI, Arial, sans-serif">')
$L.Add('<defs><marker id="a" viewBox="0 0 10 10" refX="8.5" refY="5" markerWidth="5" markerHeight="5" orient="auto"><path d="M0,0 L10,5 L0,10 z" fill="#94a3b8"/></marker></defs>')
$L.Add('<rect width="1920" height="1080" fill="#ffffff"/>')
$L.Add('<text x="48" y="58" font-size="30" font-weight="700" fill="#0f172a">' + $A + 'rbol del AG ' + $EM + ' Perfil 2027-01</text>')
$L.Add('<text x="48" y="88" font-size="15" fill="#64748b">35 cohortes ' + $DOT + ' 11 sedes ' + $DOT + ' Sede ' + $ARR + ' Semestre (grupo A/B) ' + $DOT + ' Fuente: Corrida #140 (24/11/2026)</text>')
$L.Add('<rect x="1448" y="42" width="16" height="16" rx="4" fill="#eff6ff" stroke="#1d4ed8" stroke-width="2"/>')
$L.Add('<text x="1472" y="56" font-size="14" fill="#0f172a">APT-53 (pensum 53)</text>')
$L.Add('<rect x="1630" y="42" width="16" height="16" rx="4" fill="#fffbeb" stroke="#b45309" stroke-width="2"/>')
$L.Add('<text x="1654" y="56" font-size="14" fill="#0f172a">APT-52 (pensum 52)</text>')
$L.Add('<line x1="48" y1="112" x2="1872" y2="112" stroke="#e2e8f0" stroke-width="1"/>')

$bandTop = 130
for ($i = 0; $i -lt $sedes.Count; $i++) {
    $s = $sedes[$i]
    $n = $s.c.Count
    $rows = 1
    if ($n -gt 9) { $rows = 2 }
    $bandH = 78
    if ($rows -eq 2) { $bandH = 132 }
    $cy = [Math]::Round($bandTop + $bandH / 2, 1)

    $L.Add('<rect x="48" y="' + ($cy - 27) + '" width="252" height="54" rx="12" fill="#0f172a"/>')
    $fs = 17
    if ($s.n.Length -gt 16) { $fs = 15 }
    $L.Add('<text x="68" y="' + ($cy - 2) + '" font-size="' + $fs + '" font-weight="700" fill="#ffffff">' + $s.n + '</text>')
    $ct = '1 cohorte'
    if ($n -gt 1) { $ct = ('{0} cohortes' -f $n) }
    $L.Add('<text x="68" y="' + ($cy + 15) + '" font-size="11.5" fill="#94a3b8">' + $ct + '</text>')

    $perRow = [Math]::Ceiling($n / $rows)
    $slot = 1550 / $perRow
    $w = [Math]::Round([Math]::Min(150, $slot - 14), 1)
    $chips = New-Object System.Collections.Generic.List[object]
    $idx = 0
    for ($r = 0; $r -lt $rows; $r++) {
        $cnt = [Math]::Min($perRow, $n - $r * $perRow)
        if ($cnt -le 0) { continue }
        $rc = [Math]::Round($bandTop + ($bandH / (2 * $rows)) * (2 * $r + 1), 1)
        $off = ($perRow - $cnt) * $slot / 2
        for ($k = 0; $k -lt $cnt; $k++) {
            $cx = [Math]::Round(330 + $off + $slot * ($k + 0.5), 1)
            $x = [Math]::Round($cx - $w / 2, 1)
            $p = $s.c[$idx].Split('|')
            $chips.Add([pscustomobject]@{ cx = $cx; x = $x; y = $rc; sem = $p[0]; pr = $p[1] })
            $idx++
        }
    }

    foreach ($ch in $chips) {
        $L.Add('<path d="M302 ' + $cy + ' H316 V ' + $ch.y + ' H ' + ($ch.x - 8) + '" fill="none" stroke="#cbd5e1" stroke-width="1.4" marker-end="url(#a)"/>')
    }
    foreach ($ch in $chips) {
        $fill = '#fffbeb'
        $stroke = '#b45309'
        if ($ch.pr -eq '53') { $fill = '#eff6ff'; $stroke = '#1d4ed8' }
        $L.Add('<rect x="' + $ch.x + '" y="' + ($ch.y - 23) + '" width="' + $w + '" height="46" rx="9" fill="' + $fill + '" stroke="' + $stroke + '" stroke-width="1.8"/>')
        $L.Add('<text x="' + $ch.cx + '" y="' + ($ch.y - 5) + '" font-size="10" fill="#64748b" text-anchor="middle">Semestre</text>')
        $L.Add('<text x="' + $ch.cx + '" y="' + ($ch.y + 14) + '" font-size="17" font-weight="700" fill="#0f172a" text-anchor="middle">' + $ch.sem + '</text>')
    }

    $bandTop = $bandTop + $bandH
}

$footer = 'Modalidad: Presencial (CALI I-A/I-B, TULU' + $A + ' I-A) ' + $DOT + ' Mixta (resto) ' + $DOT + ' Plataforma (FLORIDA IX-A, CAICEDONIA VIII-A, GUADALAJARA DE BUGA V-A y X-A) ' + $DOT + ' Subsedes: FLORIDA/PLATAFORMA, CAICEDONIA/PLATAFORMA, GUADALAJARA DE BUGA/SEVILLA y /FLORIDA-PRADERA'
$L.Add('<text x="48" y="1066" font-size="12" fill="#94a3b8">' + $footer + '</text>')
$L.Add('</svg>')

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$full = (Resolve-Path $OutDir).Path
$svgPath = Join-Path $full ($Name + '.svg')
$enc = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($svgPath, ($L -join [char]10), $enc)

Write-Output ('SVG: ' + $svgPath)
Write-Output ('Bytes: ' + (Get-Item $svgPath).Length)
Write-Output ('Elementos: ' + $L.Count)
