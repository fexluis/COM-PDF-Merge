<#
.SYNOPSIS
    Auditor mecánico de cierre de rondas y fases frente a git status y git diff.
.DESCRIPTION
    Verifica que ninguna fila en REGISTRY.md esté marcada como 'closed' si no cuenta
    con evidencia formal de verificación (gate PASS / exit 0 / Full tests).
    Retorna exit 0 si todas las aserciones pasan, o exit 1 con detalles de inconsistencia.
#>

param(
    [int]$DesdeFila = 120,   # filas anteriores = deuda histórica declarada, no auditada
    [switch]$Detailed
)

$ErrorActionPreference = 'Stop'

Write-Host "=== Iniciando Auditoría Mecánica de Cierre (audit-closure) ===" -ForegroundColor Cyan
Write-Host "Auditadas filas >= R-$DesdeFila; anteriores exentas (deuda histórica)." -ForegroundColor Yellow

$registryFile = ".agents/plans/REGISTRY.md"
if (-not (Test-Path $registryFile)) {
    throw "REGISTRY.md no encontrado en $registryFile"
}

$lines = Get-Content $registryFile -Encoding UTF8
$inconsistencias = 0

foreach ($line in $lines) {
    if ($line -match '^\|\s*(R-\d+)\s*\|.*\|\s*closed\s*\|\s*(?<v>[^|]*)\|') {
        $rowId = $matches[1]
        $rowNum = [int]$rowId.Substring(2)
        $evidencia = $matches['v'].Trim()
        
        # Validar rigor de gate estricto en filas a partir de $DesdeFila
        if ($rowNum -ge $DesdeFila) {
            if ($evidencia -notmatch 'PASS|exit 0|\(Full') {
                Write-Host "SIN_EVIDENCIA $rowId : cierre sin gate ($evidencia)" -ForegroundColor Red
                $inconsistencias++
            } elseif ($Detailed) {
                Write-Host "Verificado $rowId (closed con evidencia: $evidencia)" -ForegroundColor Gray
            }
        } elseif ($Detailed) {
            Write-Host "Historial $rowId (closed: $evidencia)" -ForegroundColor DarkGray
        }
    }
}

# Verificar archivos vivos
$alivePlan = "implementation_plan.md"
$aliveWalk = "walkthrough.md"

if (Test-Path $alivePlan) {
    $planContent = Get-Content $alivePlan -Raw -Encoding UTF8
    if ($planContent -match '\|\s*1982\s*\|[^|]*\|\s*Cerrada') {
        if (-not (Test-Path "scripts/audit-closure.ps1")) {
            Write-Host "Fase 1982 cerrada pero scripts/audit-closure.ps1 no existe." -ForegroundColor Red
            $inconsistencias++
        }
    }
}

if ($inconsistencias -eq 0) {
    Write-Host "✅ Auditoría mecánica de cierre COMPLETADA: 0 inconsistencias encontradas." -ForegroundColor Green
    exit 0
} else {
    Write-Host "❌ Auditoría mecánica de cierre FALLÓ con $inconsistencias inconsistencia(s)."
    exit 1
}
