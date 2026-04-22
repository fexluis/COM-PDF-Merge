---
description: Sincronizar copias de seguridad (.txt) de los archivos editados en _FIXES/fix-YYYY-MM-DD
---

# Contexto Previo (Sincronización de Auditoría)

Este workflow automatiza la captura de archivos modificados y documentación, generando copias `.txt` con encabezados de trazabilidad para auditoría de cambios.

---

## Pasos

### 1. Configuración de Entorno y Versión

Calcula rutas, fecha y extrae la versión actual de la aplicación.

// turbo
```powershell
$date      = Get-Date -Format "yyyyMMdd-HHmm.ss.fff"
$targetDir = "_FIXES\fix-$date"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$checkpointPath = ".agents\checkpoints"
$checkpointFile = "$checkpointPath\contexto-previo.checkpoint"

if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force }
if (-not (Test-Path $checkpointPath)) { New-Item -ItemType Directory -Path $checkpointPath -Force }

# Extraer versión desde App.cs
$appVersion = "1.0.0.0"
$appFile = "ConversorPDF.XLL\Core\Helpers\App.cs"
if (Test-Path $appFile) {
    if ((Get-Content $appFile -Raw) -match "public const string APP_VERSION\s*=\s*`"(.+?)`"") {
        $appVersion = $Matches[1]
    }
}

Write-Host "Contexto: $timestamp | Versión: $appVersion"
Write-Host "Destino: $targetDir"
```

### 2. Sincronizar archivos fuente (.cs) modificados

Detecta archivos modificados desde la última ejecución y genera copias con encabezado de auditoría.

// turbo
```powershell
# Obtener fecha de último checkpoint
$lastExec = (Get-Date).AddHours(-4) # Fallback: últimas 4 horas
if (Test-Path $checkpointFile) {
    $lastExec = (Get-Item $checkpointFile).LastWriteTime
}
Write-Host "Buscando cambios después de: $lastExec"

# Captura archivos .cs modificados (excluyendo carpetas de build/temp)
$sourceFiles = Get-ChildItem -Path . -Filter "*.cs" -Recurse | 
    Where-Object { $_.LastWriteTime -gt $lastExec -and $_.FullName -notmatch "(\\obj\\|\\bin\\|\\_FIXES\\)" }

foreach ($f in $sourceFiles) {
    if (Test-Path $f.FullName) {
        $name    = [System.IO.Path]::GetFileNameWithoutExtension($f.FullName)
        $destino = "$targetDir\$name.txt"

        $header = @"
# ============================================================
# AUDITORÍA DE CAMBIO
# ============================================================
# Versión de la aplicación : $appVersion
# Archivo original          : $($f.FullName)
# Fecha de copia            : $timestamp
# ------------------------------------------------------------
# DIAGNÓSTICO (problema identificado antes del cambio)
# [Completar: descripción del bug/issue/necesidad detectada]
#
# IMPLEMENTADO (solución aplicada en esta versión)
# [Completar: descripción de los cambios realizados]
# ============================================================

"@
        $header + (Get-Content $f.FullName -Raw) | Set-Content $destino -Encoding UTF8
        Write-Host "Sincronizado Fuente: $($f.Name) --> $destino"
    }
}

if ($sourceFiles.Count -eq 0) { Write-Host "No se detectaron archivos fuente modificados." }
```

### 3. Sincronizar documentación (Brain / Artifacts)

Copia los artefactos de la sesión actual (walkthrough, implementation_plan) con encabezado.

// turbo
```powershell
$brainDir = "C:\Users\luis.bechara\.gemini\antigravity\brain"
$activeConv = Get-ChildItem -Path $brainDir -Directory | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($activeConv) {
    $docFiles = @("walkthrough.md", "implementation_plan.md", "task.md")
    
    foreach ($docName in $docFiles) {
        $fullPath = Join-Path $activeConv.FullName $docName
        if (Test-Path $fullPath) {
            $name    = [System.IO.Path]::GetFileNameWithoutExtension($fullPath)
            $destino = "$targetDir\$name.txt"

            $header = @"
# ============================================================
# AUDITORÍA DE DOCUMENTACIÓN
# ============================================================
# Versión de la aplicación : $appVersion
# Archivo original          : $fullPath
# Fecha de copia            : $timestamp
# ============================================================

"@
            $header + (Get-Content $fullPath -Raw) | Set-Content $destino -Encoding UTF8
            Write-Host "Sincronizada Doc: $docName --> $destino"
        }
    }
}
```

### 4. Actualizar Checkpoint y Verificar

Actualiza la marca de tiempo para la próxima ejecución.

// turbo
```powershell
# Actualizar marca de tiempo
Set-Content $checkpointFile $timestamp
Write-Host "Checkpoint actualizado: $timestamp"

# Resumen
Get-ChildItem $targetDir
```