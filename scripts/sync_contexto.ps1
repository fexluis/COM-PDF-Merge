$date = Get-Date -Format "yyyyMMdd-HHmm.ss.fff"
$targetDir = "_FIXES\fix-$date"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$checkpointPath = ".agents\checkpoints"
$checkpointFile = "$checkpointPath\contexto-previo.checkpoint"

if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
if (-not (Test-Path $checkpointPath)) { New-Item -ItemType Directory -Path $checkpointPath -Force | Out-Null }

$appVersion = "1.6.3.65"
$appFile = "ConversorPDF.XLL\Core\Helpers\App.cs"
if (Test-Path $appFile) {
    if ((Get-Content $appFile -Raw) -match "public const string APP_VERSION\s*=\s*`"(.+?)`"") {
        $appVersion = $Matches[1]
    }
}

$lastExec = (Get-Date).AddHours(-4)
if (Test-Path $checkpointFile) {
    $lastExec = (Get-Item $checkpointFile).LastWriteTime
}

$sourceFiles = Get-ChildItem -Path . -Filter "*.cs" -Recurse | 
    Where-Object { $_.LastWriteTime -gt $lastExec -and $_.FullName -notmatch "(\\obj\\|\\bin\\|\\_FIXES\\)" }

$count = 0
foreach ($f in $sourceFiles) {
    if (Test-Path $f.FullName) {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($f.FullName)
        $relPath = $f.FullName.Replace((Get-Location).Path, "").TrimStart("\")
        $dest = Join-Path $targetDir "$name.txt"
        $header = "// ============================================================================`r`n// AUDITORIA DE ARCHIVO FUENTE`r`n// ============================================================================`r`n// Archivo original: $relPath`r`n// Version:          $appVersion`r`n// Fecha captura:    $timestamp`r`n// ============================================================================`r`n`r`n"
        $content = Get-Content $f.FullName -Raw
        Set-Content -Path $dest -Value ($header + $content) -NoNewline
        $count++
    }
}

Set-Content -Path $checkpointFile -Value $timestamp
Write-Host "Sincronizados $count archivos en $targetDir"
