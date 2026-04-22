$path = 'ConversorPDF.XLL\bas\Via_Cumplidos.bas'
if (-not (Test-Path $path)) {
    Write-Error "File not found: $path"
    exit 1
}

$content = Get-Content $path -Encoding Default -Raw
# Update the pattern to include optional PARA
$oldPattern = '.Pattern = "PAGO\s+" & viaticosPattern & "(?:\s+\d*" & VIGENCIA_SUFFIX & ")?\b"'
$newPattern = '.Pattern = "PAGO\s+" & viaticosPattern & "(?:\s+\d*" & VIGENCIA_SUFFIX & ")?(?:\s+PARA)?\b"'

if ($content -contains $oldPattern) {
    $content = $content.Replace($oldPattern, $newPattern)
    Set-Content $path -Value $content -Encoding Default
    Write-Host "VBA file updated successfully with PARA support."
} else {
    Write-Error "Could not find the old pattern in the VBA file."
    exit 1
}
