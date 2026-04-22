$path = 'ConversorPDF.XLL\bas\Via_Cumplidos.bas'
if (-not (Test-Path $path)) {
    Write-Error "File not found: $path"
    exit 1
}

$content = Get-Content $path -Encoding Default -Raw

# Configurable parts
$viaticosPattern = '(?:VIATICOS|VITICOS|VIATICA|VIATICO|VAITICOS|VAITICO|VIATCIOS)'
$VIGENCIA_SUFFIX = 'VIGENCIA_SUFFIX' # Placeholder if needed, but the file uses the variable name literal here

# The exact line from Select-String:
#         .Pattern = "PAGO\s+" & viaticosPattern & "(?:\s+\d*" & VIGENCIA_SUFFIX & ")?(?:\s+PARA)?\b"
$newPattern = '        .Pattern = "PAGO\s+" & viaticosPattern & "(?:(?:\s+\d+)?\s+PARA|(?:\s+\d*" & VIGENCIA_SUFFIX & ")?)\b"'

# Use regex to find and replace the .Pattern line regardless of minor spacing
$oldPatternRegex = '(?m)^\s*\.Pattern\s*=\s*"PAGO\\s\+"\s*&\s*viaticosPattern\s*&\s*"\(\?:\\s\+\\d\*" & VIGENCIA_SUFFIX & "\)\?\(\?:\\s\+PARA\)\?\\b"'

if ($content -match $oldPatternRegex) {
    $content = $content -replace $oldPatternRegex, $newPattern
    Set-Content $path -Value $content -Encoding Default
    Write-Host "VBA file updated successfully."
} else {
    Write-Error "Could not find the pattern in the VBA file using regex."
    exit 1
}
