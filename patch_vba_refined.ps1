$path = 'ConversorPDF.XLL\bas\Via_Cumplidos.bas'
if (-not (Test-Path $path)) {
    Write-Error "File not found: $path"
    exit 1
}

$content = Get-Content $path -Encoding Default -Raw

# The pattern logic to find:
# .Pattern = "PAGO\s+" & viaticosPattern & ...
# We want to replace it with the branching logic version.

$viaticosPattern = '(?:VIATICOS|VITICOS|VIATICA|VIATICO|VAITICOS|VAITICO|VIATCIOS)'
$newPattern = '        .Pattern = "PAGO\s+" & viaticosPattern & "(?:(?:\s+\d+)?\s+PARA|(?:\s+\d*" & VIGENCIA_SUFFIX & ")?)\b"'

# Flexible regex to find the .Pattern line
$patternRegex = '(?m)^\s*\.Pattern\s*=\s*"PAGO\\s\+"\s*&\s*viaticosPattern\s*&\s*".*?"\b'

if ($content -match $patternRegex) {
    $content = $content -replace $patternRegex, $newPattern
    Set-Content $path -Value $content -Encoding Default
    Write-Host "VBA file updated successfully with refined branching logic."
} else {
    Write-Error "Could not find the .Pattern line in the VBA file to replace."
    exit 1
}
