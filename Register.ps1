param(
  [ValidateSet('register','unregister')]
  [string]$Action = 'register',
  [ValidateSet('Debug','Release')]
  [string]$Configuration = 'Debug'
)
$regasm = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\regasm.exe'
if (-not (Test-Path -Path $regasm)) {
  Write-Error 'No se encontró regasm de 64 bits.'
  exit 1
}
$projectDir = Join-Path $PSScriptRoot 'ConversorPDF'
$dll = Join-Path $projectDir ("bin\\x64\\{0}\\ConversorPDF.dll" -f $Configuration)
if (-not (Test-Path -Path $dll)) {
  Write-Error ('No se encontró la DLL: {0}' -f $dll)
  exit 1
}
if ($Action -eq 'register') {
  & $regasm $dll /tlb /codebase
  exit $LASTEXITCODE
} else {
  & $regasm $dll /u
  exit $LASTEXITCODE
}
