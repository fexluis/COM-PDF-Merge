param(
    [string]$Command = "SELECT id, idvinculacion, idtercero, fechasesionInicial, duracion, horas FROM sesiones ORDER BY id DESC;",
    [string]$DbPath = ""
)

$scriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($DbPath)) {
    $DbPath = [System.IO.Path]::GetFullPath((Join-Path $scriptDir "decrypted.db"))
}

$dllPath = [System.IO.Path]::GetFullPath((Join-Path $scriptDir "..\ConversorPDF.XLL\bin\x64\Debug\System.Data.SQLite.dll"))

[Reflection.Assembly]::LoadFile($dllPath) | Out-Null

$conn = New-Object System.Data.SQLite.SQLiteConnection("Data Source=$DbPath;Version=3;")
$conn.Open()

$cmd = New-Object System.Data.SQLite.SQLiteCommand($Command, $conn)
$reader = $cmd.ExecuteReader()
$results = @()
while ($reader.Read()) {
    $row = [ordered]@{ }
    for ($i = 0; $i -lt $reader.FieldCount; $i++) {
        $row[$reader.GetName($i)] = $reader.GetValue($i)
    }
    $results += New-Object PSObject -Property $row
}
$conn.Close()

if ($results.Count -gt 0) {
    Write-Host "Base de datos: $DbPath"
    Write-Host "Ejecutando: $Command"
    Write-Host "Resultados ($($results.Count) encontrados):"
    $results | Format-Table -AutoSize
} else {
    Write-Host "No se encontraron resultados en $DbPath."
}
