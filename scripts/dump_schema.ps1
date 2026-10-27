$dllPath = "C:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.2.6.21\ConversorPDF.Tests\bin\Debug\net48\System.Data.SQLite.dll"
Add-Type -Path $dllPath

$tempDir = "C:\Users\USUARIO\AppData\Local\Temp"
$files = Get-ChildItem -Path $tempDir -Filter "esap_acad_tmp_AcademicData_BOLIVAR*.db" | Sort-Object LastWriteTime -Descending

if ($files.Count -eq 0) {
    Write-Host "No decrypted db found in Temp."
    exit
}

$targetDb = $files[0].FullName
Write-Host "Using DB: $targetDb"

$connString = "Data Source=$targetDb;Version=3;"
$conn = New-Object System.Data.SQLite.SQLiteConnection($connString)
$conn.Open()

try {
    $conn.EnableExtensions($true)
    $conn.LoadExtension("C:\Users\USUARIO\AppData\Local\Temp\crsqlite_esap.dll", "sqlite3_crsqlite_init")
    Write-Host "crsqlite loaded."
} catch {
    Write-Host "Could not load crsqlite: $_"
}

$cmd = New-Object System.Data.SQLite.SQLiteCommand("SELECT type, name, sql FROM sqlite_master WHERE name LIKE '%usuarios%';", $conn)
$reader = $cmd.ExecuteReader()

while ($reader.Read()) {
    Write-Host "--- $($reader['type']) : $($reader['name']) ---"
    Write-Host $reader['sql']
    Write-Host ""
}

$reader.Close()

# Let's also test the crsql_changes query to see if it throws!
try {
    $cmd2 = New-Object System.Data.SQLite.SQLiteCommand("SELECT * FROM crsql_changes WHERE `"table`" = 'usuarios' LIMIT 1;", $conn)
    $reader2 = $cmd2.ExecuteReader()
    Write-Host "Query to crsql_changes successful"
    $reader2.Close()
} catch {
    Write-Host "Error querying crsql_changes: $_"
}

$conn.Close()
