$files = Get-ChildItem -Path 'ConversorPDF.XLL' -Filter '*Repository.cs' -Recurse
foreach ($f in $files) {
    $content = Get-Content $f.FullName -Raw
    $newContent = $content -replace 'return (SqliteResult<int>\.Success\([^\)]+\));', 'ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty(); return $1;'
    if ($content -ne $newContent) {
        Set-Content -Path $f.FullName -Value $newContent -NoNewline
        Write-Host "Updated $($f.Name)"
    }
}
