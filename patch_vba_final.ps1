$path = 'ConversorPDF.XLL\bas\Via_Cumplidos.bas'
if (-not (Test-Path $path)) {
    Write-Error "File not found: $path"
    exit 1
}

$lines = Get-Content $path -Encoding Default
$startLine = 641
$endLine = 675 # Increased range to find the whole function if needed

$startIdx = $startLine - 1
$foundStart = -1
$foundEnd = -1

for ($i = $startIdx; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match "Public Function ExtraerDespuesPagoViaticos") {
        $foundStart = $i
    }
    if ($foundStart -ne -1 -and $lines[$i] -match "End Function") {
        $foundEnd = $i
        break
    }
}

if ($foundStart -eq -1) {
    # If not found renamed, try original name
    for ($i = $startIdx; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match "Public Function ExtraerDespuesDe26") {
            $foundStart = $i
        }
        if ($foundStart -ne -1 -and $lines[$i] -match "End Function") {
            $foundEnd = $i
            break
        }
    }
}

if ($foundStart -eq -1) {
    Write-Error "Could not find function definition."
    exit 1
}

$newFunc = @"
Public Function ExtraerDespuesPagoViaticos(ByVal cadena As String) As String
    If Len(cadena) = 0 Then
        ExtraerDespuesPagoViaticos = vbNullString
        Exit Function
    End If

    ' Configuración de la vigencia (últimos 2 dígitos)
    ' En C# se obtiene de App.APP_VIGENCIA.
    Const VIGENCIA_SUFFIX As String = "26"
    
    On Error Resume Next
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    On Error GoTo 0
    
    If regEx Is Nothing Then
        ' Fallback if Regex fails (should not happen on Windows)
        ExtraerDespuesPagoViaticos = LCase(Trim(cadena))
        Exit Function
    End If
    
    ' VITICOS, VIATICA, VIATICO, VAITICOS, VAITICO, VIATCIOS
    Dim viaticosPattern As String
    viaticosPattern = "(?:VIATICOS|VITICOS|VIATICA|VIATICO|VAITICOS|VAITICO|VIATCIOS)"
    
    With regEx
        .IgnoreCase = True
        .Global = False
        .Pattern = "PAGO\s+" & viaticosPattern & "(?:\s+\d*" & VIGENCIA_SUFFIX & ")?(?:\s+PARA)?\b"
    End With
    
    Dim matches As Object
    Set matches = regEx.Execute(cadena)
    
    If matches.Count > 0 Then
        Dim match As Object
        Set match = matches(0)
        ExtraerDespuesPagoViaticos = LCase(Trim(Mid(cadena, match.FirstIndex + match.Length + 1)))
    Else
        ExtraerDespuesPagoViaticos = LCase(Trim(cadena))
    End If
End Function
"@

$before = $lines[0..($foundStart-1)]
$after = $lines[($foundEnd+1)..($lines.Count-1)]

$combined = @()
if ($before) { $combined += $before }
$combined += $newFunc
if ($after) { $combined += $after }

$finalContent = $combined -join [Environment]::NewLine
$finalContent = $finalContent -replace 'ExtraerDespuesDe26\(', 'ExtraerDespuesPagoViaticos('

Set-Content $path -Value $finalContent -Encoding Default
Write-Host "VBA file updated correctly."
