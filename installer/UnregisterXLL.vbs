' =============================================
' UnregisterXLL.vbs - Custom Action para Excel 365
' =============================================

Dim shell, keyPath, i, keyName, existingVal

Set shell = CreateObject("WScript.Shell")

keyPath = "HKCU\Software\Microsoft\Office\16.0\Excel\Options"

i = 0
Do
    If i = 0 Then
        keyName = "OPEN"
    Else
        keyName = "OPEN" & CStr(i)
    End If

    On Error Resume Next
    existingVal = shell.RegRead(keyPath & "\" & keyName)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Do
    End If
    On Error GoTo 0

    If InStr(1, existingVal, "conversorpdf", vbTextCompare) > 0 Then
        On Error Resume Next
        shell.RegDelete keyPath & "\" & keyName
        On Error GoTo 0
    End If

    i = i + 1
Loop While i < 50
