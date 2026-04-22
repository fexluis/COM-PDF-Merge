' =============================================
' RegisterXLL.vbs - Custom Action para Excel 365
' =============================================

Dim shell, addinPath, keyPath, i, keyName

Set shell = CreateObject("WScript.Shell")

' Ruta completa del .xll (AppData del usuario actual)
addinPath = shell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\AddIns\ConversorPDF.xll"

keyPath = "HKCU\Software\Microsoft\Office\16.0\Excel\Options"

i = 0
Do
    If i = 0 Then
        keyName = "OPEN"
    Else
        keyName = "OPEN" & CStr(i)
    End If

    On Error Resume Next
    dummy = shell.RegRead(keyPath & "\" & keyName)
    If Err.Number <> 0 Then
        ' Clave libre encontrada → la creamos
        On Error Goto 0
        shell.RegWrite keyPath & "\" & keyName, "/R """ & addinPath & """", "REG_SZ"
        Exit Do
    End If
    On Error Goto 0
    i = i + 1
Loop While i < 50   ' seguridad
