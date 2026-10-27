' =============================================
' RegisterXLL.vbs - Custom Action para Excel 365
' =============================================

Dim shell, addinPath, keyPath, i, keyName, existingVal, bRegistered

Set shell = CreateObject("WScript.Shell")

' Ruta completa del .xll (AppData del usuario actual)
addinPath = shell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\AddIns\ConversorPDF.xll"

keyPath = "HKCU\Software\Microsoft\Office\16.0\Excel\Options"

' Paso 1: buscar y actualizar una key existente que ya apunte a ConversorPDF
Dim firstFreeKey
firstFreeKey = ""
bRegistered = False
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
        ' Key libre — la guardamos como candidata y salimos del scan
        If firstFreeKey = "" Then firstFreeKey = keyName
        On Error GoTo 0
        Exit Do
    End If
    On Error GoTo 0

    ' ¿La key existente apunta a alguna versión de ConversorPDF?
    If InStr(1, existingVal, "conversorpdf", vbTextCompare) > 0 Then
        ' Actualizar en lugar de añadir nueva
        shell.RegWrite keyPath & "\" & keyName, "/R """ & addinPath & """", "REG_SZ"
        bRegistered = True
        Exit Do
    End If

    i = i + 1
Loop While i < 50

' Paso 2: no se encontró key de ConversorPDF — escribir en la primera key libre
If Not bRegistered And firstFreeKey <> "" Then
    shell.RegWrite keyPath & "\" & firstFreeKey, "/R """ & addinPath & """", "REG_SZ"
End If
