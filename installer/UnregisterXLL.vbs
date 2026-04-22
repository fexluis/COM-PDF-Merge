' ===============================================
' UnregisterXLL.vbs - Quita el registro de Excel
' ===============================================

Dim shell, addinFileName, keyPath, i, keyName, val

Set shell = CreateObject("WScript.Shell")
addinFileName = "ConversorPDF.xll"
keyPath = "HKCU\Software\Microsoft\Office\16.0\Excel\Options"

i = 0
Do
    If i = 0 Then
        keyName = "OPEN"
    Else
        keyName = "OPEN" & CStr(i)
    End If

    On Error Resume Next
    val = ""
    val = shell.RegRead(keyPath & "\" & keyName)
    
    If Err.Number = 0 Then
        ' Clave existe, verificar si contiene nuestro Add-in
        If InStr(1, val, addinFileName, vbTextCompare) > 0 Then
            ' Encontrado -> Borrar clave
            shell.RegDelete keyPath & "\" & keyName
            ' También borramos la configuración propia
            shell.RegDelete "HKCU\Software\ESAP\ConversorPDF\"
            Exit Do
        End If
    Else
        ' No hay más claves OPEN correlativas
        Exit Do
    End If
    On Error Goto 0
    
    i = i + 1
Loop While i < 50
