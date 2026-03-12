Attribute VB_Name = "PNGtoText"
' ============================================================================
' CONVERTIR ARCHIVO PNG A BASE64 EN VBA - GENERA FUNCIÓN GetFirma()
' Versión corregida: Evita error de compilación por constantes largas
' ============================================================================

Option Explicit

' APIs para codificación Base64

    Private Declare PtrSafe Function CryptBinaryToString Lib "Crypt32" _
        Alias "CryptBinaryToStringA" ( _
        ByRef pbBinary As Byte, _
        ByVal cbBinary As Long, _
        ByVal dwFlags As Long, _
        ByVal pszString As String, _
        ByRef pcchString As Long) As Long
        
    Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" _
        Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Const CRYPT_STRING_BASE64 As Long = &H1
Private Const CRYPT_STRING_BASE64HEADER As Long = &H0

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
End Type


' ============================================================================
' PROCEDIMIENTO PRINCIPAL: GENERAR ARCHIVO CON FUNCIÓN GetFirma()
' ============================================================================
Sub GenerarFuncionGetFirma()
    Dim rutaSeleccionada As String
    Dim firmaBase64 As String
    Dim nombreArchivo As String
    Dim nombreFuncion As String
    Dim rutaSalida As String
    Dim numArchivo As Integer
    Dim i As Long
    Dim numLinea As Integer
    Const MAX_LINEA As Long = 200  ' Caracteres por línea en el código VBA
    
    ' Seleccionar archivo
    rutaSeleccionada = SeleccionarArchivo( _
        titulo:="Seleccionar imagen PNG para convertir a función VBA", _
        filtro:="Imágenes PNG (*.png)|*.png|Todos los archivos (*.*)|*.*")
    
    If rutaSeleccionada = "" Then
        MsgBox "No se seleccionó ningún archivo.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener nombres
    nombreArchivo = Mid(rutaSeleccionada, InStrRev(rutaSeleccionada, "\") + 1)
    nombreArchivo = Left(nombreArchivo, InStrRev(nombreArchivo, ".") - 1)
    
    ' Crear nombre de función válido (quitar espacios, caracteres especiales)
    nombreFuncion = "Get" & LimpiarNombreFuncion(nombreArchivo)
    
    rutaSalida = Left(rutaSeleccionada, InStrRev(rutaSeleccionada, "\")) & nombreFuncion & ".txt"
    
    ' Convertir a Base64
    Debug.Print "Convirtiendo archivo a Base64..."
    firmaBase64 = ArchivoABase64(rutaSeleccionada)
    
    If firmaBase64 = "" Then
        MsgBox "ERROR: No se pudo convertir el archivo a Base64", vbCritical
        Exit Sub
    End If
    
    Debug.Print "Longitud Base64: " & Len(firmaBase64) & " caracteres"
    
    ' Generar archivo con función
    On Error GoTo ErrorHandler
    numArchivo = FreeFile
    Open rutaSalida For Output As #numArchivo
    
    ' Encabezado
    Print #numArchivo, "' ============================================================================"
    Print #numArchivo, "' FUNCIÓN VBA: " & nombreFuncion & "()"
    Print #numArchivo, "' Descripción: Retorna imagen " & nombreArchivo & ".png como String Base64"
    Print #numArchivo, "' Archivo origen: " & nombreArchivo & ".png"
    Print #numArchivo, "' Fecha generación: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    Print #numArchivo, "' Longitud Base64: " & Len(firmaBase64) & " caracteres"
    Print #numArchivo, "'"
    Print #numArchivo, "' USO:"
    Print #numArchivo, "'   Dim base64String As String"
    Print #numArchivo, "'   base64String = " & nombreFuncion & "()"
    Print #numArchivo, "'"
    Print #numArchivo, "' NOTA: Esta función evita el error de compilación por constantes largas"
    Print #numArchivo, "' ============================================================================"
    Print #numArchivo, ""
    Print #numArchivo, "Public Function " & nombreFuncion & "() As String"
    Print #numArchivo, "    Dim resultado As String"
    Print #numArchivo, "    resultado = """""
    Print #numArchivo, ""
    
    ' Generar líneas de concatenación
    numLinea = 0
    i = 1
    
    Do While i <= Len(firmaBase64)
        Dim finSegmento As Long
        Dim segmento As String
        
        finSegmento = i + MAX_LINEA - 1
        If finSegmento > Len(firmaBase64) Then finSegmento = Len(firmaBase64)
        
        segmento = Mid(firmaBase64, i, finSegmento - i + 1)
        
        ' Escapar comillas dobles en el segmento
        segmento = Replace(segmento, """", """""")
        
        Print #numArchivo, "    resultado = resultado & """ & segmento & """"
        
        numLinea = numLinea + 1
        i = i + MAX_LINEA
    Loop
    
    ' Cierre de función
    Print #numArchivo, ""
    Print #numArchivo, "    " & nombreFuncion & " = resultado"
    Print #numArchivo, "End Function"
    Print #numArchivo, ""
    Print #numArchivo, "' ============================================================================"
    Print #numArchivo, "' FIN DE " & nombreFuncion
    Print #numArchivo, "' ============================================================================"
    
    Close #numArchivo
    
    ' Éxito
    Debug.Print "=========================================="
    Debug.Print "ARCHIVO GENERADO EXITOSAMENTE:"
    Debug.Print rutaSalida
    Debug.Print "=========================================="
    Debug.Print "Líneas de código: " & numLinea + 10
    Debug.Print ""
    Debug.Print "Para usar:"
    Debug.Print "1. Copia el contenido del archivo a un módulo VBA"
    Debug.Print "2. Llama a: base64String = " & nombreFuncion & "()"
    
    ' Abrir archivo
    Shell "notepad.exe """ & rutaSalida & """", vbNormalFocus
    
    MsgBox "¡Archivo generado exitosamente!" & vbCrLf & vbCrLf & _
           "Función: " & nombreFuncion & "()" & vbCrLf & _
           "Ubicación: " & rutaSalida & vbCrLf & vbCrLf & _
           "Copia el contenido a un módulo VBA y usa:" & vbCrLf & _
           "  base64String = " & nombreFuncion & "()", vbInformation
           
    Exit Sub
    
ErrorHandler:
    MsgBox "ERROR: " & Err.Description, vbCritical
    On Error Resume Next
    Close #numArchivo
End Sub

' ============================================================================
' FUNCIÓN: MOSTRAR DIÁLOGO DE SELECCIÓN DE ARCHIVO
' ============================================================================
Function SeleccionarArchivo(ByVal titulo As String, _
                            Optional ByVal filtro As String = "Todos los archivos (*.*)|*.*") As String
    
    Dim ofn As OPENFILENAME
    Dim resultado As Long
    
    ofn.lStructSize = LenB(ofn)
    ofn.lpstrFile = String(260, vbNullChar)
    ofn.nMaxFile = 260
    ofn.lpstrFileTitle = String(260, vbNullChar)
    ofn.nMaxFileTitle = 260
    ofn.lpstrTitle = titulo
    ofn.lpstrFilter = Replace(filtro, "|", vbNullChar) & vbNullChar & vbNullChar
    ofn.flags = &H80000 Or &H4
    
    resultado = GetOpenFileName(ofn)
    
    If resultado <> 0 Then
        SeleccionarArchivo = Left(ofn.lpstrFile, InStr(ofn.lpstrFile, vbNullChar) - 1)
    Else
        SeleccionarArchivo = ""
    End If
End Function

' ============================================================================
' FUNCIÓN: ARCHIVO A BASE64 (sin saltos de línea)
' ============================================================================
Function ArchivoABase64(ByVal rutaArchivo As String) As String
    Dim datos() As Byte
    Dim base64 As String
    Dim longitud As Long
    Dim resultado As Long
    Dim numArchivo As Integer
    
    On Error GoTo ErrorHandler
    
    If Dir(rutaArchivo) = "" Then
        ArchivoABase64 = ""
        Exit Function
    End If
    
    ' Leer archivo binario
    numArchivo = FreeFile
    Open rutaArchivo For Binary As #numArchivo
    ReDim datos(0 To LOF(numArchivo) - 1)
    Get #numArchivo, , datos
    Close #numArchivo
    
    ' Primera llamada: obtener tamaño necesario
    resultado = CryptBinaryToString(datos(0), UBound(datos) + 1, CRYPT_STRING_BASE64, vbNullString, longitud)
    
    If resultado = 0 Then
        ArchivoABase64 = ""
        Exit Function
    End If
    
    ' Segunda llamada: obtener string Base64
    base64 = String(longitud, vbNullChar)
    resultado = CryptBinaryToString(datos(0), UBound(datos) + 1, CRYPT_STRING_BASE64, base64, longitud)
    
    If resultado = 0 Then
        ArchivoABase64 = ""
        Exit Function
    End If
    
    ' Limpiar null terminator y saltos de línea que agrega Windows
    base64 = Left(base64, InStr(base64, vbNullChar) - 1)
    base64 = Replace(base64, vbCr, "")
    base64 = Replace(base64, vbLf, "")
    
    ArchivoABase64 = base64
    
    Exit Function
    
ErrorHandler:
    ArchivoABase64 = ""
    On Error Resume Next
    Close #numArchivo
End Function

' ============================================================================
' FUNCIÓN AUXILIAR: Limpiar nombre para función VBA válida
' ============================================================================
Private Function LimpiarNombreFuncion(ByVal nombre As String) As String
    Dim resultado As String
    Dim i As Integer
    Dim char As String
    
    resultado = ""
    
    For i = 1 To Len(nombre)
        char = Mid(nombre, i, 1)
        
        ' Permitir letras, números (excepto al inicio) y underscore
        Select Case char
            Case "a" To "z", "A" To "Z"
                resultado = resultado & char
            Case "0" To "9"
                If resultado <> "" Then resultado = resultado & char
            Case " ", "-", "_", ".", "(", ")", "[", "]", "{", "}"
                ' Convertir separadores a underscore
                If resultado <> "" And Right(resultado, 1) <> "_" Then
                    resultado = resultado & "_"
                End If
            Case Else
                ' Ignorar otros caracteres especiales
        End Select
    Next i
    
    ' Eliminar underscore final si existe
    If Right(resultado, 1) = "_" Then
        resultado = Left(resultado, Len(resultado) - 1)
    End If
    
    ' Si quedó vacío, usar nombre genérico
    If resultado = "" Then resultado = "Firma"
    
    LimpiarNombreFuncion = resultado
End Function

