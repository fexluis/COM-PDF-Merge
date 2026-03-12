Attribute VB_Name = "PNGtoUsar"
' ============================================================================
' MÓDULO: Mod_Base64Temp
' Descripción: Convierte Base64 a archivo temporal PNG
' ============================================================================

Option Explicit

' ============================================================================
' PROCEDIMIENTO DE PRUEBA
' ============================================================================
Sub TestCompleto()
    Dim rutaTemp As String
    Dim shp As Shape
    
    Debug.Print "=== INICIANDO TEST ==="
    
    ' Verificar que GetFirmaLuis no está vacía
    Dim testString As String
    testString = GetFirmaLuis()
    
    If Len(testString) < 100 Then
        MsgBox "ERROR: GetFirmaLuis() está vacía o no contiene datos" & vbCrLf & _
               "Longitud actual: " & Len(testString) & vbCrLf & vbCrLf & _
               "EJECUTA PRIMERO: GenerarFuncionGetFirma()" & vbCrLf & _
               "y copia la función generada a este módulo.", vbCritical
        Exit Sub
    End If
    
    ' Obtener archivo temporal
    rutaTemp = FirmaToTempFile("Luis")
    
    If rutaTemp = "" Then
        MsgBox "Error al crear archivo temporal", vbCritical
        Exit Sub
    End If
    
    ' Verificar archivo
    If Dir(rutaTemp) = "" Then
        MsgBox "Archivo no encontrado: " & rutaTemp, vbExclamation
        Exit Sub
    End If
    
    ' Insertar imagen
    On Error Resume Next
    Set shp = ActiveSheet.Shapes.AddPicture( _
        fileName:=rutaTemp, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=100, Top:=100, Width:=-1, Height:=-1)
    
    If Err.Number <> 0 Then
        MsgBox "Error al insertar imagen: " & Err.Description, vbExclamation
        Err.Clear
    Else
        shp.Name = "FirmaLuis"
        MsgBox "¡Firma insertada correctamente!" & vbCrLf & _
               "Archivo: " & rutaTemp & vbCrLf & _
               "Tamaño: " & FileLen(rutaTemp) & " bytes", vbInformation
    End If
    On Error GoTo 0
    
    ' Limpiar temporal
    Kill rutaTemp
End Sub

' ============================================================================
' FUNCIÓN PRINCIPAL: Obtiene FirmaLuis y crea archivo temporal
' ============================================================================
Public Function FirmaToTempFile(cadena As String) As String
    Dim firma As String
    If cadena = "MONTANO" Then
        firma = GetFirmaMontano()
    ElseIf cadena = "VILLEGAS" Then
        firma = GetFirmaVillegas()
    Else
        firma = GetFirmaLuis()
    End If
    FirmaToTempFile = Base64ToTempFile(firma, "frmx")
End Function

' ============================================================================
' FUNCIÓN GENÉRICA: Cualquier cadena Base64 a archivo temporal
' ============================================================================
Public Function Base64ToTempFile( _
    ByVal base64String As String, _
    Optional ByVal nombreBase As String = "Imagen") As String
    
    Dim bytes() As Byte
    Dim tempPath As String
    Dim fileNum As Integer
    
    On Error GoTo ErrorHandler
    
    ' Validar
    If Len(base64String) < 100 Then
        Debug.Print "ERROR: Cadena Base64 demasiado corta (" & Len(base64String) & ")"
        Base64ToTempFile = ""
        Exit Function
    End If
    
    ' Decodificar
    bytes = DecodeBase64(base64String)
    If UBound(bytes) < 0 Then
        Debug.Print "ERROR: Decodificación fallida"
        Base64ToTempFile = ""
        Exit Function
    End If
    
    ' Crear archivo temporal
    tempPath = CrearRutaTemporal(nombreBase, "png")
    
    fileNum = FreeFile
    Open tempPath For Binary Access Write As #fileNum
    Put #fileNum, 1, bytes
    Close #fileNum
    
    If Dir(tempPath) <> "" Then
        Base64ToTempFile = tempPath
    Else
        Base64ToTempFile = ""
    End If
    
    Exit Function
    
ErrorHandler:
    Base64ToTempFile = ""
    On Error Resume Next
    Close #fileNum
End Function

' ============================================================================
' FUNCIÓN: Decodificar Base64 usando MSXML
' ============================================================================
Private Function DecodeBase64(ByVal base64String As String) As Byte()
    Dim xmlDoc As Object
    Dim xmlNode As Object
    Dim bytes() As Byte
    
    On Error GoTo ErrorHandler
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    If xmlDoc Is Nothing Then Set xmlDoc = CreateObject("MSXML2.DOMDocument.3.0")
    If xmlDoc Is Nothing Then Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    
    Set xmlNode = xmlDoc.createElement("b64")
    xmlNode.DataType = "bin.base64"
    xmlNode.text = base64String
    bytes = xmlNode.NodeTypedValue
    
    DecodeBase64 = bytes
    Exit Function
    
ErrorHandler:
    ReDim bytes(0)
    DecodeBase64 = bytes
End Function

' ============================================================================
' FUNCIÓN: Crear ruta temporal única
' ============================================================================
Private Function CrearRutaTemporal(ByVal nombreBase As String, ByVal extension As String) As String
    Dim tempFolder As String
    Dim fileName As String
    
    tempFolder = Environ$("TEMP")
    If tempFolder = "" Then tempFolder = Environ$("TMP")
    If tempFolder = "" Then tempFolder = ThisWorkbook.path
    
    ' Limpiar caracteres inválidos
    nombreBase = Replace(nombreBase, "\", "")
    nombreBase = Replace(nombreBase, "/", "")
    nombreBase = Replace(nombreBase, ":", "")
    nombreBase = Replace(nombreBase, "*", "")
    nombreBase = Replace(nombreBase, "?", "")
    nombreBase = Replace(nombreBase, """", "")
    nombreBase = Replace(nombreBase, "<", "")
    nombreBase = Replace(nombreBase, ">", "")
    nombreBase = Replace(nombreBase, "|", "")
    
    fileName = nombreBase & "_" & Format(Now, "yyyymmdd_hhmmss") & "_" & Int(Rnd() * 10000)
    CrearRutaTemporal = tempFolder & "\" & fileName & "." & extension
End Function

