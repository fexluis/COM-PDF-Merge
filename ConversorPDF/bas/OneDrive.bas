Attribute VB_Name = "OneDrive"
'*****************************************************************************
' GetLocalPath
'*****************************************************************************
'Attribute VB_Name = "GetLocalPath"
'
' Cross-platform VBA Function to get the local path of OneDrive/SharePoint
' synchronized Microsoft Office files (Works on Windows and on macOS)
'
' Author: Guido Witt-Dörring
' Created: 2022/07/01
' Updated: 2024/07/08
' License: MIT
'
' ----------------------------------------------------------------
' https://gist.github.com/guwidoe/038398b6be1b16c458365716a921814d
' https://stackoverflow.com/a/73577057/12287457
' ----------------------------------------------------------------
'
' Copyright (c) 2024 Guido Witt-Dörring
'
' ============================================================================
' FUNCIÓN PRINCIPAL: GetOneDriveLocalPath CORREGIDA
' ============================================================================

' ============================================================================
' FUNCIÓN PRINCIPAL: GetOneDriveLocalPath CORREGIDA
' ============================================================================

Public Function GetOneDriveLocalPath(ByVal path As String) As String
    ' ========================================
    ' VERSIÓN OPTIMIZADA - MÁXIMA VELOCIDAD
    ' ========================================
    
    Dim resultado As String
    Dim pos As Long
    Dim oneDriveFullPath As String
    Dim rutaRelativa As String
    
    ' CASO 1: Ya es ruta local (más común) - Salida rápida
    If Mid(path, 2, 1) = ":" Then  ' Formato X: más rápido que Like
        GetOneDriveLocalPath = path
        Exit Function
    End If
    
    ' CASO 2: No es URL de SharePoint/OneDrive - Devolver original
    If InStr(1, path, "sharepoint", vbTextCompare) = 0 And _
       InStr(1, path, "onedrive", vbTextCompare) = 0 Then
        GetOneDriveLocalPath = path
        Exit Function
    End If
    
    ' ========================================
    ' CASO 3: URL de SharePoint/OneDrive - Requiere conversión
    ' ========================================
    
    ' Obtener ruta base de OneDrive (registro es más rápido que buscar carpetas)
    oneDriveFullPath = GetOneDrivePathFromRegistryFast()
    
    ' Fallback solo si es necesario
    If oneDriveFullPath = "" Then
        oneDriveFullPath = GetOneDrivePathFromFolderFast(Environ("UserName"))
    End If
    
    ' Si no se puede determinar, devolver original
    If oneDriveFullPath = "" Then
        GetOneDriveLocalPath = path
        Exit Function
    End If
    
    ' Buscar punto de corte en URL (solo las variantes más comunes)
    pos = InStr(1, path, "/Documents/", vbTextCompare)
    If pos = 0 Then pos = InStr(1, path, "/documentos/", vbTextCompare)
    
    If pos > 0 Then
        ' Extraer ruta relativa (pos + 11 = después de "/Documents/")
        rutaRelativa = Mid(path, pos + 11)
        
        ' Reemplazos en una sola pasada donde sea posible
        rutaRelativa = Replace(Replace(rutaRelativa, "/", "\"), "%20", " ")
        
        ' Eliminar barra inicial si existe
        If Left(rutaRelativa, 1) = "\" Then rutaRelativa = Mid(rutaRelativa, 2)
        
        resultado = oneDriveFullPath & "\" & rutaRelativa
    Else
        ' Fallback: extraer desde el último /
        pos = InStrRev(path, "/")
        If pos > 0 Then
            rutaRelativa = Replace(Mid(path, pos + 1), "%20", " ")
            resultado = oneDriveFullPath & "\" & rutaRelativa
        Else
            resultado = oneDriveFullPath
        End If
    End If
    
    ' Limpiar dobles barras (una sola pasada)
    resultado = Replace(resultado, "\\", "\")
    
    GetOneDriveLocalPath = resultado
End Function

' ============================================================================
' VERSIÓN RÁPIDA: Obtiene ruta desde registro (solo claves esenciales)
' ============================================================================

Private Function GetOneDrivePathFromRegistryFast() As String
    Dim wsh As Object
    Dim oneDrivePath As String
    
    On Error Resume Next
    Set wsh = CreateObject("WScript.Shell")
    
    ' Solo las 2 claves más comunes (Business1 y Personal)
    oneDrivePath = wsh.RegRead("HKCU\Software\Microsoft\OneDrive\Accounts\Business1\UserFolder")
    If oneDrivePath = "" Then
        oneDrivePath = wsh.RegRead("HKCU\Software\Microsoft\OneDrive\Accounts\Personal\UserFolder")
    End If
    
    GetOneDrivePathFromRegistryFast = oneDrivePath
    Set wsh = Nothing
End Function

' ============================================================================
' VERSIÓN RÁPIDA: Obtiene ruta desde carpeta (solo verifica existencia)
' ============================================================================

Private Function GetOneDrivePathFromFolderFast(ByVal usuario As String) As String
    Dim fso As Object
    Dim userPath As String
    Dim folder As Object
    Dim subFolder As Object
    
    On Error Resume Next
    
    userPath = "C:\Users\" & usuario
    If Dir(userPath, vbDirectory) = "" Then Exit Function
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(userPath)
    
    ' Buscar solo la primera carpeta OneDrive comercial (con " - ")
    For Each subFolder In folder.SubFolders
        If Left(subFolder.Name, 8) = "OneDrive" Then
            If InStr(subFolder.Name, " - ") > 0 Then
                GetOneDrivePathFromFolderFast = subFolder.path
                GoTo CleanExit
            End If
        End If
    Next subFolder
    
    ' Si no hay comercial, buscar personal
    For Each subFolder In folder.SubFolders
        If Left(subFolder.Name, 8) = "OneDrive" Then
            GetOneDrivePathFromFolderFast = subFolder.path
            Exit For
        End If
    Next subFolder

CleanExit:
    Set fso = Nothing
End Function

