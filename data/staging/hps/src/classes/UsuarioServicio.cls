VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsuarioServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_TestDb As DAO.Database
Private m_TestAttachmentsRootActual As String
Private m_TestAttachmentsRootHistorico As String
Private m_TestFailureStage As String

Public Sub ConfigureLifecycleTransactionTestSeams(Optional ByRef p_Db As DAO.Database = Nothing, _
                                                  Optional ByVal p_AttachmentsRootActual As String = "", _
                                                  Optional ByVal p_AttachmentsRootHistorico As String = "", _
                                                  Optional ByVal p_FailureStage As String = "", _
                                                  Optional ByRef p_Error As String)
    On Error GoTo errores

    Set m_TestDb = p_Db
    m_TestAttachmentsRootActual = p_AttachmentsRootActual
    m_TestAttachmentsRootHistorico = p_AttachmentsRootHistorico
    m_TestFailureStage = p_FailureStage
    Exit Sub
errores:
    p_Error = "El método UsuarioServicio.ConfigureLifecycleTransactionTestSeams ha devuelto el error: " & Err.Description
End Sub

Public Function PasarDeActualAHistorico(p_IDActual As String, Optional ByRef p_Error As String) As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    On Error GoTo errores

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarActualAHistorico(p_IDActual, m_TestDb, m_TestAttachmentsRootActual, m_TestAttachmentsRootHistorico, m_TestFailureStage, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
limpiar:
    On Error Resume Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioServicio.PasarDeActualAHistorico ha devuelto el error: " & Err.Description
    End If
    GoTo limpiar
End Function

Public Function PasarDeHistoricoAActual(p_IDHistorico As String, Optional ByRef p_Error As String) As String
    Dim coordinator As UsuarioLifecycleTransactionCoordinator
    On Error GoTo errores

    Set coordinator = New UsuarioLifecycleTransactionCoordinator
    Call coordinator.PasarHistoricoAActual(p_IDHistorico, m_TestDb, m_TestAttachmentsRootActual, m_TestAttachmentsRootHistorico, m_TestFailureStage, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
limpiar:
    On Error Resume Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioServicio.PasarDeHistoricoAActual ha devuelto el error: " & Err.Description
    End If
    GoTo limpiar
End Function

Public Function URLCarpetaAnexosActuales() As String
    ' [config-as-source-of-truth] Inyección explícita gana (test seam);
    ' si no, delega a m_ObjEntorno.URLCarpetaAnexos que ya implementa
    ' el selector APP_ROOT_REMOTO/APP_ROOT_LOCAL según accesibilidad.
    If Len(m_TestAttachmentsRootActual) > 0 Then
        URLCarpetaAnexosActuales = EnsureServicioSlash(m_TestAttachmentsRootActual)
        Exit Function
    End If
    URLCarpetaAnexosActuales = m_ObjEntorno.URLCarpetaAnexos
End Function

Public Function URLCarpetaAnexosHistoricos() As String
    ' [config-as-source-of-truth] Mirror de URLCarpetaAnexosActuales.
    If Len(m_TestAttachmentsRootHistorico) > 0 Then
        URLCarpetaAnexosHistoricos = EnsureServicioSlash(m_TestAttachmentsRootHistorico)
        Exit Function
    End If
    URLCarpetaAnexosHistoricos = m_ObjEntorno.URLCarpetaAnexosHistoricos
End Function
Public Function URLCarpetaAnexosActualesPorDNI(p_DNI As String) As String
    URLCarpetaAnexosActualesPorDNI = URLCarpetaAnexosActuales() & p_DNI & "\"
End Function
Public Function URLCarpetaAnexosHistoricosPorDNI(p_DNI As String) As String
    URLCarpetaAnexosHistoricosPorDNI = URLCarpetaAnexosHistoricos() & p_DNI & "\"
    
    
End Function

Private Function EnsureServicioSlash(ByVal folderPath As String) As String
    If Len(folderPath) = 0 Then Exit Function
    If Right$(folderPath, 1) = "\" Then
        EnsureServicioSlash = folderPath
    Else
        EnsureServicioSlash = folderPath & "\"
    End If
End Function


Public Function FsoDeleteFolder(ByVal folderPath As String, Optional ByRef errorMsg As String) As Boolean
    On Error GoTo ErrHandler
    Dim fso As Object ' Scripting.FileSystemObject (late binding)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Right(folderPath, 1) = "\" Then
        folderPath = Left(folderPath, Len(folderPath) - 1)
    End If
    ' Si no existe, lo consideramos "eliminado"
    If Not fso.FolderExists(folderPath) Then
        FsoDeleteFolder = True
        Exit Function
    End If

    ' Limpia atributos en todo el árbol para evitar errores de permiso
    ClearAttributesRecursive folderPath

    ' Borrar carpeta (True = Forzar, también elimina archivos de solo lectura)
    fso.DeleteFolder folderPath, True

    ' Verificar que realmente desapareció
    If Not fso.FolderExists(folderPath) Then
        FsoDeleteFolder = True
    Else
        errorMsg = "No se pudo eliminar la carpeta por una causa desconocida."
        FsoDeleteFolder = False
    End If
    Exit Function

ErrHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description & vbCrLf & "Carpeta: " & folderPath
    FsoDeleteFolder = False
End Function


' --- Helpers ---

Private Sub ClearAttributesRecursive(ByVal folderPath As String)
    On Error GoTo ErrHandler
    Dim fso As Object, fld As Object, subFld As Object, fil As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fld = fso.GetFolder(folderPath)

    ' Quitar atributos de archivos en esta carpeta
    For Each fil In fld.Files
        SafeClearFileAttributes fil
    Next fil

    ' Procesar subcarpetas primero (profundidad)
    For Each subFld In fld.SubFolders
        ClearAttributesRecursive subFld.path
    Next subFld

    ' Quitar atributos de la carpeta actual (por si está marcada como oculta/sistema/solo lectura)
    SafeClearFolderAttributes fld
    Exit Sub

ErrHandler:
    ' Si algún elemento falla (bloqueado en uso), seguimos intentando con el resto.
    Resume Next
End Sub

Private Sub SafeClearFileAttributes(ByVal f As Object)
    On Error Resume Next
    ' Quitar ReadOnly, Hidden y System (dejar otros como Archive si quieres)
    Dim attrs As Long
    attrs = f.Attributes
    attrs = attrs And Not vbReadOnly
    attrs = attrs And Not vbHidden
    attrs = attrs And Not vbSystem
    f.Attributes = attrs
End Sub

Private Sub SafeClearFolderAttributes(ByVal d As Object)
    On Error Resume Next
    Dim attrs As Long
    attrs = d.Attributes
    attrs = attrs And Not vbReadOnly
    attrs = attrs And Not vbHidden
    attrs = attrs And Not vbSystem
    d.Attributes = attrs
End Sub


