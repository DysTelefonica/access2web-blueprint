VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RechazoServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: RechazoServicio.cls
' RESPONSABILIDAD: Lógica de negocio para la gestión de rechazos técnicos.
' ==========================================================================

Public Sub LimpiarHistorialRechazos(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database)
    ' Responsabilidad: Orquestar la eliminación de todos los rechazos de una solicitud.
    Dim dbConexion As DAO.Database
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    ' Delegar en el repositorio (Capa de datos)
    Call RechazoRepositorio.EliminarRechazosPorSolicitud(idSolicitud, dbConexion)
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RechazoServicio.LimpiarHistorialRechazos"
    errObj.Raise
End Sub

Public Function GetUltimoRechazoActivo(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As rechazo
    Dim dbConexion As DAO.Database
    On Error GoTo Errores
    
    If db Is Nothing Then
        Set dbConexion = getdb()
    Else
        Set dbConexion = db
    End If
    
    Set GetUltimoRechazoActivo = RechazoRepositorio.GetUltimoRechazoActivo(idSolicitud)
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RechazoServicio.GetUltimoRechazoActivo"
    errObj.Raise
End Function

Public Sub DesactivarRechazosPrevios(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database)
    Dim dbConexion As DAO.Database
    On Error GoTo Errores
    
    If db Is Nothing Then
        Set dbConexion = getdb()
    Else
        Set dbConexion = db
    End If
    
    Call RechazoRepositorio.DesactivarRechazosPrevios(idSolicitud, dbConexion)
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RechazoServicio.DesactivarRechazosPrevios"
    errObj.Raise
End Sub

