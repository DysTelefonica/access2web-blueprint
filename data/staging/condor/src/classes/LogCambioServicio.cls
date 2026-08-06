VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogCambioServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: LogCambioServicio.cls
' RESPONSABILIDAD: Lógica de negocio para la gestión de logs de cambios técnicos.
' ==========================================================================

Public Sub LimpiarHistorialCambios(ByVal idSolicitud As Long, ByRef db As DAO.Database)
    ' Responsabilidad: Orquestar la eliminación de todos los registros de cambios de una solicitud.
    On Error GoTo Errores
    
    ' Delegar en el repositorio (Capa de datos)
    Call LogCambioRepositorio.EliminarPorSolicitud(idSolicitud, db)
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogCambioServicio.LimpiarHistorialCambios"
    errObj.Raise
End Sub

