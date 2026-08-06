VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidacionRevisionServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: ValidacionRevisionServicio.cls
' RESPONSABILIDAD: Gestión del ciclo de vida de validaciones y revisiones (RAC).
' ==========================================================================

Public Function GetUltimoOrdinal(ByVal idSolicitud As Long) As Long
    On Error GoTo Errores
    GetUltimoOrdinal = ValidacionRevisionRepositorio.GetUltimoOrdinal(idSolicitud)
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionServicio.GetUltimoOrdinal"
    errObj.Raise
End Function

Public Sub RegistrarEnvioRAC(ByVal idSolicitud As Long, ByVal fechaEnvio As Variant, ByVal usuario As String)
    On Error GoTo Errores
    Call ValidacionRevisionRepositorio.GuardarHitoEnvio(idSolicitud, fechaEnvio, usuario)
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionServicio.RegistrarEnvioRAC"
    errObj.Raise
End Sub

Public Sub RegistrarRespuestaRAC(ByVal idSolicitud_o_IdHito As Long, ByVal resultado As String, ByVal comentarios As String, ByVal fechaRecepcion As Variant, Optional ByVal esIdHitoDirecto As Boolean = False)
    Dim idHito As Long
    On Error GoTo Errores
    
    If esIdHitoDirecto Then
        idHito = idSolicitud_o_IdHito
    Else
        idHito = ValidacionRevisionRepositorio.GetUltimoPendienteId(idSolicitud_o_IdHito)
    End If
    
    If idHito > 0 Then
        Call ValidacionRevisionRepositorio.CerrarHitoRespuesta(idHito, resultado, comentarios, fechaRecepcion)
    End If
    
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionServicio.RegistrarRespuestaRAC"
    errObj.Raise
End Sub

Public Sub RegistrarRespuestaRACPorId(ByVal idValidacion As Long, ByVal resultado As String, ByVal comentarios As String)
    On Error GoTo Errores
    Call ValidacionRevisionRepositorio.ActualizarRespuestaPorId(idValidacion, resultado, comentarios)
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionServicio.RegistrarRespuestaRACPorId"
    errObj.Raise
End Sub

Public Sub RegistrarBorrador(ByVal idSolicitud As Long, ByVal idAdjunto As Long, ByVal usuario As String)
    On Error GoTo Errores
    Call ValidacionRevisionRepositorio.GuardarHitoBorrador(idSolicitud, idAdjunto, usuario)
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionServicio.RegistrarBorrador"
    errObj.Raise
End Sub

Public Sub LimpiarPorSolicitud(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database)
    On Error GoTo Errores
    Call ValidacionRevisionRepositorio.EliminarPorSolicitud(idSolicitud, db)
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionServicio.LimpiarPorSolicitud"
    errObj.Raise
End Sub
Public Function GetUltimoPendienteId(ByVal idSolicitud As Long) As Long
    On Error GoTo Errores
    GetUltimoPendienteId = ValidacionRevisionRepositorio.GetUltimoPendienteId(idSolicitud)
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionServicio.GetUltimoPendienteId"
    errObj.Raise
End Function


