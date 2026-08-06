Attribute VB_Name = "LogCambioRepositorio"

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: LogCambioRepositorio.bas (REFACTORIZADO)
' ==========================================================================
Public Sub GuardarLog(ByRef objLog As LogCambio, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    On Error GoTo Errores
    
    ' Gestión de conexión dual
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    objLog.idLogCambio = getSiguienteIDLogCambios(dbConexion)
    
    Set rcd = dbConexion.OpenRecordset("tbLogCambios", dbOpenDynaset)
    rcd.AddNew
    Call RepositorioComun.RellenarRecordsetDesdeObjeto(rcd, objLog)
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogCambioRepositorio.GuardarLog"
    errObj.Raise
End Sub

Private Function getSiguienteIDLogCambios(ByRef db As DAO.Database) As Long
    Dim rcd As DAO.Recordset
    Dim maxID As Long
    On Error GoTo Errores
    
    Set rcd = db.OpenRecordset("SELECT Max(idLogCambio) AS MaxID FROM tbLogCambios")
    If Not rcd.EOF Then maxID = Nz(rcd!maxID, 0)
    getSiguienteIDLogCambios = maxID + 1
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
    
Errores:
    
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogCambioRepositorio.getSiguienteIDLogCambios"
    errObj.Raise
End Function

Public Function getLogCambioPorID(ByVal id As Long, Optional ByRef db As DAO.Database) As LogCambio
    Dim sql As String
    Dim params As Object
    Dim dbConexion As DAO.Database
    On Error GoTo Errores
    
    ' Gestión de conexión dual
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbLogCambios WHERE idLogCambio = [p_ID];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", id
    
    Set getLogCambioPorID = RepositorioComun.HidratarEntidadDesdeSQL(sql, "LogCambio", dbConexion, params)
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogCambioRepositorio.getLogCambioPorID"
    errObj.Raise
End Function

Public Sub EliminarPorSolicitud(ByVal idSol As Long, ByRef db As DAO.Database)
    ' RESPONSABILIDAD: Eliminar los logs de auditoría asociados a esta solicitud
    Dim sql As String
    Dim params As Object
    On Error GoTo Errores
    
    ' Borramos los logs donde la tabla sea 'tbSolicitudes' y el registro sea el ID de la solicitud
    sql = "DELETE FROM tbLogCambios WHERE tabla = 'tbSolicitudes' AND registro = [p_ID];"
    
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", idSol
    
    Call RepositorioComun.EjecutarAccion(sql, params, db)
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogCambioRepositorio.EliminarPorSolicitud"
    errObj.Raise
End Sub

Public Function GetCambiosPorIntervalo(ByVal idSolicitud As Long, ByVal fechaDesde As Date, Optional ByRef db As DAO.Database) As DAO.Recordset
    ' RESPONSABILIDAD: Obtener cambios realizados desde el último rechazo.
    ' Filtra por registro = idSolicitud.
    
    Dim sql As String
    Dim fDesdeStr As String
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    ' Gestión de conexión dual
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    ' Formato fecha Access (US)
    fDesdeStr = Format(fechaDesde, "mm/dd/yyyy hh:nn:ss")
    
    sql = "SELECT * FROM tbLogCambios " & _
          "WHERE registro = " & idSolicitud & " " & _
          "AND fechaHora >= #" & fDesdeStr & "# " & _
          "ORDER BY fechaHora DESC;"
          
    Set GetCambiosPorIntervalo = dbConexion.OpenRecordset(sql, dbOpenSnapshot)
    
    Exit Function

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogCambioRepositorio.GetCambiosPorIntervalo"
    errObj.Raise
End Function


