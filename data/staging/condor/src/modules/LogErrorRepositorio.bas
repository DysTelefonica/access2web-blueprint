Attribute VB_Name = "LogErrorRepositorio"

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: LogErrorRepositorio.bas (REFACTORIZADO)
' ==========================================================================

'----- EN: LogErrorRepositorio.bas (REEMPLAZAR SUBRUTINA) -----

Public Sub GuardarError(ByRef objError As LogError, Optional ByRef db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    On Error GoTo Errores
    
    ' Fallback: usa conexión de logging independiente (getdb()).
    ' IMPORTANTE: semántica fire-and-forget — errores de logging se silencian
    ' deliberadamente para no romper el flujo principal (ver también GuardarDesdeCondorError).
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    objError.idLogError = getSiguienteIDLogErrores(dbConexion)
    
    Set rcd = dbConexion.OpenRecordset("tbLogErrores", dbOpenDynaset)
    rcd.AddNew
    
    ' --- REFACTORIZADO ---
    Call RepositorioComun.RellenarRecordsetDesdeObjeto(rcd, objError)
    ' --- FIN REFACTORIZADO ---
    
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
    
Errores:
    
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogErrorRepositorio.GuardarError"
    errObj.Raise
End Sub

Private Function getSiguienteIDLogErrores(ByRef db As DAO.Database) As Long
    Dim rcd As DAO.Recordset
    Dim maxID As Long
    On Error GoTo Errores
    
    Set rcd = db.OpenRecordset("SELECT Max(idLogError) AS MaxID FROM tbLogErrores")
    If Not rcd.EOF Then maxID = Nz(rcd!maxID, 0)
    getSiguienteIDLogErrores = maxID + 1
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
    
Errores:
    
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogErrorRepositorio.getSiguienteIDLogErrores"
    errObj.Raise
End Function

Public Function getLogErrorPorID(ByVal id As Long, Optional ByRef db As DAO.Database) As LogError
    Dim sql As String
    Dim params As Object
    Dim dbConexion As DAO.Database
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbLogErrores WHERE idLogError = [p_ID];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", id
    
    Set getLogErrorPorID = RepositorioComun.HidratarEntidadDesdeSQL(sql, "LogError", dbConexion, params)

    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogErrorRepositorio.getLogErrorPorID"
    errObj.Raise
End Function

' Convenience: crea LogError desde CondorError y lo persiste
Public Sub GuardarDesdeCondorError(ByRef condorErr As CondorError, Optional ByRef db As DAO.Database)
    On Error GoTo Errores
    
    Dim logErr As New LogError
    logErr.fechaHora = Now()
    logErr.usuario = IIf(m_ObjUsuarioActivo Is Nothing, "Sistema", m_ObjUsuarioActivo.nombre)
    logErr.modulo = condorErr.source
    logErr.procedimiento = IIf(InStrRev(condorErr.source, ".") > 0, Mid(condorErr.source, InStrRev(condorErr.source, ".") + 1), condorErr.source)
    logErr.numeroError = condorErr.Number
    logErr.descripcionError = condorErr.description
    logErr.contexto = condorErr.CallStackToString()
    
    Call GuardarError(logErr, db)
    Exit Sub
    
Errores:
    ' Si el logging falla, silenciamos — nunca debe romper el flujo
    On Error Resume Next
End Sub

