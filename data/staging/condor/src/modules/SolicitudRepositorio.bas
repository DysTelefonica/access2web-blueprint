Attribute VB_Name = "SolicitudRepositorio"
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: SolicitudRepositorio.bas (REFACTORIZADO COMPLETO CON CondorError)
' ==========================================================================

Public Function getSolicitudPorID(ByVal id As Long, Optional ByRef db As DAO.Database) As Solicitud
    Dim sql As String
    Dim params As Object
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbSolicitudes WHERE idSolicitud = [p_ID];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", id
    
    Set getSolicitudPorID = RepositorioComun.HidratarEntidadDesdeSQL(sql, "Solicitud", dbConexion, params)
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudRepositorio.getSolicitudPorID"
    errObj.Raise
End Function
Public Function getSolicitudPorCodigo(ByVal p_Codigo As String, Optional ByRef db As DAO.Database) As Solicitud
    Dim sql As String
    Dim params As Object
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbSolicitudes WHERE codigoSolicitud = [p_Codigo];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_Codigo", p_Codigo
    
    Set getSolicitudPorCodigo = RepositorioComun.HidratarEntidadDesdeSQL(sql, "Solicitud", dbConexion, params)
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudRepositorio.getSolicitudPorCodigo"
    errObj.Raise
End Function

' EN: SolicitudRepositorio.bas

Public Sub GuardarSolicitud(ByRef objSolicitud As Solicitud, Optional db As DAO.Database)
    Dim rcd As DAO.Recordset
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    If objSolicitud.idSolicitud > 0 And RegistroExiste("tbSolicitudes", "idSolicitud", objSolicitud.idSolicitud, dbConexion) Then
        ' UPDATE
        Dim qdf As DAO.QueryDef
        Set qdf = dbConexion.CreateQueryDef("", "SELECT * FROM tbSolicitudes WHERE idSolicitud = [p_ID]")
        qdf.Parameters("p_ID").value = objSolicitud.idSolicitud
        Set rcd = qdf.OpenRecordset(dbOpenDynaset)
        
        If rcd.EOF Then Err.Raise 513, , "No se encontró el registro para actualizar."
        
        rcd.Edit
        Call RepositorioComun.RellenarRecordsetDesdeObjeto(rcd, objSolicitud, True)
    Else
        ' INSERT
        Set rcd = dbConexion.OpenRecordset("tbSolicitudes", dbOpenDynaset)
        rcd.AddNew
        Call RepositorioComun.RellenarRecordsetDesdeObjeto(rcd, objSolicitud, False)
    End If
    
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudRepositorio.GuardarSolicitud"
    ' CORRECCIÓN: Añadimos contexto vital para depuración
    errObj.AddToCallStack "ID Solicitud: " & objSolicitud.idSolicitud
    errObj.AddToCallStack "Código: " & objSolicitud.codigoSolicitud
    errObj.Raise
End Sub

Public Function getSiguienteOrdinalParaClave(ByVal claveExpediente As String, Optional ByRef db As DAO.Database) As Long
    Dim qdf As DAO.QueryDef
    Dim rcd As DAO.Recordset
    Dim sql As String
    Dim maxOrdinal As Long
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT Max(Val(Mid(codigoSolicitud, InStrRev(codigoSolicitud, '-') + 1))) AS MaxOrdinal " & _
          "FROM tbSolicitudes WHERE codigoSolicitud LIKE [p_Patron];"
    Set qdf = dbConexion.CreateQueryDef("", sql)
    qdf.Parameters("p_Patron").value = "DC-" & claveExpediente & "-*"
    
    Set rcd = qdf.OpenRecordset(dbOpenSnapshot)
    
    If Not rcd.EOF Then maxOrdinal = Nz(rcd!maxOrdinal, 0)
    
    getSiguienteOrdinalParaClave = maxOrdinal + 1

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
    
Errores:
   
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudRepositorio.getSiguienteOrdinalParaClave"
    errObj.Raise
End Function

Public Function ExisteCodigo(ByVal codigo As String, Optional ByRef db As DAO.Database) As Boolean
    Dim qdf As DAO.QueryDef
    Dim rcd As DAO.Recordset
    Dim dbConexion As DAO.Database
    Dim sql As String
    
    On Error GoTo Errores
    ExisteCodigo = False
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT COUNT(*) AS n FROM tbSolicitudes WHERE codigoSolicitud = [p_Codigo];"
    Set qdf = dbConexion.CreateQueryDef("", sql)
    qdf.Parameters("p_Codigo").value = codigo
    
    Set rcd = qdf.OpenRecordset(dbOpenDynaset)
    
    If Not rcd.EOF Then
        If rcd!n > 0 Then ExisteCodigo = True
    End If

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudRepositorio.ExisteCodigo"
    errObj.Raise
End Function

Public Function ActualizarEstado(ByVal sol As Solicitud, Optional db As DAO.Database) As Boolean
    Dim dbConexion As DAO.Database
    Dim sqlIns As String
    Dim p As Object
    
    On Error GoTo Errores
    ActualizarEstado = False
    
    ' Gestión de conexión dual
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sqlIns = "UPDATE tbSolicitudes " & _
            "SET idEstadoInterno = [p_idestadofinal], " & _
            "fechaModificacion = [p_fechaModificacion], " & _
            "usuarioModificacion = [p_usuarioModificacion] " & _
            "WHERE idSolicitud = [p_id];"
            
    Set p = CreateObject("Scripting.Dictionary")
    p.Add "p_idestadofinal", sol.idEstadoInterno
    p.Add "p_fechaModificacion", sol.fechaModificacion
    p.Add "p_usuarioModificacion", sol.usuarioModificacion
    p.Add "p_id", sol.idSolicitud
    
    ' Usamos dbConexion
    RepositorioComun.EjecutarAccion sqlIns, p, dbConexion
    ActualizarEstado = True

LimpiarYSalir:
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudRepositorio.ActualizarEstado"
    errObj.Raise
End Function

Public Sub Eliminar(ByVal idSol As Long, ByRef db As DAO.Database)
    ' RESPONSABILIDAD: Eliminar el registro principal de tbSolicitudes.
    Dim sql As String
    Dim params As Object
    On Error GoTo Errores
    
    sql = "DELETE FROM tbSolicitudes WHERE idSolicitud = [p_IDSol];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_IDSol", idSol
    
    Call RepositorioComun.EjecutarAccion(sql, params, db)
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudRepositorio.Eliminar"
    errObj.AddToCallStack "idSolicitud: " & idSol
    errObj.Raise
End Sub


