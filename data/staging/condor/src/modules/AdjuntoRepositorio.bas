Attribute VB_Name = "AdjuntoRepositorio"

Option Compare Database
Option Explicit

Public Sub Guardar(ByRef adj As Adjunto, Optional db As DAO.Database)
    Dim sql As String, params As Object, dbConexion As DAO.Database
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    adj.idAdjunto = getSiguienteIDAdjunto(dbConexion) ' Pasamos la conexión
    
    sql = "INSERT INTO tbAdjuntos (idAdjunto, idSolicitud, etapaWF, nombreArchivo, fechaSubida, usuarioSubida, descripcion, TipoAccion) " & _
          "VALUES ([p_idAdjunto], [p_idSolicitud], [p_etapaWF], [p_nombreArchivo], [p_fechaSubida], [p_usuarioSubida], [p_descripcion], [p_TipoAccion]);"
          
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_idAdjunto", adj.idAdjunto
    params.Add "p_idSolicitud", adj.idSolicitud
    params.Add "p_etapaWF", adj.etapaWF
    params.Add "p_nombreArchivo", adj.nombreArchivo
    params.Add "p_fechaSubida", adj.fechaSubida
    params.Add "p_usuarioSubida", adj.usuarioSubida
    params.Add "p_descripcion", adj.descripcion
    params.Add "p_TipoAccion", adj.tipoAccion
    
    Call RepositorioComun.EjecutarAccion(sql, params, dbConexion)
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "AdjuntoRepositorio.Guardar"
    errObj.Raise
End Sub

' --- getSiguienteIDAdjunto MODIFICADO ---
Private Function getSiguienteIDAdjunto(Optional db As DAO.Database) As Long
    Dim dbConexion As DAO.Database, rcd As DAO.Recordset, maxID As Long
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    maxID = 0
    Set rcd = dbConexion.OpenRecordset("SELECT MAX(idAdjunto) AS MaxValue FROM tbAdjuntos", dbOpenSnapshot)
    
    If Not rcd.EOF Then maxID = Nz(rcd!MaxValue, 0)
    
    getSiguienteIDAdjunto = maxID + 1

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "AdjuntoRepositorio.getSiguienteIDAdjunto"
    errObj.Raise
End Function

Public Function ExisteAdjuntoEtapa(ByVal idSolicitud As Long, ByVal etapaWF As String, Optional db As DAO.Database) As Boolean
    Dim sql As String, params As Object, rcd As DAO.Recordset, dbConexion As DAO.Database
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    ExisteAdjuntoEtapa = False
    
    sql = "SELECT COUNT(*) AS n FROM tbAdjuntos WHERE idSolicitud=[p_id] AND etapaWF=[p_etapa];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_id", idSolicitud
    params.Add "p_etapa", etapaWF
    
    Set rcd = RepositorioComun.EjecutarConsulta(sql, params, dbConexion)
    
    If Not rcd Is Nothing Then
        If rcd.Fields(0).value > 0 Then ExisteAdjuntoEtapa = True
    End If

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
Errores:
    If Not rcd Is Nothing Then rcd.Close
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "AdjuntoRepositorio.ExisteAdjuntoEtapa"
    errObj.Raise
End Function

Public Function getPorEtapa(ByVal idSolicitud As Long, ByVal etapaWF As String, Optional db As DAO.Database) As Adjunto
    Dim sql As String, params As Object, rcd As DAO.Recordset, dbConexion As DAO.Database
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    Set getPorEtapa = Nothing
    
    sql = "SELECT TOP 1 * FROM tbAdjuntos WHERE idSolicitud=[p_id] AND etapaWF=[p_etapa] ORDER BY fechaSubida DESC;"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_id", idSolicitud
    params.Add "p_etapa", etapaWF
    
    Set getPorEtapa = RepositorioComun.HidratarEntidadDesdeSQL(sql, "Adjunto", dbConexion, params)
    
    Exit Function
Errores:
    Set getPorEtapa = Nothing
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "AdjuntoRepositorio.getPorEtapa"
    errObj.Raise
End Function

Public Function getPorIdSolicitud(ByVal idSol As Long, Optional db As DAO.Database) As Collection
    Dim sql As String
    Dim params As Object
    Dim col As Collection
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    If db Is Nothing Then
        Set dbConexion = getdb()
    Else
        Set dbConexion = db
    End If
    
    sql = "SELECT * FROM tbAdjuntos WHERE idSolicitud = [p_IDSol] ORDER BY fechaSubida DESC;"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_IDSol", idSol
    
    Dim dict As Scripting.Dictionary
    Dim key As Variant
    Set dict = RepositorioComun.HidratarColeccionDesdeSQL(sql, "Adjunto", "idAdjunto", dbConexion, params)
    
    Set col = New Collection
    For Each key In dict.Keys
        col.Add dict(key)
    Next key
        
    Set getPorIdSolicitud = col
    
LimpiarYSalir:
    On Error Resume Next
    Set dict = Nothing
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "AdjuntoRepositorio.getPorIdSolicitud"
    errObj.AddToCallStack "idSolicitud: " & idSol
    errObj.Raise
End Function


Public Function getPorID(ByVal idAdjunto As Long, Optional db As DAO.Database) As Adjunto
    Dim sql As String, params As Object, dbConexion As DAO.Database
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbAdjuntos WHERE idAdjunto = [p_ID];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", idAdjunto
    
    Set getPorID = RepositorioComun.HidratarEntidadDesdeSQL(sql, "Adjunto", dbConexion, params)
    
LimpiarYSalir:
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "AdjuntoRepositorio.getPorID"
    errObj.AddToCallStack "idAdjunto: " & idAdjunto
    errObj.Raise
End Function

Public Sub Eliminar(ByVal idAdjunto As Long, Optional db As DAO.Database)
    Dim sql As String, params As Object, dbConexion As DAO.Database
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "DELETE FROM tbAdjuntos WHERE idAdjunto = [p_ID];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", idAdjunto
    
    Call RepositorioComun.EjecutarAccion(sql, params, dbConexion)
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "AdjuntoRepositorio.Eliminar"
    errObj.AddToCallStack "idAdjunto: " & idAdjunto
    errObj.Raise
End Sub

Public Sub EliminarPorIdSolicitud(ByVal idSol As Long, ByRef db As DAO.Database)
    ' RESPONSABILIDAD: Eliminar todos los registros de tbAdjuntos asociados a una solicitud.
    Dim sql As String
    Dim params As Object
    On Error GoTo Errores
    
    sql = "DELETE FROM tbAdjuntos WHERE idSolicitud = [p_IDSol];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_IDSol", idSol
    
    Call RepositorioComun.EjecutarAccion(sql, params, db)
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "AdjuntoRepositorio.EliminarPorIdSolicitud"
    errObj.AddToCallStack "idSolicitud: " & idSol
    errObj.Raise
End Sub




