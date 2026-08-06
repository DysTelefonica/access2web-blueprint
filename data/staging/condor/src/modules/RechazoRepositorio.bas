Attribute VB_Name = "RechazoRepositorio"

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: RechazoRepositorio
' DESCRIPCIÓN: Gestión de persistencia para la tabla tbRechazos.
' ==========================================================================

Public Function GetUltimoRechazoActivo(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As rechazo
    ' RESPONSABILIDAD: Recuperar el rechazo vigente para una solicitud.
    Dim sql As String
    Dim rst As DAO.Recordset
    Dim resultado As rechazo
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    ' Consulta directa a tbRechazos buscando el activo
    sql = "SELECT TOP 1 * FROM tbRechazos " & _
          "WHERE idSolicitud = " & idSolicitud & " AND esActivo = True " & _
          "ORDER BY fechaRechazo DESC;"
          
    Set rst = dbConexion.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rst.EOF Then
        Set resultado = New rechazo
        resultado.idSolicitud = rst!idSolicitud
        resultado.fechaRechazo = Nz(rst!fechaRechazo, Date)
        resultado.motivoPrincipal = Nz(rst!motivoPrincipal, "")
        resultado.areaAfectada = Nz(rst!areaAfectada, "")
        resultado.comentarios = Nz(rst!comentarios, "")
        resultado.usuarioRechazo = Nz(rst!usuarioRechazo, "")
        resultado.CambiosTecnico = Nz(rst!CambiosTecnico, "")
        ' NOTA: notasSubsanacion eliminado por deprecacion (delta automatico)
        
        Set GetUltimoRechazoActivo = resultado
    Else
        Set GetUltimoRechazoActivo = Nothing
    End If
    
    rst.Close
    Set rst = Nothing
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RechazoRepositorio.GetUltimoRechazoActivo"
    errObj.AddToCallStack "ID Solicitud: " & idSolicitud
    errObj.Raise
End Function

Public Function GetUltimoRechazo(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As rechazo
    ' RESPONSABILIDAD: Recuperar el último rechazo (activo o inactivo) para ver cambios del técnico.
    ' Usado por JsonHelper para obtener el JSON de cambios cuando el rechazo ya está inactivo.
    Dim sql As String
    Dim rst As DAO.Recordset
    Dim resultado As rechazo
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    ' Consulta directa a tbRechazos sin filtro de esActivo
    sql = "SELECT TOP 1 * FROM tbRechazos " & _
          "WHERE idSolicitud = " & idSolicitud & " " & _
          "ORDER BY fechaRechazo DESC;"
          
    Set rst = dbConexion.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rst.EOF Then
        Set resultado = New rechazo
        resultado.idSolicitud = rst!idSolicitud
        resultado.fechaRechazo = Nz(rst!fechaRechazo, Date)
        resultado.motivoPrincipal = Nz(rst!motivoPrincipal, "")
        resultado.areaAfectada = Nz(rst!areaAfectada, "")
        resultado.comentarios = Nz(rst!comentarios, "")
        resultado.usuarioRechazo = Nz(rst!usuarioRechazo, "")
        resultado.CambiosTecnico = Nz(rst!CambiosTecnico, "")
        
        Set GetUltimoRechazo = resultado
    Else
        Set GetUltimoRechazo = Nothing
    End If
    
    rst.Close
    Set rst = Nothing
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RechazoRepositorio.GetUltimoRechazo"
    errObj.AddToCallStack "ID Solicitud: " & idSolicitud
    errObj.Raise
End Function

Public Function GetFechaUltimoRechazo(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As Date
    ' RESPONSABILIDAD: Obtener solo la fecha del rechazo activo para calcular intervalos.
    ' Retorna Date 0 si no hay rechazo activo.
    Dim sql As String
    Dim rst As DAO.Recordset
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT fechaRechazo FROM tbRechazos " & _
          "WHERE idSolicitud = " & idSolicitud & " AND esActivo = True " & _
          "ORDER BY fechaRechazo DESC;"
          
    Set rst = dbConexion.OpenRecordset(sql, dbOpenSnapshot)
    
    If Not rst.EOF Then
        GetFechaUltimoRechazo = Nz(rst!fechaRechazo, 0)
    Else
        GetFechaUltimoRechazo = 0
    End If
    
    rst.Close
    Set rst = Nothing
    Exit Function

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RechazoRepositorio.GetFechaUltimoRechazo"
    errObj.Raise
End Function

Public Sub GuardarRechazo(ByVal objRechazo As rechazo, Optional ByRef dbTransaccional As DAO.Database)
    ' RESPONSABILIDAD: Insertar un nuevo rechazo y marcarlo como activo.
    ' HOTFIX: Garantizar unicidad de rechazo activo desactivando previos.
    
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    
    On Error GoTo Errores
    
    If dbTransaccional Is Nothing Then Set db = getdb() Else Set db = dbTransaccional
    
    ' Garantizar que solo hay uno activo
    Call DesactivarRechazosPrevios(objRechazo.idSolicitud, db)
    
    Set rst = db.OpenRecordset("tbRechazos", dbOpenDynaset)
    rst.AddNew
    rst!idSolicitud = objRechazo.idSolicitud
    rst!fechaRechazo = objRechazo.fechaRechazo
    rst!motivoPrincipal = Left(objRechazo.motivoPrincipal, 255) ' Truncar por seguridad
    rst!areaAfectada = Left(objRechazo.areaAfectada, 50)
    rst!comentarios = objRechazo.comentarios
    rst!usuarioRechazo = Left(objRechazo.usuarioRechazo, 100)
    rst!CambiosTecnico = objRechazo.CambiosTecnico
    rst!esActivo = True
    rst.Update
    
    rst.Close
    Set rst = Nothing
    Exit Sub

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RechazoRepositorio.GuardarRechazo"
    errObj.Raise
End Sub

Public Sub DesactivarRechazosPrevios(ByVal idSolicitud As Long, Optional ByRef dbTransaccional As DAO.Database)
    ' RESPONSABILIDAD: Marcar como inactivos todos los rechazos anteriores de una solicitud.
    Dim db As DAO.Database
    Dim sql As String
    
    On Error GoTo Errores
    
    If dbTransaccional Is Nothing Then Set db = getdb() Else Set db = dbTransaccional
    
    sql = "UPDATE tbRechazos SET esActivo = False WHERE idSolicitud = " & idSolicitud
    db.Execute sql, dbFailOnError
    
    Exit Sub

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RechazoRepositorio.DesactivarRechazosPrevios"
    errObj.Raise
End Sub

Public Sub EliminarRechazosPorSolicitud(ByVal idSolicitud As Long, Optional ByRef dbTransaccional As DAO.Database)
    Dim db As DAO.Database
    Dim sql As String
    
    On Error GoTo Errores
    
    If dbTransaccional Is Nothing Then Set db = getdb() Else Set db = dbTransaccional
    
    sql = "DELETE FROM tbRechazos WHERE idSolicitud = " & idSolicitud
    db.Execute sql, dbFailOnError
    
    Exit Sub

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RechazoRepositorio.EliminarRechazosPorSolicitud"
    errObj.Raise
End Sub

Public Sub ActualizarComentarioRechazoActivo(ByVal idSolicitud As Long, ByVal comentarios As String, Optional ByRef dbTransaccional As DAO.Database)
    ' RESPONSABILIDAD: Actualizar el campo comentarios (Memo) del rechazo activo.
    ' Usa DAO Recordset para evitar truncamiento a 255 del SQL concatenation.
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim gestionaTransaccionLocal As Boolean
    
    On Error GoTo Errores
    
    If dbTransaccional Is Nothing Then
        gestionaTransaccionLocal = True
        Set ws = getWorkspace()
        Set db = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
    Else
        gestionaTransaccionLocal = False
        Set db = dbTransaccional
    End If
    
    Set rst = db.OpenRecordset("SELECT * FROM tbRechazos WHERE idSolicitud = " & idSolicitud & " AND esActivo = True", dbOpenDynaset)
    If Not rst.EOF Then
        rst.Edit
        rst!comentarios = comentarios  ' DAO maneja Memo correctamente - no trunca
        rst.Update
    End If
    rst.Close
    
    If gestionaTransaccionLocal Then
        ws.CommitTrans
    End If
    
LimpiarYSalir:
    On Error Resume Next
    If gestionaTransaccionLocal And Not db Is Nothing Then db.Close
    Set ws = Nothing
    Exit Sub
    
Errores:
    If gestionaTransaccionLocal And Not ws Is Nothing Then ws.Rollback
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RechazoRepositorio.ActualizarComentarioRechazoActivo"
    errObj.Raise
End Sub

Public Sub ActualizarJSON(ByVal idSolicitud As Long, ByVal jsonString As String, Optional ByRef dbTransaccional As DAO.Database)
    ' RESPONSABILIDAD: Actualizar el campo CambiosTecnico del rechazo activo (Memo).
    ' Usa DAO Recordset para evitar truncamiento a 255 del SQL concatenation.
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim gestionaTransaccionLocal As Boolean
    Dim ws As DAO.Workspace
    
    On Error GoTo Errores
    
    If dbTransaccional Is Nothing Then
        gestionaTransaccionLocal = True
        Set ws = getWorkspace()
        Set db = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
    Else
        gestionaTransaccionLocal = False
        Set db = dbTransaccional
    End If
    
    Set rst = db.OpenRecordset("SELECT * FROM tbRechazos WHERE idSolicitud = " & idSolicitud & " AND esActivo = True", dbOpenDynaset)
    If Not rst.EOF Then
        rst.Edit
        rst!CambiosTecnico = jsonString  ' DAO maneja Memo correctamente
        rst.Update
    End If
    rst.Close
    
    If gestionaTransaccionLocal Then
        ws.CommitTrans
    End If
    
LimpiarYSalir:
    On Error Resume Next
    If gestionaTransaccionLocal And Not db Is Nothing Then db.Close
    Set ws = Nothing
    Exit Sub

Errores:
    If gestionaTransaccionLocal And Not ws Is Nothing Then ws.Rollback
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RechazoRepositorio.ActualizarJSON"
    errObj.Raise
End Sub


Public Function GetHistorialPorSolicitud(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As DAO.Recordset
    Dim sql As String
    Dim dbConexion As DAO.Database
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then
        Set GetHistorialPorSolicitud = Nothing
        Exit Function
    End If
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT idSolicitud, fechaRechazo, motivoPrincipal, comentarios, usuarioRechazo, CambiosTecnico " & _
          "FROM tbRechazos WHERE idSolicitud = " & idSolicitud & " ORDER BY fechaRechazo DESC;"
    
    Set GetHistorialPorSolicitud = dbConexion.OpenRecordset(sql, dbOpenSnapshot)
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RechazoRepositorio.GetHistorialPorSolicitud"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud
    errObj.Raise
End Function





