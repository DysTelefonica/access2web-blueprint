Attribute VB_Name = "ValidacionRevisionRepositorio"

Option Compare Database
Option Explicit

Public Function GetUltimoOrdinal(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As Long
    Dim sql As String
    Dim params As Object
    Dim rcd As DAO.Recordset
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then
        GetUltimoOrdinal = 0
        Exit Function
    End If
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT ordinal FROM tbValidacionRevision WHERE idSolicitud = [p_ID] ORDER BY ordinal DESC;"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", idSolicitud
    
    Set rcd = RepositorioComun.EjecutarConsulta(sql, params, dbConexion)
    
    If Not rcd Is Nothing Then
        If Not rcd.EOF Then
            GetUltimoOrdinal = Nz(rcd!ordinal, 0)
        Else
            GetUltimoOrdinal = 0
        End If
    Else
        GetUltimoOrdinal = 0
    End If
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionRepositorio.GetUltimoOrdinal"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud
    errObj.Raise
End Function

Public Function GetUltimoCicloCerrado(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As ValidacionRevision
    Dim sql As String
    Dim params As Object
    Dim rcd As DAO.Recordset
    Dim dbConexion As DAO.Database
    Dim ciclo As ValidacionRevision
    
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then
        Set GetUltimoCicloCerrado = Nothing
        Exit Function
    End If
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT TOP 1 * FROM tbValidacionRevision WHERE idSolicitud = [p_ID] ORDER BY ordinal DESC;"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", idSolicitud
    
    Set rcd = RepositorioComun.EjecutarConsulta(sql, params, dbConexion)
    
    If Not rcd Is Nothing Then
        If Not rcd.EOF Then
            Set ciclo = New ValidacionRevision
            ciclo.id = rcd!id
            ciclo.idSolicitud = rcd!idSolicitud
            ciclo.ordinal = rcd!ordinal
            ciclo.resultado = Nz(rcd!resultado, "")
            ciclo.comentarios = Nz(rcd!comentarios, "")
            ciclo.usuario = Nz(rcd!usuario, "")
            ciclo.hashDatos = Nz(rcd!hashDatos, "")
            ciclo.idAdjunto = Nz(rcd!idAdjunto, 0)
            If Not IsNull(rcd!fechaEnvio) Then ciclo.fechaEnvio = rcd!fechaEnvio
            Set GetUltimoCicloCerrado = ciclo
        Else
            Set GetUltimoCicloCerrado = Nothing
        End If
    Else
        Set GetUltimoCicloCerrado = Nothing
    End If
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionRepositorio.GetUltimoCicloCerrado"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud
    errObj.Raise
End Function

Public Function GetHistorialPorSolicitud(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As DAO.Recordset
    Dim sql As String
    Dim params As Object
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then
        Set GetHistorialPorSolicitud = Nothing
        Exit Function
    End If
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT Id, idSolicitud, ordinal, FechaEnvio, FechaRecepcion, Resultado, Comentarios, Usuario, idAdjunto, HashDatos " & _
          "FROM tbValidacionRevision WHERE idSolicitud = [p_ID] ORDER BY ordinal DESC;"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", idSolicitud
    
    Set GetHistorialPorSolicitud = RepositorioComun.EjecutarConsulta(sql, params, dbConexion)
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionRepositorio.GetHistorialPorSolicitud"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud
    errObj.Raise
End Function

Public Sub EliminarPorSolicitud(ByVal idSolicitud As Long, ByRef db As DAO.Database)
    Dim sql As String
    Dim params As Object
    
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then Exit Sub
    
    sql = "DELETE FROM tbValidacionRevision WHERE idSolicitud = [p_ID];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", idSolicitud
    
    Call RepositorioComun.EjecutarAccion(sql, params, db)
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionRepositorio.EliminarPorSolicitud"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud
    errObj.Raise
End Sub

Public Sub ActualizarHashRevision(ByVal idSolicitud As Long, ByVal ordinal As Long, ByVal hashDatos As String, Optional ByRef db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim sql As String
    Dim params As Object
    
    On Error GoTo Errores
    
    If idSolicitud <= 0 Or ordinal <= 0 Then Exit Sub
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "UPDATE tbValidacionRevision SET HashDatos = [p_hashDatos] WHERE idSolicitud = [p_idSolicitud] AND ordinal = [p_ordinal];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_hashDatos", hashDatos
    params.Add "p_idSolicitud", idSolicitud
    params.Add "p_ordinal", ordinal
    Call RepositorioComun.EjecutarAccion(sql, params, dbConexion)
    
    Exit Sub
Errores:
    Set dbConexion = Nothing
    Set params = Nothing
End Sub

Public Function GetUltimoPendienteId(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As Long
    Dim sql As String
    Dim params As Object
    Dim rcd As DAO.Recordset
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then
        GetUltimoPendienteId = 0
        Exit Function
    End If
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT Id FROM tbValidacionRevision WHERE idSolicitud = [p_ID] AND (FechaRecepcion Is Null OR FechaRecepcion = 0) ORDER BY ordinal DESC;"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", idSolicitud
    
    Set rcd = RepositorioComun.EjecutarConsulta(sql, params, dbConexion)
    
    If Not rcd Is Nothing Then
        If Not rcd.EOF Then
            GetUltimoPendienteId = Nz(rcd!id, 0)
        Else
            GetUltimoPendienteId = 0
        End If
    Else
        GetUltimoPendienteId = 0
    End If
    
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionRepositorio.GetUltimoPendienteId"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud
    errObj.Raise
End Function

Public Function GetCicloAbierto(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As ValidacionRevision
    Dim sql As String
    Dim params As Object
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then
        Set GetCicloAbierto = Nothing
        Exit Function
    End If
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT TOP 1 Id, idSolicitud, ordinal, FechaEnvio, FechaRecepcion, Resultado, Comentarios, Usuario, idAdjunto, HashDatos " & _
          "FROM tbValidacionRevision WHERE idSolicitud = [p_ID] AND (FechaRecepcion Is Null OR FechaRecepcion = 0 OR Comentarios LIKE '*Importación Word*') ORDER BY ordinal DESC;"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", idSolicitud
    
    Set GetCicloAbierto = RepositorioComun.HidratarEntidadDesdeSQL(sql, "ValidacionRevision", dbConexion, params)
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionRepositorio.GetCicloAbierto"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud
    errObj.Raise
End Function

Public Function GetAdjuntoIdCicloActual(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As Long
    Dim dbConexion As DAO.Database
    Dim cicloActual As ValidacionRevision
    
    On Error GoTo Errores
    
    GetAdjuntoIdCicloActual = 0
    
    If idSolicitud <= 0 Then Exit Function
    
    If db Is Nothing Then
        Set dbConexion = getdb()
    Else
        Set dbConexion = db
    End If
    
    Set cicloActual = GetCicloAbierto(idSolicitud, dbConexion)
    If Not cicloActual Is Nothing Then
        GetAdjuntoIdCicloActual = Nz(cicloActual.idAdjunto, 0)
    End If
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionRepositorio.GetAdjuntoIdCicloActual"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud
    errObj.Raise
End Function

Public Sub GuardarHitoEnvio(ByVal idSolicitud As Long, ByVal fechaEnvio As Variant, ByVal usuario As Variant, Optional ByVal comentarios As String = "", Optional ByRef db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim sql As String
    Dim params As Object
    Dim nuevoOrdinal As Long
    Dim fechaEnvioValue As Variant
    Dim usuarioValue As String
    Dim cicloActual As ValidacionRevision
    
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then Exit Sub
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    If IsDate(fechaEnvio) Then
        fechaEnvioValue = CDate(fechaEnvio)
    Else
        fechaEnvioValue = Now
    End If
    
    usuarioValue = CStr(Nz(usuario, ""))

    Set cicloActual = GetCicloAbierto(idSolicitud, dbConexion)
    If Not cicloActual Is Nothing Then
        sql = "UPDATE tbValidacionRevision SET FechaEnvio = [p_fechaEnvio], Usuario = [p_usuario] " & _
              "WHERE Id = [p_id] AND FechaEnvio Is Null;"
        Set params = CreateObject("Scripting.Dictionary")
        params.Add "p_fechaEnvio", fechaEnvioValue
        params.Add "p_usuario", usuarioValue
        params.Add "p_id", cicloActual.id
        Call RepositorioComun.EjecutarAccion(sql, params, dbConexion)
    Else
        nuevoOrdinal = GetUltimoOrdinal(idSolicitud, dbConexion) + 1
        sql = "INSERT INTO tbValidacionRevision (idSolicitud, ordinal, FechaEnvio, Resultado, Comentarios, Usuario, idAdjunto) " & _
              "VALUES ([p_idSolicitud], [p_ordinal], [p_fechaEnvio], [p_resultado], [p_comentarios], [p_usuario], [p_idAdjunto]);"
        Set params = CreateObject("Scripting.Dictionary")
        params.Add "p_idSolicitud", idSolicitud
        params.Add "p_ordinal", nuevoOrdinal
        params.Add "p_fechaEnvio", fechaEnvioValue
        params.Add "p_resultado", "PENDIENTE"
        params.Add "p_comentarios", ""
        params.Add "p_usuario", usuarioValue
        params.Add "p_idAdjunto", 0
        Call RepositorioComun.EjecutarAccion(sql, params, dbConexion)
    End If
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionRepositorio.GuardarHitoEnvio"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud
    errObj.Raise
End Sub

Public Sub GuardarHitoBorrador(ByVal idSolicitud As Long, ByVal idAdjunto As Long, ByVal usuario As Variant, Optional ByVal hashDatos As String = "", Optional ByRef db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim sql As String
    Dim params As Object
    Dim nuevoOrdinal As Long
    Dim usuarioValue As String
    Dim cicloActual As ValidacionRevision
    
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then Exit Sub
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    usuarioValue = CStr(Nz(usuario, ""))
    Set cicloActual = GetCicloAbierto(idSolicitud, dbConexion)
    
    If Not cicloActual Is Nothing Then
        sql = "UPDATE tbValidacionRevision SET idAdjunto = [p_idAdjunto], Usuario = [p_usuario], HashDatos = [p_hashDatos] " & _
              "WHERE Id = [p_id];"
        Set params = CreateObject("Scripting.Dictionary")
        params.Add "p_idAdjunto", idAdjunto
        params.Add "p_usuario", usuarioValue
        params.Add "p_hashDatos", hashDatos
        params.Add "p_id", cicloActual.id
        Call RepositorioComun.EjecutarAccion(sql, params, dbConexion)
    Else
        nuevoOrdinal = GetUltimoOrdinal(idSolicitud, dbConexion) + 1
        sql = "INSERT INTO tbValidacionRevision (idSolicitud, ordinal, Resultado, Comentarios, Usuario, idAdjunto, HashDatos) " & _
              "VALUES ([p_idSolicitud], [p_ordinal], [p_resultado], [p_comentarios], [p_usuario], [p_idAdjunto], [p_hashDatos]);"
        Set params = CreateObject("Scripting.Dictionary")
        params.Add "p_idSolicitud", idSolicitud
        params.Add "p_ordinal", nuevoOrdinal
        params.Add "p_resultado", "PENDIENTE"
        params.Add "p_comentarios", ""
        params.Add "p_usuario", usuarioValue
        params.Add "p_idAdjunto", idAdjunto
        params.Add "p_hashDatos", hashDatos
        Call RepositorioComun.EjecutarAccion(sql, params, dbConexion)
    End If
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionRepositorio.GuardarHitoBorrador"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud
    errObj.Raise
End Sub

Public Function GetHashVersionAnterior(ByVal idSolicitud As Long, ByVal ordinalActual As Long, Optional ByRef db As DAO.Database) As String
    Dim sql As String
    Dim params As Object
    Dim rcd As DAO.Recordset
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    GetHashVersionAnterior = ""
    
    If idSolicitud <= 0 Or ordinalActual <= 1 Then Exit Function
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    ' Buscamos el hash del ordinal anterior (ordinalActual - 1)
    sql = "SELECT HashDatos FROM tbValidacionRevision WHERE idSolicitud = [p_ID] AND ordinal = [p_ordinal];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", idSolicitud
    params.Add "p_ordinal", ordinalActual - 1
    
    Set rcd = RepositorioComun.EjecutarConsulta(sql, params, dbConexion)
    
    If Not rcd Is Nothing Then
        If Not rcd.EOF Then
            GetHashVersionAnterior = Nz(rcd!hashDatos, "")
        End If
        rcd.Close
    End If
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionRepositorio.GetHashVersionAnterior"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud & ", ordinal=" & ordinalActual
    errObj.Raise
End Function

Public Sub CerrarHitoRespuesta(ByVal idValidacion As Long, ByVal resultado As Variant, ByVal comentarios As Variant, ByVal fechaRecepcion As Variant, Optional ByRef db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim sql As String
    Dim params As Object
    Dim resultadoValue As String
    Dim fechaRecepcionValue As Variant
    
    On Error GoTo Errores
    
    If idValidacion <= 0 Then Exit Sub
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    resultadoValue = UCase$(CStr(Nz(resultado, "")))
    If resultadoValue = "" Then resultadoValue = "COMENTARIOS"
    
    If IsDate(fechaRecepcion) Then
        fechaRecepcionValue = CDate(fechaRecepcion)
    Else
        fechaRecepcionValue = Now
    End If
    
    sql = "UPDATE tbValidacionRevision SET FechaRecepcion = [p_fechaRecepcion], Resultado = [p_resultado], Comentarios = [p_comentarios] " & _
          "WHERE Id = [p_id] AND (FechaRecepcion Is Null OR FechaRecepcion = 0);"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_fechaRecepcion", fechaRecepcionValue
    params.Add "p_resultado", resultadoValue
    params.Add "p_comentarios", CStr(Nz(comentarios, ""))
    params.Add "p_id", idValidacion
    
    Call RepositorioComun.EjecutarAccion(sql, params, dbConexion)
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionRepositorio.CerrarHitoRespuesta"
    errObj.AddToCallStack "idValidacion=" & idValidacion
    errObj.Raise
End Sub

Public Sub ActualizarRespuestaPorId(ByVal idValidacion As Long, ByVal resultado As String, ByVal comentarios As String, Optional ByRef db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim sql As String
    Dim params As Object
    
    On Error GoTo Errores
    
    Debug.Print "REPO_ActualizarRespuestaPorId: Inicio idValidacion=" & idValidacion
    
    If idValidacion <= 0 Then Exit Sub
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    Debug.Print "REPO_ActualizarRespuestaPorId: dbConexion esta asignada"
    
    sql = "UPDATE tbValidacionRevision SET Resultado = [p_resultado], " & _
          "Comentarios = [p_comentarios], FechaRecepcion = Now() " & _
          "WHERE Id = [p_id]"
    
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_resultado", UCase$(resultado)
    params.Add "p_comentarios", comentarios
    params.Add "p_id", idValidacion
    
    Debug.Print "REPO_ActualizarRespuestaPorId: Llamando EjecutarAccion"
    Call RepositorioComun.EjecutarAccion(sql, params, dbConexion)
    Debug.Print "REPO_ActualizarRespuestaPorId: EjecutarAccion completado"
    
    Exit Sub
Errores:
    Debug.Print "REPO_ActualizarRespuestaPorId ERROR: " & Err.Number & " - " & Err.description
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionRepositorio.ActualizarRespuestaPorId"
    errObj.AddToCallStack "idValidacion=" & idValidacion
    errObj.Raise
End Sub

Public Function GetVersionActualBorrador(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As String
    Dim dbConexion As DAO.Database
    Dim sql As String
    Dim params As Object
    Dim rcd As DAO.Recordset
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    ' 1. Buscamos si hay un ciclo abierto (sin FechaRecepcion)
    sql = "SELECT ordinal FROM tbValidacionRevision WHERE idSolicitud = [p_ID] AND FechaRecepcion Is Null ORDER BY ordinal DESC;"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", idSolicitud
    Set rcd = RepositorioComun.EjecutarConsulta(sql, params, dbConexion)
    
    If Not rcd Is Nothing Then
        If Not rcd.EOF Then
            GetVersionActualBorrador = "v" & CStr(rcd!ordinal)
            GoTo LimpiarYSalir
        End If
    End If
    
    ' 2. Si no hay ciclo abierto, devolvemos el siguiente al máximo
    GetVersionActualBorrador = "v" & CStr(GetUltimoOrdinal(idSolicitud, dbConexion) + 1)
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
Errores:
    GetVersionActualBorrador = "v1"
    Resume LimpiarYSalir
End Function

Public Sub GetUltimoEstadoValidacion(ByVal idSolicitud As Long, ByRef resultado As String, ByRef comentarios As String, Optional ByRef db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim sql As String
    Dim params As Object
    Dim rcd As DAO.Recordset
    
    On Error GoTo Errores
    
    resultado = ""
    comentarios = ""
    
    If idSolicitud <= 0 Then Exit Sub
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    ' Buscamos el último ciclo (abierto o cerrado)
    sql = "SELECT TOP 1 Resultado, Comentarios, FechaRecepcion FROM tbValidacionRevision " & _
          "WHERE idSolicitud = [p_ID] " & _
          "ORDER BY ordinal DESC, Id DESC;"
          
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", idSolicitud
    
    Set rcd = RepositorioComun.EjecutarConsulta(sql, params, dbConexion)
    
    If Not rcd Is Nothing Then
        If Not rcd.EOF Then
            If IsNull(rcd!fechaRecepcion) Then
                resultado = "PENDIENTE"
                comentarios = ""
            Else
                resultado = CStr(Nz(rcd!resultado, ""))
                comentarios = CStr(Nz(rcd!comentarios, ""))
            End If
        End If
        rcd.Close
    End If
    
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionRepositorio.GetUltimoEstadoValidacion"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud
    errObj.Raise
End Sub

Public Sub CrearNuevoCicloValidacion(ByVal idSolicitud As Long, ByVal usuario As String, ByVal hashActual As String, Optional ByRef db As DAO.Database, Optional ByVal comentarios As String = "")
    Dim dbConexion As DAO.Database
    Dim sql As String
    Dim params As Object
    Dim nuevoOrdinal As Long
    
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then Exit Sub
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    nuevoOrdinal = GetUltimoOrdinal(idSolicitud, dbConexion) + 1
    
    sql = "INSERT INTO tbValidacionRevision (idSolicitud, ordinal, FechaEnvio, Resultado, Comentarios, Usuario, idAdjunto, HashDatos) " & _
          "VALUES ([p_idSolicitud], [p_ordinal], [p_fechaEnvio], [p_resultado], [p_comentarios], [p_usuario], [p_idAdjunto], [p_hashDatos]);"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_idSolicitud", idSolicitud
    params.Add "p_ordinal", nuevoOrdinal
    params.Add "p_fechaEnvio", Now()
    params.Add "p_resultado", "PENDIENTE"
    params.Add "p_comentarios", comentarios
    params.Add "p_usuario", usuario
    params.Add "p_idAdjunto", 0
    params.Add "p_hashDatos", hashActual
    
    Call RepositorioComun.EjecutarAccion(sql, params, dbConexion)
    
    Debug.Print "CREAR_CICLO: Nuevo ciclo ordinal=" & nuevoOrdinal & " para idSolicitud=" & idSolicitud & ", HashDatos=" & Left(hashActual, 16) & "..."
    
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "ValidacionRevisionRepositorio.CrearNuevoCicloValidacion"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud
    errObj.Raise
End Sub





