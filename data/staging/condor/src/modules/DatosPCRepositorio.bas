Attribute VB_Name = "DatosPCRepositorio"

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: DatosPCRepositorio.bas
' ACCIÓN: AÑADIR nuevo método "Upsert" genérico.
' ==========================================================================
Public Sub Guardar(ByRef objPC As DatosPC, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    ' Buscamos por idSolicitud, que es nuestra clave de negocio
    sql = "SELECT * FROM tbDatosPC WHERE idSolicitud = " & objPC.idSolicitud & " ORDER BY idDatosPC"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If rcd.EOF Then
        ' MODO INSERT: El registro no existe
        rcd.AddNew
        ' Asignamos el nuevo ID autoincremental si la tabla lo tuviera,
        ' o calculamos el siguiente ID si es numérico.
        If objPC.idDatosPC = 0 Then
            objPC.idDatosPC = getSiguienteIDDatosPC(dbConexion)
        End If
        ' Rellenamos todos los campos desde el objeto
        Call RepositorioComun.RellenarRecordsetDesdeObjeto(rcd, objPC)
    Else
        ' MODO UPDATE: El registro ya existe
        rcd.Edit
        ' Rellenamos todos los campos, omitiendo la clave primaria
        Call RepositorioComun.RellenarRecordsetDesdeObjeto(rcd, objPC, True)
    End If
    
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCRepositorio.Guardar"
    errObj.AddToCallStack "idSolicitud: " & objPC.idSolicitud
    errObj.Raise
End Sub

Public Function getPorIdSolicitud(ByVal idSol As Long, Optional db As DAO.Database) As DatosPC
    Dim sql As String, params As Object, dbConexion As DAO.Database
    On Error GoTo Errores
    
    ' Si nos pasan transacción la usamos, si no, nueva conexión
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosPC WHERE idSolicitud = [p_IDSol] ORDER BY idDatosPC;"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_IDSol", idSol
    
    Set getPorIdSolicitud = RepositorioComun.HidratarEntidadDesdeSQL(sql, "DatosPC", dbConexion, params)
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCRepositorio.getPorIdSolicitud"
    errObj.Raise
End Function



Private Function getSiguienteIDDatosPC(ByRef db As DAO.Database) As Long
    Dim rcd As DAO.Recordset
    Dim maxID As Long
    On Error GoTo Errores
    
    Set rcd = db.OpenRecordset("SELECT Max(idDatosPC) AS MaxID FROM tbDatosPC")
    
    If Not rcd.EOF Then
        maxID = Nz(rcd!maxID, 0)
    End If
    
    getSiguienteIDDatosPC = maxID + 1
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
    
Errores:
    
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCRepositorio.getSiguienteIDDatosPC"
    errObj.Raise
End Function
' === FUNCIONES A AÑADIR EN DatosPCRepositorio.bas ===

' ==========================================================================
' MÓDULO: DatosPCRepositorio.bas
' ACCIÓN: REEMPLAZAR ActualizarDatosGenerales para usar DAO.Recordset y soportar campos Memo.
' ==========================================================================
Public Sub ActualizarDatosGenerales(ByRef objPC As DatosPC)
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores

    Set db = getdb()

    ' 1. Abrimos un Recordset de la tabla, filtrando por el idSolicitud para editar o comprobar si existe.
    sql = "SELECT * FROM tbDatosPC WHERE idSolicitud = " & objPC.idSolicitud & " ORDER BY idDatosPC"
    Set rcd = db.OpenRecordset(sql, dbOpenDynaset)

    If rcd.EOF Then
        ' --- MODO INSERT: El registro no existe, así que lo creamos ---
        rcd.AddNew
        
        ' Asignamos el nuevo ID al objeto y al recordset
        objPC.idDatosPC = getSiguienteIDDatosPC(db)
        rcd!idDatosPC = objPC.idDatosPC
        rcd!idSolicitud = objPC.idSolicitud
        
        ' Usamos nuestro ayudante RellenarRecordsetDesdeObjeto para los demás campos
        ' Omitimos el PK porque ya lo hemos asignado.
        Call RepositorioComun.RellenarRecordsetDesdeObjeto(rcd, objPC, True)

    Else
        ' --- MODO UPDATE: El registro ya existe, lo editamos ---
        rcd.Edit
        
        ' Asignamos los valores directamente. DAO.Recordset maneja los campos Memo sin problemas.
        rcd!refContratoInspeccionOficial = Nz(objPC.refContratoInspeccionOficial, "")
        rcd!refSuministrador = Nz(objPC.refSuministrador, "")
        rcd!denominacionContrato = Nz(objPC.denominacionContrato, "")
        rcd!SuministradorNombreDir = Nz(objPC.SuministradorNombreDir, "")
        rcd!objetoContrato = Nz(objPC.objetoContrato, "")
    End If
    
    ' 2. Guardamos los cambios (ya sea el AddNew o el Edit)
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Set rcd = Nothing
    Set db = Nothing
    Exit Sub

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCRepositorio.ActualizarDatosGenerales"
    errObj.AddToCallStack "idSolicitud=" & Nz(objPC.idSolicitud, 0)
    errObj.Raise
End Sub
' ==========================================================================
' MÓDULO: DatosPCRepositorio.bas (REFACTORIZADO CON RECORDSET PARA MEMO)
' ==========================================================================

Public Sub ActualizarPropuesta(ByRef objPC As DatosPC, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    If db Is Nothing Then
        Set dbConexion = getdb()
    Else
        Set dbConexion = db
    End If

    ' 1. Construir una consulta SELECT parametrizada para obtener el registro a editar.
    sql = "SELECT * FROM tbDatosPC WHERE idSolicitud = [p_id] ORDER BY idDatosPC;"
    
    Dim qdf As DAO.QueryDef
    Set qdf = dbConexion.CreateQueryDef("", sql)
    qdf.Parameters("p_id").value = objPC.idSolicitud
    
    ' 2. Abrir el Recordset en modo Dynaset (editable).
    Set rs = qdf.OpenRecordset(dbOpenDynaset)

    ' 3. Verificar que el registro existe y entrar en modo edición.
    If rs.EOF Then
        Dim ce As New CondorError
        ce.Create 513, "No existe registro en tbDatosPC para esta solicitud. Debes guardar primero el bloque 'Datos Generales'.", "DatosPCRepositorio.ActualizarPropuesta"
        ce.AddToCallStack "idSolicitud=" & Nz(objPC.idSolicitud, 0)
        ce.Raise
    End If
    
    rs.Edit
    
    ' 4. Asignar los valores a los campos. Esto funciona sin problemas para campos MEMO.
    rs!descripcionMaterialAfectado = Nz(objPC.descripcionMaterialAfectado, "")
    rs!numPlanoEspecificacion = Nz(objPC.numPlanoEspecificacion, "")
    rs!descripcionPropuestaCambio = Nz(objPC.descripcionPropuestaCambio, "")
    
    ' 5. Guardar los cambios.
    rs.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set dbConexion = Nothing
    Exit Sub

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCRepositorio.ActualizarPropuesta"
    errObj.AddToCallStack "idSolicitud=" & Nz(objPC.idSolicitud, 0)
    errObj.Raise
End Sub

' ==========================================================================
' MÓDULO: DatosPCRepositorio.bas
' ACCIÓN: REEMPLAZAR ActualizarImpacto para usar DAO.Recordset.
' ==========================================================================
Public Sub ActualizarImpacto(ByRef objPC As DatosPC, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    If db Is Nothing Then
        Set dbConexion = getdb()
    Else
        Set dbConexion = db
    End If

    ' 1. Abrimos un Recordset filtrado por idSolicitud para editar.
    sql = "SELECT * FROM tbDatosPC WHERE idSolicitud = " & objPC.idSolicitud & " ORDER BY idDatosPC"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)

    ' 2. Verificamos que el registro base exista.
    If rcd.EOF Then
        Dim ce As New CondorError
        ce.Create 513, "No existe registro en tbDatosPC para esta solicitud. Debes guardar primero el bloque 'Datos Generales'.", "DatosPCRepositorio.ActualizarImpacto"
        ce.AddToCallStack "idSolicitud=" & Nz(objPC.idSolicitud, 0)
        ce.Raise
    End If
    
    ' 3. Entramos en modo edición y asignamos todos los valores.
    rcd.Edit
    rcd!motivoCorregirDeficiencias = objPC.motivoCorregirDeficiencias
    rcd!motivoMejorarCapacidad = objPC.motivoMejorarCapacidad
    rcd!motivoAumentarNacionalizacion = objPC.motivoAumentarNacionalizacion
    rcd!motivoMejorarSeguridad = objPC.motivoMejorarSeguridad
    rcd!motivoMejorarFiabilidad = objPC.motivoMejorarFiabilidad
    rcd!motivoMejorarCosteEficacia = objPC.motivoMejorarCosteEficacia
    rcd!motivoOtros = objPC.motivoOtros
    rcd!motivoOtrosDetalle = Nz(objPC.motivoOtrosDetalle, "") ' <-- Campo Memo
    
    rcd!incidenciaCoste = Nz(objPC.incidenciaCoste, "")
    rcd!incidenciaPlazo = Nz(objPC.incidenciaPlazo, "")
    rcd!incidenciaSeguridad = objPC.incidenciaSeguridad
    rcd!incidenciaFiabilidad = objPC.incidenciaFiabilidad
    rcd!incidenciaMantenibilidad = objPC.incidenciaMantenibilidad
    rcd!incidenciaIntercambiabilidad = objPC.incidenciaIntercambiabilidad
    rcd!incidenciaVidaUtilAlmacen = objPC.incidenciaVidaUtilAlmacen
    rcd!incidenciaFuncionamientoFuncion = objPC.incidenciaFuncionamientoFuncion
    
    rcd!impactoClasificacion = Nz(objPC.impactoClasificacion, "")
    rcd!CambioAfectaAMaterial = Nz(objPC.CambioAfectaAMaterial, "")
    
    ' 4. Guardamos los cambios.
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Set rcd = Nothing
    Set dbConexion = Nothing
    Exit Sub

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCRepositorio.ActualizarImpacto"
    errObj.AddToCallStack "idSolicitud=" & Nz(objPC.idSolicitud, 0)
    errObj.Raise
End Sub



' ==========================================================================
' MÓDULO: DatosPCRepositorio.bas
' ACCIÓN: REEMPLAZAR ActualizarAprobacionSuministrador para usar DAO.Recordset por consistencia.
' ==========================================================================
Public Sub ActualizarAprobacionSuministrador(ByRef objPC As DatosPC, Optional db As DAO.Database)
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    If db Is Nothing Then Set db = getdb()

    ' 1. Abrimos un Recordset filtrado por idSolicitud para editar.
    sql = "SELECT * FROM tbDatosPC WHERE idSolicitud = " & objPC.idSolicitud & " ORDER BY idDatosPC"
    Set rcd = db.OpenRecordset(sql, dbOpenDynaset)

    ' 2. Verificamos que el registro base exista.
    If rcd.EOF Then
        Dim ce As New CondorError
        ce.Create 513, "No existe registro en tbDatosPC para esta solicitud. Debes guardar primero el bloque 'Datos Generales'.", "DatosPCRepositorio.ActualizarAprobacionSuministrador"
        ce.AddToCallStack "idSolicitud=" & Nz(objPC.idSolicitud, 0)
        ce.Raise
    End If
    
    ' 3. Entramos en modo edición y asignamos los valores.
    rcd.Edit
    rcd!firmaOficinaTecnicaNombre = Nz(objPC.firmaOficinaTecnicaNombre, "")
    rcd!firmaRepSuministradorNombre = Nz(objPC.firmaRepSuministradorNombre, "")
    
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Set rcd = Nothing
    Set db = Nothing
    Exit Sub

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCRepositorio.ActualizarAprobacionSuministrador"
    errObj.AddToCallStack "idSolicitud=" & Nz(objPC.idSolicitud, 0)
    errObj.Raise
End Sub

' ==========================================================================
' MÓDULO: DatosPCRepositorio.bas
' ACCIÓN: Actualización de datos RAC con soporte transaccional.
' ==========================================================================
Public Sub ActualizarDictamenRAC(ByRef objPC As DatosPC, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    ' 1. GESTIÓN DE CONEXIÓN DUAL
    ' Si nos llega una transacción abierta (db), la usamos. Si no, abrimos una nueva.
    If db Is Nothing Then
        Set dbConexion = getdb()
    Else
        Set dbConexion = db
    End If
    
    sql = "SELECT * FROM tbDatosPC WHERE idSolicitud = " & objPC.idSolicitud & " ORDER BY idDatosPC"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)

    If rcd.EOF Then
        Err.Raise 513, "ActualizarDictamenRAC", "No se encontró el registro de datos PC."
    End If
    
    rcd.Edit
    
    ' 2. CAMPOS INFORMATIVOS
    rcd!racCodigo = Nz(objPC.racCodigo, "")
    rcd!observacionesRAC = Nz(objPC.observacionesRAC, "")
    rcd!racNombre = Nz(objPC.racNombre, "")
    
    ' 3. CAMPOS DE DECISIÓN (CRÍTICOS)
    rcd!racDecision = Nz(objPC.racDecision, "")
    rcd!racRechazoMotivos = Nz(objPC.racRechazoMotivos, "")
    
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCRepositorio.ActualizarDictamenRAC"
    errObj.AddToCallStack "idSolicitud: " & objPC.idSolicitud
    errObj.Raise
End Sub

Public Sub ActualizarDecisionFinal(ByRef objPC As DatosPC, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    ' Lógica de conexión dual: Usamos la transaccional si viene, o una nueva si no.
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosPC WHERE idSolicitud = [p_id] ORDER BY idDatosPC;"
    Dim qdf As DAO.QueryDef
    Set qdf = dbConexion.CreateQueryDef("", sql)
    qdf.Parameters("p_id").value = objPC.idSolicitud
    Set rs = qdf.OpenRecordset(dbOpenDynaset)

    If rs.EOF Then
        Dim ce As New CondorError
        ce.Create 513, "No existe registro en tbDatosPC.", "DatosPCRepositorio.ActualizarDecisionFinal"
        ce.Raise
    End If
    
    rs.Edit
    rs!obsAprobacionAutoridadDiseno = Nz(objPC.obsAprobacionAutoridadDiseno)
    rs!NombreAutoridadDiseno = Nz(objPC.NombreAutoridadDiseno)
    rs!decisionFinal = Nz(objPC.decisionFinal)
    rs!obsDecisionFinal = Nz(objPC.obsDecisionFinal)
    rs!NombreFirmanteFinal = Nz(objPC.NombreFirmanteFinal)
    rs.Update
    
LimpiarYSalir:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCRepositorio.ActualizarDecisionFinal"
    errObj.Raise
End Sub


Public Sub EliminarPorIdSolicitud(ByVal idSol As Long, ByRef db As DAO.Database)
    ' RESPONSABILIDAD: Eliminar el registro de tbDatosPC asociado a una solicitud.
    Dim sql As String
    Dim params As Object
    On Error GoTo Errores
    
    sql = "DELETE FROM tbDatosPC WHERE idSolicitud = [p_IDSol];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_IDSol", idSol
    
    ' Se ejecuta la acción usando la conexión a la base de datos transaccional
    ' que se pasa como parámetro desde el servicio.
    Call RepositorioComun.EjecutarAccion(sql, params, db)
    
LimpiarYSalir:
    Exit Sub
Errores:
    ' Si falla, se propaga el error para que el servicio pueda hacer Rollback.
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCRepositorio.EliminarPorIdSolicitud"
    errObj.AddToCallStack "idSolicitud: " & idSol
    errObj.Raise
End Sub
Public Sub LimpiarDictamenRAC(ByVal idSolicitud As Long, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    ' Gestión de conexión dual (Transaccional o Global)
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosPC WHERE idSolicitud = " & idSolicitud & " ORDER BY idDatosPC"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)

    If Not rcd.EOF Then
        rcd.Edit
        ' 1. Limpieza de datos generales del RAC
        rcd!racCodigo = Null
        rcd!observacionesRAC = Null
        rcd!racNombre = Null
        
        ' 2. CRÍTICO: Limpieza de los campos de decisión/rechazo
        rcd!racDecision = Null
        rcd!racRechazoMotivos = Null
        
        rcd.Update
    End If

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCRepositorio.LimpiarDictamenRAC"
    errObj.AddToCallStack "idSolicitud: " & idSolicitud
    errObj.Raise
End Sub

Public Sub LimpiarAprobacionSuministrador(ByVal idSolicitud As Long, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosPC WHERE idSolicitud = " & idSolicitud
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If Not rcd.EOF Then
        rcd.Edit
        rcd!firmaOficinaTecnicaNombre = Null
        rcd!firmaRepSuministradorNombre = Null
        rcd.Update
    End If

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCRepositorio.LimpiarAprobacionSuministrador"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud: errObj.Raise
End Sub

Public Sub LimpiarDecisionFinal(ByVal idSolicitud As Long, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosPC WHERE idSolicitud = " & idSolicitud
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If Not rcd.EOF Then
        rcd.Edit
        rcd!obsAprobacionAutoridadDiseno = Null
        rcd!NombreAutoridadDiseno = Null
        rcd!decisionFinal = Null
        rcd!obsDecisionFinal = Null
        rcd!NombreFirmanteFinal = Null
        rcd.Update
    End If

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCRepositorio.LimpiarDecisionFinal"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud: errObj.Raise
End Sub


