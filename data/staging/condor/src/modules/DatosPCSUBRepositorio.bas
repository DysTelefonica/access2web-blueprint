Attribute VB_Name = "DatosPCSUBRepositorio"
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: DatosPCSUBRepositorio.bas
' RESPONSABILIDAD: Repositorio DAO para solicitudes PC-Subcontratista.
' CLONE DE: DatosPCRepositorio.bas (misma estructura, cambia tbDatosPC ? tbDatosPCSUB)
' ==========================================================================

' === MÉTODO: Guardar (Insert or Update) ===
Public Sub Guardar(ByRef objPCSUB As DatosPCSUB, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String

    On Error GoTo Errores

    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db

    sql = "SELECT * FROM tbDatosPCSUB WHERE idSolicitud = " & objPCSUB.idSolicitud & " ORDER BY idDatosPCSUB"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)

    If rcd.EOF Then
        ' MODO INSERT: El registro no existe
        rcd.AddNew
        If objPCSUB.idDatosPCSUB = 0 Then
            objPCSUB.idDatosPCSUB = getSiguienteIDDatosPCSUB(dbConexion)
        End If
        Call RepositorioComun.RellenarRecordsetDesdeObjeto(rcd, objPCSUB)
    Else
        ' MODO UPDATE: El registro ya existe
        rcd.Edit
        Call RepositorioComun.RellenarRecordsetDesdeObjeto(rcd, objPCSUB, True)
    End If

    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBRepositorio.Guardar"
    errObj.AddToCallStack "idSolicitud: " & objPCSUB.idSolicitud
    errObj.Raise
End Sub

' === MÉTODO: GetById ===
Public Function getPorIdSolicitud(ByVal idSol As Long, Optional db As DAO.Database) As DatosPCSUB
    Dim sql As String, params As Object, dbConexion As DAO.Database
    On Error GoTo Errores

    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db

    sql = "SELECT * FROM tbDatosPCSUB WHERE idSolicitud = [p_IDSol] ORDER BY idDatosPCSUB;"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_IDSol", idSol

    Set getPorIdSolicitud = RepositorioComun.HidratarEntidadDesdeSQL(sql, "DatosPCSUB", dbConexion, params)

    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBRepositorio.getPorIdSolicitud"
    errObj.Raise
End Function

' === MÉTODO: GetById (alias) ===
Public Function GetById(ByVal idSol As Long, Optional db As DAO.Database) As DatosPCSUB
    Set GetById = getPorIdSolicitud(idSol, db)
End Function

Private Function getSiguienteIDDatosPCSUB(ByRef db As DAO.Database) As Long
    Dim rcd As DAO.Recordset
    Dim maxID As Long
    On Error GoTo Errores

    Set rcd = db.OpenRecordset("SELECT Max(idDatosPCSUB) AS MaxID FROM tbDatosPCSUB")

    If Not rcd.EOF Then
        maxID = Nz(rcd!maxID, 0)
    End If

    getSiguienteIDDatosPCSUB = maxID + 1

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBRepositorio.getSiguienteIDDatosPCSUB"
    errObj.Raise
End Function

' === ACTUALIZAR DATOS GENERALES ===
Public Sub ActualizarDatosGenerales(ByRef objPCSUB As DatosPCSUB)
    Dim db As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String

    On Error GoTo Errores

    Set db = getdb()

    sql = "SELECT * FROM tbDatosPCSUB WHERE idSolicitud = " & objPCSUB.idSolicitud & " ORDER BY idDatosPCSUB"
    Set rcd = db.OpenRecordset(sql, dbOpenDynaset)

    If rcd.EOF Then
        rcd.AddNew
        objPCSUB.idDatosPCSUB = getSiguienteIDDatosPCSUB(db)
        rcd!idDatosPCSUB = objPCSUB.idDatosPCSUB
        rcd!idSolicitud = objPCSUB.idSolicitud
        Call RepositorioComun.RellenarRecordsetDesdeObjeto(rcd, objPCSUB, True)
    Else
        rcd.Edit
        rcd!refContratoInspeccionOficial = Nz(objPCSUB.refContratoInspeccionOficial, "")
        rcd!refSubSuministrador = Nz(objPCSUB.refSubSuministrador, "")
        rcd!denominacionContrato = Nz(objPCSUB.denominacionContrato, "")
        rcd!SubsuministradorNombreDir = Nz(objPCSUB.SubsuministradorNombreDir, "")
        rcd!objetoContrato = Nz(objPCSUB.objetoContrato, "")
    End If

    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Set rcd = Nothing
    Set db = Nothing
    Exit Sub

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBRepositorio.ActualizarDatosGenerales"
    errObj.AddToCallStack "idSolicitud=" & Nz(objPCSUB.idSolicitud, 0)
    errObj.Raise
End Sub

' === ACTUALIZAR PROPUESTA ===
Public Sub ActualizarPropuesta(ByRef objPCSUB As DatosPCSUB, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    On Error GoTo Errores

    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db

    sql = "SELECT * FROM tbDatosPCSUB WHERE idSolicitud = [p_id] ORDER BY idDatosPCSUB;"

    Dim qdf As DAO.QueryDef
    Set qdf = dbConexion.CreateQueryDef("", sql)
    qdf.Parameters("p_id").value = objPCSUB.idSolicitud

    Set rs = qdf.OpenRecordset(dbOpenDynaset)

    If rs.EOF Then
        Dim ce As New CondorError
        ce.Create 513, "No existe registro en tbDatosPCSUB para esta solicitud.", "DatosPCSUBRepositorio.ActualizarPropuesta"
        ce.AddToCallStack "idSolicitud=" & Nz(objPCSUB.idSolicitud, 0)
        ce.Raise
    End If

    rs.Edit
    rs!descripcionMaterialAfectado = Nz(objPCSUB.descripcionMaterialAfectado, "")
    rs!numPlanoEspecificacion = Nz(objPCSUB.numPlanoEspecificacion, "")
    rs!descripcionPropuestaCambio = Nz(objPCSUB.descripcionPropuestaCambio, "")
    rs!descripcionPropuestaCambioCont = Nz(objPCSUB.descripcionPropuestaCambioCont, "")
    rs.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set dbConexion = Nothing
    Exit Sub

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBRepositorio.ActualizarPropuesta"
    errObj.AddToCallStack "idSolicitud=" & Nz(objPCSUB.idSolicitud, 0)
    errObj.Raise
End Sub

' === ACTUALIZAR IMPACTO ===
Public Sub ActualizarImpacto(ByRef objPCSUB As DatosPCSUB, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String

    On Error GoTo Errores

    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db

    sql = "SELECT * FROM tbDatosPCSUB WHERE idSolicitud = " & objPCSUB.idSolicitud & " ORDER BY idDatosPCSUB"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)

    If rcd.EOF Then
        Dim ce As New CondorError
        ce.Create 513, "No existe registro en tbDatosPCSUB.", "DatosPCSUBRepositorio.ActualizarImpacto"
        ce.AddToCallStack "idSolicitud=" & Nz(objPCSUB.idSolicitud, 0)
        ce.Raise
    End If

    rcd.Edit
    rcd!motivoCorregirDeficiencias = objPCSUB.motivoCorregirDeficiencias
    rcd!motivoMejorarCapacidad = objPCSUB.motivoMejorarCapacidad
    rcd!motivoAumentarNacionalizacion = objPCSUB.motivoAumentarNacionalizacion
    rcd!motivoMejorarSeguridad = objPCSUB.motivoMejorarSeguridad
    rcd!motivoMejorarFiabilidad = objPCSUB.motivoMejorarFiabilidad
    rcd!motivoMejorarCosteEficacia = objPCSUB.motivoMejorarCosteEficacia
    rcd!motivoOtros = objPCSUB.motivoOtros
    rcd!motivoOtrosDetalle = Nz(objPCSUB.motivoOtrosDetalle, "")

    rcd!incidenciaCoste = Nz(objPCSUB.incidenciaCoste, "")
    rcd!incidenciaPlazo = Nz(objPCSUB.incidenciaPlazo, "")
    rcd!incidenciaSeguridad = objPCSUB.incidenciaSeguridad
    rcd!incidenciaFiabilidad = objPCSUB.incidenciaFiabilidad
    rcd!incidenciaMantenibilidad = objPCSUB.incidenciaMantenibilidad
    rcd!incidenciaIntercambiabilidad = objPCSUB.incidenciaIntercambiabilidad
    rcd!incidenciaVidaUtilAlmacen = objPCSUB.incidenciaVidaUtilAlmacen
    rcd!incidenciaFuncionamientoFuncion = objPCSUB.incidenciaFuncionamientoFuncion

    rcd!impactoClasificacion = Nz(objPCSUB.impactoClasificacion, "")
    rcd!CambioAfectaAMaterial = Nz(objPCSUB.CambioAfectaAMaterial, "")

    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Set rcd = Nothing
    Set dbConexion = Nothing
    Exit Sub

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBRepositorio.ActualizarImpacto"
    errObj.AddToCallStack "idSolicitud=" & Nz(objPCSUB.idSolicitud, 0)
    errObj.Raise
End Sub

' === ACTUALIZAR APROBACIÓN SUMINISTRADOR ===
Public Sub ActualizarAprobacionSuministrador(ByRef objPCSUB As DatosPCSUB, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String

    On Error GoTo Errores
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db

    sql = "SELECT * FROM tbDatosPCSUB WHERE idSolicitud = " & objPCSUB.idSolicitud & " ORDER BY idDatosPCSUB"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)

    If rcd.EOF Then
        Dim ce As New CondorError
        ce.Create 513, "No existe registro en tbDatosPCSUB.", "DatosPCSUBRepositorio.ActualizarAprobacionSuministrador"
        ce.AddToCallStack "idSolicitud=" & Nz(objPCSUB.idSolicitud, 0)
        ce.Raise
    End If

    rcd.Edit
    rcd!firmaOficinaTecnicaSubSuministradorNombre = Nz(objPCSUB.firmaOficinaTecnicaSubSuministradorNombre, "")
    rcd!firmaRepSubSuministradorNombre = Nz(objPCSUB.firmaRepSubSuministradorNombre, "")
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Set rcd = Nothing
    Set dbConexion = Nothing
    Exit Sub

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBRepositorio.ActualizarAprobacionSuministrador"
    errObj.AddToCallStack "idSolicitud=" & Nz(objPCSUB.idSolicitud, 0)
    errObj.Raise
End Sub

' === ACTUALIZAR DICTAMEN RAC (racDelegado*) ===
Public Sub ActualizarDictamenRAC(ByRef objPCSUB As DatosPCSUB, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String

    On Error GoTo Errores

    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db

    sql = "SELECT * FROM tbDatosPCSUB WHERE idSolicitud = " & objPCSUB.idSolicitud & " ORDER BY idDatosPCSUB"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)

    If rcd.EOF Then
        Err.Raise 513, "ActualizarDictamenRAC", "No se encontró el registro de datos PCSUB."
    End If

    rcd.Edit

    ' Campos RAC estándar (tbDatosPCSUB no tiene racDelegado*)
    rcd!racCodigo = Nz(objPCSUB.racCodigo, "")
    rcd!observacionesRAC = Nz(objPCSUB.observacionesRAC, "")
    rcd!racNombre = Nz(objPCSUB.racNombre, "")
    rcd!racDecision = Nz(objPCSUB.racDecision, "")
    rcd!racRechazoMotivos = Nz(objPCSUB.racRechazoMotivos, "")

    ' Campos RAC Delegador
    rcd!observacionesRACDelegador = Nz(objPCSUB.observacionesRACDelegador, "")
    rcd!racNombreDelegador = Nz(objPCSUB.racNombreDelegador, "")

    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBRepositorio.ActualizarDictamenRAC"
    errObj.AddToCallStack "idSolicitud: " & objPCSUB.idSolicitud
    errObj.Raise
End Sub

' === ACTUALIZAR DECISIÓN FINAL ===
Public Sub ActualizarDecisionFinal(ByRef objPCSUB As DatosPCSUB, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    On Error GoTo Errores

    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db

    sql = "SELECT * FROM tbDatosPCSUB WHERE idSolicitud = [p_id] ORDER BY idDatosPCSUB;"
    Dim qdf As DAO.QueryDef
    Set qdf = dbConexion.CreateQueryDef("", sql)
    qdf.Parameters("p_id").value = objPCSUB.idSolicitud
    Set rs = qdf.OpenRecordset(dbOpenDynaset)

    If rs.EOF Then
        Dim ce As New CondorError
        ce.Create 513, "No existe registro en tbDatosPCSUB.", "DatosPCSUBRepositorio.ActualizarDecisionFinal"
        ce.Raise
    End If

    rs.Edit
    rs!obsAprobacionAutoridadDiseno = Nz(objPCSUB.obsAprobacionAutoridadDiseno)
    rs!NombreAutoridadDiseno = Nz(objPCSUB.NombreAutoridadDiseno)
    rs!decisionFinal = Nz(objPCSUB.decisionFinal)
    rs!obsDecisionFinal = Nz(objPCSUB.obsDecisionFinal)
    rs!NombreFirmanteFinal = Nz(objPCSUB.NombreFirmanteFinal)
    rs.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBRepositorio.ActualizarDecisionFinal"
    errObj.Raise
End Sub

' === ELIMINAR POR IDSOLICITUD ===
Public Sub EliminarPorIdSolicitud(ByVal idSol As Long, ByRef db As DAO.Database)
    Dim sql As String
    Dim params As Object
    On Error GoTo Errores

    sql = "DELETE FROM tbDatosPCSUB WHERE idSolicitud = [p_IDSol];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_IDSol", idSol

    Call RepositorioComun.EjecutarAccion(sql, params, db)

LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBRepositorio.EliminarPorIdSolicitud"
    errObj.AddToCallStack "idSolicitud: " & idSol
    errObj.Raise
End Sub

' === LIMPIAR DICTAMEN RAC ===
Public Sub LimpiarDictamenRAC(ByVal idSolicitud As Long, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String

    On Error GoTo Errores

    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db

    sql = "SELECT * FROM tbDatosPCSUB WHERE idSolicitud = " & idSolicitud & " ORDER BY idDatosPCSUB"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)

    If Not rcd.EOF Then
        rcd.Edit
        rcd!racCodigo = Null
        rcd!observacionesRAC = Null
        rcd!racNombre = Null
        rcd!racDecision = Null
        rcd!racRechazoMotivos = Null
        rcd!observacionesRACDelegador = Null
        rcd!racNombreDelegador = Null
        rcd.Update
    End If

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBRepositorio.LimpiarDictamenRAC"
    errObj.AddToCallStack "idSolicitud: " & idSolicitud
    errObj.Raise
End Sub

' === LIMPIAR APROBACIÓN SUMINISTRADOR ===
Public Sub LimpiarAprobacionSuministrador(ByVal idSolicitud As Long, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String

    On Error GoTo Errores

    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db

    sql = "SELECT * FROM tbDatosPCSUB WHERE idSolicitud = " & idSolicitud
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)

    If Not rcd.EOF Then
        rcd.Edit
        rcd!firmaOficinaTecnicaSubSuministradorNombre = Null
        rcd!firmaRepSubSuministradorNombre = Null
        rcd.Update
    End If

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBRepositorio.LimpiarAprobacionSuministrador"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud: errObj.Raise
End Sub

' === LIMPIAR DECISIÓN FINAL ===
Public Sub LimpiarDecisionFinal(ByVal idSolicitud As Long, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String

    On Error GoTo Errores

    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db

    sql = "SELECT * FROM tbDatosPCSUB WHERE idSolicitud = " & idSolicitud
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
    errObj.Create Err.Number, Err.description, "DatosPCSUBRepositorio.LimpiarDecisionFinal"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud: errObj.Raise
End Sub
