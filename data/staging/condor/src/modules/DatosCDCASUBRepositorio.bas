Attribute VB_Name = "DatosCDCASUBRepositorio"

Option Compare Database
Option Explicit

Private Function getSiguienteIDDatosCDCASUB(ByRef db As DAO.Database) As Long
    Dim rcd As DAO.Recordset
    Dim maxID As Long
    On Error GoTo Errores
    
    Set rcd = db.OpenRecordset("SELECT Max(idDatosCDCASUB) AS MaxID FROM tbDatosCDCASUB")
    If Not rcd.EOF Then
        maxID = Nz(rcd!maxID, 0)
    End If
    
    getSiguienteIDDatosCDCASUB = maxID + 1
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
Errores:
    Dim errObjID As New CondorError
    errObjID.Create Err.Number, Err.description, "DatosCDCASUBRepositorio.getSiguienteIDDatosCDCASUB"
    errObjID.Raise
End Function

Public Sub Guardar(ByRef obj As DatosCDCASUB, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database, rcd As DAO.Recordset, sql As String
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosCDCASUB WHERE idSolicitud = " & obj.idSolicitud & " ORDER BY idDatosCDCASUB"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If rcd.EOF Then
        If obj.idDatosCDCASUB = 0 Then
            obj.idDatosCDCASUB = getSiguienteIDDatosCDCASUB(dbConexion)
        End If
        rcd.AddNew
        rcd!idDatosCDCASUB = obj.idDatosCDCASUB
        rcd!idSolicitud = obj.idSolicitud
        Call RepositorioComun.RellenarRecordsetDesdeObjeto(rcd, obj, True)
        rcd!idDatosCDCASUB = obj.idDatosCDCASUB
    Else
        rcd.Edit
        Call RepositorioComun.RellenarRecordsetDesdeObjeto(rcd, obj, True)
    End If
    rcd.Update
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCASUBRepositorio.Guardar"
    errObj.Raise
End Sub

Public Function getPorIdSolicitud(ByVal idSol As Long, Optional db As DAO.Database) As DatosCDCASUB
    Dim sql As String, params As Object, dbConexion As DAO.Database
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosCDCASUB WHERE idSolicitud = [p_ID] ORDER BY idDatosCDCASUB;"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", idSol
    
    Set getPorIdSolicitud = RepositorioComun.HidratarEntidadDesdeSQL(sql, "DatosCDCASUB", dbConexion, params)
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCASUBRepositorio.getPorIdSolicitud"
    errObj.Raise
End Function

Public Sub ActualizarDatosGenerales(ByRef obj As DatosCDCASUB, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    ' Gestión de conexión dual
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosCDCASUB WHERE idSolicitud = " & obj.idSolicitud & " ORDER BY idDatosCDCASUB"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If rcd.EOF Then
        rcd.AddNew
        obj.idDatosCDCASUB = getSiguienteIDDatosCDCASUB(dbConexion)
        rcd!idDatosCDCASUB = obj.idDatosCDCASUB
        rcd!idSolicitud = obj.idSolicitud
        
        rcd!refSuministrador = Nz(obj.refSuministrador, "")
        rcd!refSubSuministrador = Nz(obj.refSubSuministrador, "")
        rcd!suministradorPrincipalNombreDir = Nz(obj.suministradorPrincipalNombreDir, "")
        rcd!subSuministradorNombreDir = Nz(obj.subSuministradorNombreDir, "")
        rcd!refDesviacionesPrevias = Nz(obj.refDesviacionesPrevias, "")
        rcd!requiereModificacionContrato = obj.requiereModificacionContrato
    Else
        rcd.Edit
        rcd!refSuministrador = Nz(obj.refSuministrador, "")
        rcd!refSubSuministrador = Nz(obj.refSubSuministrador, "")
        rcd!suministradorPrincipalNombreDir = Nz(obj.suministradorPrincipalNombreDir, "")
        rcd!subSuministradorNombreDir = Nz(obj.subSuministradorNombreDir, "")
        rcd!refDesviacionesPrevias = Nz(obj.refDesviacionesPrevias, "")
        rcd!requiereModificacionContrato = obj.requiereModificacionContrato
    End If
    
    rcd.Update
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObjDG As New CondorError
    errObjDG.Create Err.Number, Err.description, "DatosCDCASUBRepositorio.ActualizarDatosGenerales"
    errObjDG.AddToCallStack "idSolicitud=" & Nz(obj.idSolicitud, 0)
    errObjDG.Raise
End Sub

Public Sub ActualizarPropuesta(ByRef obj As DatosCDCASUB, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    If db Is Nothing Then
        Set dbConexion = getdb()
    Else
        Set dbConexion = db
    End If
    
    sql = "SELECT * FROM tbDatosCDCASUB WHERE idSolicitud = " & obj.idSolicitud & " ORDER BY idDatosCDCASUB"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If rcd.EOF Then Err.Raise 513, , "No se encontro el registro base CDCASUB."
    
    rcd.Edit
    rcd!identificacionMaterial = Nz(obj.identificacionMaterial, "")
    rcd!numPlanoEspecificacion = Nz(obj.numPlanoEspecificacion, "")
    rcd!cantidadPeriodo = Nz(obj.cantidadPeriodo, "")
    rcd!numSerieLote = Nz(obj.numSerieLote, "")
    rcd!causaNC = Nz(obj.causaNC, "")
    rcd!descripcionImpactoNC = Nz(obj.descripcionImpactoNC, "")
    rcd!descripcionImpactoNCCont = Nz(obj.descripcionImpactoNCCont, "")
    rcd.Update
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObjP As New CondorError
    errObjP.Create Err.Number, Err.description, "DatosCDCASUBRepositorio.ActualizarPropuesta"
    errObjP.Raise
End Sub

Public Sub ActualizarImpacto(ByRef obj As DatosCDCASUB, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    If db Is Nothing Then
        Set dbConexion = getdb()
    Else
        Set dbConexion = db
    End If
    
    sql = "SELECT * FROM tbDatosCDCASUB WHERE idSolicitud = " & obj.idSolicitud & " ORDER BY idDatosCDCASUB"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If rcd.EOF Then Err.Raise 513, , "No se encontro el registro base CDCASUB."
    
    rcd.Edit
    rcd!afectaPrestaciones = obj.afectaPrestaciones
    rcd!afectaSeguridad = obj.afectaSeguridad
    rcd!afectaFiabilidad = obj.afectaFiabilidad
    rcd!afectaVidaUtil = obj.afectaVidaUtil
    rcd!afectaMedioambiente = obj.afectaMedioambiente
    rcd!afectaMantenibilidad = obj.afectaMantenibilidad
    rcd!afectaIntercambiabilidad = obj.afectaIntercambiabilidad
    rcd!afectaApariencia = obj.afectaApariencia
    rcd!afectaOtros = obj.afectaOtros
    
    rcd!impactoCoste = Nz(obj.impactoCoste, "")
    rcd!clasificacionNC = Nz(obj.clasificacionNC, "")
    rcd!esSubSuministradorAD = obj.esSubSuministradorAD
    rcd!identificacionAutoridadDiseno = Nz(obj.identificacionAutoridadDiseno, "")
    rcd!efectoFechaEntrega = Nz(obj.efectoFechaEntrega, "")
    
    rcd.Update
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObjI As New CondorError
    errObjI.Create Err.Number, Err.description, "DatosCDCASUBRepositorio.ActualizarImpacto"
    errObjI.Raise
End Sub

Public Sub ActualizarAprobacionSuministrador(ByRef obj As DatosCDCASUB, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    ' Gestión de conexión dual
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosCDCASUB WHERE idSolicitud = " & obj.idSolicitud & " ORDER BY idDatosCDCASUB"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If rcd.EOF Then Err.Raise 513, , "No se encontro el registro base CDCASUB."
    
    rcd.Edit
    rcd!firmaAprobacionRespIngenieriaNombre = Nz(obj.firmaAprobacionRespIngenieriaNombre, "")
    rcd!firmaAprobacionRespProduccionNombre = Nz(obj.firmaAprobacionRespProduccionNombre, "")
    rcd!firmaAprobacionRespCalidadNombre = Nz(obj.firmaAprobacionRespCalidadNombre, "")
    rcd!firmaAprobacionRespDisenioNombre = Nz(obj.firmaAprobacionRespDisenioNombre, "")
    rcd!firmaAprobacionRepresentanteSumNombre = Nz(obj.firmaAprobacionRepresentanteSumNombre, "")
    
    rcd.Update
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObjA As New CondorError
    errObjA.Create Err.Number, Err.description, "DatosCDCASUBRepositorio.ActualizarAprobacion"
    errObjA.Raise
End Sub

Public Sub ActualizarDictamenRAC(ByRef obj As DatosCDCASUB, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    If db Is Nothing Then
        Set dbConexion = getdb()
    Else
        Set dbConexion = db
    End If
    
    sql = "SELECT * FROM tbDatosCDCASUB WHERE idSolicitud = " & obj.idSolicitud & " ORDER BY idDatosCDCASUB"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If rcd.EOF Then Err.Raise 513, , "No se encontro el registro base CDCASUB."
    
    rcd.Edit
    rcd!racCodigo = Nz(obj.racCodigo, "")
    rcd!observacionesRAC = Nz(obj.observacionesRAC, "")
    rcd!racNombre = Nz(obj.racNombre, "")
    rcd!racDecision = Nz(obj.racDecision, "")
    rcd!racRechazoMotivos = Nz(obj.racRechazoMotivos, "")
    rcd!observacionesRACDelegador = Nz(obj.observacionesRACDelegador, "")
    rcd!racNombreDelegador = Nz(obj.racNombreDelegador, "")
    
    rcd.Update
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObjR As New CondorError
    errObjR.Create Err.Number, Err.description, "DatosCDCASUBRepositorio.ActualizarDictamenRAC"
    errObjR.Raise
End Sub

Public Sub ActualizarDecisionFinal(ByRef obj As DatosCDCASUB, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    If db Is Nothing Then
        Set dbConexion = getdb()
    Else
        Set dbConexion = db
    End If
    
    sql = "SELECT * FROM tbDatosCDCASUB WHERE idSolicitud = " & obj.idSolicitud & " ORDER BY idDatosCDCASUB"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If rcd.EOF Then Err.Raise 513, , "No se encontro el registro base CDCASUB."
    
    rcd.Edit
    rcd!decisionFinal = Nz(obj.decisionFinal, "")
    rcd!observacionesFinales = Nz(obj.observacionesFinales, "")
    rcd!NombreFirmanteFinal = Nz(obj.NombreFirmanteFinal, "")
    
    rcd.Update
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObjDF As New CondorError
    errObjDF.Create Err.Number, Err.description, "DatosCDCASUBRepositorio.ActualizarDecisionFinal"
    errObjDF.Raise
End Sub

Public Sub EliminarPorIdSolicitud(ByVal idSol As Long, ByRef db As DAO.Database)
    Dim sql As String, params As Object
    On Error GoTo Errores
    sql = "DELETE FROM tbDatosCDCASUB WHERE idSolicitud = [p_ID];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_ID", idSol
    Call RepositorioComun.EjecutarAccion(sql, params, db)
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCASUBRepositorio.EliminarPorIdSolicitud"
    errObj.Raise
End Sub

Public Sub LimpiarDictamenRAC(ByVal idSolicitud As Long, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database, rcd As DAO.Recordset, sql As String
    On Error GoTo Errores
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosCDCASUB WHERE idSolicitud = " & idSolicitud & " ORDER BY idDatosCDCASUB"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    If Not rcd.EOF Then
        rcd.Edit
        rcd!racCodigo = Null: rcd!observacionesRAC = Null: rcd!racNombre = Null
        rcd!racDecision = Null: rcd!racRechazoMotivos = Null
        ' Campos extra CDCASUB
        rcd!observacionesRACDelegador = Null: rcd!racNombreDelegador = Null
        rcd.Update
    End If
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCASUBRepositorio.LimpiarDictamenRAC"
    errObj.Raise
End Sub

Public Sub LimpiarAprobacionSuministrador(ByVal idSolicitud As Long, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosCDCASUB WHERE idSolicitud = " & idSolicitud & " ORDER BY idDatosCDCASUB"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If Not rcd.EOF Then
        rcd.Edit
        rcd!firmaAprobacionRespIngenieriaNombre = Null
        rcd!firmaAprobacionRespCalidadNombre = Null
        rcd!firmaAprobacionRespProduccionNombre = Null
        rcd!firmaAprobacionRespDisenioNombre = Null
        rcd!firmaAprobacionRepresentanteSumNombre = Null
        rcd.Update
    End If

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCASUBRepositorio.LimpiarAprobacionSuministrador"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud: errObj.Raise
End Sub

Public Sub LimpiarDecisionFinal(ByVal idSolicitud As Long, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database, rcd As DAO.Recordset, sql As String
    On Error GoTo Errores
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosCDCASUB WHERE idSolicitud = " & idSolicitud & " ORDER BY idDatosCDCASUB"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    If Not rcd.EOF Then
        rcd.Edit
        ' ACTUALIZADO CON FUENTE DE VERDAD
        rcd!decisionFinal = Null
        rcd!observacionesFinales = Null
        rcd!NombreFirmanteFinal = Null
        rcd.Update
    End If
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCASUBRepositorio.LimpiarDecisionFinal"
    errObj.Raise
End Sub


