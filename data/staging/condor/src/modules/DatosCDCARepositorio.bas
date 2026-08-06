Attribute VB_Name = "DatosCDCARepositorio"

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: DatosCDCARepositorio.bas (Refactorizado: Patrón Granular tipo PC)
' RESPONSABILIDAD: Acceso a datos seguro y por secciones para CD/CA.
' ==========================================================================

' --- FUNCIONES AUXILIARES PRIVADAS ---

Private Function getSiguienteIDDatosCDCA(ByRef db As DAO.Database) As Long
    Dim rcd As DAO.Recordset
    Dim maxID As Long
    On Error GoTo Errores
    
    Set rcd = db.OpenRecordset("SELECT Max(idDatosCDCA) AS MaxID FROM tbDatosCDCA")
    
    If Not rcd.EOF Then
        maxID = Nz(rcd!maxID, 0)
    End If
    
    getSiguienteIDDatosCDCA = maxID + 1
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCARepositorio.getSiguienteIDDatosCDCA"
    errObj.Raise
End Function

' --- MÉTODOS DE ACTUALIZACIÓN GRANULAR (UPSERT) ---

Public Sub ActualizarDatosGenerales(ByRef objCDCA As DatosCDCA, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    ' Gestión de conexión dual
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db

    ' Buscamos por idSolicitud para editar o crear
    sql = "SELECT * FROM tbDatosCDCA WHERE idSolicitud = " & objCDCA.idSolicitud & " ORDER BY idDatosCDCA"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)

    If rcd.EOF Then
        ' CREACIÓN (Solo ocurre en la primera pestaña)
        rcd.AddNew
        objCDCA.idDatosCDCA = getSiguienteIDDatosCDCA(dbConexion)
        rcd!idDatosCDCA = objCDCA.idDatosCDCA
        rcd!idSolicitud = objCDCA.idSolicitud
        ' Rellenamos el resto de campos nuevos
        rcd!numContrato = Nz(objCDCA.numContrato, "")
        rcd!refSuministrador = Nz(objCDCA.refSuministrador, "")
        rcd!SuministradorNombreDir = Nz(objCDCA.SuministradorNombreDir, "")
        rcd!refDesviacionesPrevias = Nz(objCDCA.refDesviacionesPrevias, "")
        rcd!requiereModificacionContrato = objCDCA.requiereModificacionContrato
    Else
        ' EDICIÓN
        rcd.Edit
        rcd!numContrato = Nz(objCDCA.numContrato, "")
        rcd!refSuministrador = Nz(objCDCA.refSuministrador, "")
        rcd!SuministradorNombreDir = Nz(objCDCA.SuministradorNombreDir, "")
        rcd!refDesviacionesPrevias = Nz(objCDCA.refDesviacionesPrevias, "")
        rcd!requiereModificacionContrato = objCDCA.requiereModificacionContrato
    End If
    
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCARepositorio.ActualizarDatosGenerales"
    errObj.AddToCallStack "idSolicitud=" & Nz(objCDCA.idSolicitud, 0)
    errObj.Raise
End Sub

Public Sub ActualizarPropuesta(ByRef objCDCA As DatosCDCA, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database, rcd As DAO.Recordset, sql As String
    On Error GoTo Errores
    
    If db Is Nothing Then
        Set dbConexion = getdb()
    Else
        Set dbConexion = db
    End If
    
    sql = "SELECT * FROM tbDatosCDCA WHERE idSolicitud = " & objCDCA.idSolicitud & " ORDER BY idDatosCDCA"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If rcd.EOF Then Err.Raise 513, , "No se encontró el registro base CDCA."
    
    rcd.Edit
    rcd!identificacionMaterial = Nz(objCDCA.identificacionMaterial, "")
    rcd!numPlanoEspecificacion = Nz(objCDCA.numPlanoEspecificacion, "")
    rcd!cantidadPeriodo = Nz(objCDCA.cantidadPeriodo, "")
    rcd!numSerieLote = Nz(objCDCA.numSerieLote, "")
    rcd!causaNC = Nz(objCDCA.causaNC, "")
    rcd!descripcionImpactoNC = Nz(objCDCA.descripcionImpactoNC, "")
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCARepositorio.ActualizarPropuesta"
    errObj.Raise
End Sub

Public Sub ActualizarImpacto(ByRef objCDCA As DatosCDCA, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database, rcd As DAO.Recordset, sql As String
    On Error GoTo Errores
    
    If db Is Nothing Then
        Set dbConexion = getdb()
    Else
        Set dbConexion = db
    End If
    
    sql = "SELECT * FROM tbDatosCDCA WHERE idSolicitud = " & objCDCA.idSolicitud & " ORDER BY idDatosCDCA"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If rcd.EOF Then Err.Raise 513, , "No se encontró el registro base CDCA."
    
    rcd.Edit
    ' Checkboxes
    rcd!afectaPrestaciones = objCDCA.afectaPrestaciones
    rcd!afectaSeguridad = objCDCA.afectaSeguridad
    rcd!afectaFiabilidad = objCDCA.afectaFiabilidad
    rcd!afectaVidaUtil = objCDCA.afectaVidaUtil
    rcd!afectaMedioambiente = objCDCA.afectaMedioambiente
    rcd!afectaMantenibilidad = objCDCA.afectaMantenibilidad
    rcd!afectaIntercambiabilidad = objCDCA.afectaIntercambiabilidad
    rcd!afectaApariencia = objCDCA.afectaApariencia
    rcd!afectaOtros = objCDCA.afectaOtros
    
    
    ' Selectores y Textos
    rcd!impactoCoste = Nz(objCDCA.impactoCoste, "")
    rcd!clasificacionNC = Nz(objCDCA.clasificacionNC, "")
    rcd!esSuministradorAD = objCDCA.esSuministradorAD
    rcd!identificacionAutoridadDiseno = Nz(objCDCA.identificacionAutoridadDiseno, "")
    rcd!efectoFechaEntrega = Nz(objCDCA.efectoFechaEntrega, "")
    
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCARepositorio.ActualizarImpacto"
    errObj.Raise
End Sub

Public Sub ActualizarAprobacionSuministrador(ByRef objCDCA As DatosCDCA, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database, rcd As DAO.Recordset, sql As String
    On Error GoTo Errores
    
    ' Gestión de conexión dual
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosCDCA WHERE idSolicitud = " & objCDCA.idSolicitud & " ORDER BY idDatosCDCA"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If rcd.EOF Then Err.Raise 513, , "No se encontró el registro base CDCA."
    
    rcd.Edit
    ' Firmas Obligatorias
    rcd!firmaAprobacionRespIngenieriaNombre = Nz(objCDCA.firmaAprobacionRespIngenieriaNombre, "")
    rcd!firmaAprobacionRespCalidadNombre = Nz(objCDCA.firmaAprobacionRespCalidadNombre, "")
    
    ' Firmas Opcionales
    rcd!firmaAprobacionRespProduccionNombre = Nz(objCDCA.firmaAprobacionRespProduccionNombre, "")
    rcd!firmaAprobacionRespDisenioNombre = Nz(objCDCA.firmaAprobacionRespDisenioNombre, "")
    rcd!firmaAprobacionRepresentanteSumNombre = Nz(objCDCA.firmaAprobacionRepresentanteSumNombre, "")
    
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCARepositorio.ActualizarAprobacionSuministrador"
    errObj.Raise
End Sub

' En DatosCDCARepositorio.bas

Public Sub ActualizarDictamenRAC(ByRef objCDCA As DatosCDCA, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    ' Gestión de conexión dual
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosCDCA WHERE idSolicitud = " & objCDCA.idSolicitud & " ORDER BY idDatosCDCA"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If rcd.EOF Then Err.Raise 513, , "No se encontró el registro base CDCA."
    
    rcd.Edit
    rcd!racCodigo = Nz(objCDCA.racCodigo, "")
    rcd!observacionesRAC = Nz(objCDCA.observacionesRAC, "")
    rcd!racNombre = Nz(objCDCA.racNombre, "")
    rcd!racDecision = Nz(objCDCA.racDecision, "")
    
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCARepositorio.ActualizarDictamenRAC"
    errObj.Raise
End Sub

Public Sub ActualizarDecisionFinal(ByRef objCDCA As DatosCDCA, Optional db As DAO.Database)
    ' Soporta transacción opcional (db)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset, sql As String
    
    On Error GoTo Errores
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosCDCA WHERE idSolicitud = " & objCDCA.idSolicitud & " ORDER BY idDatosCDCA"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If rcd.EOF Then Err.Raise 513, , "No se encontró el registro base CDCA."
    
    rcd.Edit
    rcd!decisionFinal = Nz(objCDCA.decisionFinal, "")
    rcd!observacionesFinales = Nz(objCDCA.observacionesFinales, "")
    rcd!NombreFirmanteFinal = Nz(objCDCA.NombreFirmanteFinal, "")
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCARepositorio.ActualizarDecisionFinal"
    errObj.Raise
End Sub


Public Function getPorIdSolicitud(ByVal idSol As Long, Optional db As DAO.Database) As DatosCDCA
    Dim sql As String
    Dim params As Object
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    
    ' Gestión de conexión dual: Si viene transacción (db), la usamos.
    If db Is Nothing Then
        Set dbConexion = getdb()
    Else
        Set dbConexion = db
    End If
    
    sql = "SELECT * FROM tbDatosCDCA WHERE idSolicitud = [p_IDSol] ORDER BY idDatosCDCA;"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_IDSol", idSol
    
    Set getPorIdSolicitud = RepositorioComun.HidratarEntidadDesdeSQL(sql, "DatosCDCA", dbConexion, params)

LimpiarYSalir:
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCARepositorio.getPorIdSolicitud"
    errObj.AddToCallStack "idSolicitud: " & idSol
    errObj.Raise
End Function

Public Sub EliminarPorIdSolicitud(ByVal idSol As Long, ByRef db As DAO.Database)
    Dim sql As String
    Dim params As Object
    On Error GoTo Errores
    
    sql = "DELETE FROM tbDatosCDCA WHERE idSolicitud = [p_IDSol];"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_IDSol", idSol
    
    Call RepositorioComun.EjecutarAccion(sql, params, db)
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCARepositorio.EliminarPorIdSolicitud"
    errObj.AddToCallStack "idSolicitud: " & idSol
    errObj.Raise
End Sub
' ==========================================================================
' MÓDULO: DatosCDCARepositorio.bas
' ACCIÓN: Añadir el método genérico Guardar para compatibilidad con transacciones de rechazo.
' ==========================================================================
Public Sub Guardar(ByRef objCDCA As DatosCDCA, Optional db As DAO.Database)
    ' ESTE MÉTODO SE MANTIENE SOLO PARA SOPORTAR EL SERVICIO DE RECHAZO TRANSACCIONAL
    ' QUE USA LA ESTRATEGIA DE RELLENAR TODO EL OBJETO.
    ' NO DEBE USARSE DESDE EL FORMULARIO DE EDICIÓN NORMAL.
    
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosCDCA WHERE idSolicitud = " & objCDCA.idSolicitud & " ORDER BY idDatosCDCA"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If rcd.EOF Then
        rcd.AddNew
        If objCDCA.idDatosCDCA = 0 Then
            objCDCA.idDatosCDCA = getSiguienteIDDatosCDCA(dbConexion)
        End If
        Call RepositorioComun.RellenarRecordsetDesdeObjeto(rcd, objCDCA)
    Else
        rcd.Edit
        Call RepositorioComun.RellenarRecordsetDesdeObjeto(rcd, objCDCA, True) ' Omitir PK
    End If
    
    rcd.Update

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCARepositorio.Guardar (Compatibilidad)"
    errObj.Raise
End Sub
' ==========================================================================
' MÓDULO: DatosCDCARepositorio.bas
' ACCIÓN: Nuevo método para limpiar atómicamente la sección RAC
' ==========================================================================
Public Sub LimpiarDictamenRAC(ByVal idSolicitud As Long, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosCDCA WHERE idSolicitud = " & idSolicitud & " ORDER BY idDatosCDCA"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)

    If Not rcd.EOF Then
        rcd.Edit
        rcd!racCodigo = Null
        rcd!observacionesRAC = Null
        rcd!racNombre = Null
        
        ' CRÍTICO: Limpieza de los campos de decisión/rechazo
        rcd!racDecision = Null
        
        rcd.Update
    End If

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCARepositorio.LimpiarDictamenRAC"
    errObj.AddToCallStack "idSolicitud: " & idSolicitud
    errObj.Raise
End Sub

Public Sub LimpiarAprobacionSuministrador(ByVal idSolicitud As Long, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosCDCA WHERE idSolicitud = " & idSolicitud & " ORDER BY idDatosCDCA"
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
    errObj.Create Err.Number, Err.description, "DatosCDCARepositorio.LimpiarAprobacionSuministrador"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud: errObj.Raise
End Sub

Public Sub LimpiarDecisionFinal(ByVal idSolicitud As Long, Optional db As DAO.Database)
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    
    sql = "SELECT * FROM tbDatosCDCA WHERE idSolicitud = " & idSolicitud & " ORDER BY idDatosCDCA"
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If Not rcd.EOF Then
        rcd.Edit
        rcd!decisionFinal = Null
        rcd!observacionesFinales = Null
        rcd!NombreFirmanteFinal = Null
        rcd.Update
    End If

LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCARepositorio.LimpiarDecisionFinal"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud: errObj.Raise
End Sub


