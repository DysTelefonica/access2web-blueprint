VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatosCDCASUBServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Public Function getDatosCDCASUB(ByVal idSolicitud As Long, Optional db As DAO.Database) As DatosCDCASUB
    On Error GoTo Errores
    Set getDatosCDCASUB = DatosCDCASUBRepositorio.getPorIdSolicitud(idSolicitud, db)
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCASUBServicio.getDatosCDCASUB"
    errObj.Raise
End Function

Public Function GetDatosGeneralesIniciales(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As DatosCDCASUB
    Dim cdca As New DatosCDCASUB
    Dim solServ As New SolicitudServicio, expServ As New ExpedienteServicio
    Dim sol As Solicitud, exp As Expediente
    Dim sumServ As New SuministradorServicio
    Dim colRelaciones As Scripting.Dictionary
    Dim colNombresPrincipales As New Collection
    Dim colNombresSub As New Collection
    Dim key As Variant
    Dim relacion As ExpedienteSuministrador
    Dim unSuministrador As suministrador
    Dim nombrePrincipal As String
    Dim nombreSub As String
    Dim firstPrincipalDir As String
    Dim firstSubDir As String
    Dim direccionUTE As String
    
    On Error GoTo Errores
    
    Set sol = solServ.getSolicitudPorID(idSolicitud)
    Set exp = expServ.getExpedientePorID(CStr(sol.idExpediente))
    
    Set colRelaciones = expServ.getSuministradoresPorExpediente(exp.idExpediente, db)
    If Not colRelaciones Is Nothing Then
        For Each key In colRelaciones.Keys
            Set relacion = colRelaciones(key)
            Set unSuministrador = sumServ.getSuministradorPorID(CStr(relacion.idSuministrador), db)
            If Not unSuministrador Is Nothing Then
                If relacion.ContratistaPrincipal = "Sí" Then
                    colNombresPrincipales.Add UCase(Trim(unSuministrador.nombre))
                    If colNombresPrincipales.count = 1 Then firstPrincipalDir = Nz(unSuministrador.Direccion, "")
                ElseIf relacion.SubContratista = "Sí" Then
                    colNombresSub.Add UCase(Trim(unSuministrador.nombre))
                    If colNombresSub.count = 1 Then firstSubDir = Nz(unSuministrador.Direccion, "")
                End If
            End If
        Next key
    End If
    
    If colNombresPrincipales.count > 1 Then
        nombrePrincipal = "UNION TEMPORAL DE EMPRESAS: " & JoinCollection(colNombresPrincipales, " y ")
    ElseIf colNombresPrincipales.count = 1 Then
        nombrePrincipal = colNombresPrincipales(1)
    End If
    
    If colNombresSub.count > 1 Then
        nombreSub = "UNION TEMPORAL DE EMPRESAS: " & JoinCollection(colNombresSub, " y ")
    ElseIf colNombresSub.count = 1 Then
        nombreSub = colNombresSub(1)
    End If
    
    If nombreSub = "" Then
        nombreSub = nombrePrincipal
    End If
    
    If nombrePrincipal = "" Then
        nombrePrincipal = nombreSub
    End If
    
    direccionUTE = Trim(Nz(m_ObjEntorno.DirUTE, ""))
    If exp.ContratistaPrincipal = "Sí" Then
        If direccionUTE = "" Then direccionUTE = Trim(Nz(firstPrincipalDir, ""))
        If Nz(nombrePrincipal, "") <> "" And direccionUTE <> "" Then
            nombrePrincipal = nombrePrincipal & " (" & direccionUTE & ")"
        End If
    Else
        If direccionUTE = "" Then direccionUTE = Trim(Nz(firstSubDir, ""))
        If Nz(nombreSub, "") <> "" And direccionUTE <> "" Then
            nombreSub = nombreSub & " (" & direccionUTE & ")"
        End If
    End If
    
    With cdca
        .idSolicitud = idSolicitud
        .refSubSuministrador = sol.codigoSolicitud
        
        .suministradorPrincipalNombreDir = nombrePrincipal
        .subSuministradorNombreDir = nombreSub
    End With
    
    Set GetDatosGeneralesIniciales = cdca
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCASUBServicio.GetDatosGeneralesIniciales"
    errObj.Raise
End Function

' ... (Helpers y otros métodos se mantienen igual, solo actualizamos los que tocan los campos renombrados) ...

Public Function GetDatosAprobacionIniciales(ByVal idSolicitud As Long) As DatosCDCASUB
    Dim cdca As New DatosCDCASUB
    Dim solServ As New SolicitudServicio, expServ As New ExpedienteServicio
    Dim sol As Solicitud, exp As Expediente
    On Error GoTo Errores
    Set sol = solServ.getSolicitudPorID(idSolicitud)
    Set exp = expServ.getExpedientePorID(CStr(sol.idExpediente))
    With cdca
        .firmaAprobacionRespIngenieriaNombre = AgregarSufijoRol(exp.ResponsableTecnico, "Jefe de Proyecto")
        .firmaAprobacionRespCalidadNombre = AgregarSufijoRol(m_ObjEntorno.NombreJefeCalidad, "Responsable de Calidad")
    End With
    Set GetDatosAprobacionIniciales = cdca
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCASUBServicio.GetDatosAprobacionIniciales"
    errObj.Raise
End Function

Public Function GuardarDatosGenerales(obj As DatosCDCASUB, Optional ByRef db As DAO.Database) As Boolean
    Dim wfServ As New WorkflowServicio, solServ As New SolicitudServicio
    Dim ws As DAO.Workspace, dbConexion As DAO.Database, sol As Solicitud
    Dim transicionRealizada As Boolean
    Dim gestionaTransaccionLocal As Boolean
    
    On Error GoTo Errores
    transicionRealizada = False
    
    ' ~~ELIMINADO: refSuministrador ya no es obligatorio en CDCASUB~~
    If Trim(Nz(obj.refSubSuministrador, "")) = "" Then Err.Raise 513, "Validación", "Ref. Sub-Suministrador obligatoria."
    If Trim(Nz(obj.suministradorPrincipalNombreDir, "")) = "" Then Err.Raise 513, "Validación", "Suministrador Principal obligatorio."
    If Trim(Nz(obj.subSuministradorNombreDir, "")) = "" Then Err.Raise 513, "Validación", "Sub-Suministrador obligatorio."
    If IsNull(obj.requiereModificacionContrato) Then Err.Raise 513, "Validación", "Indique si requiere modificación."

    ' Dual-resolution: usar db entregada o crear una local
    If db Is Nothing Then
        gestionaTransaccionLocal = True
        Set ws = getWorkspace()
        Set dbConexion = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
    Else
        gestionaTransaccionLocal = False
        Set dbConexion = db
    End If
    
    Call DatosCDCASUBRepositorio.Guardar(obj, dbConexion)
    
    Set sol = solServ.getSolicitudPorID(obj.idSolicitud, dbConexion)
    If sol.idEstadoInterno = estadoPreregistro Then
        Call wfServ.EjecutarTransicion(sol, estadoRegistro, m_ObjUsuarioActivo, dbConexion)
        transicionRealizada = True
    End If
    
    If gestionaTransaccionLocal Then ws.CommitTrans
    GuardarDatosGenerales = transicionRealizada
    
LimpiarYSalir:
    On Error Resume Next
    If gestionaTransaccionLocal And Not dbConexion Is Nothing Then dbConexion.Close
    Set ws = Nothing
    Exit Function
Errores:
    Dim savedErr As Long, savedDesc As String
    savedErr = Err.Number: savedDesc = Err.description
    If gestionaTransaccionLocal And Not dbConexion Is Nothing Then dbConexion.Close
    If Not ws Is Nothing Then ws.Rollback
    Dim errObj As New CondorError
    errObj.Create savedErr, savedDesc, "DatosCDCASUBServicio.GuardarDatosGenerales"
    errObj.Raise
End Function
Public Sub GuardarPropuesta(obj As DatosCDCASUB, Optional db As DAO.Database)
    On Error GoTo Errores
    
    Dim ws As DAO.Workspace
    Dim dbTrabajo As DAO.Database
    Dim gestionaTransaccionLocal As Boolean
    Dim requiereTransaccion As Boolean
    If Trim(Nz(obj.identificacionMaterial, "")) = "" Then Err.Raise 513, "Validacion", "Identificacion obligatoria."
    If Trim(Nz(obj.causaNC, "")) = "" Then Err.Raise 513, "Validacion", "Causa NC obligatoria."
    If Trim(Nz(obj.descripcionImpactoNC, "")) = "" Then Err.Raise 513, "Validacion", "Descripcion Impacto obligatoria."
    ' --- TRAZABILIDAD RECHAZO (PREPARACION) ---
    ' Inicializar dbTrabajo primero para consistencia transaccional
    requiereTransaccion = (db Is Nothing And EsSolicitudEnValidacion(obj.idSolicitud))

    If requiereTransaccion Then
        Set ws = getWorkspace()
        Set dbTrabajo = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
        gestionaTransaccionLocal = True
    Else
        Set dbTrabajo = IIf(db Is Nothing, getdb(), db)
    End If

    Dim rechazoActivo As rechazo
    Dim objOriginal As DatosCDCASUB
    Set rechazoActivo = RechazoRepositorio.GetUltimoRechazoActivo(obj.idSolicitud, dbTrabajo)

    If Not rechazoActivo Is Nothing Then
        ' Leer estado original
        Set objOriginal = Me.getDatosCDCASUB(obj.idSolicitud, dbTrabajo)
    End If
    ' ------------------------------------------
    
    Call DatosCDCASUBRepositorio.ActualizarPropuesta(obj, dbTrabajo)

    ' --- TRAZABILIDAD RECHAZO (REGISTRO) ---
    If Not rechazoActivo Is Nothing And Not objOriginal Is Nothing Then
        Dim jsonDict As Object
        Set jsonDict = JsonHelper.ObtenerJsonRechazo(obj.idSolicitud)
        
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Propuesta", "identificacionMaterial", objOriginal.identificacionMaterial, obj.identificacionMaterial
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Propuesta", "numPlanoEspecificacion", objOriginal.numPlanoEspecificacion, obj.numPlanoEspecificacion
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Propuesta", "cantidadPeriodo", objOriginal.cantidadPeriodo, obj.cantidadPeriodo
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Propuesta", "numSerieLote", objOriginal.numSerieLote, obj.numSerieLote
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Propuesta", "causaNC", objOriginal.causaNC, obj.causaNC
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Propuesta", "descripcionImpactoNC", objOriginal.descripcionImpactoNC, obj.descripcionImpactoNC
        
        JsonHelper.GuardarJsonRechazo obj.idSolicitud, jsonDict
    End If
    ' ---------------------------------------
    
    If gestionaTransaccionLocal Then
        ws.CommitTrans
    End If
    
LimpiarYSalir:
    If gestionaTransaccionLocal And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Exit Sub
Errores:
    Dim savedErr As Long, savedDesc As String
    savedErr = Err.Number: savedDesc = Err.description
    If gestionaTransaccionLocal And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    If gestionaTransaccionLocal And Not (ws Is Nothing) Then ws.Rollback
    Set ws = Nothing
    Dim errObj As New CondorError: errObj.Create savedErr, savedDesc, "DatosCDCASUBServicio.GuardarPropuesta": errObj.Raise
End Sub

Public Sub GuardarImpacto(obj As DatosCDCASUB, Optional db As DAO.Database)
' PROPOSITO: Validar y orquestar el guardado del bloque "Impacto".
' REGLAS: En Impacto todos son obligatorios.
' -----------------------------------------------------------------------
    On Error GoTo Errores
    
    Dim ws As DAO.Workspace
    Dim dbTrabajo As DAO.Database
    Dim gestionaTransaccionLocal As Boolean
    Dim requiereTransaccion As Boolean
    Dim coste As String: coste = UCase(Trim(Nz(obj.impactoCoste, "")))
    Dim clasif As String: clasif = UCase(Trim(Nz(obj.clasificacionNC, "")))
    If coste = "" Or clasif = "" Or IsNull(obj.esSubSuministradorAD) Then Err.Raise 513, "Validacion", "Complete campos obligatorios de Impacto."
    ' --- TRAZABILIDAD RECHAZO (PREPARACION) ---
    ' Inicializar dbTrabajo primero para consistencia transaccional
    requiereTransaccion = (db Is Nothing And EsSolicitudEnValidacion(obj.idSolicitud))

    If requiereTransaccion Then
        Set ws = getWorkspace()
        Set dbTrabajo = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
        gestionaTransaccionLocal = True
    Else
        Set dbTrabajo = IIf(db Is Nothing, getdb(), db)
    End If

    Dim rechazoActivo As rechazo
    Dim objOriginal As DatosCDCASUB
    Set rechazoActivo = RechazoRepositorio.GetUltimoRechazoActivo(obj.idSolicitud, dbTrabajo)

    If Not rechazoActivo Is Nothing Then
        ' Leer estado original
        Set objOriginal = Me.getDatosCDCASUB(obj.idSolicitud, dbTrabajo)
    End If
    ' ------------------------------------------
    
    Call DatosCDCASUBRepositorio.ActualizarImpacto(obj, dbTrabajo)

    ' --- TRAZABILIDAD RECHAZO (REGISTRO) ---
    If Not rechazoActivo Is Nothing And Not objOriginal Is Nothing Then
        Dim jsonDict As Object
        Set jsonDict = JsonHelper.ObtenerJsonRechazo(obj.idSolicitud)
        
        ' Caracteristicas Afectadas (Booleanos)
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Impacto", "afectaPrestaciones", objOriginal.afectaPrestaciones, obj.afectaPrestaciones
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Impacto", "afectaSeguridad", objOriginal.afectaSeguridad, obj.afectaSeguridad
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Impacto", "afectaFiabilidad", objOriginal.afectaFiabilidad, obj.afectaFiabilidad
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Impacto", "afectaVidaUtil", objOriginal.afectaVidaUtil, obj.afectaVidaUtil
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Impacto", "afectaMedioambiente", objOriginal.afectaMedioambiente, obj.afectaMedioambiente
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Impacto", "afectaMantenibilidad", objOriginal.afectaMantenibilidad, obj.afectaMantenibilidad
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Impacto", "afectaIntercambiabilidad", objOriginal.afectaIntercambiabilidad, obj.afectaIntercambiabilidad
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Impacto", "afectaApariencia", objOriginal.afectaApariencia, obj.afectaApariencia
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Impacto", "afectaOtros", objOriginal.afectaOtros, obj.afectaOtros
        
        ' Campos de Texto y Radios
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Impacto", "impactoCoste", objOriginal.impactoCoste, obj.impactoCoste
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Impacto", "clasificacionNC", objOriginal.clasificacionNC, obj.clasificacionNC
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Impacto", "esSubSuministradorAD", objOriginal.esSubSuministradorAD, obj.esSubSuministradorAD
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Impacto", "identificacionAutoridadDiseno", objOriginal.identificacionAutoridadDiseno, obj.identificacionAutoridadDiseno
        JsonHelper.RegistrarCambio jsonDict, "CDCASUB_Impacto", "efectoFechaEntrega", objOriginal.efectoFechaEntrega, obj.efectoFechaEntrega
        
        JsonHelper.GuardarJsonRechazo obj.idSolicitud, jsonDict
    End If
    ' ---------------------------------------
    
    If gestionaTransaccionLocal Then
        ws.CommitTrans
    End If
    
LimpiarYSalir:
    If gestionaTransaccionLocal And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Exit Sub
Errores:
    Dim savedErr As Long, savedDesc As String
    savedErr = Err.Number: savedDesc = Err.description
    If gestionaTransaccionLocal And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    If gestionaTransaccionLocal And Not (ws Is Nothing) Then ws.Rollback
    Set ws = Nothing
    Dim errObj As New CondorError: errObj.Create savedErr, savedDesc, "DatosCDCASUBServicio.GuardarImpacto": errObj.Raise
End Sub

Public Sub GuardarAprobacionSuministrador(obj As DatosCDCASUB, Optional ByRef db As DAO.Database)
    On Error GoTo Errores
    If db Is Nothing Then Set db = getdb()
    If Trim(Nz(obj.firmaAprobacionRespIngenieriaNombre, "")) = "" Then Err.Raise 513, "Validación", "Firma Ingeniería obligatoria."
    If Trim(Nz(obj.firmaAprobacionRespCalidadNombre, "")) = "" Then Err.Raise 513, "Validación", "Firma Calidad obligatoria."
    Call DatosCDCASUBRepositorio.ActualizarAprobacionSuministrador(obj, db)
    Exit Sub
Errores:
    Dim errObj As New CondorError: errObj.Create Err.Number, Err.description, "DatosCDCASUBServicio.GuardarAprobacionSuministrador": errObj.Raise
End Sub

Public Sub GuardarDictamenRAC(obj As DatosCDCASUB, Optional ByRef db As DAO.Database)
    On Error GoTo Errores

    Dim ws As DAO.Workspace
    Dim dbTrabajo As DAO.Database
    Dim gestionaTransaccionLocal As Boolean
    If obj.racDecision <> "RECHAZADO" Then
        If Trim(Nz(obj.racCodigo, "")) = "" Then Err.Raise 513, "Validación", "Código RAC obligatorio."
    End If

    ' Slice 2 fix (#18, SDD fix-cdcASUB-servicio-transactional-isolation): separate
    ' the connection-source decision (caller vs service) from the transaction-policy
    ' decision (wrap in BeginTrans or not). Previously the SAME conditional
    ' conflated both, causing `dbTrabajo = Nothing` for solicitudes NOT in
    ' validation when the caller passed `db = Nothing`.
    If db Is Nothing Then
        ' Caller wants the service to manage everything. Always open a
        ' connection; wrap in a transaction only when the solicitud is in
        ' validation (a single UPDATE outside validation auto-commits).
        Set ws = getWorkspace()
        Set dbTrabajo = ws.OpenDatabase(getdb().name, False, False, ";PWD=" & GetPasswordDB())
        If EsSolicitudEnValidacion(obj.idSolicitud) Then
            ws.BeginTrans
            gestionaTransaccionLocal = True
        End If
    Else
        Set dbTrabajo = db
    End If

    Call DatosCDCASUBRepositorio.ActualizarDictamenRAC(obj, dbTrabajo)

    If gestionaTransaccionLocal Then
        ws.CommitTrans
    End If

LimpiarYSalir:
    If gestionaTransaccionLocal And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Exit Sub
Errores:
    Dim savedErr As Long, savedDesc As String
    savedErr = Err.Number: savedDesc = Err.description
    If gestionaTransaccionLocal And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    If gestionaTransaccionLocal And Not (ws Is Nothing) Then ws.Rollback
    Set ws = Nothing
    Dim errObj As New CondorError: errObj.Create savedErr, savedDesc, "DatosCDCASUBServicio.GuardarDictamenRAC": errObj.Raise
End Sub

Public Sub GuardarDecisionFinal(ByRef obj As DatosCDCASUB, Optional ByRef db As DAO.Database)
    On Error GoTo Errores
    ' ACTUALIZADO CON FUENTE DE VERDAD
    If Trim(Nz(obj.decisionFinal, "")) = "" Then Err.Raise 513, "Validación", "Decisión Final obligatoria."
    Call DatosCDCASUBRepositorio.ActualizarDecisionFinal(obj, db)
    Exit Sub
Errores:
    Dim savedErr As Long, savedDesc As String
    savedErr = Err.Number: savedDesc = Err.description
    Dim errObj As New CondorError
    errObj.Create savedErr, savedDesc, "DatosCDCASUBServicio.GuardarDecisionFinal"
    errObj.Raise
End Sub

Public Function EsDatosGeneralesCompleta(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objCDCASUB As DatosCDCASUB
    On Error GoTo Errores
    
    EsDatosGeneralesCompleta = False
    ' PASAMOS DB
    Set objCDCASUB = Me.getDatosCDCASUB(idSolicitud, db)
    If objCDCASUB Is Nothing Then Exit Function
    
    
    
    If Trim(Nz(objCDCASUB.refSubSuministrador, "")) = "" Then Exit Function
    If Trim(Nz(objCDCASUB.suministradorPrincipalNombreDir, "")) = "" Then Exit Function
    If Trim(Nz(objCDCASUB.subSuministradorNombreDir, "")) = "" Then Exit Function
    If Trim(Nz(objCDCASUB.requiereModificacionContrato, "")) = "" Then Exit Function
    
    EsDatosGeneralesCompleta = True
    
LimpiarYSalir:
    Exit Function
Errores:
    EsDatosGeneralesCompleta = False
    Resume LimpiarYSalir
End Function
Public Function EsParteTecnicaCompleta(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objCDCASUB As DatosCDCASUB
    On Error GoTo Errores
    
    EsParteTecnicaCompleta = False
    ' PASAMOS DB
    Set objCDCASUB = Me.getDatosCDCASUB(idSolicitud, db)
    If objCDCASUB Is Nothing Then Exit Function

    ' Propuesta
    If Trim(Nz(objCDCASUB.identificacionMaterial, "")) = "" Then Exit Function
    If Trim(Nz(objCDCASUB.causaNC, "")) = "" Then Exit Function
    If Trim(Nz(objCDCASUB.descripcionImpactoNC, "")) = "" Then Exit Function

    ' Impacto
    Dim coste As String, clasificacion As String
    coste = UCase(Trim(Nz(objCDCASUB.impactoCoste, "")))
    clasificacion = UCase(Trim(Nz(objCDCASUB.clasificacionNC, "")))
    
    If clasificacion <> "MAYOR" And clasificacion <> "MENOR" Then Exit Function
    If IsNull(objCDCASUB.esSubSuministradorAD) Then Exit Function
    
    EsParteTecnicaCompleta = True
    
LimpiarYSalir:
    Exit Function
Errores:
    EsParteTecnicaCompleta = False
    Resume LimpiarYSalir
End Function

Public Function EsDetalleCompleto(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objCDCASUB As DatosCDCASUB
    On Error GoTo Errores
    
    EsDetalleCompleto = False
    Set objCDCASUB = Me.getDatosCDCASUB(idSolicitud, db)
    If objCDCASUB Is Nothing Then Exit Function
    
    If Trim(Nz(objCDCASUB.identificacionMaterial, "")) = "" Then Exit Function
    If Trim(Nz(objCDCASUB.causaNC, "")) = "" Then Exit Function
    If Trim(Nz(objCDCASUB.descripcionImpactoNC, "")) = "" Then Exit Function
    
    EsDetalleCompleto = True
    Exit Function

Errores:
    EsDetalleCompleto = False
    
End Function

Public Function EsMotivosCompleto(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objCDCASUB As DatosCDCASUB

    On Error GoTo Errores

    EsMotivosCompleto = False
    Set objCDCASUB = Me.getDatosCDCASUB(idSolicitud, db)
    If objCDCASUB Is Nothing Then Exit Function

    If Trim(Nz(objCDCASUB.descripcionImpactoNC, "")) = "" Then Exit Function
    If Trim(Nz(objCDCASUB.descripcionImpactoNCCont, "")) = "" Then Exit Function

    EsMotivosCompleto = True
    Exit Function

Errores:
    EsMotivosCompleto = False
    
End Function



Public Function EsDictamenRACCompleto(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objCDCASUB As DatosCDCASUB
    On Error GoTo Errores
    
    EsDictamenRACCompleto = False
    
    Set objCDCASUB = Me.getDatosCDCASUB(idSolicitud, db)
    
    If objCDCASUB Is Nothing Then Exit Function

    If Trim(Nz(objCDCASUB.racCodigo, "")) = "" Then Exit Function
    If Trim(Nz(objCDCASUB.racNombre, "")) = "" Then Exit Function
    If Trim(Nz(objCDCASUB.racDecision, "")) = "" Then Exit Function
    
    EsDictamenRACCompleto = True
LimpiarYSalir:
    Exit Function
Errores:
    EsDictamenRACCompleto = False
    Resume LimpiarYSalir
End Function

Public Function EsAprobacionSuministradorCompleta(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objCDCASUB As DatosCDCASUB
    On Error GoTo Errores
    
    EsAprobacionSuministradorCompleta = False
   
    Set objCDCASUB = Me.getDatosCDCASUB(idSolicitud, db)
    If objCDCASUB Is Nothing Then Exit Function
    
    If Trim(Nz(objCDCASUB.firmaAprobacionRespIngenieriaNombre, "")) = "" Then Exit Function
    If Trim(Nz(objCDCASUB.firmaAprobacionRespCalidadNombre, "")) = "" Then Exit Function
    If Trim(Nz(objCDCASUB.decisionFinal, "")) = "" Then Exit Function
    
    EsAprobacionSuministradorCompleta = True
LimpiarYSalir:
    Exit Function
Errores:
    EsAprobacionSuministradorCompleta = False
    Resume LimpiarYSalir
End Function
Public Function EsDecisionFinalCompleta(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objCDCASUB As DatosCDCASUB
    On Error GoTo Errores
    
    EsDecisionFinalCompleta = False
    
    
    Set objCDCASUB = Me.getDatosCDCASUB(idSolicitud, db)
    
    If objCDCASUB Is Nothing Then Exit Function

    If Not DecisionFinalHelper_EsCompleta(objCDCASUB.decisionFinal, objCDCASUB.NombreFirmanteFinal) Then Exit Function
    
    EsDecisionFinalCompleta = True
    
LimpiarYSalir:
    Exit Function
Errores:
    EsDecisionFinalCompleta = False
    Resume LimpiarYSalir
End Function


Public Sub EliminarDictamenRAC(idSol As Long, Optional db As DAO.Database)
    Call DatosCDCASUBRepositorio.LimpiarDictamenRAC(idSol, db)
End Sub

Public Sub EliminarAprobacionSuministrador(ByVal idSolicitud As Long, Optional db As DAO.Database)
    On Error GoTo Errores
    Call DatosCDCASUBRepositorio.LimpiarAprobacionSuministrador(idSolicitud, db)
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCASUBServicio.EliminarAprobacionSuministrador"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud: errObj.Raise
End Sub

Public Sub EliminarDecisionFinal(idSol As Long, Optional db As DAO.Database)
    Call DatosCDCASUBRepositorio.LimpiarDecisionFinal(idSol, db)
End Sub

Public Sub ActualizarCamposDependientesDeExpediente(ByRef sol As Solicitud, ByRef exp As Expediente, Optional db As DAO.Database)
    Dim data As DatosCDCASUB
    Dim dictNombres As Scripting.Dictionary
    
    On Error GoTo Errores
    Set data = Me.getDatosCDCASUB(sol.idSolicitud, db)
    If data Is Nothing Then Set data = New DatosCDCASUB: data.idSolicitud = sol.idSolicitud
    
    data.refSubSuministrador = sol.codigoSolicitud
    Dim expServ As New ExpedienteServicio
    Set dictNombres = expServ.getNombresSuministradores(exp.idExpediente)
    
    If exp.ContratistaPrincipal = "Sí" Then
        data.suministradorPrincipalNombreDir = dictNombres("principal") & " (" & m_ObjEntorno.DirUTE & ")"
        data.subSuministradorNombreDir = dictNombres("subcontratista")
    Else
        data.suministradorPrincipalNombreDir = dictNombres("principal")
        data.subSuministradorNombreDir = dictNombres("subcontratista") & " (" & m_ObjEntorno.DirUTE & ")"
    End If
    
    Call DatosCDCASUBRepositorio.Guardar(data, db)
    Exit Sub
Errores:
    Dim errObj As New CondorError: errObj.Create Err.Number, Err.description, "DatosCDCASUBServicio.ActualizarCampos": errObj.Raise
End Sub

Public Sub PurgaTecnica(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database)
    ' Responsabilidad: Tabla Rasa de la parte técnica (Propuesta e Impacto)
    Dim data As DatosCDCASUB
    On Error GoTo Errores
    Set data = Me.getDatosCDCASUB(idSolicitud, db)
    If data Is Nothing Then Exit Sub
    
    With data
        ' Limpiar Propuesta
        .identificacionMaterial = ""
        .numPlanoEspecificacion = ""
        .cantidadPeriodo = ""
        .numSerieLote = ""
        .causaNC = ""
        .descripcionImpactoNC = ""
        .descripcionImpactoNCCont = ""
        
        ' Limpiar Impacto
        .afectaPrestaciones = False
        .afectaSeguridad = False
        .afectaFiabilidad = False
        .afectaVidaUtil = False
        .afectaMedioambiente = False
        .afectaIntercambiabilidad = False
        .afectaApariencia = False
        .afectaOtros = False
        .impactoCoste = ""
        .clasificacionNC = ""
        .efectoFechaEntrega = ""
    End With
    
    Call DatosCDCASUBRepositorio.Guardar(data, db)
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCASUBServicio.PurgaTecnica"
    errObj.Raise
End Sub

Public Sub RegistrarRechazoDesdeFormalizacion(ByVal idSolicitud As Long, ByVal motivoRechazo As String, ByVal usuarioActivo As usuario, Optional ByRef db As DAO.Database)
    Dim ws As DAO.Workspace
    Dim dbConexion As DAO.Database
    Dim wfServ As New WorkflowServicio
    Dim subData As DatosCDCASUB
    
    On Error GoTo Errores
    
    ' Gestión de conexión transaccional
    If db Is Nothing Then
        Set ws = getWorkspace()
        Set dbConexion = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
    Else
        Set dbConexion = db
    End If
    
    ' 2. Obtener datos
    Set subData = Me.getDatosCDCASUB(idSolicitud, dbConexion)
    If subData Is Nothing Then Err.Raise 513, , "Datos no encontrados."

    ' 3. Limpieza de datos (En el objeto)
    ' A. Borrar datos de Formalización (Decisión Final)
    subData.decisionFinal = ""
    subData.observacionesFinales = ""
    subData.NombreFirmanteFinal = ""
    
    ' B. Borrar datos de Aprobación RAC
    subData.racCodigo = ""
    subData.racNombre = ""
    subData.observacionesRAC = ""
    subData.observacionesRACDelegador = ""
    
    ' C. Establecer el Rechazo
    subData.racDecision = "RECHAZADO"
    subData.racRechazoMotivos = motivoRechazo
    
    ' 4. Guardar Datos (Usando la conexión DB de la transacción)
    Call DatosCDCASUBRepositorio.Guardar(subData, dbConexion)
    
    ' 5. Transición de Estado (Usando la conexión DB de la transacción)
    Call wfServ.EjecutarTransicion(SolicitudRepositorio.getSolicitudPorID(idSolicitud, dbConexion), estadoRechazada, usuarioActivo, dbConexion)
    
    ' 6. Commit (solo si abrimos nosotros)
    If db Is Nothing Then ws.CommitTrans

LimpiarYSalir:
    On Error Resume Next
    If db Is Nothing And Not dbConexion Is Nothing Then dbConexion.Close
    If db Is Nothing Then Set ws = Nothing
    Exit Sub
    
Errores:
    If db Is Nothing And Not ws Is Nothing Then ws.Rollback
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCASUBServicio.RegistrarRechazoDesdeFormalizacion"
    errObj.Raise
End Sub

Private Function EsSolicitudEnValidacion(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As Boolean
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String
    
    On Error GoTo Errores
    
    If Not db Is Nothing Then
        Set dbConexion = db
    Else
        Set dbConexion = getdb()
    End If
    
    sql = "SELECT idEstadoInterno FROM tbSolicitudes WHERE idSolicitud = " & idSolicitud
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    
    If rcd.EOF Then
        rcd.Close
        Exit Function
    End If
    
    EsSolicitudEnValidacion = (rcd!idEstadoInterno = estadoValidacion)
    
    rcd.Close
    
    Exit Function
Errores:
    EsSolicitudEnValidacion = False
End Function

Private Function AgregarSufijoRol(ByVal nombre As String, ByVal rol As String) As String
    Dim resultado As String
    Dim sufijo As String
    
    resultado = Nz(nombre, "")
    sufijo = " (" & rol & ")"
    
    If resultado <> "" And InStr(resultado, sufijo) = 0 Then
        resultado = resultado & sufijo
    End If
    
    AgregarSufijoRol = resultado
End Function

' ==========================================================================
' MÉTODOS DE MAPEO VM (Extraídos del form para lógica de negocio reutilizable)
' ==========================================================================

Public Sub ImportarEntidad_aVM(ByRef ent As DatosCDCASUB, ByRef vm As DatosCDCASUBViewModel)
    On Error GoTo Errores

    If ent Is Nothing Then Exit Sub

    ' 1. Sincronizar el objeto de datos interno del VM
    Set vm.datos = ent

    ' 2. Sección Generales
    vm.vm_refSuministrador = ent.refSuministrador
    vm.vm_refSubSuministrador = ent.refSubSuministrador
    vm.vm_suministradorPrincipalNombreDir = ent.suministradorPrincipalNombreDir
    vm.vm_subSuministradorNombreDir = ent.subSuministradorNombreDir
    vm.vm_refDesviacionesPrevias = ent.refDesviacionesPrevias
    vm.vm_requiereModificacionContrato = ent.requiereModificacionContrato

    ' 3. Sección Propuesta
    vm.vm_identificacionMaterial = ent.identificacionMaterial
    vm.vm_numPlanoEspecificacion = ent.numPlanoEspecificacion
    vm.vm_cantidadPeriodo = ent.cantidadPeriodo
    vm.vm_numSerieLote = ent.numSerieLote
    vm.vm_causaNC = ent.causaNC
    vm.vm_descripcionImpactoNC = ent.descripcionImpactoNC

    ' 4. Sección Impacto
    vm.vm_afectaPrestaciones = ent.afectaPrestaciones
    vm.vm_afectaSeguridad = ent.afectaSeguridad
    vm.vm_afectaFiabilidad = ent.afectaFiabilidad
    vm.vm_afectaVidaUtil = ent.afectaVidaUtil
    vm.vm_afectaMedioambiente = ent.afectaMedioambiente
    vm.vm_afectaMantenibilidad = ent.afectaMantenibilidad
    vm.vm_afectaIntercambiabilidad = ent.afectaIntercambiabilidad
    vm.vm_afectaApariencia = ent.afectaApariencia
    vm.vm_afectaOtros = ent.afectaOtros
    vm.vm_impactoCoste = ent.impactoCoste
    vm.vm_clasificacionNC = ent.clasificacionNC
    vm.vm_efectoFechaEntrega = ent.efectoFechaEntrega
    vm.vm_identificacionAutoridadDiseno = ent.identificacionAutoridadDiseno
    vm.vm_esSubSuministradorAD = ent.esSubSuministradorAD

    ' 5. Sección Aprobación Suministrador
    vm.vm_firmaAprobacionRespIngenieriaNombre = ent.firmaAprobacionRespIngenieriaNombre
    vm.vm_firmaAprobacionRespCalidadNombre = ent.firmaAprobacionRespCalidadNombre
    vm.vm_firmaAprobacionRespProduccionNombre = ent.firmaAprobacionRespProduccionNombre
    vm.vm_firmaAprobacionRespDisenioNombre = ent.firmaAprobacionRespDisenioNombre
    vm.vm_firmaAprobacionRepresentanteSumNombre = ent.firmaAprobacionRepresentanteSumNombre

    ' 6. Sección RAC (Campos específicos CDCASUB)
    vm.vm_racCodigo = ent.racCodigo
    vm.vm_racDecision = ent.racDecision
    vm.vm_observacionesRAC = ent.observacionesRAC
    vm.vm_racNombre = ent.racNombre
    vm.vm_observacionesRACDelegador = ent.observacionesRACDelegador
    vm.vm_racNombreDelegador = ent.racNombreDelegador

    ' 7. Sección Decisión Final
    vm.vm_decisionFinal = ent.decisionFinal
    vm.vm_obsDecisionFinal = ent.observacionesFinales
    vm.vm_NombreFirmanteFinal = ent.NombreFirmanteFinal

    Exit Sub

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCASUBServicio.ImportarEntidad_aVM"
    errObj.Raise
End Sub

Public Sub ExportarVM_aEntidad(ByVal bloque As enumBloqueFormulario, ByRef vm As DatosCDCASUBViewModel, ByRef ent As DatosCDCASUB)
    On Error GoTo Errores

    ' Sincronizamos los IDs siempre
    ent.idSolicitud = vm.Solicitud.idSolicitud
    ent.idDatosCDCASUB = vm.datos.idDatosCDCASUB

    Select Case bloque
        Case Bloque_Generales
            ent.refSuministrador = CStr(Nz(vm.vm_refSuministrador, ""))
            ent.refSubSuministrador = CStr(Nz(vm.vm_refSubSuministrador, ""))
            ent.suministradorPrincipalNombreDir = CStr(Nz(vm.vm_suministradorPrincipalNombreDir, ""))
            ent.subSuministradorNombreDir = CStr(Nz(vm.vm_subSuministradorNombreDir, ""))
            ent.refDesviacionesPrevias = CStr(Nz(vm.vm_refDesviacionesPrevias, ""))
            Dim reqMod As Variant
            reqMod = vm.vm_requiereModificacionContrato
            If IsNull(reqMod) Then
                ent.requiereModificacionContrato = Null
            ElseIf VarType(reqMod) = vbBoolean Then
                ent.requiereModificacionContrato = reqMod
            ElseIf IsNumeric(reqMod) Then
                ent.requiereModificacionContrato = (CLng(reqMod) <> 0)
            Else
                ent.requiereModificacionContrato = Null
            End If

        Case Bloque_Propuesta
            ent.identificacionMaterial = vm.vm_identificacionMaterial
            ent.numPlanoEspecificacion = vm.vm_numPlanoEspecificacion
            ent.cantidadPeriodo = vm.vm_cantidadPeriodo
            ent.numSerieLote = vm.vm_numSerieLote
            ent.causaNC = vm.vm_causaNC
            ent.descripcionImpactoNC = vm.vm_descripcionImpactoNC

        Case Bloque_Impacto
            ent.afectaPrestaciones = vm.vm_afectaPrestaciones
            ent.afectaSeguridad = vm.vm_afectaSeguridad
            ent.afectaFiabilidad = vm.vm_afectaFiabilidad
            ent.afectaVidaUtil = vm.vm_afectaVidaUtil
            ent.afectaMedioambiente = vm.vm_afectaMedioambiente
            ent.afectaMantenibilidad = vm.vm_afectaMantenibilidad
            ent.afectaIntercambiabilidad = vm.vm_afectaIntercambiabilidad
            ent.afectaApariencia = vm.vm_afectaApariencia
            ent.afectaOtros = vm.vm_afectaOtros
            ent.impactoCoste = vm.vm_impactoCoste
            ent.clasificacionNC = vm.vm_clasificacionNC
            ent.efectoFechaEntrega = vm.vm_efectoFechaEntrega
            ent.identificacionAutoridadDiseno = vm.vm_identificacionAutoridadDiseno
            ent.esSubSuministradorAD = vm.vm_esSubSuministradorAD

        Case Bloque_AprobacionSuministrador
            ent.firmaAprobacionRespIngenieriaNombre = vm.vm_firmaAprobacionRespIngenieriaNombre
            ent.firmaAprobacionRespCalidadNombre = vm.vm_firmaAprobacionRespCalidadNombre
            ent.firmaAprobacionRespProduccionNombre = vm.vm_firmaAprobacionRespProduccionNombre
            ent.firmaAprobacionRespDisenioNombre = vm.vm_firmaAprobacionRespDisenioNombre
            ent.firmaAprobacionRepresentanteSumNombre = vm.vm_firmaAprobacionRepresentanteSumNombre

        Case Bloque_DictamenRAC
            ent.racCodigo = vm.vm_racCodigo
            ent.racDecision = vm.vm_racDecision
            ent.observacionesRAC = vm.vm_observacionesRAC
            ent.racNombre = vm.vm_racNombre
            ent.observacionesRACDelegador = vm.vm_observacionesRACDelegador
            ent.racNombreDelegador = vm.vm_racNombreDelegador

        Case Bloque_DecisionFinal
            ent.decisionFinal = vm.vm_decisionFinal
            ent.observacionesFinales = vm.vm_obsDecisionFinal
            ent.NombreFirmanteFinal = vm.vm_NombreFirmanteFinal

    End Select

    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCASUBServicio.ExportarVM_aEntidad"
    errObj.Raise
End Sub

Public Function ObtenerViewModelCompleto(ByVal idSolicitud As Long, Optional db As DAO.Database) As DatosCDCASUBViewModel
    Dim vm As New DatosCDCASUBViewModel
    Dim sol As Solicitud
    Dim exp As Expediente
    Dim ent As DatosCDCASUB
    Dim expServ As New ExpedienteServicio
    Dim solServ As New SolicitudServicio
    Dim colHistorial As Scripting.Dictionary
    Dim ultimaFecha As Date
    
    On Error GoTo Errores
    
    ' Cargar Solicitud
    Set sol = solServ.getSolicitudPorID(idSolicitud, db)
    If sol Is Nothing Then Err.Raise 513, "ObtenerViewModelCompleto", "Solicitud no encontrada."
    
    ' Cargar Expediente
    Set exp = expServ.getExpedientePorID(CStr(sol.idExpediente))
    If exp Is Nothing Then Err.Raise 513, "ObtenerViewModelCompleto", "Expediente no encontrado."
    
    ' Obtener o crear entidad en memoria
    Set ent = Me.getDatosCDCASUB(idSolicitud, db)
    If ent Is Nothing Then
        Set ent = Me.GetDatosGeneralesIniciales(idSolicitud)
    End If
    
    ' Asignar al VM
    Set vm.Solicitud = sol
    Set vm.Expediente = exp
    Set vm.datos = ent
    
    ' Estado
    If m_ObjEntorno.estados.Exists(CStr(sol.idEstadoInterno)) Then
        vm.EstadoNombre = m_ObjEntorno.estados(CStr(sol.idEstadoInterno)).nombreEstado
    Else
        vm.EstadoNombre = "Estado " & sol.idEstadoInterno
    End If
    
    ' Tiempos
    Set colHistorial = solServ.getHistorialDeEstados(idSolicitud)
    ultimaFecha = sol.fechaCreacion
    If colHistorial.count > 0 Then ultimaFecha = colHistorial.items(colHistorial.count - 1).fechaTransicion
    
    If sol.idEstadoInterno = estadoAprobada Or sol.idEstadoInterno = estadoRechazada Then
         vm.TiempoCiclo = CalcularDuracion(sol.fechaCreacion, ultimaFecha)
    Else
         vm.TiempoCiclo = CalcularDuracion(sol.fechaCreacion, Now()) & " (en curso)"
    End If
    
    ' Mapear entidad a VM
    Call ImportarEntidad_aVM(ent, vm)
    
    ' Datos de revisión Calidad
    vm.vm_revisionCalidadEstado = sol.revisionCalidadEstado
    vm.vm_revisionCalidadComentarios = sol.revisionCalidadComentarios
    
    ' Permisos
    Dim wfServ As New WorkflowServicio
    vm.PuedeGuardar_Generales = wfServ.PuedeEditarBloque(idSolicitud, Bloque_Generales)
    vm.PuedeGuardar_Propuesta = wfServ.PuedeEditarBloque(idSolicitud, Bloque_Propuesta)
    vm.PuedeGuardar_Impacto = wfServ.PuedeEditarBloque(idSolicitud, Bloque_Impacto)
    vm.PuedeGuardar_AprobacionSuministrador = wfServ.PuedeEditarBloque(idSolicitud, Bloque_AprobacionSuministrador)
    vm.PuedeGuardar_DictamenRAC = wfServ.PuedeEditarBloque(idSolicitud, Bloque_DictamenRAC)
    vm.PuedeGuardar_DecisionFinal = wfServ.PuedeEditarBloque(idSolicitud, Bloque_DecisionFinal)
    
    Set ObtenerViewModelCompleto = vm
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCASUBServicio.ObtenerViewModelCompleto"
    errObj.AddToCallStack "idSolicitud: " & idSolicitud
    errObj.Raise
End Function



