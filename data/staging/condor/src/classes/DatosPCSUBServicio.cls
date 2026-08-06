VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatosPCSUBServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: DatosPCSUBServicio.cls
' SERVICIO: Solicitudes PC-Subcontratista (PCSUB)
' FUENTE DE VERDAD: tbDatosPCSUB en condor_datos.accdb (ERD 2026-05-01)
' ==========================================================================

' === getDatosPCSUB ===
Public Function getDatosPCSUB(ByVal idSolicitud As Long, Optional db As DAO.Database) As DatosPCSUB
    On Error GoTo Errores
    Set getDatosPCSUB = DatosPCSUBRepositorio.getPorIdSolicitud(idSolicitud, db)
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBServicio.getDatosPCSUB"
    errObj.Raise
End Function

Public Sub PurgaTecnica(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database)
    ' Responsabilidad: Tabla rasa de la parte técnica (Propuesta e Impacto) para PC_SUB.
    Dim pcsub As DatosPCSUB

    On Error GoTo Errores

    Set pcsub = Me.getDatosPCSUB(idSolicitud, db)
    If pcsub Is Nothing Then Exit Sub

    With pcsub
        ' Limpiar Propuesta
        .descripcionPropuestaCambio = ""
        .descripcionPropuestaCambioCont = ""
        .numPlanoEspecificacion = ""

        ' Limpiar Motivos
        .motivoCorregirDeficiencias = False
        .motivoMejorarCapacidad = False
        .motivoAumentarNacionalizacion = False
        .motivoMejorarSeguridad = False
        .motivoMejorarFiabilidad = False
        .motivoMejorarCosteEficacia = False
        .motivoOtros = False
        .motivoOtrosDetalle = ""

        ' Limpiar Impacto / Incidencias
        .descripcionMaterialAfectado = ""
        .incidenciaCoste = ""
        .incidenciaPlazo = ""
        .incidenciaSeguridad = False
        .incidenciaFiabilidad = False
        .incidenciaMantenibilidad = False
        .incidenciaIntercambiabilidad = False
        .incidenciaVidaUtilAlmacen = False
        .incidenciaFuncionamientoFuncion = False
        .impactoClasificacion = ""
        .CambioAfectaAMaterial = ""
    End With

    Call DatosPCSUBRepositorio.Guardar(pcsub, db)

LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBServicio.PurgaTecnica"
    errObj.Raise
End Sub

' === GetDatosGeneralesIniciales ===
Public Function GetDatosGeneralesIniciales(ByVal idSolicitud As Long) As DatosPCSUB
    Dim pcsub As New DatosPCSUB
    Dim solServ As New SolicitudServicio
    Dim expServ As New ExpedienteServicio
    Dim sumServ As New SuministradorServicio
    Dim sol As Solicitud
    Dim exp As Expediente
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

    ' Datos immutables del expediente
    pcsub.refContratoInspeccionOficial = exp.CodExpLargo
    pcsub.denominacionContrato = exp.titulo
    pcsub.objetoContrato = exp.objetoContrato

    ' Referencia subcontratista
    pcsub.refSubSuministrador = sol.codigoSolicitud

    ' Cargar nombres de suministradores
    Set colRelaciones = expServ.getSuministradoresPorExpediente(exp.idExpediente)
    If Not colRelaciones Is Nothing Then
        For Each key In colRelaciones.Keys
            Set relacion = colRelaciones(key)
            Set unSuministrador = sumServ.getSuministradorPorID(CStr(relacion.idSuministrador))
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

    If nombreSub = "" Then nombreSub = nombrePrincipal
    If nombrePrincipal = "" Then nombrePrincipal = nombreSub

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

    With pcsub
        .idSolicitud = idSolicitud
        .SubsuministradorNombreDir = nombreSub
    End With

    Set GetDatosGeneralesIniciales = pcsub
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBServicio.GetDatosGeneralesIniciales"
    errObj.Raise
End Function

' === GetDatosAprobacionIniciales ===
Public Function GetDatosAprobacionIniciales(ByVal idSolicitud As Long) As DatosPCSUB
    Dim pcsub As New DatosPCSUB
    Dim solServ As New SolicitudServicio
    Dim expServ As New ExpedienteServicio
    Dim sol As Solicitud
    Dim exp As Expediente
    On Error GoTo Errores
    Set sol = solServ.getSolicitudPorID(idSolicitud)
    Set exp = expServ.getExpedientePorID(CStr(sol.idExpediente))
    With pcsub
        .firmaOficinaTecnicaSubSuministradorNombre = AgregarSufijoRol(exp.ResponsableTecnico, "Jefe de Proyecto")
        .firmaRepSubSuministradorNombre = AgregarSufijoRol(m_ObjEntorno.NombreJefeCalidad, "Responsable de Calidad")
    End With
    Set GetDatosAprobacionIniciales = pcsub
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBServicio.GetDatosAprobacionIniciales"
    errObj.Raise
End Function

' === GuardarDatosGenerales ===
Public Function GuardarDatosGenerales(obj As DatosPCSUB, Optional db As DAO.Database) As Boolean
    Dim wfServ As New WorkflowServicio
    Dim solServ As New SolicitudServicio
    Dim ws As DAO.Workspace
    Dim dbTrabajo As DAO.Database
    Dim sol As Solicitud
    Dim transicionRealizada As Boolean
    Dim gestionaTransaccion As Boolean
    Dim requiereTransaccion As Boolean

    On Error GoTo Errores
    transicionRealizada = False
    gestionaTransaccion = False

    If Trim(Nz(obj.refContratoInspeccionOficial, "")) = "" Then Err.Raise 513, "Validación", "El campo 'Referencia del Contrato/Inspección Oficial' es obligatorio."
    If Trim(Nz(obj.refSubSuministrador, "")) = "" Then Err.Raise 513, "Validación", "Ref. Sub-Suministrador obligatoria."
    If Trim(Nz(obj.SubsuministradorNombreDir, "")) = "" Then Err.Raise 513, "Validación", "Sub-Suministrador (Nombre y Dir) obligatorio."
    If Trim(Nz(obj.denominacionContrato, "")) = "" Then Err.Raise 513, "Validación", "El campo 'Denominación del Contrato' es obligatorio."
    If Trim(Nz(obj.objetoContrato, "")) = "" Then Err.Raise 513, "Validación", "Objeto del contrato obligatorio."

    requiereTransaccion = (db Is Nothing And EsSolicitudEnValidacion(obj.idSolicitud))
    If requiereTransaccion Then
        Set ws = getWorkspace()
        Set dbTrabajo = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
        gestionaTransaccion = True
    Else
        Set dbTrabajo = db
    End If

    Call DatosPCSUBRepositorio.Guardar(obj, dbTrabajo)

    Set sol = solServ.getSolicitudPorID(obj.idSolicitud, dbTrabajo)
    If sol.idEstadoInterno = estadoPreregistro Then
        Call wfServ.EjecutarTransicion(sol, estadoRegistro, m_ObjUsuarioActivo, dbTrabajo)
        transicionRealizada = True
    End If

    If requiereTransaccion Then ws.CommitTrans
    GuardarDatosGenerales = transicionRealizada
    GoTo LimpiarYSalir

LimpiarYSalir:
    On Error Resume Next
    If gestionaTransaccion And Not dbTrabajo Is Nothing Then dbTrabajo.Close
    Set ws = Nothing
    Exit Function
Errores:
    Dim savedErr As Long, savedDesc As String
    savedErr = Err.Number: savedDesc = Err.description
    If requiereTransaccion And Not ws Is Nothing Then ws.Rollback
    Dim errObj As New CondorError
    errObj.Create savedErr, savedDesc, "DatosPCSUBServicio.GuardarDatosGenerales"
    errObj.Raise
End Function

' === GuardarPropuesta ===
Public Sub GuardarPropuesta(obj As DatosPCSUB, Optional db As DAO.Database)
    On Error GoTo Errores

    Dim ws As DAO.Workspace
    Dim dbTrabajo As DAO.Database
    Dim gestionaTransaccion As Boolean
    Dim requiereTransaccion As Boolean

    If Trim(Nz(obj.descripcionMaterialAfectado, "")) = "" Then Err.Raise 513, "Regla de Negocio Incumplida", "La descripción del material afectado es obligatoria."
    If Trim(Nz(obj.descripcionPropuestaCambio, "")) = "" Then Err.Raise 513, "Regla de Negocio Incumplida", "La descripción de la propuesta de cambio es obligatoria."

    requiereTransaccion = (db Is Nothing And EsSolicitudEnValidacion(obj.idSolicitud))
    If requiereTransaccion Then
        Set ws = getWorkspace()
        Set dbTrabajo = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
        gestionaTransaccion = True
    Else
        Set dbTrabajo = db
    End If

    Dim rechazo As rechazo
    Dim objOriginal As DatosPCSUB
    Set rechazo = RechazoRepositorio.GetUltimoRechazoActivo(obj.idSolicitud, dbTrabajo)
    If Not rechazo Is Nothing Then Set objOriginal = Me.getDatosPCSUB(obj.idSolicitud, dbTrabajo)

    Call DatosPCSUBRepositorio.ActualizarPropuesta(obj, dbTrabajo)

    If Not rechazo Is Nothing And Not objOriginal Is Nothing Then
        Dim jsonDict As Object
        Set jsonDict = JsonHelper.ObtenerJsonRechazo(obj.idSolicitud)
        JsonHelper.RegistrarCambio jsonDict, "PCSUB_Propuesta", "descripcionPropuestaCambio", objOriginal.descripcionPropuestaCambio, obj.descripcionPropuestaCambio
        JsonHelper.RegistrarCambio jsonDict, "PCSUB_Propuesta", "numPlanoEspecificacion", objOriginal.numPlanoEspecificacion, obj.numPlanoEspecificacion
        JsonHelper.RegistrarCambio jsonDict, "PCSUB_Propuesta", "descripcionMaterialAfectado", objOriginal.descripcionMaterialAfectado, obj.descripcionMaterialAfectado
        JsonHelper.GuardarJsonRechazo obj.idSolicitud, jsonDict
    End If

    If gestionaTransaccion Then ws.CommitTrans

LimpiarYSalir:
    If gestionaTransaccion And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Exit Sub
Errores:
    If gestionaTransaccion And Not (ws Is Nothing) Then ws.Rollback
    If gestionaTransaccion And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBServicio.GuardarPropuesta"
    errObj.Raise
End Sub

' === GuardarImpacto ===
Public Sub GuardarImpacto(obj As DatosPCSUB, Optional db As DAO.Database)
    On Error GoTo Errores

    Dim ws As DAO.Workspace
    Dim dbTrabajo As DAO.Database
    Dim gestionaTransaccion As Boolean
    Dim requiereTransaccion As Boolean

    ' 1) Motivos
    Dim algunMotivo As Boolean
    algunMotivo = (obj.motivoCorregirDeficiencias Or obj.motivoMejorarCapacidad Or _
                   obj.motivoAumentarNacionalizacion Or obj.motivoMejorarSeguridad Or _
                   obj.motivoMejorarFiabilidad Or obj.motivoMejorarCosteEficacia Or _
                   obj.motivoOtros)
    If Not algunMotivo Then Err.Raise 513, "Regla de Negocio Incumplida", ">> (Impacto) Seleccione al menos un Motivo."
    If obj.motivoOtros And Trim(Nz(obj.motivoOtrosDetalle, "")) = "" Then Err.Raise 513, "Regla de Negocio Incumplida", ">> (Impacto) Debe indicar el detalle cuando marca 'Otros'."
    If obj.motivoOtros = False And Trim(Nz(obj.motivoOtrosDetalle, "")) <> "" Then Err.Raise 513, "Regla de Negocio Incumplida", ">> (Impacto) Si se indica otros motivos se ha de marcar 'Otros'."

    ' 2) Incidencias
    Dim v As String
    v = UCase$(Trim$(Nz(obj.incidenciaCoste, "")))
    If (v <> "AUMENTARÁ" And v <> "DISMINUIRÁ" And v <> "NO VARIARÁ") Then Err.Raise 513, "Regla de Negocio Incumplida", ">> (Impacto) Incidencia sobre el Coste es obligatoria."
    v = UCase$(Trim$(Nz(obj.incidenciaPlazo, "")))
    If (v <> "AUMENTARÁ" And v <> "DISMINUIRÁ" And v <> "NO VARIARÁ") Then Err.Raise 513, "Regla de Negocio Incumplida", ">> (Impacto) Incidencia sobre el Plazo es obligatoria."

    ' 3) Clasificación
    v = UCase$(Trim$(Nz(obj.impactoClasificacion, "")))
    If (v <> "MAYOR" And v <> "MENOR") Then Err.Raise 513, "Regla de Negocio Incumplida", ">> (Impacto) Clasificación (MAYOR/MENOR) es obligatoria."

    v = Trim$(Nz(obj.CambioAfectaAMaterial, ""))
    If (v <> "Material ya entregado" And v <> "Material por entregar") Then Err.Raise 513, "Regla de Negocio Incumplida", ">> (Impacto) Debe indicar si el cambio afecta a material."

    requiereTransaccion = (db Is Nothing And EsSolicitudEnValidacion(obj.idSolicitud))
    If requiereTransaccion Then
        Set ws = getWorkspace()
        Set dbTrabajo = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
        gestionaTransaccion = True
    Else
        Set dbTrabajo = db
    End If

    Dim rechazoActivo As rechazo
    Dim objOriginal As DatosPCSUB
    Set rechazoActivo = RechazoRepositorio.GetUltimoRechazoActivo(obj.idSolicitud, dbTrabajo)
    If Not rechazoActivo Is Nothing Then Set objOriginal = Me.getDatosPCSUB(obj.idSolicitud, dbTrabajo)

    Call DatosPCSUBRepositorio.ActualizarImpacto(obj, dbTrabajo)

    If Not rechazoActivo Is Nothing And Not objOriginal Is Nothing Then
        Dim jsonDict As Object
        Set jsonDict = JsonHelper.ObtenerJsonRechazo(obj.idSolicitud)
        JsonHelper.RegistrarCambio jsonDict, "PCSUB_Impacto", "motivoCorregirDeficiencias", objOriginal.motivoCorregirDeficiencias, obj.motivoCorregirDeficiencias
        JsonHelper.RegistrarCambio jsonDict, "PCSUB_Impacto", "motivoMejorarCapacidad", objOriginal.motivoMejorarCapacidad, obj.motivoMejorarCapacidad
        JsonHelper.RegistrarCambio jsonDict, "PCSUB_Impacto", "motivoAumentarNacionalizacion", objOriginal.motivoAumentarNacionalizacion, obj.motivoAumentarNacionalizacion
        JsonHelper.RegistrarCambio jsonDict, "PCSUB_Impacto", "motivoMejorarSeguridad", objOriginal.motivoMejorarSeguridad, obj.motivoMejorarSeguridad
        JsonHelper.RegistrarCambio jsonDict, "PCSUB_Impacto", "motivoMejorarFiabilidad", objOriginal.motivoMejorarFiabilidad, obj.motivoMejorarFiabilidad
        JsonHelper.RegistrarCambio jsonDict, "PCSUB_Impacto", "motivoMejorarCosteEficacia", objOriginal.motivoMejorarCosteEficacia, obj.motivoMejorarCosteEficacia
        JsonHelper.RegistrarCambio jsonDict, "PCSUB_Impacto", "motivoOtros", objOriginal.motivoOtros, obj.motivoOtros
        JsonHelper.RegistrarCambio jsonDict, "PCSUB_Impacto", "motivoOtrosDetalle", objOriginal.motivoOtrosDetalle, obj.motivoOtrosDetalle
        JsonHelper.RegistrarCambio jsonDict, "PCSUB_Impacto", "incidenciaCoste", objOriginal.incidenciaCoste, obj.incidenciaCoste
        JsonHelper.RegistrarCambio jsonDict, "PCSUB_Impacto", "incidenciaPlazo", objOriginal.incidenciaPlazo, obj.incidenciaPlazo
        JsonHelper.RegistrarCambio jsonDict, "PCSUB_Impacto", "impactoClasificacion", objOriginal.impactoClasificacion, obj.impactoClasificacion
        JsonHelper.RegistrarCambio jsonDict, "PCSUB_Impacto", "CambioAfectaAMaterial", objOriginal.CambioAfectaAMaterial, obj.CambioAfectaAMaterial
        JsonHelper.GuardarJsonRechazo obj.idSolicitud, jsonDict
    End If

    If gestionaTransaccion Then ws.CommitTrans

LimpiarYSalir:
    If gestionaTransaccion And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Exit Sub
Errores:
    If gestionaTransaccion And Not (ws Is Nothing) Then ws.Rollback
    If gestionaTransaccion And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBServicio.GuardarImpacto"
    errObj.Raise
End Sub

' === GuardarAprobacionSuministrador ===
Public Sub GuardarAprobacionSuministrador(obj As DatosPCSUB, Optional db As DAO.Database)
    On Error GoTo Errores
    If Trim(Nz(obj.firmaOficinaTecnicaSubSuministradorNombre, "")) = "" Then Err.Raise 513, "Regla Incumplida", ">> (Aprobación Sub) Nombre en firma de Oficina Técnica."
    If Trim(Nz(obj.firmaRepSubSuministradorNombre, "")) = "" Then Err.Raise 513, "Regla de Negocio Incumplida", ">> (Aprobación Sub) Nombre en firma del Representante del Sub."
    If db Is Nothing Then Set db = getdb()
    Call DatosPCSUBRepositorio.ActualizarAprobacionSuministrador(obj, db)
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBServicio.GuardarAprobacionSuministrador"
    errObj.Raise
End Sub

' === GuardarDictamenRAC ===
Public Sub GuardarDictamenRAC(obj As DatosPCSUB, Optional db As DAO.Database)
    On Error GoTo Errores

    Dim ws As DAO.Workspace
    Dim dbTrabajo As DAO.Database
    Dim gestionaTransaccion As Boolean
    Dim requiereTransaccion As Boolean

    ' Si NO es RECHAZADO, código obligatorio
    If UCase(Trim(Nz(obj.racDecision, ""))) <> "RECHAZADO" Then
        If Trim(Nz(obj.racCodigo, "")) = "" Then Err.Raise 513, "Validación", "Código RAC obligatorio."
    End If

    requiereTransaccion = (db Is Nothing And EsSolicitudEnValidacion(obj.idSolicitud))
    If requiereTransaccion Then
        Set ws = getWorkspace()
        Set dbTrabajo = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
        gestionaTransaccion = True
    Else
        Set dbTrabajo = db
    End If

    Call DatosPCSUBRepositorio.ActualizarDictamenRAC(obj, dbTrabajo)

    If gestionaTransaccion Then ws.CommitTrans

LimpiarYSalir:
    If gestionaTransaccion And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Exit Sub
Errores:
    If gestionaTransaccion And Not (ws Is Nothing) Then ws.Rollback
    If gestionaTransaccion And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBServicio.GuardarDictamenRAC"
    errObj.Raise
End Sub

' === GuardarDecisionFinal ===
Public Sub GuardarDecisionFinal(obj As DatosPCSUB, Optional db As DAO.Database)
    On Error GoTo Errores
    If Trim(Nz(obj.decisionFinal, "")) = "" Then Err.Raise 513, "Validación", "Decisión Final obligatoria."
    Call DatosPCSUBRepositorio.ActualizarDecisionFinal(obj, db)
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBServicio.GuardarDecisionFinal"
    errObj.Raise
End Sub

' === EsDatosGeneralesCompleta ===
Public Function EsDatosGeneralesCompleta(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objPCSUB As DatosPCSUB
    On Error GoTo Errores
    EsDatosGeneralesCompleta = False
    Set objPCSUB = Me.getDatosPCSUB(idSolicitud, db)
    If objPCSUB Is Nothing Then Exit Function
    If Trim(Nz(objPCSUB.refContratoInspeccionOficial, "")) = "" Then Exit Function
    If Trim(Nz(objPCSUB.refSubSuministrador, "")) = "" Then Exit Function
    If Trim(Nz(objPCSUB.SubsuministradorNombreDir, "")) = "" Then Exit Function
    If Trim(Nz(objPCSUB.denominacionContrato, "")) = "" Then Exit Function
    If Trim(Nz(objPCSUB.objetoContrato, "")) = "" Then Exit Function
    EsDatosGeneralesCompleta = True
LimpiarYSalir:
    Exit Function
Errores:
    EsDatosGeneralesCompleta = False
    Resume LimpiarYSalir
End Function

' === EsParteTecnicaCompleta (propuesta + impacto) ===
Public Function EsParteTecnicaCompleta(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objPCSUB As DatosPCSUB
    On Error GoTo Errores
    EsParteTecnicaCompleta = False
    Set objPCSUB = Me.getDatosPCSUB(idSolicitud, db)
    If objPCSUB Is Nothing Then Exit Function
    If Trim(Nz(objPCSUB.descripcionMaterialAfectado, "")) = "" Then Exit Function
    If Trim(Nz(objPCSUB.descripcionPropuestaCambio, "")) = "" Then Exit Function
    Dim coste As String, clasificacion As String
    coste = UCase(Trim(Nz(objPCSUB.incidenciaCoste, "")))
    clasificacion = UCase(Trim(Nz(objPCSUB.impactoClasificacion, "")))
    If coste <> "AUMENTARÁ" And coste <> "DISMINUIRÁ" And coste <> "NO VARIARÁ" Then Exit Function
    If clasificacion <> "MAYOR" And clasificacion <> "MENOR" Then Exit Function
    Dim algunMotivo As Boolean
    algunMotivo = (objPCSUB.motivoCorregirDeficiencias Or objPCSUB.motivoMejorarCapacidad Or _
                   objPCSUB.motivoAumentarNacionalizacion Or objPCSUB.motivoMejorarSeguridad Or _
                   objPCSUB.motivoMejorarFiabilidad Or objPCSUB.motivoMejorarCosteEficacia Or _
                   objPCSUB.motivoOtros)
    If Not algunMotivo Then Exit Function
    EsParteTecnicaCompleta = True
LimpiarYSalir:
    Exit Function
Errores:
    EsParteTecnicaCompleta = False
    Resume LimpiarYSalir
End Function

' === EsDetalleCompleto (SPEC-132) ===
Public Function EsDetalleCompleto(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim obj As DatosPCSUB
    EsDetalleCompleto = False
    Set obj = Me.getDatosPCSUB(idSolicitud, db)
    If obj Is Nothing Then Exit Function
    If Trim(Nz(obj.descripcionMaterialAfectado, "")) = "" Then Exit Function
    If Trim(Nz(obj.descripcionPropuestaCambio, "")) = "" Then Exit Function
    If Trim(Nz(obj.descripcionPropuestaCambioCont, "")) = "" Then Exit Function
    EsDetalleCompleto = True
End Function

' === EsMotivosCompleto (SPEC-132) ===
Public Function EsMotivosCompleto(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim obj As DatosPCSUB
    EsMotivosCompleto = False
    Set obj = Me.getDatosPCSUB(idSolicitud, db)
    If obj Is Nothing Then Exit Function
    Dim algunMotivo As Boolean
    algunMotivo = (obj.motivoCorregirDeficiencias Or obj.motivoMejorarCapacidad Or _
                   obj.motivoAumentarNacionalizacion Or obj.motivoMejorarSeguridad Or _
                   obj.motivoMejorarFiabilidad Or obj.motivoMejorarCosteEficacia Or _
                   obj.motivoOtros)
    If Not algunMotivo Then Exit Function
    EsMotivosCompleto = True
End Function

' === EsDictamenRACCompleto ===
Public Function EsDictamenRACCompleto(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objPCSUB As DatosPCSUB
    On Error GoTo Errores
    EsDictamenRACCompleto = False
    Set objPCSUB = Me.getDatosPCSUB(idSolicitud, db)
    If objPCSUB Is Nothing Then Exit Function
    If Trim(Nz(objPCSUB.racCodigo, "")) = "" Then Exit Function
    If Trim(Nz(objPCSUB.racNombre, "")) = "" Then Exit Function
    If Trim(Nz(objPCSUB.racDecision, "")) = "" Then Exit Function
    EsDictamenRACCompleto = True
LimpiarYSalir:
    Exit Function
Errores:
    EsDictamenRACCompleto = False
    Resume LimpiarYSalir
End Function

' === EsAprobacionSuministradorCompleta ===
Public Function EsAprobacionSuministradorCompleta(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objPCSUB As DatosPCSUB
    On Error GoTo Errores
    EsAprobacionSuministradorCompleta = False
    Set objPCSUB = Me.getDatosPCSUB(idSolicitud, db)
    If objPCSUB Is Nothing Then Exit Function
    If Trim(Nz(objPCSUB.firmaOficinaTecnicaSubSuministradorNombre, "")) = "" Then Exit Function
    If Trim(Nz(objPCSUB.firmaRepSubSuministradorNombre, "")) = "" Then Exit Function
    If Trim(Nz(objPCSUB.decisionFinal, "")) = "" Then Exit Function
    EsAprobacionSuministradorCompleta = True
LimpiarYSalir:
    Exit Function
Errores:
    EsAprobacionSuministradorCompleta = False
    Resume LimpiarYSalir
End Function

' === EsDecisionFinalCompleta ===
Public Function EsDecisionFinalCompleta(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objPCSUB As DatosPCSUB
    On Error GoTo Errores
    EsDecisionFinalCompleta = False
    Set objPCSUB = Me.getDatosPCSUB(idSolicitud, db)
    If objPCSUB Is Nothing Then Exit Function
    If Not DecisionFinalHelper_EsCompleta(objPCSUB.decisionFinal, objPCSUB.NombreFirmanteFinal) Then Exit Function
    EsDecisionFinalCompleta = True
LimpiarYSalir:
    Exit Function
Errores:
    EsDecisionFinalCompleta = False
    Resume LimpiarYSalir
End Function

' === ImportarEntidad_aVM ===
Public Sub ImportarEntidad_aVM(ByRef ent As DatosPCSUB, ByRef vm As DatosPCSUBViewModel)
    On Error GoTo Errores
    If ent Is Nothing Then Exit Sub
    Set vm.datos = ent

    ' Generales
    vm.vm_refContratoInspeccionOficial = Nz(ent.refContratoInspeccionOficial, "")
    vm.vm_refSubSuministrador = Nz(ent.refSubSuministrador, "")
    vm.vm_denominacionContrato = Nz(ent.denominacionContrato, "")
    vm.vm_objetoContrato = Nz(ent.objetoContrato, "")
    vm.vm_subSuministradorNombreDir = Nz(ent.SubsuministradorNombreDir, "")

    ' Propuesta
    vm.vm_numPlanoEspecificacion = Nz(ent.numPlanoEspecificacion, "")
    vm.vm_descripcionMaterialAfectado = Nz(ent.descripcionMaterialAfectado, "")
    vm.vm_descripcionPropuestaCambio = Nz(ent.descripcionPropuestaCambio, "")
    vm.vm_descripcionPropuestaCambioCont = Nz(ent.descripcionPropuestaCambioCont, "")

    ' Motivos
    vm.vm_motivoCorregirDeficiencias = ent.motivoCorregirDeficiencias
    vm.vm_motivoMejorarCapacidad = ent.motivoMejorarCapacidad
    vm.vm_motivoAumentarNacionalizacion = ent.motivoAumentarNacionalizacion
    vm.vm_motivoMejorarSeguridad = ent.motivoMejorarSeguridad
    vm.vm_motivoMejorarFiabilidad = ent.motivoMejorarFiabilidad
    vm.vm_motivoMejorarCosteEficacia = ent.motivoMejorarCosteEficacia
    vm.vm_motivoOtros = ent.motivoOtros
    vm.vm_motivoOtrosDetalle = Nz(ent.motivoOtrosDetalle, "")

    ' Incidencias
    vm.vm_incidenciaCoste = Nz(ent.incidenciaCoste, "")
    vm.vm_incidenciaPlazo = Nz(ent.incidenciaPlazo, "")
    vm.vm_incidenciaSeguridad = ent.incidenciaSeguridad
    vm.vm_incidenciaFiabilidad = ent.incidenciaFiabilidad
    vm.vm_incidenciaMantenibilidad = ent.incidenciaMantenibilidad
    vm.vm_incidenciaIntercambiabilidad = ent.incidenciaIntercambiabilidad
    vm.vm_incidenciaVidaUtilAlmacen = ent.incidenciaVidaUtilAlmacen
    vm.vm_incidenciaFuncionamientoFuncion = ent.incidenciaFuncionamientoFuncion
    vm.vm_impactoClasificacion = Nz(ent.impactoClasificacion, "")
    vm.vm_CambioAfectaAMaterial = Nz(ent.CambioAfectaAMaterial, "")

    ' Aprobación Sub
    vm.vm_firmaOficinaTecnicaSubSuministradorNombre = Nz(ent.firmaOficinaTecnicaSubSuministradorNombre, "")
    vm.vm_firmaRepSubSuministradorNombre = Nz(ent.firmaRepSubSuministradorNombre, "")

    ' Dictamen RAC
    vm.vm_racCodigo = Nz(ent.racCodigo, "")
    vm.vm_observacionesRAC = Nz(ent.observacionesRAC, "")
    vm.vm_racNombre = Nz(ent.racNombre, "")
    vm.vm_racDecision = Nz(ent.racDecision, "")
    vm.vm_racRechazoMotivos = Nz(ent.racRechazoMotivos, "")

    ' Delegador
    vm.vm_observacionesRACDelegador = Nz(ent.observacionesRACDelegador, "")
    vm.vm_racNombreDelegador = Nz(ent.racNombreDelegador, "")

    ' Decisión Final
    vm.vm_obsAprobacionAutoridadDiseno = Nz(ent.obsAprobacionAutoridadDiseno, "")
    vm.vm_NombreAutoridadDiseno = Nz(ent.NombreAutoridadDiseno, "")
    vm.vm_decisionFinal = Nz(ent.decisionFinal, "")
    vm.vm_obsDecisionFinal = Nz(ent.obsDecisionFinal, "")
    vm.vm_NombreFirmanteFinal = Nz(ent.NombreFirmanteFinal, "")
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBServicio.ImportarEntidad_aVM"
    errObj.Raise
End Sub

' === ExportarVM_aEntidad ===
Public Sub ExportarVM_aEntidad(ByVal bloque As enumBloqueFormulario, ByRef vm As DatosPCSUBViewModel, ByRef ent As DatosPCSUB)
    On Error GoTo Errores
    ent.idSolicitud = vm.Solicitud.idSolicitud
    ent.idDatosPCSUB = vm.datos.idDatosPCSUB

    Select Case bloque
        Case Bloque_Generales
            ent.refContratoInspeccionOficial = Nz(vm.vm_refContratoInspeccionOficial, "")
            ent.refSubSuministrador = Nz(vm.vm_refSubSuministrador, "")
            ent.denominacionContrato = Nz(vm.vm_denominacionContrato, "")
            ent.SubsuministradorNombreDir = Nz(vm.vm_subSuministradorNombreDir, "")
            ent.objetoContrato = Nz(vm.vm_objetoContrato, "")

        Case Bloque_Propuesta
            ent.numPlanoEspecificacion = Nz(vm.vm_numPlanoEspecificacion, "")
            ent.descripcionMaterialAfectado = Nz(vm.vm_descripcionMaterialAfectado, "")
            ent.descripcionPropuestaCambio = Nz(vm.vm_descripcionPropuestaCambio, "")
            ent.descripcionPropuestaCambioCont = Nz(vm.vm_descripcionPropuestaCambioCont, "")

        Case Bloque_Impacto
            ent.motivoCorregirDeficiencias = vm.vm_motivoCorregirDeficiencias
            ent.motivoMejorarCapacidad = vm.vm_motivoMejorarCapacidad
            ent.motivoAumentarNacionalizacion = vm.vm_motivoAumentarNacionalizacion
            ent.motivoMejorarSeguridad = vm.vm_motivoMejorarSeguridad
            ent.motivoMejorarFiabilidad = vm.vm_motivoMejorarFiabilidad
            ent.motivoMejorarCosteEficacia = vm.vm_motivoMejorarCosteEficacia
            ent.motivoOtros = vm.vm_motivoOtros
            ent.motivoOtrosDetalle = Nz(vm.vm_motivoOtrosDetalle, "")
            ent.incidenciaCoste = Nz(vm.vm_incidenciaCoste, "")
            ent.incidenciaPlazo = Nz(vm.vm_incidenciaPlazo, "")
            ent.incidenciaSeguridad = vm.vm_incidenciaSeguridad
            ent.incidenciaFiabilidad = vm.vm_incidenciaFiabilidad
            ent.incidenciaMantenibilidad = vm.vm_incidenciaMantenibilidad
            ent.incidenciaIntercambiabilidad = vm.vm_incidenciaIntercambiabilidad
            ent.incidenciaVidaUtilAlmacen = vm.vm_incidenciaVidaUtilAlmacen
            ent.incidenciaFuncionamientoFuncion = vm.vm_incidenciaFuncionamientoFuncion
            ent.impactoClasificacion = Nz(vm.vm_impactoClasificacion, "")
            ent.CambioAfectaAMaterial = Nz(vm.vm_CambioAfectaAMaterial, "")

        Case Bloque_AprobacionSuministrador
            ent.firmaOficinaTecnicaSubSuministradorNombre = Nz(vm.vm_firmaOficinaTecnicaSubSuministradorNombre, "")
            ent.firmaRepSubSuministradorNombre = Nz(vm.vm_firmaRepSubSuministradorNombre, "")

        Case Bloque_DictamenRAC
            ent.racCodigo = Nz(vm.vm_racCodigo, "")
            ent.observacionesRAC = Nz(vm.vm_observacionesRAC, "")
            ent.racNombre = Nz(vm.vm_racNombre, "")
            ent.racDecision = Nz(vm.vm_racDecision, "")
            ent.racRechazoMotivos = Nz(vm.vm_racRechazoMotivos, "")
            ent.observacionesRACDelegador = Nz(vm.vm_observacionesRACDelegador, "")
            ent.racNombreDelegador = Nz(vm.vm_racNombreDelegador, "")

        Case Bloque_DecisionFinal
            ent.obsAprobacionAutoridadDiseno = Nz(vm.vm_obsAprobacionAutoridadDiseno, "")
            ent.NombreAutoridadDiseno = Nz(vm.vm_NombreAutoridadDiseno, "")
            ent.decisionFinal = Nz(vm.vm_decisionFinal, "")
            ent.obsDecisionFinal = Nz(vm.vm_obsDecisionFinal, "")
            ent.NombreFirmanteFinal = Nz(vm.vm_NombreFirmanteFinal, "")
    End Select
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBServicio.ExportarVM_aEntidad"
    errObj.Raise
End Sub

' === ObtenerViewModelCompleto ===
Public Function ObtenerViewModelCompleto(ByVal idSolicitud As Long, Optional db As DAO.Database) As DatosPCSUBViewModel
    Dim vm As New DatosPCSUBViewModel
    Dim sol As Solicitud
    Dim exp As Expediente
    Dim ent As DatosPCSUB
    Dim expServ As New ExpedienteServicio
    Dim solServ As New SolicitudServicio
    Dim colHistorial As Scripting.Dictionary
    Dim ultimaFecha As Date
    Dim estadoActual As estado

    On Error GoTo Errores

    Set sol = solServ.getSolicitudPorID(idSolicitud, db)
    If sol Is Nothing Then Err.Raise 513, "ObtenerViewModelCompleto", "Solicitud no encontrada."
    Set exp = expServ.getExpedientePorID(CStr(sol.idExpediente))
    If exp Is Nothing Then Err.Raise 513, "ObtenerViewModelCompleto", "Expediente no encontrado."

    Set ent = Me.getDatosPCSUB(idSolicitud, db)
    If ent Is Nothing Then Set ent = Me.GetDatosGeneralesIniciales(idSolicitud)

    Set vm.Solicitud = sol
    Set vm.Expediente = exp
    Set vm.datos = ent

    If m_ObjEntorno.estados.Exists(CStr(sol.idEstadoInterno)) Then
        Set estadoActual = m_ObjEntorno.estados(CStr(sol.idEstadoInterno))
        vm.EstadoNombre = estadoActual.nombreEstado
        vm.EstadoDescripcion = estadoActual.descripcion
    Else
        vm.EstadoNombre = "Estado " & sol.idEstadoInterno
    End If

    Set colHistorial = solServ.getHistorialDeEstados(idSolicitud)
    ultimaFecha = sol.fechaCreacion
    If colHistorial.count > 0 Then ultimaFecha = colHistorial.items(colHistorial.count - 1).fechaTransicion

    If sol.idEstadoInterno = estadoAprobada Or sol.idEstadoInterno = estadoRechazada Then
         vm.TiempoCiclo = CalcularDuracion(sol.fechaCreacion, ultimaFecha)
    Else
         vm.TiempoCiclo = CalcularDuracion(sol.fechaCreacion, Now()) & " (en curso)"
    End If

    Call ImportarEntidad_aVM(ent, vm)

    vm.vm_revisionCalidadEstado = sol.revisionCalidadEstado
    vm.vm_revisionCalidadComentarios = sol.revisionCalidadComentarios

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
    errObj.Create Err.Number, Err.description, "DatosPCSUBServicio.ObtenerViewModelCompleto"
    errObj.AddToCallStack "idSolicitud: " & idSolicitud
    errObj.Raise
End Function

' === EliminarDictamenRAC ===
Public Sub EliminarDictamenRAC(idSol As Long, Optional db As DAO.Database)
    Call DatosPCSUBRepositorio.LimpiarDictamenRAC(idSol, db)
End Sub

' === EliminarAprobacionSuministrador ===
Public Sub EliminarAprobacionSuministrador(ByVal idSolicitud As Long, Optional db As DAO.Database)
    On Error GoTo Errores
    Call DatosPCSUBRepositorio.LimpiarAprobacionSuministrador(idSolicitud, db)
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBServicio.EliminarAprobacionSuministrador"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud: errObj.Raise
End Sub

' ==========================================================================
' ACCIÓN: Método para limpiar la Decisión Final (gemelo de CDCA/PC)
' ==========================================================================
Public Sub EliminarDecisionFinal(ByVal idSolicitud As Long, Optional db As DAO.Database)
    On Error GoTo Errores
    Call DatosPCSUBRepositorio.LimpiarDecisionFinal(idSolicitud, db)
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBServicio.EliminarDecisionFinal"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud: errObj.Raise
End Sub

' === RegistrarRechazoDesdeFormalizacion ===
Public Sub RegistrarRechazoDesdeFormalizacion(ByVal idSolicitud As Long, ByVal motivoRechazo As String, ByVal usuarioActivo As usuario, Optional ByRef db As DAO.Database)
    Dim ws As DAO.Workspace
    Dim dbTrabajo As DAO.Database
    Dim wfServ As New WorkflowServicio
    Dim pcsub As DatosPCSUB
    Dim gestionaTransaccion As Boolean
    Dim requiereTransaccion As Boolean

    On Error GoTo Errores

    requiereTransaccion = (db Is Nothing)
    If requiereTransaccion Then
        Set ws = getWorkspace()
        Set dbTrabajo = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
        gestionaTransaccion = True
    Else
        Set dbTrabajo = db
        gestionaTransaccion = False
    End If

    Set pcsub = Me.getDatosPCSUB(idSolicitud, dbTrabajo)
    If pcsub Is Nothing Then Err.Raise 513, , "Datos no encontrados."

    pcsub.decisionFinal = ""
    pcsub.obsDecisionFinal = ""
    pcsub.NombreFirmanteFinal = ""

    pcsub.racCodigo = ""
    pcsub.racNombre = ""
    pcsub.observacionesRAC = ""
    pcsub.racDecision = "RECHAZADO"
    pcsub.observacionesRACDelegador = ""
    pcsub.racNombreDelegador = ""

    Call DatosPCSUBRepositorio.Guardar(pcsub, dbTrabajo)
    Call wfServ.EjecutarTransicion(SolicitudRepositorio.getSolicitudPorID(idSolicitud, dbTrabajo), estadoRechazada, usuarioActivo, dbTrabajo)
    If requiereTransaccion Then ws.CommitTrans
    GoTo LimpiarYSalir

LimpiarYSalir:
    On Error Resume Next
    If gestionaTransaccion And Not dbTrabajo Is Nothing Then dbTrabajo.Close
    Set ws = Nothing
    Set wfServ = Nothing
    Set pcsub = Nothing
    Exit Sub
Errores:
    If requiereTransaccion And Not ws Is Nothing Then ws.Rollback
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBServicio.RegistrarRechazoDesdeFormalizacion"
    errObj.Raise
End Sub

' ==========================================================================
' HELPERS PRIVADOS
' ==========================================================================

Private Function EsSolicitudEnValidacion(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As Boolean
    Dim dbConexion As DAO.Database
    Dim rcd As DAO.Recordset
    Dim sql As String

    On Error GoTo Errores
    If Not db Is Nothing Then Set dbConexion = db Else Set dbConexion = getdb()
    sql = "SELECT idEstadoInterno FROM tbSolicitudes WHERE idSolicitud = " & idSolicitud
    Set rcd = dbConexion.OpenRecordset(sql, dbOpenDynaset)
    If rcd.EOF Then rcd.Close: Exit Function
    EsSolicitudEnValidacion = (rcd!idEstadoInterno = estadoValidacion)
    rcd.Close
    Exit Function
Errores:
    EsSolicitudEnValidacion = False
End Function

Public Sub ActualizarCamposDependientesDeExpediente(ByRef sol As Solicitud, ByRef exp As Expediente, Optional db As DAO.Database)
    Dim pcsub As DatosPCSUB
    Dim dictNombres As Scripting.Dictionary
    Dim expServ As New ExpedienteServicio
    Dim nombreSub As String
    Dim direccionUTE As String

    On Error GoTo Errores

    Set pcsub = Me.getDatosPCSUB(sol.idSolicitud, db)
    If pcsub Is Nothing Then
        Set pcsub = New DatosPCSUB
        pcsub.idSolicitud = sol.idSolicitud
    End If

    pcsub.refContratoInspeccionOficial = exp.CodExpLargo
    pcsub.refSubSuministrador = sol.codigoSolicitud
    pcsub.denominacionContrato = exp.titulo
    pcsub.objetoContrato = exp.objetoContrato

    Set dictNombres = expServ.getNombresSuministradores(exp.idExpediente, db)
    nombreSub = Nz(dictNombres("subcontratista"), "")
    If nombreSub = "" Then nombreSub = Nz(dictNombres("principal"), "")

    If exp.ContratistaPrincipal <> "Sí" Then
        direccionUTE = Trim(Nz(m_ObjEntorno.DirUTE, ""))
        If nombreSub <> "" And direccionUTE <> "" Then nombreSub = nombreSub & " (" & direccionUTE & ")"
    End If

    pcsub.SubsuministradorNombreDir = nombreSub

    Call DatosPCSUBRepositorio.Guardar(pcsub, db)
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCSUBServicio.ActualizarCamposDependientesDeExpediente"
    errObj.Raise
End Sub

Private Function AgregarSufijoRol(ByVal nombre As String, ByVal rol As String) As String
    Dim resultado As String
    Dim sufijo As String
    resultado = Nz(nombre, "")
    sufijo = " (" & rol & ")"
    If resultado <> "" And InStr(resultado, sufijo) = 0 Then resultado = resultado & sufijo
    AgregarSufijoRol = resultado
End Function





