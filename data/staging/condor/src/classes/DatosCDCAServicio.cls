VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatosCDCAServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Compare Database
Option Explicit


Public Sub ActualizarCamposDependientesDeExpediente(ByRef sol As Solicitud, ByRef exp As Expediente, Optional db As DAO.Database)
    Dim cdca As DatosCDCA
    
    On Error GoTo Errores
    
    ' 1. Obtener o crear
    Set cdca = Me.GetDatosCDCA(sol.idSolicitud, db)
    If cdca Is Nothing Then
        Set cdca = New DatosCDCA
        cdca.idSolicitud = sol.idSolicitud
    End If
    
    ' 2. Actualizar campos
    cdca.numContrato = exp.CodExpLargo
    cdca.refSuministrador = sol.codigoSolicitud
    cdca.firmaAprobacionRespIngenieriaNombre = exp.ResponsableTecnico
    
    cdca.SuministradorNombreDir = Me.GetNombreSuministradorDefault(exp.idExpediente)
    
    ' 3. Persistir (PASANDO LA TRANSACCIÓN db)
    Call DatosCDCARepositorio.Guardar(cdca, db)

LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCAServicio.ActualizarCamposDependientesDeExpediente"
    errObj.Raise
End Sub

Public Sub PurgaTecnica(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database)
    ' Responsabilidad: Tabla Rasa de la parte técnica (Propuesta e Impacto)
    Dim cdca As DatosCDCA
    On Error GoTo Errores
    Set cdca = Me.GetDatosCDCA(idSolicitud, db)
    If cdca Is Nothing Then Exit Sub
    
    With cdca
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
    
    Call DatosCDCARepositorio.Guardar(cdca, db)
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCAServicio.PurgaTecnica"
    errObj.Raise
End Sub

' ==========================================================================
' MÓDULO: DatosCDCAServicio.cls
' ACCIÓN: Propagar la transacción en las funciones de lectura y validación
' ==========================================================================

' 1. Actualizar getDatosCDCA
Public Function GetDatosCDCA(ByVal idSolicitud As Long, Optional db As DAO.Database) As DatosCDCA
    On Error GoTo Errores
    If idSolicitud <= 0 Then Err.Raise 513, "getDatosCDCA", "El ID de la Solicitud no es válido."
    
    ' Pasamos la conexión db al repositorio
    Set GetDatosCDCA = DatosCDCARepositorio.getPorIdSolicitud(idSolicitud, db)
    
LimpiarYSalir:
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCAServicio.getDatosCDCA"
    errObj.AddToCallStack "idSolicitud: " & idSolicitud
    errObj.Raise
End Function






' --- VALORES POR DEFECTO ---

Public Function GetDatosGeneralesIniciales(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As DatosCDCA
    Dim cdca As New DatosCDCA
    Dim solServ As New SolicitudServicio, expServ As New ExpedienteServicio
    Dim sumServ As New SuministradorServicio
    Dim sol As Solicitud, exp As Expediente
    Dim colRelaciones As Scripting.Dictionary
    Dim colNombresPrincipales As New Collection
    Dim colNombresSub As New Collection
    Dim key As Variant
    Dim relacion As ExpedienteSuministrador
    Dim unSuministrador As suministrador
    Dim nombreFinalSuministrador As String
    Dim coleccionAUsar As Collection
    Dim firstPrincipalDir As String
    Dim firstSubDir As String
    Dim direccionUTE As String
    On Error GoTo Errores
    
    Set sol = solServ.getSolicitudPorID(idSolicitud)
    If sol Is Nothing Then Err.Raise 513, "GetDatosGeneralesIniciales", "No se encontro la solicitud."
    
    Set exp = expServ.getExpedientePorID(CStr(sol.idExpediente))
    If exp Is Nothing Then Err.Raise 513, "GetDatosGeneralesIniciales", "No se encontro el expediente asociado."
    
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
    
    If colNombresSub.count > 0 Then
        Set coleccionAUsar = colNombresSub
    Else
        Set coleccionAUsar = colNombresPrincipales
    End If
    
    If Not coleccionAUsar Is Nothing Then
        If coleccionAUsar.count > 1 Then
            nombreFinalSuministrador = "UNION TEMPORAL DE EMPRESAS: " & JoinCollection(coleccionAUsar, " y ")
        ElseIf coleccionAUsar.count = 1 Then
            nombreFinalSuministrador = coleccionAUsar(1)
        End If
    End If
    
    direccionUTE = Trim(Nz(m_ObjEntorno.DirUTE, ""))
    If direccionUTE = "" Then
        If colNombresSub.count > 0 Then
            direccionUTE = Trim(Nz(firstSubDir, ""))
        Else
            direccionUTE = Trim(Nz(firstPrincipalDir, ""))
        End If
    End If
    If Nz(nombreFinalSuministrador, "") <> "" And direccionUTE <> "" Then
        nombreFinalSuministrador = nombreFinalSuministrador & " (" & direccionUTE & ")"
    End If
    
    With cdca
        .idSolicitud = idSolicitud
        .numContrato = exp.CodExpLargo
        .refSuministrador = sol.codigoSolicitud
        .SuministradorNombreDir = nombreFinalSuministrador
    End With
    
    Set GetDatosGeneralesIniciales = cdca
LimpiarYSalir:
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCAServicio.GetDatosGeneralesIniciales"
    errObj.Raise
End Function

Public Function GetNombreSuministradorDefault(ByVal idExpediente As Long) As String
    Dim expServ As New ExpedienteServicio
    Dim sumServ As New SuministradorServicio
    Dim exp As Expediente
    Dim colRelaciones As Scripting.Dictionary
    Dim colNombresPrincipales As New Collection
    Dim colNombresSub As New Collection
    Dim key As Variant
    Dim relacion As ExpedienteSuministrador
    Dim unSuministrador As suministrador
    Dim nombreFinalSuministrador As String
    
    On Error GoTo Errores

    ' 1. Obtener el objeto Expediente
    Set exp = expServ.getExpedientePorID(CStr(idExpediente))
    If exp Is Nothing Then GoTo LimpiarYSalir
    
    ' 2. Obtener las relaciones con suministradores
    Set colRelaciones = expServ.getSuministradoresPorExpediente(exp.idExpediente)
    If colRelaciones Is Nothing Then GoTo LimpiarYSalir

    ' 3. Recorrer las relaciones para obtener los nombres
    For Each key In colRelaciones.Keys
        Set relacion = colRelaciones(key)
        Set unSuministrador = sumServ.getSuministradorPorID(CStr(relacion.idSuministrador))
        If Not unSuministrador Is Nothing Then
            If relacion.ContratistaPrincipal = "Sí" Then
                colNombresPrincipales.Add UCase(Trim(unSuministrador.nombre))
            End If
            If relacion.SubContratista = "Sí" Then
                colNombresSub.Add UCase(Trim(unSuministrador.nombre))
            End If
        End If
    Next key
    
    ' 4. Aplicar la regla de negocio para construir el nombre final
    Dim coleccionAUsar As Collection
    If colNombresSub.count > 0 Then
        Set coleccionAUsar = colNombresSub
    Else
        Set coleccionAUsar = colNombresPrincipales
    End If

    If coleccionAUsar.count > 1 Then
        nombreFinalSuministrador = "UNIÓN TEMPORAL DE EMPRESAS: " & JoinCollection(coleccionAUsar, " y ")
    ElseIf coleccionAUsar.count = 1 Then
        nombreFinalSuministrador = coleccionAUsar(1)
    End If
    
    ' 5. Añadir la Dirección
    If Nz(nombreFinalSuministrador, "") <> "" Then
         nombreFinalSuministrador = nombreFinalSuministrador & " (" & m_ObjEntorno.DirUTE & ")"
    End If
    
    GetNombreSuministradorDefault = nombreFinalSuministrador

LimpiarYSalir:
    Exit Function
Errores:
    GetNombreSuministradorDefault = "Error al cargar suministrador"
    Resume LimpiarYSalir
End Function

Public Function GetDatosAprobacionIniciales(ByVal idSolicitud As Long) As DatosCDCA
    Dim cdca As New DatosCDCA
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
LimpiarYSalir:
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCAServicio.GetDatosAprobacionIniciales"
    errObj.Raise
End Function



Public Function GuardarDatosGenerales(objCDCA As DatosCDCA, Optional ByRef db As DAO.Database) As Boolean
    Dim solServ As New SolicitudServicio, wfServ As New WorkflowServicio, sol As Solicitud
    Dim transicionRealizada As Boolean
    
    ' Variables Transacción
    Dim ws As DAO.Workspace
    Dim dbConexion As DAO.Database
    
    On Error GoTo Errores
    transicionRealizada = False
    
    ' Validaciones
    If Trim(Nz(objCDCA.numContrato, "")) = "" Then Err.Raise 513, "Validación", "El campo 'Nº del contrato' es obligatorio."
    If Trim(Nz(objCDCA.refSuministrador, "")) = "" Then Err.Raise 513, "Validación", "El campo 'Referencia del Suministrador' es obligatorio."
    If Trim(Nz(objCDCA.SuministradorNombreDir, "")) = "" Then Err.Raise 513, "Validación", "El campo 'Suministrador (Nombre y Dirección)' es obligatorio."
    If IsNull(objCDCA.requiereModificacionContrato) Then Err.Raise 513, "Validación", "Debe indicar si 'Requiere Modificación del contrato' (Sí/No)."
    
    ' INICIO TRANSACCIÓN
    If db Is Nothing Then
        Set ws = getWorkspace()
        Set dbConexion = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
    Else
        Set dbConexion = db
    End If
    
    ' Persistencia
    ' Usamos el método Guardar compatible con transacciones
    Call DatosCDCARepositorio.Guardar(objCDCA, dbConexion)
    
    ' Transición automática
    Set sol = solServ.getSolicitudPorID(objCDCA.idSolicitud, dbConexion)
    If sol.idEstadoInterno = estadoPreregistro Then
        ' Pasamos la conexión 'dbConexion'
        Call wfServ.EjecutarTransicion(sol, estadoRegistro, m_ObjUsuarioActivo, dbConexion)
        transicionRealizada = True
    End If
    
    ' COMMIT
    If db Is Nothing Then ws.CommitTrans
    
    GuardarDatosGenerales = transicionRealizada
    
LimpiarYSalir:
    On Error Resume Next
    If db Is Nothing And Not dbConexion Is Nothing Then dbConexion.Close
    If db Is Nothing Then Set ws = Nothing
    Exit Function
Errores:
     ' 1. CAPTURA
    Dim savedErr As Long, savedDesc As String
    savedErr = Err.Number: savedDesc = Err.description

    ' 2. ROLLBACK SEGURO
    If db Is Nothing And Not ws Is Nothing Then
        On Error Resume Next
        ws.Rollback
    End If
    
    ' 3. PROPAGACIÓN
    Dim errObj As New CondorError
    errObj.Create savedErr, savedDesc, "DatosCDCAServicio.GuardarDatosGenerales"
    errObj.AddToCallStack "idSolicitud: " & objCDCA.idSolicitud
    errObj.Raise
End Function


Public Sub GuardarPropuesta(objCDCA As DatosCDCA, Optional db As DAO.Database)
    On Error GoTo Errores
    
    Dim ws As DAO.Workspace
    Dim dbTrabajo As DAO.Database
    Dim gestionaTransaccionLocal As Boolean
    Dim requiereTransaccion As Boolean
    ' Validacion
    If Trim(Nz(objCDCA.identificacionMaterial, "")) = "" Then Err.Raise 513, "Validacion", "El campo 'Identificacion del Material' es obligatorio."
    If Trim(Nz(objCDCA.causaNC, "")) = "" Then Err.Raise 513, "Validacion", "El campo 'Causa de la No Conformidad' es obligatorio."
    If Trim(Nz(objCDCA.descripcionImpactoNC, "")) = "" Then Err.Raise 513, "Validacion", "El campo 'Descripcion/Impacto de la No Conformidad' es obligatorio."
    
    ' --- TRAZABILIDAD RECHAZO (PREPARACION) ---
    Dim rechazoActivo As rechazo
    Dim objOriginal As DatosCDCA
    Set rechazoActivo = RechazoRepositorio.GetUltimoRechazoActivo(objCDCA.idSolicitud, dbTrabajo)
    
    If Not rechazoActivo Is Nothing Then
        Set objOriginal = Me.GetDatosCDCA(objCDCA.idSolicitud, dbTrabajo)
    End If
    ' ------------------------------------------
    
    requiereTransaccion = (db Is Nothing And EsSolicitudEnValidacion(objCDCA.idSolicitud))
    
    If requiereTransaccion Then
        Set ws = getWorkspace()
        Set dbTrabajo = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
        gestionaTransaccionLocal = True
    ElseIf db Is Nothing Then
        Set dbTrabajo = getdb()
    Else
        Set dbTrabajo = db
    End If
    
    ' Persistencia GRANULAR
    Call DatosCDCARepositorio.ActualizarPropuesta(objCDCA, dbTrabajo)
    
    ' --- TRAZABILIDAD RECHAZO (REGISTRO) ---
    If Not rechazoActivo Is Nothing And Not objOriginal Is Nothing Then
        Dim jsonDict As Object
        Set jsonDict = JsonHelper.ObtenerJsonRechazo(objCDCA.idSolicitud)
        
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Propuesta", "identificacionMaterial", objOriginal.identificacionMaterial, objCDCA.identificacionMaterial
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Propuesta", "numPlanoEspecificacion", objOriginal.numPlanoEspecificacion, objCDCA.numPlanoEspecificacion
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Propuesta", "cantidadPeriodo", objOriginal.cantidadPeriodo, objCDCA.cantidadPeriodo
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Propuesta", "numSerieLote", objOriginal.numSerieLote, objCDCA.numSerieLote
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Propuesta", "causaNC", objOriginal.causaNC, objCDCA.causaNC
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Propuesta", "descripcionImpactoNC", objOriginal.descripcionImpactoNC, objCDCA.descripcionImpactoNC
        
        JsonHelper.GuardarJsonRechazo objCDCA.idSolicitud, jsonDict
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
    If gestionaTransaccionLocal And Not (ws Is Nothing) Then ws.Rollback
    If gestionaTransaccionLocal And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCAServicio.GuardarPropuesta"
    errObj.Raise
End Sub

Public Sub GuardarImpacto(objCDCA As DatosCDCA, Optional db As DAO.Database)
' PROPOSITO: Validar y orquestar el guardado del bloque "Impacto".
' REGLAS: En Impacto todos son obligatorios.
' -----------------------------------------------------------------------
    On Error GoTo Errores
    
    Dim ws As DAO.Workspace
    Dim dbTrabajo As DAO.Database
    Dim gestionaTransaccionLocal As Boolean
    Dim requiereTransaccion As Boolean
    
    ' Validacion
    Dim coste As String, clasificacion As String
    coste = UCase(Trim(Nz(objCDCA.impactoCoste, "")))
    clasificacion = UCase(Trim(Nz(objCDCA.clasificacionNC, "")))

    If coste <> "AUMENTARÁ" And coste <> "DISMINUIRÁ" And coste <> "NO VARIARÁ" Then Err.Raise 513, "Validacion", "Debe seleccionar una opcion para 'Coste'."
    If clasificacion <> "MAYOR" And clasificacion <> "MENOR" Then Err.Raise 513, "Validacion", "Debe seleccionar si la 'NC Considerada' es Mayor o Menor."
If IsNull(objCDCA.esSuministradorAD) Then Err.Raise 513, "Validacion", "Debe indicar si 'El suministrador es la Autoridad de Diseno' (Si/No)."
    
    ' --- INICIALIZAR dbTrabajo PRIMERO para consistencia transaccional ---
    requiereTransaccion = (db Is Nothing And EsSolicitudEnValidacion(objCDCA.idSolicitud))
    
    If requiereTransaccion Then
        Set ws = getWorkspace()
        Set dbTrabajo = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
        gestionaTransaccionLocal = True
    ElseIf db Is Nothing Then
        Set dbTrabajo = getdb()
    Else
        Set dbTrabajo = db
    End If
    
    ' --- TRAZABILIDAD RECHAZO (PREPARACION) ---
    Dim rechazoActivo As rechazo
    Dim objOriginal As DatosCDCA
    Set rechazoActivo = RechazoRepositorio.GetUltimoRechazoActivo(objCDCA.idSolicitud, dbTrabajo)
    
    If Not rechazoActivo Is Nothing Then
        Set objOriginal = Me.GetDatosCDCA(objCDCA.idSolicitud, dbTrabajo)
    End If
    ' ------------------------------------------
    
    ' Persistencia GRANULAR
    Call DatosCDCARepositorio.ActualizarImpacto(objCDCA, dbTrabajo)
    
    ' --- TRAZABILIDAD RECHAZO (REGISTRO) ---
    If Not rechazoActivo Is Nothing And Not objOriginal Is Nothing Then
        Dim jsonDict As Object
        Set jsonDict = JsonHelper.ObtenerJsonRechazo(objCDCA.idSolicitud)
        
        ' Caracteristicas Afectadas (Booleanos)
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Impacto", "afectaPrestaciones", objOriginal.afectaPrestaciones, objCDCA.afectaPrestaciones
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Impacto", "afectaSeguridad", objOriginal.afectaSeguridad, objCDCA.afectaSeguridad
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Impacto", "afectaFiabilidad", objOriginal.afectaFiabilidad, objCDCA.afectaFiabilidad
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Impacto", "afectaVidaUtil", objOriginal.afectaVidaUtil, objCDCA.afectaVidaUtil
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Impacto", "afectaMedioambiente", objOriginal.afectaMedioambiente, objCDCA.afectaMedioambiente
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Impacto", "afectaMantenibilidad", objOriginal.afectaMantenibilidad, objCDCA.afectaMantenibilidad
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Impacto", "afectaIntercambiabilidad", objOriginal.afectaIntercambiabilidad, objCDCA.afectaIntercambiabilidad
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Impacto", "afectaApariencia", objOriginal.afectaApariencia, objCDCA.afectaApariencia
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Impacto", "afectaOtros", objOriginal.afectaOtros, objCDCA.afectaOtros
        
        ' Campos de Texto y Radios
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Impacto", "impactoCoste", objOriginal.impactoCoste, objCDCA.impactoCoste
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Impacto", "clasificacionNC", objOriginal.clasificacionNC, objCDCA.clasificacionNC
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Impacto", "esSuministradorAD", objOriginal.esSuministradorAD, objCDCA.esSuministradorAD
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Impacto", "identificacionAutoridadDiseno", objOriginal.identificacionAutoridadDiseno, objCDCA.identificacionAutoridadDiseno
        JsonHelper.RegistrarCambio jsonDict, "CDCA_Impacto", "efectoFechaEntrega", objOriginal.efectoFechaEntrega, objCDCA.efectoFechaEntrega
        
        JsonHelper.GuardarJsonRechazo objCDCA.idSolicitud, jsonDict
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
    If gestionaTransaccionLocal And Not (ws Is Nothing) Then ws.Rollback
    If gestionaTransaccionLocal And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCAServicio.GuardarImpacto"
    errObj.Raise
End Sub
Public Sub GuardarAprobacionSuministrador(objCDCA As DatosCDCA, Optional ByRef db As DAO.Database)
    On Error GoTo Errores
    
    If db Is Nothing Then Set db = getdb()
    
    ' Validaciones Fijas
    If Trim(Nz(objCDCA.firmaAprobacionRespIngenieriaNombre, "")) = "" Then Err.Raise 513, "Validación", "El nombre del Responsable de Ingeniería es obligatorio."
    If Trim(Nz(objCDCA.firmaAprobacionRespCalidadNombre, "")) = "" Then Err.Raise 513, "Validación", "El nombre del Responsable de Calidad es obligatorio."
    
    ' Persistencia GRANULAR
    Call DatosCDCARepositorio.ActualizarAprobacionSuministrador(objCDCA, db)
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCAServicio.GuardarAprobacionSuministrador"
    errObj.Raise
End Sub

' En DatosCDCAServicio.cls

Public Sub GuardarDictamenRAC(objCDCA As DatosCDCA, Optional db As DAO.Database)
    On Error GoTo Errores
    
    Dim ws As DAO.Workspace
    Dim dbTrabajo As DAO.Database
    Dim gestionaTransaccionLocal As Boolean
    Dim requiereTransaccion As Boolean
    ' Validación
    If objCDCA.racDecision <> "RECHAZADO" Then
        If Trim(Nz(objCDCA.racCodigo, "")) = "" Then Err.Raise 513, "Validación", "El campo 'Código' es obligatorio."
        If Trim(Nz(objCDCA.racNombre, "")) = "" Then Err.Raise 513, "Validación", "El campo 'Nombre' es obligatorio."
    End If
    
    requiereTransaccion = (db Is Nothing And EsSolicitudEnValidacion(objCDCA.idSolicitud))
    
    If requiereTransaccion Then
        Set ws = getWorkspace()
        Set dbTrabajo = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
        gestionaTransaccionLocal = True
    Else
        Set dbTrabajo = db
    End If
    
    ' Persistencia GRANULAR con soporte transaccional
    Call DatosCDCARepositorio.ActualizarDictamenRAC(objCDCA, dbTrabajo)
    
    If gestionaTransaccionLocal Then
        ws.CommitTrans
    End If
    
LimpiarYSalir:
    If gestionaTransaccionLocal And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Exit Sub
Errores:
    If gestionaTransaccionLocal And Not (ws Is Nothing) Then ws.Rollback
    If gestionaTransaccionLocal And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCAServicio.GuardarDictamenRAC"
    errObj.Raise
End Sub

Public Sub GuardarDecisionFinal(ByRef objCDCA As DatosCDCA, Optional ByRef db As DAO.Database)
    On Error GoTo Errores

    If Trim(Nz(objCDCA.decisionFinal, "")) = "" Then
        Err.Raise 513, "Regla de Negocio Incumplida", "« (Decisión Final) La decisión es obligatoria."
    End If

    ' Persistencia GRANULAR (Ya existía y es correcta)
    Call DatosCDCARepositorio.ActualizarDecisionFinal(objCDCA, db)

    Exit Sub

Errores:
    Dim savedErr As Long, savedDesc As String
    savedErr = Err.Number: savedDesc = Err.description
    Dim errObj As New CondorError
    errObj.Create savedErr, savedDesc, "DatosCDCAServicio.GuardarDecisionFinal"
    errObj.Raise
End Sub

Public Function EsDatosGeneralesCompleta(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objCDCA As DatosCDCA
    On Error GoTo Errores
    
    EsDatosGeneralesCompleta = False
    ' PASAMOS DB
    Set objCDCA = Me.GetDatosCDCA(idSolicitud, db)
    If objCDCA Is Nothing Then Exit Function
    If Trim(Nz(objCDCA.numContrato, "")) = "" Then Exit Function
    If Trim(Nz(objCDCA.refSuministrador, "")) = "" Then Exit Function
    If Trim(Nz(objCDCA.SuministradorNombreDir, "")) = "" Then Exit Function
    If Trim(Nz(objCDCA.requiereModificacionContrato, "")) = "" Then Exit Function
    
    
    EsDatosGeneralesCompleta = True
    
LimpiarYSalir:
    Exit Function
Errores:
    EsDatosGeneralesCompleta = False
    Resume LimpiarYSalir
End Function
Public Function EsParteTecnicaCompleta(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objCDCA As DatosCDCA
    On Error GoTo Errores
    
    EsParteTecnicaCompleta = False
    ' PASAMOS DB
    Set objCDCA = Me.GetDatosCDCA(idSolicitud, db)
    If objCDCA Is Nothing Then Exit Function

    ' Propuesta
    If Trim(Nz(objCDCA.identificacionMaterial, "")) = "" Then Exit Function
    If Trim(Nz(objCDCA.causaNC, "")) = "" Then Exit Function
    If Trim(Nz(objCDCA.descripcionImpactoNC, "")) = "" Then Exit Function

    ' Impacto
    Dim coste As String, clasificacion As String
    coste = UCase(Trim(Nz(objCDCA.impactoCoste, "")))
    clasificacion = UCase(Trim(Nz(objCDCA.clasificacionNC, "")))
    If coste <> "AUMENTARÁ" And coste <> "DISMINUIRÁ" And coste <> "NO VARIARÁ" Then Exit Function
    If clasificacion <> "MAYOR" And clasificacion <> "MENOR" Then Exit Function
    If IsNull(objCDCA.esSuministradorAD) Then Exit Function
    
    EsParteTecnicaCompleta = True
    
LimpiarYSalir:
    Exit Function
Errores:
    EsParteTecnicaCompleta = False
    Resume LimpiarYSalir
End Function

Public Function EsDetalleCompleto(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objCDCA As DatosCDCA
    On Error GoTo Errores
    
    EsDetalleCompleto = False
    Set objCDCA = Me.GetDatosCDCA(idSolicitud, db)
    If objCDCA Is Nothing Then Exit Function
    
    If Trim(Nz(objCDCA.identificacionMaterial, "")) = "" Then Exit Function
    If Trim(Nz(objCDCA.causaNC, "")) = "" Then Exit Function
    If Trim(Nz(objCDCA.descripcionImpactoNC, "")) = "" Then Exit Function
    
    EsDetalleCompleto = True
    Exit Function

Errores:
    EsDetalleCompleto = False
   
End Function

Public Function EsMotivosCompleto(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objCDCA As DatosCDCA

    On Error GoTo Errores

    EsMotivosCompleto = False
    Set objCDCA = Me.GetDatosCDCA(idSolicitud, db)
    If objCDCA Is Nothing Then Exit Function

    If Trim(Nz(objCDCA.descripcionImpactoNC, "")) = "" Then Exit Function
    If Trim(Nz(objCDCA.descripcionImpactoNCCont, "")) = "" Then Exit Function

    EsMotivosCompleto = True
    Exit Function

Errores:
    EsMotivosCompleto = False
    
End Function

Public Function EsDictamenRACCompleto(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objCDCA As DatosCDCA
    On Error GoTo Errores
    
    EsDictamenRACCompleto = False
    
    ' --- CAMBIO CLAVE: Pasamos 'db' al método getDatosCDCA ---
    ' Esto permite leer los datos pendientes de confirmar dentro de la transacción
    Set objCDCA = Me.GetDatosCDCA(idSolicitud, db)
    
    If objCDCA Is Nothing Then Exit Function

    If Trim(Nz(objCDCA.racCodigo, "")) = "" Then Exit Function
    If Trim(Nz(objCDCA.racNombre, "")) = "" Then Exit Function
    If Trim(Nz(objCDCA.racDecision, "")) = "" Then Exit Function
    
    EsDictamenRACCompleto = True
LimpiarYSalir:
    Exit Function
Errores:
    EsDictamenRACCompleto = False
    Resume LimpiarYSalir
End Function


Public Function EsAprobacionSuministradorCompleta(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objCDCA As DatosCDCA
    On Error GoTo Errores
    
    EsAprobacionSuministradorCompleta = False
    ' PASAMOS DB
    Set objCDCA = Me.GetDatosCDCA(idSolicitud, db)
    If objCDCA Is Nothing Then Exit Function
    
    If Trim(Nz(objCDCA.firmaAprobacionRespIngenieriaNombre, "")) = "" Then Exit Function
    If Trim(Nz(objCDCA.firmaAprobacionRespCalidadNombre, "")) = "" Then Exit Function
    If Trim(Nz(objCDCA.decisionFinal, "")) = "" Then Exit Function
    
    EsAprobacionSuministradorCompleta = True
LimpiarYSalir:
    Exit Function
Errores:
    EsAprobacionSuministradorCompleta = False
    Resume LimpiarYSalir
End Function
Public Function EsDecisionFinalCompleta(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objCDCA As DatosCDCA
    On Error GoTo Errores
    
    EsDecisionFinalCompleta = False
    
    ' Leemos usando la conexión transaccional para ver los datos recién guardados
    Set objCDCA = Me.GetDatosCDCA(idSolicitud, db)
    
    If objCDCA Is Nothing Then Exit Function

    If Not DecisionFinalHelper_EsCompleta(objCDCA.decisionFinal, objCDCA.NombreFirmanteFinal) Then Exit Function
    
    EsDecisionFinalCompleta = True
    
LimpiarYSalir:
    Exit Function
Errores:
    EsDecisionFinalCompleta = False
    Resume LimpiarYSalir
End Function

Public Sub EliminarDictamenRAC(ByVal idSolicitud As Long, Optional db As DAO.Database)
    ' RESPONSABILIDAD: Orquestar el blanqueo de datos del RAC delegando en el repositorio.
    On Error GoTo Errores
    
    ' Delegación directa al repositorio pasando la conexión transaccional si existe
    Call DatosCDCARepositorio.LimpiarDictamenRAC(idSolicitud, db)
    
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCAServicio.EliminarDictamenRAC"
    errObj.Raise
End Sub

Public Sub EliminarAprobacionSuministrador(ByVal idSolicitud As Long, Optional db As DAO.Database)
    On Error GoTo Errores
    Call DatosCDCARepositorio.LimpiarAprobacionSuministrador(idSolicitud, db)
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCAServicio.EliminarAprobacionSuministrador"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud: errObj.Raise
End Sub

' ==========================================================================
' MÓDULO: DatosCDCAServicio.cls
' ACCIÓN: Método transaccional para limpiar la Decisión Final
' ==========================================================================
Public Sub EliminarDecisionFinal(ByVal idSolicitud As Long, Optional db As DAO.Database)
    On Error GoTo Errores
    Call DatosCDCARepositorio.LimpiarDecisionFinal(idSolicitud, db)
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCAServicio.EliminarDecisionFinal"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud: errObj.Raise
End Sub

Public Sub RegistrarRechazoDesdeFormalizacion(ByVal idSolicitud As Long, ByVal motivoRechazo As String, ByVal usuarioActivo As usuario, Optional ByRef db As DAO.Database)
    Dim ws As DAO.Workspace
    Dim dbConexion As DAO.Database
    Dim wfServ As New WorkflowServicio
    Dim cdca As DatosCDCA
    
    On Error GoTo Errores
    
    ' Gestión de conexión transaccional
    If db Is Nothing Then
        Set ws = getWorkspace()
        Set dbConexion = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
    Else
        Set dbConexion = db
    End If
    
    Set cdca = Me.GetDatosCDCA(idSolicitud, dbConexion)
    If cdca Is Nothing Then Err.Raise 513, "RechazoFormalizacion", "No se encontraron datos para la solicitud."

    ' Limpiar Formalización
    cdca.decisionFinal = ""
    cdca.observacionesFinales = ""
    cdca.NombreFirmanteFinal = ""
    
    ' Limpiar RAC
    cdca.racCodigo = ""
    cdca.racNombre = ""
    cdca.observacionesRAC = ""
    
    ' Establecer Rechazo
    cdca.racDecision = "RECHAZADO"
    
    ' Persistencia Transaccional (Usamos el Guardar completo para este caso especial transaccional)
    Call DatosCDCARepositorio.Guardar(cdca, dbConexion)
    
    Call wfServ.EjecutarTransicion(SolicitudRepositorio.getSolicitudPorID(idSolicitud, dbConexion), estadoRechazada, usuarioActivo, dbConexion)
    
    If db Is Nothing Then ws.CommitTrans

LimpiarYSalir:
    On Error Resume Next
    If db Is Nothing And Not dbConexion Is Nothing Then dbConexion.Close
    If db Is Nothing Then Set ws = Nothing
    Exit Sub
    
Errores:
    If db Is Nothing And Not ws Is Nothing Then ws.Rollback
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosCDCAServicio.RegistrarRechazoDesdeFormalizacion"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud: errObj.Raise
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

Public Sub ImportarEntidad_aVM(ByRef cdca As DatosCDCA, ByRef vm As DatosCDCAViewModel)
    ' Campos Generales
    vm.vm_numContrato = cdca.numContrato
    vm.vm_refSuministrador = cdca.refSuministrador
    vm.vm_SuministradorNombreDir = cdca.SuministradorNombreDir
    vm.vm_refDesviacionesPrevias = cdca.refDesviacionesPrevias
    vm.vm_requiereModificacionContrato = cdca.requiereModificacionContrato
    
    ' Propuesta
    vm.vm_identificacionMaterial = cdca.identificacionMaterial
    vm.vm_numPlanoEspecificacion = cdca.numPlanoEspecificacion
    vm.vm_cantidadPeriodo = cdca.cantidadPeriodo
    vm.vm_numSerieLote = cdca.numSerieLote
    vm.vm_causaNC = cdca.causaNC
    vm.vm_descripcionImpactoNC = cdca.descripcionImpactoNC
    
    ' Impacto
    vm.vm_afectaPrestaciones = cdca.afectaPrestaciones
    vm.vm_afectaSeguridad = cdca.afectaSeguridad
    vm.vm_afectaFiabilidad = cdca.afectaFiabilidad
    vm.vm_afectaVidaUtil = cdca.afectaVidaUtil
    vm.vm_afectaMedioambiente = cdca.afectaMedioambiente
    vm.vm_afectaMantenibilidad = cdca.afectaMantenibilidad
    vm.vm_afectaIntercambiabilidad = cdca.afectaIntercambiabilidad
    vm.vm_afectaApariencia = cdca.afectaApariencia
    vm.vm_afectaOtros = cdca.afectaOtros
    vm.vm_impactoCoste = cdca.impactoCoste
    vm.vm_clasificacionNC = cdca.clasificacionNC
    vm.vm_esSuministradorAD = cdca.esSuministradorAD
    vm.vm_identificacionAutoridadDiseno = cdca.identificacionAutoridadDiseno
    vm.vm_efectoFechaEntrega = cdca.efectoFechaEntrega
    
    ' Aprobación Suministrador
    vm.vm_firmaAprobacionRespIngenieriaNombre = cdca.firmaAprobacionRespIngenieriaNombre
    vm.vm_firmaAprobacionRespCalidadNombre = cdca.firmaAprobacionRespCalidadNombre
    vm.vm_firmaAprobacionRespProduccionNombre = cdca.firmaAprobacionRespProduccionNombre
    vm.vm_firmaAprobacionRespDisenioNombre = cdca.firmaAprobacionRespDisenioNombre
    vm.vm_firmaAprobacionRepresentanteSumNombre = cdca.firmaAprobacionRepresentanteSumNombre
    
    ' Dictamen RAC
    vm.vm_racCodigo = cdca.racCodigo
    vm.vm_observacionesRAC = cdca.observacionesRAC
    vm.vm_racNombre = cdca.racNombre
    vm.vm_racDecision = Nz(cdca.racDecision, "")
    
    ' Decisión Final
    vm.vm_decisionFinal = cdca.decisionFinal
    vm.vm_obsDecisionFinal = cdca.observacionesFinales
    vm.vm_NombreFirmanteFinal = cdca.NombreFirmanteFinal
    
    ' -------------------------------------------------------
    ' LÓGICA DE INFERENCIA DE ESTADO (Sí/No)
    ' -------------------------------------------------------
    If Len(Trim(Nz(cdca.firmaAprobacionRespProduccionNombre, ""))) > 0 Then
        vm.vm_aplicaRespProduccion = True
    Else
        vm.vm_aplicaRespProduccion = False
    End If
    
    If Len(Trim(Nz(cdca.firmaAprobacionRespDisenioNombre, ""))) > 0 Then
        vm.vm_aplicaRespDisenio = True
    Else
        vm.vm_aplicaRespDisenio = False
    End If
    
    If Len(Trim(Nz(cdca.firmaAprobacionRepresentanteSumNombre, ""))) > 0 Then
        vm.vm_aplicaRepresentanteSum = True
    Else
        vm.vm_aplicaRepresentanteSum = False
    End If
End Sub

Public Sub ExportarVM_aEntidad(ByVal bloque As enumBloqueFormulario, ByRef vm As DatosCDCAViewModel, ByRef cdca As DatosCDCA)
    Select Case bloque
        Case Bloque_Generales
            cdca.numContrato = vm.vm_numContrato
            cdca.refSuministrador = vm.vm_refSuministrador
            cdca.SuministradorNombreDir = vm.vm_SuministradorNombreDir
            cdca.refDesviacionesPrevias = vm.vm_refDesviacionesPrevias
            cdca.requiereModificacionContrato = vm.vm_requiereModificacionContrato
        Case Bloque_Propuesta
            cdca.identificacionMaterial = vm.vm_identificacionMaterial
            cdca.numPlanoEspecificacion = vm.vm_numPlanoEspecificacion
            cdca.cantidadPeriodo = vm.vm_cantidadPeriodo
            cdca.numSerieLote = vm.vm_numSerieLote
            cdca.causaNC = vm.vm_causaNC
            cdca.descripcionImpactoNC = vm.vm_descripcionImpactoNC
        Case Bloque_Impacto
            cdca.afectaPrestaciones = vm.vm_afectaPrestaciones
            cdca.afectaSeguridad = vm.vm_afectaSeguridad
            cdca.afectaFiabilidad = vm.vm_afectaFiabilidad
            cdca.afectaVidaUtil = vm.vm_afectaVidaUtil
            cdca.afectaMedioambiente = vm.vm_afectaMedioambiente
            cdca.afectaMantenibilidad = vm.vm_afectaMantenibilidad
            cdca.afectaIntercambiabilidad = vm.vm_afectaIntercambiabilidad
            cdca.afectaApariencia = vm.vm_afectaApariencia
            cdca.afectaOtros = vm.vm_afectaOtros
            cdca.impactoCoste = vm.vm_impactoCoste
            cdca.clasificacionNC = vm.vm_clasificacionNC
            cdca.esSuministradorAD = vm.vm_esSuministradorAD
            cdca.identificacionAutoridadDiseno = vm.vm_identificacionAutoridadDiseno
            cdca.efectoFechaEntrega = vm.vm_efectoFechaEntrega
        Case Bloque_AprobacionSuministrador
            cdca.firmaAprobacionRespIngenieriaNombre = vm.vm_firmaAprobacionRespIngenieriaNombre
            cdca.firmaAprobacionRespCalidadNombre = vm.vm_firmaAprobacionRespCalidadNombre
            cdca.firmaAprobacionRespProduccionNombre = vm.vm_firmaAprobacionRespProduccionNombre
            cdca.firmaAprobacionRespDisenioNombre = vm.vm_firmaAprobacionRespDisenioNombre
            cdca.firmaAprobacionRepresentanteSumNombre = vm.vm_firmaAprobacionRepresentanteSumNombre
        Case Bloque_DictamenRAC
            cdca.racCodigo = vm.vm_racCodigo
            cdca.observacionesRAC = vm.vm_observacionesRAC
            cdca.racDecision = vm.vm_racDecision
            cdca.racNombre = vm.vm_racNombre
        Case Bloque_DecisionFinal
            cdca.decisionFinal = vm.vm_decisionFinal
            cdca.observacionesFinales = vm.vm_obsDecisionFinal
            cdca.NombreFirmanteFinal = vm.vm_NombreFirmanteFinal
    End Select
End Sub

Public Function ObtenerViewModelCompleto(ByVal idSolicitud As Long, Optional db As DAO.Database) As DatosCDCAViewModel
    Dim vm As New DatosCDCAViewModel
    Dim cdca As DatosCDCA
    Dim sol As Solicitud
    Dim exp As Expediente
    Dim expServ As New ExpedienteServicio
    Dim solServ As New SolicitudServicio
    Dim estadoServ As New EstadoServicio
    Dim estadoActual As estado
    Dim idEstado As Long
    
    On Error GoTo Errores
    
    ' Cargar Solicitud
    Set sol = solServ.getSolicitudPorID(idSolicitud, db)
    If sol Is Nothing Then Err.Raise 513, "ObtenerViewModelCompleto", "No se encontró la solicitud."
    
    ' Cargar Expediente
    Set exp = expServ.getExpedientePorID(CStr(sol.idExpediente))
    If exp Is Nothing Then Err.Raise 513, "ObtenerViewModelCompleto", "No se encontró el expediente."
    
    ' Cargar DatosCDCA
    Set cdca = Me.GetDatosCDCA(idSolicitud, db)
    If cdca Is Nothing Then Set cdca = Me.GetDatosGeneralesIniciales(idSolicitud)
    
    ' Armar VM
    Set vm.Solicitud = sol
    Set vm.Expediente = exp
    Set vm.datos = cdca
    
    ' Estado
    Set estadoActual = estadoServ.getEstadoPorID(sol.idEstadoInterno)
    If Not estadoActual Is Nothing Then
        vm.EstadoNombre = estadoActual.nombreEstado
    End If
    
    ' Tiempo Ciclo
    If estadoActual.esEstadoFinal Then
        vm.TiempoCiclo = CalcularDuracion(sol.fechaCreacion, sol.fechaModificacion)
    Else
        vm.TiempoCiclo = CalcularDuracion(sol.fechaCreacion, Now()) & " (en curso)"
    End If
    
    ' Mapear entidad a VM
    Call ImportarEntidad_aVM(cdca, vm)
    
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
    errObj.Create Err.Number, Err.description, "DatosCDCAServicio.ObtenerViewModelCompleto"
    errObj.AddToCallStack "idSolicitud: " & idSolicitud
    errObj.Raise
End Function



