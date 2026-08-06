VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatosPCServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit


Public Sub ActualizarCamposDependientesDeExpediente(ByRef sol As Solicitud, ByRef exp As Expediente, Optional db As DAO.Database)
    Dim pc As DatosPC
    Dim expServ As New ExpedienteServicio
    Dim dictNombres As Scripting.Dictionary
    
    On Error GoTo Errores
    
    ' 1. Obtener o crear el objeto de datos tecnicos
    ' Nota: getDatosPC solo lee, no necesita transaccion obligatoria aqui,
    ' pero si estamos en una, es bueno pasarla para leer "lo que acabamos de escribir" si fuera el caso.
    Set pc = Me.getDatosPC(sol.idSolicitud, db)
    
    If pc Is Nothing Then
        Set pc = New DatosPC
        pc.idSolicitud = sol.idSolicitud
    End If
    
    ' 2. Actualizar los campos dependientes del expediente
    pc.refContratoInspeccionOficial = exp.CodExpLargo
    pc.refSuministrador = sol.codigoSolicitud
    pc.denominacionContrato = exp.titulo
    pc.objetoContrato = exp.objetoContrato
    pc.firmaOficinaTecnicaNombre = exp.ResponsableTecnico
    
    Set dictNombres = expServ.getNombresSuministradores(exp.idExpediente)
    pc.SuministradorNombreDir = dictNombres("principal")
    
    ' 3. Persistir los cambios (PASANDO LA TRANSACCION db)
    Call DatosPCRepositorio.Guardar(pc, db)
    
LimpiarYSalir:
    Set expServ = Nothing
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCServicio.ActualizarCamposDependientesDeExpediente"
    errObj.Raise
End Sub

Public Sub PurgaTecnica(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database)
    ' Responsabilidad: Tabla Rasa de la parte técnica (Propuesta e Impacto)
    Dim pc As DatosPC
    On Error GoTo Errores
    Set pc = Me.getDatosPC(idSolicitud, db)
    If pc Is Nothing Then Exit Sub
    
    With pc
        ' Limpiar Propuesta
        .descripcionPropuestaCambio = ""
        .descripcionPropuestaCambioCont = ""
        .numPlanoEspecificacion = ""
        
        ' Limpiar Motivos (Booleanos)
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
    
    Call DatosPCRepositorio.Guardar(pc, db)
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCServicio.PurgaTecnica"
    errObj.Raise
End Sub
Public Function GetDatosAprobacionSuministradorIniciales(ByVal idSolicitud As Long) As DatosPC
    ' RESPONSABILIDAD: Construye un objeto DatosPC NO PERSISTIDO con los valores
    '                  por defecto para la seccion de Aprobacion del Suministrador.
    Dim pc As New DatosPC
    Dim solServ As New SolicitudServicio
    Dim expServ As New ExpedienteServicio
    Dim sol As Solicitud
    Dim exp As Expediente
    
    On Error GoTo Errores
    
    Set sol = solServ.getSolicitudPorID(idSolicitud)
    If sol Is Nothing Then Err.Raise 513, "GetDatosAprobacionSuministradorIniciales", "No se encontro la solicitud."
    
    Set exp = expServ.getExpedientePorID(CStr(sol.idExpediente))
    If exp Is Nothing Then Err.Raise 513, "GetDatosAprobacionSuministradorIniciales", "No se encontro el expediente asociado."

    With pc
        ' --- CAMPOS CON VALORES POR DEFECTO ---
        .firmaOficinaTecnicaNombre = AgregarSufijoRol(exp.ResponsableTecnico, "Jefe de Proyecto")
        
        .firmaRepSuministradorNombre = AgregarSufijoRol(m_ObjEntorno.NombreJefeCalidad, "Responsable de Calidad")
    End With
    
    Set GetDatosAprobacionSuministradorIniciales = pc
    
LimpiarYSalir:
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCServicio.GetDatosAprobacionSuministradorIniciales"
    errObj.Raise
End Function
' ==========================================================================
' MODULO: DatosPCServicio.cls
' ACCION: REEMPLAZO DEFINITIVO de GetDatosPCIniciales con la logica completa.
' ==========================================================================
Public Function GetDatosPCIniciales(ByVal idSolicitud As Long) As DatosPC
    Dim pc As New DatosPC
    Dim solServ As New SolicitudServicio
    Dim expServ As New ExpedienteServicio
    Dim sumServ As New SuministradorServicio
    Dim sol As Solicitud
    Dim exp As Expediente
    
    On Error GoTo Errores
    
    Set sol = solServ.getSolicitudPorID(idSolicitud)
    If sol Is Nothing Then Err.Raise 513, "GetDatosPCIniciales", "No se encontro la solicitud."
    
    Set exp = expServ.getExpedientePorID(CStr(sol.idExpediente))
    If exp Is Nothing Then Err.Raise 513, "GetDatosPCIniciales", "No se encontro el expediente asociado."

    With pc
        .idSolicitud = idSolicitud
        
        ' --- CAMPOS INMUTABLES ---
        .refContratoInspeccionOficial = exp.CodExpLargo
        .refSuministrador = sol.codigoSolicitud
        
        ' --- CAMPOS MUTABLES (Valores por defecto) ---
        .denominacionContrato = exp.titulo
        .objetoContrato = exp.objetoContrato
        
        ' --- INICIO DE LA LOGICA RESTAURADA PARA SUMINISTRADOR ---
        Dim colRelaciones As Scripting.Dictionary
        Dim colNombresPrincipales As New Collection
        Dim colNombresSub As New Collection
        Dim key As Variant, relacion As ExpedienteSuministrador, unSuministrador As suministrador
        Dim nombreFinalSuministrador As String
        Dim firstPrincipalDir As String
        Dim firstSubDir As String
        Dim direccionUTE As String
        
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
        
        ' Regla de negocio: "Si somos subcontratistas, mostrar subcontratistas; si no, principales"
        Dim coleccionAUsar As Collection
        If colNombresSub.count > 0 Then
            Set coleccionAUsar = colNombresSub
        Else
            Set coleccionAUsar = colNombresPrincipales
        End If

        If coleccionAUsar.count > 1 Then
            nombreFinalSuministrador = "UNION TEMPORAL DE EMPRESAS: " & JoinCollection(coleccionAUsar, " y ")
        ElseIf coleccionAUsar.count = 1 Then
            nombreFinalSuministrador = coleccionAUsar(1)
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

        .SuministradorNombreDir = nombreFinalSuministrador
        ' --- FIN DE LA LÓGICA RESTAURADA ---
    End With
    
    Set GetDatosPCIniciales = pc
    
LimpiarYSalir:
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCServicio.GetDatosPCIniciales"
    errObj.Raise
End Function


Public Function GuardarDatosGenerales(objPC As DatosPC) As Boolean
    ' DEVUELVE: True si se ejecutó una transición de estado (1->2).
    Dim solServ As New SolicitudServicio
    Dim wfServ As New WorkflowServicio
    Dim sol As Solicitud
    Dim transicionRealizada As Boolean
    
    ' Variables Transacción
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    
    On Error GoTo Errores
    transicionRealizada = False
    
    ' 1. Obtener solicitud (Lectura)
    Set sol = solServ.getSolicitudPorID(objPC.idSolicitud)
    If sol Is Nothing Then Err.Raise 513, "GuardarDatosGenerales", "No se encontró la solicitud."
    
    ' 2. Validaciones
    If Trim(Nz(objPC.refContratoInspeccionOficial, "")) = "" Then Err.Raise 513, "Validación", "El campo 'Referencia del Contrato/Inspección Oficial' es obligatorio."
    If Trim(Nz(objPC.refSuministrador, "")) = "" Then Err.Raise 513, "Validación", "El campo 'Referencia del Suministrador' es obligatorio."
    If Trim(Nz(objPC.denominacionContrato, "")) = "" Then Err.Raise 513, "Validación", "El campo 'Denominación del Contrato' es obligatorio."
    If Trim(Nz(objPC.SuministradorNombreDir, "")) = "" Then Err.Raise 513, "Validación", "El campo 'Suministrador (Nombre y Dirección)' es obligatorio."
    If Trim(Nz(objPC.objetoContrato, "")) = "" Then Err.Raise 513, "Validación", "El campo 'Objeto del Contrato' es obligatorio."
    
    ' 3. INICIO TRANSACCION
    Set ws = getWorkspace()
    Set db = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
    ws.BeginTrans
    
    ' 4. Persistir Datos (Pasando DB a la version compatible del repositorio)
    ' Nota: Usamos DatosPCRepositorio.Guardar que ya creamos/verificamos que acepta db
    Call DatosPCRepositorio.Guardar(objPC, db)
    
    ' 5. LOGICA DE TRANSICION AUTOMATICA (Si aplica)
    If sol.idEstadoInterno = estadoPreregistro Then
        ' IMPORTANTE: Pasamos la conexion 'db' a EjecutarTransicion para que se una a esta transaccion
        Call wfServ.EjecutarTransicion(sol, estadoRegistro, m_ObjUsuarioActivo, db)
        transicionRealizada = True
    End If
    
    ' 6. COMMIT
    ws.CommitTrans
    
    GuardarDatosGenerales = transicionRealizada
    
LimpiarYSalir:
    On Error Resume Next
    If Not db Is Nothing Then db.Close
    Set ws = Nothing
    Exit Function

Errores:
    If Not ws Is Nothing Then ws.Rollback
    GuardarDatosGenerales = False
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCServicio.GuardarDatosGenerales"
    errObj.Raise
End Function
Public Sub GuardarPropuesta(objPC As DatosPC, Optional db As DAO.Database)
' PROPOSITO: Valida y orquesta el guardado de la seccion de propuesta tecnica.
'-------------------------------------------------------------------------------
    On Error GoTo Errores
    
    Dim ws As DAO.Workspace
    Dim dbTrabajo As DAO.Database
    Dim gestionaTransaccion As Boolean
    Dim requiereTransaccion As Boolean
    
    If IsNull(objPC.descripcionMaterialAfectado) Or Trim(objPC.descripcionMaterialAfectado) = "" Then
        Err.Raise 513, "Regla de Negocio Incumplida", "La descripcion del material afectado es obligatoria."
    End If
    If Trim(objPC.descripcionPropuestaCambio) = "" Then
        Err.Raise 513, "Regla de Negocio Incumplida", "La descripcion de la propuesta de cambio obligatoria."
    End If

    ' --- TRAZABILIDAD RECHAZO (PREPARACION) ---
    ' Primero inicializar dbTrabajo para poder usarlo en las lecturas de seguimiento
    requiereTransaccion = (db Is Nothing And EsSolicitudEnValidacion(objPC.idSolicitud))

    If requiereTransaccion Then
        Set ws = getWorkspace()
        Set dbTrabajo = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
        gestionaTransaccion = True
    Else
        Set dbTrabajo = IIf(db Is Nothing, getdb(), db)
    End If

    Dim rechazoActivo As rechazo
    Dim objOriginal As DatosPC
    Set rechazoActivo = RechazoRepositorio.GetUltimoRechazoActivo(objPC.idSolicitud, dbTrabajo)

    If Not rechazoActivo Is Nothing Then
        ' IMPORTANTE: Leer el estado original ANTES de guardar los cambios
        Set objOriginal = Me.getDatosPC(objPC.idSolicitud, dbTrabajo)
    End If
    ' ------------------------------------------
    
    DatosPCRepositorio.ActualizarPropuesta objPC, dbTrabajo

    ' --- TRAZABILIDAD RECHAZO (REGISTRO) ---
    If Not rechazoActivo Is Nothing And Not objOriginal Is Nothing Then
        Dim jsonDict As Object
        Set jsonDict = JsonHelper.ObtenerJsonRechazo(objPC.idSolicitud)
        
        JsonHelper.RegistrarCambio jsonDict, "PC_Propuesta", "descripcionPropuestaCambio", objOriginal.descripcionPropuestaCambio, objPC.descripcionPropuestaCambio
        JsonHelper.RegistrarCambio jsonDict, "PC_Propuesta", "numPlanoEspecificacion", objOriginal.numPlanoEspecificacion, objPC.numPlanoEspecificacion
        JsonHelper.RegistrarCambio jsonDict, "PC_Propuesta", "descripcionMaterialAfectado", objOriginal.descripcionMaterialAfectado, objPC.descripcionMaterialAfectado
        
        JsonHelper.GuardarJsonRechazo objPC.idSolicitud, jsonDict
    End If
    ' ---------------------------------------

    If gestionaTransaccion Then
        ws.CommitTrans
    End If
    
LimpiarYSalir:
    If gestionaTransaccion And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Exit Sub

Errores:
    If gestionaTransaccion And Not (ws Is Nothing) Then ws.Rollback
    If gestionaTransaccion And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCServicio.GuardarPropuesta"
    errObj.Raise
End Sub

Public Sub GuardarImpacto(objPC As DatosPC, Optional db As DAO.Database)
' PROPOSITO: Validar y orquestar el guardado del bloque "Impacto".
' REGLAS: En Impacto todos son obligatorios (segun tu ultima indicacion).
' -----------------------------------------------------------------------
    On Error GoTo Errores
    
    Dim ws As DAO.Workspace
    Dim dbTrabajo As DAO.Database
    Dim gestionaTransaccion As Boolean
    Dim requiereTransaccion As Boolean

    ' 1) Motivos: al menos uno marcado; si "Otros" entonces detalle obligatorio
    Dim algunMotivo As Boolean
    algunMotivo = (objPC.motivoCorregirDeficiencias Or objPC.motivoMejorarCapacidad Or _
                   objPC.motivoAumentarNacionalizacion Or objPC.motivoMejorarSeguridad Or _
                   objPC.motivoMejorarFiabilidad Or objPC.motivoMejorarCosteEficacia Or _
                   objPC.motivoOtros)
    If Not algunMotivo Then
        Err.Raise 513, "Regla de Negocio Incumplida", _
                   ">> (Impacto) Seleccione al menos un Motivo de la Propuesta."
    End If
    If objPC.motivoOtros And Trim(Nz(objPC.motivoOtrosDetalle, "")) = "" Then
        Err.Raise 513, "Regla de Negocio Incumplida", _
                   ">> (Impacto) Debe indicar el detalle cuando marca 'Otros'."
    End If
    If objPC.motivoOtros = False And Trim(Nz(objPC.motivoOtrosDetalle, "")) <> "" Then
        Err.Raise 513, "Regla de Negocio Incumplida", _
                   ">> (Impacto) Si se indica otros motivos se ha de marcar 'Otros'"
    End If
    ' 2) Incidencias de coste y plazo (valores permitidos por la UI)
    Dim v As String
    v = UCase$(Trim$(Nz(objPC.incidenciaCoste, "")))
    If (v <> "AUMENTARÁ" And v <> "DISMINUIRÁ" And v <> "NO VARIARÁ") Then
        Err.Raise 513, "Regla de Negocio Incumplida", _
                   ">> (Impacto) Incidencia sobre el Coste es obligatoria."
    End If
    v = UCase$(Trim$(Nz(objPC.incidenciaPlazo, "")))
    If (v <> "AUMENTARÁ" And v <> "DISMINUIRÁ" And v <> "NO VARIARÁ") Then
        Err.Raise 513, "Regla de Negocio Incumplida", _
                   ">> (Impacto) Incidencia sobre el Plazo es obligatoria."
    End If

    ' 3) Campos de texto obligatorios nuevos
    v = UCase$(Trim$(Nz(objPC.impactoClasificacion, "")))  ' MAYOR | MENOR
    If (v <> "MAYOR" And v <> "MENOR") Then
        Err.Raise 513, "Regla de Negocio Incumplida", _
                   ">> (Impacto) Clasificacion (MAYOR/MENOR) es obligatoria."
    End If

    v = Trim$(Nz(objPC.CambioAfectaAMaterial, ""))          ' Material ya entregado | Material por entregar
    If (v <> "Material ya entregado" And v <> "Material por entregar") Then
        Err.Raise 513, "Regla de Negocio Incumplida", _
                   ">> (Impacto) Debe indicar si el cambio afecta a material ya entregado o por entregar."
    End If

    ' --- TRAZABILIDAD RECHAZO (INICIO) ---
    ' Primero inicializar dbTrabajo para consistencia transaccional
    requiereTransaccion = (db Is Nothing And EsSolicitudEnValidacion(objPC.idSolicitud))

    If requiereTransaccion Then
        Set ws = getWorkspace()
        Set dbTrabajo = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
        gestionaTransaccion = True
    Else
        Set dbTrabajo = IIf(db Is Nothing, getdb(), db)
    End If

    Dim rechazoActivo As rechazo
    Set rechazoActivo = RechazoRepositorio.GetUltimoRechazoActivo(objPC.idSolicitud, dbTrabajo)

    If Not rechazoActivo Is Nothing Then
        Dim objOriginal As DatosPC
        Set objOriginal = Me.getDatosPC(objPC.idSolicitud, dbTrabajo)

        If Not objOriginal Is Nothing Then
            Dim jsonDict As Object
            Set jsonDict = JsonHelper.ObtenerJsonRechazo(objPC.idSolicitud)

            ' Motivos (Booleanos)
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "motivoCorregirDeficiencias", objOriginal.motivoCorregirDeficiencias, objPC.motivoCorregirDeficiencias
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "motivoMejorarCapacidad", objOriginal.motivoMejorarCapacidad, objPC.motivoMejorarCapacidad
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "motivoAumentarNacionalizacion", objOriginal.motivoAumentarNacionalizacion, objPC.motivoAumentarNacionalizacion
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "motivoMejorarSeguridad", objOriginal.motivoMejorarSeguridad, objPC.motivoMejorarSeguridad
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "motivoMejorarFiabilidad", objOriginal.motivoMejorarFiabilidad, objPC.motivoMejorarFiabilidad
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "motivoMejorarCosteEficacia", objOriginal.motivoMejorarCosteEficacia, objPC.motivoMejorarCosteEficacia
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "motivoOtros", objOriginal.motivoOtros, objPC.motivoOtros
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "motivoOtrosDetalle", objOriginal.motivoOtrosDetalle, objPC.motivoOtrosDetalle

            ' Incidencias
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "incidenciaCoste", objOriginal.incidenciaCoste, objPC.incidenciaCoste
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "incidenciaPlazo", objOriginal.incidenciaPlazo, objPC.incidenciaPlazo
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "incidenciaSeguridad", objOriginal.incidenciaSeguridad, objPC.incidenciaSeguridad
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "incidenciaFiabilidad", objOriginal.incidenciaFiabilidad, objPC.incidenciaFiabilidad
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "incidenciaMantenibilidad", objOriginal.incidenciaMantenibilidad, objPC.incidenciaMantenibilidad
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "incidenciaIntercambiabilidad", objOriginal.incidenciaIntercambiabilidad, objPC.incidenciaIntercambiabilidad
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "incidenciaVidaUtilAlmacen", objOriginal.incidenciaVidaUtilAlmacen, objPC.incidenciaVidaUtilAlmacen
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "incidenciaFuncionamientoFuncion", objOriginal.incidenciaFuncionamientoFuncion, objPC.incidenciaFuncionamientoFuncion

            ' Clasificacion
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "impactoClasificacion", objOriginal.impactoClasificacion, objPC.impactoClasificacion
            JsonHelper.RegistrarCambio jsonDict, "PC_Impacto", "CambioAfectaAMaterial", objOriginal.CambioAfectaAMaterial, objPC.CambioAfectaAMaterial

            JsonHelper.GuardarJsonRechazo objPC.idSolicitud, jsonDict
        End If
    End If
    ' --- TRAZABILIDAD RECHAZO (FIN) ---
    
    DatosPCRepositorio.ActualizarImpacto objPC, dbTrabajo   ' OK con el repo actual refactorizado
    
    If gestionaTransaccion Then
        ws.CommitTrans
    End If
    
LimpiarYSalir:
    If gestionaTransaccion And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Exit Sub

Errores:
    If gestionaTransaccion And Not (ws Is Nothing) Then ws.Rollback
    If gestionaTransaccion And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCServicio.GuardarImpacto"
    errObj.Raise
End Sub
Public Sub GuardarAprobacionSuministrador(objPC As DatosPC, Optional db As DAO.Database)
' PROPOSITO: Validar y guardar "Aprobacion Suministrador" (firmas).
' OBLIGATORIOS: Todos (segun tu regla).
' -------------------------------------------------------------------
    On Error GoTo Errores

    If Trim(Nz(objPC.firmaOficinaTecnicaNombre, "")) = "" Then
        Err.Raise 513, "Regla Incumplida", _
                   ">> de Negocio (Aprobacion Suministrador) Nombre en firma de Oficina Tecnica."
    End If
    If Trim(Nz(objPC.firmaRepSuministradorNombre, "")) = "" Then
        Err.Raise 513, "Regla de Negocio Incumplida", _
                   ">> (Aprobacion Suministrador) Nombre en firma del Representante del Suministrador."
    End If

    If db Is Nothing Then Set db = getdb()
    DatosPCRepositorio.ActualizarAprobacionSuministrador objPC, db
    
    Exit Sub

Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCServicio.GuardarAprobacionSuministrador"
    errObj.Raise
End Sub

Public Sub GuardarDictamenRAC(ByRef objPC As DatosPC, Optional db As DAO.Database)
    On Error GoTo Errores
    
    Dim ws As DAO.Workspace
    Dim dbTrabajo As DAO.Database
    Dim gestionaTransaccion As Boolean
    Dim requiereTransaccion As Boolean
    
    If objPC.racDecision <> "RECHAZADO" Then
        If Trim(Nz(objPC.racCodigo, "")) = "" Then
            Err.Raise 513, "Validacion", "El Codigo RAC es obligatorio."
        End If
    End If
    
    requiereTransaccion = (db Is Nothing And EsSolicitudEnValidacion(objPC.idSolicitud))
    
    If requiereTransaccion Then
        Set ws = getWorkspace()
        Set dbTrabajo = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
        gestionaTransaccion = True
    Else
        Set dbTrabajo = db
    End If
    
    Call DatosPCRepositorio.ActualizarDictamenRAC(objPC, dbTrabajo)
    
    If gestionaTransaccion Then
        ws.CommitTrans
    End If
    
LimpiarYSalir:
    If gestionaTransaccion And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Exit Sub
Errores:
    If gestionaTransaccion And Not (ws Is Nothing) Then ws.Rollback
    If gestionaTransaccion And Not (dbTrabajo Is Nothing) Then dbTrabajo.Close
    Set ws = Nothing
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCServicio.GuardarDictamenRAC"
    errObj.Raise
End Sub

Public Sub GuardarDecisionFinal(ByRef objPC As DatosPC, Optional db As DAO.Database)
    On Error GoTo Errores
    
    If Trim(Nz(objPC.decisionFinal, "")) = "" Then Err.Raise 513, "Validacion", "Decision Final obligatoria."
    
    Call DatosPCRepositorio.ActualizarDecisionFinal(objPC, db)
    
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCServicio.GuardarDecisionFinal"
    errObj.Raise
End Sub

Public Sub EliminarDictamenRAC(ByVal idSolicitud As Long, Optional db As DAO.Database)
    Call DatosPCRepositorio.LimpiarDictamenRAC(idSolicitud, db)
End Sub

Public Sub EliminarAprobacionSuministrador(ByVal idSolicitud As Long, Optional db As DAO.Database)
    On Error GoTo Errores
    Call DatosPCRepositorio.LimpiarAprobacionSuministrador(idSolicitud, db)
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCServicio.EliminarAprobacionSuministrador"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud: errObj.Raise
End Sub

Public Sub EliminarDecisionFinal(ByVal idSolicitud As Long, Optional db As DAO.Database)
    On Error GoTo Errores
    Call DatosPCRepositorio.LimpiarDecisionFinal(idSolicitud, db)
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCServicio.EliminarDecisionFinal"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud: errObj.Raise
End Sub





' === RegistrarRechazoDesdeFormalizacion ===
Public Sub RegistrarRechazoDesdeFormalizacion(ByVal idSolicitud As Long, ByVal motivoRechazo As String, ByVal usuarioActivo As usuario, Optional ByRef db As DAO.Database)
    Dim ws As DAO.Workspace
    Dim dbTrabajo As DAO.Database
    Dim wfServ As New WorkflowServicio
    Dim pc As DatosPC
    Dim gestionaTransaccion As Boolean
    Dim requiereTransaccion As Boolean

    On Error GoTo Errores

    ' Dual-resolution: usar db entregada o crear una local
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

    ' 2. Obtener datos
    Set pc = Me.getDatosPC(idSolicitud, dbTrabajo)
    If pc Is Nothing Then Err.Raise 513, , "Datos no encontrados."

    ' 3. Limpieza de datos (En el objeto)
    ' A. Borrar datos de Formalizacion (Decision Final)
    pc.decisionFinal = ""
    pc.obsDecisionFinal = ""
    pc.NombreFirmanteFinal = ""

    ' B. Borrar datos de Aprobacion RAC
    pc.racCodigo = ""
    pc.racNombre = ""
    pc.observacionesRAC = ""

    ' C. Establecer el Rechazo
    pc.racDecision = "RECHAZADO"
    pc.racRechazoMotivos = motivoRechazo

    ' 4. Guardar Datos (Usando la conexion DB de la transaccion)
    Call DatosPCRepositorio.Guardar(pc, dbTrabajo)

    ' 5. Transicion de Estado
    Call wfServ.EjecutarTransicion(SolicitudRepositorio.getSolicitudPorID(idSolicitud, dbTrabajo), estadoRechazada, usuarioActivo, dbTrabajo)

    ' 6. Commit
    If requiereTransaccion Then ws.CommitTrans

LimpiarYSalir:
    On Error Resume Next
    If gestionaTransaccion And Not dbTrabajo Is Nothing Then dbTrabajo.Close
    Set ws = Nothing
    Set wfServ = Nothing
    Set pc = Nothing
    Exit Sub

Errores:
    If requiereTransaccion And Not ws Is Nothing Then ws.Rollback
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCServicio.RegistrarRechazoDesdeFormalizacion"
    errObj.Raise
End Sub

' ==========================================================================
' FUNCIONES DE SOPORTE Y VALIDACION
' ==========================================================================

Public Function getDatosPC(ByVal idSolicitud As Long, Optional db As DAO.Database) As DatosPC
    On Error GoTo Errores
    Set getDatosPC = DatosPCRepositorio.getPorIdSolicitud(idSolicitud, db)
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DatosPCServicio.getDatosPC"
    errObj.Raise
End Function


Public Function EsDatosGeneralesCompleta(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objPC As DatosPC
    On Error GoTo Errores
    
    EsDatosGeneralesCompleta = False
    ' PASAMOS DB
    Set objPC = Me.getDatosPC(idSolicitud, db)
    If objPC Is Nothing Then Exit Function
    
    
    If Trim(Nz(objPC.refContratoInspeccionOficial, "")) = "" Then Exit Function
    If Trim(Nz(objPC.refSuministrador, "")) = "" Then Exit Function
    If Trim(Nz(objPC.denominacionContrato, "")) = "" Then Exit Function
    If Trim(Nz(objPC.SuministradorNombreDir, "")) = "" Then Exit Function
    If Trim(Nz(objPC.objetoContrato, "")) = "" Then Exit Function
    
    EsDatosGeneralesCompleta = True
    
LimpiarYSalir:
    Exit Function
Errores:
    EsDatosGeneralesCompleta = False
    Resume LimpiarYSalir
End Function


Public Function EsParteTecnicaCompleta(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objPC As DatosPC
    On Error GoTo Errores
    
    EsParteTecnicaCompleta = False
    
    ' Leemos usando la conexion transaccional si existe
    Set objPC = Me.getDatosPC(idSolicitud, db)
    If objPC Is Nothing Then Exit Function
    
    ' 1. Validar Propuesta
    If Trim(Nz(objPC.descripcionMaterialAfectado, "")) = "" Then Exit Function
    If Trim(Nz(objPC.descripcionPropuestaCambio, "")) = "" Then Exit Function
    
    ' 2. Validar Impacto
    Dim coste As String, clasificacion As String
    coste = UCase(Trim(Nz(objPC.incidenciaCoste, "")))
    clasificacion = UCase(Trim(Nz(objPC.impactoClasificacion, "")))
    
    If coste <> "AUMENTARÁ" And coste <> "DISMINUIRÁ" And coste <> "NO VARIARÁ" Then Exit Function
    If clasificacion <> "MAYOR" And clasificacion <> "MENOR" Then Exit Function
    
    ' Validar Motivos (al menos uno)
    Dim algunMotivo As Boolean
    algunMotivo = (objPC.motivoCorregirDeficiencias Or objPC.motivoMejorarCapacidad Or _
                   objPC.motivoAumentarNacionalizacion Or objPC.motivoMejorarSeguridad Or _
                   objPC.motivoMejorarFiabilidad Or objPC.motivoMejorarCosteEficacia Or _
                   objPC.motivoOtros)
    If Not algunMotivo Then Exit Function
    
    EsParteTecnicaCompleta = True
    Exit Function

Errores:
    EsParteTecnicaCompleta = False
End Function

Public Function EsDetalleCompleto(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objPC As DatosPC
    On Error GoTo Errores
    
    EsDetalleCompleto = False
    Set objPC = Me.getDatosPC(idSolicitud, db)
    If objPC Is Nothing Then Exit Function
    
    If Trim(Nz(objPC.descripcionMaterialAfectado, "")) = "" Then Exit Function
    If Trim(Nz(objPC.descripcionPropuestaCambio, "")) = "" Then Exit Function
    If Trim(Nz(objPC.descripcionPropuestaCambioCont, "")) = "" Then Exit Function
    
    EsDetalleCompleto = True
    Exit Function

Errores:
    EsDetalleCompleto = False
End Function

Public Function EsMotivosCompleto(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objPC As DatosPC
    Dim coste As String, clasificacion As String
    Dim algumMotivo As Boolean
    
    On Error GoTo Errores
    
    EsMotivosCompleto = False
    Set objPC = Me.getDatosPC(idSolicitud, db)
    If objPC Is Nothing Then Exit Function
    
    coste = UCase(Trim(Nz(objPC.incidenciaCoste, "")))
    clasificacion = UCase(Trim(Nz(objPC.impactoClasificacion, "")))
    
    If coste <> "AUMENTARÁ" And coste <> "DISMINUIRÁ" And coste <> "NO VARIARÁ" Then Exit Function
    If clasificacion <> "MAYOR" And clasificacion <> "MENOR" Then Exit Function
    
    algumMotivo = (objPC.motivoCorregirDeficiencias Or objPC.motivoMejorarCapacidad Or _
                   objPC.motivoAumentarNacionalizacion Or objPC.motivoMejorarSeguridad Or _
                   objPC.motivoMejorarFiabilidad Or objPC.motivoMejorarCosteEficacia Or _
                   objPC.motivoOtros)
    If Not algumMotivo Then Exit Function
    
    EsMotivosCompleto = True
    
LimpiarYSalir:
    Exit Function
Errores:
    EsMotivosCompleto = False
    Resume LimpiarYSalir
End Function

Public Function EsDictamenRACCompleto(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objPC As DatosPC
    On Error GoTo Errores
    
    Set objPC = Me.getDatosPC(idSolicitud, db)
    If objPC Is Nothing Then Exit Function
    If Trim(Nz(objPC.racCodigo, "")) = "" Then Exit Function
    If Trim(Nz(objPC.racNombre, "")) = "" Then Exit Function
    If Trim(Nz(objPC.racDecision, "")) = "" Then Exit Function
    EsDictamenRACCompleto = True
    
LimpiarYSalir:
    Exit Function
Errores:
    EsDictamenRACCompleto = False
    Resume LimpiarYSalir
End Function

Public Function EsAprobacionSuministradorCompleta(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objPC As DatosPC
    On Error GoTo Errores
    
    Set objPC = Me.getDatosPC(idSolicitud, db)
    If objPC Is Nothing Then Exit Function
    
    If Trim(Nz(objPC.firmaOficinaTecnicaNombre, "")) = "" Then Exit Function
    If Trim(Nz(objPC.firmaRepSuministradorNombre, "")) = "" Then Exit Function
    If Trim(Nz(objPC.decisionFinal, "")) = "" Then Exit Function
    
    EsAprobacionSuministradorCompleta = True
    
LimpiarYSalir:
    Exit Function
Errores:
    EsAprobacionSuministradorCompleta = False
    Resume LimpiarYSalir
End Function


Public Function EsDecisionFinalCompleta(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim objPC As DatosPC
    On Error GoTo Errores
    
    Set objPC = Me.getDatosPC(idSolicitud, db)
    If objPC Is Nothing Then Exit Function
    If Not DecisionFinalHelper_EsCompleta(objPC.decisionFinal, objPC.NombreFirmanteFinal) Then Exit Function
    EsDecisionFinalCompleta = True
    
LimpiarYSalir:
    Exit Function
Errores:
    EsDecisionFinalCompleta = False
    Resume LimpiarYSalir
End Function

Public Function EsSolicitudEnValidacion(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As Boolean
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

Public Sub ImportarEntidad_aVM(ByRef pc As DatosPC, ByRef vm As DatosPCViewModel)
    ' --- GENERALES ---
    vm.vm_refContratoInspeccionOficial = Nz(pc.refContratoInspeccionOficial, "")
    vm.vm_refSuministrador = Nz(pc.refSuministrador, "")
    vm.vm_denominacionContrato = Nz(pc.denominacionContrato, "")
    vm.vm_SuministradorNombreDir = Nz(pc.SuministradorNombreDir, "")
    vm.vm_objetoContrato = Nz(pc.objetoContrato, "")
    
    ' --- PROPUESTA ---
    vm.vm_numPlanoEspecificacion = Nz(pc.numPlanoEspecificacion, "")
    vm.vm_descripcionMaterialAfectado = Nz(pc.descripcionMaterialAfectado, "")
    vm.vm_descripcionPropuestaCambio = Nz(pc.descripcionPropuestaCambio, "")
    
    ' --- IMPACTO (Motivos) ---
    vm.vm_motivoCorregirDeficiencias = pc.motivoCorregirDeficiencias
    vm.vm_motivoMejorarCapacidad = pc.motivoMejorarCapacidad
    vm.vm_motivoAumentarNacionalizacion = pc.motivoAumentarNacionalizacion
    vm.vm_motivoMejorarSeguridad = pc.motivoMejorarSeguridad
    vm.vm_motivoMejorarFiabilidad = pc.motivoMejorarFiabilidad
    vm.vm_motivoMejorarCosteEficacia = pc.motivoMejorarCosteEficacia
    vm.vm_motivoOtros = pc.motivoOtros
    vm.vm_motivoOtrosDetalle = Nz(pc.motivoOtrosDetalle, "")
    
    ' --- IMPACTO (Incidencias y Clasificación) ---
    vm.vm_incidenciaCoste = Nz(pc.incidenciaCoste, "")
    vm.vm_incidenciaPlazo = Nz(pc.incidenciaPlazo, "")
    vm.vm_incidenciaSeguridad = pc.incidenciaSeguridad
    vm.vm_incidenciaFiabilidad = pc.incidenciaFiabilidad
    vm.vm_incidenciaMantenibilidad = pc.incidenciaMantenibilidad
    vm.vm_incidenciaIntercambiabilidad = pc.incidenciaIntercambiabilidad
    vm.vm_incidenciaVidaUtilAlmacen = pc.incidenciaVidaUtilAlmacen
    vm.vm_incidenciaFuncionamientoFuncion = pc.incidenciaFuncionamientoFuncion
    vm.vm_impactoClasificacion = Nz(pc.impactoClasificacion, "")
    vm.vm_CambioAfectaAMaterial = Nz(pc.CambioAfectaAMaterial, "")
    
    ' --- Aprobación SUMINISTRADOR ---
    vm.vm_firmaOficinaTecnicaNombre = Nz(pc.firmaOficinaTecnicaNombre, "")
    vm.vm_firmaRepSuministradorNombre = Nz(pc.firmaRepSuministradorNombre, "")
    
    ' --- DICTAMEN RAC (NOMBRES NORMALIZADOS) ---
    vm.vm_racCodigo = Nz(pc.racCodigo, "")
    vm.vm_racDecision = Nz(pc.racDecision, "")
    vm.vm_observacionesRAC = Nz(pc.observacionesRAC, "")
    vm.vm_racNombre = Nz(pc.racNombre, "")
    
    ' --- Decisión FINAL ---
    vm.vm_obsAprobacionAutoridadDiseno = Nz(pc.obsAprobacionAutoridadDiseno, "")
    vm.vm_NombreAutoridadDiseno = Nz(pc.NombreAutoridadDiseno, "")
    vm.vm_decisionFinal = Nz(pc.decisionFinal, "")
    vm.vm_obsDecisionFinal = Nz(pc.obsDecisionFinal, "")
    vm.vm_NombreFirmanteFinal = Nz(pc.NombreFirmanteFinal, "")
End Sub

Public Sub ExportarVM_aEntidad(ByVal bloque As enumBloqueFormulario, ByRef vm As DatosPCViewModel, ByRef pc As DatosPC)
    Select Case bloque
        Case Bloque_Generales
            pc.refContratoInspeccionOficial = vm.vm_refContratoInspeccionOficial
            pc.refSuministrador = vm.vm_refSuministrador
            pc.denominacionContrato = vm.vm_denominacionContrato
            pc.SuministradorNombreDir = vm.vm_SuministradorNombreDir
            pc.objetoContrato = vm.vm_objetoContrato
        Case Bloque_Propuesta
            pc.numPlanoEspecificacion = vm.vm_numPlanoEspecificacion
            pc.descripcionMaterialAfectado = vm.vm_descripcionMaterialAfectado
            pc.descripcionPropuestaCambio = vm.vm_descripcionPropuestaCambio
        Case Bloque_Impacto
            pc.motivoCorregirDeficiencias = vm.vm_motivoCorregirDeficiencias
            pc.motivoMejorarCapacidad = vm.vm_motivoMejorarCapacidad
            pc.motivoAumentarNacionalizacion = vm.vm_motivoAumentarNacionalizacion
            pc.motivoMejorarSeguridad = vm.vm_motivoMejorarSeguridad
            pc.motivoMejorarFiabilidad = vm.vm_motivoMejorarFiabilidad
            pc.motivoMejorarCosteEficacia = vm.vm_motivoMejorarCosteEficacia
            pc.motivoOtros = vm.vm_motivoOtros
            pc.motivoOtrosDetalle = vm.vm_motivoOtrosDetalle
            pc.incidenciaCoste = vm.vm_incidenciaCoste
            pc.incidenciaPlazo = vm.vm_incidenciaPlazo
            pc.incidenciaSeguridad = vm.vm_incidenciaSeguridad
            pc.incidenciaFiabilidad = vm.vm_incidenciaFiabilidad
            pc.incidenciaMantenibilidad = vm.vm_incidenciaMantenibilidad
            pc.incidenciaIntercambiabilidad = vm.vm_incidenciaIntercambiabilidad
            pc.incidenciaVidaUtilAlmacen = vm.vm_incidenciaVidaUtilAlmacen
            pc.incidenciaFuncionamientoFuncion = vm.vm_incidenciaFuncionamientoFuncion
            pc.impactoClasificacion = vm.vm_impactoClasificacion
            pc.CambioAfectaAMaterial = vm.vm_CambioAfectaAMaterial
        Case Bloque_AprobacionSuministrador
            pc.firmaOficinaTecnicaNombre = vm.vm_firmaOficinaTecnicaNombre
            pc.firmaRepSuministradorNombre = vm.vm_firmaRepSuministradorNombre
        Case Bloque_DictamenRAC
             ' --- SECCIÓN RAC ACTUALIZADA ---
            pc.racCodigo = vm.vm_racCodigo
            pc.racDecision = vm.vm_racDecision
            pc.observacionesRAC = vm.vm_observacionesRAC
            pc.racNombre = vm.vm_racNombre
        Case Bloque_DecisionFinal
            pc.obsAprobacionAutoridadDiseno = vm.vm_obsAprobacionAutoridadDiseno
            pc.NombreAutoridadDiseno = vm.vm_NombreAutoridadDiseno
            pc.decisionFinal = vm.vm_decisionFinal
            pc.obsDecisionFinal = vm.vm_obsDecisionFinal
            pc.NombreFirmanteFinal = vm.vm_NombreFirmanteFinal
    End Select
End Sub

Public Function ObtenerViewModelCompleto(ByVal idSolicitud As Long, Optional db As DAO.Database) As DatosPCViewModel
    Dim vm As New DatosPCViewModel
    Dim pc As DatosPC
    Dim sol As Solicitud
    Dim exp As Expediente
    Dim expServ As New ExpedienteServicio
    Dim solServ As New SolicitudServicio
    Dim idEstado As Long
    Dim estadoActual As estado
    Dim colHistorial As Scripting.Dictionary
    Dim ultimaFechaTransicion As Date
    
    On Error GoTo Errores
    
    ' Cargar Solicitud
    Set sol = solServ.getSolicitudPorID(idSolicitud, db)
    If sol Is Nothing Then Err.Raise 513, "ObtenerViewModelCompleto", "No se encontró la solicitud."
    
    ' Cargar Expediente
    Set exp = expServ.getExpedientePorID(CStr(sol.idExpediente))
    If exp Is Nothing Then Err.Raise 513, "ObtenerViewModelCompleto", "No se encontró el expediente."
    
    ' Cargar DatosPC
    Set pc = Me.getDatosPC(idSolicitud, db)
    If pc Is Nothing Then
        Set pc = New DatosPC
        pc.idSolicitud = idSolicitud
        pc.refContratoInspeccionOficial = exp.CodExpLargo
        pc.refSuministrador = sol.codigoSolicitud
    End If
    
    ' Armar VM
    vm.idSolicitud = idSolicitud
    Set vm.datos = pc
    Set vm.Solicitud = sol
    Set vm.Expediente = exp
    
    ' Estado y Tiempo Ciclo
    idEstado = sol.idEstadoInterno
    If Not m_ObjEntorno Is Nothing Then
        If m_ObjEntorno.estados.Exists(CStr(idEstado)) Then
            Set estadoActual = m_ObjEntorno.estados(CStr(idEstado))
            vm.EstadoNombre = estadoActual.nombreEstado
            vm.EstadoDescripcion = estadoActual.descripcion
        Else
            vm.EstadoNombre = "Desconocido (" & idEstado & ")"
        End If
    Else
        vm.EstadoNombre = "Estado ID: " & idEstado
    End If
    
    Set colHistorial = solServ.getHistorialDeEstados(idSolicitud)
    ultimaFechaTransicion = sol.fechaCreacion
    If colHistorial.count > 0 Then
        Dim ultimoLog As LogEstado
        Set ultimoLog = colHistorial.items(colHistorial.count - 1)
        ultimaFechaTransicion = ultimoLog.fechaTransicion
    End If
    If Not estadoActual Is Nothing Then
        If estadoActual.esEstadoFinal Then
            vm.TiempoCiclo = CalcularDuracion(sol.fechaCreacion, ultimaFechaTransicion)
        Else
            vm.TiempoCiclo = CalcularDuracion(sol.fechaCreacion, Now()) & " (en curso)"
        End If
    Else
        vm.TiempoCiclo = CalcularDuracion(sol.fechaCreacion, Now()) & " (en curso)"
    End If
    
    ' Mapear entidad a VM
    Call ImportarEntidad_aVM(pc, vm)
    
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
    errObj.Create Err.Number, Err.description, "DatosPCServicio.ObtenerViewModelCompleto"
    errObj.AddToCallStack "idSolicitud: " & idSolicitud
    errObj.Raise
End Function




