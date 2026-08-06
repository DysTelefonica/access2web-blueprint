VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolicitudServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

' =============================================================================
' TEST SEAM: NC Cleanup Failure Simulation (ESC Hardening Step 1)
' Minimal internal flag to enable testing of non-blocking NC cleanup behavior.
' =============================================================================
Private m_ForceNCCleanupFailure As Boolean    ' When True, simulates NC cleanup failure
Private m_NCCleanupFailureRaised As Boolean   ' Tracks if NC cleanup failure was raised

' -----------------------------------------------------------------------------
' TEST HOOK: Enable/disable NC cleanup failure simulation
' Only for use in ESC canonical tests - never call in production code.
' -----------------------------------------------------------------------------
Public Sub SetForceNCCleanupFailure(ByVal forceFailure As Boolean)
    m_ForceNCCleanupFailure = forceFailure
    m_NCCleanupFailureRaised = False
End Sub

' -----------------------------------------------------------------------------
' TEST HOOK: Check if NC cleanup failure was raised during last call
' Returns True if the NC cleanup failure path was triggered.
' Only for use in ESC canonical tests.
' -----------------------------------------------------------------------------
Public Function GetNCCleanupFailureRaised() As Boolean
    GetNCCleanupFailureRaised = m_NCCleanupFailureRaised
End Function



Public Sub GuardarNuevaSolicitud(ByVal idExpediente As Long, ByVal tipoSolicitud As String, _
                               ByVal codigoSolicitud As String, ByVal idNC As Long, _
                               Optional ByRef db As DAO.Database)
    Dim nuevaSol As Solicitud
    On Error GoTo Errores
    If db Is Nothing Then Set db = getdb()
    
    Set nuevaSol = CrearNuevaSolicitud(idExpediente, tipoSolicitud, codigoSolicitud, idNC, db)
    Call Validar(nuevaSol)
    Call SolicitudRepositorio.GuardarSolicitud(nuevaSol, db)
    
LimpiarYSalir:
    On Error Resume Next
    Set nuevaSol = Nothing
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudServicio.GuardarNuevaSolicitud"
    errObj.Raise
End Sub

' ==========================================================================
' MÓDULO: SolicitudServicio.cls
' ACCIÓN: REEMPLAZO DEFINITIVO de getSolicitudesViewModel con una consulta SQL corregida.
' ==========================================================================
Public Function getSolicitudesViewModel(Optional ByVal palabraClave As String, _
                                      Optional ByRef filtros As FiltrosSolicitud, _
                                      Optional ByVal soloMisSolicitudes As Boolean = False, _
                                      Optional ByVal soloPendientes As Boolean = False, _
                                      Optional ByRef db As DAO.Database) As Scripting.Dictionary

    Dim sql As String, whereClause As String
    Dim wheres As Object, params As Object
    On Error GoTo Errores
    
    If db Is Nothing Then Set db = getdb()
    
    Set wheres = CreateObject("Scripting.Dictionary")
    Set params = CreateObject("Scripting.Dictionary")
    
    ' --- INICIO DE LA CONSULTA CORREGIDA Y SIMPLIFICADA ---
    
    ' Subconsulta para obtener el Jefe de Proyecto (Responsable Técnico)
    Dim subQueryRT As String
    subQueryRT = "(SELECT T_Resp_RT.IdExpediente, FIRST(T_Usr_RT.Nombre) AS NombreRT, FIRST(T_Usr_RT.CorreoUsuario) AS EmailRT " & _
                 "FROM TbExpedientesResponsables AS T_Resp_RT INNER JOIN TbUsuariosAplicaciones AS T_Usr_RT ON T_Resp_RT.IdUsuario = T_Usr_RT.Id " & _
                 "WHERE T_Resp_RT.EsJefeProyecto = 'SÍ' " & _
                 "GROUP BY T_Resp_RT.IdExpediente) AS SubRT"

    ' Construcción de la consulta principal
    sql = "SELECT " & _
            "S.idSolicitud, " & _
            "IIf(S.tipoSolicitud='PC', 'PC - Propuesta de Cambio', IIf(S.tipoSolicitud='CD_CA', 'CD/CA - Concesión/Desviación', IIf(S.tipoSolicitud='CD_CA_SUB', 'CD/CA-SUB - Subcontratista', IIf(S.tipoSolicitud='PC_SUB', 'PC-SUB - Subcontratista', S.tipoSolicitud)))) AS TipoSolicitud, " & _
            "S.codigoSolicitud, " & _
            "E.nombreEstado AS EstadoNombre, " & _
            "T_Exp.IDExpediente, T_Exp.CodExp, T_Exp.Nemotecnico, T_Exp.FechaInicioContrato, T_Exp.FechaFinContrato, " & _
            "SubRT.NombreRT AS ResponsableTecnico, SubRT.EmailRT AS EmailResponsableTecnico, " & _
            "T_RC.Nombre AS ResponsableCalidad, T_RC.CorreoUsuario AS EmailResponsableCalidad, " & _
            "'' AS RacAsignado, " & _
            "S.fechaCreacion " & _
          "FROM " & _
            "(((tbSolicitudes AS S " & _
            "INNER JOIN tbEstados AS E ON S.idEstadoInterno = E.idEstado) " & _
            "INNER JOIN TbExpedientes AS T_Exp ON S.idExpediente = T_Exp.IDExpediente) " & _
            "LEFT JOIN TbUsuariosAplicaciones AS T_RC ON T_Exp.IDResponsableCalidad = T_RC.Id) " & _
            "LEFT JOIN " & subQueryRT & " ON T_Exp.IDExpediente = SubRT.IdExpediente"
    
    If soloPendientes Then
        If rolUsuario = rol.Tecnico Then
            ' Para el Técnico, sus tareas son las que están en DESARROLLO TÉCNICO
            wheres.Add "S.idEstadoInterno = " & estadoDesarrolloTecnico, ""
        Else ' Para Calidad / Admin
            ' Sus tareas son TODAS las que NO están en DESARROLLO TÉCNICO (del técnico)
            ' y que NO están en un estado final (APROBADA).
            wheres.Add "S.idEstadoInterno NOT IN (" & estadoDesarrolloTecnico & "," & estadoAprobada & ")", ""
        End If
    ElseIf Not filtros Is Nothing Then
        If Nz(filtros.idEstado, "") <> "" And filtros.idEstado <> "0" Then
            wheres.Add "S.idEstadoInterno IN (" & filtros.idEstado & ")", ""
        End If
    End If
    If soloMisSolicitudes Then
        If rolUsuario = rol.Tecnico Then
            wheres.Add "EXISTS (SELECT 1 FROM TbExpedientesResponsables RT WHERE RT.IdExpediente = T_Exp.IDExpediente AND RT.IdUsuario = [p_IDUsuarioActivo])", ""
            params.Add "p_IDUsuarioActivo", m_ObjUsuarioActivo.id
        Else
            wheres.Add "T_Exp.IDResponsableCalidad = [p_IDUsuarioActivo]", ""
            params.Add "p_IDUsuarioActivo", m_ObjUsuarioActivo.id
        End If
    End If
    
    If Nz(palabraClave, "") <> "" Then
        wheres.Add "(S.codigoSolicitud LIKE [p_PalabraClave] OR T_Exp.Nemotecnico LIKE [p_PalabraClave] OR T_Exp.CodExp LIKE [p_PalabraClave])", ""
        params.Add "p_PalabraClave", "*" & palabraClave & "*"
    End If
    
    If Not filtros Is Nothing Then
        If filtros.idResponsableTecnico > 0 And Not soloMisSolicitudes Then
            wheres.Add "EXISTS (SELECT 1 FROM TbExpedientesResponsables RT WHERE RT.IdExpediente = T_Exp.IDExpediente AND RT.IdUsuario = [p_IDTecnico])", ""
            params.Add "p_IDTecnico", filtros.idResponsableTecnico
        End If
        
        If filtros.idResponsableCalidad > 0 And Not soloMisSolicitudes Then
            wheres.Add "T_Exp.IDResponsableCalidad = [p_IDCalidad]", ""
            params.Add "p_IDCalidad", filtros.idResponsableCalidad
        End If

        If filtros.idSuministrador > 0 Then
             wheres.Add "EXISTS (SELECT 1 FROM TbExpedientesSuministradores TES WHERE TES.IDExpediente = T_Exp.IDExpediente AND TES.IDSuministrador = [p_IDSuministrador])", ""
             params.Add "p_IDSuministrador", filtros.idSuministrador
        End If
        
        If Trim(Nz(filtros.tipoSolicitud, "")) <> "" Then
            wheres.Add "S.tipoSolicitud = [p_TipoSol]", ""
            params.Add "p_TipoSol", filtros.tipoSolicitud
        End If
    End If
    
    If wheres.count > 0 Then whereClause = " WHERE " & Join(wheres.Keys, " AND ")
    
    sql = sql & whereClause & " ORDER BY S.fechaCreacion DESC;"
    
    Set getSolicitudesViewModel = RepositorioComun.HidratarColeccionDesdeSQL(sql, "SolicitudBusquedaViewModel", "idSolicitud", db, params)

LimpiarYSalir:
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudServicio.getSolicitudesViewModel"
    errObj.Raise
End Function
Public Function getSolicitudViewModelPorID(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As SolicitudViewModel
    Dim vm As New SolicitudViewModel
    Dim sol As Solicitud
    Dim expServ As New ExpedienteServicio
    Dim estado As estado
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set db = getdb()
    
    ' 1. Obtener la solicitud
    Set sol = Me.getSolicitudPorID(idSolicitud, db)
    If sol Is Nothing Then Exit Function
    Set vm.Solicitud = sol
    
    ' 2. Obtener el expediente asociado
    Set vm.Expediente = expServ.getExpedientePorID(CStr(sol.idExpediente))
    If vm.Expediente Is Nothing Then Exit Function
    
    ' 3. Poblar las propiedades aplanadas del ViewModel
    Dim dictNombres As Scripting.Dictionary
    Set dictNombres = expServ.getNombresSuministradores(sol.idExpediente)
    vm.nombreContratistaPrincipal = dictNombres("principal")
    vm.nombreSubcontratista = dictNombres("subcontratista")
    vm.JefeProyecto = Nz(vm.Expediente.ResponsableTecnico, "N/A")
    vm.JefeProyectoEmail = Nz(vm.Expediente.EmailResponsableTecnico, "")
    vm.TecnicoCalidad = Nz(vm.Expediente.ResponsableCalidad, "N/A")
    vm.TecnicoCalidadEmail = Nz(vm.Expediente.EmailResponsableCalidad, "")
    
    ' --- INICIO DE LA MODIFICACIÓN ---
    ' 4. Traducir el tipo de solicitud y guardarlo en el ViewModel
    vm.TipoSolicitudNombre = getNombreAmigableTipoSolicitud(sol.tipoSolicitud)
    ' --- FIN DE LA MODIFICACIÓN ---

    ' 5. Obtener el nombre del estado
    If m_ObjEntorno.estados.Exists(CStr(sol.idEstadoInterno)) Then
        Set estado = m_ObjEntorno.estados(CStr(sol.idEstadoInterno))
        vm.EstadoNombre = estado.nombreEstado
        vm.EstadoDescripcion = estado.descripcion
        
        ' --- SPEC-011a: Override de descripción para mensajes contextuales ---
        If sol.idEstadoInterno = estadoModificacion And sol.revisionCalidadEstado = "RECHAZADO" Then
             If rolUsuario = rol.Tecnico Then
                 vm.EstadoDescripcion = "Cambios realizados pendiente decisión calidad"
             ElseIf rolUsuario = rol.Calidad Or rolUsuario = rol.Administrador Then
                 vm.EstadoDescripcion = "Cambios del técnico realizado"
             End If
        End If
        ' ---------------------------------------------------------------------
    Else
        vm.EstadoNombre = "Desconocido"
        vm.EstadoDescripcion = "ID " & sol.idEstadoInterno
    End If
     ' --- INICIO DE LA NUEVA LÓGICA ---
    ' 5. Obtener el Código de la No Conformidad, si existe.
    vm.CodigoNCAsociada = "N/A" ' Valor por defecto
    If sol.idNCAsociada > 0 Then
        Dim ncServ As New NoConformidadServicio
        Dim colNCs As Scripting.Dictionary
        Set colNCs = ncServ.getNoConformidades() ' Obtenemos todas (usamos caché si está implementada)
        
        If Not colNCs Is Nothing Then
            If colNCs.Exists(CStr(sol.idNCAsociada)) Then
                Dim nc As NoConformidad
                Set nc = colNCs(CStr(sol.idNCAsociada))
                vm.CodigoNCAsociada = nc.CodigoNoConformidad
            Else
                vm.CodigoNCAsociada = "ID: " & sol.idNCAsociada & " (No encontrada)"
            End If
        End If
    End If
    ' --- FIN DE LA NUEVA LÓGICA ---
    ' 6. Calcular Tiempo de Ciclo
    Dim colHistorial As Scripting.Dictionary
    Dim ultimaFechaTransicion As Date
    Set colHistorial = Me.getHistorialDeEstados(idSolicitud)
    ultimaFechaTransicion = sol.fechaCreacion
    If colHistorial.count > 0 Then
        Dim ultimoLog As LogEstado
        Set ultimoLog = colHistorial.items(colHistorial.count - 1)
        ultimaFechaTransicion = ultimoLog.fechaTransicion
    End If
    If Not estado Is Nothing Then
        If estado.esEstadoFinal Then
            vm.TiempoCiclo = CalcularDuracion(sol.fechaCreacion, ultimaFechaTransicion)
        Else
            vm.TiempoCiclo = CalcularDuracion(sol.fechaCreacion, Now()) & " (en curso)"
        End If
    Else
        vm.TiempoCiclo = CalcularDuracion(sol.fechaCreacion, Now()) & " (en curso)"
    End If
    vm.vm_revisionCalidadEstado = sol.revisionCalidadEstado
    vm.vm_revisionCalidadComentarios = sol.revisionCalidadComentarios
    
    ' --- SPEC-099: Inicializacion de propiedades para notificaciones inteligentes ---
    ' NOTA: Estas propiedades requieren evolucion en el modelo de datos (TbSolicitudes).
    ' Por ahora se inicializan con valores por defecto para asegurar la estabilidad del motor.
    
    ' Lógica de Subsanación (Deducida)
    ' Si el estado de calidad sigue constando como RECHAZADO pero la solicitud ha avanzado a Revisión/Validación,
    ' asumimos que es una subsanación recién enviada.
    If (sol.revisionCalidadEstado = "RECHAZADO") And (sol.idEstadoInterno = enumEstados.estadoRevision Or sol.idEstadoInterno = enumEstados.estadoAprobada) Then
        vm.vm_tecnicoSubsanacionRealizada = True
        vm.vm_tecnicoComentariosSubsanacion = "Se han aplicado correcciones tras el rechazo de calidad."
    Else
        vm.vm_tecnicoSubsanacionRealizada = False
        vm.vm_tecnicoComentariosSubsanacion = ""
    End If
    
    ' Lógica RAC (Validación/Revisión)
    Dim racEstado As String
    Dim racComentarios As String
    ' Usamos el repositorio para obtener el último estado real
    Call ValidacionRevisionRepositorio.GetUltimoEstadoValidacion(idSolicitud, racEstado, racComentarios)
    
    vm.vm_racValidacionEstado = racEstado
    vm.vm_racValidacionComentarios = racComentarios
    ' --------------------------------------------------------------------------------

    Set getSolicitudViewModelPorID = vm

LimpiarYSalir:
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudServicio.getSolicitudViewModelPorID"
    errObj.AddToCallStack "ID Solicitud: " & idSolicitud
    errObj.Raise
End Function

Public Function getSolicitudPorID(ByVal id As Long, Optional ByRef db As DAO.Database) As Solicitud
    On Error GoTo Errores
    If id <= 0 Then Err.Raise 513, "getSolicitudPorID", "ID de solicitud no válido."
    Dim dbConexion As DAO.Database
    If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db
    Set getSolicitudPorID = SolicitudRepositorio.getSolicitudPorID(id, dbConexion)
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudServicio.getSolicitudPorID"
    errObj.Raise
End Function

Public Function CrearNuevaSolicitud(ByVal idExpediente As Long, ByVal tipoSolicitud As String, _
                                  ByVal codigoSolicitud As String, ByVal idNC As Long, _
                                  Optional ByRef db As DAO.Database) As Solicitud
    Dim sol As New Solicitud
    On Error GoTo Errores
    If db Is Nothing Then Set db = getdb()
    With sol
        .idSolicitud = DameID("tbSolicitudes", "idSolicitud", db)
        .idExpediente = idExpediente
        .tipoSolicitud = tipoSolicitud
        .codigoSolicitud = codigoSolicitud
        .idEstadoInterno = estadoPreregistro
        .fechaCreacion = Now()
        .usuarioCreacion = m_ObjUsuarioActivo.nombre
        .fechaModificacion = .fechaCreacion
        .usuarioModificacion = .usuarioCreacion
        .idNCAsociada = idNC ' <<-- AÑADIDO
         ' --- INICIO DE LA CORRECCIÓN ---
        ' Inicializamos el estado de revisión para evitar error de "longitud cero"
        .revisionCalidadEstado = "PENDIENTE"
        .revisionCalidadComentarios = ""
        ' --- FIN DE LA CORRECCIÓN ---
    End With
    Set CrearNuevaSolicitud = sol
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudServicio.CrearNuevaSolicitud"
    errObj.Raise
End Function
' EN: SolicitudServicio.cls

Public Sub Validar(ByRef entidad As Solicitud, Optional ByVal esEdicion As Boolean = False, Optional ByRef db As DAO.Database)
    On Error GoTo Errores
    
    If entidad Is Nothing Then Err.Raise 513, "Validación", "La entidad solicitud es nula."
    
    ' Validaciones básicas
    If entidad.idExpediente <= 0 Then Err.Raise 513, "Validación", "• Debe seleccionar un expediente."
    If Trim(entidad.tipoSolicitud) = "" Then Err.Raise 513, "Validación", "• Debe seleccionar un tipo de solicitud."
    If Trim(entidad.codigoSolicitud) = "" Then Err.Raise 513, "Validación", "• El código de solicitud no puede estar vacío."
    
    ' Validación de revisionCalidadEstado - debe ser PENDIENTE, APROBADO, RECHAZADO o vacío
    If Trim(entidad.revisionCalidadEstado) <> "" Then
        Select Case UCase(Trim(entidad.revisionCalidadEstado))
            Case "PENDIENTE", "APROBADO", "RECHAZADO"
                ' Válido
            Case Else
                Err.Raise 513, "Validación", "• El estado de revisión de calidad debe ser PENDIENTE, APROBADO o RECHAZADO."
        End Select
    End If
    
    ' --- SDD tipo-solicitud-pc-subcontratista (Spec-001) ---
    ' Defensa de dominio tipoSolicitud antes del chequeo de duplicados.
    ' Canonical set: PC, CD_CA, CD_CA_SUB, PC_SUB.
    ' Legacy aliases PCSUB / CDCASUB se aceptan por evidencia (task 1.0 audit):
    '   - SolicitudServicio.getNombreAmigableTipoSolicitud ya los acepta
    '   - SnapshotHelper los usa como claves de mapeo (3 sitios)
    ' Cualquier otro valor (incluso en mayúsculas) se rechaza con Err 513.
    Select Case UCase$(Trim$(entidad.tipoSolicitud))
        Case "PC", "CD_CA", "CD_CA_SUB", "PC_SUB", "PCSUB", "CDCASUB"
            ' Válido
        Case Else
            Err.Raise 513, "Validación", "• Tipo de solicitud no válido: '" & entidad.tipoSolicitud & "'. Valores permitidos: PC, CD_CA, CD_CA_SUB, PC_SUB."
    End Select
    ' --- FIN SDD tipo-solicitud-pc-subcontratista ---

    ' Validación de Duplicados
    If esEdicion Then
        Dim solOriginal As Solicitud
        Set solOriginal = SolicitudRepositorio.getSolicitudPorCodigo(entidad.codigoSolicitud, db)
        
        If Not solOriginal Is Nothing Then
            If solOriginal.idSolicitud <> entidad.idSolicitud Then
                Err.Raise 513, "Validación", "• El código '" & entidad.codigoSolicitud & "' ya pertenece a otra solicitud."
            End If
        End If
    Else
        If SolicitudRepositorio.ExisteCodigo(entidad.codigoSolicitud, db) Then
            Err.Raise 513, "Validación", "• El código '" & entidad.codigoSolicitud & "' ya existe en el sistema."
        End If
    End If

    Exit Sub
    
Errores:
    ' CORRECCIÓN: Usamos CondorError para transportar la validación hasta la UI
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudServicio.Validar"
    If Not entidad Is Nothing Then
        errObj.AddToCallStack "CodigoSolicitud: " & entidad.codigoSolicitud
        errObj.AddToCallStack "ID: " & entidad.idSolicitud
    End If
    errObj.Raise
End Sub
' --- INICIO DEL NUEVO MÉTODO ---
Public Function ProponerCodigoSolicitud(ByVal idExpediente As Long) As String
    Dim expServ As New ExpedienteServicio
    Dim exp As Expediente
    Dim claveExp As String
    Dim siguienteOrdinal As Long
    On Error GoTo Errores
    
    If idExpediente <= 0 Then
        ProponerCodigoSolicitud = ""
        Exit Function
    End If
    
    Set exp = expServ.getExpedientePorID(CStr(idExpediente))
    If exp Is Nothing Then
        Err.Raise 513, "ProponerCodigo", "No se encontró el expediente con ID " & idExpediente
    End If
    
    claveExp = expServ.getClaveParaCodigoSolicitud(exp)
    siguienteOrdinal = SolicitudRepositorio.getSiguienteOrdinalParaClave(claveExp)
    
    ProponerCodigoSolicitud = "DC-" & claveExp & "-" & siguienteOrdinal
    
LimpiarYSalir:
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudServicio.ProponerCodigoSolicitud"
    errObj.Raise
End Function
Public Function getHistorialDeEstados(ByVal idSolicitud As Long) As Scripting.Dictionary
    ' Responsabilidad: Orquestar la obtención del historial de cambios
    ' de estado para una solicitud específica.
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then
        Err.Raise 513, "Validación", "El ID de la Solicitud para obtener el historial no es válido."
    End If
    
    ' Delegamos la llamada directamente al repositorio especializado.
    Set getHistorialDeEstados = LogEstadoRepositorio.getHistorialPorIdSolicitud(idSolicitud)
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudServicio.getHistorialDeEstados"
    
    errObj.Raise
End Function


' ==========================================================================
' MÓDULO: SolicitudServicio.cls (ACTUALIZADO)
' ACCIÓN: Permitir que el rol Calidad también pueda eliminar solicitudes.
' ==========================================================================
' ==========================================================================
' MÓDULO: SolicitudServicio.cls
' ACCIÓN: REEMPLAZO de EliminarSolicitudCompleta para que sea polimórfico.
' ==========================================================================
Public Sub EliminarSolicitudCompleta(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database)
    Dim ws As DAO.Workspace
    Dim dbTransaccional As DAO.Database
    Dim dbCreadaLocalmente As Boolean
    Dim adjServ As New AdjuntosServicio
    Dim colAdjuntos As Collection
    Dim adj As Adjunto
    Dim fso As Object
    Dim archivosAEliminar As New Collection
    Dim sol As Solicitud ' <<< AÑADIDO para consultar el tipo
    Dim notifServ As New NotificacionServicio
    Dim vm As SolicitudViewModel
    Dim destinatario As String, copia As String, asunto As String
    
    On Error GoTo Errores
    
    dbCreadaLocalmente = (db Is Nothing)
    If dbCreadaLocalmente Then Set db = getdb()
    
    ' --- REGLA DE NEGOCIO DE PERMISOS (sin cambios) ---
    If rolUsuario <> rol.Administrador And rolUsuario <> rol.Calidad Then
        Err.Raise 513, "EliminarSolicitud", "No tiene los permisos necesarios (Administrador o Calidad) para eliminar una solicitud."
    End If
    
    ' --- OBTENER SOLICITUD PARA DETERMINAR EL TIPO ---
    Call modBattery_Canonical.Canonical_Log_Trace("[TRACE][EliminarSolicitudCompleta] Inicio. idSolicitud=" & idSolicitud & ", dbCreadaLocalmente=" & dbCreadaLocalmente)
    Set sol = Me.getSolicitudPorID(idSolicitud, db)
    Call modBattery_Canonical.Canonical_Log_Trace("[TRACE][EliminarSolicitudCompleta] getSolicitudPorID completado. sol Is Nothing=" & (sol Is Nothing))
    If sol Is Nothing Then
        Err.Raise 513, "EliminarSolicitud", "La solicitud con ID " & idSolicitud & " no existe."
    End If
    Set vm = Me.getSolicitudViewModelPorID(idSolicitud, db)
    Call modBattery_Canonical.Canonical_Log_Trace("[TRACE][EliminarSolicitudCompleta] getSolicitudViewModelPorID completado. vm Is Nothing=" & (vm Is Nothing))
    If Not vm Is Nothing Then
        Call modBattery_Canonical.Canonical_Log_Trace("[TRACE][EliminarSolicitudCompleta] vm.Solicitud Is Nothing=" & (vm.Solicitud Is Nothing))
        Call modBattery_Canonical.Canonical_Log_Trace("[TRACE][EliminarSolicitudCompleta] vm.Expediente Is Nothing=" & (vm.Expediente Is Nothing))
    End If
    
    ' --- SPEC-101: Notificación PRE-Borrado ---
    ' Requerimiento: Enviar aviso antes de eliminar los datos.
    Call modBattery_Canonical.Canonical_Log_Trace("[TRACE][EliminarSolicitudCompleta] Preparando datos de notificación")
    
    ' Guardia: si vm o vm.Expediente es Nothing, no se puede preparar notificación
    If vm Is Nothing Or vm.Expediente Is Nothing Then
        Err.Raise 513, "EliminarSolicitudCompleta", _
            "No se pudo obtener el expediente asociado a la solicitud " & idSolicitud & ". La eliminación no puede continuar."
    End If
    
    destinatario = vm.Expediente.EmailResponsableCalidad
    copia = vm.Expediente.EmailResponsableTecnico
    asunto = "CONDOR - ELIMINADO | " & vm.Expediente.Nemotecnico & " | " & vm.TipoSolicitudNombre & " | " & vm.Solicitud.codigoSolicitud
    Call modBattery_Canonical.Canonical_Log_Trace("[TRACE][EliminarSolicitudCompleta] Notificación preparada. destinatario='" & destinatario & "', copia='" & copia & "'")
    
    On Error Resume Next
    Call notifServ.EnviarNotificacionEliminacion(vm, asunto, destinatario, copia)
    If Err.Number <> 0 Then Call modBattery_Canonical.Canonical_Log_Trace("[TRACE][EliminarSolicitudCompleta] EnviarNotificacionEliminacion falló: " & Err.Number & " - " & Err.description)
    On Error GoTo Errores
    ' -------------------------------------------
    
    ' --- PREPARACIÓN Y VALIDACIÓN DE PRECONDICIONES (sin cambios) ---
    Set colAdjuntos = adjServ.getAdjuntosPorSolicitud(idSolicitud)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not colAdjuntos Is Nothing Then
        For Each adj In colAdjuntos
            Dim rutaCompleta As String
            rutaCompleta = m_ObjEntorno.URLDirectorioDocumentacion & adj.nombreArchivo
            If fso.FileExists(rutaCompleta) Then
                If FicheroAbierto(rutaCompleta) Then
                    Err.Raise 513, "EliminarSolicitud", "No se puede continuar. El archivo adjunto '" & adj.nombreArchivo & "' está actualmente en uso."
                End If
                archivosAEliminar.Add rutaCompleta
            End If
        Next adj
    End If

    ' --- INICIO DE LA TRANSACCIÓN (sin cambios) ---
    Set ws = getWorkspace()
    Set dbTransaccional = db
    ws.BeginTrans
    
    ' --- LIMPIEZA DE INTEGRIDAD REFERENCIAL (BARRIDO COMPLETO) ---
    ' Ejecutamos el borrado en TODAS las tablas de datos técnicos.
    ' Esto asegura que no queden registros huérfanos (aunque se haya cambiado el tipo de solicitud)
    ' y evita el error de bloqueo por relaciones.
    
    Call DatosPCRepositorio.EliminarPorIdSolicitud(idSolicitud, dbTransaccional)
    Call DatosCDCARepositorio.EliminarPorIdSolicitud(idSolicitud, dbTransaccional)
    Call DatosCDCASUBRepositorio.EliminarPorIdSolicitud(idSolicitud, dbTransaccional)
    
    ' Borrado de tablas comunes
    Call AdjuntoRepositorio.EliminarPorIdSolicitud(idSolicitud, dbTransaccional)
    Call LogEstadoRepositorio.EliminarPorIdSolicitud(idSolicitud, dbTransaccional)
    Call LogCambioRepositorio.EliminarPorSolicitud(idSolicitud, dbTransaccional)
    Call ValidacionRevisionRepositorio.EliminarPorSolicitud(idSolicitud, dbTransaccional)
    Call RechazoRepositorio.EliminarRechazosPorSolicitud(idSolicitud, dbTransaccional)
    Call SolicitudRepositorio.Eliminar(idSolicitud, dbTransaccional)
    ' --- FIN DE LA MODIFICACIÓN ---
    
    ' 5. Operaciones de Borrado Físico (Archivos) (sin cambios)
    Dim rutaArchivo As Variant
    For Each rutaArchivo In archivosAEliminar
        If fso.FileExists(rutaArchivo) Then
            fso.DeleteFile rutaArchivo, True
        End If
    Next rutaArchivo
    
    ' 6. Confirmación de la Transacción (sin cambios)
    ws.CommitTrans
    
    ' --- SPEC-131: Limpieza de CodConcesionAsociada en BD de No Conformidades (POST-COMMIT) ---
    ' NOTA: Esta operación se ejecuta FUERA de la transacción principal.
    ' FIX ESC-131: Si falla, se LOGUEA/WARN pero NO se propaga - la eliminación principal YA fue confirmada.
    If Len(Nz(sol.codigoSolicitud, "")) > 0 Then
        Dim dbNC As DAO.Database
        On Error Resume Next
        Set dbNC = getdb()
        
        ' TEST SEAM: Force NC cleanup failure for test scenarios (ESC hardening)
        If m_ForceNCCleanupFailure Then
            m_NCCleanupFailureRaised = True
            Err.Raise 513, "SolicitudServicio.NC_Cleanup.Test", "Test forced NC cleanup failure"
        End If
        
        Call NoConformidadRepositorio.LimpiarCodConcesionAsociada(sol.codigoSolicitud, dbNC)
        If Err.Number <> 0 Then
            m_NCCleanupFailureRaised = True
            ' FIX ESC-131: No se propaga el error - solo se loguea/warn
            ' La eliminación principal ya fue confirmada por el COMMIT anterior
            Call modBattery_Canonical.Canonical_Log_Trace("[WARN][EliminarSolicitudCompleta] NC cleanup falló post-commit: " & Err.Number & " - " & Err.description & " | SolicitudID: " & idSolicitud & " | Codigo: " & sol.codigoSolicitud)
        End If
        On Error GoTo Errores
    End If
    ' ---------------------------------------------------------------------------
    
    ' (Notificación movida al inicio por SPEC-101)
    
LimpiarYSalir:
    On Error Resume Next
    ' FIX: Solo cerrar la conexión si la creamos localmente.
    ' Si db fue pasada por el llamador, NO la cerramos (el llamador es dueño).
    If dbCreadaLocalmente And Not dbTransaccional Is Nothing Then dbTransaccional.Close
    Set ws = Nothing
    Set fso = Nothing
    Exit Sub
    
Errores:
    ' 1. CAPTURA INMEDIATA (Snapshot)
    ' Guardamos los detalles del error ANTES de que el Rollback resetee el objeto Err.
    Dim savedErr As Long, savedDesc As String, savedSource As String
    savedErr = Err.Number
    savedDesc = Err.description
    savedSource = Err.source

    ' 2. ROLLBACK SEGURO
    If Not ws Is Nothing Then
        On Error Resume Next ' Evitamos que un fallo en el rollback oculte el error original
        ws.Rollback
        On Error GoTo 0      ' Restauramos el control de errores
    End If
    
    ' 3. PROPAGACIÓN INTELIGENTE (La clave de la mejora)
    ' Verificamos si ya existe un error detallado en el objeto global (que viene del Repositorio)
    If Not g_objLastError Is Nothing Then
        ' ESCENARIO A: El error viene de abajo (ej: RepositorioComun capturó el error de Integridad Referencial).
        ' En este caso, el objeto g_objLastError ya contiene el mensaje "El registro no se puede eliminar..."
        ' Lo único que hacemos es añadir la traza de pila actual y volver a lanzarlo.
        g_objLastError.AddToCallStack "SolicitudServicio.EliminarSolicitudCompleta (ID: " & idSolicitud & ")"
        g_objLastError.Raise ' Esto vuelve a activar el error preservando el mensaje original.
    Else
        ' ESCENARIO B: Es un error ocurrido aquí mismo en el Servicio (ej: Variable no definida, error de lógica).
        ' Creamos un nuevo error Condor usando los datos que capturamos en el paso 1.
        Dim errObj As New CondorError
        errObj.Create savedErr, savedDesc, "SolicitudServicio.EliminarSolicitudCompleta"
        errObj.AddToCallStack "idSolicitud: " & idSolicitud
        errObj.Raise
    End If
End Sub

Public Function getNombreAmigableTipoSolicitud(ByVal codigoInterno As String) As String
    ' RESPONSABILIDAD: Traducir un código interno de tipo de solicitud
    '                  a su nombre descriptivo y amigable para el usuario.
    On Error GoTo Fin
    
    Select Case UCase(Trim(Nz(codigoInterno, "")))
        Case "PC"
            getNombreAmigableTipoSolicitud = "PC - Propuesta de Cambio"
        Case "CD_CA"
            getNombreAmigableTipoSolicitud = "CD/CA - Concesión/Desviación"

        ' --- CORRECCIÓN: Añadido el código exacto que viene de la BD ---
        Case "CD_CA_SUB"
            getNombreAmigableTipoSolicitud = "CD/CA-SUB - Subcontratista"
        Case "CDCASUB" ' (Mantenemos por compatibilidad defensiva)
            getNombreAmigableTipoSolicitud = "CD/CA-SUB - Subcontratista"

        ' --- PC_SUB: Gemelo de PC para subcontratistas ---
        Case "PC_SUB", "PCSUB"
            getNombreAmigableTipoSolicitud = "PC-SUB - Subcontratista"

        Case Else
            ' Si el código no se reconoce, devolvemos el propio código
            getNombreAmigableTipoSolicitud = codigoInterno
    End Select
    
Fin:
    ' Si algo falla, devolvemos el código original.
    If getNombreAmigableTipoSolicitud = "" Then getNombreAmigableTipoSolicitud = codigoInterno
End Function

Public Sub Guardar(ByRef sol As Solicitud, Optional ByRef db As DAO.Database)
    ' Responsabilidad: Delegar la persistencia de la solicitud en el repositorio.
    On Error GoTo Errores
    Call SolicitudRepositorio.GuardarSolicitud(sol, db)
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudServicio.Guardar"
    errObj.Raise
End Sub

Public Function ExisteCodigo(ByVal codigo As String) As Boolean
    ' Responsabilidad: Verificar existencia de código delegando en el repositorio.
    On Error GoTo Errores
    ExisteCodigo = SolicitudRepositorio.ExisteCodigo(codigo)
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudServicio.ExisteCodigo"
    errObj.Raise
End Function

' ==========================================================================
' MÓDULO: SolicitudServicio.cls
' ACCIÓN: CalcularHashFaseModificacion (SPEC-105)
' ==========================================================================
Public Function CalcularHashFaseModificacion(ByVal idSolicitud As Long) As String
    Dim snapshotServ As New SnapshotServicio
    
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then
        CalcularHashFaseModificacion = ""
        Exit Function
    End If
    
    CalcularHashFaseModificacion = snapshotServ.CalcularHashSnapshot(idSolicitud)
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudServicio.CalcularHashFaseModificacion"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud
    errObj.Raise
End Function

Public Function GuardarConVinculacionNC(ByVal sol As Solicitud, ByRef dbCondor As DAO.Database, _
                                         Optional ByVal idNCAnterior As Long = 0) As Boolean
    Dim ncServ As New NoConformidadServicio
    Dim dbNC As DAO.Database
    On Error GoTo Errores
    
    Set dbNC = getdb()
    
    If sol.idNCAsociada > 0 Then
        If Not ncServ.actualizarCodigoConcesionEnTransaccion(sol.idNCAsociada, sol.codigoSolicitud, dbCondor, dbNC) Then
            Err.Raise 513, "GuardarConVinculacionNC", "Error al actualizar CodConcesionAsociada en NC"
        End If
    End If
    
    If idNCAnterior > 0 And idNCAnterior <> sol.idNCAsociada Then
        If Not ncServ.limpiarCodigoConcesionEnTransaccion(idNCAnterior, dbCondor, dbNC) Then
            Err.Raise 513, "GuardarConVinculacionNC", "Error al limpiar CodConcesionAsociada en NC anterior"
        End If
    End If
    
    GuardarConVinculacionNC = True
    
    Exit Function
Errores:
    GuardarConVinculacionNC = False
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "SolicitudServicio.GuardarConVinculacionNC"
    errObj.Raise
End Function







