

Option Compare Database
Option Explicit

Public Event SolicitudCreadaConExito(ByVal idNuevaSolicitud As Long, ByVal tipoSolicitud As String)
Public Event SolicitudModificada()

Private m_idExpedienteSeleccionado As Long
Private m_idSolicitudEnEdicion As Long

' --- VARIABLES PARA DETECCIÓN DE CAMBIOS (DIRTY CHECK) ---
Private m_orig_IdExpediente As Long
Private m_orig_TipoSolicitud As String
Private m_orig_Codigo As String
Private m_orig_IdNC As Long
Private m_orig_Ordinal As String

' --- VARIABLES PARA DETECCIÓN DE NC EN BD EXTERNA ---
Private m_NCVinculadaEnBDExternaa As Boolean
Private m_CodigoNCEnBDExternaa As String
Private m_idNCAnterior As Long

Private Sub Form_Load()
    On Error GoTo Errores
    
    Ajustar Me
    m_idSolicitudEnEdicion = CLng(Nz(Me.OpenArgs, 0))
    
    Call PoblarCombosFijos
    
    If m_idSolicitudEnEdicion > 0 Then
        Me.Caption = "Editar Pre-registro de Solicitud"
        Me.cmdAltaSolicitud.Caption = "Guardar Cambios"
        Call CargarDatosSolicitudExistente(m_idSolicitudEnEdicion)
    Else
        Me.Caption = "Alta de Nueva Solicitud"
        Me.cmdAltaSolicitud.Caption = "Crear Solicitud"
        Call LimpiarCampos_EstadoInicial
    End If
    Me.lblTitulo1.Caption = Me.Caption
    
    ' Inicializamos los valores originales para la comparación posterior
    GuardarValoresOriginales
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("Form_Load", Me)
    DoCmd.Close acForm, Me.name
    Resume LimpiarYSalir
End Sub

Private Sub GuardarValoresOriginales()
    ' Guardamos el estado actual de los controles relevantes
    m_orig_IdExpediente = CLng(val(Nz(Me.txtIDExpediente.value, 0)))
    m_orig_TipoSolicitud = Nz(Me.cboTipoSolicitud.value, "")
    m_orig_Codigo = Nz(Me.txtCodigoSolicitudPropuesto.value, "")
    m_orig_IdNC = CLng(val(Nz(Me.cboNCAsociada.value, 0)))
    m_orig_Ordinal = Nz(Me.txtOrdinalManual.value, "")
End Sub

Private Function HayCambiosSinGuardar() As Boolean
    ' Comparamos valores actuales con los originales
    If CLng(Nz(Me.txtIDExpediente.value, 0)) <> m_orig_IdExpediente Then HayCambiosSinGuardar = True: Exit Function
    If Nz(Me.cboTipoSolicitud.value, "") <> m_orig_TipoSolicitud Then HayCambiosSinGuardar = True: Exit Function
    If Nz(Me.txtCodigoSolicitudPropuesto.value, "") <> m_orig_Codigo Then HayCambiosSinGuardar = True: Exit Function
    If CLng(Nz(Me.cboNCAsociada.value, 0)) <> m_orig_IdNC Then HayCambiosSinGuardar = True: Exit Function
    
    HayCambiosSinGuardar = False
End Function

Private Sub CargarDatosSolicitudExistente(ByVal idSol As Long)
    Dim solServ As New SolicitudServicio
    Dim sol As Solicitud
    Dim ncServ As New NoConformidadServicio
    Dim idNCActual As Long
    Dim nc As NoConformidad
    On Error GoTo Errores
    
    Set sol = solServ.getSolicitudPorID(idSol)
    If sol Is Nothing Then Err.Raise 513, "CargarDatos", "No se encontró la solicitud con ID " & idSol
    
    Call CargarDatosExpediente(sol.idExpediente)
    
    Me.txtCodigoSolicitudPropuesto.value = sol.codigoSolicitud
    
    ' Ahora asignamos la NC guardada. Como el combo ya tiene las NCs filtradas correctas,
    ' el valor se mostrará bien.
    Me.cboNCAsociada.value = Nz(sol.idNCAsociada, 0)
    
    ' Detectar vinculación en BD externa
    idNCActual = CLng(Nz(Me.cboNCAsociada.value, 0))
    m_idNCAnterior = idNCActual
    m_NCVinculadaEnBDExternaa = False
    m_CodigoNCEnBDExternaa = ""
    
    If idNCActual > 0 Then
        m_NCVinculadaEnBDExternaa = ncServ.estaRegistradaEnBaseDatosExternaa(idNCActual)
        If m_NCVinculadaEnBDExternaa Then
            Set nc = ncServ.getNoConformidadPorID(idNCActual)
            If Not nc Is Nothing Then
                m_CodigoNCEnBDExternaa = nc.CodigoNoConformidad
            End If
        End If
    End If
    
    Me.cboTipoSolicitud.value = sol.tipoSolicitud
    
    If Me.cboTipoSolicitud.ListIndex = -1 Then
        Me.cboTipoSolicitud.value = Null
    End If
    
    If sol.idEstadoInterno <> estadoPreregistro Then
        Me.cboTipoSolicitud.Enabled = False
    End If
LimpiarYSalir:
    Exit Sub
Errores:
    Err.Raise Err.Number, "CargarDatosSolicitudExistente", Err.description
End Sub

Private Sub CargarDatosExpediente(ByVal idExpediente As Long)
    Dim expServ As New ExpedienteServicio
    Dim solServ As New SolicitudServicio
    Dim exp As Expediente
    Dim dictNombres As Scripting.Dictionary
    Dim sol As Solicitud
    
    On Error GoTo Errores
    
    Call LimpiarCampos_Solicitud
    
    m_idExpedienteSeleccionado = idExpediente
    ' --- Refrescar NCs filtradas por el proyecto ---
    Call ActualizarComboNC(m_idExpedienteSeleccionado)
    
    ' --- Detectar NC vinculada en BD externa por código CONDOR ---
    If m_idExpedienteSeleccionado > 0 Then
        Dim codigoCondorTemp As String
        codigoCondorTemp = Nz(Me.txtCodigoSolicitudPropuesto.value, "")
        If codigoCondorTemp <> "" Then
            Call DetectarNCVinculadaEnBD(m_idExpedienteSeleccionado, codigoCondorTemp)
        End If
    End If
    
    If m_idExpedienteSeleccionado > 0 Then
        Set exp = expServ.getExpedientePorID(CStr(m_idExpedienteSeleccionado))
        If exp Is Nothing Then Err.Raise 513, "CargarDatosExpediente", "No se encontró el expediente."
        
        Me.txtPalabraClave.value = exp.Nemotecnico
        Me.txtIDExpediente.value = m_idExpedienteSeleccionado
        
        ' Solo proponemos código nuevo si estamos creando o si hemos cambiado de expediente
        If m_idSolicitudEnEdicion = 0 Or (m_idSolicitudEnEdicion > 0 And idExpediente <> m_orig_IdExpediente) Then
             Me.txtCodigoSolicitudPropuesto.value = solServ.ProponerCodigoSolicitud(m_idExpedienteSeleccionado)
        End If
        
        ' UI: Nombres de suministradores
        Set dictNombres = expServ.getNombresSuministradores(m_idExpedienteSeleccionado)
        Me.txtContratistaPrincipal.value = dictNombres("principal")
        Me.txtSubContratista.value = dictNombres("subcontratista")
        
        ' --- SDD tipo-solicitud-pc-subcontratista (Spec-002) ---
        ' Lógica delegada a helper puro para mantener formulario fino.
        ' Principal: PC + CD_CA. Subcontratista (incl. empty/unknown): PC + PC_SUB + CD_CA_SUB.
        Me.cboTipoSolicitud.rowSource = DatosTipoSolicitudHelper.ConstruirRowSourceComboTipo(exp.ContratistaPrincipal)
        ' --- FIN SDD tipo-solicitud-pc-subcontratista ---
        
        Dim puedeCambiarTipo As Boolean
        puedeCambiarTipo = True
        
        If m_idSolicitudEnEdicion > 0 Then
            Set sol = solServ.getSolicitudPorID(m_idSolicitudEnEdicion)
            If Not sol Is Nothing Then
                If sol.idEstadoInterno <> estadoPreregistro Then
                    puedeCambiarTipo = False
                End If
            End If
        End If
        
        Me.txtPalabraClave.Enabled = True
        Me.cmdBuscarExpediente.Enabled = True
        Me.cmdIrABusquedaExpedientes.Enabled = True
        Me.cmdVerDetalleExpediente.Enabled = True
        Me.cboTipoSolicitud.Enabled = puedeCambiarTipo
        Me.cboNCAsociada.Enabled = True
        Me.txtOrdinalManual.Enabled = True
        Me.cmdAltaSolicitud.Enabled = True
        
        If Not puedeCambiarTipo And Not sol Is Nothing Then
            Me.cboTipoSolicitud.value = sol.tipoSolicitud
        End If
    End If
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("CargarDatosExpediente", Me)
    Call LimpiarCampos_EstadoInicial
    Resume LimpiarYSalir
End Sub
Private Sub LimpiarCampos_EstadoInicial()
    m_idExpedienteSeleccionado = 0
    Me.txtPalabraClave.value = ""
    Me.txtIDExpediente.value = ""
    Me.cboTipoSolicitud.rowSource = ""
    Me.cboTipoSolicitud.value = Null
    Me.txtCodigoSolicitudPropuesto.value = ""
    Me.txtOrdinalManual.value = ""
    Me.cboNCAsociada.value = 0
    Me.txtContratistaPrincipal.value = ""
    Me.txtSubContratista.value = ""
    
    Me.txtPalabraClave.Enabled = True
    Me.cmdBuscarExpediente.Enabled = True
    Me.cmdIrABusquedaExpedientes.Enabled = True
    Me.cmdVerDetalleExpediente.Enabled = False
    Me.cboTipoSolicitud.Enabled = False
    Me.cboNCAsociada.Enabled = False
    Me.txtOrdinalManual.Enabled = False
    Me.cmdAltaSolicitud.Enabled = False
    
    Me.txtPalabraClave.SetFocus
End Sub

Private Sub LimpiarCampos_Solicitud()
    Me.cboTipoSolicitud.rowSource = ""
    Me.cboTipoSolicitud.value = Null
    Me.txtCodigoSolicitudPropuesto.value = ""
    Me.txtOrdinalManual.value = ""
    Me.cboNCAsociada.value = 0
End Sub

Private Sub cmdBuscarExpediente_Click()
    Dim palabraClave As String
    Dim expServ As New ExpedienteServicio
    Dim colExpedientes As Scripting.Dictionary
    
    On Error GoTo Errores
    
    If m_idExpedienteSeleccionado > 0 Then
        If MsgBox("Ya hay un expediente seleccionado. Si busca uno nuevo, se perderán los datos no guardados.", vbQuestion + vbYesNo, "CONDOR") = vbNo Then Exit Sub
    End If
    
    palabraClave = Trim(Nz(Me.txtPalabraClave.value, ""))
    If palabraClave = "" Then MsgBox "Introduzca un término de búsqueda.", vbInformation, "CONDOR": Exit Sub
    
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    Set colExpedientes = expServ.getExpedientes(palabraClave)
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False

    Select Case colExpedientes.count
        Case 0
            MsgBox "No se encontraron resultados.", vbInformation, "CONDOR"
            Call LimpiarCampos_EstadoInicial
            Me.txtPalabraClave.value = palabraClave
        Case 1
            Call CargarDatosExpediente(colExpedientes.items(0).idExpediente)
        Case Else
            DoCmd.OpenForm "frmBuscarExpediente", acNormal, , , , acDialog, palabraClave
            On Error Resume Next
            Dim idExpedienteSeleccionado As Long
            idExpedienteSeleccionado = 0
            If Not Forms Is Empty Then
                Dim frmBuscador As Form
                Set frmBuscador = Forms("frmBuscarExpediente")
                If Not frmBuscador Is Nothing Then
                    If IsNumeric(frmBuscador.Tag) And frmBuscador.Tag <> "" Then
                        idExpedienteSeleccionado = CLng(frmBuscador.Tag)
                    End If
                End If
            End If
            On Error GoTo Errores
            If idExpedienteSeleccionado > 0 Then
                Call CargarDatosExpediente(idExpedienteSeleccionado)
            End If
            DoCmd.Close acForm, "frmBuscarExpediente"
    End Select
LimpiarYSalir:
    Exit Sub
Errores:
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Call ManejadorErroresFormulario("cmdBuscarExpediente_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdIrABusquedaExpedientes_Click()
    On Error GoTo Errores
    
    If m_idExpedienteSeleccionado > 0 Then
        If MsgBox("Se perderán los cambios no guardados. ¿Continuar?", vbQuestion + vbYesNo, "CONDOR") = vbNo Then Exit Sub
    End If

    DoCmd.OpenForm "frmBuscarExpediente", acNormal, , , , acDialog
    On Error Resume Next
    Dim idExpedienteSeleccionado2 As Long
    idExpedienteSeleccionado2 = 0
    If Not Forms Is Empty Then
        Dim frmBuscador2 As Form
        Set frmBuscador2 = Forms("frmBuscarExpediente")
        If Not frmBuscador2 Is Nothing Then
            If IsNumeric(frmBuscador2.Tag) And frmBuscador2.Tag <> "" Then
                idExpedienteSeleccionado2 = CLng(frmBuscador2.Tag)
            End If
        End If
    End If
    On Error GoTo Errores
    If idExpedienteSeleccionado2 > 0 Then
        Call CargarDatosExpediente(idExpedienteSeleccionado2)
    End If
    DoCmd.Close acForm, "frmBuscarExpediente"
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdIrABusquedaExpedientes_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdVerDetalleExpediente_Click()
    On Error GoTo Errores
    If m_idExpedienteSeleccionado > 0 Then
        DoCmd.OpenForm "frmDetalleExpediente", acNormal, , , , acDialog, CStr(m_idExpedienteSeleccionado)
    End If
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdVerDetalleExpediente_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub PoblarCombosFijos()
    ' En esta versión, el combo de NC depende del Expediente seleccionado.
    ' Inicialmente lo dejamos vacío o con la opción "Ninguna".
    Me.cboNCAsociada.rowSource = "0;'(Ninguna)'"
    Me.cboNCAsociada.value = 0
End Sub
Private Sub ActualizarComboNC(ByVal idExpediente As Long)
    Dim ncServ As New NoConformidadServicio
    Dim colNCs As Scripting.Dictionary
    Dim key As Variant, nc As NoConformidad
    Dim rowSourceNC As String
    Dim valorActual As Long
    
    On Error GoTo Errores
    
    ' Guardamos el valor actual por si podemos mantenerlo tras el refresco
    valorActual = CLng(Nz(Me.cboNCAsociada.value, 0))
    
    rowSourceNC = "0;'(Ninguna)'"
    
    If idExpediente > 0 Then
        ' Llamamos al servicio con el filtro
        Set colNCs = ncServ.getNoConformidadesPorExpediente(idExpediente)
        
        If Not colNCs Is Nothing Then
            For Each key In colNCs.Keys
                Set nc = colNCs(key)
                ' Formato: ID; Código - Descripción (opcional)
                rowSourceNC = rowSourceNC & ";" & nc.IDNoConformidad & ";" & nc.CodigoNoConformidad
            Next key
        End If
    End If
    
    Me.cboNCAsociada.rowSource = rowSourceNC
    
    ' Intentamos restaurar el valor si aún existe en la nueva lista, si no, a 0
    ' (Access maneja esto visualmente, si el valor no está en la lista y LimitToList=Yes, dará error,
    ' pero si lo asignamos por código suele quedar en blanco o aceptar si no es LimitToList estricto.
    ' Para seguridad, lo reseteamos si cambiamos de expediente).
    Me.cboNCAsociada.value = 0
    
    ' Si teníamos un valor previo y coincide con uno de la lista (ej: al cargar edición), lo ponemos.
    ' (Esta lógica refinada se maneja mejor al cargar los datos de edición).

LimpiarYSalir:
    Set ncServ = Nothing
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("ActualizarComboNC", Me)
    Resume LimpiarYSalir
End Sub

Private Sub DetectarNCVinculadaEnBD(ByVal idExpediente As Long, ByVal codigoCondor As String)
    Dim nc As NoConformidad
    Dim ncServ As New NoConformidadServicio
    
    m_NCVinculadaEnBDExternaa = False
    m_CodigoNCEnBDExternaa = ""
    
    If idExpediente > 0 And codigoCondor <> "" Then
        Set nc = ncServ.getNoConformidadPorCodigoCondor(codigoCondor)
        If Not nc Is Nothing Then
            m_NCVinculadaEnBDExternaa = True
            m_CodigoNCEnBDExternaa = nc.CodigoNoConformidad
            Me.cboNCAsociada.value = nc.IDNoConformidad
        End If
    End If
End Sub

Private Sub cboNCAsociada_AfterUpdate()
    Call ValidarCambioNC
End Sub

Private Sub ValidarCambioNC()
    If m_NCVinculadaEnBDExternaa Then
        Dim respuesta As VbMsgBoxResult
        respuesta = MsgBox( _
            "Esta No Conformidad está registrada en la base de datos de No Conformidades." & vbCrLf & vbCrLf & _
            "Código NC: " & m_CodigoNCEnBDExternaa & vbCrLf & _
            "¿Está seguro de que desea cambiar o eliminar esta relación?", _
            vbExclamation + vbYesNo, _
            "CONDOR" _
        )
        
        If respuesta = vbNo Then
            Me.cboNCAsociada.value = m_orig_IdNC
            Exit Sub
        End If
        
        m_NCVinculadaEnBDExternaa = False
        m_CodigoNCEnBDExternaa = ""
    End If
End Sub

Private Sub cmdAltaSolicitud_Click()
    Dim solServ As New SolicitudServicio, sol As Solicitud
    Dim codigoFinal As String, idNC As Long
    
    Dim expServ As New ExpedienteServicio
    Dim exp As Expediente
    Dim pcServ As DatosPCServicio
    Dim cdcaServ As DatosCDCAServicio
    Dim cdcaSubServ As DatosCDCASUBServicio
    Dim pcsubServ As DatosPCSUBServicio

    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim usaTransaccion As Boolean

    On Error GoTo Errores
    
    If m_idSolicitudEnEdicion > 0 Then
        If Not HayCambiosSinGuardar() Then
            MsgBox "No se han detectado cambios en la solicitud.", vbInformation, "CONDOR"
            DoCmd.Close acForm, Me.name
            Exit Sub
        End If
    End If
    
    DoCmd.Hourglass True
    Me.lblBloqueo.Visible = True
    
    codigoFinal = Nz(Me.txtCodigoSolicitudPropuesto.value, "")
    idNC = CLng(Nz(Me.cboNCAsociada.value, 0))
    
    If m_idSolicitudEnEdicion = 0 Then
        ' =====================================================================
        ' ALTA NUEVA
        ' =====================================================================
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(getdb().name, False, False, ";PWD=dpddpd")
        ws.BeginTrans
        usaTransaccion = True
        
        Set sol = solServ.CrearNuevaSolicitud(m_idExpedienteSeleccionado, Nz(Me.cboTipoSolicitud.value, ""), codigoFinal, idNC)
        Call solServ.Validar(sol, False)
        Dim srvAlta As SolicitudServicio
        Set srvAlta = New SolicitudServicio
        Call srvAlta.Guardar(sol, db)
        
        ' CRÍTICO: Actualizar NC en BD externa dentro de la misma transacción
        If Not srvAlta.GuardarConVinculacionNC(sol, db, 0) Then
            Err.Raise 513, "cmdAltaSolicitud_Click", "Error al vincular NC en BD externa"
        End If
        
        ws.CommitTrans
        usaTransaccion = False
        
        Call RefrescarTablerosPrincipales
        Me.lblBloqueo.Visible = False
        DoCmd.Hourglass False
        MsgBox "Solicitud creada correctamente con el código:" & vbCrLf & vbCrLf & _
               sol.codigoSolicitud & vbCrLf & vbCrLf & _
               "A continuación se abrirá el formulario de datos para completar la información.", _
               vbInformation, "CONDOR"
        RaiseEvent SolicitudCreadaConExito(sol.idSolicitud, sol.tipoSolicitud)
        
    Else
        ' =====================================================================
        ' EDICIÓN
        ' =====================================================================
        Set ws = DBEngine.Workspaces(0)
        Set db = ws.OpenDatabase(getdb().name, False, False, ";PWD=dpddpd")
        ws.BeginTrans
        usaTransaccion = True
        
        Set sol = solServ.getSolicitudPorID(m_idSolicitudEnEdicion)
        If sol Is Nothing Then Err.Raise 513, "Guardar", "La solicitud ya no existe."
        
        Dim haCambiadoExpediente As Boolean
        haCambiadoExpediente = (sol.idExpediente <> CLng(Nz(Me.txtIDExpediente.value, 0)))
        
        sol.idExpediente = CLng(Nz(Me.txtIDExpediente.value, 0))
        sol.tipoSolicitud = Nz(Me.cboTipoSolicitud.value, "")
        sol.codigoSolicitud = codigoFinal
        sol.idNCAsociada = idNC
        sol.fechaModificacion = Now()
        sol.usuarioModificacion = m_ObjUsuarioActivo.nombre
        
        Call solServ.Validar(sol, True)
        Dim srvEdit As SolicitudServicio
        Set srvEdit = New SolicitudServicio
        Call srvEdit.Guardar(sol, db)
        
        ' CRÍTICO: Actualizar NC en BD externa dentro de la misma transacción
        If Not srvEdit.GuardarConVinculacionNC(sol, db, m_idNCAnterior) Then
            Err.Raise 513, "cmdAltaSolicitud_Click", "Error al vincular NC en BD externa"
        End If
        
        If haCambiadoExpediente Then
            Set exp = expServ.getExpedientePorID(CStr(sol.idExpediente))
            Select Case sol.tipoSolicitud
                Case "PC"
                    Set pcServ = New DatosPCServicio
                    Call pcServ.ActualizarCamposDependientesDeExpediente(sol, exp, db)
                Case "CD_CA"
                    Set cdcaServ = New DatosCDCAServicio
                    Call cdcaServ.ActualizarCamposDependientesDeExpediente(sol, exp, db)
                Case "CD_CA_SUB"
                    Set cdcaSubServ = New DatosCDCASUBServicio
                    Call cdcaSubServ.ActualizarCamposDependientesDeExpediente(sol, exp, db)
                Case "PC_SUB"
                    Set pcsubServ = New DatosPCSUBServicio
                    Call pcsubServ.ActualizarCamposDependientesDeExpediente(sol, exp, db)
            End Select
        End If
        
        ws.CommitTrans
        usaTransaccion = False
        
        Me.lblBloqueo.Visible = False
        DoCmd.Hourglass False
        MsgBox "Los cambios en el pre-registro se han guardado correctamente.", vbInformation, "CONDOR"
        RaiseEvent SolicitudModificada
    End If
    
    DoCmd.Close acForm, Me.name

LimpiarYSalir:
    On Error Resume Next
    If usaTransaccion And Not db Is Nothing Then db.Close
    DoCmd.Hourglass False
    Set solServ = Nothing: Set expServ = Nothing: Set exp = Nothing: Set pcServ = Nothing: Set cdcaServ = Nothing: Set cdcaSubServ = Nothing: Set pcsubServ = Nothing
    Set ws = Nothing
    Exit Sub
    
Errores:
    If usaTransaccion And Not ws Is Nothing Then
        ws.Rollback
    End If
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Call ManejadorErroresFormulario("cmdAltaSolicitud_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub txtOrdinalManual_AfterUpdate()
    Dim codigoPropuesto As String, ordinalManual As String, baseCodigo As String, codigoFinal As String
    Dim solServ As New SolicitudServicio
    On Error GoTo Errores
    
    codigoPropuesto = Nz(Me.txtCodigoSolicitudPropuesto.value, "")
    ordinalManual = Trim(Nz(Me.txtOrdinalManual.value, ""))
    If codigoPropuesto = "" Then Exit Sub
    
    If ordinalManual <> "" And IsNumeric(ordinalManual) Then
        baseCodigo = Left(codigoPropuesto, InStrRev(codigoPropuesto, "-"))
        codigoFinal = baseCodigo & CLng(ordinalManual)
        Me.txtCodigoSolicitudPropuesto.value = codigoFinal
        
        Dim srvCheck As SolicitudServicio
        Set srvCheck = New SolicitudServicio
        If srvCheck.ExisteCodigo(codigoFinal) Then
            Dim esMiPropioCodigo As Boolean
            esMiPropioCodigo = False
            If m_idSolicitudEnEdicion > 0 Then
                Dim solActual As Solicitud
                Set solActual = solServ.getSolicitudPorID(m_idSolicitudEnEdicion)
                If Not solActual Is Nothing Then
                    If solActual.codigoSolicitud = codigoFinal Then
                        esMiPropioCodigo = True
                    End If
                End If
            End If
            If Not esMiPropioCodigo Then
                Me.txtCodigoSolicitudPropuesto.BackColor = vbRed
                MsgBox "¡ATENCIÓN! El código '" & codigoFinal & "' ya existe.", vbCritical, "CONDOR"
            Else
                Me.txtCodigoSolicitudPropuesto.BackColor = vbWhite
            End If
        Else
            Me.txtCodigoSolicitudPropuesto.BackColor = vbWhite
        End If
    Else
        If m_idExpedienteSeleccionado > 0 Then
             Me.txtCodigoSolicitudPropuesto.value = solServ.ProponerCodigoSolicitud(m_idExpedienteSeleccionado)
             Me.txtCodigoSolicitudPropuesto.BackColor = vbWhite
        End If
    End If
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("txtOrdinalManual_AfterUpdate", Me)
    Resume LimpiarYSalir
End Sub

Private Sub txtPalabraClave_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        Me.cmdBuscarExpediente.SetFocus
        cmdBuscarExpediente_Click
    End If
End Sub

Private Sub cmdSalir_Click()
    DoCmd.Close acForm, Me.name
End Sub
