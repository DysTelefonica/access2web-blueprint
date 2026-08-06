
Option Compare Database
Option Explicit
' ==========================================================================
' MÓDULO: Form_frmDatosPCSUB.cls
' RESPONSABILIDAD: Orquestador principal para la edición de datos de una
'                  Propuesta de Cambio Subcontratista (PCSUB).
'                  Gestiona la Navegación, el guardado y el flujo de trabajo.
' FUENTE DE VERDAD: tbDatosPCSUB en condor_datos.accdb (ERD 2026-05-01)
' ==========================================================================

'=== Estado (Vista desconectada con VM) ===
Private m_ViewModel As DatosPCSUBViewModel
Private m_PCSrv As DatosPCSUBServicio
Private m_SolSrv As SolicitudServicio
Private m_WFSrv As WorkflowServicio
Public Event SolicitudHaCambiado()

Private Sub cmdCerrar_Click()
    DoCmd.Close acForm, Me.name
End Sub

Private Sub cmdVolverResumen_Click()
    On Error GoTo Errores
    Me.Visible = False
    If FormularioAbierto("frmGestionSolicitud") Then
        Forms("frmGestionSolicitud").Visible = True
        Forms("frmGestionSolicitud").SetFocus
    Else
        DoCmd.OpenForm "frmGestionSolicitud", acNormal, , , , acWindowNormal, CStr(m_ViewModel.Solicitud.idSolicitud)
    End If
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdVolverResumen_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub Form_Load()
    On Error GoTo Errores
    Ajustar Me
    DoCmd.Hourglass True

    Set m_PCSrv = New DatosPCSUBServicio
    Set m_SolSrv = New SolicitudServicio
    Set m_WFSrv = New WorkflowServicio

    Dim idSol As Long
    idSol = CLng(Nz(Me.OpenArgs, 0))
    If idSol <= 0 Then Err.Raise 513, "Form_Load", "No se recibió un idSolicitud válido."

    Call RecargarViewModel(idSol)
    Dim paginaInicial As String
    paginaInicial = m_WFSrv.getPaginaActivaPCSUB(m_ViewModel.Solicitud.idEstadoInterno)
    Call NavegarAPestana(paginaInicial)
LimpiarYSalir:
    DoCmd.Hourglass False
    Exit Sub
Errores:
    ManejadorErroresFormulario "frmDatosPCSUB.Form_Load", Me
    DoCmd.Close acForm, Me.name
    Resume LimpiarYSalir
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set m_ViewModel = Nothing
    Set m_PCSrv = Nothing
    Set m_SolSrv = Nothing
    Set m_WFSrv = Nothing
End Sub

'================================================================
' NAVEGACIÓN Y POBLADO DE SUBFORMULARIOS
'================================================================
Private Sub NavegarAPestana(ByVal nombreTab As String)
    On Error Resume Next

    Dim nombreSubform As String
    nombreSubform = FormulariosPadreAuxiliares.GetSubformNameFromTabName(nombreTab, TipoForm_PCSUB)
    If nombreSubform = "" Then Exit Sub

    Dim focoEstablecido As Boolean
    focoEstablecido = False
    On Error Resume Next
    Me.SetFocus
    If Err.Number = 0 Then focoEstablecido = True
    On Error GoTo 0

    If focoEstablecido Then
        DoCmd.BrowseTo acBrowseToForm, nombreSubform, Me.name & ".NavigationSubform"
        If Err.Number <> 0 Then
            DoCmd.OpenForm Me.name, , , , , , Me.name
            Me.SetFocus
            DoCmd.BrowseTo acBrowseToForm, nombreSubform, Me.name & ".NavigationSubform"
        End If
    End If

    On Error Resume Next
    Me.Controls(nombreTab).SetFocus
    If Err.Number = 0 Then
        DoEvents
        SendKeys "{Enter}", True
    End If
    On Error GoTo 0

    Call ConfigurarYPoblarInterfaz
    Call PoblarSubformActual
End Sub

Private Sub NavigationButton_Click()
    On Error GoTo Errores
    Call ConfigurarYPoblarInterfaz
    Call PoblarSubformActual
    Exit Sub
Errores:
    ManejadorErroresFormulario "frmDatosPCSUB.NavigationButton_Click", Me
End Sub

Public Sub PoblarSubformActual()
    On Error Resume Next
    If Me.NavigationSubform Is Nothing Then Exit Sub
    If Me.NavigationSubform.Form Is Nothing Then Exit Sub
    If m_ViewModel Is Nothing Then Exit Sub
    CallByName Me.NavigationSubform.Form, "PoblarControles", VbMethod, m_ViewModel
LimpiarYSalir:
    Exit Sub
Errores:
    Resume Next
End Sub

Public Sub GuardarDesdeSubform(ByVal nombreSub As String, Optional ByVal silenciarErrorValidacion As Boolean = False)
    Dim bloque As enumBloqueFormulario
    Dim pcDB As DatosPCSUB, pcUI As DatosPCSUB
    Dim transicionAutomaticaRealizada As Boolean
    Dim solicitudParaTransicion As Solicitud
    Dim huboAvanceAutomatico As Boolean
    Dim estadoAnterior As Long
    Dim helperResult As Object
    Dim promptResult As VbMsgBoxResult
    Dim messageAlreadyRendered As Boolean

    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim gestionaTransaccion As Boolean

    On Error GoTo Errores
    Me.lblBloqueo.Visible = True: DoCmd.Hourglass True
    transicionAutomaticaRealizada = False: huboAvanceAutomatico = False

    If m_ViewModel Is Nothing Then Err.Raise 91, "GuardarDesdeSubform", "m_ViewModel no está inicializado."
    If m_PCSrv Is Nothing Then Set m_PCSrv = New DatosPCSUBServicio
    If m_ViewModel.Solicitud Is Nothing Then Err.Raise 91, "GuardarDesdeSubform", "m_ViewModel.Solicitud es Nothing."

    Set solicitudParaTransicion = m_ViewModel.Solicitud

    bloque = FormulariosPadreAuxiliares.MapearBloquePorNombreSub(nombreSub, TipoForm_PCSUB)
    If bloque = 0 Then Err.Raise 513, "GuardarDesdeSubform", "Subformulario no reconocido."

    If bloque = Bloque_Generales Then
        CallByName Me.NavigationSubform.Form, "RecogerValores", VbMethod, m_ViewModel

        Set helperResult = DatosPCSUBGuardarHelper_GuardarDesdeSubformPlan(nombreSub, m_ViewModel, db)
        If helperResult Is Nothing Then Err.Raise 513, "GuardarDesdeSubform", "El helper de guardado no devolvió resultado."
        If Not CBool(helperResult("ok")) Then Err.Raise 513, "GuardarDesdeSubform", CStr(helperResult("errorText"))

        messageAlreadyRendered = False
        If CBool(helperResult("requiresUserPrompt")) And Not silenciarErrorValidacion Then
            promptResult = MsgBox(CStr(helperResult("message")) & vbCrLf & "¿Asignar tarea al Técnico ahora?", vbQuestion + vbYesNo, "CONDOR")
            messageAlreadyRendered = True
            If promptResult = vbYes Then
                Set helperResult = DatosPCSUBGuardarHelper_AsignarTecnicoPostGuardado(m_ViewModel, db)
                If helperResult Is Nothing Then Err.Raise 513, "GuardarDesdeSubform", "El helper de asignación no devolvió resultado."
                If Not CBool(helperResult("ok")) Then Err.Raise 513, "GuardarDesdeSubform", CStr(helperResult("errorText"))
                messageAlreadyRendered = False
            End If
        End If

        If Not silenciarErrorValidacion Then
            If Not messageAlreadyRendered Then
                If Len(Trim$(Nz(helperResult("message"), ""))) > 0 Then MsgBox CStr(helperResult("message")), vbInformation, "CONDOR"
            End If
        End If

        RaiseEvent SolicitudHaCambiado
        Call RecargarViewModel(m_ViewModel.Solicitud.idSolicitud)
        Call ConfigurarYPoblarInterfaz
        On Error Resume Next
        If Not Me.NavigationSubform.Form Is Nothing Then Me.NavigationSubform.Form.Controls("cmdGuardar").SetFocus
        On Error GoTo Errores
        Call PoblarSubformActual
        GoTo SalidaSinGuardar
    End If

    Set pcDB = m_PCSrv.getDatosPCSUB(solicitudParaTransicion.idSolicitud)
    If Not pcDB Is Nothing Then
        Set pcUI = FormulariosPadreAuxiliares.ClonarEntidad(pcDB, TipoForm_PCSUB)
    Else
        Set pcUI = New DatosPCSUB
        pcUI.idSolicitud = solicitudParaTransicion.idSolicitud
    End If

    CallByName Me.NavigationSubform.Form, "RecogerValores", VbMethod, m_ViewModel
    Call ExportarVM_aEntidad(bloque, m_ViewModel, pcUI)

    ' Dirty check
    If Not pcDB Is Nothing Then
        If FormulariosPadreAuxiliares.SonEntidadesIguales(pcDB, pcUI) Then
            Dim esRacDecisionFinal As Boolean
                esRacDecisionFinal = (bloque = Bloque_DictamenRAC And pcUI.racDecision = "APROBADO")
            If Not esRacDecisionFinal Then
                Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
                If Not silenciarErrorValidacion Then MsgBox "No se han detectado cambios.", vbInformation, "CONDOR"
                GoTo SalidaSinGuardar
            End If
        End If
    End If

    ' =========================================================================================
    ' LÓGICA DE TRANSACCIÓN PARA DICTAMEN RAC
    ' =========================================================================================
    If bloque = Bloque_DictamenRAC Then
        Set ws = getWorkspace()
        Set db = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
        gestionaTransaccion = True
        Call m_PCSrv.GuardarDictamenRAC(pcUI, db)
        ws.CommitTrans
        gestionaTransaccion = False
        If Not silenciarErrorValidacion Then MsgBox "Datos guardados correctamente.", vbInformation, "CONDOR"
    Else
        Select Case bloque
            Case Bloque_Generales:               transicionAutomaticaRealizada = m_PCSrv.GuardarDatosGenerales(pcUI, db)
            Case Bloque_Propuesta:               Call m_PCSrv.GuardarPropuesta(pcUI)
            Case Bloque_Impacto:                 Call m_PCSrv.GuardarImpacto(pcUI)
            Case Bloque_AprobacionSuministrador: Call m_PCSrv.GuardarAprobacionSuministrador(pcUI)
            Case Bloque_DecisionFinal:           Call m_PCSrv.GuardarDecisionFinal(pcUI, db)
        End Select
    End If

    ' Spec-151: Recargar checklist tras guardar AprobSum en Modificación
    If bloque = Bloque_AprobacionSuministrador And solicitudParaTransicion.idEstadoInterno = estadoModificacion Then
        If FormularioAbierto("frmGestionSolicitud") Then
            Call Forms("frmGestionSolicitud").CargarDatosSolicitud(solicitudParaTransicion.idSolicitud, mantenerVistaActual:=True)
            DoEvents
            Me.Repaint
        End If
    End If

    ' =========================================================================================
    ' BLOQUE 1: Avance desde REGISTRO -> DESARROLLO TÉCNICO (Datos Generales guardados)
    ' =========================================================================================
    If bloque = Bloque_Generales And (solicitudParaTransicion.idEstadoInterno = estadoRegistro Or transicionAutomaticaRealizada) Then
        If transicionAutomaticaRealizada Then
           solicitudParaTransicion.idEstadoInterno = estadoRegistro
        End If
        estadoAnterior = solicitudParaTransicion.idEstadoInterno
        If MsgBox("Datos Generales guardados." & vbCrLf & "¿Asignar tarea al Técnico ahora?", vbQuestion + vbYesNo, "CONDOR") = vbYes Then
            Call m_WFSrv.EjecutarTransicion(solicitudParaTransicion, estadoDesarrolloTecnico, m_ObjUsuarioActivo)
            solicitudParaTransicion.idEstadoInterno = estadoDesarrolloTecnico
            huboAvanceAutomatico = True
            MsgBox "Solicitud asignada a Desarrollo Técnico.", vbInformation, "CONDOR"
        End If
    End If

    ' =========================================================================================
    ' BLOQUE 2: SPEC-132/137 — Mensajes diferenciados y navegación automática en DESARROLLO TÉCNICO
    ' =========================================================================================
    If solicitudParaTransicion.idEstadoInterno = estadoDesarrolloTecnico And _
       estadoAnterior <> estadoRegistro Then

       Dim vieneDeRechazo As Boolean
       vieneDeRechazo = (solicitudParaTransicion.revisionCalidadEstado = "RECHAZADO")

       Dim tieneDetalle As Boolean
       tieneDetalle = m_PCSrv.EsDetalleCompleto(solicitudParaTransicion.idSolicitud)

       Dim tieneMotivos As Boolean
       tieneMotivos = m_PCSrv.EsMotivosCompleto(solicitudParaTransicion.idSolicitud)

       ' Scenario 1: Sin motivos - navegar a Motivos
       If Not tieneMotivos And Not vieneDeRechazo Then
           MsgBox "Se ha guardado correctamente, completa el apartado 'Motivos'", vbInformation, "CONDOR"
           Call NavegarAPestana("tabImpacto")
           GoTo LimpiarYSalir
       End If

       ' Scenario 2: En Motivos sin detalle - navegar a Detalle
       If bloque = Bloque_Impacto And Not tieneDetalle Then
           MsgBox "Se ha guardado correctamente, completa el apartado 'Detalle y Descripcion'", vbInformation, "CONDOR"
           Call NavegarAPestana("tabPropuesta")
           GoTo LimpiarYSalir
       End If

       ' Scenario 3: Ambas partes completas - preguntar promoción una sola vez
       If (bloque = Bloque_Propuesta Or bloque = Bloque_Impacto) And tieneMotivos And tieneDetalle Then
           Dim respPromocion As VbMsgBoxResult
           Dim msgPromocion As String
           If solicitudParaTransicion.revisionCalidadEstado = "RECHAZADO" Then
               msgPromocion = "Ha corregido los datos técnicos." & vbCrLf & vbCrLf & _
                            "¿Desea enviar la solicitud de nuevo a Calidad para su validación?"
           Else
               msgPromocion = "Ha completado todos los datos técnicos obligatorios." & vbCrLf & vbCrLf & _
                            "¿Desea finalizar su tarea y enviar la solicitud a Calidad?" & vbCrLf & _
                            "(Esto avanzará el estado a 'MODIFICACIÓN')."
           End If
           respPromocion = MsgBox(msgPromocion, vbQuestion + vbYesNo + vbDefaultButton1, "CONDOR")
           If respPromocion = vbYes Then
               Me.lblBloqueo.Visible = True: DoCmd.Hourglass True
               Call m_WFSrv.EjecutarTransicion(solicitudParaTransicion, estadoModificacion, m_ObjUsuarioActivo)
               solicitudParaTransicion.idEstadoInterno = estadoModificacion
               huboAvanceAutomatico = True
               MsgBox "Tarea completada. La solicitud ha sido enviada a Calidad.", vbInformation, "CONDOR"
           Else
               If bloque = Bloque_Propuesta Then
                   Call NavegarAPestana("tabImpacto")
               Else
                   Call NavegarAPestana("tabPropuesta")
               End If
               GoTo LimpiarYSalir
           End If
       End If
    End If

    RaiseEvent SolicitudHaCambiado
    Call RecargarViewModel(solicitudParaTransicion.idSolicitud)
    Call ConfigurarYPoblarInterfaz

    On Error Resume Next
    If Not Me.NavigationSubform.Form Is Nothing Then Me.NavigationSubform.Form.Controls("cmdGuardar").SetFocus
    On Error GoTo Errores
    Call PoblarSubformActual

SalidaSinGuardar:
    On Error Resume Next
    Me.SetFocus: Me.NavigationSubform.SetFocus
    If Not Me.NavigationSubform.Form Is Nothing Then Me.NavigationSubform.Form.Controls("cmdGuardar").SetFocus
    On Error GoTo Errores

LimpiarYSalir:
    On Error Resume Next
    If gestionaTransaccion And Not db Is Nothing Then db.Close
    Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
    Exit Sub

Errores:
    If gestionaTransaccion And Not ws Is Nothing Then ws.Rollback
    Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
    If Err.Number = 513 And silenciarErrorValidacion Then Exit Sub
    ManejadorErroresFormulario "frmDatosPCSUB.GuardarDesdeSubform", Me
End Sub

Public Sub ConfigurarYPoblarInterfaz()
    On Error GoTo Errores

    If Me.NavigationSubform Is Nothing Then Exit Sub
    If Me.NavigationSubform.Form Is Nothing Then Exit Sub
    If m_ViewModel Is Nothing Then Exit Sub
    If m_ViewModel.Expediente Is Nothing Then Exit Sub

    Me.Caption = "Propuesta de Cambio Subcontratista: " & m_ViewModel.Solicitud.codigoSolicitud & " (" & m_ViewModel.Expediente.Nemotecnico & ")"
    Me.lblTitulo.Caption = "Solicitud de PCSUB " & m_ViewModel.Expediente.Nemotecnico
    Me.txtEstado.value = m_ViewModel.EstadoNombre

    Dim bloque As enumBloqueFormulario
    bloque = FormulariosPadreAuxiliares.MapearBloquePorNombreSub(Me.NavigationSubform.Form.name, TipoForm_PCSUB)

    Dim puedeGuardar As Boolean
    Select Case bloque
        Case Bloque_Generales:               puedeGuardar = m_ViewModel.PuedeGuardar_Generales
        Case Bloque_Propuesta:               puedeGuardar = m_ViewModel.PuedeGuardar_Propuesta
        Case Bloque_Impacto:                 puedeGuardar = m_ViewModel.PuedeGuardar_Impacto
        Case Bloque_AprobacionSuministrador: puedeGuardar = m_ViewModel.PuedeGuardar_AprobacionSuministrador
        Case Bloque_DictamenRAC:             puedeGuardar = m_ViewModel.PuedeGuardar_DictamenRAC
        Case Bloque_DecisionFinal:           puedeGuardar = m_ViewModel.PuedeGuardar_DecisionFinal
        Case Else:                           puedeGuardar = False
    End Select

    Call FormulariosPadreAuxiliares.HabilitarGuardarEnSubform(Me.NavigationSubform.Form, puedeGuardar)
    Call HabilitarControlesSubformulario(Me.NavigationSubform.Form, puedeGuardar)

    If m_ViewModel.vm_revisionCalidadEstado = "RECHAZADO" And m_ViewModel.Solicitud.idEstadoInterno = estadoModificacion Then
        Me.lblBannerRechazo.Visible = True
        Me.lblBannerRechazo.Caption = "RECHAZADO POR CALIDAD: " & m_ViewModel.vm_revisionCalidadComentarios
    Else
        Me.lblBannerRechazo.Visible = False
    End If
LimpiarYSalir:
    Exit Sub
Errores:
    ManejadorErroresFormulario "frmDatosPCSUB.ConfigurarYPoblarInterfaz", Me
    Resume LimpiarYSalir
End Sub

Public Sub RecargarViewModel(ByVal idSolicitud As Long)
    Dim vm As DatosPCSUBViewModel

    If m_SolSrv Is Nothing Then Set m_SolSrv = New SolicitudServicio
    If m_PCSrv Is Nothing Then Set m_PCSrv = New DatosPCSUBServicio
    If m_WFSrv Is Nothing Then Set m_WFSrv = New WorkflowServicio

    On Error GoTo Errores
    Set vm = m_PCSrv.ObtenerViewModelCompleto(idSolicitud)
    Set m_ViewModel = vm
    Exit Sub
Errores:
    Dim e As New CondorError
    e.Create Err.Number, Err.description, "frmDatosPCSUB.RecargarViewModel"
    e.AddToCallStack "idSolicitud: " & idSolicitud
    e.Raise
End Sub

Private Sub ImportarEntidad_aVM(ByRef pc As DatosPCSUB, ByRef vm As DatosPCSUBViewModel)
    Call m_PCSrv.ImportarEntidad_aVM(pc, vm)
End Sub

Private Sub ExportarVM_aEntidad(ByVal bloque As enumBloqueFormulario, ByRef vm As DatosPCSUBViewModel, ByRef pc As DatosPCSUB)
    Call m_PCSrv.ExportarVM_aEntidad(bloque, vm, pc)
End Sub

Private Sub tabAprobacionSuministrador_Click()
    NavigationButton_Click
End Sub

Private Sub tabDecisionFinal_Click()
    NavigationButton_Click
End Sub

Private Sub tabDictamenRAC_Click()
    NavigationButton_Click
End Sub

Private Sub tabGeneral_Click()
    NavigationButton_Click
End Sub

Private Sub tabImpacto_Click()
    NavigationButton_Click
End Sub

Private Sub tabPropuesta_Click()
    NavigationButton_Click
End Sub

Public Sub NotificarCambioDeEstado()
    On Error GoTo Errores
    RaiseEvent SolicitudHaCambiado
    Exit Sub
Errores:
    ManejadorErroresFormulario "frmDatosPCSUB.NotificarCambioDeEstado", Me
End Sub

Public Sub ProcesarCierreFormalizacion()
    Dim wfServ As New WorkflowServicio
    Dim pcUI As DatosPCSUB, pcDB As DatosPCSUB
    Dim fDialog As FileDialog
    Dim rutaSeleccionada As String

    On Error GoTo Errores

    If MsgBox("Para APROBAR definitivamente la solicitud, es obligatorio adjuntar el documento de Cierre." & vbCrLf & vbCrLf & _
              "Al pulsar Aceptar, se abrirá una ventana para que seleccione el archivo PDF firmado.", _
              vbInformation + vbOKCancel, "CONDOR") = vbCancel Then Exit Sub

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    fDialog.Title = "SELECCIONE EL PDF DE CIERRE (Obligatorio)"
    fDialog.Filters.Clear
    fDialog.Filters.Add "Documentos PDF", "*.pdf"
    fDialog.AllowMultiSelect = False

    If fDialog.Show = -1 Then
        rutaSeleccionada = fDialog.SelectedItems(1)
    Else
        MsgBox "Operación cancelada. No se ha seleccionado el documento firmado.", vbExclamation, "CONDOR - Cierre"
        Exit Sub
    End If

    Me.lblBloqueo.Visible = True
    Me.lblBloqueo.Caption = "Procesando cierre, guardando archivo y firmando..."
    DoCmd.Hourglass True

    Set pcDB = m_PCSrv.getDatosPCSUB(m_ViewModel.Solicitud.idSolicitud)
    Set pcUI = FormulariosPadreAuxiliares.ClonarEntidad(pcDB, TipoForm_PCSUB)

    pcUI.decisionFinal = m_ViewModel.vm_decisionFinal
    pcUI.obsDecisionFinal = m_ViewModel.vm_obsDecisionFinal
    pcUI.NombreAutoridadDiseno = m_ViewModel.vm_NombreAutoridadDiseno
    pcUI.NombreFirmanteFinal = m_ViewModel.vm_NombreFirmanteFinal

    Call wfServ.EjecutarCierreFormalizacion(m_ViewModel.Solicitud, pcUI, m_ObjUsuarioActivo, rutaSeleccionada)

    Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
    MsgBox "Solicitud APROBADA y CERRADA correctamente." & vbCrLf & _
           "El documento de Cierre ha sido archivado.", vbInformation, "CONDOR"

    RaiseEvent SolicitudHaCambiado
    Call RecargarViewModel(m_ViewModel.Solicitud.idSolicitud)
    Call ConfigurarYPoblarInterfaz
    Call PoblarSubformActual
    Exit Sub

Errores:
    Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
    Call ManejadorErroresFormulario("ProcesarCierreFormalizacion", Me)
End Sub
