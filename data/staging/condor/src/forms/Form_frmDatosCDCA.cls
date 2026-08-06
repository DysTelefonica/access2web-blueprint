Option Compare Database
Option Explicit
' ==========================================================================
' MÓDULO: Form_frmDatosCDCA.cls (VERSIÓN 2.0 - HOMOGENEIZADA CON PC)
' RESPONSABILIDAD: Orquestador principal CD/CA con Dirty Checking y Transacciones.
' ==========================================================================

Public Event SolicitudHaCambiado()

Private m_ViewModel As DatosCDCAViewModel
Private m_CDCA_Srv As DatosCDCAServicio
Private m_SolSrv As SolicitudServicio
Private m_WFSrv As WorkflowServicio

' --- GESTIÓN DEL CICLO DE VIDA DEL FORMULARIO ---

Private Sub Form_Load()
    On Error GoTo Errores
    
    Ajustar Me
    DoCmd.Hourglass True
    
    ' Inicializamos servicios
    Set m_CDCA_Srv = New DatosCDCAServicio
    Set m_SolSrv = New SolicitudServicio
    Set m_WFSrv = New WorkflowServicio
    
    Dim idSol As Long
    idSol = CLng(Nz(Me.OpenArgs, 0))
    If idSol <= 0 Then Err.Raise 513, "Form_Load", "No se recibió un idSolicitud válido."
    
    Call RecargarViewModel(idSol)
    
    ' 3. Navegar a la pestaña correcta según estado
    Dim paginaInicial As String
    paginaInicial = m_WFSrv.getPaginaActivaCDCA(m_ViewModel.Solicitud.idEstadoInterno)
    
    If paginaInicial = "" Then paginaInicial = "tabGeneral"
    
    Call NavegarAPestana(paginaInicial)
    
    Me.Visible = True
    
LimpiarYSalir:
    DoCmd.Hourglass False
    Exit Sub
Errores:
    ManejadorErroresFormulario "frmDatosCDCA.Form_Load", Me
    DoCmd.Close acForm, Me.name
    Resume LimpiarYSalir
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set m_ViewModel = Nothing
    Set m_CDCA_Srv = Nothing
    Set m_SolSrv = Nothing
    Set m_WFSrv = Nothing
End Sub

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
        ' Si no estuviera abierto por algún motivo, lo abrimos en modo lectura
        DoCmd.OpenForm "frmGestionSolicitud", acNormal, , , , acWindowNormal, CStr(m_ViewModel.Solicitud.idSolicitud)
    End If
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdVolverResumen_Click", Me)
    Resume LimpiarYSalir
End Sub

' --- NAVEGACIÓN Y POBLADO DE SUBFORMULARIOS ---

Private Sub NavigationButton_Click()
    On Error GoTo Errores
    
    Call ConfigurarYPoblarInterfaz
    Call PoblarSubformActual
LimpiarYSalir:
    Exit Sub
Errores:
    ManejadorErroresFormulario "frmDatosCDCA.NavigationButton_Click", Me
    Resume LimpiarYSalir
End Sub

Public Sub PoblarSubformActual()
    On Error GoTo Errores
    If Me.NavigationSubform Is Nothing Then Exit Sub
    If Me.NavigationSubform.Form Is Nothing Then Exit Sub
    If m_ViewModel Is Nothing Then Exit Sub
    CallByName Me.NavigationSubform.Form, "PoblarControles", VbMethod, m_ViewModel
LimpiarYSalir:
    Exit Sub
Errores:
    ' Silencioso si el subform aún no está listo
    Resume Next
End Sub

' --- EVENTOS DE PESTAÑAS (Wiring manual necesario en Access) ---
Private Sub tabAprobacionSuministrador_Click(): NavigationButton_Click: End Sub
Private Sub tabDecisionFinal_Click(): NavigationButton_Click: End Sub
Private Sub tabDictamenRAC_Click(): NavigationButton_Click: End Sub
Private Sub tabGeneral_Click(): NavigationButton_Click: End Sub
Private Sub tabImpacto_Click(): NavigationButton_Click: End Sub
Private Sub tabPropuesta_Click(): NavigationButton_Click: End Sub

' ==========================================================================
' FUNCIONES AUXILIARES DE NAVEGACIÓN
' ==========================================================================

Private Sub NavegarAPestana(ByVal nombreTab As String)
    On Error Resume Next
    
    Dim nombreSubform As String
    nombreSubform = FormulariosPadreAuxiliares.GetSubformNameFromTabName(nombreTab, TipoForm_CDCA)
    If nombreSubform = "" Then Exit Sub
    
    Dim focoEstablecido As Boolean
    focoEstablecido = False
    
    On Error Resume Next
    Me.SetFocus
    If Err.Number = 0 Then focoEstablecido = True
    On Error GoTo 0
    
    If focoEstablecido Then
        DoCmd.BrowseTo acBrowseToForm, nombreSubform, Me.name & ".NavigationSubform"
    End If
    
    On Error Resume Next
    Me.Controls(nombreTab).SetFocus
    If Err.Number = 0 Then
        DoEvents
        SendKeys "{Enter}", True
    End If
    On Error GoTo 0
    
    On Error Resume Next
    Call ConfigurarYPoblarInterfaz
    Call PoblarSubformActual
    On Error GoTo 0
End Sub

Public Sub GuardarDesdeSubform(ByVal nombreSub As String, Optional ByVal silenciarErrorValidacion As Boolean = False)
    Dim bloque As enumBloqueFormulario
    Dim cdcaDB As DatosCDCA, cdcaUI As DatosCDCA
    Dim transicionAutomaticaRealizada As Boolean
    Dim solicitudParaTransicion As Solicitud
    Dim huboAvanceAutomatico As Boolean
    Dim estadoAnterior As Long
    
    ' Variables Transacción
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim gestionaTransaccion As Boolean
    
    On Error GoTo Errores
    Me.lblBloqueo.Visible = True: DoCmd.Hourglass True
    transicionAutomaticaRealizada = False: huboAvanceAutomatico = False
    
    ' --- DEFENSA CONTRA ERROR 91 (Spec-167) ---
    If m_ViewModel Is Nothing Then
        Err.Raise 91, "GuardarDesdeSubform", "m_ViewModel no está inicializado."
    End If
    If m_CDCA_Srv Is Nothing Then Set m_CDCA_Srv = New DatosCDCAServicio
    If m_ViewModel.Solicitud Is Nothing Then
        Err.Raise 91, "GuardarDesdeSubform", "m_ViewModel.Solicitud es Nothing."
    End If
    ' --------------------------------------------

    Set solicitudParaTransicion = m_ViewModel.Solicitud
    bloque = FormulariosPadreAuxiliares.MapearBloquePorNombreSub(nombreSub, TipoForm_CDCA)
    If bloque = 0 Then Err.Raise 513, "GuardarDesdeSubform", "Subformulario no reconocido."
    
    ' 1. Obtener y Validar Datos UI
    Set cdcaDB = m_CDCA_Srv.GetDatosCDCA(solicitudParaTransicion.idSolicitud)
    If Not cdcaDB Is Nothing Then Set cdcaUI = FormulariosPadreAuxiliares.ClonarEntidad(cdcaDB, TipoForm_CDCA) Else Set cdcaUI = New DatosCDCA: cdcaUI.idSolicitud = solicitudParaTransicion.idSolicitud
    
    CallByName Me.NavigationSubform.Form, "RecogerValores", VbMethod, m_ViewModel
    Call ExportarVM_aEntidad(bloque, m_ViewModel, cdcaUI)
    
    ' 2. Dirty Check (Optimización)
    If Not cdcaDB Is Nothing Then
        If FormulariosPadreAuxiliares.SonEntidadesIguales(cdcaDB, cdcaUI) Then
            Dim esRacAprobado As Boolean
            esRacAprobado = (bloque = Bloque_DictamenRAC And cdcaUI.racDecision = "APROBADO")
            ' Si es RAC Aprobado, seguimos adelante aunque no haya cambios para intentar el avance
            If Not esRacAprobado Then
                Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
                If Not silenciarErrorValidacion Then MsgBox "No se han detectado cambios.", vbInformation, "CONDOR"
                GoTo SalidaSinGuardar
            End If
        End If
    End If
    
    ' =========================================================================================
    ' LÓGICA DE TRANSACCIÓN PARA DICTAMEN RAC (ATOMICIDAD: GUARDAR + AVANZAR)
    ' =========================================================================================
    If bloque = Bloque_DictamenRAC Then
        Set ws = getWorkspace()
        Set db = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
        gestionaTransaccion = True
        
        ' A. Guardar Datos (Pasando la conexión db)
        Call m_CDCA_Srv.GuardarDictamenRAC(cdcaUI, db)
        
        ' D. Commit de TODO
        ws.CommitTrans
        gestionaTransaccion = False
        
        If Not silenciarErrorValidacion Then
            MsgBox "Datos guardados correctamente.", vbInformation, "CONDOR"
        End If
        
    Else
        ' =========================================================================================
        ' LÓGICA ESTÁNDAR PARA OTROS BLOQUES (Sin transacción compleja por ahora)
        ' =========================================================================================
        Select Case bloque
            Case Bloque_Generales:           transicionAutomaticaRealizada = m_CDCA_Srv.GuardarDatosGenerales(cdcaUI, db)
            Case Bloque_Propuesta:           Call m_CDCA_Srv.GuardarPropuesta(cdcaUI, db)
            Case Bloque_Impacto:             Call m_CDCA_Srv.GuardarImpacto(cdcaUI, db)
            Case Bloque_AprobacionSuministrador
                If Not Me.NavigationSubform.Form.hayCambios() Then
                    Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
                    If Not silenciarErrorValidacion Then MsgBox "No se han detectado cambios.", vbInformation, "CONDOR"
                    GoTo SalidaSinGuardar
                End If
                Call m_CDCA_Srv.GuardarAprobacionSuministrador(cdcaUI, db)
            Case Bloque_DecisionFinal:       Call m_CDCA_Srv.GuardarDecisionFinal(cdcaUI, db)
        End Select
        
        ' Spec-151: Recargar checklist tras guardar AprobSum en Modificación
        If bloque = Bloque_AprobacionSuministrador And solicitudParaTransicion.idEstadoInterno = estadoModificacion Then
            If FormularioAbierto("frmGestionSolicitud") Then
                Call Forms("frmGestionSolicitud").CargarDatosSolicitud(solicitudParaTransicion.idSolicitud, mantenerVistaActual:=True)
            End If
        End If
        
        ' Lógica de avance Datos Generales
        If bloque = Bloque_Generales And (solicitudParaTransicion.idEstadoInterno = estadoRegistro Or transicionAutomaticaRealizada) Then
             
             ' === CORRECCIÓN CRÍTICA DE ESTADO EN MEMORIA ===
             If transicionAutomaticaRealizada Then
                solicitudParaTransicion.idEstadoInterno = estadoRegistro
             End If
             ' ===============================================

             ' Spec-137: Guardar estado anterior para evitar msgbox de Motivos cuando Calidad pasa de REGISTRO a DESARROLLO TÉCNICO
             estadoAnterior = solicitudParaTransicion.idEstadoInterno

             If MsgBox("Datos Generales guardados." & vbCrLf & "¿Asignar tarea al Técnico ahora?", vbQuestion + vbYesNo, "CONDOR") = vbYes Then
                Call m_WFSrv.EjecutarTransicion(solicitudParaTransicion, estadoDesarrolloTecnico, m_ObjUsuarioActivo)
                huboAvanceAutomatico = True
                MsgBox "Solicitud asignada a Desarrollo Técnico.", vbInformation, "CONDOR"
             End If
        End If
        
        ' =========================================================================================
        ' SPEC-132: Mensajes diferenciados y navegación automática en Desarrollo Técnico
        ' SPEC-137: Excluir transición REGISTRO -> DESARROLLO TÉCNICO (hecha por Calidad)
        ' =========================================================================================
        If solicitudParaTransicion.idEstadoInterno = estadoDesarrolloTecnico And _
           estadoAnterior <> estadoRegistro Then
            
            Dim vieneDeRechazo As Boolean
            vieneDeRechazo = (solicitudParaTransicion.revisionCalidadEstado = "RECHAZADO")
            
            Dim tieneDetalle As Boolean
            tieneDetalle = m_CDCA_Srv.EsDetalleCompleto(solicitudParaTransicion.idSolicitud)
            
            Dim tieneMotivos As Boolean
            tieneMotivos = m_CDCA_Srv.EsMotivosCompleto(solicitudParaTransicion.idSolicitud)
            
            ' Scenario 1: Sin motivos - navegar a Motivos
            If Not tieneMotivos And Not vieneDeRechazo Then
                MsgBox "Se ha guardado correctamente, completa el apartado 'Motivos'", vbInformation, "CONDOR"
                Call NavegarAPestana("tabImpacto")
                GoTo LimpiarYSalir
            End If
            
            ' Scenario 2: En Motivos sin detalle - navegar a Propuesta
            If bloque = Bloque_Impacto And Not tieneDetalle Then
                MsgBox "Se ha guardado correctamente, completa el apartado 'Detalle y Descripcion'", vbInformation, "CONDOR"
                Call NavegarAPestana("tabPropuesta")
                GoTo LimpiarYSalir
            End If
            
            ' Scenario 3: Ambas partes completas - preguntar promoción
            If tieneDetalle And tieneMotivos Then
                Dim respPromocion As VbMsgBoxResult
                respPromocion = MsgBox("¿Desea finalizar su tarea y enviar la solicitud para que Calidad la revise? (Esto avanzará de Estado)", vbQuestion + vbYesNo + vbDefaultButton1, "CONDOR")
                
                If respPromocion = vbYes Then
                    Me.lblBloqueo.Visible = True: DoCmd.Hourglass True
                    Call m_WFSrv.EjecutarTransicion(solicitudParaTransicion, estadoModificacion, m_ObjUsuarioActivo)
                    MsgBox "Tarea completada. La solicitud ha sido enviada a Calidad.", vbInformation, "CONDOR"
                    GoTo LimpiarYSalir
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
        
        ' --- NUEVO: AVANCE GUIADO TECNICO (DESARROLLO -> MODIFICACION) ---
        Dim yaSePreguntoPromocion As Boolean
        yaSePreguntoPromocion = False
        If (bloque = Bloque_Propuesta Or bloque = Bloque_Impacto) And solicitudParaTransicion.idEstadoInterno = estadoDesarrolloTecnico Then
            
            Dim cdcaServCheckTec As New DatosCDCAServicio
            
            ' Validamos completitud tecnica
            If cdcaServCheckTec.EsParteTecnicaCompleta(solicitudParaTransicion.idSolicitud) Then
                
                Dim respTecCDCA As VbMsgBoxResult
                Dim msgPromocion As String
                
                If solicitudParaTransicion.revisionCalidadEstado = "RECHAZADO" Then
                    msgPromocion = "Ha corregido los datos tecnicos." & vbCrLf & vbCrLf & _
                                 "¿Desea enviar la solicitud de nuevo a Calidad para su validacion?"
                Else
                    msgPromocion = "Ha completado todos los datos tecnicos obligatorios." & vbCrLf & vbCrLf & _
                                 "¿Desea finalizar su tarea y enviar la solicitud a Calidad?" & vbCrLf & _
                                 "(Esto avanzara el estado a 'MODIFICACION')."
                End If
                
                respTecCDCA = MsgBox(msgPromocion, vbQuestion + vbYesNo + vbDefaultButton1, "CONDOR")
                yaSePreguntoPromocion = True
                
                If respTecCDCA = vbYes Then
                    Me.lblBloqueo.Visible = True: DoCmd.Hourglass True
                    
                    ' Ejecutamos la transicion al estado 4 (MODIFICACION)
                    Call m_WFSrv.EjecutarTransicion(solicitudParaTransicion, estadoModificacion, m_ObjUsuarioActivo)
                    
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
    End If
    
    ' --- AQUÍ SE HA ELIMINADO EL BLOQUE DE RETORNO DE CORRECCIÓN ---

    ' 7. FEEDBACK GENÉRICO
    If Not huboAvanceAutomatico And Not transicionAutomaticaRealizada And bloque <> Bloque_DictamenRAC Then
        If Not silenciarErrorValidacion And Not yaSePreguntoPromocion Then MsgBox "Datos guardados correctamente.", vbInformation, "CONDOR"
    End If
    
    ' 8. REFRESCO UI
    RaiseEvent SolicitudHaCambiado
    Call RecargarViewModel(solicitudParaTransicion.idSolicitud)
    Call ConfigurarYPoblarInterfaz
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
    ' Rollback si estábamos en el bloque transaccional
    If gestionaTransaccion And Not ws Is Nothing Then
        ws.Rollback
    End If
    
    Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
    If Err.Number = 513 And silenciarErrorValidacion Then Exit Sub
    ManejadorErroresFormulario "frmDatosCDCA.GuardarDesdeSubform", Me
End Sub

' ==========================================================================
' MÉTODOS DE MAPEO (Delegados al servicio)
' ==========================================================================

Private Sub ImportarEntidad_aVM(ByRef cdca As DatosCDCA, ByRef vm As DatosCDCAViewModel)
    ' Delegado al servicio - código de negocio extraído
    Call m_CDCA_Srv.ImportarEntidad_aVM(cdca, vm)
End Sub

Private Sub ExportarVM_aEntidad(ByVal bloque As enumBloqueFormulario, ByRef vm As DatosCDCAViewModel, ByRef cdca As DatosCDCA)
    ' Delegado al servicio - código de negocio extraído
    Call m_CDCA_Srv.ExportarVM_aEntidad(bloque, vm, cdca)
End Sub

Public Sub RecargarViewModel(ByVal idSolicitud As Long)
    Dim vm As DatosCDCAViewModel
    
    ' --- DEFENSA CONTRA ERROR 91 ---
    If m_SolSrv Is Nothing Then Set m_SolSrv = New SolicitudServicio
    If m_CDCA_Srv Is Nothing Then Set m_CDCA_Srv = New DatosCDCAServicio
    If m_WFSrv Is Nothing Then Set m_WFSrv = New WorkflowServicio
    ' -------------------------------
    
    On Error GoTo Errores
    
    ' Obtener ViewModel completo del servicio (lógica de negocio centralizada)
    Set vm = m_CDCA_Srv.ObtenerViewModelCompleto(idSolicitud)
    
    ' Asignar al estado del form (única responsabilidad del form)
    Set m_ViewModel = vm
    
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "frmDatosCDCA.RecargarViewModel"
    errObj.Raise
End Sub

' ==========================================================================
' MÉTODOS PÚBLICOS DE INTERFAZ (Botones superiores)
' ==========================================================================

Public Sub ConfigurarYPoblarInterfaz()
    On Error GoTo Errores
    
    If Me.NavigationSubform Is Nothing Then Exit Sub
    If Me.NavigationSubform.Form Is Nothing Then Exit Sub
    If m_ViewModel Is Nothing Then Exit Sub
    If m_ViewModel.Expediente Is Nothing Then Exit Sub
    
    ' Cabecera
    Me.Caption = "Concesión/Desviación: " & m_ViewModel.Solicitud.codigoSolicitud & " (" & m_ViewModel.Expediente.Nemotecnico & ")"
    Me.lblTitulo.Caption = "Solicitud de Concesión/Desviación " & m_ViewModel.Expediente.Nemotecnico
    Me.txtEstado.value = m_ViewModel.EstadoNombre
    
    ' Permisos subformulario
    Dim bloque As enumBloqueFormulario
    bloque = FormulariosPadreAuxiliares.MapearBloquePorNombreSub(Me.NavigationSubform.Form.name, TipoForm_CDCA)
    
    Dim puedeGuardar As Boolean
    Select Case bloque
        Case Bloque_Generales:               puedeGuardar = m_ViewModel.PuedeGuardar_Generales
        Case Bloque_Propuesta:               puedeGuardar = m_ViewModel.PuedeGuardar_Propuesta
        Case Bloque_Impacto:                 puedeGuardar = m_ViewModel.PuedeGuardar_Impacto
        Case Bloque_AprobacionSuministrador: puedeGuardar = m_ViewModel.PuedeGuardar_AprobacionSuministrador
        Case Bloque_DictamenRAC:             puedeGuardar = m_ViewModel.PuedeGuardar_DictamenRAC
        Case Bloque_DecisionFinal:           puedeGuardar = m_ViewModel.PuedeGuardar_DecisionFinal
    End Select
    
    Call FormulariosPadreAuxiliares.HabilitarGuardarEnSubform(Me.NavigationSubform.Form, puedeGuardar)
    Call HabilitarControlesSubformulario(Me.NavigationSubform.Form, puedeGuardar)
    
    ' Botones de Decisión (Excluyentes)
    Dim estadoID As Long
    Dim esRolCalidad As Boolean
    estadoID = m_ViewModel.Solicitud.idEstadoInterno
    esRolCalidad = (rolUsuario = rol.Calidad Or rolUsuario = rol.Administrador)
    
    ' Banner Rechazo
    If m_ViewModel.vm_revisionCalidadEstado = "RECHAZADO" And estadoID = estadoModificacion Then
        Me.lblBannerRechazo.Visible = True
        Me.lblBannerRechazo.Caption = "RECHAZADO POR CALIDAD: " & m_ViewModel.vm_revisionCalidadComentarios
    Else
        Me.lblBannerRechazo.Visible = False
    End If
    
LimpiarYSalir:
    Exit Sub
Errores:
    ManejadorErroresFormulario "frmDatosCDCA.ConfigurarYPoblarInterfaz", Me
    Resume LimpiarYSalir
End Sub

Public Sub NotificarCambioDeEstado()
    RaiseEvent SolicitudHaCambiado
End Sub
' EN: Form_frmDatosCDCA.cls

Public Sub ProcesarCierreFormalizacion()
    Dim wfServ As New WorkflowServicio
    Dim cdcaUI As DatosCDCA, cdcaDB As DatosCDCA
    Dim fDialog As FileDialog
    Dim rutaSeleccionada As String
    
    On Error GoTo Errores
    
    ' 1. Aviso previo obligatorio
    If MsgBox("Para APROBAR definitivamente la solicitud, es obligatorio adjuntar el documento de Cierre." & vbCrLf & vbCrLf & _
              "Al pulsar Aceptar, se abrira una ventana para que seleccione el archivo PDF firmado.", _
              vbInformation + vbOKCancel, "CONDOR") = vbCancel Then
        Exit Sub
    End If
    
    ' 2. Selección del Archivo (Obligatorio)
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    fDialog.Title = "SELECCIONE EL PDF DE CIERRE (Obligatorio)"
    fDialog.Filters.Clear
    fDialog.Filters.Add "Documentos PDF", "*.pdf"
    fDialog.AllowMultiSelect = False
    
    If fDialog.Show = -1 Then
        rutaSeleccionada = fDialog.SelectedItems(1)
    Else
        MsgBox "Operacion cancelada. No se ha seleccionado el documento firmado.", vbExclamation, "CONDOR"
        Exit Sub
    End If
    
    Me.lblBloqueo.Visible = True
    Me.lblBloqueo.Caption = "Procesando cierre, guardando archivo y firmando..."
    DoCmd.Hourglass True
    
    ' 3. Preparar Datos
    Set cdcaDB = m_CDCA_Srv.GetDatosCDCA(m_ViewModel.Solicitud.idSolicitud)
    Set cdcaUI = FormulariosPadreAuxiliares.ClonarEntidad(cdcaDB, TipoForm_CDCA)
    
    cdcaUI.decisionFinal = m_ViewModel.vm_decisionFinal
    cdcaUI.observacionesFinales = m_ViewModel.vm_obsDecisionFinal
    cdcaUI.NombreFirmanteFinal = m_ViewModel.vm_NombreFirmanteFinal
    
    ' 4. Llamada a la Transacción Maestra (Pasando la ruta del archivo)
    Call wfServ.EjecutarCierreFormalizacion(m_ViewModel.Solicitud, cdcaUI, m_ObjUsuarioActivo, rutaSeleccionada)
    
    ' 5. Éxito
    Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
    MsgBox "Solicitud APROBADA y CERRADA correctamente." & vbCrLf & _
           "El documento de Cierre ha sido archivado.", vbInformation, "CONDOR"
           
    Call NotificarCambioDeEstado
    Call RecargarViewModel(m_ViewModel.Solicitud.idSolicitud)
    Call ConfigurarYPoblarInterfaz
    Call PoblarSubformActual
    
    Exit Sub
    
Errores:
    Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
    Call ManejadorErroresFormulario("ProcesarCierreFormalizacion", Me)
End Sub







