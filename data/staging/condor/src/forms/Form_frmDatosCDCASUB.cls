
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_frmDatosCDCASUB.cls
' RESPONSABILIDAD: Orquestador principal CD/CA SUB (Alineado con PC/CDCA)
' ==========================================================================

Public Event SolicitudHaCambiado()

Private m_ViewModel As DatosCDCASUBViewModel
Private m_CDCASUB_Srv As DatosCDCASUBServicio
Private m_SolSrv As SolicitudServicio
Private m_WFSrv As WorkflowServicio

' --- GESTIÓN DEL CICLO DE VIDA ---

Private Sub Form_Load()
    On Error GoTo Errores
    Ajustar Me
    DoCmd.Hourglass True
    
    ' Inicializamos servicios
    Set m_CDCASUB_Srv = New DatosCDCASUBServicio
    Set m_SolSrv = New SolicitudServicio
    Set m_WFSrv = New WorkflowServicio
    
    Dim idSol As Long
    idSol = CLng(Nz(Me.OpenArgs, 0))
    
    ' Si recibimos ID, cargamos. Si no, esperamos a llamada externa.
    If idSol > 0 Then
        Call CargarDatos(idSol)
    End If
    
LimpiarYSalir:
    DoCmd.Hourglass False
    Exit Sub
Errores:
    ManejadorErroresFormulario "frmDatosCDCASUB.Form_Load", Me
    DoCmd.Close acForm, Me.name
    Resume LimpiarYSalir
End Sub

' --- MÉTODO PÚBLICO PARA FORZAR RECARGA (Soluciona el problema de datos antiguos) ---
Public Sub CargarDatos(ByVal idSol As Long)
    On Error GoTo Errores
    
    ' 1. Cargar el ViewModel con los datos frescos de la BD
    Call RecargarViewModel(idSol)
    
    ' 2. Configurar la interfaz (Título, botones)
    Call ConfigurarYPoblarInterfaz
    
    ' 3. Navegar a la pestaña correcta según estado
    Dim paginaInicial As String
    paginaInicial = m_WFSrv.getPaginaActivaCDCASUB(m_ViewModel.Solicitud.idEstadoInterno)
    
    ' DEBUG: Asegurar que paginaInicial no es nula
    If paginaInicial = "" Then paginaInicial = "tabGeneral"
    
    Call NavegarAPestana(paginaInicial)
    
    Me.Visible = True
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "frmDatosCDCASUB.CargarDatos", Me
End Sub
' ==========================================================================
' INTERFAZ PÚBLICA (CONFIGURACIÓN VISUAL)
' ==========================================================================

Public Sub ConfigurarYPoblarInterfaz()
    On Error GoTo Errores
    
    If Me.NavigationSubform Is Nothing Then Exit Sub
    If Me.NavigationSubform.Form Is Nothing Then Exit Sub
    If m_ViewModel Is Nothing Then Exit Sub
    If m_ViewModel.Expediente Is Nothing Then Exit Sub
    
    ' 1. Cabecera y Títulos
    Me.Caption = "Concesión/Desviación SUB: " & m_ViewModel.Solicitud.codigoSolicitud & " (" & m_ViewModel.Expediente.Nemotecnico & ")"
    Me.lblTitulo.Caption = "Solicitud de Concesión/Desviación " & m_ViewModel.Expediente.Nemotecnico
    Me.txtEstado.value = m_ViewModel.EstadoNombre
    
    ' 2. Permisos del subformulario activo
    Dim bloque As enumBloqueFormulario
    bloque = FormulariosPadreAuxiliares.MapearBloquePorNombreSub(Me.NavigationSubform.Form.name, TipoForm_CDCASUB)
    
    Dim habilitar As Boolean
    Select Case bloque
        Case Bloque_Generales: habilitar = m_ViewModel.PuedeGuardar_Generales
        Case Bloque_Propuesta: habilitar = m_ViewModel.PuedeGuardar_Propuesta
        Case Bloque_Impacto: habilitar = m_ViewModel.PuedeGuardar_Impacto
        Case Bloque_AprobacionSuministrador: habilitar = m_ViewModel.PuedeGuardar_AprobacionSuministrador
        Case Bloque_DictamenRAC: habilitar = m_ViewModel.PuedeGuardar_DictamenRAC
        Case Bloque_DecisionFinal: habilitar = m_ViewModel.PuedeGuardar_DecisionFinal
        Case Else: habilitar = False
    End Select
    
    Call FormulariosPadreAuxiliares.HabilitarGuardarEnSubform(Me.NavigationSubform.Form, habilitar)
    Call HabilitarControlesSubformulario(Me.NavigationSubform.Form, habilitar)
    
    ' 3. Botones de Decisión (Excluyentes según rol y estado)
    
    Dim esEditor As Boolean
    esEditor = (rolUsuario = rol.Calidad Or rolUsuario = rol.Administrador)
    
    ' Botón Decisión Técnica (Estados intermedios)
    Dim mostrarBtnTecnica As Boolean
    mostrarBtnTecnica = (m_ViewModel.Solicitud.idEstadoInterno = estadoModificacion Or _
                         m_ViewModel.Solicitud.idEstadoInterno = estadoValidacion Or _
                         (m_ViewModel.Solicitud.idEstadoInterno = estadoDesarrolloTecnico And m_ViewModel.vm_revisionCalidadEstado = "RECHAZADO")) _
                        And esEditor

    
    ' 4. Banner de Rechazo (Aviso visual para el técnico)
    If m_ViewModel.vm_revisionCalidadEstado = "RECHAZADO" And m_ViewModel.Solicitud.idEstadoInterno = estadoModificacion Then
        Me.lblBannerRechazo.Visible = True
        Me.lblBannerRechazo.Caption = "RECHAZADO POR CALIDAD: " & m_ViewModel.vm_revisionCalidadComentarios
    Else
        Me.lblBannerRechazo.Visible = False
    End If

    Exit Sub
Errores:
    ManejadorErroresFormulario "frmDatosCDCASUB.ConfigurarYPoblarInterfaz", Me
End Sub

' ==========================================================================
' FUNCIONES AUXILIARES DE UI (Mapeo y Botones)
' ==========================================================================

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set m_ViewModel = Nothing
    Set m_CDCASUB_Srv = Nothing
    Set m_SolSrv = Nothing
    Set m_WFSrv = Nothing
End Sub

Private Sub cmdCerrar_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.name
End Sub

Private Sub cmdVolverResumen_Click()
    On Error GoTo Errores
    
    ' Spec-048 GAP: Solo foco al padre, sin triggers ni refrescos
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

' --- NAVEGACIÓN ---

Private Sub NavigationButton_Click()
    On Error GoTo Errores
    Call ConfigurarYPoblarInterfaz
    Call PoblarSubformActual
LimpiarYSalir:
    Exit Sub
Errores:
    ManejadorErroresFormulario "frmDatosCDCASUB.NavigationButton_Click", Me
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
    Resume Next
End Sub

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
    nombreSubform = FormulariosPadreAuxiliares.GetSubformNameFromTabName(nombreTab, TipoForm_CDCASUB)
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

' ==========================================================================
' GUARDADO
' ==========================================================================

Public Sub GuardarDesdeSubform(ByVal nombreSub As String, Optional ByVal silenciarErrorValidacion As Boolean = False)
    Dim bloque As enumBloqueFormulario
    Dim dataDB As DatosCDCASUB, dataUI As DatosCDCASUB
    Dim transicionRealizada As Boolean, huboAvanceAutomatico As Boolean
    Dim solicitudParaTransicion As Solicitud
    Dim estadoAnterior As Long
    
    ' Variables Transacción
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim gestionaTransaccion As Boolean
    
    On Error GoTo Errores
    Me.lblBloqueo.Visible = True: DoCmd.Hourglass True
    transicionRealizada = False: huboAvanceAutomatico = False
    
    ' --- DEFENSA CONTRA ERROR 91 (Spec-167) ---
    If m_ViewModel Is Nothing Then
        Err.Raise 91, "GuardarDesdeSubform", "m_ViewModel no está inicializado."
    End If
    If m_CDCASUB_Srv Is Nothing Then Set m_CDCASUB_Srv = New DatosCDCASUBServicio
    If m_ViewModel.Solicitud Is Nothing Then
        Err.Raise 91, "GuardarDesdeSubform", "m_ViewModel.Solicitud es Nothing."
    End If
    ' --------------------------------------------

    Set solicitudParaTransicion = m_ViewModel.Solicitud
    bloque = FormulariosPadreAuxiliares.MapearBloquePorNombreSub(nombreSub, TipoForm_CDCASUB)
    If bloque = 0 Then Err.Raise 513, "GuardarDesdeSubform", "Subformulario no reconocido."
    
    ' 1. Obtener Datos
    Set dataDB = m_CDCASUB_Srv.getDatosCDCASUB(solicitudParaTransicion.idSolicitud)
    
    ' 2. Preparar UI (Si no existe en BD, creamos nuevo, si existe, usamos ese como base)
    If Not dataDB Is Nothing Then
        Set dataUI = FormulariosPadreAuxiliares.ClonarEntidad(dataDB, TipoForm_CDCASUB)
    Else
        Set dataUI = New DatosCDCASUB
        dataUI.idSolicitud = solicitudParaTransicion.idSolicitud
    End If
    
    ' 3. Recoger
    CallByName Me.NavigationSubform.Form, "RecogerValores", VbMethod, m_ViewModel
    Call ExportarVM_aEntidad(bloque, m_ViewModel, dataUI)
    
    ' 4. Dirty Check
    If bloque = Bloque_AprobacionSuministrador Then
        If Not Me.NavigationSubform.Form.hayCambios() Then
            Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
            If Not silenciarErrorValidacion Then MsgBox "No se han detectado cambios.", vbInformation, "CONDOR"
            GoTo SalidaSinGuardar
        End If
    Else
        If Not dataDB Is Nothing Then
            If FormulariosPadreAuxiliares.SonEntidadesIguales(dataDB, dataUI) Then
                Dim esRacAprobado As Boolean
                esRacAprobado = (bloque = Bloque_DictamenRAC And dataUI.racDecision = "APROBADO")
                If Not esRacAprobado Then
                    Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
                    If Not silenciarErrorValidacion Then MsgBox "No se han detectado cambios.", vbInformation, "CONDOR"
                    GoTo SalidaSinGuardar
                End If
            End If
        End If
    End If
    
    ' --- TRANSACCIÓN RAC ---
    If bloque = Bloque_DictamenRAC Then
        ' Issue #17 (claim-without-proof): el formulario ya no abre DAO directo
        ' para Bloque_DictamenRAC. El servicio DatosCDCASUBServicio.GuardarDictamenRAC
        ' gestiona internamente BeginTrans/CommitTrans/Rollback cuando el parámetro
        ' db es Nothing (ver required check en el servicio para EsSolicitudEnValidacion).
        Call m_CDCASUB_Srv.GuardarDictamenRAC(dataUI)

        ' Feedback al usuario. Antes del fix, este MsgBox vivía después de
        ' CommitTrans en este mismo bloque. Ahora el CommitTrans está dentro
        ' del servicio, pero el feedback sigue siendo responsabilidad del
        ' formulario (la lógica condicional silenciarErrorValidacion se respeta).
        If Not silenciarErrorValidacion Then
            MsgBox "Datos guardados correctamente.", vbInformation, "CONDOR"
        End If
    Else
        ' --- GUARDADO ESTÁNDAR ---
        Select Case bloque
            Case Bloque_Generales: transicionRealizada = m_CDCASUB_Srv.GuardarDatosGenerales(dataUI, db)
            Case Bloque_Propuesta: Call m_CDCASUB_Srv.GuardarPropuesta(dataUI, db)
            Case Bloque_Impacto: Call m_CDCASUB_Srv.GuardarImpacto(dataUI, db)
            Case Bloque_AprobacionSuministrador: Call m_CDCASUB_Srv.GuardarAprobacionSuministrador(dataUI, db)
            Case Bloque_DecisionFinal: Call m_CDCASUB_Srv.GuardarDecisionFinal(dataUI, db)
        End Select
        
        ' Spec-151: Recargar checklist tras guardar AprobSum en Modificación
        If bloque = Bloque_AprobacionSuministrador And solicitudParaTransicion.idEstadoInterno = estadoModificacion Then
            If FormularioAbierto("frmGestionSolicitud") Then
                Call Forms("frmGestionSolicitud").CargarDatosSolicitud(solicitudParaTransicion.idSolicitud, mantenerVistaActual:=True)
            End If
        End If
        
        ' >>> GENERAR TRAZA DE AUDITORÍA (Vital para desbloqueo técnico) <<<
        If Not transicionRealizada Then
            Dim logSrv As New LogServicio
            Dim campoTraza As String
            Select Case bloque
                Case Bloque_Propuesta: campoTraza = "descripcionImpactoNC"
                Case Bloque_Impacto: campoTraza = "impactoCoste"
                Case Bloque_Generales: campoTraza = "refSuministrador"
                Case Else: campoTraza = ""
            End Select
            If campoTraza <> "" Then
                Call logSrv.RegistrarCambio("tbDatosCDCASUB", solicitudParaTransicion.idSolicitud, campoTraza, "...", "Actualizado", "EDICION_TECNICA")
            End If
        End If
        ' >>> FIN TRAZA <<<
        
        Set solicitudParaTransicion = m_SolSrv.getSolicitudPorID(m_ViewModel.Solicitud.idSolicitud)
        
        ' Avance Generales
        If bloque = Bloque_Generales And (solicitudParaTransicion.idEstadoInterno = estadoRegistro Or transicionRealizada) Then
            If transicionRealizada Then solicitudParaTransicion.idEstadoInterno = estadoRegistro
            
            ' Spec-137: Guardar estado anterior para evitar msgbox de Motivos cuando Calidad pasa de REGISTRO a DESARROLLO TÉCNICO
            estadoAnterior = solicitudParaTransicion.idEstadoInterno

            If MsgBox("Datos Generales guardados." & vbCrLf & "¿Asignar tarea al Técnico ahora?", vbQuestion + vbYesNo, "CONDOR") = vbYes Then
                Call m_WFSrv.EjecutarTransicion(solicitudParaTransicion, estadoDesarrolloTecnico, m_ObjUsuarioActivo)
                solicitudParaTransicion.idEstadoInterno = estadoDesarrolloTecnico
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
            tieneDetalle = m_CDCASUB_Srv.EsDetalleCompleto(solicitudParaTransicion.idSolicitud)
            
            Dim tieneMotivos As Boolean
            tieneMotivos = m_CDCASUB_Srv.EsMotivosCompleto(solicitudParaTransicion.idSolicitud)
            
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
        
        ' Avance Tecnico
        Dim yaSePreguntoPromocion As Boolean
        yaSePreguntoPromocion = False
        If (bloque = Bloque_Propuesta Or bloque = Bloque_Impacto) And solicitudParaTransicion.idEstadoInterno = estadoDesarrolloTecnico Then
            Dim cdcaSubServCheckTec As New DatosCDCASUBServicio
            If cdcaSubServCheckTec.EsParteTecnicaCompleta(solicitudParaTransicion.idSolicitud) Then
                
                Dim msgPromocion As String
                Dim respTec As VbMsgBoxResult
                If solicitudParaTransicion.revisionCalidadEstado = "RECHAZADO" Then
                    msgPromocion = "Ha corregido los datos técnicos." & vbCrLf & vbCrLf & _
                                   "¿Desea enviar la solicitud de nuevo a Calidad para su validación?"
                Else
                    msgPromocion = "Ha completado todos los datos técnicos obligatorios." & vbCrLf & vbCrLf & _
                                   "¿Desea finalizar su tarea y enviar la solicitud a Calidad?" & vbCrLf & _
                                   "(Esto avanzará el estado a 'MODIFICACIÓN')."
                End If
                
                respTec = MsgBox(msgPromocion, vbQuestion + vbYesNo + vbDefaultButton1, "CONDOR")
                yaSePreguntoPromocion = True
                
                If respTec = vbYes Then
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
    End If
    
    If Not huboAvanceAutomatico And Not transicionRealizada And bloque <> Bloque_DictamenRAC Then
        If Not silenciarErrorValidacion And Not yaSePreguntoPromocion Then MsgBox "Datos guardados correctamente.", vbInformation, "CONDOR"
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
    GoTo LimpiarYSalir

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
    ManejadorErroresFormulario "frmDatosCDCASUB.GuardarDesdeSubform", Me
End Sub
Public Sub ProcesarCierreFormalizacion()
    Dim wfServ As New WorkflowServicio
    Dim subUI As DatosCDCASUB, subDB As DatosCDCASUB
    Dim fDialog As FileDialog
    Dim rutaSeleccionada As String
    
    On Error GoTo Errores
    
    If MsgBox("Para APROBAR definitivamente la solicitud, es obligatorio adjuntar el documento de Cierre." & vbCrLf & vbCrLf & _
              "Al pulsar Aceptar, se abrira una ventana para que seleccione el archivo PDF firmado.", _
              vbInformation + vbOKCancel, "CONDOR") = vbCancel Then
        Exit Sub
    End If
    
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
    
    Set subDB = m_CDCASUB_Srv.getDatosCDCASUB(m_ViewModel.Solicitud.idSolicitud)
    Set subUI = FormulariosPadreAuxiliares.ClonarEntidad(subDB, TipoForm_CDCASUB)
    
    subUI.decisionFinal = m_ViewModel.vm_decisionFinal
    subUI.observacionesFinales = m_ViewModel.vm_obsDecisionFinal
    subUI.NombreFirmanteFinal = m_ViewModel.vm_NombreFirmanteFinal
    
    Call wfServ.EjecutarCierreFormalizacion(m_ViewModel.Solicitud, subUI, m_ObjUsuarioActivo, rutaSeleccionada)
    
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
' ==========================================================================
' MÉTODOS DE MAPEO (Delegados al servicio)
' ==========================================================================

Private Sub ExportarVM_aEntidad(ByVal bloque As enumBloqueFormulario, ByVal vm As DatosCDCASUBViewModel, ByRef ent As DatosCDCASUB)
    ' Delegado al servicio - código de negocio extraído
    Call m_CDCASUB_Srv.ExportarVM_aEntidad(bloque, vm, ent)
End Sub

' ==========================================================================
' RECARGA Y MAPEO (CORE - Delegado al servicio)
' ==========================================================================

Public Sub RecargarViewModel(ByVal idSolicitud As Long)
    Dim vm As DatosCDCASUBViewModel
    
    ' --- DEFENSA CONTRA ERROR 91 ---
    If m_SolSrv Is Nothing Then Set m_SolSrv = New SolicitudServicio
    If m_CDCASUB_Srv Is Nothing Then Set m_CDCASUB_Srv = New DatosCDCASUBServicio
    If m_WFSrv Is Nothing Then Set m_WFSrv = New WorkflowServicio
    ' -------------------------------
    
    On Error GoTo Errores
    
    ' Obtener ViewModel completo del servicio (lógica de negocio centralizada)
    Set vm = m_CDCASUB_Srv.ObtenerViewModelCompleto(idSolicitud)
    
    ' Asignar al estado del form (única responsabilidad del form)
    Set m_ViewModel = vm
    
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "frmDatosCDCASUB.RecargarViewModel"
End Sub
' ==========================================================================
' MÉTODOS DE MAPEO (Delegados al servicio)
' ==========================================================================

Private Sub ImportarEntidad_aVM(ByRef ent As DatosCDCASUB, ByRef vm As DatosCDCASUBViewModel)
    ' Delegado al servicio - código de negocio extraído
    Call m_CDCASUB_Srv.ImportarEntidad_aVM(ent, vm)
End Sub

Public Sub NotificarCambioDeEstado()
    On Error GoTo Errores
    
    RaiseEvent SolicitudHaCambiado
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "frmDatosCDCASUB.NotificarCambioDeEstado", Me
End Sub
