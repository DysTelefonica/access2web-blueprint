
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_frmDatosPC.cls
' RESPONSABILIDAD: Orquestador principal para la edición de datos de una Propuesta de Cambio (PC).
'                  Gestiona la Navegación, el guardado y el flujo de trabajo.
' VERSIÓN: Final con flujo de cierre guiado y correcciones de errores de reentrada.
' ==========================================================================

'=== Estado (Vista desconectada con VM) ===
Private m_ViewModel As DatosPCViewModel
Private m_PCSrv As DatosPCServicio
Private m_SolSrv As SolicitudServicio
Private m_WFSrv As WorkflowServicio
Public Event SolicitudHaCambiado()

Private Sub cmdCerrar_Click()
    DoCmd.Close acForm, Me.name
End Sub

' EN: Form_frmDatosPC.cls


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
' ==========================================================================
' MÓDULO: Form_frmDatosPC.cls
' ACCIÓN: Simplificar Form_Load para delegar la configuración de UI a sus eventos.
' ==========================================================================
Private Sub Form_Load()
    On Error GoTo Errores
    Ajustar Me
    
    DoCmd.Hourglass True
    
    ' Inicializamos los servicios
    Set m_PCSrv = New DatosPCServicio
    Set m_SolSrv = New SolicitudServicio
    Set m_WFSrv = New WorkflowServicio

    Dim idSol As Long
    idSol = CLng(Nz(Me.OpenArgs, 0))
    If idSol <= 0 Then Err.Raise 513, "Form_Load", "No se recibió un idSolicitud válido."

    ' ÚNICA RESPONSABILIDAD: Cargar los datos en el ViewModel.
    Call RecargarViewModel(idSol)
    
    ' La configuración de la UI se hará automáticamente cuando el NavigationControl
    ' cargue la primera pestaña y dispare su evento Click.
    ' CORRECCIÓN: Usar NavegarAPestana para asegurar la pestaña correcta según estado
    Dim paginaInicial As String
    paginaInicial = m_WFSrv.getPaginaActivaPC(m_ViewModel.Solicitud.idEstadoInterno)
    Call NavegarAPestana(paginaInicial)
    
LimpiarYSalir:
    DoCmd.Hourglass False
    Exit Sub
Errores:
    ManejadorErroresFormulario "frmDatosPC.Form_Load", Me
    DoCmd.Close acForm, Me.name
    Resume LimpiarYSalir
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' No ponemos LÓGICA aquí porque puede conflictos con validaciones de Access.
    ' Confiamos en los botones de Navegación.
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
    nombreSubform = FormulariosPadreAuxiliares.GetSubformNameFromTabName(nombreTab, TipoForm_PC)
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
    ManejadorErroresFormulario "frmDatosPC.NavigationButton_Click", Me
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
    Dim pcDB As DatosPC, pcUI As DatosPC
    Dim transicionAutomaticaRealizada As Boolean
    Dim solicitudParaTransicion As Solicitud
    Dim huboAvanceAutomatico As Boolean
    Dim estadoAnterior As Long
    
    ' Variables TransACCIÓN
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
    If m_PCSrv Is Nothing Then Set m_PCSrv = New DatosPCServicio
    If m_ViewModel.Solicitud Is Nothing Then
        Err.Raise 91, "GuardarDesdeSubform", "m_ViewModel.Solicitud es Nothing."
    End If
    ' --------------------------------------------

    Set solicitudParaTransicion = m_ViewModel.Solicitud

    ' 1. Determinar bloque
    bloque = FormulariosPadreAuxiliares.MapearBloquePorNombreSub(nombreSub, TipoForm_PC)
    If bloque = 0 Then Err.Raise 513, "GuardarDesdeSubform", "Subformulario no reconocido."

    ' 2. OBTENER ESTADO ACTUAL BD
    Set pcDB = m_PCSrv.getDatosPC(solicitudParaTransicion.idSolicitud)
    If Not pcDB Is Nothing Then
        Set pcUI = FormulariosPadreAuxiliares.ClonarEntidad(pcDB, TipoForm_PC)
    Else
        Set pcUI = New DatosPC
        pcUI.idSolicitud = solicitudParaTransicion.idSolicitud
    End If
    
    ' 3. RECOGER Y VALIDAR UI
    CallByName Me.NavigationSubform.Form, "RecogerValores", VbMethod, m_ViewModel
    Call ExportarVM_aEntidad(bloque, m_ViewModel, pcUI)
    
    ' 4. DIRTY CHECK
    If Not pcDB Is Nothing Then
        If FormulariosPadreAuxiliares.SonEntidadesIguales(pcDB, pcUI) Then
            Dim esRacAprobado As Boolean
            esRacAprobado = (bloque = Bloque_DictamenRAC And pcUI.racDecision = "APROBADO")
            
            ' Si es RAC Aprobado, seguimos adelante para intentar el avance de fase
            If Not esRacAprobado Then
                Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
                If Not silenciarErrorValidacion Then MsgBox "No se han detectado cambios.", vbInformation, "CONDOR"
                GoTo SalidaSinGuardar
            End If
        End If
    End If

    ' =========================================================================================
    ' LÓGICA DE TRANSACCIÓN PARA DICTAMEN RAC (ATOMICIDAD: GUARDAR + AVANZAR) - PC
    ' =========================================================================================
    If bloque = Bloque_DictamenRAC Then
        Set ws = getWorkspace()
        Set db = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
        gestionaTransaccion = True
        
        ' A. Guardar Datos
        Call m_PCSrv.GuardarDictamenRAC(pcUI, db)
        
        ' D. Commit de TODO
        ws.CommitTrans
        gestionaTransaccion = False
        
        If Not silenciarErrorValidacion Then
            MsgBox "Datos guardados correctamente.", vbInformation, "CONDOR"
        End If

    Else
        ' =========================================================================================
        ' LÓGICA ESTANDAR PARA OTROS BLOQUES
        ' =========================================================================================
        Select Case bloque
            Case Bloque_Generales:               transicionAutomaticaRealizada = m_PCSrv.GuardarDatosGenerales(pcUI)
            Case Bloque_Propuesta:               Call m_PCSrv.GuardarPropuesta(pcUI)
            Case Bloque_Impacto:                 Call m_PCSrv.GuardarImpacto(pcUI)
            Case Bloque_AprobacionSuministrador
                If Not Me.NavigationSubform.Form.hayCambios() Then
                    Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
                    If Not silenciarErrorValidacion Then MsgBox "No se han detectado cambios.", vbInformation, "CONDOR"
                    GoTo SalidaSinGuardar
                End If
                
                Call m_PCSrv.GuardarAprobacionSuministrador(pcUI)
            Case Bloque_DecisionFinal:           Call m_PCSrv.GuardarDecisionFinal(pcUI)
        End Select
        
        ' Spec-151: Recargar checklist tras guardar AprobSum en Modificación
        If bloque = Bloque_AprobacionSuministrador And solicitudParaTransicion.idEstadoInterno = estadoModificacion Then
            If FormularioAbierto("frmGestionSolicitud") Then
                Call Forms("frmGestionSolicitud").CargarDatosSolicitud(solicitudParaTransicion.idSolicitud, mantenerVistaActual:=True)
            End If
        End If
        
       ' LÓGICA de avance Datos Generales (PC)
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
            
            ' Scenario 3: Ambas partes completas - preguntar promoción
            If tieneMotivos And tieneDetalle Then
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
        
        ' =========================================================================================
        ' Lógica existente de promoción a Calidad (se ejecuta si no se activó SPEC-132)
        ' =========================================================================================
        Dim yaSePreguntoPromocion As Boolean
        yaSePreguntoPromocion = False
        If (bloque = Bloque_Propuesta Or bloque = Bloque_Impacto) And solicitudParaTransicion.idEstadoInterno = estadoDesarrolloTecnico Then
            
            Dim pcServCheckTec As New DatosPCServicio
            
            ' Validamos contra la BD si el conjunto de datos Técnicos está completo
            If pcServCheckTec.EsParteTecnicaCompleta(solicitudParaTransicion.idSolicitud) Then
                
                Dim respTec As VbMsgBoxResult
                Dim msgPromocion As String
                
                If solicitudParaTransicion.revisionCalidadEstado = "RECHAZADO" Then
                    msgPromocion = "Ha corregido los datos técnicos." & vbCrLf & vbCrLf & _
                                 "¿Desea enviar la solicitud de nuevo a Calidad para su validacón?"
                Else
                    msgPromocion = "Ha completado todos los datos técnicos obligatorios." & vbCrLf & vbCrLf & _
                                 "¿Desea finalizar su tarea y enviar la solicitud a Calidad?" & vbCrLf & _
                                 "(Esto avanzará el estado a 'MODIFICACIÓN')."
                End If
                
                respTec = MsgBox(msgPromocion, vbQuestion + vbYesNo + vbDefaultButton1, "CONDOR")
                yaSePreguntoPromocion = True
                
                If respTec = vbYes Then
                    Me.lblBloqueo.Visible = True: DoCmd.Hourglass True
                    
                    ' Ejecutamos la transición al estado 4 (MODIFICACIÓN)
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

    ' --- aquí ES DONDE ANTES ESTABA EL BLOQUE DE "RETORNO DE CORRECCIÓN". YA NO está. ---

    ' 7. FEEDBACK GENÉRICO
    If Not huboAvanceAutomatico And Not transicionAutomaticaRealizada And bloque <> Bloque_DictamenRAC Then
        If Not silenciarErrorValidacion And Not yaSePreguntoPromocion Then MsgBox "Datos guardados correctamente.", vbInformation, "CONDOR"
    End If
    
    ' 8. REFRESCO FINAL
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
    ' Rollback si estábamos en transACCIÓN
    If gestionaTransaccion And Not ws Is Nothing Then
        ws.Rollback
    End If

    Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
    If Err.Number = 513 And silenciarErrorValidacion Then Exit Sub
    ManejadorErroresFormulario "frmDatosPC.GuardarDesdeSubform", Me
End Sub

' --- FUNCIONES AUXILIARES ELIMINADAS (migradas a FormulariosPadreAuxiliares.bas) ---
' EN: Form_frmDatosPC.cls
Public Sub ConfigurarYPoblarInterfaz()
    On Error GoTo Errores
    
    ' --- Validaciones iniciales ---
    If Me.NavigationSubform Is Nothing Then Exit Sub
    If Me.NavigationSubform.Form Is Nothing Then Exit Sub
    If m_ViewModel Is Nothing Then Exit Sub
    If m_ViewModel.Expediente Is Nothing Then Exit Sub
    
    ' --- Rellenar cabecera ---
    Me.Caption = "Propuesta de Cambio: " & m_ViewModel.Solicitud.codigoSolicitud & " (" & m_ViewModel.Expediente.Nemotecnico & ")"
    Me.lblTitulo.Caption = "Solicitud de Propuesta de Cambio " & m_ViewModel.Expediente.Nemotecnico
    Me.txtEstado.value = m_ViewModel.EstadoNombre
    
    ' --- LÓGICA de permisos del subformulario activo ---
    Dim bloque As enumBloqueFormulario
    bloque = FormulariosPadreAuxiliares.MapearBloquePorNombreSub(Me.NavigationSubform.Form.name, TipoForm_PC)
    
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
    
    ' ==========================================================================================
    ' Gestión DE BOTONES DE Decisión (LÓGICA EXCLUYENTE)
    ' ==========================================================================================
    Dim estadoID As Long
    Dim esRolCalidad As Boolean
    
    estadoID = m_ViewModel.Solicitud.idEstadoInterno
    esRolCalidad = (rolUsuario = rol.Calidad Or rolUsuario = rol.Administrador)
    
    ' 1. BOTÓN Decisión Técnica (Fases 4 y 5)
    ' ---------------------------------------
    Dim mostrarBtnTecnica As Boolean
    mostrarBtnTecnica = (estadoID = estadoModificacion Or _
                         estadoID = estadoValidacion Or _
                         (estadoID = estadoDesarrolloTecnico And m_ViewModel.vm_revisionCalidadEstado = "RECHAZADO")) _
                        And esRolCalidad
    
    

    
    
    ' ==========================================================================================
    
    ' --- Banner de Rechazo ---
    If m_ViewModel.vm_revisionCalidadEstado = "RECHAZADO" And estadoID = estadoModificacion Then
        Me.lblBannerRechazo.Visible = True
        
        Me.lblBannerRechazo.Caption = "RECHAZADO POR CALIDAD: " & m_ViewModel.vm_revisionCalidadComentarios
    Else
        Me.lblBannerRechazo.Visible = False
        
    End If
    
LimpiarYSalir:
    Exit Sub
Errores:
    ManejadorErroresFormulario "frmDatosPC.ConfigurarYPoblarInterfaz", Me
    Resume LimpiarYSalir
End Sub
' ==========================================================================
' MÓDULO: Form_frmDatosPC.cls
' ACCIÓN: Refactorizado para delegar carga de VM al servicio
' ==========================================================================
Public Sub RecargarViewModel(ByVal idSolicitud As Long)
    Dim vm As DatosPCViewModel
    
    ' --- DEFENSA CONTRA ERROR 91 (Pérdida de referencias) ---
    If m_SolSrv Is Nothing Then Set m_SolSrv = New SolicitudServicio
    If m_PCSrv Is Nothing Then Set m_PCSrv = New DatosPCServicio
    If m_WFSrv Is Nothing Then Set m_WFSrv = New WorkflowServicio
    ' --------------------------------------------------------

    On Error GoTo Errores
    
    ' Obtener ViewModel completo del servicio (lógica de negocio centralizada)
    Set vm = m_PCSrv.ObtenerViewModelCompleto(idSolicitud)
    
    ' Asignar al estado del form (única responsabilidad del form)
    Set m_ViewModel = vm
    
    Exit Sub
Errores:
    Dim e As New CondorError
    e.Create Err.Number, Err.description, "frmDatosPC.RecargarViewModel"
    e.AddToCallStack "idSolicitud: " & idSolicitud
    e.Raise
End Sub

' ==========================================================================
' MÓDULO: Form_frmDatosPC.cls
' ACCIÓN: Refactorizado para delegar mapeo VM a DatosPCServicio
' ==========================================================================

' ==========================================================================
' MÉTODOS DE MAPEO (Delegados al servicio)
' ==========================================================================
Private Sub ImportarEntidad_aVM(ByRef pc As DatosPC, ByRef vm As DatosPCViewModel)
    ' Delegado al servicio - código de negocio extraído
    Call m_PCSrv.ImportarEntidad_aVM(pc, vm)
End Sub

Private Sub ExportarVM_aEntidad(ByVal bloque As enumBloqueFormulario, ByRef vm As DatosPCViewModel, ByRef pc As DatosPC)
    ' Delegado al servicio - código de negocio extraído
    Call m_PCSrv.ExportarVM_aEntidad(bloque, vm, pc)
End Sub

' ==========================================================================
' MÓDULO: Form_frmDatosPC.cls
' ACCIÓN: Gestión centralizada de botones de Decisión (Técnica vs RAC).
' ==========================================================================
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

' ==========================================================================
' MÓDULO: Form_frmDatosPC.cls
' ACCIÓN: añadir este nuevo método público para la comunicación desde subformularios.
' ==========================================================================
Public Sub NotificarCambioDeEstado()
    ' RESPONSABILIDAD: Recibir una notificación de un subformulario y propagarla
    '                  hacia el exterior a trav?s de su propio evento público.
    
    On Error GoTo Errores
    
    ' Lanza el evento de este formulario, que puede ser escuchado por frmGestionSolicitud
    RaiseEvent SolicitudHaCambiado
    
    Exit Sub
Errores:
    ' No deber?a fallar, pero por seguridad...
    ManejadorErroresFormulario "frmDatosPC.NotificarCambioDeEstado", Me
End Sub

Public Sub ProcesarCierreFormalizacion()
    Dim wfServ As New WorkflowServicio
    Dim pcUI As DatosPC, pcDB As DatosPC
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
        MsgBox "Operacion cancelada. No se ha seleccionado el documento firmado.", vbExclamation, "CONDOR - Cierre"
        Exit Sub
    End If
    
    Me.lblBloqueo.Visible = True
    Me.lblBloqueo.Caption = "Procesando cierre, guardando archivo y firmando..."
    DoCmd.Hourglass True
    
    ' 3. Preparar Datos
    Set pcDB = m_PCSrv.getDatosPC(m_ViewModel.Solicitud.idSolicitud)
    Set pcUI = FormulariosPadreAuxiliares.ClonarEntidad(pcDB, TipoForm_PC)
    
    pcUI.decisionFinal = m_ViewModel.vm_decisionFinal
    pcUI.obsDecisionFinal = m_ViewModel.vm_obsDecisionFinal
    pcUI.NombreAutoridadDiseno = m_ViewModel.vm_NombreAutoridadDiseno
    pcUI.NombreFirmanteFinal = m_ViewModel.vm_NombreFirmanteFinal
    
    ' 4. Llamada a la Transacción Maestra (Pasando la ruta del archivo)
    Call wfServ.EjecutarCierreFormalizacion(m_ViewModel.Solicitud, pcUI, m_ObjUsuarioActivo, rutaSeleccionada)
    
    ' 5. éxito
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







