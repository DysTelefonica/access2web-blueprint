

Option Compare Database
Option Explicit


Public Event SolicitudModificada()
Public Event SolicitudEliminada()
Public Event SolicitudCambiada(ByVal tipoCambio As String)

Private m_ViewModel As SolicitudViewModel
Private m_solServ As SolicitudServicio

Private Enum VistaWeb
    VistaResumen = 1
    VistaCambioEstado = 2
    VistaHistorial = 3
End Enum
Private m_VistaActual As VistaWeb

' Variable de módulo para el script de inyección
Private m_jsScriptParaInyectar As String
Private m_AccionPendiente As String
Private m_ParametroPendiente As Long
Private m_ParametroTextoPendiente As String
Private m_TimerIntervalAnterior As Long
Private m_ModoRechazoCalidad As String

Private WithEvents m_frmDatosPC As Form_frmDatosPC



Private WithEvents m_frmDatosCDCA As Form_frmDatosCDCA



Private WithEvents m_frmDatosCDCASUB As Form_frmDatosCDCASUB



Private WithEvents m_frmDatosPCSUB As Form_frmDatosPCSUB



Private Function EsRenderizadoFicheroActivo() As Boolean
    Dim flagValue As String
    On Error Resume Next
    flagValue = CStr(Application.TempVars("UsarRenderizadoFichero"))
    If Err.Number <> 0 Then
        Err.Clear
        EsRenderizadoFicheroActivo = False
        Exit Function
    End If
    flagValue = UCase$(flagValue)
    EsRenderizadoFicheroActivo = (flagValue = "SI" Or flagValue = "SÍ")
End Function




' ==========================================================================
' MÓDULO: Form_frmGestionSolicitud.cls
' ==========================================================================
Private Sub Form_Load()
    On Error GoTo Errores
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    DoEvents
    Ajustar Me
    Set m_solServ = New SolicitudServicio
    Dim idSol As Long
    idSol = CLng(Nz(Me.OpenArgs, 0))
    If idSol <= 0 Then Err.Raise 513, "Form_Load", "No se proporcion un ID de Solicitud válido."
    Call CargarDatosSolicitud(idSol)

    ' --- SOLUCIÓN EDGE: ACTIVAR TIMER PARA POLLING ---
    If EsRenderizadoFicheroActivo Then
        Me.TimerInterval = 300 ' 300ms para sondeo
    End If
    ' -------------------------------------------------

LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.Form_Load")
    DoCmd.Close acForm, Me.name
    Exit Sub
End Sub

Public Sub CargarDatosSolicitud(ByVal idSolicitud As Long, Optional ByVal mantenerVistaActual As Boolean = False, Optional ByVal soloMemoria As Boolean = False)
    On Error GoTo Errores
    Set m_ViewModel = m_solServ.getSolicitudViewModelPorID(idSolicitud)
    If m_ViewModel Is Nothing Then Err.Raise 513, "CargarDatos", "No se pudo cargar la información."
    If Not soloMemoria Then Me.Caption = "Gestión de Solicitud: " & m_ViewModel.Solicitud.codigoSolicitud
    
    ' LÓGICA DE PERSISTENCIA DE VISTA (Spec-004)
    ' Si se solicita mantener la vista y existe una vista válida previa, la restauramos.
    ' Esto es crítico para "descongelar" el control Edge tras acciones JS.
    If Not soloMemoria Then
        If mantenerVistaActual And m_VistaActual <> 0 Then
            Call MostrarVista(m_VistaActual)
        Else
            Call MostrarVista(VistaResumen)
        End If
    End If
    
    If Not soloMemoria Then Call PoblarControlesDesdeViewModel
    Exit Sub
Errores:
    Err.Raise Err.Number, "CargarDatosSolicitud", Err.description
End Sub

' ==========================================================================
' MÓDULO: Form_frmGestionSolicitud.cls
' ACCIÓN: MostrarVista (SOPORTE DUAL: Memoria vs Fichero según Flag)
' ==========================================================================
' ==========================================================================
' MÓDULO: Form_frmGestionSolicitud.cls
' ACCIÓN: MostrarVista (Adaptado para Control Edge con https://msaccess/)
' ==========================================================================
Private Sub MostrarVista(ByVal vista As VistaWeb)
    Dim wfServ As New WorkflowServicio
    Dim html As String
    Dim rutaTemporal As String
    Dim stream As Object
    
    On Error GoTo Errores
    
    ' Spec-018: Refresco delegado a generación de ruta única (Anti-Caché Estricto)
    ' Se ha eliminado about:blank por incompatibilidad con Reglas de Oro (Sincronía)
    DoEvents
    
    m_VistaActual = vista
    
    ' 1. Generar el HTML
    Select Case vista
        Case VistaResumen
            html = wfServ.GenerarHTML_FichaResumen(m_ViewModel.Solicitud.idSolicitud)
        Case VistaCambioEstado
            html = wfServ.GenerarHTML_VisualizadorDeEstado(m_ViewModel.Solicitud.idSolicitud)
        Case VistaHistorial
            html = wfServ.GenerarHTML_GraficoTiempos(m_ViewModel.Solicitud.idSolicitud)
    End Select
    
    ' Actualizar botones
    Me.cmdCambiarEstado.Caption = IIf(m_VistaActual = VistaCambioEstado, "Ver Ficha Resumen", "Cambiar Estado")
    Me.cmdVerHistorial.Caption = IIf(m_VistaActual = VistaHistorial, "Ver Ficha Resumen", "Ver Tiempos")

    ' ==========================================================================
    ' LÓGICA ESPECÍFICA PARA CONTROL EDGE (https://msaccess/)
    ' ==========================================================================
    If EsRenderizadoFicheroActivo Then
        
        ' 2. Definir ruta física (Debe estar en una ubicación de confianza)
    ' TRUCO: Añadimos timestamp con milisegundos y RANDOM para evitar caché del navegador (Spec-014)
    Dim timestamp As String
    Randomize
    timestamp = Format(Now, "hhmmss") & "_" & Right(Format(Timer, "0.00"), 2) & "_" & Int((9999 - 1000 + 1) * Rnd + 1000)
    rutaTemporal = Environ("Temp") & "\CONDOR_Vista_" & m_ViewModel.Solicitud.idSolicitud & "_" & timestamp & ".html"
        
        ' 3. Guardar archivo (UTF-8)
        ' NOTA: En Edge NO hace falta el comentario "MOTW" (saved from url...),
        ' el protocolo msaccess ya le dice que es seguro.
        Dim ts As Object, bs As Object, bytes As Variant
        Set ts = CreateObject("ADODB.Stream")
        ts.Type = 2 ' adTypeText
        ts.Charset = "utf-8"
        ts.Open
        ts.WriteText html
        ts.Position = 0
        ts.Type = 1 ' adTypeBinary
        bytes = ts.Read
        ts.Close
        
        If IsArray(bytes) Then
            Dim start As Long, i As Long, out() As Byte
            start = 0
            If UBound(bytes) >= 2 Then
                If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then start = 3
            End If
            ReDim out(UBound(bytes) - start)
            For i = start To UBound(bytes)
                out(i - start) = bytes(i)
            Next i
            Set bs = CreateObject("ADODB.Stream")
            bs.Type = 1 ' adTypeBinary
            bs.Open
            bs.Write out
            bs.SaveToFile rutaTemporal, 2 ' adSaveCreateOverWrite
            bs.Close
        End If
        
        ' 4. CONSTRUIR LA URL ESPECIAL DE ACCESS
        Dim rutaNavegacion As String
        ' Transformamos: "C:\Temp\archivo.html" -> "https://msaccess/C:\Temp\archivo.html"
        rutaNavegacion = "https://msaccess/" & rutaTemporal
        
        ' 5. Navegar
        ' IMPORTANTE: El control Edge usa .Navigate, NO .Navigate2
        On Error Resume Next
        
        ' Spec-014: Retiramos about:blank porque causa condiciones de carrera (pantalla blanca)
        ' La unicidad del nombre de archivo (timestamp) ya evita la caché.
        DoEvents
        Me.Repaint
        Me.webInfo.Navigate rutaNavegacion
        
        If Err.Number <> 0 Then
            Dim logServEdge As New LogServicio
            logServEdge.RegistrarError Err.Number, "Error al navegar en Edge: " & Err.description, "Form_frmGestionSolicitud", "cmdCambiarEstado_Click", ""
        End If
        On Error GoTo Errores
        
        m_jsScriptParaInyectar = ""
        
    Else
        ' MODO COMPATIBILIDAD (Inyección en memoria)
        html = Replace(Replace(Replace(html, "\", "\\"), "'", "\'"), vbCrLf, "\n")
        m_jsScriptParaInyectar = "document.documentElement.innerHTML = '" & html & "';"
        Me.webInfo.Navigate "about:blank"
    End If

LimpiarYSalir:
    Set stream = Nothing
    Set wfServ = Nothing
    Exit Sub

Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.MostrarVista")
    Resume LimpiarYSalir
End Sub
' --- EVENTOS Y ACCIONES DE LOS BOTONES PRINCIPALES ---

Private Sub cmdCambiarEstado_Click()
    On Error GoTo Errores
    If m_ViewModel.Solicitud Is Nothing Then Exit Sub
    If m_VistaActual = VistaCambioEstado Then
        Call MostrarVista(VistaResumen)
    Else
        Call MostrarVista(VistaCambioEstado)
    End If
LimpiarYSalir:
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.cmdCambiarEstado_Click")
    Resume LimpiarYSalir
End Sub

Private Sub cmdVerHistorial_Click()
    On Error GoTo Errores
    If m_ViewModel.Solicitud Is Nothing Then Exit Sub
    If m_VistaActual = VistaHistorial Then
        Call MostrarVista(VistaResumen)
    Else
        Call MostrarVista(VistaHistorial)
    End If
LimpiarYSalir:
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.cmdVerHistorial_Click")
    Resume LimpiarYSalir
End Sub

Private Sub cmdRefrescarWeb_Click()
    On Error GoTo Errores
    Call ForzarResetWebView
LimpiarYSalir:
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.cmdRefrescarWeb_Click")
    Resume LimpiarYSalir
End Sub

Public Sub ForzarResetWebView()
    On Error GoTo Errores
    If m_ViewModel Is Nothing Then Exit Sub
    If m_ViewModel.Solicitud Is Nothing Then Exit Sub
    Me.webInfo.Navigate "about:blank"
    DoEvents
    DoEvents
    DoEvents
    m_jsScriptParaInyectar = ""
    Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud)
    Me.Repaint
LimpiarYSalir:
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.ForzarResetWebView")
    Resume LimpiarYSalir
End Sub

' ==========================================================================
' MÓDULO: Form_frmGestionSolicitud.cls
' ACCIÓN: REEMPLAZO de cmdVerDatos_Click para gestionar la visibilidad de formularios ocultos.
' ==========================================================================
Private Sub cmdVerDatos_Click()
    Dim formName As String
    On Error GoTo Errores

    If m_ViewModel Is Nothing Then Exit Sub
    
    ' 1. Determinar el nombre del formulario (sin cambios)
    Select Case m_ViewModel.Solicitud.tipoSolicitud
        Case "PC": formName = "frmDatosPC"
        Case "CD_CA": formName = "frmDatosCDCA"
        Case "CD_CA_SUB": formName = "frmDatosCDCASUB"
        Case "PC_SUB": formName = "frmDatosPCSUB"
        Case Else
            MsgBox "No hay un formulario de datos específico para este tipo de solicitud.", vbInformation, "CONDOR"
            Exit Sub
    End Select
    
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    DoEvents
    
    ' --- INICIO DE LA LÓGICA REFACTORIZADA ---
    
    Dim frm As Form
    Dim formularioExiste As Boolean
    
    On Error Resume Next
    Set frm = Forms(formName)
    formularioExiste = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    
    If formularioExiste Then
        If Not frm.Visible Then
            DoCmd.SelectObject acForm, formName, False
            DoCmd.Restore
            frm.Visible = True
        End If
        DoCmd.SelectObject acForm, formName, True
    Else
        On Error Resume Next
        DoCmd.OpenForm formName, acNormal, , , , acWindowNormal, CStr(m_ViewModel.Solicitud.idSolicitud)
        
        If Err.Number <> 0 Then
            On Error GoTo 0
            GoTo LimpiarYSalir
        End If
        
        On Error Resume Next
        Select Case formName
            Case "frmDatosPC": Set m_frmDatosPC = Forms(formName)
            Case "frmDatosCDCA": Set m_frmDatosCDCA = Forms(formName)
            Case "frmDatosCDCASUB": Set m_frmDatosCDCASUB = Forms(formName)
            Case "frmDatosPCSUB": Set m_frmDatosPCSUB = Forms(formName)
        End Select
        
        If Err.Number <> 0 Then
            Err.Clear
        End If
        On Error GoTo 0
    End If
    
    ' --- FIN DE LA LÓGICA REFACTORIZADA ---
    
LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.cmdVerDatos_Click")
    Resume LimpiarYSalir
End Sub

Private Sub cmdGestionarAdjuntos_Click()
    On Error GoTo Errores
    If m_ViewModel Is Nothing Then Exit Sub
    Me.lblBloqueo.Visible = True
    DoCmd.OpenForm "frmGestionAdjuntos", acNormal, , , , acDialog, CStr(m_ViewModel.Solicitud.idSolicitud)
LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.cmdGestionarAdjuntos_Click")
    Resume LimpiarYSalir
End Sub

Private Sub cmdSalir_Click()
    DoCmd.Close acForm, Me.name
End Sub
' EN: Form_frmGestionSolicitud.cls

' ==========================================================================
' FUNCIÓN: LeerColaComandosConTimeout
' DESCRIPCIÓN: Lee window.colaComandos con timeout de 2 segundos.
'              Evita que RetrieveJavascriptValue quede bloqueado indefinidamente.
' ==========================================================================
Private Function LeerColaComandosConTimeout() As String
    Dim sResultado As String
    Dim dTimeout As Date
    Dim bExito As Boolean
    
    sResultado = ""
    bExito = False
    dTimeout = Now + TimeSerial(0, 0, 2)  ' Timeout de 2 segundos
    
    Do While Now < dTimeout
        On Error Resume Next
        sResultado = Me.webInfo.RetrieveJavascriptValue("window.colaComandos")
        If Err.Number = 0 Then
            bExito = True
            Exit Do
        End If
        If InStr(1, Err.description, "timeout", vbTextCompare) > 0 Then
            Exit Do
        End If
        On Error GoTo 0
        DoEvents
    Loop
    
    On Error GoTo 0
    If Not bExito Then sResultado = ""
    LeerColaComandosConTimeout = sResultado
End Function

Private Sub Form_Timer()
    ' ==========================================================================
    ' SOLUCIÓN HÍBRIDA: SOPORTE EDGE (POLLING) Y LEGACY (ONE-SHOT)
    ' ==========================================================================
    On Error Resume Next
    
    ' --- MODO EDGE: SONDEO CONTINUO ---
    If Application.TempVars("UsarRenderizadoFichero") = "Sí" Then
        Dim cmd As String
        Dim errNum As Long
        Dim errDesc As String
        
        ' 1. LEER: Preguntamos al navegador con timeout
        cmd = LeerColaComandosConTimeout()
        
        errNum = Err.Number
        errDesc = Err.description
        
        If cmd = "" Then
            ' Solo logueamos si hay error
            ' LogDebug removido — Debug.Print "Error LeerColaComandosConTimeout: " & errNum & " - " & errDesc
            Err.Clear
        End If
        
        cmd = Nz(cmd, "")
        
        ' 2. VALIDAR: Si hay comando válido (y no es un mensaje de error del control)
        If cmd <> "" And cmd <> "null" And cmd <> "undefined" And InStr(1, cmd, "tiempo de espera", vbTextCompare) = 0 Then
            
            ' Comando DETECTADO — sin debug
            ' Debug.Print "Comando DETECTADO: " & cmd
            
            ' 3. LIMPIAR: Borramos inmediatamente en JS para no repetir
            ' Solo intentamos limpiar si no hubo timeout previo para no saturar
            Me.webInfo.ExecuteJavascript "window.colaComandos = null;"
            
            ' 4. PARSEAR
            Dim partes() As String
            partes = Split(cmd, ":")
            
            If UBound(partes) >= 0 Then
                m_AccionPendiente = partes(0)
                If UBound(partes) >= 1 Then
                    If IsNumeric(partes(1)) Then
                        m_ParametroPendiente = CLng(partes(1))
                        m_ParametroTextoPendiente = ""
                    Else
                        m_ParametroPendiente = 0
                        m_ParametroTextoPendiente = partes(1)
                    End If
                Else
                    m_ParametroPendiente = 0
                    m_ParametroTextoPendiente = ""
                End If
                
                ' 5. EJECUTAR (Desacoplado del hilo del navegador)
                Call EjecutarAccionPendiente
            End If
        ElseIf InStr(1, cmd, "tiempo de espera", vbTextCompare) > 0 Then
            ' Si es un timeout, no hacemos nada, simplemente esperamos al siguiente ciclo
            ' Esto evita que tratemos el mensaje de error como un comando
            Err.Clear
        End If
        
    ' --- MODO LEGACY (IE): DISPARO ÚNICO ---
    Else
        ' Desactivamos timer tras la primera ejecución (disparado por BeforeNavigate)
        Me.TimerInterval = 0
        If m_AccionPendiente <> "" Then
            Call EjecutarAccionPendiente
        End If
    End If
    
    On Error GoTo 0
End Sub

Private Sub EjecutarAccionPendiente()
    Dim fsoGen As Object  ' Declarado aquí para sea accesible desde los labels de cleanup
    
    On Error GoTo Errores
    
    ' Ejecutamos la lógica según la acción que guardamos
    Select Case m_AccionPendiente
        
        Case "ERROR_SIN_CAMBIOS"
            MsgBox "No se han detectado cambios en la propuesta técnica desde el último rechazo." & vbCrLf & vbCrLf & _
                   "Debe realizar modificaciones en las pestañas 'Detalle y Descripción' o 'Motivos e Impacto' para poder reenviar la solicitud.", _
                   vbExclamation, "Sin Cambios Detectados"
                   
        Case "AVANZAR_ESTADO"
            If m_ParametroPendiente = estadoRechazada Then
                If MsgBox("¿Confirma el RECHAZO definitivo de la solicitud?" & vbCrLf & vbCrLf & _
                          "Se asignará automáticamente el motivo: 'Rechazado desde el visor de cambio de estado'.", _
                          vbQuestion + vbYesNo + vbDefaultButton2, "Rechazo Rápido") = vbYes Then
                    Call InyectarMotivoRechazoAutomatico(m_ViewModel.Solicitud)
                    Call AccionAvanzarEstado(m_ParametroPendiente)
                End If
            ElseIf m_ViewModel.Solicitud.idEstadoInterno = estadoAprobada And m_ParametroPendiente = estadoFormalizacion Then
                Dim msgResetAprob As String
                msgResetAprob = "ATENCIÓN: Esta acción devolverá la solicitud al estado 'FORMALIZACIÓN'." & vbCrLf & vbCrLf & _
                                "Se eliminaran los datos de la 'Decision Final' y el documento de Cierre." & vbCrLf & vbCrLf & _
                                "¿Desea continuar con la pérdida irreversible de los datos de esta fase y del adjunto firmado?"
                If MsgBox(msgResetAprob, vbExclamation + vbYesNo + vbDefaultButton2, "Reset de Aprobación") = vbYes Then
                    Call AccionRevertirGenerica(estadoFormalizacion, "Reapertura desde Aprobada")
                End If
            Else
                If MsgBox("¿Desea cambiar el estado de la solicitud?", vbQuestion + vbYesNo, "Confirmar Transición") = vbYes Then
                    Call AccionAvanzarEstado(m_ParametroPendiente)
                End If
            End If
            
        Case "APROBAR_TECNICA"
            Call AccionAprobarYAvanzar(m_ParametroPendiente)
            
        Case "REVERTIR_TECNICA"
            Dim msgReset As String
            msgReset = "ATENCIÓN: Esta acción borrará todos los datos de la fase actual y devolverá la solicitud al estado anterior. ¿Desea continuar?"
            If MsgBox(msgReset, vbExclamation + vbYesNo + vbDefaultButton2, "Reset de Fase") = vbYes Then
                Call AccionRevertirTecnica(m_ParametroPendiente)
            End If
            
        Case "REVERTIR_DESARROLLO"
            msgReset = "ATENCIÓN: Esta acción borrará todos los datos de la fase actual y devolverá la solicitud al estado anterior. ¿Desea continuar?"
            If MsgBox(msgReset, vbExclamation + vbYesNo + vbDefaultButton2, "Reset de Fase") = vbYes Then
                Call AccionRevertirGenerica(m_ParametroPendiente, "Retroceso a Desarrollo Tecnico con purga")
            End If
            
        Case "REVERTIR_VALIDACION"
            msgReset = "ATENCIÓN: Esta acción borrará todos los datos de la fase actual y devolverá la solicitud al estado anterior. ¿Desea continuar?"
            If MsgBox(msgReset, vbExclamation + vbYesNo + vbDefaultButton2, "Reset de Fase") = vbYes Then
                Call AccionRevertirGenerica(m_ParametroPendiente, "Hard reset validacion/revision")
            End If
            
        Case "REVERTIR_FORMALIZACION"
            msgReset = "ATENCION: Esta accion realizara un RESET de la fase de Formalizacion. Se ELIMINARA el documento de Cierre y se borrara la Decision Final. Desea continuar con la perdida irreversible de los documentos?"
            If MsgBox(msgReset, vbExclamation + vbYesNo + vbDefaultButton2, "Reset de Formalización") = vbYes Then
                Call AccionRevertirGenerica(m_ParametroPendiente, "Corrección desde Formalización")
            End If

        Case "REVERTIR_RECHAZO"
            If MsgBox("¿Está seguro de que desea ANULAR el rechazo y reabrir la solicitud en fase de 'REVISIÓN'?" & vbCrLf & vbCrLf & _
                      "Esta acción ELIMINARÁ la decisión de rechazo y sus motivos para permitir una nueva evaluación." & vbCrLf & vbCrLf & _
                      "¿Continuar?", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Anular Rechazo") = vbYes Then
                Call AccionRevertirGenerica(m_ParametroPendiente, "Anulación de Rechazo desde Web")
            End If
            
        Case "VER_MOTIVO_CALIDAD", "VER_MOTIVOS_RECHAZO"
            Dim wfServMotivo As New WorkflowServicio
            Dim revServMotivo As New RevisionServicio
            Dim rechazoActivo As rechazo
            Dim htmlMotivo As String
            Dim modoRechazo As String
            Dim comentariosRechazo As String
            Dim hayCambios As Boolean
            Dim rolUsuario As Long
            
            rolUsuario = m_ObjUsuarioActivo.rol
            hayCambios = wfServMotivo.VerificarCambiosTrasRechazo(m_ViewModel.Solicitud)
            Set rechazoActivo = revServMotivo.ObtenerRechazoActivo(m_ViewModel.Solicitud.idSolicitud)
            
            If Not rechazoActivo Is Nothing Then
                comentariosRechazo = rechazoActivo.comentarios
            Else
                comentariosRechazo = Nz(m_ViewModel.Solicitud.revisionCalidadComentarios, "")
            End If
            
            If rechazoActivo Is Nothing Then
                modoRechazo = "LECTURA"
            ElseIf rolUsuario = rol.Tecnico Then
                modoRechazo = "LECTURA"
            ElseIf (rolUsuario = rol.Calidad Or rolUsuario = rol.Administrador) And Not hayCambios Then
                modoRechazo = "EDICION"
            Else
                modoRechazo = "LECTURA"
            End If
            
            m_ModoRechazoCalidad = modoRechazo
            htmlMotivo = wfServMotivo.GenerarHTML_CapturaRechazoCalidad(comentariosRechazo, modoRechazo)
            DoCmd.OpenForm "frmWebVisor", acNormal
            Form_frmWebVisor.CargarContenido htmlMotivo, "Motivo de rechazo"
            
        Case "VER_CAMBIOS", "VER_DELTA"
            Dim wfServ As New WorkflowServicio
            Dim htmlDeltas As String
            htmlDeltas = wfServ.GenerarHTML_DeltasRechazo(m_ViewModel.Solicitud.idSolicitud)
            DoCmd.OpenForm "frmWebVisor", acNormal
            Form_frmWebVisor.CargarContenido htmlDeltas, "Cambios detectados (Subsanación)"
            
        Case "VER_HISTORIAL_RAC"
            Dim wfServHist As New WorkflowServicio
            Dim htmlHist As String
            htmlHist = wfServHist.GenerarHTML_HistorialRAC(m_ViewModel.Solicitud.idSolicitud)
            DoCmd.OpenForm "frmWebVisor", acNormal
            Form_frmWebVisor.CargarContenido htmlHist, "Historial de Validación y Revisión"
            
        Case "ABRIR_ADJUNTO"
            Dim adj As Adjunto
            Dim rutaAdjunto As String
            Dim dirAdjunto As String
            If m_ParametroPendiente > 0 Then
                Dim adjServ As New AdjuntosServicio
                Set adj = adjServ.getAdjuntoPorID(m_ParametroPendiente)
                If Not adj Is Nothing Then
                    dirAdjunto = m_ObjEntorno.URLDirectorioDocumentacion
                    If Right$(dirAdjunto, 1) <> "\" Then
                        dirAdjunto = dirAdjunto & "\"
                    End If
                    rutaAdjunto = dirAdjunto & adj.nombreArchivo
                    If rutaAdjunto <> "" Then
                        Call AbrirEnLocal(rutaAdjunto)
                    End If
                End If
            End If
            
        Case "EXPORTAR_BORRADOR"
            ' Spec-144: Manejo de warning por retorno vacío (sin cambios detectados)
            Dim wfServVersion As New WorkflowServicio
            Dim rutaFinal As String
            On Error Resume Next
            rutaFinal = wfServVersion.ExportarBorradorSingleton(m_ViewModel.Solicitud.idSolicitud)
            
            If Err.Number <> 0 Then
                ' Error real -> propagar al manejador principal
                Dim tmpErrNum As Long: tmpErrNum = Err.Number
                Dim tmpErrDesc As String: tmpErrDesc = Err.description
                On Error GoTo Errores
                Err.Raise tmpErrNum, "Form_frmGestionSolicitud.EXPORTAR_BORRADOR", tmpErrDesc
            End If
            
            On Error GoTo Errores
            
            ' Spec-144: Detectar retorno vacío como señal de warning
            If rutaFinal = "" Then
                MsgBox "No se puede obtener una nueva versión al no haber detectado cambios con respecto a la última.", vbExclamation, "Sin cambios detectados"
            ElseIf Len(Nz(rutaFinal, "")) > 0 Then
                Call AbrirEnLocal(rutaFinal)
                Dim jsUpdateChecklist As String
                jsUpdateChecklist = "(function(){var items = document.querySelectorAll('.req-nok, .req-ok'); for(var i=0; i<items.length; i++){ var text = items[i].querySelector('.req-text'); if(text && text.textContent.indexOf('Exportar borrador') !== -1){ items[i].className = 'req-ok'; items[i].querySelector('.check-icon').innerHTML = '&#10003;'; break; } } })();"
                Me.webInfo.ExecuteJavascript jsUpdateChecklist
            Else
                MsgBox "No se pudo generar el documento.", vbExclamation, "CONDOR"
            End If
            
        Case "CAPTURA_ENVIO_RAC", "NOTIFICAR_ENVIO_RAC"
            Dim wfServEnvio As New WorkflowServicio
            Dim htmlEnvio As String
            If Not wfServEnvio.TieneAdjuntoEnCicloActual(m_ViewModel.Solicitud.idSolicitud) Then
                MsgBox "Debe exportar el borrador antes de notificar el envio.", vbExclamation, "CONDOR"
                Exit Sub
            End If
            htmlEnvio = wfServEnvio.GenerarHTML_CapturaFechaRAC(Date)
            DoCmd.OpenForm "frmWebVisor", acNormal
            Form_frmWebVisor.CargarContenido htmlEnvio, "Notificar envio RAC", "CAPTURA_FECHA_RAC"
            
        Case "CAPTURA_COMENTARIOS_RAC", "COMENTARIOS_RAC"
            Dim wfServCom As New WorkflowServicio
            Dim htmlCom As String
            htmlCom = wfServCom.GenerarHTML_CapturaComentariosRAC("")
            DoCmd.OpenForm "frmWebVisor", acNormal
            Form_frmWebVisor.CargarContenido htmlCom, "Comentarios RAC", "CAPTURA_COMENTARIOS_RAC"
            
        Case "IMPORTAR_WORD_RAC_DESDE_BOTON"
            Call ProcesarImportacionWordRAC_BotonDirecto
            
        Case "APROBAR_RAC"
            Dim idValAprobado As Long
            Dim srvHito As ValidacionRevisionServicio
            Set srvHito = New ValidacionRevisionServicio
            idValAprobado = srvHito.GetUltimoPendienteId(m_ViewModel.Solicitud.idSolicitud)
            If idValAprobado <= 0 Then
                MsgBox "No hay una validación RAC pendiente para aprobar.", vbExclamation, "CONDOR"
                GoTo AccionPendienteFinalizar
            End If

            Dim srvResRAC As ValidacionRevisionServicio
            Set srvResRAC = New ValidacionRevisionServicio
            Call srvResRAC.RegistrarRespuestaRAC(idValAprobado, "APROBADO", "", Now, esIdHitoDirecto:=True)
            Call AccionAvanzarEstado(estadoFormalizacion)
            Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, True)
            DoEvents
            Me.Repaint
            
        Case "GENERAR_PDF_FINAL"
            Dim wfServFinal As New WorkflowServicio
            Dim rutaPDFfinal As String
            Dim okAprob As Boolean
            Dim tipo As String
            Dim pcServCheck As New DatosPCServicio
            Dim cdcaServCheck As New DatosCDCAServicio
            Dim subServCheck As New DatosCDCASUBServicio
            Dim pcsubServCheck As New DatosPCSUBServicio
            On Error GoTo Errores
            
            tipo = m_ViewModel.Solicitud.tipoSolicitud
            okAprob = False
            If tipo = "PC" Then
                okAprob = pcServCheck.EsAprobacionSuministradorCompleta(m_ViewModel.Solicitud.idSolicitud)
            End If
            If tipo = "CD_CA" Then
                okAprob = cdcaServCheck.EsAprobacionSuministradorCompleta(m_ViewModel.Solicitud.idSolicitud)
            End If
            If tipo = "CD_CA_SUB" Then
                okAprob = subServCheck.EsAprobacionSuministradorCompleta(m_ViewModel.Solicitud.idSolicitud)
            End If
            If tipo = "PC_SUB" Then
                okAprob = pcsubServCheck.EsAprobacionSuministradorCompleta(m_ViewModel.Solicitud.idSolicitud)
            End If
            If Not okAprob Then
                MsgBox "Complete la pestaña 'Aprobación Suministrador' antes de generar el PDF.", vbExclamation, "CONDOR"
                Set fsoGen = Nothing: GoTo AccionPendienteFinalizar
            End If
            If MsgBox("Se va a generar el documento PDF final con los datos actuales para su envio a firma." & vbCrLf & vbCrLf & "Desea continuar?", vbQuestion + vbYesNo, "Generar PDF Final") = vbNo Then
                Set fsoGen = Nothing: GoTo AccionPendienteFinalizar
            End If
            Me.lblBloqueo.Visible = True
            DoCmd.Hourglass True
            rutaPDFfinal = wfServFinal.GenerarYAdjuntarPDFfinal(m_ViewModel.Solicitud.idSolicitud)
            Me.lblBloqueo.Visible = False
            DoCmd.Hourglass False
            
            If Len(Trim$(Nz(rutaPDFfinal, ""))) = 0 Then
                MsgBox "No se pudo generar el PDF final. Ruta invalida.", vbExclamation, "CONDOR"
                Set fsoGen = Nothing: GoTo AccionPendienteFinalizar
            End If
            
            Set fsoGen = CreateObject("Scripting.FileSystemObject")
            If fsoGen.FileExists(rutaPDFfinal) Then
                Call AbrirEnLocal(rutaPDFfinal)
            Else
                MsgBox "No se pudo generar el PDF final. Archivo no encontrado.", vbExclamation
            End If
            Set fsoGen = Nothing: GoTo AccionPendienteFinalizar
            
        Case "ADJUNTAR_PDF_FIRMADO"
            Dim wfServGate As New WorkflowServicio
            Dim okFirmas As Boolean
            Dim adjServPDF As New AdjuntosServicio
            
            ' GAP-139: Usar método genérico GeneradoPDFParaFirma
            If Not adjServPDF.GeneradoPDFParaFirma(m_ViewModel.Solicitud.idSolicitud) Then
                MsgBox "Debe generar el PDF desde el boton 'Generar PDF firma' antes de adjuntar el documento firmado.", vbExclamation, "CONDOR"
                GoTo AccionPendienteFinalizar
            End If
            
            okFirmas = wfServGate.EsAprobacionSuministradorCompleta(m_ViewModel.Solicitud.idSolicitud)
            If Not okFirmas Then
                MsgBox "Debe cumplimentar y guardar la pestaña Aprobación Suministrador antes de adjuntar el documento firmado.", vbExclamation, "CONDOR"
                GoTo AccionPendienteFinalizar
            End If
            Application.TempVars.Add "AutoUploadEtapa", "Cierre"
            DoCmd.OpenForm "frmGestionAdjuntos", , , , , acDialog, CStr(m_ViewModel.Solicitud.idSolicitud)
            Dim adjServAuto As New AdjuntosServicio
            If adjServAuto.ExisteAdjuntoEtapa(m_ViewModel.Solicitud.idSolicitud, "Cierre") _
               Or adjServAuto.ExisteAdjuntoEtapa(m_ViewModel.Solicitud.idSolicitud, "Documento Final Firmado") Then
                Call AccionAprobarFinal
            End If
            
        Case "APROBAR_FINAL"
            Call AccionAprobarFinal
            
        Case "DATOS_RAC"
            Call ProcesarDatosRAC(m_ParametroTextoPendiente)
            
        Case "IMPORTAR_WORD_RAC"
            Call ProcesarImportacionWordRAC(m_ParametroTextoPendiente)
            
        Case "IMPORTAR_WORD_RAC_DESDE_BOTON"
            Call ProcesarImportacionWordRAC_BotonDirecto
            
        Case "DATOS_RECHAZO"
            Call ProcesarDatosRechazoCalidad(m_ParametroTextoPendiente)
            
        Case "RECHAZAR_TECNICA"
            Dim wfServRechazo As New WorkflowServicio
            Dim htmlRechazo As String
            m_ModoRechazoCalidad = "CAPTURA"
            htmlRechazo = wfServRechazo.GenerarHTML_CapturaRechazoCalidad("")
            DoCmd.OpenForm "frmWebVisor", acNormal
            Form_frmWebVisor.CargarContenido htmlRechazo, "Rechazo calidad", "CAPTURA_RECHAZO_CALIDAD"
            
        Case "GET_HAY_CAMBIOS"
            Dim wfServCambios As New WorkflowServicio
            Dim hayCambiosResult As String
            hayCambiosResult = wfServCambios.ObtenerHayCambiosValidacion(m_ParametroPendiente)
            Me.webInfo.ExecuteJavascript "window.ajaxResult = '" & hayCambiosResult & "';"

    End Select
    
AccionPendienteFinalizar:
    m_AccionPendiente = ""
    m_ParametroPendiente = 0
    m_ParametroTextoPendiente = ""
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.EjecutarAccionPendiente")
    Resume LimpiarYSalir
End Sub

Private Sub AccionAprobarFinal()
    Dim adjServ As New AdjuntosServicio
    Dim wfServ As New WorkflowServicio
    Dim tipo As String
    Dim okAdj As Boolean
    Dim ws As DAO.Workspace, db As DAO.Database
    Dim sol As Solicitud
    
    On Error GoTo Errores
    
    tipo = m_ViewModel.Solicitud.tipoSolicitud
    okAdj = adjServ.ExisteAdjuntoEtapa(m_ViewModel.Solicitud.idSolicitud, "Cierre")
    If Not okAdj Then
        okAdj = adjServ.ExisteAdjuntoEtapa(m_ViewModel.Solicitud.idSolicitud, "Documento Final Firmado")
    End If
    If Not okAdj Then
        MsgBox "Debe adjuntar el documento de Cierre antes de aprobar.", vbExclamation, "CONDOR - Cierre"
        Exit Sub
    End If
    
    Me.lblBloqueo.Visible = True: DoCmd.Hourglass True
    Set ws = getWorkspace()
    Set db = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
    ws.BeginTrans
    
    Set sol = m_ViewModel.Solicitud
    Select Case tipo
        Case "PC"
            Dim pc As New DatosPC
            Dim pcServ As New DatosPCServicio
            pc.idSolicitud = sol.idSolicitud
            pc.decisionFinal = "APROBADO"
            Call pcServ.GuardarDecisionFinal(pc, db)
        Case "CD_CA"
            Dim cd As New DatosCDCA
            Dim cdServ As New DatosCDCAServicio
            cd.idSolicitud = sol.idSolicitud
            cd.decisionFinal = "APROBADO"
            Call cdServ.GuardarDecisionFinal(cd, db)
        Case "CD_CA_SUB"
            Dim subd As New DatosCDCASUB
            Dim subServ As New DatosCDCASUBServicio
            subd.idSolicitud = sol.idSolicitud
            subd.decisionFinal = "APROBADO"
            Call subServ.GuardarDecisionFinal(subd, db)
        Case "PC_SUB"
            Dim pcsub As New DatosPCSUB
            Dim pcsubServ As New DatosPCSUBServicio
            pcsub.idSolicitud = sol.idSolicitud
            pcsub.decisionFinal = "APROBADO"
            Call pcsubServ.GuardarDecisionFinal(pcsub, db)
    End Select
    
    Call wfServ.EjecutarTransicion(sol, estadoAprobada, m_ObjUsuarioActivo, db)
    ws.CommitTrans
    
    Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
    Call CargarDatosSolicitud(sol.idSolicitud, mantenerVistaActual:=True)
    DoEvents
    Me.Repaint
    Exit Sub
    
Errores:
    ' Rollback con trace si falla la limpieza
    Dim rollbackErrNum As Long, rollbackErrDesc As String
    On Error Resume Next
    If Not ws Is Nothing Then ws.Rollback
    rollbackErrNum = Err.Number: rollbackErrDesc = Err.description: On Error GoTo 0
    If rollbackErrNum <> 0 Then
        Dim logSvcRollback As LogServicio
        Set logSvcRollback = New LogServicio
        logSvcRollback.RegistrarError rollbackErrNum, rollbackErrDesc, "frmGestionSolicitud", "AccionAprobarFinal_Rollback"
    End If
    Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
End Sub

Public Sub RecibirDatosRAC(ByVal data As String)
    m_AccionPendiente = "DATOS_RAC"
    m_ParametroTextoPendiente = data
    Call EjecutarAccionPendiente
End Sub

Public Sub RecibirDatosRechazoCalidad(ByVal data As String)
    m_AccionPendiente = "DATOS_RECHAZO"
    m_ParametroTextoPendiente = data
    Call EjecutarAccionPendiente
End Sub

Public Sub RecibirImportacionWordRAC()
    m_AccionPendiente = "IMPORTAR_WORD_RAC"
    m_ParametroTextoPendiente = ""
    Call EjecutarAccionPendiente
End Sub

Public Sub RecibirImportacionWordRAC_Boton()
    m_AccionPendiente = "IMPORTAR_WORD_RAC_DESDE_BOTON"
    m_ParametroTextoPendiente = ""
    Call EjecutarAccionPendiente
End Sub

Public Sub Debug_SimularEntradaRAC(ByVal data As String)
    Call ProcesarDatosRAC(data)
End Sub

Private Sub ProcesarDatosRAC(ByVal data As String)
    Dim valor As String
    Dim tipo As String
    Dim pos As Long
    Dim idVal As Long
    
    If Trim(Nz(data, "")) = "" Then Exit Sub

    If CurrentProject.AllForms("frmWebVisor").IsLoaded Then DoCmd.Close acForm, "frmWebVisor"
    
    data = UrlDecode(data)
    
    If UCase$(data) = "CANCELAR" Then
        If CurrentProject.AllForms("frmWebVisor").IsLoaded Then DoCmd.Close acForm, "frmWebVisor"
        Exit Sub
    End If
    
    pos = InStr(1, data, "|", vbBinaryCompare)
    If pos > 0 Then
        tipo = Left$(data, pos - 1)
        valor = Mid$(data, pos + 1)
    Else
        tipo = data
        valor = ""
    End If
    
    tipo = UCase$(tipo)
    
    If tipo = "FECHA" Then
        If Not IsDate(valor) Then
            DoCmd.SelectObject acForm, Me.name, False
            DoEvents
            MsgBox "Fecha invalida.", vbExclamation, "CONDOR"
            Exit Sub
        End If
        Dim wfServEnvio As New WorkflowServicio
        If Not wfServEnvio.TieneAdjuntoEnCicloActual(m_ViewModel.Solicitud.idSolicitud) Then
            DoCmd.SelectObject acForm, Me.name, False
            DoEvents
            MsgBox "Debe exportar el borrador antes de notificar el envio.", vbExclamation, "CONDOR"
            Exit Sub
        End If
        Dim srvVR As ValidacionRevisionServicio
        Set srvVR = New ValidacionRevisionServicio
        Call srvVR.RegistrarEnvioRAC(m_ViewModel.Solicitud.idSolicitud, CDate(valor), m_ObjUsuarioActivo.nombre)
        Call AccionAvanzarEstado(estadoRevision)
    ElseIf tipo = "COMENTARIOS" Then
        Dim srvValId As ValidacionRevisionServicio
        Set srvValId = New ValidacionRevisionServicio
        idVal = srvValId.GetUltimoPendienteId(m_ViewModel.Solicitud.idSolicitud)
        If idVal = 0 Then
            DoCmd.SelectObject acForm, Me.name, False
            DoEvents
            MsgBox "No existe ciclo pendiente.", vbExclamation, "CONDOR"
            Exit Sub
        End If
        Dim srvResRAC As ValidacionRevisionServicio
            Set srvResRAC = New ValidacionRevisionServicio
            Call srvResRAC.RegistrarRespuestaRAC(idVal, "COMENTARIOS", valor, Now, esIdHitoDirecto:=True)
        Call AccionRetrocederFlujo(estadoValidacion, "Revision con Comentarios")
        MsgBox "Se ha guardado correctamente, recuerda completar el formulario segun las observaciones del RAC", vbInformation, "Confirmacion"
    End If
    
    If CurrentProject.AllForms("frmWebVisor").IsLoaded Then DoCmd.Close acForm, "frmWebVisor"
    Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, mantenerVistaActual:=True)
    DoEvents
    Me.Repaint
End Sub

Private Function VerificarDictamenRACCompleto(ByVal idSolicitud As Long, ByVal tipoSolicitud As String) As Boolean
    Dim pcServ As DatosPCServicio
    Dim cdcaServ As DatosCDCAServicio
    Dim cdcaSubServ As DatosCDCASUBServicio
    Dim pcsubServ As DatosPCSUBServicio
    
    On Error GoTo Errores
    VerificarDictamenRACCompleto = False
    
    Select Case tipoSolicitud
        Case "PC"
            Set pcServ = New DatosPCServicio
            VerificarDictamenRACCompleto = pcServ.EsDictamenRACCompleto(idSolicitud)
        Case "CD_CA"
            Set cdcaServ = New DatosCDCAServicio
            VerificarDictamenRACCompleto = cdcaServ.EsDictamenRACCompleto(idSolicitud)
        Case "CD_CA_SUB"
            Set cdcaSubServ = New DatosCDCASUBServicio
            VerificarDictamenRACCompleto = cdcaSubServ.EsDictamenRACCompleto(idSolicitud)
        Case "PC_SUB"
            Set pcsubServ = New DatosPCSUBServicio
            VerificarDictamenRACCompleto = pcsubServ.EsDictamenRACCompleto(idSolicitud)
        Case Else
            VerificarDictamenRACCompleto = False
    End Select
    
    Exit Function
Errores:
    VerificarDictamenRACCompleto = False
End Function

Private Sub ProcesarImportacionWordRAC(ByVal fileName As String)
    Dim docServ As New DocumentoServicio
    Dim tipoSolicitud As String
    Dim idSolicitud As Long
    Dim fDialog As FileDialog
    Dim filePath As String
    Dim observacionesImportadas As String
    Dim db As DAO.Database
    Dim sqlUpdate As String
    Dim camposSet As String
    Dim pcDatos As DatosPC
    Dim cdDatos As DatosCDCA
    Dim subDatos As DatosCDCASUB
    Dim pcsubDatos As Object
    
    On Error GoTo Errores
    
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    fDialog.Title = "Seleccione el documento Word con las observaciones del RAC"
    fDialog.Filters.Clear
    fDialog.Filters.Add "Documentos de Word", "*.docx; *.doc"
    If fDialog.Show <> -1 Then
        If CurrentProject.AllForms("frmWebVisor").IsLoaded Then DoCmd.Close acForm, "frmWebVisor"
        Exit Sub
    End If
    filePath = fDialog.SelectedItems(1)
    
    DoEvents
    Sleep 500
    
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    
    idSolicitud = m_ViewModel.Solicitud.idSolicitud
    tipoSolicitud = m_ViewModel.Solicitud.tipoSolicitud
    
    Dim hashAntesImportWeb As String
    hashAntesImportWeb = SnapshotHelper.CalcularHashFaseValidacion(idSolicitud, tipoSolicitud, db)
    
    If tipoSolicitud = "PC" Then
        Set pcDatos = docServ.ImportarDatosRACDesdeWord(filePath)
        If pcDatos Is Nothing Then
            MsgBox "No se pudieron leer datos del archivo.", vbExclamation, "Importacion Word"
            GoTo LimpiarYSalir
        End If
        
        Set db = getdb()
        camposSet = ""
        If Not IsNull(pcDatos.observacionesRAC) And Trim(pcDatos.observacionesRAC) <> "" Then
            camposSet = "observacionesRAC = '" & Replace(pcDatos.observacionesRAC, "'", "''") & "'"
        End If
        If Not IsNull(pcDatos.racCodigo) And Trim(pcDatos.racCodigo) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racCodigo = '" & Replace(pcDatos.racCodigo, "'", "''") & "'"
        End If
        If Not IsNull(pcDatos.racNombre) And Trim(pcDatos.racNombre) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racNombre = '" & Replace(pcDatos.racNombre, "'", "''") & "'"
        End If
        
        If camposSet <> "" Then
            sqlUpdate = "UPDATE tbDatosPC SET " & camposSet & " WHERE idSolicitud = " & idSolicitud
            db.Execute sqlUpdate, dbFailOnError
        End If
        
    ElseIf tipoSolicitud = "CD_CA" Then
        Set cdDatos = docServ.ImportarDatosRACDesdeWord_CDCA(filePath)
        If cdDatos Is Nothing Then
            MsgBox "No se pudieran leer datos del archivo.", vbExclamation, "Importacion Word"
            GoTo LimpiarYSalir
        End If
        
        Set db = getdb()
        camposSet = ""
        If Not IsNull(cdDatos.observacionesRAC) And Trim(cdDatos.observacionesRAC) <> "" Then
            camposSet = "observacionesRAC = '" & Replace(cdDatos.observacionesRAC, "'", "''") & "'"
        End If
        If Not IsNull(cdDatos.racCodigo) And Trim(cdDatos.racCodigo) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racCodigo = '" & Replace(cdDatos.racCodigo, "'", "''") & "'"
        End If
        If Not IsNull(cdDatos.racNombre) And Trim(cdDatos.racNombre) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racNombre = '" & Replace(cdDatos.racNombre, "'", "''") & "'"
        End If
        
        If camposSet <> "" Then
            sqlUpdate = "UPDATE tbDatosCDCA SET " & camposSet & " WHERE idSolicitud = " & idSolicitud
            db.Execute sqlUpdate, dbFailOnError
        End If
        
    ElseIf tipoSolicitud = "CD_CA_SUB" Then
        Set subDatos = docServ.ImportarDatosRACDesdeWord_CDCASUB(filePath)
        If subDatos Is Nothing Then
            MsgBox "No se pudieran leer datos del archivo.", vbExclamation, "Importacion Word"
            GoTo LimpiarYSalir
        End If
        
        Set db = getdb()
        camposSet = ""
        If Not IsNull(subDatos.observacionesRAC) And Trim(subDatos.observacionesRAC) <> "" Then
            camposSet = "observacionesRAC = '" & Replace(subDatos.observacionesRAC, "'", "''") & "'"
        End If
        If Not IsNull(subDatos.racCodigo) And Trim(subDatos.racCodigo) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racCodigo = '" & Replace(subDatos.racCodigo, "'", "''") & "'"
        End If
        If Not IsNull(subDatos.racNombre) And Trim(subDatos.racNombre) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racNombre = '" & Replace(subDatos.racNombre, "'", "''") & "'"
        End If
        
        If camposSet <> "" Then
            sqlUpdate = "UPDATE tbDatosCDCASUB SET " & camposSet & " WHERE idSolicitud = " & idSolicitud
            db.Execute sqlUpdate, dbFailOnError
            
            ' GAP FIX: si no existe fila en tbDatosCDCASUB, el UPDATE no afecta registros.
            ' Hacemos UPSERT para garantizar que se guarden comentarios/código/nombre importados.
            If db.RecordsAffected = 0 Then
                Dim subSrvUpsert As New DatosCDCASUBServicio
                Dim subEntUpsert As DatosCDCASUB
                
                Set subEntUpsert = subSrvUpsert.getDatosCDCASUB(idSolicitud)
                If subEntUpsert Is Nothing Then
                    Set subEntUpsert = subSrvUpsert.GetDatosGeneralesIniciales(idSolicitud)
                    subEntUpsert.idSolicitud = idSolicitud
                End If
                
                If Not IsNull(subDatos.observacionesRAC) And Trim(subDatos.observacionesRAC) <> "" Then
                    subEntUpsert.observacionesRAC = subDatos.observacionesRAC
                End If
                If Not IsNull(subDatos.racCodigo) And Trim(subDatos.racCodigo) <> "" Then
                    subEntUpsert.racCodigo = subDatos.racCodigo
                End If
                If Not IsNull(subDatos.racNombre) And Trim(subDatos.racNombre) <> "" Then
                    subEntUpsert.racNombre = subDatos.racNombre
                End If
                
                Call DatosCDCASUBRepositorio.Guardar(subEntUpsert, db)
            End If
        End If
    ElseIf tipoSolicitud = "PC_SUB" Then
        ' PC_SUB usa la importación de PC (misma estructura RAC)
        Set pcsubDatos = docServ.ImportarDatosRACDesdeWord(filePath)
        If pcsubDatos Is Nothing Then
            MsgBox "No se pudieron leer datos del archivo.", vbExclamation, "Importacion Word"
            GoTo LimpiarYSalir
        End If

        Set db = getdb()
        camposSet = ""
        If Not IsNull(pcsubDatos.observacionesRAC) And Trim(pcsubDatos.observacionesRAC) <> "" Then
            camposSet = "observacionesRAC = '" & Replace(pcsubDatos.observacionesRAC, "'", "''") & "'"
        End If
        If Not IsNull(pcsubDatos.racCodigo) And Trim(pcsubDatos.racCodigo) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racCodigo = '" & Replace(pcsubDatos.racCodigo, "'", "''") & "'"
        End If
        If Not IsNull(pcsubDatos.racNombre) And Trim(pcsubDatos.racNombre) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racNombre = '" & Replace(pcsubDatos.racNombre, "'", "''") & "'"
        End If
        ' PC_SUB también tiene campos RACDelegado
        If Not IsNull(pcsubDatos.racDelegadoCodigo) And Trim(pcsubDatos.racDelegadoCodigo) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racDelegadoCodigo = '" & Replace(pcsubDatos.racDelegadoCodigo, "'", "''") & "'"
        End If
        If Not IsNull(pcsubDatos.racDelegadoObservaciones) And Trim(pcsubDatos.racDelegadoObservaciones) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racDelegadoObservaciones = '" & Replace(pcsubDatos.racDelegadoObservaciones, "'", "''") & "'"
        End If
        If Not IsNull(pcsubDatos.racDelegadoNombre) And Trim(pcsubDatos.racDelegadoNombre) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racDelegadoNombre = '" & Replace(pcsubDatos.racDelegadoNombre, "'", "''") & "'"
        End If

        If camposSet <> "" Then
            sqlUpdate = "UPDATE tbDatosPCSUB SET " & camposSet & " WHERE idSolicitud = " & idSolicitud
            db.Execute sqlUpdate, dbFailOnError
        End If
    Else
        MsgBox "Tipo de solicitud no valido para importacion: " & tipoSolicitud, vbExclamation, "Importacion Word"
        GoTo LimpiarYSalir
    End If
    
    Select Case tipoSolicitud
        Case "PC"
            If Not pcDatos Is Nothing Then observacionesImportadas = Nz(pcDatos.observacionesRAC, "")
        Case "CD_CA"
            If Not cdDatos Is Nothing Then observacionesImportadas = Nz(cdDatos.observacionesRAC, "")
        Case "CD_CA_SUB"
            If Not subDatos Is Nothing Then observacionesImportadas = Nz(subDatos.observacionesRAC, "")
        Case "PC_SUB"
            If Not pcsubDatos Is Nothing Then observacionesImportadas = Nz(pcsubDatos.observacionesRAC, "")
    End Select
    
    Dim esCompleto As Boolean
    esCompleto = VerificarDictamenRACCompleto(idSolicitud, tipoSolicitud)
    
    Dim respuesta As Integer
    If esCompleto Then
        MsgBox "Comentarios importados correctamente." & vbCrLf & vbCrLf & "Se retrocederá a la fase de VALIDACIÓN para obtener una nueva versión del documento.", vbInformation, "Importación Word"
        Dim srvResRAC As ValidacionRevisionServicio
        Set srvResRAC = New ValidacionRevisionServicio
        Dim idVal As Long
        idVal = srvResRAC.GetUltimoPendienteId(idSolicitud)
        If idVal > 0 Then
            Call srvResRAC.RegistrarRespuestaRACPorId(idVal, "COMENTARIOS", "Importación Word desde RAC")
        End If
        Dim nuevoHash As String
        nuevoHash = SnapshotHelper.CalcularHashFaseValidacion(idSolicitud, tipoSolicitud, db)
        If Len(nuevoHash) > 0 Then
            Dim ultimoOrdinal As Long
            ultimoOrdinal = ValidacionRevisionRepositorio.GetUltimoOrdinal(idSolicitud)
            If ultimoOrdinal > 0 Then
                Call ValidacionRevisionRepositorio.ActualizarHashRevision(idSolicitud, ultimoOrdinal, hashAntesImportWeb)
            End If
        End If
        Call AccionRetrocederFlujo(estadoValidacion, "Revisión con Comentarios (Importación Word)")
    Else
        MsgBox "Comentarios importados correctamente." & vbCrLf & vbCrLf & "Se retrocederá a la fase de VALIDACIÓN para obtener una nueva versión del documento.", vbInformation, "Importación Word"
        Dim srvValIncompleto As ValidacionRevisionServicio
        Set srvValIncompleto = New ValidacionRevisionServicio
        Dim idValIncompleto As Long
        idValIncompleto = srvValIncompleto.GetUltimoPendienteId(idSolicitud)
        If idValIncompleto > 0 Then
            Dim comentariosIncompletos As String
            comentariosIncompletos = IIf(Len(observacionesImportadas) > 0, observacionesImportadas, "Importación Word (campos incompletos)")
            Call srvValIncompleto.RegistrarRespuestaRACPorId(idValIncompleto, "COMENTARIOS", "Importación Word (campos incompletos)")
        End If
        Dim nuevoHashIncompleto As String
        nuevoHashIncompleto = SnapshotHelper.CalcularHashFaseValidacion(idSolicitud, tipoSolicitud, db)
        If Len(nuevoHashIncompleto) > 0 Then
            Dim ultimoOrdinalIncompleto As Long
            ultimoOrdinalIncompleto = ValidacionRevisionRepositorio.GetUltimoOrdinal(idSolicitud)
            If ultimoOrdinalIncompleto > 0 Then
                Call ValidacionRevisionRepositorio.ActualizarHashRevision(idSolicitud, ultimoOrdinalIncompleto, hashAntesImportWeb)
            End If
        End If
        Call AccionRetrocederFlujo(estadoValidacion, "Revisión con Comentarios (Importación Word)")
    End If
    
    If CurrentProject.AllForms("frmWebVisor").IsLoaded Then DoCmd.Close acForm, "frmWebVisor"
    
    Call CargarDatosSolicitud(idSolicitud, mantenerVistaActual:=True)
    DoEvents
    Me.Repaint
    
LimpiarYSalir:
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Exit Sub
    
Errores:
    If CurrentProject.AllForms("frmWebVisor").IsLoaded Then DoCmd.Close acForm, "frmWebVisor"
    MsgBox "Error al importar datos: " & Err.description, vbExclamation, "Importacion Word"
    Resume LimpiarYSalir
End Sub

Private Sub ProcesarImportacionWordRAC_BotonDirecto()
    Dim docServ As New DocumentoServicio
    Dim tipoSolicitud As String
    Dim idSolicitud As Long
    Dim fDialog As FileDialog
    Dim filePath As String
    Dim db As DAO.Database
    Dim sqlUpdate As String
    Dim camposSet As String
    Dim observacionesImportadasBot As String
    Dim pcDatos As DatosPC
    Dim cdDatos As DatosCDCA
    Dim subDatos As DatosCDCASUB
    Dim pcsubDatosBot As Object
    
    On Error GoTo Errores
    
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    fDialog.Title = "Seleccione el documento Word con las observaciones del RAC"
    fDialog.Filters.Clear
    fDialog.Filters.Add "Documentos de Word", "*.docx; *.doc"
    If fDialog.Show <> -1 Then
        Exit Sub
    End If
    filePath = fDialog.SelectedItems(1)
    
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    
    idSolicitud = m_ViewModel.Solicitud.idSolicitud
    tipoSolicitud = m_ViewModel.Solicitud.tipoSolicitud
    
    Dim hashAntesImport As String
    hashAntesImport = SnapshotHelper.CalcularHashFaseValidacion(idSolicitud, tipoSolicitud, db)
    
    If tipoSolicitud = "PC" Then
        Set pcDatos = docServ.ImportarDatosRACDesdeWord(filePath)
        If pcDatos Is Nothing Then
            MsgBox "No se pudieron leer datos del archivo.", vbExclamation, "Importacion Word"
            GoTo LimpiarYSalir
        End If
        
        Set db = getdb()
        camposSet = ""
        If Not IsNull(pcDatos.observacionesRAC) And Trim(pcDatos.observacionesRAC) <> "" Then
            camposSet = "observacionesRAC = '" & Replace(pcDatos.observacionesRAC, "'", "''") & "'"
        End If
        If Not IsNull(pcDatos.racCodigo) And Trim(pcDatos.racCodigo) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racCodigo = '" & Replace(pcDatos.racCodigo, "'", "''") & "'"
        End If
        If Not IsNull(pcDatos.racNombre) And Trim(pcDatos.racNombre) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racNombre = '" & Replace(pcDatos.racNombre, "'", "''") & "'"
        End If
        
        If camposSet <> "" Then
            sqlUpdate = "UPDATE tbDatosPC SET " & camposSet & " WHERE idSolicitud = " & idSolicitud
            db.Execute sqlUpdate, dbFailOnError
        End If
        
    ElseIf tipoSolicitud = "CD_CA" Then
        Set cdDatos = docServ.ImportarDatosRACDesdeWord_CDCA(filePath)
        If cdDatos Is Nothing Then
            MsgBox "No se pudieron leer datos del archivo.", vbExclamation, "Importacion Word"
            GoTo LimpiarYSalir
        End If
        
        Set db = getdb()
        camposSet = ""
        If Not IsNull(cdDatos.observacionesRAC) And Trim(cdDatos.observacionesRAC) <> "" Then
            camposSet = "observacionesRAC = '" & Replace(cdDatos.observacionesRAC, "'", "''") & "'"
        End If
        If Not IsNull(cdDatos.racCodigo) And Trim(cdDatos.racCodigo) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racCodigo = '" & Replace(cdDatos.racCodigo, "'", "''") & "'"
        End If
        If Not IsNull(cdDatos.racNombre) And Trim(cdDatos.racNombre) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racNombre = '" & Replace(cdDatos.racNombre, "'", "''") & "'"
        End If
        
        If camposSet <> "" Then
            sqlUpdate = "UPDATE tbDatosCDCA SET " & camposSet & " WHERE idSolicitud = " & idSolicitud
            db.Execute sqlUpdate, dbFailOnError
        End If
        
    ElseIf tipoSolicitud = "CD_CA_SUB" Then
        Set subDatos = docServ.ImportarDatosRACDesdeWord_CDCASUB(filePath)
        If subDatos Is Nothing Then
            MsgBox "No se pudieron leer datos del archivo.", vbExclamation, "Importacion Word"
            GoTo LimpiarYSalir
        End If
        
        Set db = getdb()
        camposSet = ""
        If Not IsNull(subDatos.observacionesRAC) And Trim(subDatos.observacionesRAC) <> "" Then
            camposSet = "observacionesRAC = '" & Replace(subDatos.observacionesRAC, "'", "''") & "'"
        End If
        If Not IsNull(subDatos.racCodigo) And Trim(subDatos.racCodigo) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racCodigo = '" & Replace(subDatos.racCodigo, "'", "''") & "'"
        End If
        If Not IsNull(subDatos.racNombre) And Trim(subDatos.racNombre) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racNombre = '" & Replace(subDatos.racNombre, "'", "''") & "'"
        End If
        
        If camposSet <> "" Then
            sqlUpdate = "UPDATE tbDatosCDCASUB SET " & camposSet & " WHERE idSolicitud = " & idSolicitud
            db.Execute sqlUpdate, dbFailOnError
            
            ' GAP FIX: si no existe fila en tbDatosCDCASUB, el UPDATE no afecta registros.
            ' Hacemos UPSERT para garantizar que se guarden comentarios/código/nombre importados.
            If db.RecordsAffected = 0 Then
                Dim subSrvUpsertBot As New DatosCDCASUBServicio
                Dim subEntUpsertBot As DatosCDCASUB
                
                Set subEntUpsertBot = subSrvUpsertBot.getDatosCDCASUB(idSolicitud)
                If subEntUpsertBot Is Nothing Then
                    Set subEntUpsertBot = subSrvUpsertBot.GetDatosGeneralesIniciales(idSolicitud)
                    subEntUpsertBot.idSolicitud = idSolicitud
                End If
                
                If Not IsNull(subDatos.observacionesRAC) And Trim(subDatos.observacionesRAC) <> "" Then
                    subEntUpsertBot.observacionesRAC = subDatos.observacionesRAC
                End If
                If Not IsNull(subDatos.racCodigo) And Trim(subDatos.racCodigo) <> "" Then
                    subEntUpsertBot.racCodigo = subDatos.racCodigo
                End If
                If Not IsNull(subDatos.racNombre) And Trim(subDatos.racNombre) <> "" Then
                    subEntUpsertBot.racNombre = subDatos.racNombre
                End If
                
                Call DatosCDCASUBRepositorio.Guardar(subEntUpsertBot, db)
            End If
        End If
    ElseIf tipoSolicitud = "PC_SUB" Then
        ' PC_SUB usa la importación de PC (misma estructura RAC)
        Set pcsubDatosBot = docServ.ImportarDatosRACDesdeWord(filePath)
        If pcsubDatosBot Is Nothing Then
            MsgBox "No se pudieron leer datos del archivo.", vbExclamation, "Importacion Word"
            GoTo LimpiarYSalir
        End If

        Set db = getdb()
        camposSet = ""
        If Not IsNull(pcsubDatosBot.observacionesRAC) And Trim(pcsubDatosBot.observacionesRAC) <> "" Then
            camposSet = "observacionesRAC = '" & Replace(pcsubDatosBot.observacionesRAC, "'", "''") & "'"
        End If
        If Not IsNull(pcsubDatosBot.racCodigo) And Trim(pcsubDatosBot.racCodigo) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racCodigo = '" & Replace(pcsubDatosBot.racCodigo, "'", "''") & "'"
        End If
        If Not IsNull(pcsubDatosBot.racNombre) And Trim(pcsubDatosBot.racNombre) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racNombre = '" & Replace(pcsubDatosBot.racNombre, "'", "''") & "'"
        End If
        ' PC_SUB también tiene campos RACDelegado
        If Not IsNull(pcsubDatosBot.racDelegadoCodigo) And Trim(pcsubDatosBot.racDelegadoCodigo) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racDelegadoCodigo = '" & Replace(pcsubDatosBot.racDelegadoCodigo, "'", "''") & "'"
        End If
        If Not IsNull(pcsubDatosBot.racDelegadoObservaciones) And Trim(pcsubDatosBot.racDelegadoObservaciones) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racDelegadoObservaciones = '" & Replace(pcsubDatosBot.racDelegadoObservaciones, "'", "''") & "'"
        End If
        If Not IsNull(pcsubDatosBot.racDelegadoNombre) And Trim(pcsubDatosBot.racDelegadoNombre) <> "" Then
            If camposSet <> "" Then camposSet = camposSet & ", "
            camposSet = camposSet & "racDelegadoNombre = '" & Replace(pcsubDatosBot.racDelegadoNombre, "'", "''") & "'"
        End If

        If camposSet <> "" Then
            sqlUpdate = "UPDATE tbDatosPCSUB SET " & camposSet & " WHERE idSolicitud = " & idSolicitud
            db.Execute sqlUpdate, dbFailOnError
        End If
    Else
        MsgBox "Tipo de solicitud no valido: " & tipoSolicitud, vbExclamation, "Importacion Word"
        GoTo LimpiarYSalir
    End If
    
    
    Select Case tipoSolicitud
        Case "PC"
            If Not pcDatos Is Nothing Then observacionesImportadasBot = Nz(pcDatos.observacionesRAC, "")
        Case "CD_CA"
            If Not cdDatos Is Nothing Then observacionesImportadasBot = Nz(cdDatos.observacionesRAC, "")
        Case "CD_CA_SUB"
            If Not subDatos Is Nothing Then observacionesImportadasBot = Nz(subDatos.observacionesRAC, "")
        Case "PC_SUB"
            If Not pcsubDatosBot Is Nothing Then observacionesImportadasBot = Nz(pcsubDatosBot.observacionesRAC, "")
    End Select
    
    Dim esCompletoBot As Boolean
    esCompletoBot = VerificarDictamenRACCompleto(idSolicitud, tipoSolicitud)
    
    Dim respuestaBot As Integer
    If esCompletoBot Then
        MsgBox "Comentarios importados correctamente." & vbCrLf & vbCrLf & "Se retrocederá a la fase de VALIDACIÓN para obtener una nueva versión del documento.", vbInformation, "Importación Word"
        Dim srvResRACBot As ValidacionRevisionServicio
        Set srvResRACBot = New ValidacionRevisionServicio
        Dim idValBot As Long
        idValBot = srvResRACBot.GetUltimoPendienteId(idSolicitud)
        If idValBot > 0 Then
            Call srvResRACBot.RegistrarRespuestaRACPorId(idValBot, "COMENTARIOS", "Importación Word desde RAC")
        End If
        Dim nuevoHashBot As String
        nuevoHashBot = SnapshotHelper.CalcularHashFaseValidacion(idSolicitud, tipoSolicitud, db)
        If Len(nuevoHashBot) > 0 Then
            Dim ultimoOrdinalBot As Long
            ultimoOrdinalBot = ValidacionRevisionRepositorio.GetUltimoOrdinal(idSolicitud)
            If ultimoOrdinalBot > 0 Then
                Call ValidacionRevisionRepositorio.ActualizarHashRevision(idSolicitud, ultimoOrdinalBot, hashAntesImport)
            End If
        End If
        Call AccionRetrocederFlujo(estadoValidacion, "Revisión con Comentarios (Importación Word)")
    Else
        MsgBox "Comentarios importados correctamente." & vbCrLf & vbCrLf & "Se retrocederá a la fase de VALIDACIÓN para obtener una nueva versión del documento.", vbInformation, "Importación Word"
        Dim srvValIncompletoBot As ValidacionRevisionServicio
        Set srvValIncompletoBot = New ValidacionRevisionServicio
        Dim idValIncompletoBot As Long
        idValIncompletoBot = srvValIncompletoBot.GetUltimoPendienteId(idSolicitud)
        If idValIncompletoBot > 0 Then
            Dim comentariosBotIncompletos As String
            comentariosBotIncompletos = IIf(Len(observacionesImportadasBot) > 0, observacionesImportadasBot, "Importación Word (campos incompletos)")
            Call srvValIncompletoBot.RegistrarRespuestaRACPorId(idValIncompletoBot, "COMENTARIOS", "Importación Word (campos incompletos)")
        End If
        Dim nuevoHashIncompletoBot As String
        nuevoHashIncompletoBot = SnapshotHelper.CalcularHashFaseValidacion(idSolicitud, tipoSolicitud, db)
        If Len(nuevoHashIncompletoBot) > 0 Then
            Dim ultimoOrdinalIncompletoBot As Long
            ultimoOrdinalIncompletoBot = ValidacionRevisionRepositorio.GetUltimoOrdinal(idSolicitud)
            If ultimoOrdinalIncompletoBot > 0 Then
                Call ValidacionRevisionRepositorio.ActualizarHashRevision(idSolicitud, ultimoOrdinalIncompletoBot, hashAntesImport)
            End If
        End If
        Call AccionRetrocederFlujo(estadoValidacion, "Revisión con Comentarios (Importación Word)")
    End If
    
    Call CargarDatosSolicitud(idSolicitud, mantenerVistaActual:=True)
    DoEvents
    Me.Repaint
    
LimpiarYSalir:
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Exit Sub
    
Errores:
    MsgBox "Error al importar: " & Err.description, vbExclamation, "Importacion Word"
    Resume LimpiarYSalir
End Sub

Private Sub ProcesarDatosRechazoCalidad(ByVal data As String)
    Dim valor As String
    Dim modoFinal As String
    
    If Trim(Nz(data, "")) = "" Then Exit Sub
    
    If CurrentProject.AllForms("frmWebVisor").IsLoaded Then DoCmd.Close acForm, "frmWebVisor"
    
    data = UrlDecode(data)
    
    If UCase$(data) = "CANCELAR" Then
        If CurrentProject.AllForms("frmWebVisor").IsLoaded Then DoCmd.Close acForm, "frmWebVisor"
        Exit Sub
    End If
    
    valor = Trim$(data)
    If valor = "" Then
        MsgBox "Debe indicar un motivo.", vbExclamation, "CONDOR"
        Exit Sub
    End If
    
    modoFinal = UCase$(Trim$(m_ModoRechazoCalidad))
    
    If modoFinal = "EDICION" Then
        Dim revServ As New RevisionServicio
        Call revServ.ActualizarMotivoRechazoActivo(m_ViewModel.Solicitud.idSolicitud, valor)
    ElseIf modoFinal = "CAPTURA" Then
        Call AccionRechazarTecnica(valor)
    End If
    
    m_ModoRechazoCalidad = ""
    
    If CurrentProject.AllForms("frmWebVisor").IsLoaded Then DoCmd.Close acForm, "frmWebVisor"
    Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, mantenerVistaActual:=True)
    DoEvents
    Me.Repaint
End Sub

Private Function UrlDecode(ByVal value As String) As String
    Dim bytes() As Byte
    Dim i As Long
    Dim ch As String
    Dim hexPair As String
    Dim byteCount As Long
    Dim b As Byte
    Dim stm As Object
    
    If Len(value) = 0 Then
        UrlDecode = ""
        Exit Function
    End If
    
    ReDim bytes(0 To Len(value) - 1)
    
    i = 1
    Do While i <= Len(value)
        ch = Mid$(value, i, 1)
        If ch = "+" Then
            b = 32
        ElseIf ch = "%" And i + 2 <= Len(value) Then
            hexPair = Mid$(value, i + 1, 2)
            If IsHexPair(hexPair) Then
                b = CByte("&H" & hexPair)
                i = i + 2
            Else
                b = Asc(ch)
            End If
        Else
            b = Asc(ch)
        End If
        
        bytes(byteCount) = b
        byteCount = byteCount + 1
        i = i + 1
    Loop
    
    If byteCount = 0 Then
        UrlDecode = ""
        Exit Function
    End If
    
    ReDim Preserve bytes(0 To byteCount - 1)
    
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write bytes
    stm.Position = 0
    stm.Type = 2
    stm.Charset = "utf-8"
    UrlDecode = stm.ReadText
    stm.Close
End Function

Private Function IsHexPair(ByVal value As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim validChars As String
    
    If Len(value) <> 2 Then
        IsHexPair = False
        Exit Function
    End If
    
    validChars = "0123456789ABCDEF"
    For i = 1 To 2
        ch = UCase$(Mid$(value, i, 1))
        If InStr(1, validChars, ch, vbBinaryCompare) = 0 Then
            IsHexPair = False
            Exit Function
        End If
    Next i
    
    IsHexPair = True
End Function
Private Sub webInfo_BeforeNavigate(Cancel As Integer, URL As String)
    Dim mensaje As String, partes() As String
    
    On Error GoTo Errores
    
    ' Interceptamos solo las llamadas que empiezan por nuestro protocolo personalizado "vba:"
    If Left(URL, 4) = "vba:" Then
        ' 1. Cancelamos la navegación para que el navegador no haga nada más
        Cancel = True
        
        ' 2. Quitamos el foco del navegador
        Me.cmdCambiarEstado.SetFocus
        
        ' 3. Parseamos el mensaje
        mensaje = Mid(URL, 5)
        partes = Split(mensaje, ":")
        
        If UBound(partes) >= 0 Then
            ' 4. GUARDAMOS LA INTENCIÓN (Siempre existe la parte 0)
            m_AccionPendiente = partes(0)
            
            ' 5. GUARDAMOS EL PARÁMETRO (SOLO SI EXISTE LA PARTE 1)
            ' --- CORRECCIÓN DEL ERROR 9 ---
            If UBound(partes) >= 1 Then
                If IsNumeric(partes(1)) Then
                    m_ParametroPendiente = CLng(partes(1))
                    m_ParametroTextoPendiente = ""
                Else
                    m_ParametroPendiente = 0
                    m_ParametroTextoPendiente = partes(1)
                End If
            Else
                m_ParametroPendiente = 0
                m_ParametroTextoPendiente = ""
            End If
            ' ------------------------------
            
            ' 6. ACTIVAMOS EL TIMER
            Me.TimerInterval = 50
        End If
    End If

LimpiarYSalir:
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.webInfo_BeforeNavigate")
    Resume LimpiarYSalir
End Sub
' --- PROCEDIMIENTOS DE ACCIÓN (Llamados desde eventos) ---

Private Sub AccionAvanzarEstado(ByVal idEstadoDestino As Long)
    Dim wfServicio As New WorkflowServicio
    On Error GoTo Errores
    Me.lblBloqueo.Visible = True
    
    DoCmd.Hourglass True
    
    ' --- REGLA DE NEGOCIO: PROMOCIÓN MANUAL POR CALIDAD (3 -> 4) ---
    ' Si Calidad decide avanzar desde Desarrollo Técnico, se asume que toma el control
    ' y se limpia cualquier rechazo previo para reiniciar el ciclo.
    If m_ViewModel.Solicitud.idEstadoInterno = estadoDesarrolloTecnico And _
       idEstadoDestino = estadoModificacion And _
       m_ObjUsuarioActivo.rol = rol.Calidad Then
       
       ' Verificamos si existe un rechazo activo que deba ser limpiado
       If Not RechazoRepositorio.GetUltimoRechazoActivo(m_ViewModel.Solicitud.idSolicitud) Is Nothing Then
           If MsgBox("ATENCIÓN: La solicitud tiene un rechazo técnico activo." & vbCrLf & vbCrLf & _
                     "Al promocionarla manualmente a 'MODIFICACIÓN', se eliminará el rechazo y sus motivos." & vbCrLf & _
                     "El historial de cambios (Delta) se reiniciará y no será visible." & vbCrLf & vbCrLf & _
                     "¿Está seguro de que desea continuar?", _
                     vbExclamation + vbYesNo + vbDefaultButton2, "Confirmar Promoción Manual") = vbNo Then
               GoTo LimpiarYSalir
           End If
           
           ' Ejecutar limpieza (Fuera de la transacción de EjecutarTransicion, pero aceptable aquí)
           Call RechazoRepositorio.DesactivarRechazosPrevios(m_ViewModel.Solicitud.idSolicitud)
       End If
    End If
    ' ---------------------------------------------------------------
    
    Call wfServicio.EjecutarTransicion(m_ViewModel.Solicitud, idEstadoDestino, m_ObjUsuarioActivo)
    Call RefrescarTablerosPrincipales
    MsgBox "El estado de la solicitud ha sido actualizado correctamente.", vbInformation, "Cambio de Estado"
    RaiseEvent SolicitudModificada
    RaiseEvent SolicitudCambiada("Avanzar")
    Call NotificarCambioContadorTareas
    
    ' Spec-014: Asegurar refresco visual tras cambio de estado
    DoEvents
    Me.Repaint ' Spec-014: Forzar repintado de Access antes de cargar HTML
    
    ' Spec-018: Refresco visual delegado a CargarDatosSolicitud -> MostrarVista
    ' La generación de URL única garantiza el repintado sin about:blank
    
    ' PERSISTENCIA VISTA: True
    Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, True)
    If m_VistaActual = VistaCambioEstado Then
        Call MostrarVista(VistaCambioEstado)
    End If
LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.AccionAvanzarEstado")
    Resume LimpiarYSalir
End Sub

Private Sub AccionRechazarTecnica(ByVal motivo As String)
    Dim wfServ As New WorkflowServicio
    
    On Error GoTo Errores
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    
    ' 2. Llamar al servicio
    ' NOTA: RechazarFaseTecnica ya gestiona la transacción y el cambio de estado (4 -> 3)
    ' CORRECCIÓN: Pasamos el motivo como String, el servicio se encarga de crear el objeto Rechazo.
    Call wfServ.RechazarFaseTecnica(m_ViewModel.Solicitud, motivo, m_ObjUsuarioActivo)
    
    ' 3. Refrescar
    Call RefrescarTablerosPrincipales
    MsgBox "La solicitud ha sido devuelta a la fase de Desarrollo Técnico.", vbInformation, "Rechazo Completado"
    RaiseEvent SolicitudModificada
    RaiseEvent SolicitudCambiada("Rechazar")
    Call NotificarCambioContadorTareas
    
    ' Spec-014: Asegurar refresco visual tras cambio de estado
    DoEvents
    Me.Repaint ' Spec-014: Forzar repintado de Access antes de cargar HTML
    
    ' Spec-018: Refresco visual delegado a CargarDatosSolicitud -> MostrarVista
    ' La generación de URL única garantiza el repintado sin about:blank
    
    ' PERSISTENCIA VISTA: True
    Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, True)

LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.AccionRechazarTecnica")
    Resume LimpiarYSalir
End Sub



Private Sub PoblarControlesDesdeViewModel()
    Dim esCerrado As Boolean
    Dim estadoActualID As Long
    
    ' CORRECCIÓN: Usamos manejo de errores estándar para detectar fallos de datos
    On Error GoTo Errores
    
    estadoActualID = m_ViewModel.Solicitud.idEstadoInterno
    
    ' Uso seguro del diccionario de estados
    If m_ObjEntorno.estados.Exists(CStr(estadoActualID)) Then
        esCerrado = m_ObjEntorno.estados(CStr(estadoActualID)).esEstadoFinal
    Else
        esCerrado = False
    End If
    
    ' Asignación de botones (Protegemos cada bloque si tememos que el control no exista,
    ' pero lo ideal es saber si fallan)
    
    SetPropiedadControl Me, "cmdCambiarEstado", "Enabled", True
    SetPropiedadControl Me, "cmdVerHistorial", "Enabled", True
    SetPropiedadControl Me, "cmdGestionarAdjuntos", "Enabled", True
    SetPropiedadControl Me, "cmdVerDatos", "Enabled", True

    ' --- GRUPO 2: ACCIONES DE FLUJO ---
    Dim puedeEditar As Boolean
    puedeEditar = (estadoActualID < estadoValidacion) And (rolUsuario = rol.Calidad Or rolUsuario = rol.Administrador) And Not esCerrado
    SetPropiedadControl Me, "cmdEditarPreregistro", "Enabled", puedeEditar
    
    Dim puedeEliminar As Boolean
    puedeEliminar = (rolUsuario = rol.Administrador Or rolUsuario = rol.Calidad) And (Not esCerrado Or estadoActualID = estadoRechazada)
    SetPropiedadControl Me, "cmdEliminarSolicitud", "Enabled", puedeEliminar

    ' --- GRUPO 3: DOCUMENTOS ---
    
    ' Estado Visual
    On Error Resume Next
    Me.txtEstado.value = m_ViewModel.EstadoNombre
    If Err.Number <> 0 Then
        ' LogDebug removed — Debug.Print "PoblarControles: txtEstado assign failed..."
        Err.Clear
    End If
    On Error GoTo Errores
   
    Exit Sub
    
Errores:
    ' Si falló algo visual, lo registramos pero no bloqueamos al usuario
    Resume Next
End Sub

' Helper local para asignar propiedades sin romper si el control no existe (Sustituto seguro de Resume Next)
Private Sub SetPropiedadControl(frm As Form, ctlName As String, propName As String, valor As Variant)
    On Error Resume Next
    frm.Controls(ctlName).Properties(propName) = valor
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub m_frmDatosPC_SolicitudHaCambiado()
    On Error GoTo Errores
    Dim idEstadoPrevio As Long
    idEstadoPrevio = m_ViewModel.Solicitud.idEstadoInterno
    
    ' Cargamos datos solo en memoria para no robar el foco (Spec-048 V3)
    Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, mantenerVistaActual:=True, soloMemoria:=True)
    
    If m_ViewModel.Solicitud.idEstadoInterno = idEstadoPrevio Then
        ' Refresco inyectado por JS (Silencioso)
        Call RefrescarTimelineSilencioso
    Else
        ' Cambio de fase real: refresco visual total necesario
        Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, mantenerVistaActual:=True)
    End If
    
    RaiseEvent SolicitudModificada
LimpiarYSalir:
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.m_frmDatosPC_SolicitudHaCambiado")
    Resume LimpiarYSalir
End Sub

' --- OTROS PROCEDIMIENTOS (Restaurados del Bundle) ---

Private Sub cmdEditarPreregistro_Click()
    On Error GoTo Errores
    If m_ViewModel Is Nothing Then Exit Sub
    DoCmd.OpenForm "frmAltaSolicitud", acNormal, , , acFormEdit, acDialog, CStr(m_ViewModel.Solicitud.idSolicitud)
    Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud)
    RaiseEvent SolicitudModificada
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.cmdEditarPreregistro_Click")
End Sub


Private Sub cmdEliminarSolicitud_Click()
    On Error GoTo Errores
    If m_ViewModel Is Nothing Then Exit Sub
    
    ' --- SPEC-131: Verificar si hay NC vinculada antes de eliminar ---
    Dim ncServ As New NoConformidadServicio
    Dim ncVinculada As NoConformidad
    Dim mensajeConfirmacion As String
    
    Set ncVinculada = ncServ.getNoConformidadPorCodigoCondor(m_ViewModel.Solicitud.codigoSolicitud)
    
    If Not ncVinculada Is Nothing Then
        mensajeConfirmacion = "La solicitud '" & m_ViewModel.Solicitud.codigoSolicitud & "' está asociada a la No Conformidad '" & ncVinculada.CodigoNoConformidad & "'." & vbCrLf & vbCrLf & "¿Está seguro de que desea eliminarla? Se eliminará la relación en la base de datos de No Conformidades." & vbCrLf & vbCrLf & "Esta acción borrará la solicitud, todos sus datos asociados, el historial y los archivos adjuntos." & vbCrLf & "ESTA ACCIÓN NO SE PUEDE DESHACER."
    Else
        mensajeConfirmacion = "¿Está seguro de que desea eliminar permanentemente la solicitud '" & m_ViewModel.Solicitud.codigoSolicitud & "'?" & vbCrLf & vbCrLf & "Esta acción borrará la solicitud, todos sus datos asociados, el historial y los archivos adjuntos." & vbCrLf & "ESTA ACCIÓN NO SE PUEDE DESHACER."
    End If
    
    If MsgBox(mensajeConfirmacion, vbCritical + vbYesNo + vbDefaultButton2, "Confirmar Eliminación Permanente") = vbNo Then Exit Sub
    ' ---------------------------------------------------------------------
    
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    Call m_solServ.EliminarSolicitudCompleta(m_ViewModel.Solicitud.idSolicitud)
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    MsgBox "La solicitud ha sido eliminada con éxito.", vbInformation, "Operación Completada"
    Call RefrescarTablerosPrincipales
    RaiseEvent SolicitudEliminada
    RaiseEvent SolicitudCambiada("Eliminar")
    Call NotificarCambioContadorTareas
    DoCmd.Close acForm, Me.name
LimpiarYSalir:
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.cmdEliminarSolicitud_Click")
    Resume LimpiarYSalir
End Sub

Private Sub AccionReabrir()
    Dim wfServ As New WorkflowServicio
    Dim db As DAO.Database
    On Error GoTo Errores
    If MsgBox("¿Está seguro de que desea reabrir esta solicitud?", vbQuestion + vbYesNo, "Confirmar Reapertura") = vbNo Then Exit Sub
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    Set db = getdb()
    Call wfServ.ReabrirSolicitudCerrada(m_ViewModel.Solicitud, m_ObjUsuarioActivo, db)
    
    MsgBox "La solicitud ha sido reabierta con éxito.", vbInformation, "Operación Completada"
    RaiseEvent SolicitudModificada
    
    ' Spec-014: Asegurar refresco visual tras cambio de estado
    DoEvents
    Me.Repaint ' Spec-014: Forzar repintado de Access antes de cargar HTML
    
    ' Spec-018: Refresco visual delegado a CargarDatosSolicitud -> MostrarVista
    ' La generación de URL única garantiza el repintado sin about:blank
    
    ' PERSISTENCIA VISTA: True
    Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, True)
LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Set wfServ = Nothing
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("Form_frmGestionSolicitud.AccionReabrir")
    Resume LimpiarYSalir
End Sub

Private Sub AccionCerrar()
    Dim adjServ As New AdjuntosServicio
    Dim wfServ As New WorkflowServicio
    Const ETAPA_CIERRE As String = "Documento de Cierre"
    On Error GoTo Errores
    If Not adjServ.ExisteAdjuntoEtapa(m_ViewModel.Solicitud.idSolicitud, ETAPA_CIERRE) Then
        If MsgBox("Para cerrar la solicitud es necesario adjuntar el '" & ETAPA_CIERRE & "'." & vbCrLf & vbCrLf & "¿Desea adjuntarlo ahora?", vbQuestion + vbYesNo, "Adjunto Requerido") = vbYes Then
            Dim fDialog As FileDialog, filePath As String
            Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
            fDialog.Title = "Seleccione el " & ETAPA_CIERRE & " firmado"
            fDialog.AllowMultiSelect = False
            If fDialog.Show = -1 Then
                filePath = fDialog.SelectedItems(1)
                Me.lblBloqueo.Visible = True
                DoCmd.Hourglass True
                Call adjServ.GuardarAdjuntoDesdeArchivo(m_ViewModel.Solicitud.idSolicitud, ETAPA_CIERRE, filePath)
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    If MsgBox("¿Desea cerrar definitivamente la solicitud '" & m_ViewModel.Solicitud.codigoSolicitud & "'?", vbQuestion + vbYesNo, "Confirmar Cierre") = vbYes Then
        Me.lblBloqueo.Visible = True
        DoCmd.Hourglass True
        Call wfServ.EjecutarTransicion(m_ViewModel.Solicitud, estadoAprobada, m_ObjUsuarioActivo)
        
        MsgBox "La solicitud ha sido cerrada con éxito.", vbInformation, "Proceso Completado"
        RaiseEvent SolicitudModificada
        
        ' Spec-014: Asegurar refresco visual tras cambio de estado
        DoEvents
        Me.Repaint ' Spec-014: Forzar repintado de Access antes de cargar HTML
        ' PERSISTENCIA VISTA: True
        Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, True)
    End If
LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Set adjServ = Nothing
    Set wfServ = Nothing
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("Form_frmGestionSolicitud.AccionCerrar")
    Resume LimpiarYSalir
End Sub

' ==========================================================================
' MÓDULO: Form_frmGestionSolicitud.cls
' ACCIÓN: AÑADIR este nuevo manejador de eventos para el formulario CDCA.
' ==========================================================================
Private Sub m_frmDatosCDCA_SolicitudHaCambiado()
    On Error GoTo Errores
    Dim idEstadoPrevio As Long
    idEstadoPrevio = m_ViewModel.Solicitud.idEstadoInterno
    
    ' Cargamos datos solo en memoria para no robar el foco (Spec-048 V3)
    Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, mantenerVistaActual:=True, soloMemoria:=True)
    
    If m_ViewModel.Solicitud.idEstadoInterno = idEstadoPrevio Then
        ' Refresco inyectado por JS (Silencioso)
        Call RefrescarTimelineSilencioso
    Else
        ' Cambio de fase real: refresco visual total necesario
        Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, mantenerVistaActual:=True)
    End If
    
    RaiseEvent SolicitudModificada
LimpiarYSalir:
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.m_frmDatosCDCA_SolicitudHaCambiado")
    Resume LimpiarYSalir
End Sub


Public Sub NotificarCambioExterno()
    On Error Resume Next
    
    ' 1. Actualizar contadores globales (Buscadores)
    Call RefrescarTablerosPrincipales
    
    ' 2. Lanzar evento para otros escuchas
    RaiseEvent SolicitudModificada
    
    ' 3. CRÍTICO: Recargar ESTE formulario para actualizar la UI
    If Not m_ViewModel Is Nothing Then
         Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, mantenerVistaActual:=True)
    End If
    
    If Err.Number <> 0 Then
        Dim logSvcExt As LogServicio
        Set logSvcExt = New LogServicio
        logSvcExt.RegistrarError Err.Number, Err.description, "frmGestionSolicitud", "NotificarCambioExterno"
        Err.Clear
    End If
End Sub

Private Sub AccionAprobarYAvanzar(ByVal idEstadoDestino As Long)
    Dim revServ As New RevisionServicio
    Dim wfServ As New WorkflowServicio
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim comentarioAuto As String
    
    On Error GoTo Errores
    
    If MsgBox("Va a aprobar la parte técnica y avanzar a 'VALIDACIÓN'." & vbCrLf & vbCrLf & _
              "Esto registrará internamente la decisión de Calidad como 'APROBADO'." & vbCrLf & _
              "¿Desea continuar?", vbQuestion + vbYesNo, "Aprobar y Avanzar") <> vbYes Then
        Exit Sub
    End If
    
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    
    comentarioAuto = "Aprobado automáticamente al avanzar fase."
    
    ' 1. INICIAR TRANSACCIÓN
    Set ws = getWorkspace()
    Set db = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
    ws.BeginTrans
    
    ' 2. Guardar en BD (Aprobación)
    Call revServ.GuardarDecisionRevision(m_ViewModel.Solicitud.idSolicitud, "APROBADO", comentarioAuto, db)
    
    ' --- CORRECCIÓN CRÍTICA: ACTUALIZAR MEMORIA ("Estado Rancio") ---
    ' Actualizamos el objeto en memoria para que coincida con lo que acabamos de guardar en BD.
    ' Si no hacemos esto, EjecutarTransicion guardará el valor antiguo ("PENDIENTE") encima del nuevo.
    m_ViewModel.Solicitud.revisionCalidadEstado = "APROBADO"
    m_ViewModel.Solicitud.revisionCalidadComentarios = comentarioAuto
    ' ----------------------------------------------------------------
    
    ' 3. Ejecutar transición (Usando el objeto ya actualizado en memoria)
    Call wfServ.EjecutarTransicion(m_ViewModel.Solicitud, idEstadoDestino, m_ObjUsuarioActivo, db)
    
    ' 4. COMMIT
    ws.CommitTrans
    
    Call RefrescarTablerosPrincipales
    MsgBox "La fase técnica ha sido aprobada y la solicitud ha avanzado a 'VALIDACIÓN'.", vbInformation, "Aprobación Exitosa"
    RaiseEvent SolicitudModificada
    RaiseEvent SolicitudCambiada("Aprobar")
    Call NotificarCambioContadorTareas
    
    ' Spec-014: Asegurar refresco visual tras cambio de estado
    DoEvents
    Me.Repaint ' Spec-014: Forzar repintado de Access antes de cargar HTML
    
    ' PERSISTENCIA VISTA: True
    Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, True)

LimpiarYSalir:
    On Error Resume Next
    If Not db Is Nothing Then db.Close
    Set ws = Nothing
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Set revServ = Nothing
    Set wfServ = Nothing
    Exit Sub
    
Errores:
    If Not ws Is Nothing Then ws.Rollback
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.AccionAprobarYAvanzar")
    Resume LimpiarYSalir
End Sub

' ==========================================================================
' MÓDULO: Form_frmGestionSolicitud.cls
' ACCIÓN: Procedimiento para devolver a Modificación revocando la aprobación técnica.
' ==========================================================================
Private Sub AccionRevertirTecnica(ByVal idEstadoDestino As Long)
    Dim wfServicio As New WorkflowServicio
    Dim sol As Solicitud
    
    On Error GoTo Errores
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    
    ' 1. Obtener la versión más reciente de la solicitud
    Set sol = m_solServ.getSolicitudPorID(m_ViewModel.Solicitud.idSolicitud)
    
    ' 2. RESETEO EXPLÍCITO DE LA APROBACIÓN
    ' Antes de moverla, "rompemos" la aprobación para que quede pendiente de nuevo.
    sol.revisionCalidadEstado = "PENDIENTE"
    
    ' Guardamos este cambio intermedio
    Call SolicitudRepositorio.GuardarSolicitud(sol)
    
    ' 3. EJECUTAR REVERSIÓN DE ESTADO (Validación -> Modificación)
    ' Usamos RevertirAFaseAnterior para que quede constancia en el log de que fue un retroceso
    Call wfServicio.RevertirAFaseAnterior_FromUI(sol, idEstadoDestino, m_ObjUsuarioActivo, "Revocación de Aprobación Técnica")
    
    ' 4. Refresco y Feedback
    Call RefrescarTablerosPrincipales
    
    MsgBox "La solicitud ha vuelto a 'MODIFICACIÓN' y la aprobación técnica ha sido revocada.", vbInformation, "Estado Revertido"
    
    RaiseEvent SolicitudModificada
    
    ' Spec-014: Asegurar refresco visual tras cambio de estado
    DoEvents
    Me.Repaint ' Spec-014: Forzar repintado de Access antes de cargar HTML
    
    ' PERSISTENCIA VISTA: True
    Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, True)

LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Set wfServicio = Nothing
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.AccionRevertirTecnica")
    Resume LimpiarYSalir
End Sub
' ==========================================================================
' DEBUG HELPER: Función temporal para depurar AccionAprobarYAvanzar
' Ejecutar desde la Ventana Inmediato: Call Forms("frmGestionSolicitud").Debug_ForzarAprobacion
' ==========================================================================
Public Sub Debug_ForzarAprobacion()
    On Error GoTo Errores
    
    ' 1. Definimos el ID del estado destino (VALIDACIÓN = 5)
    ' Usamos el número directo o la constante si es accesible aquí.
    Dim idEstadoDestino As Long
    idEstadoDestino = estadoValidacion
    
    ' 2. Llamada directa: Aquí puedes poner tu Punto de Interrupción (F9)
    Call AccionAprobarYAvanzar(idEstadoDestino)
    
    Exit Sub

Errores:
    MsgBox "Error en Debug: " & Err.description, vbExclamation, "CONDOR"
End Sub
' ==========================================================================
' DEBUG HELPER: Función para depurar la reversión a REVISIÓN sin usar el navegador
' Ejecutar desde Inmediato: Call Forms("frmGestionSolicitud").Debug_RevertirARevision
' ==========================================================================
Public Sub Debug_RevertirARevision()
    Dim wfServ As New WorkflowServicio
    Dim sol As Solicitud
    Dim idEstadoDestino As Long
    
    On Error GoTo Errores
    
    ' 1. Definimos el estado destino: REVISIÓN (ID 6)
    idEstadoDestino = 6
    
    ' 2. Obtenemos la solicitud fresca
    If m_ViewModel Is Nothing Then
        Exit Sub
    End If
    Set sol = m_solServ.getSolicitudPorID(m_ViewModel.Solicitud.idSolicitud)
    
    ' 3. PUNTO DE INTERRUPCIÓN RECOMENDADO AQUÍ (F9)
    ' Llamamos a la función que limpia datos y cambia el estado
    Call wfServ.RevertirAFaseAnterior_FromUI(sol, idEstadoDestino, m_ObjUsuarioActivo, "DEBUG: Reversión manual")
    
    ' 4. Refrescar la pantalla
    Call RefrescarTablerosPrincipales
    RaiseEvent SolicitudModificada
    Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud)
    
    MsgBox "Debug Finalizado: Solicitud revertida a REVISIÓN.", vbInformation, "CONDOR"
    Exit Sub

Errores:
    MsgBox "Error en Debug: " & Err.Number & " - " & Err.description, vbExclamation, "CONDOR"
End Sub

' ==========================================================================
' DEBUG HELPER: Función para depurar AccionRechazarTecnica
' Ejecutar desde Inmediato: Call Forms("frmGestionSolicitud").Debug_RechazarTecnica
' ==========================================================================
Public Sub Debug_RechazarTecnica()
    Dim motivo As String
    
    On Error GoTo Errores
    
    motivo = "DEBUG: Rechazo forzado desde ventana inmediato."
    
    If m_ViewModel Is Nothing Then
        Exit Sub
    End If
    
    ' Simular la llamada
    ' RECOMENDACIÓN: Poner punto de interrupción (F9) en la línea siguiente para entrar en la función
    Call AccionRechazarTecnica(motivo)
    
    Exit Sub
    
Errores:
    MsgBox "Error en Debug: " & Err.Number & " - " & Err.description, vbExclamation, "CONDOR"
End Sub
' EN: Form_frmGestionSolicitud.cls

Private Sub InyectarMotivoRechazoAutomatico(ByRef sol As Solicitud)
    Dim motivoAuto As String
    motivoAuto = "Rechazado desde el visor de cambio de estado"
    
    On Error GoTo Errores
    
    If sol.tipoSolicitud = "PC" Then
        Dim pcServ As New DatosPCServicio
        Dim pc As DatosPC
        
        ' Obtenemos el objeto para editarlo
        Set pc = pcServ.getDatosPC(sol.idSolicitud)
        If pc Is Nothing Then
            Set pc = New DatosPC
            pc.idSolicitud = sol.idSolicitud
        End If
        
        ' Escribimos directamente la decisión y el motivo
        pc.racDecision = "RECHAZADO"
        pc.racRechazoMotivos = motivoAuto
        
        ' Guardamos usando el método específico que valida reglas de RAC
        ' (Al ser RECHAZADO, permitirá guardar aunque falten fecha/firma)
        Call pcServ.GuardarDictamenRAC(pc)
        
    ElseIf sol.tipoSolicitud = "CD_CA" Then
        Dim cdcaServ As New DatosCDCAServicio
        Dim cdca As DatosCDCA
        
        Set cdca = cdcaServ.GetDatosCDCA(sol.idSolicitud)
        If cdca Is Nothing Then
            Set cdca = New DatosCDCA
            cdca.idSolicitud = sol.idSolicitud
        End If
        
        
        cdca.racDecision = "RECHAZADO"

        Call cdcaServ.GuardarDictamenRAC(cdca)
    ElseIf sol.tipoSolicitud = "PC_SUB" Then
        Dim pcsubServ As New DatosPCSUBServicio
        Dim pcsub As DatosPCSUB

        Set pcsub = pcsubServ.getDatosPCSUB(sol.idSolicitud)
        If pcsub Is Nothing Then
            Set pcsub = New DatosPCSUB
            pcsub.idSolicitud = sol.idSolicitud
        End If

        pcsub.racDecision = "RECHAZADO"
        pcsub.racRechazoMotivos = motivoAuto

        Call pcsubServ.GuardarDictamenRAC(pcsub)
    End If
    
    Exit Sub
Errores:
    ' Si falló la inyección automática, avisamos pero no detenemos la ejecución drásticamente
    MsgBox "Advertencia: No se pudo registrar el motivo automático del rechazo.", vbExclamation, "CONDOR"
End Sub

Private Sub webInfo_DocumentComplete(URL As Variant)
     On Error GoTo Errores
    
    ' Solo ejecutamos la lógica antigua si el flag NO está activado
    If Not EsRenderizadoFicheroActivo Then
        If URL = "about:blank" And m_jsScriptParaInyectar <> "" Then
            Me.webInfo.ExecuteJavascript m_jsScriptParaInyectar
            m_jsScriptParaInyectar = ""
        End If
    End If
    
LimpiarYSalir:
    Exit Sub
Errores:
    m_jsScriptParaInyectar = ""
    ' Call ErrorPresenter.MostrarError("frmGestionSolicitud.webInfo_DocumentComplete") ' Opcional silenciar
    Resume LimpiarYSalir
End Sub
' ==========================================================================
' DEBUG HELPER: Simular Clic en "Aprobar y Avanzar" (Caso Especial Modificación)
' Ejecutar desde Inmediato: Call Forms("frmGestionSolicitud").Debug_BotonAprobarTecnica
' ==========================================================================
Public Sub Debug_BotonAprobarTecnica()
    On Error GoTo Errores
    
    ' 1. Validaciones de seguridad para el Debug
    If m_ViewModel Is Nothing Then
        Exit Sub
    End If
    
    ' Verificamos que estamos en el estado correcto (MODIFICACIÓN = 4)
    If m_ViewModel.Solicitud.idEstadoInterno <> estadoModificacion Then ' 4 = estadoModificacion
        ' Estado incorrecto - la acción puede no funcionar como esperado
    End If

    ' 2. LLAMADA DIRECTA A LA ACCIÓN
    ' El botón HTML llama a 'vba:APROBAR_TECNICA:5'.
    ' Esto se traduce en el método AccionAprobarYAvanzar pasando el ID 5 (VALIDACIÓN).
    
    Call AccionAprobarYAvanzar(estadoValidacion)
    
    Exit Sub

Errores:
    MsgBox "Error Debug: " & Err.description, vbExclamation, "CONDOR"
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
End Sub
' EN: Form_frmGestionSolicitud.cls


Private Sub AccionRetrocederFlujo(ByVal idEstadoDestino As Long, ByVal motivo As String)
    Dim wfServicio As New WorkflowServicio
    Dim sol As Solicitud
    Dim tipoSolicitud As String
    On Error GoTo Errores
    Me.lblBloqueo.Visible = True: DoCmd.Hourglass True
    
    tipoSolicitud = m_ViewModel.Solicitud.tipoSolicitud
    
    Set sol = m_solServ.getSolicitudPorID(m_ViewModel.Solicitud.idSolicitud)
    Call wfServicio.RevertirFlujo(sol, idEstadoDestino, m_ObjUsuarioActivo, motivo)
    Call RefrescarTablerosPrincipales
    RaiseEvent SolicitudModificada
    Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, True)
LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False: DoCmd.Hourglass False
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.AccionRetrocederFlujo")
    Resume LimpiarYSalir
End Sub

Private Sub AccionRevertirGenerica(ByVal idEstadoDestino As Long, ByVal motivo As String)
    Dim wfServicio As New WorkflowServicio
    Dim sol As Solicitud
    
    On Error GoTo Errores
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    
    Set sol = m_solServ.getSolicitudPorID(m_ViewModel.Solicitud.idSolicitud)
    
    ' Usamos el servicio de workflow que ya sabe limpiar datos según el estado origen
    Call wfServicio.RevertirAFaseAnterior_FromUI(sol, idEstadoDestino, m_ObjUsuarioActivo, motivo)
    
    Call RefrescarTablerosPrincipales
    MsgBox "La solicitud ha retrocedido correctamente.", vbInformation, "Estado Revertido"
    
    RaiseEvent SolicitudModificada
    
    ' Spec-014: Asegurar refresco visual tras cambio de estado
    DoEvents
    Me.Repaint ' Spec-014: Forzar repintado de Access antes de cargar HTML
    
    ' PERSISTENCIA VISTA: True
    Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, True)

LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Set wfServicio = Nothing
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.AccionRevertirGenerica")
    Resume LimpiarYSalir
End Sub
' ==========================================================================
' MÓDULO: Form_frmGestionSolicitud.cls
' ACCIÓN: Añadir método público para forzar la apertura de datos desde fuera
' ==========================================================================

Public Sub AbrirFormularioDeDatos()
    ' RESPONSABILIDAD: Método público para que frm0Ppal pueda forzar la apertura
    '                  de la ficha de datos inmediatamente después de abrir este formulario.
    '                  Reutiliza la lógica del botón "Ver Datos".
    
    Call cmdVerDatos_Click
End Sub
Private Sub m_frmDatosCDCASUB_SolicitudHaCambiado()
    On Error GoTo Errores
    Dim idEstadoPrevio As Long
    idEstadoPrevio = m_ViewModel.Solicitud.idEstadoInterno
    
    ' Cargamos datos solo en memoria para no robar el foco (Spec-048 V3)
    Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, mantenerVistaActual:=True, soloMemoria:=True)
    
    If m_ViewModel.Solicitud.idEstadoInterno = idEstadoPrevio Then
        ' Refresco inyectado por JS (Silencioso)
        Call RefrescarTimelineSilencioso
    Else
        ' Cambio de fase real: refresco visual total necesario
        Call CargarDatosSolicitud(m_ViewModel.Solicitud.idSolicitud, mantenerVistaActual:=True)
    End If
    
    RaiseEvent SolicitudModificada
LimpiarYSalir:
    Exit Sub
Errores:
    Call ErrorPresenter.MostrarError("frmGestionSolicitud.m_frmDatosCDCASUB_SolicitudHaCambiado")
    Resume LimpiarYSalir
End Sub

Public Sub RefrescarTimelineSilencioso()
    Dim wfServ As New WorkflowServicio
    Dim html As String
    Dim errNum As Long, errDesc As String
    
    On Error Resume Next
    
    ' Solo aplica si estamos en la vista de cambio de estado (Timeline)
    If m_VistaActual <> VistaCambioEstado Then Exit Sub
    
    html = wfServ.GenerarHTML_VisualizadorDeEstado(m_ViewModel.Solicitud.idSolicitud)
    errNum = Err.Number: errDesc = Err.description: If errNum <> 0 Then Err.Clear
    
    ' Extraer solo el contenido de la clase timeline para inyectar
    Dim inicio As Long, Fin As Long
    inicio = InStr(1, html, "class='timeline'>")
    If inicio > 0 Then
        inicio = inicio + Len("class='timeline'>")
        Fin = InStrRev(html, "</div><script>")
        If Fin > inicio Then
            html = Mid(html, inicio, Fin - inicio)
            ' Escapar para JS
            html = Replace(html, "\", "\\")
            html = Replace(html, "'", "\'")
            html = Replace(html, vbCrLf, " ")
            html = Replace(html, vbLf, " ")
            
            Me.webInfo.ExecuteJavascript "if(window.actualizarTimeline) { window.actualizarTimeline('" & html & "'); }"
        End If
    End If
    
    ' Loguear falla silenciosa si hubo error en generacion HTML
    If errNum <> 0 Then
        Dim logSvcTimeline As LogServicio
        Set logSvcTimeline = New LogServicio
        logSvcTimeline.RegistrarError errNum, errDesc, "frmGestionSolicitud", "RefrescarTimelineSilencioso"
    End If
End Sub

Private Sub NotificarCambioContadorTareas()
    On Error Resume Next
    
    If FormularioAbierto("frm0Ppal") Then
        Forms("frm0Ppal").ActualizarContadorTareas
    ElseIf FormularioAbierto("frm0PpalTecnico") Then
        Forms("frm0PpalTecnico").ActualizarContadorTareas
    End If
    
    If Err.Number <> 0 Then
        Dim logSvc As LogServicio
        Set logSvc = New LogServicio
        logSvc.RegistrarError Err.Number, Err.description, "frmGestionSolicitud", "NotificarCambioContadorTareas"
        Err.Clear
    End If
End Sub



