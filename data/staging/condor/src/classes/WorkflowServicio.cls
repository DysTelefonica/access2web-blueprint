VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkflowServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' Spec-118: Cache de HTML del Timeline
Private m_cacheTimeline As Object
Private m_cacheTTL_Minutos As Integer

Private Sub InicializarCacheTimeline()
    If m_cacheTimeline Is Nothing Then
        Set m_cacheTimeline = CreateObject("Scripting.Dictionary")
        m_cacheTTL_Minutos = 5
    End If
End Sub

Public Sub InvalidarCacheTimeline(ByVal idSolicitud As Long)
    If m_cacheTimeline Is Nothing Then Exit Sub
    Dim cacheKey As String
    cacheKey = "TL_" & idSolicitud
    If m_cacheTimeline.Exists(cacheKey) Then
        m_cacheTimeline.Remove cacheKey
    End If
End Sub

Public Sub LimpiarCacheTimeline()
    If Not m_cacheTimeline Is Nothing Then
        m_cacheTimeline.RemoveAll
    End If
End Sub

Private Function GetFromCacheTimeline(ByVal idSolicitud As Long) As String
    If m_cacheTimeline Is Nothing Then
        GetFromCacheTimeline = ""
        Exit Function
    End If
    
    Dim cacheKey As String
    cacheKey = "TL_" & idSolicitud
    
    If m_cacheTimeline.Exists(cacheKey) Then
        Dim fechaCache As Date
        fechaCache = m_cacheTimeline(cacheKey)
        
        If DateDiff("n", fechaCache, Now()) < m_cacheTTL_Minutos Then
            GetFromCacheTimeline = m_cacheTimeline("HTML_" & cacheKey)
        Else
            m_cacheTimeline.Remove cacheKey
            m_cacheTimeline.Remove "HTML_" & cacheKey
            GetFromCacheTimeline = ""
        End If
    Else
        GetFromCacheTimeline = ""
    End If
End Function

Private Sub SaveToCacheTimeline(ByVal idSolicitud As Long, ByVal html As String)
    Call InicializarCacheTimeline
    Dim cacheKey As String
    cacheKey = "TL_" & idSolicitud
    m_cacheTimeline(cacheKey) = Now()
    m_cacheTimeline("HTML_" & cacheKey) = html
End Sub

Private Sub GenerarYAdjuntarDocumentoBorrador(ByVal idSolicitud As Long, ByVal etapa As String, ByRef dbTransaccional As DAO.Database)
    Dim docServ As New DocumentoServicio
    Dim adjServ As New AdjuntosServicio
    Dim rutaDocumento As String
    On Error GoTo Errores
    rutaDocumento = docServ.GenerarDocumentoParaSolicitud(idSolicitud, 0, dbTransaccional)
    If fso.FileExists(rutaDocumento) Then
        Call adjServ.GuardarAdjuntoDesdeArchivo(idSolicitud, etapa, rutaDocumento, "Generado", dbTransaccional)
    Else
        Err.Raise 513, "GenerarYAdjuntar", "Falló la creación del archivo de documento Word."
    End If
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.GenerarYAdjuntarDocumentoBorrador"
    errObj.Raise
End Sub
' ==========================================================================
' MÓDULO: WorkflowServicio.cls
' ACCIÓN: CORRECCIÓN DE ERROR 91 (Objeto no establecido) en Visualizador.
'         REFACTORIZACIÓN (Spec-007): Inyección de botones y lógica de visibilidad.
' ==========================================================================
Public Function GenerarHTML_VisualizadorDeEstado(ByVal idSolicitud As Long) As String
    Dim solServ As New SolicitudServicio
    Dim sol As Solicitud
    Dim html As String
    Dim estadoActual As estado
    Dim ordenActual As Long
    
    On Error GoTo Errores
    
    Dim htmlCacheado As String
    htmlCacheado = GetFromCacheTimeline(idSolicitud)
    If htmlCacheado <> "" Then
        GenerarHTML_VisualizadorDeEstado = htmlCacheado
        Exit Function
    End If
    
    Set sol = solServ.getSolicitudPorID(idSolicitud)
    If sol Is Nothing Then GoTo Finalizar
    
    Set estadoActual = m_ObjEntorno.estados(CStr(sol.idEstadoInterno))
    ordenActual = estadoActual.orden
    
    html = "<!DOCTYPE html><html lang='es'><head><meta charset='UTF-8'><style>" & GetTimelineCSS() & "</style></head><body>"
    html = html & "<div class='header'></div>"
    html = html & "<div id='timeline-container' class='timeline'>"
    
    Dim estadosOrdenados As Collection
    Set estadosOrdenados = SortDictionaryByProperty(m_ObjEntorno.estados, "orden")
    
    Dim estadoActualDelBucle As estado
    
    For Each estadoActualDelBucle In estadosOrdenados
        
        ' Lógica de visibilidad (Rechazada/Aprobada)
        If estadoActualDelBucle.idEstado = enumEstados.estadoRechazada Then
            If sol.idEstadoInterno <> enumEstados.estadoRechazada Then GoTo SiguienteEstado
        End If
        If sol.idEstadoInterno = enumEstados.estadoRechazada And estadoActualDelBucle.idEstado = enumEstados.estadoAprobada Then GoTo SiguienteEstado
        
        Dim claseCss As String, icono As String, contenidoExtra As String, anclaId As String
        Dim descripcionMostrar As String
        
        contenidoExtra = ""
        anclaId = ""
        icono = ""
        descripcionMostrar = estadoActualDelBucle.descripcion
        
        If estadoActualDelBucle.orden < ordenActual Then
            ' PASADO
            claseCss = "step-done": icono = "&#10003;"
            
        ElseIf estadoActualDelBucle.orden = ordenActual Then
            ' PRESENTE (ACTIVO)
            anclaId = " id='estado-activo'"
            
            ' Lógica específica por estado
            If estadoActualDelBucle.idEstado = estadoPreregistro Then
                 claseCss = "step-active": icono = "&#9658;": contenidoExtra = GenerarChecklistHTML(sol)
            ElseIf estadoActualDelBucle.idEstado = estadoDesarrolloTecnico Then
                 claseCss = "step-active"
                 icono = "&#9658;"
                 
                 ' Spec-007: Lógica de Rechazo Activo para Técnico
                 If sol.revisionCalidadEstado = "RECHAZADO" Then
                     Dim hayCambios As Boolean
                     hayCambios = VerificarCambiosTrasRechazo(sol)
                     
                     If hayCambios Then
                         ' Spec-011: Mensajes diferenciados por rol cuando hay cambios pero no se ha avanzado
                         If rolUsuario = rol.Calidad Or rolUsuario = rol.Administrador Then
                             contenidoExtra = "<div class='requirements-ok'>" & _
                                              "<strong>Cambios detectados:</strong><br>" & _
                                              "El Técnico ha realizado cambios pero aún no ha promocionado el estado a Modificación." & _
                                              "</div>"
                         Else
                             contenidoExtra = "<div class='requirements-ok'>" & _
                                              "<strong>Cambios detectados:</strong><br>" & _
                                              "Cambios realizados a falta de avanzar la fase a Modificación." & _
                                              "</div>"
                         End If
                     Else
                         ' Spec-010: Diferenciacion de mensajes por rol
                         If rolUsuario = rol.Tecnico Then
                             contenidoExtra = "<div class='rework-comments'>" & _
                                              "<strong>Esperando correcciones:</strong><br>" & _
                                               "Debe realizar cambios en las pestañas técnicas antes de poder avanzar." & _
                                              "</div>"
                         Else
                             contenidoExtra = "<div class='rework-comments'>" & _
                                              "<strong>Estado:</strong><br>" & _
                                               "Esperando cambios del técnico" & _
                                              "</div>"
                         End If
                     End If
                     
                     ' Spec-007: Botón siempre visible para ver los motivos (Spec-007: B.1)
                    ' Spec-012: Indicar si es solo lectura para Calidad si hay cambios
                    Dim textoBotonMotivos As String
                    textoBotonMotivos = "Ver Motivos Rechazo"
                    
                    If (rolUsuario = rol.Calidad Or rolUsuario = rol.Administrador) And hayCambios Then
                        textoBotonMotivos = "Ver Motivos Rechazo (Solo Lectura)"
                    End If
                    
                    contenidoExtra = contenidoExtra & "<br><button class='button' onclick=""window.colaComandos = 'VER_MOTIVOS_RECHAZO:0';"">" & textoBotonMotivos & "</button>"
                    
                    ' Spec-009a1: Botón para ver cambios detectados
                    If hayCambios Then
                        contenidoExtra = contenidoExtra & " <button class='button' onclick=""window.colaComandos = 'VER_CAMBIOS:0';"">Ver cambios</button>"
                    End If
                 Else
                     contenidoExtra = GenerarChecklistHTML(sol)
                 End If
                 
            ElseIf estadoActualDelBucle.idEstado = estadoModificacion Then
                 ' --- LÓGICA DE REVISIÓN EN MODIFICACIÓN (Spec-003 / Spec-007) ---
                 Dim esSubsanacion As Boolean
                 ' Si viene de un rechazo técnico, el estado de revisión persiste como RECHAZADO hasta que Calidad apruebe de nuevo.
                 esSubsanacion = (sol.revisionCalidadEstado = "RECHAZADO")
                 
                 If esSubsanacion Then
                      ' ESTADO: REVISIÓN DE SUBSANACIÓN
                      claseCss = "step-active"
                      icono = "&#9998;" ' Lápiz
                      
                      ' Spec-007: Descripción dinámica para subsanación (Gap 6.2)
                      descripcionMostrar = "El técnico ha hecho cambios a la espera de decisión de calidad"
                      
                      contenidoExtra = "<div class='requirements' style='border-left-color: #0066FF; background-color: #F0F8FF; color: #0049B8;'>" & _
                                       "<strong>Revisión de Subsanación:</strong><br>" & _
                                       "Esta solicitud ha sido corregida tras un rechazo.<br><br>"
                      
                      ' Spec-007: Botón Delta SOLO si hay cambios ACTIVOS tras rechazo (Spec-007: A.5, A.6)
                      If VerificarCambiosTrasRechazo(sol) Then
                          contenidoExtra = contenidoExtra & "<button class='button' onclick=""window.colaComandos = 'VER_DELTA:0';"">Ver Cambios (Delta)</button>"
                      End If
                      
                      contenidoExtra = contenidoExtra & "</div>"
                      contenidoExtra = contenidoExtra & GenerarChecklistHTML(sol)
                 Else
                      ' ESTADO: REVISIÓN NORMAL
                      claseCss = "step-active"
                      icono = "&#9658;"
                      contenidoExtra = GenerarChecklistHTML(sol)
                 End If
                
            ElseIf estadoActualDelBucle.idEstado = estadoRevision Then
                claseCss = "step-active"
                icono = "&#9658;"
                contenidoExtra = ""
                
                Dim ultimoOrdinalRevision As Long
                contenidoExtra = contenidoExtra & "<div>"
                contenidoExtra = contenidoExtra & "<button class='action-button' onclick=""window.colaComandos = 'APROBAR_RAC:0';"">Aprobado sin comentarios</button>"
                contenidoExtra = contenidoExtra & "<button class='action-button-back' onclick=""window.colaComandos = 'COMENTARIOS_RAC:0';"">Revisado con comentarios</button>"
                contenidoExtra = contenidoExtra & "<button class='action-button-back' onclick=""window.colaComandos = 'IMPORTAR_WORD_RAC_DESDE_BOTON:0';"">Importar Word con comentarios RAC</button>"
                
                ultimoOrdinalRevision = ValidacionRevisionRepositorio.GetUltimoOrdinal(sol.idSolicitud)
                If ultimoOrdinalRevision > 0 Then
                    contenidoExtra = contenidoExtra & "<div style='margin-top:6px;'><a href='javascript:void(0);' onclick=""window.colaComandos = 'VER_HISTORIAL_RAC:0';"">Ver historial de validaci&oacute;n y revisi&oacute;n</a></div>"
                End If
                
                contenidoExtra = contenidoExtra & "</div>"
                contenidoExtra = contenidoExtra & GenerarChecklistHTML(sol)
                contenidoExtra = contenidoExtra & "<div class='danger-zone'><button class='action-button-reset' onclick=""window.colaComandos = 'REVERTIR_VALIDACION:5';"">RESET FASE</button></div>"
                
            ElseIf estadoActualDelBucle.idEstado = enumEstados.estadoRechazada Then
                 claseCss = "step-rework": icono = "&#10007;": contenidoExtra = "<div class='rework-comments'>SOLICITUD RECHAZADA.</div>"
                 
            ElseIf estadoActualDelBucle.idEstado = enumEstados.estadoAprobada Then
                 claseCss = "step-active": icono = "&#10003;": contenidoExtra = "<div class='info-box'>Solicitud cerrada.</div>"
                 
            ElseIf estadoActualDelBucle.idEstado = enumEstados.estadoValidacion Then
                 claseCss = "step-active": icono = "&#9658;": contenidoExtra = GenerarChecklistHTML(sol)
                 
                 If rolUsuario = rol.Calidad Or rolUsuario = rol.Administrador Then
                     Dim versionBorrador As String
                     Dim ultimoOrdinalValidacion As Long
                     
                    versionBorrador = GetProximaVersionBorrador(sol.idSolicitud)
                    contenidoExtra = contenidoExtra & "<div>"
                    contenidoExtra = contenidoExtra & "<button class='action-button' onclick=""window.colaComandos = 'EXPORTAR_BORRADOR:0';"">Exportar borrador (" & versionBorrador & ")</button>"
                    contenidoExtra = contenidoExtra & "<button class='action-button' onclick=""window.colaComandos = 'NOTIFICAR_ENVIO_RAC:0';"">Enviado a RAC</button>"
                     
                     ultimoOrdinalValidacion = ValidacionRevisionRepositorio.GetUltimoOrdinal(sol.idSolicitud)
                     If ultimoOrdinalValidacion > 0 Then
                         contenidoExtra = contenidoExtra & "<div style='margin-top:6px;'><a href='javascript:void(0);' onclick=""window.colaComandos = 'VER_HISTORIAL_RAC:0';"">Ver historial de validaci&oacute;n y revisi&oacute;n</a></div>"
                     End If
                     
                     contenidoExtra = contenidoExtra & "</div>"
                     
                     contenidoExtra = contenidoExtra & "<div class='danger-zone'><button class='action-button-reset' onclick=""window.colaComandos = 'REVERTIR_TECNICA:" & estadoModificacion & "';"">RESET FASE</button></div>"
                 End If
                 
                 ' (Spec-002/003): En Validación NO mostramos Delta ni opciones de rechazo técnico.
                 ' Se asume que si llegó aquí, la fase técnica está cerrada.

            ElseIf estadoActualDelBucle.idEstado = estadoFormalizacion Then
                 claseCss = "step-active": icono = "&#9658;": contenidoExtra = GenerarChecklistHTML(sol)
                 
                 contenidoExtra = contenidoExtra & "<div>"
                 contenidoExtra = contenidoExtra & "<button class='action-button' onclick=""window.colaComandos = 'GENERAR_PDF_FINAL:0';"">Generar PDF firma</button>"
                 contenidoExtra = contenidoExtra & "<button class='action-button' onclick=""window.colaComandos = 'ADJUNTAR_PDF_FIRMADO:0';"">Adjuntar PDF firmado (Aprobar)</button>"
                 contenidoExtra = contenidoExtra & "</div>"
                 
            Else
                 claseCss = "step-active": icono = "&#9658;": contenidoExtra = GenerarChecklistHTML(sol)
            End If
            
            If estadoActualDelBucle.idEstado <> estadoValidacion And estadoActualDelBucle.idEstado <> estadoRevision Then
                Dim transicionesValidas As Scripting.Dictionary
                Set transicionesValidas = getTransicionesValidas(idSolicitud, rolUsuario)
                Dim botonesHtml As String: botonesHtml = ""
                Dim botonesResetHtml As String: botonesResetHtml = ""
                
                If estadoActualDelBucle.idEstado = estadoModificacion And (rolUsuario = rol.Calidad Or rolUsuario = rol.Administrador) Then
                    botonesHtml = botonesHtml & "<button class='action-button' onclick=""window.colaComandos = 'AVANZAR_ESTADO:5';"">Aceptar y Avanzar</button>"
                    botonesHtml = botonesHtml & "<button class='action-button-back' onclick=""window.colaComandos = 'RECHAZAR_TECNICA:3';"">Rechazar (Devolver a Técnico)</button>"
                    botonesResetHtml = botonesResetHtml & "<button class='action-button-reset' onclick=""window.colaComandos = 'REVERTIR_DESARROLLO:" & estadoDesarrolloTecnico & "';"">RESET FASE</button>"
                End If
    
                If Not transicionesValidas Is Nothing Then
                    Dim key As Variant, estDestino As estado
                    Dim comandoVBA As String, textoBotón As String, claseBotón As String
                    
                    For Each key In transicionesValidas.Keys
                        Set estDestino = transicionesValidas(key)
                        
                        If estDestino.orden > ordenActual Then
                            comandoVBA = "AVANZAR_ESTADO"
                            claseBotón = "action-button"
                            textoBotón = "Avanzar a '" & GetNombreEstadoSafe(estDestino.idEstado) & "'"
                            
                            If estadoActualDelBucle.idEstado = estadoModificacion And estDestino.idEstado = estadoValidacion And (rolUsuario = rol.Calidad Or rolUsuario = rol.Administrador) Then
                                textoBotón = ""
                            End If
                            
                            If estadoActualDelBucle.idEstado = enumEstados.estadoRevision Then
                                If estDestino.idEstado = enumEstados.estadoRechazada Then claseBotón = "action-button-back": textoBotón = "Cerrar como '" & GetNombreEstadoSafe(estDestino.idEstado) & "'"
                            End If
    
                            If estadoActualDelBucle.idEstado = estadoFormalizacion Then textoBotón = ""
                            
                            If estadoActualDelBucle.idEstado = estadoDesarrolloTecnico And rolUsuario = rol.Tecnico And sol.revisionCalidadEstado = "RECHAZADO" Then
                                If Not VerificarCambiosTrasRechazo(sol) Then
                                    comandoVBA = "ERROR_SIN_CAMBIOS"
                                End If
                            End If
    
                            If textoBotón <> "" Then
                                botonesHtml = botonesHtml & "<button class='" & claseBotón & "' onclick=""window.colaComandos = '" & comandoVBA & ":" & estDestino.idEstado & "';"">" & textoBotón & "</button>"
                            End If
                            
                        ElseIf estDestino.orden < ordenActual Then
                            comandoVBA = "AVANZAR_ESTADO"
                            claseBotón = "action-button-reset"
                            textoBotón = "RESET FASE"
                            
                            If estadoActualDelBucle.idEstado = estadoValidacion And estDestino.idEstado = estadoModificacion Then
                                 comandoVBA = "REVERTIR_TECNICA"
                            End If
                            
                            If estadoActualDelBucle.idEstado = estadoModificacion And estDestino.idEstado = estadoDesarrolloTecnico Then
                                 If (rolUsuario = rol.Calidad Or rolUsuario = rol.Administrador) Then
                                     textoBotón = ""
                                 Else
                                     comandoVBA = "RECHAZAR_TECNICA"
                                     textoBotón = "Rechazar (Devolver a Técnico)"
                                 End If
                            End If
                            
                            If estadoActualDelBucle.idEstado = enumEstados.estadoRechazada And estDestino.idEstado = enumEstados.estadoRevision Then
                                 comandoVBA = "REVERTIR_RECHAZO"
                                 textoBotón = "REABRIR SOLICITUD"
                            End If
                            
                            If textoBotón <> "" Then
                                 botonesResetHtml = botonesResetHtml & "<button class='" & claseBotón & "' onclick=""window.colaComandos = '" & comandoVBA & ":" & estDestino.idEstado & "';"">" & textoBotón & "</button>"
                            End If
                        End If
                    Next key
                End If
                
                ' BOTÓN RESET FASE EN DESARROLLO TÉCNICO (Hacia Registro)
                If estadoActualDelBucle.idEstado = estadoDesarrolloTecnico Then
                    If rolUsuario = rol.Calidad Or rolUsuario = rol.Administrador Then
                        botonesResetHtml = botonesResetHtml & "<button class='action-button-reset' onclick=""window.colaComandos = 'REVERTIR_DESARROLLO:2';"">RESET FASE</button>"
                    End If
                End If
                
                If estadoActualDelBucle.idEstado = estadoFormalizacion Then
                    If rolUsuario = rol.Calidad Or rolUsuario = rol.Administrador Then
                        botonesResetHtml = botonesResetHtml & "<button class='action-button-reset' onclick=""window.colaComandos = 'REVERTIR_FORMALIZACION:6';"">RESET FASE</button>"
                    End If
                End If
                
                If botonesHtml <> "" Then contenidoExtra = contenidoExtra & "<div>" & botonesHtml & "</div>"
                If botonesResetHtml <> "" Then contenidoExtra = contenidoExtra & "<div class='danger-zone'>" & botonesResetHtml & "</div>"
            End If
            
        Else
            ' FUTURO
            claseCss = "step-future": icono = "&#9711;"
        End If
        
        Dim tituloHtml As String
        tituloHtml = "<div class='title'>" & UCase(GetNombreEstadoSafe(estadoActualDelBucle.idEstado))
        
        If claseCss = "step-done" Then
            Dim histId As String
            histId = "hist-" & CStr(estadoActualDelBucle.idEstado)
            tituloHtml = tituloHtml & " <span class='history-toggle' data-target='" & histId & "' title='Ver detalle'>&#128269;</span>"
            
            Dim tiemposDict As Scripting.Dictionary
            Dim cronometroTxt As String
            Set tiemposDict = GetTiemposPorEstado(idSolicitud)
            If Not tiemposDict Is Nothing Then
                If tiemposDict.Exists(CStr(estadoActualDelBucle.idEstado)) Then
                    cronometroTxt = FormatearDuracion(CLng(tiemposDict(CStr(estadoActualDelBucle.idEstado))))
                Else
                    cronometroTxt = "0m"
                End If
            Else
                cronometroTxt = "0m"
            End If
            
            Dim colHistorial As Scripting.Dictionary
            Dim usuarioInicio As String, usuarioCierre As String
            Dim fechaInicio As Date, fechaCierre As Date
            Set colHistorial = solServ.getHistorialDeEstados(idSolicitud)
            usuarioInicio = "": usuarioCierre = ""
            If Not colHistorial Is Nothing Then
                Dim k As Variant, log As LogEstado
                For Each k In colHistorial.Keys
                    Set log = colHistorial(k)
                    If log.idEstadoNuevo = estadoActualDelBucle.idEstado Then
                        usuarioInicio = log.usuarioTransicion
                        fechaInicio = log.fechaTransicion
                    End If
                    If log.idEstadoAnterior = estadoActualDelBucle.idEstado Then
                        usuarioCierre = log.usuarioTransicion
                        fechaCierre = log.fechaTransicion
                    End If
                Next k
            End If
            
            ' --- MODIFICACIÓN FORENSE (Spec-032.2): Resumen Limpio ---
            Dim inicioStr As String, cierreStr As String
            inicioStr = IIf(fechaInicio <> 0, Format(fechaInicio, "dd/mm/yyyy hh:nn"), "-")
            cierreStr = IIf(fechaCierre <> 0, Format(fechaCierre, "dd/mm/yyyy hh:nn"), "-")
            
            ' Cálculo de Ciclos
            Dim ciclosStr As String: ciclosStr = ""
            If estadoActualDelBucle.idEstado = estadoDesarrolloTecnico Or estadoActualDelBucle.idEstado = estadoModificacion Then
                Dim rsCiclos As DAO.Recordset
                Set rsCiclos = RechazoRepositorio.GetHistorialPorSolicitud(idSolicitud)
                If Not rsCiclos Is Nothing Then
                    If Not rsCiclos.EOF Then
                        rsCiclos.MoveLast
                        ciclosStr = CStr(rsCiclos.RecordCount) & " iteraciones"
                    End If
                    rsCiclos.Close
                End If
            ElseIf estadoActualDelBucle.idEstado = estadoValidacion Or estadoActualDelBucle.idEstado = estadoRevision Then
                Dim ordRAC As Long
                ordRAC = ValidacionRevisionRepositorio.GetUltimoOrdinal(idSolicitud)
                If ordRAC > 0 Then ciclosStr = "v" & CStr(ordRAC)
            End If
            
            Dim detalleHtml As String
            detalleHtml = "<div class='history-row'><span class='history-label'>Inicio:</span> <span class='history-value'>" & inicioStr & "</span></div>" & _
                          "<div class='history-row'><span class='history-label'>Fin:</span> <span class='history-value'>" & cierreStr & "</span></div>" & _
                          "<div class='history-row'><span class='history-label'>Total:</span> <span class='history-value'>" & cronometroTxt & "</span></div>" & _
                          "<div class='history-row'><span class='history-label'>Actor cierre:</span> <span class='history-value'>" & HtmlEncode(CStr(Nz(usuarioCierre, ""))) & "</span></div>"
            
            If ciclosStr <> "" Then
                detalleHtml = detalleHtml & "<div class='history-row'><span class='history-label'>Ciclos:</span> <span class='history-value'>" & ciclosStr & "</span></div>"
            End If
            
            ' Adjuntos
            Dim idAdj As Long
            idAdj = GetAdjuntoIdPorEtapa(idSolicitud, GetNombreEstadoSafe(estadoActualDelBucle.idEstado))
            If idAdj > 0 Then
                detalleHtml = detalleHtml & "<div class='history-row history-links'><button type='button' class='link-button' onclick=""window.colaComandos = 'ABRIR_ADJUNTO:" & CStr(idAdj) & "';"">Abrir Word</button></div>"
            End If
            
            ' Botones y Listas Específicas
            If estadoActualDelBucle.idEstado = estadoDesarrolloTecnico Then
                ' Spec-032.3: Botón Ver Detalle recuperando info forense
                detalleHtml = detalleHtml & "<div style='margin-top:8px;'><button class='button' onclick=""window.colaComandos = 'VER_DELTA:0';"">Ver Detalle (Forense)</button></div>"
            ElseIf estadoActualDelBucle.idEstado = estadoModificacion Then
                detalleHtml = detalleHtml & ConstruirListadoRechazos(idSolicitud)
            End If
            
            If estadoActualDelBucle.idEstado = estadoValidacion Or estadoActualDelBucle.idEstado = estadoRevision Then
                detalleHtml = detalleHtml & ConstruirListadoRAC(idSolicitud)
            End If
            
            contenidoExtra = contenidoExtra & "<div id='" & histId & "' class='history-container' style='display:none;'>" & detalleHtml & "</div>"
        End If
        
        tituloHtml = tituloHtml & "</div>"
        
        html = html & "<div class='step " & claseCss & "'" & anclaId & "><div class='icon'>" & icono & "</div><div class='content'>" & tituloHtml & "<div class='description'>" & descripcionMostrar & "</div>" & contenidoExtra & "</div></div>"

SiguienteEstado:
    Next estadoActualDelBucle

    html = html & "</div><script>" & vbCrLf & _
           "function bindInteractions() {" & vbCrLf & _
           "    var t=document.getElementsByClassName('history-toggle');" & vbCrLf & _
           "    for(var i=0;i<t.length;i++){" & vbCrLf & _
           "        t[i].onclick = function(){" & vbCrLf & _
           "            var id=this.getAttribute('data-target');" & vbCrLf & _
           "            var box=document.getElementById(id);" & vbCrLf & _
           "            if(!box) return;" & vbCrLf & _
           "            box.style.display = (box.style.display==='none' || !box.style.display) ? 'block' : 'none';" & vbCrLf & _
           "        };" & vbCrLf & _
           "    }" & vbCrLf & _
           "}" & vbCrLf & _
           "window.actualizarTimeline = function(nuevoHtml) {" & vbCrLf & _
           "    document.getElementById('timeline-container').innerHTML = nuevoHtml;" & vbCrLf & _
           "    bindInteractions();" & vbCrLf & _
           "};" & vbCrLf & _
           "document.addEventListener('DOMContentLoaded', function() {" & vbCrLf & _
           "    var el = document.getElementById('estado-activo');" & vbCrLf & _
           "    if(el) el.scrollIntoView({behavior: 'auto', block: 'center'});" & vbCrLf & _
           "    bindInteractions();" & vbCrLf & _
           "});" & vbCrLf & _
           "</script></body></html>"

Finalizar:
    Call SaveToCacheTimeline(idSolicitud, html)
    GenerarHTML_VisualizadorDeEstado = html
    Exit Function
Errores:
    Dim logServ As New LogServicio
    logServ.RegistrarError Err.Number, "Error en línea " & Erl() & " — " & Err.description, "WorkflowServicio", "GenerarHTML_VisualizadorDeEstado", ""
    html = "<html><body>Error al generar visualizador: " & Err.description & " (" & Err.Number & ")</body></html>"
    Resume Finalizar
End Function
' ==========================================================================
' MÓDULO: WorkflowServicio.cls
' FUNCIÓN: GetTimelineCSS (ESTILOS COMPLETOS PARA VISUALIZADOR)
' ==========================================================================
Private Function GetTimelineCSS() As String
    Dim css As String
    
    ' --- 1. Estilos Generales ---
    css = "body { font-family: 'Telefónica Sans', Arial, sans-serif; margin: 0; padding: 0; background-color: #FFFFFF; font-size: 12px; color: #2B3447; } "
    css = css & ".header { display: none; height: 0; margin: 0; padding: 0; } "
    
    ' --- 2. Estilos de Botónes ---
    ' Botón genérico pequeño
    css = css & ".button { background-color: #FFFFFF; color: #0066FF; border: 1px solid #0066FF; border-radius: 4px; padding: 5px 10px; font-size: 11px; font-weight: bold; cursor: pointer; } "
    css = css & ".button:hover { background-color: #F2F4FF; } "
    
    ' Botón de Acción Principal (AVANZAR - Azul)
    css = css & ".action-button { background-color: #0066FF; color: white; border: 1px solid #0066FF; border-radius: 4px; padding: 8px 15px; font-size: 12px; font-weight: bold; cursor: pointer; display: inline-block; margin-top: 10px; margin-right: 5px; text-decoration: none; } "
    css = css & ".action-button:hover { background-color: #0049B8; border-color: #0049B8; } "
    
    ' Botón de Acción Secundaria/Negativa (RETROCEDER - Coral/Rojo)
    css = css & ".action-button-back { background-color: #FFFFFF; color: #D6786B; border: 1px solid #D6786B; border-radius: 4px; padding: 8px 15px; font-size: 12px; font-weight: bold; cursor: pointer; display: inline-block; margin-top: 10px; margin-right: 5px; text-decoration: none; } "
    css = css & ".action-button-back:hover { background-color: #F9EEED; } "
    
    ' Botón de RESET (Acción destructiva de retroceso)
    css = css & ".action-button-reset { background-color: #FFF5F5; color: #C53030; border: 1.5px dashed #C53030; border-radius: 4px; padding: 6px 12px; font-size: 11px; font-weight: bold; cursor: pointer; display: inline-block; margin-top: 15px; text-transform: uppercase; } "
    css = css & ".action-button-reset:hover { background-color: #C53030; color: white; border-style: solid; } "
    
    ' Contenedor de Zona de Peligro (Separador visual para acciones destructivas)
    css = css & ".danger-zone { border-top: 1px solid #EEE; margin-top: 10px; padding-top: 5px; } "
    
    ' --- 3. Estructura del Timeline ---
    css = css & ".timeline { display: flex; flex-direction: column; position: relative; } "
    css = css & ".timeline::before { content: ''; position: absolute; left: 14px; top: 0; bottom: 0; width: 2px; background: #D1D5E4; } "
    css = css & ".step { display: flex; align-items: flex-start; margin-bottom: 6px; position: relative; padding: 6px 0 6px 28px; } "
    css = css & ".icon { position: absolute; left: 0; top: 0; width: 22px; height: 22px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 12px; font-weight: bold; z-index: 1; } "
    css = css & ".content { width: 100%; } "
    css = css & ".title { font-weight: bold; font-size: 14px; margin-bottom: 4px; } "
    css = css & ".description { color: #58617A; margin-bottom: 4px; font-size: 10.5px; } "
    
    ' --- 4. Estados de los Pasos ---
    ' Pasado (Hecho)
    css = css & ".step-done .icon { background-color: #0066FF; color: white; border: 2px solid #F2F4FF; } "
    css = css & ".step-done .title { color: #0066FF; } "
    
    ' Actual (Activo)
    css = css & ".step-active .icon { background-color: #00A551; color: white; border: 2px solid #F2F4FF; } "
    css = css & ".step-active .title { color: #00A551; } "
    
    ' Futuro
    css = css & ".step-future .icon { background-color: #D1D5E4; color: #6E7894; border: 2px solid #F2F4FF; } "
    css = css & ".step-future .title, .step-future .description { color: #8F97AF; } "
    
    ' --- 5. Variaciones y Animaciones ---
    ' Siguiente acción sugerida (Pulso Amarillo)
    css = css & ".step-next .icon { background-color: #E4C35C; color: #473515; animation: pulse 1.5s infinite; } "
    css = css & "@keyframes pulse { 0% { box-shadow: 0 0 0 0 rgba(228, 195, 92, 0.7); } 70% { box-shadow: 0 0 0 10px rgba(228, 195, 92, 0); } 100% { box-shadow: 0 0 0 0 rgba(228, 195, 92, 0); } } "
    
    ' Estado de Rework/Rechazo (Rebote Rojo)
    css = css & ".step-rework .icon { background-color: #D6786B; color: white; animation: bounce 1s; } "
    css = css & ".step-rework .title { color: #D6786B; } "
    css = css & "@keyframes bounce { 0%, 20%, 50%, 80%, 100% {transform: translateY(0);} 40% {transform: translateY(-10px);} 60% {transform: translateY(-5px);} } "
    
    ' --- 6. Cajas de Mensajes (Alertas) ---
    ' Requisitos Pendientes (Amarillo)
    css = css & ".requirements { margin-top: 8px; padding: 10px; background-color: #FFF8F3; border-left: 4px solid #E4C35C; color: #A7863E; font-size: 12px; border-radius: 0 4px 4px 0; } "
    css = css & ".requirements ul { margin: 5px 0 0 15px; padding: 0; } "
    
    ' Requisitos Cumplidos / OK (Cyan/Verde)
    css = css & ".requirements-ok { margin-top: 8px; padding: 10px; background-color: #EBFFFF; border-left: 4px solid #75C0C7; color: #528889; font-size: 12px; border-radius: 0 4px 4px 0; }"
    
    ' Comentarios de Rechazo (Coral)
    css = css & ".rework-comments { margin-top: 10px; padding: 10px; background-color: #F9EEED; border-left: 4px solid #D6786B; color: #843C34; font-size: 12px; border-radius: 0 4px 4px 0; } "
    css = css & ".rework-comments strong { display: block; margin-bottom: 5px; }"
    
    ' Caja de Información Neutra (Gris/Azul) - Usada en estado APROBADA
    css = css & ".info-box { margin-top: 8px; padding: 10px; background-color: #F2F4FF; border-left: 4px solid #0066FF; color: #58617A; font-size: 12px; border-radius: 0 4px 4px 0; }"
     ' --- 7. Estilos para Checklist de Requisitos ---
    css = css & ".checklist { list-style: none; padding: 0; margin: 10px 0; }"
    css = css & ".checklist li { display: flex; align-items: center; margin-bottom: 6px; font-size: 12px; }"
    css = css & ".check-icon { margin-right: 8px; font-weight: bold; font-size: 14px; }"
    css = css & ".req-ok { color: #00A551; } " ' Verde
    css = css & ".req-nok { color: #D6786B; } " ' Rojo Coral
    css = css & ".req-text { color: #58617A; }"
    
    css = css & ".history-toggle { display: inline-block; margin-left: 8px; cursor: pointer; font-size: 12px; color: #0066FF; } "
    css = css & ".history-toggle:hover { color: #0049B8; } "
    css = css & ".history-container { margin-top: 6px; padding: 8px; background: #F9FAFC; border: 1px solid #E6E9F2; border-radius: 6px; font-size: 11px; color: #58617A; } "
    css = css & ".history-row { display: flex; gap: 8px; align-items: center; margin-bottom: 6px; } "
    css = css & ".history-label { font-weight: bold; color: #8F97AF; } "
    css = css & ".history-value { color: #2B3447; } "
    css = css & ".history-links .link-button { margin-right: 8px; text-decoration: none; font-weight: bold; color: #0066FF; cursor: pointer; background: none; border: none; padding: 0; } "
    css = css & ".history-group { margin-top: 8px; } "
    css = css & ".history-subtitle { font-weight: bold; color: #0066FF; margin-bottom: 6px; } "
    css = css & ".history-item { border: 1px solid #E6E9F2; border-radius: 6px; padding: 6px; margin-bottom: 6px; } "
    GetTimelineCSS = css
End Function

' ==========================================================================
' MÓDULO: WorkflowServicio.cls
' ACCIÓN: REEMPLAZAR esta función para eliminar la dependencia de .NET (ArrayList).
'         Ahora usa un algoritmo de ordenación 100% nativo de VBA.
' ==========================================================================
Private Function SortDictionaryByProperty(ByVal dict As Scripting.Dictionary, ByVal propName As String) As Collection
    ' Responsabilidad: Ordena un diccionario de objetos en una nueva colección,
    '                  basándose en el valor numérico de una de sus propiedades.
    Dim arr() As Object
    Dim col As New Collection
    Dim key As Variant
    Dim i As Long, j As Long
    Dim temp As Object
    
    On Error GoTo Errores
    
    ' 1. Volcar los objetos del diccionario a un array para poder ordenarlos
    ReDim arr(1 To dict.count)
    i = 1
    For Each key In dict.Keys
        Set arr(i) = dict(key)
        i = i + 1
    Next key
    
    ' 2. Ordenar el array usando el algoritmo de burbuja (sencillo y eficaz para N pequeño)
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i).getPropiedad(propName) > arr(j).getPropiedad(propName) Then
                Set temp = arr(i)
                Set arr(i) = arr(j)
                Set arr(j) = temp
            End If
        Next j
    Next i
    
    ' 3. Volcar el array ordenado a la colección de salida
    For i = 1 To UBound(arr)
        col.Add arr(i)
    Next i
    
    Set SortDictionaryByProperty = col
    
LimpiarYSalir:
    Exit Function
Errores:
    ' En caso de error, devuelve una colección vacía para no romper el visualizador
    Set SortDictionaryByProperty = New Collection
    Resume LimpiarYSalir
End Function

Private Function ConstruirListadoRechazos(ByVal idSolicitud As Long) As String
    Dim html As String
    Dim rcd As DAO.Recordset
    On Error GoTo Errores
    
    Set rcd = RechazoRepositorio.GetHistorialPorSolicitud(idSolicitud)
    If rcd Is Nothing Then GoTo Finalizar
    If rcd.EOF Then GoTo Finalizar
    
    rcd.MoveFirst
    html = "<div class='history-group'><div class='history-subtitle'>Rechazos de Calidad</div>"
    Do While Not rcd.EOF
        Dim fechaStr As String
        Dim usuarioStr As String
        Dim motivoStr As String
        Dim comentariosStr As String
        fechaStr = IIf(IsDate(rcd!fechaRechazo), Format(rcd!fechaRechazo, "dd/mm/yyyy"), "-")
        usuarioStr = HtmlEncode(CStr(Nz(rcd!usuarioRechazo, "")))
        motivoStr = HtmlEncode(CStr(Nz(rcd!motivoPrincipal, "")))
        comentariosStr = HtmlEncode(CStr(Nz(rcd!comentarios, "")))
        comentariosStr = Replace(comentariosStr, vbCrLf, "<br>")
        comentariosStr = Replace(comentariosStr, vbLf, "<br>")
        
        html = html & "<div class='history-item'>" & _
                      "<div class='history-row'><span class='history-label'>Fecha:</span> <span class='history-value'>" & fechaStr & "</span></div>" & _
                      "<div class='history-row'><span class='history-label'>Autor:</span> <span class='history-value'>" & usuarioStr & "</span></div>" & _
                      "<div class='history-row'><span class='history-label'>Motivo:</span> <span class='history-value'>" & motivoStr & "</span></div>" & _
                      "<div class='history-row'><span class='history-label'>Comentarios:</span> <span class='history-value'>" & comentariosStr & "</span></div>" & _
                      "</div>"
        rcd.MoveNext
    Loop
    html = html & "</div>"
    
Finalizar:
    ConstruirListadoRechazos = html
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
Errores:
    ConstruirListadoRechazos = ""
    Resume Finalizar
End Function

Private Function ConstruirListadoRAC(ByVal idSolicitud As Long) As String
    Dim html As String
    Dim rcd As DAO.Recordset
    On Error GoTo Errores
    
    Set rcd = ValidacionRevisionRepositorio.GetHistorialPorSolicitud(idSolicitud)
    If rcd Is Nothing Then GoTo Finalizar
    If rcd.EOF Then GoTo Finalizar
    
    rcd.MoveFirst
    html = "<div class='history-group'><div class='history-subtitle'>Hitos RAC</div>"
    Do While Not rcd.EOF
        Dim fechaEnvioStr As String, fechaRecepcionStr As String
        Dim usuarioStr As String, resultadoValue As String, claseEstado As String
        Dim idAdjunto As Long, wordLink As String
        
        If IsDate(rcd!fechaEnvio) Then fechaEnvioStr = Format(rcd!fechaEnvio, "dd/mm/yyyy") Else fechaEnvioStr = "-"
        If IsDate(rcd!fechaRecepcion) Then fechaRecepcionStr = Format(rcd!fechaRecepcion, "dd/mm/yyyy") Else fechaRecepcionStr = "-"
        usuarioStr = HtmlEncode(CStr(Nz(rcd!usuario, "")))
        
        resultadoValue = UCase$(CStr(Nz(rcd!resultado, "PENDIENTE")))
        If resultadoValue = "" Then resultadoValue = "PENDIENTE"
        Select Case resultadoValue
            Case "APROBADO": claseEstado = "aprobado"
            Case "COMENTARIOS": claseEstado = "comentarios"
            Case Else: claseEstado = "pendiente": resultadoValue = "PENDIENTE"
        End Select
        
        idAdjunto = Nz(rcd!idAdjunto, 0)
        wordLink = ""
        If idAdjunto > 0 Then
            wordLink = "<button type='button' class='word-link' onclick=""window.colaComandos = 'ABRIR_ADJUNTO:" & CStr(idAdjunto) & "';"">[Word]</button>"
        End If
        
        Dim comentariosHtml As String
        comentariosHtml = HtmlEncode(CStr(Nz(rcd!comentarios, "")))
        If Trim(comentariosHtml) = "" Then comentariosHtml = "(Vacio)"
        comentariosHtml = Replace(comentariosHtml, vbCrLf, "<br>")
        comentariosHtml = Replace(comentariosHtml, vbLf, "<br>")
        
        html = html & "<div class='history-item " & claseEstado & "'>" & _
                      "<div class='history-row'><span class='history-label'>Envio:</span> <span class='history-value'>" & fechaEnvioStr & "</span></div>" & _
                      "<div class='history-row'><span class='history-label'>Recepcion:</span> <span class='history-value'>" & fechaRecepcionStr & "</span></div>" & _
                      "<div class='history-row'><span class='history-label'>Usuario:</span> <span class='history-value'>" & usuarioStr & "</span></div>" & _
                      "<div class='history-row'><span class='history-label'>Resultado:</span> <span class='history-value'>" & resultadoValue & "</span> " & wordLink & "</div>" & _
                      "<div class='history-row'><span class='history-label'>Obs:</span> <span class='history-value'>" & comentariosHtml & "</span></div>" & _
                      "</div>"
        rcd.MoveNext
    Loop
    html = html & "</div>"
    
Finalizar:
    ConstruirListadoRAC = html
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
Errores:
    ConstruirListadoRAC = ""
    Resume Finalizar
End Function

Private Function ConstruirResumenCambiosTecnico(ByVal idSolicitud As Long) As String
    Dim html As String
    Dim jsonDict As Object
    Dim bloqueKey As Variant, campoKey As Variant
    Dim bloques As Long, campos As Long
    On Error GoTo Errores
    
    Set jsonDict = JsonHelper.ObtenerJsonRechazo(idSolicitud)
    If jsonDict Is Nothing Then GoTo Finalizar
    
    For Each bloqueKey In jsonDict.Keys
        bloques = bloques + 1
        Dim bloque As Object
        Set bloque = jsonDict(bloqueKey)
        For Each campoKey In bloque.Keys
            campos = campos + 1
        Next campoKey
    Next bloqueKey
    
    html = "<div class='history-group'><div class='history-subtitle'>Cambios Tecnico</div>" & _
           "<div class='history-item'>" & _
           "<div class='history-row'><span class='history-label'>Bloques:</span> <span class='history-value'>" & CStr(bloques) & "</span></div>" & _
           "<div class='history-row'><span class='history-label'>Campos:</span> <span class='history-value'>" & CStr(campos) & "</span></div>" & _
           "</div></div>"
    
Finalizar:
    ConstruirResumenCambiosTecnico = html
    Exit Function
Errores:
    ConstruirResumenCambiosTecnico = ""
    Resume Finalizar
End Function
Public Function getTransicionesValidas(ByVal idSolicitud As Long, ByVal rolUsuarioActual As rol) As Scripting.Dictionary
    ' --- CAMBIO: Ahora recibe idSolicitud en lugar de idEstadoOrigen ---
    Dim colPosiblesDestinos As Scripting.Dictionary
    Dim colEstadosValidos As New Scripting.Dictionary
    Dim idEstadoDestino As Variant
    Dim rolRequerido As String
    
    Dim solServ As New SolicitudServicio
    Dim sol As Solicitud
    Dim mensajeValidacion As String
    
    On Error GoTo Errores
    
    ' 1. Obtener la solicitud para conocer su estado actual.
    Set sol = solServ.getSolicitudPorID(idSolicitud)
    If sol Is Nothing Then
        Set getTransicionesValidas = colEstadosValidos
        Exit Function
    End If
    
    ' 2. Obtener las transiciones teóricamente posibles desde el estado actual.
    Set colPosiblesDestinos = WorkflowRepositorio.getPosiblesDestinos(sol.idEstadoInterno)
    
    ' 3. Iterar y aplicar lógica de negocio para filtrar las que son realmente válidas.
    For Each idEstadoDestino In colPosiblesDestinos.Keys
        rolRequerido = colPosiblesDestinos(idEstadoDestino)
        
        ' Condición 1: El rol del usuario es suficiente.
        If PermisoSuficiente(rolUsuarioActual, rolRequerido) Then
            
            ' Condición 2 (NUEVA): Se cumplen las precondiciones de negocio para esa transición.
            mensajeValidacion = ""
            If PrecondicionesCumplidas(sol, CLng(idEstadoDestino), mensajeValidacion) Then
            
                ' Si todo es OK, añadimos el estado a la lista de destinos válidos.
                If m_ObjEntorno.estados.Exists(CStr(idEstadoDestino)) Then
                    colEstadosValidos.Add idEstadoDestino, m_ObjEntorno.estados(idEstadoDestino)
                End If
                
            End If
        End If
    Next idEstadoDestino
    
    Set getTransicionesValidas = colEstadosValidos
    
LimpiarYSalir:
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.getTransicionesValidas"
    errObj.AddToCallStack "idSolicitud: " & idSolicitud
    errObj.Raise
End Function
Private Function PrecondicionesCumplidas(ByRef sol As Solicitud, ByVal idEstadoDestino As Long, ByRef outMensaje As String, Optional db As DAO.Database) As Boolean
    On Error GoTo Errores
    PrecondicionesCumplidas = True
    
    Select Case sol.tipoSolicitud
        Case "PC"
            If Not PrecondicionesCumplidas_PC(sol, idEstadoDestino, outMensaje, db) Then PrecondicionesCumplidas = False
        Case "CD_CA"
            If Not PrecondicionesCumplidas_CDCA(sol, idEstadoDestino, outMensaje, db) Then PrecondicionesCumplidas = False
        Case "CD_CA_SUB"
            If Not PrecondicionesCumplidas_CDCASUB(sol, idEstadoDestino, outMensaje, db) Then PrecondicionesCumplidas = False
        Case "PC_SUB"
            If Not PrecondicionesCumplidas_PCSUB(sol, idEstadoDestino, outMensaje, db) Then PrecondicionesCumplidas = False
        Case Else
            ' Sin precondiciones
    End Select

LimpiarYSalir:
    Exit Function
Errores:
    PrecondicionesCumplidas = False
    Resume LimpiarYSalir
End Function
Private Function PrecondicionesCumplidas_CDCA(ByRef sol As Solicitud, ByVal idEstadoDestino As Long, ByRef outMensaje As String, Optional db As DAO.Database) As Boolean
    On Error GoTo Errores
    PrecondicionesCumplidas_CDCA = True
    
    Select Case idEstadoDestino
        Case estadoRegistro
            PrecondicionesCumplidas_CDCA = CumplePasoARegistro(sol, outMensaje, db)
        Case estadoModificacion
            PrecondicionesCumplidas_CDCA = CumplePasoAModificacion(sol, outMensaje, db)
        Case estadoValidacion
            PrecondicionesCumplidas_CDCA = CumplePasoAValidacion(sol, outMensaje, db)
        Case estadoRevision
            PrecondicionesCumplidas_CDCA = CumplePasoARevision(sol, outMensaje, db)
        Case estadoFormalizacion
            PrecondicionesCumplidas_CDCA = CumplePasoAFormalizacion(sol, outMensaje, db)
        Case estadoAprobada
            PrecondicionesCumplidas_CDCA = CumplePasoAAprobada(sol, outMensaje, db)
    End Select

LimpiarYSalir:
    Exit Function
Errores:
    PrecondicionesCumplidas_CDCA = False
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.PrecondicionesCumplidas_CDCA"
    errObj.Raise
End Function
' ==========================================================================
' MÓDULO: WorkflowServicio.cls
' ACCIÓN: Nueva función auxiliar para validar estados de CD-CA-SUB
' ==========================================================================
Private Function PrecondicionesCumplidas_CDCASUB(ByRef sol As Solicitud, ByVal idEstadoDestino As Long, ByRef outMensaje As String, Optional db As DAO.Database) As Boolean
    On Error GoTo Errores
    PrecondicionesCumplidas_CDCASUB = True
    
    Select Case idEstadoDestino
        Case estadoRegistro
            PrecondicionesCumplidas_CDCASUB = CumplePasoARegistro(sol, outMensaje, db)
        Case estadoModificacion
            PrecondicionesCumplidas_CDCASUB = CumplePasoAModificacion(sol, outMensaje, db)
        Case estadoValidacion
            PrecondicionesCumplidas_CDCASUB = CumplePasoAValidacion(sol, outMensaje, db)
        Case estadoRevision
            PrecondicionesCumplidas_CDCASUB = CumplePasoARevision(sol, outMensaje, db)
        Case estadoFormalizacion
            PrecondicionesCumplidas_CDCASUB = CumplePasoAFormalizacion(sol, outMensaje, db)
        Case estadoAprobada
            PrecondicionesCumplidas_CDCASUB = CumplePasoAAprobada(sol, outMensaje, db)
    End Select

LimpiarYSalir:
    Exit Function
Errores:
    PrecondicionesCumplidas_CDCASUB = False
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.PrecondicionesCumplidas_CDCASUB"
    errObj.Raise
End Function

Private Function PrecondicionesCumplidas_PC(ByRef sol As Solicitud, ByVal idEstadoDestino As Long, ByRef outMensaje As String, Optional db As DAO.Database) As Boolean
    On Error GoTo Errores
    PrecondicionesCumplidas_PC = True
    
    Select Case idEstadoDestino
        Case estadoRegistro
            PrecondicionesCumplidas_PC = CumplePasoARegistro(sol, outMensaje, db)
        Case estadoModificacion
            PrecondicionesCumplidas_PC = CumplePasoAModificacion(sol, outMensaje, db)
        Case estadoValidacion
            PrecondicionesCumplidas_PC = CumplePasoAValidacion(sol, outMensaje, db)
        Case estadoRevision
            PrecondicionesCumplidas_PC = CumplePasoARevision(sol, outMensaje, db)
        Case estadoFormalizacion
            PrecondicionesCumplidas_PC = CumplePasoAFormalizacion(sol, outMensaje, db)
        Case estadoAprobada
            PrecondicionesCumplidas_PC = CumplePasoAAprobada(sol, outMensaje, db)
    End Select

LimpiarYSalir:
    Exit Function
Errores:
    PrecondicionesCumplidas_PC = False
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.PrecondicionesCumplidas_PC"
    errObj.Raise
End Function

Private Function PrecondicionesCumplidas_PCSUB(ByRef sol As Solicitud, ByVal idEstadoDestino As Long, ByRef outMensaje As String, Optional db As DAO.Database) As Boolean
    On Error GoTo Errores
    PrecondicionesCumplidas_PCSUB = True
    
    Select Case idEstadoDestino
        Case estadoRegistro
            PrecondicionesCumplidas_PCSUB = CumplePasoARegistro(sol, outMensaje, db)
        Case estadoModificacion
            PrecondicionesCumplidas_PCSUB = CumplePasoAModificacion(sol, outMensaje, db)
        Case estadoValidacion
            PrecondicionesCumplidas_PCSUB = CumplePasoAValidacion(sol, outMensaje, db)
        Case estadoRevision
            PrecondicionesCumplidas_PCSUB = CumplePasoARevision(sol, outMensaje, db)
        Case estadoFormalizacion
            PrecondicionesCumplidas_PCSUB = CumplePasoAFormalizacion(sol, outMensaje, db)
        Case estadoAprobada
            PrecondicionesCumplidas_PCSUB = CumplePasoAAprobada(sol, outMensaje, db)
    End Select

LimpiarYSalir:
    Exit Function
Errores:
    PrecondicionesCumplidas_PCSUB = False
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.PrecondicionesCumplidas_PCSUB"
    errObj.Raise
End Function

Private Function CumplePasoARegistro(ByRef sol As Solicitud, ByRef outMensaje As String, Optional db As DAO.Database) As Boolean
    Dim pcServ As New DatosPCServicio
    Dim cdcaServ As New DatosCDCAServicio
    Dim cdcaSubServ As New DatosCDCASUBServicio
    Dim pcsubServ As New DatosPCSUBServicio
    
    On Error GoTo Errores
    CumplePasoARegistro = True
    outMensaje = ""
    
    Select Case sol.tipoSolicitud
        Case "PC"
            If Not pcServ.EsDatosGeneralesCompleta(sol.idSolicitud, db) Then
                outMensaje = "Guarde primero la seccion 'Datos Generales'."
                CumplePasoARegistro = False
                GoTo LimpiarYSalir
            End If
        Case "CD_CA"
            If Not cdcaServ.EsDatosGeneralesCompleta(sol.idSolicitud, db) Then
                outMensaje = "Guarde primero la seccion 'Datos Generales'."
                CumplePasoARegistro = False
                GoTo LimpiarYSalir
            End If
        Case "CD_CA_SUB"
            If Not cdcaSubServ.EsDatosGeneralesCompleta(sol.idSolicitud, db) Then
                outMensaje = "Guarde primero la seccion 'Datos Generales'."
                CumplePasoARegistro = False
                GoTo LimpiarYSalir
            End If
        Case "PC_SUB"
            If Not pcsubServ.EsDatosGeneralesCompleta(sol.idSolicitud, db) Then
                outMensaje = "Guarde primero la seccion 'Datos Generales'."
                CumplePasoARegistro = False
                GoTo LimpiarYSalir
            End If
    End Select
    
LimpiarYSalir:
    Exit Function
Errores:
    CumplePasoARegistro = False
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.CumplePasoARegistro"
    errObj.Raise
End Function

Private Function CumplePasoAModificacion(ByRef sol As Solicitud, ByRef outMensaje As String, Optional db As DAO.Database) As Boolean
    Dim pcServ As New DatosPCServicio
    Dim cdcaServ As New DatosCDCAServicio
    Dim cdcaSubServ As New DatosCDCASUBServicio
    Dim pcsubServ As New DatosPCSUBServicio
    Dim parteTecnicaCompleta As Boolean
    Dim pRazonFallo As String

    On Error GoTo Errores
    CumplePasoAModificacion = True
    outMensaje = ""

    ' Slice 3.3 (D1): pre-leemos el flag por gemelo (acceso DAO inevitable) y
    ' delegamos la evaluación booleana al helper puro BR-002. El wording del
    ' callsite se preserva integro.
    Select Case sol.tipoSolicitud
        Case "PC"
            parteTecnicaCompleta = pcServ.EsParteTecnicaCompleta(sol.idSolicitud, db)
        Case "CD_CA"
            parteTecnicaCompleta = cdcaServ.EsParteTecnicaCompleta(sol.idSolicitud, db)
        Case "CD_CA_SUB"
            parteTecnicaCompleta = cdcaSubServ.EsParteTecnicaCompleta(sol.idSolicitud, db)
        Case "PC_SUB"
            parteTecnicaCompleta = pcsubServ.EsParteTecnicaCompleta(sol.idSolicitud, db)
        Case Else
            parteTecnicaCompleta = True
    End Select

    If Not PrecondicionesWorkflowHelper_ParteTecnicaCompleta(sol.tipoSolicitud, "estadoModificacion", parteTecnicaCompleta, pRazonFallo) Then
        outMensaje = "Faltan campos obligatorios en las pestañas 'Propuesta' o 'Impacto'."
        CumplePasoAModificacion = False
        GoTo LimpiarYSalir
    End If
    
    If sol.revisionCalidadEstado = "RECHAZADO" Then
        If Not VerificarCambiosTrasRechazo(sol) Then
            Dim msg As String
            msg = "CALIDAD RECHAZÓ ESTA SOLICITUD." & vbCrLf & vbCrLf & _
                  "Motivo: " & sol.revisionCalidadComentarios & vbCrLf & vbCrLf & _
                  "El sistema NO ha detectado cambios en los apartados técnicos (Propuesta/Impacto) desde el rechazo." & vbCrLf & vbCrLf & _
                  "Debe modificar la información técnica según los comentarios de Calidad antes de volver a enviar."
            outMensaje = msg
            CumplePasoAModificacion = False
            GoTo LimpiarYSalir
        End If
    End If
    
LimpiarYSalir:
    Exit Function
Errores:
    CumplePasoAModificacion = False
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.CumplePasoAModificacion"
    errObj.Raise
End Function

Private Function CumplePasoAValidacion(ByRef sol As Solicitud, ByRef outMensaje As String, Optional db As DAO.Database) As Boolean
    Dim pcServ As New DatosPCServicio
    Dim cdcaServ As New DatosCDCAServicio
    Dim cdcaSubServ As New DatosCDCASUBServicio
    Dim pcsubServ As New DatosPCSUBServicio
    Dim parteTecnicaCompleta As Boolean
    Dim pRazonFallo As String

    On Error GoTo Errores
    CumplePasoAValidacion = True
    outMensaje = ""

    ' Slice 3.3 (D1): pre-leemos el flag por gemelo y delegamos la evaluación
    ' booleana al helper puro BR-002. El wording del callsite se preserva.
    Select Case sol.tipoSolicitud
        Case "PC"
            parteTecnicaCompleta = pcServ.EsParteTecnicaCompleta(sol.idSolicitud, db)
        Case "CD_CA"
            parteTecnicaCompleta = cdcaServ.EsParteTecnicaCompleta(sol.idSolicitud, db)
        Case "CD_CA_SUB"
            parteTecnicaCompleta = cdcaSubServ.EsParteTecnicaCompleta(sol.idSolicitud, db)
        Case "PC_SUB"
            parteTecnicaCompleta = pcsubServ.EsParteTecnicaCompleta(sol.idSolicitud, db)
        Case Else
            parteTecnicaCompleta = True
    End Select

    If Not PrecondicionesWorkflowHelper_ParteTecnicaCompleta(sol.tipoSolicitud, "estadoValidacion", parteTecnicaCompleta, pRazonFallo) Then
        outMensaje = "La parte técnica debe estar completa antes de validar."
        CumplePasoAValidacion = False
        GoTo LimpiarYSalir
    End If
    
    If Not EsDictamenRACCompleto(sol, db) Then
                outMensaje = "Debe cumplimentar y guardar los datos obligatorios en la pestaña 'RAC' antes de pasar a Validación."
        CumplePasoAValidacion = False
        GoTo LimpiarYSalir
    End If
    
    If Not EsAprobacionSuministradorCompleta(sol.idSolicitud, db) Then
                outMensaje = "Debe cumplimentar y guardar la pestaña 'Aprobación Suministrador' antes de pasar a Validación."
        CumplePasoAValidacion = False
        GoTo LimpiarYSalir
    End If

    ' Slice 3.3 TODO: delegar tambien a PrecondicionesWorkflowHelper_RACCompleto
    ' (BR-003) requiere exponer los campos RAC (codigo/nombre/decision) por
    ' gemelo desde los repositorios. EsDictamenRACCompleto encapsula esa lectura
    ' internamente y devuelve un Boolean; conservamos esa semántica para no
    ' romper a los callers. Cubrir en un slice posterior de extracción de
    ' campos RAC.
    
LimpiarYSalir:
    Exit Function
Errores:
    CumplePasoAValidacion = False
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.CumplePasoAValidacion"
    errObj.Raise
End Function

Private Function CumplePasoARevision(ByRef sol As Solicitud, ByRef outMensaje As String, Optional db As DAO.Database) As Boolean
    Dim adjServ As New AdjuntosServicio
    Dim etapaValidacionNombre As String
    Dim adjuntoEnEtapa As Boolean

    On Error GoTo Errores
    CumplePasoARevision = True
    outMensaje = ""

    ' Slice 3.3 (D3): etapaValidacionNombre puede incluir tildes
    ' (p.ej. "Validación"). El helper tolera ambas formas.
    etapaValidacionNombre = GetNombreEstadoSafe(estadoValidacion)
    adjuntoEnEtapa = adjServ.ExisteAdjuntoEtapa(sol.idSolicitud, etapaValidacionNombre, db)

    If Not PrecondicionesWorkflowHelper_AdjuntoEnEtapa(etapaValidacionNombre, adjuntoEnEtapa) Then
        outMensaje = "Debe generar y adjuntar el borrador Word antes de pasar a Revision."
        CumplePasoARevision = False
        GoTo LimpiarYSalir
    End If
    
LimpiarYSalir:
    Exit Function
Errores:
    CumplePasoARevision = False
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.CumplePasoARevision"
    errObj.Raise
End Function

Private Function CumplePasoAFormalizacion(ByRef sol As Solicitud, ByRef outMensaje As String, Optional db As DAO.Database) As Boolean
    Dim resultadoRAC As String
    Dim comentariosRAC As String

    On Error GoTo Errores
    CumplePasoAFormalizacion = True
    outMensaje = ""

    Call ValidacionRevisionRepositorio.GetUltimoEstadoValidacion(sol.idSolicitud, resultadoRAC, comentariosRAC, db)
    resultadoRAC = UCase$(Trim$(Nz(resultadoRAC, "")))

    If resultadoRAC = "" Or resultadoRAC = "PENDIENTE" Then
        outMensaje = "Debe registrar la respuesta del RAC antes de pasar a Formalización."
        CumplePasoAFormalizacion = False
        GoTo LimpiarYSalir
    End If

    ' Slice 3.3 (D1): delegamos al helper puro BR-005 la decisión APROBADO.
    ' El wording del callsite se preserva integro.
    If Not PrecondicionesWorkflowHelper_ResultadoValidacionAprobado(resultadoRAC) Then
        outMensaje = "El RAC debe estar aprobado antes de pasar a Formalización."
        CumplePasoAFormalizacion = False
        GoTo LimpiarYSalir
    End If
     
LimpiarYSalir:
    Exit Function
Errores:
    CumplePasoAFormalizacion = False
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.CumplePasoAFormalizacion"
    errObj.Raise
End Function

Private Function EsDictamenRACCompleto(ByRef sol As Solicitud, Optional db As DAO.Database) As Boolean
    Dim pcServ As New DatosPCServicio
    Dim cdcaServ As New DatosCDCAServicio
    Dim cdcaSubServ As New DatosCDCASUBServicio
    Dim pcsubServ As New DatosPCSUBServicio
    
    On Error GoTo Errores
    EsDictamenRACCompleto = False
    
    Select Case sol.tipoSolicitud
        Case "PC"
            EsDictamenRACCompleto = pcServ.EsDictamenRACCompleto(sol.idSolicitud, db)
        Case "CD_CA"
            EsDictamenRACCompleto = cdcaServ.EsDictamenRACCompleto(sol.idSolicitud, db)
        Case "CD_CA_SUB"
            EsDictamenRACCompleto = cdcaSubServ.EsDictamenRACCompleto(sol.idSolicitud, db)
        Case "PC_SUB"
            EsDictamenRACCompleto = pcsubServ.EsDictamenRACCompleto(sol.idSolicitud, db)
        Case Else
            EsDictamenRACCompleto = False
    End Select
    
    Exit Function
Errores:
    EsDictamenRACCompleto = False
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.EsDictamenRACCompleto"
    errObj.Raise
End Function

Private Function CumplePasoAAprobada(ByRef sol As Solicitud, ByRef outMensaje As String, Optional db As DAO.Database) As Boolean
    Dim adjServ As New AdjuntosServicio
    Dim etapaFormalizacionNombre As String
    Dim adjuntoEnEtapa As Boolean

    On Error GoTo Errores
    CumplePasoAAprobada = True
    outMensaje = ""

    ' Slice 3.3 (D3): etapaFormalizacionNombre puede incluir tildes. El helper
    ' tolera tanto el canónico ("Cierre") como la variante localizada.
    etapaFormalizacionNombre = GetNombreEstadoSafe(estadoFormalizacion)
    adjuntoEnEtapa = adjServ.ExisteAdjuntoEtapa(sol.idSolicitud, etapaFormalizacionNombre, db)

    If Not PrecondicionesWorkflowHelper_AdjuntoEnEtapa(etapaFormalizacionNombre, adjuntoEnEtapa) Then
        outMensaje = "Es obligatorio adjuntar el documento de Cierre antes de aprobar la solicitud."
        CumplePasoAAprobada = False
        GoTo LimpiarYSalir
    End If
    
LimpiarYSalir:
    Exit Function
Errores:
    CumplePasoAAprobada = False
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.CumplePasoAAprobada"
    errObj.Raise
End Function
Public Sub EjecutarTransicion(ByRef sol As Solicitud, ByVal idEstadoDestino As Long, ByVal usuario As usuario, Optional ByRef dbTransaccional As DAO.Database)
    Dim estadoAnteriorID As Long
    Dim logServ As New LogServicio
    Dim logEst As New LogEstado
    
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim fso As Object
    Dim requiereGuardarSolicitud As Boolean
    Dim sqlPurga As String
    Dim gestionaTransaccionLocal As Boolean
    
    Dim pcServ As New DatosPCServicio
    Dim cdcaServ As New DatosCDCAServicio
    Dim cdcaSubServ As New DatosCDCASUBServicio
    Dim racCompleto As Boolean
    Dim mensajeValidacion As String
    
    On Error GoTo Errores
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' --- 1. TRANSACCIÓN ---
    If dbTransaccional Is Nothing Then
        gestionaTransaccionLocal = True
        Set ws = getWorkspace() ' Workspace aislado
        Set db = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
    Else
        gestionaTransaccionLocal = False
        Set db = dbTransaccional
    End If
    
    ' Validaciones
    If Not EsTransicionPermitida(sol.idEstadoInterno, idEstadoDestino, rolUsuario) Then
        Err.Raise 513, , "Transición no permitida para su rol."
    End If
    mensajeValidacion = ""
    If Not PrecondicionesCumplidas(sol, idEstadoDestino, mensajeValidacion, db) Then
        Err.Raise 513, "EjecutarTransicion", mensajeValidacion
    End If
    
    ' --- 3. ESCRITURA BD (CRÍTICO) ---
    estadoAnteriorID = sol.idEstadoInterno
    sol.idEstadoInterno = idEstadoDestino
    sol.fechaModificacion = Now()
    sol.usuarioModificacion = usuario.nombre
    
    ' Desactivar rechazos activos al promover desde Desarrollo Técnico (3 -> 4)
    ' GAP-004 FIX: NO desactivamos el rechazo para que Calidad pueda ver los cambios del técnico
    ' If estadoAnteriorID = estadoDesarrolloTecnico And idEstadoDestino = estadoModificacion Then
    '     Call RechazoRepositorio.DesactivarRechazosPrevios(sol.idSolicitud, db)
    '     sol.revisionCalidadEstado = "PENDIENTE"
    '     Debug.Print "DEBUG DT-08-001: Limpiando revisionCalidadEstado a PENDIENTE para solicitud " & sol.idSolicitud
    ' End If
    
    ' Desactivar rechazos activos al promover desde Modificación (4 -> 5)
    If estadoAnteriorID = estadoModificacion And idEstadoDestino = estadoValidacion Then
        Call RechazoRepositorio.DesactivarRechazosPrevios(sol.idSolicitud, db)
    End If
    
    ' Crear ciclo de validación al llegar a Validación (cualquier origen)
    If idEstadoDestino = estadoValidacion Then
        ' Siempre pasar hash vacío para permitir exportar la primera versión o la siguiente
        ' sin requerir cambios adicionales
        Call ValidacionRevisionRepositorio.CrearNuevoCicloValidacion(sol.idSolicitud, usuario.nombre, "", db)
    End If
    
    If idEstadoDestino = estadoDesarrolloTecnico Then
        sol.revisionCalidadEstado = "PENDIENTE"
    End If
    
    Call SolicitudRepositorio.GuardarSolicitud(sol, db)
    
    With logEst
        .idSolicitud = sol.idSolicitud
        .idEstadoAnterior = estadoAnteriorID
        .idEstadoNuevo = idEstadoDestino
        .fechaTransicion = sol.fechaModificacion
        .usuarioTransicion = sol.usuarioModificacion
    End With
    Call LogEstadoRepositorio.Guardar(logEst, db)
    
    Dim estadoAnteriorNombre As String, estadoNuevoNombre As String
    estadoAnteriorNombre = GetNombreEstadoSafe(estadoAnteriorID)
    estadoNuevoNombre = GetNombreEstadoSafe(idEstadoDestino)
    Call logServ.RegistrarCambio("tbSolicitudes", sol.idSolicitud, "idEstadoInterno", estadoAnteriorID & " (" & estadoAnteriorNombre & ")", idEstadoDestino & " (" & estadoNuevoNombre & ")", "TRANSICIÓN_WF", db)

    ' --- 4. COMMIT (CRÍTICO) ---
    If gestionaTransaccionLocal Then
        ws.CommitTrans
    End If

    ' Spec-118: Invalidar cache del Timeline tras transicion
    Call Me.InvalidarCacheTimeline(sol.idSolicitud)
   ' Debug.Print "Spec-118 CACHE: Cache invalidado para solicitud " & sol.idSolicitud & " tras transicion a estado " & idEstadoDestino

    ' =======================================================================
    ' --- 6. NOTIFICACIONES (NO BLOQUEANTE) ---
    ' Si el correo falla, NO queremos hacer Rollback de los datos.
    ' Slice 3.3 (D4): la política no-bloqueante se delega a
    ' NotificacionHelper_NotificarSiPosible, que envuelve la llamada con su
    ' propio manejo silencioso de errores y devuelve "OK" o
    ' "NOTIF_ERROR_<n>_<desc>". Eliminamos el chequeo manual de Err inline.
    ' =======================================================================

    Dim notifServ As New NotificacionServicio
    Dim solServ As New SolicitudServicio
    Dim vm As SolicitudViewModel
    Dim destinatario As String, copia As String, asunto As String, infoExpediente As String
    Dim descriptorAccion As String
    Dim sufijoSubsanacion As String
    Dim nombreEstadoNotif As String
    Dim notifResult As String

    Set vm = solServ.getSolicitudViewModelPorID(sol.idSolicitud, db)
    copia = m_ObjUsuarioReal.CorreoUsuario
    infoExpediente = vm.Expediente.Nemotecnico & " | " & vm.TipoSolicitudNombre

    nombreEstadoNotif = GetNombreEstadoSafe(idEstadoDestino)
    If idEstadoDestino = estadoValidacion Then
        If estadoAnteriorID = estadoModificacion Then
            nombreEstadoNotif = "VALIDADO_POR_CALIDAD"
        End If
    ElseIf UCase$(Trim$(Nz(sol.revisionCalidadEstado, ""))) = "RECHAZADO" Then
        If idEstadoDestino = estadoDesarrolloTecnico Then
            nombreEstadoNotif = "RECHAZADO_POR_CALIDAD"
        End If
    End If
    If idEstadoDestino = estadoRevision And estadoAnteriorID = estadoValidacion Then
        nombreEstadoNotif = "REVISION_RAC"
    End If
    If (vm.vm_racValidacionEstado = "RECHAZADO" Or vm.vm_racValidacionEstado = "COMENTARIOS") And idEstadoDestino = estadoValidacion Then
        nombreEstadoNotif = "REVISION_RAC"
    End If

    Select Case idEstadoDestino
        Case estadoDesarrolloTecnico
            descriptorAccion = "DESARROLLO TÉCNICO"
            asunto = "CONDOR - " & nombreEstadoNotif & " | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
        Case estadoRegistro
            asunto = "CONDOR - " & nombreEstadoNotif & " | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
        Case estadoRevision
            asunto = "CONDOR - " & nombreEstadoNotif & " | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
        Case estadoModificacion
            If nombreEstadoNotif <> "PENDIENTE DE VALIDACION" Then
                sufijoSubsanacion = " (RECHAZADA POR CALIDAD)"
            Else
                sufijoSubsanacion = ""
            End If
            asunto = "CONDOR - " & nombreEstadoNotif & sufijoSubsanacion & " | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
        Case estadoValidacion
            asunto = "CONDOR - " & nombreEstadoNotif & " | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
            If (vm.vm_racValidacionEstado = "RECHAZADO" Or vm.vm_racValidacionEstado = "COMENTARIOS") Then
                asunto = "CONDOR - " & nombreEstadoNotif & " TRAS COMENTARIOS RAC/CLIENTE | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
            End If
        Case estadoFormalizacion
            nombreEstadoNotif = "APROBADO_RAC"
            asunto = "CONDOR - " & nombreEstadoNotif & " | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
        Case estadoAprobada
            nombreEstadoNotif = "APROBACIÓN"
            asunto = "CONDOR - " & nombreEstadoNotif & " | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
        Case Else
            asunto = "CONDOR - " & nombreEstadoNotif & " | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
    End Select

    destinatario = notifServ.ObtenerDestinatariosPorEstado(nombreEstadoNotif, vm)

    If destinatario <> "" Then
        ' El helper absorbe errores internamente y devuelve "OK" o un
        ' marcador "NOTIF_ERROR_<n>_<desc>" sin propagar el error.
        notifResult = NotificacionHelper_NotificarSiPosible( _
            notifServ, _
            vm, _
            asunto, _
            destinatario, _
            copia, _
            nombreEstadoNotif)
        If notifResult <> "OK" Then
            ' TODO: loggear notifResult si NO es "OK" (logErrores o equivalente).
            ' LogNotificacionError "TRANSICION_WF", sol.idSolicitud, notifResult
        End If
    End If
    ' =======================================================================

LimpiarYSalir:
    On Error Resume Next
    If gestionaTransaccionLocal And Not db Is Nothing Then db.Close
    Set ws = Nothing
    Exit Sub
    
Errores:
    Dim savedErrNum As Long
    Dim savedErrDesc As String
    savedErrNum = Err.Number
    savedErrDesc = Err.description
    
    If gestionaTransaccionLocal And Not ws Is Nothing Then
        ws.Rollback
    End If
    
    Dim errObj As New CondorError
    errObj.Create savedErrNum, savedErrDesc, "WorkflowServicio.EjecutarTransicion"
    errObj.AddToCallStack "Transición de ID " & sol.idSolicitud & " a Estado " & idEstadoDestino
    errObj.Raise
End Sub

Private Function PermisoSuficiente(ByVal rolUsuarioActual As rol, ByVal rolRequerido As String) As Boolean
    ' Slice 3.3 nota: la tabla inline se preserva porque modela "rol minimo
    ' requerido" (Admin/Calidad/Tecnico/Ingenieria), abstracción distinta del
    ' helper JerarquiaRolesHelper_UsuarioPuedeTransicionarDesde, que evalua
    ' "este usuario puede transicionar origen->destino". Delegar requiere
    ' recibir contexto de estado en cada callsite y cambiar firmas de
    ' callers; queda como slice posterior para evitar invasion.
    PermisoSuficiente = False
    Select Case rolRequerido
        Case "Admin": If rolUsuarioActual = rol.Administrador Then PermisoSuficiente = True
        Case "Calidad": If rolUsuarioActual = rol.Administrador Or rolUsuarioActual = rol.Calidad Then PermisoSuficiente = True
        Case "Ingenieria", "Tecnico": If rolUsuarioActual = rol.Administrador Or rolUsuarioActual = rol.Calidad Or rolUsuarioActual = rol.Tecnico Then PermisoSuficiente = True
        Case Else: PermisoSuficiente = True ' Si no se especifica rol, se permite
    End Select
End Function
Public Function PuedeEditarBloque(ByVal idSolicitud As Long, ByVal bloque As enumBloqueFormulario) As Boolean
    Dim solServ As New SolicitudServicio
    Dim sol As Solicitud
    On Error GoTo Errores

    Set sol = solServ.getSolicitudPorID(idSolicitud)
    If sol Is Nothing Then Err.Raise 513, "PuedeEditarBloque", "No existe la solicitud."

    Select Case sol.tipoSolicitud
        Case "PC"
            PuedeEditarBloque = PermiteEdicion_PC(sol, rolUsuario, bloque)
        Case "CD_CA"
            PuedeEditarBloque = PermiteEdicion_CDCA(sol, rolUsuario, bloque)
        Case "CD_CA_SUB"
            ' Reutilizamos lógica de CDCA ya que las reglas de roles/estados son idénticas
            PuedeEditarBloque = PermiteEdicion_CDCASUB(sol, rolUsuario, bloque)
        Case "PC_SUB"
            ' Gemelo de PC — las reglas de roles/estados son idénticas
            PuedeEditarBloque = PermiteEdicion_PCSUB(sol, rolUsuario, bloque)
        Case Else
            PuedeEditarBloque = False
    End Select
    
    Exit Function
Errores:
    PuedeEditarBloque = False
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.PuedeEditarBloque"
    errObj.Raise
End Function

Public Function EsAprobacionSuministradorCompleta(ByVal idSolicitud As Long, Optional db As DAO.Database) As Boolean
    Dim solServ As New SolicitudServicio
    Dim sol As Solicitud
    Dim pcServ As New DatosPCServicio
    Dim cdcaServ As New DatosCDCAServicio
    Dim cdcaSubServ As New DatosCDCASUBServicio
    Dim pcsubServ As New DatosPCSUBServicio
    
    On Error GoTo Errores
    
    Set sol = solServ.getSolicitudPorID(idSolicitud, db)
    If sol Is Nothing Then Err.Raise 513, "EsAprobacionSuministradorCompleta", "No existe la solicitud."
    
    Select Case sol.tipoSolicitud
        Case "PC"
            EsAprobacionSuministradorCompleta = pcServ.EsAprobacionSuministradorCompleta(idSolicitud, db)
        Case "CD_CA"
            EsAprobacionSuministradorCompleta = cdcaServ.EsAprobacionSuministradorCompleta(idSolicitud, db)
        Case "CD_CA_SUB"
            EsAprobacionSuministradorCompleta = cdcaSubServ.EsAprobacionSuministradorCompleta(idSolicitud, db)
        Case "PC_SUB"
            EsAprobacionSuministradorCompleta = pcsubServ.EsAprobacionSuministradorCompleta(idSolicitud, db)
        Case Else
            EsAprobacionSuministradorCompleta = False
    End Select
    
    Exit Function
Errores:
    EsAprobacionSuministradorCompleta = False
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.EsAprobacionSuministradorCompleta"
    errObj.Raise
End Function

Private Function PermiteEdicion_CDCA(ByRef sol As Solicitud, ByVal rolUsuarioActual As rol, ByVal bloque As enumBloqueFormulario) As Boolean
    On Error GoTo FalloSilencioso
    PermiteEdicion_CDCA = False
    
    Dim esEditorPrincipal As Boolean
    esEditorPrincipal = (rolUsuarioActual = rol.Calidad Or rolUsuarioActual = rol.Administrador)
    Dim esTecnico As Boolean
    esTecnico = (rolUsuarioActual = rol.Tecnico)

    ' --- 1. BLOQUEOS DE ESTADO FINAL (Rechazada / Aprobada) ---
    If sol.idEstadoInterno = estadoRechazada Then Exit Function ' Nadie edita nada en Rechazada
    
    ' En Aprobada solo se toca Decisión (para borrarla/reabrir)
    If sol.idEstadoInterno = estadoAprobada Then
        If bloque = Bloque_DecisionFinal And esEditorPrincipal Then
            PermiteEdicion_CDCA = True
        End If
        Exit Function
    End If

    ' --- 2. BLOQUEO ESTRICTO EN FORMALIZACIÓN (Fase 7) ---
    If sol.idEstadoInterno = estadoFormalizacion Then
        ' En Formalización, el RAC ya aprobó. El contenido técnico y general es SAGRADO.
        ' Solo permitimos editar las Firmas (para el documento) y la Decisión Final.
        If esEditorPrincipal Then
            If bloque = Bloque_AprobacionSuministrador Or bloque = Bloque_DecisionFinal Then
                PermiteEdicion_CDCA = True
            End If
        End If
        ' Salimos inmediatamente. Si el bloque no es uno de esos dos, devuelve False (Bloqueado).
        Exit Function
    End If

    ' --- 3. LÓGICA PARA FASES ANTERIORES (< Formalización) ---
    Select Case bloque
        
        ' --- DATOS GENERALES ---
        Case Bloque_Generales
            If esEditorPrincipal Then PermiteEdicion_CDCA = True
            
        ' --- DATOS TÉCNICOS ---
        Case Bloque_Propuesta, Bloque_Impacto
             ' Fase 3: Desarrollo Técnico (Técnico rellena)
            If sol.idEstadoInterno = estadoDesarrolloTecnico Then
                If esTecnico Then PermiteEdicion_CDCA = True
            
            ' Fase 4: Modificación (Solo Calidad Revisa)
            ElseIf sol.idEstadoInterno = estadoModificacion Then
                ' REGLA ESTRICTA (Spec-003): Técnicos NO editan en Modificación. Solo Calidad.
                If esEditorPrincipal Then PermiteEdicion_CDCA = True
                
            ' Fases Posteriores: Validación, Revisión... (Solo Calidad)
            ElseIf sol.idEstadoInterno > estadoModificacion Then
                If esEditorPrincipal Then PermiteEdicion_CDCA = True
            End If
            
        ' --- APROBACIÓN (FIRMAS) ---
        Case Bloque_AprobacionSuministrador
            If sol.idEstadoInterno > estadoDesarrolloTecnico Then
                If esEditorPrincipal Then PermiteEdicion_CDCA = True
            End If

        ' --- DICTAMEN RAC ---
        Case Bloque_DictamenRAC
            If sol.idEstadoInterno > estadoDesarrolloTecnico Then
                If esEditorPrincipal Then PermiteEdicion_CDCA = True
            End If
            
        ' (DecisionFinal no llega aquí porque ya se manejó en el bloque de Formalización)
            
    End Select
    
    Exit Function
FalloSilencioso:
    PermiteEdicion_CDCA = False
End Function
' -----------------------------------------------------------------------------------
' 3. PERMISOS PARA CD_CA_SUB (Subcontratista) - NUEVA FUNCIÓN INDEPENDIENTE
' -----------------------------------------------------------------------------------
Private Function PermiteEdicion_CDCASUB(ByRef sol As Solicitud, ByVal rolUsuarioActual As rol, ByVal bloque As enumBloqueFormulario) As Boolean
    On Error GoTo FalloSilencioso
    PermiteEdicion_CDCASUB = False
    
    Dim esEditorPrincipal As Boolean: esEditorPrincipal = (rolUsuarioActual = rol.Calidad Or rolUsuarioActual = rol.Administrador)
    Dim esTecnico As Boolean: esTecnico = (rolUsuarioActual = rol.Tecnico)

    If sol.idEstadoInterno = estadoRechazada Then Exit Function
    If sol.idEstadoInterno = estadoAprobada Then
        If bloque = Bloque_DecisionFinal And esEditorPrincipal Then PermiteEdicion_CDCASUB = True
        Exit Function
    End If
    If sol.idEstadoInterno = estadoFormalizacion Then
        If esEditorPrincipal And (bloque = Bloque_AprobacionSuministrador Or bloque = Bloque_DecisionFinal) Then PermiteEdicion_CDCASUB = True
        Exit Function
    End If

    Select Case bloque
        Case Bloque_Generales
            If esEditorPrincipal Then PermiteEdicion_CDCASUB = True
            
        Case Bloque_Propuesta, Bloque_Impacto
            ' CORRECCIÓN APLICADA AQUÍ:
            If sol.idEstadoInterno = estadoDesarrolloTecnico Then
                If esTecnico Then PermiteEdicion_CDCASUB = True
            ElseIf sol.idEstadoInterno = estadoModificacion Then
                ' REGLA ESTRICTA (Spec-003): Técnicos NO editan en Modificación. Solo Calidad.
                If esEditorPrincipal Then PermiteEdicion_CDCASUB = True
            ElseIf sol.idEstadoInterno > estadoModificacion Then
                If esEditorPrincipal Then PermiteEdicion_CDCASUB = True
            End If
            
        Case Bloque_AprobacionSuministrador
            If sol.idEstadoInterno > estadoDesarrolloTecnico And esEditorPrincipal Then PermiteEdicion_CDCASUB = True
            
        Case Bloque_DictamenRAC
            If sol.idEstadoInterno > estadoDesarrolloTecnico And esEditorPrincipal Then PermiteEdicion_CDCASUB = True
    End Select
    Exit Function
FalloSilencioso:
    PermiteEdicion_CDCASUB = False
End Function
Private Function PermiteEdicion_PC(ByRef sol As Solicitud, ByVal rolUsuarioActual As rol, ByVal bloque As enumBloqueFormulario) As Boolean
    ' ESTRUCTURA GEMELA A CDCA
    On Error GoTo FalloSilencioso
    PermiteEdicion_PC = False
    
    Dim esEditorPrincipal As Boolean
    esEditorPrincipal = (rolUsuarioActual = rol.Calidad Or rolUsuarioActual = rol.Administrador)
    Dim esTecnico As Boolean
    esTecnico = (rolUsuarioActual = rol.Tecnico)

    ' --- 1. BLOQUEOS DE ESTADO FINAL ---
    If sol.idEstadoInterno = estadoRechazada Then Exit Function
    
    If sol.idEstadoInterno = estadoAprobada Then
        If bloque = Bloque_DecisionFinal And esEditorPrincipal Then
            PermiteEdicion_PC = True
        End If
        Exit Function
    End If

    ' --- 2. BLOQUEO ESTRICTO EN FORMALIZACIÓN (Fase 7) ---
    If sol.idEstadoInterno = estadoFormalizacion Then
        ' Solo permitimos editar Firmas y Decisión Final.
        If esEditorPrincipal Then
            If bloque = Bloque_AprobacionSuministrador Or bloque = Bloque_DecisionFinal Then
                PermiteEdicion_PC = True
            End If
        End If
        Exit Function
    End If

    ' --- 3. LÓGICA PARA FASES ANTERIORES ---
    Select Case bloque
        
        ' --- DATOS GENERALES ---
        Case Bloque_Generales
            If esEditorPrincipal Then PermiteEdicion_PC = True
            
        ' --- DATOS TÉCNICOS ---
        Case Bloque_Propuesta, Bloque_Impacto
            ' Fase 3: Desarrollo Técnico
            If sol.idEstadoInterno = estadoDesarrolloTecnico Then
                If esTecnico Then PermiteEdicion_PC = True
                
            ' Fase 4: Modificación (Solo Calidad Revisa)
            ElseIf sol.idEstadoInterno = estadoModificacion Then
                ' REGLA ESTRICTA (Spec-003): Técnicos NO editan en Modificación. Solo Calidad.
                If esEditorPrincipal Then PermiteEdicion_PC = True
                
            ' Fases Posteriores
            ElseIf sol.idEstadoInterno > estadoModificacion Then
                If esEditorPrincipal Then PermiteEdicion_PC = True
            End If
            
        ' --- APROBACIÓN (FIRMAS) ---
        Case Bloque_AprobacionSuministrador
            If sol.idEstadoInterno > estadoDesarrolloTecnico Then
                If esEditorPrincipal Then PermiteEdicion_PC = True
            End If
            
        ' --- DICTAMEN RAC ---
        Case Bloque_DictamenRAC
            If sol.idEstadoInterno > estadoDesarrolloTecnico Then
                If esEditorPrincipal Then PermiteEdicion_PC = True
            End If
            
    End Select
    
    Exit Function
FalloSilencioso:
    PermiteEdicion_PC = False
End Function

' -----------------------------------------------------------------------------------
' 4. PERMISOS PARA PC_SUB (Subcontratista) — GEMELO EXACTO DE PC
' -----------------------------------------------------------------------------------
Private Function PermiteEdicion_PCSUB(ByRef sol As Solicitud, ByVal rolUsuarioActual As rol, ByVal bloque As enumBloqueFormulario) As Boolean
    ' ESTRUCTURA GEMELA A PC (mismas reglas de roles/estados para subcontratistas)
    On Error GoTo FalloSilencioso
    PermiteEdicion_PCSUB = False
    
    Dim esEditorPrincipal As Boolean
    esEditorPrincipal = (rolUsuarioActual = rol.Calidad Or rolUsuarioActual = rol.Administrador)
    Dim esTecnico As Boolean
    esTecnico = (rolUsuarioActual = rol.Tecnico)

    ' --- 1. BLOQUEOS DE ESTADO FINAL ---
    If sol.idEstadoInterno = estadoRechazada Then Exit Function
    
    If sol.idEstadoInterno = estadoAprobada Then
        If bloque = Bloque_DecisionFinal And esEditorPrincipal Then
            PermiteEdicion_PCSUB = True
        End If
        Exit Function
    End If

    ' --- 2. BLOQUEO ESTRICTO EN FORMALIZACIÓN (Fase 7) ---
    If sol.idEstadoInterno = estadoFormalizacion Then
        ' Solo permitimos editar Firmas y Decisión Final.
        If esEditorPrincipal Then
            If bloque = Bloque_AprobacionSuministrador Or bloque = Bloque_DecisionFinal Then
                PermiteEdicion_PCSUB = True
            End If
        End If
        Exit Function
    End If

    ' --- 3. LÓGICA PARA FASES ANTERIORES ---
    Select Case bloque
        
        ' --- DATOS GENERALES ---
        Case Bloque_Generales
            If esEditorPrincipal Then PermiteEdicion_PCSUB = True
            
        ' --- DATOS TÉCNICOS ---
        Case Bloque_Propuesta, Bloque_Impacto
            ' Fase 3: Desarrollo Técnico
            If sol.idEstadoInterno = estadoDesarrolloTecnico Then
                If esTecnico Then PermiteEdicion_PCSUB = True
                
            ' Fase 4: Modificación (Solo Calidad Revisa)
            ElseIf sol.idEstadoInterno = estadoModificacion Then
                ' REGLA ESTRICTA (Spec-003): Técnicos NO editan en Modificación. Solo Calidad.
                If esEditorPrincipal Then PermiteEdicion_PCSUB = True
                
            ' Fases Posteriores
            ElseIf sol.idEstadoInterno > estadoModificacion Then
                If esEditorPrincipal Then PermiteEdicion_PCSUB = True
            End If
            
        ' --- APROBACIÓN (FIRMAS) ---
        Case Bloque_AprobacionSuministrador
            If sol.idEstadoInterno > estadoDesarrolloTecnico Then
                If esEditorPrincipal Then PermiteEdicion_PCSUB = True
            End If
            
        ' --- DICTAMEN RAC ---
        Case Bloque_DictamenRAC
            If sol.idEstadoInterno > estadoDesarrolloTecnico Then
                If esEditorPrincipal Then PermiteEdicion_PCSUB = True
            End If
            
    End Select
    
    Exit Function
FalloSilencioso:
    PermiteEdicion_PCSUB = False
End Function

Private Function UsuarioPuedeTransicionarDesde(ByVal idEstadoOrigen As Long, ByVal rolUsuarioActual As rol) As Boolean
    ' RESPONSABILIDAD: Comprueba si un usuario, basado en su rol, tiene al menos
    '                  una transición de salida válida desde el estado origen.
    On Error GoTo Errores
    
    Dim colPosiblesDestinos As Scripting.Dictionary
    Dim rolRequerido As Variant
    
    Set colPosiblesDestinos = WorkflowRepositorio.getPosiblesDestinos(idEstadoOrigen)
    
    If colPosiblesDestinos.count = 0 Then
        UsuarioPuedeTransicionarDesde = False
        Exit Function
    End If
    
    For Each rolRequerido In colPosiblesDestinos.items
        If PermisoSuficiente(rolUsuarioActual, CStr(rolRequerido)) Then
            UsuarioPuedeTransicionarDesde = True
            Exit Function
        End If
    Next rolRequerido
    
    UsuarioPuedeTransicionarDesde = False
    Exit Function

Errores:
    UsuarioPuedeTransicionarDesde = False
End Function

Private Function EsTransicionPermitida(ByVal idEstadoOrigen As Long, ByVal idEstadoDestino As Long, ByVal rolUsuarioActual As rol) As Boolean
    ' RESPONSABILIDAD: Valida de forma atómica y eficiente si una transición
    '                  específica está permitida por rol, sin comprobar precondiciones.
    Dim colPosiblesDestinos As Scripting.Dictionary
    Dim rolRequerido As String
    
    On Error GoTo Errores
    
    ' Por defecto, la transición no es permitida.
    EsTransicionPermitida = False
    
    ' 1. Obtener del repositorio solo las transiciones teóricas desde el origen.
    Set colPosiblesDestinos = WorkflowRepositorio.getPosiblesDestinos(idEstadoOrigen)
    
    ' 2. Comprobar si el destino solicitado está en la lista de posibilidades.
    If colPosiblesDestinos.Exists(CStr(idEstadoDestino)) Then
        ' 3. Si está, comprobar si el rol del usuario es suficiente.
        rolRequerido = colPosiblesDestinos(CStr(idEstadoDestino))
        If PermisoSuficiente(rolUsuarioActual, rolRequerido) Then
            EsTransicionPermitida = True
        End If
    End If

LimpiarYSalir:
    Exit Function
Errores:
    EsTransicionPermitida = False
    Resume LimpiarYSalir
End Function
' ==========================================================================
' MÓDULO: WorkflowServicio.cls
' ACCIÓN: Permitir subir documento de Cierre en FORMALIZACION y APROBADA.
' ==========================================================================
Public Function getEtapasDocumentalesAlcanzadas(ByVal idEstadoActual As Long) As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary
    Dim estadoActual As estado
    Dim ordenActual As Long
    Dim key As Variant
    Dim unEstado As estado
    
    On Error GoTo Errores
    
    If m_ObjEntorno.estados.Exists(CStr(idEstadoActual)) Then
        Set estadoActual = m_ObjEntorno.estados(CStr(idEstadoActual))
        ordenActual = estadoActual.orden
    Else
        dict.Add "Otro", "Otro"
        Set getEtapasDocumentalesAlcanzadas = dict
        Exit Function
    End If
    
    ' 1. Estados alcanzados
    For Each key In m_ObjEntorno.estados.Keys
        If Nz(key, "") = "" Then GoTo siguiente
        Set unEstado = m_ObjEntorno.estados(key)
        If unEstado.orden <= ordenActual Then
            Dim nombreEstadoSafe As String
            nombreEstadoSafe = GetNombreEstadoSafe(unEstado.idEstado)
            If Not dict.Exists(nombreEstadoSafe) Then
                dict.Add unEstado.idEstado, nombreEstadoSafe
            End If
        End If
siguiente:
    Next key
    
    ' 2. --- LÓGICA AMPLIADA: Documento de Cierre ---
    ' Permitimos subirlo en FORMALIZACION (para avanzar)
    ' Y TAMBIEN en APROBADA (para sustituir/corregir sin revertir)
    If idEstadoActual = estadoFormalizacion Or idEstadoActual = estadoAprobada Then
        If Not dict.Exists("Cierre") Then
            dict.Add "Cierre", "Cierre"
        End If
    End If
    ' -------------------------------------------
    
    ' 3. Opción genérica
    If Not dict.Exists("Otro") Then
        dict.Add "Otro", "Otro"
    End If
    
    Set getEtapasDocumentalesAlcanzadas = dict
    
LimpiarYSalir:
    Exit Function
Errores:
    Set getEtapasDocumentalesAlcanzadas = New Scripting.Dictionary
    Resume LimpiarYSalir
End Function
' ==========================================================================
' MÓDULO: WorkflowServicio.cls
' ACCIÓN: REEMPLAZO de ReabrirSolicitudCerrada para que sea polimórfica.
' ==========================================================================
Public Sub ReabrirSolicitudCerrada(ByRef sol As Solicitud, ByRef usuario As usuario, ByRef db As DAO.Database)
    Dim idEstadoDestino As Long
    Dim logServ As New LogServicio

    On Error GoTo Errores

    ' 1. VALIDACIONES (sin cambios)
    If usuario.rol <> rol.Administrador And usuario.rol <> rol.Calidad Then
        Err.Raise 513, "ReabrirSolicitud", "No tiene permisos para reabrir una solicitud."
    End If
     If sol.idEstadoInterno <> estadoAprobada Then
        Err.Raise 513, "ReabrirSolicitud", "La solicitud no se encuentra en estado 'APROBADA'."
    End If
    
    ' --- INICIO DE LA CORRECCIÓN ---
    ' 2. BORRAR LOS DATOS DE LA FASE FINAL DE FORMA POLIMÓRFICA
    Select Case sol.tipoSolicitud
        Case "PC"
            Dim pcServ As New DatosPCServicio
            Call pcServ.EliminarDecisionFinal(sol.idSolicitud, db)
        Case "CD_CA"
            Dim cdcaServ As New DatosCDCAServicio
            Call cdcaServ.EliminarDecisionFinal(sol.idSolicitud, db)
        Case "CD_CA_SUB"
            ' --- LÓGICA IMPLEMENTADA ---
            Dim cdcaSubServ As New DatosCDCASUBServicio
            Call cdcaSubServ.EliminarDecisionFinal(sol.idSolicitud, db)
        Case "PC_SUB"
            Dim pcsubServReabrir As New DatosPCSUBServicio
            Call pcsubServReabrir.EliminarDecisionFinal(sol.idSolicitud, db)
    End Select
    ' --- FIN DE LA CORRECCIÓN ---
    
    ' 3. OBTENER ESTADO ANTERIOR Y REVERTIR (sin cambios)
    idEstadoDestino = LogEstadoRepositorio.getUltimoEstadoAnterior(sol.idSolicitud, db)
    If idEstadoDestino = 0 Then idEstadoDestino = estadoFormalizacion ' Fallback seguro
    
    ' Llamamos a nuestára función de retroceso genérica para completar la acción
    Call Me.RevertirAFaseAnterior(sol, idEstadoDestino, usuario, "Reapertura de Solicitud Aprobada", db)
        
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.ReabrirSolicitudCerrada"
    errObj.Raise
End Sub

' ==========================================================================
' MÓDULO: WorkflowServicio.cls
' ACCIÓN: Reemplazar la función getPaginaActivíaPC con la lógica de estados canónicos.
' ==========================================================================
Public Function getPaginaActivaGenerica(ByVal idEstadoActual As Long) As String
    ' RESPONSABILIDAD: Devuelve el nombre del control de navegación (pestaña)
    '                  que corresponde al estado actual del workflow.
    '                  Lógica común para PC, CD_CA y CD_CA_SUB.
    ' VERSIÓN: Refactorizada con los nuevos estados canónicos.
    On Error GoTo Errores
    
    ' Por defecto, si algo falla, devolvemos la primera pestaña.
    getPaginaActivaGenerica = "tabGeneral"
    
    Select Case idEstadoActual
        Case estadoPreregistro, estadoRegistro ' 1, 2
            ' En las etapas iniciales, el foco está en los datos generales.
            getPaginaActivaGenerica = "tabGeneral"
            
        Case estadoDesarrolloTecnico ' 3
            ' El técnico trabaja en Propuesta e Impacto. Lo llevamos a la primera de ellas.
            getPaginaActivaGenerica = "tabPropuesta"
            
        Case estadoModificacion ' 4
            ' Calidad revisa lo que hizo el técnico, así que lo llevamos a esa sección.
            getPaginaActivaGenerica = "tabPropuesta"
            
        Case estadoValidacion ' 5
            ' La acción es generar el documento, pero la información relevante es la de aprobación.
            getPaginaActivaGenerica = "tabAprobacionSuministrador"
            
        Case estadoRevision ' 6
            ' El foco está en el dictamen del RAC.
            getPaginaActivaGenerica = "tabDictamenRAC"
            
        Case estadoFormalizacion ' 7
            ' El foco está en la decisión final y firmas competentes.
            getPaginaActivaGenerica = "tabDecisionFinal"
            
        Case estadoAprobada ' 8
            ' Al estar aprobada, la última pantalla relevante es la de la decisión.
            getPaginaActivaGenerica = "tabDecisionFinal"
            
    End Select
    
LimpiarYSalir:
    Exit Function
Errores:
    ' Ante cualquier error, se devuelve el valor por defecto "tabGeneral".
    Resume LimpiarYSalir
End Function

Public Function getPaginaActivaPC(ByVal idEstadoActual As Long) As String
    getPaginaActivaPC = getPaginaActivaGenerica(idEstadoActual)
End Function

Public Function getPaginaActivaCDCA(ByVal idEstadoActual As Long) As String
    getPaginaActivaCDCA = getPaginaActivaGenerica(idEstadoActual)
End Function

Public Function getPaginaActivaCDCASUB(ByVal idEstadoActual As Long) As String
    getPaginaActivaCDCASUB = getPaginaActivaGenerica(idEstadoActual)
End Function

Public Function getPaginaActivaPCSUB(ByVal idEstadoActual As Long) As String
    getPaginaActivaPCSUB = getPaginaActivaGenerica(idEstadoActual)
End Function

Public Function GetProximaVersionBorrador(ByVal idSolicitud As Long) As String
    Dim ultimoOrdinal As Long
    
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then
        GetProximaVersionBorrador = "v1"
        Exit Function
    End If
    
    ultimoOrdinal = ValidacionRevisionRepositorio.GetUltimoOrdinal(idSolicitud)
    GetProximaVersionBorrador = ValidacionRevisionRepositorio.GetVersionActualBorrador(idSolicitud)
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.GetProximaVersionBorrador"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud
    errObj.Raise
End Function

Public Function ObtenerHayCambiosValidacion(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As String
    Dim cicloActual As ValidacionRevision
    Dim ordinalActual As Long
    Dim hashEnTbSolicitud As String
    Dim hashEnTbRevision As String
    Dim sol As Solicitud
    Dim solServ As New SolicitudServicio
    Dim comentarios As String
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set db = getdb()
    Set sol = solServ.getSolicitudPorID(idSolicitud)
    If sol Is Nothing Then
        ObtenerHayCambiosValidacion = "ERROR"
        Exit Function
    End If
    
    Set cicloActual = ValidacionRevisionRepositorio.GetCicloAbierto(idSolicitud, db)
    
    If cicloActual Is Nothing Then
        ObtenerHayCambiosValidacion = "V1"
        Exit Function
    End If
    
    ordinalActual = cicloActual.ordinal
    
    If ordinalActual <= 1 Then
        ObtenerHayCambiosValidacion = "V1"
        Exit Function
    End If
    
    comentarios = Nz(cicloActual.comentarios, "")
    If InStr(1, comentarios, "Importación Word", vbTextCompare) > 0 _
       Or InStr(1, comentarios, "Rechazo", vbTextCompare) > 0 _
       Or InStr(1, comentarios, "Revisión", vbTextCompare) > 0 Then
        ObtenerHayCambiosValidacion = "SI"
        Exit Function
    End If
    
    hashEnTbRevision = cicloActual.hashDatos
    
    If Len(hashEnTbRevision) = 0 Then
        ObtenerHayCambiosValidacion = "V1"
    Else
        hashEnTbSolicitud = SnapshotHelper.CalcularHashFaseValidacion(idSolicitud, sol.tipoSolicitud, db)
        
        If StrComp(hashEnTbSolicitud, hashEnTbRevision, vbBinaryCompare) = 0 Then
            ObtenerHayCambiosValidacion = "NO"
        Else
            Call ValidacionRevisionRepositorio.ActualizarHashRevision(idSolicitud, cicloActual.ordinal, hashEnTbSolicitud, db)
            ObtenerHayCambiosValidacion = "SI"
        End If
    End If
    
    Exit Function
Errores:
    ObtenerHayCambiosValidacion = "ERROR"
End Function

Public Sub RegistrarBorradorRAC(ByVal idSolicitud As Long, ByVal idAdjunto As Long, Optional ByVal hashDatos As String = "", Optional ByVal hashActual As String = "")
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then Err.Raise 513, "RegistrarBorradorRAC", "ID solicitud no valido."
    If idAdjunto <= 0 Then Err.Raise 513, "RegistrarBorradorRAC", "ID adjunto no valido."
    
    Call ValidacionRevisionRepositorio.GuardarHitoBorrador(idSolicitud, idAdjunto, m_ObjUsuarioActivo.nombre, hashDatos)
    
LimpiarYSalir:
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.RegistrarBorradorRAC"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud
    errObj.Raise
End Sub

Public Function TieneAdjuntoEnCicloActual(ByVal idSolicitud As Long) As Boolean
    Dim cicloActual As ValidacionRevision
    
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then
        TieneAdjuntoEnCicloActual = False
        Exit Function
    End If
    
    Set cicloActual = ValidacionRevisionRepositorio.GetCicloAbierto(idSolicitud)
    If cicloActual Is Nothing Then
        TieneAdjuntoEnCicloActual = False
    Else
        TieneAdjuntoEnCicloActual = (cicloActual.idAdjunto > 0)
    End If
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.TieneAdjuntoEnCicloActual"
    errObj.AddToCallStack "idSolicitud=" & idSolicitud
    errObj.Raise
End Function
' --- EN: WorkflowServicio.cls ---

Public Function ExportarBorradorSingleton(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As String
    Dim docServ As New DocumentoServicio
    Dim adjServ As New AdjuntosServicio
    Dim ws As DAO.Workspace, dbLocal As DAO.Database
    Dim rutaGenerada As String
    Dim ciclo As ValidacionRevision
    Dim destinoDir As String
    Dim destinoFull As String
    Dim adj As Adjunto
    Dim idAdjNuevo As Long
    Dim etapaValidacionNombre As String
    Dim fso As Object
    Dim ordinalCiclo As Long
    Dim descBorrador As String
    Dim errNum As Long
    Dim errDesc As String
    Dim solServ As New SolicitudServicio
    Dim sol As Solicitud
    Dim hashActual As String
    Dim hashAnterior As String
    Dim tipoSolicitud As String
    Dim dbCheck As DAO.Database
    Dim cicloCheck As ValidacionRevision
    Dim ordinalCheck As Long
    Dim hashActualCheck As String
    Dim hashAnteriorCheck As String
    Dim ultimoOrdinal As Long
    
    On Error GoTo Errores
    
    If db Is Nothing Then Set db = getdb()
    
    ' --- SPEC-133: Verificación canónica de cambios ---
    Dim hayCambiosValidacion As String
    hayCambiosValidacion = ObtenerHayCambiosValidacion(idSolicitud, db)
    
    If hayCambiosValidacion = "SI" Then
    ElseIf hayCambiosValidacion = "NOCICLO" Or hayCambiosValidacion = "V1" Then
    End If
    ' --- FIN SPEC-133 ---
    
    If m_ObjEntorno Is Nothing Then Set m_ObjEntorno = New entorno
    Set fso = CreateObject("Scripting.FileSystemObject")
    etapaValidacionNombre = GetNombreEstadoSafe(estadoValidacion)
    
    Set sol = solServ.getSolicitudPorID(idSolicitud)
    If sol Is Nothing Then
        Err.Raise 513, "ExportarBorradorSingleton", "No se encontró la solicitud con ID: " & idSolicitud
    End If
    tipoSolicitud = sol.tipoSolicitud
    If tipoSolicitud = "" Or IsNull(tipoSolicitud) Then
        Err.Raise 513, "ExportarBorradorSingleton", "La solicitud ID=" & idSolicitud & " no tiene tipo de solicitud definido"
    End If
    
    ' Calcular hash para guardar (Spec-105) - usa el hash simple de validación
    hashActual = SnapshotHelper.CalcularHashFaseValidacion(idSolicitud, tipoSolicitud, db)
    
    ' --- INICIO BLOQUEO SPEC-105 (Validación Previa) ---
    ' Si Spec-133 ya determinó que hay cambios (Importación Word), omitir verificación de hash
    ' Si es V1 o NOCICLO, también omitir verificación (primera versión no tiene hash anterior)
    ' Si hay ERROR, permitir por seguridad (fallback)
    If hayCambiosValidacion = "SI" Or hayCambiosValidacion = "V1" Or hayCambiosValidacion = "NOCICLO" Or hayCambiosValidacion = "ERROR" Then
    ElseIf hayCambiosValidacion = "NO" Then
        ' Comparar hash calculado vs HashDatos del ciclo actual
        Set dbCheck = db
        Set cicloCheck = ValidacionRevisionRepositorio.GetCicloAbierto(idSolicitud, dbCheck)
        ultimoOrdinal = ValidacionRevisionRepositorio.GetUltimoOrdinal(idSolicitud, dbCheck)
        
        If Not cicloCheck Is Nothing Then
            ordinalCheck = cicloCheck.ordinal
        Else
            ordinalCheck = ultimoOrdinal + 1
        End If
        
        If ordinalCheck > 1 Then
            hashActualCheck = hashActual
            hashAnteriorCheck = ValidacionRevisionRepositorio.GetHashVersionAnterior(idSolicitud, ordinalCheck, dbCheck)
            If StrComp(hashActualCheck, hashAnteriorCheck, vbBinaryCompare) = 0 And Len(hashActualCheck) > 0 Then
                ExportarBorradorSingleton = ""
                Exit Function
            End If
        End If
    End If
    ' --- FIN BLOQUEO SPEC-105 ---
    
    If docServ Is Nothing Then Set docServ = New DocumentoServicio
    On Error Resume Next
    rutaGenerada = docServ.GenerarDocumentoParaSolicitud(idSolicitud, 0, db)
    errNum = Err.Number
    errDesc = Err.description
    On Error GoTo Errores
    If errNum <> 0 Then
        Err.Raise 514, "ExportarBorradorSingleton", "Error al generar documento para solicitud " & idSolicitud & ": " & errDesc & " (Código: " & errNum & ")"
    End If
    If Len(Trim(rutaGenerada)) = 0 Then
        Err.Raise 514, "ExportarBorradorSingleton", "GenerarDocumentoParaSolicitud retornó vacío para ID=" & idSolicitud
    End If
    If Not fso.FileExists(rutaGenerada) Then
        Err.Raise 514, "ExportarBorradorSingleton", "El documento no existe en disco. Ruta: " & rutaGenerada
    End If
    
    Set ws = getWorkspace()
    Set dbLocal = ws.OpenDatabase(db.name, False, False, "MS Access;PWD=" & GetPasswordDB())
    ws.BeginTrans
    
    Set ciclo = ValidacionRevisionRepositorio.GetCicloAbierto(idSolicitud, dbLocal)
    
    ' --- SPEC-142/143: Solo sobrescribir archivo físico si ya existe ---
    ' Si el ciclo ya tiene idAdjunto > 0, solo sobreescribir archivo físico
    If Not ciclo Is Nothing Then
        If ciclo.idAdjunto > 0 Then
            On Error Resume Next
            Call adjServ.ActualizarFicheroAdjunto(ciclo.idAdjunto, rutaGenerada, dbLocal)
            errNum = Err.Number
            errDesc = Err.description
            On Error GoTo Errores
            
            If errNum = 0 Then
                On Error Resume Next
                Call ValidacionRevisionRepositorio.GuardarHitoBorrador(idSolicitud, ciclo.idAdjunto, m_ObjUsuarioActivo.nombre, hashActual, dbLocal)
                errNum = Err.Number
                errDesc = Err.description
                On Error GoTo Errores
                
                Set adj = AdjuntoRepositorio.getPorID(ciclo.idAdjunto, dbLocal)
                GoTo FinalizarRegistro
            End If
        End If
    End If
    ' --- FIN SPEC-142/143 ---
    
    ' Si no existe ciclo con idAdjunto, crear nuevo
    GoTo CrearNuevoAdjunto

CrearNuevoAdjunto:
    ' Usar ordinal del ciclo si existe, sino GetUltimoOrdinal + 1
    If Not ciclo Is Nothing Then
        ordinalCiclo = ciclo.ordinal
    Else
        ordinalCiclo = ValidacionRevisionRepositorio.GetUltimoOrdinal(idSolicitud, dbLocal) + 1
    End If
    
    descBorrador = "Borrador v" & ordinalCiclo
    
    On Error Resume Next
    Call adjServ.GuardarAdjuntoDesdeArchivo(idSolicitud, etapaValidacionNombre, rutaGenerada, "Generado", dbLocal, idAdjNuevo, descBorrador)
    errNum = Err.Number
    errDesc = Err.description
    On Error GoTo Errores
    If errNum <> 0 Then
        Err.Raise 515, "ExportarBorradorSingleton", "Error al guardar adjunto del borrador: " & errDesc
    End If
    If idAdjNuevo <= 0 Then
        Err.Raise 515, "ExportarBorradorSingleton", "El guardado del adjunto no retornó un ID válido. idAdjNuevo=" & idAdjNuevo
    End If
    Call RegistrarBorradorRAC(idSolicitud, idAdjNuevo, hashActual, hashActual)
    
    On Error Resume Next
    Set adj = AdjuntoRepositorio.getPorID(idAdjNuevo, dbLocal)
    errNum = Err.Number
    errDesc = Err.description
    On Error GoTo Errores
    If errNum <> 0 Then
        Err.Raise 515, "ExportarBorradorSingleton", "Error al recuperar el adjunto ID=" & idAdjNuevo & " tras guardarlo: " & errDesc
    End If
    If adj Is Nothing Then
        Err.Raise 515, "ExportarBorradorSingleton", "Se guardó el adjunto ID=" & idAdjNuevo & " pero al recuperarlo retorna Nothing"
    End If

FinalizarRegistro:
    ' 3. Calcular ruta final para el retorno de la función
    If Not adj Is Nothing Then
        destinoDir = m_ObjEntorno.URLDirectorioDocumentacion
        If Right$(destinoDir, 1) <> "\" Then destinoDir = destinoDir & "\"
        destinoFull = destinoDir & adj.nombreArchivo
        ExportarBorradorSingleton = destinoFull
    End If
    
    ws.CommitTrans
    If Not dbLocal Is Nothing Then dbLocal.Close: Set dbLocal = Nothing
    Exit Function
    
Errores:
    If Not ws Is Nothing Then
        ws.Rollback
    End If
    If Not dbLocal Is Nothing Then dbLocal.Close: Set dbLocal = Nothing
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.ExportarBorradorSingleton"
    errObj.Raise
End Function

Public Function getProximaTransicionValida(ByVal idSolicitud As Long, ByVal rolUsuarioActual As rol) As Scripting.Dictionary
    ' RESPONSABILIDAD: Devuelve la siguiente transición de estado "NATURAL" (Camino Feliz).
    '                  Ignora transiciones de rechazo o retroceso para la sugerencia automática.
    
    Dim colTransiciones As Scripting.Dictionary
    Dim idEstadoDestino As Variant
    Dim estadoDestino As estado
    Dim solServ As New SolicitudServicio
    Dim sol As Solicitud
    
    On Error GoTo Errores
    
    ' 1. Obtenemos todas las transiciones válidas
    Set colTransiciones = Me.getTransicionesValidas(idSolicitud, rolUsuarioActual)
    Set sol = solServ.getSolicitudPorID(idSolicitud)
    
    If colTransiciones Is Nothing Then GoTo SalidaSinPropuesta
    If colTransiciones.count = 0 Then GoTo SalidaSinPropuesta

    ' 2. FILTRADO INTELIGENTE: Buscamos el único camino de AVANCE positivo
    Dim candidatos As New Collection
    
    For Each idEstadoDestino In colTransiciones.Keys
        Set estadoDestino = m_ObjEntorno.estados(CStr(idEstadoDestino))
        
        ' Regla A: El destino debe tener un orden SUPERIOR al actual (Avance)
        ' Regla B: El destino NO puede ser "RECHAZADA" (Estado 9), salvo que sea la única opción lógica
        ' Regla C: El destino NO puede ser "DESARROLLO TÉCNICO" si venimos de adelante (eso es retroceso)
        
        Dim esAvance As Boolean
        Dim estadoActualOrden As Long
        estadoActualOrden = m_ObjEntorno.estados(CStr(sol.idEstadoInterno)).orden
        
        esAvance = (estadoDestino.orden > estadoActualOrden)
        
        If esAvance And estadoDestino.idEstado <> estadoRechazada Then
            candidatos.Add estadoDestino
        End If
    Next idEstadoDestino
    
    ' 3. Evaluación de Candidatos
    If candidatos.count = 1 Then
        ' ¡BINGO! Tenemos un único camino feliz claro.
        Set estadoDestino = candidatos(1)
        
        Set getProximaTransicionValida = New Scripting.Dictionary
        getProximaTransicionValida.Add "id", estadoDestino.idEstado
        getProximaTransicionValida.Add "nombre", GetNombreEstadoSafe(estadoDestino.idEstado)
        Exit Function
    Else
        ' Si hay 0 candidatos (estamos en fin de flujo o solo hay rechazo)
        ' O si hay > 1 candidatos (bifurcación real), no sugerimos nada.
        GoTo SalidaSinPropuesta
    End If

SalidaSinPropuesta:
    Set getProximaTransicionValida = Nothing
    Exit Function

Errores:
    Set getProximaTransicionValida = Nothing
    Resume LimpiarYSalir
LimpiarYSalir:
End Function

Public Sub RevertirFlujo(ByVal sol As Solicitud, ByVal idEstadoDestino As Long, ByVal usuario As usuario, ByVal motivo As String, Optional ByRef db As DAO.Database)
    ' RESPONSABILIDAD: Retroceder el estado de la solicitud en el flujo normal (SIN PURGA).
    Dim ws As DAO.Workspace
    Dim dbLocal As DAO.Database
    Dim logServ As New LogServicio
    Dim estadoAnteriorID As Long
    Dim dbOrigIsNothing As Boolean
    
    On Error GoTo Errores
    
    dbOrigIsNothing = (db Is Nothing)
    If dbOrigIsNothing Then Set db = getdb()
    
    estadoAnteriorID = sol.idEstadoInterno
    
    If dbOrigIsNothing Then
        Set ws = getWorkspace()
        Set dbLocal = db  ' Use db directly; ws is used only for BeginTrans/CommitTrans/Rollback
        ws.BeginTrans
    Else
        Set dbLocal = db
    End If
    
    ' Crear ciclo de validación al revertir a Validación
    If idEstadoDestino = estadoValidacion Then
        Dim hashParaNuevoCiclo As String
        Dim comentarioNuevoCiclo As String
        hashParaNuevoCiclo = ""
        comentarioNuevoCiclo = ""
        
        ' Si viene de REVISIÓN, verificar si hubo Importación Word
        If estadoAnteriorID = estadoRevision Then
            Dim cicloAnterior As ValidacionRevision
            Set cicloAnterior = ValidacionRevisionRepositorio.GetUltimoCicloCerrado(sol.idSolicitud, dbLocal)
            If Not cicloAnterior Is Nothing Then
                If InStr(1, cicloAnterior.comentarios, "Importación Word", vbTextCompare) > 0 Then
                    hashParaNuevoCiclo = ""
                    comentarioNuevoCiclo = cicloAnterior.comentarios
                Else
                    Dim ordinalNuevoCiclo As Long
                    ordinalNuevoCiclo = ValidacionRevisionRepositorio.GetUltimoOrdinal(sol.idSolicitud, dbLocal) + 1
                    Call SnapshotHelper.ExportarSnapshotParaDebug(sol.idSolicitud, sol.tipoSolicitud, "crear_ciclo", ordinalNuevoCiclo, db)
                    hashParaNuevoCiclo = SnapshotHelper.CalcularHashFaseValidacion(sol.idSolicitud, sol.tipoSolicitud, db)
                End If
            End If
        End If
        
        Call ValidacionRevisionRepositorio.CrearNuevoCicloValidacion(sol.idSolicitud, usuario.nombre, hashParaNuevoCiclo, dbLocal, comentarioNuevoCiclo)
    End If
    
    ' 1. Actualizar Estado
    sol.idEstadoInterno = idEstadoDestino
    sol.fechaModificacion = Now()
    sol.usuarioModificacion = usuario.nombre
    Call SolicitudRepositorio.ActualizarEstado(sol, dbLocal)
    
    ' 2. Log de Estados
    Dim logEst As New LogEstado
    logEst.idSolicitud = sol.idSolicitud
    logEst.idEstadoAnterior = estadoAnteriorID
    logEst.idEstadoNuevo = idEstadoDestino
    logEst.fechaTransicion = sol.fechaModificacion
    logEst.usuarioTransicion = sol.usuarioModificacion
    Call LogEstadoRepositorio.Guardar(logEst, dbLocal)
    
    ' 3. Log de Cambios
    Call logServ.RegistrarCambio("tbSolicitudes", sol.idSolicitud, "idEstadoInterno", _
        CStr(estadoAnteriorID), CStr(idEstadoDestino), "FLUJO: " & motivo, dbLocal)
        
    If dbOrigIsNothing Then ws.CommitTrans
    
    If idEstadoDestino = estadoValidacion And estadoAnteriorID = estadoRevision Then
        Call NotificarTransicion(sol, idEstadoDestino, usuario, estadoAnteriorID)
    End If
    
    Exit Sub
Errores:
    If dbOrigIsNothing And Not ws Is Nothing Then ws.Rollback
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.RevertirFlujo"
    errObj.Raise
End Sub

Public Sub RevertirAFaseAnterior_FromUI(ByVal sol As Solicitud, ByVal idEstadoDestino As Long, ByVal usuario As usuario, ByVal motivo As String)
    Dim ws As DAO.Workspace
    Dim db As DAO.Database

    On Error GoTo Errores

    Set ws = getWorkspace()
    Set db = getdb()
    ws.BeginTrans

    Call Me.RevertirAFaseAnterior(sol, idEstadoDestino, usuario, motivo, db)

    ws.CommitTrans

LimpiarYSalir:
    On Error Resume Next
    If Not db Is Nothing Then db.Close
    Set db = Nothing
    Set ws = Nothing
    Exit Sub

Errores:
    If Not ws Is Nothing Then ws.Rollback
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.RevertirAFaseAnterior_FromUI"
    errObj.AddToCallStack "Solicitud: " & sol.codigoSolicitud
    errObj.Raise
End Sub

Public Sub RevertirAFaseAnterior(ByVal sol As Solicitud, ByVal idEstadoDestino As Long, ByVal usuario As usuario, ByVal motivo As String, ByRef db As DAO.Database)
    Dim estadoAnteriorID As Long
    Dim logServ As New LogServicio
    Dim pcServ As New DatosPCServicio
    Dim cdcaServ As New DatosCDCAServicio
    Dim cdcaSubServ As New DatosCDCASUBServicio
    Dim pcsubServ As New DatosPCSUBServicio
    Dim adjServ As New AdjuntosServicio
    Dim rechazoServ As New RechazoServicio
    Dim logCambioServ As New LogCambioServicio
    Dim detalleReset As String
    
    Dim dbLocal As DAO.Database
    Dim requiereGuardarSolicitud As Boolean
    
    On Error GoTo Errores
    
    Set dbLocal = db
    
    ' 1. VALIDACIÓN PREVIA
    If usuario.rol <> rol.Administrador And usuario.rol <> rol.Calidad Then
        Err.Raise 513, "RevertirAFaseAnterior", "No tiene los permisos necesarios (Admin/Calidad) para revertir estados."
    End If
    
    estadoAnteriorID = sol.idEstadoInterno
    
    ' 2. USE INJECTED TRANSACTION CONTEXT
    
    ' 3. LIMPIEZA DE DATOS (Pasando 'dbLocal' para atomicidad)
    Select Case estadoAnteriorID
        
        Case estadoPreregistro ' 1 -> ?
             ' No hay limpieza definida
             
        Case estadoRegistro ' 2 -> 1
             ' No hay limpieza definida
             
        Case estadoDesarrolloTecnico ' 3 -> 2
             ' RESET FASE TÉCNICA usando los servicios correspondientes (que delegan en repositorios)
             If sol.tipoSolicitud = "PC" Then
                 Call pcServ.PurgaTecnica(sol.idSolicitud, dbLocal)
             ElseIf sol.tipoSolicitud = "CD_CA" Then
                 Call cdcaServ.PurgaTecnica(sol.idSolicitud, dbLocal)
              ElseIf sol.tipoSolicitud = "CD_CA_SUB" Then
                  Call cdcaSubServ.PurgaTecnica(sol.idSolicitud, dbLocal)
              ElseIf sol.tipoSolicitud = "PC_SUB" Then
                  Call pcsubServ.PurgaTecnica(sol.idSolicitud, dbLocal)
              End If
             
             ' Limpieza de adjuntos y rastro de revisiones
             Call adjServ.EliminarAdjuntosPorEtapas(sol.idSolicitud, Array("Propuesta", "Impacto", "Tecnica"), dbLocal)
             Call rechazoServ.LimpiarHistorialRechazos(sol.idSolicitud, dbLocal)
             Call logCambioServ.LimpiarHistorialCambios(sol.idSolicitud, dbLocal)
             sol.revisionCalidadEstado = "PENDIENTE"
             sol.revisionCalidadComentarios = ""
             requiereGuardarSolicitud = True
             detalleReset = "Purga tecnica aplicada" & vbCrLf & "Adjuntos Propuesta/Impacto/Tecnica eliminados" & vbCrLf & "Historial rechazos eliminado" & vbCrLf & "Historial cambios eliminado" & vbCrLf & "Revision calidad reiniciada"

        Case estadoRechazada ' 9 -> 6
            ' Aquí limpiamos el RAC porque el rechazo se escribió ahí
            If sol.tipoSolicitud = "PC" Then
                Call pcServ.EliminarDictamenRAC(sol.idSolicitud, dbLocal)
            ElseIf sol.tipoSolicitud = "CD_CA" Then
                Call cdcaServ.EliminarDictamenRAC(sol.idSolicitud, dbLocal)
            ElseIf sol.tipoSolicitud = "CD_CA_SUB" Then
                Call cdcaSubServ.EliminarDictamenRAC(sol.idSolicitud, dbLocal)
            ElseIf sol.tipoSolicitud = "PC_SUB" Then
                Call pcsubServ.EliminarDictamenRAC(sol.idSolicitud, dbLocal)
            End If
            detalleReset = "Dictamen RAC eliminado"
            
        Case estadoFormalizacion ' 7 -> 6
            Dim etapaFormalizacionNombre As String
            Dim etapaFormalizacionSafe As String
            etapaFormalizacionNombre = ""
            If m_ObjEntorno.estados.Exists(CStr(estadoFormalizacion)) Then
                etapaFormalizacionNombre = GetNombreEstadoSafe(estadoFormalizacion)
            End If
            etapaFormalizacionSafe = GetNombreEstadoSafe(estadoFormalizacion)
            
            Call adjServ.EliminarAdjuntosPorEtapas(sol.idSolicitud, _
                Array( _
                    "Cierre", _
                    "Documento Final Firmado", _
                    "PDF Firmado", _
                    "Formalización", _
                    "Formalizacion", _
                    etapaFormalizacionNombre, _
                    etapaFormalizacionSafe, _
                    CStr(estadoFormalizacion) _
                ), _
                dbLocal)
            
            ' NOTA: NO se eliminan DecisionFinal ni AprobacionSuministrador
            ' porque ya existían en la fase de Revisión (Spec-170)
            detalleReset = "Adjuntos de formalizacion eliminados"
            
        Case estadoAprobada ' 8 -> 7
            Dim colAdjFinal As Collection
            Dim adjFinal As Adjunto
            Dim etapaKeyFinal As String
            
            Set colAdjFinal = adjServ.getAdjuntosPorSolicitudTx(sol.idSolicitud, dbLocal)
            If Not colAdjFinal Is Nothing Then
                For Each adjFinal In colAdjFinal
                    etapaKeyFinal = CStr(adjFinal.etapaWF)
                    If StrComp(etapaKeyFinal, "Cierre", vbTextCompare) = 0 _
                       Or StrComp(etapaKeyFinal, "Documento Final Firmado", vbTextCompare) = 0 _
                       Or StrComp(etapaKeyFinal, "PDF Firmado", vbTextCompare) = 0 _
                       Or StrComp(etapaKeyFinal, CStr(estadoFormalizacion), vbTextCompare) = 0 _
                       Or StrComp(etapaKeyFinal, CStr(estadoAprobada), vbTextCompare) = 0 _
                       Or UCase$(adjFinal.tipoAccion) = "CIERRE_OFICIAL" Then
                        Call adjServ.EliminarAdjunto(adjFinal.idAdjunto, dbLocal)
                    End If
                Next adjFinal
            End If
            
            If sol.tipoSolicitud = "PC" Then Call pcServ.EliminarDecisionFinal(sol.idSolicitud, dbLocal)
            If sol.tipoSolicitud = "CD_CA" Then Call cdcaServ.EliminarDecisionFinal(sol.idSolicitud, dbLocal)
            If sol.tipoSolicitud = "CD_CA_SUB" Then Call cdcaSubServ.EliminarDecisionFinal(sol.idSolicitud, dbLocal)
            If sol.tipoSolicitud = "PC_SUB" Then Call pcsubServ.EliminarDecisionFinal(sol.idSolicitud, dbLocal)
            detalleReset = "Adjuntos finales eliminados" & vbCrLf & "Decision final eliminada"
            
        Case estadoRevision ' 6 -> 5
             Dim etapaValidacionNombre As String
             Dim etapaRevisionNombre As String
             
             etapaValidacionNombre = GetNombreEstadoSafe(estadoValidacion)
             etapaRevisionNombre = GetNombreEstadoSafe(estadoRevision)
             
             Call ValidacionRevisionRepositorio.EliminarPorSolicitud(sol.idSolicitud, dbLocal)
             Call adjServ.EliminarAdjuntosPorEtapas(sol.idSolicitud, Array(CStr(estadoValidacion), CStr(estadoRevision), etapaValidacionNombre, etapaRevisionNombre), dbLocal)
             
              If sol.tipoSolicitud = "PC" Then Call pcServ.EliminarDictamenRAC(sol.idSolicitud, dbLocal)
              If sol.tipoSolicitud = "CD_CA" Then Call cdcaServ.EliminarDictamenRAC(sol.idSolicitud, dbLocal)
              If sol.tipoSolicitud = "CD_CA_SUB" Then Call cdcaSubServ.EliminarDictamenRAC(sol.idSolicitud, dbLocal)
              If sol.tipoSolicitud = "PC_SUB" Then Call pcsubServ.EliminarDictamenRAC(sol.idSolicitud, dbLocal)
              detalleReset = "Validacion/Revision limpiadas" & vbCrLf & "Adjuntos validacion/revision eliminados" & vbCrLf & "Dictamen RAC eliminado"

        Case estadoModificacion ' 4 -> 3
             Call RechazoRepositorio.EliminarRechazosPorSolicitud(sol.idSolicitud, dbLocal)
             
             If sol.tipoSolicitud = "PC" Then
                 Call pcServ.EliminarAprobacionSuministrador(sol.idSolicitud, dbLocal)
             ElseIf sol.tipoSolicitud = "CD_CA" Then
                 Call cdcaServ.EliminarAprobacionSuministrador(sol.idSolicitud, dbLocal)
              ElseIf sol.tipoSolicitud = "CD_CA_SUB" Then
                  Call cdcaSubServ.EliminarAprobacionSuministrador(sol.idSolicitud, dbLocal)
              ElseIf sol.tipoSolicitud = "PC_SUB" Then
                  Call pcsubServ.EliminarAprobacionSuministrador(sol.idSolicitud, dbLocal)
              End If
             
             sol.revisionCalidadEstado = "PENDIENTE"
             sol.revisionCalidadComentarios = ""
             requiereGuardarSolicitud = True
             detalleReset = "Rechazos eliminados" & vbCrLf & "Aprobacion Suministrador eliminada" & vbCrLf & "Revision calidad reiniciada"
             
        Case estadoValidacion ' 5 -> 4
             Dim etapaValidacionNombre5 As String
             Dim etapaRevisionNombre5 As String

             etapaValidacionNombre5 = GetNombreEstadoSafe(estadoValidacion)
             etapaRevisionNombre5 = GetNombreEstadoSafe(estadoRevision)

             Call ValidacionRevisionRepositorio.EliminarPorSolicitud(sol.idSolicitud, dbLocal)
             Call adjServ.EliminarAdjuntosPorEtapas(sol.idSolicitud, Array(CStr(estadoValidacion), CStr(estadoRevision), etapaValidacionNombre5, etapaRevisionNombre5), dbLocal)
             
             If sol.tipoSolicitud = "PC" Then
                 Call pcServ.EliminarDictamenRAC(sol.idSolicitud, dbLocal)
                 Call pcServ.EliminarAprobacionSuministrador(sol.idSolicitud, dbLocal)
             ElseIf sol.tipoSolicitud = "CD_CA" Then
                 Call cdcaServ.EliminarDictamenRAC(sol.idSolicitud, dbLocal)
                 Call cdcaServ.EliminarAprobacionSuministrador(sol.idSolicitud, dbLocal)
              ElseIf sol.tipoSolicitud = "CD_CA_SUB" Then
                  Call cdcaSubServ.EliminarDictamenRAC(sol.idSolicitud, dbLocal)
                  Call cdcaSubServ.EliminarAprobacionSuministrador(sol.idSolicitud, dbLocal)
              ElseIf sol.tipoSolicitud = "PC_SUB" Then
                  Call pcsubServ.EliminarDictamenRAC(sol.idSolicitud, dbLocal)
                  Call pcsubServ.EliminarAprobacionSuministrador(sol.idSolicitud, dbLocal)
              End If

             sol.revisionCalidadEstado = "PENDIENTE"
             requiereGuardarSolicitud = True
              detalleReset = "Validacion/Revision limpiadas" & vbCrLf & "Adjuntos validacion/revision eliminados" & vbCrLf & "Revision calidad reiniciada" & vbCrLf & "Aprobacion Suministrador eliminada" & vbCrLf & "Dictamen RAC eliminado"
              
    End Select
    
    ' 4. ACTUALIZAR ESTADO
    sol.idEstadoInterno = idEstadoDestino
    sol.fechaModificacion = Now()
    sol.usuarioModificacion = usuario.nombre
    
    If requiereGuardarSolicitud Then
        Call SolicitudRepositorio.GuardarSolicitud(sol, dbLocal)
    Else
        Call SolicitudRepositorio.ActualizarEstado(sol, dbLocal)
    End If
    
    ' 5. LOG DE ESTADOS
    Dim logEst As New LogEstado
    With logEst
        .idSolicitud = sol.idSolicitud
        .idEstadoAnterior = estadoAnteriorID
        .idEstadoNuevo = idEstadoDestino
        .fechaTransicion = sol.fechaModificacion
        .usuarioTransicion = sol.usuarioModificacion
    End With
    Call LogEstadoRepositorio.Guardar(logEst, dbLocal)
    
    ' 6. LOG DE CAMBIOS
    Dim estadoAnteriorNombre As String, estadoNuevoNombre As String
    estadoAnteriorNombre = GetNombreEstadoSafe(estadoAnteriorID)
    estadoNuevoNombre = GetNombreEstadoSafe(idEstadoDestino)
    
    Call logServ.RegistrarCambio("tbSolicitudes", sol.idSolicitud, "idEstadoInterno", _
        estadoAnteriorID & " (" & estadoAnteriorNombre & ")", _
        idEstadoDestino & " (" & estadoNuevoNombre & ")", _
        "REVERSION: " & motivo, dbLocal)
         
    ' 7. TRANSACTION OWNERSHIP REMAINS WITH THE CALLER
    
    On Error Resume Next
    Dim notifServ As New NotificacionServicio
    Dim solServ As New SolicitudServicio
    Dim vm As SolicitudViewModel
    Dim destinatario As String
    Dim copia As String
    Dim asunto As String
    Dim extraHtml As String
    Dim nombreEstadoReset As String
    Set vm = solServ.getSolicitudViewModelPorID(sol.idSolicitud)
    If idEstadoDestino = estadoDesarrolloTecnico Then
        destinatario = vm.Expediente.EmailResponsableTecnico
    Else
        destinatario = vm.Expediente.EmailResponsableCalidad
    End If
    copia = m_ObjUsuarioReal.CorreoUsuario
    nombreEstadoReset = UCase$(GetNombreEstadoSafe(idEstadoDestino))
    ' SPEC-100: Estandarización de asunto (NOMBRE_FASE + Contexto)
    asunto = "CONDOR - " & nombreEstadoReset & " (RESET FASE) | " & vm.Expediente.Nemotecnico & " | " & vm.TipoSolicitudNombre & " | " & vm.Solicitud.codigoSolicitud
    
    ' SPEC-100: Tarjeta contextualizada
    extraHtml = "<div class='card'><div class='card-header'>AVISO DE REVERSIÓN</div><div class='card-body'><strong>Detalle purga:</strong><br>" & Replace(HtmlEncode(Nz(detalleReset, "")), vbCrLf, "<br>") & "</div></div>"
    If Trim$(Nz(destinatario, "")) <> "" Then
        Call notifServ.EnviarNotificacionConDetalle(vm, asunto, destinatario, copia, extraHtml)
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
        
    Exit Sub  ' Notificación completada sin errores - salir aquí
        
LimpiarYSalir:
    On Error Resume Next
    Set logServ = Nothing
    Set pcServ = Nothing
    Set cdcaServ = Nothing
    Set cdcaSubServ = Nothing
    Set adjServ = Nothing
    Set rechazoServ = Nothing
    Set logCambioServ = Nothing
    Exit Sub
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.RevertirAFaseAnterior"
    errObj.AddToCallStack "Solicitud: " & sol.codigoSolicitud
    errObj.Raise
End Sub
' ==========================================================================
' MÓDULO: WorkflowServicio.cls
' ACCIÓN: Mostrar aviso visual de falta de PDF Firmado en el HTML.
' ==========================================================================
Private Function GetPrecondicionesPendientes(ByRef sol As Solicitud, ByVal idEstadoDestino As Long, ByVal rolUsuarioActual As rol, Optional ByRef db As DAO.Database) As Collection
    Dim col As New Collection
    Dim rolRequerido As String
    Dim pcServ As New DatosPCServicio
    Dim cdcaServ As New DatosCDCAServicio
    Dim cdcaSubServ As New DatosCDCASUBServicio
    Dim pcsubServ As New DatosPCSUBServicio
    Dim adjServ As New AdjuntosServicio ' <<< Necesario
    
    On Error GoTo Fin
    
    If db Is Nothing Then Set db = getdb()
    
    ' 1. Rol (sin cambios)
    rolRequerido = WorkflowRepositorio.getPosiblesDestinos(sol.idEstadoInterno)(CStr(idEstadoDestino))
    If Not PermisoSuficiente(rolUsuarioActual, rolRequerido) Then
        col.Add "Tener el rol: <strong>" & rolRequerido & "</strong>"
    End If

    ' 2. Datos y Adjuntos
    Select Case idEstadoDestino
        Case estadoRegistro
            If sol.tipoSolicitud = "PC" Then
                If Not pcServ.EsDatosGeneralesCompleta(sol.idSolicitud, db) Then col.Add "Completar y guardará la pestaña 'General'."
            ElseIf sol.tipoSolicitud = "CD_CA" Then
                If Not cdcaServ.EsDatosGeneralesCompleta(sol.idSolicitud, db) Then col.Add "Completar y guardará la pestaña 'General'."
            ElseIf sol.tipoSolicitud = "CD_CA_SUB" Then
                If Not cdcaSubServ.EsDatosGeneralesCompleta(sol.idSolicitud, db) Then col.Add "Completar y guardará la pestaña 'General'."
            ElseIf sol.tipoSolicitud = "PC_SUB" Then
                If Not pcsubServ.EsDatosGeneralesCompleta(sol.idSolicitud, db) Then col.Add "Completar y guardará la pestaña 'General'."
            End If
            
        Case estadoModificacion, estadoValidacion
             ' (Lógica existente para estos estados...)
             Dim parteTecnicaCompleta As Boolean
             If sol.tipoSolicitud = "PC" Then parteTecnicaCompleta = pcServ.EsParteTecnicaCompleta(sol.idSolicitud, db)
             If sol.tipoSolicitud = "CD_CA" Then parteTecnicaCompleta = cdcaServ.EsParteTecnicaCompleta(sol.idSolicitud, db)
            If sol.tipoSolicitud = "CD_CA_SUB" Then parteTecnicaCompleta = cdcaSubServ.EsParteTecnicaCompleta(sol.idSolicitud, db)
            If sol.tipoSolicitud = "PC_SUB" Then parteTecnicaCompleta = pcsubServ.EsParteTecnicaCompleta(sol.idSolicitud, db)
            If Not parteTecnicaCompleta Then col.Add "Datos de ingeniería aún sin cumplimentar."

         Case estadoFormalizacion
              ' (Lógica existente...)
If sol.tipoSolicitud = "PC" Then
                  If Not pcServ.EsDictamenRACCompleto(sol.idSolicitud, db) Then col.Add "Completar y guardará la pestaña 'RAC'."
                  If Not pcServ.EsAprobacionSuministradorCompleta(sol.idSolicitud, db) Then col.Add "Completar y guardará la pestaña 'Aprobación Suministrador'."
               ElseIf sol.tipoSolicitud = "CD_CA" Then
                  If Not cdcaServ.EsDictamenRACCompleto(sol.idSolicitud, db) Then col.Add "Completar y guardará la pestaña 'RAC'."
                  If Not cdcaServ.EsAprobacionSuministradorCompleta(sol.idSolicitud, db) Then col.Add "Completar y guardará la pestaña 'Aprobación Suministrador'."
               ElseIf sol.tipoSolicitud = "CD_CA_SUB" Then
                  If Not cdcaSubServ.EsDictamenRACCompleto(sol.idSolicitud, db) Then col.Add "Completar y guardará la pestaña 'RAC'."
                  If Not cdcaSubServ.EsAprobacionSuministradorCompleta(sol.idSolicitud, db) Then col.Add "Completar y guardará la pestaña 'Aprobación Suministrador'."
               ElseIf sol.tipoSolicitud = "PC_SUB" Then
                  If Not pcsubServ.EsDictamenRACCompleto(sol.idSolicitud, db) Then col.Add "Completar y guardará la pestaña 'RAC'."
                  If Not pcsubServ.EsAprobacionSuministradorCompleta(sol.idSolicitud, db) Then col.Add "Completar y guardará la pestaña 'Aprobación Suministrador'."
               End If
        
        Case estadoAprobada
            ' 1. Datos
            If sol.tipoSolicitud = "PC" Then
                If Not pcServ.EsDecisionFinalCompleta(sol.idSolicitud, db) Then col.Add "Completar y guardará la pestaña 'Autoridad Competente'."
            ElseIf sol.tipoSolicitud = "CD_CA" Then
                If Not cdcaServ.EsDecisionFinalCompleta(sol.idSolicitud, db) Then col.Add "Completar y guardará la pestaña 'Autoridad Competente'."
            ElseIf sol.tipoSolicitud = "CD_CA_SUB" Then
                If Not cdcaSubServ.EsDecisionFinalCompleta(sol.idSolicitud, db) Then col.Add "Completar y guardará la pestaña 'Autoridad Competente'."
            ElseIf sol.tipoSolicitud = "PC_SUB" Then
                If Not pcsubServ.EsDecisionFinalCompleta(sol.idSolicitud, db) Then col.Add "Completar y guardará la pestaña 'Autoridad Competente'."
            End If
            
            ' 2. --- NUEVO AVISO VISUAL ---
            Dim etapaFormalizacionNombre As String
            etapaFormalizacionNombre = GetNombreEstadoSafe(estadoFormalizacion)
            If Not adjServ.ExisteAdjuntoEtapa(sol.idSolicitud, etapaFormalizacionNombre, db) Then
                col.Add "Adjuntar el documento de Cierre en la gestion de adjuntos."
            End If
            ' -----------------------------
    End Select
    
Fin:
    Set GetPrecondicionesPendientes = col
End Function

Private Function getProximaTransicionTeorica(ByVal idSolicitud As Long, ByVal rolUsuarioActual As rol) As Scripting.Dictionary
    ' RESPONSABILIDAD: Devuelve el siguiente estado en el flujo, basándose en el
    '                  orden numérico, independientemente de los permisos del usuario.
    '                  Es una función "teórica" para ayudar a construir la UI.
    Dim solServ As New SolicitudServicio
    Dim sol As Solicitud
    Dim estadoActual As estado
    Dim ordenActual As Long
    Dim colPosiblesDestinos As Scripting.Dictionary
    Dim idEstadoDestino As Variant
    Dim estadoDestino As estado
    
    On Error GoTo Errores
    
    ' 1. Obtener el estado y orden actual de la solicitud
    Set sol = solServ.getSolicitudPorID(idSolicitud)
    If sol Is Nothing Then Exit Function
    
    Set estadoActual = m_ObjEntorno.estados(CStr(sol.idEstadoInterno))
    ordenActual = estadoActual.orden
    
    ' 2. Obtener todas las transiciones posibles desde el estado actual desde la BD
    Set colPosiblesDestinos = WorkflowRepositorio.getPosiblesDestinos(sol.idEstadoInterno)
    If colPosiblesDestinos.count = 0 Then Exit Function ' Es un estado final, no hay próximo.
    
    ' 3. Encontrar el estado destino con el siguiente número de orden más bajo
    Dim mejorCandidatoID As Long
    Dim minOrdenSuperior As Long: minOrdenSuperior = 9999 ' Un número muy alto para empezar
    
    For Each idEstadoDestino In colPosiblesDestinos.Keys
        Set estadoDestino = m_ObjEntorno.estados(CStr(idEstadoDestino))
        
        ' Buscamos el estado cuyo orden sea el más cercano por encima del actual
        If estadoDestino.orden > ordenActual And estadoDestino.orden < minOrdenSuperior Then
            minOrdenSuperior = estadoDestino.orden
            mejorCandidatoID = estadoDestino.idEstado
        End If
    Next
    
    ' 4. Si encontramos un candidato, lo empaquíetamos y lo devolvemos
    If mejorCandidatoID > 0 Then
        Dim estadoFinal As estado
        Set estadoFinal = m_ObjEntorno.estados(CStr(mejorCandidatoID))
        
        Set getProximaTransicionTeorica = New Scripting.Dictionary
        getProximaTransicionTeorica.Add "id", estadoFinal.idEstado
        getProximaTransicionTeorica.Add "nombre", GetNombreEstadoSafe(estadoFinal.idEstado)
    Else
        Set getProximaTransicionTeorica = Nothing
    End If
    
LimpiarYSalir:
    Exit Function
Errores:
    ' En caso de error, devolvemos Nothing para que el llamador no fallóe.
    Set getProximaTransicionTeorica = Nothing
    Resume LimpiarYSalir
End Function
' ==========================================================================
' MÓDULO: WorkflowServicio.cls
' ACCIÓN: Modificar GenerarHTML_FichaResumen para mostrar estado de revisión en MODIFICACIÓN.
' ==========================================================================
Public Function GenerarHTML_FichaResumen(ByVal idSolicitud As Long) As String
    Dim solServ As New SolicitudServicio
    Dim vm As SolicitudViewModel
    Dim html As String
    
    On Error GoTo Errores
    
    Set vm = solServ.getSolicitudViewModelPorID(idSolicitud)
    If vm Is Nothing Then
        html = "<html><body>Error: No se encontró la solicitud.</body></html>"
        GoTo Finalizar
    End If
    
    html = "<!DOCTYPE html><html lang='es'><head><meta charset='UTF-8'><style>" & GetFichaCSS() & "</style></head><body>"
    html = html & "<div class='card'>"
    
    html = html & "  <div class='card-header'>Resumen: " & vm.Solicitud.codigoSolicitud & " (" & UCase(vm.EstadoNombre) & ")</div>"
    html = html & "  <div class='card-body'>"
    
    html = html & "    <table class='data-table'>"
    html = html & "      <tr><th>Tipo:</th><td>" & vm.TipoSolicitudNombre & "</td></tr>"
    html = html & "      <tr><th>F. Solicitud:</th><td>" & Format(vm.Solicitud.fechaCreacion, "dd/mm/yyyy") & "</td></tr>"
    
    ' --- INICIO DE LA MODIFICACIÓN: ESTADO REVISIÓN EN MODIFICACIÓN ---
    ' Solo mostramos esta fila si estamos en el estado 4 (MODIFICACIÓN)
    If vm.Solicitud.idEstadoInterno = estadoModificacion Then
        Dim estadoRev As String
        Dim estiloRev As String
        
        estadoRev = UCase(Nz(vm.vm_revisionCalidadEstado, "PENDIENTE"))
        If estadoRev = "" Then estadoRev = "PENDIENTE"
        
        Select Case estadoRev
            Case "APROBADO"
                estiloRev = "color: #00A551; font-weight: bold;" ' Verde Telefónica
            Case "RECHAZADO"
                estiloRev = "color: #D6786B; font-weight: bold;" ' Coral (Rojo)
                estadoRev = "RECHAZADO PTE REVISÓN CAMBIOS" ' <--- REQ: ETIQUETA ESPECÍFICA
            Case Else
                estiloRev = "color: #E4C35C; font-weight: bold;" ' Amarillo Oscuro (Pendiente)
        End Select
        
        html = html & "      <tr><th>Decisión Calidad:</th><td style='" & estiloRev & "'>" & estadoRev & "</td></tr>"
    End If
    ' --- FIN DE LA MODIFICACIÓN ---

    ' --- INICIO DE LA CORRECCIÓN: VISUALIZACIÓN RECHAZO EN FASE TÉCNICA (Spec-003) ---
    If vm.Solicitud.idEstadoInterno = estadoDesarrolloTecnico Then
         Dim rechazoActivo As rechazo
         Set rechazoActivo = RechazoRepositorio.GetUltimoRechazoActivo(vm.Solicitud.idSolicitud)
         
         If Not rechazoActivo Is Nothing Then
             ' REQ: Visualización simplificada en una sola línea
             html = html & "      <tr><th>Decisión Calidad:</th><td style='color: #D6786B; font-weight: bold;'>RECHAZADO <span onclick=""window.colaComandos='VER_MOTIVO_CALIDAD:0'"" style='font-size: 0.9em; font-weight:normal; color:#0066FF; text-decoration:underline; margin-left:5px; cursor:pointer;'>[Ver Detalle]</span></td></tr>"
         End If
    End If
    ' --------------------------------------------------------------------------------
    
    html = html & "      <tr><th>NC Asociada:</th><td>" & vm.CodigoNCAsociada & "</td></tr>"
    html = html & "      <tr class='spacer'><td colspan='2'></td></tr>"
    html = html & "      <tr><th>Expediente:</th><td>" & vm.Expediente.Nemotecnico & " (" & vm.Expediente.CodExp & ")</td></tr>"
    html = html & "      <tr><th>Contratista:</th><td>" & vm.nombreContratistaPrincipal & "</td></tr>"
    html = html & "      <tr><th>Sub-suministradores:</th><td>" & Nz(vm.nombreSubcontratista, "N/A") & "</td></tr>"
    html = html & "      <tr class='spacer'><td colspan='2'></td></tr>"
    html = html & "      <tr><th>Resp. Técnico:</th><td>" & vm.JefeProyecto & "</td></tr>"
    html = html & "      <tr><th>Técnico Calidad:</th><td>" & vm.TecnicoCalidad & "</td></tr>"
    html = html & "      <tr class='spacer'><td colspan='2'></td></tr>"
    html = html & "      <tr><th>Creado por:</th><td>" & vm.Solicitud.usuarioCreacion & " el " & Format(vm.Solicitud.fechaCreacion, "dd/mm/yyyy hh:nn") & "</td></tr>"
    html = html & "      <tr><th>Tiempo de Ciclo:</th><td>" & vm.TiempoCiclo & "</td></tr>"
    html = html & "    </table>"
    html = html & "  </div>"
    html = html & "</div></body></html>"

Finalizar:
    GenerarHTML_FichaResumen = html
    Exit Function
    
Errores:
    html = "<html><body>Error al generar la ficha resumen.</body></html>"
    Resume Finalizar
End Function

' ==========================================================================
' MÓDULO: WorkflowServicio.cls
' ACCIÓN: REEMPLAZO del CSS para optimizar el espacio vertical y eliminar el scroll.
' ==========================================================================
Private Function GetFichaCSS() As String
    Dim css As String
    
    ' --- Ajustes: Reducido padding en .card-body, .data-table th/td y margin-bottom en .status-banner ---
    css = "body { font-family: 'Telefónica Sans', Arial, sans-serif; margin: 10px; background-color: transparent; font-size: 12px; } " & _
          ".card { background-color: #FFFFFF; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); border: 1px solid #D1D5E4; } " & _
          ".card-header { background-color: #F2F4FF; color: #0066FF; padding: 8px 12px; font-weight: bold; border-bottom: 1px solid #D1D5E4; border-radius: 7px 7px 0 0; font-size: 13px; } " & _
          ".card-body { padding: 10px 12px; } " & _
          ".status-banner { padding: 6px; margin-bottom: 10px; text-align: center; font-weight: bold; color: white; border-radius: 4px; } " & _
          ".status-revisión { background-color: #D6786B; } " & _
          ".status-aprobada { background-color: #528889; } " & _
          ".status-formalización { background-color: #E4C35C; color: #473515; } " & _
          ".data-table { border-collapse: collapse; width: 100%; } " & _
          ".data-table th, .data-table td { text-align: left; padding: 5px 6px; border-bottom: 1px solid #F2F4FF; vertical-align: top; } " & _
          ".data-table th { color: #58617A; width: 30%; font-weight: bold; } " & _
          ".data-table .spacer td { padding: 6px; border-bottom: 1px solid #D1D5E4; }"
          
    GetFichaCSS = css
End Function
' EN: WorkflowServicio.cls

Public Function GenerarHTML_GraficoTiempos(ByVal idSolicitud As Long) As String
    Dim solServ As New SolicitudServicio, sol As Solicitud
    Dim colHistorial As Scripting.Dictionary
    Dim tiemposPorEstado As New Scripting.Dictionary ' Almacenará los tiempos calculados (ID_Estado, Minutos)
    Dim html As String
    
    On Error GoTo Errores
    
    Set sol = solServ.getSolicitudPorID(idSolicitud)
    If sol Is Nothing Then GoTo Finalizar
    
    Set colHistorial = solServ.getHistorialDeEstados(idSolicitud)
    
    ' --- 1. Calcular los tiempos de los estados ya completados ---
    Dim log As LogEstado, key As Variant, fechaAnterior As Date
    Dim strEstadoID As String
    Dim tiempoCalculado As Long
    
    fechaAnterior = sol.fechaCreacion
    
    If Not colHistorial Is Nothing Then
        For Each key In colHistorial.Keys
            Set log = colHistorial(key)
            
            strEstadoID = CStr(log.idEstadoAnterior)
            tiempoCalculado = DateDiff("n", fechaAnterior, log.fechaTransicion)
            
            ' CORRECCIÓN CRÍTICA: PROTECCIÓN CONTRA CLAVES DUPLICADAS
            ' Si el estado ya existe (reproceso), sumamos el tiempo. Si no, lo creamos.
            If tiemposPorEstado.Exists(strEstadoID) Then
                tiemposPorEstado(strEstadoID) = CLng(tiemposPorEstado(strEstadoID)) + tiempoCalculado
            Else
                tiemposPorEstado.Add strEstadoID, tiempoCalculado
            End If
            
            fechaAnterior = log.fechaTransicion
        Next key
    End If
    
    ' --- 2. Calcular el tiempo del estado actual (si no es final) ---
    Dim estadoActual As estado
    Set estadoActual = m_ObjEntorno.estados(CStr(sol.idEstadoInterno))
    
    If Not estadoActual.esEstadoFinal Then
        strEstadoID = CStr(sol.idEstadoInterno)
        tiempoCalculado = DateDiff("n", fechaAnterior, Now())
        
        ' CORRECCIÓN TAMBIÉN AQUÍ:
        ' Puede que el estado actual ya haya ocurrido antes en el historial.
        If tiemposPorEstado.Exists(strEstadoID) Then
            tiemposPorEstado(strEstadoID) = CLng(tiemposPorEstado(strEstadoID)) + tiempoCalculado
        Else
            tiemposPorEstado.Add strEstadoID, tiempoCalculado
        End If
    End If
    
    ' --- 3. Preparar la estructura fija y encontrar el tiempo máximo para la escala ---
    Dim estadosOrdenados As Collection
    Set estadosOrdenados = SortDictionaryByProperty(m_ObjEntorno.estados, "orden")
    
    Dim datosParaGrafico As New Collection ' Colección de arrays (NombreEstado, Minutos)
    Dim maxMinutos As Double
    Dim unEstado As estado
    
    For Each unEstado In estadosOrdenados
        Dim tiempo As Long
        If tiemposPorEstado.Exists(CStr(unEstado.idEstado)) Then
            tiempo = tiemposPorEstado(CStr(unEstado.idEstado))
        Else
            tiempo = 0 ' Tiempo 0 para estados futuros o no pisados
        End If
        
        datosParaGrafico.Add Array(GetNombreEstadoSafe(unEstado.idEstado), tiempo)
        
        If tiempo > maxMinutos Then maxMinutos = tiempo
    Next unEstado
    
    If maxMinutos = 0 Then maxMinutos = 1 ' Evitar división por cero
    
    ' --- 4. Calcular tiempo total ---
    Dim tiempoTotal As String
    If estadoActual.esEstadoFinal And colHistorial.count > 0 Then
        ' Si está cerrada, calculamos desde creación hasta la última transición del log
        tiempoTotal = CalcularDuracion(sol.fechaCreacion, colHistorial.items(colHistorial.count - 1).fechaTransicion)
    Else
        ' Si está abierta, calculamos hasta AHORA
        tiempoTotal = CalcularDuracion(sol.fechaCreacion, Now()) & " (en curso)"
    End If

    ' --- 5. Construir el HTML ---
    html = "<!DOCTYPE html><html lang='es'><head><meta charset='UTF-8'><style>" & GetGraficoCSS() & "</style></head><body>"
    html = html & "<div class='card'>"
    html = html & "  <div class='card-header'>Tiempos por Fase (Acumulado)</div>"
    html = html & "  <div class='card-body'>"
    
    Dim colores(1 To 8) As String
    colores(1) = "#0066FF": colores(2) = "#75C0C7": colores(3) = "#E4C35C": colores(4) = "#F4C495"
    colores(5) = "#D6786B": colores(6) = "#B86BE8": colores(7) = "#528889": colores(8) = "#A7863E"
    
    html = html & "<ul class='chart-container'>"
    Dim item As Variant, i As Long
    i = 1
    For Each item In datosParaGrafico
        Dim porcentajeAncho As Double: porcentajeAncho = (item(1) / maxMinutos) * 100
        html = html & "<li>"
        html = html & "<div class='bar-label' title='" & item(0) & "'>" & item(0) & "</div>"
        html = html & "<div class='bar-container'><div class='bar' style='width:" & Format(porcentajeAncho, "0.00") & "%; background-color:" & colores((i - 1) Mod 8 + 1) & ";'></div></div>"
        html = html & "<div class='bar-time'>" & FormatearDuracion(item(1)) & "</div>"
        html = html & "</li>"
        i = i + 1
    Next item
    html = html & "</ul>"
    
    html = html & "</div>"
    html = html & "<div class='footer-card'><div class='footer-label'>Tiempo Total de Ciclo</div><div class='footer-value'>" & tiempoTotal & "</div></div>"
    html = html & "</div></body></html>"
    
Finalizar:
    GenerarHTML_GraficoTiempos = Replace(html, ",", ".")
    Exit Function
Errores:
    html = "<html><body>Error al generar el gráfico de tiempos: " & Err.description & "</body></html>"
    Resume Finalizar
End Function

' ==========================================================================
' MÓDULO: WorkflowServicio.cls
' ACCIÓN: REEMPLAZO del CSS del gráfico para hacerlo más espaciado y legible.
' ==========================================================================
Private Function GetGraficoCSS() As String
    Dim css As String
    
    ' --- Ajustes: Aumentado padding en .card-body y margin-bottom en .chart-container li ---
    css = "body { font-family: 'Telefónica Sans', Arial, sans-serif; margin: 0; padding: 15px; background-color: transparent; font-size: 12px; } " & _
          ".card { background-color: #FFFFFF; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); border: 1px solid #D1D5E4; display: flex; flex-direction: column; height: 95%; } " & _
          ".card-header { background-color: #F2F4FF; color: #0066FF; padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #D1D5E4; border-radius: 7px 7px 0 0; font-size: 13px; } " & _
          ".card-body { padding: 20px 15px; flex-grow: 1; } " & _
          ".chart-container { list-style: none; padding: 0; margin: 0; } " & _
          ".chart-container li { display: grid; grid-template-columns: 140px 1fr 100px; align-items: center; margin-bottom: 15px; gap: 10px; font-size: 12px; color: #2B3447; } " & _
          ".bar-label { text-align: right; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; } " & _
          ".bar-container { background-color: #F2F4FF; border-radius: 4px; height: 22px; } " & _
          ".bar { height: 100%; border-radius: 4px; transition: width 0.5s ease-in-out; } " & _
          ".bar-time { font-weight: bold; } " & _
          ".footer-card { border-top: 1px solid #D1D5E4; padding: 15px; display: flex; justify-content: center; align-items: center; gap: 15px; } " & _
          ".footer-label { font-size: 11pt; font-weight: bold; color: #0066FF; } " & _
          ".footer-value { font-size: 11pt; background-color: #F2F4FF; padding: 5px 10px; border-radius: 4px; color: #2B3447; }"
          
    GetGraficoCSS = css
End Function
Private Function FormatearDuracion(ByVal totalMinutos As Long) As String
    If totalMinutos < 0 Then totalMinutos = 0
    Dim dias As Long, horas As Long, minutos As Long
    dias = totalMinutos \ 1440
    horas = (totalMinutos Mod 1440) \ 60
    minutos = (totalMinutos Mod 1440) Mod 60
    Dim resultado As String
    If dias > 0 Then resultado = dias & "d "
    If horas > 0 Then resultado = resultado & horas & "h "
    resultado = resultado & minutos & "m"
    FormatearDuracion = Trim(resultado)
End Function
' ==========================================================================
' MÓDULO: WorkflowServicio.cls
' ACCIÓN: AÑADIR esta nueva función pública que devuelve solo el script.
' ==========================================================================
Public Function GetPestanasJS() As String
    Dim js As String
    js = "function openTab(evt, tabName) { " & _
         "  var i, tabcontent, tablinks; " & _
         "  tabcontent = document.getElementsByClassName('tab-content'); " & _
         "  for (i = 0; i < tabcontent.length; i++) { tabcontent[i].style.display = 'none'; } " & _
         "  tablinks = document.getElementsByClassName('tab-link'); " & _
         "  for (i = 0; i < tablinks.length; i++) { tablinks[i].className = tablinks[i].className.replace(' active', ''); } " & _
         "  document.getElementById(tabName).style.display = 'block'; " & _
         "  evt.currentTarget.className += ' active'; " & _
         "}"
    GetPestanasJS = js
End Function

Private Function GetTiemposPorEstado(ByVal idSolicitud As Long) As Scripting.Dictionary
    Dim solServ As New SolicitudServicio
    Dim sol As Solicitud
    Dim colHistorial As Scripting.Dictionary
    Dim tiemposPorEstado As New Scripting.Dictionary
    Dim log As LogEstado, key As Variant
    Dim fechaAnterior As Date
    Dim strEstadoID As String
    Dim tiempoCalculado As Long
    
    On Error GoTo Errores
    
    Set sol = solServ.getSolicitudPorID(idSolicitud)
    If sol Is Nothing Then GoTo Finalizar
    
    Set colHistorial = solServ.getHistorialDeEstados(idSolicitud)
    fechaAnterior = sol.fechaCreacion
    
    If Not colHistorial Is Nothing Then
        For Each key In colHistorial.Keys
            Set log = colHistorial(key)
            strEstadoID = CStr(log.idEstadoAnterior)
            tiempoCalculado = DateDiff("n", fechaAnterior, log.fechaTransicion)
            If tiemposPorEstado.Exists(strEstadoID) Then
                tiemposPorEstado(strEstadoID) = CLng(tiemposPorEstado(strEstadoID)) + tiempoCalculado
            Else
                tiemposPorEstado.Add strEstadoID, tiempoCalculado
            End If
            fechaAnterior = log.fechaTransicion
        Next key
    End If
    
Finalizar:
    Set GetTiemposPorEstado = tiemposPorEstado
    Exit Function
Errores:
    Set GetTiemposPorEstado = Nothing
End Function

Public Function ValidarYExportarDocumentoParaRAC(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As String
    ' RESPONSABILIDAD: Orquesta el proceso completo de "Validar y Exportar".
    ' Es transaccional para el acceso a BD.
    Dim docServ As New DocumentoServicio
    Dim adjServ As New AdjuntosServicio
    Dim ws As DAO.Workspace
    Dim dbTrans As DAO.Database
    Dim rutaDocumento As String
    Dim sol As Solicitud
    Dim dbOrigIsNothing As Boolean
    
    On Error GoTo Errores
    
    dbOrigIsNothing = (db Is Nothing)
    If dbOrigIsNothing Then Set db = getdb()
    
    ' 1. Generar el documento (lógica polimórfica)
    rutaDocumento = docServ.GenerarDocumentoParaSolicitud(idSolicitud)
    If Not fso.FileExists(rutaDocumento) Then
        Err.Raise 513, "ValidarExportar", "Falló la generación del documento Word."
    End If

    ' 2. Iniciar transacción para adjuntar y cambiará estado
    If dbOrigIsNothing Then
        Set ws = getWorkspace()
        Set dbTrans = db  ' Use db directly; ws is used only for BeginTrans/CommitTrans/Rollback
        ws.BeginTrans
    Else
        Set dbTrans = db
    End If
    
    ' 3. Guardar el documento como adjunto oficial
     Dim nombreEtapaDestino As String
    nombreEtapaDestino = GetNombreEstadoSafe(estadoRevision)
    
    Call adjServ.GuardarAdjuntoDesdeArchivo(idSolicitud, nombreEtapaDestino, rutaDocumento, "Generado", dbTrans)
    
    ' --- INICIO DE LA CORRECCIÓN: Lógica Condicional de Avance ---
    Dim solServ As New SolicitudServicio
    Set sol = solServ.getSolicitudPorID(idSolicitud)
    
    ' Solo ejecutamos la transición si NO estamos ya en REVISIÓN.
    ' Si ya estamos en revisión, significa que estamos regenerando el documento (iteración),
    ' por lo que no hace falta cambiará el estado.
    If sol.idEstadoInterno <> estadoRevision Then
        Call Me.EjecutarTransicion(sol, estadoRevision, m_ObjUsuarioActivo, dbTrans)
    End If
    ' --- FIN DE LA CORRECCIÓN ---
    
    ' 5. Confirmar transacción
    If dbOrigIsNothing Then ws.CommitTrans
    
    ValidarYExportarDocumentoParaRAC = rutaDocumento
    
LimpiarYSalir:
    On Error Resume Next
    If dbOrigIsNothing And Not dbTrans Is Nothing Then
        dbTrans.Close
    End If
    Set ws = Nothing
    Exit Function
    
Errores:
    If dbOrigIsNothing And Not ws Is Nothing Then
        ws.Rollback
    End If
    ' Lógica de compensación: si el documento se creó pero falló el resto, lo borramos.
    ' Si la compensación falla, es error crítico — no silenciamos.
    If fso.FileExists(rutaDocumento) Then fso.DeleteFile rutaDocumento, True

    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.ValidarYExportarDocumentoParaRAC"
    errObj.Raise
End Function

Public Function GenerarYAdjuntarPDFfinal(ByVal idSolicitud As Long) As String
    ' RESPONSABILIDAD: Orquesta la generación y adjunto del PDF final.
    Dim docServ As New DocumentoServicio
    Dim adjServ As New AdjuntosServicio
    Dim rutaPDF As String
    
    On Error GoTo Errores
    
    ' 1. Generar el PDF (lógica polimórfica)
    rutaPDF = docServ.GenerarDocumentoPDFparaSolicitud(idSolicitud)
    
    ' 2. Adjuntar el PDF a la solicitud (Spec-135: etapa=Formalización, descripcion=Borrador)
    Dim nombreEtapaActual As String
    nombreEtapaActual = GetNombreEstadoSafe(estadoFormalizacion)
    
    Call adjServ.GuardarAdjuntoDesdeArchivo(idSolicitud, nombreEtapaActual, rutaPDF, "Generado", , , "Borrador")
    
    GenerarYAdjuntarPDFfinal = rutaPDF
    
LimpiarYSalir:
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.GenerarYAdjuntarPDFfinal"
    errObj.Raise
End Function
' --- EN: WorkflowServicio.cls ---

Private Function GenerarChecklistHTML(ByRef sol As Solicitud) As String
    Dim html As String
    Dim pcServ As New DatosPCServicio
    Dim cdcaServ As New DatosCDCAServicio
    Dim cdcaSubS As New DatosCDCASUBServicio
    Dim pcsubServ As New DatosPCSUBServicio
    Dim adjServ As New AdjuntosServicio ' <<< ESTA ES LA LÍNEA QUE FALTABA
    Dim itemHtml As String
    Dim hayDecision As Boolean
    
    html = "<ul class='checklist'>"
    
    Select Case sol.idEstadoInterno
        ' ... [Casos anteriores: Preregistro, Registro, Desarrollo, Modificación, Validación SIN CAMBIOS] ...
        Case estadoPreregistro
             Dim hayGenerales As Boolean
If sol.tipoSolicitud = "PC" Then hayGenerales = pcServ.EsDatosGeneralesCompleta(sol.idSolicitud)
              If sol.tipoSolicitud = "CD_CA" Then hayGenerales = cdcaServ.EsDatosGeneralesCompleta(sol.idSolicitud)
              If sol.tipoSolicitud = "CD_CA_SUB" Then
                  hayGenerales = cdcaSubS.EsDatosGeneralesCompleta(sol.idSolicitud)
              End If
              If sol.tipoSolicitud = "PC_SUB" Then
                  hayGenerales = pcsubServ.EsDatosGeneralesCompleta(sol.idSolicitud)
              End If
              html = html & GetItemChecklist(hayGenerales, "Completar pestaña 'Datos Generales'")
        Case estadoRegistro
             html = html & GetItemChecklist(True, "Revisión del alcance (Manual)")
        Case estadoDesarrolloTecnico
             Dim hayTecnica As Boolean
If sol.tipoSolicitud = "PC" Then hayTecnica = pcServ.EsParteTecnicaCompleta(sol.idSolicitud)
              If sol.tipoSolicitud = "CD_CA" Then hayTecnica = cdcaServ.EsParteTecnicaCompleta(sol.idSolicitud)
              If sol.tipoSolicitud = "CD_CA_SUB" Then hayTecnica = cdcaSubS.EsParteTecnicaCompleta(sol.idSolicitud)
              If sol.tipoSolicitud = "PC_SUB" Then hayTecnica = pcsubServ.EsParteTecnicaCompleta(sol.idSolicitud)
              html = html & GetItemChecklist(hayTecnica, "Completar pestañas 'Detalle y Descripción' y 'Motivos'")
        Case estadoModificacion
             Dim itemsStrMod As String
             itemsStrMod = GetItemsChecklistPorEstado(estadoModificacion, sol.tipoSolicitud, sol.revisionCalidadEstado)
             Dim arrItemsMod As Variant
             arrItemsMod = Split(itemsStrMod, vbCrLf)
             
             Dim hayRACMod As Boolean
             Dim hayAprobMod As Boolean
             If sol.tipoSolicitud = "PC" Then
                 hayRACMod = pcServ.EsDictamenRACCompleto(sol.idSolicitud)
                 hayAprobMod = pcServ.EsAprobacionSuministradorCompleta(sol.idSolicitud)
             End If
             If sol.tipoSolicitud = "CD_CA" Then
                 hayRACMod = cdcaServ.EsDictamenRACCompleto(sol.idSolicitud)
                 hayAprobMod = cdcaServ.EsAprobacionSuministradorCompleta(sol.idSolicitud)
             End If
If sol.tipoSolicitud = "CD_CA_SUB" Then
                  hayRACMod = cdcaSubS.EsDictamenRACCompleto(sol.idSolicitud)
                  Dim subSMod2 As New DatosCDCASUBServicio
                  hayAprobMod = subSMod2.EsAprobacionSuministradorCompleta(sol.idSolicitud)
              End If
              If sol.tipoSolicitud = "PC_SUB" Then
                  hayRACMod = pcsubServ.EsDictamenRACCompleto(sol.idSolicitud)
                  hayAprobMod = pcsubServ.EsAprobacionSuministradorCompleta(sol.idSolicitud)
              End If
             
             Dim idxMod As Long
             Dim estadoItemMod As Boolean
             For idxMod = LBound(arrItemsMod) To UBound(arrItemsMod)
                 If Trim$(arrItemsMod(idxMod)) <> "" Then
                     If idxMod = LBound(arrItemsMod) Then
                         estadoItemMod = False
                     ElseIf idxMod = LBound(arrItemsMod) + 1 Then
                         estadoItemMod = hayAprobMod
                     ElseIf idxMod = LBound(arrItemsMod) + 2 Then
                         estadoItemMod = hayRACMod
                     Else
                         estadoItemMod = False
                     End If
                     html = html & GetItemChecklist(estadoItemMod, CStr(arrItemsMod(idxMod)))
                 End If
             Next idxMod
        Case estadoValidacion
             Dim hayBorradorGenerado As Boolean
             Dim cicloValidacion As ValidacionRevision
             Set cicloValidacion = ValidacionRevisionRepositorio.GetCicloAbierto(sol.idSolicitud)
             If Not cicloValidacion Is Nothing Then
                 hayBorradorGenerado = (cicloValidacion.idAdjunto > 0)
             Else
                 hayBorradorGenerado = False
             End If
html = html & GetItemChecklist(hayBorradorGenerado, "Exportar borrador")
              html = html & GetItemChecklist(False, "Registrar Fecha Envio RAC")
              
         ' --- ESTADO REVISIÓN ---
         Case estadoRevision
              Dim hayTecnicaVal As Boolean
              If sol.tipoSolicitud = "PC" Then hayTecnicaVal = pcServ.EsParteTecnicaCompleta(sol.idSolicitud)
              If sol.tipoSolicitud = "CD_CA" Then hayTecnicaVal = cdcaServ.EsParteTecnicaCompleta(sol.idSolicitud)
              If sol.tipoSolicitud = "CD_CA_SUB" Then hayTecnicaVal = cdcaSubS.EsParteTecnicaCompleta(sol.idSolicitud)
              If sol.tipoSolicitud = "PC_SUB" Then hayTecnicaVal = pcsubServ.EsParteTecnicaCompleta(sol.idSolicitud)
              html = html & GetItemChecklist(hayTecnicaVal, "Validación de Datos Técnicos")
             
        Case estadoFormalizacion
             Dim pdfCierre As Boolean
             pdfCierre = adjServ.ExisteAdjuntoEtapa(sol.idSolicitud, "Cierre")
             html = html & GetItemChecklist(pdfCierre, "Adjuntar documento de Cierre firmado")
    End Select
    
    html = html & "</ul>"
    
    GenerarChecklistHTML = html
End Function

Private Function GetItemChecklist(cumplido As Boolean, texto As String) As String
    If cumplido Then
        GetItemChecklist = "<li class='req-ok'><span class='check-icon'>&#10003;</span> <span class='req-text'>" & texto & "</span></li>"
    Else
        GetItemChecklist = "<li class='req-nok'><span class='check-icon'>&#10007;</span> <span class='req-text'>" & texto & "</span></li>"
    End If
End Function
' ==========================================================================
' MÓDULO: WorkflowServicio.cls
' ACCIÓN: Función auxiliar para obtener nombres de estado dinámicamente.
' ==========================================================================
Private Function GetNombreEstadoSafe(ByVal idEstado As Long) As String
    ' Obtiene el nombre real de la base de datos (vía caché de Entorno)
    ' Si fallóa, devuelve el ID como fallóback.
    On Error Resume Next
    If m_ObjEntorno.estados.Exists(CStr(idEstado)) Then
        GetNombreEstadoSafe = m_ObjEntorno.estados(CStr(idEstado)).nombreEstado
    Else
        GetNombreEstadoSafe = "Estado " & idEstado
    End If
End Function
' EN: WorkflowServicio.cls
' EN: WorkflowServicio.cls

Public Sub EjecutarCierreFormalizacion(ByRef sol As Solicitud, ByRef objDatos As Object, ByVal usuario As usuario, ByVal rutaArchivoFisico As String, Optional ByRef db As DAO.Database)
    Dim ws As DAO.Workspace
    Dim dbTrans As DAO.Database
    Dim adjServ As New AdjuntosServicio
    Dim pcServ As New DatosPCServicio
    Dim pcsubServ As New DatosPCSUBServicio
    Dim cdcaServ As New DatosCDCAServicio
    Dim cdcaSubServ As New DatosCDCASUBServicio
    Dim nombreEtapaPDF As String
    Dim nombreEtapaPDFLegacy As String
    Dim nombreEtapaPDFLegacy2 As String
    Dim dbOrigIsNothing As Boolean
    
    On Error GoTo Errores
    If Trim(rutaArchivoFisico) = "" Or Not fso.FileExists(rutaArchivoFisico) Then Err.Raise 514, , "Archivo inválido."
    
    dbOrigIsNothing = (db Is Nothing)
    If dbOrigIsNothing Then Set db = getdb()
    
    If dbOrigIsNothing Then
        Set ws = getWorkspace()
        Set dbTrans = db  ' Use db directly; ws is used only for BeginTrans/CommitTrans/Rollback
        ws.BeginTrans
    Else
        Set dbTrans = db
    End If
    
    ' 1. Guardar Datos
    Select Case sol.tipoSolicitud
        Case "PC": Call pcServ.GuardarDecisionFinal(objDatos, dbTrans)
        Case "PC_SUB": Call pcsubServ.GuardarDecisionFinal(objDatos, dbTrans)
        Case "CD_CA": Call cdcaServ.GuardarDecisionFinal(objDatos, dbTrans)
        Case "CD_CA_SUB": Call cdcaSubServ.GuardarDecisionFinal(objDatos, dbTrans)
    End Select
    
    ' 2. Adjunto
    nombreEtapaPDF = GetNombreEstadoSafe(estadoFormalizacion)
    nombreEtapaPDFLegacy = "Cierre"
    nombreEtapaPDFLegacy2 = "Documento Final Firmado"
    Dim colAdj As Collection, adj As Adjunto
    Set colAdj = adjServ.getAdjuntosPorSolicitud(sol.idSolicitud)
    If Not colAdj Is Nothing Then
        For Each adj In colAdj
            If adj.etapaWF = nombreEtapaPDF Or adj.etapaWF = nombreEtapaPDFLegacy Or adj.etapaWF = nombreEtapaPDFLegacy2 Then
                Call AdjuntoRepositorio.Eliminar(adj.idAdjunto, dbTrans)
            End If
        Next adj
    End If
    Call adjServ.GuardarAdjuntoDesdeArchivo(sol.idSolicitud, nombreEtapaPDF, rutaArchivoFisico, "Generado", dbTrans)
    
    ' 3. Transición
    Call Me.EjecutarTransicion(sol, estadoAprobada, usuario, dbTrans)
    
    If dbOrigIsNothing Then ws.CommitTrans
    
LimpiarYSalir:
    On Error Resume Next
    If dbOrigIsNothing And Not dbTrans Is Nothing Then dbTrans.Close
    Set ws = Nothing
    Exit Sub
Errores:
    If dbOrigIsNothing And Not ws Is Nothing Then ws.Rollback
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.EjecutarCierreFormalizacion"
    errObj.Raise
End Sub

' Helper privado para creóar el objeto adjunto rápidamente
Private Function creóarObjetoAdjunto(idSol As Long, etapa As String, ruta As String, user As String) As Adjunto
    Dim a As New Adjunto
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    a.idSolicitud = idSol
    a.etapaWF = etapa
    a.nombreArchivo = fso.GetFileName(ruta) ' Asumimos que la copia física se gestáiona aparte o usamos el mismo nombre
    ' NOTA IMPORTANTE: El método GuardarAdjuntoDesdeArchivo hacía la copia física.
    ' Aquí estamos registrando en BD. Para simplificar y reutilizar, asumiremos que
    ' el archivo ya está en la ruta destino o usamos el helper existente.
    
    ' RECTIFICACIÓN: Para garantizar que el archivo físico está en la carpeta de red,
    ' debemos copiarlo antes.
    Dim destinoFull As String
    destinoFull = m_ObjEntorno.URLDirectorioDocumentacion & fso.GetFileName(ruta)
    fso.CopyFile ruta, destinoFull, True
    
    a.nombreArchivo = fso.GetFileName(ruta)
    a.fechaSubida = Now()
    a.usuarioSubida = user
    a.descripcion = "Generado Automáticamente al Aprobar"
    a.tipoAccion = "Generado"
    Set creóarObjetoAdjunto = a
End Function
Public Function VerificarCambiosTrasRechazo(ByRef sol As Solicitud) As Boolean
    ' RESPONSABILIDAD: Verificar si existen cambios registrados en el JSON de trazabilidad (tbRechazos).
    ' REFACTORIZADO (Spec-007/009): Se sustituye la consulta a tbLogCambios por la lectura directa
    ' del JSON de cambios acumulados gestionado por JsonHelper.
    
    On Error GoTo Errores
    
    ' Si no está rechazado, pase libre.
    If sol.revisionCalidadEstado <> "RECHAZADO" Then
        VerificarCambiosTrasRechazo = True
        Exit Function
    End If
    
    Dim jsonDict As Object
    Set jsonDict = JsonHelper.ObtenerJsonRechazo(sol.idSolicitud)
    
    ' Si el diccionario tiene claves, significa que hay cambios técnicos registrados y activos.
    If jsonDict.count > 0 Then
        VerificarCambiosTrasRechazo = True
    Else
        VerificarCambiosTrasRechazo = False
    End If
    
    Exit Function
Errores:
    VerificarCambiosTrasRechazo = False ' Ante duda técnica, bloquear para seguridad
End Function

Public Sub RechazarFaseTecnica(ByVal sol As Solicitud, ByVal motivo As String, ByVal usuario As usuario, Optional ByRef db As DAO.Database)
    Dim logServ As New LogServicio
    Dim logEst As New LogEstado
    ' RechazoRepositorio es un módulo estándar, lo llamamos directamente.
    
    Dim ws As DAO.Workspace
    Dim dbTrans As DAO.Database
    Dim gestionaTransaccionLocal As Boolean
    
    On Error GoTo Errores
    
    ' --- TRANSACCIÓN HÍBRIDA (Spec-XXX: compatible con llamada aislada o en contexto de llamador) ---
    If db Is Nothing Then
        ' Modo Stand-alone -> Creamos transacción local
        gestionaTransaccionLocal = True
        Set ws = getWorkspace()
        Set dbTrans = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
    Else
        ' Modo Heredado (caller ya tiene db con transacción abierta) -> reutilizamos
        gestionaTransaccionLocal = False
        Set dbTrans = db
    End If
    
    If usuario.rol <> rol.Administrador And usuario.rol <> rol.Calidad Then
        Err.Raise 513, "RechazarFaseTecnica", "No tiene permisos para rechazar."
    End If
    
    ' 1. Guardar el rechazo en el histórico
    Dim objRechazo As New rechazo
    With objRechazo
        .idSolicitud = sol.idSolicitud
        .fechaRechazo = Now()
        .motivoPrincipal = "Rechazo Fase Técnica"
        .areaAfectada = "Desarrollo Técnico"
        .comentarios = motivo
        .usuarioRechazo = usuario.nombre
    End With
    Call RechazoRepositorio.GuardarRechazo(objRechazo, dbTrans)
    
    ' 2. Actualizar estado de revisión en la solicitud (redundancia útil)
    sol.revisionCalidadEstado = "RECHAZADO"
    sol.revisionCalidadComentarios = motivo
    
    ' 3. Retroceder estado (4 -> 3)
    Dim estadoAnteriorID As Long
    estadoAnteriorID = sol.idEstadoInterno
    sol.idEstadoInterno = estadoDesarrolloTecnico
    sol.fechaModificacion = Now()
    sol.usuarioModificacion = usuario.nombre
    
    Call SolicitudRepositorio.GuardarSolicitud(sol, dbTrans)
    
    ' 4. Log de Estado
    With logEst
        .idSolicitud = sol.idSolicitud
        .idEstadoAnterior = estadoAnteriorID
        .idEstadoNuevo = estadoDesarrolloTecnico
        .fechaTransicion = sol.fechaModificacion
        .usuarioTransicion = sol.usuarioModificacion
    End With
    Call LogEstadoRepositorio.Guardar(logEst, dbTrans)
    
    ' 5. Log de Auditoría
    Dim estadoAnteriorNombre As String, estadoNuevoNombre As String
    estadoAnteriorNombre = GetNombreEstadoSafe(estadoAnteriorID)
    estadoNuevoNombre = GetNombreEstadoSafe(estadoDesarrolloTecnico)
    
    Call logServ.RegistrarCambio("tbSolicitudes", sol.idSolicitud, "idEstadoInterno", _
        estadoAnteriorID & " (" & estadoAnteriorNombre & ")", _
        estadoDesarrolloTecnico & " (" & estadoNuevoNombre & ")", _
        "RECHAZO CALIDAD: " & motivo, dbTrans)
        
    ' 6. Commit (solo si somos dueños de la transacción)
    If gestionaTransaccionLocal Then
        ws.CommitTrans
    End If
    
    ' 7. Llamar a NotificarTransicion para enviar notificaciones (sin validar)
    Call NotificarTransicion(sol, estadoDesarrolloTecnico, usuario)
    
LimpiarYSalir:
    On Error Resume Next
    ' Cerrar solo el dbTrans que CREAMOS nosotros; nunca el proporcionado por el caller
    If gestionaTransaccionLocal And Not dbTrans Is Nothing Then dbTrans.Close
    Set ws = Nothing
    Exit Sub
Errores:
    ' Rollback solo si somos dueños de la transacción
    If gestionaTransaccionLocal And Not ws Is Nothing Then
        ws.Rollback
    End If
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "WorkflowServicio.RechazarFaseTecnica"
    errObj.AddToCallStack "Solicitud: " & sol.codigoSolicitud
    errObj.Raise
End Sub

' ==========================================================================
' MÓDULO: WorkflowServicio.cls
' FUNCIÓN: NotificarTransicion
' DESCRIPCIÓN: Envía notificaciones de transición SIN validar ni escribir en BD.
'              Usado por funciones que ya han hecho la transición (ej: RechazarFaseTecnica).
' ==========================================================================
Private Sub NotificarTransicion(ByRef sol As Solicitud, ByVal idEstadoDestino As Long, ByVal usuario As usuario, Optional ByVal idEstadoOrigen As Variant)
    On Error Resume Next
    
    Dim notifServ As New NotificacionServicio
    Dim solServ As New SolicitudServicio
    Dim vm As SolicitudViewModel
    Dim destinatario As String, copia As String, asunto As String, infoExpediente As String
    Dim descriptorAccion As String
    Dim sufijoSubsanacion As String
    Dim nombreEstadoNotif As String
    
    Set vm = solServ.getSolicitudViewModelPorID(sol.idSolicitud)
    copia = m_ObjUsuarioReal.CorreoUsuario
    infoExpediente = vm.Expediente.Nemotecnico & " | " & vm.TipoSolicitudNombre
    
    nombreEstadoNotif = GetNombreEstadoSafe(idEstadoDestino)
    If UCase$(Trim$(Nz(sol.revisionCalidadEstado, ""))) = "RECHAZADO" And (idEstadoDestino = estadoModificacion Or idEstadoDestino = estadoDesarrolloTecnico) Then
        nombreEstadoNotif = "RECHAZADO_POR_CALIDAD"
    End If
    If (vm.vm_racValidacionEstado = "RECHAZADO" Or vm.vm_racValidacionEstado = "COMENTARIOS") And idEstadoDestino = estadoValidacion Then
        nombreEstadoNotif = "REVISION_RAC"
    End If
    If Not IsEmpty(idEstadoOrigen) And idEstadoOrigen = estadoRevision And idEstadoDestino = estadoValidacion Then
        nombreEstadoNotif = "RECHAZO_RAC"
    End If
    
    destinatario = notifServ.ObtenerDestinatariosPorEstado(nombreEstadoNotif, vm)
    
    Select Case idEstadoDestino
        Case estadoDesarrolloTecnico
            descriptorAccion = "DESARROLLO TÉCNICO"
            asunto = "CONDOR - " & nombreEstadoNotif & " | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
        Case estadoRegistro
            asunto = "CONDOR - " & nombreEstadoNotif & " | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
        Case estadoRevision
            asunto = "CONDOR - " & nombreEstadoNotif & " | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
        Case estadoModificacion
            sufijoSubsanacion = ""
            If UCase$(Trim$(Nz(sol.revisionCalidadEstado, ""))) = "RECHAZADO" Then
                sufijoSubsanacion = " (RECHAZADA POR CALIDAD)"
            End If
            asunto = "CONDOR - " & nombreEstadoNotif & sufijoSubsanacion & " | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
        Case estadoValidacion
            asunto = "CONDOR - " & nombreEstadoNotif & " | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
            If (vm.vm_racValidacionEstado = "RECHAZADO" Or vm.vm_racValidacionEstado = "COMENTARIOS") Then
                asunto = "CONDOR - " & nombreEstadoNotif & " TRAS COMENTARIOS RAC/CLIENTE | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
            End If
        Case estadoFormalizacion
            nombreEstadoNotif = "APROBADO_RAC"
            asunto = "CONDOR - " & nombreEstadoNotif & " | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
        Case estadoAprobada
            nombreEstadoNotif = "APROBACIÓN"
            asunto = "CONDOR - " & nombreEstadoNotif & " | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
        Case Else
            asunto = "CONDOR - " & nombreEstadoNotif & " | " & infoExpediente & " | " & vm.Solicitud.codigoSolicitud
    End Select
    
    If destinatario <> "" Then
        Call notifServ.EnviarNotificacion(vm, asunto, destinatario, copia, nombreEstadoNotif)
            If Err.Number <> 0 Then
                Err.Clear
            End If
    End If
End Sub

' ==========================================================================
' MÓDULO: WorkflowServicio.cls
' FUNCIÓN: GenerarHTML_DeltasRechazo (Spec-009a)
' DESCRIPCIÓN: Genera un visualizador HTML atractivo para los cambios técnicos.
' ==========================================================================
Public Function GenerarHTML_DeltasRechazo(ByVal idSolicitud As Long) As String
    Dim jsonDict As Object
    Dim html As String
    Dim bloqueKey As Variant, campoKey As Variant
    Dim bloque As Object, campo As Object
    Dim hayCambios As Boolean
    Dim infoTecnico As String
    Dim htmlPropuesta As String
    Dim htmlImpacto As String
    Dim tienePropuesta As Boolean
    Dim tieneImpacto As Boolean
    Dim grupo As String
    Dim fila As String
    
    On Error GoTo Errores
    
    Set jsonDict = JsonHelper.ObtenerJsonRechazo(idSolicitud)
    
    ' --- INICIO MODIFICACIÓN FORENSE ---
    ' Si no hay rechazo activo, buscamos el último histórico (Spec-032.2)
    If jsonDict.count = 0 Then
        Dim rsHist As DAO.Recordset
        Set rsHist = RechazoRepositorio.GetHistorialPorSolicitud(idSolicitud)
        If Not rsHist Is Nothing Then
            If Not rsHist.EOF Then
                ' Intentamos recuperar el JSON del último rechazo histórico
                Dim rawJson As String
                rawJson = Nz(rsHist!CambiosTecnico, "")
                If Len(rawJson) > 0 Then
                    Set jsonDict = JsonConverter.ParseJson(rawJson)
                End If
            End If
            rsHist.Close
        End If
    End If
    ' --- FIN MODIFICACIÓN FORENSE ---
    
    html = "<!DOCTYPE html><html lang='es'><head><meta charset='UTF-8'><style>" & GetDeltasCSS() & "</style></head><body>"
    html = html & "<div class='container'>"
    
    If jsonDict.count = 0 Then
        html = html & "<div class='empty-state'>No se han detectado cambios técnicos registrados para este rechazo.</div>"
    Else
        ' 1. Extraer info del primer técnico (Spec-009a: Simplificación UI)
        For Each bloqueKey In jsonDict.Keys
            Set bloque = jsonDict(bloqueKey)
            For Each campoKey In bloque.Keys
                Set campo = bloque(campoKey)
                infoTecnico = "<div class='tecnico-header'>Modificado por: <strong>" & campo("usuario_mod") & "</strong> el " & Format(campo("fecha_ult_mod"), "dd/mm/yyyy hh:nn") & "</div>"
                Exit For
            Next campoKey
            If infoTecnico <> "" Then Exit For
        Next bloqueKey
        
        html = html & infoTecnico
        
        hayCambios = False
        For Each bloqueKey In jsonDict.Keys
            Set bloque = jsonDict(bloqueKey)
            grupo = ObtenerGrupoBloque(CStr(bloqueKey))
            For Each campoKey In bloque.Keys
                Set campo = bloque(campoKey)
                If ValoresDiferentes(campo("valor_inicial"), campo("valor_final")) Then
                    fila = "<tr><td class='col-campo'>" & HtmlEncode(NormalizarNombreCampo(CStr(campoKey))) & "</td>" & _
                           "<td class='col-antes'>" & ValorParaHtml(campo("valor_inicial")) & "</td>" & _
                           "<td class='col-despues'>" & ValorParaHtml(campo("valor_final")) & "</td></tr>"
                    If grupo = "PROPUESTA" Then
                        htmlPropuesta = htmlPropuesta & fila
                        tienePropuesta = True
                    Else
                        htmlImpacto = htmlImpacto & fila
                        tieneImpacto = True
                    End If
                    hayCambios = True
                End If
            Next campoKey
        Next bloqueKey
        
        If Not hayCambios Then
            html = html & "<div class='empty-state'>No se han detectado cambios técnicos registrados para este rechazo.</div>"
        Else
            html = html & ConstruirSeccion("DETALLE Y DESCRIPCIÓN", htmlPropuesta, tienePropuesta)
            html = html & ConstruirSeccion("MOTIVOS", htmlImpacto, tieneImpacto)
        End If
    End If
    
    html = html & "</div></body></html>"
    
Finalizar:
    GenerarHTML_DeltasRechazo = html
    Exit Function
    
Errores:
    html = "<html><body>Error al generar visor de cambios: " & Err.description & "</body></html>"
    Resume Finalizar
End Function

Public Function GenerarHTML_HistorialRAC(ByVal idSolicitud As Long) As String
    Dim rcd As DAO.Recordset
    Dim html As String
    Dim ordinalValue As Long
    Dim resultadoValue As String
    Dim claseEstado As String
    Dim fechaEnvioStr As String
    Dim fechaRecepcionStr As String
    Dim usuarioStr As String
    Dim comentariosRaw As String
    Dim comentariosHtml As String
    Dim idAdjunto As Long
    Dim etapaBorrador As String
    Dim wordLink As String
    
    On Error GoTo Errores
    
    html = "<!DOCTYPE html><html lang='es'><head><meta charset='UTF-8'><style>" & GetHistorialRACCSS() & "</style></head><body>"
    html = html & "<div class='container'>"
    html = html & "<div class='page-title'>Historial de Validaci&oacute;n y Revisi&oacute;n</div>"
    
    If idSolicitud <= 0 Then GoTo SinDatos
    
    Set rcd = ValidacionRevisionRepositorio.GetHistorialPorSolicitud(idSolicitud)
    If rcd Is Nothing Then GoTo SinDatos
    If rcd.EOF Then GoTo SinDatos
    
    rcd.MoveFirst
    Do While Not rcd.EOF
        ordinalValue = Nz(rcd!ordinal, 0)
        resultadoValue = UCase$(CStr(Nz(rcd!resultado, "PENDIENTE")))
        If resultadoValue = "" Then resultadoValue = "PENDIENTE"
        
        Select Case resultadoValue
            Case "APROBADO"
                claseEstado = "aprobado"
            Case "COMENTARIOS"
                claseEstado = "comentarios"
            Case Else
                resultadoValue = "PENDIENTE"
                claseEstado = "pendiente"
        End Select
        
        If IsDate(rcd!fechaEnvio) Then
            fechaEnvioStr = Format(rcd!fechaEnvio, "dd/mm/yyyy")
        Else
            fechaEnvioStr = "-"
        End If
        
        If IsDate(rcd!fechaRecepcion) Then
            fechaRecepcionStr = Format(rcd!fechaRecepcion, "dd/mm/yyyy")
        Else
            fechaRecepcionStr = "-"
        End If
        
        usuarioStr = HtmlEncode(CStr(Nz(rcd!usuario, "")))
        
        comentariosRaw = CStr(Nz(rcd!comentarios, ""))
        If Trim(comentariosRaw) = "" Then comentariosRaw = "(Vacio)"
        comentariosHtml = HtmlEncode(comentariosRaw)
        comentariosHtml = Replace(comentariosHtml, vbCrLf, "<br>")
        comentariosHtml = Replace(comentariosHtml, vbLf, "<br>")
        
        idAdjunto = Nz(rcd!idAdjunto, 0)
        wordLink = ""
        If idAdjunto > 0 Then
            wordLink = "<a class='word-link' href='vba:ABRIR_ADJUNTO:" & CStr(idAdjunto) & "'>[Word]</a>"
        End If
        
        html = html & "<div class='card " & claseEstado & "'>"
        html = html & "<div class='card-header'>"
        html = html & "<div class='card-title'>Historial de Validaci&oacute;n y Revisi&oacute;n (v" & CStr(ordinalValue) & ")</div>"
        html = html & "<div class='card-actions'>"
        If wordLink <> "" Then html = html & wordLink
        html = html & "<div class='badge " & claseEstado & "'>" & resultadoValue & "</div>"
        html = html & "</div>"
        html = html & "</div>"
        html = html & "<div class='card-meta'>"
        html = html & "<span><strong>Envio:</strong> " & fechaEnvioStr & "</span>"
        html = html & "<span><strong>Recepcion:</strong> " & fechaRecepcionStr & "</span>"
        html = html & "<span><strong>Usuario:</strong> " & usuarioStr & "</span>"
        html = html & "</div>"
        html = html & "<div class='card-body'>" & comentariosHtml & "</div>"
        html = html & "</div>"
        
        rcd.MoveNext
    Loop
    
    GoTo Finalizar
    
SinDatos:
    html = html & "<div class='empty-state'>Sin historial registrado.</div>"
    
Finalizar:
    html = html & "</div></body></html>"
    GenerarHTML_HistorialRAC = html
    
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
    
Errores:
    html = "<html><body>Error al generar historial de validacion y revision: " & Err.description & "</body></html>"
    Resume Finalizar
End Function

Public Function GenerarHTML_CapturaFechaRAC(ByVal fechaDefault As Variant) As String
    Dim html As String
    Dim valorFecha As String
    
    If IsDate(fechaDefault) Then
        valorFecha = Format(fechaDefault, "yyyy-mm-dd")
    Else
        valorFecha = Format(Date, "yyyy-mm-dd")
    End If
    
    html = "<!DOCTYPE html><html lang='es'><head><meta charset='UTF-8'><style>" & GetCapturaRACCSS() & "</style></head><body>"
    html = html & "<div class='container'>"
    html = html & "<div class='page-title'>Notificar envio al RAC</div>"
    html = html & "<div class='card'>"
    html = html & "<label class='label'>Fecha de envio</label>"
    html = html & "<input id='fechaEnvio' class='input' type='date' value='" & valorFecha & "' />"
    html = html & "<div id='error' class='error' style='display:none;'>Debe indicar una fecha valida.</div>"
    html = html & "<div class='actions'>"
    html = html & "<button class='btn secondary' onclick=""cancelar()"">Cancelar</button>"
    html = html & "<button class='btn primary' onclick=""confirmar()"">Confirmar envio</button>"
    html = html & "</div>"
    html = html & "</div></div>"
    html = html & "<script>"
    html = html & "function confirmar(){var v=document.getElementById('fechaEnvio').value; if(!v){document.getElementById('error').style.display='block'; return;} window.location.href='vba:DATOS_RAC:FECHA|' + encodeURIComponent(v);} "
    html = html & "function cancelar(){window.location.href='vba:DATOS_RAC:CANCELAR';}"
    html = html & "setTimeout(() => { document.getElementById('fechaEnvio').focus(); }, 100);"
    html = html & "</script></body></html>"
    
    html = Replace(html, vbCrLf, "\n")
    html = Replace(html, vbCr, "\n")
    html = Replace(html, vbLf, "\n")
    
    GenerarHTML_CapturaFechaRAC = html
End Function

Public Function GenerarHTML_CapturaComentariosRAC(ByVal comentariosDefault As String) As String
    Dim html As String
    Dim valorDefault As String
    
    valorDefault = HtmlEncode(CStr(Nz(comentariosDefault, "")))
    
    html = "<!DOCTYPE html><html lang='es'><head><meta charset='UTF-8'><style>" & GetCapturaRACCSS() & "</style></head><body>"
    html = html & "<div class='container'>"
    html = html & "<div class='page-title'>Comentarios del RAC</div>"
    html = html & "<div class='card'>"
    html = html & "<label class='label'>Observaciones</label>"
    html = html & "<textarea id='comentarios' class='textarea' rows='6'>" & valorDefault & "</textarea>"
    html = html & "<div id='error' class='error' style='display:none;'>Debe indicar comentarios.</div>"
    html = html & "<div class='actions'>"
    html = html & "<button class='btn secondary' onclick=""cancelar()"">Cancelar</button>"
    html = html & "<button class='btn primary' onclick=""confirmar()"">Confirmar</button>"
    html = html & "</div>"
    html = html & "</div></div>"
    html = html & "<script>"
    html = html & "function confirmar(){var v=document.getElementById('comentarios').value; if(!v||v.trim()===''){document.getElementById('error').style.display='block'; return;} window.location.href='vba:DATOS_RAC:COMENTARIOS|' + encodeURIComponent(v);} "
    html = html & "function cancelar(){window.location.href='vba:DATOS_RAC:CANCELAR';} "
    html = html & "</script></body></html>"
    
    html = Replace(html, vbCrLf, "\n")
    html = Replace(html, vbCr, "\n")
    html = Replace(html, vbLf, "\n")
    
    GenerarHTML_CapturaComentariosRAC = html
End Function

Public Function GenerarHTML_CapturaRechazoCalidad(ByVal comentariosDefault As String, Optional ByVal modo As String = "CAPTURA") As String
    Dim html As String
    Dim valorDefault As String
    Dim modoFinal As String
    Dim titulo As String
    Dim etiqueta As String
    Dim botonPrimario As String
    Dim botonSecundario As String
    Dim readOnlyAttr As String
    
    modoFinal = UCase$(Trim$(modo))
    If modoFinal = "" Then modoFinal = "CAPTURA"
    
    If modoFinal = "LECTURA" Then
        readOnlyAttr = " readonly"
        titulo = "Motivo de rechazo"
        etiqueta = "Comentarios"
        botonPrimario = "Cerrar"
        botonSecundario = ""
    ElseIf modoFinal = "EDICION" Then
        titulo = "Actualizar motivo de rechazo"
        etiqueta = "Comentarios"
        botonPrimario = "Actualizar"
        botonSecundario = "Cancelar"
    Else
        modoFinal = "CAPTURA"
        titulo = "Registrar rechazo"
        etiqueta = "Motivo"
        botonPrimario = "Confirmar rechazo"
        botonSecundario = "Cancelar"
    End If
    
    valorDefault = HtmlEncode(CStr(Nz(comentariosDefault, "")))
    
    html = "<!DOCTYPE html><html lang='es'><head><meta charset='UTF-8'><style>" & GetCapturaRACCSS() & "</style></head><body>"
    html = html & "<div class='container'>"
    html = html & "<div class='page-title'>" & titulo & "</div>"
    html = html & "<div class='card'>"
    html = html & "<label class='label'>" & etiqueta & "</label>"
    html = html & "<textarea id='comentarios' class='textarea' rows='6'" & readOnlyAttr & ">" & valorDefault & "</textarea>"
    html = html & "<div id='error' class='error' style='display:none;'>Debe indicar comentarios.</div>"
    html = html & "<div class='actions'>"
    If botonSecundario <> "" Then html = html & "<button class='btn secondary' onclick=""cancelar()"">" & botonSecundario & "</button>"
    html = html & "<button class='btn primary' onclick=""confirmar()"">" & botonPrimario & "</button>"
    html = html & "</div>"
    html = html & "</div></div>"
    html = html & "<script>"
    If modoFinal = "LECTURA" Then
        html = html & "function confirmar(){window.location.href='vba:DATOS_RECHAZO:CANCELAR';}"
    Else
        html = html & "function confirmar(){var v=document.getElementById('comentarios').value; if(!v||v.trim()===''){document.getElementById('error').style.display='block'; return;} window.location.href='vba:DATOS_RECHAZO:' + encodeURIComponent(v);} "
    End If
    html = html & "function cancelar(){window.location.href='vba:DATOS_RECHAZO:CANCELAR';}"
    html = html & "</script></body></html>"
    
    html = Replace(html, vbCrLf, "\n")
    html = Replace(html, vbCr, "\n")
    html = Replace(html, vbLf, "\n")
    
    GenerarHTML_CapturaRechazoCalidad = html
End Function

Private Function GetCapturaRACCSS() As String
    Dim css As String
    css = "* { box-sizing: border-box; } " & _
          "body { font-family: Telefonica Sans, Arial, sans-serif; margin: 0; padding: 16px; background-color: #F7F9FC; color: #2B3447; font-size: 13px; } " & _
          ".container { max-width: 520px; margin: 0 auto; } " & _
          ".page-title { font-size: 16px; font-weight: bold; margin-bottom: 12px; color: #0066FF; } " & _
          ".card { background: #FFFFFF; border: 1px solid #E6E9F2; border-radius: 8px; padding: 14px; box-shadow: 0 2px 4px rgba(0,0,0,0.04); } " & _
          ".label { display: block; font-weight: bold; margin-bottom: 6px; } " & _
          ".input { width: 100%; padding: 8px 10px; border: 1px solid #D1D5E4; border-radius: 6px; font-size: 13px; } " & _
          ".textarea { width: 100%; padding: 8px 10px; border: 1px solid #D1D5E4; border-radius: 6px; font-size: 13px; resize: vertical; } " & _
          ".actions { display: flex; justify-content: flex-end; gap: 8px; margin-top: 12px; flex-wrap: wrap; } " & _
          ".btn { border: none; border-radius: 6px; padding: 8px 14px; font-weight: bold; cursor: pointer; } " & _
          ".btn.primary { background: #0066FF; color: #FFFFFF; } " & _
          ".btn.secondary { background: #E6E9F2; color: #2B3447; } " & _
          ".btn.word-btn { background: #2D7D46; color: #FFFFFF; } " & _
          ".error { margin-top: 8px; color: #D6786B; font-size: 12px; } "
    GetCapturaRACCSS = css
End Function

Private Function GetHistorialRACCSS() As String
    Dim css As String
    css = "body { font-family: 'Telefonica Sans', Arial, sans-serif; margin: 0; padding: 16px; background-color: #F7F9FC; color: #2B3447; font-size: 13px; } " & _
          ".container { max-width: 900px; margin: 0 auto; } " & _
          ".page-title { font-size: 16px; font-weight: bold; margin-bottom: 12px; color: #0066FF; } " & _
          ".card { background: #FFFFFF; border: 1px solid #E6E9F2; border-left: 6px solid #ADB5BD; border-radius: 8px; padding: 12px 14px; margin-bottom: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.04); } " & _
          ".card.aprobado { border-left-color: #28A745; } " & _
          ".card.comentarios { border-left-color: #FFC107; } " & _
          ".card.pendiente { border-left-color: #ADB5BD; } " & _
          ".card-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; gap: 8px; } " & _
          ".card-title { font-size: 14px; font-weight: bold; color: #2B3447; } " & _
          ".card-actions { display: flex; align-items: center; gap: 8px; } " & _
          ".badge { padding: 3px 8px; border-radius: 12px; font-size: 11px; font-weight: bold; color: #FFFFFF; } " & _
          ".badge.aprobado { background: #28A745; } " & _
          ".badge.comentarios { background: #FFC107; color: #2B3447; } " & _
          ".badge.pendiente { background: #ADB5BD; } " & _
          ".card-meta { display: flex; gap: 12px; flex-wrap: wrap; font-size: 11px; color: #58617A; margin-bottom: 8px; } " & _
          ".card-body { background: #F9FAFC; border: 1px solid #E6E9F2; border-radius: 6px; padding: 10px; white-space: normal; line-height: 1.4; } " & _
          ".word-link { display: inline-flex; align-items: center; justify-content: center; padding: 3px 8px; border-radius: 6px; border: 1px solid #D1D5E4; background: #F2F4FF; color: #0066FF; font-size: 11px; font-weight: bold; text-decoration: none; cursor: pointer; } " & _
          ".word-link:hover { background: #E6E9F2; } " & _
          ".empty-state { text-align: center; padding: 30px; color: #58617A; font-style: italic; background: #FFFFFF; border-radius: 8px; border: 1px dashed #D1D5E4; } "
    GetHistorialRACCSS = css
End Function

Private Function GetAdjuntoIdPorEtapa(ByVal idSolicitud As Long, ByVal etapa As String, Optional ByRef db As DAO.Database) As Long
    Dim sql As String
    Dim params As Object
    Dim rcd As DAO.Recordset
    Dim idAdjunto As Long
    
    On Error GoTo Errores
    
    ' Usar el db proporcionado o obtener uno fresco
    If db Is Nothing Then Set db = getdb()
    
    If idSolicitud <= 0 Then Exit Function
    If Trim(Nz(etapa, "")) = "" Then Exit Function
    
    sql = "SELECT TOP 1 idAdjunto FROM tbAdjuntos WHERE idSolicitud = [p_id] AND etapaWF = [p_etapa] ORDER BY fechaSubida DESC;"
    Set params = CreateObject("Scripting.Dictionary")
    params.Add "p_id", idSolicitud
    params.Add "p_etapa", etapa
    
    Set rcd = RepositorioComun.EjecutarConsulta(sql, params, db)
    If Not rcd Is Nothing Then
        If Not rcd.EOF Then idAdjunto = Nz(rcd!idAdjunto, 0)
    End If
    
    GetAdjuntoIdPorEtapa = idAdjunto
    
LimpiarYSalir:
    On Error Resume Next
    If Not rcd Is Nothing Then rcd.Close
    Exit Function
Errores:
    GetAdjuntoIdPorEtapa = 0
    Resume LimpiarYSalir
End Function

Private Function GetDeltasCSS() As String
    Dim css As String
    css = "body { font-family: 'Telefónica Sans', Arial, sans-serif; margin: 0; padding: 15px; background-color: #F9FAFC; color: #2B3447; font-size: 13px; } " & _
          ".container { max-width: 800px; margin: 0 auto; } " & _
          ".section-title { font-weight: bold; color: #0066FF; margin: 20px 0 10px 0; font-size: 14px; border-bottom: 2px solid #E6E9F2; padding-bottom: 5px; text-transform: uppercase; } " & _
          ".delta-card { background: white; border-radius: 8px; border: 1px solid #D1D5E4; box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 15px; overflow: hidden; } " & _
          ".delta-header { background: #F2F4FF; padding: 10px 15px; border-bottom: 1px solid #D1D5E4; display: flex; justify-content: space-between; align-items: center; } " & _
          ".field-name { font-weight: bold; color: #0066FF; } " & _
          ".meta-info { font-size: 11px; color: #58617A; } " & _
          ".delta-body { padding: 15px; display: flex; align-items: center; gap: 15px; } " & _
          ".value-box { flex: 1; padding: 10px; border-radius: 6px; position: relative; min-height: 40px; } " & _
          ".value-label { font-size: 9px; font-weight: bold; margin-bottom: 4px; color: #8F97AF; } " & _
          ".value-text { font-size: 12px; word-break: break-all; } " & _
          ".value-old { background-color: #F9EEED; border: 1px solid #D6786B; color: #D6786B; } " & _
          ".value-old .value-text { text-decoration: line-through; opacity: 0.8; } " & _
          ".value-new { background-color: #E6F6EF; border: 1px solid #00A551; color: #00A551; } " & _
          ".delta-arrow { color: #D1D5E4; font-size: 20px; } " & _
          ".delta-table { width: 100%; border-collapse: collapse; background: white; border: 1px solid #D1D5E4; border-radius: 8px; overflow: hidden; margin-bottom: 15px; } " & _
          ".delta-table th { text-align: left; background: #F2F4FF; color: #2B3447; font-size: 11px; letter-spacing: 0.5px; text-transform: uppercase; padding: 8px 10px; border-bottom: 1px solid #D1D5E4; } " & _
          ".delta-table td { padding: 8px 10px; border-bottom: 1px solid #EEF1F7; font-size: 12px; vertical-align: top; } " & _
          ".delta-table tr:last-child td { border-bottom: none; } " & _
          ".col-antes { color: #D6786B; } " & _
          ".col-despues { color: #00A551; } " & _
          ".empty-row { text-align: center; color: #58617A; font-style: italic; } " & _
          ".empty-state { text-align: center; padding: 40px; color: #58617A; font-style: italic; background: white; border-radius: 8px; border: 1px dashed #D1D5E4; } "
    GetDeltasCSS = css
End Function

Private Function ConstruirSeccion(ByVal titulo As String, ByVal filas As String, ByVal tieneFilas As Boolean) As String
    Dim html As String
    html = "<div class='section-title'>" & titulo & "</div>"
    html = html & "<table class='delta-table'><thead><tr><th>Campo</th><th>Antes</th><th>Después</th></tr></thead><tbody>"
    If tieneFilas Then
        html = html & filas
    Else
        html = html & "<tr><td colspan='3' class='empty-row'>Sin cambios</td></tr>"
    End If
    html = html & "</tbody></table>"
    ConstruirSeccion = html
End Function

Private Function ObtenerGrupoBloque(ByVal nombre As String) As String
    If InStr(1, nombre, "PROPUESTA", vbTextCompare) > 0 Then
        ObtenerGrupoBloque = "PROPUESTA"
    Else
        ObtenerGrupoBloque = "IMPACTO"
    End If
End Function

Private Function ValoresDiferentes(ByVal valorInicial As Variant, ByVal valorFinal As Variant) As Boolean
    Dim a As String
    Dim b As String
    a = CStr(Nz(valorInicial, ""))
    b = CStr(Nz(valorFinal, ""))
    ValoresDiferentes = (a <> b)
End Function

Private Function ValorParaHtml(ByVal valor As Variant) As String
    Dim texto As String
    If IsNull(valor) Then
        texto = "(Vacío)"
    ElseIf VarType(valor) = vbBoolean Then
        If valor Then
            texto = "Sí"
        Else
            texto = "No"
        End If
    Else
        texto = CStr(valor)
        If Trim(texto) = "" Then
            texto = "(Vacío)"
        End If
    End If
    texto = Replace(texto, vbCrLf, "<br>")
    texto = Replace(texto, vbLf, "<br>")
    ValorParaHtml = HtmlEncode(texto)
End Function

Private Function HtmlEncode(ByVal texto As String) As String
    Dim s As String
    s = texto
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&#39;")
    HtmlEncode = s
End Function

Private Function NormalizarNombreBloque(ByVal nombre As String) As String
    ' Mapeo amigable de nombres de tablas/bloques
    Select Case nombre
        Case "PC_Propuesta": NormalizarNombreBloque = "Datos de la Propuesta"
        Case "PC_Impacto": NormalizarNombreBloque = "Impactos de la Propuesta"
        Case "CDCA_Impacto": NormalizarNombreBloque = "Impactos CD/CA"
        Case "CDCASUB_Impacto": NormalizarNombreBloque = "Impactos CD/CA (Sub)"
        Case Else: NormalizarNombreBloque = nombre
    End Select
End Function

Private Function NormalizarNombreCampo(ByVal nombre As String) As String
    ' Mapeo amigable de campos técnicos
    Select Case nombre
        Case "descripcionPropuestaCambio": NormalizarNombreCampo = "Descripción del Cambio"
        Case "numPlanoEspecificacion": NormalizarNombreCampo = "Plano / Especificación"
        Case "descripcionMaterialAfectado": NormalizarNombreCampo = "Material Afectado"
        Case "afectaPrestaciones": NormalizarNombreCampo = "Afecta Prestaciones"
        Case "afectaSeguridad": NormalizarNombreCampo = "Afecta Seguridad"
        Case "afectaFiabilidad": NormalizarNombreCampo = "Afecta Fiabilidad"
        Case "afectaVidaUtil": NormalizarNombreCampo = "Afecta Vida Útil"
        Case "afectaMedioambiente": NormalizarNombreCampo = "Afecta Medioambiente"
        Case Else: NormalizarNombreCampo = nombre
    End Select
End Function













