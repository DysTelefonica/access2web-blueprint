VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NotificacionServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Public Sub EnviarNotificacion(ByRef vm As SolicitudViewModel, ByVal asunto As String, ByVal destinatarioEmail As String, ByVal copiaEmail As String, Optional ByVal nombreEstadoNotif As String = "")
    ' RESPONSABILIDAD: Enviar una notificación de cambio de estado. Es "tonto",
    '                  recibe toda la información que necesita para operar.
    Dim objCorreo As New Correo
    Dim cuerpo As String
    Dim devNotice As String
    Dim m_Originador As String
    On Error GoTo Errores
    
    ' --- 1. Validaciones Preliminares ---
    If vm Is Nothing Then Err.Raise 513, "EnviarNotificacion", "El ViewModel es nulo."
    If Trim(Nz(destinatarioEmail, "")) = "" Then Err.Raise 513, "EnviarNotificacion", "El destinatario no puede estar vacío."
    
    ' --- 2. Gestión de Entorno de Desarrollo ---
    devNotice = ""
    If Application.TempVars("EnDesarrollo") = "Sí" Then
        devNotice = "<div class='dev-notice'>" & _
                    "<b>AVISO DE DESARROLLO:</b> Este correo ha sido redirigido para pruebas." & _
                    "<br>En producción, se habría enviado a: <b>" & destinatarioEmail & "</b></div>"
        destinatarioEmail = m_ObjUsuarioReal.CorreoUsuario
        copiaEmail = "" ' Vaciamos la copia para no recibirlo duplicado
    End If
    'destinatarioEmail = "andres.romandelperal@telefonica.com"
    If copiaEmail <> "andres.romandelperal@telefonica.com" And copiaEmail <> "" Then
        copiaEmail = ""
    End If
cuerpo = GenerarCuerpoHTML(vm, devNotice, "", nombreEstadoNotif)
    
    ' --- 4. Construcción y Envío del Objeto Correo ---
    With objCorreo
        .IDCorreo = CorreoRepositorio.getSiguienteIDCorreo(getdb())
        .Aplicacion = "CONDOR"
        .asunto = asunto
        .Destinatarios = destinatarioEmail
        .DestinatariosConCopia = copiaEmail
        .DestinatariosConCopiaOculta = "andres.romandelperal@telefonica.com" ' CCO Fijo
        .CuerpoHTML = True
        .cuerpo = cuerpo
        .URLAdjunto = "" ' Ya no hay imagen que adjuntar
        .FechaGrabacion = Now() ' <<< AÑADIR ESTA LÍNEA
        If g_blnImpersonando Then
            m_Originador = m_ObjUsuarioReal.nombre & " (en nombre de " & m_ObjUsuarioActivo.nombre & ")"
        Else
             m_Originador = m_ObjUsuarioActivo.nombre
        End If
        If Len(m_Originador) > 50 Then
            m_Originador = Left(m_Originador, 50)
        End If
        .Originador = m_Originador
    End With
    
    Call ValidarCorreo(objCorreo)
    Call CorreoRepositorio.GuardarCorreo(objCorreo)

LimpiarYSalir:
    On Error Resume Next
    Set objCorreo = Nothing
    Exit Sub
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NotificacionServicio.EnviarNotificacion"
    errObj.Raise
End Sub
Public Sub EnviarNotificacionConDetalle(ByRef vm As SolicitudViewModel, ByVal asunto As String, ByVal destinatarioEmail As String, ByVal copiaEmail As String, ByVal extraHtml As String)
    Dim objCorreo As New Correo
    Dim cuerpo As String
    Dim devNotice As String
    Dim m_Originador As String
    On Error GoTo Errores
    
    If vm Is Nothing Then Err.Raise 513, "EnviarNotificacionConDetalle", "El ViewModel es nulo."
    If Trim(Nz(destinatarioEmail, "")) = "" Then Err.Raise 513, "EnviarNotificacionConDetalle", "El destinatario no puede estar vacío."
    
    devNotice = ""
    If Application.TempVars("EnDesarrollo") = "Sí" Then
        devNotice = "<div class='dev-notice'><b>AVISO DE DESARROLLO:</b> Este correo ha sido redirigido para pruebas.<br>En producción, se habría enviado a: <b>" & destinatarioEmail & "</b></div>"
        destinatarioEmail = m_ObjUsuarioReal.CorreoUsuario
        copiaEmail = ""
    End If
    If copiaEmail <> "andres.romandelperal@telefonica.com" And copiaEmail <> "" Then
        copiaEmail = ""
    End If
    
    cuerpo = GenerarCuerpoHTML(vm, devNotice, extraHtml)
    
    With objCorreo
        .IDCorreo = CorreoRepositorio.getSiguienteIDCorreo(getdb())
        .Aplicacion = "CONDOR"
        .asunto = asunto
        .Destinatarios = destinatarioEmail
        .DestinatariosConCopia = copiaEmail
        .DestinatariosConCopiaOculta = "andres.romandelperal@telefonica.com"
        .CuerpoHTML = True
        .cuerpo = cuerpo
        .URLAdjunto = ""
        .FechaGrabacion = Now()
        If g_blnImpersonando Then
            m_Originador = m_ObjUsuarioReal.nombre & " (en nombre de " & m_ObjUsuarioActivo.nombre & ")"
        Else
            m_Originador = m_ObjUsuarioActivo.nombre
        End If
        If Len(m_Originador) > 50 Then
            m_Originador = Left(m_Originador, 50)
        End If
        .Originador = m_Originador
    End With
    
    Call ValidarCorreo(objCorreo)
    Call CorreoRepositorio.GuardarCorreo(objCorreo)
    
LimpiarYSalir:
    On Error Resume Next
    Set objCorreo = Nothing
    Exit Sub
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NotificacionServicio.EnviarNotificacionConDetalle"
    errObj.Raise
End Sub

Public Sub EnviarNotificacionEliminacion(ByRef vm As SolicitudViewModel, ByVal asunto As String, ByVal destinatarioEmail As String, ByVal copiaEmail As String)
    Dim objCorreo As New Correo
    Dim cuerpo As String
    Dim devNotice As String
    Dim m_Originador As String
    Dim nombreActor As String
    
    On Error GoTo Errores
    
    If vm Is Nothing Then Err.Raise 513, "EnviarNotificacionEliminacion", "El ViewModel es nulo."
    If Trim(Nz(destinatarioEmail, "")) = "" Then Err.Raise 513, "EnviarNotificacionEliminacion", "El destinatario no puede estar vacío."
    
    devNotice = ""
    If Application.TempVars("EnDesarrollo") = "Sí" Then
        devNotice = "<div class='dev-notice'><b>AVISO DE DESARROLLO:</b> Este correo ha sido redirigido para pruebas.<br>En producción, se habría enviado a: <b>" & destinatarioEmail & "</b></div>"
        destinatarioEmail = m_ObjUsuarioReal.CorreoUsuario
        copiaEmail = ""
    End If
    If copiaEmail <> "andres.romandelperal@telefonica.com" And copiaEmail <> "" Then
        copiaEmail = ""
    End If
    
    ' Determinar el nombre del actor para el cuerpo del mensaje
    If g_blnImpersonando Then
        nombreActor = m_ObjUsuarioReal.nombre & " (en nombre de " & m_ObjUsuarioActivo.nombre & ")"
    Else
        nombreActor = m_ObjUsuarioActivo.nombre
    End If
    
    cuerpo = GenerarCuerpoHTML_Eliminacion(vm, devNotice, nombreActor)
    
    With objCorreo
        .IDCorreo = CorreoRepositorio.getSiguienteIDCorreo(getdb())
        .Aplicacion = "CONDOR"
        .asunto = asunto
        .Destinatarios = destinatarioEmail
        .DestinatariosConCopia = copiaEmail
        .DestinatariosConCopiaOculta = "andres.romandelperal@telefonica.com"
        .CuerpoHTML = True
        .cuerpo = cuerpo
        .URLAdjunto = ""
        .FechaGrabacion = Now()
        .Originador = Left(nombreActor, 50)
    End With
    
    Call ValidarCorreo(objCorreo)
    Call CorreoRepositorio.GuardarCorreo(objCorreo)
    
LimpiarYSalir:
    On Error Resume Next
    Set objCorreo = Nothing
    Exit Sub
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NotificacionServicio.EnviarNotificacionEliminacion"
    errObj.Raise
End Sub

Private Function GenerarCuerpoHTML_Eliminacion(ByRef vm As SolicitudViewModel, ByVal devNotice As String, ByVal nombreActor As String) As String
    ' RESPONSABILIDAD: Cuerpo específico para eliminación (Spec-101)
    Dim html As String
    
    html = "<!DOCTYPE html><html><head><style>" & m_ObjEntorno.css & "</style></head>"
    html = html & "<body><div class='container'>"
    
    html = html & "<div class='header'><h2>Notificacion de CONDOR</h2></div>"
    html = html & "<div class='content'>"
    html = html & devNotice
    
    ' MENSAJE DE ELIMINACIÓN EN ROJO
    html = html & "<p style='color: #D6786B; font-weight: bold; font-size: 1.1em;'>Solicitud ELIMINADA por " & nombreActor & "</p>"
    
    ' --- Resumen del Expediente ---
    Dim tablaExp As String
    tablaExp = "<table class='data-table'>" & _
               "<tr><th>Nemotecnico:</th><td>" & vm.Expediente.Nemotecnico & "</td></tr>" & _
               "<tr><th>Titulo:</th><td>" & vm.Expediente.titulo & "</td></tr>" & _
               "<tr><th>Contratista:</th><td>" & vm.nombreContratistaPrincipal & "</td></tr>" & _
               "<tr><th>Jefe Proyecto:</th><td>" & vm.JefeProyecto & "</td></tr>" & _
               "<tr><th>Calidad:</th><td>" & vm.TecnicoCalidad & "</td></tr>" & _
               "</table>"
    html = html & GenerarTarjeta("Resumen del Expediente", tablaExp)

    ' --- Detalles de la Solicitud ---
    Dim tablaSol As String
    tablaSol = "<table class='data-table'>" & _
               "<tr><th>Estado Anterior:</th><td class='status-highlight'>" & vm.EstadoNombre & "</td></tr>" & _
               "<tr><th>Codigo:</th><td>" & vm.Solicitud.codigoSolicitud & "</td></tr>" & _
               "<tr><th>Tipo:</th><td>" & vm.TipoSolicitudNombre & "</td></tr>" & _
               "<tr><th>Fecha Creacion:</th><td>" & Format(vm.Solicitud.fechaCreacion, "dd/mm/yyyy hh:nn") & "</td></tr>" & _
               "</table>"
    html = html & GenerarTarjeta("Detalles de la Solicitud", tablaSol)
    
    html = html & "</div>" ' Cierre de .content
    
    html = html & "<div class='footer'>"
    html = html & "  Este es un correo generado automáticamente por la aplicación CONDOR.<br>"
    html = html & "  Por favor, no responda a este mensaje."
    html = html & "</div>"
    
    html = html & "</div></body></html>"
    
    GenerarCuerpoHTML_Eliminacion = html
End Function

Private Function GenerarCuerpoHTML(ByRef vm As SolicitudViewModel, ByVal devNotice As String, Optional ByVal extraHtml As String = "", Optional ByVal nombreEstadoNotif As String = "") As String
    ' RESPONSABILIDAD: Construye el cuerpo del correo en formato HTML, con un diseño moderno de tarjetas.
    ' REFACTOR: Spec-099 Estructura de 4 partes.
    Dim html As String
    
    ' --- 0. Inicialización y CSS ---
    html = "<!DOCTYPE html><html><head><style>" & m_ObjEntorno.css & "</style></head>"
    html = html & "<body><div class='container'>"
    
    ' --- Cabecera ---
    html = html & "<div class='header'><h2>Notificacion de CONDOR</h2></div>"
    html = html & "<div class='content'>"
    html = html & devNotice
    html = html & "<p>La solicitud <b>" & vm.Solicitud.codigoSolicitud & "</b> ha cambiado de estado y requiere su atencion.</p>"

    ' --- PARTE 1: El "Por Que" (Contexto Dinamico) ---
    Dim idEstado As Long
    idEstado = vm.Solicitud.idEstadoInterno

    ' Escenario SUBSANACION (Verde) - Cuando técnico ha subsanado (estados 6,8 o 4 con rechazo previo)
    If vm.vm_tecnicoSubsanacionRealizada And (idEstado = 6 Or idEstado = 8) Then
        html = html & GenerarTarjetaAlerta("SUBSANACION REALIZADA", vm.vm_tecnicoComentariosSubsanacion, "#2ECC71")
    ElseIf idEstado = 4 And vm.vm_revisionCalidadEstado = "RECHAZADO" Then
        html = html & GenerarTarjetaAlerta("SUBSANACION REALIZADA", "Se han aplicado correcciones tras el rechazo de calidad.", "#2ECC71")
    
    ' Escenario RECHAZO (Rojo) - Solo cuando está en Desarrollo Técnico (3)
    ElseIf vm.vm_revisionCalidadEstado = "RECHAZADO" And idEstado = 3 Then
        html = html & GenerarTarjetaAlerta("MOTIVO DEL RECHAZO", vm.vm_revisionCalidadComentarios, "#D6786B")
    
    ' Escenario VALIDACION RAC (Naranja) - SPEC-100
    ' Visible en fases técnicas (3,4) si ha sido devuelto, o en fases de revisión (6,8) si hay comentarios pendientes
    ElseIf (vm.vm_racValidacionEstado = "RECHAZADO" Or vm.vm_racValidacionEstado = "COMENTARIOS") And (idEstado = 3 Or idEstado = 4 Or idEstado = 6 Or idEstado = 8) Then
        html = html & GenerarTarjetaAlerta("OBSERVACIONES RAC", vm.vm_racValidacionComentarios, "#E67E22")
    End If

    ' --- PARTE 2: Resumen del Expediente (Fija) ---
    Dim tablaExp As String
    tablaExp = "<table class='data-table'>" & _
               "<tr><th>Nemotecnico:</th><td>" & vm.Expediente.Nemotecnico & "</td></tr>" & _
               "<tr><th>Titulo:</th><td>" & vm.Expediente.titulo & "</td></tr>" & _
               "<tr><th>Contratista:</th><td>" & vm.nombreContratistaPrincipal & "</td></tr>" & _
               "<tr><th>Jefe Proyecto:</th><td>" & vm.JefeProyecto & "</td></tr>" & _
               "<tr><th>Calidad:</th><td>" & vm.TecnicoCalidad & "</td></tr>" & _
               "</table>"
    html = html & GenerarTarjeta("Resumen del Expediente", tablaExp)

    ' --- PARTE 3: Detalles de la Solicitud (Fija) ---
    Dim tablaSol As String
    tablaSol = "<table class='data-table'>" & _
               "<tr><th>Estado Actual:</th><td class='status-highlight'>" & vm.EstadoNombre & "</td></tr>" & _
               "<tr><th>Codigo:</th><td>" & vm.Solicitud.codigoSolicitud & "</td></tr>" & _
               "<tr><th>Tipo:</th><td>" & vm.TipoSolicitudNombre & "</td></tr>" & _
               "<tr><th>Fecha:</th><td>" & Format(vm.Solicitud.fechaCreacion, "dd/mm/yyyy hh:nn") & "</td></tr>" & _
               "<tr><th>Actor:</th><td>" & vm.Solicitud.usuarioModificacion & "</td></tr>" & _
               "</table>"
    html = html & GenerarTarjeta("Detalles de la Solicitud", tablaSol)

    ' --- PARTE 4: Proximos Pasos / Checklist (Nueva) ---
    html = html & GenerarChecklistSiguienteFase(vm, nombreEstadoNotif)

    ' --- Extra HTML (Legacy support) ---
    If Trim$(Nz(extraHtml, "")) <> "" Then
        html = html & extraHtml
    End If

    html = html & "</div>" ' Cierre de .content
    
    ' --- Pie de página ---
    html = html & "<div class='footer'>"
    html = html & "  Este es un correo generado automáticamente por la aplicación CONDOR.<br>"
    html = html & "  Por favor, no responda a este mensaje."
    html = html & "</div>"
    
    html = html & "</div></body></html>" ' Cierre de .container
    
    GenerarCuerpoHTML = html
End Function

' --- Helper Functions for Spec-099 ---

Private Function GenerarTarjetaAlerta(ByVal titulo As String, ByVal contenido As String, ByVal colorHex As String) As String
    Dim html As String
    Dim bgColor As String
    
    ' Determinar color de fondo suave basado en el color principal
    Select Case colorHex
        Case "#D6786B": bgColor = "#F9EEED" ' Rojo suave
        Case "#2ECC71": bgColor = "#EAFAF1" ' Verde suave
        Case "#3498DB": bgColor = "#EBF5FB" ' Azul suave
        Case "#E67E22": bgColor = "#FDEBD0" ' Naranja suave
        Case Else: bgColor = "#FFFFFF"
    End Select

    html = "<div class='card' style='border-color: " & colorHex & ";'>"
    html = html & "  <div class='card-header' style='background-color: " & colorHex & "; color: white; font-weight: bold;'>" & titulo & "</div>"
    html = html & "  <div class='card-body' style='background-color: " & bgColor & "; color: #333333;'>"
    html = html & "    " & Replace(Nz(contenido, ""), vbCrLf, "<br>")
    html = html & "  </div>"
    html = html & "</div>"
    GenerarTarjetaAlerta = html
End Function

Private Function GenerarTarjeta(ByVal titulo As String, ByVal contenidoHtml As String) As String
    Dim html As String
    html = "<div class='card'>"
    html = html & "  <div class='card-header'>" & titulo & "</div>"
    html = html & "  <div class='card-body'>"
    html = html & "    " & contenidoHtml
    html = html & "  </div>"
    html = html & "</div>"
    GenerarTarjeta = html
End Function

Private Function GenerarChecklistSiguienteFase(ByRef vm As SolicitudViewModel, Optional ByVal nombreEstadoNotif As String = "") As String
    Dim html As String
    Dim items As String
    items = ""

    ' Si se proporciona el nombre del estado de notificación, usarlo directamente
    Dim itemsNotif As String
    If nombreEstadoNotif <> "" Then
        itemsNotif = ObtenerProximosPasosPorEstado(nombreEstadoNotif, vm.Solicitud.tipoSolicitud)
    Else
        ' Lógica basada en el estado actual (vm.Solicitud.idEstadoInterno)
        ' Spec-150: Usa helper unificado para todos los estados
        itemsNotif = GetItemsChecklistPorEstado(vm.Solicitud.idEstadoInterno, vm.Solicitud.tipoSolicitud, vm.vm_revisionCalidadEstado, vm.vm_racValidacionEstado)
    End If
    
    If itemsNotif <> "" Then
        Dim arrNotif As Variant
        arrNotif = Split(itemsNotif, vbCrLf)
        Dim j As Long
        For j = LBound(arrNotif) To UBound(arrNotif)
            If Trim$(arrNotif(j)) <> "" Then
                items = items & "<li>" & arrNotif(j) & "</li>"
            End If
        Next j
    End If

    If items <> "" Then
        html = "<div class='card'>"
        html = html & "  <div class='card-header'>Proximos Pasos (Checklist)</div>"
        html = html & "  <div class='card-body'>"
        html = html & "    <ul style='margin-top: 0; margin-bottom: 0; padding-left: 20px;'>" & items & "</ul>"
        html = html & "  </div>"
        html = html & "</div>"
    End If

    GenerarChecklistSiguienteFase = html
End Function

Public Function ObtenerDestinatariosPorEstado(ByVal estado As String, ByRef vm As SolicitudViewModel) As String
    Dim dest As String
    Dim sep As String
    
    Select Case UCase$(estado)
        Case "REGISTRO"
            dest = vm.JefeProyectoEmail
            If vm.TecnicoCalidadEmail <> "" Then
                dest = dest & ";" & vm.TecnicoCalidadEmail
            End If
            
        Case "MODIFICACION", "PENDIENTE DE VALIDACION"
            dest = vm.TecnicoCalidadEmail
            
        Case "RECHAZADO_POR_CALIDAD", "RECHAZADO_CALIDAD"
            dest = vm.JefeProyectoEmail
            
        Case "VALIDADO_POR_CALIDAD", "VALIDADO_CALIDAD"
            dest = vm.JefeProyectoEmail
            
        Case "REVISION_RAC"
            dest = vm.JefeProyectoEmail
            If vm.TecnicoCalidadEmail <> "" Then
                dest = dest & ";" & vm.TecnicoCalidadEmail
            End If
            
        Case "RECHAZO_RAC"
            dest = vm.TecnicoCalidadEmail
            
        Case "ACEPTACION_RAC"
            dest = vm.TecnicoCalidadEmail
            
        Case "APROBADO_RAC"
            dest = vm.JefeProyectoEmail
            If vm.TecnicoCalidadEmail <> "" Then
                dest = dest & ";" & vm.TecnicoCalidadEmail
            End If
            
        Case "APROBACIÓN"
            dest = vm.JefeProyectoEmail
            If vm.TecnicoCalidadEmail <> "" Then
                dest = dest & ";" & vm.TecnicoCalidadEmail
            End If
            
        Case Else
            dest = vm.JefeProyectoEmail
    End Select
    
    ObtenerDestinatariosPorEstado = dest
End Function

Public Function ObtenerProximosPasosPorEstado(ByVal estado As String, ByVal tipoSolicitud As String) As String
    Dim pasos As String
    
    Select Case UCase$(estado)
        Case "REGISTRO"
            pasos = "Completar en Desarrollo Técnico el formulario: apartados General, Detalle y Descripción, y Motivos (responsable: JP)."
            
        Case "MODIFICACION", "PENDIENTE DE VALIDACION"
            pasos = "Validar la información indicada por el JP; completar los apartados Aprobación Suministrador y RAC."
            
        Case "RECHAZADO_POR_CALIDAD", "RECHAZADO_CALIDAD"
            pasos = "Modificar el formulario con las indicaciones de Calidad."
            
        Case "VALIDADO_POR_CALIDAD", "VALIDADO_CALIDAD"
            pasos = "El formulario ha sido aprobado. Próximo paso: envío del borrador al RAC."
            
        Case "REVISION_RAC"
            pasos = "Aprobación del RAC y formalización."
            
        Case "RECHAZO_RAC"
            pasos = "Revisar el motivo del rechazo del RAC y tomar las acciones correspondientes."
            
        Case "ACEPTACION_RAC"
            pasos = "El RAC ha aceptado la solicitud. Proceder con la formalización."
            
        Case "APROBADO_RAC"
            pasos = "Adjuntar documento de Cierre firmado."
            
        Case "APROBACIÓN"
            pasos = "El expediente ha sido aprobado y formalizado."
            
        Case Else
            pasos = "Consultar el estado de la solicitud en el sistema CONDOR."
    End Select
    
    ObtenerProximosPasosPorEstado = pasos
End Function

Private Sub ValidarCorreo(ByRef objCorreo As Correo)
    On Error GoTo Errores
    
    If objCorreo Is Nothing Then Err.Raise 513, "Validación", "El objeto Correo es nulo."

    If objCorreo.IDCorreo <= 0 Then Err.Raise 513, "Validación", "Error interno: ID de correo no válido."
    If Trim(Nz(objCorreo.Aplicacion, "")) = "" Then Err.Raise 513, "Validación", "Falta la aplicación origen."
    If Trim(Nz(objCorreo.Destinatarios, "")) = "" Then Err.Raise 513, "Validación", "El destinatario es obligatorio."
    If Trim(Nz(objCorreo.asunto, "")) = "" Then Err.Raise 513, "Validación", "El asunto es obligatorio."
    If Trim(Nz(objCorreo.Originador, "")) = "" Then Err.Raise 513, "Validación", "El remitente (Originador) es obligatorio."

    Exit Sub
Errores:
    ' CORRECCIÓN: Encapsulamiento en CondorError
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "NotificacionServicio.ValidarCorreo"
    If Not objCorreo Is Nothing Then
        errObj.AddToCallStack "Destinatario: " & objCorreo.Destinatarios
    End If
    errObj.Raise
End Sub





