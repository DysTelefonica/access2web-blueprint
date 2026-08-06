Attribute VB_Name = "Automiatizacion"

Option Compare Database
Option Explicit

' =========================================================================================
' WRAPPER (PUNTO DE ENTRADA ÚNICO)
' Mantiene la firma original para retroactividad, pero decide la lógica interna.
' =========================================================================================
Public Function EnviarCorreoNoAutomatico( _
                                        p_EnumTipoEnvioCorreo As EnumTipoEnvioCorreo, _
                                        Optional p_IDSolicitud As String, _
                                        Optional p_Solicitud As solicitud, _
                                        Optional p_DESTINATARIOS As String, _
                                        Optional p_DestinatariosConCopia As String, _
                                        Optional p_asunto As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Correo
    
    ' 1. Determinar el objeto solicitud si no viene dado (necesario para ambos modos para extraer datos)
    If p_Solicitud Is Nothing Then
        Set p_Solicitud = constructor.getSolicitud(p_IDSolicitud:=p_IDSolicitud, p_Error:=p_Error)
        If p_Error <> "" Then Exit Function
    End If
    
    ' 2. Decisión de camino
    If m_ObjEntorno.ModoNuevoEnvioCorreo Then
        ' CAMINO NUEVO (Desacoplado)
        ' Extraemos los datos aquí y llamamos al genérico
        
        Dim sAsunto As String
        Dim sDestinatarios As String
        Dim sRecursos As String
        Dim sPlantilla As String
        Dim sAccion As String
        Dim sAdjunto As String
        
        ' Extracción de datos (Hidratación)
        If p_asunto = "" Then sAsunto = p_Solicitud.AsuntoCalculado(p_EnumTipoEnvioCorreo) Else sAsunto = p_asunto
        If p_DESTINATARIOS = "" Then sDestinatarios = p_Solicitud.DestinatariosCalculados Else sDestinatarios = p_DESTINATARIOS
        
        sRecursos = p_Solicitud.ColParametrosParaHTMLTexto
        sPlantilla = p_Solicitud.NombrePlantillaParaCorreoCalculada(p_EnumTipoEnvioCorreo)
        sAccion = p_Solicitud.AccionParaCorreoCalculada(p_EnumTipoEnvioCorreo)
        
        If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.EnvioExcel Then
            sAdjunto = p_Solicitud.URLPlantillaExcel
        End If
        
        ' Llamada al nuevo método genérico
        ' NOTA: Pasamos p_Solicitud.IDSolicitud real para mantener el vínculo histórico si existe,
        ' pero la función genérica está preparada para aceptar "-1" si quisiéramos.
        ' Al usar el wrapper, mantenemos el ID real.
        Set EnviarCorreoNoAutomatico = EnviarCorreoGenerico( _
                                            p_DESTINATARIOS:=sDestinatarios, _
                                            p_asunto:=sAsunto, _
                                            p_NombrePlantilla:=sPlantilla, _
                                            p_CadenaRecursos:=sRecursos, _
                                            p_Accion:=sAccion, _
                                            p_TipoCorreo:=p_EnumTipoEnvioCorreo, _
                                            p_DestinatariosCopia:=p_DestinatariosConCopia, _
                                            p_URLAdjunto:=sAdjunto, _
                                            p_IDSolicitudVinculada:=p_Solicitud.IDSolicitud, _
                                            p_Error:=p_Error)
                                            
    Else
        ' CAMINO VIEJO (Legacy)
        Set EnviarCorreoNoAutomatico = EnviarCorreoLegacy( _
                                            p_EnumTipoEnvioCorreo, _
                                            p_IDSolicitud, _
                                            p_Solicitud, _
                                            p_DESTINATARIOS, _
                                            p_DestinatariosConCopia, _
                                            p_asunto, _
                                            p_Error)
    End If

End Function

' =========================================================================================
' NUEVA FUNCIÓN DESACOPLADA (CORE)
' No depende de un objeto Solicitud. Recibe primitivos.
' =========================================================================================


Public Function EnviarCorreoGenerico( _
                                    p_DESTINATARIOS As String, _
                                    p_asunto As String, _
                                    p_NombrePlantilla As String, _
                                    p_CadenaRecursos As String, _
                                    p_Accion As String, _
                                    p_TipoCorreo As EnumTipoEnvioCorreo, _
                                    Optional p_DestinatariosCopia As String = "", _
                                    Optional p_URLAdjunto As String = "", _
                                    Optional p_IDSolicitudVinculada As String = CONST_IDSOLICITUD_PRUEBA, _
                                    Optional p_FechaOrden As Variant = Null, _
                                    Optional ByRef p_Error As String _
                                    ) As Correo
    
    Dim m_Correo As Correo
    Dim m_CorreoOp As CorreoOperaciones
    
    On Error GoTo errores
    
    Set m_Correo = New Correo
    With m_Correo
        .DESTINATARIOS = p_DESTINATARIOS
        .DestinatariosConCopia = p_DestinatariosCopia
        .Asunto = p_asunto
        .nombrePlantilla = p_NombrePlantilla
        .VersionPlantilla = m_ObjEntorno.Configuracion.VersionPlantillasHTML
        .cadenaRecursos = p_CadenaRecursos
        .Accion = p_Accion
        .URLAdjunto = p_URLAdjunto
        
        ' Usamos el -1 (o el ID que pases) para cumplir con la tabla numérica
        .IDSolicitud = p_IDSolicitudVinculada
        
        .tipoCorreo = p_TipoCorreo
        
        ' Lógica para la fecha programada
        If IsDate(p_FechaOrden) Then
            .FechaOrdenEnvio = p_FechaOrden
            .Programado = "Sí"
            .DesencadenadoPor = m_ObjUsuarioConectado.UsuarioRed & " (Prueba Programada)"
        Else
            .Programado = "No"
            .DesencadenadoPor = m_ObjUsuarioConectado.UsuarioRed & " (Prueba Manual)"
        End If
        
    End With
    
    Set m_CorreoOp = New CorreoOperaciones
    With m_CorreoOp
        Set .Correo = m_Correo
        .Registrar , p_Error
    End With
    
    Set EnviarCorreoGenerico = m_Correo
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en EnviarCorreoGenerico: " & Err.Description
    End If
End Function

' =========================================================================================
' FUNCIÓN LEGACY (Antigua EnviarCorreoNoAutomatico renombrada)
' =========================================================================================
Private Function EnviarCorreoLegacy( _
                                        p_EnumTipoEnvioCorreo As EnumTipoEnvioCorreo, _
                                        Optional p_IDSolicitud As String, _
                                        Optional p_Solicitud As solicitud, _
                                        Optional p_DESTINATARIOS As String, _
                                        Optional p_DestinatariosConCopia As String, _
                                        Optional p_asunto As String, _
                                        Optional ByRef p_Error As String _
                                        ) As Correo
    Dim m_Correo As Correo
    Dim m_CorreoOp As CorreoOperaciones
    
    On Error GoTo errores
    
    ' (Lógica original intacta)
    If p_EnumTipoEnvioCorreo = Empty Then
        p_Error = "No se ha podido determinar el tipo de correo"
        Err.Raise 1000
    End If
    If p_Solicitud Is Nothing Then
        Set p_Solicitud = constructor.getSolicitud(p_IDSolicitud:=p_IDSolicitud, p_Error:=p_Error)
        If p_Error <> "" Then
            p_Error = "No se ha podido determinar la solicitud"
            Err.Raise 1000
        End If
    End If
    
    Set m_Correo = New Correo
    With m_Correo
        If p_EnumTipoEnvioCorreo = EnumTipoEnvioCorreo.EnvioExcel Then
            .URLAdjunto = p_Solicitud.URLPlantillaExcel
        End If
        If p_DESTINATARIOS = "" Then
            p_DESTINATARIOS = p_Solicitud.DestinatariosCalculados
        End If
        .DESTINATARIOS = p_DESTINATARIOS
        .DestinatariosConCopia = p_DestinatariosConCopia
        If p_asunto = "" Then
            p_asunto = p_Solicitud.AsuntoCalculado(p_EnumTipoEnvioCorreo)
        End If
        .Asunto = p_asunto
        .nombrePlantilla = p_Solicitud.NombrePlantillaParaCorreoCalculada(p_EnumTipoEnvioCorreo)
        .VersionPlantilla = m_ObjEntorno.Configuracion.VersionPlantillasHTML
        .cadenaRecursos = p_Solicitud.ColParametrosParaHTMLTexto
        .IDSolicitud = p_Solicitud.IDSolicitud
        .Accion = p_Solicitud.AccionParaCorreoCalculada(p_EnumTipoEnvioCorreo)
        .Programado = "No"
        .tipoCorreo = p_EnumTipoEnvioCorreo
    End With
    
    Set m_CorreoOp = New CorreoOperaciones
    With m_CorreoOp
        Set .Correo = m_Correo
        .Registrar , p_Error
    End With
    
    Set EnviarCorreoLegacy = m_Correo
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnviarCorreoLegacy ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

