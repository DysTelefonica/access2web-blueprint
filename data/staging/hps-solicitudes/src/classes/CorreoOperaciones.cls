VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CorreoOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Correo As Correo
Private m_Error As String
Public Error As String

Public Function Eliminar(Optional ByRef p_Error As String) As String
    
    Dim m_SQL As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    With Me.Correo
        If .IDCORREO = "" Then
            Exit Function
        End If
        m_SQL = "DELETE * FROM TbCorreosEnviados " & _
                "WHERE IDCorreo=" & .IDCORREO & ";"
        getdb().Execute m_SQL
    End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreoOperaciones.Eliminar ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function
Public Function MotivoNoOK( _
                                Optional ByRef p_Error As String _
                                ) As String
        
    
   
    On Error GoTo errores
        
    With Me.Correo
        
        
        If Not IsDate(.FechaOrdenEnvio) And .Programado = "Sí" Then
            MotivoNoOK = "No se conoce la fecha de orden de envío"
            Exit Function
        End If
        If InStr(1, .DESTINATARIOS, "@") = 0 _
            And InStr(1, .DestinatariosConCopia, "@") = 0 _
            And InStr(1, .DestinatariosConCopiaOculta, "@") = 0 Then
            MotivoNoOK = "No hay destinatarios"
            Exit Function
        End If
        If .Asunto = "" Then
            MotivoNoOK = "Ha de indicar el asunto"
            Exit Function
        End If
        If .nombrePlantilla = "" Then
            MotivoNoOK = "Ha de indicar el NombrePlantilla"
            Exit Function
        End If
        If .VersionPlantilla = "" Then
            MotivoNoOK = "Ha de indicar la versión de la plantilla"
            Exit Function
        End If
        If .URLAdjunto <> "" Then
            If Not fso.FileExists(.URLAdjunto) Then
                MotivoNoOK = "El Adjunto no es alcanzable"
                Exit Function
            End If
        End If
        If .cadenaRecursos = "" Then
            MotivoNoOK = "Ha de indicar el CadenaRecursos"
            Exit Function
        End If
        If .IDSolicitud = "" Then
            MotivoNoOK = "Ha de indicar el IDSolicitud"
            Exit Function
        End If
        If .Accion = "" Then
            MotivoNoOK = "Ha de indicar el Accion"
            Exit Function
        End If
        If .Programado <> "Sí" And .Programado <> "No" Then
            MotivoNoOK = "Se ha de indicar si es o no programado"
            Exit Function
        End If
        If Not m_ObjEntorno.ColTiposCorreos.Exists(CStr(.tipoCorreo)) Then
            MotivoNoOK = "Se ha de indicar un tipo de correo existente"
            Exit Function
        End If
        
        
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método CorreoOperaciones.MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function


Public Function Registrar( _
                                Optional p_CorreoAlInicio As Correo, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_IDCorreo As String
    Dim dato As Variant
     Dim m_MotivoNoOK As String
    On Error GoTo errores
    
    m_MotivoNoOK = MotivoNoOK(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    With Me.Correo
        
        If .DesencadenadoPor = "" Then
            .DesencadenadoPor = m_ObjUsuarioConectado.UsuarioRed
        End If
        If p_CorreoAlInicio Is Nothing Then
            
            .IDCORREO = .IDCorreoCalculado
            .FECHAGRABACION = Ahora
            m_SQL = "TbCorreosEnviados"
        Else
            m_SQL = "SELECT * " & _
                    "FROM TbCorreosEnviados " & _
                    "WHERE IDCorreo=" & p_CorreoAlInicio.IDCORREO & ";"
        End If
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        If p_CorreoAlInicio Is Nothing Then
            rcdDatos.AddNew
                rcdDatos.Fields("IDCorreo") = .IDCORREO
                If .URLAdjunto <> "" Then
                    rcdDatos.Fields("URLAdjunto") = .URLAdjunto
                End If
                rcdDatos.Fields("Aplicacion") = "HPS_TRAMITACION"
                
                
                If .DESTINATARIOS <> "" Then
                    rcdDatos.Fields("Destinatarios") = .DESTINATARIOS
                End If
                If InStr(1, .DestinatariosConCopia, "@") <> 0 Then
                    rcdDatos.Fields("DestinatariosConCopia") = .DestinatariosConCopia
                End If
                If InStr(1, .DestinatariosConCopiaOculta, "andres.romandelperal@telefonica.com") = 0 Then
                    If .DestinatariosConCopiaOculta = "" Then
                        .DestinatariosConCopiaOculta = "andres.romandelperal@telefonica.com"
                    Else
                        .DestinatariosConCopiaOculta = .DestinatariosConCopiaOculta & ";" & "andres.romandelperal@telefonica.com"
                    End If
                    
                End If
                If InStr(1, .DestinatariosConCopiaOculta, "@") <> 0 Then
                    rcdDatos.Fields("DestinatariosConCopiaOculta") = .DestinatariosConCopiaOculta
                End If
                rcdDatos.Fields("Asunto") = .Asunto
                
                If IsDate(.FechaOrdenEnvio) Then
                    rcdDatos.Fields("FechaOrdenEnvio") = .FechaOrdenEnvio
                End If
                rcdDatos.Fields("FechaGrabacion") = .FECHAGRABACION
                rcdDatos.Fields("NombrePlantilla") = .nombrePlantilla
                rcdDatos.Fields("VersionPlantilla") = .VersionPlantilla
                rcdDatos.Fields("CadenaRecursos") = .cadenaRecursos
                rcdDatos.Fields("IDSolicitud") = .IDSolicitud
                rcdDatos.Fields("Accion") = .Accion
                rcdDatos.Fields("DesencadenadoPor") = .DesencadenadoPor
                rcdDatos.Fields("Programado") = .Programado
                rcdDatos.Fields("TipoCorreo") = .tipoCorreo
            rcdDatos.Update
            rcdDatos.Close
            Set rcdDatos = Nothing
            Registrar = .IDCORREO
        Else
            rcdDatos.Edit
                
                If .DESTINATARIOS <> "" Then
                    rcdDatos.Fields("Destinatarios") = .DESTINATARIOS
                End If
                If InStr(1, .DestinatariosConCopia, "@") <> 0 Then
                    rcdDatos.Fields("DestinatariosConCopia") = .DestinatariosConCopia
                Else
                    rcdDatos.Fields("DestinatariosConCopia") = Null
                End If
                If .DestinatariosConCopiaOculta <> p_CorreoAlInicio.DestinatariosConCopiaOculta Then
                    If .DestinatariosConCopiaOculta <> "" Then
                    
                    
                        If InStr(1, .DestinatariosConCopiaOculta, "andres.romandelperal@telefonica.com") = 0 Then
                            If .DestinatariosConCopiaOculta = "" Then
                                .DestinatariosConCopiaOculta = "andres.romandelperal@telefonica.com"
                            Else
                                .DestinatariosConCopiaOculta = .DestinatariosConCopiaOculta & ";" & "andres.romandelperal@telefonica.com"
                            End If
                            
                        End If
                    End If
                End If
                
                If InStr(1, .DestinatariosConCopiaOculta, "@") <> 0 Then
                    rcdDatos.Fields("DestinatariosConCopiaOculta") = .DestinatariosConCopiaOculta
                Else
                    rcdDatos.Fields("DestinatariosConCopiaOculta") = "andres.romandelperal@telefonica.com"
                End If
                rcdDatos.Fields("Asunto") = .Asunto
                
                If IsDate(.FechaOrdenEnvio) Then
                    rcdDatos.Fields("FechaOrdenEnvio") = .FechaOrdenEnvio
                Else
                    rcdDatos.Fields("FechaOrdenEnvio") = Null
                End If
                
                rcdDatos.Fields("NombrePlantilla") = .nombrePlantilla
                rcdDatos.Fields("VersionPlantilla") = .VersionPlantilla
                rcdDatos.Fields("CadenaRecursos") = .cadenaRecursos
                rcdDatos.Fields("IDSolicitud") = .IDSolicitud
                rcdDatos.Fields("Accion") = .Accion
                rcdDatos.Fields("DesencadenadoPor") = .DesencadenadoPor
                rcdDatos.Fields("Programado") = .Programado
                rcdDatos.Fields("TipoCorreo") = .tipoCorreo
            rcdDatos.Update
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
        
        
    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreoOperaciones.Registrar ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function



Public Function CambiarFechaOrdenEnvio( _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    
    Dim m_SQL As String
    
    On Error GoTo errores
    With Me.Correo
        If IsDate(.FechaOrdenEnvio) Then
            m_SQL = "UPDATE TbCorreosEnviados SET FechaOrdenEnvio = #" & Format(.FechaOrdenEnvio, "mm/dd/yyyy") & "# " & _
                    "WHERE IDCorreo=" & .IDCORREO & ";"
        Else
            m_SQL = "UPDATE TbCorreosEnviados SET FechaOrdenEnvio = Null " & _
                    "WHERE IDCorreo=" & .IDCORREO & ";"
        End If
    End With
    getdb().Execute m_SQL
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreoOperaciones.CambiarFechaOrdenEnvio ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function



