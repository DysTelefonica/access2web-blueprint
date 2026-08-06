VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Correo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDCorreo As String
Public Originador As String
Public Destinatarios As String
Public DestinatariosConCopia As String
Public DestinatariosConCopiaOculta As String
Public Asunto As String
Public Cuerpo As String
Public URLAdjunto As String
Public FechaEnvio As String
Public FechaGrabacion As String
Public IDEdicion As String

Private m_sIDCorreoCalculado As String
Private m_EnPruebas As Boolean
Private m_objEdicion As Edicion
Private m_ObjProyecto As Proyecto
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String



Public Property Get IDCorreoCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDCorreoCalculado = DameID("TbCorreosEnviados", "IDCorreo", getdb(), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDCorreoCalculado = m_sIDCorreoCalculado
    
    Exit Property
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método Correo.IDCorreoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get Edicion() As Edicion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not m_objEdicion Is Nothing Then
        Set Edicion = m_objEdicion
        Exit Function
    End If
    If Me.IDEdicion = "" Then
        Exit Property
    End If
    Set m_objEdicion = Constructor.getEdicion(Me.IDEdicion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Edicion = m_objEdicion
    Exit Function
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método Correo.Edicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set Edicion(ByVal NewValue As Variant)

    Set m_objEdicion = NewValue

End Property


Public Function EnviarCorreo(Optional ByRef p_Error As String) As String

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not Me.Edicion Is Nothing Then
        If Not Me.Edicion.Proyecto Is Nothing Then
            If Me.Edicion.Proyecto.ParaInformeAvisos <> "Sí" Then
                ' Fix #66 (issue #64 cluster): explicit no-op sentinel so the caller
                ' can detect "we did NOT send" without relying on IDCorreo="" heuristic.
                ' The "NoOp:" prefix is the contract that the caller uses to
                ' distinguish a clean no-op return from a real error.
                p_Error = "NoOp:ParaInformeAvisos<>Sí (no se envia correo para este proyecto)"
                Exit Function
            End If
        End If
    End If
    
    Dim vRedir() As Variant
    If Application.TempVars("EnPruebas") = "Sí" Then
        vRedir = RedirigeCorreoPruebas(Me.Destinatarios, Me.DestinatariosConCopia, Me.DestinatariosConCopiaOculta, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Me.Destinatarios = vRedir(0)
        Me.DestinatariosConCopia = vRedir(1)
        Me.DestinatariosConCopiaOculta = vRedir(2)
        If vRedir(3) <> "" Then
            Me.Cuerpo = vRedir(3) & vbCrLf & Me.Cuerpo
        End If
    End If
    
    If Not IsDate(Me.FechaGrabacion) Then
        Me.FechaGrabacion = Me.FechaEnvio
    End If
    If Me.IDCorreo = "" Then
        Me.IDCorreo = Me.IDCorreoCalculado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    
    Me.Registrar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Correo.EnviarCorreo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Function

' RedirigeCorreoPruebas - retorna array [destFinal, copiaFinal, ocultaFinal, devBanner]
' Llamar ANTES de Registrar() cuando EnPruebas="Sí"
Public Function RedirigeCorreoPruebas( _
    ByVal sDestOrig As String, _
    ByVal sCopiaOrig As String, _
    ByVal sOcultaOrig As String, _
    Optional ByRef p_Error As String _
) As Variant()
    RedirigeCorreoPruebas = AplicarRedireccionCorreoPruebas( _
        sDestOrig, _
        sCopiaOrig, _
        sOcultaOrig, _
        m_ObjUsuarioConectado.CorreoUsuario, _
        Application.TempVars("EnPruebas") = "Sí", _
        p_Error)
End Function

Public Function AplicarRedireccionCorreoPruebas( _
    ByVal sDestOrig As String, _
    ByVal sCopiaOrig As String, _
    ByVal sOcultaOrig As String, _
    ByVal sCorreoUsuarioConectado As String, _
    ByVal bEnPruebas As Boolean, _
    Optional ByRef p_Error As String _
) As Variant()
    Dim vResult(0 To 3) As Variant
    Dim sBanner As String
    
    If bEnPruebas Then
        vResult(0) = sCorreoUsuarioConectado
        vResult(1) = ""
        vResult(2) = ""

        sBanner = "<div style='background:#FFF3CD;border:1px solid #856404;padding:10px;margin-bottom:15px;color:#856404;'>"
        sBanner = sBanner & "<strong>AVISO DE PRUEBAS:</strong> Este correo ha sido redirigido para pruebas.<br>"
        sBanner = sBanner & "<strong>Destinatarios reales:</strong> " & HTMLSafe(sDestOrig)
        If sCopiaOrig <> "" Then
            sBanner = sBanner & "<br><strong>Con copia:</strong> " & HTMLSafe(sCopiaOrig)
        End If
        If sOcultaOrig <> "" Then
            sBanner = sBanner & "<br><strong>Con copia oculta:</strong> " & HTMLSafe(sOcultaOrig)
        End If
        sBanner = sBanner & "</div>"
        vResult(3) = sBanner
    Else
        vResult(0) = sDestOrig
        vResult(1) = sCopiaOrig
        vResult(2) = sOcultaOrig
        vResult(3) = ""
    End If
    
    AplicarRedireccionCorreoPruebas = vResult
End Function

Public Function MotivoNoOK( _
                                Optional ByRef p_Error As String _
                                ) As String
    
    
    
    
    On Error GoTo errores
        
    With Me
        If .Destinatarios = "" And .DestinatariosConCopia = "" And .DestinatariosConCopiaOculta = "" Then
            MotivoNoOK = "Al menos ha de haber un destinatario, o en copia o en oculta"
            Exit Function
        End If
        If .Asunto = "" Then
            MotivoNoOK = "El asunto es obligatorio"
            Exit Function
        End If
'        If .Cuerpo = "" Then
'            MotivoNoOK = "El Cuerpo es obligatorio"
'            Exit Function
'        End If
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function


Public Function Registrar( _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
     Dim m_SQL As String
    Dim m_MotivoNoOK As String
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    m_MotivoNoOK = MotivoNoOK(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.IDCorreo = "" Then
        Me.IDCorreo = Me.IDCorreoCalculado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
'    Me.Destinatarios = ""
'    Me.DestinatariosConCopia = ""
'    Me.DestinatariosConCopiaOculta = "andres.romandelperal@telefonica.com"
    m_SQL = "TbCorreosEnviados"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("IDCorreo") = Me.IDCorreo
            If Application.TempVars("EnPruebas") = "Sí" Then
                .Fields("Aplicacion") = "RIESGOS_PRUEBAS"
            Else
                .Fields("Aplicacion") = "RIESGOS"
            End If
            
            .Fields("Originador") = m_ObjUsuarioConectado.UsuarioRed
            If Me.Destinatarios <> "" Then
                .Fields("Destinatarios") = Me.Destinatarios
            End If
            If Me.DestinatariosConCopia <> "" Then
                .Fields("DestinatariosConCopia") = Me.DestinatariosConCopia
            End If
            If Me.DestinatariosConCopiaOculta <> "" Then
                .Fields("DestinatariosConCopiaOculta") = Me.DestinatariosConCopiaOculta
            End If
            If Me.Cuerpo <> "" Then
                Me.Cuerpo = CuerpoHTMLConEstiloCorporativo(Me.Asunto, Me.Cuerpo)
                .Fields("Cuerpo") = Me.Cuerpo
            End If
            .Fields("Asunto") = Me.Asunto
            If Me.Cuerpo <> "" Then
                .Fields("Cuerpo") = Me.Cuerpo
            End If
            
            If Me.URLAdjunto <> "" Then
                .Fields("URLAdjunto") = Me.URLAdjunto
            End If
            .Fields("FechaGrabacion") = Now()
            If Me.IDEdicion <> "" Then
                .Fields("IDEdicion") = Me.IDEdicion
            End If
            
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing

fin:
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Registrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function SetCorreoAltaProyecto( _
                                        ByRef p_ObjProyecto As Proyecto, _
                                        Optional ByRef p_Error As String _
                                        ) As CORREO
    
    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    
    On Error GoTo errores
    p_Error = ""
    If Application.TempVars("EnPruebas") = "Sí" Then
        m_Asunto = "Alta de gestión de riesgos EN PRUEBAS: " & p_ObjProyecto.Proyecto & " (gestión de riesgos EN PRUEBAS)"
    Else
        m_Asunto = "Alta de gestión de riesgos: " & p_ObjProyecto.Proyecto & " (gestión de riesgos)"
    End If
    
m_Cuerpo = GetBodyCorreoAltaProyecto(p_ObjProyecto, p_Error)
    
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Application.TempVars("EnPruebas") = "Sí" Then
        m_Asunto = "Alta de gestión de riesgos EN PRUEBAS: " & p_ObjProyecto.Proyecto & " (gestión de riesgos EN PRUEBAS)"
    Else
        m_Asunto = "Alta de gestión de riesgos: " & p_ObjProyecto.Proyecto & " (gestión de riesgos)"
    End If
    
    If Application.TempVars("EnPruebas") = "Sí" Then
        m_Destinatarios = m_ObjUsuarioConectadoInicialmente.CorreoUsuario
    Else
        m_Destinatarios = p_ObjProyecto.CadenaCorreoAutorizados
        p_Error = p_ObjProyecto.Error
        If p_Error <> "" Then
            p_Error = "El método CorreoGestor.SetCorreoAltaProyecto ha llamado al método CorreoGestor.getCadenaCorreoAutorizados el error: " & _
                    vbNewLine & p_Error
            Err.Raise 1000
        End If
        m_DestinatariosConCopia = m_ObjEntorno.CorreoCalidad
    End If
    If Application.TempVars("EnPruebas") = "Sí" Then
        m_Destinatarios = m_ObjUsuarioConectadoInicialmente.CorreoUsuario
    Else
        m_Destinatarios = p_ObjProyecto.CadenaCorreoAutorizados
        p_Error = p_ObjProyecto.Error
        If p_Error <> "" Then
            p_Error = "El método CorreoGestor.SetCorreoAltaProyecto ha llamado al método CorreoGestor.getCadenaCorreoAutorizados el error: " & _
                    vbNewLine & p_Error
            Err.Raise 1000
        End If
        m_DestinatariosConCopia = m_ObjEntorno.CorreoCalidad
    End If
    
    
    If m_Destinatarios = "" Then
        Exit Function
    End If
    
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set SetCorreoAltaProyecto = New CORREO
    With SetCorreoAltaProyecto
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            ' Fix #66 (issue #64 cluster): NoOp sentinels (e.g., ParaInformeAvisos<>Sí)
            ' are clean returns, not errors. See EsNoOpError in Funciones Generales.bas.
            If EsNoOpError(p_Error) Then
                p_Error = ""
            Else
                Err.Raise 1000
            End If
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SetCorreoAltaProyecto ha devuelto el error: " & Err.Description
    End If
End Function
Public Function SetCorreoRiesgoAceptadoRetirado( _
                                                ByRef p_ObjRiesgo As riesgo, _
                                                Optional p_EsAceptado As EnumSiNo = EnumSiNo.Sí, _
                                                Optional ByRef p_Error As String _
                                                ) As CORREO
    
    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    Dim m_ObjUsuarioCalidad As Usuario
    
    
    On Error GoTo errores
    p_Error = ""
    
    If p_ObjRiesgo Is Nothing Then
        p_Error = "No se ha indicado el riesgo para el correo"
        Err.Raise 1000
    End If
    Set m_ObjProyecto = p_ObjRiesgo.Edicion.Proyecto
    If m_ObjProyecto Is Nothing Then
        p_Error = "No se conoce la gestión de riesgos a la que pertenece el Riesgo"
        Err.Raise 1000
    End If
    If p_EsAceptado = EnumSiNo.Sí Then
        If Application.TempVars("EnPruebas") = "Sí" Then
            m_Asunto = "Riesgo Aceptado por el Técnico EN PRUEBAS: " & m_ObjProyecto.Proyecto & " (gestión de riesgos EN PRUEBAS)"
        Else
            m_Asunto = "Riesgo Aceptado por el Técnico: " & m_ObjProyecto.Proyecto & " (gestión de riesgos)"
        End If
        
    Else
        If Application.TempVars("EnPruebas") = "Sí" Then
            m_Asunto = "Riesgo Retirado por el Técnico EN PRUEBAS: " & m_ObjProyecto.Proyecto & " (gestión de riesgos EN PRUEBAS)"
        Else
            m_Asunto = "Riesgo Retirado por el Técnico: " & m_ObjProyecto.Proyecto & " (gestión de riesgos)"
        End If
        
    End If
    
    m_Cuerpo = p_ObjRiesgo.getHTMLRiesgoAceptadoRetirado(p_EsAceptado, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjUsuarioCalidad = getUsuario(, , m_ObjProyecto.NombreUsuarioCalidad, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Application.TempVars("EnPruebas") = "Sí" Then
        m_Destinatarios = m_ObjUsuarioConectadoInicialmente.CorreoUsuario
    Else
        m_Destinatarios = m_ObjUsuarioCalidad.CorreoUsuario
        m_DestinatariosConCopia = m_ObjProyecto.CadenaCorreoAutorizados
        p_Error = m_ObjProyecto.Error
        If p_Error <> "" Then
            p_Error = "El método CorreoGestor.SetCorreoRiesgoAceptadoRetirado ha llamado al método CorreoGestor.getCadenaCorreoAutorizados el error: " & _
                    vbNewLine & p_Error
            Err.Raise 1000
        End If
    End If
    
    
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set SetCorreoRiesgoAceptadoRetirado = New CORREO
    With SetCorreoRiesgoAceptadoRetirado
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            ' Fix #66 (issue #64 cluster): NoOp sentinels (e.g., ParaInformeAvisos<>Sí)
            ' are clean returns, not errors. See EsNoOpError in Funciones Generales.bas.
            If EsNoOpError(p_Error) Then
                p_Error = ""
            Else
                Err.Raise 1000
            End If
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SetCorreoRiesgoAceptadoRetirado ha devuelto el error: " & Err.Description
    End If
End Function

Public Function SetCorreoNuevaPublicacion( _
                                            ByRef p_ObjProyecto As Proyecto, _
                                            p_URLInforme As String, _
                                            Optional p_ObjEdicion As Edicion, _
                                            Optional p_FechaPublicacion As String, _
                                            Optional p_EsUltimaEdicion As String, _
                                            Optional p_ConEnvioAlRAC As EnumSiNo, _
                                            Optional ByRef p_Error As String _
                                            ) As CORREO
    
    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_CorreoRAC As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    Dim m_ObjUltimaEdicion As Edicion
    Dim m_Edicion As String
    
    On Error GoTo errores
    p_Error = ""
    
    If p_ConEnvioAlRAC = Empty Then
        p_ConEnvioAlRAC = EnumSiNo.No
    End If
    
    If Not p_ObjEdicion Is Nothing Then
        m_Edicion = p_ObjEdicion.Edicion
        If p_ObjProyecto.FechaCierre <> "" Then
            If Application.TempVars("EnPruebas") = "Sí" Then
                m_Asunto = "Última Edición Publicada EN PRUEBAS" & m_Edicion & ": " & p_ObjProyecto.NombreProyecto & " (gestión de riesgos EN PRUEBAS)"
            Else
                m_Asunto = "Última Edición Publicada " & m_Edicion & ": " & p_ObjProyecto.NombreProyecto & " (gestión de riesgos)"
            End If
            
        Else
            If Application.TempVars("EnPruebas") = "Sí" Then
                m_Asunto = "Edición Publicada EN PRUEBAS " & m_Edicion & ": " & p_ObjProyecto.NombreProyecto & " (gestión de riesgos EN PRUEBAS)"
            Else
                m_Asunto = "Edición Publicada " & m_Edicion & ": " & p_ObjProyecto.NombreProyecto & " (gestión de riesgos)"
            End If
        End If
        
    Else
        If p_ObjProyecto.FechaCierre <> "" Then
            If Application.TempVars("EnPruebas") = "Sí" Then
                m_Asunto = "Última Edición Publicada EN PRUEBAS: " & p_ObjProyecto.NombreProyecto & " (gestión de riesgos EN PRUEBAS)"
            Else
                m_Asunto = "Última Edición Publicada: " & p_ObjProyecto.NombreProyecto & " (gestión de riesgos)"
            End If
        Else
            If Application.TempVars("EnPruebas") = "Sí" Then
                m_Asunto = "Edición Publicada EN PRUEBAS: " & p_ObjProyecto.NombreProyecto & " (gestión de riesgos EN PRUEBAS)"
            Else
                m_Asunto = "Edición Publicada: " & p_ObjProyecto.NombreProyecto & " (gestión de riesgos)"
            End If
        End If
        
    End If
    With p_ObjProyecto
        m_Cuerpo = GetBodyCorreoNuevaPublicacion(p_ObjProyecto, p_ObjEdicion, p_FechaPublicacion, p_EsUltimaEdicion, p_Error)
        If Application.TempVars("EnPruebas") = "Sí" Then
            m_Destinatarios = m_ObjEntorno.CorreoCalidad
        Else
            m_Destinatarios = .CadenaCorreoAutorizados
            p_Error = .Error
            If p_Error <> "" Then
                p_Error = "El método CorreoGestor.SetCorreoNuevaPublicacion ha llamado al método CorreoGestor.getCadenaCorreoAutorizados el error: " & _
                        vbNewLine & p_Error
                Err.Raise 1000
            End If
        End If
        m_CorreoRAC = .CorreoRAC
        If InStr(1, m_CorreoRAC, "@") <> 0 Then
            If m_DestinatariosConCopia = "" Then
                m_DestinatariosConCopia = m_CorreoRAC
            Else
                m_DestinatariosConCopia = m_DestinatariosConCopia & ";" & m_CorreoRAC
            End If
            
        End If
        If p_ConEnvioAlRAC = EnumSiNo.Sí And m_CorreoRAC <> "" Then
            m_DestinatariosConCopia = m_ObjEntorno.CorreoCalidad & ";" & m_CorreoRAC
        Else
            m_DestinatariosConCopia = m_ObjEntorno.CorreoCalidad
        End If
    End With
    
    
    
    
    
    
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set SetCorreoNuevaPublicacion = New CORREO
    With SetCorreoNuevaPublicacion
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .URLAdjunto = p_URLInforme
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            ' Fix #66 (issue #64 cluster): NoOp sentinels (e.g., ParaInformeAvisos<>Sí)
            ' are clean returns, not errors. See EsNoOpError in Funciones Generales.bas.
            If EsNoOpError(p_Error) Then
                p_Error = ""
            Else
                Err.Raise 1000
            End If
        End If
    End With
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SetCorreoNuevaPublicacion ha devuelto el error: " & Err.Description
    End If
End Function

Public Function NuevaPublicacion( _
                                    ByRef p_ObjProyecto As Proyecto, _
                                    p_URLInforme As String, _
                                    Optional p_ObjEdicion As Edicion, _
                                    Optional p_FechaPublicacion As String, _
                                    Optional p_EsUltimaEdicion As String, _
                                    Optional p_ConEnvioAlRAC As EnumSiNo, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim m_ObjCorreo As CORREO
    On Error GoTo errores
    
    Set m_ObjCorreo = getCorreoNuevaPublicacion( _
                    p_ObjProyecto, p_URLInforme, _
                    p_ObjEdicion, p_FechaPublicacion, _
                    p_EsUltimaEdicion, _
                    p_ConEnvioAlRAC, p_Error)
    If p_Error <> "" Then
       Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método NuevaPublicacion ha devuelto el error: " & Err.Description
    End If
End Function


Public Function AltaProyecto( _
                                ByRef p_ObjProyecto As Proyecto, _
                                Optional ByRef p_Error As String _
                                ) As String

    Dim m_ObjCorreo As CORREO
    On Error GoTo errores
    
    Set m_ObjCorreo = getCorreoByProyecto(p_ObjProyecto, p_Error)
    If p_Error <> "" Then
       Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AltaProyecto ha devuelto el error: " & Err.Description
    End If
End Function

Public Function SetCorreoRiesgoRequiereRetipificacion( _
                                                        ByRef p_ObjRiesgo As riesgo, _
                                                        Optional ByRef p_Error As String _
                                                        ) As CORREO
    
    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    Dim m_ObjUsuarioCalidad As Usuario
    
    On Error GoTo errores
    p_Error = ""
    
    Set m_ObjProyecto = p_ObjRiesgo.Edicion.Proyecto
    If m_ObjProyecto Is Nothing Then
        p_Error = "No se conoce la gestión de riesgos a la que pertenece el Riesgo"
        Err.Raise 1000
    End If
    With m_ObjProyecto
        If Application.TempVars("EnPruebas") = "Sí" Then
            m_Asunto = "Riesgo No encontrado en Biblioteca por el Técnico EN PRUEBAS: " & .Proyecto & " (gestión de riesgos EN PRUEBAS)"
        Else
            m_Asunto = "Riesgo No encontrado en Biblioteca por el Técnico: " & .Proyecto & " (gestión de riesgos)"
        End If
        
        m_Cuerpo = p_ObjRiesgo.HTMLRiesgoPorRetipificar
        p_Error = p_ObjRiesgo.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_ObjUsuarioCalidad = getUsuario(, , .NombreUsuarioCalidad, , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Application.TempVars("EnPruebas") = "Sí" Then
            m_Destinatarios = m_ObjUsuarioConectadoInicialmente.CorreoUsuario
        Else
            m_Destinatarios = m_ObjUsuarioCalidad.CorreoUsuario
            m_DestinatariosConCopia = .CadenaCorreoAutorizados
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        
    End With
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set SetCorreoRiesgoRequiereRetipificacion = New CORREO
    With SetCorreoRiesgoRequiereRetipificacion
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            ' Fix #66 (issue #64 cluster): NoOp sentinels (e.g., ParaInformeAvisos<>Sí)
            ' are clean returns, not errors. See EsNoOpError in Funciones Generales.bas.
            If EsNoOpError(p_Error) Then
                p_Error = ""
            Else
                Err.Raise 1000
            End If
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SetCorreoRiesgoRequiereRetipificacion ha devuelto el error: " & Err.Description
    End If
End Function
Public Function SetCorreoRiesgoRetipificacion( _
                                                ByRef p_ObjRiesgo As riesgo, _
                                                ByRef p_ObjRiesgoAlInicio As riesgo, _
                                                Optional ByRef p_Error As String _
                                                ) As CORREO
    
    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    Dim m_ObjUsuarioCalidad As Usuario
    
    On Error GoTo errores
    p_Error = ""
    
    Set m_ObjProyecto = p_ObjRiesgo.Edicion.Proyecto
    If m_ObjProyecto Is Nothing Then
        p_Error = "No se conoce la gestión de riesgos a la que pertenece el Riesgo"
        Err.Raise 1000
    End If
    If Application.TempVars("EnPruebas") = "Sí" Then
        m_Asunto = "Riesgo Asignado de la Biblioteca por Calidad EN PRUEBAS: " & m_ObjProyecto.Proyecto & " (gestión de riesgos EN PRUEBAS)"
    Else
        m_Asunto = "Riesgo Asignado de la Biblioteca por Calidad: " & m_ObjProyecto.Proyecto & " (gestión de riesgos)"
    End If
    
    m_Cuerpo = p_ObjRiesgo.getHTMLRiesgoRetipificado(p_ObjRiesgoAlInicio, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjUsuarioCalidad = getUsuario(, , m_ObjProyecto.NombreUsuarioCalidad, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Application.TempVars("EnPruebas") = "Sí" Then
        m_Destinatarios = m_ObjUsuarioConectadoInicialmente.CorreoUsuario
    Else
        m_Destinatarios = m_ObjProyecto.CadenaCorreoAutorizados
        p_Error = m_ObjProyecto.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_DestinatariosConCopia = m_ObjEntorno.CorreoCalidad
    End If
    
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set SetCorreoRiesgoRetipificacion = New CORREO
    With SetCorreoRiesgoRetipificacion
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            ' Fix #66 (issue #64 cluster): NoOp sentinels (e.g., ParaInformeAvisos<>Sí)
            ' are clean returns, not errors. See EsNoOpError in Funciones Generales.bas.
            If EsNoOpError(p_Error) Then
                p_Error = ""
            Else
                Err.Raise 1000
            End If
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SetCorreoRiesgoRetipificacion ha devuelto el error: " & Err.Description
    End If
End Function
Public Function SetCorreoRiesgoMaterializado( _
                                                ByRef p_ObjRiesgo As riesgo, _
                                                Optional ByRef p_Error As String _
                                                ) As CORREO
    
    Dim m_Destinatarios As String
    Dim m_DestinatariosConCopia As String
    Dim m_DestinatariosConCopiaOculta As String
    Dim m_Asunto As String
    Dim m_Cuerpo As String
    Dim m_FechaGrabacion As String
    Dim m_ObjUsuarioCalidad As Usuario
    
    On Error GoTo errores
    p_Error = ""
    
    Set m_ObjProyecto = p_ObjRiesgo.Edicion.Proyecto
    If m_ObjProyecto Is Nothing Then
        p_Error = "No se conoce la gestión de riesgos a la que pertenece el Riesgo"
        Err.Raise 1000
    End If
    If Application.TempVars("EnPruebas") = "Sí" Then
        m_Asunto = "Riesgo Materializado EN PRUEBAS: " & m_ObjProyecto.Proyecto & " (gestión de riesgos EN PRUEBAS)"
    Else
        m_Asunto = "Riesgo Materializado: " & m_ObjProyecto.Proyecto & " (gestión de riesgos)"
    End If
    
    m_Cuerpo = p_ObjRiesgo.HTMLRiesgoMaterializado
    p_Error = p_ObjRiesgo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjUsuarioCalidad = getUsuario(, , m_ObjProyecto.NombreUsuarioCalidad, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Application.TempVars("EnPruebas") = "Sí" Then
        m_Destinatarios = m_ObjUsuarioConectadoInicialmente.CorreoUsuario
    Else
        m_Destinatarios = m_ObjEntorno.CorreoCalidad
        m_DestinatariosConCopia = m_ObjProyecto.CadenaCorreoAutorizados
        p_Error = m_ObjProyecto.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    m_DestinatariosConCopiaOculta = m_ObjEntorno.DestinatariosCorreoAdministradores
    m_FechaGrabacion = Now()
    Set SetCorreoRiesgoMaterializado = New CORREO
    With SetCorreoRiesgoMaterializado
        .Asunto = m_Asunto
        .Cuerpo = m_Cuerpo
        .Destinatarios = m_Destinatarios
        .DestinatariosConCopia = m_DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_DestinatariosConCopiaOculta
        .FechaGrabacion = m_FechaGrabacion
        .EnviarCorreo p_Error
        If p_Error <> "" Then
            ' Fix #66 (issue #64 cluster): NoOp sentinels (e.g., ParaInformeAvisos<>Sí)
            ' are clean returns, not errors. See EsNoOpError in Funciones Generales.bas.
            If EsNoOpError(p_Error) Then
                p_Error = ""
            Else
                Err.Raise 1000
            End If
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SetCorreoRiesgoMaterializado ha devuelto el error: " & Err.Description
    End If
End Function


Public Property Get ColCampos() As Collection

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColCampos Is Nothing Then
        Set ColCampos = m_objColCampos
        Exit Property
    End If
    Set m_objColCampos = New Collection
    With m_objColCampos
        .Add "IDCorreo"
        .Add "Originador"
        .Add "Destinatarios"
        .Add "DestinatariosConCopia"
        .Add "DestinatariosConCopiaOculta"
        .Add "Asunto"
        .Add "Cuerpo"
        .Add "URLAdjunto"
        .Add "FechaEnvio"
        .Add "FechaGrabacion"
        .Add "IDEdicion"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método CorreosEnviado.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDCorreo" Then
        getPropiedad = Me.IDCorreo
        Exit Function
    End If
    If m_NombrePropiedad = "Originador" Then
        getPropiedad = Me.Originador
        Exit Function
    End If
    If m_NombrePropiedad = "Destinatarios" Then
        getPropiedad = Me.Destinatarios
        Exit Function
    End If
    If m_NombrePropiedad = "DestinatariosConCopia" Then
        getPropiedad = Me.DestinatariosConCopia
        Exit Function
    End If
    If m_NombrePropiedad = "DestinatariosConCopiaOculta" Then
        getPropiedad = Me.DestinatariosConCopiaOculta
        Exit Function
    End If
    If m_NombrePropiedad = "Asunto" Then
        getPropiedad = Me.Asunto
        Exit Function
    End If
    If m_NombrePropiedad = "Cuerpo" Then
        getPropiedad = Me.Cuerpo
        Exit Function
    End If
    If m_NombrePropiedad = "URLAdjunto" Then
        getPropiedad = Me.URLAdjunto
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEnvio" Then
        getPropiedad = Me.FechaEnvio
        Exit Function
    End If
    If m_NombrePropiedad = "FechaGrabacion" Then
        getPropiedad = Me.FechaGrabacion
        Exit Function
    End If
    If m_NombrePropiedad = "IDEdicion" Then
        getPropiedad = Me.IDEdicion
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en CorreosEnviado.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreosEnviado.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDCorreo" Then
        Me.IDCorreo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Originador" Then
        Me.Originador = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Destinatarios" Then
        Me.Destinatarios = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DestinatariosConCopia" Then
        Me.DestinatariosConCopia = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DestinatariosConCopiaOculta" Then
        Me.DestinatariosConCopiaOculta = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Asunto" Then
        Me.Asunto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Cuerpo" Then
        Me.Cuerpo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "URLAdjunto" Then
        Me.URLAdjunto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEnvio" Then
        Me.FechaEnvio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaGrabacion" Then
        Me.FechaGrabacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEdicion" Then
        Me.IDEdicion = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en CorreosEnviado.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreosEnviado.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function




