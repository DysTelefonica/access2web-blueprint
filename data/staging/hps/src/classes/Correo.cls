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
Public FechaEnvio As String
Public FechaGrabacion As String
Public IDEdicion As String

Private m_sIDCorreoCalculado As String

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String



Public Property Get IDCorreoCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_EnOficina <> EnumSiNo.Sí Then
        Exit Property
    End If
    m_sIDCorreoCalculado = DameID("TbCorreosEnviados", "IDCorreo", getdbCorreo(), m_Error)
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

Public Function EnviarCorreo(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    
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
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
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
    If m_EnOficina <> EnumSiNo.Sí Then
        Exit Function
    End If
    m_MotivoNoOK = MotivoNoOK(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.IDCorreo = Me.IDCorreoCalculado
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    'TODO:QUITAR
    'GoTo fin
    m_SQL = "TbCorreosEnviados"
    Set rcdDatos = getdbCorreo().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("IDCorreo") = Me.IDCorreo
            .Fields("Aplicacion") = "HPS"
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
            .Fields("Asunto") = Me.Asunto
            .Fields("Cuerpo") = Me.Cuerpo
            .Fields("FechaGrabacion") = Now()
            
            
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
        .Add "FechaEnvio"
        .Add "FechaGrabacion"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método CorreosEnviado.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
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
    If m_NombrePropiedad = "FechaEnvio" Then
        getPropiedad = Me.FechaEnvio
        Exit Function
    End If
    If m_NombrePropiedad = "FechaGrabacion" Then
        getPropiedad = Me.FechaGrabacion
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
    If m_NombrePropiedad = "FechaEnvio" Then
        Me.FechaEnvio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaGrabacion" Then
        Me.FechaGrabacion = m_ValorPropiedad
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

