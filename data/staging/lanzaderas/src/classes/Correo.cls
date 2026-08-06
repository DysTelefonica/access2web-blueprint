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
Public Destinatarios As String
Public DestinatariosConCopia As String
Public DestinatariosConCopiaOculta As String
Public Asunto As String
Public Cuerpo As String
Public FechaEnvio As String
Public FechaGrabacion As String
Public URLAdjunto As String

Public Error As String
Private m_sIDCorreoCalculado As String
Private m_objColCampos As Collection


Public Property Get IDCorreoCalculado() As String

    On Error GoTo errores
    Me.Error = ""
    m_sIDCorreoCalculado = DameID("TbCorreosEnviados", "IDCorreo", getdb(), Me.Error)
    If Me.Error <> "" Then
        Err.Raise 1000
    End If
    IDCorreoCalculado = m_sIDCorreoCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Correo.IDCorreoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property


Public Function EnviarCorreo( _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim rcdDatos As DAO.Recordset
    
    On Error GoTo errores
    'GoTo fin
    With Me
        If .IDCorreo = "" Then
            .IDCorreo = .IDCorreoCalculado
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If InStr(1, .Destinatarios, "@") = 0 And InStr(1, .DestinatariosConCopia, "@") = 0 Then
            p_Error = "No hay destinatarios de ningún tipo"
            Err.Raise 1000
        End If
        If .DestinatariosConCopiaOculta = "" Then
            .DestinatariosConCopiaOculta = "andres.romandeleperal@telefonica.com"
        Else
            If InStr(1, .DestinatariosConCopiaOculta, "andres.romandeleperal@telefonica.com") = 0 Then
                .DestinatariosConCopiaOculta = .DestinatariosConCopiaOculta & ";" & "andres.romandeleperal@telefonica.com"
            End If
        End If
        If .Asunto = "" Then
            p_Error = "Falta el asunto"
            Err.Raise 1000
        End If
'        If .Cuerpo = "" Then
'            p_Error = "Falta el Cuerpo"
'            Err.Raise 1000
'        End If
        If Me.FechaGrabacion = "" Then
            Me.FechaGrabacion = Now()
        End If
        
    End With
    m_SQL = "TbCorreosEnviados"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("IDCorreo") = Me.IDCorreo
            .Fields("Aplicacion") = "Lanzadera"
            If Not m_ObjUsuarioConectadoLogin Is Nothing Then
                .Fields("Originador") = m_ObjUsuarioConectadoLogin.UsuarioRed
            End If
            
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
            .Fields("FechaGrabacion") = Me.FechaGrabacion
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
fin:
    EnviarCorreo = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Correo.EnviarCorreo ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Public Property Get ColCampos() As Collection
    
    Dim tdf As DAO.TableDef
    Dim fld As Object
    On Error GoTo errores
    
    Me.Error = ""
    If Not m_objColCampos Is Nothing Then
        Set ColCampos = m_objColCampos
        Exit Property
    End If
    Set tdf = getdb().TableDefs("TbUsuariosCorreosEnvio")
    For Each fld In tdf.Fields
        If m_objColCampos Is Nothing Then
            Set m_objColCampos = New Collection
        
        End If
        m_objColCampos.Add fld.Name
    Next
    Set tdf = Nothing
    
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El metodo Correo.ColCampos ha devuelto el error: " & vbNewLine & Err.Description
    End If


End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If m_NombrePropiedad = "IDCorreo" Then
        getPropiedad = Me.IDCorreo
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
    If m_NombrePropiedad = "URLAdjunto" Then
        getPropiedad = Me.URLAdjunto
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Correo.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Correo.getPropiedad ha devuelto el error: " & vbNewLine & Err.Description
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
    If m_NombrePropiedad = "URLAdjunto" Then
        Me.URLAdjunto = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Correo.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Correo.SetPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function








