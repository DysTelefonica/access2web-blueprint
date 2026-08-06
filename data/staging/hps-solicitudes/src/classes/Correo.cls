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

Public IDCORREO As String
Public URLAdjunto As String
Public Aplicacion As String
Public DESTINATARIOS As String
Public DestinatariosConCopia As String
Public DestinatariosConCopiaOculta As String
Public Asunto As String
Public FechaEnvio As String
Public FechaOrdenEnvio As String
Public FECHAGRABACION As String
Public nombrePlantilla As String
Public VersionPlantilla As String
Public cadenaRecursos As String
Public IDSolicitud As String
Public Accion As String
Public DesencadenadoPor As String
Public Programado As String
Public tipoCorreo As String

Private m_sIDCorreoCalculado As String
Private m_objSolicitud As solicitud
Private m_sURLPlantilla As String

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String
Public Property Get EstadoTexto() As String
    On Error Resume Next
    
    If IsDate(Me.FechaEnvio) Then
        EstadoTexto = "Enviado"
        Exit Property
    End If
    
    ' Si no tiene fecha envio, miramos si es futuro
    If IsDate(Me.FechaOrdenEnvio) Then
        If CDate(Me.FechaOrdenEnvio) > Now() Then
            EstadoTexto = "Programado"
            Exit Property
        End If
    End If
    
    ' Si no es enviado ni futuro, es pendiente de salir ya
    EstadoTexto = "Pendiente"
End Property
' SUSTITUIR LA PROPIEDAD ACTUAL Get solicitud POR ESTA:

Public Property Get solicitud() As solicitud
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    ' Si ya tenemos el objeto en memoria, lo devolvemos
    If Not m_objSolicitud Is Nothing Then
        Set solicitud = m_objSolicitud
        Exit Property
    End If
    
    ' Validaciones de IDs vacíos
    If Me.IDSolicitud = "" Then
        Exit Property
    End If
    
    ' --- NUEVO: Si es un correo de prueba/genérico, no buscamos en BD ---
    If Me.IDSolicitud = CONST_IDSOLICITUD_PRUEBA Then
        ' Devolvemos Nothing explícitamente, ya que no hay solicitud real vinculada
        Set solicitud = Nothing
        Exit Property
    End If
    ' -------------------------------------------------------------------

    Set m_objSolicitud = constructor.getSolicitud(p_IDSolicitud:=Me.IDSolicitud, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set solicitud = m_objSolicitud
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Correo.Solicitud ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get URLPlantilla() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLPlantilla <> "" Then
        URLPlantilla = m_sURLPlantilla
        Exit Property
    End If
    If Me.VersionPlantilla = "" Then
        Exit Property
    End If
    
    m_sURLPlantilla = m_ObjEntorno.URLDirectorioPlantillasHTML & Me.VersionPlantilla & "\" & Me.nombrePlantilla
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    URLPlantilla = m_sURLPlantilla
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Correo.URLPlantilla  ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get IDCorreoCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDCorreoCalculado = DameID("TbCorreosEnviados", "IDCorreo", getdb, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDCorreoCalculado = m_sIDCorreoCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Correo.IDCorreoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



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
        .Add "Aplicacion"
        .Add "Destinatarios"
        .Add "DestinatariosConCopia"
        .Add "DestinatariosConCopiaOculta"
        .Add "Asunto"
        .Add "FechaEnvio"
        .Add "FechaOrdenEnvio"
        .Add "URLAdjunto"
        .Add "FechaGrabacion"
        .Add "NombrePlantilla"
        .Add "VersionPlantilla"
        .Add "CadenaRecursos"
        .Add "IDSolicitud"
        .Add "Accion"
        .Add "DesencadenadoPor"
        .Add "Programado"
        .Add "TipoCorreo"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Correo.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDCorreo" Then
        getPropiedad = Me.IDCORREO
        Exit Function
    End If
    If m_NombrePropiedad = "Aplicacion" Then
        getPropiedad = Me.Aplicacion
        Exit Function
    End If
    If m_NombrePropiedad = "Destinatarios" Then
        getPropiedad = Me.DESTINATARIOS
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
    If m_NombrePropiedad = "FechaOrdenEnvio" Then
        getPropiedad = Me.FechaOrdenEnvio
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEnvio" Then
        getPropiedad = Me.FechaEnvio
        Exit Function
    End If
    If m_NombrePropiedad = "URLAdjunto" Then
        getPropiedad = Me.URLAdjunto
        Exit Function
    End If
    If m_NombrePropiedad = "FechaGrabacion" Then
        getPropiedad = Me.FECHAGRABACION
        Exit Function
    End If
    If m_NombrePropiedad = "NombrePlantilla" Then
        getPropiedad = Me.nombrePlantilla
        Exit Function
    End If
    If m_NombrePropiedad = "VersionPlantilla" Then
        getPropiedad = Me.VersionPlantilla
        Exit Function
    End If
    If m_NombrePropiedad = "CadenaRecursos" Then
        getPropiedad = Me.cadenaRecursos
        Exit Function
    End If
    If m_NombrePropiedad = "IDSolicitud" Then
        getPropiedad = Me.IDSolicitud
        Exit Function
    End If
    If m_NombrePropiedad = "Accion" Then
        getPropiedad = Me.Accion
        Exit Function
    End If
    If m_NombrePropiedad = "DesencadenadoPor" Then
        getPropiedad = Me.DesencadenadoPor
        Exit Function
    End If
    If m_NombrePropiedad = "Programado" Then
        getPropiedad = Me.Programado
        Exit Function
    End If
    If m_NombrePropiedad = "TipoCorreo" Then
        getPropiedad = Me.tipoCorreo
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Correo.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Correo.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function setPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDCorreo" Then
        Me.IDCORREO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Aplicacion" Then
        Me.Aplicacion = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "Destinatarios" Then
        Me.DESTINATARIOS = m_ValorPropiedad
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
    If m_NombrePropiedad = "FechaOrdenEnvio" Then
        Me.FechaOrdenEnvio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEnvio" Then
        Me.FechaEnvio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "URLAdjunto" Then
        Me.URLAdjunto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaGrabacion" Then
        Me.FECHAGRABACION = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombrePlantilla" Then
        Me.nombrePlantilla = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "VersionPlantilla" Then
        Me.VersionPlantilla = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CadenaRecursos" Then
        Me.cadenaRecursos = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDSolicitud" Then
        Me.IDSolicitud = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Accion" Then
        Me.Accion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DesencadenadoPor" Then
        Me.DesencadenadoPor = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Programado" Then
        Me.Programado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TipoCorreo" Then
        Me.tipoCorreo = m_ValorPropiedad
        Exit Function
    End If

    p_Error = "Propiedad no reconocida en Correo.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Correo.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function








