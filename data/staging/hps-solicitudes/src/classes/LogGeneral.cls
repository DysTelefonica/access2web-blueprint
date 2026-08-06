VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDLog As String
Public FECHA As String
Public Accion As String
Public Usuario As String
Public Descripcion As String
Public DNI As String
Public CorreosAutomaticos As String
Public IDSolicitud As String

Private m_sIDLogCalculado As String
Private m_sIDSolicitudCalculado As String
Private m_ObjUsuarioObj As Usuario
Private m_ObjSolicitudPrimera As solicitud

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get SolicitudPrimera() As solicitud
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjSolicitudPrimera Is Nothing Then
        Set SolicitudPrimera = m_ObjSolicitudPrimera
        Exit Property
    End If
    Set m_ObjSolicitudPrimera = constructor.getSolicitudPrimeraDeLog(p_DNI:=Me.DNI, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    Set SolicitudPrimera = m_ObjSolicitudPrimera
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Log.SolicitudPrimera ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get IDSolicitudCalculado() As String
    
    Dim m_SolicitudPrimera As solicitud
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.DNI = "" Then
        Exit Property
    End If
    If m_sIDSolicitudCalculado <> "" Then
        IDSolicitudCalculado = m_sIDSolicitudCalculado
        Exit Property
    End If
    Set m_SolicitudPrimera = Me.SolicitudPrimera
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_SolicitudPrimera Is Nothing Then
        Exit Property
    End If
    m_sIDSolicitudCalculado = m_SolicitudPrimera.IDSolicitud
    
    IDSolicitudCalculado = m_sIDSolicitudCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Log.IDSolicitudCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get UsuarioObj() As Usuario
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjUsuarioObj Is Nothing Then
        Set UsuarioObj = m_ObjUsuarioObj
        Exit Property
    End If
    Set m_ObjUsuarioObj = constructor.getUsuario(p_UsuarioRed:=Me.Usuario, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    Set UsuarioObj = m_ObjUsuarioObj
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Log.UsuarioObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get IDLogCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDLogCalculado = DameID("TbLogsGeneral", "IDLog", getdb(), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDLogCalculado = m_sIDLogCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Log.IDLogCalculado ha devuelto el error: " & vbNewLine & Err.Description
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
        .Add "IDLog"
        .Add "Fecha"
        .Add "Accion"
        .Add "Usuario"
        .Add "Descripcion"
        .Add "DNI"
        .Add "CorreosAutomaticos"
        .Add "IDSolicitud"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Log.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDLog" Then
        getPropiedad = Me.IDLog
        Exit Function
    End If
    If m_NombrePropiedad = "Fecha" Then
        getPropiedad = Me.FECHA
        Exit Function
    End If
    If m_NombrePropiedad = "Accion" Then
        getPropiedad = Me.Accion
        Exit Function
    End If
    If m_NombrePropiedad = "Usuario" Then
        getPropiedad = Me.Usuario
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    If m_NombrePropiedad = "DNI" Then
        getPropiedad = Me.DNI
        Exit Function
    End If
    If m_NombrePropiedad = "CorreosAutomaticos" Then
        getPropiedad = Me.CorreosAutomaticos
        Exit Function
    End If
    If m_NombrePropiedad = "IDSolicitud" Then
        getPropiedad = Me.IDSolicitud
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Log.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Log.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function setPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDLog" Then
        Me.IDLog = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "Fecha" Then
        Me.FECHA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Accion" Then
        Me.Accion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Usuario" Then
        Me.Usuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DNI" Then
        Me.DNI = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CorreosAutomaticos" Then
        Me.CorreosAutomaticos = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDSolicitud" Then
        Me.IDSolicitud = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Log.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Log.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

