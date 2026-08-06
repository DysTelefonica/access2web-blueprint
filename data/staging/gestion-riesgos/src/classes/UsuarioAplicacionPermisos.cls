VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsuarioAplicacionPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Campos de TbUsuariosAplicacionesPermisos realmente usados en la app
Public CorreoUsuario As String
Public IDAplicacion As String
Public EsUsuarioAdministrador As String
Public EsUsuarioCalidad As String
Public EsUsuarioCalidadAvisos As String
Public EsUsuarioTecnico As String

Private m_ObjUsuario As Usuario

Private m_eEsUsuarioAdministradorCalculado As EnumSiNo
Private m_eEsUsuarioCalidadCalculado As EnumSiNo
Private m_eEsUsuarioCalidadAvisosCalculado As EnumSiNo
Private m_eEsUsuarioTecnicoCalculado As EnumSiNo

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get Usuario() As Usuario

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjUsuario Is Nothing Then
        Set Usuario = m_ObjUsuario
        Exit Property
    End If
    If Me.CorreoUsuario = "" Then
        Exit Property
    End If
    Set m_ObjUsuario = Constructor.getUsuario(, , , Me.CorreoUsuario, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Usuario = m_ObjUsuario
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método UsuarioAplicacionPermiso.Usuario ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set Usuario(ByVal objNewValue As Usuario)

    Set m_ObjUsuario = objNewValue

End Property

Public Property Get EsUsuarioAdministradorCalculado() As EnumSiNo

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.EsUsuarioAdministrador = "Sí" Then
        m_eEsUsuarioAdministradorCalculado = EnumSiNo.Sí
    ElseIf Me.EsUsuarioAdministrador = "No" Then
        m_eEsUsuarioAdministradorCalculado = EnumSiNo.No
    End If
    EsUsuarioAdministradorCalculado = m_eEsUsuarioAdministradorCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método AplicacionAplicacionPermiso.EsUsuarioAdministradorCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Let EsUsuarioAdministradorCalculado(ByVal eNewValue As EnumSiNo)

    m_eEsUsuarioAdministradorCalculado = eNewValue

End Property

Public Property Get EsUsuarioCalidadCalculado() As EnumSiNo

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.EsUsuarioCalidad = "Sí" Then
        m_eEsUsuarioCalidadCalculado = EnumSiNo.Sí
    ElseIf Me.EsUsuarioCalidad = "No" Then
        m_eEsUsuarioCalidadCalculado = EnumSiNo.No
    End If
    EsUsuarioCalidadCalculado = m_eEsUsuarioCalidadCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método AplicacionAplicacionPermiso.EsUsuarioCalidadCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Let EsUsuarioCalidadCalculado(ByVal eNewValue As EnumSiNo)

    m_eEsUsuarioCalidadCalculado = eNewValue

End Property

Public Property Get EsUsuarioCalidadAvisosCalculado() As EnumSiNo

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.EsUsuarioCalidadAvisos = "Sí" Then
        m_eEsUsuarioCalidadAvisosCalculado = EnumSiNo.Sí
    ElseIf Me.EsUsuarioCalidadAvisos = "No" Then
        m_eEsUsuarioCalidadAvisosCalculado = EnumSiNo.No
    End If
    EsUsuarioCalidadAvisosCalculado = m_eEsUsuarioCalidadAvisosCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método AplicacionAplicacionPermiso.EsUsuarioCalidadAvisosCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Let EsUsuarioCalidadAvisosCalculado(ByVal eNewValue As EnumSiNo)

    m_eEsUsuarioCalidadAvisosCalculado = eNewValue

End Property

Public Property Get EsUsuarioTecnicoCalculado() As EnumSiNo

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.EsUsuarioTecnico = "Sí" Then
        m_eEsUsuarioTecnicoCalculado = EnumSiNo.Sí
    ElseIf Me.EsUsuarioTecnico = "No" Then
        m_eEsUsuarioTecnicoCalculado = EnumSiNo.No
    End If
    EsUsuarioTecnicoCalculado = m_eEsUsuarioTecnicoCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método AplicacionAplicacionPermiso.EsUsuarioTecnicoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Let EsUsuarioTecnicoCalculado(ByVal eNewValue As EnumSiNo)

    m_eEsUsuarioTecnicoCalculado = eNewValue

End Property

' ============================================================
' ColCampos — solo los campos realmente usados en la app
' ============================================================
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
        .Add "CorreoUsuario"
        .Add "IDAplicacion"
        .Add "EsUsuarioAdministrador"
        .Add "EsUsuarioCalidad"
        .Add "EsUsuarioCalidadAvisos"
        .Add "EsUsuarioTecnico"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método UsuarioAplicacionPermisos.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    p_Error = ""

    If m_NombrePropiedad = "CorreoUsuario" Then
        getPropiedad = Me.CorreoUsuario
        Exit Function
    End If
    If m_NombrePropiedad = "IDAplicacion" Then
        getPropiedad = Me.IDAplicacion
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioAdministrador" Then
        getPropiedad = Me.EsUsuarioAdministrador
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioCalidad" Then
        getPropiedad = Me.EsUsuarioCalidad
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioCalidadAvisos" Then
        getPropiedad = Me.EsUsuarioCalidadAvisos
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioTecnico" Then
        getPropiedad = Me.EsUsuarioTecnico
        Exit Function
    End If

    p_Error = "Propiedad no reconocida en UsuarioAplicacionPermisos.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo UsuarioAplicacionPermisos.getPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String

    On Error GoTo errores

    Me.Error = ""
    p_Error = ""

    If m_NombrePropiedad = "CorreoUsuario" Then
        Me.CorreoUsuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDAplicacion" Then
        Me.IDAplicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioAdministrador" Then
        Me.EsUsuarioAdministrador = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioCalidad" Then
        Me.EsUsuarioCalidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioCalidadAvisos" Then
        Me.EsUsuarioCalidadAvisos = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsUsuarioTecnico" Then
        Me.EsUsuarioTecnico = m_ValorPropiedad
        Exit Function
    End If

    p_Error = "Propiedad no reconocida en UsuarioAplicacionPermisos.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo UsuarioAplicacionPermisos.SetPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function











