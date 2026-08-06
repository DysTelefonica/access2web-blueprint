VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Campos de TbUsuariosAplicaciones realmente usados en la app
Public CorreoUsuario As String
Public UsuarioRed As String
Public Nombre As String
Public ID As String
Public FechaBaja As String

Private m_ObjUltimoProyecto As UltimoProyecto
Private m_objColProyectos As Scripting.Dictionary
Private m_objColProyectosActivos As Scripting.Dictionary
Private m_eEsAdministradorCalculado As EnumSiNo

Private m_eEsCalidadCalculado As EnumSiNo
Private m_eEsTecnicoCalculado As EnumSiNo
Private m_eActivo As EnumSiNo
Private m_objColAplicacionesPermisos As Scripting.Dictionary
Private m_objPermisos As UsuarioAplicacionPermisos
Public Error As String
Private m_Error As String
Private m_objColCampos As Collection


Public Property Get EsAdministradorCalculado() As EnumSiNo

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_ObjEntorno.ColUsuariosAdministradores Is Nothing Then
        m_eEsAdministradorCalculado = EnumSiNo.No
        EsAdministradorCalculado = m_eEsAdministradorCalculado
        Exit Property
    End If
    If m_ObjEntorno.ColUsuariosAdministradores.Exists(Me.UsuarioRed) Then
        m_eEsAdministradorCalculado = EnumSiNo.Sí
    Else
        m_eEsAdministradorCalculado = EnumSiNo.No
    End If
    
    EsAdministradorCalculado = m_eEsAdministradorCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Usuario.EsAdministradorCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get ColAplicacionesPermisos() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColAplicacionesPermisos Is Nothing Then
        Set ColAplicacionesPermisos = m_objColAplicacionesPermisos
        Exit Property
    End If
    If Me.CorreoUsuario = "" Then
        Exit Property
    End If
    Set m_objColAplicacionesPermisos = Constructor.getAplicacionesPermisos(Me.CorreoUsuario, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAplicacionesPermisos = m_objColAplicacionesPermisos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Usuario.colAplicacionesPermisos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get Permisos() As UsuarioAplicacionPermisos

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objPermisos Is Nothing Then
        Set Permisos = m_objPermisos
        Exit Property
    End If
    If Me.ColAplicacionesPermisos Is Nothing Then
        If Me.Error <> "" Then
            Exit Property
        End If
        Exit Property
    End If
    If Me.ColAplicacionesPermisos.Exists(CStr(m_IDAplicacion)) Then
        Set m_objPermisos = Me.ColAplicacionesPermisos(CStr(m_IDAplicacion))
    End If
    Set Permisos = m_objPermisos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Usuario.Permisos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get EsCalidadCalculado() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_ObjEntorno.ColUsuariosCalidad Is Nothing Then
        m_eEsCalidadCalculado = EnumSiNo.No
        EsCalidadCalculado = m_eEsCalidadCalculado
        Exit Property
    End If
    If m_ObjEntorno.ColUsuariosCalidad.Exists(Me.UsuarioRed) Then
        m_eEsCalidadCalculado = EnumSiNo.Sí
    Else
        m_eEsCalidadCalculado = EnumSiNo.No
    End If
    
    EsCalidadCalculado = m_eEsCalidadCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Usuario.EsCalidadCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EsTecnicoCalculado() As EnumSiNo

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_ObjEntorno.ColUsuariosCalidad.Exists(Me.UsuarioRed) Then
        m_eEsTecnicoCalculado = EnumSiNo.No
        EsTecnicoCalculado = m_eEsTecnicoCalculado
        Exit Property
    End If
    If m_ObjEntorno.ColUsuariosAdministradores.Exists(Me.UsuarioRed) Then
        m_eEsTecnicoCalculado = EnumSiNo.No
        EsTecnicoCalculado = m_eEsTecnicoCalculado
        Exit Property
    End If
    m_eEsTecnicoCalculado = EnumSiNo.Sí
    EsTecnicoCalculado = m_eEsTecnicoCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Usuario.EsTecnicoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get ColProyectos() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColProyectos Is Nothing Then
        Set ColProyectos = m_objColProyectos
        Exit Property
    End If
    If Me.UsuarioRed = "" Then
        Me.Error = "No se conoce el Usuario de Red"
        Err.Raise 1000
    End If
    Set m_objColProyectos = Constructor.getProyectosDeUsuarios(p_NombreUsuario:=Me.Nombre, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColProyectos = m_objColProyectos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Usuario.colProyectos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Set ColProyectos(ByVal sNewValue As Variant)

    Set m_objColProyectos = sNewValue

End Property

Public Property Get ColProyectosActivos() As Scripting.Dictionary
    
    Dim m_ID As Variant
    Dim m_ObjProyecto As Proyecto
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objColProyectosActivos Is Nothing Then
        Set ColProyectosActivos = m_objColProyectosActivos
        Exit Property
    End If
    If Me.ColProyectos Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    For Each m_ID In Me.ColProyectos
        Set m_ObjProyecto = Me.ColProyectos(m_ID)
        If m_ObjProyecto.FechaCierre = "" Then
            If m_objColProyectosActivos Is Nothing Then
                Set m_objColProyectosActivos = New Scripting.Dictionary
            End If
            If Not m_objColProyectosActivos.Exists(m_ObjProyecto.IDProyecto) Then
                m_objColProyectosActivos.Add m_ObjProyecto.IDProyecto, m_ObjProyecto
            End If
        End If
        Set m_ObjProyecto = Nothing
    Next
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Usuario.ColProyectosActivos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set ColProyectosActivos(ByVal objNewValue As Variant)

    Set m_objColProyectosActivos = objNewValue

End Property


Public Property Get Activo() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_eActivo <> Empty Then
        Activo = m_eActivo
        Exit Property
    End If
    If Not IsDate(Me.FechaBaja) Then
        m_eActivo = EnumSiNo.Sí
        Activo = m_eActivo
        Exit Property
    End If
    If CDate(Date) >= CDate(Me.FechaBaja) Then
        m_eActivo = EnumSiNo.No
    Else
        m_eActivo = EnumSiNo.Sí
    End If
    Activo = m_eActivo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Usuario.Activo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
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
        .Add "ID"
        .Add "CorreoUsuario"
        .Add "UsuarioRed"
        .Add "Nombre"
        .Add "FechaBaja"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Usuario.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
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
    If m_NombrePropiedad = "UsuarioRed" Then
        getPropiedad = Me.UsuarioRed
        Exit Function
    End If
    If m_NombrePropiedad = "Nombre" Then
        getPropiedad = Me.Nombre
        Exit Function
    End If
    If m_NombrePropiedad = "FechaBaja" Then
        getPropiedad = Me.FechaBaja
        Exit Function
    End If
    If m_NombrePropiedad = "Activo" Then
        getPropiedad = Me.Activo
        Exit Function
    End If
    If m_NombrePropiedad = "Id" Then
        getPropiedad = Me.ID
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Usuario.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Usuario.getPropiedad ha devuelto el error: " & vbNewLine & Err.Description
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
    If m_NombrePropiedad = "UsuarioRed" Then
        Me.UsuarioRed = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Nombre" Then
        Me.Nombre = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaBaja" Then
        Me.FechaBaja = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Id" Then
        Me.ID = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Usuario.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Usuario.SetPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function




