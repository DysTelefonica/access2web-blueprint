VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AplicacionApertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDApertura As String
Public IDAplicacion As String
Public NombreUsuario As String
Public FechaApertura As String
Public HoraApertura As String
Public FechaCierre As String
Public HoraCierre As String
Public NombreAplicacion As String
Public FechaEnvioCorreoAdministrador As String
Public EnOficina As String
Public UsuarioConectadoMaquina As String
Public VersionAplicacion As String
Public NombreMaquina As String
Public UsuarioMaquina As String
Public Observaciones As String
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

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
        .Add "IDApertura"
        .Add "IDAplicacion"
        .Add "NombreUsuario"
        .Add "FechaApertura"
        .Add "HoraApertura"
        .Add "FechaCierre"
        .Add "HoraCierre"
        .Add "NombreAplicacion"
        .Add "FechaEnvioCorreoAdministrador"
        .Add "EnOficina"
        .Add "UsuarioConectadoMaquina"
        .Add "VersionAplicacion"
        .Add "NombreMaquina"
        .Add "UsuarioMaquina"
        .Add "Observaciones"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método AplicacionesApertura.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDApertura" Then
        getPropiedad = Me.IDApertura
        Exit Function
    End If
    If m_NombrePropiedad = "IDAplicacion" Then
        getPropiedad = Me.IDAplicacion
        Exit Function
    End If
    If m_NombrePropiedad = "NombreUsuario" Then
        getPropiedad = Me.NombreUsuario
        Exit Function
    End If
    If m_NombrePropiedad = "FechaApertura" Then
        getPropiedad = Me.FechaApertura
        Exit Function
    End If
    If m_NombrePropiedad = "HoraApertura" Then
        getPropiedad = Me.HoraApertura
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCierre" Then
        getPropiedad = Me.FechaCierre
        Exit Function
    End If
    If m_NombrePropiedad = "HoraCierre" Then
        getPropiedad = Me.HoraCierre
        Exit Function
    End If
    If m_NombrePropiedad = "NombreAplicacion" Then
        getPropiedad = Me.NombreAplicacion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEnvioCorreoAdministrador" Then
        getPropiedad = Me.FechaEnvioCorreoAdministrador
        Exit Function
    End If
    If m_NombrePropiedad = "EnOficina" Then
        getPropiedad = Me.EnOficina
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioConectadoMaquina" Then
        getPropiedad = Me.UsuarioConectadoMaquina
        Exit Function
    End If
    If m_NombrePropiedad = "VersionAplicacion" Then
        getPropiedad = Me.VersionAplicacion
        Exit Function
    End If
    If m_NombrePropiedad = "NombreMaquina" Then
        getPropiedad = Me.NombreMaquina
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioMaquina" Then
        getPropiedad = Me.UsuarioMaquina
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        getPropiedad = Me.Observaciones
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en AplicacionesApertura.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AplicacionesApertura.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDApertura" Then
        Me.IDApertura = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDAplicacion" Then
        Me.IDAplicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreUsuario" Then
        Me.NombreUsuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaApertura" Then
        Me.FechaApertura = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HoraApertura" Then
        Me.HoraApertura = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCierre" Then
        Me.FechaCierre = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HoraCierre" Then
        Me.HoraCierre = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreAplicacion" Then
        Me.NombreAplicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEnvioCorreoAdministrador" Then
        Me.FechaEnvioCorreoAdministrador = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EnOficina" Then
        Me.EnOficina = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioConectadoMaquina" Then
        Me.UsuarioConectadoMaquina = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "VersionAplicacion" Then
        Me.VersionAplicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreMaquina" Then
        Me.NombreMaquina = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioMaquina" Then
        Me.UsuarioMaquina = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        Me.Observaciones = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en AplicacionesApertura.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AplicacionesApertura.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

