VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Suministrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public IDSuministrador As String
Public Nombre As String
Public CIF As String
Public Descripcion As String
Public TramitadoraHPS As String
Public Nemotecnico As String
Public Direccion As String
Public CP As String
Public Ciudad As String


Private m_sNombreConCif As String
Private m_sSolicitudesHPS As String

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get NombreConCif() As String
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    m_sNombreConCif = Me.Nombre & " (" & Me.CIF & ")"
    NombreConCif = m_sNombreConCif
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Suministrador.NombreConCif ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    m_Error = Me.Error
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
        .Add "IDSuministrador"
        .Add "Nombre"
        .Add "CIF"
        .Add "DESCRIPCION"
        .Add "TramitadoraHPS"
        .Add "Nemotecnico"
        .Add "Direccion"
        .Add "CP"
        .Add "Ciudad"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Suministrador.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDSuministrador" Then
        getPropiedad = Me.IDSuministrador
        Exit Function
    End If
    If m_NombrePropiedad = "Nombre" Then
        getPropiedad = Me.Nombre
        Exit Function
    End If
    If m_NombrePropiedad = "CIF" Then
        getPropiedad = Me.CIF
        Exit Function
    End If
    If m_NombrePropiedad = "DESCRIPCION" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    If m_NombrePropiedad = "TramitadoraHPS" Then
        getPropiedad = Me.TramitadoraHPS
        Exit Function
    End If
    If m_NombrePropiedad = "Nemotecnico" Then
        getPropiedad = Me.Nemotecnico
        Exit Function
    End If
    If m_NombrePropiedad = "Direccion" Then
        getPropiedad = Me.Direccion
        Exit Function
    End If
    If m_NombrePropiedad = "CP" Then
        getPropiedad = Me.CP
        Exit Function
    End If
    If m_NombrePropiedad = "Ciudad" Then
        getPropiedad = Me.Ciudad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Suministrador.getPropiedad"
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Suministrador.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function setPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDSuministrador" Then
        Me.IDSuministrador = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Nombre" Then
        Me.Nombre = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CIF" Then
        Me.CIF = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DESCRIPCION" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TramitadoraHPS" Then
        Me.TramitadoraHPS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Nemotecnico" Then
        Me.Nemotecnico = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Direccion" Then
        Me.Direccion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CP" Then
        Me.CP = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Ciudad" Then
        Me.Ciudad = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Suministrador.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Suministrador.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function



