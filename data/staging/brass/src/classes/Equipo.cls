VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Equipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public IDEquipo As String
Public NODO As String
Public BUI As String
Public SubSistema As String
Public Equipo As String
Public Descripcion As String
Public FechaObsoleto As String

Private m_eEquipoActivo As EnumSino
Private m_sIDEquipoCalculado As String
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get IDEquipoCalculado() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sIDEquipoCalculado = DameID("TbEquipo", "IDEquipo")
    
    IDEquipoCalculado = m_sIDEquipoCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Equipo.IDEquipoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get EquipoActivo() As EnumSino
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If IsDate(Me.FechaObsoleto) Then
        m_eEquipoActivo = EnumSino.No
    Else
        m_eEquipoActivo = EnumSino.Sí
    End If
    
    EquipoActivo = m_eEquipoActivo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Equipo.EquipoActivo ha devuelto el error: " & vbNewLine & Err.Description
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
        .Add "IDEquipo"
        .Add "NODO"
        .Add "BUI"
        .Add "SUBSISTEMA"
        .Add "Equipo"
        .Add "Descripcion"
        .Add "FechaObsoleto"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Equipo.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDEquipo" Then
        getPropiedad = Me.IDEquipo
        Exit Function
    End If
    If m_NombrePropiedad = "NODO" Then
        getPropiedad = Me.NODO
        Exit Function
    End If
    If m_NombrePropiedad = "BUI" Then
        getPropiedad = Me.BUI
        Exit Function
    End If
    If m_NombrePropiedad = "SUBSISTEMA" Then
        getPropiedad = Me.SubSistema
        Exit Function
    End If
    If m_NombrePropiedad = "Equipo" Then
        getPropiedad = Me.Equipo
        Exit Function
    End If
    If m_NombrePropiedad = "Equipo" Then
        getPropiedad = Me.Equipo
        Exit Function
    End If
    If m_NombrePropiedad = "FechaObsoleto" Then
        getPropiedad = Me.FechaObsoleto
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Equipo.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Equipo.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDEquipo" Then
        Me.IDEquipo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NODO" Then
        Me.NODO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "BUI" Then
        Me.BUI = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "SUBSISTEMA" Then
        Me.SubSistema = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Equipo" Then
        Me.Equipo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaObsoleto" Then
        Me.FechaObsoleto = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Equipo.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Equipo.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function




