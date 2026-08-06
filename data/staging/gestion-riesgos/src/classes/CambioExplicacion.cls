VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CambioExplicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Apartado As String
Public Explicacion As String
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
        .Add "Apartado"
        .Add "Explicacion"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método CambiosExplicacion.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "Apartado" Then
        getPropiedad = Me.Apartado
        Exit Function
    End If
    If m_NombrePropiedad = "Explicacion" Then
        getPropiedad = Me.Explicacion
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en CambiosExplicacion.getPropiedad"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CambiosExplicacion.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "Apartado" Then
        Me.Apartado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Explicacion" Then
        Me.Explicacion = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en CambiosExplicacion.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CambiosExplicacion.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

