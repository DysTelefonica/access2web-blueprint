VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDCambio As String
Public IDVersion As String
Public CodCambio As String
Public NombreCambio As String
Public DescripcionCambio As String
Public EsBug As String
Public EsMejora As String
Public MostrarEninformeGeneral As String

Private m_objDocumentos As Scripting.Dictionary
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get Documentos() As Scripting.Dictionary

     On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objDocumentos Is Nothing Then
        Set Documentos = m_objDocumentos
        Exit Property
    End If
    If Me.IDCambio = "" Then
       Exit Property
    End If
    Set m_objDocumentos = Constructor.getDocumentosCCCambio(Me.IDCambio, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Documentos = m_objDocumentos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método CCCambio.Documentos ha devuelto el error: " & vbNewLine & Err.Description
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
        .Add "IDCambio"
        .Add "IDVersion"
        .Add "NombreCambio"
        .Add "CodCambio"
        .Add "DescripcionCambio"
        .Add "EsBug"
        .Add "EsMejora"
        .Add "MostrarEninformeGeneral"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método CCCambio.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDCambio" Then
        getPropiedad = Me.IDCambio
        Exit Function
    End If
    If m_NombrePropiedad = "IDVersion" Then
        getPropiedad = Me.IDVersion
        Exit Function
    End If
    If m_NombrePropiedad = "NombreCambio" Then
        getPropiedad = Me.NombreCambio
        Exit Function
    End If
    If m_NombrePropiedad = "CodCambio" Then
        getPropiedad = Me.CodCambio
        Exit Function
    End If
    If m_NombrePropiedad = "DescripcionCambio" Then
        getPropiedad = Me.DescripcionCambio
        Exit Function
    End If
    If m_NombrePropiedad = "EsBug" Then
        getPropiedad = Me.EsBug
        Exit Function
    End If
    If m_NombrePropiedad = "EsMejora" Then
        getPropiedad = Me.EsMejora
        Exit Function
    End If
    If m_NombrePropiedad = "MostrarEninformeGeneral" Then
        getPropiedad = Me.MostrarEninformeGeneral
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en CCCambio.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CCCambio.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDCambio" Then
        Me.IDCambio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDVersion" Then
        Me.IDVersion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreCambio" Then
        Me.NombreCambio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodCambio" Then
        Me.CodCambio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DescripcionCambio" Then
        Me.DescripcionCambio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsBug" Then
        Me.EsBug = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsMejora" Then
        Me.EsMejora = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "MostrarEninformeGeneral" Then
        Me.MostrarEninformeGeneral = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en CCCambio.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CCCambio.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function



