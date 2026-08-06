VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCDocumentoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDDoc As String
Public IDCambio As String
Public URLDocumento As String

Private m_sURLDocumentoCalculado As String
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get URLDocumentoCalculado() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If m_sURLDocumentoCalculado <> "" Then
        URLDocumentoCalculado = m_sURLDocumentoCalculado
        Exit Property
    End If
    m_sURLDocumentoCalculado = m_ObjEntorno.URLDirectorioDocumentosCambio & fso.GetFileName(Me.URLDocumento)
    
    
    URLDocumentoCalculado = m_sURLDocumentoCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método DocumentoCambio.URLDocumentoCalculado ha devuelto el error:" & vbNewLine & Err.Description
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
        .Add "IDDoc"
        .Add "IDCambio"
        .Add "URLDocumento"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método DocumentoCambio.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDDoc" Then
        getPropiedad = Me.IDDoc
        Exit Function
    End If
    If m_NombrePropiedad = "IDCambio" Then
        getPropiedad = Me.IDCambio
        Exit Function
    End If
    If m_NombrePropiedad = "URLDocumento" Then
        getPropiedad = Me.URLDocumento
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en DocumentoCambio.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DocumentoCambio.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDDoc" Then
        Me.IDDoc = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDCambio" Then
        Me.IDCambio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "URLDocumento" Then
        Me.URLDocumento = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en DocumentoCambio.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DocumentoCambio.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


