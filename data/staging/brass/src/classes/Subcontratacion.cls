VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Subcontratacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private fso As New Scripting.FileSystemObject

Public IDSubContratacion As String
Public FechaAlta As String
Public FechaFin As String
Public Empresa As String
Public Descripcion As String
Public Importe As String
Public Observaciones As String

Private m_sFechaFacturado As String
Private m_ObjFactura As Factura
Private m_objColAnexos As Scripting.Dictionary
Private m_eAnexosCerrados As EnumSino
Private m_sURLAnexos As String

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get ColAnexos() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    If Not m_objColAnexos Is Nothing Then
        Set ColAnexos = m_objColAnexos
        Exit Property
    End If
    If Me.IDSubContratacion = "" Then
        Exit Property
        
    End If
    Set m_objColAnexos = Constructor.getAnexosDeSubContratacion(Me.IDSubContratacion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAnexos = m_objColAnexos
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método subContratacion.ColAnexos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Set ColAnexos(ByVal objNewValue As Variant)

    Set m_objColAnexos = objNewValue

End Property

Public Property Get AnexosCerrados() As EnumSino
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Anexo As Anexo
    Dim m_URL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_eAnexosCerrados <> 0 Then
        AnexosCerrados = m_eAnexosCerrados
        Exit Property
    End If
    Set m_Col = Me.ColAnexos
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_Anexo = m_Col(m_ID)
            If fso.FileExists(m_Anexo.URLAnexo) Then
                If FicheroAbierto(m_Anexo.URLAnexo) Then
                    m_eAnexosCerrados = EnumSino.No
                    AnexosCerrados = m_eAnexosCerrados
                    Exit Property
                End If
            End If
            Set m_Anexo = Nothing
        Next
    End If
    m_eAnexosCerrados = EnumSino.Sí
    AnexosCerrados = m_eAnexosCerrados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método subContratacion.AnexosCerrados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get URLAnexos() As String
    
    Dim m_Anexo As Anexo
    Dim m_ID As Variant
    Dim m_Col As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    Set m_Col = Me.ColAnexos
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Property
    End If
    For Each m_ID In m_Col
        Set m_Anexo = m_Col(m_ID)
        If m_sURLAnexos = "" Then
            m_sURLAnexos = m_Anexo.URLAnexo
        Else
            m_sURLAnexos = m_sURLAnexos & ";" & m_Anexo.URLAnexo
        End If
        Set m_Anexo = Nothing
    Next
    URLAnexos = m_sURLAnexos
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método subContratacion.URLAnexos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get FechaFacturado() As String

    
    
    On Error GoTo errores
    
    Me.Error = ""
    
    
    If IsDate(m_sFechaFacturado) Then
        FechaFacturado = m_sFechaFacturado
        Exit Property
    End If
    If Me.Factura Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 100
        End If
        m_sFechaFacturado = Me.Factura.FechaFactura
        
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 100
        End If
    End If
    FechaFacturado = m_sFechaFacturado
        
   
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método subContratacion.FechaFacturado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get Factura() As Factura
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_ObjFactura Is Nothing Then
        Set Factura = m_ObjFactura
        Exit Property
    End If
    If Me.IDSubContratacion = "" Then
        Exit Property
    End If
    Set m_ObjFactura = Constructor.getFacturaDeSubContratacion(Me.IDSubContratacion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Factura = m_ObjFactura
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Subcontratacion.Factura ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function EliminarAnexos( _
                                Optional p_Error As String _
                                ) As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Anexo As Anexo
    Dim m_AnexosCerrados As EnumSino
    On Error GoTo errores
    
    Me.Error = ""
    
    Set m_Col = Me.ColAnexos
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    m_AnexosCerrados = Me.AnexosCerrados
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_AnexosCerrados <> EnumSino.Sí Then
        p_Error = "Al menos un anexo está abierto"
        Err.Raise 1000
    End If
    For Each m_ID In m_Col
        Set m_Anexo = m_Col(m_ID)
        m_Anexo.Eliminar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_Anexo = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Subcontratacion.EliminarAnexos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

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
        .Add "IDSubcontratacion"
        .Add "FechaAlta"
        .Add "FechaFin"
        .Add "Empresa"
        .Add "Descripcion"
        .Add "Importe"
        .Add "Observaciones"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Subcontratacion.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDSubcontratacion" Then
        getPropiedad = Me.IDSubContratacion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAlta" Then
        getPropiedad = Me.FechaAlta
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFin" Then
        getPropiedad = Me.FechaFin
        Exit Function
    End If
    If m_NombrePropiedad = "Empresa" Then
        getPropiedad = Me.Empresa
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    If m_NombrePropiedad = "Importe" Then
        getPropiedad = Me.Importe
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        getPropiedad = Me.Observaciones
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Subcontratacion.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Subcontratacion.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDSubcontratacion" Then
        Me.IDSubContratacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAlta" Then
        Me.FechaAlta = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFin" Then
        Me.FechaFin = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Empresa" Then
        Me.Empresa = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Importe" Then
        Me.Importe = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        Me.Observaciones = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Subcontratacion.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Subcontratacion.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

