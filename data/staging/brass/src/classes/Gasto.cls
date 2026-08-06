VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Gasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private fso As New Scripting.FileSystemObject

Public IDGasto As String
Public TipoGasto As String
Public ImporteUnitario As String
Public NumeroUnidades As String
Public FechaImputacionGasto As String
Public Alias As String
Public Observaciones As String


Private m_sIDGastoCalculado As String
Private m_ObjFactura As Factura
Private m_ObjTecnico As Tecnico
Private m_objColAnexos As Scripting.Dictionary
Private m_sURLAnexos As String
Private m_eAnexosCerrados As EnumSino
Private m_sFechaFacturado As String

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

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
        m_Error = "El método Gasto.AnexosCerrados ha devuelto el error: " & vbNewLine & Err.Description
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
        m_Error = "El método Gasto.URLAnexos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get ColAnexos() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    If Not m_objColAnexos Is Nothing Then
        Set ColAnexos = m_objColAnexos
        Exit Property
    End If
    If Me.IDGasto = "" Then
        Exit Property
        
    End If
    Set m_objColAnexos = Constructor.getAnexosDeGastos(Me.IDGasto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAnexos = m_objColAnexos
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Gasto.ColAnexos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property


Public Property Get IDGastoCalculado() As String

    IDGastoCalculado = m_sIDGastoCalculado

End Property

Public Property Get Factura() As Factura

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjFactura Is Nothing Then
        Set Factura = m_ObjFactura
        Exit Property
    End If
    If Me.IDGasto = "" Then
        Exit Property
    End If
   
    Set m_ObjFactura = Constructor.getFacturaDeGasto(Me.IDGasto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Factura = m_ObjFactura
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Gasto.Factura ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get Tecnico() As Tecnico

    On Error GoTo errores
    
    Me.Error = ""
    
    If Not m_ObjTecnico Is Nothing Then
        Set Tecnico = m_ObjTecnico
        Exit Property
    End If
    If Me.Alias = "" Then
        Exit Property
    End If
    Set m_ObjTecnico = Constructor.getTecnico(Me.Alias, m_Error)
    If m_Error <> "" Then
        Err.r1000
    End If
    Set Tecnico = m_ObjTecnico
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Gasto.Tecnico ha devuelto el error: " & vbNewLine & Err.Description
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
            Err.Raise 1000
        End If
        m_sFechaFacturado = Me.Factura.FechaFactura
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    FechaFacturado = m_sFechaFacturado
        
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Gasto.FechaFacturado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
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
        p_Error = "El método Gasto.EliminarAnexos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Public Function Eliminar( _
                            Optional p_Error As String _
                            ) As String
    
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    
    EliminarAnexos p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.FechaFacturado <> "" Then
        p_Error = "Está facturado"
        Err.Raise 1000
    End If
    m_SQL = "DELETE * " & _
            "FROM TbGastos " & _
            "WHERE IDGasto=" & Me.IDGasto & ";"
    getdb().Execute m_SQL
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Gasto.Eliminar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
        .Add "IDGasto"
        .Add "TipoGasto"
        .Add "ImporteUnitario"
        .Add "NumeroUnidades"
        .Add "FechaImputacionGasto"
        .Add "ALIAS"
        .Add "Observaciones"
        
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Gasto.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDGasto" Then
        getPropiedad = Me.IDGasto
        Exit Function
    End If
    If m_NombrePropiedad = "TipoGasto" Then
        getPropiedad = Me.TipoGasto
        Exit Function
    End If
    If m_NombrePropiedad = "ImporteUnitario" Then
        getPropiedad = Me.ImporteUnitario
        Exit Function
    End If
    If m_NombrePropiedad = "NumeroUnidades" Then
        getPropiedad = Me.NumeroUnidades
        Exit Function
    End If
    If m_NombrePropiedad = "FechaImputacionGasto" Then
        getPropiedad = Me.FechaImputacionGasto
        Exit Function
    End If
    If m_NombrePropiedad = "ALIAS" Then
        getPropiedad = Me.Alias
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        getPropiedad = Me.Observaciones
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Gasto.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Gasto.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDGasto" Then
        Me.IDGasto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TipoGasto" Then
        Me.TipoGasto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ImporteUnitario" Then
        Me.ImporteUnitario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NumeroUnidades" Then
        Me.NumeroUnidades = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaImputacionGasto" Then
        Me.FechaImputacionGasto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ALIAS" Then
        Me.Alias = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        Me.Observaciones = m_ValorPropiedad
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Gasto.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Gasto.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

