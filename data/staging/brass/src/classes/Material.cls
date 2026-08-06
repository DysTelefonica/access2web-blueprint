VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Material"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private fso As New FileSystemObject

Public IDMaterial As String
Public IDEquipo As String
Public FechaIncio As String
Public Garantia As String
Public IDActividad As String
Public TipoAccion As String
Public Material As String
Public PN As String
Public NS As String
Public COSTE As String
Public ReparadoPor As String
Public FechaEntrega As String
Public Descripcion As String
Public CodImportacion As String
Public EsReparacion As String

Private m_sIDMaterialCalculado As String
Private m_ObjActividad As Actividad
Private m_objcolSeguimientos As Scripting.Dictionary
Private m_eTieneSeguimientos As EnumSino
Private m_eEntregado As EnumSino
Private m_ObjFactura As Factura
Private m_sFechaFacturado As String
Private m_ObjEquipo As Equipo
Private m_eSeguimientoFinalizado As EnumSino
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
    If Me.IDMaterial = "" Then
        Exit Property
        
    End If
    Set m_objColAnexos = Constructor.getAnexosDeMaterial(Me.IDMaterial, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAnexos = m_objColAnexos
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Material.ColAnexos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Set ColAnexos(ByVal objNewValue As Variant)

    Set m_objColAnexos = objNewValue

End Property

Public Property Get AnexosCerrados() As EnumSino
    
    
    
    Dim m_ColAnexosDeMateriales As Scripting.Dictionary
    Dim m_ColAnexosDeMatSeguimiento As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Anexo As Anexo
    Dim m_MaterialSet As MaterialSeguimiento
    
    Dim m_URL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_eAnexosCerrados <> 0 Then
        AnexosCerrados = m_eAnexosCerrados
        Exit Property
    End If
    Set m_ColAnexosDeMateriales = Me.ColAnexos
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ColAnexosDeMateriales Is Nothing Then
        For Each m_ID In m_ColAnexosDeMateriales
            Set m_Anexo = m_ColAnexosDeMateriales(m_ID)
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
    Set m_ColAnexosDeMatSeguimiento = Me.colSeguimientos
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ColAnexosDeMatSeguimiento Is Nothing Then
        For Each m_ID In m_ColAnexosDeMatSeguimiento
            Set m_MaterialSet = m_ColAnexosDeMatSeguimiento(m_ID)
            If m_MaterialSet.AnexosCerrados = EnumSino.No Then
                m_Error = m_MaterialSet.Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                m_eAnexosCerrados = EnumSino.No
                AnexosCerrados = m_eAnexosCerrados
                Exit Property
            Else
                m_Error = m_MaterialSet.Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
            Set m_MaterialSet = Nothing
        Next
    End If
    
    m_eAnexosCerrados = EnumSino.Sí
    AnexosCerrados = m_eAnexosCerrados
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Parte.AnexosCerrados ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Entregado() As EnumSino
    
    On Error GoTo errores
    
    Me.Error = ""
    If m_eEntregado <> 0 Then
        Entregado = m_eEntregado
        Exit Property
    End If
    If IsDate(Me.FechaEntrega) Then
        m_eEntregado = EnumSino.Sí
    Else
        m_eEntregado = EnumSino.No
    End If
    
    Entregado = m_eEntregado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Material.Entregado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get IDMaterialCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sIDMaterialCalculado = DameIDMaterial(Me.IDActividad)
    
    IDMaterialCalculado = m_sIDMaterialCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Material.IDMaterialCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Actividad() As Actividad
    
    
    On Error GoTo errores
    
    Me.Error = ""
    If Not m_ObjActividad Is Nothing Then
        Set Actividad = m_ObjActividad
        Exit Property
    End If
    If Me.IDActividad = "" Then
        Exit Property
    End If
    Set m_ObjActividad = Constructor.getActividad(Me.IDActividad, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Actividad = m_ObjActividad
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Material.Actividad ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get colSeguimientos() As Scripting.Dictionary
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not m_objcolSeguimientos Is Nothing Then
        Set colSeguimientos = m_objcolSeguimientos
        Exit Property
    End If
    If Me.IDMaterial = "" Then
        Exit Property
    End If
    Set m_objcolSeguimientos = Constructor.getSeguimientos(Me.IDMaterial, m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set colSeguimientos = m_objcolSeguimientos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Material.colSeguimientos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get TieneSeguimientos() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If m_eTieneSeguimientos <> 0 Then
        TieneSeguimientos = m_eTieneSeguimientos
        Exit Property
    End If
    If Not Me.colSeguimientos Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneSeguimientos = Sí
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTieneSeguimientos = No
    End If
    TieneSeguimientos = m_eTieneSeguimientos
    
     Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Material.TieneSeguimientos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    
End Property

Public Property Get Factura() As Factura
    
    
    On Error GoTo errores
    
    Me.Error = ""
    
    
    If Not m_ObjFactura Is Nothing Then
        Set Factura = m_ObjFactura
        Exit Property
    End If
    If Me.IDMaterial = "" Then
        Exit Property
    End If
    Set m_ObjFactura = Constructor.getFacturaDeMaterial(Me.IDMaterial, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Factura = m_ObjFactura
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Material.Factura ha devuelto el error: " & vbNewLine & Err.Description
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
        Exit Property
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_sFechaFacturado = Me.Factura.FechaFactura
    End If
    FechaFacturado = m_sFechaFacturado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Material.FechaFacturado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Equipo() As Equipo
    Dim m_Evento As Evento
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_ObjEquipo Is Nothing Then
        Set Equipo = m_ObjEquipo
        Exit Property
    End If
    If Me.Actividad Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Set m_Evento = Me.Actividad.Evento
    m_Error = Me.Actividad.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Evento Is Nothing Then
        Exit Property
    End If
    Set m_ObjEquipo = m_Evento.Equipo
    m_Error = m_Evento.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set Equipo = m_ObjEquipo
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Material.Equipo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Property Get SeguimientoFinalizado() As EnumSino
    
    Dim m_ID As Variant
    Dim m_Seg As MaterialSeguimiento
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    
    If m_eSeguimientoFinalizado <> 0 Then
        SeguimientoFinalizado = m_eSeguimientoFinalizado
        Exit Property
    End If
    If Me.colSeguimientos Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    Else
         m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    For Each m_ID In Me.colSeguimientos
        Set m_Seg = Me.colSeguimientos(m_ID)
        If m_Seg.UltimaIntervencion = "Sí" Then
            m_eSeguimientoFinalizado = EnumSino.Sí
            SeguimientoFinalizado = m_eSeguimientoFinalizado
            Exit Property
        End If
        Set m_Seg = Nothing
    Next
    m_eSeguimientoFinalizado = EnumSino.No
    SeguimientoFinalizado = m_eSeguimientoFinalizado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Material.SeguimientoFinalizado ha devuelto el error: " & vbNewLine & Err.Description
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
        m_Error = "El método Material.URLAnexos ha devuelto el error: " & vbNewLine & Err.Description
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
        p_Error = "El método Material.EliminarAnexos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
        .Add "IDMaterial"
        .Add "IDEquipo"
        .Add "FechaIncio"
        .Add "Garantia"
        .Add "IDActividad"
        .Add "TipoAccion"
        .Add "Material"
        .Add "PN"
        .Add "NS"
        .Add "COSTE"
        .Add "ReparadoPor"
        .Add "FechaEntrega"
        .Add "Descripcion"
        .Add "CodImportacion"
        .Add "EsReparacion"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Material.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDMaterial" Then
        getPropiedad = Me.IDMaterial
        Exit Function
    End If
    If m_NombrePropiedad = "IDEquipo" Then
        getPropiedad = Me.IDEquipo
        Exit Function
    End If
    If m_NombrePropiedad = "FechaIncio" Then
        getPropiedad = Me.FechaIncio
        Exit Function
    End If
    If m_NombrePropiedad = "Garantia" Then
        getPropiedad = Me.Garantia
        Exit Function
    End If
    If m_NombrePropiedad = "IDActividad" Then
        getPropiedad = Me.IDActividad
        Exit Function
    End If
    If m_NombrePropiedad = "TipoAccion" Then
        getPropiedad = Me.TipoAccion
        Exit Function
    End If
    If m_NombrePropiedad = "Material" Then
        getPropiedad = Me.Material
        Exit Function
    End If
    If m_NombrePropiedad = "PN" Then
        getPropiedad = Me.PN
        Exit Function
    End If
    If m_NombrePropiedad = "NS" Then
        getPropiedad = Me.NS
        Exit Function
    End If
    If m_NombrePropiedad = "COSTE" Then
        getPropiedad = Me.COSTE
        Exit Function
    End If
    If m_NombrePropiedad = "ReparadoPor" Then
        getPropiedad = Me.ReparadoPor
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEntrega" Then
        getPropiedad = Me.FechaEntrega
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    If m_NombrePropiedad = "CodImportacion" Then
        getPropiedad = Me.CodImportacion
        Exit Function
    End If
    If m_NombrePropiedad = "EsReparacion" Then
        getPropiedad = Me.EsReparacion
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Material.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Material.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDMaterial" Then
        Me.IDMaterial = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEquipo" Then
        Me.IDEquipo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaIncio" Then
        Me.FechaIncio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Garantia" Then
        Me.Garantia = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDActividad" Then
        Me.IDActividad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TipoAccion" Then
        Me.TipoAccion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Material" Then
        Me.Material = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PN" Then
        Me.PN = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NS" Then
        Me.NS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "COSTE" Then
        Me.COSTE = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ReparadoPor" Then
        Me.ReparadoPor = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEntrega" Then
        Me.FechaEntrega = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodImportacion" Then
        Me.CodImportacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsReparacion" Then
        Me.EsReparacion = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Material.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Material.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function





