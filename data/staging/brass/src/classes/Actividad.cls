VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Actividad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private fso As New Scripting.FileSystemObject

Public IDActividad As String
Public IDEVENTO As String
Public ALIASTECNICO As String
Public Actividad As String
Public FechaAlta As String
Public HorasLaborables As String
Public HorasExtras As String
Public Ubicacion As String
Public Descripcion As String
Public IDEventoGenerado As String
Public IDExportacion As String
Public CodImportacion As String
Public IDFacturacion As String
Public Observaciones As String

Private m_sNombreTecnico As String
Private m_ObjFactura As Factura
Private m_sFechaFacturada As String
Private m_ObjColMateriales As Scripting.Dictionary
Public ACTIVIDADNOPROGRAMADA As String
Private m_ObjEvento As Evento
Public m_ObjTecnico As Tecnico
Private m_objColAnexos As Scripting.Dictionary
Private m_sIDActividadCalculada As String
Private m_eAnexosCerrados As EnumSino
Private m_sURLAnexos As String

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

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
        m_Error = "El método Actividad.URLAnexos ha devuelto el error: " & vbNewLine & Err.Description
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
    If Me.IDActividad = "" Then
        Exit Property
        
    End If
    Set m_objColAnexos = Constructor.getAnexosDeActividad(Me.IDActividad, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAnexos = m_objColAnexos
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Actividad.ColAnexos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Set ColAnexos(ByVal objNewValue As Variant)

    Set m_objColAnexos = objNewValue

End Property

Public Property Get AnexosCerrados() As EnumSino
    
    
    Dim m_ColAnexosDeActividades As Scripting.Dictionary
    Dim m_ColAnexosDeMateriales As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Anexo As Anexo
    Dim m_Material As Material
    
    Dim m_URL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_eAnexosCerrados <> 0 Then
        AnexosCerrados = m_eAnexosCerrados
        Exit Property
    End If
    Set m_ColAnexosDeActividades = Me.ColAnexos
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ColAnexosDeActividades Is Nothing Then
        For Each m_ID In m_ColAnexosDeActividades
            Set m_Anexo = m_ColAnexosDeActividades(m_ID)
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
    Set m_ColAnexosDeMateriales = Me.ColMateriales
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ColAnexosDeMateriales Is Nothing Then
        For Each m_ID In m_ColAnexosDeMateriales
            Set m_Material = m_ColAnexosDeMateriales(m_ID)
            If m_Material.AnexosCerrados = EnumSino.No Then
                m_Error = m_Material.Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                m_eAnexosCerrados = EnumSino.No
                AnexosCerrados = m_eAnexosCerrados
                Exit Property
            Else
                m_Error = m_Material.Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
            Set m_Material = Nothing
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


Public Property Get ColMateriales() As Scripting.Dictionary

    On Error GoTo errores
    
    Me.Error = ""
    
    If Not m_ObjColMateriales Is Nothing Then
        Set ColMateriales = m_ObjColMateriales
        Exit Property
    End If
    If Me.IDActividad = "" Then
        Exit Property
    End If
    
    
    Set m_ObjColMateriales = Constructor.getMaterialesEnActividad(Me.IDActividad, m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColMateriales = m_ObjColMateriales
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Actividad.ColMateriales ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property




Public Property Get Evento() As Evento
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjEvento Is Nothing Then
        Set Evento = m_ObjEvento
        Exit Property
    End If
    If Me.IDEVENTO = "" Then
        Exit Property
    End If
   
    Set m_ObjEvento = Constructor.getEvento(Me.IDEVENTO, m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Evento = m_ObjEvento
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Actividad.Evento ha devuelto el error: " & vbNewLine & Err.Description
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
    If Me.IDActividad = "" Then
        Exit Property
    End If
   
    Set m_ObjFactura = Constructor.getFacturaDeActividad(Me.IDActividad, m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Factura = m_ObjFactura
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Actividad.Factura ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get FechaFacturada() As String

    
    On Error GoTo errores
    
    Me.Error = ""
    
    If IsDate(m_sFechaFacturada) Then
        FechaFacturada = m_sFechaFacturada
        Exit Property
    End If
    If Me.Factura Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_sFechaFacturada = Me.Factura.FechaFactura
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    FechaFacturada = m_sFechaFacturada
        
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Actividad.FechaFacturada ha devuelto el error: " & vbNewLine & Err.Description
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
    If Me.ALIASTECNICO = "" Then
        Exit Property
    End If
    Set m_ObjTecnico = Constructor.getTecnico(Me.ALIASTECNICO, m_Error)
    If m_Error <> "" Then
        Err.r1000
    End If
    Set Tecnico = m_ObjTecnico
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Actividad.Tecnico ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get NombreTecnico() As String

    On Error GoTo errores
    
    Me.Error = ""
    If Not m_sNombreTecnico = "" Then
        NombreTecnico = m_sNombreTecnico
        Exit Property
    End If
    
    If Me.Tecnico Is Nothing Then
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
        m_sNombreTecnico = Me.Tecnico.Nombre
    End If
    
    NombreTecnico = m_sNombreTecnico
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Actividad.NombreTecnico ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get IDActividadCalculada() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sIDActividadCalculada = DameIDActividad(Me.ALIASTECNICO, Me.FechaAlta)
    
    IDActividadCalculada = m_sIDActividadCalculada
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Actividad.IDActividadCalculada ha devuelto el error: " & vbNewLine & Err.Description
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
        p_Error = "El método Actividad.EliminarAnexos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
        .Add "IDActividad"
        .Add "IDEvento"
        .Add "ALIASTECNICO"
        .Add "Actividad"
        .Add "FechaAlta"
        .Add "HorasLaborables"
        .Add "HorasExtras"
        .Add "Ubicacion"
        .Add "DESCRIPCION"
        .Add "IDEVENTOGENERADO"
        .Add "IDExportacion"
        .Add "CodImportacion"
        .Add "IDFacturacion"
        .Add "Observaciones"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Actividade.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDActividad" Then
        getPropiedad = Me.IDActividad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEvento" Then
        getPropiedad = Me.IDEVENTO
        Exit Function
    End If
    If m_NombrePropiedad = "ALIASTECNICO" Then
        getPropiedad = Me.ALIASTECNICO
        Exit Function
    End If
    If m_NombrePropiedad = "Actividad" Then
        getPropiedad = Me.Actividad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAlta" Then
        getPropiedad = Me.FechaAlta
        Exit Function
    End If
    If m_NombrePropiedad = "HorasLaborables" Then
        getPropiedad = Me.HorasLaborables
        Exit Function
    End If
    If m_NombrePropiedad = "HorasExtras" Then
        getPropiedad = Me.HorasExtras
        Exit Function
    End If
    If m_NombrePropiedad = "Ubicacion" Then
        getPropiedad = Me.Ubicacion
        Exit Function
    End If
    If m_NombrePropiedad = "DESCRIPCION" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    If m_NombrePropiedad = "IDEVENTOGENERADO" Then
        getPropiedad = Me.IDEventoGenerado
        Exit Function
    End If
    If m_NombrePropiedad = "IDExportacion" Then
        getPropiedad = Me.IDExportacion
        Exit Function
    End If
    If m_NombrePropiedad = "CodImportacion" Then
        getPropiedad = Me.CodImportacion
        Exit Function
    End If
    If m_NombrePropiedad = "IDFacturacion" Then
        getPropiedad = Me.IDFacturacion
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        getPropiedad = Me.Observaciones
        Exit Function
    End If
    If m_NombrePropiedad = "NombreTecnico" Then
        getPropiedad = Me.NombreTecnico
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Actividade.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Actividade.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDActividad" Then
        Me.IDActividad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEvento" Then
        Me.IDEVENTO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ALIASTECNICO" Then
        Me.ALIASTECNICO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Actividad" Then
        Me.Actividad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAlta" Then
        Me.FechaAlta = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HorasLaborables" Then
        Me.HorasLaborables = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HorasExtras" Then
        Me.HorasExtras = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Ubicacion" Then
        Me.Ubicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DESCRIPCION" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEVENTOGENERADO" Then
        Me.IDEventoGenerado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDExportacion" Then
        Me.IDExportacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodImportacion" Then
        Me.CodImportacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDFacturacion" Then
        Me.IDFacturacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        Me.Observaciones = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Actividade.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Actividade.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function






