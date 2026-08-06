VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MaterialSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private fso As Scripting.FileSystemObject

Public IDSeguimiento As String
Public IDMaterial As String
Public FechaInicioIntervencion As String
Public DESTINO As String
Public MOTIVO As String
Public FechaFinIntervencion As String
Public RESULTADO As String
Public UltimaIntervencion As String
Public TituloAnexo As String
Public DescripcionAnexo As String
Public NombreArchivoAnexo As String

Public IDSeguimientoCalculado As String
Private m_ObjMaterial As Material
Private m_sURLAnexos As String
Private m_eAnexosCerrados As EnumSino
Private m_objColAnexos As Scripting.Dictionary

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
    If Me.IDSeguimiento = "" Then
        Exit Property
        
    End If
    Set m_objColAnexos = Constructor.getAnexosDeMaterialSeguimiento(Me.IDSeguimiento, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColAnexos = m_objColAnexos
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método MaterialSeguimiento.ColAnexos ha devuelto el error: " & vbNewLine & Err.Description
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
        m_Error = "El método Parte.AnexosCerrados ha devuelto el error: " & vbNewLine & Err.Description
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
        m_Error = "El método MaterialSeguimiento.URLAnexos ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Material() As Material
    
    
    On Error GoTo errores
    
    Me.Error = ""
    If Not m_ObjMaterial Is Nothing Then
        Set Material = m_ObjMaterial
        Exit Property
    End If
    If Me.IDMaterial = "" Then
        Exit Property
    End If
    
    Set m_ObjMaterial = Constructor.getMaterial(Me.IDMaterial, m_Error)
    
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Material = m_ObjMaterial
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método MaterialSeguimiento.Material ha devuelto el error: " & vbNewLine & Err.Description
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
        p_Error = "El método MaterialSeguimiento.EliminarAnexos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
        .Add "IDSeguimiento"
        .Add "IDMaterial"
        .Add "FechaInicioIntervencion"
        .Add "Destino"
        .Add "Motivo"
        .Add "FechaFinIntervencion"
        .Add "Resultado"
        .Add "UltimaIntervencion"
        .Add "TituloAnexo"
        .Add "DescripcionAnexo"
        .Add "NombreArchivoAnexo"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método MaterialSeguimiento.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDSeguimiento" Then
        getPropiedad = Me.IDSeguimiento
        Exit Function
    End If
    If m_NombrePropiedad = "IDMaterial" Then
        getPropiedad = Me.IDMaterial
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicioIntervencion" Then
        getPropiedad = Me.FechaInicioIntervencion
        Exit Function
    End If
    If m_NombrePropiedad = "Destino" Then
        getPropiedad = Me.DESTINO
        Exit Function
    End If
    If m_NombrePropiedad = "Motivo" Then
        getPropiedad = Me.MOTIVO
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinIntervencion" Then
        getPropiedad = Me.FechaFinIntervencion
        Exit Function
    End If
    If m_NombrePropiedad = "Resultado" Then
        getPropiedad = Me.RESULTADO
        Exit Function
    End If
    If m_NombrePropiedad = "UltimaIntervencion" Then
        getPropiedad = Me.UltimaIntervencion
        Exit Function
    End If
    If m_NombrePropiedad = "TituloAnexo" Then
        getPropiedad = Me.TituloAnexo
        Exit Function
    End If
    If m_NombrePropiedad = "DescripcionAnexo" Then
        getPropiedad = Me.DescripcionAnexo
        Exit Function
    End If
    If m_NombrePropiedad = "NombreArchivoAnexo" Then
        getPropiedad = Me.NombreArchivoAnexo
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en MaterialSeguimiento.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MaterialSeguimiento.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDSeguimiento" Then
        Me.IDSeguimiento = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDMaterial" Then
        Me.IDMaterial = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicioIntervencion" Then
        Me.FechaInicioIntervencion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Destino" Then
        Me.DESTINO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Motivo" Then
        Me.MOTIVO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinIntervencion" Then
        Me.FechaFinIntervencion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Resultado" Then
        Me.RESULTADO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UltimaIntervencion" Then
        Me.UltimaIntervencion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TituloAnexo" Then
        Me.TituloAnexo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DescripcionAnexo" Then
        Me.DescripcionAnexo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreArchivoAnexo" Then
        Me.NombreArchivoAnexo = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en MaterialSeguimiento.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MaterialSeguimiento.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function



