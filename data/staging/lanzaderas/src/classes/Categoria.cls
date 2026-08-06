VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Categoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public IDCategoria As String
Public NombreCategoria As String
Private m_eCategoriaCalculada As EnumVideoCategoria
Private m_objColCampos As Collection
Public Error As String

Public Property Get CategoriaCalculada() As EnumVideoCategoria

    On Error GoTo errores
    Me.Error = ""
    If m_eCategoriaCalculada <> 0 Then
        CategoriaCalculada = m_eCategoriaCalculada
        Exit Property
    End If
    If Me.NombreCategoria = "" Then
        Exit Property
    End If
    If m_ObjEntorno.ColCategoriasVideo Is Nothing Then
        Me.Error = m_ObjEntorno.Error
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    If Not m_ObjEntorno.ColCategoriasVideo.Exists(Me.NombreCategoria) Then
        Exit Property
    End If
    m_eCategoriaCalculada = m_ObjEntorno.ColCategoriasVideo(Me.NombreCategoria)
    
    
    CategoriaCalculada = m_eCategoriaCalculada
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Categoria.CategoriaCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property


Public Property Get ColCampos() As Collection
    
    Dim tdf As DAO.TableDef
    Dim fld As Object
    On Error GoTo errores
    
    Me.Error = ""
    If Not m_objColCampos Is Nothing Then
        Set ColCampos = m_objColCampos
        Exit Property
    End If
    Set tdf = getdb().TableDefs("TbCategorias")
    For Each fld In tdf.Fields
        If m_objColCampos Is Nothing Then
            Set m_objColCampos = New Collection
        
        End If
        m_objColCampos.Add fld.Name
    Next
    Set tdf = Nothing
    
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El metodo Categoria.ColCampos ha devuelto el error: " & vbNewLine & Err.Description
    End If


End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If m_NombrePropiedad = "IDCategoria" Then
        getPropiedad = Me.IDCategoria
        Exit Function
    End If
    If m_NombrePropiedad = "NombreCategoria" Then
        getPropiedad = Me.NombreCategoria
        Exit Function
    End If
   
    
    p_Error = "Propiedad no reconocida en Categoria.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Categoria.getPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If m_NombrePropiedad = "IDCategoria" Then
        Me.IDCategoria = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreCategoria" Then
        Me.NombreCategoria = m_ValorPropiedad
        Exit Function
    End If
        
    p_Error = "Propiedad no reconocida en Categoria.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Categoria.SetPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



