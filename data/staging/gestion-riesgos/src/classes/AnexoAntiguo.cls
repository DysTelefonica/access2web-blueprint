VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnexoAntiguo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDAnexo As String
Public IDProyecto As String
Public IDEdicion As String
Public FechaAnexo As String
Public CodigoUnico As String
Public Titulo As String
Public Descripcion As String
Public NombreArchivo As String
Public EvidenciaUTE As String


Private m_ObjProyecto As Proyecto
Private m_objEdicion As Edicion
Private m_sURLAnexo As String
Private m_sError As String
Private m_sEvidenciaUTE As String
Private m_ObjRiesgo As riesgo

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get Proyecto() As Proyecto
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjProyecto Is Nothing Then
        Set Proyecto = m_ObjProyecto
        Exit Property
    End If
    If Me.IDProyecto = "" Then
        Exit Property
    End If
    
    Set m_ObjProyecto = Constructor.getProyecto(Me.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Proyecto = m_ObjProyecto
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método AnexoAntiguo.Proyecto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get Edicion() As Edicion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not m_objEdicion Is Nothing Then
        Set Edicion = m_objEdicion
        Exit Property
    End If
    If Me.IDEdicion = "" Then
        m_Error = "No se ha podido determinar la Edición del Riesgo"
        Err.Raise 1000
    End If
    Set m_objEdicion = Constructor.getEdicion(Me.IDEdicion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Edicion = m_objEdicion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método AnexoAntiguo.Edicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Set Edicion(ByVal sNewValue As Variant)

    m_objEdicion = sNewValue

End Property

Public Property Get riesgo() As riesgo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjRiesgo Is Nothing Then
        Set riesgo = m_ObjRiesgo
        Exit Property
    End If
    If Me.IDEdicion = "" Then
        m_Error = "No se ha podido determinar la Edición del Riesgo"
        Err.Raise 1000
    End If
    If Me.CodigoUnico = "" Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    Set m_ObjRiesgo = Constructor.getRiesgo(, Me.IDEdicion, Me.CodigoUnico, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set riesgo = m_ObjRiesgo
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Edicion.AnexoAntiguo ha devuelto el error: " & vbNewLine & Err.Description
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
        .Add "IDAnexo"
        .Add "IDProyecto"
        .Add "IDEdicion"
        .Add "FechaAnexo"
        .Add "CodigoUnico"
        .Add "Titulo"
        .Add "Descripcion"
        .Add "NombreArchivo"
        .Add "EvidenciaUTE"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método AnexosAntigua.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDAnexo" Then
        getPropiedad = Me.IDAnexo
        Exit Function
    End If
    If m_NombrePropiedad = "IDProyecto" Then
        getPropiedad = Me.IDProyecto
        Exit Function
    End If
    If m_NombrePropiedad = "IDEdicion" Then
        getPropiedad = Me.IDEdicion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAnexo" Then
        getPropiedad = Me.FechaAnexo
        Exit Function
    End If
    If m_NombrePropiedad = "CodigoUnico" Then
        getPropiedad = Me.CodigoUnico
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo" Then
        getPropiedad = Me.Titulo
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    If m_NombrePropiedad = "NombreArchivo" Then
        getPropiedad = Me.NombreArchivo
        Exit Function
    End If
    If m_NombrePropiedad = "EvidenciaUTE" Then
        getPropiedad = Me.EvidenciaUTE
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en AnexosAntigua.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexosAntigua.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDAnexo" Then
        Me.IDAnexo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDProyecto" Then
        Me.IDProyecto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEdicion" Then
        Me.IDEdicion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAnexo" Then
        Me.FechaAnexo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodigoUnico" Then
        Me.CodigoUnico = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo" Then
        Me.Titulo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreArchivo" Then
        Me.NombreArchivo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EvidenciaUTE" Then
        Me.EvidenciaUTE = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en AnexosAntigua.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexosAntigua.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

