VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDVersion As String
Public IDAplicacion As String
Public Version As String
Public Descripcion As String
Public FechaInicio As String
Public fechaFin As String
Public VersionPrincipal As String
Public URLRequisitosPrevios As String

Private m_eHayCambiosDeMejoras As EnumSiNo
Private m_eHayCambiosDeErrores As EnumSiNo
Private m_sCambiosDeErrores As String
Private m_sCambiosDeMejoras As String
Private m_objCambios As Scripting.Dictionary
Private m_sURLIcono As String


Private m_sURLRequisitosPreviosCalculada As String
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get URLIcono() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If m_sURLIcono <> "" Then
        URLIcono = m_sURLIcono
        Exit Property
    End If
    
    If Me.URLRequisitosPrevios = "" Then
        Exit Property
    End If
    If fso.GetExtensionName(Me.URLRequisitosPrevios) = "zip" Then
        m_sURLIcono = m_ObjEntorno.URLIconoZIP
    Else
         m_sURLIcono = m_ObjEntorno.URLIconoPDF
    End If
    URLIcono = m_sURLIcono
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método DocumentoCambio.URLIcono ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLRequisitosPreviosCalculada() As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If m_sURLRequisitosPreviosCalculada <> "" Then
        URLRequisitosPreviosCalculada = m_sURLRequisitosPreviosCalculada
        Exit Property
    End If
    If Me.URLRequisitosPrevios = "" Then
        Exit Property
    End If
    m_sURLRequisitosPreviosCalculada = m_ObjEntorno.URLDirectorioDocumentosRequisitosCambio & fso.GetFileName(Me.URLRequisitosPrevios)
    
    
    URLRequisitosPreviosCalculada = m_sURLRequisitosPreviosCalculada
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método DocumentoCambio.URLRequisitosPreviosCalculada ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get HayCambiosDeMejoras() As EnumSiNo
    
    Dim m_ID As Variant
    Dim m_Cambio As CCCambio
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If m_eHayCambiosDeMejoras <> Empty Then
        HayCambiosDeMejoras = m_eHayCambiosDeMejoras
        Exit Property
    End If
    If Me.Cambios Is Nothing Then
        m_eHayCambiosDeMejoras = EnumSiNo.No
        HayCambiosDeMejoras = m_eHayCambiosDeMejoras
        Exit Property
    End If
    For Each m_ID In Me.Cambios
        Set m_Cambio = Me.Cambios(m_ID)
        With m_Cambio
            If .EsMejora = "Sí" Then
                m_eHayCambiosDeMejoras = EnumSiNo.Sí
                HayCambiosDeMejoras = m_eHayCambiosDeMejoras
                Exit Property
            End If
        End With
        Set m_Cambio = Nothing
    Next
    m_eHayCambiosDeMejoras = EnumSiNo.No
    HayCambiosDeMejoras = m_eHayCambiosDeMejoras
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método CCVersion.HayCambiosDeMejoras ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get HayCambiosDeErrores() As EnumSiNo

    Dim m_ID As Variant
    Dim m_Cambio As CCCambio
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If m_eHayCambiosDeErrores <> Empty Then
        HayCambiosDeErrores = m_eHayCambiosDeErrores
        Exit Property
    End If
    If Me.Cambios Is Nothing Then
        m_eHayCambiosDeErrores = EnumSiNo.No
        HayCambiosDeErrores = m_eHayCambiosDeErrores
        Exit Property
    End If
    For Each m_ID In Me.Cambios
        Set m_Cambio = Me.Cambios(m_ID)
        With m_Cambio
            If .EsBug = "Sí" Then
                m_eHayCambiosDeErrores = EnumSiNo.Sí
                HayCambiosDeErrores = m_eHayCambiosDeErrores
                Exit Property
            End If
        End With
        Set m_Cambio = Nothing
    Next
    m_eHayCambiosDeErrores = EnumSiNo.No
    HayCambiosDeErrores = m_eHayCambiosDeErrores
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método CCVersion.HayCambiosDeErrores ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property

Public Property Get CambiosDeErrores() As String
    
    Dim m_ID As Variant
    Dim m_Cambio As CCCambio
    Dim intCambio As Integer
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sCambiosDeErrores) Then
        CambiosDeErrores = m_sCambiosDeErrores
        Exit Property
    End If
    
    If Me.Cambios Is Nothing Then
        m_sCambiosDeErrores = "0"
        CambiosDeErrores = m_sCambiosDeErrores
        Exit Property
    End If
    For Each m_ID In Me.Cambios
        Set m_Cambio = Me.Cambios(m_ID)
        With m_Cambio
            If .EsBug = "Sí" Then
                intCambio = intCambio + 1
            End If
        End With
        Set m_Cambio = Nothing
    Next
    m_sCambiosDeErrores = CStr(intCambio)
    CambiosDeErrores = m_sCambiosDeErrores
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método CCVersion.CambiosDeErrores ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get CambiosDeMejoras() As String
    
    Dim m_ID As Variant
    Dim m_Cambio As CCCambio
    Dim intCambio As Integer
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If IsNumeric(m_sCambiosDeMejoras) Then
        CambiosDeMejoras = m_sCambiosDeMejoras
        Exit Property
    End If
    
    If Me.Cambios Is Nothing Then
        m_sCambiosDeMejoras = "0"
        CambiosDeMejoras = m_sCambiosDeMejoras
        Exit Property
    End If
    For Each m_ID In Me.Cambios
        Set m_Cambio = Me.Cambios(m_ID)
        With m_Cambio
            If .EsMejora = "Sí" Then
                intCambio = intCambio + 1
            End If
        End With
        Set m_Cambio = Nothing
    Next
    m_sCambiosDeMejoras = CStr(intCambio)
    CambiosDeMejoras = m_sCambiosDeMejoras
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método CCVersion.CambiosDeMejoras ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get Cambios() As Scripting.Dictionary
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objCambios Is Nothing Then
        Set Cambios = m_objCambios
        Exit Property
    End If
    
    If Me.IDVersion = "" Then
        Exit Property
    End If
    Set m_objCambios = Constructor.getCCCambios(Me.IDVersion, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Cambios = m_objCambios
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método CCVersion.m_objCambios ha devuelto el error:" & vbNewLine & Err.Description
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
        .Add "IDVersion"
        .Add "IDAplicacion"
        .Add "Version"
        .Add "Descripcion"
        .Add "FechaInicio"
        .Add "FechaFin"
        .Add "VersionPrincipal"
        .Add "URLRequisitosPrevios"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método CCVersion.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDVersion" Then
        getPropiedad = Me.IDVersion
        Exit Function
    End If
    If m_NombrePropiedad = "IDAplicacion" Then
        getPropiedad = Me.IDAplicacion
        Exit Function
    End If
    If m_NombrePropiedad = "Version" Then
        getPropiedad = Me.Version
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicio" Then
        getPropiedad = Me.FechaInicio
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFin" Then
        getPropiedad = Me.fechaFin
        Exit Function
    End If
    If m_NombrePropiedad = "VersionPrincipal" Then
        getPropiedad = Me.VersionPrincipal
        Exit Function
    End If
    If m_NombrePropiedad = "URLRequisitosPrevios" Then
        getPropiedad = Me.URLRequisitosPrevios
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en CCVersion.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CCVersion.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDVersion" Then
        Me.IDVersion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDAplicacion" Then
        Me.IDAplicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Version" Then
        Me.Version = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicio" Then
        Me.FechaInicio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFin" Then
        Me.fechaFin = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "VersionPrincipal" Then
        Me.VersionPrincipal = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "URLRequisitosPrevios" Then
        Me.URLRequisitosPrevios = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en CCVersion.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CCVersion.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function



