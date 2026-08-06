VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Visionado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public IDVisionado As String
Public IDVideo As String
Public TiempoVisionado As String
Public IDUsuario As String
Public FechaVisionado As String
Public TiempoVideo As String

Private m_sIDVisionadoCalculado As String
Private m_objVideo As Video
Private m_ObjUsuario As usuario
Private m_objColCampos As Collection
Private m_sPorcentaje As String
Public Error As String

Public Property Get Video() As Video

    On Error GoTo errores
    Me.Error = ""
    If Not m_objVideo Is Nothing Then
        Set Video = m_objVideo
        Exit Property
    End If
    If Me.IDVideo = "" Then
        Exit Property
    End If
    Set m_objVideo = Constructor.getVideo(Me.IDVideo, Me.Error)
    Set Video = m_objVideo
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Visionado.Video ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get usuario() As usuario

    On Error GoTo errores
    Me.Error = ""
    If Not m_ObjUsuario Is Nothing Then
        Set usuario = m_ObjUsuario
        Exit Property
    End If
    If Me.IDUsuario = "" Then
        Exit Property
    End If
    Set m_ObjUsuario = Constructor.getUsuario(Me.IDUsuario, , , , Me.Error)
    Set usuario = m_ObjUsuario
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Visionado.Usuario ha devuelto el error: " & vbNewLine & Err.Description
    End If


End Property

Public Property Get IDVisionadoCalculado() As String

    On Error GoTo errores
    Me.Error = ""
    m_sIDVisionadoCalculado = DameID("TbVideosVisionados", "IDVisionado", getdb(), Me.Error)
    If Me.Error <> "" Then
        Err.Raise 1000
    End If
    IDVisionadoCalculado = m_sIDVisionadoCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Visionado.IDVisionadoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get Porcentaje() As String

    On Error GoTo errores
    Me.Error = ""
    If IsNumeric(m_sPorcentaje) Then
        Porcentaje = m_sPorcentaje
        Exit Property
    End If
    If Not IsNumeric(Me.TiempoVideo) Then
        Exit Property
    End If
    If Not IsNumeric(Me.TiempoVisionado) Then
        m_sPorcentaje = "0"
        Porcentaje = m_sPorcentaje
        Exit Property
    End If
    m_sPorcentaje = CStr(Round(CDbl(Me.TiempoVisionado) / CDbl(Me.TiempoVideo) * 100, 2))
    
    Porcentaje = m_sPorcentaje
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Visionado.Porcentaje ha devuelto el error: " & vbNewLine & Err.Description
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
    Set tdf = getdb().TableDefs("TbVideosVisionados")
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
        Me.Error = "El metodo Visionado.ColCampos ha devuelto el error: " & vbNewLine & Err.Description
    End If


End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If m_NombrePropiedad = "IDVisionado" Then
        getPropiedad = Me.IDVisionado
        Exit Function
    End If
    If m_NombrePropiedad = "IDVideo" Then
        getPropiedad = Me.IDVideo
        Exit Function
    End If
    If m_NombrePropiedad = "TiempoVisionado" Then
        getPropiedad = Me.TiempoVisionado
        Exit Function
    End If
    If m_NombrePropiedad = "IDUsuario" Then
        getPropiedad = Me.IDUsuario
        Exit Function
    End If
    If m_NombrePropiedad = "FechaVisionado" Then
        getPropiedad = Me.FechaVisionado
        Exit Function
    End If
    If m_NombrePropiedad = "TiempoVideo" Then
        getPropiedad = Me.TiempoVideo
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Visionado.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Visionado.getPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If m_NombrePropiedad = "IDVisionado" Then
        Me.IDVisionado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDVideo" Then
        Me.IDVideo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TiempoVisionado" Then
        Me.TiempoVisionado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDUsuario" Then
        Me.IDUsuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaVisionado" Then
        Me.FechaVisionado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TiempoVideo" Then
        Me.TiempoVideo = m_ValorPropiedad
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Visionado.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Visionado.SetPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function Alta(Optional ByRef p_Error As String)
    
    Dim rcdDatos As DAO.Recordset
    On Error GoTo errores
    If Me.IDUsuario = "" Then
        p_Error = "No se conoce el ID del Usuario del visionado"
        Err.Raise 1000
    End If
    If Me.IDVideo = "" Then
        p_Error = "No se conoce el ID del Vídeo del visionado"
        Err.Raise 1000
    End If
    If Not IsNumeric(Me.TiempoVisionado) Then
        p_Error = "No se ha indicado el tiempo de visionado"
        Err.Raise 1000
    End If
    If Not IsDate(Me.FechaVisionado) Then
        Me.FechaVisionado = Now()
    End If
    If Not IsNumeric(Me.TiempoVideo) Then
        p_Error = "No se ha indicado el tiempo del vídeo"
        Err.Raise 1000
    End If
    Me.IDVisionado = Me.IDVisionadoCalculado
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "TbVideosVisionados"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("IDVisionado") = Me.IDVisionado
            .Fields("IDVideo") = Me.IDVideo
            .Fields("TiempoVisionado") = Me.TiempoVisionado
            .Fields("IDUsuario") = Me.IDUsuario
            .Fields("FechaVisionado") = Me.FechaVisionado
            .Fields("TiempoVideo") = Me.TiempoVideo
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Visionado.Alta ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function Eliminar(Optional ByRef p_Error As String)
    
    
    On Error GoTo errores
    If Me.IDVisionado = "" Then
        p_Error = "No se conoce el ID del visionado"
        Err.Raise 1000
    End If
    
    m_SQL = "DELETE * FROM TbVideosVisionados " & _
            "WHERE IDVISIONADO=" & Me.IDVisionado & ";"
    getdb().Execute m_SQL
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Visionado.Eliminar ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

