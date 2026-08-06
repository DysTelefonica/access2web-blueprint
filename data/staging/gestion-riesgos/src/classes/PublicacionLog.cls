VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PublicacionLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public ID As String
Public IDEdicion As String
Public FechaPreparadaParaPublicar As String
Public UsuarioFechaPreparadaParaPublicar As String
Public FechaPreparadaParaPublicarQuitar As String
Public UsuarioFechaPreparadaParaPublicarQuitar As String
Public FechaPublicacion As String
Public UsuarioFechaPublicacion As String
Public PropuestaRechazadaPorCalidadFecha As String
Public PropuestaRechazadaPorCalidadMotivo As String
Public FechaRegistro As String

Private m_sIDCalculado As String
Private m_objEdicion As Edicion
Private m_ObjProyecto As Proyecto

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get IDCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDCalculado = DameID("TbLogPublicaciones", "ID", , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    IDCalculado = m_sIDCalculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PublicacionLog.IDCalculado ha devuelto el error: " & vbNewLine & Err.Description
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
        Exit Property
    End If
    Set m_objEdicion = Constructor.getEdicion(Me.IDEdicion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Edicion = m_objEdicion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PublicacionLog.Edicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set Edicion(ByVal sNewValue As Variant)

    Set m_objEdicion = sNewValue

End Property


Public Function Registrar( _
                            Optional ByRef p_Error As String _
                            ) As String
                            
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    
    
    On Error GoTo errores
    
    Me.Error = ""
    
    
    With Me
        .ID = .IDCalculado
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If .IDEdicion = "" Then
            p_Error = "No se sabe la Edición"
            Err.Raise 1000
        End If
    End With
    
    
    'if not isdate(FechaPreparadaParaPublicar) and not isdate(me.
    m_SQL = "TbLogPublicaciones"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("ID") = Me.ID
            .Fields("IDEdicion") = Me.IDEdicion
            If IsDate(Me.FechaPreparadaParaPublicar) Then
                .Fields("FechaPreparadaParaPublicar") = Me.FechaPreparadaParaPublicar
            End If
            If Me.UsuarioFechaPreparadaParaPublicar <> "" Then
                .Fields("UsuarioFechaPreparadaParaPublicar") = Me.UsuarioFechaPreparadaParaPublicar
            End If
            If IsDate(Me.FechaPreparadaParaPublicarQuitar) Then
                .Fields("FechaPreparadaParaPublicarQuitar") = Me.FechaPreparadaParaPublicarQuitar
            End If
            If Me.UsuarioFechaPreparadaParaPublicarQuitar <> "" Then
                .Fields("UsuarioFechaPreparadaParaPublicarQuitar") = Me.UsuarioFechaPreparadaParaPublicarQuitar
            End If
            If IsDate(Me.FechaPublicacion) Then
                .Fields("FechaPublicacion") = Me.FechaPublicacion
            End If
            If Me.UsuarioFechaPublicacion <> "" Then
                .Fields("UsuarioFechaPublicacion") = Me.UsuarioFechaPublicacion
            End If
            If IsDate(Me.PropuestaRechazadaPorCalidadFecha) Then
                .Fields("PropuestaRechazadaPorCalidadFecha") = Me.PropuestaRechazadaPorCalidadFecha
            End If
            If Me.PropuestaRechazadaPorCalidadMotivo <> "" Then
                .Fields("PropuestaRechazadaPorCalidadMotivo") = Me.PropuestaRechazadaPorCalidadMotivo
            End If
            .Fields("FechaRegistro") = Now()
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Registrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Eliminar( _
                            Optional ByRef p_Error As String _
                            ) As String
                            
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    
    m_SQL = "DELETE * FROM TbLogPublicaciones " & _
            "WHERE ID=" & Me.ID & ";"
    getdb().Execute m_SQL
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Eliminar ha devuelto el error: " & Err.Description
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
        .Add "ID"
        .Add "IDEdicion"
        .Add "FechaPreparadaParaPublicar"
        .Add "UsuarioFechaPreparadaParaPublicar"
        .Add "FechaPreparadaParaPublicarQuitar"
        .Add "UsuarioFechaPreparadaParaPublicarQuitar"
        .Add "FechaPublicacion"
        .Add "UsuarioFechaPublicacion"
        .Add "PropuestaRechazadaPorCalidadFecha"
        .Add "PropuestaRechazadaPorCalidadMotivo"
        .Add "FechaRegistro"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método PublicacionLog.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_NombrePropiedad = "ID" Then
        getPropiedad = Me.ID
        Exit Function
    End If
    If m_NombrePropiedad = "IDEdicion" Then
        getPropiedad = Me.IDEdicion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPreparadaParaPublicar" Then
        getPropiedad = Me.FechaPreparadaParaPublicar
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioFechaPreparadaParaPublicar" Then
        getPropiedad = Me.UsuarioFechaPreparadaParaPublicar
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPreparadaParaPublicarQuitar" Then
        getPropiedad = Me.FechaPreparadaParaPublicarQuitar
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioFechaPreparadaParaPublicarQuitar" Then
        getPropiedad = Me.UsuarioFechaPreparadaParaPublicarQuitar
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPublicacion" Then
        getPropiedad = Me.FechaPublicacion
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioFechaPublicacion" Then
        getPropiedad = Me.UsuarioFechaPublicacion
        Exit Function
    End If
    If m_NombrePropiedad = "PropuestaRechazadaPorCalidadFecha" Then
        getPropiedad = Me.PropuestaRechazadaPorCalidadFecha
        Exit Function
    End If
    If m_NombrePropiedad = "PropuestaRechazadaPorCalidadMotivo" Then
        getPropiedad = Me.PropuestaRechazadaPorCalidadMotivo
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRegistro" Then
        getPropiedad = Me.FechaRegistro
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en PublicacionLog.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PublicacionLog.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "ID" Then
        Me.ID = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEdicion" Then
        Me.IDEdicion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPreparadaParaPublicar" Then
        Me.FechaPreparadaParaPublicar = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioFechaPreparadaParaPublicar" Then
        Me.UsuarioFechaPreparadaParaPublicar = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPreparadaParaPublicarQuitar" Then
        Me.FechaPreparadaParaPublicarQuitar = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioFechaPreparadaParaPublicarQuitar" Then
        Me.UsuarioFechaPreparadaParaPublicarQuitar = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPublicacion" Then
        Me.FechaPublicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioFechaPublicacion" Then
        Me.UsuarioFechaPublicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PropuestaRechazadaPorCalidadFecha" Then
        Me.PropuestaRechazadaPorCalidadFecha = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PropuestaRechazadaPorCalidadMotivo" Then
        Me.PropuestaRechazadaPorCalidadMotivo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRegistro" Then
        Me.FechaRegistro = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en PublicacionLog.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PublicacionLog.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


