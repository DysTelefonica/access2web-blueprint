VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComprobacionAccesoVideos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



'Public UsuarioRed As String
'Public UsuarioLogin As String
'Public FechaComprobacion As String
'Public AccesoCompleto As String
'Public CarpetasSinAccceso As String
'Public TiempoComprobacion As String
'Public AccesoADirectorioPrincipal As String
'
'Private m_objUsuarioRedObj As Usuario
'Private m_objUsuarioLoginObj As Usuario
'Private m_objColCampos As Collection
'Private m_eAccesoCompletoCalculado As EnumSino
'Private m_eAccesoADirectorioPrincipalCalculado As EnumSino
'Private t1 As Single
'Private t2 As Single
'Public Error As String
'
'
'Public Property Get AccesoCompletoCalculado() As EnumSino
'
'    On Error GoTo errores
'
'    Me.Error = ""
'    If Me.AccesoCompleto = "Sí" Then
'        m_eAccesoCompletoCalculado = EnumSino.Sí
'    ElseIf Me.AccesoCompleto = "No" Then
'        m_eAccesoCompletoCalculado = EnumSino.No
'    End If
'
'    AccesoCompletoCalculado = m_eAccesoCompletoCalculado
'    Exit Property
'errores:
'    If Err.Number <> 1000 Then
'        Me.Error = "El método ComprobacionAccesoVideos.AccesoCompletoCalculado ha devuelto el error: " & vbNewLine & Err.description
'    End If
'
'End Property
'
'Public Property Get AccesoADirectorioPrincipalCalculado() As EnumSino
'
'    On Error GoTo errores
'
'    Me.Error = ""
'    If Me.AccesoADirectorioPrincipal = "Sí" Then
'        m_eAccesoADirectorioPrincipalCalculado = EnumSino.Sí
'    ElseIf Me.AccesoADirectorioPrincipal = "No" Then
'        m_eAccesoADirectorioPrincipalCalculado = EnumSino.No
'    End If
'
'    AccesoADirectorioPrincipalCalculado = m_eAccesoADirectorioPrincipalCalculado
'    Exit Property
'errores:
'    If Err.Number <> 1000 Then
'        Me.Error = "El método ComprobacionAccesoVideos.AccesoADirectorioPrincipalCalculado ha devuelto el error: " & vbNewLine & Err.description
'    End If
'
'End Property
'
'
'Public Property Get UsuarioRedObj() As Usuario
'
'    On Error GoTo errores
'
'    Me.Error = ""
'    If Not m_objUsuarioRedObj Is Nothing Then
'        Set UsuarioRedObj = m_objUsuarioRedObj
'        Exit Property
'    End If
'    If Me.UsuarioRed = "" Then
'        Exit Property
'    End If
'    Set m_objUsuarioRedObj = Constructor.getUsuario(, , , Me.UsuarioRed, Me.Error)
'    Set UsuarioRedObj = m_objUsuarioRedObj
'    Exit Property
'errores:
'    If Err.Number <> 1000 Then
'        Me.Error = "El método ComprobacionAccesoVideos.UsuarioRedObj ha devuelto el error: " & vbNewLine & Err.description
'    End If
'
'End Property
'
'Public Property Get UsuarioLoginObj() As Usuario
'
'    On Error GoTo errores
'
'    Me.Error = ""
'    If Not m_objUsuarioLoginObj Is Nothing Then
'        Set UsuarioLoginObj = m_objUsuarioLoginObj
'        Exit Property
'    End If
'    If Me.UsuarioLogin = "" Then
'        Exit Property
'    End If
'    Set m_objUsuarioLoginObj = Constructor.getUsuario(, , , Me.UsuarioLogin, Me.Error)
'    Set UsuarioLoginObj = m_objUsuarioLoginObj
'    Exit Property
'errores:
'    If Err.Number <> 1000 Then
'        Me.Error = "El método ComprobacionAccesoVideos.UsuarioLoginObj ha devuelto el error: " & vbNewLine & Err.description
'    End If
'
'End Property
'
'Public Function Registrar(Optional ByRef p_Error As String) As String
'
'    Dim rcdDatos As DAO.Recordset
'    Dim m_Mensaje As String
'    Dim m_NombreUsuario As String
'    Dim m_ColCarpetasSinAcceso As Scripting.Dictionary
'    Dim m_ID As Variant
'
'    On Error GoTo errores
'
'    t1 = Timer
'    If Me.UsuarioRed = "" Then
'        Me.UsuarioRed = m_UsuarioRedConectado
'        If Me.UsuarioRed = "" Then
'            p_Error = "No se conoce el usuario de red conectado"
'            Err.Raise 1000
'        End If
'    End If
'    If Me.UsuarioLogin = "" Then
'        If Not m_ObjUsuarioConectadoLogin Is Nothing Then
'            Me.UsuarioLogin = m_ObjUsuarioConectadoLogin.UsuarioRed
'        End If
'    End If
'    If Me.UsuarioLogin <> "" Then
'        m_NombreUsuario = Me.UsuarioLogin
'    Else
'        m_NombreUsuario = Me.UsuarioRed
'    End If
'    If FSO.FolderExists(m_ObjEntorno.URLDirVideos) Then
'        Me.AccesoADirectorioPrincipal = "Sí"
'    Else
'        Me.AccesoADirectorioPrincipal = "No"
'    End If
'    If Me.AccesoADirectorioPrincipal = "No" Then
'        t2 = Timer
'        Me.TiempoComprobacion = Round((t2 - t1), 0)
'        GoTo grabar
'    End If
'    If Not IsDate(Me.FechaComprobacion) Then
'        Me.FechaComprobacion = Now()
'    End If
'    Set m_ColCarpetasSinAcceso = getColCarpetasSinAcceso(p_Error)
'    If p_Error <> "" Then
'        Err.Raise 1000
'    End If
'    If Not m_ColCarpetasSinAcceso Is Nothing Then
'        For Each m_ID In m_ColCarpetasSinAcceso
'            If Me.CarpetasSinAccceso = "" Then
'                Me.CarpetasSinAccceso = m_ID
'            Else
'                Me.CarpetasSinAccceso = Me.CarpetasSinAccceso & ";" & m_ID
'            End If
'        Next
'    End If
'    If Me.CarpetasSinAccceso = "" And Me.AccesoADirectorioPrincipal = "Sí" Then
'        Me.AccesoCompleto = "Sí"
'    Else
'        Me.AccesoCompleto = "No"
'    End If
'    t2 = Timer
'    Me.TiempoComprobacion = Round((t2 - t1), 0)
'grabar:
'    m_SQL = "SELECT * FROM TbComprobacionCarpetasVideos " & _
'            "WHERE UsuarioRed='" & Me.UsuarioRed & "';"
'    Set rcdDatos = getdb().OpenRecordset(m_SQL)
'    With rcdDatos
'        If .EOF Then
'            .AddNew
'                .Fields("UsuarioRed") = Me.UsuarioRed
'                If Me.UsuarioLogin <> "" Then
'                    .Fields("UsuarioLogin") = Me.UsuarioLogin
'                End If
'                .Fields("FechaComprobacion") = Me.FechaComprobacion
'                .Fields("AccesoCompleto") = Me.AccesoCompleto
'                If Me.CarpetasSinAccceso <> "" Then
'                    .Fields("CarpetasSinAccceso") = Me.CarpetasSinAccceso
'                End If
'                .Fields("TiempoComprobacion") = Me.TiempoComprobacion
'                .Fields("AccesoADirectorioPrincipal") = Me.AccesoADirectorioPrincipal
'            .Update
'        Else
'            .Edit
'
'                .Fields("FechaComprobacion") = Me.FechaComprobacion
'                .Fields("AccesoCompleto") = Me.AccesoCompleto
'                If Me.CarpetasSinAccceso <> "" Then
'                    .Fields("CarpetasSinAccceso") = Me.CarpetasSinAccceso
'                Else
'                    .Fields("CarpetasSinAccceso") = Null
'                End If
'                .Fields("TiempoComprobacion") = Me.TiempoComprobacion
'                .Fields("AccesoADirectorioPrincipal") = Me.AccesoADirectorioPrincipal
'            .Update
'        End If
'
'    End With
'    rcdDatos.Close
'    Set rcdDatos = Nothing
'    If Me.AccesoADirectorioPrincipal = "No" Then
'        m_Mensaje = "El usuario " & m_NombreUsuario & " No se puede acceder a la carpeta de videos principal"
'        CorreoNoAccesoAVideos m_Mensaje, m_NombreUsuario
'    End If
'    If Me.CarpetasSinAccceso <> "" Then
'        m_Mensaje = "El usuario " & m_NombreUsuario & " No se puede acceder a las subcarpetas de vídeos:"
'        dato = Split(Me.CarpetasSinAccceso, ";")
'        For Each m_ID In dato
'            m_Mensaje = m_Mensaje & "<br/>" & m_ID
'        Next
'        m_Mensaje = m_Mensaje & "<br/><br/>" & "Tiempo invertido: " & t2 - t1
'        CorreoNoAccesoAVideos m_Mensaje, m_NombreUsuario
'    End If
'    Exit Function
'errores:
'    If Err.Number <> 1000 Then
'        p_Error = "El metodo ComprobacionAccesoVideos.Registrar ha devuelto el error: " & vbNewLine & Err.description
'    End If
'End Function
'
'
'
'Private Function getColCarpetasSinAcceso(Optional ByRef p_Error As String) As Scripting.Dictionary
'
'
'
'    Dim m_ID As Variant
'    Dim Carpeta As Folder
'    Dim m_NombreTxt As String
'    Dim m_URLTxt As String
'    Dim m_URLCarpetaVideo As String
'
'
'    On Error GoTo errores
'    For Each Carpeta In FSO.GetFolder(m_ObjEntorno.URLDirVideos).SubFolders
'        m_NombreTxt = FSO.GetTempName
'        m_NombreTxt = Replace(m_NombreTxt, ".tmp", ".txt")
'        m_URLTxt = Carpeta.Path & "\" & m_NombreTxt
'        On Error Resume Next
'        FSO.CreateTextFile m_URLTxt
'        If Err.Number <> 0 Then
'            If getColCarpetasSinAcceso Is Nothing Then
'                Set getColCarpetasSinAcceso = New Scripting.Dictionary
'                getColCarpetasSinAcceso.CompareMode = TextCompare
'            End If
'            If Not getColCarpetasSinAcceso.Exists(Carpeta.Name) Then
'                getColCarpetasSinAcceso.Add Carpeta.Name, Carpeta.Name
'            End If
'            Err.Clear
'        Else
'            FSO.DeleteFile m_URLTxt, True
'        End If
'        On Error GoTo errores
'    Next
'
'
'    Exit Function
'errores:
'    If Err.Number <> 1000 Then
'        p_Error = "El método AvisoNoAccesoVideos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.description
'    End If
'
'End Function
'
'
'Public Property Get ColCampos() As Collection
'
'    Dim tdf As DAO.TableDef
'    Dim fld As Object
'    On Error GoTo errores
'
'    Me.Error = ""
'    If Not m_objColCampos Is Nothing Then
'        Set ColCampos = m_objColCampos
'        Exit Property
'    End If
'    Set tdf = getdb().TableDefs("TbComprobacionCarpetasVideos")
'    For Each fld In tdf.Fields
'        If m_objColCampos Is Nothing Then
'            Set m_objColCampos = New Collection
'
'        End If
'        m_objColCampos.Add fld.Name
'    Next
'    Set tdf = Nothing
'
'    Set ColCampos = m_objColCampos
'    Exit Property
'errores:
'    If Err.Number <> 1000 Then
'        Me.Error = "El metodo ComprobacionAccesoVideos.ColCampos ha devuelto el error: " & vbNewLine & Err.description
'    End If
'
'
'End Property
'
'Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant
'
'    On Error GoTo errores
'
'    Me.Error = ""
'    p_Error = ""
'
'
'    If m_NombrePropiedad = "UsuarioRed" Then
'        getPropiedad = Me.UsuarioRed
'        Exit Function
'    End If
'    If m_NombrePropiedad = "UsuarioLogin" Then
'        getPropiedad = Me.UsuarioLogin
'        Exit Function
'    End If
'    If m_NombrePropiedad = "FechaComprobacion" Then
'        getPropiedad = Me.FechaComprobacion
'        Exit Function
'    End If
'    If m_NombrePropiedad = "AccesoCompleto" Then
'        getPropiedad = Me.AccesoCompleto
'        Exit Function
'    End If
'    If m_NombrePropiedad = "CarpetasSinAccceso" Then
'        getPropiedad = Me.CarpetasSinAccceso
'        Exit Function
'    End If
'    If m_NombrePropiedad = "TiempoComprobacion" Then
'        getPropiedad = Me.TiempoComprobacion
'        Exit Function
'    End If
'    If m_NombrePropiedad = "AccesoADirectorioPrincipal" Then
'        getPropiedad = Me.AccesoADirectorioPrincipal
'        Exit Function
'    End If
'
'    p_Error = "Propiedad no reconocida en ComprobacionAccesoVideos.getPropiedad"
'    Err.Raise 1000
'    Exit Function
'errores:
'    If Err.Number <> 1000 Then
'        p_Error = "El metodo ComprobacionAccesoVideos.getPropiedad ha devuelto el error: " & vbNewLine & Err.description
'    End If
'End Function
'
'
'
'Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
'
'    On Error GoTo errores
'
'    Me.Error = ""
'    p_Error = ""
'
'    If m_NombrePropiedad = "UsuarioRed" Then
'        Me.UsuarioRed = m_ValorPropiedad
'        Exit Function
'    End If
'    If m_NombrePropiedad = "UsuarioLogin" Then
'        Me.UsuarioLogin = m_ValorPropiedad
'        Exit Function
'    End If
'    If m_NombrePropiedad = "FechaComprobacion" Then
'        Me.FechaComprobacion = m_ValorPropiedad
'        Exit Function
'    End If
'    If m_NombrePropiedad = "AccesoCompleto" Then
'        Me.AccesoCompleto = m_ValorPropiedad
'        Exit Function
'    End If
'    If m_NombrePropiedad = "CarpetasSinAccceso" Then
'        Me.CarpetasSinAccceso = m_ValorPropiedad
'        Exit Function
'    End If
'    If m_NombrePropiedad = "TiempoComprobacion" Then
'        Me.TiempoComprobacion = m_ValorPropiedad
'        Exit Function
'    End If
'    If m_NombrePropiedad = "AccesoADirectorioPrincipal" Then
'        Me.AccesoADirectorioPrincipal = m_ValorPropiedad
'        Exit Function
'    End If
'
'
'    p_Error = "Propiedad no reconocida en ComprobacionAccesoVideos.SetPropiedad"
'    Err.Raise 1000
'    Exit Function
'errores:
'    If Err.Number <> 1000 Then
'        p_Error = "El metodo ComprobacionAccesoVideos.SetPropiedad ha devuelto el error: " & vbNewLine & Err.description
'    End If
'End Function
'
'
'
'
'
'
