VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Video"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDVideo As String
Public Titulo As String
Public NombreArchivo As String
Public IDAplicacion As String
Public Observaciones As String
Public Descripcion As String
Public URLLocal As String
Public SubidoPor As String
Public FechaSubido As String

Public ParaCalidad As String
Public ParaAdministrador As String
Public ParaTecnicos As String

Private m_objAplicacion As Aplicacion
Private m_objColCategorias As Scripting.Dictionary

Private m_ObjCategoria As Categoria
Private m_sURL As String
Private m_eEsDeAyuda As EnumSino
Private m_eEsDeEncuesta As EnumSino
Private m_sIDVideoCalculado As String
Private m_objColCampos As Collection
Private m_objSubidoPorObj As usuario
Private m_eParaCalidadCalculado As EnumSino
Private m_eParaAdministradorCalculado As EnumSino
Private m_eParaTecnicosCalculado As EnumSino
Private m_eParaTodosCalculado As EnumSino


Public Error As String




Public Property Get Aplicacion() As Aplicacion

    On Error GoTo errores
    Me.Error = ""
    If Not m_objAplicacion Is Nothing Then
        Set Aplicacion = m_objAplicacion
        Exit Property
    End If
    If Me.IDAplicacion = "" Then
        Exit Property
    End If
    Set m_objAplicacion = Constructor.getAplicacion(Me.IDAplicacion, Me.Error)
    Set Aplicacion = m_objAplicacion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Video.Aplicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get ColCategorias() As Scripting.Dictionary

    On Error GoTo errores
    Me.Error = ""
    If Not m_objColCategorias Is Nothing Then
        Set ColCategorias = m_objColCategorias
        Exit Property
    End If
    
    If Me.IDVideo = "" Then
        Exit Property
    End If
    Set m_objColCategorias = Constructor.getCategorias(Me.IDVideo, Me.Error)
    Set ColCategorias = m_objColCategorias
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Video.ColCategorias ha devuelto el error: " & vbNewLine & Err.Description
    End If


End Property

Public Property Get Categoria() As Categoria
    Dim m_ID As Variant
    
    On Error GoTo errores
    Me.Error = ""
    If Not m_ObjCategoria Is Nothing Then
        Set Categoria = m_ObjCategoria
        Exit Property
    End If
    If Me.ColCategorias Is Nothing Then
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    If Me.ColCategorias.Count <> 1 Then
        Exit Property
    End If
    For Each m_ID In Me.ColCategorias
        Set m_ObjCategoria = Me.ColCategorias(m_ID)
    Next
    
    Set Categoria = m_ObjCategoria
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Video.Categoria ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get EsDeAyuda() As EnumSino
    Dim m_ID As Variant
    Dim m_ObjCategoria As Categoria
    On Error GoTo errores
    Me.Error = ""
    If m_eEsDeAyuda <> 0 Then
        EsDeAyuda = m_eEsDeAyuda
        Exit Property
    End If
    If Me.ColCategorias Is Nothing Then
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    If m_ObjEntorno.ColCategoriasVideo Is Nothing Then
        Me.Error = m_ObjEntorno.Error
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    For Each m_ID In Me.ColCategorias
        Set m_ObjCategoria = Me.ColCategorias(m_ID)
        If m_ObjCategoria.CategoriaCalculada = EnumVideoCategoria.Ayuda Then
            Set m_ObjCategoria = Nothing
            m_eEsDeAyuda = EnumSino.Sí
            EsDeAyuda = m_eEsDeAyuda
            Exit Property
        End If
        
    Next
    m_eEsDeAyuda = EnumSino.No
    EsDeAyuda = m_eEsDeAyuda
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Categoria.EsDeAyuda ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Property Let EsDeAyuda(p_Valor As EnumSino)
    m_eEsDeAyuda = p_Valor
End Property

Public Property Get EsDeEncuesta() As EnumSino

    Dim m_ID As Variant
    Dim m_ObjCategoria As Categoria
    On Error GoTo errores
    Me.Error = ""
    If m_eEsDeEncuesta <> 0 Then
        EsDeEncuesta = m_eEsDeEncuesta
        Exit Property
    End If
    If Me.ColCategorias Is Nothing Then
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    If m_ObjEntorno.ColCategoriasVideo Is Nothing Then
        Me.Error = m_ObjEntorno.Error
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    For Each m_ID In Me.ColCategorias
        Set m_ObjCategoria = Me.ColCategorias(m_ID)
        If m_ObjCategoria.CategoriaCalculada = EnumVideoCategoria.Encuesta Then
            Set m_ObjCategoria = Nothing
            m_eEsDeEncuesta = EnumSino.Sí
            EsDeEncuesta = m_eEsDeEncuesta
            Exit Property
        End If
        
    Next
    m_eEsDeEncuesta = EnumSino.No
    EsDeEncuesta = m_eEsDeEncuesta
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Categoria.EsDeEncuesta ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property
Property Let EsDeEncuesta(p_Valor As EnumSino)
    m_eEsDeEncuesta = p_Valor
End Property

Public Property Get url() As String
    
    
    Dim m_URLDirAppVideo As String
    On Error GoTo errores
    Me.Error = ""
    
    If m_sURL <> "" Then
        url = m_sURL
        Exit Property
    End If
    If Me.NombreArchivo = "" Then
        Exit Property
    End If
    If Me.Aplicacion Is Nothing Then
        If Me.Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    
    m_URLDirAppVideo = Me.Aplicacion.URLDirectorioVideos
    m_sURL = m_URLDirAppVideo & Me.NombreArchivo
    
    url = m_sURL
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Video.URL ha devuelto el error: " & vbNewLine & Err.Description
    End If


End Property

Public Property Get IDVideoCalculado() As String

    On Error GoTo errores
    Me.Error = ""
    m_sIDVideoCalculado = DameID("TbVideos", "IDVideo", getdb(), Me.Error)
    If Me.Error <> "" Then
        Err.Raise 1000
    End If
    IDVideoCalculado = m_sIDVideoCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Video.IDVideoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get SubidoPorObj() As usuario

    On Error GoTo errores
    
    Me.Error = ""
    If Not m_objSubidoPorObj Is Nothing Then
        Set SubidoPorObj = m_objSubidoPorObj
        Exit Property
    End If
    If Me.SubidoPor = "" Then
        Exit Property
    End If
    Set m_objSubidoPorObj = Constructor.getUsuario(, Me.SubidoPor, , , Me.Error)
    Set SubidoPorObj = m_objSubidoPorObj
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Video.SubidoPorObj ha devuelto el error: " & vbNewLine & Err.Description
    End If


End Property

Public Property Get ParaCalidadCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    
    If Me.ParaCalidad = "Sí" Then
        m_eParaCalidadCalculado = EnumSino.Sí
    ElseIf Me.ParaCalidad = "No" Then
        m_eParaCalidadCalculado = EnumSino.No
    End If
    ParaCalidadCalculado = m_eParaCalidadCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Video.ParaCalidadCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get ParaAdministradorCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    
    If Me.ParaAdministrador = "Sí" Then
        m_eParaAdministradorCalculado = EnumSino.Sí
    ElseIf Me.ParaAdministrador = "No" Then
        m_eParaAdministradorCalculado = EnumSino.No
    End If
    ParaAdministradorCalculado = m_eParaAdministradorCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Video.ParaAdministradorCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get ParaTecnicosCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    
    If Me.ParaTecnicos = "Sí" Then
        m_eParaTecnicosCalculado = EnumSino.Sí
    ElseIf Me.ParaTecnicos = "No" Then
        m_eParaTecnicosCalculado = EnumSino.No
    End If
    ParaTecnicosCalculado = m_eParaTecnicosCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Video.ParaTecnicosCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Get ParaTodosCalculado() As EnumSino

    On Error GoTo errores
    Me.Error = ""
    
    If Me.ParaTodosCalculado = 0 And Me.ParaCalidadCalculado = 0 And Me.ParaTecnicosCalculado = 0 Then
        m_eParaTodosCalculado = EnumSino.Sí
        ParaTodosCalculado = m_eParaTodosCalculado
        Exit Property
    End If
    If Me.ParaCalidadCalculado = EnumSino.Sí And Me.ParaCalidadCalculado = EnumSino.Sí And _
        Me.ParaTecnicosCalculado = EnumSino.Sí Then
        
        m_eParaTodosCalculado = EnumSino.Sí
    Else
        m_eParaTodosCalculado = EnumSino.No
    End If
    
    ParaTodosCalculado = m_eParaTodosCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Video.ParaTodosCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property


Private Function CampoRepetido( _
                                p_NombreCampo As String, _
                                p_Valor As String, _
                                Optional p_ObjVideoAlInicio As Video, _
                                Optional ByRef p_Error As String _
                                ) As EnumSino
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Where As String
    On Error GoTo errores
    If p_NombreCampo = "" Or p_Valor = "" Then
        p_Error = "Faltan datos"
        Err.Raise 1000
    End If
    
    If Not p_ObjVideoAlInicio Is Nothing Then
        m_Where = "WHERE " & p_NombreCampo & "='" & p_Valor & "' AND " & _
                    "IDAplicacion=" & Me.IDAplicacion & " AND " & _
                    "IDVideo<>" & p_ObjVideoAlInicio.IDVideo & ";"
        
    Else
        m_Where = "WHERE " & p_NombreCampo & "='" & p_Valor & "' AND " & _
                            "IDAplicacion=" & Me.IDAplicacion & ";"
    End If
    m_SQL = "SELECT " & p_NombreCampo & " " & _
            "FROM TbVideos " & _
            m_Where
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            CampoRepetido = EnumSino.No
        Else
            CampoRepetido = EnumSino.Sí
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Video.CampoRepetido ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Function Registrar( _
                    Optional ByRef p_ObjVideoAlInicio As Video, _
                    Optional ByRef p_Error As String _
                    ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_ID As String
    Dim m_CampoRepetido As EnumSino
    Dim m_URLFinal As String
    
    Dim m_IDVideoCategoria As String
    On Error GoTo errores
    
    
    If p_ObjVideoAlInicio Is Nothing Then
        If Not fso.FileExists(Me.URLLocal) Then
            p_Error = "No es alcanzable el video en local"
            Err.Raise 1000
        End If
        Me.NombreArchivo = fso.GetFileName(URLLocal)
        m_CampoRepetido = CampoRepetido("NombreArchivo", Me.NombreArchivo, , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_CampoRepetido = EnumSino.Sí Then
            p_Error = "Ya existe para esta misma aplicación un archivo con el mismo nombre"
            Err.Raise 1000
        End If
        m_URLFinal = Me.url
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If fso.FileExists(m_URLFinal) Then
            p_Error = "En el servidor hay un vídeo con el mismo nombre"
            Err.Raise 1000
        End If
    End If
    
    If Me.IDAplicacion = "" Then
        p_Error = "No sabemos para qué aplicación es el vídeo"
        Err.Raise 1000
    End If
    If Me.Titulo = "" Then
        p_Error = "Ha de indicar el Título"
        Err.Raise 1000
    End If
    m_CampoRepetido = CampoRepetido("Titulo", Me.Titulo, p_ObjVideoAlInicio, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_CampoRepetido = EnumSino.Sí Then
        p_Error = "Ya existe para esta misma aplicación un título idéntico para otro vídeo"
        Err.Raise 1000
    End If
    If Me.EsDeAyuda = 0 Then
        p_Error = "Ha de indicar si el vídeo es de ayuda o no"
        Err.Raise 1000
    End If
    If Me.EsDeEncuesta = 0 Then
        p_Error = "Ha de indicar si el vídeo es de encuesta o no"
        Err.Raise 1000
    End If
    If Me.EsDeAyuda = EnumSino.No And Me.EsDeEncuesta = EnumSino.No Then
        p_Error = "El vídeo ha de ser de ayuda o de encuesta"
        Err.Raise 1000
    End If
    If Me.Descripcion = "" Then
        p_Error = "Ha de indicar la descripción"
        Err.Raise 1000
    End If
    If Me.ParaTodosCalculado = EnumSino.Sí Then
        Me.ParaAdministrador = "Sí"
        Me.ParaCalidad = "Sí"
        Me.ParaTecnicos = "Sí"
    Else
        If Me.ParaAdministrador <> "Sí" And Me.ParaAdministrador <> "No" Then
            Me.ParaAdministrador = "Sí"
        End If
        If Me.ParaCalidad <> "Sí" And Me.ParaCalidad <> "No" Then
            Me.ParaCalidad = "No"
        End If
        If Me.ParaTecnicos <> "Sí" And Me.ParaTecnicos <> "No" Then
            Me.ParaTecnicos = "No"
        End If
    End If
    If p_ObjVideoAlInicio Is Nothing Then
    
        m_SQL = "TbVideos"
    Else
        m_SQL = "SELECT TbVideos.* " & _
            "FROM TbVideos " & _
            "WHERE IDVideo=" & p_ObjVideoAlInicio.IDVideo & ";"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If p_ObjVideoAlInicio Is Nothing Then
            .AddNew
                m_ID = Me.IDVideoCalculado
                p_Error = Me.Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                Me.IDVideo = m_ID
                
                .Fields("IDVideo") = Me.IDVideo
                .Fields("IDAplicacion") = Me.IDAplicacion
                .Fields("NombreArchivo") = Me.NombreArchivo
                .Fields("Titulo") = Me.Titulo
                .Fields("Descripcion") = Me.Descripcion
                If Nz(Me.Observaciones, "") <> "" Then
                    .Fields("Observaciones") = Me.Observaciones
                End If
                .Fields("SubidoPor") = m_ObjUsuarioConectadoLogin.UsuarioRed
                .Fields("FechaSubido") = Now()
                .Fields("ParaAdministrador") = Me.ParaAdministrador
                .Fields("ParaCalidad") = Me.ParaCalidad
                .Fields("ParaTecnicos") = Me.ParaTecnicos
            .Update
        Else
            Me.IDAplicacion = p_ObjVideoAlInicio.IDAplicacion
            .Edit
                .Fields("Titulo") = Me.Titulo
                .Fields("Descripcion") = Me.Descripcion
                If Nz(Me.Observaciones, "") <> "" Then
                    .Fields("Observaciones") = Me.Observaciones
                Else
                    .Fields("Observaciones") = Null
                End If
                .Fields("ParaAdministrador") = Me.ParaAdministrador
                .Fields("ParaCalidad") = Me.ParaCalidad
                .Fields("ParaTecnicos") = Me.ParaTecnicos
            .Update
        End If
       
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If p_ObjVideoAlInicio Is Nothing Then
        If Not fso.FolderExists(fso.GetParentFolderName(Me.url)) Then
            fso.CreateFolder fso.GetParentFolderName(Me.url)
        End If
        fso.CopyFile Me.URLLocal, Me.url, True
        m_SQL = "DELETE TbVideosCategorias.* " & _
                "FROM TbVideosCategorias " & _
                "WHERE IDVideo=" & Me.IDVideo & ";"
        getdb().Execute m_SQL
        m_SQL = "TbVideosCategorias"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        With rcdDatos
            If Me.EsDeAyuda = EnumSino.Sí Then
                m_IDVideoCategoria = DameID("TbVideosCategorias", "IDCategoriaVideo", getdb(), p_Error)
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                .AddNew
                    .Fields("IDCategoriaVideo") = m_IDVideoCategoria
                    .Fields("IDCategoria") = EnumVideoCategoria.Ayuda
                    .Fields("IDVideo") = Me.IDVideo
                .Update
            End If
            If Me.EsDeEncuesta = EnumSino.Sí Then
                m_IDVideoCategoria = DameID("TbVideosCategorias", "IDCategoriaVideo", getdb(), p_Error)
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                .AddNew
                    .Fields("IDCategoriaVideo") = m_IDVideoCategoria
                    .Fields("IDCategoria") = EnumVideoCategoria.Encuesta
                    .Fields("IDVideo") = Me.IDVideo
                .Update
            End If
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
    Else
        If p_ObjVideoAlInicio.EsDeAyuda <> Me.EsDeAyuda Or p_ObjVideoAlInicio.EsDeEncuesta <> Me.EsDeEncuesta Then
            m_SQL = "DELETE TbVideosCategorias.* " & _
                    "FROM TbVideosCategorias " & _
                    "WHERE IDVideo=" & p_ObjVideoAlInicio.IDVideo & ";"
            getdb().Execute m_SQL
            m_SQL = "TbVideosCategorias"
            Set rcdDatos = getdb().OpenRecordset(m_SQL)
            With rcdDatos
                If Me.EsDeAyuda = EnumSino.Sí Then
                    m_IDVideoCategoria = DameID("TbVideosCategorias", "IDCategoriaVideo", getdb(), p_Error)
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    .AddNew
                        .Fields("IDCategoriaVideo") = m_IDVideoCategoria
                        .Fields("IDCategoria") = EnumVideoCategoria.Ayuda
                        .Fields("IDVideo") = p_ObjVideoAlInicio.IDVideo
                    .Update
                End If
                If Me.EsDeEncuesta = EnumSino.Sí Then
                    m_IDVideoCategoria = DameID("TbVideosCategorias", "IDCategoriaVideo", getdb(), p_Error)
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    .AddNew
                        .Fields("IDCategoriaVideo") = m_IDVideoCategoria
                        .Fields("IDCategoria") = EnumVideoCategoria.Encuesta
                        .Fields("IDVideo") = p_ObjVideoAlInicio.IDVideo
                    .Update
                End If
            End With
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Registrar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function


Function Eliminar( _
                    Optional ByRef p_Error As String _
                    ) As String
    
    
    
    Dim m_URL As String
    On Error GoTo errores
   
    If Me.IDVideo = "" Then
        p_Error = "Ha de indicar el IDVideo"
        Err.Raise 1000
    End If
    m_URL = Me.url
    If fso.FileExists(m_URL) Then
        If FicheroAbierto(m_URL) Then
            p_Error = "El vídeo parece estar siendo usado o estar de sólo lectura"
            Err.Raise 1000
        End If
    End If
    'categorías
    m_SQL = "DELETE TbVideosCategorias.* " & _
                "FROM TbVideosCategorias " & _
                "WHERE IDVideo=" & Me.IDVideo & ";"
    getdb().Execute m_SQL
    m_SQL = "DELETE TbVideos.* " & _
            "FROM TbVideos " & _
            "WHERE IDVideo=" & Me.IDVideo & ";"
    getdb().Execute m_SQL
    If fso.FileExists(m_URL) Then
        fso.DeleteFile m_URL, True
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Eliminar ha producido el error nº: " & Err.Number & _
        vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Public Property Get ColCampos() As Collection
    
    Dim tdf As DAO.TableDef
    Dim fld As Object
    On Error GoTo errores
    
    Me.Error = ""
    If Not m_objColCampos Is Nothing Then
        Set ColCampos = m_objColCampos
        Exit Property
    End If
    Set tdf = getdb().TableDefs("TbVideos")
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
        Me.Error = "El metodo Video.ColCampos ha devuelto el error: " & vbNewLine & Err.Description
    End If


End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If m_NombrePropiedad = "IDVideo" Then
        getPropiedad = Me.IDVideo
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo" Then
        getPropiedad = Me.Titulo
        Exit Function
    End If
    If m_NombrePropiedad = "NombreArchivo" Then
        getPropiedad = Me.NombreArchivo
        Exit Function
    End If
    If m_NombrePropiedad = "IDAplicacion" Then
        getPropiedad = Me.IDAplicacion
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        getPropiedad = Me.Observaciones
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    If m_NombrePropiedad = "SubidoPor" Then
        getPropiedad = Me.SubidoPor
        Exit Function
    End If
    If m_NombrePropiedad = "FechaSubido" Then
        getPropiedad = Me.FechaSubido
        Exit Function
    End If
    If m_NombrePropiedad = "ParaCalidad" Then
        getPropiedad = Me.ParaCalidad
        Exit Function
    End If
    If m_NombrePropiedad = "ParaAdministrador" Then
        getPropiedad = Me.ParaAdministrador
        Exit Function
    End If
    If m_NombrePropiedad = "ParaTecnicos" Then
        getPropiedad = Me.ParaTecnicos
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Video.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Video.getPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If m_NombrePropiedad = "IDVideo" Then
        Me.IDVideo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo" Then
        Me.Titulo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreArchivo" Then
        Me.NombreArchivo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDAplicacion" Then
        Me.IDAplicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        Me.Observaciones = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "SubidoPor" Then
        Me.SubidoPor = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaSubido" Then
        Me.FechaSubido = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ParaCalidad" Then
        Me.ParaCalidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ParaAdministrador" Then
        Me.ParaAdministrador = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ParaTecnicos" Then
        Me.ParaTecnicos = m_ValorPropiedad
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Video.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Video.SetPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function




