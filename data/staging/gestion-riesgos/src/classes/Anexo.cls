VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Anexo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDAnexo As String
Public IDProyecto As String
Public IDEdicion As String
Public IDRiesgo As String
Public FechaAnexo As String
Public Titulo As String
Public NombreArchivo As String
Public EvidenciaUTE As String
Public EvidenciaSuministrador As String

Private m_sIDAnexoCalculado As String
Private m_sURLAnexo As String
Private m_objEdicion As Edicion
Private m_ObjProyecto As Proyecto
Private m_ObjRiesgo As riesgo
Private m_eUsuarioConectadoAutorizado As EnumSiNo
Private m_eEsActivo As EnumSiNo
Private m_sTipo As String

Private m_objColCampos As Collection
Private m_URLFinal As String
Public Error As String
Private m_Error As String

Public Property Get IDAnexoCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDAnexoCalculado = DameID("TbAnexos", "IDAnexo", , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    IDAnexoCalculado = m_sIDAnexoCalculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.IDAnexoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get URLAnexo() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.NombreArchivo = "" Then
        Exit Property
    End If
    If m_ObjEntorno.URLDirectorioDocumentacion = "" Then
        m_Error = m_ObjEntorno.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    m_sURLAnexo = m_ObjEntorno.URLDirectorioDocumentacion & Me.NombreArchivo
    URLAnexo = m_sURLAnexo
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.URLAnexo ha devuelto el error: " & vbNewLine & Err.Description
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
        m_Error = "El método Anexo.Edicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set Edicion(ByVal sNewValue As Variant)

    Set m_objEdicion = sNewValue

End Property

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
        m_Error = "El método Anexo.Proyecto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set Proyecto(ByVal sNewValue As Variant)

    Set Proyecto = sNewValue

End Property

Public Property Get riesgo() As riesgo

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjRiesgo Is Nothing Then
        Set riesgo = m_ObjRiesgo
        Exit Property
    End If
    If Me.IDRiesgo = "" Then
        Exit Property
    End If
    Set m_ObjRiesgo = Constructor.getRiesgo(Me.IDRiesgo, , , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set riesgo = m_ObjRiesgo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.Riesgo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get UsuarioConectadoAutorizado() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eUsuarioConectadoAutorizado <> 0 Then
        UsuarioConectadoAutorizado = m_eUsuarioConectadoAutorizado
        Exit Property
    End If
    m_eUsuarioConectadoAutorizado = UsuarioAutorizado(Me, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    UsuarioConectadoAutorizado = m_eUsuarioConectadoAutorizado
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.UsuarioConectadoAutorizado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EsActivo() As EnumSiNo
    
    Dim m_ObjEdicionPertenece As Edicion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eEsActivo <> 0 Then
        EsActivo = m_eEsActivo
        Exit Property
    End If
    If Not Me.Edicion Is Nothing Then
        Set m_ObjEdicionPertenece = Me.Edicion
        m_Error = Me.Error
        
    ElseIf Not Me.riesgo Is Nothing Then
        Set m_ObjEdicionPertenece = Me.riesgo.Edicion
        m_Error = Me.riesgo.Error
    End If
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_eEsActivo = m_ObjEdicionPertenece.EsActivo
    EsActivo = m_eEsActivo
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.EsActivo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Tipo() As String

    ' #83: Tipo es propiedad computada (no leida de SQL porque el campo no existe
    ' en TbAnexos). Mismo patron que modAnexosListPresenter.DeriveTipoFromAnexo:
    ' "R" si IDRiesgo poblado (anexo de riesgo), "E" en caso contrario (anexo directo
    ' de edicion). Test fixture B2: directos con IDEdicion+IDRiesgo vacio = "E",
    ' riesgos con IDRiesgo poblado = "R".
    If LenB(Nz(Me.IDRiesgo, "")) > 0 Then
        Tipo = "R"
    Else
        Tipo = "E"
    End If
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.Tipo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property



Public Function MotivoNoOK( _
                            p_Obj As Object, _
                            p_URLLocal As String, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_NombreProximoAnexoSinExtension As String
    Dim m_Extension As String
    
    
    On Error GoTo errores
    
    If p_Obj Is Nothing Then
        p_Error = "No se conoce el objeto"
        Exit Function
    End If
    If TypeOf p_Obj Is Proyecto Then
        Me.IDProyecto = p_Obj.IDProyecto
    ElseIf TypeOf p_Obj Is Edicion Then
        Me.IDEdicion = p_Obj.IDEdicion
    ElseIf TypeOf p_Obj Is riesgo Then
        Me.IDRiesgo = p_Obj.IDRiesgo
    Else
        p_Error = "Este tipo de objeto no permite anexar"
        Err.Raise 1000
    End If
    
    If p_Obj.UsuarioConectadoAutorizado = EnumSiNo.No Then
        p_Error = p_Obj.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        MotivoNoOK = "El perfil de su usuario no tiene autorización para realizar esta operación"
        Exit Function
    End If
    If Not fso.FileExists(p_URLLocal) Then
        MotivoNoOK = "El Archivo a Anexar no es alcanzable en estos momentos"
        Exit Function
    End If
    m_Extension = fso.GetExtensionName(p_URLLocal)
    
    If Me.Titulo = "" Then
        MotivoNoOK = "Falta la descripción del anexo"
        Exit Function
    End If
    If Me.EvidenciaUTE <> "Sí" And Me.EvidenciaUTE <> "No" Then
        Me.EvidenciaUTE = "No"
    End If
    If Me.EvidenciaSuministrador <> "Sí" And Me.EvidenciaSuministrador <> "No" Then
        Me.EvidenciaSuministrador = "No"
    End If
    m_NombreProximoAnexoSinExtension = NombreProximoAnexoSinExtension(p_Obj, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Extension <> "" Then
        m_URLFinal = m_ObjEntorno.URLDirectorioDocumentacion & m_NombreProximoAnexoSinExtension & "." & m_Extension
    Else
        m_URLFinal = m_ObjEntorno.URLDirectorioDocumentacion & m_NombreProximoAnexoSinExtension
    End If
    If Not fso.FileExists(m_URLFinal) Then
        If FicheroAbierto(m_URLFinal, p_Error) = True Then
            MotivoNoOK = "Existe en el repositorio un anexo con el mismo nombre y está abierto o es de sólo lectura"
            Exit Function
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MotivoNoOK ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Registrar( _
                            p_Obj As Object, _
                            p_URLLocal As String, _
                            Optional ByRef p_Error As String _
                            ) As String
                            
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    
    m_MotivoNoOK = MotivoNoOK(p_Obj, p_URLLocal, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.IDAnexo = Me.IDAnexoCalculado
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "TbAnexos"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            If Me.IDAnexo <> "" Then
                .Fields("IDAnexo") = Me.IDAnexo
            End If
            If Me.IDProyecto <> "" Then
                .Fields("IDProyecto") = Me.IDProyecto
            End If
            If Me.IDEdicion <> "" Then
                .Fields("IDEdicion") = Me.IDEdicion
            End If
            If Me.IDRiesgo <> "" Then
                .Fields("IDRiesgo") = Me.IDRiesgo
            End If
            Me.NombreArchivo = fso.GetFileName(m_URLFinal)
            .Fields("NombreArchivo") = Me.NombreArchivo
            If Me.Titulo <> "" Then
                .Fields("Titulo") = Me.Titulo
            End If
            Me.FechaAnexo = Now()
            If Me.EvidenciaUTE <> "" Then
                .Fields("EvidenciaUTE") = Me.EvidenciaUTE
            End If
            If Me.EvidenciaSuministrador <> "" Then
                .Fields("EvidenciaSuministrador") = Me.EvidenciaSuministrador
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    fso.CopyFile p_URLLocal, m_URLFinal, True
    Set p_Obj.ColAnexos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Registrar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function CambiarNombre( _
                            p_NuevoNombre As String, _
                            Optional ByRef p_Error As String _
                            ) As String
                            
    
    Dim m_SQL As String
    
    
    
    On Error GoTo errores
    
    Me.Error = ""
    
    If p_NuevoNombre = "" Then
        p_Error = "Falta el Nuevo nombre"
        Err.Raise 1000
    End If
    m_SQL = "UPDATE TbAnexos SET TbAnexos.Titulo = '" & p_NuevoNombre & "' " & _
            "WHERE (((TbAnexos.IDAnexo)=" & Me.IDAnexo & "));"
    getdb().Execute m_SQL
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CambiarNombre ha devuelto el error: " & Err.Description
    End If
End Function
Private Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    On Error GoTo errores
    
    If fso.FileExists(Me.URLAnexo) Then
        If FicheroAbierto(Me.URLAnexo) Then
            p_Error = "El anexo está abierto"
            Err.Raise 1000
        End If
    End If
    
    If Me.UsuarioConectadoAutorizado = EnumSiNo.No Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "El perfil de su usuario no tiene autorización para realizar esta operación"
        Err.Raise 1000
    End If
    If Me.EsActivo = EnumSiNo.No Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "El anexo no pertence a una edición activa"
        Err.Raise 1000
    End If
       
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ValidacionEliminar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function EliminarAnexo( _
                                Optional ByRef p_Error As String _
                                ) As String

    Dim m_ObjProyecto As Proyecto
    Dim m_SQL As String
    On Error GoTo errores

    Me.Error = ""

    ValidacionEliminar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "DELETE * FROM TbAnexos " & _
            "WHERE IDAnexo=" & Me.IDAnexo & ";"
    getdb().Execute m_SQL
    If fso.FileExists(Me.URLAnexo) Then
        fso.DeleteFile Me.URLAnexo, True
    End If

    ' #83: invalidar caches afectados. Si el anexo es de un riesgo, tambien
    ' aparece en la coleccion de la edicion padre, por lo que invalidamos
    ' edicion ademas del riesgo. Si es directo de edicion, solo edicion.
    If LenB(Nz(Me.IDRiesgo, "")) > 0 Then
        modFormRiesgoDocumentosHelper.InvalidarCachePorIDRiesgo CStr(Me.IDRiesgo)
    End If
    If LenB(Nz(Me.IDEdicion, "")) > 0 Then
        modFormRiesgoDocumentosHelper.InvalidarCachePorIDEdicion CStr(Me.IDEdicion)
    End If

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EliminarAnexo ha devuelto el error: " & Err.Description
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
        .Add "IDAnexo"
        .Add "IDProyecto"
        .Add "IDEdicion"
        .Add "IDRiesgo"
        .Add "FechaAnexo"
        .Add "Titulo"
        .Add "NombreArchivo"
        .Add "EvidenciaUTE"
        .Add "EvidenciaSuministrador"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
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
    If m_NombrePropiedad = "IDRiesgo" Then
        getPropiedad = Me.IDRiesgo
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAnexo" Then
        getPropiedad = Me.FechaAnexo
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
    If m_NombrePropiedad = "EvidenciaUTE" Then
        getPropiedad = Me.EvidenciaUTE
        Exit Function
    End If
    If m_NombrePropiedad = "EvidenciaSuministrador" Then
        getPropiedad = Me.EvidenciaSuministrador
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Anexo.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Anexo.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
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
    If m_NombrePropiedad = "IDRiesgo" Then
        Me.IDRiesgo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAnexo" Then
        Me.FechaAnexo = m_ValorPropiedad
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
    If m_NombrePropiedad = "EvidenciaUTE" Then
        Me.EvidenciaUTE = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EvidenciaSuministrador" Then
        Me.EvidenciaSuministrador = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Anexo.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Anexo.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


