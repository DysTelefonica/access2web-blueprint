VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SuministradorParaEvidencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public ID As String
Public IDEdicion As String
Public IDSuministrador As String
Public IDAnexo As String

Private m_sIDCalculado As String
Private m_objEdicion As Edicion
Private m_ObjSuministrador As Suministrador
Private m_ObjAnexo As Anexo
Private m_sURLAnexo As String
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get IDCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDCalculado = DameID("TbProyectosEdicionesSuministradores", "ID", getdb(), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDCalculado = m_sIDCalculado
    
    Exit Property
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método EdicionSuministrador.IDCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Suministrador() As Suministrador
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjSuministrador Is Nothing Then
        Set Suministrador = m_ObjSuministrador
        Exit Property
    End If
    
    Set m_ObjSuministrador = Constructor.getSuministrador(p_Id:=Me.IDSuministrador, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Suministrador = m_ObjSuministrador
    Exit Property
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método EdicionSuministrador.Suministrador ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get Anexo() As Anexo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjAnexo Is Nothing Then
        Set Anexo = m_ObjAnexo
        Exit Property
    End If
    If Me.IDAnexo = "" Then
        Exit Property
    End If
    
    Set m_ObjAnexo = Constructor.getAnexo(Me.IDAnexo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    Set Anexo = m_ObjAnexo
    
    Exit Property
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método EdicionSuministrador.Anexo ha devuelto el error: " & vbNewLine & Err.Description
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
    Set m_objEdicion = Constructor.getEdicion(Me.IDEdicion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    Set Edicion = m_objEdicion
    
    
     Exit Property
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método EdicionSuministrador.Edicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get URLAnexo() As String
    
    Dim m_Anexo As Anexo
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sURLAnexo <> "" Then
        URLAnexo = m_sURLAnexo
        Exit Property
    End If
    Set m_Anexo = Me.Anexo
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Anexo Is Nothing Then
        Exit Property
    End If
    m_sURLAnexo = m_Anexo.URLAnexo
    m_Error = m_Anexo.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    URLAnexo = m_sURLAnexo
    
    Exit Property
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método EdicionSuministrador.URLAnexo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Function MotivoAnexarTecnicoNoOK( _
                                        p_URLLocal As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_URLFinal As String
    Dim m_Edicion As Edicion
    On Error GoTo errores
    
    Set m_Edicion = Me.Edicion
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Edicion Is Nothing Then
        MotivoAnexarTecnicoNoOK = "No se conoce la edición"
        Exit Function
    End If
    If m_Edicion.EsActivo = EnumSiNo.No Then
        MotivoAnexarTecnicoNoOK = "Se pretende anexar una evidencia de suministrador en una edición no activa"
        Exit Function
    End If
    If Not fso.FileExists(p_URLLocal) Then
        MotivoAnexarTecnicoNoOK = "No se alcanza en estos momentos el anexo"
        Exit Function
    End If
    If FicheroAbierto(p_URLLocal) Then
        MotivoAnexarTecnicoNoOK = "Ha de cerrar primero el anexo"
        Exit Function
    End If
    m_URLFinal = NombreProximoAnexoSinExtension(Me.Edicion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If fso.FileExists(m_URLFinal) Then
        If FicheroAbierto(m_URLFinal) Then
            MotivoAnexarTecnicoNoOK = "Existe el anexo en el servidor y parece estar abierto."
            Exit Function
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MotivoAnexarTecnicoNoOK ha devuelto el error: " & Err.Description
    End If
End Function
Public Function AnexarPorTecnico( _
                                p_URLLocal As String, _
                                Optional ByRef p_Error As String _
                                ) As String
                            
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    Dim m_Anexo As Anexo
    
    On Error GoTo errores
    
    Me.Error = ""
    
    m_MotivoNoOK = MotivoAnexarTecnicoNoOK(p_URLLocal, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set m_Anexo = New Anexo
    With m_Anexo
        .IDEdicion = Me.IDEdicion
        .EvidenciaUTE = "No"
        .EvidenciaSuministrador = "Sí"
        .Titulo = "ED " & Me.Edicion.Edicion & " Acuerdo tratamiento Riesgos con Suministrador " & Me.Suministrador.Nombre
        .Registrar Me.Edicion, p_URLLocal, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    If m_Anexo.IDAnexo = "" Then
        p_Error = "El alta del anexo no ha dado los resultados esperados"
        Err.Raise 1000
    End If
    Me.IDAnexo = m_Anexo.IDAnexo
    m_SQL = "UPDATE TbProyectosEdicionesSuministradores SET IDAnexo = " & Me.IDAnexo & " " & _
            "WHERE ID=" & Me.ID & ";"
    getdb.Execute m_SQL

    Me.Edicion.ColSuministradoresVaciar
    InvalidarPublicabilidadPorCambioEvidenciaEdicion Me.IDEdicion, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexarPorTecnico ha devuelto el error: " & Err.Description
    End If
End Function

Public Function EliminarEvidencia( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                            
    Dim m_SQL As String
    Dim m_Anexo As Anexo
    
    On Error GoTo errores
    
    Me.Error = ""
    
    Set m_Anexo = Me.Anexo
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Anexo Is Nothing Then
        m_Anexo.EliminarAnexo p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Me.IDAnexo = ""
    m_SQL = "UPDATE TbProyectosEdicionesSuministradores SET IDAnexo =Null " & _
            "WHERE ID=" & Me.ID & ";"
    getdb.Execute m_SQL
    Me.Edicion.ColSuministradoresVaciar
    InvalidarPublicabilidadPorCambioEvidenciaEdicion Me.IDEdicion, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EliminarEvidencia ha devuelto el error: " & Err.Description
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
        .Add "IDSuministrador"
        .Add "IDAnexo"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EdicionSuministrador.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "ID" Then
        getPropiedad = Me.ID
        Exit Function
    End If
    If m_NombrePropiedad = "IDEdicion" Then
        getPropiedad = Me.IDEdicion
        Exit Function
    End If
    If m_NombrePropiedad = "IDSuministrador" Then
        getPropiedad = Me.IDSuministrador
        Exit Function
    End If
    If m_NombrePropiedad = "IDAnexo" Then
        getPropiedad = Me.IDAnexo
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en EdicionSuministrador.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EdicionSuministrador.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
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
    If m_NombrePropiedad = "IDSuministrador" Then
        Me.IDSuministrador = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDAnexo" Then
        Me.IDAnexo = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en EdicionSuministrador.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EdicionSuministrador.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function




