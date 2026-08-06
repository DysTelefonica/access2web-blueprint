VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EdicionCorreoRevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public IDEnvioCorreoTecnico As String
Public IDEdicion As String
Public FechaCorreoRevision As String
Public UsuarioCalidad As String
Public IDCorreo As String


Private m_sIDCorreoCalculado As String
Private m_objEdicion As Edicion
Private m_ObjUsuarioCalidadObj As Usuario

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Private Function SQLTexto(ByVal p_Valor As String) As String
    SQLTexto = "'" & Replace(p_Valor, "'", "''") & "'"
End Function

Private Function DebeRegistrarCorreoRevision(Optional ByRef p_Error As String) As Boolean

    On Error GoTo errores
    p_Error = ""
    DebeRegistrarCorreoRevision = False

    If Me.IDEdicion = "" Then
        p_Error = "No se ha informado IDEdicion"
        Err.Raise 1000
    End If
    If Me.Edicion Is Nothing Then
        p_Error = "No se ha encontrado la edición para registrar el correo de revisión"
        Err.Raise 1000
    End If
    If Me.Edicion.Proyecto Is Nothing Then
        p_Error = "No se ha encontrado el proyecto de la edición para registrar el correo de revisión"
        Err.Raise 1000
    End If
    If Me.Edicion.Proyecto.ParaInformeAvisos <> "Sí" Then
        Exit Function
    End If

    DebeRegistrarCorreoRevision = True
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EdicionCorreoRevision.DebeRegistrarCorreoRevision ha devuelto el error: " & Err.Description
    End If
End Function

Public Property Get IDCorreoCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDCorreoCalculado = DameID("TbProyectoEdicionesCorreoRevision", "IDEnvioCorreoTecnico", getdb(), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDCorreoCalculado = m_sIDCorreoCalculado
    
    Exit Property
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método EdicionCorreoRevision.IDCorreoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get Edicion() As Edicion
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not m_objEdicion Is Nothing Then
        Set Edicion = m_objEdicion
        Exit Function
    End If
    If Me.IDEdicion = "" Then
        Exit Property
    End If
    Set m_objEdicion = Constructor.getEdicion(Me.IDEdicion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Edicion = m_objEdicion
    Exit Function
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método EdicionCorreoRevision.Edicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get UsuarioCalidadObj() As Usuario
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not m_ObjUsuarioCalidadObj Is Nothing Then
        Set UsuarioCalidadObj = m_ObjUsuarioCalidadObj
        Exit Property
    End If
    If Me.UsuarioCalidad = "" Then
        Exit Property
    End If
    Set m_ObjUsuarioCalidadObj = Constructor.getUsuario(p_UsuarioRed:=Me.UsuarioCalidad, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioCalidadObj = m_ObjUsuarioCalidadObj
    Exit Property
    
errores:
    If Err.Property <> 1000 Then
        m_Error = "El método EdicionCorreoRevision.UsuarioCalidadObj ha devuelto el error: " & vbNewLine & Err.Description
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
        .Add "IDEnvioCorreoTecnico"
        .Add "IDEdicion"
        .Add "FechaCorreoRevision"
        .Add "UsuarioCalidad"
        .Add "IDCorreo"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EdicionCorreoRevision.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDEnvioCorreoTecnico" Then
        getPropiedad = Me.IDEnvioCorreoTecnico
        Exit Function
    End If
    If m_NombrePropiedad = "IDEdicion" Then
        getPropiedad = Me.IDEdicion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCorreoRevision" Then
        getPropiedad = Me.FechaCorreoRevision
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioCalidad" Then
        getPropiedad = Me.UsuarioCalidad
        Exit Function
    End If
    If m_NombrePropiedad = "IDCorreo" Then
        getPropiedad = Me.IDCorreo
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en EdicionCorreoRevision.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EdicionCorreoRevision.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDEnvioCorreoTecnico" Then
        Me.IDEnvioCorreoTecnico = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEdicion" Then
        Me.IDEdicion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCorreoRevision" Then
        Me.FechaCorreoRevision = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioCalidad" Then
        Me.UsuarioCalidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDCorreo" Then
        Me.IDCorreo = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en EdicionCorreoRevision.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EdicionCorreoRevision.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function Registrar( _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Correo As CORREO
    On Error GoTo errores
    p_Error = ""
    
    If Not DebeRegistrarCorreoRevision(p_Error) Then
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    Set m_Correo = EnviarCorreoRevisionEdicion(Me.Edicion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Correo Is Nothing Then
        p_Error = "No se ha podido enviar el correo"
        Err.Raise 1000
    End If
     m_SQL = "TbProyectoEdicionesCorreoRevision"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            Me.IDCorreo = Me.IDCorreoCalculado
            .Fields("IDEnvioCorreoTecnico") = Me.IDCorreo
            .Fields("IDEdicion") = Me.IDEdicion
            Me.FechaCorreoRevision = Now()
            .Fields("FechaCorreoRevision") = Me.FechaCorreoRevision
            Me.IDCorreo = m_Correo.IDCorreo
            .Fields("IDCorreo") = Me.IDCorreo
            Me.UsuarioCalidad = m_ObjUsuarioConectado.UsuarioRed
            .Fields("UsuarioCalidad") = Me.UsuarioCalidad
            
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

Public Function getSQLRegistrarCorreoEnviado( _
                                            ByVal p_IDEnvioCorreoTecnico As String, _
                                            ByVal p_IDEdicion As String, _
                                            ByVal p_IDCorreo As String, _
                                            ByVal p_FechaCorreoRevision As Date, _
                                            ByVal p_UsuarioCalidad As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String

    On Error GoTo errores
    p_Error = ""

    If p_IDEnvioCorreoTecnico = "" Then
        p_Error = "No se ha informado IDEnvioCorreoTecnico"
        Err.Raise 1000
    End If
    If p_IDEdicion = "" Then
        p_Error = "No se ha informado IDEdicion"
        Err.Raise 1000
    End If
    If p_IDCorreo = "" Then
        p_Error = "No se ha informado IDCorreo"
        Err.Raise 1000
    End If

    getSQLRegistrarCorreoEnviado = "INSERT INTO TbProyectoEdicionesCorreoRevision " & _
        "(IDEnvioCorreoTecnico, IDEdicion, FechaCorreoRevision, UsuarioCalidad, IDCorreo) " & _
        "VALUES (" & p_IDEnvioCorreoTecnico & ", " & _
        p_IDEdicion & ", #" & Format$(p_FechaCorreoRevision, "yyyy\/mm\/dd hh\:nn\:ss") & "#, " & _
        SQLTexto(p_UsuarioCalidad) & ", " & p_IDCorreo & ");"
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EdicionCorreoRevision.getSQLRegistrarCorreoEnviado ha devuelto el error: " & Err.Description
    End If
End Function

Public Function RegistrarCorreoEnviado( _
                                        ByVal p_IDEdicion As String, _
                                        ByVal p_IDCorreo As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    Dim m_SQL As String
    On Error GoTo errores
    p_Error = ""

    Me.IDEdicion = p_IDEdicion
    If Not DebeRegistrarCorreoRevision(p_Error) Then
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    Me.IDEnvioCorreoTecnico = Me.IDCorreoCalculado
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.IDCorreo = p_IDCorreo
    Me.FechaCorreoRevision = Now()
    Me.UsuarioCalidad = m_ObjUsuarioConectado.UsuarioRed

    m_SQL = getSQLRegistrarCorreoEnviado(Me.IDEnvioCorreoTecnico, Me.IDEdicion, Me.IDCorreo, CDate(Me.FechaCorreoRevision), Me.UsuarioCalidad, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    getdb().Execute m_SQL, dbFailOnError
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EdicionCorreoRevision.RegistrarCorreoEnviado ha devuelto el error: " & Err.Description
    End If
End Function


