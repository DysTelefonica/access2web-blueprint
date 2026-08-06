VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UltimoProyecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDUltimoProyecto As String
Public IDProyecto As String
Public Usuario As String
Public Fecha As String

Private m_sIDUltimoProyectoCalculado As String
Private m_ObjProyecto As Proyecto
Private m_objUsuarioObj As Usuario

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get IDUltimoProyectoCalculado() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sIDUltimoProyectoCalculado = DameID("TbUltimoProyecto", "IDUltimoProyecto", , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDUltimoProyectoCalculado = m_sIDUltimoProyectoCalculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método UltimoProyecto.IDUltimoProyectoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

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
        m_Error = "No se conoce el IDProyecto"
        Err.Raise 1000
    End If
    Set m_ObjProyecto = Constructor.getProyecto(Me.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Proyecto = m_ObjProyecto
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método UltimoProyecto.Proyecto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    

End Property

Public Property Set Proyecto(ByVal objNewValue As Proyecto)

    Set m_ObjProyecto = objNewValue

End Property

Public Property Get UsuarioObj() As Usuario

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objUsuarioObj Is Nothing Then
        Set UsuarioObj = m_objUsuarioObj
        Exit Property
    End If
    If Me.Usuario = "" Then
        m_Error = "No se conoce el Usuario"
        Err.Raise 1000
    End If
    Set m_objUsuarioObj = Constructor.getUsuario(, Me.Usuario, , , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioObj = m_objUsuarioObj
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método UltimoProyecto.UsuarioObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    

End Property

Public Property Set UsuarioObj(ByVal objNewValue As Usuario)

    Set m_objUsuarioObj = objNewValue

End Property

Public Function Registrar( _
                            p_IDProyecto As String, _
                            p_UsuarioRed As String, _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_ID As String
    Dim m_SQL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    If p_IDProyecto = "" Or p_UsuarioRed = "" Then
        Exit Function
    End If
    
    m_SQL = "SELECT TbUltimoProyecto.* " & _
            "FROM TbUltimoProyecto " & _
            "WHERE Usuario='" & p_UsuarioRed & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            m_ID = Me.IDUltimoProyectoCalculado
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            .AddNew
                .Fields("IDUltimoProyecto") = m_ID
                .Fields("IDProyecto") = p_IDProyecto
                .Fields("Usuario") = p_UsuarioRed
                .Fields("Fecha") = Now()
            .Update
        Else
            .Edit
                .Fields("IDProyecto") = p_IDProyecto
                .Fields("Fecha") = Now()
            .Update
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UltimoProyecto.Registrar ha devuelto el error: " & Err.Description
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
        .Add "IDUltimoProyecto"
        .Add "IDProyecto"
        .Add "Usuario"
        .Add "Fecha"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método UltimoProyecto.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDUltimoProyecto" Then
        getPropiedad = Me.IDUltimoProyecto
        Exit Function
    End If
    If m_NombrePropiedad = "IDProyecto" Then
        getPropiedad = Me.IDProyecto
        Exit Function
    End If
    If m_NombrePropiedad = "Usuario" Then
        getPropiedad = Me.Usuario
        Exit Function
    End If
    If m_NombrePropiedad = "Fecha" Then
        getPropiedad = Me.Fecha
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en UltimoProyecto.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UltimoProyecto.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDUltimoProyecto" Then
        Me.IDUltimoProyecto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDProyecto" Then
        Me.IDProyecto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Usuario" Then
        Me.Usuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Fecha" Then
        Me.Fecha = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en UltimoProyecto.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UltimoProyecto.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


