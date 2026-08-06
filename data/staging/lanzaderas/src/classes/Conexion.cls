VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Conexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public IDConexion As String
Public usuario As String
Public FechaConexion As String
Public FechaCierre As String
Public ConContraseña As Boolean
Public UsuarioSSID As String
Public EnOficina As Boolean
Public Vertical As String
Public Horizontal As String


Private m_objUsuarioObj As usuario
Private m_sIDConexionCalculada As String
Private m_objColCampos As Collection
Public Error As String

Public Property Get UsuarioObj() As usuario
    
    On Error GoTo errores
    
    Me.Error = ""
    If Not m_objUsuarioObj Is Nothing Then
        Set UsuarioObj = m_objUsuarioObj
        Exit Property
    End If
    If Me.usuario = "" Then
        Exit Property
    End If
    Set m_objUsuarioObj = Constructor.getUsuario(, , , Me.usuario, Me.Error)
    Set UsuarioObj = m_objUsuarioObj
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.UsuarioObj ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property

Public Property Set UsuarioObj(ByVal objNewValue As usuario)

    Set m_objUsuarioObj = objNewValue

End Property

Public Property Get IDConexionCalculada() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_sIDConexionCalculada = DameID("TbConexionesRegistro", "IDConexion", getdb(), Me.Error)
    If Me.Error <> "" Then
        Exit Property
    End If
    IDConexionCalculada = m_sIDConexionCalculada
    Exit Property
errores:
    If Err.Number <> 1000 Then
        Me.Error = "El método Aplicacion.IDConexionCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If

End Property



Public Function RegistroConexion( _
                                p_SSID As String, _
                                p_EnOficina As Boolean, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_Fecha As String
    
    On Error GoTo errores
    
    Me.Error = ""
    If p_SSID = "" Then
        p_Error = "No se conoce el SSID del usuario del registro"
        Err.Raise 1000
    End If
    
    Me.IDConexion = Me.IDConexionCalculada
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.usuario = "" And Me.UsuarioObj Is Nothing Then
        p_Error = "No se conoce el usuario del registro"
        Err.Raise 1000
    End If
    If Me.UsuarioObj Is Nothing Then
        Set Me.UsuarioObj = Constructor.getUsuario(, , , Me.usuario, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If Me.UsuarioObj Is Nothing Then
        p_Error = "No se conoce el usuario del registro"
        Err.Raise 1000
    End If
    
    
    m_Fecha = Now()
    m_SQL = "TbConexionesRegistro"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("IDConexion") = Me.IDConexion
            .Fields("Usuario") = Me.UsuarioObj.CorreoUsuario
            .Fields("FechaConexion") = m_Fecha
            .Fields("ConContraseña") = True
            .Fields("UsuarioSSID") = p_SSID
            .Fields("EnOficina") = p_EnOficina
            
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    RegistroConexion = Me.IDConexion
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistroConexion ha devuelto el error: " & vbNewLine & Err.Description
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
    Set tdf = getdb().TableDefs("TbConexionesRegistro")
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
        Me.Error = "El metodo Conexion.ColCampos ha devuelto el error: " & vbNewLine & Err.Description
    End If


End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If m_NombrePropiedad = "IDConexion" Then
        getPropiedad = Me.IDConexion
        Exit Function
    End If
    If m_NombrePropiedad = "Usuario" Then
        getPropiedad = Me.usuario
        Exit Function
    End If
    If m_NombrePropiedad = "FechaConexion" Then
        getPropiedad = Me.FechaConexion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCierre" Then
        getPropiedad = Me.FechaCierre
        Exit Function
    End If
    If m_NombrePropiedad = "ConContraseña" Then
        getPropiedad = Me.ConContraseña
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioSSID" Then
        getPropiedad = Me.UsuarioSSID
        Exit Function
    End If
    If m_NombrePropiedad = "EnOficina" Then
        getPropiedad = Me.EnOficina
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Conexion.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Conexion.getPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function



Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If m_NombrePropiedad = "IDConexion" Then
        Me.IDConexion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Usuario" Then
        Me.usuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaConexion" Then
        Me.FechaConexion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCierre" Then
        Me.FechaCierre = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ConContraseña" Then
        Me.ConContraseña = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioSSID" Then
        Me.UsuarioSSID = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EnOficina" Then
        Me.EnOficina = m_ValorPropiedad
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Conexion.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El metodo Conexion.SetPropiedad ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function





