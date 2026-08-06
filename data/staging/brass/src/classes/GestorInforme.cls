VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GestorInforme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public ID As String
Public Fecha As String
Public Datos As String
Public Usuario As String
Public Visto As String

Private m_sIDCalculado As String
Private m_objUsuarioObj As Usuario
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get IDCalculado() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    
    m_sIDCalculado = DameID("Tb0FiltroGestion", "ID")
    IDCalculado = m_sIDCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método GestionInforme.IDCalculado ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

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
        Exit Property
    End If
    Set m_objUsuarioObj = Constructor.getUsuario(, Me.Usuario, , , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioObj = m_objUsuarioObj
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método GestionInforme.UsuarioObj ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property

Public Function Registrar( _
                            p_ColDatos As Scripting.Dictionary, _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_ID As String
    Dim m_SQL As String
    Dim m_Datos As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    m_ID = Me.IDCalculado
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ID = "" Then
        p_Error = "No se puede obtener el IDCalculado"
        Err.Raise 1000
    End If
    Me.ID = m_ID
    
    m_Datos = ColADatos(p_ColDatos, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_SQL = "Tb0FiltroGestion"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("ID") = Me.ID
            .Fields("Fecha") = Now()
            If m_Datos <> Empty Then
                .Fields("Datos") = m_Datos
                Me.Datos = m_Datos
            End If
            .Fields("Usuario") = m_ObjUsuarioConectado.UsuarioRed
            .Fields("Visto") = "No"
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GestionInforme.Registrar ha devuelto el error: " & Err.Description
    End If
End Function

Private Function ColADatos( _
                            p_ColDatos As Scripting.Dictionary, _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim m_Campo As Variant
    Dim m_Resultado As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    If p_ColDatos Is Nothing Then
        Exit Function
    End If
    For Each m_Campo In p_ColDatos
        If m_Resultado = "" Then
            m_Resultado = m_Campo
        Else
            m_Resultado = m_Resultado & "|" & m_Campo
        End If
    Next
    ColADatos = m_Resultado
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GestionInforme.Registrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
                                
    On Error GoTo errores
    
    Me.Error = ""
    If Not IsDate(Me.ID) Then
        p_Error = "No se ha indicado el ID"
        Err.Raise 1000
    End If
    m_SQL = "DELETE * FROM Tb0FiltroGestion " & _
            "WHERE ID=" & Me.ID & ";"
    CurrentDb().Execute m_SQL
   
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GestionInforme.Borrar ha devuelto el error: " & Err.Description
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
        .Add "Fecha"
        .Add "Datos"
        .Add "Usuario"
        .Add "Visto"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método GestionInforme.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
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
    If m_NombrePropiedad = "Fecha" Then
        getPropiedad = Me.Fecha
        Exit Function
    End If
    If m_NombrePropiedad = "Datos" Then
        getPropiedad = Me.Datos
        Exit Function
    End If
    If m_NombrePropiedad = "Usuario" Then
        getPropiedad = Me.Usuario
        Exit Function
    End If
    If m_NombrePropiedad = "Visto" Then
        getPropiedad = Me.Visto
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en GestionInforme.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GestionInforme.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
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
    If m_NombrePropiedad = "Fecha" Then
        Me.Fecha = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Datos" Then
        Me.Datos = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Usuario" Then
        Me.Usuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Visto" Then
        Me.Visto = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en GestionInforme.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GestionInforme.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

