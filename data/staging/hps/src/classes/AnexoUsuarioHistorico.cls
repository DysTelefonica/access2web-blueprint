VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnexoUsuarioHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDAnexo As String
Public NombreAnexo As String
Public idUsuario As String
Public Hash As String



Public URLLocal As String
Private m_sURLAnexo As String
Private m_ObjUsuario As UsuarioHistorico
Private m_sIDNuevo As String

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get Usuario() As UsuarioHistorico

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_ObjUsuario Is Nothing Then
        Set Usuario = m_ObjUsuario
        Exit Property
    End If
    If Me.idUsuario = "" Then
       Exit Property
    End If
    Set m_ObjUsuario = getUsuarioHistorico(Me.idUsuario, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Usuario = m_ObjUsuario
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método AnexoUsuarioHistorico.Usuario ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get URLAnexo() As String
    
    Dim m_URLDirectorioAnexos As String
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Me.NombreAnexo = "" Then
        m_Error = "No se ha podido determinar el nombre del anexo"
        Err.Raise 1000
    End If
    If Me.Usuario Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_URLDirectorioAnexos = Me.Usuario.URLCarpetaAnexos
    
    m_sURLAnexo = m_URLDirectorioAnexos & Me.NombreAnexo
    URLAnexo = m_sURLAnexo
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método AnexoUsuarioHistorico.URLAnexo ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Function Registrar( _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    Dim blnAnexado As Boolean
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    If Me.IDAnexo = "" Then
        p_Error = "No se ha indicado el IDAnexo"
        Err.Raise 1000
    End If
    AdjuntarAnexo p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    blnAnexado = True
    m_SQL = "TbAnexosUsuariosHistoricos"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("IDAnexo") = DameID("TbAnexosUsuariosHistoricos", "IDAnexo", getdb())
            .Fields("NombreAnexo") = Me.NombreAnexo
            .Fields("IDUsuario") = Me.idUsuario
        .Update
    
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexoUsuarioHistorico.registrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function RegistroParaNormalizar( _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                                    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    If Me.IDAnexo = "" Then
        p_Error = "No se ha indicado el IDAnexo"
        Err.Raise 1000
    End If
    
    m_SQL = "TbAnexosUsuariosHistoricos"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("IDAnexo") = Me.IDAnexo
            .Fields("NombreAnexo") = Me.NombreAnexo
            .Fields("IDUsuario") = Me.idUsuario
        .Update
    
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexoUsuarioHistorico.RegistroParaNormalizar ha devuelto el error: " & Err.Description
    End If
End Function
Private Function AdjuntarAnexo( _
                                Optional ByRef p_Error As String _
                                ) As String
                            
    
    Dim m_URLLocal As String
    Dim m_URLRemoto As String
    Dim m_URLDirAnexos As String
    Dim m_URLDirectorioFinal As String
    
    
    On Error GoTo errores
    
    
    
    If Me.Usuario Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = "No se conoce el Usuario histórico"
        Err.Raise 1000
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
    End If
     m_URLDirectorioFinal = Me.Usuario.URLCarpetaAnexos
    
    p_Error = Me.Usuario.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not fso.FolderExists(m_URLDirectorioFinal) Then
        If Not fso.FolderExists(fso.GetParentFolderName(m_URLDirectorioFinal)) Then
            fso.CreateFolder fso.GetParentFolderName(m_URLDirectorioFinal)
        End If
        fso.CreateFolder m_URLDirectorioFinal
    End If
    m_URLRemoto = m_URLDirectorioFinal & Me.NombreAnexo
    fso.CopyFile Me.URLLocal, m_URLRemoto, True
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexoUsuarioHistorico.AdjuntarAnexo ha devuelto el error: " & Err.Description
    End If
    
End Function



Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
                           
    Dim m_SQL As String
    
    On Error GoTo errores
    
    
    
    If fso.FileExists(Me.URLAnexo) Then
        fso.DeleteFile Me.URLAnexo, True
    End If
    
    m_SQL = "DELETE * " & _
            "FROM TbAnexosUsuariosHistoricos " & _
            "WHERE IDAnexo=" & Me.IDAnexo & ";"
    getdb().Execute m_SQL
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexoUsuarioHistorico.Borrar ha devuelto el error: " & Err.Description
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
        .Add "NombreAnexo"
        .Add "IDUsuario"
        .Add "Hash"
        
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método AnexoUsuarioHistorico.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDAnexo" Then
        getPropiedad = Me.IDAnexo
        Exit Function
    End If
    If m_NombrePropiedad = "NombreAnexo" Then
        getPropiedad = Me.NombreAnexo
        Exit Function
    End If
    If m_NombrePropiedad = "IDUsuario" Then
        getPropiedad = Me.idUsuario
        Exit Function
    End If
    If m_NombrePropiedad = "Hash" Then
        getPropiedad = Me.Hash
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en AnexoUsuarioHistorico.getPropiedad"
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexoUsuarioHistorico.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
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
    If m_NombrePropiedad = "NombreAnexo" Then
        Me.NombreAnexo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDUsuario" Then
        Me.idUsuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Hash" Then
        Me.Hash = m_ValorPropiedad
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en AnexoUsuarioHistorico.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexoUsuarioHistorico.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function








