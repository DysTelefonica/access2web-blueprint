VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnexoUsuarioSICA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDAnexo As String
Public NombreAnexo As String
Public IDUsuarioSICA As String
Public Hash As String

Public URLLocal As String
Private m_sURLAnexo As String
Private m_objUsuarioSICA As UsuarioSICA
Private m_sIDNuevo As String
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String



Public Property Get UsuarioSICA() As UsuarioSICA

    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If Not m_objUsuarioSICA Is Nothing Then
        Set UsuarioSICA = m_objUsuarioSICA
        Exit Property
    End If
    If Me.IDUsuarioSICA = "" Then
        Exit Property
    End If
    Set m_objUsuarioSICA = getUsuarioSICA(Me.IDUsuarioSICA, , , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioSICA = m_objUsuarioSICA
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método AnexoUsuarioSICA.UsuarioSICA ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get URLAnexo() As String
    
    Dim m_URLDirectorioAnexos As String
    
    Dim m_DNI As String
    
    On Error GoTo errores
    Me.Error = ""
    If Me.NombreAnexo = "" Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se ha podido determinar el nombre del anexo SICA"
        Err.Raise 1000
    End If
    m_URLDirectorioAnexos = m_ObjEntorno.URLCarpetaAnexosSICA
    m_Error = m_ObjEntorno.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_sURLAnexo = m_URLDirectorioAnexos & "\" & Me.NombreAnexo
    URLAnexo = m_sURLAnexo
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método AnexoUsuarioSICA.URLAnexo ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get IDNuevo() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sIDNuevo = getID("TbAnexosUsuariosSICA", "IDAnexo", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    IDNuevo = m_sIDNuevo
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método AnexoUsuarioSICA.IDNuevo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Function MotivoNoOK( _
                            Optional ByRef p_Error As String _
                                ) As String
   
    
    
    On Error GoTo errores
    If Application.TempVars("SoloLectura") = "Sí" Then
        p_Error = "El Sistema está en modo sólo lectura"
        Err.Raise 1000
    End If
    With Me
        If .UsuarioSICA Is Nothing Then
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            p_Error = "No se ha podido determinar el usuario SICA del anexo"
            Err.Raise 1000
        Else
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        
        If Not fso.FileExists(.URLLocal) Then
            MotivoNoOK = "El URLLocal es obligatorio"
            Exit Function
'        Else
'            If FicheroAbierto(.URLLocal) Then
'                MotivoNoOK = "El URLLocal parece estar abierto"
'                Exit Function
'            End If
        End If
        If fso.FileExists(.URLAnexo) Then
            If FicheroAbierto(.URLAnexo) Then
                MotivoNoOK = "El URLAnexo ya existe y parece estar abierto"
                Exit Function
            End If
        End If
        If Not Me.UsuarioSICA.AnexoSICA Is Nothing Then
            MotivoNoOK = "Ya existe un anexo para este usuario SICA"
            Exit Function
        End If
        
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método UsuarioSICA.MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function

Private Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                    
   
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Application.TempVars("SoloLectura") = "Sí" Then
        p_Error = "El Sistema está en modo sólo lectura"
        Err.Raise 1000
    End If
    With Me
        If fso.FileExists(.URLAnexo) Then
           If FicheroAbierto(.URLAnexo) Then
                p_Error = "El URLAnexo parece estar abierto"
                Err.Raise 1000
            End If
        End If
    End With
    
        
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexoUsuarioSICA.ValidacionEliminar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function Registrar( _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    Dim blnAnexado As Boolean
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    
    m_MotivoNoOK = MotivoNoOK(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.NombreAnexo = fso.GetFileName(Me.URLLocal)
    Me.IDAnexo = Me.IDNuevo
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    AdjuntarAnexo p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    blnAnexado = True
    
    m_SQL = "TbAnexosUsuariosSICA"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("IDAnexo") = Me.IDAnexo
            
            .Fields("NombreAnexo") = Me.NombreAnexo
            .Fields("IDUsuarioSICA") = Me.IDUsuarioSICA
        .Update
    
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Anexo.registrar ha devuelto el error: " & Err.Description
    End If
End Function


Private Function AdjuntarAnexo( _
                                Optional ByRef p_Error As String _
                                ) As String
                            
    
    Dim m_URLLocal As String
    Dim m_URLRemoto As String
    Dim m_URLDirectorioFinal As String
    
    
    On Error GoTo errores
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        p_Error = "El Sistema está en modo sólo lectura"
        Err.Raise 1000
    End If
    
    m_URLLocal = Me.URLLocal
    m_URLRemoto = Me.URLAnexo
    m_URLDirectorioFinal = fso.GetParentFolderName(m_URLRemoto)
    If Not fso.FolderExists(m_URLDirectorioFinal) Then
        p_Error = "No se puede acceder al directorio de Anexos SICA: " & vbNewLine & m_URLDirectorioFinal
        Err.Raise 1000
    End If
    fso.CopyFile m_URLLocal, m_URLRemoto, True
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioSICA.AdjuntarAnexo ha devuelto el error: " & Err.Description
    End If
    
End Function


Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
                           
    Dim m_SQL As String
    
    On Error GoTo errores
    ValidacionEliminar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If fso.FileExists(Me.URLAnexo) Then
        fso.DeleteFile Me.URLAnexo, True
    End If
    m_SQL = "DELETE * " & _
            "FROM TbAnexosUsuariosSICA " & _
            "WHERE IDAnexo=" & Me.IDAnexo & ";"
    getdb().Execute m_SQL
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioSICA.Borrar ha devuelto el error: " & Err.Description
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
        .Add "IDUsuarioSICA"
        .Add "Hash"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método AnexoUsuarioSICA.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
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
    If m_NombrePropiedad = "IDUsuarioSICA" Then
        getPropiedad = Me.IDUsuarioSICA
        Exit Function
    End If
    If m_NombrePropiedad = "Hash" Then
        getPropiedad = Me.Hash
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en AnexoUsuarioSICA.getPropiedad"
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexoUsuarioSICA.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
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
    If m_NombrePropiedad = "IDUsuarioSICA" Then
        Me.IDUsuarioSICA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Hash" Then
        Me.Hash = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en AnexoUsuarioSICA.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexoUsuarioSICA.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function





