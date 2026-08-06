VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnexoUsuarioHPS"
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
Public EsHistorico As String


Public URLLocal As String
Private m_sURLAnexo As String
Private m_ObjUsuario As UsuarioHPS
Private m_sIDNuevo As String
Private m_ObjDb As DAO.Database

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

' Holds the DAO.Database the instance was loaded through (set by
' Constructor.getAnexoUsuarioHPS). Lets the lazy Usuario lookup reuse
' the same backend instead of falling back to getdb(), which can point
' at a different store in the test context (TempVars("DatosEnLocal")
' resolves to m_URLRutaAplicacionesLocal, not to the injected HPST.accdb).
' Nothing = legacy path: getUsuarioHPS uses getdb().
Public Property Set p_Db(ByVal value As DAO.Database)
    Set m_ObjDb = value
End Property

Public Property Get Usuario() As UsuarioHPS

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
    Set m_ObjUsuario = getUsuarioHPS(Me.idUsuario, , m_Error, m_ObjDb)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Usuario = m_ObjUsuario
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método AnexoUsuarioHPS.Usuario ha devuelto el error: " & vbNewLine & Err.Description
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
    
    If EsHistorico = "Sí" Then
        m_URLDirectorioAnexos = Me.Usuario.URLCarpetaAnexosHistoricos
    ElseIf EsHistorico = "No" Then
        m_URLDirectorioAnexos = Me.Usuario.URLCarpetaAnexosActuales
    End If
    m_sURLAnexo = m_URLDirectorioAnexos & Me.NombreAnexo
    URLAnexo = m_sURLAnexo
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método AnexoUsuarioHPS.URLAnexo ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property

Public Property Get IDNuevo() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDNuevo = getID("TbAnexosUsuariosHPS", "IDAnexo", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDNuevo = m_sIDNuevo
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método AnexoUsuarioHPS.IDNuevo ha devuelto el error: " & vbNewLine & Err.Description
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
        If .Usuario Is Nothing Then
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            p_Error = "No se ha podido determinar el usuario del anexo"
            Err.Raise 1000
        Else
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If .EsHistorico <> "Sí" And .EsHistorico <> "No" Then
            MotivoNoOK = "Se ha de indicar si es o no histórico el Anexo del Usuario HPS"
            Exit Function
        End If
        If .URLLocal = "" Then
            MotivoNoOK = "El URLLocal es obligatorio"
            Exit Function
        End If
'        If FSO.FileExists(.URLLocal) Then
'            If FicheroAbierto(.URLLocal) Then
'                MotivoNoOK = "El URLLocal parece estar abierto"
'                Exit Function
'            End If
'
'
'        End If
        If fso.FileExists(.URLAnexo) Then
            If FicheroAbierto(.URLAnexo) Then
                MotivoNoOK = "El URLAnexo ya existe y parece estar abierto"
                Exit Function
            End If
        End If
        If .NombreAnexo = "" Then
            .NombreAnexo = fso.GetFileName(.URLLocal)
            
        End If
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método AnexoUsuarioHPS.MotivoNoOK ha devuelto el error: " & Err.Description
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
        p_Error = "El método AnexoUsuarioHPS.ValidacionEliminar ha devuelto el error: " & Err.Description
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
    If Not fso.FileExists(Me.URLLocal) Then
        Exit Function
    End If
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
    m_SQL = "TbAnexosUsuariosHPS"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("IDAnexo") = Me.IDAnexo
            .Fields("NombreAnexo") = Me.NombreAnexo
            .Fields("IDUsuario") = Me.idUsuario
            .Fields("EsHistorico") = Me.EsHistorico
        .Update
    
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexoUsuarioHPS.registrar ha devuelto el error: " & Err.Description
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
    
    
    
    Me.IDAnexo = Me.IDNuevo
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_SQL = "TbAnexosUsuariosHPS"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("IDAnexo") = Me.IDAnexo
            .Fields("NombreAnexo") = Me.NombreAnexo
            .Fields("IDUsuario") = Me.idUsuario
            .Fields("EsHistorico") = Me.EsHistorico
        .Update
    
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexoUsuarioHPS.RegistroParaNormalizar ha devuelto el error: " & Err.Description
    End If
End Function

Private Function AdjuntarAnexo( _
                                Optional ByRef p_Error As String _
                                ) As String
                            
    
    Dim m_URLLocal As String
    Dim m_URLRemoto As String
    Dim m_URLDirDNI As String
    Dim m_URLDirAnexos As String
    Dim m_URLDirectorioFinal As String
    Dim m_Hash As String
    
    On Error GoTo errores
    
    
    m_URLLocal = Me.URLLocal
    If Me.EsHistorico <> "Sí" And Me.EsHistorico <> "No" Then
       p_Error = "No se sabe si el anexo es histórico o no"
       Err.Raise 1000
    End If
    If Me.EsHistorico = "Sí" Then
        m_URLDirectorioFinal = Me.Usuario.URLCarpetaAnexosHistoricos
    Else
        m_URLDirectorioFinal = Me.Usuario.URLCarpetaAnexosActuales
    End If
    p_Error = Me.Usuario.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Dim m_DirDNI As String
    If Not fso.FolderExists(m_URLDirectorioFinal) Then
        If Not fso.FolderExists(fso.GetParentFolderName(m_URLDirectorioFinal)) Then
            fso.CreateFolder fso.GetParentFolderName(m_URLDirectorioFinal)
        End If
        fso.CreateFolder m_URLDirectorioFinal
    End If
    m_URLRemoto = m_URLDirectorioFinal & Me.NombreAnexo
    fso.CopyFile m_URLLocal, m_URLRemoto, True
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexoUsuarioHPS.AdjuntarAnexo ha devuelto el error: " & Err.Description
    End If
    
End Function



Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
                           
    Dim m_SQL As String
    
    On Error GoTo errores
    
    
'    ValidacionEliminar p_Error
'    If p_Error <> "" Then
'        Err.Raise 1000
'    End If
'
    If fso.FileExists(Me.URLAnexo) Then
        fso.DeleteFile Me.URLAnexo, True
    End If
    
    m_SQL = "DELETE * " & _
            "FROM TbAnexosUsuariosHPS " & _
            "WHERE IDAnexo=" & Me.IDAnexo & ";"
    getdb().Execute m_SQL
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexoUsuarioHPS.Borrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function PasarAnexo( _
                            Optional ByRef p_AHistorico As EnumSiNo = EnumSiNo.Sí, _
                            Optional ByRef p_Error As String _
                            ) As String

    ' [anexos-batch-hotfix] Compatibility wrapper. Delegates the move to
    ' AnexoSelectionTransactionCoordinator so a single-row call benefits
    ' from the same preflight + snapshot + transaction + filesystem
    ' compensation machinery that batch callers use. EsHistorico on the
    ' in-memory instance is updated ONLY after a confirmed commit; the
    ' coordinator itself rejects any concurrent-direction or missing-row
    ' state before opening the transaction.
    Dim m_Coordinador As AnexoSelectionTransactionCoordinator
    Dim m_IDs As Collection

    On Error GoTo errores

    p_Error = ""
    If p_AHistorico = Empty Then
        p_Error = "Se ha de indicar si se quiere pasar de actual a histórico o viceversa"
        Err.Raise 1000
    End If

    Set m_IDs = New Collection
    m_IDs.Add CLng(Me.IDAnexo)

    Set m_Coordinador = New AnexoSelectionTransactionCoordinator
    If m_Coordinador.MoverLote(m_IDs, (p_AHistorico = EnumSiNo.Sí), Nothing, "", p_Error) Then
        ' Update in-memory state only after a confirmed commit. The
        ' coordinator already flipped the DB row; mirroring it here keeps
        ' any caller that re-reads Me.EsHistorico in sync.
        If p_AHistorico = EnumSiNo.Sí Then
            Me.EsHistorico = "Sí"
        Else
            Me.EsHistorico = "No"
        End If
    ElseIf Len(p_Error) = 0 Then
        p_Error = "PasarAnexo: el coordinador no comunicó detalle del fallo."
        Err.Raise 1000
    Else
        Err.Raise 1000
    End If

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexoUsuarioHPS.PasarAnexo ha devuelto el error: " & Err.Description
    End If

End Function


Public Function PasarAnexoAHistorico( _
                                    Optional ByRef p_Error As String _
                                    ) As String

    ' [anexos-batch-hotfix] Compatibility wrapper. The historical entry point
    ' has no external callers in the current source tree, but its signature
    ' is preserved (delegating to the coordinator) so any legacy code path
    ' that still references it picks up the same atomicity guarantees as
    ' the modern PasarAnexo path.
    Dim m_Coordinador As AnexoSelectionTransactionCoordinator
    Dim m_IDs As Collection

    On Error GoTo errores

    Set m_IDs = New Collection
    m_IDs.Add CLng(Me.IDAnexo)

    Set m_Coordinador = New AnexoSelectionTransactionCoordinator
    If m_Coordinador.MoverLote(m_IDs, True, Nothing, "", p_Error) Then
        Me.EsHistorico = "Sí"
    ElseIf Len(p_Error) = 0 Then
        p_Error = "PasarAnexoAHistorico: el coordinador no comunicó detalle del fallo."
        Err.Raise 1000
    Else
        Err.Raise 1000
    End If

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexoUsuarioHPS.PasarAnexoAHistorico ha devuelto el error: " & Err.Description
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
        .Add "EsHistorico"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método AnexoUsuarioHPS.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
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
    If m_NombrePropiedad = "EsHistorico" Then
        getPropiedad = Me.EsHistorico
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en AnexoUsuarioHPS.getPropiedad"
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexoUsuarioHPS.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
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
    If m_NombrePropiedad = "EsHistorico" Then
        Me.EsHistorico = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en AnexoUsuarioHPS.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AnexoUsuarioHPS.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function






