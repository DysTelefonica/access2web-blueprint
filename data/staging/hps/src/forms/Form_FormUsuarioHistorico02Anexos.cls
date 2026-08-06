Attribute VB_Name = "Form_FormUsuarioHistorico02Anexos"
Option Compare Database
Option Explicit
Private m_Error As String
Dim m_AnexoSeleccionado As AnexoUsuarioHistorico


Private Function EstablecerLista(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_Anexo As AnexoUsuarioHistorico
    Dim m_Col As Scripting.Dictionary
    Dim m_IDAnexo As Variant
    
    
    On Error GoTo errores
    m_Error = ""
    Set lst = Me.ListaAnexos
    lst.RowSource = ""
    Set m_Col = m_ObjUsuarioHistoricoActivo.ColAnexos
    p_Error = m_ObjUsuarioHistoricoActivo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        For Each m_IDAnexo In m_Col
            Set m_Anexo = m_Col(m_IDAnexo)
            lst.AddItem m_IDAnexo & ";" & m_Anexo.NombreAnexo
            
            Set m_Anexo = Nothing
        Next
        If lst.ListCount = 1 Then
            lst.Selected(1) = True
            ListaAnexos_Click
        End If
        
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerLista ha dado el error: " & vbNewLine & Err.Description
    End If
End Function



Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    Me.AllowEdits = False
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    If m_ObjUsuarioHistoricoActivo Is Nothing Then
        m_Error = "No se ha seleccionado ningún usuario"
        Err.Raise 1000
    End If
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.AllowEdits = True
    Me.ListaAnexos.Enabled = True
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo

End Sub

Private Sub ListaAnexos_Click()
    
    Dim m_IDAnexoSeleccionado As String
    
    
    On Error GoTo errores
    m_Error = ""
    m_IDAnexoSeleccionado = Nz(Me.ListaAnexos.Column(0), "")
    If m_IDAnexoSeleccionado = "" Then
        Set m_AnexoSeleccionado = Nothing
        Exit Sub
    End If
    If m_AnexoSeleccionado Is Nothing Then
        Set m_AnexoSeleccionado = Constructor.getAnexoUsuarioHistorico(m_IDAnexoSeleccionado, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_AnexoSeleccionado Is Nothing Then
            m_Error = "El Anexo del Usuario Histórico seleccionado no parece estar registrado"
            Err.Raise 1000
        End If
    Else
        If m_IDAnexoSeleccionado <> m_AnexoSeleccionado.IDAnexo Then
            Set m_AnexoSeleccionado = getAnexoUsuarioHistorico(m_IDAnexoSeleccionado, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            If m_AnexoSeleccionado Is Nothing Then
                m_Error = "El Anexo del Usuario Histórico seleccionado no parece estar registrado"
                Err.Raise 1000
            End If
        End If
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaAnexos_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ListaAnexos_DblClick(Cancel As Integer)
    
    
    On Error GoTo errores
    m_Error = ""
    Application.Echo False
    If m_AnexoSeleccionado Is Nothing Then
        Set m_AnexoSeleccionado = Constructor.getAnexoUsuarioHistorico(Nz(Me.ListaAnexos.Column(0), ""), m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_AnexoSeleccionado Is Nothing Then
            m_Error = "El Anexo del Usuario Histórico seleccionado no parece estar registrado"
            Err.Raise 1000
        End If
    End If
    
    If Not fso.FileExists(m_AnexoSeleccionado.URLAnexo) Then
        m_Error = "No es alcanzable en estos momentos el anexo Histórico seleccionado" & vbNewLine & m_AnexoSeleccionado.URLAnexo
        Err.Raise 1000
    End If
    Ejecutar Me.hWnd, "open", m_AnexoSeleccionado.URLAnexo, "", "", 1
    Application.Echo True

    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El intentar ver el anexo del usuario histórico ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Application.Echo True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
