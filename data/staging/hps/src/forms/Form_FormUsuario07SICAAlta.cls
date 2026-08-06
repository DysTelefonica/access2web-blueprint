Attribute VB_Name = "Form_FormUsuario07SICAAlta"
Option Compare Database
Option Explicit
Private m_Error As String
Private m_TramitacionExternaAnterior As String
Private m_URLLocal As String
Public Event AltaRegistro()


Private Function RellenarObjetoConDatosLocal(Optional ByRef p_Error As String) As String
    
    Dim m_NombreCampo As String
    Dim m_ValorCampo As String
    Dim m_objColCampos As Collection
    Dim m_ID As String
    Dim ctl As Control
    On Error GoTo errores
    
    p_Error = ""
    RellenarObjetoConDatos m_ObjUsuarioSICAActivo, Me, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Nz(Me.UsuarioHPS, "") <> "" Then
        m_ID = getIDUsuario(Me, , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ID <> "" Then
            m_ObjUsuarioSICAActivo.IDHPS = m_ID
        End If
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarObjetoConDatosLocal ha producido el error: " & Err.Description
    End If
End Function

Private Sub CmdSalir_Click()
    DoCmd.Close acForm, Me.Name, acSaveYes
End Sub

Private Sub ComandoAnexar_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_URLLocal = Seleccionar1(True, "Seleccione el anexo", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not fso.FileExists(m_URLLocal) Then
        m_URLLocal = ""
        Me.ImagenAnexo.Visible = False
        Me.ComandoEliminarAnexoActual.Enabled = True
    Else
        Me.ImagenAnexo.Visible = True
        Me.ComandoEliminarAnexoActual.Enabled = True
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al usar el método ComandoAnexar_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarAnexoActual_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    m_URLLocal = ""
    Me.ImagenAnexo.Visible = False
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al usar el método ComandoEliminarAnexoActual_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub ComandoGrabar_Click()
    
    Dim m_ID As String
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    RellenarObjetoConDatosLocal m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
   
    m_ObjUsuarioSICAActivo.Alta m_URLLocal, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_ID = m_ObjUsuarioSICAActivo.ID
    pregunta = MsgBox("Alta correcta", vbInformation, "Alta de UsuarioSICA")
    ActualizarUsuarioSICA m_ID, EnumSiNo.No, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    'DoCmd.Close acForm, Me.Name, acSaveNo
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGrabar_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub Form_Load()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    AjustarTamaño Me
    If Application.TempVars("SoloLectura") = "Sí" Then
        Me.AllowEdits = False
    Else
        Me.AllowEdits = True
    End If
    
    
    Set m_ObjUsuarioSICAActivo = New UsuarioSICA
    m_ObjUsuarioSICAActivo.ParaAlta = EnumSiNo.Sí
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.AllowEdits = True Then
        Me.ComandoGrabar.Enabled = True
        Me.ComandoEmpresas.Enabled = True
    Else
        Me.ComandoGrabar.Enabled = False
        Me.ComandoEmpresas.Enabled = False
    End If
    RellenarComboEmpresas Me.IDEmpresaTramitadora, EnumSiNo.No, , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    RellenarComboUsuarioHPS Me.UsuarioHPS, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarComboUsuariosHistoricos Me.UsuarioHistorico, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Caption = Me.lblTitulo.Caption
    CopiarEstiloBoton Me.ComandoGrabar
    CopiarEstiloBoton Me.cmdSalir
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ImagenAnexo_Click()
    
   On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    
    If Not fso.FileExists(m_URLLocal) Then
        m_Error = "No es alcanzable en estos momentos el anexo seleccionado" & vbNewLine & m_URLLocal
        Err.Raise 1000
    End If
    
    Ejecutar Me.hWnd, "open", m_URLLocal, "", "", 1
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ImagenAnexo_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub TramitacionExterna_BeforeUpdate(Cancel As Integer)
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Me.AllowEdits = False Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    
    
    If Nz(Me.TramitacionExterna, "") = "" Then
        Me.IDEmpresaTramitadora = Null
        Me.IDEmpresaTramitadora.Enabled = False
        Me.ComandoEmpresas.Enabled = False
        Me.Nombre = Null
        Me.Nombre.Enabled = False
        Me.Apellido_1 = Null
        Me.Apellido_1.Enabled = False
        Me.Apellido_2 = Null
        Me.Apellido_2.Enabled = False
        Me.Correo_e = Null
        Me.Correo_e.Enabled = False
        Me.UsuarioSICATSOL = Null
        Me.UsuarioSICATSOL.Enabled = False
        Me.UsuarioSICATDE = Null
        Me.UsuarioSICATDE.Enabled = False
        Me.FAltaTSOL = Null
        Me.FAltaTSOL.Enabled = False
        Me.FAltaTdE = Null
        Me.FAltaTdE.Enabled = False
        Me.FbajaTSOL = Null
        Me.FbajaTSOL.Enabled = False
        Me.Observaciones = Null
        Me.Observaciones.Enabled = False
        m_TramitacionExternaAnterior = ""
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Me.ComandoAnexar.Enabled = True
    
    If Me.TramitacionExterna = "Sí" Then
        
        
        If m_TramitacionExternaAnterior = "Sí" Then
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
            Exit Sub
        End If
        Me.UsuarioHPS = Null
        Me.UsuarioHPS.Enabled = False
        Me.IDEmpresaTramitadora.Enabled = True
        Me.ComandoEmpresas.Enabled = True
        Me.Nombre.Enabled = True
        Me.Apellido_1.Enabled = True
        Me.Apellido_2.Enabled = True
        Me.Correo_e.Enabled = True
        Me.UsuarioSICATSOL.Enabled = True
        Me.UsuarioSICATDE.Enabled = True
        Me.FAltaTSOL.Enabled = True
        Me.FAltaTdE.Enabled = True
        Me.FbajaTSOL.Enabled = True
        Me.FBajaTdE.Enabled = True
        Me.Observaciones.Enabled = True
        Me.ComandoEmpresas.Enabled = True
        m_TramitacionExternaAnterior = "Sí"
    Else
        If m_TramitacionExternaAnterior = "No" Then
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
            Exit Sub
        End If
        Me.UsuarioHPS = Null
        Me.UsuarioHPS.Enabled = True
        Me.Nombre = Null
        Me.Nombre.Enabled = False
        Me.Apellido_1 = Null
        Me.Apellido_1.Enabled = False
        Me.Apellido_2 = Null
        Me.Apellido_2.Enabled = False
        Me.Correo_e = Null
        Me.Correo_e.Enabled = False
        Me.UsuarioSICATSOL = Null
        Me.UsuarioSICATSOL.Enabled = False
        Me.UsuarioSICATDE = Null
        Me.UsuarioSICATDE.Enabled = False
        Me.FAltaTSOL = Null
        Me.FAltaTSOL.Enabled = False
        Me.FAltaTdE = Null
        Me.FAltaTdE.Enabled = False
        Me.FbajaTSOL = Null
        Me.FbajaTSOL.Enabled = False
        Me.Observaciones = Null
        Me.Observaciones.Enabled = False
        Me.IDEmpresaTramitadora.Enabled = False
        Me.ComandoEmpresas.Enabled = False
        m_TramitacionExternaAnterior = "No"
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al TramitacionExterna_BeforeUpdate se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub UsuarioHistorico_Exit(Cancel As Integer)
    
    Dim m_IDUsuarioHistorico As String
    Dim m_ObjUsuarioHistorico As UsuarioHistorico
    
    On Error GoTo errores
    
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Me.AllowEdits = False Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    m_IDUsuarioHistorico = getIDUsuarioHistorico(Me, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_IDUsuarioHistorico = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Set m_ObjUsuarioHistorico = getUsuarioHistorico(m_IDUsuarioHistorico, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjUsuarioHistorico Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    With m_ObjUsuarioHistorico
        RellenarComboEmpresas Me.IDEmpresaTramitadora, EnumSiNo.Sí, , m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Me.IDEmpresaTramitadora = .IDEmpresaHPS
        Me.IDEmpresaTramitadora.Enabled = False
        Me.ComandoEmpresas.Enabled = False
        Me.Nombre = .Nombre
        Me.Nombre.Enabled = False
        Me.Apellido_1 = .Apellido_1
        Me.Apellido_1.Enabled = False
        If .Apellido_2 <> "" Then
            Me.Apellido_2 = .Apellido_2
        Else
            Me.Apellido_2 = Null
        End If
        Me.Apellido_2.Enabled = False
        Me.Correo_e = .Correo_e
        Me.Correo_e.Enabled = False
        Me.UsuarioSICATSOL.Enabled = True
        Me.FAltaTSOL.Enabled = True
        Me.FbajaTSOL.Enabled = True
        Me.UsuarioSICATDE.Enabled = True
        Me.FAltaTdE.Enabled = True
        Me.FBajaTdE.Enabled = True
        Me.Observaciones.Enabled = True
        Me.UsuarioSICATDE.SetFocus
    End With
    Set m_ObjUsuarioHistorico = Nothing
   
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al UsuarioHistorico_Exit se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub UsuarioHPS_Exit(Cancel As Integer)
    
    Dim m_ID As String
    Dim m_objUsuarioHPS As UsuarioHPS
    
    On Error GoTo errores
    
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Me.AllowEdits = False Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    m_ID = getIDUsuario(Me, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ID = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    Set m_objUsuarioHPS = getUsuarioHPS(m_ID, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_objUsuarioHPS Is Nothing Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    With m_objUsuarioHPS
        RellenarComboEmpresas Me.IDEmpresaTramitadora, EnumSiNo.Sí, , m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Me.IDEmpresaTramitadora = .IDEmpresaHPS
        Me.IDEmpresaTramitadora.Enabled = False
        Me.ComandoEmpresas.Enabled = False
        Me.Nombre = .Nombre
        Me.Nombre.Enabled = False
        Me.Apellido_1 = .Apellido_1
        Me.Apellido_1.Enabled = False
        If .Apellido_2 <> "" Then
            Me.Apellido_2 = .Apellido_2
        Else
            Me.Apellido_2 = Null
        End If
        Me.Apellido_2.Enabled = False
        Me.Correo_e = .Correo_e
        Me.Correo_e.Enabled = False
        Me.UsuarioSICATSOL.Enabled = True
        Me.FAltaTSOL.Enabled = True
        Me.FbajaTSOL.Enabled = True
        Me.UsuarioSICATDE.Enabled = True
        Me.FAltaTdE.Enabled = True
        Me.FBajaTdE.Enabled = True
        Me.Observaciones.Enabled = True
        Me.UsuarioSICATDE.SetFocus
    End With
    Set m_objUsuarioHPS = Nothing
   
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al UsuarioHPS_Exit se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub