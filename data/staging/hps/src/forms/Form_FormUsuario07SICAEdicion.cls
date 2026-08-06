Attribute VB_Name = "Form_FormUsuario07SICAEdicion"
Option Compare Database
Option Explicit
Private m_Error As String

Private m_AnexoAnInicio As AnexoUsuarioSICA

Private m_TramitacionExternaAnterior As String
Private m_IDHPSAlInicio As String
Private m_IDHPSHistoricoAlInicio As String

Public Event EdicionRegistro()

Private Function RellenarObjetoConDatosLocal(Optional ByRef p_Error As String) As String
    
    Dim m_NombreCampo As String
    Dim m_ValorCampo As String
    Dim ctl As Control
    Dim m_ID As String
    On Error GoTo errores
    
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
            m_ObjUsuarioSICAActivo.IDHPSHistorico = ""
        End If
    Else
        m_ID = getIDUsuarioHistorico(Me, , p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ID <> "" Then
            m_ObjUsuarioSICAActivo.IDHPSHistorico = m_ID
            m_ObjUsuarioSICAActivo.IDHPS = ""
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarObjetoConDatosLocal ha producido el error: " & Err.Description
    End If
End Function
Public Function HaHabidoCambiosLocal(Optional ByRef p_Error As String) As Boolean
    
    
    Dim m_IDHPS As String
    Dim m_IDHPSHistorico As String
    Dim m_HaHabidoCambios As Boolean
    On Error GoTo errores
    
    m_HaHabidoCambios = HaHabidoCambios(m_ObjUsuarioSICAAlInicio, Me, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_HaHabidoCambios = True Then
        HaHabidoCambiosLocal = m_HaHabidoCambios
        Exit Function
    End If
    If Nz(Me.UsuarioHPS, "") <> "" Then
        m_IDHPS = getIDUsuario(, Me.UsuarioHPS, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_IDHPS <> m_IDHPSAlInicio Then
            HaHabidoCambiosLocal = True
            Exit Function
        End If
    End If
    If Nz(Me.UsuarioHistorico, "") <> "" Then
        m_IDHPSHistorico = getIDUsuarioHistorico(, Me.UsuarioHistorico, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_IDHPSHistorico <> m_IDHPSHistoricoAlInicio Then
            HaHabidoCambiosLocal = True
            Exit Function
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambiosLocal ha producido el error: " & Err.Description
    End If

End Function

Public Function getUsuarioSICAAlInicio(Optional ByRef p_Error As String) As UsuarioSICA
    
    On Error GoTo errores
    
    Set getUsuarioSICAAlInicio = m_ObjUsuarioSICAAlInicio
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuarioSICAAlInicio ha producido el error: " & Err.Description
    End If
End Function


Private Sub CmdSalir_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ComandoAnexar_Click()
    On Error GoTo errores
    Dim m_URLLocal As String
    
    
    m_Error = ""
    If Not m_AnexoAnInicio Is Nothing Then
        pregunta = MsgBox("¿Desea quitar el anexo actual y cambiarlo por otro?", vbExclamation + vbYesNo + vbDefaultButton2, "Anexar")
        If pregunta <> vbYes Then
            Exit Sub
        End If
    End If
    m_URLLocal = Seleccionar1(True, "Seleccione el anexo", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    If Not fso.FileExists(m_URLLocal) Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If Not m_AnexoAnInicio Is Nothing Then
        m_AnexoAnInicio.Borrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
    End If
    Set m_AnexoAnInicio = New AnexoUsuarioSICA
    With m_AnexoAnInicio
        .IDUsuarioSICA = m_ObjUsuarioSICAAlInicio.ID
        .URLLocal = m_URLLocal
        .Registrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Me.ImagenAnexo.Visible = True
    Me.ComandoEliminarAnexoActual.Enabled = True
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

Private Sub ComandoDetalleUsuarioHPS_Click()
    Dim frm As Form
    Dim frmLateral As Form
    Dim rst As DAO.Recordset
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Not FormularioAbierto("FormInicial00Principal") Then
        DoCmd.OpenForm "FormInicial00Principal"
    End If
    Set frm = Forms("FormInicial00Principal")
    Set frmLateral = frm.Controls("SubFormLateral").Form
    If frm.Controls("SubFormCentral").SourceObject <> "FormInicial02ConsultasPrincipal" Then
        If m_ObjUsuarioSICAActivo.IDHPS <> "" Then
            Form_FormInicial01Lateral.lblUsuariosHPS_Click
        Else
            Form_FormInicial01Lateral.lblUsuariosHistoricos_Click
        End If
       
        
    End If
    Set frm = frm.Controls("SubFormCentral").Form
    Set frm = frm.Controls("SubFormResultados").Form
    If m_ObjUsuarioSICAActivo.IDHPS <> "" Then
        Set rst = frm.RecordsetClone 'Clonamos el recordset
        rst.FindFirst "[ID] =" & m_ObjUsuarioSICAActivo.IDHPS   'Buscamos
        If Not rst.NoMatch Then 'Si  se encuentra
            frm.Bookmark = rst.Bookmark 'Nos vamos al registro encontrado
            frm.IDExpediente.SetFocus 'Activamos el campo ID (ahora es cuando se produce el scroll)
            Form_FormInicial02ConsultasPrincipal.ComandoEditarUsuario_Click
        End If
    Else
        Form_FormInicial08UsuariosHistoricos.Nombre = Trim(m_ObjUsuarioSICAActivo.Nombre & _
                                                            " " & m_ObjUsuarioSICAActivo.Apellido_1)
        Form_FormInicial08UsuariosHistoricos.Nombre_Exit False
        Form_FormInicial08UsuariosHistoricos.ComandoVerUsuarioHistorico_Click
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoDetalleUsuarioHPS_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarAnexoActual_Click()
    On Error GoTo errores
    
    m_Error = ""
    If m_AnexoAnInicio Is Nothing Then
        m_Error = "No hay anexo que eliminar"
        Err.Raise 1000
    End If
    pregunta = MsgBox("¿Desea eliminar definitivamente del Sistema el anexo seleccionado?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Anexo")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_AnexoAnInicio.Borrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_AnexoAnInicio = Nothing
    Set m_ObjUsuarioSICAActivo.AnexoSICA = Nothing
    Set m_AnexoAnInicio = m_ObjUsuarioSICAActivo.AnexoSICA
    If m_AnexoAnInicio Is Nothing Then
        m_ObjUsuarioSICAActivo.URLAnexo = ""
        Me.ImagenAnexo.Visible = False
        Me.ComandoAnexar.Enabled = True
        Me.ComandoAnexar.SetFocus
        Me.ComandoEliminarAnexoActual.Enabled = False
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EliminarAnexoActual ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
     DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub ComandoGrabar_Click()
    
    Dim m_HaHabidoCambios As Boolean
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
    m_HaHabidoCambios = HaHabidoCambiosLocal(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_HaHabidoCambios = False Then
        m_Error = "No hay cambios que guardar"
        Err.Raise 1000
    End If
    m_ObjUsuarioSICAActivo.ParaAlta = EnumSiNo.No
    m_ObjUsuarioSICAActivo.Edicion m_ObjUsuarioSICAAlInicio, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_ID = m_ObjUsuarioSICAActivo.ID
    pregunta = MsgBox("Edición correcta", vbInformation, "Edición de UsuarioSICA")
    ActualizarUsuarioSICA m_ID, EnumSiNo.No, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RaiseEvent EdicionRegistro
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    DoCmd.Close acForm, Me.Name, acSaveNo
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGrabar_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Sub Form_Load()
    
    On Error GoTo errores
    
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If Application.TempVars("SoloLectura") = "Sí" Then
        Me.AllowEdits = False
    Else
        Me.AllowEdits = True
    End If
    
    If m_ObjUsuarioSICAActivo Is Nothing Then
        m_Error = "No se ha abierto el formulario con los datos necesarios"
        Err.Raise 1000
    End If
    Set m_ObjUsuarioSICAAlInicio = Constructor.getUsuarioSICA(m_ObjUsuarioSICAActivo.ID, , , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.ComandoGrabar.Enabled = False
    Me.ComandoEmpresas.Enabled = False
    If Me.AllowEdits = True Then
        Me.ComandoGrabar.Enabled = True
        
    End If

    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    CopiarEstiloBoton Me.ComandoGrabar
    CopiarEstiloBoton Me.cmdSalir
    CopiarEstiloBoton Me.ComandoDetalleUsuarioHPS
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al abrir el formulario se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim ctl As Control
    
    Dim m_Dato As String
    Dim m_NombreCampo As String
    Dim m_Titulo As String
    
    
    On Error GoTo errores
    
    p_Error = ""
    
    
    With m_ObjUsuarioSICAAlInicio
        
        Me.TramitacionExterna = .TramitacionExterna
        If Me.TramitacionExterna = "Sí" Then
            Me.UsuarioHPS = Null
            Me.UsuarioHPS.Enabled = False
            Me.UsuarioHistorico = Null
            Me.UsuarioHistorico.Enabled = False
            Me.ComandoDetalleUsuarioHPS.Enabled = False
        Else
            RellenarComboUsuarioHPS Me.UsuarioHPS, m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            RellenarComboUsuariosHistoricos Me.UsuarioHistorico, m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            If .IDHPS <> "" Then
                Set m_ObjUsuarioHistoricoActivo = Nothing
                Set m_ObjUsuarioActivo = .UsuarioHPS
                p_Error = .Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If Not m_ObjUsuarioActivo Is Nothing Then
                    Me.UsuarioHPS = m_ObjUsuarioActivo.Nombre & " " & _
                                m_ObjUsuarioActivo.Apellido_1 & " " & _
                                m_ObjUsuarioActivo.Apellido_1 & " " & _
                                "(" & m_ObjUsuarioActivo.ID & ")"
                     Me.ComandoDetalleUsuarioHPS.Enabled = True
                     UsuarioHPS_Exit False
                    If Me.AllowEdits = True Then
                        Me.UsuarioHPS.Enabled = True
                    Else
                        Me.UsuarioHPS.Enabled = False
                    End If
                Else
                    Me.ComandoDetalleUsuarioHPS.Enabled = False
                End If
            Else
                Set m_ObjUsuarioHistoricoActivo = .UsuarioHistorico
                p_Error = .Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                
                If Not m_ObjUsuarioHistoricoActivo Is Nothing Then
                    Me.UsuarioHistorico = m_ObjUsuarioHistoricoActivo.Nombre & " " & _
                                m_ObjUsuarioHistoricoActivo.Apellido_1 & " " & _
                                m_ObjUsuarioHistoricoActivo.Apellido_1 & " " & _
                                "(" & m_ObjUsuarioHistoricoActivo.ID & ")"
                     Me.ComandoDetalleUsuarioHPS.Enabled = True
                     UsuarioHistorico_Exit False
                    If Me.AllowEdits = True Then
                        Me.UsuarioHistorico.Enabled = True
                    Else
                        Me.UsuarioHistorico.Enabled = False
                    End If
                Else
                    Me.ComandoDetalleUsuarioHPS.Enabled = False
                End If
            End If
            
            
        End If
        If .Nombre <> "" Then
            Me.Nombre = .Nombre
        End If
        If .Apellido_1 <> "" Then
            Me.Apellido_1 = .Apellido_1
        End If
        If .Apellido_2 <> "" Then
            Me.Apellido_2 = .Apellido_2
        End If
        If .Correo_e <> "" Then
            Me.Correo_e = .Correo_e
        End If
        If .UsuarioSICATDE <> "" Then
            Me.UsuarioSICATDE = .UsuarioSICATDE
        End If
        If IsDate(.FAltaTdE) Then
            Me.FAltaTdE = .FAltaTdE
        End If
        If IsDate(.FBajaTdE) Then
            Me.FBajaTdE = .FBajaTdE
        End If
        
        If .UsuarioSICATSOL <> "" Then
            Me.UsuarioSICATSOL = .UsuarioSICATSOL
        End If
        If IsDate(.FAltaTSOL) Then
            Me.FAltaTSOL = .FAltaTSOL
        End If
        If IsDate(.FbajaTSOL) Then
            Me.FbajaTSOL = .FbajaTSOL
        End If
        If .Observaciones <> "" Then
            Me.Observaciones = .Observaciones
        End If
        
        Set m_AnexoAnInicio = .AnexoSICA
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Me.AllowEdits = True Then
            Me.ComandoAnexar.Enabled = True
        Else
            Me.ComandoAnexar.Enabled = False
        End If
        If Not m_AnexoAnInicio Is Nothing Then
            If m_AnexoAnInicio.URLAnexo <> "" Then
                Me.ImagenAnexo.Visible = True
                If Me.AllowEdits = True Then
                    Me.ComandoEliminarAnexoActual.Enabled = True
                Else
                    Me.ComandoEliminarAnexoActual.Enabled = False
                End If
            Else
                Me.ImagenAnexo.Visible = False
                If Me.AllowEdits = True Then
                    Me.ComandoEliminarAnexoActual.Enabled = False
                Else
                    Me.ComandoEliminarAnexoActual.Enabled = False
                End If
            End If
        Else
            Me.ImagenAnexo.Visible = False
            If Me.AllowEdits = True Then
                Me.ComandoEliminarAnexoActual.Enabled = False
            Else
                Me.ComandoEliminarAnexoActual.Enabled = False
            End If
        End If
        
        RellenarComboEmpresas Me.IDEmpresaTramitadora, EnumSiNo.No, , m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_IDHPSAlInicio = .IDHPS
        m_IDHPSHistoricoAlInicio = .IDHPSHistorico
        m_Titulo = "DETALLE USUARIO SICA (" & .Nombre & " " & _
                                        .Apellido_1 & " " & _
                                        .Apellido_2 & ")"
            Me.lblTitulo.Caption = m_Titulo
            Me.Caption = Me.lblTitulo.Caption
    End With
    If Me.AllowEdits = True Then
        If Me.TramitacionExterna = "Sí" Then
            Me.UsuarioHPS.Enabled = False
            Me.UsuarioHistorico.Enabled = False
            Me.Nombre.Enabled = True
            Me.Apellido_1.Enabled = True
            Me.Apellido_2.Enabled = True
            Me.Correo_e.Enabled = True
            Me.Nombre.SetFocus
        Else
            Me.UsuarioHPS.Enabled = True
            Me.UsuarioHistorico.Enabled = True
            Me.Nombre.Enabled = False
            Me.Apellido_1.Enabled = False
            Me.Apellido_2.Enabled = False
            Me.Correo_e.Enabled = False
            
        End If
        Me.IDEmpresaTramitadora.Enabled = True
        
        Me.UsuarioSICATDE.Enabled = True
        Me.FAltaTdE.Enabled = True
        Me.FBajaTdE.Enabled = True
        Me.UsuarioSICATSOL.Enabled = True
        Me.FAltaTSOL.Enabled = True
        Me.FbajaTSOL.Enabled = True
        Me.Observaciones.Enabled = True
       
    End If
    If Nz(Me.IDEmpresaTramitadora, "") = "" Then
        If m_ObjUsuarioSICAAlInicio.IDEmpresaTramitadora <> "" Then
            Me.IDEmpresaTramitadora = m_ObjUsuarioSICAAlInicio.IDEmpresaTramitadora
        End If
    End If
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Then
            If InStr(1, ctl.Tag, "DATOS;INTOCABLE") <> 0 And Me.AllowEdits = True Then
                ctl.Enabled = False
            End If
        End If

    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & Err.Description
    End If
End Function


Private Sub ImagenAnexo_Click()
    
    Dim m_URL As String
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_AnexoAnInicio Is Nothing Then
        m_Error = "No se ha abierto el formulario con los datos necesarios"
        Err.Raise 1000
    End If
    If m_AnexoAnInicio Is Nothing Then
        m_Error = "Este usuario no tiene anexo"
        Err.Raise 1000
    End If
    m_URL = m_AnexoAnInicio.URLAnexo
    
    If Not fso.FileExists(m_URL) Then
        m_Error = "No es alcanzable en estos momentos el anexo seleccionado" & vbNewLine & m_URL
        Err.Raise 1000
    End If
    Ejecutar Me.hWnd, "open", m_URL, "", "", 1
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
        Me.UsuarioHistorico = Null
        Me.UsuarioHPS.Enabled = False
        Me.UsuarioHistorico.Enabled = False
        Me.IDEmpresaTramitadora.Enabled = True
        Me.Nombre.Enabled = True
        Me.Apellido_1.Enabled = True
        Me.Apellido_2.Enabled = True
        Me.Correo_e.Enabled = True
        Me.UsuarioSICATSOL.Enabled = True
        Me.UsuarioSICATDE.Enabled = True
        Me.FAltaTSOL.Enabled = True
        Me.FAltaTdE.Enabled = True
        Me.FbajaTSOL.Enabled = True
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
        Me.UsuarioHistorico = Null
        Me.UsuarioHPS.Enabled = True
        Me.UsuarioHistorico.Enabled = True
        Me.IDEmpresaTramitadora = Null
        Me.IDEmpresaTramitadora.Enabled = False
        Me.Nombre = Null
        Me.Nombre.Enabled = False
        Me.Apellido_1 = Null
        Me.Apellido_1.Enabled = False
        Me.Apellido_2 = Null
        Me.Apellido_2.Enabled = False
        Me.Correo_e = Null
        Me.Correo_e.Enabled = False
        
        
        m_TramitacionExternaAnterior = "No"
        RellenarComboUsuarioHPS Me.UsuarioHPS, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        RellenarComboUsuariosHistoricos Me.UsuarioHistorico, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
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