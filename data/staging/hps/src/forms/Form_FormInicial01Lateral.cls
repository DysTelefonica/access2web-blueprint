Attribute VB_Name = "Form_FormInicial01Lateral"
Option Compare Database
Option Explicit
Private WithEvents m_Indicadores As clsIndicadoresBus
Private m_Error As String

Private Sub ComandoCambiarEstado_Click()
    
    On Error GoTo errores
    m_Error = ""
    DoCmd.OpenForm "FormInicial"
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al cambiar de estado se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub





Private Sub Form_Load()

    On Error GoTo errores

    m_Error = ""


    Me.lblVersion.Caption = "Versión: " & m_ObjEntorno.VersionAplicacion
    Me.lblDatosEn.Caption = "Datos en: " & m_ObjEntorno.DatosEn
    Me.lblModoEscritura.Caption = "Modo Escritura: " & m_ObjEntorno.ModoEscritura

    EstablecerContadores m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_Indicadores = modIndicadores.m_Indicadores
    CopiarEstiloBoton Me.ComandoCambiarEstado, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Application.TempVars("SoloLectura") = "Sí" Then
        Me.ComandoCambiarEstado.Caption = "Pasar a Escritura"

    Else
        Me.ComandoCambiarEstado.Caption = "Pasar a Lectura"
    End If
    If Application.TempVars("SoloLectura") = "Sí" Then
        Me.ComandoCambiarEstado.Enabled = False
    Else
        Me.ComandoCambiarEstado.Enabled = True
    End If


    SeleccionarEnlace , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If

    lblUsuariosHPS_Click
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al abrir el formulario se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    m_Error = ""
    EstablecerContadores m_Error
    On Error GoTo 0
End Sub

Private Sub m_Indicadores_LifecycleMoveCompleted(p_Direction As String, p_UserId As String)
    On Error Resume Next
    m_Error = ""
    EstablecerContadores m_Error
    On Error GoTo 0
End Sub



Private Sub lblConfiguracion_Click()
    On Error GoTo errores
    m_Error = ""
    SeleccionarEnlace Me.lblConfiguracion, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al hacer click en el control se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub lblHPSAPuntoDeCaducar_Click()
    On Error GoTo errores
    m_Error = ""
    SeleccionarEnlace Me.lblHPSAPuntoDeCaducar, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al hacer click en el control se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub lblHPSCaducados_Click()
    On Error GoTo errores
    m_Error = ""
    SeleccionarEnlace Me.lblHPSCaducados, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al hacer click en el control se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub lblHPSEstadoIrregular_Click()
    On Error GoTo errores
    m_Error = ""
    SeleccionarEnlace Me.lblHPSEstadoIrregular, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al hacer click en el control se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub lblIndicadores_Click()
    On Error GoTo errores
    m_Error = ""
    SeleccionarEnlace Me.lblIndicadores, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al lblIndicadores_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub lblRequiereCurso_Click()
    On Error GoTo errores
    m_Error = ""
    SeleccionarEnlace Me.lblRequiereCurso, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al lblRequiereCurso_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub lblRequiereCursoPteCorreoJefeSeguridad_Click()
    On Error GoTo errores
    m_Error = ""
    SeleccionarEnlace Me.lblRequiereCursoPteCorreoJefeSeguridad, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al lblRequiereCursoPteCorreoJefeSeguridad_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub lblRequiereCursoPtePrimeraConvocatoria_Click()
    On Error GoTo errores
    m_Error = ""
    SeleccionarEnlace Me.lblRequiereCursoPtePrimeraConvocatoria, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al lblRequiereCursoPtePrimeraConvocatoria_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub lblRequiereCursoPteSegundaConvocatoria_Click()
    On Error GoTo errores
    m_Error = ""
    SeleccionarEnlace Me.lblRequiereCursoPteSegundaConvocatoria, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al lblRequiereCursoPteSegundaConvocatoria_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub lblSolicitados_Click()
    On Error GoTo errores
    m_Error = ""
    SeleccionarEnlace Me.lblSolicitados, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al lblSolicitados_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Sub lblUsuariosHistoricos_Click()
    On Error GoTo errores
    m_Error = ""
    SeleccionarEnlace Me.lblUsuariosHistoricos, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al hacer click en el control se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Sub lblUsuariosHPS_Click()
    
    On Error GoTo errores
    m_Error = ""
    SeleccionarEnlace Me.lblUsuariosHPS, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al hacer click en el control se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Function SeleccionarEnlace(Optional ByRef p_lbl As Label, Optional ByRef p_Error As String) As String
    
    Dim ctl As Control
    Dim m_Frm As Form
    Dim m_frmPpal As Form
    Dim m_frmParametros As Form
    Dim m_NombreControl As String
    Dim m_nombreControlSeleccionado As String
    Dim m_Particula As String
    
    On Error GoTo errores
    
    If p_lbl Is Nothing Then
        For Each ctl In Me.Controls
            If TypeOf ctl Is Label Then
                If InStr(1, ctl.Tag, ";") <> 0 Then
                    ctl.Tag = "Elegibles;NoSeleccionado"
                End If
            End If
        Next
        Exit Function
    End If
    m_nombreControlSeleccionado = p_lbl.Name
    p_lbl.Tag = "Elegibles;Seleccionado"
    For Each ctl In Me.Controls
        If TypeOf ctl Is Label Then
            m_NombreControl = ctl.Name
            If m_NombreControl <> m_nombreControlSeleccionado Then
                If InStr(1, Nz(ctl.Tag, ";"), ";") <> 0 Then
                    ctl.Tag = "Elegibles;NoSeleccionado"
                End If
            End If
            
        End If
    Next
    
    Set m_Frm = Me.Parent
    'Set ctl = m_frm.lblTitulo
    'With ctl
        If p_lbl.Name = "lblUsuariosHPS" Then
            If m_Frm.SubFormCentral.SourceObject <> "FormInicial02ConsultasPrincipal" Then
                m_Frm.SubFormCentral.SourceObject = "FormInicial02ConsultasPrincipal"
                '.Left = (m_frm.InsideWidth / 2) - (ctl.Width / 2)
                EstablecerControlesformlateral Me, p_Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                Set m_Frm = m_Frm.SubFormCentral.Form
                m_Frm.Controls("ListaConsultas").Selected(0) = True
                Form_FormInicial02ConsultasPrincipal.ListaConsultas_Click
                Form_FormInicial03ConsultasEstado.ComandoLimpiarEstadoHPS_Click
            Else
                If Me.lblHPSEstadoIrregular.BackColor = m_ObjEntorno.EtiquetaLinkBackColorSeleccionado Or _
                    Me.lblHPSAPuntoDeCaducar.BackColor = m_ObjEntorno.EtiquetaLinkBackColorSeleccionado Or _
                    Me.lblHPSCaducados.BackColor = m_ObjEntorno.EtiquetaLinkBackColorSeleccionado Then
                    '.Left = (m_frm.InsideWidth / 2) - (ctl.Width / 2)
                    EstablecerControlesformlateral Me, p_Error
                    If m_Error <> "" Then
                        Err.Raise 1000
                    End If
                    Set m_Frm = m_Frm.SubFormCentral.Form
                    m_Frm.Controls("ListaConsultas").Selected(0) = True
                    Form_FormInicial02ConsultasPrincipal.ListaConsultas_Click
                    Form_FormInicial03ConsultasEstado.ComandoLimpiarEstadoHPS_Click
                End If
            End If
            
            
          
        
        ElseIf p_lbl.Name = "lblHPSAPuntoDeCaducar" Then
            
            lblIndicadores_Click
            Set m_Frm = m_Frm.SubFormCentral.Form
            If m_Frm.Controls("ComandoAPuntoDeCaducar").Enabled = True Then
                Form_FormInicial09Indicadores.ComandoAPuntoDeCaducar_Click
            End If
        ElseIf p_lbl.Name = "lblHPSCaducados" Then
            
            lblIndicadores_Click
            Set m_Frm = m_Frm.SubFormCentral.Form
            If m_Frm.Controls("ComandoCaducados").Enabled = True Then
                Form_FormInicial09Indicadores.ComandoCaducados_Click
            End If
        ElseIf p_lbl.Name = "lblHPSEstadoIrregular" Then
            m_Particula = "Irregular"
        
            If m_Frm.SubFormCentral.SourceObject <> "FormInicial02ConsultasPrincipal" Then
                m_Frm.SubFormCentral.SourceObject = "FormInicial02ConsultasPrincipal"
            End If
            '.Left = (m_frm.InsideWidth / 2) - (ctl.Width / 2)
            EstablecerControlesformlateral Me, p_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            
            '.Width = 3059
            Set m_frmPpal = m_Frm.SubFormCentral.Form
            m_frmPpal.Controls("ListaConsultas").Selected(0) = True
            
            Form_FormInicial02ConsultasPrincipal.ListaConsultas_Click
            Set m_frmPpal = m_frmPpal.SubFormParametrosConsulta.Form
            m_frmPpal.Controls("ComboEstadoHPS").value = "Irregular"
            Form_FormInicial03ConsultasEstado.ComboEstadoHPS_AfterUpdate
            
        
        ElseIf p_lbl.Name = "lblRequiereCurso" Then
            
            lblIndicadores_Click
            Set m_Frm = m_Frm.SubFormCentral.Form
            If m_Frm.Controls("ComandoPendientesCursoHPS").Enabled = True Then
                Form_FormInicial09Indicadores.ComandoPendientesCursoHPS_Click
            End If
        ElseIf p_lbl.Name = "lblRequiereCursoPtePrimeraConvocatoria" Then
            
            lblIndicadores_Click
            Set m_Frm = m_Frm.SubFormCentral.Form
            If m_Frm.Controls("ComandoPendientesConvocatoria1").Enabled = True Then
                Form_FormInicial09Indicadores.ComandoPendientesConvocatoria1_Click
            End If
        ElseIf p_lbl.Name = "lblRequiereCursoPteSegundaConvocatoria" Then
            
            lblIndicadores_Click
            Set m_Frm = m_Frm.SubFormCentral.Form
            If m_Frm.Controls("ComandoPendientesConvocatoria2").Enabled = True Then
                Form_FormInicial09Indicadores.ComandoPendientesConvocatoria2_Click
            End If
        ElseIf p_lbl.Name = "lblRequiereCursoPteCorreoJefeSeguridad" Then
            
            lblIndicadores_Click
            Set m_Frm = m_Frm.SubFormCentral.Form
            If m_Frm.Controls("ComandoPendientesCorreoJefeSeguridad").Enabled = True Then
                Form_FormInicial09Indicadores.ComandoPendientesCorreoJefeSeguridad_Click
            End If
        ElseIf p_lbl.Name = "lblConfiguracion" Then
            If m_Frm.SubFormCentral.SourceObject <> "FormInicial05Configuracion" Then
                m_Frm.SubFormCentral.SourceObject = "FormInicial05Configuracion"
                '.Left = (m_frm.InsideWidth / 2) - (ctl.Width / 2)
                EstablecerControlesformlateral Me, p_Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
        
        ElseIf p_lbl.Name = "lblIndicadores" Then
            If m_Frm.SubFormCentral.SourceObject <> "FormInicial09Indicadores" Then
                m_Frm.SubFormCentral.SourceObject = "FormInicial09Indicadores"
                '.Left = (m_frm.InsideWidth / 2) - (ctl.Width / 2)
                EstablecerControlesformlateral Me, p_Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
            '.Width = 10258
        ElseIf p_lbl.Name = "lblUsuariosSICA" Then
            If m_Frm.SubFormCentral.SourceObject <> "FormInicial07UsuariosSICA" Then
                m_Frm.SubFormCentral.SourceObject = "FormInicial07UsuariosSICA"
                '.Left = (m_frm.InsideWidth / 2) - (ctl.Width / 2)
                EstablecerControlesformlateral Me, p_Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
        ElseIf p_lbl.Name = "lblUsuariosHistoricos" Then
            If m_Frm.SubFormCentral.SourceObject <> "FormInicial08UsuariosHistoricos" Then
                m_Frm.SubFormCentral.SourceObject = "FormInicial08UsuariosHistoricos"
                '.Left = (m_frm.InsideWidth / 2) - (ctl.Width / 2)
                EstablecerControlesformlateral Me, p_Error
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
        ElseIf p_lbl.Name = "lblSolicitados" Then
            lblIndicadores_Click
            Set m_Frm = m_Frm.SubFormCentral.Form
            If m_Frm.Controls("ComandoEnSolicitud").Enabled = True Then
                Form_FormInicial09Indicadores.ComandoEnSolicitud_Click
            End If
        End If
        
    'End With
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "Al SeleccionarEnlace se ha producido el error: " & Err.Description
    End If
End Function



Public Sub lblUsuariosSICA_Click()
    On Error GoTo errores
    m_Error = ""
    SeleccionarEnlace Me.lblUsuariosSICA, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al hacer click en el control se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub