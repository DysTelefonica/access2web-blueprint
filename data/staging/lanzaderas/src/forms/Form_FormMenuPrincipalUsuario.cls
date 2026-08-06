Option Compare Database
Option Explicit
Private m_Error As String
Private m_ObjMenu As Menu

Private Sub Form_Load()
    
    Dim m_NumeroEtiquetaUltima As String
    Dim ctlIndicador As Control
    Dim ctlMenu As Control
    Dim frm1 As Form
    Dim m_Nombre As String
    
    
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Ajustar Me
    Set m_ObjMenu = getMenu(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_Nombre = getNombreUsuarioConectado()
    
    Me.lblVer.Caption = "Versión " & m_ObjEntorno.VersionAplicacion
    Me.Caption = "Usuario conectado: " & m_Nombre
    Me.EncabezadoDelFormulario.BackColor = m_ObjEntorno.ColorMenuDesactivado
    Me.lblUsuarioConectado.Caption = "Usuario Conectado: " & m_ObjUsuarioConectadoLogin.Nombre & _
                                    " Miembro del equipo de " & m_ObjUsuarioConectadoLogin.Equipo
    If m_EnOficina = EnumSino.Sí Then
    
        Me.lblUsuarioUbicacion.Caption = "Ubicación: " & "En Oficina"
    Else
        Me.lblUsuarioUbicacion.Caption = "Ubicación: " & "Fuera de la Oficina"
    End If
    Me.FormDetalle.SourceObject = ""
    m_ObjMenu.RatonSobreDetalle
    ActivarObjetoMenu
    m_ObjMenu.ActDesActMenu 1, EnumSino.Sí
    MENU1_Click
    Set frm1 = Me.FormDetalle.Form
    If frm1.IND1.Visible = False Then
        frm1.IND1.Visible = True
    End If
    If frm1.IND2.Visible = True Then
        frm1.IND2.Visible = False
    End If
    If frm1.MENU1.BackColor = m_ObjEntorno.ColorMenuDesactivado Then
        frm1.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado
    End If
    If frm1.MENU2.BackColor = m_ObjEntorno.ColorMenuActivado Then
        frm1.MENU2.BackColor = m_ObjEntorno.ColorMenuDesactivado
    End If
    If frm1.FormDetalle.SourceObject <> "FormAplicacionesFueraOficina" Then
        frm1.FormDetalle.SourceObject = "FormAplicacionesFueraOficina"
    End If
'    If IsNumeric(m_ObjUsuarioConectadoLogin.EtiquetaAplicacionesAPulsar) Then
'        ActivarObjetoMenu
'        m_ObjMenu.ActDesActMenu 1, EnumSino.Sí
'        MENU1_Click
'        Set frm1 = Me.FormDetalle.Form
'        m_NumeroEtiquetaUltima = m_ObjUsuarioConectadoLogin.EtiquetaAplicacionesAPulsar
'
'        If m_NumeroEtiquetaUltima = "1" Then
'            If frm1.IND1.Visible = False Then
'                frm1.IND1.Visible = True
'            End If
'            If frm1.IND2.Visible = True Then
'                frm1.IND2.Visible = False
'            End If
'            If frm1.MENU1.BackColor = m_ObjEntorno.ColorMenuDesactivado Then
'                frm1.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado
'            End If
'            If frm1.MENU2.BackColor = m_ObjEntorno.ColorMenuActivado Then
'                frm1.MENU2.BackColor = m_ObjEntorno.ColorMenuDesactivado
'            End If
'            If frm1.FormDetalle.SourceObject <> "FormAplicacionesFueraOficina" Then
'                frm1.FormDetalle.SourceObject = "FormAplicacionesFueraOficina"
'            End If
'        ElseIf m_NumeroEtiquetaUltima = "2" Then
'            If frm1.IND2.Visible = False Then
'                frm1.IND2.Visible = True
'            End If
'            If frm1.IND1.Visible = True Then
'                frm1.IND1.Visible = False
'            End If
'            If frm1.MENU2.BackColor = m_ObjEntorno.ColorMenuDesactivado Then
'                frm1.MENU2.BackColor = m_ObjEntorno.ColorMenuActivado
'            End If
'            If frm1.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado Then
'                frm1.MENU1.BackColor = m_ObjEntorno.ColorMenuDesactivado
'            End If
'            If frm1.FormDetalle.SourceObject <> "FormAplicacionesOficina" Then
'                frm1.FormDetalle.SourceObject = "FormAplicacionesOficina"
'            End If
'
'        End If
'    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Activate se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ImagenLogout_Click()
    On Error GoTo errores
    Pregunta = MsgBox("¿Desea cerrar sesión?", vbExclamation + vbYesNo + vbDefaultButton2, "Cerrar sesión")
    If Pregunta <> vbYes Then
        Exit Sub
    End If
    
    HechoLogout = EnumSino.Sí
    DoCmd.Close acForm, Me.Name
    
    
    DoCmd.OpenForm "FormLogin"
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ImagenLogout_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description & vbNewLine & Err.Source
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ImagenSalir_Click()
    On Error Resume Next
    Pregunta = MsgBox("¿Desea salir de la lanzadera", vbExclamation + vbYesNo + vbDefaultButton2, "Salir de la Lanzadera")
    If Pregunta <> vbYes Then
        Exit Sub
    End If
    
    DoCmd.Quit acQuitSaveAll
End Sub

Private Sub MENU1_Click()
    Dim m_NumeroEtiquetaUltima As String
    Dim m_Form As Form
    On Error Resume Next
    Application.Echo False
    If Me.FormDetalle.SourceObject <> "FormAplicaciones" Then
        Me.FormDetalle.SourceObject = "FormAplicaciones"
    End If
'    If Not m_ObjUsuarioConectadoLogin.UltimaAplicacionAbierta Is Nothing Then
'        m_NumeroEtiquetaUltima = m_ObjUsuarioConectadoLogin.EtiquetaAplicacionesAPulsar
'        If IsNumeric(m_NumeroEtiquetaUltima) Then
'            Set m_Form = Me.FormDetalle.Form
'            If m_NumeroEtiquetaUltima = "1" Then
'                If m_Form.IND1.Visible = False Then
'                    m_Form.IND1.Visible = True
'                End If
'                    If m_Form.IND2.Visible = True Then
'                    m_Form.IND2.Visible = False
'                End If
'                If m_Form.MENU1.BackColor = m_ObjEntorno.ColorMenuDesactivado Then
'                    m_Form.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado
'                End If
'                If m_Form.MENU2.BackColor = m_ObjEntorno.ColorMenuActivado Then
'                    m_Form.MENU2.BackColor = m_ObjEntorno.ColorMenuDesactivado
'                End If
'                Form_FormAplicaciones.MENU1_Click
'            ElseIf m_NumeroEtiquetaUltima = "2" Then
'                If m_Form.IND2.Visible = False Then
'                    m_Form.IND2.Visible = True
'                End If
'                If m_Form.IND1.Visible = True Then
'                    m_Form.IND1.Visible = False
'                End If
'                If m_Form.MENU2.BackColor = m_ObjEntorno.ColorMenuDesactivado Then
'                    m_Form.MENU2.BackColor = m_ObjEntorno.ColorMenuActivado
'                End If
'                If m_Form.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado Then
'                    m_Form.MENU1.BackColor = m_ObjEntorno.ColorMenuDesactivado
'                End If
'                Form_FormAplicaciones.MENU2_Click
'            End If
'        End If
'    End If
    Application.Echo True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al MENU1_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub IND1_Click()
    On Error Resume Next
    Application.Echo False
    MENU1_Click
    Application.Echo True
End Sub
Private Sub imagen1_Click()
    On Error Resume Next
    Application.Echo False
    MENU1_Click
    Application.Echo True
End Sub
Private Sub lblMenu1_Click()
    On Error Resume Next
    Application.Echo False
    MENU1_Click
    Application.Echo True
End Sub
Private Sub MENU2_Click()
    Pregunta = MsgBox("En construcción", vbExclamation, "En construcción")
'
'    If Me.FormDetalle.SourceObject <> "FormVideosUsuarios" Then
'        Me.FormDetalle.SourceObject = "FormVideosUsuarios"
'    End If
End Sub
Private Sub IND2_Click()
    On Error Resume Next
    Application.Echo False
    MENU2_Click
    Application.Echo True
End Sub
Private Sub imagen2_Click()
    On Error Resume Next
    Application.Echo False
    MENU2_Click
    Application.Echo True
End Sub
Private Sub lblMenu2_Click()
    On Error Resume Next
    Application.Echo False
    MENU2_Click
    Application.Echo True
End Sub
Private Sub MENU3_Click()
    On Error Resume Next
    Application.Echo False
    If Me.FormDetalle.SourceObject <> "FormPerfil" Then
        Me.FormDetalle.SourceObject = "FormPerfil"
    End If
    Application.Echo True
End Sub
Private Sub IND3_Click()
    On Error Resume Next
    Application.Echo False
    MENU3_Click
End Sub
Private Sub imagen3_Click()
    On Error Resume Next
    Application.Echo False
    MENU3_Click
    Application.Echo True
End Sub
Private Sub lblMenu3_Click()
    On Error Resume Next
    Application.Echo False
    MENU3_Click
    Application.Echo True
End Sub

Private Sub Desplegar()
    On Error Resume Next
    Application.Echo False
    ActivarObjetoMenu
    m_ObjMenu.MostrarSubMenus
    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
    End If
    Application.Echo True
End Sub

Private Sub Plegar()
    On Error Resume Next
    Application.Echo False
    ActivarObjetoMenu
    m_ObjMenu.OcultarSubMenus
    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
    End If
    Application.Echo True
End Sub

    
Private Sub MENU1_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Application.Echo False
    ActivarObjetoMenu
    m_ObjMenu.ActDesActMenu 1, EnumSino.Sí
    Application.Echo True
End Sub


Private Sub MENU2_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Application.Echo False
    ActivarObjetoMenu
    m_ObjMenu.ActDesActMenu 2, EnumSino.Sí
    Application.Echo True
End Sub

Private Sub MENU3_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Application.Echo False
    ActivarObjetoMenu
    m_ObjMenu.ActDesActMenu 3, EnumSino.Sí
    Application.Echo True
End Sub


Private Sub PANEL1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Application.Echo False
    ActivarObjetoMenu
     m_ObjMenu.RatonSobrePANEL1
    Application.Echo True
    
End Sub

Private Sub Detalle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Application.Echo False
    ActivarObjetoMenu
    If Me.FormDetalle.SourceObject = "" Then
        m_ObjMenu.RatonSobreDetalle
        Application.Echo True
        Exit Sub
    End If
    If Me.FormDetalle.SourceObject = "FormAplicaciones" Then
        If Me.IND1.Visible = False Then
            Me.IND1.Visible = True
        End If
        If Me.IND2.Visible = True Then
            Me.IND2.Visible = False
        End If
        If Me.IND3.Visible = True Then
            Me.IND3.Visible = False
        End If
        
        If Me.MENU1.BackColor = m_ObjEntorno.ColorMenuDesactivado Then
            Me.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado
        End If
        If Me.MENU2.BackColor = m_ObjEntorno.ColorMenuActivado Then
            Me.MENU2.BackColor = m_ObjEntorno.ColorMenuDesactivado
        End If
        If Me.MENU3.BackColor = m_ObjEntorno.ColorMenuActivado Then
            Me.MENU3.BackColor = m_ObjEntorno.ColorMenuDesactivado
        End If
        
    End If
    Application.Echo True
End Sub


Private Sub ActivarObjetoMenu()
    On Error Resume Next
    Application.Echo False
    If m_ObjMenu Is Nothing Then
        Set m_ObjMenu = getMenu(Me)
        
    End If
    Application.Echo True
End Sub
