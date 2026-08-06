Option Compare Database
Option Explicit
Private m_Error As String
Private m_ObjMenu As Menu


Private Sub Desplegar()
    ActivarObjetoMenu
    m_ObjMenu.MostrarSubMenus
    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
    End If
End Sub

Private Sub Form_Load()
    Dim m_NumeroEtiquetaUltima As String
    Dim ctlIndicador As Control
    Dim ctlMenu As Control
    Dim m_Nombre As String
    
    Dim frm1 As Form
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
  
    Ajustar Me

    Set m_ObjMenu = getMenu(Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_Nombre = getNombreUsuarioConectado()
    
    
    Me.lblVer.Caption = "Versión " & m_ObjEntorno.VersionAplicacion
    Me.Caption = "Usuario conectado: " & m_Nombre
    Me.EncabezadoDelFormulario.BackColor = m_ObjEntorno.ColorMenuDesactivado
    Me.lblUsuarioConectado.Caption = "Usuario Conectado: " & m_Nombre & _
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

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
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
    
    
    DoCmd.Quit acQuitSaveNone
End Sub

Private Sub lblSubMenu3_Click()
    SUBMENU3_Click
End Sub

Private Sub MENU1_Click()
    Dim m_NumeroEtiquetaUltima As String
    Dim m_Form As Form
    On Error GoTo errores
    
    If Me.FormDetalle.SourceObject <> "FormAplicaciones" Then
        Me.FormDetalle.SourceObject = "FormAplicaciones"
    End If
    If Not m_ObjUsuarioConectadoLogin.UltimaAplicacionAbierta Is Nothing Then
        m_NumeroEtiquetaUltima = m_ObjUsuarioConectadoLogin.EtiquetaAplicacionesAPulsar
        If IsNumeric(m_NumeroEtiquetaUltima) Then
            Set m_Form = Me.FormDetalle.Form
            If m_NumeroEtiquetaUltima = "1" Then
                If m_Form.IND1.Visible = False Then
                    m_Form.IND1.Visible = True
                End If
                If m_Form.IND2.Visible = True Then
                    m_Form.IND2.Visible = False
                End If
                If m_Form.MENU1.BackColor = m_ObjEntorno.ColorMenuDesactivado Then
                    m_Form.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado
                End If
                If m_Form.MENU2.BackColor = m_ObjEntorno.ColorMenuActivado Then
                    m_Form.MENU2.BackColor = m_ObjEntorno.ColorMenuDesactivado
                End If
                Form_FormAplicaciones.MENU1_Click
            ElseIf m_NumeroEtiquetaUltima = "2" Then
                If m_Form.IND2.Visible = False Then
                    m_Form.IND2.Visible = True
                End If
                If m_Form.IND1.Visible = True Then
                    m_Form.IND1.Visible = False
                End If
                If m_Form.MENU2.BackColor = m_ObjEntorno.ColorMenuDesactivado Then
                    m_Form.MENU2.BackColor = m_ObjEntorno.ColorMenuActivado
                End If
                If m_Form.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado Then
                    m_Form.MENU1.BackColor = m_ObjEntorno.ColorMenuDesactivado
                End If
                Form_FormAplicaciones.MENU2_Click
            
            End If
        End If
    End If
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
    MENU1_Click
End Sub
Private Sub imagen1_Click()
    MENU1_Click
End Sub
Private Sub lblMenu1_Click()
    MENU1_Click
End Sub
Private Sub MENU2_Click()
    Pregunta = MsgBox("En construcción", vbExclamation, "En construcción")
'
'    If Me.FormDetalle.SourceObject <> "FormVideosUsuarios" Then
'        Me.FormDetalle.SourceObject = "FormVideosUsuarios"
'    End If
End Sub
Private Sub IND2_Click()
    MENU2_Click
End Sub
Private Sub imagen2_Click()
    MENU2_Click
End Sub
Private Sub lblMenu2_Click()
    MENU2_Click
End Sub

Private Sub MENU3_Click()
    If Me.SUBMENU1.Visible = True Then
        Plegar
    Else
        Desplegar
    End If
End Sub
Private Sub IND3_Click()
    MENU3_Click
End Sub
Private Sub imagen3_Click()
    MENU3_Click
End Sub
Private Sub lblMenu3_Click()
    MENU3_Click
End Sub

Private Sub SUBMENU1_Click()
    Dim frm As Form
    Set m_ObjUsuarioActivo = Nothing
    If Me.FormDetalle.SourceObject <> "FormUsuarioGeneral" Then
        Me.FormDetalle.SourceObject = "FormUsuarioGeneral"
    End If
    Set frm = Me.FormDetalle.Form
    With frm
        If .Controls("IND1").Visible <> True Then
            .Controls("IND1").Visible = True
        End If
        If .Controls("MENU1").BackColor <> m_ObjEntorno.ColorMenuActivado Then
            .Controls("MENU1").BackColor = m_ObjEntorno.ColorMenuActivado
        End If
    End With
    Form_FormUsuarioGeneral.MENU1_Click
End Sub
Private Sub lblSubMenu1_Click()
    SUBMENU1_Click
End Sub
Private Sub SUBMENU2_Click()
    Dim frm As Form
    Set m_ObjUsuarioActivo = Nothing
    If Me.FormDetalle.SourceObject <> "FormAplicacionGeneral" Then
        Me.FormDetalle.SourceObject = "FormAplicacionGeneral"
    End If
    Set frm = Me.FormDetalle.Form
    With frm
        If .Controls("IND1").Visible <> True Then
            .Controls("IND1").Visible = True
        End If
        If .Controls("MENU1").BackColor <> m_ObjEntorno.ColorMenuActivado Then
            .Controls("MENU1").BackColor = m_ObjEntorno.ColorMenuActivado
        End If
    End With
    Form_FormAplicacionGeneral.MENU1_Click
End Sub
Private Sub lblSubMenu2_Click()
    SUBMENU2_Click
End Sub
Private Sub Plegar()
    ActivarObjetoMenu
    m_ObjMenu.OcultarSubMenus
    If Me.FormDetalle.SourceObject <> "" Then
        Me.FormDetalle.SourceObject = ""
    End If
End Sub
Private Sub SUBMENU3_Click()
    
    Set m_ObjUsuarioActivo = Nothing
    If Me.FormDetalle.SourceObject <> "FormVideosGestion" Then
        Me.FormDetalle.SourceObject = "FormVideosGestion"
    End If
    
End Sub
Private Sub lblSubMenu13_Click()
    SUBMENU3_Click
End Sub

Private Sub MENU1_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ActivarObjetoMenu
    m_ObjMenu.ActDesActMenu 1, EnumSino.Sí
End Sub
Private Sub MENU2_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ActivarObjetoMenu
    m_ObjMenu.ActDesActMenu 2, EnumSino.Sí
End Sub
Private Sub MENU3_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ActivarObjetoMenu
    m_ObjMenu.ActDesActMenu 3, EnumSino.Sí
End Sub

Private Sub SUBMENU1_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ActivarObjetoMenu
    m_ObjMenu.ActivarSubMenu 1
End Sub
Private Sub SUBMENU2_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ActivarObjetoMenu
    m_ObjMenu.ActivarSubMenu 2
End Sub
Private Sub SUBMENU3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ActivarObjetoMenu
    m_ObjMenu.ActivarSubMenu 3
End Sub
Private Sub PANEL1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ActivarObjetoMenu
     m_ObjMenu.RatonSobrePANEL1
End Sub

Private Sub Detalle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ActivarObjetoMenu
    m_ObjMenu.RatonSobreDetalle
End Sub

Private Sub ActivarObjetoMenu()
    If m_ObjMenu Is Nothing Then
        Set m_ObjMenu = getMenu(Me)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
End Sub

