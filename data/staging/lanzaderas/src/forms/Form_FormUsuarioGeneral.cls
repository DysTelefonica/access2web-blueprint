Option Compare Database
Option Explicit
Private WithEvents m_Form As Form_FormUsuariosGestion
Private m_Error As String


Private Sub EncabezadoDelFormulario_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.IND1.Visible = True Then
        Me.IND1.Visible = False
    End If
    If Me.IND2.Visible = True Then
        Me.IND2.Visible = False
    End If
    If Me.IND3.Visible = True Then
        Me.IND3.Visible = False
    End If
    If Me.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado Then
        Me.MENU1.BackColor = m_ObjEntorno.ColorMenuDesactivado
    End If
    If Me.MENU2.BackColor = m_ObjEntorno.ColorMenuActivado Then
        Me.MENU2.BackColor = m_ObjEntorno.ColorMenuDesactivado
    End If
    If Me.MENU3.BackColor = m_ObjEntorno.ColorMenuActivado Then
        Me.MENU3.BackColor = m_ObjEntorno.ColorMenuDesactivado
    End If
End Sub

Private Sub Form_Load()

    On Error GoTo errores
    
    m_Error = ""
    
    Me.IND1.Visible = True
    
    Me.IND2.Visible = False
    
    Me.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado
    Me.MENU2.BackColor = m_ObjEntorno.ColorMenuDesactivado
    Me.MENU3.BackColor = m_ObjEntorno.ColorMenuDesactivado
    
    
    Me.FormDetalle.SourceObject = ""
    Set m_ObjUsuarioActivo = Nothing
    Me.lblUsuarioActivo.Caption = "Usuario Activo: "
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar Form_Load se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_Form_QuitarUsuarioSeleccionado()
    Me.lblUsuarioActivo.Caption = "Usuario Activo: "
End Sub

Private Sub m_Form_UsuarioSeleccionado(usuario As usuario)
    Me.lblUsuarioActivo.Caption = "Usuario Activo: " & usuario.Nombre
End Sub

Public Sub MENU1_Click()
    If Me.FormDetalle.SourceObject <> "FormUsuariosGestion" Then
        Me.FormDetalle.SourceObject = "FormUsuariosGestion"
    End If
    Set m_ObjUsuarioActivo = Nothing
    Set m_Form = Me.FormDetalle.Form
End Sub

Private Sub lblMenu1_Click()
    MENU1_Click
End Sub
Private Sub MENU2_Click()
    If Me.FormDetalle.SourceObject <> "FormUsuarioDatosGenerales" Then
        Me.FormDetalle.SourceObject = "FormUsuarioDatosGenerales"
    End If
End Sub

Private Sub lblMenu2_Click()
    MENU2_Click
End Sub
Private Sub MENU3_Click()
    If m_ObjUsuarioActivo Is Nothing Then
        Pregunta = MsgBox("No se ha establecido un usuario activo", vbCritical, "Sin usuario activo")
        Exit Sub
    End If
    If Me.FormDetalle.SourceObject <> "FormUsuarioPerfilAplicaciones" Then
        Me.FormDetalle.SourceObject = "FormUsuarioPerfilAplicaciones"
    End If
End Sub

Private Sub lblMenu3_Click()
    MENU3_Click
End Sub

Private Sub MENU1_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
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
End Sub
Private Sub lblMenu1_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    MENU1_Mousemove Button, Shift, x, y
    
End Sub

Private Sub MENU2_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Me.IND2.Visible = False Then
        Me.IND2.Visible = True
    End If
    If Me.IND1.Visible = True Then
        Me.IND1.Visible = False
    End If
    If Me.IND3.Visible = True Then
        Me.IND3.Visible = False
    End If
    If Me.MENU2.BackColor = m_ObjEntorno.ColorMenuDesactivado Then
        Me.MENU2.BackColor = m_ObjEntorno.ColorMenuActivado
    End If
    If Me.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado Then
        Me.MENU1.BackColor = m_ObjEntorno.ColorMenuDesactivado
    End If
    If Me.MENU3.BackColor = m_ObjEntorno.ColorMenuActivado Then
        Me.MENU3.BackColor = m_ObjEntorno.ColorMenuDesactivado
    End If
End Sub

Private Sub lblMenu2_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    MENU2_Mousemove Button, Shift, x, y
End Sub

Private Sub MENU3_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Me.IND3.Visible = False Then
        Me.IND3.Visible = True
    End If
    If Me.IND1.Visible = True Then
        Me.IND1.Visible = False
    End If
    If Me.IND2.Visible = True Then
        Me.IND2.Visible = False
    End If
    If Me.MENU3.BackColor = m_ObjEntorno.ColorMenuDesactivado Then
        Me.MENU3.BackColor = m_ObjEntorno.ColorMenuActivado
    End If
    If Me.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado Then
        Me.MENU1.BackColor = m_ObjEntorno.ColorMenuDesactivado
    End If
    If Me.MENU2.BackColor = m_ObjEntorno.ColorMenuActivado Then
        Me.MENU2.BackColor = m_ObjEntorno.ColorMenuDesactivado
    End If
End Sub

Private Sub lblMenu3_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    MENU3_Mousemove Button, Shift, x, y
End Sub


