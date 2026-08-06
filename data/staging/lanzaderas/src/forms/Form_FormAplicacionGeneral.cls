Option Compare Database
Option Explicit
Private WithEvents m_Form As Form_FormAplicacionesGestion
Private m_Error As String


Private Sub EncabezadoDelFormulario_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.IND1.Visible = True Then
        Me.IND1.Visible = False
    End If
    If Me.IND2.Visible = True Then
        Me.IND2.Visible = False
    End If
    
    If Me.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado Then
        Me.MENU1.BackColor = m_ObjEntorno.ColorMenuDesactivado
    End If
    If Me.MENU2.BackColor = m_ObjEntorno.ColorMenuActivado Then
        Me.MENU2.BackColor = m_ObjEntorno.ColorMenuDesactivado
    End If
    
End Sub

Private Sub Form_Load()

    On Error GoTo errores
    
    m_Error = ""
    
    Me.IND1.Visible = True
    
    Me.IND2.Visible = False
    
    Me.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado
    Me.MENU2.BackColor = m_ObjEntorno.ColorMenuDesactivado
    
    
    
    Me.FormDetalle.SourceObject = ""
    Set m_ObjUsuarioActivo = Nothing
    Me.lblAplicacionActiva.Caption = "Aplicación Activa: "
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar Form_Load se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub m_Form_QuitarAplicacionSeleccionada()
    Me.lblAplicacionActiva.Caption = "Aplicación Activa: "
End Sub

Private Sub m_Form_AplicacionSeleccionada(m_objAplicacion As Aplicacion)
    Me.lblAplicacionActiva.Caption = "Aplicación Activa: " & m_objAplicacion.NombreAplicacion
End Sub

Public Sub MENU1_Click()
    If Me.FormDetalle.SourceObject <> "FormAplicacionesGestion" Then
        Me.FormDetalle.SourceObject = "FormAplicacionesGestion"
    End If
    Set m_ObjAplicacionActiva = Nothing
    Set m_Form = Me.FormDetalle.Form
End Sub

Private Sub lblMenu1_Click()
    MENU1_Click
End Sub
Private Sub MENU2_Click()
    If Me.FormDetalle.SourceObject <> "FormAplicacionDatosGenerales" Then
        Me.FormDetalle.SourceObject = "FormAplicacionDatosGenerales"
    End If
End Sub

Private Sub lblMenu2_Click()
    MENU2_Click
End Sub


Private Sub MENU1_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Me.IND1.Visible = False Then
        Me.IND1.Visible = True
    End If
    If Me.IND2.Visible = True Then
        Me.IND2.Visible = False
    End If
    
    If Me.MENU1.BackColor = m_ObjEntorno.ColorMenuDesactivado Then
        Me.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado
    End If
    If Me.MENU2.BackColor = m_ObjEntorno.ColorMenuActivado Then
        Me.MENU2.BackColor = m_ObjEntorno.ColorMenuDesactivado
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
    
    If Me.MENU2.BackColor = m_ObjEntorno.ColorMenuDesactivado Then
        Me.MENU2.BackColor = m_ObjEntorno.ColorMenuActivado
    End If
    If Me.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado Then
        Me.MENU1.BackColor = m_ObjEntorno.ColorMenuDesactivado
    End If
    
End Sub

Private Sub lblMenu2_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    MENU2_Mousemove Button, Shift, x, y
End Sub

