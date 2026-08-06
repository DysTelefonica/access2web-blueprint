Option Compare Database
Option Explicit
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
    Me.IND1.Visible = False
    Me.IND2.Visible = False
    Me.IND3.Visible = False
    Me.MENU1.BackColor = m_ObjEntorno.ColorMenuDesactivado
    Me.MENU2.BackColor = m_ObjEntorno.ColorMenuDesactivado
    Me.MENU3.BackColor = m_ObjEntorno.ColorMenuDesactivado
    
    Me.FormDetalle.SourceObject = ""
    
    
End Sub
Public Sub MENU1_Click()
    On Error Resume Next
    Me.AllowEdits = False
    If Me.FormDetalle.SourceObject <> "FormAplicacionesFueraOficina" Then
        Me.FormDetalle.SourceObject = "FormAplicacionesFueraOficina"
    End If
    Me.AllowEdits = True
End Sub
Private Sub lblMenu1_Click()
    
    MENU1_Click
End Sub
Public Sub MENU2_Click()
    On Error Resume Next
    Me.AllowEdits = True
    If Me.FormDetalle.SourceObject <> "FormAplicacionesOficina" Then
        Me.FormDetalle.SourceObject = "FormAplicacionesOficina"
    End If
    Me.AllowEdits = False
End Sub
Private Sub lblMenu2_Click()
    On Error Resume Next
    MENU2_Click
End Sub
Public Sub MENU3_Click()
    On Error Resume Next
    Me.AllowEdits = True
    If Me.FormDetalle.SourceObject <> "FormAplicacionesPrueba" Then
        Me.FormDetalle.SourceObject = "FormAplicacionesPrueba"
    End If
    Me.AllowEdits = False
End Sub
Private Sub lblMenu3_Click()
    On Error Resume Next
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

Private Sub IND1_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
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
