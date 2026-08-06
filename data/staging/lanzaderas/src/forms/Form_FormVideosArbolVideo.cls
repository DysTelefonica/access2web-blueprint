Option Compare Database
Option Explicit
Private m_Error As String

Private Sub EncabezadoDelFormulario_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.IND1.Visible = True Then
        Me.IND1.Visible = False
    End If
    
    
    If Me.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado Then
        Me.MENU1.BackColor = m_ObjEntorno.ColorMenuDesactivado
    End If
    
    
End Sub



Public Sub MENU1_Click()
    If Me.IND1.Visible = False Then
        Me.IND1.Visible = True
    End If
    
    If Me.FormDetalle.SourceObject <> "FormVideoDetalle" Then
        Me.FormDetalle.SourceObject = "FormVideoDetalle"
       
    End If
End Sub
Private Sub lblMenu1_Click()
    MENU1_Click
End Sub



Private Sub MENU1_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Me.IND1.Visible = False Then
        Me.IND1.Visible = True
    End If
    
    
    If Me.MENU1.BackColor = m_ObjEntorno.ColorMenuDesactivado Then
        Me.MENU1.BackColor = m_ObjEntorno.ColorMenuActivado
    End If
    
    
End Sub
Private Sub lblMenu1_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
     MENU1_Mousemove Button, Shift, x, y
End Sub

Private Sub IND1_Mousemove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
     MENU1_Mousemove Button, Shift, x, y
End Sub

