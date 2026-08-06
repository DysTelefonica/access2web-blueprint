Option Compare Database
Option Explicit

'````````````````````````````````````````````````````````````````````````````
'`                    Code from Tassos Filoxenidis                          `
'````````````````````````````````````````````````````````````````````````````

#If Win64 = 1 Then
    Private Declare PtrSafe Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
#Else
    Private Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
#End If
Private m_sngLeftResizePos As Integer
Private MeWidth As Integer
Private MeWidthHalf As Integer
Private ArbolLeft As Integer
Private SubFormLeft As Integer
Private Sepleft As Integer
Const SepWMin As Integer = 60
Const SepW As Integer = 60.1
Private m_Error As String
Private m_NodoSeleccionado As MSComctlLib.Node

Private Sub Arbol_NodeClick(ByVal Node As Object)
    
    Dim m_IDVideo As String
    
    On Error GoTo errores
    
    If Left(Node.key, 1) = "V" Then
        dato = Split(Node.key, "V")
        m_IDVideo = CStr(CInt(dato(1)))
        If m_ObjVideoActivo Is Nothing Then
            Set m_ObjVideoActivo = Constructor.getVideo(m_IDVideo, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        Else
            If m_ObjVideoActivo.IDVideo <> m_IDVideo Then
                Set m_ObjVideoActivo = Constructor.getVideo(m_IDVideo, m_Error)
                If m_Error <> "" Then
                    Err.Raise 1000
                End If
                Me.FormDetalle.SourceObject = ""
                
            End If
        End If
        
    Else
        Set m_ObjVideoActivo = Nothing
    End If
    If m_ObjVideoActivo Is Nothing Then
        Me.FormDetalle.SourceObject = ""
    Else
        If Me.FormDetalle.SourceObject = "" Then
            Me.FormDetalle.SourceObject = "FormVideosArbolVideo"
            
        End If
        Form_FormVideoDetalle.EstablecerDatos
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar Arbol_NodeClick se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub Form_Load()
   
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar Form_Load se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    
    
    On Error GoTo errores
    
    CargarArbol Me.Arbol.Object, Me.ListaImagenes.Object, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Sep.Left = Me.FormDetalle.Left - Me.Sep.Width
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Private Sub Sep_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        MouseCursor IDC_SIZEWE
        m_sngLeftResizePos = x
        MeWidthHalf = Me.InsideWidth / 2
        MeWidth = Me.InsideWidth - 100
        ArbolLeft = Me.Arbol.Left
        SubFormLeft = Me.FormDetalle.Left
    End If
    
End Sub


Private Sub Sep_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseCursor IDC_SIZEWE
    If Button = 1 Then
        On Error Resume Next
        With Me.Sep
            If .Left > MeWidthHalf Then Exit Sub    '.Left = MeWidthHalf - 1:
            If .Left < SepWMin Then
                .Left = SepW
                Me.Arbol.Width = .Left - ArbolLeft
                Me.FormDetalle.Left = .Left + .Width
                Me.FormDetalle.Width = MeWidth - .Left
               
            Else
                .Move .Left + x - m_sngLeftResizePos
                Me.Arbol.Width = .Left - ArbolLeft
                Me.FormDetalle.Left = .Left + .Width
                Me.FormDetalle.Width = MeWidth - .Left
                
            End If
        End With
    End If
End Sub

Private Sub Sep_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        With Me.Sep
            If .Left > MeWidthHalf Then .Left = MeWidthHalf - 0.1
            Me.Arbol.Width = IIf(Me.Arbol.Width < 60, 60, .Left - ArbolLeft)
            .Left = Me.Arbol.Width + Me.Arbol.Left
            Me.FormDetalle.Left = .Left + .Width
            Me.FormDetalle.Width = MeWidth - .Left
            
        End With
    End If
End Sub

