Option Compare Database
Option Explicit



Private m_Error As String


Private Sub Form_Load()
  
    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    Ajustar Me
    'DoCmd.Maximize
    
   
    If Not fso.FileExists(m_URLHTMLActivo) Then
        DoCmd.Close acForm, Me.Name
        DoCmd.Hourglass False
        Exit Sub
    End If
    
    'Me.Navegador.Hyperlink.CreateNewDocument
    'Me.Navegador.Hyperlink.Address = "https://msaccess/C:\Users\adm.DEFENSA\AppData\Roaming\Aplicaciones DYSN\Gestion Riesgos Prueba\HTML1.html"
    Me.Navegador.Navigate m_URLHTMLActivo
    'MiWeb.Document
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub




Private Sub Form_Resize()
'    Debug.Print "Ancho: " & Me.InsideWidth
'    Debug.Print "Alto: " & Me.InsideHeight
    Me.Navegador.Width = Me.InsideWidth
    Me.Navegador.Height = Me.InsideHeight
End Sub