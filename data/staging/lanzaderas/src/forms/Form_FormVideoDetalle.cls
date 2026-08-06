Option Compare Database
Option Explicit
Private m_Error As String


Private Sub ComandoVerVideo_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_ObjVideoActivo Is Nothing Then
        m_Error = "Ningún video seleccionado"
        Err.Raise 1000
    End If
    If Not fso.FileExists(m_ObjVideoActivo.url) Then
        m_Error = "En estos momentos no es accesible el video"
        Err.Raise 1000
    End If
    Ejecutar Me.hWnd, "open", m_ObjVideoActivo.url, "", "", 1
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerVideo_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo errores
    
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    
    On Error GoTo errores
    
    p_Error = ""
    If m_ObjVideoActivo Is Nothing Then
        Me.lblTitulo.Caption = "TÍTULO  DEL VIDEO"
        Me.lblExplicacion = "EXPLICACIÓN DEL APARTADO"
    Else
        With m_ObjVideoActivo
            Me.lblTitulo.Caption = .Titulo
            If .Descripcion <> "" Then
                Me.lblExplicacion.Caption = .Descripcion
            Else
                Me.lblExplicacion.Caption = "Sin descripción registrada"
            End If
            
        End With
    End If
    
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function