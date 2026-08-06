Option Compare Database
Option Explicit


Private m_Error As String

Private Sub cmdSalir_Click()
    
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function EstablecerLista(Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Nombre As String
    
    On Error GoTo errores
    Set lst = Me.ListaPlantillas
    lst.RowSource = ""
    Set m_Col = m_ObjEntorno.ColPlantillas
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    For Each m_ID In m_Col
        m_Nombre = m_Col(m_ID)
        lst.AddItem m_ID & ";" & m_Nombre
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerLista ha devuelto el error: " & Err.Description
    End If
End Function



Private Sub ComandoAyuda_Click()
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    AbrirAyuda m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoAyuda_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoVerPlantilla_Click()
    
    Dim m_URLDirectorioPlantillas As String
    Dim m_NombrePlantilla As String
    Dim m_URLPlantillaRemota As String
    Dim m_URLPlantillaLocal As String
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Not IsNumeric(Me.ListaPlantillas.Column(0)) Then
        m_Error = "Seleccione un elemento de la lista"
        Err.Raise 1000
    End If
    m_NombrePlantilla = m_ObjEntorno.ColPlantillas(CStr(Me.ListaPlantillas.Column(0)))
    If m_NombrePlantilla = "" Then
        m_Error = "No se conoce el nombre de la plantilla que se ha seleccionado"
        Err.Raise 1000
    End If
    m_URLDirectorioPlantillas = m_ObjEntorno.URLDirectorioPlantillasHTMLVersionEstablecida
    m_URLPlantillaRemota = m_URLDirectorioPlantillas & m_NombrePlantilla
    If Not fso.FileExists(m_URLPlantillaRemota) Then
        m_Error = "No se encuetra la URL de la plantilla seleccionada"
        Err.Raise 1000
    End If
    m_URLPlantillaLocal = m_ObjEntorno.URLDirectorioLocal & fso.GetFileName(m_URLPlantillaRemota)
    fso.CopyFile m_URLPlantillaRemota, m_URLPlantillaLocal, True
    Ejecutar 1, "open", m_URLPlantillaLocal, "", "", Me.hwnd
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerHTMLEnvioExcel_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Ajustar Me
    
    Me.Caption = Me.lblTitulo1.Caption
    
    Me.ComandoVerPlantilla.Enabled = True
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
   
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ListaPlantillas_Click()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Me.ComandoVerPlantilla.Enabled = False
    If Nz(Me.ListaPlantillas.Column(0), "") <> "" Then
        Me.ComandoVerPlantilla.Enabled = True
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaPlantillas_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ListaPlantillas_DblClick(Cancel As Integer)
    
    On Error Resume Next
    If Me.ComandoVerPlantilla.Enabled = True Then
        ComandoVerPlantilla_Click
    End If
End Sub