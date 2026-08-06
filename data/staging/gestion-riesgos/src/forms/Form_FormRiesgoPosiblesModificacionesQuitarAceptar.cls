Option Compare Database
Option Explicit
Private m_RiesgoAlInicio As riesgo
Private m_Estado As EnumRiesgoEstado
Public Event AceptacionQuitada(m_Mitigacion As String)

Private m_Error As String

Private Sub ComandoGrabar_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""

    m_ObjRiesgoActivo.AceptacionRegistrarQuitar Me.ListaMitigacion.Column(0), Date, m_Estado, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
        
   
    RaiseEvent AceptacionQuitada(Me.ListaMitigacion.Column(0))
    DoCmd.Close acForm, Me.Name, acSaveNo
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGrabar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub



Private Sub ComandoSalir_Click()
    On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub



Private Sub Form_Load()
    
    
    On Error GoTo errores
    
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    AjustarTamaño Me

    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function SeleccionarValorEnLIsta(p_Valor As String, Optional ByRef p_Error As String) As String
    
    Dim lst As ListBox
    Dim i As Long
    On Error GoTo errores
    
    If p_Valor = "" Then
        Exit Function
    End If
    Set lst = Me.ListaMitigacion
    For i = 1 To lst.ListCount - 1
        If Nz(lst.Column(0, i), "") = p_Valor Then
            lst.Selected(i) = True
            ListaMitigacion_Click
            Exit For
        End If
    Next
    
   
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SeleccionarValorEnLIsta ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.description
    End If
    
End Function

Private Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    Dim lst As ListBox
    Dim m_Mitigacion As Variant
    Dim m_Titulo As String
    On Error GoTo errores
    
    If m_RiesgoAlInicio Is Nothing Then
        p_Error = "No hay datos del Riesgo"
        Err.Raise 1000
    End If
    Me.AllowEdits = True
    Set lst = Me.ListaMitigacion
    lst.RowSource = "Mitigación;Descripción"
    
    For Each m_Mitigacion In m_ObjEntorno.ColMitigacionValores.keys
        lst.AddItem m_Mitigacion & ";" & Replace(m_ObjEntorno.ColMitigacionValores(m_Mitigacion), ";", ":")
    Next
    SeleccionarValorEnLIsta m_ObjRiesgoActivo.Mitigacion, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Estado = m_ObjRiesgoActivo.EstadoEnum
    Me.ComandoGrabar.Enabled = True
    Me.lblTitulo1.Caption = m_Titulo
    
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Establecerdatos ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Private Sub ListaMitigacion_Click()
    
    On Error GoTo errores
    
    Me.DescripcionMitigacion = Nz(Me.ListaMitigacion.Column(1), "")
    
        
        
    
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ListaMitigacion_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub
