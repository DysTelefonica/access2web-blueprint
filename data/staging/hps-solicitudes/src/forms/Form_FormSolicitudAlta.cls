 Option Compare Database
Option Explicit


Private m_Error As String

Private Sub cmdSalir_Click()
     On Error Resume Next
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ComandoAyuda_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
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



Private Sub ComandoRegistrar_Click()
    Dim m_Tipo As String
    Dim m_NombreForm As String
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    m_Tipo = Nz(Me.TIPO, "")
    If m_Tipo = "" Then
        m_Error = "Seleccione un tipo"
        Err.Raise 1000
    End If
    Set m_ObjSolicitudActiva = New solicitud
    Set m_ObjSolicitudAlInicio = Nothing
    
    If m_Tipo = "Alta" Or m_Tipo = "Traspaso" Then
        m_NombreForm = "FormSolicitudAltaTraspaso"
        m_ObjSolicitudActiva.TIPO = m_Tipo
    Else
        If m_EnOficina <> EnumSiNo.Sí Then
            m_Error = "Para poder hacer una renovación se ha de estar en la oficina. No se puede acceder a los datos de HPS"
            Err.Raise 1000
        End If
        m_NombreForm = "FormUsuariosHPSBusqueda"
    End If
    If FormularioAbierto(m_NombreForm) Then
        DoCmd.Close acForm, m_NombreForm, acSaveNo
    End If
    
    DoCmd.OpenForm m_NombreForm, OpenArgs:=m_Tipo
    If m_NombreForm = "FormUsuariosHPSBusqueda" Then
        AvanceCerrar
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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
    Me.ComandoRegistrar.Enabled = False
    Me.AllowEdits = False
    Me.Caption = m_TituloFormulario
    If EsTecnico <> EnumSiNo.Sí Then
        Me.ComandoRegistrar.Enabled = True
        Me.AllowEdits = True
    End If
    
    
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub