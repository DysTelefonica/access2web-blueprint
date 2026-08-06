 Option Compare Database
Option Explicit

Public Event Adjuntado()
Private m_URLLocal As String
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
    
    Dim m_SolOp As SolicitudOperaciones

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If m_URLLocal = "" Then
        m_Error = "Debe seleccionar el archivo a registrar"
        Err.Raise 1000
    End If
    Set m_SolOp = New SolicitudOperaciones
    With m_SolOp
        Set .solicitud = m_ObjSolicitudActiva
        .RegistroExcelDevuelto m_URLLocal, Nz(Me.FECHA, ""), m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        RaiseEvent Adjuntado
    End With
    PintarIndicadores
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    If m_ObjSolicitudActiva Is Nothing Then
        p_Error = "No hay datos previos de la solicitud"
        Err.Raise 1000
    End If
    With m_ObjSolicitudActiva
        If .URLAdjuntoCalculado <> "" Then
            Me.ImagenURL.Visible = True
            Me.URL = "..\" & fso.GetFileName(.URLAdjuntoCalculado)
        Else
            Me.ImagenURL.Visible = False
            Me.URL = Null
        End If
        If Not .SolicitudFechas Is Nothing Then
            If .SolicitudFechas.FechaRecepcionExcel = "" Then
                Me.FECHA = Date
            Else
                Me.FECHA = .SolicitudFechas.FechaRecepcionExcel
            End If
        End If
    End With
    Me.Caption = m_TituloFormulario
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function


Private Sub ComandoSeleccionar_Click()
      
    On Error GoTo errores
    m_Error = ""
    
    m_URLLocal = Seleccionar(True, "Seleccione el archivo", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_URLLocal = "" Then
        Me.URL = Null
    Else
        Me.URL = "..\" & fso.GetFileName(m_URLLocal)
    End If
    
    
    
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoSeleccionar_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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
    
    Me.AllowEdits = False
        
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
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

Private Sub ImagenURL_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If fso.FileExists(m_URLLocal) Then
        Ejecutar Me.hwnd, "open", m_URLLocal, "", "", 1
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Me.AllowEdits = True
        Exit Sub
    End If
    AbrirEnLocal m_ObjSolicitudActiva.URLAdjuntoCalculado, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ImagenURL_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    
End Sub