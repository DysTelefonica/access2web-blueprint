Option Compare Database
Option Explicit


Private m_Error As String

Private Sub cmdSalir_Click()
    
    
    On Error GoTo errores
    m_Error = ""
    pregunta = MsgBox("¿Desea realmente salir de la aplicación", vbExclamation + vbYesNo, "Salir de la aplicación")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
   
   
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    DoCmd.Quit acQuitSaveAll
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al CmdSalir_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Quit acQuitSaveAll
End Sub

Private Sub ComandoAlta_Click()
    Dim parsed As Object
    Dim payload As Object

    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Set parsed = JsonConverter.ParseJson(mod0BDOpcionesHelper.Opciones_OpenFormAction("alta", m_Error))
    If m_Error <> "" Then Err.Raise 1000
    Set payload = parsed("payload")
    DoCmd.OpenForm CStr(payload("formName")), OpenArgs:=CStr(payload("openArgs"))
    
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "Al ComandoAlta_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

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

Private Sub ComandoBuscadorTecnico_Click()
    Dim parsed As Object
    Dim payload As Object
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set parsed = JsonConverter.ParseJson(mod0BDOpcionesHelper.Opciones_OpenFormAction("buscador-tecnico", m_Error))
    If m_Error <> "" Then Err.Raise 1000
    Set payload = parsed("payload")
    If CBool(payload("closeIfOpen")) Then
        If FormularioAbierto(CStr(payload("formName"))) Then
            DoCmd.Close acForm, CStr(payload("formName")), acSaveNo
        End If
    End If
    DoCmd.OpenForm CStr(payload("formName"))
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoBuscadorTecnico_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoGestionExpedientes_Click()
    Dim m_NombreForm As String
    Dim parsed As Object
    Dim payload As Object
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Set parsed = JsonConverter.ParseJson(mod0BDOpcionesHelper.Opciones_GestionExpedientesTarget(EsTecnico = EnumSiNo.Sí, m_Error))
    If m_Error <> "" Then Err.Raise 1000
    Set payload = parsed("payload")
    m_NombreForm = CStr(payload("formName"))
    If FormularioAbierto(m_NombreForm) Then
        DoCmd.Close acForm, m_NombreForm, acSaveNo
    End If

    MostrarPopupProgreso "Cargando...", "Preparando gestión de expedientes..."

    DoCmd.OpenForm m_NombreForm

    ' Rellenar la barra antes de cerrar (mínimo visible)
    Dim k As Long
    For k = 0 To Forms("frmBusy").lblProgresoFondo.Width Step 30
        Forms("frmBusy").lblProgresoBarra.Width = k
        DoEvents
    Next k

    DoCmd.Close acForm, "frmBusy"
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Close acForm, "frmBusy", acSaveNo
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGestionExpedientes_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub lblGestionBatchE2E_Click()
    Dim parsed As Object
    Dim payload As Object

    On Error GoTo errores

    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""

    Set parsed = JsonConverter.ParseJson(mod0BDOpcionesHelper.Opciones_OpenFormAction("batch-e2e", m_Error))
    If m_Error <> "" Then Err.Raise 1000
    Set payload = parsed("payload")
    If CBool(payload("closeIfOpen")) Then
        If FormularioAbierto(CStr(payload("formName"))) Then
            DoCmd.Close acForm, CStr(payload("formName")), acSaveNo
        End If
    End If
    DoCmd.OpenForm CStr(payload("formName"))

    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        m_Error = "Al lblGestionBatchE2E_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoGestorEntidades_Click()
    Dim parsed As Object
    Dim payload As Object

    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set parsed = JsonConverter.ParseJson(mod0BDOpcionesHelper.Opciones_OpenFormAction("gestor-entidades", m_Error))
    If m_Error <> "" Then Err.Raise 1000
    Set payload = parsed("payload")
    If CBool(payload("closeIfOpen")) Then
        If FormularioAbierto(CStr(payload("formName"))) Then
            DoCmd.Close acForm, CStr(payload("formName")), acSaveNo
        End If
    End If
    DoCmd.OpenForm CStr(payload("formName"))
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGestorEntidades_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub Form_Open(Cancel As Integer)
    Dim parsed As Object
    Dim payload As Object
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Ajustar Me
    
    If m_ObjEntorno Is Nothing Then
        EVE m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    PintarTareas EnumSiNo.Sí, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Application.TempVars("EnDesarrollo") <> "Sí" Then
        DoCmd.ShowToolbar "Ribbon", acToolbarNo
    Else
        DoCmd.ShowToolbar "Ribbon", acToolbarYes
    End If
    
    Set parsed = JsonConverter.ParseJson(mod0BDOpcionesHelper.Opciones_FormOpenState(EsAdministrador = EnumSiNo.Sí, m_ObjEntorno.VersionAplicacion, m_Error))
    If m_Error <> "" Then Err.Raise 1000
    Set payload = parsed("payload")

    Me.ComandoAlta.Enabled = CBool(payload("altaEnabled"))
    Me.Caption = m_TituloFormulario
    Me.lblVersion.Caption = CStr(payload("versionCaption"))
    Me.ComandoGestionExpedientes.Enabled = CBool(payload("gestionEnabled"))
    Me.ComandoGestorEntidades.Enabled = CBool(payload("gestorEntidadesEnabled"))
    Me.ComandoBuscadorTecnico.Visible = CBool(payload("buscadorTecnicoVisible"))
    AvanceCerrar
    Me.lblEstado.Visible = CBool(payload("estadoVisible"))
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
    If Application.TempVars("EnDesarrollo") <> "Sí" Then
        DoCmd.Quit acQuitSaveAll
    End If
End Sub


Private Sub ImagenTareasOFF_Click()
    Dim parsed As Object
    Dim payload As Object

    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set parsed = JsonConverter.ParseJson(mod0BDOpcionesHelper.Opciones_OpenFormAction("tareas", m_Error))
    If m_Error <> "" Then Err.Raise 1000
    Set payload = parsed("payload")
    DoCmd.OpenForm CStr(payload("formName"))
    
    AvanceCerrar
    Me.lblEstado.Visible = False
    VBA.DoEvents
    DoCmd.Hourglass False
    Me.AllowEdits = True
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ImagenTareasOFF_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub ImagenTareasON_Click()
    Dim parsed As Object
    Dim payload As Object

    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set parsed = JsonConverter.ParseJson(mod0BDOpcionesHelper.Opciones_OpenFormAction("tareas", m_Error))
    If m_Error <> "" Then Err.Raise 1000
    Set payload = parsed("payload")
    DoCmd.OpenForm CStr(payload("formName"))
    
    AvanceCerrar
    Me.lblEstado.Visible = False
    VBA.DoEvents
    DoCmd.Hourglass False
    Me.AllowEdits = True
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ImagenTareasON_Click se ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub lblGestionExpedientes_Click()
    If Me.ComandoGestionExpedientes.Enabled = True Then
        ComandoGestionExpedientes_Click
    End If
End Sub

Private Sub lblAlta_Click()
    If Me.ComandoAlta.Enabled = True Then
        ComandoAlta_Click
    End If
    
End Sub
Private Sub lblGestorEntidades_Click()
    If Me.ComandoGestorEntidades.Enabled = True Then
        ComandoGestorEntidades_Click
    End If
    
End Sub
