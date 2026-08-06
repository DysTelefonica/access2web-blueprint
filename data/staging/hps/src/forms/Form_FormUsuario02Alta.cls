Attribute VB_Name = "Form_FormUsuario02Alta"
Option Compare Database
Option Explicit
Private m_Error As String



Private WithEvents m_FormExpedientes As Form_FormExpedientesBusqueda

Private Sub ComandoActualizar_Click()
    Set m_ObjEntorno.Expedientes = Nothing
End Sub

Private Sub ComandoBuscarExpediente_Click()
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Exp As Expediente
    
    Dim m_PalabraClave As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    
    m_PalabraClave = Nz(Me.PalabraClave, "")
    If m_PalabraClave = "" Then
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    
    Set m_Col = Constructor.getExpedientesBusqueda(p_PalabraClave:=m_PalabraClave, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Me.IDExpediente = Null
        Me.Titulo = Null
       
        If Nz(Me.IDExpediente, "") = "" Then
            Me.ComandoActualizar.Enabled = False
        Else
            Me.ComandoActualizar.Enabled = True
        End If
        
        ComandoIrABusquedaExpedientes_Click
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    
    If m_Col.Count = 1 Then
        For Each m_ID In m_Col
            Set m_Exp = m_Col(m_ID)
            If m_Exp.Titulo <> "" Then
                Me.Titulo = m_Exp.CodExp & " " & m_Exp.Titulo
                
            Else
                Me.Titulo = m_Exp.CodExp & " " & m_Exp.TextoExpediente
            End If
            Me.CadenaContratistas = m_Exp.CadenaContratistas
            Me.IDExpediente = m_Exp.IDExpediente
            Me.ComandoActualizar.Enabled = True
            Me.PalabraClave = Null
            Me.Titulo.ForeColor = Me.Correo_e.ForeColor
            
        Next
        
        
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    If FormularioAbierto("FormExpedientesBusqueda") Then
        DoCmd.Close acForm, "FormExpedientesBusqueda", acSaveNo
    End If
    DoCmd.OpenForm "FormExpedientesBusqueda"
    If FormularioAbierto("FormExpedientesBusqueda") Then
        Set m_FormExpedientes = Forms("FormExpedientesBusqueda")
        m_FormExpedientes.PalabraClave = Nz(Me.PalabraClave, "")
        Form_FormExpedientesBusqueda.Filtrar
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoBuscarExpediente_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If

End Sub

Private Sub ComandoIrABusquedaExpedientes_Click()
    On Error GoTo errores
    
   
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormExpedientesBusqueda") Then
        DoCmd.Close acForm, "FormExpedientesBusqueda", acSaveNo
    End If
    DoCmd.OpenForm "FormExpedientesBusqueda"
    If FormularioAbierto("FormExpedientesBusqueda") Then
        Set m_FormExpedientes = Forms("FormExpedientesBusqueda")
        If Nz(Me.PalabraClave, "") <> "" Then
            m_FormExpedientes.PalabraClave = Nz(Me.PalabraClave, "")
            Form_FormExpedientesBusqueda.Filtrar
        End If
        
    End If
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoIrABusquedaExpedientes_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    
End Sub



Private Sub PalabraClave_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Me.ComandoBuscarExpediente.Enabled = False Then
        Exit Sub
    End If
    
    If KeyCode = vbKeyReturn Then
        Me.ComandoBuscarExpediente.SetFocus
        Me.PalabraClave.SetFocus
        'Debug.Print Me.Pass
        ' Código que deseas ejecutar cuando se pulse Enter
        ComandoBuscarExpediente_Click
        
        ' Si deseas evitar que Access realice la acción predeterminada de Enter,
        ' puedes utilizar la siguiente línea:
        KeyCode = 0
    End If
End Sub

Private Sub CmdSalir_Click()
    DoCmd.Close acForm, Me.Name, acSaveYes
End Sub

Private Sub ComandoMotivos_Click()
    Dim m_Motivo As String
    On Error GoTo errores
    
    m_Error = ""
    m_Motivo = Nz(InputBox("Introduzca el motivo si no existe en la lista anterior", "Motivo nuevo"), "")
    If m_Motivo = "" Then
        Exit Sub
    End If
    MeterUnMotivoEnListaMOtivos m_Motivo, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Motivo_HPS.Requery
    Me.Motivo_HPS = m_Motivo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoMotivos_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub




Private Sub ComandoGrabar_Click()
    
    Dim m_Observacion As String
    Dim m_ObservacionCurso As String
    Dim m_ID As String
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    RellenarObjetoConDatos m_ObjUsuarioActivo, Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    'parte del curso
    
    m_ObjUsuarioActivo.Registrar , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_ID = m_ObjUsuarioActivo.ID
    
    
    m_Observacion = Nz(Me.Observaciones, "")
    If m_Observacion <> "" Then
        Set m_ObjObservacionActiva = New Observacion
        m_ObjObservacionActiva.ParaAlta = EnumSiNo.Sí
        m_ObjObservacionActiva.ID = m_ID
        m_ObjObservacionActiva.TIPO = "General"
        m_ObjObservacionActiva.Observacion = m_Observacion
        m_ObjObservacionActiva.Fecha = Date
        m_ObjObservacionActiva.Registrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_ObservacionCurso = Nz(Me.Obs_Curso, "")
    If m_ObservacionCurso <> "" Then
        Set m_ObjObservacionActiva = New Observacion
        m_ObjObservacionActiva.ParaAlta = EnumSiNo.Sí
        m_ObjObservacionActiva.ID = m_ID
        m_ObjObservacionActiva.TIPO = "Curso"
        m_ObjObservacionActiva.Observacion = m_ObservacionCurso
        m_ObjObservacionActiva.Fecha = Date
        m_ObjObservacionActiva.Registrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    cache_usuario_regenerar p_ID:=m_ObjUsuarioActivo.ID, p_Error:=m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    pregunta = MsgBox("Alta correcta", vbInformation, "Alta de Usuario")
    
    
    Form_FormInicial02ConsultasPrincipal.ComandoEditarUsuario_Click
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGrabar_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Correo_e_Exit(Cancel As Integer)
    Dim m_Valor As String
    
    On Error GoTo errores
    
    m_Error = ""
    m_Valor = Nz(Me.Correo_e, "")
    If m_Valor <> "" Then
        If Not FormatoDatoValido(UCase(m_Valor), "email", m_Error) Then
            pregunta = MsgBox("Formato de correo electrónico incorrecto", vbCritical, "Error")
            Exit Sub
        End If
        
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al abrir el salir del control se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub DNI_Exit(Cancel As Integer)
    Dim m_Valor As String
    Dim m_UsuarioConDNI As String
    On Error GoTo errores
    
    m_Error = ""
    m_Valor = Nz(Me.DNI, "")
    If Not m_ObjUsuarioAlInicio Is Nothing Then
        If m_ObjUsuarioAlInicio.DNI <> m_Valor Then
            If Len(m_Valor) <> 9 Then
                pregunta = MsgBox("Formato de DNI incorrecto", vbCritical, "Error")
                Exit Sub
            End If
            If Not FormatoDatoValido(UCase(m_Valor), "DNI", m_Error) Then
                pregunta = MsgBox("Formato de DNI incorrecto", vbCritical, "Error")
                Exit Sub
            End If
            m_UsuarioConDNI = DNIEnActivo(m_Valor, m_Error)
            If m_UsuarioConDNI <> "" Then
            Me.DNI = Null
                m_Error = "Hay un usuario registrado con ese DNI " & vbNewLine & m_UsuarioConDNI
                Err.Raise 1000
            End If
            m_UsuarioConDNI = DNIEnHistorico(m_Valor, m_Error)
            If m_UsuarioConDNI <> "" Then
                Me.DNI = Null
                m_Error = "Hay un usuario Histórico registrado con ese DNI " & vbNewLine & m_UsuarioConDNI
                Err.Raise 1000
            End If
        End If
    End If
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al abrir el salir del control se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Load()
    
    On Error GoTo errores
    
    m_Error = ""
   
    AjustarTamaño Me
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        Me.AllowEdits = False
    Else
        Me.AllowEdits = True
    End If
    Set m_ObjUsuarioActivo = New UsuarioHPS
    m_ObjUsuarioActivo.ParaAlta = EnumSiNo.Sí
    
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.AllowEdits = True Then
        Me.ComandoGrabar.Enabled = True
    Else
        Me.ComandoGrabar.Enabled = False
    End If
    RellenarComboEmpresas Me.IDEmpresaUsuario, EnumSiNo.No, , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarComboEmpresas Me.IDEmpresaHPS, EnumSiNo.Sí, , m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    RellenarComboMotivos Me.Motivo_HPS, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    CopiarEstiloBoton Me.ComandoGrabar
    CopiarEstiloBoton Me.cmdSalir
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al abrir el formulario se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub




Private Sub m_FormExpedientes_Seleccionado(m_Exp As Expediente)
    On Error GoTo errores
    
   
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.IDExpediente = m_Exp.IDExpediente
    If m_Exp.Titulo <> "" Then
        Me.Titulo = m_Exp.CodExp & " " & m_Exp.Titulo
        
    Else
        Me.Titulo = m_Exp.CodExp & " " & m_Exp.TextoExpediente
    End If
    Me.CadenaContratistas = m_Exp.CadenaContratistas
    If Nz(Me.IDExpediente, "") = "" Then
        Me.ComandoActualizar.Enabled = False
    Else
        Me.ComandoActualizar.Enabled = True
    End If
    Me.PalabraClave = Null
    Me.Titulo.ForeColor = Me.Correo_e.ForeColor
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormExpedientes_Seleccionado se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub