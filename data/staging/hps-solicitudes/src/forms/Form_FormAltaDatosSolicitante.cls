Option Compare Database
Option Explicit

Public Event Alta()


Private WithEvents m_FormExpedientes As Form_FormExpedientesBusqueda
Private m_Error As String

Private Sub cmdSalir_Click()
   

    
    DoCmd.Close acForm, Me.Name
   
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

Private Sub ComandoBuscarExpediente_Click()
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_Expediente As Expediente
    
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
    
    Set m_Col = constructor.getExpedientesBusqueda(p_PalabraClave:=m_PalabraClave, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Me.IDExpediente = Null
        Me.Titulo = Null
       
       
        
        ComandoIrABusquedaExpedientes_Click
        VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    
    If m_Col.Count = 1 Then
        For Each m_ID In m_Col
            Set m_Expediente = m_Col(m_ID)
            Me.IDExpediente = m_Expediente.IDExpediente
            Me.Titulo = m_Expediente.Titulo
            
        Next
        Me.PalabraClave = Null
        
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
Private Sub ComandoActualizar_Click()
    Set m_ObjEntorno.Expedientes = Nothing
End Sub
Private Sub m_FormExpedientes_Seleccionado(m_Exp As Expediente)
    On Error GoTo errores
    
   
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.IDExpediente = m_Exp.IDExpediente
    If m_Exp.Titulo <> "" Then
        Me.Titulo = m_Exp.Titulo
    Else
        Me.Titulo = m_Exp.TextoExpediente
    End If
    If Nz(Me.IDExpediente, "") = "" Then
        Me.ComandoActualizar.Enabled = False
    Else
        Me.ComandoActualizar.Enabled = True
    End If
    Me.PalabraClave = Null
   
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
Private Sub ComandoActualizarar_Click()
    Set m_ObjEntorno.Expedientes = Nothing
End Sub
Private Sub PalabraClave_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Me.ComandoBuscarExpediente.Enabled = False Then
        Exit Sub
    End If
    
    If KeyCode = vbKeyReturn Or KeyCode = 9 Then
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
Private Sub ComandoAyuda_Click()
     On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    Me.AllowEdits = False
    AbrirAyuda m_Error
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
        m_Error = "Al ComandoAyuda_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        On Error Resume Next
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")

End Sub

Private Sub ComandoGenerarDNIValido_Click()
    Me.DNI = GenerarDNINIEValido()
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





Private Sub Form_Load()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.AllowEdits = False
    Ajustar Me
    EstablecerDatos m_Error
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
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    
    On Error GoTo errores
    
    EstablecerCombos p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Gestor = m_ObjUsuarioConectado.Nombre
    Me.CorreosAutomaticos = m_ObjEntorno.CorreosAutomaticos
    If Application.TempVars("EnPruebas") = "Sí" Then
        Me.ComandoGenerarDNIValido.Visible = True
    Else
        Me.ComandoGenerarDNIValido.Visible = False
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function

Public Function RellenarObjeto(Optional ByRef p_Error As String) As String
    
    
    On Error GoTo errores
    
    
    Set m_ObjSolicitudActiva = New solicitud
    With m_ObjSolicitudActiva
        .TIPO = Nz(Me.TIPO, "")
        .CorreosAutomaticos = Nz(Me.CorreosAutomaticos, "")
        
        .Gestor = Nz(Me.Gestor, "")
        .Nombre = Nz(Me.Nombre, "")
        .Apellido1 = Nz(Me.Apellido1, "")
        .Apellido2 = Nz(Me.Apellido2, "")
        .FNacimiento = Nz(Me.FNacimiento, "")
        .LugarNacimiento = Nz(Me.LugarNacimiento, "")
        .email = Nz(Me.email, "")
        .emailResponsable = Nz(Me.emailResponsable, "")
        .Telefono = Nz(Me.Telefono, "")
        .IDEmpresaUsuario = Nz(Me.IDEmpresaUsuario, "")
        .IDEmpresaTramitadora = Nz(Me.IDEmpresaTramitadora, "")
        .IDExpediente = Nz(Me.IDExpediente, "")
        .DNI = Nz(Me.DNI, "")
        .Motivo_HPS = Nz(Me.Motivo_HPS, "")
        .Observaciones = Nz(Me.Observaciones, "")
        
    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarObjeto ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function
Private Sub ComandoRegistrar_Click()
    
    
    Dim m_SolicitudOp As SolicitudOperaciones
    On Error GoTo errores
    m_Error = ""
    DoCmd.Hourglass True
    Me.AllowEdits = False
    
    RellenarObjeto m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_SolicitudOp = New SolicitudOperaciones
    With m_SolicitudOp
        Set .solicitud = m_ObjSolicitudActiva
        .RegistrarDatos , m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        RaiseEvent Alta
    End With
    Set m_ObjSolicitudActiva = constructor.getSolicitud(p_IDSolicitud:=m_ObjSolicitudActiva.IDSolicitud, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.AllowEdits = True
    If FormularioAbierto("FormSolicitudAltaTraspaso") Then
        DoCmd.Close acForm, "FormSolicitudAltaTraspaso", acSaveNo
    End If
    DoCmd.OpenForm "FormSolicitudAltaTraspaso"
    DoCmd.Close acForm, Me.Name, acSaveNo
    
    DoCmd.Hourglass False
    
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "El método ComandoRegistrar_Click ha devuelto el error: " & Err.Numbr & vbCrLf & "Detalle: " & Err.Description
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Public Function EstablecerCombos(Optional ByRef p_Error As String) As String
    
    Dim cmb As ComboBox
    Dim m_ID As Variant
    Dim m_Col As Scripting.Dictionary
    Dim m_Expediente As Expediente
    Dim m_Suministrador As Suministrador
    Dim m_Nombre As String
    
    On Error GoTo errores
    RellenarComboTramitadores Me.Gestor, EnumSiNo.no, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Set cmb = Me.IDEmpresaTramitadora
    cmb.RowSource = ""
    Set m_Col = m_ObjEntorno.SumTramitadoras
   
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
            Set m_Suministrador = m_Col(m_ID)
            m_Nombre = m_Suministrador.Nemotecnico
            If m_Nombre = "" Then
                m_Nombre = m_Suministrador.Nombre
            End If
            cmb.AddItem m_Suministrador.IDSuministrador & ";" & m_Nombre
            
            Set m_Suministrador = Nothing
        Next
    End If
    Set cmb = Me.IDEmpresaUsuario
    cmb.RowSource = ""
    Set m_Col = m_ObjEntorno.Suministradores
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
             Set m_Suministrador = m_Col(m_ID)
            m_Nombre = m_Suministrador.Nemotecnico
            If m_Nombre = "" Then
                m_Nombre = m_Suministrador.Nombre
            End If
            cmb.AddItem m_Suministrador.IDSuministrador & ";" & m_Nombre
            
            Set m_Suministrador = Nothing
        Next
    End If
    
    
    Set cmb = Me.Motivo_HPS
    RellenarComboMotivos cmb, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerCombos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function

