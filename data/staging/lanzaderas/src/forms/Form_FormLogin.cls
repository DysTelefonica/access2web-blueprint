Option Compare Database
Option Explicit

Private m_Error As String

Private NumeroVecesIntentado As Integer

Private Sub ComandoCambioDeContrasenia_Click()
    
    Dim m_CorreoUsuario As String
    
    Dim m_ObjUsuario As usuario
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    m_CorreoUsuario = Nz(Me.CorreoUsuario, "")
   
    
    Set m_ObjUsuario = Constructor.getUsuario(, , , m_CorreoUsuario, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ObjUsuario Is Nothing Then
        m_Error = "No hay registrado ningún usuario con ese correo. Hable con su administrador"
        Err.Raise 1000
    End If
    
    m_ObjUsuario.ResetearPass m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
   
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Pregunta = MsgBox("En un periodo máximo de 10 minutos recibirá un correo electrónico con una contraseña de un solo uso", vbInformation, "Lanzadera")
    
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar obtener la contraseña se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub




Private Sub ComandoIniciarSesion_Click()
    Dim m_SSID As String
    Dim m_Correo As String
    Dim m_Pass As String
    Dim m_EstadoLogin As EnumPassEstados
    Dim m_UsuarioConectadoMaquinaEsAdministrador As EnumSino
    Dim m_UsuarioMaquina As usuario
    Dim blnEstablecerArchivoSSID
    On Error GoTo errores
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If NumeroVecesIntentado > 5 Then
        m_Error = "Se ha intentado más de 5 veces y la cuenta se ha bloqueado. Hable con su administrador"
        Err.Raise 1000
    End If
    If Nz(Me.Pass, "") = "" Then
        m_Error = "La contraseña no puede estar vacía"
        Err.Raise 1000
    End If
    If InStr(1, Nz(Me.CorreoUsuario, ""), "@") = 0 Then
        m_Error = "Se ha de introducir un correo válido"
        Err.Raise 1000
    End If
   
    If Not EVEEjecutado Then
        Avance "Estableciendo las variables de entorno"
        EVE m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Set m_UsuarioMaquina = m_ObjEntorno.UsuarioConectadoMaquina
    m_Error = m_ObjEntorno.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_UsuarioMaquina Is Nothing Then
        m_UsuarioConectadoMaquinaEsAdministrador = EnumSino.No
    Else
        m_UsuarioConectadoMaquinaEsAdministrador = m_UsuarioMaquina.EsAdministradorCalculado
    End If
    m_Correo = Me.CorreoUsuario
    m_Pass = Me.Pass
    Set m_ObjUsuarioConectadoLogin = Constructor.getUsuario(, , , m_Correo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_CorreoUsuarioEnLogin = m_ObjUsuarioConectadoLogin.CorreoUsuario
    If m_ObjUsuarioConectadoLogin Is Nothing Then
        
        DoCmd.Hourglass False
        Pregunta = MsgBox("Está intentando iniciar sesión con un correo electrónico no registrado" & vbNewLine & _
                        "Ha de solicitar al alta de esta cuenta", vbExclamation, "Usuario no registrado")
        
        DoCmd.Quit acQuitSaveNone
        Exit Sub
    End If
    If m_UsuarioConectadoMaquinaEsAdministrador = EnumSino.Sí Then
        m_EstadoLogin = EnumPassEstados.Correcto
    Else
        m_EstadoLogin = Login(m_ObjUsuarioConectadoLogin, m_Pass, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    
    If m_EstadoLogin = EnumPassEstados.caducada Then
        Pregunta = MsgBox("La Pass actual ha caducado y ha de cambiarla", vbExclamation, "Lanzadera")
        'close splash form & open main form
        DoCmd.Close acForm, Me.Name
        DoCmd.OpenForm "FormUsuarioCambioPass"
        AvanceCerrar
        Exit Sub
    End If
    If m_EstadoLogin = EnumPassEstados.Correcto Then
        SetSSOFAArchivo m_Correo, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
         
        
        LoginCorrecto m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
        AvanceCerrar
        Exit Sub
    End If
    If m_EstadoLogin = EnumPassEstados.Incorrecto Then
        m_Error = "Usuario/Pass no válida. Puede solicitar una nueva Pass pulsando el botón Obtener"
        Err.Raise 1000
    End If
    If m_EstadoLogin = EnumPassEstados.NecesitaCambio Then
        Pregunta = MsgBox("La Pass actual es de un solo uso y ha de cambiarla", vbExclamation, "Lanzadera")
        
        DoCmd.Close acForm, Me.Name
        DoCmd.OpenForm "FormUsuarioCambioPass"
        AvanceCerrar
        Exit Sub
    End If
    If m_EstadoLogin = EnumPassEstados.Bloqueada Then
        CorreoCuentaBloqueada m_ObjUsuarioConectadoLogin, m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
        AvanceCerrar
        Pregunta = MsgBox("Cuenta bloqueada por exceso de intentos fallidos", vbExclamation, "Cuenta bloqueada")
        
        DoCmd.Quit acQuitSaveAll
        Exit Sub
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoIniciarSesion_Click se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub ComandoSalir_Click()
    On Error Resume Next
    Pregunta = MsgBox("¿Desea salir de la lanzadera", vbExclamation + vbYesNo + vbDefaultButton2, "Salir de la Lanzadera")
    If Pregunta <> vbYes Then
        Exit Sub
    End If
    DoCmd.Quit acQuitSaveAll
End Sub



Private Sub Form_Load()
    
    
    Dim m_Prueba As String
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Ajustar Me
    DoCmd.ShowToolbar "Ribbon", acToolbarNo
    
    EVEMin m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjUsuarioConectadoLogin Is Nothing Then
        LoginCorrecto m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        AvanceCerrar
         VBA.DoEvents
        DoCmd.Hourglass False
        VBA.DoEvents
        Exit Sub
    End If
    m_Prueba = "CualquierCosaAEncriptar"
    
    m_Prueba = SHA256(m_Prueba)
    If m_Prueba = "" Then
        Pregunta = MsgBox("No tiene instalado el NetFrameWork 3.5.Siga el manual para hacerlo", vbCritical, "Falta un componente")
        
        DoCmd.Hourglass False
        DoCmd.Quit acQuitSaveNone
        Exit Sub
    End If
   
    Me.Caption = Me.lblTitulo1.Caption
    If InStr(1, m_CorreoPorArchivo, "@") <> 0 Then
        Me.CorreoUsuario = m_CorreoPorArchivo
        Me.Pass.SetFocus
    End If
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar Form_Open se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Quit acQuitSaveNone
End Sub


Private Sub ComandoAyudaFormulario_Click()
    On Error Resume Next
    AbrirArchivoAyuda Me
End Sub
Private Function EstablecerInicio(Optional ByRef p_Error As String) As String
    
    Dim m_Correo As String
    Dim m_Pass As String
    Dim m_FormatoCorreoCorrecto As EnumSino
    
    On Error GoTo errores
    
    m_Pass = Nz(Me.Pass, "")
    If m_Pass = "" Then
       
        Exit Function
    End If
    m_Correo = Nz(Me.CorreoUsuario, "")
    If m_Correo = "" Then
       
        Exit Function
    End If
    m_FormatoCorreoCorrecto = FormatoCorreoCorrecto(m_Correo, p_Error)
    If m_FormatoCorreoCorrecto = EnumSino.No Then
        
        Exit Function
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerInicio ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Private Sub Pass_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
        Me.CorreoUsuario.SetFocus
        Me.Pass.SetFocus
        'Debug.Print Me.Pass
        ' Código que deseas ejecutar cuando se pulse Enter
        ComandoIniciarSesion_Click
        
        ' Si deseas evitar que Access realice la acción predeterminada de Enter,
        ' puedes utilizar la siguiente línea:
        KeyCode = 0
    End If
End Sub