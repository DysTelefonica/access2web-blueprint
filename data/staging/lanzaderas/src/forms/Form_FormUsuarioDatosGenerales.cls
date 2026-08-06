Option Compare Database
Option Explicit
Private m_Error As String
Private m_NombreFormulario As String
Private m_ObjUsuarioAlInicio As usuario
Public Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean
    Dim m_Activado As Boolean
    Dim m_UsuarioImborrable As Boolean
    Dim m_PasswordNuncaCaduca As Boolean
    Dim m_MantenerLanzaderaAbierta As Boolean
    
    On Error GoTo errores
    If Nz(Me.ComboActivado, "") = "Sí" Then
        m_Activado = True
    End If
    If Nz(Me.ComboUsuarioImborrable, "") = "Sí" Then
        m_UsuarioImborrable = True
    End If
    If Nz(Me.ComboPasswordNuncaCaduca, "") = "Sí" Then
        m_PasswordNuncaCaduca = True
    End If
    If Nz(Me.ComboMantenerLanzaderaAbierta, "") = "Sí" Then
        m_MantenerLanzaderaAbierta = True
    End If
    With m_ObjUsuarioAlInicio
        If .CorreoUsuario <> Nz(Me.CorreoUsuario, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Nombre <> Nz(Me.Nombre, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Matricula <> Nz(Me.Matricula, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If Nz(Me.ComboEsAdministrador, "") = "" Then
            Me.ComboEsAdministrador = "No"
        End If
        
        If .EsAdministrador <> Me.ComboEsAdministrador Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Activado <> m_Activado Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .UsuarioImborrable <> m_UsuarioImborrable Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .PasswordNuncaCaduca <> m_PasswordNuncaCaduca Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .MantenerLanzaderaAbierta <> m_MantenerLanzaderaAbierta Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .PermisoPruebas <> Me.ComboPermisoPruebas Then
            HaHabidoCambios = True
            Exit Function
        End If
        
        If .FechaAlta <> Me.FechaAlta Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .FechaBaja <> Me.FechaBaja Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .UsuarioRed <> Me.UsuarioRed Then
            HaHabidoCambios = True
            Exit Function
        End If
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim m_Activado As String
    Dim m_UsuarioImborrable As String
    Dim m_PasswordNuncaCaduca As String
    Dim m_MantenerLanzaderaAbierta As String
    Dim m_PermisoPruebas As String
    On Error GoTo errores
    
    EstablecerCombo p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If m_ObjUsuarioActivo Is Nothing Then
        Set m_ObjUsuarioActivo = New usuario
        m_ObjUsuarioActivo.ParaAlta = EnumSino.Sí
        m_ObjUsuarioActivo.EsAdministrador = "No"
        m_ObjUsuarioActivo.Activado = True
        m_ObjUsuarioActivo.UsuarioImborrable = False
        m_ObjUsuarioActivo.PasswordNuncaCaduca = False
        m_ObjUsuarioActivo.MantenerLanzaderaAbierta = True
        m_ObjUsuarioActivo.PermisoPruebas = "No"
        m_ObjUsuarioActivo.FechaAlta = Date
        
    
        
    End If
    With m_ObjUsuarioActivo
        If .CorreoUsuario <> "" Then
            Me.CorreoUsuario = .CorreoUsuario
        End If
        If .Nombre <> "" Then
            Me.Nombre = .Nombre
        End If
        If .Matricula <> "" Then
            Me.Matricula = .Matricula
        End If
        If .Telefono <> "" Then
            Me.Telefono = .Telefono
        End If
        If .Movil <> "" Then
            Me.Movil = .Movil
        End If
        Me.ComboEsAdministrador = .EsAdministrador
        
        If .Activado = True Then
            m_Activado = "Sí"
        Else
            m_Activado = "No"
        End If
        If .UsuarioImborrable = True Then
            m_UsuarioImborrable = "Sí"
        Else
            m_UsuarioImborrable = "No"
        End If
        If .PasswordNuncaCaduca = True Then
            m_PasswordNuncaCaduca = "Sí"
        Else
            m_PasswordNuncaCaduca = "No"
        End If
        If .MantenerLanzaderaAbierta = True Then
            m_MantenerLanzaderaAbierta = "Sí"
        Else
            m_MantenerLanzaderaAbierta = "No"
        End If
        
        Me.ComboActivado = m_Activado
        Me.ComboUsuarioImborrable = m_UsuarioImborrable
        Me.ComboPasswordNuncaCaduca = m_PasswordNuncaCaduca
        Me.ComboMantenerLanzaderaAbierta = m_MantenerLanzaderaAbierta
        
        If IsDate(.FechaAlta) Then
            Me.FechaAlta = Format(.FechaAlta, "dd/mm/yyyy")
        End If
        
        If IsDate(.FechaBaja) Then
            Me.FechaBaja = Format(.FechaBaja, "dd/mm/yyyy")
        End If
       
        If .UsuarioRed <> "" Then
            Me.UsuarioRed = .UsuarioRed
        End If
        If .PermisoPruebas = "" Then
            m_PermisoPruebas = "No"
        Else
            m_PermisoPruebas = .PermisoPruebas
        End If
        Me.ComboPermisoPruebas = m_PermisoPruebas
        Me.ID = .ID
        If .ParaAlta = EnumSino.No Then
            Set m_ObjUsuarioAlInicio = Constructor.getUsuario(.ID, , , , m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            
        End If
        If Not .JefeUsuarioObj Is Nothing Then
            Me.Responsable = .JefeUsuarioObj.ID
        End If
        Me.TextoTiempoParaExpirarPass = .TextoTiempoParaExpirarPass
    End With
    
    If m_ObjUsuarioConectadoLogin.EsAdministradorCalculado = EnumSino.Sí Then
        Me.ComandoRegistrar.Enabled = True
        Me.AllowEdits = True
    Else
        Me.ComandoRegistrar.Enabled = False
        Me.AllowEdits = False
    End If
    If m_ObjUsuarioActivo.ID = "" Then
        Me.ComandoCambiarPass.Enabled = False
    Else
        If m_ObjUsuarioActivo.ID = m_ObjUsuarioConectadoLogin.ID Then
            Me.ComandoCambiarPass.Enabled = True
        Else
            Me.ComandoCambiarPass.Enabled = False
        End If
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Private Sub ComandoCambiarPass_Click()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormUsuarioCambioPass") Then
        DoCmd.Close acForm, "FormUsuarioCambioPass", acSaveNo
    End If
    DoCmd.OpenForm "FormUsuarioCambioPass"
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoCambiarPass_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoRegistrar_Click()
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If Not m_ObjUsuarioAlInicio Is Nothing Then
        If Not HaHabidoCambios(m_Error) Then
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            m_Error = "Ningún cambio que registrar"
            Err.Raise 1000
        End If
    End If
    If m_ObjUsuarioActivo Is Nothing Then
        Set m_ObjUsuarioActivo = New usuario
    End If
    With m_ObjUsuarioActivo
        .CorreoUsuario = Nz(Me.CorreoUsuario, "")
        .Nombre = Nz(Me.Nombre, "")
        .Matricula = Nz(Me.Matricula, "")
        .Telefono = Nz(Me.Telefono, "")
        .Movil = Nz(Me.Movil, "")
        If Nz(Me.ComboEsAdministrador, "") = "Sí" Then
            .EsAdministrador = "Sí"
        Else
            .EsAdministrador = "No"
        End If
        
        If Nz(Me.ComboUsuarioImborrable, "") = "Sí" Then
            .UsuarioImborrable = True
        Else
            .UsuarioImborrable = False
        End If
        If Nz(Me.ComboPasswordNuncaCaduca, "") = "Sí" Then
            .PasswordNuncaCaduca = True
        Else
            .PasswordNuncaCaduca = False
        End If
        .FechaAlta = Nz(Me.FechaAlta, "")
        .FechaBaja = Nz(Me.FechaBaja, "")
        .UsuarioRed = Nz(Me.UsuarioRed, "")
        
        If Nz(Me.ComboMantenerLanzaderaAbierta, "") = "Sí" Then
            .MantenerLanzaderaAbierta = True
        Else
            .MantenerLanzaderaAbierta = False
        End If
       If Nz(Me.ComboPermisoPruebas, "") = "Sí" Then
            .PermisoPruebas = "Sí"
        Else
            .PermisoPruebas = "No"
        End If
        If Nz(Me.Responsable.Column(0), "") <> "" Then
            .JefeDelUsuario = Constructor.getUsuario(Me.Responsable.Column(0)).UsuarioRed
        Else
            .JefeDelUsuario = ""
        End If
        If Nz(Me.ComboActivado, "") <> "No" Then
            .Activado = True
            .FechaBloqueo = ""
        Else
            .Activado = False
            .FechaBloqueo = Now()
        End If
        .Registrar m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    
    Set m_ObjUsuarioAlInicio = Constructor.getUsuario(m_ObjUsuarioActivo.ID, , , , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.ID = m_ObjUsuarioAlInicio.ID
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoRegistrar_Click se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    m_NombreFormulario = Me.Name
    If m_ObjEntorno Is Nothing Then
        EVE
    End If
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    
    Pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, m_NombreFormulario, acSaveNo
End Sub

Public Function EstablecerCombo(Optional ByRef p_Error As String) As String
    
    Dim cmb As ComboBox
    Dim m_ID As Variant
    Dim m_ObjUsuario As usuario
    Dim m_ObjCol As Scripting.Dictionary

    On Error GoTo errores
    
    Set cmb = Me.Responsable
    cmb.RowSource = ""
    Set m_ObjCol = Constructor.getUsuarios(p_Error)
    
   
    For Each m_ID In m_ObjCol
        Set m_ObjUsuario = m_ObjCol(m_ID)
        cmb.AddItem m_ObjUsuario.ID & ";" & m_ObjUsuario.Nombre
        Set m_ObjUsuario = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerCombo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
