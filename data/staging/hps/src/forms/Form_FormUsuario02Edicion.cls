Attribute VB_Name = "Form_FormUsuario02Edicion"
Option Compare Database
Option Explicit
Private m_Error As String
Private m_ObjObservacionSeleccionada As Observacion

Private WithEvents m_FormUsuario02DatosObservacion As Form_FormUsuario02DatosObservacion

Private WithEvents m_FormExpedientes As Form_FormExpedientesBusqueda

Public Event EdicionRegistro()
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
    
   
    If KeyCode = vbKeyReturn Then
        If Me.AllowEdits = False Then
            Exit Sub
        End If
        Me.ComandoBuscarExpediente.Enabled = True
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

Public Function getUsuarioAlInicio(Optional ByRef p_Error As String) As UsuarioHPS
    
    On Error GoTo errores
    
    Set getUsuarioAlInicio = m_ObjUsuarioAlInicio
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuarioAlInicio ha producido el error: " & Err.Description
    End If
End Function



Private Sub ComandoDetalleUsuarioSICA_Click()
    
    Dim m_HaHabidoCambios As Boolean
    On Error GoTo errores
    
    m_Error = ""
    If m_ObjUsuarioSICAActivo Is Nothing Then
        Exit Sub
    End If
    If Not m_ObjUsuarioAlInicio Is Nothing Then
        m_HaHabidoCambios = HaHabidoCambios(m_ObjUsuarioAlInicio, Me, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_HaHabidoCambios Then
            ComandoGrabar_Click
        End If
    End If
    
    
    
    If FormularioAbierto("FormUsuario07SICAEdicion") Then
        DoCmd.Close acForm, "FormUsuario07SICAEdicion", acSaveNo
    End If
    m_ObjUsuarioSICAActivo.ParaAlta = EnumSiNo.No
    DoCmd.OpenForm "FormUsuario07SICAEdicion"
    
    
    
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al usar el método ComandoDetalleUsuarioSICA_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoGrabar_Click()
    
    Dim m_ID As String
    Dim m_HaHabidoCambios As Boolean
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If m_ObjUsuarioActivo Is Nothing Then
        m_Error = "No se conoce el usuario activo"
        Err.Raise 1000
    End If
    If m_ObjUsuarioAlInicio Is Nothing Then
        m_Error = "No se conoce el usuario activo al inicio"
        Err.Raise 1000
    End If
    m_ID = m_ObjUsuarioAlInicio.ID
    m_HaHabidoCambios = HaHabidoCambios(m_ObjUsuarioAlInicio, Me, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_HaHabidoCambios Then
        m_Error = "No hay cambios que grabar"
        Err.Raise 1000
    End If
    RellenarObjetoConDatos m_ObjUsuarioActivo, Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    m_ObjUsuarioActivo.ParaAlta = EnumSiNo.No
    m_ObjUsuarioActivo.Registrar m_ObjUsuarioAlInicio, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RaiseEvent EdicionRegistro
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents

    pregunta = MsgBox("Edición correcta", vbInformation, "Edición de Usuario")

    DoCmd.Close acForm, "FormUsuario00Principal", acSaveNo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoGrabar_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoObservacionAlta_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    Set m_ObjObservacionSeleccionada = Nothing
    If FormularioAbierto("FormUsuario02DatosObservacion") Then
        DoCmd.Close acForm, "FormUsuario02DatosObservacion", acSaveNo
    End If
    Set m_ObjObservacionActiva = m_ObjObservacionSeleccionada
    DoCmd.OpenForm "FormUsuario02DatosObservacion", OpenArgs:="General"
    Set m_FormUsuario02DatosObservacion = Forms("FormUsuario02DatosObservacion")
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar abrir el formulario se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoObservacionCursoAlta_Click()
    
    
    On Error GoTo errores
    
    m_Error = ""
    Set m_ObjObservacionSeleccionada = Nothing
    If FormularioAbierto("FormUsuario02DatosObservacion") Then
        DoCmd.Close acForm, "FormUsuario02DatosObservacion", acSaveNo
    End If
    Set m_ObjObservacionActiva = m_ObjObservacionSeleccionada
    DoCmd.OpenForm "FormUsuario02DatosObservacion", OpenArgs:="Curso"
    Set m_FormUsuario02DatosObservacion = Forms("FormUsuario02DatosObservacion")
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar abrir el formulario se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoObservacionCursoEditar_Click()
    On Error GoTo errores
    
    m_Error = ""
    If m_ObjObservacionSeleccionada Is Nothing Then
        m_Error = "Se la de seleccionar una observación de la lista"
        Err.Raise 1000
    End If
    If m_ObjObservacionSeleccionada.TIPO <> "Curso" Then
        m_Error = "Se la de seleccionar una observación de la lista"
        Err.Raise 1000
    End If
    Set m_ObjObservacionActiva = m_ObjObservacionSeleccionada
    If FormularioAbierto("FormUsuario02DatosObservacion") Then
        DoCmd.Close acForm, "FormUsuario02DatosObservacion", acSaveNo
    End If
    DoCmd.OpenForm "FormUsuario02DatosObservacion"
    Set m_FormUsuario02DatosObservacion = Forms("FormUsuario02DatosObservacion")
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar abrir el formulario se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoObservacionCursoEliminar_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    If m_ObjObservacionSeleccionada Is Nothing Then
        m_Error = "Se la de seleccionar una observación de la lista"
        Err.Raise 1000
    End If
    If m_ObjObservacionSeleccionada.TIPO <> "Curso" Then
        m_Error = "Se la de seleccionar una observación de la lista"
        Err.Raise 1000
    End If
    pregunta = MsgBox("¿Desea elmiminar la observación seleccionada?", _
                        vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar observación")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    Set m_ObjObservacionActiva = m_ObjObservacionSeleccionada
    m_ObjObservacionActiva.Borrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjObservacionActiva = Nothing
    Set m_ObjObservacionSeleccionada = Nothing
    EstablecerListasObservaciones "Curso", m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoObservacionCursoEliminar_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoObservacionEditar_Click()
    
    On Error GoTo errores
    m_Error = ""
    
    If m_ObjObservacionSeleccionada Is Nothing Then
        m_Error = "Se la de seleccionar una observación de la lista"
        Err.Raise 1000
    End If
    If m_ObjObservacionSeleccionada.TIPO <> "General" Then
        m_Error = "Se la de seleccionar una observación de la lista"
        Err.Raise 1000
    End If
    Set m_ObjObservacionActiva = m_ObjObservacionSeleccionada
    If FormularioAbierto("FormUsuario02DatosObservacion") Then
        DoCmd.Close acForm, "FormUsuario02DatosObservacion", acSaveNo
    End If
    DoCmd.OpenForm "FormUsuario02DatosObservacion"
    Set m_FormUsuario02DatosObservacion = Forms("FormUsuario02DatosObservacion")
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar abrir el formulario se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoObservacionEliminar_Click()
    
    On Error GoTo errores
    
    m_Error = ""
    If m_ObjObservacionSeleccionada Is Nothing Then
        m_Error = "Se la de seleccionar una observación de la lista"
        Err.Raise 1000
    End If
    If m_ObjObservacionSeleccionada.TIPO <> "General" Then
        m_Error = "Se la de seleccionar una observación de la lista"
        Err.Raise 1000
    End If
    pregunta = MsgBox("¿Desea elmiminar la observación seleccionada?", _
                        vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar observación")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    Set m_ObjObservacionActiva = m_ObjObservacionSeleccionada
    m_ObjObservacionActiva.Borrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjObservacionActiva = Nothing
    Set m_ObjObservacionSeleccionada = Nothing
    EstablecerListasObservaciones "General", m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al intentar eliminar la observación se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoPasarAHistorico_Click()
    Dim m_Servicio As UsuarioServicio
    On Error GoTo errores

    m_Error = ""
    pregunta = MsgBox("¿Desea realmente pasar este usuario a Histórico?", vbExclamation + vbYesNo + vbDefaultButton2, "Pasar a Actual")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    Set m_Servicio = New UsuarioServicio

    ' Reloj de arena durante el move (puede tardar con muchos anexos / observación)
    DoCmd.Hourglass True
    m_Servicio.PasarDeActualAHistorico m_ObjUsuarioActivo.ID, m_Error
    If m_Error <> "" Then
        DoCmd.Hourglass False
        Err.Raise 1000
    End If
    DoCmd.Hourglass False

    If m_Error = "" Then
        Indicadores_NotificarMovimiento "c2h", CStr(m_ObjUsuarioActivo.ID)
    End If

    MsgBox "El usuario ha sido pasado a Histórico correctamente.", vbInformation, "Operación completada"

    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoPasarAHistorico_Click se ha producido el error: " & Err.Description
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
Public Sub Form_Load()
    
    Dim bSoloLectura As Boolean
    Dim ctl As Control
    
    On Error GoTo errores
    
    m_Error = ""
    
    ' --- 1. DETECTAR MODO (Sin cambiar AllowEdits todavía) ---
    If Application.TempVars("SoloLectura") = "Sí" Then
        bSoloLectura = True
    Else
        bSoloLectura = False
    End If
    
    ' --- IMPORTANTE: SIEMPRE Permitimos Edits a nivel de formulario ---
    ' Si lo ponemos en False, las listas se congelan.
    ' La protección la haremos control a control más abajo.
    Me.AllowEdits = True
    
    ' 2. Carga de objetos básicos
    If m_ObjUsuarioActivo Is Nothing Then
        m_Error = "No se ha abierto el formulario con los datos necesarios"
        Err.Raise 1000
    End If
    
    Set m_ObjUsuarioAlInicio = getUsuarioHPS(m_ObjUsuarioActivo.ID, , m_Error)
    If m_Error <> "" Then Err.Raise 1000

    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    ' 3. ESTADO INICIAL DE BOTONES
    ' Botones de Observaciones (Edición) apagados al inicio
    Me.ComandoObservacionEditar.Enabled = False
    Me.ComandoObservacionEliminar.Enabled = False
    Me.ComandoObservacionCursoEditar.Enabled = False
    Me.ComandoObservacionCursoEliminar.Enabled = False
    Me.ComandoPasarAHistorico.Visible = False

    ' Lógica según modo Lectura/Escritura
    If bSoloLectura = False Then
        ' --- MODO ESCRITURA ---
        Me.ComandoGrabar.Enabled = True
        Me.ComandoObservacionAlta.Enabled = True
        Me.ComandoObservacionCursoAlta.Enabled = True
        
        Me.PalabraClave.Enabled = True
        Me.ComandoBuscarExpediente.Enabled = True
        Me.ComandoIrABusquedaExpedientes.Enabled = True
        Me.ComandoActualizar.Enabled = True
    Else
        ' --- MODO SOLO LECTURA ---
        Me.ComandoGrabar.Enabled = False
        Me.ComandoObservacionAlta.Enabled = False
        Me.ComandoObservacionCursoAlta.Enabled = False
        
        ' Búsqueda y cambios de expediente BLOQUEADOS
        Me.PalabraClave.Enabled = False
        Me.ComandoBuscarExpediente.Enabled = False
        Me.ComandoIrABusquedaExpedientes.Enabled = False
        Me.ComandoActualizar.Enabled = False
    End If
    
    ' Copiar Estilos
    CopiarEstiloBoton Me.ComandoGrabar
    CopiarEstiloBoton Me.ComandoBuscarExpediente
    CopiarEstiloBoton Me.ComandoIrABusquedaExpedientes
    CopiarEstiloBoton Me.ComandoActualizar
    CopiarEstiloBoton Me.ComandoDetalleUsuarioSICA
    CopiarEstiloBoton Me.ComandoPasarAHistorico
    
    ' 4. Cargar Datos
    EstablecerDatos m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    ' --- 5. EL TRUCO FINAL: BLOQUEO QUIRÚRGICO ---
    ' Si es solo lectura, bloqueamos los Inputs para que no escriban,
    ' PERO dejamos las Listas vivas.
    If bSoloLectura = True Then
        For Each ctl In Me.Controls
            ' Bloqueamos Textbox, Combos y Checkbox
            If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
                ctl.Locked = True   ' No se puede escribir
                ' ctl.Enabled = False ' Opcional: Descomenta si quieres que se vean grises
            End If
            ' Aseguramos que las listas SI funcionen
            If TypeOf ctl Is ListBox Then
                ctl.Locked = False
                ctl.Enabled = True
            End If
        Next
    End If

    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al abrir el formulario se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    Dim ctl As Control
    Dim m_Exp As Expediente
    Dim m_Dato As String
    Dim m_NombreCampo As String
    Dim m_Lugar As String
    
    On Error GoTo errores
    
    p_Error = ""
    RellenarComboEmpresas Me.IDEmpresaUsuario, EnumSiNo.No, , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    RellenarComboEmpresas Me.IDEmpresaHPS, EnumSiNo.Sí, , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    RellenarComboMotivos Me.Motivo_HPS, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_Exp = Constructor.getExpediente(p_IDExpediente:=m_ObjUsuarioActivo.IDExpediente, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Exp Is Nothing Then
        Me.Titulo = m_Exp.CodExp & " " & m_Exp.Titulo
        Me.CadenaContratistas = m_Exp.CadenaContratistas
        Me.IDExpediente = Me.IDExpediente
    End If
    Me.ListaObservaciones.RowSource = ""
    Me.ListaObservacionesCurso.RowSource = ""
    If m_ObjUsuarioAlInicio Is Nothing Then
        Exit Function
    End If
    DatosDelObjetoAlFormulario m_ObjUsuarioAlInicio, Me, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Then
            
            If Me.AllowEdits = True Then
                If InStr(1, ctl.Tag, "DATOS;INTOCABLE") <> 0 Then
                    ctl.Enabled = False
                End If
            End If
        End If
    Next
   
    EstablecerListasObservaciones , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjUsuarioSICAActivo = getUsuarioSICA(, m_ObjUsuarioActivo.ID, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjUsuarioSICAActivo Is Nothing Then
        Me.ComandoDetalleUsuarioSICA.Enabled = True
    Else
        Me.ComandoDetalleUsuarioSICA.Enabled = False
    End If
    If Me.AllowEdits = True Then
        If IsDate(m_ObjUsuarioActivo.F_BajaCalculada) Then
            Me.F_Baja = m_ObjUsuarioActivo.F_BajaCalculada
            Me.ComandoPasarAHistorico.Visible = True
            
        Else
            Me.F_Baja = Null
            Me.ComandoPasarAHistorico.Visible = False
            
        End If
        
    Else
        Me.ComandoPasarAHistorico.Visible = False
    End If
    'Me.ComandoPasarAHistorico.Visible = True
    Me.Titulo.ForeColor = Me.Correo_e.ForeColor
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha devuelto el error: " & Err.Description
    End If
End Function
Private Function EstablecerListasObservaciones( _
                                                Optional p_Tipo As String, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    Dim m_IDObservacion As Variant
    Dim m_ObjObServacion As Observacion
    Dim lst As ListBox
    Dim m_Col As Scripting.Dictionary
    
    On Error GoTo errores
    
    p_Error = ""
    m_Error = ""
    
    If p_Tipo = "Curso" Then
        Set lst = Me.ListaObservacionesCurso
        Set m_ObjUsuarioActivo.ColObservacionesCurso = Nothing
        Set m_Col = m_ObjUsuarioActivo.ColObservacionesCurso
    ElseIf p_Tipo = "General" Then
        Set lst = Me.ListaObservaciones
        Set m_ObjUsuarioActivo.ColObservaciones = Nothing
        Set m_Col = m_ObjUsuarioActivo.ColObservaciones
    Else
        EstablecerListasObservaciones = EstablecerListasObservaciones("General")
        EstablecerListasObservaciones = EstablecerListasObservaciones("Curso")
        Exit Function
    End If
    lst.RowSource = ""
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_IDObservacion In m_Col
        Set m_ObjObServacion = m_Col(m_IDObservacion)
        With m_ObjObServacion
            lst.AddItem .IDObservacion & ";" & .Fecha & ";" & .Observacion
        End With
        Set m_ObjObServacion = Nothing
    Next
    If lst.ListCount = 1 Then
        lst.Selected(1) = True
        If p_Tipo = "Curso" Then
            ListaObservacionesCurso_Click
        ElseIf p_Tipo = "General" Then
            ListaObservaciones_Click
        End If
        
    End If

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerListasObservaciones ha devuelto el error: " & Err.Description
    End If
End Function

Private Sub ListaObservaciones_Click()
    
    Dim m_IDObservacionSeleccionada As String
    On Error GoTo errores
    
    m_Error = ""
    
    m_IDObservacionSeleccionada = Nz(Me.ListaObservaciones.Column(0), "")
    If m_IDObservacionSeleccionada = "" Then
        Me.ComandoObservacionEditar.Enabled = False
        Me.ComandoObservacionEliminar.Enabled = False
        Set m_ObjObservacionSeleccionada = Nothing
        Exit Sub
    End If
    
    If m_ObjObservacionSeleccionada Is Nothing Then
        Set m_ObjObservacionSeleccionada = getObservacion(m_IDObservacionSeleccionada, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_ObjObservacionSeleccionada.IDObservacion <> m_IDObservacionSeleccionada Then
            Set m_ObjObservacionSeleccionada = getObservacion(m_IDObservacionSeleccionada, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    Me.ComandoObservacionEditar.Enabled = True
    If Me.AllowEdits = True Then
        
        Me.ComandoObservacionEliminar.Enabled = True
    Else
        
        Me.ComandoObservacionEliminar.Enabled = False
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al hacer click en la lista se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ListaObservaciones_DblClick(Cancel As Integer)
    
    On Error Resume Next
    Application.Echo False
    If Not m_ObjObservacionSeleccionada Is Nothing Then
        If m_ObjObservacionSeleccionada.TIPO = "General" Then
            If Me.ComandoObservacionEditar.Enabled Then
                ComandoObservacionEditar_Click
            End If
        End If
        
    End If
    Application.Echo True
End Sub

Private Sub ListaObservacionesCurso_Click()
    
    Dim m_IDObservacionSeleccionada As String
    On Error GoTo errores
    
    m_Error = ""
    
    m_IDObservacionSeleccionada = Nz(Me.ListaObservacionesCurso.Column(0), "")
    If m_IDObservacionSeleccionada = "" Then
        Me.ComandoObservacionCursoEditar.Enabled = False
        Me.ComandoObservacionCursoEliminar.Enabled = False
        Set m_ObjObservacionSeleccionada = Nothing
        Exit Sub
    End If
    
    If m_ObjObservacionSeleccionada Is Nothing Then
        Set m_ObjObservacionSeleccionada = getObservacion(m_IDObservacionSeleccionada, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        If m_ObjObservacionSeleccionada.IDObservacion <> m_IDObservacionSeleccionada Then
            Set m_ObjObservacionSeleccionada = getObservacion(m_IDObservacionSeleccionada, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
        End If
    End If
    Me.ComandoObservacionCursoEditar.Enabled = True
    If Me.AllowEdits = True Then
        
        Me.ComandoObservacionCursoEliminar.Enabled = True
    Else
        
        Me.ComandoObservacionCursoEliminar.Enabled = False
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al hacer click en la lista se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ListaObservacionesCurso_DblClick(Cancel As Integer)
     On Error Resume Next
    Application.Echo False
    If Not m_ObjObservacionSeleccionada Is Nothing Then
        If m_ObjObservacionSeleccionada.TIPO = "Curso" Then
            If Me.ComandoObservacionCursoEditar.Enabled Then
                ComandoObservacionCursoEditar_Click
            End If
        End If
        
    End If
    Application.Echo True
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