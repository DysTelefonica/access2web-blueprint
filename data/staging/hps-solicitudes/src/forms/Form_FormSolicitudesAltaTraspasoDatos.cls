Option Compare Database
Option Explicit

Public Event Alta()
Public Event Edicion()
Private m_Justificacion As Justificacion
Private m_JServicio As JustificacionServicio
Private WithEvents m_FormResponsable As Form_FormResponsable
Private WithEvents m_FormCambioTipo As Form_FormCambioTipo
Private WithEvents m_FormExpedientes As Form_FormExpedientesBusqueda
Private WithEvents f_BusquedaJustif As Form_frmBuscarJustificacion
Private m_CorreosAutoAlInicio As String ' Variable para detectar el cambio de No -> Sí
Private m_Error As String

Private Sub cmdLimpiarJustificacion_Click()
    Me.txtIdJustificacion = Null
    Me.txtJustificacion = Null
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
    If IsNumeric(m_PalabraClave) Then
        Set m_Expediente = constructor.getExpediente(m_PalabraClave, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_Expediente Is Nothing Then
            Me.IDExpediente = m_Expediente.IDExpediente
            If m_Expediente.Titulo <> "" Then
                Me.Titulo = m_Expediente.Titulo
            Else
                Me.Titulo = m_Expediente.TextoExpediente
            End If
            Me.PalabraClave = Null
            VBA.DoEvents
            DoCmd.Hourglass False
            VBA.DoEvents
            Exit Sub
        End If
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

Private Sub ComandoVerDetalleExpediente_Click()
    On Error GoTo errores
    
   
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    Set m_ObjExpedienteActivo = constructor.getExpediente(p_IDExpediente:=Nz(Me.IDExpediente, ""), p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If FormularioAbierto("FormExpedienteDetalle") Then
        DoCmd.Close acForm, "FormExpedienteDetalle", acSaveNo
    End If
    DoCmd.OpenForm "FormExpedienteDetalle"
    
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoVerDetalleExpediente_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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
Public Sub ComandoActualizar_Click()
    On Error GoTo errores
    m_Error = ""
    DoCmd.Hourglass True
    
    Me.lblEstadoSolicitud.Caption = m_ObjSolicitudActiva.EstadoCalculadoTitulo
    DoCmd.Hourglass False
   
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "El método ComandoActualizar_Click ha devuelto el error: " & Err.Numbr & vbCrLf & "Detalle: " & Err.Description
    End If
    
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoCambioTipo_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormCambioTipo") Then
        DoCmd.Close acForm, "FormCambioTipo"
    End If
    
    DoCmd.OpenForm "FormCambioTipo"
    If FormularioAbierto("FormCambioTipo") Then
        Set m_FormCambioTipo = Forms("FormCambioTipo")
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoCambioTipo_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoMotivos_Click()
    Dim m_Motivo As String
    Dim cmb As ComboBox
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
    Set m_ObjEntorno.ColMotivos = Nothing
    Set cmb = Me.Motivo_HPS
    RellenarComboMotivos cmb, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Me.Motivo_HPS = m_Motivo
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoMotivos_Click se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub


Private Sub ComandoNuevoResponsable_Click()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    If FormularioAbierto("FormResponsable") Then
        DoCmd.Close acForm, "FormResponsable"
    End If
    Set m_ObjResposableActivo = Nothing
    DoCmd.OpenForm "FormResponsable"
    If FormularioAbierto("FormResponsable") Then
        Set m_FormResponsable = Forms("FormResponsable")
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoNuevoResponsable_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub



Private Sub Form_Load()
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    'Me.AllowEdits = False
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    'Me.AllowEdits = True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
   ' Me.AllowEdits = True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
Private Sub RellenarJustificacion(Optional id As Long, Optional p_Justificacion As Justificacion)
    Set m_JServicio = New JustificacionServicio
    Set m_Justificacion = m_JServicio.ObtenerPorID(CLng(id))
    If Not m_Justificacion Is Nothing Then
        Me.txtIdJustificacion = id
        Me.txtJustificacion = m_JServicio.getJustificacionParaCampo(p_Justificacion:=m_Justificacion)
    End If
End Sub
Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    
    
    On Error GoTo errores
    
    EstablecerCombos p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If m_ObjSolicitudActiva.IDSolicitud <> "" Then
        Set m_ObjSolicitudAlInicio = constructor.getSolicitud(p_IDSolicitud:=m_ObjSolicitudActiva.IDSolicitud, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    If m_ObjSolicitudAlInicio Is Nothing Then
        Me.Gestor = Capitalizar(m_ObjUsuarioConectado.Nombre)
        Me.TIPO = m_ObjSolicitudActiva.TIPO
        Me.CorreosAutomaticos.SetFocus
        Me.lblEstadoSolicitud.Visible = False
        Me.ComandoActualizar.Visible = False
        Me.ComandoCambioTipo.Enabled = False
        Exit Function
    End If
    
    
    
    With m_ObjSolicitudAlInicio
        If IsNumeric(.IDJustificacion) Then
            RellenarJustificacion id:=.IDJustificacion
        End If
        If .TIPO <> "" Then
            Me.TIPO = .TIPO
            If .TIPO = "Traspaso" Then
                Me.ComandoCambioTipo.Enabled = False
            Else
                Me.ComandoCambioTipo.Enabled = True
            End If
        Else
            Me.ComandoCambioTipo.Enabled = False
        End If
        If .CorreosAutomaticos <> "" Then
            Me.CorreosAutomaticos = .CorreosAutomaticos
        End If
        If Not .GestorObj Is Nothing Then
            Me.Gestor = Capitalizar(.GestorObj.Nombre)
        End If
        
        If .Nombre <> "" Then
            Me.Nombre = Capitalizar(.Nombre)
        End If
        If .Apellido1 <> "" Then
            Me.Apellido1 = Capitalizar(.Apellido1)
        End If
        If .Apellido2 <> "" Then
            Me.Apellido2 = Capitalizar(.Apellido2)
        End If
        If IsDate(.FNacimiento) Then
            Me.FNacimiento = .FNacimiento
        End If
        If .LugarNacimiento <> "" Then
            Me.LugarNacimiento = .LugarNacimiento
        End If
        If .email <> "" Then
            Me.email = .email
        End If
        If .emailResponsable <> "" Then
            Me.emailResponsable = .emailResponsable
        End If
        If .Telefono <> "" Then
            Me.Telefono = .Telefono
        End If
        If .IDEmpresaUsuario <> "" Then
            Me.IDEmpresaUsuario = .IDEmpresaUsuario
        End If
        If .IDEmpresaTramitadora <> "" Then
            Me.IDEmpresaTramitadora = .IDEmpresaTramitadora
        End If
        
        If Not .Expediente Is Nothing Then
            Me.ComandoVerDetalleExpediente.Enabled = True
            Me.IDExpediente = .Expediente.IDExpediente
            If .Expediente.Titulo <> "" Then
                Me.Titulo = .Expediente.Titulo
            Else
                Me.Titulo = .Expediente.TextoExpediente
            End If
            
        End If
        If .DNI <> "" Then
            Me.DNI = .DNI
        End If
        If .Motivo_HPS <> "" Then
            Me.Motivo_HPS = .Motivo_HPS
        End If
        If .Observaciones <> "" Then
            Me.Observaciones = .Observaciones
        End If
        If .IDSolicitud <> "" Then
            Me.IDSolicitud = .IDSolicitud
        End If
    End With
    Me.ComandoActualizar.Visible = True
    Me.lblEstadoSolicitud.Visible = True
    Me.lblEstadoSolicitud.Caption = m_ObjSolicitudAlInicio.EstadoCalculadoTitulo
    If FormularioAbierto("FormSolicitudesGestion") Then
        Form_FormSolicitudesGestion.Filtrar
    End If
    m_CorreosAutoAlInicio = Nz(Me.CorreosAutomaticos, "No")
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function
Public Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean
    
    
    On Error GoTo errores
    
    
    If m_ObjSolicitudAlInicio Is Nothing Then
        HaHabidoCambios = True
        Exit Function
    End If
    With m_ObjSolicitudAlInicio
        If .TIPO <> Nz(Me.TIPO, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .CorreosAutomaticos <> Nz(Me.CorreosAutomaticos, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Gestor <> Nz(Me.Gestor, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Nombre <> Nz(Me.Nombre, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Apellido1 <> Nz(Me.Apellido1, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Apellido2 <> Nz(Me.Apellido2, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .FNacimiento <> Nz(Me.FNacimiento, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .LugarNacimiento <> Nz(Me.LugarNacimiento, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .email <> Nz(Me.email, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .emailResponsable <> Nz(Me.emailResponsable, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Telefono <> Nz(Me.Telefono, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .IDEmpresaUsuario <> Nz(Me.IDEmpresaUsuario, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .IDEmpresaTramitadora <> Nz(Me.IDEmpresaTramitadora, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .IDExpediente <> Nz(Me.IDExpediente, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Motivo_HPS <> Nz(Me.Motivo_HPS, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .IDJustificacion <> Nz(Me.txtIdJustificacion, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .Observaciones <> Nz(Me.Observaciones, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If .DNI <> Nz(Me.DNI, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function
Public Function RellenarObjeto(Optional ByRef p_Error As String) As String
    
    
    On Error GoTo errores
    
    
   
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
        .IDJustificacion = Nz(Me.txtIdJustificacion, "")
        .Observaciones = Nz(Me.Observaciones, "")
        
    End With
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarObjeto ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function


' En Form_FormSolicitudesAltaTraspasoDatos.cls

Private Sub ComandoRegistrar_Click()
    Dim m_SolOp As SolicitudOperaciones
    Dim procesoIniciado As Boolean
    
    On Error GoTo errores
    m_Error = ""
    DoCmd.Hourglass True
    Me.AllowEdits = False
    
    ' 1. Guardar los datos
    Registrar EnumSiNo.Sí, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    ' 2. Lógica de Cambio de Estado en Correos Automáticos
    If m_ObjEntorno.Configuracion.CorreosAutomaticos = "Sí" Then
        
        ' CASO 1: DE "NO" A "SÍ"
        If (m_CorreosAutoAlInicio <> "Sí") And (Me.CorreosAutomaticos = "Sí") Then
            
            ' --- NUEVA VALIDACIÓN DE REGLA DE NEGOCIO ---
            ' Solo tiene sentido reprogramar si el proceso ya ha comenzado (se ha disparado algún evento)
            procesoIniciado = False
            If Not m_ObjSolicitudActiva.SolicitudFechas Is Nothing Then
                ' Miramos si existe alguno de los dos disparadores temporales
                If IsDate(m_ObjSolicitudActiva.SolicitudFechas.FechaEnvioExcel) Or _
                   IsDate(m_ObjSolicitudActiva.SolicitudFechas.FechaTramitacionAltaMarga) Then
                    procesoIniciado = True
                End If
            End If
            ' --------------------------------------------
            
            ' Solo preguntamos si el proceso ya está vivo
            If procesoIniciado Then
                DoCmd.Hourglass False
                
                If MsgBox("Ha activado los correos automáticos para esta solicitud." & vbNewLine & vbNewLine & _
                          "¿Desea REPROGRAMAR los correos pendientes basándose en las fechas de los trámites ya realizados?", _
                          vbQuestion + vbYesNo, "Sincronizar Correos Automáticos") = vbYes Then
                    
                    DoCmd.Hourglass True
                    Set m_SolOp = New SolicitudOperaciones
                    Set m_SolOp.solicitud = m_ObjSolicitudActiva
                    m_SolOp.ReprogramarAvisosFuturos m_Error
                    
                    If m_Error <> "" Then
                        MsgBox "Datos guardados, pero hubo un error reprogramando: " & vbNewLine & m_Error, vbExclamation
                        m_Error = ""
                    End If
                End If
                DoCmd.Hourglass True
            End If
            
        ' CASO 2: DE "SÍ" A "NO" (Limpieza)
        ElseIf (m_CorreosAutoAlInicio = "Sí") And (Me.CorreosAutomaticos = "No") Then
            
            Set m_SolOp = New SolicitudOperaciones
            Set m_SolOp.solicitud = m_ObjSolicitudActiva
            
            ' Aquí sí tiene sentido limpiar siempre, por si acaso hubiera basura programada
            If m_SolOp.AnularTodosLosProgramados(m_Error) Then
                MsgBox "Al desactivar los correos automáticos, se han eliminado las programaciones de envío pendientes para esta solicitud.", vbInformation, "Limpieza de Cola"
            Else
                MsgBox "Error al intentar limpiar los correos programados: " & m_Error, vbExclamation
                m_Error = ""
            End If
            
        End If
    End If

    ' 3. Finalización estándar
    Me.AllowEdits = True
    If FormularioAbierto("FormSolicitudAlta") Then
        DoCmd.Close acForm, "FormSolicitudAlta", acSaveNo
    End If
    
    DoCmd.Hourglass False
    MsgBox "Datos grabados correctamente", vbInformation, "Datos Grabados"
    
    m_CorreosAutoAlInicio = Nz(Me.CorreosAutomaticos, "No")
    
    Exit Sub
errores:
     If Err.Number <> 1000 Then
        m_Error = "El método ComandoRegistrar_Click ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    Me.AllowEdits = True
    DoCmd.Hourglass False
    MsgBox m_Error, vbCritical, "Error"
End Sub

Public Function EstablecerCombos(Optional ByRef p_Error As String) As String
    
    Dim cmb As ComboBox
    Dim m_ID As Variant
    
    Dim m_Col As Scripting.Dictionary
    On Error GoTo errores
    
    
    Dim m_Suministrador As Suministrador
    Dim m_Nombre As String
    Dim m_Responsable  As Responsable
    
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
    Set cmb = Me.emailResponsable
    cmb.RowSource = ""
    Set m_Col = m_ObjEntorno.Responsables
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        For Each m_ID In m_Col
             Set m_Responsable = m_Col(m_ID)
            cmb.AddItem m_Responsable.Correo & ";" & m_Responsable.Nombre & ";" & m_Responsable.Descripcion
            
            Set m_Responsable = Nothing
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
Public Function Registrar( _
                            Optional p_VerSiHayCambios As EnumSiNo = EnumSiNo.Sí, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_HaHabidoCambios As Boolean
    Dim m_SolicitudOp As SolicitudOperaciones
    On Error GoTo errores
    If p_VerSiHayCambios = EnumSiNo.Sí Then
        m_HaHabidoCambios = HaHabidoCambios(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_HaHabidoCambios Then
            p_Error = "No hay cambios que guardar"
            Err.Raise 1000
        End If
    End If
    RellenarObjeto p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_SolicitudOp = New SolicitudOperaciones
    With m_SolicitudOp
        Set .solicitud = m_ObjSolicitudActiva
        .RegistrarDatos m_ObjSolicitudAlInicio, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    If Me.Parent.Name = "FormSolicitudAltaTraspaso" Then
        Form_FormSolicitudAltaTraspaso.EstablecerDatos
    ElseIf Me.Parent.Name = "FormSolicitudRenovacion" Then
        Form_FormSolicitudRenovacion.EstablecerDatos
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Registrar ha producido el error n: " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If

End Function

Private Sub m_FormCambioTipo_TipoCambiado()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    DoCmd.Close acForm, (Me.Parent.Name), acSaveNo
    
    
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormCambioTipo_TipoCambiado se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub





Private Sub m_FormResponsable_Alta()
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Set m_ObjEntorno.Responsables = Nothing
    EstablecerCombos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjResposableActivo Is Nothing Then
        Me.emailResponsable = m_ObjResposableActivo.Correo
        
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al m_FormResponsable_Alta se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
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
        Me.ComandoVerDetalleExpediente.Enabled = False
    Else
        Me.ComandoActualizar.Enabled = True
        Me.ComandoVerDetalleExpediente.Enabled = True
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
Private Sub btnBuscarJustif_Click()
    ' ERROR COMÚN: No uses acDialog aquí si quieres usar WithEvents.
    ' CORRECTO: Lo abrimos normal, pero el formulario hijo debe tener sus propiedades configuradas.
    
    DoCmd.OpenForm "frmBuscarJustificacion", OpenArgs:="MODO_SELECCION"
    
    ' Como no usamos acDialog, el código pasa inmediatamente a esta línea
    ' y conecta el "cable" de los eventos.
    Set f_BusquedaJustif = Forms("frmBuscarJustificacion")
End Sub

' Capturamos el evento del componente
Private Sub f_BusquedaJustif_JustificacionElegida(id As Long, Titulo As String)
    ' Actualizamos los controles
    If IsNumeric(id) Then
        RellenarJustificacion id:=id
    End If
    
    ' Limpiamos la referencia
    Set f_BusquedaJustif = Nothing
    
    ' Si necesitas guardar inmediatamente o habilitar otros campos:
    ' Me.Dirty = False
End Sub