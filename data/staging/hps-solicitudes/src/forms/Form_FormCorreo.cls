
Option Compare Database
Option Explicit

Public Event Alta()
Public Event Editado()
Private m_CorreoAlInicio As Correo
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
    MsgBox m_Error, vbCritical, "Error"
End Sub

Private Sub ComandoRegistrar_Click()
    
    Dim m_HaHabidoCambios As Boolean
    Dim m_CorreoOp As CorreoOperaciones
    
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    
    m_Error = ""
    
    ' Si estamos en MODO GESTION y ya se envió, este botón debería estar deshabilitado,
    ' pero por seguridad:
    If Nz(Me.OpenArgs, "") = "MODO_GESTION" And IsDate(m_CorreoAlInicio.FechaEnvio) Then
        MsgBox "Este correo ya fue enviado. Use el botón 'Volver a Enviar' para generar uno nuevo.", vbInformation
        DoCmd.Hourglass False
        Exit Sub
    End If

    m_HaHabidoCambios = HaHabidoCambios(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_HaHabidoCambios = False Then
        m_Error = "No hay cambios que guardar"
        Err.Raise 1000
    End If
    
    ' Actualizamos el objeto activo con los valores del formulario
    If m_ObjCorreoActivo Is Nothing Then
        Set m_ObjCorreoActivo = New Correo
    End If
    With m_ObjCorreoActivo
        .Asunto = Nz(Me.Asunto, "")
        
        ' Gestión inteligente de la fecha programada
        If IsDate(Me.FechaOrdenEnvio) Then
            .FechaOrdenEnvio = Me.FechaOrdenEnvio
            .Programado = "Sí"
        Else
            .FechaOrdenEnvio = ""
            .Programado = "No" ' Si borran la fecha, pasa a envío manual inmediato o pendiente
        End If
        
        .FechaEnvio = Nz(Me.FechaEnvio, "")
        .DESTINATARIOS = Nz(Me.DESTINATARIOS, "")
        .DestinatariosConCopia = Nz(Me.DestinatariosConCopia, "")
    End With
    
    Set m_CorreoOp = New CorreoOperaciones
    With m_CorreoOp
        Set .Correo = m_ObjCorreoActivo
        .Registrar p_CorreoAlInicio:=m_CorreoAlInicio, p_Error:=m_Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        
        ' Si cambiamos la fecha de orden, asegúrate de que se actualiza en BD explícitamente si Registrar no lo hiciera
        .CambiarFechaOrdenEnvio m_Error
    End With
    
    If m_CorreoAlInicio Is Nothing Then
        RaiseEvent Alta
    Else
        RaiseEvent Editado
    End If
    
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
    MsgBox m_Error, vbCritical, "Error"
End Sub

' Botón NUEVO para reenviar correos ya procesados
Private Sub ComandoReenviar_Click()
    Dim m_NuevoCorreo As Correo
    Dim m_CorreoOp As CorreoOperaciones
    
    On Error GoTo errores
    
    If MsgBox("¿Está seguro de que desea generar un NUEVO envío para este correo?", vbQuestion + vbYesNo, "Confirmar Reenvío") = vbNo Then Exit Sub
    
    DoCmd.Hourglass True
    
    ' 1. Clonamos los datos del correo actual a uno nuevo
    Set m_NuevoCorreo = New Correo
    With m_NuevoCorreo
        .Asunto = m_CorreoAlInicio.Asunto
        .DESTINATARIOS = m_CorreoAlInicio.DESTINATARIOS
        .DestinatariosConCopia = m_CorreoAlInicio.DestinatariosConCopia
        .DestinatariosConCopiaOculta = m_CorreoAlInicio.DestinatariosConCopiaOculta
        .Aplicacion = m_CorreoAlInicio.Aplicacion
        .Accion = m_CorreoAlInicio.Accion
        .IDSolicitud = m_CorreoAlInicio.IDSolicitud
        .nombrePlantilla = m_CorreoAlInicio.nombrePlantilla
        .VersionPlantilla = m_CorreoAlInicio.VersionPlantilla
        .cadenaRecursos = m_CorreoAlInicio.cadenaRecursos
        .URLAdjunto = m_CorreoAlInicio.URLAdjunto
        .tipoCorreo = m_CorreoAlInicio.tipoCorreo
        
        ' IMPORTANTE: Reseteamos fechas para que salga YA
        .FechaOrdenEnvio = ""
        .FechaEnvio = ""
        .Programado = "No"
        .DesencadenadoPor = m_ObjUsuarioConectado.UsuarioRed & " (Reenvío)"
    End With
    
    ' 2. Lo registramos como nuevo
    Set m_CorreoOp = New CorreoOperaciones
    Set m_CorreoOp.Correo = m_NuevoCorreo
    m_CorreoOp.Registrar , m_Error
    
    If m_Error <> "" Then Err.Raise 1000
    
    DoCmd.Hourglass False
    MsgBox "Se ha generado un nuevo correo en la cola de envío.", vbInformation
    
    RaiseEvent Alta ' Para refrescar la lista
    DoCmd.Close acForm, Me.Name, acSaveNo
    Exit Sub

errores:
    DoCmd.Hourglass False
    MsgBox "Error al reenviar: " & Err.Description & vbNewLine & m_Error, vbCritical
End Sub

Public Function HaHabidoCambios(Optional ByRef p_Error As String) As Boolean
    
    On Error GoTo errores
    If m_CorreoAlInicio Is Nothing Then
        HaHabidoCambios = True
        Exit Function
    End If
    With m_CorreoAlInicio
        
        If .Asunto <> Nz(Me.Asunto, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        
        ' Comparar fechas teniendo en cuenta nulos
        Dim fOrdenUI As String
        fOrdenUI = Nz(Me.FechaOrdenEnvio, "")
        If .FechaOrdenEnvio <> fOrdenUI Then
            HaHabidoCambios = True
            Exit Function
        End If
        
        If .FechaEnvio <> Nz(Me.FechaEnvio, "") Then
            HaHabidoCambios = True
            Exit Function
        End If
        If Nz(Me.DESTINATARIOS, "") <> .DESTINATARIOS Then
            HaHabidoCambios = True
            Exit Function
        End If
        If Nz(Me.DestinatariosConCopia, "") <> .DestinatariosConCopia Then
            HaHabidoCambios = True
            Exit Function
        End If
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha producido el error num " & Err.Number & _
        vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function EstablecerDatos(Optional ByRef p_Error As String) As String
    
    On Error GoTo errores
    
    Set m_CorreoAlInicio = constructor.getCorreo(p_IDCorreo:=m_ObjCorreoActivo.IDCORREO, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_CorreoAlInicio Is Nothing Then
        p_Error = "No se ha podido encontrar la Correo registrado"
        Err.Raise 1000
    End If
   
    With m_CorreoAlInicio
        If .Asunto <> "" Then Me.Asunto = .Asunto
        If .Accion <> "" Then Me.Accion = .Accion
        If .FechaOrdenEnvio <> "" Then Me.FechaOrdenEnvio = .FechaOrdenEnvio
        If .FechaEnvio <> "" Then Me.FechaEnvio = .FechaEnvio
        If .DESTINATARIOS <> "" Then Me.DESTINATARIOS = .DESTINATARIOS
        If .DestinatariosConCopia <> "" Then Me.DestinatariosConCopia = .DestinatariosConCopia
    End With
    
    ' APLICAR LÓGICA DE VISUALIZACIÓN SEGÚN MODO
    AplicarReglasNegocioVisuales
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerDatos ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Private Sub AplicarReglasNegocioVisuales()
    ' Detectamos si venimos desde el formulario de Gestión de Correos
    If Nz(Me.OpenArgs, "") = "MODO_GESTION" Then
        
        ' --- CAMBIO SOLICITADO ---
        ' Desactivamos el botón de registrar/guardar al venir de gestión
        Me.ComandoRegistrar.Enabled = False
        ' -------------------------
        
        Me.ComandoReenviar.Visible = False
        
        ' CASO A: YA ENVIADO
        If IsDate(m_CorreoAlInicio.FechaEnvio) Then
            BloquearControles True
            ' Permitimos reenviar (crear copia) si ya fue enviado
            Me.ComandoReenviar.Visible = True
            Me.FechaOrdenEnvio.Enabled = False
            
        ' CASO B: NO ENVIADO (Pendiente o Programado)
        Else
            ' Como desactivamos el botón de registrar, por coherencia
            ' bloqueamos también la edición de la fecha, ya que no se podrían guardar los cambios.
            Me.FechaOrdenEnvio.Enabled = False
        End If
        
    Else
        ' Modo normal (abierto desde otros flujos):
        ' El estado del botón dependerá de lo que se haya definido en Form_Open
        ' (por ejemplo, si es Técnico o no).
        Me.ComandoReenviar.Visible = False
    End If
End Sub

Private Sub BloquearControles(bloquear As Boolean)
    On Error Resume Next
    Me.Asunto.Locked = bloquear
    Me.DESTINATARIOS.Locked = bloquear
    Me.DestinatariosConCopia.Locked = bloquear
    Me.Accion.Locked = bloquear
    Me.FechaEnvio.Locked = bloquear
    Me.FechaOrdenEnvio.Locked = bloquear
End Sub

Private Sub ComandoVerHTML_Click()
    Dim m_NombrePlantilla As String
    Dim m_URLPlantilla As String
    
    On Error GoTo errores
    
    ' 1. Preparar la interfaz
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    ' 2. Validaciones de seguridad
    If m_CorreoAlInicio Is Nothing Then
        m_Error = "No hay datos del correo cargados para poder previsualizar."
        Err.Raise 1000
    End If
    
    ' 3. Obtener la plantilla adecuada según el tipo de correo
    ' El objeto Correo ya guarda qué plantilla se usó al crearse (ej: RECORDATORIO_EXCEL1.html)
    m_NombrePlantilla = m_CorreoAlInicio.nombrePlantilla
    
    If m_NombrePlantilla = "" Then
        ' Si por algún motivo antiguo no tuviera plantilla, intentamos deducirla del Tipo (Safety net)
        If Not m_CorreoAlInicio.solicitud Is Nothing And m_CorreoAlInicio.tipoCorreo <> "" Then
             m_NombrePlantilla = m_CorreoAlInicio.solicitud.NombrePlantillaParaCorreoCalculada(m_CorreoAlInicio.tipoCorreo)
        End If
        
        If m_NombrePlantilla = "" Then
            m_Error = "No se ha podido determinar la plantilla HTML asociada a este correo."
            Err.Raise 1000
        End If
    End If
    
    ' 4. Construir la ruta completa a la plantilla
    ' Usamos la ruta versionada del entorno (donde está la carpeta IMG también)
    m_URLPlantilla = m_ObjEntorno.URLDirectorioPlantillasHTMLVersionEstablecida & m_NombrePlantilla
    
    If Not fso.FileExists(m_URLPlantilla) Then
        m_Error = "No se encuentra el archivo de plantilla en la ruta:" & vbNewLine & m_URLPlantilla
        Err.Raise 1000
    End If
    
    ' 5. Generar y Visualizar
    ' Llamamos a la función que ya actualizamos para copiar las imágenes y generar el HTML temporal
    ' Pasamos el ID de solicitud para que el motor rellene las variables (||Gestor||, etc.)
    ObtenerHTMLEnLocal p_URLPlantillaHTML:=m_URLPlantilla, _
                       p_IDSolicitud:=m_CorreoAlInicio.IDSolicitud, _
                       p_Error:=m_Error
                       
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    ' Finalizar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    
    Exit Sub

errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Error al previsualizar HTML: " & Err.Description
    End If
    MsgBox m_Error, vbCritical, "Error Visualización"
End Sub

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo errores
    
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    
    Ajustar Me
    Me.ComandoRegistrar.Enabled = False
    Me.AllowEdits = True ' Permitimos edits para poder usar los controles, pero bloqueamos selectivamente
    Me.Caption = m_TituloFormulario
    
    ' Por defecto técnico no edita, pero si es MODO_GESTION, la lógica de AplicarReglas manda
    If EsTecnico <> EnumSiNo.Sí Then
        Me.ComandoRegistrar.Enabled = True
    End If
    
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    AvanceCerrar
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha devuelto el error: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    MsgBox m_Error, vbCritical, "Error"
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
