' ==========================================================================
' MÓDULO: Form_frm0PpalTecnico
' RESPONSABILIDAD: Menú principal para usuarios con el rol de Técnico.
' ==========================================================================
Option Compare Database
Option Explicit


Private Sub cmdTareasPendientes_Click()
    On Error GoTo Errores
    
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    
    ' --- INICIO DE LA CORRECCIÓN ---
    ' Este método ya era casi correcto, pero lo hacemos explícito y robusto.
    ' Abre el formulario y le pasa un argumento de texto simple.
    ' La lógica para interpretar "PENDIENTES" ya está en el Form_Load de frmBuscarSolicitudes.
    DoCmd.OpenForm "frmBuscarSolicitudes", OpenArgs:="PENDIENTES"
    ' --- FIN DE LA CORRECCIÓN ---

LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdTareasPendientes_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub ComandoAyuda_Click()
    Dim ruta As String
    On Error GoTo Errores
    
    ' --- CÓDIGO SIMPLIFICADO Y POLIMÓRFICO ---
    ruta = m_ObjEntorno.URLManual
    
    
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    
    If fso.FileExists(ruta) Then
        AbrirEnLocal ruta
    Else
        MsgBox "No se pudo obtener el documento de ayuda.", vbExclamation, "Error"
    End If
    
LimpiarYSalir:
    Exit Sub
Errores:
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Call ManejadorErroresFormulario("cmdAbrirAyuda_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub Form_Load()
    On Error GoTo Errores
    ' Al cargar el formulario, ajustamos la interfaz.
    Ajustar Me
    
    ' Actualizamos el título de la ventana con la información del usuario.
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    Me.lblVersion.Caption = m_ObjEntorno.VersionAplicacion
    ' Hacemos visible el botón para "dejar de suplantar"
    ' SOLO si se ha iniciado sesión en modo de suplantación.
    Me.cmdVolverASesionAdmin.Visible = g_blnImpersonando
    Me.cmdSalir.Visible = Not g_blnImpersonando
    Call ActualizarContadorTareas
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("Form_Load", Me)
End Sub


Private Sub cmdBuscarSolicitudes_Click()
    Dim filtros As FiltrosSolicitud
    Dim frmBusqueda As Form_frmBuscarSolicitudes
    
    On Error GoTo Errores
    
    ' --- INICIO DE LA CORRECCIÓN ---
    
    ' 1. Abrimos el formulario de manera estándar. Access ahora es responsable de su ciclo de vida.
    DoCmd.OpenForm "frmBuscarSolicitudes"
    
    ' 2. Obtenemos una referencia a la instancia que Access acaba de crear.
    Set frmBusqueda = Forms("frmBuscarSolicitudes")
    
    ' 3. Preparamos el filtro que queremos "inyectar".
    Set filtros = New FiltrosSolicitud
    filtros.idResponsableTecnico = m_ObjUsuarioActivo.id
    
    ' 4. Usamos la referencia para pasar el objeto de filtro a la propiedad pública del formulario de búsqueda.
    Set frmBusqueda.FiltroInicial = filtros
    
    ' 5. Forzamos una nueva búsqueda en el formulario hijo para que aplique el filtro inyectado.
    Call frmBusqueda.RealizarBusqueda
    
    ' --- FIN DE LA CORRECCIÓN ---
    
LimpiarYSalir:
    On Error Resume Next
    Set filtros = Nothing
    Set frmBusqueda = Nothing ' Ahora esto es seguro, solo liberamos nuestra variable local, no destruimos el formulario.
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdBuscarSolicitudes_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdVolverASesionAdmin_Click()
    On Error GoTo Errores
    
    ' Restauramos el usuario y el indicador de suplantación
    Set m_ObjUsuarioActivo = m_ObjUsuarioReal
    g_blnImpersonando = False
    
    ' --- INICIO DE LA CORRECCIÓN ---
    ' Re-evaluamos el rol usando nuestra nueva variable global.
    ' Esta lógica es robusta, segura y fácil de leer.
    rolUsuario = rolUsuarioReal
    ' --- FIN DE LA CORRECCIÓN ---
    
    ' Volvemos al formulario principal de Admin/Calidad y cerramos este
    DoCmd.OpenForm "frm0Ppal"
    DoCmd.Close acForm, Me.name
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdVolverASesionAdmin_Click", Me)
    Exit Sub
End Sub

Private Sub cmdSalir_Click()
    On Error GoTo Errores
    Call CerrarConexionesGlobales ' <-- AÑADIR AQUÍ
    Application.Quit acQuitSaveNone

LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdSalir_Click", Me)
    Exit Sub
End Sub
' ==========================================================================
' MÓDULO: Form_frm0PpalTecnico.cls
' ACCIÓN: REEMPLAZO para centralizar la lógica del contador en el servicio.
' ==========================================================================
Public Sub ActualizarContadorTareas()
    Dim solServ As New SolicitudServicio
    Dim colPendientes As Scripting.Dictionary
    Dim contador As Long
    
    On Error GoTo Errores
    
    Me.lblContadorTareas.Visible = False
    
    ' --- INICIO DE LA LÓGICA SIMPLIFICADA ---
    
    ' Llamamos al servicio forzando AMBOS filtros: "solo pendientes" Y "solo mis solicitudes".
    ' El servicio se encargará de aplicar la lógica correcta para el rol Técnico:
    ' 1. Filtrará por estado = DESARROLLO TÉCNICO (por 'soloPendientes').
    ' 2. Filtrará por el ID del usuario técnico activo (por 'soloMisSolicitudes').
    Set colPendientes = solServ.getSolicitudesViewModel(soloPendientes:=True, soloMisSolicitudes:=True)
    
    ' --- FIN DE LA LÓGICA SIMPLIFICADA ---
    
    If Not colPendientes Is Nothing Then
        contador = colPendientes.count
        If contador > 0 Then
            Me.lblContadorTareas.Caption = contador
            Me.lblContadorTareas.Visible = True
        End If
    End If

LimpiarYSalir:
    Exit Sub
Errores:
    Me.lblContadorTareas.Visible = False
    Resume LimpiarYSalir
End Sub

Private Sub lblContadorTareas_Click()
    cmdTareasPendientes_Click
End Sub

Private Sub shpSolicitudes_Click()
    cmdBuscarSolicitudes_Click
End Sub

Private Sub shpTareas_Click()
    cmdTareasPendientes_Click
End Sub