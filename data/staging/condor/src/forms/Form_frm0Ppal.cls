
Option Compare Database
Option Explicit

Private WithEvents m_frmAlta As Form_frmAltaSolicitud


Private Sub cmdAbrirAyuda_Click()
    Dim ruta As String
    On Error GoTo Errores
    
    ' --- CÓDIGO SIMPLIFICADO Y POLIMÓRFICO ---
    ruta = m_ObjEntorno.URLManual
    
    
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    
    If fso.FileExists(ruta) Then
        AbrirEnLocal ruta
    Else
        MsgBox "No se pudo obtener el manual de uso.", vbExclamation, "Error"
    End If
    
LimpiarYSalir:
    Exit Sub
Errores:
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Call ManejadorErroresFormulario("cmdAbrirAyuda_Click", Me)
    Resume LimpiarYSalir
End Sub

' ==========================================================================
' MÓDULO: Form_frm0Ppal.cls (VERSIÓN FINAL CORREGIDA - ByVal en Event Handler)
' ==========================================================================

Private Sub cmdAdminOpciones_Click()
    On Error GoTo Errores
    DoCmd.OpenForm "frm0OtrosAdmin"
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdAdminOpciones_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdBuscarSolicitudes_Click()
    On Error GoTo Errores
    Me.lblBloqueo.Visible = True
    Me.lblBloqueo.Caption = "Abriendo buscador..."
    DoCmd.Hourglass True
    DoEvents
    
    DoCmd.OpenForm "frmBuscarSolicitudes"
    
LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdBuscarSolicitudes_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdSalir_Click()
    Dim pregunta As Long
    On Error GoTo Errores
    
    pregunta = MsgBox("¿Desea realmente salir de la aplicación?", vbExclamation + vbYesNo, "Salir de la aplicación")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    Call CerrarConexionesGlobales ' <-- AÑADIR AQUÍ
    Application.Quit acQuitSaveNone

LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdSalir_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdTareasPendientes_Click()
    On Error GoTo Errores
    Me.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    
    ' Simplemente abre el formulario de búsqueda pasándole el argumento "PENDIENTES"
    DoCmd.OpenForm "frmBuscarSolicitudes", OpenArgs:="PENDIENTES"

LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdTareasPendientes_Click", Me)
    GoTo LimpiarYSalir
End Sub

Private Sub cmdNuevaSolicitud_Click()
    On Error GoTo Errores
    
    If FormularioAbierto("frmAltaSolicitud") Then
        Forms("frmAltaSolicitud").SetFocus
    End If
    
    ' Si no, lo abrimos con DoCmd y asignamos la referencia.
    DoCmd.OpenForm "frmAltaSolicitud"
    Set m_frmAlta = Forms("frmAltaSolicitud")
    ' --- FIN DE LA REFACTORIZACIÓN ---

LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdNuevaSolicitud_Click", Me)
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
    
    Me.lblBloqueo.Visible = True
    Me.lblBloqueo.Caption = "Cargando entorno..."
    DoCmd.Hourglass True
    DoEvents
    
    Ajustar Me
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    Me.cmdAdminOpciones.Visible = (rolUsuario = rol.Administrador)
    Me.lblVersion.Caption = m_ObjEntorno.VersionAplicacion
    Call LlenarComboDesdeDiccionario(Me.cboUsuariosTecnicos, m_ObjEntorno.ResponsablesTecnicos, "Id", "Nombre", False)
    Call ActualizarContadorTareas
LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("Form_Load", Me)
    Application.Quit acQuitSaveAll
    Resume LimpiarYSalir
End Sub

Private Sub cmdImpersonar_Click()
    Dim idTecnico As Long
    Dim objTecnico As usuario
    Dim userServicio As New UsuarioServicio
    On Error GoTo Errores
    
    idTecnico = CLng(Nz(Me.cboUsuariosTecnicos.value, 0))
    If idTecnico = 0 Then
        MsgBox "Por favor, seleccione un usuario técnico de la lista.", vbInformation, "Condor"
        Exit Sub
    End If
    
    Me.lblBloqueo.Visible = True
    Me.lblBloqueo.Caption = "Cambiando a sesión de técnico..."
    DoCmd.Hourglass True
    DoEvents
    
    Set objTecnico = userServicio.getUsuarioConPermisos(p_ID:=CStr(idTecnico))
    
    If objTecnico Is Nothing Then
        Err.Raise 513, "cmdImpersonar_Click", "No se pudo encontrar al usuario seleccionado."
    End If
    
    Set m_ObjUsuarioActivo = objTecnico
    rolUsuario = rol.Tecnico
    g_blnImpersonando = True
    
    DoCmd.OpenForm "frm0PpalTecnico"
    DoCmd.Close acForm, Me.name
    
LimpiarYSalir:
    On Error Resume Next
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Set userServicio = Nothing
    Set objTecnico = Nothing
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdImpersonar_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub lblVideo_Click()
    On Error GoTo Errores
    
    Me.lblBloqueo.Visible = True
    Me.lblBloqueo.Caption = "Abriendo vídeo..."
    DoCmd.Hourglass True
    DoEvents
    
    Ejecutar Me.hwnd, "open", m_ObjEntorno.URLVideo, "", "", 1
    
LimpiarYSalir:
    Me.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("lblVideo_Click", Me)
    Resume LimpiarYSalir
End Sub
' ==========================================================================
' MÓDULO: Form_frm0Ppal.cls
' ACCIÓN: REEMPLAZO para que el contador use la definición correcta de tareas.
' ==========================================================================
Public Sub ActualizarContadorTareas()
    Dim solServ As New SolicitudServicio
    Dim colPendientes As Scripting.Dictionary
    Dim contador As Long
    
    On Error GoTo Errores
    
    Me.lblContadorTareas.Visible = False
    
    ' Llamamos al servicio forzando el filtro de "solo pendientes".
    ' El servicio ya contiene la lógica correcta para determinar qué estados
    ' son pendientes para el rol de Calidad/Admin.
    Set colPendientes = solServ.getSolicitudesViewModel(soloPendientes:=True)
    
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


Private Sub m_frmAlta_SolicitudCreadaConExito(ByVal idNuevaSolicitud As Long, ByVal tipoSolicitud As String)
    Dim frmGestion As Form_frmGestionSolicitud
    
    On Error GoTo Errores
    
    ' 1. Limpieza de referencia al formulario de alta
    Set m_frmAlta = Nothing
    
    ' 2. Actualizar contadores del menú (para que sume la nueva tarea)
    Call ActualizarContadorTareas
    
    ' 3. Mensaje de Confirmación (Tu petición anterior)
    MsgBox "Solicitud creada correctamente." & vbCrLf & vbCrLf & _
           "A continuación se abrirá el entorno de gestión para que pueda completar los datos.", _
           vbInformation, "Creación Exitosa"
    
    ' 4. ABRIR EL ORQUESTADOR (frmGestionSolicitud)
    ' Primero abrimos el formulario "madre"
    DoCmd.OpenForm "frmGestionSolicitud", acNormal, , , , acWindowNormal, CStr(idNuevaSolicitud)
    
    ' 5. ABRIR EL FORMULARIO DE DATOS (frmDatosPC / CDCA)
    ' Una vez abierto el gestor, obtenemos una referencia y le pedimos que abra los datos.
    ' Esto mantiene intacta la cadena de eventos (WithEvents) dentro de frmGestionSolicitud.
    If FormularioAbierto("frmGestionSolicitud") Then
        Set frmGestion = Forms("frmGestionSolicitud")
        frmGestion.AbrirFormularioDeDatos
    End If
    
LimpiarYSalir:
    Exit Sub
Errores:
    Set m_frmAlta = Nothing
    Call ManejadorErroresFormulario("m_frmAlta_SolicitudCreadaConExito", Me)
    Resume LimpiarYSalir
End Sub

Private Sub m_frmAlta_Close()
    On Error Resume Next
    Set m_frmAlta = Nothing
End Sub

Private Sub shpNueva_Click()
    cmdNuevaSolicitud_Click
End Sub

Private Sub shpSolicitudes_Click()
    cmdBuscarSolicitudes_Click
End Sub

Private Sub shpTareas_Click()
    cmdTareasPendientes_Click
End Sub
