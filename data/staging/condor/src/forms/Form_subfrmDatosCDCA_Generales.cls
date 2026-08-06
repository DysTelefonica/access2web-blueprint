
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_subfrmDatosCDCA_Generales
' ACCIÓN: Implementación de Carga de Defaults con Restauración y UI Simétrica.
' ==========================================================================

Private m_ViewModel As DatosCDCAViewModel
Private m_original_numContrato As String
Private m_original_refSuministrador As String
Private m_original_SuministradorNombreDir As String
Private m_original_refDesviacionesPrevias As String

' --- COLORES EXPLÍCITOS (no dependen del diseño) ---
Private Const COLOR_FONDO_EDITABLE As Long = 16774386   ' Blanco
Private Const COLOR_FONDO_INMUTABLE As Long = 16774463  ' Gris
Private Const COLOR_TEXTO_NORMAL As Long = 4666411      ' Gris oscuro

' --- MÉTODOS DE POBLADO Y RECOGIDA ---

' ==========================================================================
' Aplica estado visual explícito según regla de negocio.
' No depende del valor por defecto en diseño.
' ==========================================================================
Private Sub AplicarEstadoVisualControl(ctl As Control, puedeEditar As Boolean)
    If ctl Is Nothing Then Exit Sub
    If ctl.Tag = "INMUTABLE" Then
        ctl.Enabled = False
        ctl.BackColor = COLOR_FONDO_INMUTABLE
        ctl.ForeColor = COLOR_TEXTO_NORMAL
    Else
        ctl.Enabled = puedeEditar
        If puedeEditar Then
            ctl.BackColor = COLOR_FONDO_EDITABLE
            ctl.ForeColor = COLOR_TEXTO_NORMAL
        Else
            ctl.BackColor = COLOR_FONDO_INMUTABLE
            ctl.ForeColor = COLOR_TEXTO_NORMAL
        End If
    End If
End Sub

Public Sub PoblarControles(ByRef vm As DatosCDCAViewModel)
    On Error GoTo Errores
    Set m_ViewModel = vm
    If vm Is Nothing Then Exit Sub
    
    ' CORRECCIÓN: Habilitación y Visibilidad simétrica con Guardar
    Dim puedeEditar As Boolean
    puedeEditar = vm.PuedeGuardar_Generales
    
    Me.numContrato.value = Nz(vm.vm_numContrato, "")
    Call AplicarEstadoVisualControl(Me.numContrato, puedeEditar)
    Me.refSuministrador.value = Nz(vm.vm_refSuministrador, "")
    Call AplicarEstadoVisualControl(Me.refSuministrador, puedeEditar)
    Me.SuministradorNombreDir.value = Nz(vm.vm_SuministradorNombreDir, "")
    Call AplicarEstadoVisualControl(Me.SuministradorNombreDir, puedeEditar)
    Me.refDesviacionesPrevias.value = Nz(vm.vm_refDesviacionesPrevias, "")
    Call AplicarEstadoVisualControl(Me.refDesviacionesPrevias, puedeEditar)
    
    ' Gestión segura del Tri-estado (Sí/No/Null)
    If IsNull(vm.vm_requiereModificacionContrato) Then
        Me.grpModContrato.value = Null
    ElseIf vm.vm_requiereModificacionContrato = True Then
        Me.grpModContrato.value = 1
    Else
        Me.grpModContrato.value = 2
    End If
    
    ' CORRECCIÓN: Habilitación y Visibilidad simétrica con Guardar
    ' (puedeEditar ya definido arriba para usar antes)
    
    ' Botón Cargar Defaults
    Me.cmdCargarDefaults.Visible = True
    Me.cmdCargarDefaults.Enabled = puedeEditar
    
    ' Botón Guardar
    Me.cmdGuardar.Enabled = puedeEditar
    
    If Me.cmdCargarDefaults.Enabled And Me.cmdCargarDefaults.Visible Then
        Me.cmdCargarDefaults.SetFocus
    End If
    
    ' Botón Restaurar
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("PoblarControles", Me)
    Resume LimpiarYSalir
End Sub

Public Sub RecogerValores(ByRef vm As DatosCDCAViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub
    
    vm.vm_numContrato = Nz(Me.numContrato.value, "")
    vm.vm_refSuministrador = Nz(Me.refSuministrador.value, "")
    vm.vm_SuministradorNombreDir = Nz(Me.SuministradorNombreDir.value, "")
    vm.vm_refDesviacionesPrevias = Nz(Me.refDesviacionesPrevias.value, "")
    
    Select Case Nz(Me.grpModContrato.value, 0)
        Case 1: vm.vm_requiereModificacionContrato = True
        Case 2: vm.vm_requiereModificacionContrato = False
        Case Else: vm.vm_requiereModificacionContrato = Null
    End Select
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("RecogerValores", Me)
    Resume LimpiarYSalir
End Sub

' --- EVENTOS DE BOTONES ---

Private Sub cmdGuardar_Click()
    On Error GoTo Errores
    CommitEditsOnForm Me
    Me.Parent.GuardarDesdeSubform Me.name
    
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False
    Me.cmdCargarDefaults.Enabled = m_ViewModel.PuedeGuardar_Generales
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdGuardar_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdCargarDefaults_Click()
    Dim cdcaServ As New DatosCDCAServicio
    Dim cdcaInicial As DatosCDCA
    On Error GoTo Errores
    If m_ViewModel Is Nothing Then Exit Sub
    
    Me.Parent.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    
    ' 1. Guardar Snapshot (BLINDADO con Nz)
    m_original_numContrato = Nz(Me.numContrato.value, "")
    m_original_refSuministrador = Nz(Me.refSuministrador.value, "")
    m_original_SuministradorNombreDir = Nz(Me.SuministradorNombreDir.value, "")
    m_original_refDesviacionesPrevias = Nz(Me.refDesviacionesPrevias.value, "")
    
    ' 2. Cargar valor por defecto desde el servicio
    Set cdcaInicial = cdcaServ.GetDatosGeneralesIniciales(m_ViewModel.Solicitud.idSolicitud)
    If Not cdcaInicial Is Nothing Then
        Me.numContrato.value = cdcaInicial.numContrato
        Me.refSuministrador.value = cdcaInicial.refSuministrador
        Me.SuministradorNombreDir.value = cdcaInicial.SuministradorNombreDir
    End If
    
    Call RecogerValores(m_ViewModel)
    
    Me.Parent.ConfigurarYPoblarInterfaz
    Me.Parent.PoblarSubformActual
    
    ' 3. Gestionar UI
    Me.cmdCargarDefaults.Enabled = False
    Me.cmdRestaurarDatos.Visible = True
    Me.cmdRestaurarDatos.Enabled = True
    
LimpiarYSalir:
    On Error Resume Next
    Me.Parent.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Set cdcaServ = Nothing
    Set cdcaInicial = Nothing
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("subfrmDatosCDCA_Generales.cmdCargarDefaults_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdRestaurarDatos_Click()
    On Error GoTo Errores
    
    ' Restaurar Snapshot
    Me.numContrato.value = m_original_numContrato
    Me.refSuministrador.value = m_original_refSuministrador
    Me.SuministradorNombreDir.value = m_original_SuministradorNombreDir
    Me.refDesviacionesPrevias.value = m_original_refDesviacionesPrevias
    
    Call RecogerValores(m_ViewModel)
    
    Me.Parent.ConfigurarYPoblarInterfaz
    Me.Parent.PoblarSubformActual
    
    ' Restaurar botones
    Me.cmdCargarDefaults.Enabled = True
    If Me.cmdCargarDefaults.Enabled And Me.cmdCargarDefaults.Visible Then
        Me.cmdCargarDefaults.SetFocus
    End If
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdRestaurarDatos_Click", Me)
    Resume LimpiarYSalir
End Sub