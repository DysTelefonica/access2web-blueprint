
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_subfrmDatosCDCASUB_Generales
' RESPONSABILIDAD: Gestión de datos generales para CD-CA-SUB.
'                  Implementa patrón de Carga de Defaults y Restauración.
' ==========================================================================

Private m_ViewModel As DatosCDCASUBViewModel

' Variables para snapshot (Restauración de datos)
Private m_original_refSuministrador As String
Private m_original_suministradorPrincipal As String
Private m_original_subSuministrador As String
Private m_original_refDesviacionesPrevias As String

' --- COLORES EXPLÍCITOS (no dependen del diseño) ---
Private Const COLOR_FONDO_EDITABLE As Long = 16774386   ' Blanco
Private Const COLOR_FONDO_INMUTABLE As Long = 16774463  ' Gris
Private Const COLOR_TEXTO_NORMAL As Long = 4666411      ' Gris oscuro

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

Public Sub PoblarControles(ByRef vm As DatosCDCASUBViewModel)
    On Error GoTo Errores
    Set m_ViewModel = vm
    If vm Is Nothing Then Exit Sub
    
    ' 2. Gestión de Estados de Botones (Simetría con PC/CDCA)
    Dim puedeEditar As Boolean
    puedeEditar = vm.PuedeGuardar_Generales
    
    ' 1. Volcar datos con estado visual
    Me.refSuministrador.value = Nz(vm.vm_refSuministrador, "") ' Manual
    Call AplicarEstadoVisualControl(Me.refSuministrador, puedeEditar)
    Me.refSubSuministrador.value = Nz(vm.vm_refSubSuministrador, "") ' Read-only
    Call AplicarEstadoVisualControl(Me.refSubSuministrador, False) ' Siempre deshabilitado
    
    Me.suministradorPrincipalNombreDir.value = Nz(vm.vm_suministradorPrincipalNombreDir, "")
    Call AplicarEstadoVisualControl(Me.suministradorPrincipalNombreDir, puedeEditar)
    Me.subSuministradorNombreDir.value = Nz(vm.vm_subSuministradorNombreDir, "")
    Call AplicarEstadoVisualControl(Me.subSuministradorNombreDir, puedeEditar)
    Me.refDesviacionesPrevias.value = Nz(vm.vm_refDesviacionesPrevias, "")
    Call AplicarEstadoVisualControl(Me.refDesviacionesPrevias, puedeEditar)
    
    If IsNull(vm.vm_requiereModificacionContrato) Then
        Me.grpModContrato.value = Null
    ElseIf vm.vm_requiereModificacionContrato Then
        Me.grpModContrato.value = 1
    Else
        Me.grpModContrato.value = 2
    End If
    
    Me.cmdGuardar.Enabled = puedeEditar
    
    ' Botón Cargar Defaults
    Me.cmdCargarDefaults.Visible = True
    Me.cmdCargarDefaults.Enabled = puedeEditar
    
    ' Botón Restaurar (oculto por defecto)
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False
    
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("subfrmDatosCDCASUB_Generales.PoblarControles", Me)
End Sub

Public Sub RecogerValores(ByRef vm As DatosCDCASUBViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub
    
    vm.vm_refSuministrador = Nz(Me.refSuministrador.value, "")
    ' refSubSuministrador es readonly, no hace falta recogerlo, ya está en el VM
    
    vm.vm_suministradorPrincipalNombreDir = Nz(Me.suministradorPrincipalNombreDir.value, "")
    vm.vm_subSuministradorNombreDir = Nz(Me.subSuministradorNombreDir.value, "")
    vm.vm_refDesviacionesPrevias = Nz(Me.refDesviacionesPrevias.value, "")
    
    Select Case Nz(Me.grpModContrato.value, 0)
        Case 1: vm.vm_requiereModificacionContrato = True
        Case 2: vm.vm_requiereModificacionContrato = False
        Case Else: vm.vm_requiereModificacionContrato = Null
    End Select
    
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("subfrmDatosCDCASUB_Generales.RecogerValores", Me)
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo Errores
    CommitEditsOnForm Me
    
    ' Al guardar, se confirman los cambios y se pierde la capacidad de restaurar
    Me.Parent.GuardarDesdeSubform Me.name
    
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False
    Me.cmdCargarDefaults.Enabled = m_ViewModel.PuedeGuardar_Generales
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("subfrmDatosCDCASUB_Generales.cmdGuardar_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdCargarDefaults_Click()
    Dim srv As New DatosCDCASUBServicio
    Dim def As DatosCDCASUB
    
    On Error GoTo Errores
    If m_ViewModel Is Nothing Then Exit Sub
    
    ' Bloqueo visual
    Me.Parent.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    
    ' 1. Snapshot (Guardar estado actual antes de sobrescribir)
    m_original_refSuministrador = Nz(Me.refSuministrador.value, "")
    m_original_suministradorPrincipal = Nz(Me.suministradorPrincipalNombreDir.value, "")
    m_original_subSuministrador = Nz(Me.subSuministradorNombreDir.value, "")
    m_original_refDesviacionesPrevias = Nz(Me.refDesviacionesPrevias.value, "")
    
    ' 2. Obtener valores calculados
    Set def = srv.GetDatosGeneralesIniciales(m_ViewModel.Solicitud.idSolicitud)
    
    ' 3. Aplicar valores
    Me.refSubSuministrador.value = Nz(def.refSubSuministrador, "")
    Me.suministradorPrincipalNombreDir.value = def.suministradorPrincipalNombreDir
    Me.subSuministradorNombreDir.value = def.subSuministradorNombreDir
    
    Call RecogerValores(m_ViewModel)
    
    Me.Parent.ConfigurarYPoblarInterfaz
    Me.Parent.PoblarSubformActual
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    ' 4. Gestión de botones (Flip)
    Me.cmdCargarDefaults.Enabled = False
    Me.cmdRestaurarDatos.Visible = True
    Me.cmdRestaurarDatos.Enabled = True
    
LimpiarYSalir:
    On Error Resume Next
    Me.Parent.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Set srv = Nothing
    Set def = Nothing
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("subfrmDatosCDCASUB_Generales.cmdCargarDefaults_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdRestaurarDatos_Click()
    On Error GoTo Errores
    
    ' 1. Restaurar valores desde el snapshot
    Me.refSuministrador.value = m_original_refSuministrador
    Me.suministradorPrincipalNombreDir.value = m_original_suministradorPrincipal
    Me.subSuministradorNombreDir.value = m_original_subSuministrador
    Me.refDesviacionesPrevias.value = m_original_refDesviacionesPrevias
    
    Call RecogerValores(m_ViewModel)
    
    Me.Parent.ConfigurarYPoblarInterfaz
    Me.cmdCargarDefaults.Enabled = True
    If Me.cmdCargarDefaults.Enabled And Me.cmdCargarDefaults.Visible Then
        Me.cmdCargarDefaults.SetFocus
    End If
    
    Me.Parent.PoblarSubformActual
    
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("subfrmDatosCDCASUB_Generales.cmdRestaurarDatos_Click", Me)
    Resume LimpiarYSalir
End Sub