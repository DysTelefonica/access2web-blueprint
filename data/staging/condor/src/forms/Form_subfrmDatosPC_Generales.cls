
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_subfrmDatosPC_Generales.cls
' VERSIÓN: DEFINITIVA con carga de defaults, restauración y lógica de habilitación simétrica.
' ==========================================================================

' -- El ViewModel del formulario padre es la fuente de verdad para los datos --
Private m_ViewModel As DatosPCViewModel

' -- Variables para guardar el estado previo a la carga de defaults --
Private m_original_refContratoInspeccionOficial As String
Private m_original_refSuministrador As String
Private m_original_denominacionContrato As String
Private m_original_SuministradorNombreDir As String
Private m_original_objetoContrato As String

' --- COLORES EXPLÍCITOS (no dependen del diseño) ---
Private Const COLOR_FONDO_EDITABLE As Long = 16774386   ' Blanco
Private Const COLOR_FONDO_INMUTABLE As Long = 16774463  ' Gris
Private Const COLOR_TEXTO_NORMAL As Long = 4666411      ' Gris oscuro

' --------------------------------------------------------------------------
' MÉTODOS DE POBLADO Y RECOGIDA
' --------------------------------------------------------------------------

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

Public Sub PoblarControles(ByRef vm As DatosPCViewModel)
    On Error GoTo Errores
    Set m_ViewModel = vm
    If m_ViewModel Is Nothing Then Exit Sub
    
    ' CORRECCIÓN: Habilitación y Visibilidad simétrica con Guardar
    Dim puedeEditar As Boolean
    puedeEditar = m_ViewModel.PuedeGuardar_Generales
    
    ' Volcar datos del ViewModel a los controles
    Me.refContratoInspeccionOficial.value = Nz(m_ViewModel.vm_refContratoInspeccionOficial, "")
    Call AplicarEstadoVisualControl(Me.refContratoInspeccionOficial, puedeEditar)
    Me.refSuministrador.value = Nz(m_ViewModel.vm_refSuministrador, "")
    Call AplicarEstadoVisualControl(Me.refSuministrador, puedeEditar)
    Me.denominacionContrato.value = Nz(m_ViewModel.vm_denominacionContrato, "")
    Call AplicarEstadoVisualControl(Me.denominacionContrato, puedeEditar)
    Me.SuministradorNombreDir.value = Nz(m_ViewModel.vm_SuministradorNombreDir, "")
    Call AplicarEstadoVisualControl(Me.SuministradorNombreDir, puedeEditar)
    Me.objetoContrato.value = Nz(m_ViewModel.vm_objetoContrato, "")
    Call AplicarEstadoVisualControl(Me.objetoContrato, puedeEditar)
    
    ' Botón Cargar Defaults: Siempre visible, habilitado si se puede editar
    Me.cmdCargarDefaults.Visible = True
    Me.cmdCargarDefaults.Enabled = puedeEditar
    
    ' Botón Guardar: Habilitado si se puede editar
    Me.cmdGuardar.Enabled = puedeEditar
    
    If Me.cmdCargarDefaults.Enabled And Me.cmdCargarDefaults.Visible Then
        Me.cmdCargarDefaults.SetFocus
    End If

    ' Botón Restaurar: Oculto por defecto (solo sale tras cargar defaults)
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPC_Generales.PoblarControles", Me
End Sub

Public Sub RecogerValores(ByRef vm As DatosPCViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub
    
    vm.vm_refContratoInspeccionOficial = Nz(Me.refContratoInspeccionOficial.value, "")
    vm.vm_refSuministrador = Nz(Me.refSuministrador.value, "")
    vm.vm_denominacionContrato = Nz(Me.denominacionContrato.value, "")
    vm.vm_SuministradorNombreDir = Nz(Me.SuministradorNombreDir.value, "")
    vm.vm_objetoContrato = Nz(Me.objetoContrato.value, "")
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPC_Generales.RecogerValores", Me
End Sub

' --------------------------------------------------------------------------
' EVENTOS DE BOTONES
' --------------------------------------------------------------------------

Private Sub cmdGuardar_Click()
    On Error GoTo Errores
    CommitEditsOnForm Me
    Me.Parent.GuardarDesdeSubform Me.name
    
    ' Si se guardó, ya no hay nada que restaurar (se consolida el cambio)
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False
    Me.cmdCargarDefaults.Enabled = m_ViewModel.PuedeGuardar_Generales
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPC_Generales.cmdGuardar_Click", Me
End Sub

Private Sub cmdCargarDefaults_Click()
    Dim pcServ As New DatosPCServicio
    Dim pcInicial As DatosPC
    
    On Error GoTo Errores
    If m_ViewModel Is Nothing Then Exit Sub
    
    Me.Parent.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    
    ' 1. Guardar el estado actual de los controles (SNAPSHOT BLINDADO)
    ' Usamos Nz() para evitar error 94 si el campo está vacío
    m_original_refContratoInspeccionOficial = Nz(Me.refContratoInspeccionOficial.value, "")
    m_original_refSuministrador = Nz(Me.refSuministrador.value, "")
    m_original_denominacionContrato = Nz(Me.denominacionContrato.value, "")
    m_original_SuministradorNombreDir = Nz(Me.SuministradorNombreDir.value, "")
    m_original_objetoContrato = Nz(Me.objetoContrato.value, "")
    
    ' 2. Llamar al servicio para obtener los datos iniciales calculados
    Set pcInicial = pcServ.GetDatosPCIniciales(m_ViewModel.idSolicitud)
    
    ' 3. Volcar estos datos en los controles del formulario
    If Not pcInicial Is Nothing Then
        Me.refContratoInspeccionOficial.value = pcInicial.refContratoInspeccionOficial
        Me.refSuministrador.value = pcInicial.refSuministrador
        Me.denominacionContrato.value = pcInicial.denominacionContrato
        Me.SuministradorNombreDir.value = pcInicial.SuministradorNombreDir
        Me.objetoContrato.value = pcInicial.objetoContrato
    End If
    
    Call RecogerValores(m_ViewModel)
    
    
    Me.Parent.ConfigurarYPoblarInterfaz
    Me.Parent.PoblarSubformActual
    Call QuitarFocoControl(Me, "cmdCargarDefaults")
    ' 4. Gestión de botones: Mostrar Restaurar, Deshabilitar Cargar
    Me.cmdCargarDefaults.Enabled = False
    Me.cmdRestaurarDatos.Visible = True
    Me.cmdRestaurarDatos.Enabled = True
    
LimpiarYSalir:
    Me.Parent.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Set pcServ = Nothing
    Set pcInicial = Nothing
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("subfrmDatosPC_Generales.cmdCargarDefaults_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdRestaurarDatos_Click()
    On Error GoTo Errores
    
    ' 1. Volcar los valores originales de nuevo en los controles
    Me.refContratoInspeccionOficial.value = m_original_refContratoInspeccionOficial
    Me.refSuministrador.value = m_original_refSuministrador
    Me.denominacionContrato.value = m_original_denominacionContrato
    Me.SuministradorNombreDir.value = m_original_SuministradorNombreDir
    Me.objetoContrato.value = m_original_objetoContrato
    
    Call RecogerValores(m_ViewModel)
   
    Me.Parent.ConfigurarYPoblarInterfaz
     
    Me.Parent.PoblarSubformActual
     
    ' 2. Gestión de botones: Ocultar Restaurar, Habilitar Cargar
    Me.cmdCargarDefaults.Enabled = True
    If Me.cmdCargarDefaults.Enabled And Me.cmdCargarDefaults.Visible Then
        Me.cmdCargarDefaults.SetFocus
    End If
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False
    
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("subfrmDatosPC_Generales.cmdRestaurarDatos_Click", Me)
End Sub
