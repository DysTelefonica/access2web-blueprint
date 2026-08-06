' ==========================================================================
' MÓDULO: subfrmDatosPCSUB_Generales.cls
' VERSIÓN: Alineado con ERD (tbDatosPCSUB) y form UI.
'           Controles en form: refContratoInspeccionOficial, denominacionContrato,
'           SubsuministradorNombreDir, refSubSuministrador, objetoContrato.
'           NO EXISTEN: refSuministrador, SuministradorNombreDir, subSuministradorTelefono,
'           subSuministradorEmail ( phantom fields removidos ).
' ==========================================================================

' -- El ViewModel del formulario padre es la fuente de verdad para los datos --
Private m_ViewModel As DatosPCSUBViewModel

' -- Variables para guardar el estado previo a la carga de defaults --
Private m_original_refContratoInspeccionOficial As String
Private m_original_denominacionContrato As String
Private m_original_SubsuministradorNombreDir As String
Private m_original_refSubSuministrador As String
Private m_original_objetoContrato As String

' --- COLORES EXPLÍCITOS (no dependen del diseño) ---
Private Const COLOR_FONDO_EDITABLE As Long = 16774386   ' Blanco
Private Const COLOR_FONDO_INMUTABLE As Long = 16774463  ' Gris
Private Const COLOR_TEXTO_NORMAL As Long = 4666411      ' Gris oscuro

' --------------------------------------------------------------------------
' MÉTODOS DE POBLADO Y RECOGIDA
' --------------------------------------------------------------------------

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

Public Sub PoblarControles(ByRef vm As DatosPCSUBViewModel)
    On Error GoTo Errores
    Set m_ViewModel = vm
    If m_ViewModel Is Nothing Then Exit Sub
    
    Dim puedeEditar As Boolean
    puedeEditar = m_ViewModel.PuedeGuardar_Generales
    
    ' Campos propios de PC (ERD: refContratoInspeccionOficial, denominacionContrato,
    ' SubsuministradorNombreDir, refSubSuministrador, objetoContrato)
    Me.refContratoInspeccionOficial.value = Nz(m_ViewModel.vm_refContratoInspeccionOficial, "")
    Call AplicarEstadoVisualControl(Me.refContratoInspeccionOficial, puedeEditar)
    Me.denominacionContrato.value = Nz(m_ViewModel.vm_denominacionContrato, "")
    Call AplicarEstadoVisualControl(Me.denominacionContrato, puedeEditar)
    Me.SubsuministradorNombreDir.value = Nz(m_ViewModel.vm_SubsuministradorNombreDir, "")
    Call AplicarEstadoVisualControl(Me.SubsuministradorNombreDir, puedeEditar)
    Me.refSubSuministrador.value = Nz(m_ViewModel.vm_refSubSuministrador, "")
    Call AplicarEstadoVisualControl(Me.refSubSuministrador, puedeEditar)
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

    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPCSUB_Generales.PoblarControles", Me
End Sub

Public Sub RecogerValores(ByRef vm As DatosPCSUBViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub
    
    vm.vm_refContratoInspeccionOficial = Nz(Me.refContratoInspeccionOficial.value, "")
    vm.vm_denominacionContrato = Nz(Me.denominacionContrato.value, "")
    vm.vm_SubsuministradorNombreDir = Nz(Me.SubsuministradorNombreDir.value, "")
    vm.vm_refSubSuministrador = Nz(Me.refSubSuministrador.value, "")
    vm.vm_objetoContrato = Nz(Me.objetoContrato.value, "")
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPCSUB_Generales.RecogerValores", Me
End Sub

' --------------------------------------------------------------------------
' EVENTOS DE BOTONES
' --------------------------------------------------------------------------

Private Sub cmdGuardar_Click()
    On Error GoTo Errores
    CommitEditsOnForm Me
    Me.Parent.GuardarDesdeSubform Me.name
    
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False
    Me.cmdCargarDefaults.Enabled = m_ViewModel.PuedeGuardar_Generales
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPCSUB_Generales.cmdGuardar_Click", Me
End Sub

Private Sub cmdCargarDefaults_Click()
    Dim pcServ As New DatosPCSUBServicio
    Dim pcInicial As DatosPCSUB
    
    On Error GoTo Errores
    If m_ViewModel Is Nothing Then Exit Sub
    
    Me.Parent.lblBloqueo.Visible = True
    DoCmd.Hourglass True
    
    ' 1. Guardar el estado actual de los controles (SNAPSHOT BLINDADO)
    m_original_refContratoInspeccionOficial = Nz(Me.refContratoInspeccionOficial.value, "")
    m_original_denominacionContrato = Nz(Me.denominacionContrato.value, "")
    m_original_SubsuministradorNombreDir = Nz(Me.SubsuministradorNombreDir.value, "")
    m_original_refSubSuministrador = Nz(Me.refSubSuministrador.value, "")
    m_original_objetoContrato = Nz(Me.objetoContrato.value, "")
    
    ' 2. Llamar al servicio para obtener los datos iniciales calculados
    Set pcInicial = pcServ.GetDatosGeneralesIniciales(m_ViewModel.idSolicitud)
    
    ' 3. Volcar estos datos en los controles del formulario
    If Not pcInicial Is Nothing Then
        Me.refContratoInspeccionOficial.value = pcInicial.refContratoInspeccionOficial
        Me.denominacionContrato.value = pcInicial.denominacionContrato
        Me.SubsuministradorNombreDir.value = pcInicial.SubsuministradorNombreDir
        Me.refSubSuministrador.value = pcInicial.refSubSuministrador
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
    Call ManejadorErroresFormulario("subfrmDatosPCSUB_Generales.cmdCargarDefaults_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdRestaurarDatos_Click()
    On Error GoTo Errores
    
    ' 1. Volcar los valores originales de nuevo en los controles
    Me.refContratoInspeccionOficial.value = m_original_refContratoInspeccionOficial
    Me.denominacionContrato.value = m_original_denominacionContrato
    Me.SubsuministradorNombreDir.value = m_original_SubsuministradorNombreDir
    Me.refSubSuministrador.value = m_original_refSubSuministrador
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
    Call ManejadorErroresFormulario("subfrmDatosPCSUB_Generales.cmdRestaurarDatos_Click", Me)
End Sub