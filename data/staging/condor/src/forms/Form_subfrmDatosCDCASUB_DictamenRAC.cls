
Option Compare Database
Option Explicit


Private m_ViewModel As DatosCDCASUBViewModel
Private m_original_racNombre As String
Private m_original_racNombreDelegador As String
Public Sub PoblarControles(ByRef vm As DatosCDCASUBViewModel)
    On Error GoTo Errores
    Set m_ViewModel = vm
    If m_ViewModel Is Nothing Then Exit Sub
    
    ' 1. Volcar datos comunes
    Me.racCodigo.value = Nz(vm.vm_racCodigo, "")
    Me.observacionesRAC.value = Nz(vm.vm_observacionesRAC, "")
    Me.racNombre.value = Nz(vm.vm_racNombre, "")

    ' 2. Volcar datos específicos (Delegador)
    Me.observacionesRACDelegador.value = Nz(vm.vm_observacionesRACDelegador, "")
    Me.racNombreDelegador.value = Nz(vm.vm_racNombreDelegador, "")

    ' 3. Lógica de Bloqueo Visual
    Dim habilitado As Boolean
    habilitado = m_ViewModel.PuedeGuardar_DictamenRAC
    
    Me.racCodigo.Enabled = habilitado
    Me.racNombre.Enabled = habilitado
    Me.observacionesRAC.Enabled = False
    Me.observacionesRACDelegador.Enabled = False
    Me.racNombreDelegador.Enabled = habilitado
    
    Me.cmdGuardar.Enabled = habilitado
    Me.cmdCargarDatosPredeterminados.Visible = True
    Me.cmdCargarDatosPredeterminados.Enabled = habilitado
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False

    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosCDCASUB_DictamenRAC.PoblarControles", Me
End Sub

Public Sub RecogerValores(ByRef vm As DatosCDCASUBViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub
    
    vm.vm_racCodigo = Nz(Me.racCodigo.value, "")
    vm.vm_observacionesRAC = Nz(Me.observacionesRAC.value, "")
    vm.vm_racNombre = Nz(Me.racNombre.value, "")
    
    ' Campos nuevos
    vm.vm_observacionesRACDelegador = Nz(Me.observacionesRACDelegador.value, "")
    vm.vm_racNombreDelegador = Nz(Me.racNombreDelegador.value, "")
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosCDCASUB_DictamenRAC.RecogerValores", Me
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo Errores
    
    Call RecogerValores(m_ViewModel)
    Me.Parent.GuardarDesdeSubform Me.name
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosCDCASUB_DictamenRAC.cmdGuardar_Click", Me
End Sub



Private Sub cmdCargarDatosPredeterminados_Click()
    Dim expServ As New ExpedienteServicio
    Dim nombreRAC As String
    
    On Error GoTo Errores
    
    If m_ViewModel Is Nothing Then Exit Sub
    
    Me.Parent.lblBloqueo.Visible = True: DoCmd.Hourglass True
    
    m_original_racNombre = Nz(Me.racNombre.value, "")
    m_original_racNombreDelegador = Nz(Me.racNombreDelegador.value, "")

    nombreRAC = expServ.GetRacAsignadoPorExpediente(m_ViewModel.Solicitud.idExpediente)
    
    If nombreRAC <> "" Then
        Me.racNombre = nombreRAC
    Else
        MsgBox "Este expediente no tiene un RAC asignado. Hable con el gestor de expedientes", vbExclamation, "Sin RAC"
    End If
    
    Me.cmdCargarDatosPredeterminados.Enabled = False
    Me.cmdRestaurarDatos.Visible = True
    Me.cmdRestaurarDatos.Enabled = True

LimpiarYSalir:
    Me.Parent.lblBloqueo.Visible = False: DoCmd.Hourglass False
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdCargarDatosPredeterminados_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdRestaurarDatos_Click()
    On Error GoTo Errores

    Me.racNombre = m_original_racNombre
    Me.racNombreDelegador = m_original_racNombreDelegador

    Call RecogerValores(m_ViewModel)

    Me.cmdCargarDatosPredeterminados.Enabled = True
    Me.cmdCargarDatosPredeterminados.SetFocus
    Me.cmdRestaurarDatos.Visible = False

    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdRestaurarDatos_Click", Me)
End Sub










