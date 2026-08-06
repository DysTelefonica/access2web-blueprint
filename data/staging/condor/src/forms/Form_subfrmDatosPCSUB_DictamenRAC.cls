' ==========================================================================
' MÓDULO: subfrmDatosPCSUB_DictamenRAC.cls
' VERSIÓN: Alineado con ERD (tbDatosPCSUB)
'           Usa campos estándar: racDelegadoNombre, racCodigo, observacionesRACDelegador, racNombreDelegador
'           Mapeo de controles a VM según especificación:
'             racDelegadoNombre     -> vm_racNombre
'             racCodigo             -> vm_racCodigo
'             observacionesRACDelegador -> vm_observacionesRACDelegador
'             racNombreDelegador    -> vm_racNombreDelegador
' ==========================================================================

Option Compare Database
Option Explicit

Private m_ViewModel As DatosPCSUBViewModel
Private m_original_racDelegadoNombre As String

Public Sub PoblarControles(ByRef vm As DatosPCSUBViewModel)
    On Error GoTo Errores
    Set m_ViewModel = vm
    If m_ViewModel Is Nothing Then Exit Sub

    ' Volcar datos a controles (campos estándar ERD / RAC Delegador)
    Me.racDelegadoNombre = Nz(vm.vm_racNombre, "")
    Me.racCodigo = Nz(vm.vm_racCodigo, "")
    Me.observacionesRACDelegador = Nz(vm.vm_observacionesRACDelegador, "")
    Me.racNombreDelegador = Nz(vm.vm_racNombreDelegador, "")

    ' Lógica de Bloqueo VISUAL
    Dim habilitado As Boolean
    habilitado = m_ViewModel.PuedeGuardar_DictamenRAC

    Me.racDelegadoNombre.Enabled = habilitado
    Me.racCodigo.Enabled = habilitado
    Me.observacionesRACDelegador.Enabled = False

    Me.cmdGuardar.Enabled = habilitado
    Me.cmdCargarDatosPredeterminados.Visible = True
    Me.cmdCargarDatosPredeterminados.Enabled = habilitado
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False

    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPCSUB_DictamenRAC.PoblarControles", Me
End Sub

Public Sub RecogerValores(ByRef vm As DatosPCSUBViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub

    ' Recogemos los campos estándar del bloque RAC
    vm.vm_racNombre = Nz(Me.racDelegadoNombre.value, "")
    vm.vm_racCodigo = Nz(Me.racCodigo.value, "")
    vm.vm_observacionesRACDelegador = Nz(Me.observacionesRACDelegador.value, "")
    vm.vm_racNombreDelegador = Nz(Me.racNombreDelegador.value, "")

    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPCSUB_DictamenRAC.RecogerValores", Me
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo Errores

    Call RecogerValores(m_ViewModel)
    Me.Parent.GuardarDesdeSubform Me.name

    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPCSUB_DictamenRAC.cmdGuardar_Click", Me
End Sub

Private Sub cmdCargarDatosPredeterminados_Click()
    Dim expServ As New ExpedienteServicio
    Dim nombreDelegador As String

    On Error GoTo Errores

    If m_ViewModel Is Nothing Then Exit Sub

    Me.Parent.lblBloqueo.Visible = True
    DoCmd.Hourglass True

    m_original_racDelegadoNombre = Nz(Me.racDelegadoNombre.value, "")

    nombreDelegador = expServ.GetRacAsignadoPorExpediente(m_ViewModel.Solicitud.idExpediente)

    If nombreDelegador <> "" Then
        Me.racDelegadoNombre = nombreDelegador
    Else
        MsgBox "Este expediente no tiene un RAC asignado. Hable con el gestor de expedientes", vbExclamation, "Sin RAC"
    End If

    Me.cmdCargarDatosPredeterminados.Enabled = False
    Me.cmdRestaurarDatos.Visible = True
    Me.cmdRestaurarDatos.Enabled = True

LimpiarYSalir:
    Me.Parent.lblBloqueo.Visible = False
    DoCmd.Hourglass False
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdCargarDatosPredeterminados_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdRestaurarDatos_Click()
    On Error GoTo Errores

    Me.racDelegadoNombre = m_original_racDelegadoNombre

    Call RecogerValores(m_ViewModel)

    Me.cmdGuardar.SetFocus
    Call QuitarFocoControl(Me, "cmdRestaurarDatos")
    Me.cmdRestaurarDatos.Visible = False
    Me.cmdCargarDatosPredeterminados.Enabled = True
    If Me.cmdCargarDatosPredeterminados.Enabled And Me.cmdCargarDatosPredeterminados.Visible Then
        Me.cmdCargarDatosPredeterminados.SetFocus
    End If

    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPCSUB_DictamenRAC.cmdRestaurarDatos_Click", Me
End Sub
