
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_subfrmDatosCDCA_Propuesta
' ==========================================================================

Private m_ViewModel As DatosCDCAViewModel

Public Sub PoblarControles(ByRef vm As DatosCDCAViewModel)
    On Error GoTo Errores
    Set m_ViewModel = vm
    If vm Is Nothing Then Exit Sub
    
    Me.identificacionMaterial.value = Nz(vm.vm_identificacionMaterial, "")
    Me.numPlanoEspecificacion.value = Nz(vm.vm_numPlanoEspecificacion, "")
    Me.cantidadPeriodo.value = Nz(vm.vm_cantidadPeriodo, "")
    Me.numSerieLote.value = Nz(vm.vm_numSerieLote, "")
    Me.causaNC.value = Nz(vm.vm_causaNC, "")
    Me.descripcionImpactoNC.value = Nz(vm.vm_descripcionImpactoNC, "")
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("PoblarControles", Me)
    Resume LimpiarYSalir
End Sub

Public Sub RecogerValores(ByRef vm As DatosCDCAViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub
    
    vm.vm_identificacionMaterial = Nz(Me.identificacionMaterial.value, "")
    vm.vm_numPlanoEspecificacion = Nz(Me.numPlanoEspecificacion.value, "")
    vm.vm_cantidadPeriodo = Nz(Me.cantidadPeriodo.value, "")
    vm.vm_numSerieLote = Nz(Me.numSerieLote.value, "")
    vm.vm_causaNC = Nz(Me.causaNC.value, "")
    vm.vm_descripcionImpactoNC = Nz(Me.descripcionImpactoNC.value, "")
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("RecogerValores", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo Errores
    Me.Parent.GuardarDesdeSubform Me.name
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdGuardar_Click", Me)
    Resume LimpiarYSalir
End Sub