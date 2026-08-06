
Option Compare Database
Option Explicit
Private m_ViewModel As DatosCDCASUBViewModel

Public Sub PoblarControles(ByRef vm As DatosCDCASUBViewModel)
    Set m_ViewModel = vm
    Me.identificacionMaterial.value = Nz(vm.vm_identificacionMaterial, "")
    Me.numPlanoEspecificacion.value = Nz(vm.vm_numPlanoEspecificacion, "")
    Me.cantidadPeriodo.value = Nz(vm.vm_cantidadPeriodo, "")
    Me.numSerieLote.value = Nz(vm.vm_numSerieLote, "")
    Me.causaNC.value = Nz(vm.vm_causaNC, "")
    Me.descripcionImpactoNC.value = Nz(vm.vm_descripcionImpactoNC, "")
End Sub

Public Sub RecogerValores(ByRef vm As DatosCDCASUBViewModel)
    vm.vm_identificacionMaterial = Nz(Me.identificacionMaterial.value, "")
    vm.vm_numPlanoEspecificacion = Nz(Me.numPlanoEspecificacion.value, "")
    vm.vm_cantidadPeriodo = Nz(Me.cantidadPeriodo.value, "")
    vm.vm_numSerieLote = Nz(Me.numSerieLote.value, "")
    vm.vm_causaNC = Nz(Me.causaNC.value, "")
    vm.vm_descripcionImpactoNC = Nz(Me.descripcionImpactoNC.value, "")
End Sub

Private Sub cmdGuardar_Click()
    Me.Parent.GuardarDesdeSubform Me.name
End Sub