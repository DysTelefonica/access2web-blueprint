

Option Compare Database
Option Explicit

Public Sub PoblarControles(ByRef vm As DatosPCSUBViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub
    Me.numPlanoEspecificacion = Nz(vm.vm_numPlanoEspecificacion, "")
    Me.descripcionMaterialAfectado = Nz(vm.vm_descripcionMaterialAfectado, "")
    Me.descripcionPropuestaCambio = Nz(vm.vm_descripcionPropuestaCambio, "")
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPCSUB_Propuesta.PoblarControles", Me
End Sub

Public Sub RecogerValores(ByRef vm As DatosPCSUBViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub
    vm.vm_numPlanoEspecificacion = Nz(Me.numPlanoEspecificacion.value, "")
    vm.vm_descripcionMaterialAfectado = Nz(Me.descripcionMaterialAfectado.value, "")
    vm.vm_descripcionPropuestaCambio = Nz(Me.descripcionPropuestaCambio.value, "")
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPCSUB_Propuesta.RecogerValores", Me
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo Errores
    
    Me.Parent.GuardarDesdeSubform Me.name
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPCSUB_Propuesta.cmdGuardar_Click", Me
End Sub