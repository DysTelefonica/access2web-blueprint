
Option Compare Database
Option Explicit
Private m_ViewModel As DatosCDCASUBViewModel

Public Sub PoblarControles(ByRef vm As DatosCDCASUBViewModel)
    On Error GoTo Errores
    Set m_ViewModel = vm
    ' Checkboxes
    Me.afectaPrestaciones = vm.vm_afectaPrestaciones
    Me.afectaSeguridad = vm.vm_afectaSeguridad
    Me.afectaFiabilidad = vm.vm_afectaFiabilidad
    Me.afectaVidaUtil = vm.vm_afectaVidaUtil
    Me.afectaMedioambiente = vm.vm_afectaMedioambiente
    Me.afectaMantenibilidad = vm.vm_afectaMantenibilidad
    Me.afectaIntercambiabilidad = vm.vm_afectaIntercambiabilidad
    Me.afectaApariencia = vm.vm_afectaApariencia
    Me.afectaOtros = vm.vm_afectaOtros
    
    ' Radios
    Select Case UCase(Nz(vm.vm_impactoCoste, ""))
        Case "AUMENTARÁ": Me.grpIncidenciaCoste = 1
        Case "DISMINUIRÁ": Me.grpIncidenciaCoste = 2
        Case "NO VARIARÁ": Me.grpIncidenciaCoste = 3
        Case Else: Me.grpIncidenciaCoste = Null
    End Select
    
    Select Case UCase(Nz(vm.vm_clasificacionNC, ""))
        Case "MAYOR": Me.grpNCConsiderada = 1
        Case "MENOR": Me.grpNCConsiderada = 2
        Case Else: Me.grpNCConsiderada = Null
    End Select
    
    Select Case vm.vm_esSubSuministradorAD
        Case True: Me.grpAutoridadDisenio = 1
        Case False: Me.grpAutoridadDisenio = 2
        Case Else: Me.grpAutoridadDisenio = Null
    End Select
    
    Me.identificacionAutoridadDiseno = Nz(vm.vm_identificacionAutoridadDiseno, "")
    Me.efectoFechaEntrega = Nz(vm.vm_efectoFechaEntrega, "")
    
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("subfrmDatosCDCASUB_Impacto.PoblarControles", Me)
End Sub

Public Sub RecogerValores(ByRef vm As DatosCDCASUBViewModel)
    On Error GoTo Errores
    vm.vm_afectaPrestaciones = Me.afectaPrestaciones
    vm.vm_afectaSeguridad = Me.afectaSeguridad
    vm.vm_afectaFiabilidad = Me.afectaFiabilidad
    vm.vm_afectaVidaUtil = Me.afectaVidaUtil
    vm.vm_afectaMedioambiente = Me.afectaMedioambiente
    vm.vm_afectaMantenibilidad = Me.afectaMantenibilidad
    vm.vm_afectaIntercambiabilidad = Me.afectaIntercambiabilidad
    vm.vm_afectaApariencia = Me.afectaApariencia
    vm.vm_afectaOtros = Me.afectaOtros
    
    Select Case Nz(Me.grpIncidenciaCoste, 0)
        Case 1: vm.vm_impactoCoste = "AUMENTARÁ"
        Case 2: vm.vm_impactoCoste = "DISMINUIRÁ"
        Case 3: vm.vm_impactoCoste = "NO VARIARÁ"
        Case Else: vm.vm_impactoCoste = ""
    End Select
    
    Select Case Nz(Me.grpNCConsiderada, 0)
        Case 1: vm.vm_clasificacionNC = "MAYOR"
        Case 2: vm.vm_clasificacionNC = "MENOR"
        Case Else: vm.vm_clasificacionNC = ""
    End Select
    
    Select Case Nz(Me.grpAutoridadDisenio, 0)
        Case 1: vm.vm_esSubSuministradorAD = True
        Case 2: vm.vm_esSubSuministradorAD = False
        Case Else: vm.vm_esSubSuministradorAD = Null
    End Select
    
    vm.vm_identificacionAutoridadDiseno = Nz(Me.identificacionAutoridadDiseno, "")
    vm.vm_efectoFechaEntrega = Nz(Me.efectoFechaEntrega, "")
    
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("subfrmDatosCDCASUB_Impacto.RecogerValores", Me)
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo Errores
    
    Call RecogerValores(m_ViewModel)
    ' NO dirty check local - DELEGA todo al formulario padre
    Me.Parent.GuardarDesdeSubform Me.name
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("subfrmDatosCDCASUB_Impacto.cmdGuardar_Click", Me)
    Resume LimpiarYSalir
End Sub