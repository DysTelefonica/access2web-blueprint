
Option Compare Database
Option Explicit

Public Sub PoblarControles(ByRef vm As DatosPCViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub

    ' Motivos
    Me.motivoCorregirDeficiencias = vm.vm_motivoCorregirDeficiencias
    Me.motivoMejorarCapacidad = vm.vm_motivoMejorarCapacidad
    Me.motivoAumentarNacionalizacion = vm.vm_motivoAumentarNacionalizacion
    Me.motivoMejorarSeguridad = vm.vm_motivoMejorarSeguridad
    Me.motivoMejorarFiabilidad = vm.vm_motivoMejorarFiabilidad
    Me.motivoMejorarCosteEficacia = vm.vm_motivoMejorarCosteEficacia
    Me.motivoOtros = vm.vm_motivoOtros
    Me.motivoOtrosDetalle = Nz(vm.vm_motivoOtrosDetalle, "")

    ' Incidencias
    
    Select Case Nz(vm.vm_incidenciaCoste, "")
        Case "Aumentará": Me.grpIncidenciaCoste = 1
        Case "Disminuirá": Me.grpIncidenciaCoste = 2
        Case "No Variará": Me.grpIncidenciaCoste = 3
        Case Else: Me.grpIncidenciaCoste = 0
    End Select
    
    Select Case Nz(vm.vm_incidenciaPlazo, "")
        Case "Aumentará": Me.grpIncidenciaPlazo = 1
        Case "Disminuirá": Me.grpIncidenciaPlazo = 2
        Case "No Variará": Me.grpIncidenciaPlazo = 3
        Case Else: Me.grpIncidenciaPlazo = 0
    End Select
    
   
    Me.incidenciaSeguridad = vm.vm_incidenciaSeguridad
    Me.incidenciaFiabilidad = vm.vm_incidenciaFiabilidad
    Me.incidenciaMantenibilidad = vm.vm_incidenciaMantenibilidad
    Me.incidenciaIntercambiabilidad = vm.vm_incidenciaIntercambiabilidad
    Me.incidenciaVidaUtilAlmacen = vm.vm_incidenciaVidaUtilAlmacen
    Me.incidenciaFuncionamientoFuncion = vm.vm_incidenciaFuncionamientoFuncion

    ' Afectación
    Select Case Nz(vm.vm_impactoClasificacion, "")
        Case "MAYOR": Me.grpClasificacion = 1
        Case "MENOR": Me.grpClasificacion = 2
        
        Case Else: Me.grpClasificacion = 0
    End Select
    
    Select Case Nz(vm.vm_CambioAfectaAMaterial, "")
        Case "Material ya entregado": Me.grpAfeccionMaterial = 1
        Case "Material por entregar": Me.grpAfeccionMaterial = 2
        
        Case Else: Me.grpAfeccionMaterial = 0
    End Select
    
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPC_Impacto.PoblarControles", Me
End Sub

Public Sub RecogerValores(ByRef vm As DatosPCViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub

    vm.vm_motivoCorregirDeficiencias = CBool(Me.motivoCorregirDeficiencias.value)
    vm.vm_motivoMejorarCapacidad = CBool(Me.motivoMejorarCapacidad.value)
    vm.vm_motivoAumentarNacionalizacion = CBool(Me.motivoAumentarNacionalizacion.value)
    vm.vm_motivoMejorarSeguridad = CBool(Me.motivoMejorarSeguridad.value)
    vm.vm_motivoMejorarFiabilidad = CBool(Me.motivoMejorarFiabilidad.value)
    vm.vm_motivoMejorarCosteEficacia = CBool(Me.motivoMejorarCosteEficacia.value)
    vm.vm_motivoOtros = CBool(Me.motivoOtros.value)
    vm.vm_motivoOtrosDetalle = Nz(Me.motivoOtrosDetalle.value, "")
    
    ' --- Incidencias (OptionGroup ? texto canónico) ---
    Select Case Nz(Me.grpIncidenciaCoste.value, 0)
        Case 1: vm.vm_incidenciaCoste = "Aumentará"
        Case 2: vm.vm_incidenciaCoste = "Disminuirá"
        Case 3: vm.vm_incidenciaCoste = "No Variará"
        Case Else: vm.vm_incidenciaCoste = ""
    End Select
    ' --- Incidencias (OptionGroup ? texto canónico) ---
    Select Case Nz(Me.grpIncidenciaPlazo.value, 0)
        Case 1: vm.vm_incidenciaPlazo = "Aumentará"
        Case 2: vm.vm_incidenciaPlazo = "Disminuirá"
        Case 3: vm.vm_incidenciaPlazo = "No Variará"
        Case Else: vm.vm_incidenciaPlazo = ""
    End Select
    
    
    vm.vm_incidenciaSeguridad = CBool(Me.incidenciaSeguridad.value)
    vm.vm_incidenciaFiabilidad = CBool(Me.incidenciaFiabilidad.value)
    vm.vm_incidenciaMantenibilidad = CBool(Me.incidenciaMantenibilidad.value)
    vm.vm_incidenciaIntercambiabilidad = CBool(Me.incidenciaIntercambiabilidad.value)
    vm.vm_incidenciaVidaUtilAlmacen = CBool(Me.incidenciaVidaUtilAlmacen.value)
    vm.vm_incidenciaFuncionamientoFuncion = CBool(Me.incidenciaFuncionamientoFuncion.value)
    
    Select Case Nz(Me.grpClasificacion.value, 0)
        Case 1: vm.vm_impactoClasificacion = "MAYOR"
        Case 2: vm.vm_impactoClasificacion = "MENOR"
        Case Else: vm.vm_incidenciaPlazo = ""
    End Select
    Select Case Nz(Me.grpAfeccionMaterial.value, 0)
        Case 1: vm.vm_CambioAfectaAMaterial = "Material ya entregado"
        Case 2: vm.vm_CambioAfectaAMaterial = "Material por entregar"
        Case Else: vm.vm_CambioAfectaAMaterial = ""
    End Select
   
    
   
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPC_Impacto.RecogerValores", Me
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo Errores
    'CommitEditsOnForm Me
    Me.Parent.GuardarDesdeSubform Me.name
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPC_Impacto.cmdGuardar_Click", Me
End Sub
