
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_subfrmDatosCDCA_Impacto
' ACCIÓN: Gestión correcta de estado inicial (Null) para opciones obligatorias.
' ==========================================================================

Private m_ViewModel As DatosCDCAViewModel

Public Sub PoblarControles(ByRef vm As DatosCDCAViewModel)
    On Error GoTo Errores
    Set m_ViewModel = vm
    If vm Is Nothing Then Exit Sub

    ' 1. Checkboxes (Siempre tienen valor True/False)
    Me.afectaPrestaciones = vm.vm_afectaPrestaciones
    Me.afectaSeguridad = vm.vm_afectaSeguridad
    Me.afectaFiabilidad = vm.vm_afectaFiabilidad
    Me.afectaVidaUtil = vm.vm_afectaVidaUtil
    Me.afectaMedioambiente = vm.vm_afectaMedioambiente
    Me.afectaIntercambiabilidad = vm.vm_afectaIntercambiabilidad
    Me.afectaMantenibilidad = vm.vm_afectaMantenibilidad
    Me.afectaApariencia = vm.vm_afectaApariencia
    Me.afectaOtros = vm.vm_afectaOtros

    ' 2. Lógica de "Sección Virgen"
    ' Si impactoCoste está vacío, asumimos que nunca se ha guardado esta sección.
    ' Forzamos Null en los grupos para obligar al usuario a elegir explícitamente.
    Dim seccionYaGuardada As Boolean
    seccionYaGuardada = (Nz(vm.vm_impactoCoste, "") <> "")
    
    If seccionYaGuardada Then
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
        
        Select Case vm.vm_esSuministradorAD
            Case True: Me.grpAutoridadDisenio = 1
            Case False: Me.grpAutoridadDisenio = 2
            Case Else: Me.grpAutoridadDisenio = Null
        End Select
    Else
        ' Estado Virgen: Todo a Null
        Me.grpIncidenciaCoste = Null
        Me.grpNCConsiderada = Null
        Me.grpAutoridadDisenio = Null
    End If
    
    ' 3. Campos de Texto
    Me.identificacionAutoridadDiseno = Nz(vm.vm_identificacionAutoridadDiseno, "")
    Me.efectoFechaEntrega = Nz(vm.vm_efectoFechaEntrega, "")
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("PoblarControles", Me)
    Resume LimpiarYSalir
End Sub

Public Sub RecogerValores(ByRef vm As DatosCDCAViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub
    
    ' Checkboxes
    vm.vm_afectaPrestaciones = Me.afectaPrestaciones.value
    vm.vm_afectaSeguridad = Me.afectaSeguridad.value
    vm.vm_afectaFiabilidad = Me.afectaFiabilidad.value
    vm.vm_afectaVidaUtil = Me.afectaVidaUtil.value
    vm.vm_afectaMedioambiente = Me.afectaMedioambiente.value
    vm.vm_afectaMantenibilidad = Me.afectaMantenibilidad.value
    vm.vm_afectaIntercambiabilidad = Me.afectaIntercambiabilidad.value
    vm.vm_afectaApariencia = Me.afectaApariencia.value
    vm.vm_afectaOtros = Me.afectaOtros.value
    
    ' Opciones (Mapeo inverso ID -> Texto)
    Select Case Nz(Me.grpIncidenciaCoste.value, 0)
        Case 1: vm.vm_impactoCoste = "AUMENTARÁ"
        Case 2: vm.vm_impactoCoste = "DISMINUIRÁ"
        Case 3: vm.vm_impactoCoste = "NO VARIARÁ"
        Case Else: vm.vm_impactoCoste = ""
    End Select
    
    Select Case Nz(Me.grpNCConsiderada.value, 0)
        Case 1: vm.vm_clasificacionNC = "MAYOR"
        Case 2: vm.vm_clasificacionNC = "MENOR"
        Case Else: vm.vm_clasificacionNC = ""
    End Select
    
    Select Case Nz(Me.grpAutoridadDisenio.value, 0)
        Case 1: vm.vm_esSuministradorAD = True
        Case 2: vm.vm_esSuministradorAD = False
        Case Else: vm.vm_esSuministradorAD = Null
    End Select
    
    ' Textos
    vm.vm_identificacionAutoridadDiseno = Nz(Me.identificacionAutoridadDiseno.value, "")
    vm.vm_efectoFechaEntrega = Nz(Me.efectoFechaEntrega.value, "")
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("RecogerValores", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo Errores
    
    Call RecogerValores(m_ViewModel)
    ' NO dirty check local - DELEGA todo al formulario padre
    Me.Parent.GuardarDesdeSubform Me.name
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdGuardar_Click", Me)
    Resume LimpiarYSalir
End Sub