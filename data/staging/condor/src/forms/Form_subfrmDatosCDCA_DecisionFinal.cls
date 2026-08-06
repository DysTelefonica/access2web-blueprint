
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_subfrmDatosCDCA_DecisionFinal
' ==========================================================================

Private m_ViewModel As DatosCDCAViewModel

Public Sub PoblarControles(ByRef vm As DatosCDCAViewModel)
    On Error GoTo Errores
    Set m_ViewModel = vm
    If vm Is Nothing Then Exit Sub
    
    ' 1. Volcar datos
    Select Case UCase(Nz(vm.vm_decisionFinal, ""))
        Case "APROBADO": Me.grpDecisionFinal = 1
        Case "NO APROBADO": Me.grpDecisionFinal = 2
        Case Else: Me.grpDecisionFinal = 0
    End Select
    
    Me.observacionesFinales = Nz(vm.vm_obsDecisionFinal, "")
    Me.NombreFirmanteFinal = Nz(vm.vm_NombreFirmanteFinal, "")
    
    Dim esAprobada As Boolean
    esAprobada = (vm.Solicitud.idEstadoInterno = 8)
    Me.cmdGuardar.Enabled = (vm.PuedeGuardar_DecisionFinal And Not esAprobada)

    ' 3. BLOQUEO VISUAL DE CAMPOS
    Dim habilitarCampos As Boolean
    If esAprobada Then
        habilitarCampos = False
    Else
        habilitarCampos = vm.PuedeGuardar_DecisionFinal
    End If
    
    Me.grpDecisionFinal.Enabled = habilitarCampos
    Me.observacionesFinales.Enabled = habilitarCampos
    Me.NombreFirmanteFinal.Enabled = habilitarCampos

    Exit Sub
Errores:
    Call ManejadorErroresFormulario("subfrmDatosCDCA_DecisionFinal.PoblarControles", Me)
   
End Sub

Public Sub RecogerValores(ByRef vm As DatosCDCAViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub
    
    Select Case Nz(Me.grpDecisionFinal.value, 0)
        Case 1: vm.vm_decisionFinal = "APROBADO"
        Case 2: vm.vm_decisionFinal = "NO APROBADO"
        Case Else: vm.vm_decisionFinal = ""
    End Select

    vm.vm_obsDecisionFinal = Nz(Me.observacionesFinales.value, "")
    vm.vm_NombreFirmanteFinal = Nz(Me.NombreFirmanteFinal.value, "")
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("RecogerValores", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo Errores
    
    ' 1. Recoger valores de la pantalla al ViewModel
    Call RecogerValores(m_ViewModel)
    
    ' 2. Validación Básica
    If m_ViewModel.vm_decisionFinal = "" Then
        MsgBox "El campo 'Decisión Final' es obligatorio.", vbExclamation, "CONDOR"
        Exit Sub
    End If

    ' --- LÓGICA DE APROBACIÓN AUTOMÁTICA ---
    If m_ViewModel.vm_decisionFinal = "APROBADO" Then
        
        ' Llamamos al padre para ejecutar la transacción de cierre
        ' El parámetro True indica "Modo Cierre Automático"
        Me.Parent.ProcesarCierreFormalizacion
        
    Else
        ' --- LÓGICA ESTÁNDAR (NO APROBADO O PENDIENTE) ---
        ' Guardado normal sin transiciones automáticas forzadas
        Me.Parent.GuardarDesdeSubform Me.name
    End If
    
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("subfrmDatosCDCA_DecisionFinal.cmdGuardar_Click", Me)
End Sub
' EN: Form_subfrmDatosCDCA_DecisionFinal.cls


