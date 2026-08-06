
Option Compare Database
Option Explicit
Public Event SolicitudHaCambiado()
Private m_ViewModel As DatosCDCASUBViewModel

Public Sub PoblarControles(ByRef vm As DatosCDCASUBViewModel)
    On Error GoTo Errores
    Set m_ViewModel = vm
    If vm Is Nothing Then Exit Sub
    
    Select Case UCase(Nz(vm.vm_decisionFinal, ""))
        Case "APROBADO": Me.grpDecisionFinal = 1
        Case "NO APROBADO": Me.grpDecisionFinal = 2
        Case Else: Me.grpDecisionFinal = 0
    End Select
    
    Me.observacionesFinales = Nz(vm.vm_obsDecisionFinal, "")
    Me.NombreFirmanteFinal = Nz(vm.vm_NombreFirmanteFinal, "")
    
    ' Lógica Botones
    Dim datosYaGuardados As Boolean, esAprobada As Boolean
    datosYaGuardados = (Nz(vm.vm_decisionFinal, "") <> "")
    esAprobada = (vm.Solicitud.idEstadoInterno = 8)
    
    
    Me.cmdGuardar.Enabled = (vm.PuedeGuardar_DecisionFinal And Not esAprobada)

    Dim habilitar As Boolean
    habilitar = IIf(esAprobada, False, vm.PuedeGuardar_DecisionFinal)
    
    Me.grpDecisionFinal.Enabled = habilitar
    Me.observacionesFinales.Enabled = habilitar
    Me.NombreFirmanteFinal.Enabled = habilitar

    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosCDCASUB_DecisionFinal.PoblarControles", Me
End Sub

Public Sub RecogerValores(ByRef vm As DatosCDCASUBViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub
    
    Select Case Nz(Me.grpDecisionFinal.value, 0)
        Case 1: vm.vm_decisionFinal = "APROBADO"
        Case 2: vm.vm_decisionFinal = "NO APROBADO"
        Case Else: vm.vm_decisionFinal = ""
    End Select

    vm.vm_obsDecisionFinal = Nz(Me.observacionesFinales.value, "")
    vm.vm_NombreFirmanteFinal = Nz(Me.NombreFirmanteFinal.value, "")
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "RecogerValores", Me
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo Errores
    Call RecogerValores(m_ViewModel)
    
    If m_ViewModel.vm_decisionFinal = "" Then
        MsgBox "El campo 'Decisión Final' es obligatorio.", vbExclamation, "CONDOR"
        Exit Sub
    End If

    If m_ViewModel.vm_decisionFinal = "APROBADO" Then
        Me.Parent.ProcesarCierreFormalizacion
    Else
        Me.Parent.GuardarDesdeSubform Me.name
    End If
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdGuardar_Click", Me)
End Sub


