
' ==========================================================================
' MÓDULO: Form_subfrmDatosPC_DecisionFinal.cls
' VERSIÓN: DEFINITIVA con lógica de eliminación y retroceso de estado.
' ==========================================================================
Option Compare Database
Option Explicit

Private m_ViewModel As DatosPCViewModel



Public Sub PoblarControles(ByRef vm As DatosPCViewModel)
    On Error GoTo Errores
    Set m_ViewModel = vm
    If vm Is Nothing Then Exit Sub
    
    ' 1. Volcar datos (USANDO LOS NOMBRES NUEVOS)
    Me.obsAprobacionAutoridadDiseno = Nz(vm.vm_obsAprobacionAutoridadDiseno, "")
    
    ' CORREGIDO: Usamos el nombre correcto del control para la Autoridad de Diseño
    Me.NombreAutoridadDiseno = Nz(vm.vm_NombreAutoridadDiseno, "")
    
    Me.obsDecisionFinal = Nz(vm.vm_obsDecisionFinal, "")
    
    Select Case vm.vm_decisionFinal
        Case "APROBADO": Me.grpDecisionFinal = 1
        Case "NO APROBADO": Me.grpDecisionFinal = 2
        Case Else: Me.grpDecisionFinal = Null
    End Select
    
    ' CORREGIDO: Nombre estandarizado
    Me.NombreFirmanteFinal = Nz(vm.vm_NombreFirmanteFinal, "")
    
    ' 2. Lógica de Botones
    Dim datosYaGuardados As Boolean
    Dim esAprobada As Boolean
    
    datosYaGuardados = (Nz(vm.vm_decisionFinal, "") <> "")
    esAprobada = (vm.Solicitud.idEstadoInterno = 8)
    
    
    Me.cmdGuardar.Enabled = (vm.PuedeGuardar_DecisionFinal And Not esAprobada)
    
    ' 3. BLOQUEO VISUAL
    Dim habilitarCampos As Boolean
    If esAprobada Then
        habilitarCampos = False
    Else
        habilitarCampos = vm.PuedeGuardar_DecisionFinal
    End If
    
    Me.grpDecisionFinal.Enabled = habilitarCampos
    Me.obsAprobacionAutoridadDiseno.Enabled = habilitarCampos
    
    ' HABILITAMOS LOS NOMBRES NUEVOS
    Me.NombreAutoridadDiseno.Enabled = habilitarCampos
    Me.NombreFirmanteFinal.Enabled = habilitarCampos
    
    Me.obsDecisionFinal.Enabled = habilitarCampos
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "subfrmDatosPC_DecisionFinal.PoblarControles", Me
End Sub

Public Sub RecogerValores(ByRef vm As DatosPCViewModel)
    On Error GoTo Errores
    If vm Is Nothing Then Exit Sub
    
    vm.vm_obsAprobacionAutoridadDiseno = Nz(Me.obsAprobacionAutoridadDiseno.value, "")
    vm.vm_NombreAutoridadDiseno = Nz(Me.NombreAutoridadDiseno.value, "")
    
    Select Case Nz(Me.grpDecisionFinal.value, 0)
        Case 1: vm.vm_decisionFinal = "APROBADO"
        Case 2: vm.vm_decisionFinal = "NO APROBADO"
        Case Else: vm.vm_decisionFinal = ""
    End Select
    
    vm.vm_obsDecisionFinal = Nz(Me.obsDecisionFinal.value, "")
    vm.vm_NombreFirmanteFinal = Nz(Me.NombreFirmanteFinal.value, "")
    
    Exit Sub
Errores:
    ManejadorErroresFormulario "Form_subfrmDatosPC_DecisionFinal.RecogerValores", Me
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
        
        ' Llamamos al padre para ejecutar la transacción de cierre (PDF + Estado 8)
        Me.Parent.ProcesarCierreFormalizacion
        
    Else
        ' --- LÓGICA ESTÁNDAR ---
        ' Guardado normal sin transiciones automáticas forzadas
        Me.Parent.GuardarDesdeSubform Me.name
    End If
    
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("Form_subfrmDatosPC_DecisionFinal.cmdGuardar_Click", Me)
End Sub

