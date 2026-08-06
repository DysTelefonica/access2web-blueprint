

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: Form_frmDetalleExpediente (REFACTORIZADO CON VIEWMODEL)
' ==========================================================================

Private m_ViewModel As ExpedienteViewModel


Private Sub Form_Load()
    On Error GoTo Errores
    
    Dim expServicio As New ExpedienteServicio
    Dim expID As String
    
    
    DoCmd.Hourglass True
    
    Ajustar Me
    expID = Nz(Me.OpenArgs, "")
    If expID = "" Then
        Err.Raise 513, "Form_Load", "No se ha proporcionado un expediente para mostrar."
    End If
    
    ' 1. ÚNICA llamada al servicio para obtener el ViewModel
    Set m_ViewModel = expServicio.getExpedienteViewModelPorID(expID)
    
    If m_ViewModel Is Nothing Then
        Err.Raise 513, "Form_Load", "El expediente con ID " & expID & " no ha sido encontrado."
    Else
        ' 2. ÚNICA llamada para poblar el formulario
        Call PoblarFormularioDesdeViewModel
    End If
    
LimpiarYSalir:
    On Error Resume Next
    
    DoCmd.Hourglass False
    Set expServicio = Nothing
    Exit Sub
    
Errores:
    Call ManejadorErroresFormulario("Form_Load", Me)
    DoCmd.Close acForm, Me.name
    GoTo LimpiarYSalir
End Sub

' ==========================================================================
' MÓDULO: Form_frmDetalleExpediente.cls
' ACCIÓN: Corregir la visualización de fechas nulas.
' ==========================================================================

Private Sub PoblarFormularioDesdeViewModel()
    ' Responsabilidad: Volcar los datos del ViewModel en los controles.
    With m_ViewModel
        Me.txtTitulo.value = .Expediente.titulo
        Me.txtNemotec.value = .Expediente.Nemotecnico
        Me.txtCodigoActividad.value = .Expediente.CodigoActividad
        Me.txtCodExp.value = .Expediente.CodExp
        Me.txtCodExpLargo.value = .Expediente.CodExpLargo
        Me.txtContratistasPrinciaples.value = .NombresContratistas
        Me.txtSubContratistas.value = .NombresSubcontratistas
        Me.txtRCalidad.value = .Expediente.ResponsableCalidad
        Me.txtRTecnico.value = .Expediente.ResponsableTecnico
        Me.txtRacAsignado.value = .Expediente.RacAsignado
        
        ' --- INICIO DE LA CORRECCIÓN ---
        ' Comprobamos si la fecha es mayor que cero antes de asignarla.
        ' Si es cero o Null, el control permanecerá vacío (Null).
        If .Expediente.FechaInicioContrato > 0 Then Me.txtFInicioContrato.value = .Expediente.FechaInicioContrato
        If .Expediente.FechaFinContrato > 0 Then Me.txtFFinContrato.value = .Expediente.FechaFinContrato
        If .Expediente.FechaFinGarantia > 0 Then Me.txtFFinGarantia.value = .Expediente.FechaFinGarantia
        ' --- FIN DE LA CORRECCIÓN ---
    End With
End Sub

Private Sub cmdSalir_Click()
    On Error GoTo Errores
     Me.Tag = "CLOSED"
    DoCmd.Close acForm, Me.name
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdSalir_Click", Me)
    Exit Sub
End Sub