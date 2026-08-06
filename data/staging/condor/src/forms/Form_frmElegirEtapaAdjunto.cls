
Option Compare Database
Option Explicit

' --- Propiedades públicas para que el llamador pueda leer el resultado ---
Public Property Get EtapaSeleccionada() As String
    EtapaSeleccionada = "Otro"
End Property

Public Property Get EtapaNombreSeleccionado() As String
    EtapaNombreSeleccionado = "Otro"
End Property

Public Property Get DescripcionAdjunto() As String
    DescripcionAdjunto = Trim$(Nz(Me.txtDescripcion.value, ""))
End Property


' ==========================================================================
' MÓDULO: Form_frmElegirEtapaAdjunto.cls
' ACCIÓN: Preseleccionar la siguiente etapa si es única, o la actual si no lo es.
' ==========================================================================

Private Sub Form_Load()
    Dim idSol As Long
    
    On Error GoTo Errores
    
    Ajustar Me
    Me.Caption = "Clasificar Documento Adjunto"
    
    idSol = CLng(Nz(Me.OpenArgs, 0))
    If idSol = 0 Then Err.Raise 513, "Form_Load", "No se proporcionó ID de Solicitud."
    
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("Form_Load", Me)
    DoCmd.Close acForm, Me.name
End Sub

Private Sub cmdAceptar_Click()
    On Error GoTo Errores
    If Trim$(Nz(Me.txtDescripcion.value, "")) = "" Then
        MsgBox "Debe indicar una descripcion para el adjunto.", vbInformation, "CONDOR"
        Exit Sub
    End If
    Me.Tag = "OK"
    Me.Visible = False
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdAceptar_Click", Me)
    Resume LimpiarYSalir
End Sub

Private Sub cmdCancelar_Click()
    On Error GoTo Errores
    Me.Tag = "Cancel"
    Me.Visible = False
LimpiarYSalir:
    Exit Sub
Errores:
    Call ManejadorErroresFormulario("cmdCancelar_Click", Me)
    Resume LimpiarYSalir
End Sub