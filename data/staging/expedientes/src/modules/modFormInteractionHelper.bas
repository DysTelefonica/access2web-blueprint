Attribute VB_Name = "modFormInteractionHelper"
Option Compare Database
Option Explicit
' modFormInteractionHelper — cross-form UI wiring utility module.
' Implements access-vba-e2e-methodology rule #10 (shared UI utility must exist
' before any feature helper that touches forms) and rule #9 (per-module prefix
' on all public names). Provides the only sanctioned entry points for forms
' to read/write controls on other forms, open/close named forms, and show
' modal prompts whose outcome is assertable by TDD atoms.
'
' Commit 3/5 of refactor-helper-expediente-fechas: REAL implementations
' (replacing the stubs from Commit 1/5). All functions honor vba-access
' §1.6.1 (no IIf/And short-circuit on the same object).
'
' Signatures MUST match spec/mod-form-interaction/spec.md.
' NO ByRef db (pure UI utilities, e2e rule #2).
' NO Me. (helpers receive the target form as explicit parameter, e2e rule #11).
' Split guards per vba-access §1.6.1.
' On Error Resume Next only inside tiny probing blocks; Err.Number checked
' immediately (vba-access §1.5 scalpel pattern).

' --- FormInteraction_FormularioAbierto ---------------------------------------------------
' Returns True if a form with p_NombreFormulario is currently open.
' Uses the Forms collection lookup. No-op safe (returns False if not open).
Public Function FormInteraction_FormularioAbierto(ByVal p_NombreFormulario As String) As Boolean
    Dim frm As Object
    On Error Resume Next
    Set frm = Forms(p_NombreFormulario)
    On Error GoTo 0
    FormInteraction_FormularioAbierto = Not (frm Is Nothing)
End Function

' --- FormInteraction_CerrarFormulario ---------------------------------------------------
' Closes the named form with the given save option. Returns True if close was attempted.
' No-op if the form is not open (returns False).
Public Function FormInteraction_CerrarFormulario(ByVal p_NombreFormulario As String, Optional ByVal p_GuardarCambios As AcCloseSave = acSaveNo) As Boolean
    If Not FormInteraction_FormularioAbierto(p_NombreFormulario) Then
        FormInteraction_CerrarFormulario = False
        Exit Function
    End If
    On Error GoTo errh
    DoCmd.Close acForm, p_NombreFormulario, p_GuardarCambios
    FormInteraction_CerrarFormulario = True
    Exit Function
errh:
    FormInteraction_CerrarFormulario = False
End Function

' --- FormInteraction_ObtenerValorControl ------------------------------------------------
' Reads the value of p_NombreControl from p_Form into p_Valor. Returns True on
' success, False if the form is Nothing, the control is missing, or any error occurs.
' Split guard per vba-access §1.6.1 (no IIf/And short-circuit).
Public Function FormInteraction_ObtenerValorControl(ByVal p_Form As Object, ByVal p_NombreControl As String, ByRef p_Valor As Variant) As Boolean
    If p_Form Is Nothing Then
        FormInteraction_ObtenerValorControl = False
        Exit Function
    End If

    Dim ctl As Object
    On Error Resume Next
    Set ctl = p_Form.Controls(p_NombreControl)
    If Err.Number <> 0 Or ctl Is Nothing Then
        On Error GoTo 0
        FormInteraction_ObtenerValorControl = False
        Exit Function
    End If
    On Error GoTo 0

    p_Valor = ctl.Value
    FormInteraction_ObtenerValorControl = True
End Function

' --- FormInteraction_EstablecerValorControl ---------------------------------------------
' Writes p_Valor to p_NombreControl on p_Form. Returns True on success, False otherwise.
' Split guard per vba-access §1.6.1.
Public Function FormInteraction_EstablecerValorControl(ByVal p_Form As Object, ByVal p_NombreControl As String, ByVal p_Valor As Variant) As Boolean
    If p_Form Is Nothing Then
        FormInteraction_EstablecerValorControl = False
        Exit Function
    End If

    Dim ctl As Object
    On Error Resume Next
    Set ctl = p_Form.Controls(p_NombreControl)
    If Err.Number <> 0 Or ctl Is Nothing Then
        On Error GoTo 0
        FormInteraction_EstablecerValorControl = False
        Exit Function
    End If
    On Error GoTo 0

    ctl.Value = p_Valor
    FormInteraction_EstablecerValorControl = True
End Function

' --- FormInteraction_Mensaje ------------------------------------------------------------
' Wraps MsgBox with a ByRef p_Resultado so atoms can assert the message
' outcome without a real modal blocking COM (e2e rule #5).
' Returns the same Long as MsgBox; p_Resultado is set to the user's choice.
' NOTE: this function WILL block if called with a real modal; TDD atoms
' must not invoke it directly. The form's error branches are stubbed to
' never reach the modal in the test path.
Public Function FormInteraction_Mensaje(ByVal p_Texto As String, ByVal p_Estilo As VbMsgBoxStyle, Optional ByRef p_Resultado As Long) As Long
    Dim result As Long
    result = MsgBox(p_Texto, p_Estilo)
    p_Resultado = result
    FormInteraction_Mensaje = result
End Function

' --- FormInteraction_CargarCombo --------------------------------------------------------
' Assigns a RowSource SQL to a named combo on p_Form. Returns True on success.
Public Function FormInteraction_CargarCombo(ByVal p_Form As Object, ByVal p_NombreCombo As String, ByVal p_RowSource As String) As Boolean
    If p_Form Is Nothing Then
        FormInteraction_CargarCombo = False
        Exit Function
    End If

    Dim ctl As Object
    On Error Resume Next
    Set ctl = p_Form.Controls(p_NombreCombo)
    If Err.Number <> 0 Or ctl Is Nothing Then
        On Error GoTo 0
        FormInteraction_CargarCombo = False
        Exit Function
    End If
    On Error GoTo 0

    ctl.RowSource = p_RowSource
    FormInteraction_CargarCombo = True
End Function

' --- FormInteraction_LimpiarControles ---------------------------------------------------
' Clears the value of every named control in p_NombresControles() on p_Form
' and returns the count actually cleared. Returns 0 if p_Form is Nothing.
Public Function FormInteraction_LimpiarControles(ByVal p_Form As Object, ByRef p_NombresControles() As String) As Long
    If p_Form Is Nothing Then
        FormInteraction_LimpiarControles = 0
        Exit Function
    End If

    Dim i As Long
    Dim cleared As Long
    cleared = 0
    For i = LBound(p_NombresControles) To UBound(p_NombresControles)
        If FormInteraction_EstablecerValorControl(p_Form, p_NombresControles(i), Null) Then
            cleared = cleared + 1
        End If
    Next i
    FormInteraction_LimpiarControles = cleared
End Function

' --- FormInteraction_RefrescarSubform ---------------------------------------------------
' Requeries a named subform control on the active form. Returns True on success.
' No-op if no active form or the named control is missing.
Public Function FormInteraction_RefrescarSubform(ByVal p_NombreSubform As String) As Boolean
    If Screen.ActiveForm Is Nothing Then
        FormInteraction_RefrescarSubform = False
        Exit Function
    End If

    Dim ctl As Object
    On Error Resume Next
    Set ctl = Screen.ActiveForm.Controls(p_NombreSubform)
    If Err.Number <> 0 Or ctl Is Nothing Then
        On Error GoTo 0
        FormInteraction_RefrescarSubform = False
        Exit Function
    End If
    On Error GoTo 0

    ctl.Requery
    FormInteraction_RefrescarSubform = True
End Function

' --- FormInteraction_EstablecerPropiedad ------------------------------------------------
' Sets an arbitrary property on a named control via CallByName. Returns True on
' success, False if the form is Nothing, the control is missing, or the property
' cannot be set. Useful for non-Value properties (e.g. .Visible, .Enabled) that
' FormInteraction_EstablecerValorControl cannot set (it only sets .Value).
' Split guard per vba-access §1.6.1.
Public Function FormInteraction_EstablecerPropiedad(ByVal p_Form As Object, ByVal p_NombreControl As String, ByVal p_Propiedad As String, ByVal p_Valor As Variant) As Boolean
    If p_Form Is Nothing Then
        FormInteraction_EstablecerPropiedad = False
        Exit Function
    End If

    Dim ctl As Object
    On Error Resume Next
    Set ctl = p_Form.Controls(p_NombreControl)
    If Err.Number <> 0 Or ctl Is Nothing Then
        On Error GoTo 0
        FormInteraction_EstablecerPropiedad = False
        Exit Function
    End If
    On Error GoTo 0

    On Error Resume Next
    CallByName ctl, p_Propiedad, VbSet, p_Valor
    If Err.Number <> 0 Then
        On Error GoTo 0
        FormInteraction_EstablecerPropiedad = False
        Exit Function
    End If
    On Error GoTo 0

    FormInteraction_EstablecerPropiedad = True
End Function
