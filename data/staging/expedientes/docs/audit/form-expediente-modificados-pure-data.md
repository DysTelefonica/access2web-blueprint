# Form_FormExpedienteModificados — pure-data audit

## Scope
- Source form: `src/forms/Form_FormExpedienteModificados.cls`
- Helper: `src/modules/modExpedienteModificadosHelper.bas`
- Tests: `src/modules/Test_ExpedienteModificadosHelper.bas`

## Extracted pure-data behavior
- Load state: derives `AllowEdits` and parent registrar visibility from edit/admin flags.
- List rendering: converts `ColModificados` into the listbox `RowSource` payload.
- Register/delete orchestration: helper validates DTO and delegates to `ExpedienteOperaciones`; the form remains responsible for UI controls, popup/hourglass, and Access form navigation.

## UI that intentionally stays in the form
- Opening `FormModificado` and assigning the `WithEvents` child form.
- Popup/hourglass rendering and MsgBox display.
- Reading/writing Access controls.

## Anti-pattern check
- Helpers/tests do not receive `ByRef p_Form`.
- Tests do not call `DoCmd.OpenForm`, `Forms(...)`, or `Screen.ActiveForm`.
- DTO/test stubs use `Scripting.Dictionary` only.
