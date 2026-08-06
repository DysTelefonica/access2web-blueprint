# Form_FormExpedienteFechas — pure-data audit

## Scope
- Source form: `src/forms/Form_FormExpedienteFechas.cls`
- Helper: `src/modules/modExpedienteFechasHelper.bas`
- Tests: `src/modules/Test_ExpedienteFechasHelper.bas`

## Extracted pure-data behavior
- Warranty end-date calculation (`GarantiaMeses`, `FechaCertificacion`, `FechaFinContrato`).
- Load state: parent registrar visibility/enabled state and edit permissions.
- Unload decision: whether a changed editable snapshot should be saved.

## UI that intentionally stays in the form
- Button click date assignment (`Date`) because it writes Access controls directly.
- `SetFormFechasFromExpediente`, `getSnapshot`, and `GuardarPestana` calls because they are existing form/DAO boundaries.

## Anti-pattern check
- Helper/tests do not receive `ByRef p_Form`.
- Tests do not call `DoCmd.OpenForm`, `Forms(...)`, or `Screen.ActiveForm`.
