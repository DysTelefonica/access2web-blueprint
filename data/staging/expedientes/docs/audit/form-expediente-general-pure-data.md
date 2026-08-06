# Form_FormExpedienteGeneral — pure-data audit

## Scope
- Source form: `src/forms/Form_FormExpedienteGeneral.cls`
- Helper: `src/modules/modExpedienteGeneralHelper.bas`
- Tests: `src/modules/Test_ExpedienteGeneralHelper.bas`

## Extracted pure-data behavior
- Load state and command enablement from edit/admin flags.
- Ambito=`HPS` decision for `HPSAplica`.
- SharePoint button enablement.
- Unload save/no-save decision based on snapshots.
- Ordinal warning display-text fallback (`CodExp` → `Nemotecnico` → `Titulo`).

## UI/DAO that intentionally stays in the form
- Picker forms and `WithEvents` selection handlers.
- Existing `GuardadoAutomaticoHelper` calls (`SetFormGeneralFromExpediente`, `getSnapshot`, `GuardarPestana`).
- `Ejecutar` for opening SharePoint URLs.

## Anti-pattern check
- Helper/tests do not receive `ByRef p_Form`.
- Tests do not call `DoCmd.OpenForm`, `Forms(...)`, or `Screen.ActiveForm`.
