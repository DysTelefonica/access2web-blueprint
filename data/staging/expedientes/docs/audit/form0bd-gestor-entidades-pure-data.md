# Form_Form0BDGestorEntidades pure-data audit

## Canonical source
- Form: `src/forms/Form_Form0BDGestorEntidades.cls`
- Helper: `src/modules/mod0BDGestorEntidadesHelper.bas`
- Tests: `src/modules/Test_0BDGestorEntidadesHelper.bas`

## Extracted seams
- List row-source generation from `m_ObjEntorno.OtrasOpciones`.
- Selection state for enabling `ComandoAbrirFormulario`.
- Open-form decision from selected option ID and form map.
- Double-click dispatch decision.

## Source-only constraints
- No test opens forms or references `Forms(...)`, `Screen.ActiveForm`, or `DoCmd.OpenForm`.
- UI actions remain in the form: `DoCmd.OpenForm`, `DoCmd.Close`, and `MsgBox` rendering.

## Runtime status
- Dysflow/runtime compile and procedure-resolution remain globally blocked; this audit records source-only readiness.
