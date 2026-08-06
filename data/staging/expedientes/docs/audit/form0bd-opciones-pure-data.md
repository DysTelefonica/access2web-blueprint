# Form_Form0BDOpciones pure-data audit

## Canonical source
- Form: `src/forms/Form_Form0BDOpciones.cls`
- Helper: `src/modules/mod0BDOpcionesHelper.bas`
- Tests: `src/modules/Test_0BDOpcionesHelper.bas`

## Extracted seams
- Form-open role/version state.
- Expedition management target (`FormExpedientesGestion` vs `FormExpedientesGestionTecnica`).
- Generic open-form action metadata for Alta, técnico search, entity manager, E2E batch, and tasks.
- Label dispatch decision.

## Source-only constraints
- Removed executable `Debug.Print` from `Form_Open`.
- UI operations (`DoCmd.OpenForm`, popup progress, busy-form animation) stay in the form as thin rendering/orchestration.

## Runtime status
- Dysflow/runtime compile and procedure-resolution remain globally blocked; this audit records source-only readiness.
