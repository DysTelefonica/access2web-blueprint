# Form_Form0BDOpcionesTecnicos pure-data audit

## Canonical source
- Form: `src/forms/Form_Form0BDOpcionesTecnicos.cls`
- Helper: `src/modules/mod0BDOpcionesTecnicosHelper.bas`
- Tests: `src/modules/Test_0BDOpcionesTecnicosHelper.bas`

## Extracted seams
- Close decision (`close-form` when main options is open; otherwise close database).
- Complete/simple search targets and `OpenArgs`.
- Form-open caption/version state.
- Label dispatch decision.

## Source-only constraints
- Tests use pure strings/booleans only; no form opening or UI references.
- The form still performs `DoCmd.OpenForm`, `DoCmd.Close`, and `Application.CloseCurrentDatabase` as UI shell actions.

## Runtime status
- Dysflow/runtime compile and procedure-resolution remain globally blocked; this audit records source-only readiness.
