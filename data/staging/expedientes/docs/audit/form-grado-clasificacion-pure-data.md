# Form_FormGradoClasificacion pure-data audit

## Canonical source
- Form: `src/forms/Form_FormGradoClasificacion.cls`
- Existing gestion helper preserved: `src/modules/modGradosClasificacionHelper.bas`
- New singular-form helper: `src/modules/modGradoClasificacionFormHelper.bas`
- Tests: `src/modules/Test_GradoClasificacionFormHelper.bas`

## Extracted seams
- Form-open state/title (`ALTA` vs `EDICIÓN`) and edit-button enablement.
- Form field value mapping into a pure dictionary.
- Change detection wrapper over the existing generic CRUD helper.
- Register decision and event name (`Alta` vs `Editado`).

## Source-only constraints
- Existing `modGradosClasificacionHelper` was not overwritten because it belongs to `Form_FormGradosClasificacionGestion`.
- Public form methods were converted to private helpers used by the UI shell.

## Runtime status
- Dysflow/runtime compile and procedure-resolution remain globally blocked; this audit records source-only readiness.
