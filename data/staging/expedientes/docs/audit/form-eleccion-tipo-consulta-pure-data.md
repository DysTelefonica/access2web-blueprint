# Form_FormEleccionTipoConsulta pure-data audit

## Canonical source
- Form: `src/forms/Form_FormEleccionTipoConsulta.cls`
- Helper: `src/modules/modEleccionTipoConsultaHelper.bas`
- Tests: `src/modules/Test_EleccionTipoConsultaHelper.bas`

## Extracted seams
- `OpenArgs` parsing and default source form/list selection.
- Initial list row-source/default selection.
- Consulta type selection (`completo` vs `simple`) and collection-load decision.
- List-load plan for AM, lotes, basados, and técnica lists.
- Export request derivation (`incluirDerivados`).

## Source-only constraints
- Public form business methods were shrunk to private helpers.
- The remaining `Forms(m_NombreFormulario).Controls(...)` calls are UI-bound list extraction; tests cover the decision plan without opening forms.

## Runtime status
- Dysflow/runtime compile and procedure-resolution remain globally blocked; this audit records source-only readiness.
