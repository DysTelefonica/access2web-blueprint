# PCSUB fixture graph — Slice 1+2 + Slice A/B1/B2a/B2b1

This is the schema-first gate for `fixture-first-tdd-suite` Slice 1+2 and `fixture-first-coverage-expansion` Slice A/B1/B2a/B2b1. No PCSUB data-touching test is accepted unless it seeds this graph explicitly in the sandbox/local backend.

## Current strict manifest

`tests/tests.vba.json` is the strict atomic manifest. It contains:

- Slice 1 foundation atoms:
  - `Test_PCSUB_FixtureGraphContract`
  - `Test_PCSUB_SqlStrEscapesApostrophe`
- Slice 2 happy-path atoms:
  - `Test_PCSUB_GetDatos_ReturnsSeededEntity`
  - `Test_PCSUB_GuardarDatosGenerales_PersistsFields`
  - `Test_PCSUB_EsDatosGeneralesCompleta_TrueForSeededCompleteRow`
  - `Test_PCSUB_EsDetalleCompleto_TrueForSeededDetail`
  - `Test_PCSUB_EsMotivosCompleto_TrueForSeededMotivo`
- Slice A completeness sad-path characterization atoms.
- Slice B1 proposal characterization GREEN atoms:
  - `Test_PCSUB_GuardarPropuesta_PersistsFields`
  - `Test_PCSUB_GuardarPropuesta_FailsWhenProposalMissing`
  - `Test_PCSUB_EsParteTecnicaCompleta_FalseWhenImpactIncomplete`
- Slice B2a impact characterization GREEN atoms accepted after Dysflow runtime 5/5 and Judgment Day PASS WITH WARNINGS:
  - `Test_PCSUB_GuardarPropuesta_FailsWhenMaterialMissing`
  - `Test_PCSUB_GuardarImpacto_PersistsFields`
  - `Test_PCSUB_GuardarImpacto_FailsWhenNoMotives`
  - `Test_PCSUB_EsParteTecnicaCompleta_TrueWhenImpactComplete`
  - `Test_PCSUB_EsParteTecnicaCompleta_FalseWhenImpactClassificationMissing`
- Slice B2b1 impact assertion-hardening characterization GREEN atoms accepted after manual compile, Dysflow runtime 2/2, and Judgment Day PASS WITH WARNINGS/no CRITICAL:
  - `Test_PCSUB_GuardarImpacto_PersistsAllImpactFields`
  - `Test_PCSUB_GuardarImpacto_RejectedSavePreservesAllImpactFields`

## Schema evidence

Sources checked in this worktree:

- `docs/ERD/condor_datos.md` generated 2026-05-08.
- Dysflow live schema reads against `condor_datos.accdb` for `TbExpedientes`, `tbSolicitudes`, and `tbDatosPCSUB` during Slice B2a.
- Frontend live schema (`CONDOR.accdb`) for `TbConfiguracionBackends` via read-only OleDb.

### `TbExpedientes`

- PK: `IDExpediente` (`Long`, required).
- Required fields from ERD: `IDExpediente` only.
- Fixture strategy: use deterministic IDs `>= 900000`; populate optional descriptive fields for traceability (`CodExp`, `Titulo`, `ObjetoContrato`, `ContratistaPrincipal`).

### `tbSolicitudes`

- PK: `idSolicitud` (`Long`, required).
- Domain parent: `idExpediente` points to `TbExpedientes.IDExpediente`.
- Required fields from ERD: `idSolicitud`, `idExpediente`, `tipoSolicitud`, `codigoSolicitud`, `idEstadoInterno`, `fechaCreacion`, `usuarioCreacion`, `revisionCalidadEstado`.
- Fixture strategy: PCSUB tests use `tipoSolicitud='PC_SUB'`, deterministic `codigoSolicitud`, and `revisionCalidadEstado='PENDIENTE'`.
- Workflow-state strategy for Slice 2: seed `idEstadoInterno = estadoRegistro` so `GuardarDatosGenerales` exercises field persistence without triggering a workflow transition. The test asserts the state remains `estadoRegistro` after Act.

### `tbDatosPCSUB`

- PK: `idDatosPCSUB` (`Long`, required).
- Domain parent: `idSolicitud` points to `tbSolicitudes.idSolicitud`.
- Required fields from ERD: `idDatosPCSUB`, `idSolicitud`, `refContratoInspeccionOficial`.
- Seeded core fields for Slice 2 assertions: identifiers, `refContratoInspeccionOficial`, general fields (`refSubSuministrador`, `denominacionContrato`, `SubsuministradorNombreDir`, `objetoContrato`), detail fields (`descripcionMaterialAfectado`, `descripcionPropuestaCambio`), and `motivoCorregirDeficiencias` when the scenario needs a motivo.
- Proposal fields used by B1 `GuardarPropuesta`: `descripcionMaterialAfectado` (`Memo`, optional in schema but required by service rule), `descripcionPropuestaCambio` (`Memo`, optional in schema but required by service rule), `numPlanoEspecificacion` (`Memo`, optional), and `descripcionPropuestaCambioCont` (`Memo`, optional).
- Impact fields used by B2a `GuardarImpacto`: motive booleans (`motivoCorregirDeficiencias`, `motivoMejorarCapacidad`, `motivoAumentarNacionalizacion`, `motivoMejorarSeguridad`, `motivoMejorarFiabilidad`, `motivoMejorarCosteEficacia`, `motivoOtros`), `motivoOtrosDetalle` (`Memo`, optional but required by service when `motivoOtros=True`), `incidenciaCoste` (`Text(50)`, service values: `AUMENTARÁ`, `DISMINUIRÁ`, `NO VARIARÁ`), `incidenciaPlazo` (`Text(50)`, same service values), incidence booleans, `impactoClasificacion` (`Text(255)`, service values: `MAYOR`, `MENOR`), and `CambioAfectaAMaterial` (`Text(255)`, service values: `Material ya entregado`, `Material por entregar`).
- Technical-completeness fields used by B1/B2a `EsParteTecnicaCompleta`: proposal fields above plus impact/motive fields `incidenciaCoste`, `impactoClasificacion`, and at least one motive boolean. `incidenciaPlazo` and `CambioAfectaAMaterial` are persisted/validated by `GuardarImpacto` but are not currently checked by `EsParteTecnicaCompleta`.
- Important stale-name guard: use `tbDatosPCSUB`/`idDatosPCSUB`; do not use stale `tbDatosPC_SUB`/`idDatosPC_SUB`.

### `TbConfiguracionBackends` (frontend-local)

- PK: `Id`.
- Required fields from live schema: `Id`, `Habilitado`.
- Sandbox routing fields: `BackendSandbox`, `BackendTest`, `PasswordBackend`.
- Lifecycle rule: `TestHelper.ForceLocalBackend` must validate the configured sandbox path with filesystem and DAO open before setting `m_TestingMode=True`.

## Seed and teardown order

Parent → child seed order:

1. `TbExpedientes`
2. `tbSolicitudes`
3. `tbDatosPCSUB`

Child → parent teardown order:

1. `tbDatosPCSUB`
2. `tbSolicitudes`
3. `TbExpedientes`

## Relationship caveat

The ERD relationship list exposes `tbSolicitudes -> tbDatosPC`, `tbDatosCDCA`, and `tbDatosCDCASUB`, but not `tbDatosPCSUB`. For tests, `TbExpedientes -> tbSolicitudes -> tbDatosPCSUB` is treated as the required domain fixture graph because the application/service code uses `idExpediente` and `idSolicitud` that way.

## Slice B1 proposal fixture semantics

B1 keeps the same parent graph and adds proposal-specific fixture shapes:

- Valid proposal persistence: seed `TbExpedientes -> tbSolicitudes -> tbDatosPCSUB` with legal required fields and initially empty proposal data; act through `DatosPCSUBServicio.GuardarPropuesta`; assert exact persisted `descripcionMaterialAfectado`, `numPlanoEspecificacion`, `descripcionPropuestaCambio`, and `descripcionPropuestaCambioCont`; assert exactly one `tbDatosPCSUB` row for the target. Accepted as characterization GREEN after the existing production behavior passed the test.
- Missing proposal data covered by B1: seed a legal row with proposal data, blank `descripcionPropuestaCambio` on the entity before Act, assert `GuardarPropuesta` raises validation and persisted proposal fields/cardinality remain unchanged. This does not cover blank `descripcionMaterialAfectado`; that sibling service-validation case remains a B1/B2 follow-up gap.
- Proposal-side technical completeness false: seed a full technical positive control and a target with proposal complete but impact/motives incomplete; assert the control returns `True` and the target returns `False` so the test cannot pass by fallback to unrelated complete data. Accepted as characterization GREEN for B1 only; B2 still owns impact-side validation/completeness behavior.

Seed order remains `TbExpedientes -> tbSolicitudes -> tbDatosPCSUB`; teardown remains `tbDatosPCSUB -> tbSolicitudes -> TbExpedientes`. All IDs remain deterministic `>= 900000` and teardown deletes only those deterministic rows.

## Slice B2a impact fixture semantics

B2a is accepted characterization GREEN after Dysflow runtime evidence (`sliceB2a`: 5 total / 5 passed / 0 failed / 0 blocked) and Judgment Day PASS WITH WARNINGS/no CRITICAL. The scope is the exact asserted behavior only; it must not be read as full impact-validation coverage.

- Missing material follow-up: seed a legal proposal row, blank only `descripcionMaterialAfectado`, call `GuardarPropuesta`, and assert validation plus no persisted mutation of the asserted material/proposal fields and target cardinality exactly one.
- Valid impact persistence: seed proposal-complete `TbExpedientes -> tbSolicitudes -> tbDatosPCSUB`, apply valid motives/incidences/classification/material-decision through `GuardarImpacto`, then assert the subset of persisted fields set/asserted by `ApplyValidImpactFields` and target cardinality exactly one. Exact-all-impact-fields persistence remains debt.
- Missing motives validation: seed a complete technical row, blank all motive booleans and `motivoOtrosDetalle` before `GuardarImpacto`, assert validation mentions motive, and assert the scoped persisted impact state/cardinality remain unchanged. Full rejected-save no-side-effect coverage across every column remains debt.
- Impact-side completeness: seed a complete positive control for every negative completeness assertion. The B2a positive atom asserts a complete impact row returns `True`; the B2a negative atom clears `impactoClasificacion` on the target, keeps the complete control present, and asserts target `False` with exact target `tbDatosPCSUB` cardinality.

Seed order remains `TbExpedientes -> tbSolicitudes -> tbDatosPCSUB`; teardown remains `tbDatosPCSUB -> tbSolicitudes -> TbExpedientes`. All IDs remain deterministic `>= 900000` and teardown deletes only those deterministic rows. Parent cardinality is deterministic and not lucky-data dependent, but broader parent-cardinality strengthening remains debt under the strictest fixture-first bar.

## Slice B2b1 impact assertion-hardening fixture semantics

B2b is split to protect the review budget. B2b1 is accepted only as characterization GREEN assertion-hardening coverage after manual compile, Dysflow `sliceB2b1` runtime evidence of 2 total / 2 passed / 0 failed / 0 blocked, and Judgment Day PASS WITH WARNINGS/no CRITICAL. No production GREEN fix was needed.

- Exact impact-field comparison: seed one complete parent graph, apply the GuardarImpacto-owned impact fields on the entity, call `GuardarImpacto`, assert the motive booleans, `motivoOtrosDetalle`, both incidence enum strings, incidence booleans, `impactoClasificacion`, `CambioAfectaAMaterial`, and exact `TbExpedientes`/`tbSolicitudes`/`tbDatosPCSUB` cardinality. Judgment Day warning: this compares against current fixture values; fields that already matched seed/default values are not proven to transition.
- Rejected-save no-motives snapshot preservation: seed one complete impact row, snapshot it, mutate every impact field while making motives invalid, expect validation, then assert the persisted row matches the pre-Act impact snapshot and exact parent/child cardinality. Judgment Day warning: this proves no-side-effects for the no-motives rejection shape only, not every invalid-domain rejection variant.

B2b2/follow-up debt remains: unchanged-field transition proof where needed, invalid `incidenciaCoste`, invalid/missing `incidenciaPlazo`, invalid `impactoClasificacion`, and invalid `CambioAfectaAMaterial` variants.

Seed order remains `TbExpedientes -> tbSolicitudes -> tbDatosPCSUB`; teardown remains `tbDatosPCSUB -> tbSolicitudes -> TbExpedientes`. All IDs remain deterministic `>= 900000` and teardown deletes only those deterministic rows.

## Known pending Slice 1+2 review debt

- Warning 3 remains intentionally pending: do not alter the shared production `getdb()` UI path in this mini-slice.
- Warning 4 remains intentionally pending: teardown error suppression and broad `TeardownAll` need a focused follow-up.
- Slice 3 sad paths remain out of scope.
- B1/B2a Judgment Day warning: parent cardinality assertions are deterministic and not lucky-data dependent, but remain weaker than the strictest parent-cardinality bar; B2b1 strengthens parent and child cardinality for the two impact assertion-hardening atoms. Broader invalid-domain variants and unchanged-field transition proof remain B2b2/follow-up debt.

## Legacy debt excluded from strict manifest

- The old root `tests.vba.json` contained legacy private/fake/lucky-data PCSUB entries and is now a redirect/debt marker.
- `tests/tests.vba.json` contains Slice 1+2 public atomic tests plus Slice A, Slice B1, Slice B2a, and Slice B2b1 characterization GREEN atoms.
- `tests/tests.vba.smoke.json` is smoke/debt only and must not be used as fixture-first acceptance.
