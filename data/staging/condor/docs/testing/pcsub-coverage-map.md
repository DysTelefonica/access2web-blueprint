# PCSUB service/repository coverage map

Change: `fixture-first-coverage-expansion`  
Current slice: Slice B2b1 — `GuardarImpacto` assertion-hardening characterization GREEN

## Classification legend

- `tested`: covered by `tests/tests.vba.json` atomic JSON tests.
- `happy-path gap`: important behavior with no deterministic positive test yet.
- `sad-path/edge gap`: important negative/edge behavior selected for a later slice.
- `characterization green`: negative/edge behavior now covered by tests that passed against already implemented production behavior; no production change was needed.
- `coupled debt`: behavior needs wider ViewModel, workflow, supplier, rejection, user, or expediente fixture graph before safe fixture-first testing.
- `trivial/no-test`: alias, mapper, or thin delegation with no standalone business value beyond covered callees.

## Service behavior

| Area/method | Current status | Classification | Slice/debt | Fixture shape | Evidence |
|---|---|---|---|---|---|
| `getDatosPCSUB` | Reads exact seeded PCSUB row. | tested | Baseline Slice 2 | `TbExpedientes -> tbSolicitudes -> tbDatosPCSUB` | `Test_PCSUB_GetDatos_ReturnsSeededEntity` |
| `GuardarDatosGenerales` | Persists seeded general fields without transition in `estadoRegistro`. | tested | Baseline Slice 2 | Same graph, transition-safe state | `Test_PCSUB_GuardarDatosGenerales_PersistsFields` |
| `EsDatosGeneralesCompleta` complete row | Positive completeness. | tested | Baseline Slice 2 | Complete general fields | `Test_PCSUB_EsDatosGeneralesCompleta_TrueForSeededCompleteRow` |
| `EsDatosGeneralesCompleta` missing fields | Negative completeness with positive control and cardinality assertions. | characterization green | Slice A accepted characterization GREEN | Complete control PCSUB row plus target PCSUB row with missing business general fields | `Test_PCSUB_EsDatosGeneralesCompleta_FalseWhenGeneralFieldsMissing` |
| `EsDatosGeneralesCompleta` no PCSUB row | Negative child-row absence with no-fallback control. | characterization green | Slice A accepted characterization GREEN | Complete control PCSUB row plus target parent solicitud only, no `tbDatosPCSUB` row | `Test_PCSUB_EsDatosGeneralesCompleta_FalseWhenNoPCSUBRow` |
| `EsDetalleCompleto` complete detail | Positive completeness. | tested | Baseline Slice 2 | Detail fields populated | `Test_PCSUB_EsDetalleCompleto_TrueForSeededDetail` |
| `EsDetalleCompleto` missing detail | Negative completeness with positive control and cardinality assertions. | characterization green | Slice A accepted characterization GREEN | Complete control PCSUB row plus target PCSUB row with empty detail fields | `Test_PCSUB_EsDetalleCompleto_FalseWhenDetailFieldsMissing` |
| `EsMotivosCompleto` one motive | Positive completeness. | tested | Baseline Slice 2 | `motivoCorregirDeficiencias=True` | `Test_PCSUB_EsMotivosCompleto_TrueForSeededMotivo` |
| `EsMotivosCompleto` no motives | Negative completeness with positive control and cardinality assertions. | characterization green | Slice A accepted characterization GREEN | Complete control PCSUB row plus target row with all motive flags `False` | `Test_PCSUB_EsMotivosCompleto_FalseWhenNoMotives` |
| `GuardarPropuesta` persistence + missing `descripcionPropuestaCambio` | Persists proposal fields and rejects a blank `descripcionPropuestaCambio` without side effects. This is B1's exact sad-path coverage; it is not full material/propuesta validation coverage. | characterization green | Slice B1 accepted characterization GREEN | Existing graph plus proposal fields, positive persisted row, missing-`descripcionPropuestaCambio` target with unchanged persisted state | `Test_PCSUB_GuardarPropuesta_PersistsFields`; `Test_PCSUB_GuardarPropuesta_FailsWhenProposalMissing` |
| `GuardarPropuesta` missing `descripcionMaterialAfectado` | Rejects blank material when proposal is populated and preserves the asserted proposal/material fields plus target `tbDatosPCSUB` cardinality. This is not a blanket no-side-effect proof across every column. | characterization green | Slice B2a accepted characterization GREEN | Same proposal graph with material description blank and proposal description populated | `Test_PCSUB_GuardarPropuesta_FailsWhenMaterialMissing`; Dysflow `sliceB2a` 5/5. |
| `GuardarImpacto` asserted persistence + missing motives | B2a persists/rejects the scoped asserted subset and is accepted characterization GREEN. B2b1 is accepted characterization GREEN assertion-hardening after Dysflow runtime 2/2 and Judgment Day PASS WITH WARNINGS: it compares all GuardarImpacto-owned fields against the current fixture values and proves the no-motives rejected-save preserves the pre-Act impact snapshot plus parent/child cardinality. It does not prove every field transitioned away from seed/default values, and it does not cover every deferred invalid-domain rejection variant. | characterization green | Slice B2b1 accepted assertion-hardening only; B2b2 remains open | Existing graph plus impacto fields: all motives, motive detail, both incidence enums, all incidence booleans, classification, affected-material decision | B2a: `Test_PCSUB_GuardarImpacto_PersistsFields`; `Test_PCSUB_GuardarImpacto_FailsWhenNoMotives`; Dysflow `sliceB2a` 5/5. B2b1: `Test_PCSUB_GuardarImpacto_PersistsAllImpactFields`; `Test_PCSUB_GuardarImpacto_RejectedSavePreservesAllImpactFields`; Dysflow `sliceB2b1` 2/2; Judgment Day PASS WITH WARNINGS/no CRITICAL. |
| `EsParteTecnicaCompleta` proposal complete / impact incomplete | Returns `False` for a row whose proposal fields are complete while impact/motive data remains incomplete, with a complete positive control proving the same method can return `True`. | characterization green | Slice B1 accepted characterization GREEN | Complete technical control row plus target row with proposal complete and impact/motives incomplete | `Test_PCSUB_EsParteTecnicaCompleta_FalseWhenImpactIncomplete`; B2 still owns impact-side persistence/validation/completeness coverage. |
| `EsParteTecnicaCompleta` impact-side variants | Complete impact returns `True`; missing `impactoClasificacion` returns `False` with explicit complete positive control and exact target `tbDatosPCSUB` cardinality. Parent cardinality is deterministic but weaker than the strictest fixture-first bar. | characterization green | Slice B2a accepted characterization GREEN | Proposal + impact complete/incomplete variants | `Test_PCSUB_EsParteTecnicaCompleta_TrueWhenImpactComplete`; `Test_PCSUB_EsParteTecnicaCompleta_FalseWhenImpactClassificationMissing`; Dysflow `sliceB2a` 5/5. Additional missing-incidence/material-decision completeness variants remain B2b gaps. |
| `GuardarAprobacionSuministrador` / `EsAprobacionSuministradorCompleta` | Approval-sub fields and completeness. | happy-path gap | Slice C | Existing graph plus approval-sub fields | Needs positive + missing signer tests. |
| `GuardarDictamenRAC` / `EsDictamenRACCompleto` | RAC fields and conditional code requirement. | happy-path gap | Slice C | Existing graph plus RAC fields | Needs accepted/rejected decision variants. |
| `GuardarDecisionFinal` / `EsDecisionFinalCompleta` | Final decision fields and completeness. | happy-path gap | Slice C | Existing graph plus final decision fields | Needs positive + empty decision tests. |
| `PurgaTecnica` | Clears proposal, motive, impact fields. | happy-path gap | Slice D | Fully populated technical row | Needs before/after field assertions. |
| `EliminarDictamenRAC`, `EliminarAprobacionSuministrador`, `EliminarDecisionFinal`, `EliminarPorIdSolicitud` | Cleanup/delete operations. | happy-path gap | Slice D | Populated rows, then cleanup/delete | Needs side-effect assertions and teardown safety. |
| `GetDatosGeneralesIniciales`, `GetDatosAprobacionIniciales`, `ObtenerViewModelCompleto`, `ActualizarCamposDependientesDeExpediente` | Reads wider expediente/supplier/environment/ViewModel graph. | coupled debt | Separate proposal or focused seam | Requires `Solicitud`, `Expediente`, supplier relations, globals, workflow state | Not safe in Slice A without expanding fixture graph. |
| `ImportarEntidad_aVM`, `ExportarVM_aEntidad` | Field mapping to/from ViewModel. | coupled debt | Separate ViewModel mapping slice | Requires `DatosPCSUBViewModel` fixture and enum block matrix | Can be unit-tested later with pure mapping fixtures. |
| `RegistrarRechazoDesdeFormalizacion` | Rejection workflow transition and data reset. | coupled debt | Separate workflow/rejection slice | Requires rejection/workflow/user fixture graph | Out of current change scope per proposal. |

## Repository behavior

| Area/method | Current status | Classification | Slice/debt | Fixture shape | Evidence |
|---|---|---|---|---|---|
| `getPorIdSolicitud` / `GetById` found | Read path through service. | tested | Baseline Slice 2 | Seeded PCSUB row | `Test_PCSUB_GetDatos_ReturnsSeededEntity` |
| `getPorIdSolicitud` not found | Returns `Nothing` through service completeness checks. | characterization green | Slice A accepted characterization GREEN | Parent solicitud only plus complete no-fallback control row | `Test_PCSUB_EsDatosGeneralesCompleta_FalseWhenNoPCSUBRow` |
| `Guardar` update | Update through service. | tested | Baseline Slice 2 | Seeded PCSUB row | `Test_PCSUB_GuardarDatosGenerales_PersistsFields` |
| `Guardar` insert / generated id | Insert branch not directly covered. | happy-path gap | Slice E | Parent solicitud plus new `DatosPCSUB` object | Needs repository-focused insert test. |
| `ActualizarDatosGenerales` | Uses implicit `getdb()` and has no `db` parameter. | coupled debt | Refactor seam before test | Needs explicit db seam or guarded sandbox routing proof | Avoid in Slice A. |
| `ActualizarPropuesta`, `ActualizarImpacto`, approval/RAC/final update methods | Direct section updates and no-row errors. | happy-path gap | Slices B/C/E | Seeded PCSUB row and no-row variants | Add by behavior cluster. |
| `LimpiarDictamenRAC`, `LimpiarAprobacionSuministrador`, `LimpiarDecisionFinal`, `EliminarPorIdSolicitud` | Cleanup/delete side effects. | happy-path gap | Slice D | Populated rows, then cleanup/delete | Assert exact cleared/deleted fields. |
| `getSiguienteIDDatosPCSUB` | Private implementation detail. | trivial/no-test | N/A | Covered by `Guardar` insert later | Do not test directly. |

## Slice A fixture discipline

Slice A uses only the existing schema-valid graph documented in `docs/testing/pcsub-fixture-graph.md`:

1. `TbExpedientes`
2. `tbSolicitudes`
3. `tbDatosPCSUB` when the scenario requires a child row

The no-row sad path deliberately seeds only `TbExpedientes -> tbSolicitudes` for the target row and asserts the service returns incomplete rather than relying on pre-existing absence.

The Slice A sad-path helper now also seeds a complete positive-control `TbExpedientes -> tbSolicitudes -> tbDatosPCSUB` row before every negative assertion. Each negative test first proves the scoped completeness method can return `True` for the control row, then asserts the target row returns `False` while the control still exists. This hardens the tests against accidental fallback to unrelated complete rows and verifies target/control `tbDatosPCSUB` cardinality explicitly. The hardened Slice A run passed 4/4 and is accepted as characterization GREEN for behavior already implemented before this slice; no production code change was needed.

## Slice B1 fixture discipline

Slice B1 uses the same `TbExpedientes -> tbSolicitudes -> tbDatosPCSUB` graph and is accepted as characterization GREEN after targeted Dysflow evidence passed 3/3 with no production code change. The B1 atoms cover exact proposal persistence, missing-`descripcionPropuestaCambio` rejection with no side effects, and proposal-side technical completeness returning `False` while impact data remains incomplete. B1 does not cover the sibling missing-`descripcionMaterialAfectado` validation even though production validates it; track that as a B1/B2 follow-up before claiming full material/propuesta validation coverage. B2 remains responsible for impact persistence, invalid impact domains, affected-material decisions, and impact-side completeness behavior.

## Slice B2a fixture discipline

B2 is split to stay under the review budget. B2a is accepted as characterization GREEN after Judgment Day PASS WITH WARNINGS and Dysflow `sliceB2a` runtime evidence of 5 total / 5 passed / 0 failed / 0 blocked. The accepted B2a behaviors are only the exact assertions present in the atoms: missing `descripcionMaterialAfectado`, asserted `GuardarImpacto` persistence, no-motives validation preserving asserted fields, and `EsParteTecnicaCompleta` impact-side true/missing-classification false controls.

Explicit remaining gaps/debt:

- Invalid `incidenciaCoste`, invalid/missing `incidenciaPlazo`, invalid `impactoClasificacion` in `GuardarImpacto`, and invalid `CambioAfectaAMaterial` variants.
- Parent cardinality strengthening to the strictest fixture-first bar; current B2a assertions are deterministic and not lucky-data dependent but focus on target `tbDatosPCSUB` cardinality.

## Slice B2b1 fixture discipline

B2b is split to protect the review budget. B2b1 is accepted as characterization GREEN assertion-hardening after manual compile, Dysflow `sliceB2b1` runtime evidence of 2 total / 2 passed / 0 failed / 0 blocked, and Judgment Day PASS WITH WARNINGS/no CRITICAL. No production GREEN fix was needed. The atoms reuse the existing `TbExpedientes -> tbSolicitudes -> tbDatosPCSUB` graph and add stricter assertions for the impact cluster:

- Exact impact-field comparison: `Test_PCSUB_GuardarImpacto_PersistsAllImpactFields` seeds one deterministic target graph, populates the GuardarImpacto-owned motive/incidence/classification/material-decision fields, calls `GuardarImpacto`, then asserts those fields match the current fixture values plus exact parent/child cardinality. Because some expected values may equal seed/default values, this is not a proof that every field transitioned away from its previous value.
- Rejected-save no-motives snapshot preservation: `Test_PCSUB_GuardarImpacto_RejectedSavePreservesAllImpactFields` snapshots a complete persisted row, mutates every impact field while making motives invalid, expects validation before persistence, then asserts every persisted impact field still matches the pre-Act snapshot plus exact parent/child cardinality. This proves no-side-effects for the no-motives rejection shape only.

Remaining B2b2 gaps/debt:

- Unchanged-field transition proof for impact fields whose expected values match seed/default values.
- Invalid `incidenciaCoste`, invalid/missing `incidenciaPlazo`, invalid `impactoClasificacion`, and invalid `CambioAfectaAMaterial` variants.
