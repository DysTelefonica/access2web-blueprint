# Tareas: guards arquitectónicos en lugar de métricas derivadas

Issue #266. Rama tracker: `feat/266-guards-arquitectonicos` (ya existente, destino final `main`, draft/no-merge hasta que las seis rebanadas estén integradas). **Solo la tracker se fusiona a `main`**; cada rebanada apunta a la rama de la anterior.

## Review Workload Forecast

| Field | Value |
|-------|-------|
| Estimated changed lines | ~1660 (30+250+500+380+300+200) |
| 400-line budget risk | High |
| Chained PRs recommended | Yes |
| Suggested split | PR 1 → PR 2 → PR 3 (size:exception) → PR 4 → PR 5 → PR 6 |
| Delivery strategy | auto-chain |
| Chain strategy | feature-branch-chain |

Decision needed before apply: No
Chained PRs recommended: Yes
Chain strategy: feature-branch-chain
400-line budget risk: High

**Verdicto sobre la rebanada 3 (`size:exception`)**: se acepta. Es borrado puro de código ya revisado (`check_crap.py`, dos árboles de fixture, un wiring test) sin lógica nueva; se verifica mecánicamente con `git grep -i crap` (cero coincidencias) y la suite completa en verde. Partirla más (script vs. fixtures) fragmentaría un único límite de revert atómico sin reducir carga cognitiva real. `size-exception-reason`: `pure-deletion-of-reviewed-code`.

### Suggested Work Units

| Unit | Rama | Base (target PR) | Presupuesto | Focused test command | Runtime harness | Rollback boundary |
|---|---|---|---|---|---|---|
| 1 | `feat/266-complexity-ceiling` | `feat/266-guards-arquitectonicos` | ~30 | `python -m pytest tests/lanzadera/test_complexity_wiring.py -q` | `python scripts/check_complexity.py --root .` | revert restaura `MAX_COMPLEXITY = 15` |
| 2 | `feat/266-crap-unwire` | `feat/266-complexity-ceiling` | ~250 | `python -m pytest tests/test_ci_workflow.py tests/test_gate_smoke.py -q` | `python scripts/quality_report.py` | revert restaura `crap` en `GATES`, `Makefile`, `ci.yml`, `pyproject.toml`, docs y expectativas de test |
| 3 | `feat/266-crap-delete` (`size:exception`) | `feat/266-crap-unwire` | ~500, borrado puro | `python -m pytest tests/lanzadera/test_crap_retired_guard.py -q` | `git grep -i crap` (0 matches) + `python scripts/quality_report.py` | revert restaura `check_crap.py`, sus fixtures y `test_crap_wiring.py` |
| 4 | `feat/266-decision-guards-gate` | `feat/266-crap-delete` | ~380 | `python -m pytest tests/lanzadera/test_decision_guards_wiring.py -q` | N/A — gate creado sin cablear en `GATES`; verificado solo por tests unitarios sobre cadenas sintéticas | revert elimina `check_decision_guards.py` y sus tests unitarios; nada más lo referencia aún |
| 5 | `feat/266-decision-guards-fixtures` | `feat/266-decision-guards-gate` | ~300 | `python -m pytest tests/lanzadera/test_decision_guards_wiring.py -q` | `python scripts/check_decision_guards.py --root tests/fixtures/decision_guards_clean --json` y `--root tests/fixtures/decision_guards_violation --json` | revert elimina los árboles de fixture y los tests tier 1/2 añadidos |
| 6 | `feat/266-decision-guards-wire` | `feat/266-decision-guards-fixtures` | ~200 | `python -m pytest tests/lanzadera/test_quality_report_smoke.py tests/test_ci_workflow.py -q` | `python scripts/quality_report.py` (exit 0, `decision_guards` presente) | revert quita la entrada de `GATES`, `COVERED_BY`/`BASELINE`, cableado CI/Makefile y la corrección de `DA-1`/`DA-13`; el gate queda dormido pero presente |

Orden obligatorio: 1 antes que 2 y 3 (con CRAP vivo el techo efectivo es 6; en orden inverso el techo flota a 15 durante toda la cadena). 6 va última (armar `GATES` antes de poblar `COVERED_BY`/`BASELINE` deja CI en rojo al fusionar).

## Phase 1: Techo de complejidad 15→10 — PR 1 (`feat/266-complexity-ceiling` → `feat/266-guards-arquitectonicos`)

Traza: DG-11, spec "Complexity ceiling lowered with review date".

- [x] 1.1 `scripts/check_complexity.py`: `MAX_COMPLEXITY = 10`, fecha de revisión `2027-02-13` junto a la constante.
- [x] 1.2 `tests/lanzadera/test_complexity_wiring.py`: pinear `MAX_COMPLEXITY == 10` y presencia de `2027-02-13` en el módulo.
- [x] 1.3 Verificar: `python scripts/check_complexity.py --root .` sale 0 sin tocar código de `app/`.
- [x] 1.4 Verificar límite de rebanada: `python scripts/quality_report.py` y `python -m pytest -q` en verde.

## Phase 2: Descableado de CRAP — PR 2 (`feat/266-crap-unwire` → `feat/266-complexity-ceiling`)

Traza: DG-12, spec "CRAP gate removal" (paso 1 de 2).

- [ ] 2.1 `scripts/quality_report.py`: quitar `("crap", "check_crap.py", ())` de `GATES`.
- [ ] 2.2 `Makefile`, `.github/workflows/ci.yml`, `app/pyproject.toml`, `app/README.md`, `.gitignore`: quitar referencias a `crap`.
- [ ] 2.3 Docstrings de `check_complexity.py` y `check_mutation.py`: quitar mención a CRAP como "binding constraint".
- [ ] 2.4 `tests/test_ci_workflow.py`: quitar `"crap"` de `REQUIRED_GATE_ORDER`.
- [ ] 2.5 `tests/test_gate_smoke.py`: quitar `"check_crap.py"` del parametrize de `test_every_gate_emits_a_well_formed_envelope`; retargetear `test_quality_report_aggregates_indicators`/`test_quality_report_names_the_failing_gate`/`test_quality_report_is_byte_identical_for_the_same_commit` fuera de `crap_clean`/`crap_violation` y de la aserción `crap.max_crap`.
- [ ] 2.6 Verificar límite de rebanada: `python scripts/quality_report.py` sin `crap` en el envelope, `python -m pytest -q` en verde (script y fixtures de CRAP siguen en disco, solo descableados).

## Phase 3: Borrado de CRAP — PR 3, `size:exception` (`feat/266-crap-delete` → `feat/266-crap-unwire`)

Traza: DG-12, spec "CRAP gate removal" (paso 2 de 2).

- [ ] 3.1 Borrar `scripts/check_crap.py`, `tests/fixtures/crap_clean/`, `tests/fixtures/crap_violation/`, `tests/lanzadera/test_crap_wiring.py`.
- [ ] 3.2 `tests/test_gate_smoke.py`: borrar la sección `check_crap` (tests que invocan `check_crap.py` directamente).
- [ ] 3.3 Crear `tests/lanzadera/test_crap_retired_guard.py` (DG-12): falla si `crap` reaparece en `GATES`, `Makefile`, `ci.yml`, `pyproject.toml` o docs.
- [ ] 3.4 Verificar: `git grep -i crap` sobre ficheros rastreados devuelve cero coincidencias.
- [ ] 3.5 Verificar límite de rebanada: `python scripts/quality_report.py` y `python -m pytest -q` en verde.

## Phase 4: Gate nuevo, sin armar — PR 4 (`feat/266-decision-guards-gate` → `feat/266-crap-delete`)

Traza: DG-1..DG-4, DG-7, DG-13, spec "Decision table traversal", "Fail-closed on malformed or absent input", "Two-tier guard test shape" (RED, nivel unitario).

- [ ] 4.1 RED: `tests/lanzadera/test_decision_guards_wiring.py` — tests unitarios sobre cadenas sintéticas para `parse_decision_tables()` (DG-2, DG-13: cabecera `id`, fila delimitadora, fila de decisión de token único vs. celda-lista) antes de que exista el módulo.
- [ ] 4.2 GREEN: crear `scripts/check_decision_guards.py` con bloque `CONFIGURATION` (`OPENSPEC_ROOT`, `TESTS_ROOT`, `SCRIPTS_ROOT`, `DECISION_ID_PREFIXES = ("DA-", "DG-", "QC-")`, `EXCLUDED_PARTS` con `fixtures`, `UNGOVERNED_DESIGNS`, `COVERED_BY`, `BASELINE` vacíos), `parse_decision_tables()`, `collect_guard_claims()` (bloque `HARNESS-PROVENANCE`, no docstring — DG-4), `validate_coverage()`, `evaluate()`, `build_report()`, `--root/--json`.
- [ ] 4.3 GREEN: tests unitarios adicionales para fallo cerrado (DG-3: sin tabla, IDs vacíos) y unicidad global de ID (DG-7: `duplicate_id`).
- [ ] 4.4 Verificar límite de rebanada: `python -m pytest tests/lanzadera/test_decision_guards_wiring.py -q` en verde; el gate no figura aún en `GATES`, así que `python scripts/quality_report.py` no lo invoca.

## Phase 5: Fixtures de dos tiers — PR 5 (`feat/266-decision-guards-fixtures` → `feat/266-decision-guards-gate`)

Traza: DG-5, DG-6, DG-8, DG-10, DG-13, spec "covered-by escape hatch", "BASELINE with mandatory target_date", "Two-tier guard test shape" (RED del caso de amenaza aplicable).

- [ ] 5.1 Crear `tests/fixtures/decision_guards_clean/`: `openspec/` mínimo con una decisión y su guard, más una tabla de cableado con celdas-lista (cabecera `Decisiones`) que el gate debe ignorar (DG-13).
- [ ] 5.2 Crear `tests/fixtures/decision_guards_violation/`: ID descubierto, tabla malformada, guard huérfano, `covered-by` obsoleto.
- [ ] 5.3 RED (amenaza única aplicable — rutas tipo documentación): fixture con cabecera `HARNESS-PROVENANCE` falsa bajo `tests/fixtures/` que cita un ID inexistente; assertar que el recorrido de producción (`--root .`) no la ve gracias a `EXCLUDED_PARTS = {"fixtures", ...}`.
- [ ] 5.4 GREEN: extender `test_decision_guards_wiring.py` con los tests tier 1 (`--root` sobre `decision_guards_violation`, exit 1 y clave de cada verdicto: `malformed_table`, `malformed_row`, `duplicate_id`, `uncovered`, `stale_coverage`, `constant_drift`, `orphan_guard`, `BASELINE` vencida) y tier 2 (`--root` sobre `decision_guards_clean`, exit 0, `decisions_total > 0`).
- [ ] 5.5 Verificar límite de rebanada: `python -m pytest tests/lanzadera/test_decision_guards_wiring.py -q` en verde; gate aún sin GATES.

## Phase 6: Armado — PR 6 (`feat/266-decision-guards-wire` → `feat/266-decision-guards-fixtures`)

Traza: DG-6, DG-9, DG-11, spec "DA-1 ROOT_PACKAGE reconciliation", "Exit codes and reporting", "GATES entry replaced".

- [ ] 6.1 `scripts/quality_report.py`: añadir `("decision_guards", "check_decision_guards.py", ())` al final de `GATES`; nuevos `INDICATOR_MEANINGS`.
- [ ] 6.2 `check_decision_guards.py`: poblar `COVERED_BY` (`DG-11→check_complexity.py asserts={"MAX_COMPLEXITY":10}`; `DA-1→check_layers.py asserts={"ROOT_PACKAGE":"app"}`; `DA-13→check_legacy_hashes.py`) y `BASELINE` para las once `DA-2..DA-12` con `target_date="2027-02-13"`.
- [ ] 6.3 `openspec/changes/lanzadera-mvp/design.md`: corregir `DA-1` a `ROOT_PACKAGE = "app"` (nota de corrección fechada); `DA-13` sin rutas `platform/`; sustituir la fila CRAP de la tabla de wiring por la fila del gate nuevo.
- [ ] 6.4 `tests/lanzadera/test_quality_report_smoke.py`: ampliar cabecera `HARNESS-PROVENANCE` a `… + architectural-guards-over-metrics DG-9`; pinear posición final de `decision_guards` en `GATES` (DG-9).
- [ ] 6.5 `tests/test_ci_workflow.py`: añadir `"decision_guards"` al final de `REQUIRED_GATE_ORDER`.
- [ ] 6.6 `.github/workflows/ci.yml`, `Makefile`: cablear el paso/target del gate nuevo.
- [ ] 6.7 Verificar criterios de éxito de la propuesta: `python scripts/quality_report.py --root .` exit 0, `decisions_uncovered == 0`, `orphan_guards == 0`; `python -m pytest -q` completo en verde (referencia local: 382 passed, 13 skipped, ~21 s).

## Next Step

Tracker `feat/266-guards-arquitectonicos` se fusiona a `main` solo tras integrar las seis rebanadas en orden. Listo para `sdd-apply` con la rebanada 1.
