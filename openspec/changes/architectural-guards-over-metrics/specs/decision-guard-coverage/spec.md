# Decision guard coverage specification

> Capability nueva del change `architectural-guards-over-metrics` (issue #266). Sustituye el gate CRAP por guards que citan la decisión arquitectónica que protegen, más un gate mecánico (`scripts/check_decision_guards.py`) que impide que una decisión se quede sin guard.

## Purpose

Esta especificación describe el gate de cruce decisión↔guard, el patrón de guard test de dos tiers que ese gate exige, la retirada de CRAP y el ajuste del techo de complejidad que la acompañan, y la reconciliación obligatoria de `DA-1`. No cubre la implementación de `scripts/check_layers.py` (fuera de alcance) ni la escritura de guards para las once decisiones `DA-*` aún no cubiertas de hecho; esas quedan calendarizadas vía `BASELINE`.

## Requirements

### Requirement: Decision table traversal and ID-to-guard matching

The system SHALL traverse every `design.md` file under `openspec/`, including archived changes under `openspec/changes/archive/`, and SHALL extract every row of a decision table that carries an ID column (`DA-`, `QC-`, `D-`, or other declared prefixes). For each extracted ID, the gate SHALL search `tests/**` for a `HARNESS-PROVENANCE` header or a guard-test docstring naming that exact ID, and SHALL mark the decision covered when found.

#### Scenario: Archived design.md is traversed

- GIVEN a decision table with an ID column under `openspec/changes/archive/**/design.md`
- WHEN the gate runs
- THEN those IDs enter the subject set exactly as if the change were active

#### Scenario: Guard header names the decision

- GIVEN a test file carries `# HARNESS-PROVENANCE: ... encodes DA-13`
- WHEN the gate matches DA-13 against `tests/**`
- THEN DA-13 is marked covered by that guard test

### Requirement: covered-by escape hatch

When a decision ID appears in a `covered-by` entry naming an existing structural gate, the gate SHALL treat that decision as covered without requiring a dedicated guard test file.

#### Scenario: DA-1 covered by check_layers

- GIVEN the wiring table lists `check_layers` with `covered-by: DA-1`
- WHEN the gate evaluates DA-1
- THEN DA-1 is marked covered without a matching test header

#### Scenario: covered-by names a nonexistent gate

- GIVEN a `covered-by` entry names a gate absent from `GATES` in `quality_report.py`
- WHEN the gate evaluates that entry
- THEN it reports the entry as an invalid escape hatch and fails

### Requirement: BASELINE with mandatory target_date

A decision not yet covered by a guard or a `covered-by` entry MAY be listed in the gate's `BASELINE` with a mandatory ISO-8601 `target_date`. The gate SHALL fail once `target_date` has passed while the decision remains uncovered.

#### Scenario: BASELINE entry before its target date

- GIVEN a decision is listed in BASELINE with `target_date` in the future, still uncovered
- WHEN the gate runs before that date
- THEN it passes for that decision and reports it as calendared debt

#### Scenario: Expired BASELINE entry

- GIVEN a BASELINE entry's `target_date` has passed and the decision is still uncovered
- WHEN the gate runs
- THEN it fails and reports the expired entry

### Requirement: Exit codes and reporting

The gate SHALL exit 0 only when every extracted decision ID is covered by a guard test, a `covered-by` entry, or an unexpired `BASELINE` entry. It SHALL exit non-zero otherwise, and every failure line SHALL name the decision ID, the source `design.md` file and line, and the reason.

#### Scenario: All decisions resolved

- GIVEN every extracted ID is covered, escaped, or calendared
- WHEN the gate runs
- THEN it exits 0

#### Scenario: One decision without resolution

- GIVEN one extracted ID has no guard, no `covered-by` entry, and no BASELINE entry
- WHEN the gate runs
- THEN it exits non-zero and reports the ID with its source file and line

### Requirement: Fail-closed on malformed or absent input

The gate SHALL exit non-zero — never 0 — when: no decision table is found under any `design.md`; a decision table is malformed (missing or duplicated ID column); a `design.md` file cannot be read; or the extracted subject set of decision IDs is empty. A gate that finds nothing to inspect MUST NOT report success, matching the fail-closed precedent set by `check_crap.py` refusing to run without `coverage.json`.

#### Scenario: No decision table anywhere

- GIVEN no `design.md` under `openspec/` contains a table with an ID column
- WHEN the gate runs
- THEN it exits non-zero and reports the absence

#### Scenario: Malformed table

- GIVEN a `design.md` table is missing its ID column
- WHEN the gate parses it
- THEN it exits non-zero and names the file and the malformed table

#### Scenario: Empty subject set

- GIVEN parsing succeeds but yields zero decision IDs
- WHEN the gate evaluates that empty set
- THEN it exits non-zero instead of trivially passing

### Requirement: Two-tier guard test shape

Each guard test module SHALL contain a detector-correctness test asserting against synthetic clean and violating fixtures, plus a production assertion test running the detector over the closed set of files the decision names. Failure messages SHALL name the file, line, rule, and remedy.

#### Scenario: Detector correctness

- GIVEN synthetic clean and violating fixtures for a guard
- WHEN the detector-correctness test runs
- THEN it flags the violating fixture and does not flag the clean one

#### Scenario: Production assertion failure message

- GIVEN the detector finds a real violation in the closed file set
- WHEN the test reports it
- THEN the message names file, line, the violated rule, and the remedy

### Requirement: CRAP gate removal

The system SHALL NOT reference `crap` (case-insensitive) in `scripts/check_crap.py`, `scripts/quality_report.py` `GATES`, `Makefile`, `.github/workflows/ci.yml`, `app/pyproject.toml`, `app/README.md`, or any test file. The `crap_clean` and `crap_violation` fixture trees SHALL NOT exist.

#### Scenario: Repository-wide search finds nothing

- GIVEN a case-insensitive search for `crap` across tracked files
- WHEN it runs after this change
- THEN it returns zero matches

#### Scenario: GATES entry replaced

- GIVEN `scripts/quality_report.py` `GATES`
- WHEN it is inspected
- THEN it contains no `crap` entry and does contain `decision_guards`

### Requirement: Complexity ceiling lowered with review date

`MAX_COMPLEXITY` in `check_complexity.py` SHALL be 10. The module SHALL declare a review date of `2027-02-13` adjacent to the ceiling, documenting that the review decides whether the ceiling moves toward 6.

#### Scenario: Ceiling passes without code change

- GIVEN every current function is at or below complexity 6
- WHEN `check_complexity.py` runs with `MAX_COMPLEXITY = 10`
- THEN it exits 0 without any application code change

#### Scenario: Review date is present

- GIVEN `check_complexity.py` source
- WHEN it is inspected
- THEN `2027-02-13` appears adjacent to `MAX_COMPLEXITY`

### Requirement: DA-1 ROOT_PACKAGE reconciliation

`DA-1` in `lanzadera-mvp/design.md` SHALL declare the same `ROOT_PACKAGE` value that `check_layers.py`, `check_complexity.py`, `check_decision_guards.py`, and every other gate declaring that constant actually use. The decision-to-guard gate SHALL read `DA-1` and detect any discrepancy.

#### Scenario: DA-1 matches the real ROOT_PACKAGE

- GIVEN `check_layers.py` declares `ROOT_PACKAGE = "app"`
- WHEN `DA-1` is inspected
- THEN it declares `"app"`, not `"platform.src.modules"`

#### Scenario: A future discrepancy is detected

- GIVEN a later edit reintroduces a mismatch between `DA-1` and a gate's declared constant
- WHEN the decision-to-guard gate runs
- THEN it fails and names the mismatched value

## Cross-references

DA-1, DA-13, QC-2, QC-9, QC-11 (retirado con CRAP). Issue #266.
