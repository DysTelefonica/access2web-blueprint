# Exploration — architectural guards over metrics

Replace the CRAP gate with architectural guard tests, plus a mechanical
decision→guard gate. Closes the exploration phase for issue #266.

## Evidence base

The exploring executor had no Bash tool, so its CRAP findings rested on the
gate's committed `BASELINE` and comments. The orchestrator re-verified them with
a live run before this document was written:

```
python -m pytest --cov --cov-report=json:coverage.json   382 passed, 13 skipped
python scripts/check_crap.py                             OK — every function at or below CRAP 6
python scripts/check_complexity.py                       OK — every function at or below complexity 15
```

The static conclusion held: zero CRAP offenders today.

## 1. What CRAP catches today

`scripts/check_crap.py` carries `BASELINE = {}` with the comment «Empty by
design … CRAP <= 6 is the absolute ceiling. No offenders exist today». The live
run confirms it.

The consequence is not that CRAP is idle. It is the opposite:

```
CRAP(f) = CC(f)² · (1 − coverage(f))³ + CC(f)
```

The `(1 − coverage)³` term is never negative, so `CRAP(f) >= CC(f)` always,
collapsing to `CC(f)` exactly at full coverage. Every function passing
`CRAP <= 6` therefore has `CC <= 6`.

**The repository's effective complexity ceiling is 6 today, not the nominal 15
of `check_complexity.py`.** CRAP is doing the work of keeping functions small
while reporting zero violations. Removing it raises the ceiling from 6 to 15.

## 2. Overlap — what else would catch it

| Gate | Replaces CRAP? | Why |
|---|---|---|
| `check_complexity.py` (ceiling 15) | No | Allows CC 7-15 regardless of coverage |
| `--cov-fail-under=69` | No | Package-wide average, not per function |
| `check_mutation.py` (cosmic-ray) | Partially | Closest analogue, but weekly and not per function |
| `check_mutation_sites.py` (ceiling 100) | No | Static surface proxy, not coverage-weighted |
| `check_dry`, `check_layers`, `check_legacy_hashes` | No | Orthogonal concerns |

**Honest conclusion.** Removing CRAP is evidence-safe *today* — zero offenders
are lost. It is not evidence-safe *going forward*: no other gate enforces «branchy
code must be well tested» per function and per PR.

Guard tests cover only functions tied to a recorded architectural decision.
Business logic that never reaches a `design.md` has no successor gate. This is
the real cost of the decision and it is named here rather than assumed away.

## 3. The guard pattern

Extracted from four artifacts in the reference repository
(`C:\00repos\codigo\gentle-ai\00_main`): `adapter_forbidden_construction_guard_test.go`,
`candidate_readonly_guard_test.go`, `review_integration_contract_guard_test.go`
and `scripts/deadcode-ratchet.sh`.

Six reusable traits:

1. **Traceable header** naming the exact decision it encodes.
2. **Two tiers**: a detector-correctness test against synthetic clean and
   violating fixtures, so the scanner cannot be vacuously green; plus a
   production assertion against real files.
3. **Closed file sets over open globs** when scope matters, so an unrelated file
   is never swept in silently.
4. **Failure messages that teach the rule**: file, line, category and remedy.
5. **Explicit non-goals** when the check is necessarily partial.
6. **The must-check set extracted mechanically from source** where feasible, so
   the guard cannot go stale unnoticed.

This repository already writes the pattern informally:
`tests/lanzadera/auth/test_no_legacy_compat_smoke.py` carries a
`HARNESS-PROVENANCE` header citing `DA-13`, and the `test_<gate>_wiring.py`
family pairs with `tests/fixtures/<gate>_clean` and `<gate>_violation` trees.

Proposed Python shape:

```python
# tests/<module>/<slice>/test_<decision-slug>_guard.py
# HARNESS-PROVENANCE: architectural-guards-over-metrics — encodes DA-<n>
"""Guard for DA-<n>: <one-line decision text copied from design.md>."""

def test_<rule>_guard_catches_known_shapes():
    # table-driven: flags a synthetic violating fixture,
    # does not flag a synthetic clean one.

def test_<rule>_guard_holds_for_production_files():
    # real scanner over the closed file set the decision names;
    # fails with file:line plus the rule text.
```

## 4. The decision→guard gate

The central question. Without it, «write a guard per decision» is a convention
people remember, and the team ends up with neither CRAP nor guards.

`openspec/changes/lanzadera-mvp/design.md` already has this half-built: an
architectural decision table with an ID column (`DA-1`..`DA-13`) at line 144, and
a «Quality gates wiring» table at line 285 mapping gate → mechanism → files →
decisions. That second table is a `covered-by` registry in embryonic form.

| Option | Mechanism | Falsos positivos | Mantenimiento | Genericidad |
|---|---|---|---|---|
| **A. Convención de identificador** | El gate lee las filas con ID de cada `design.md` y busca ese token en las cabeceras de `tests/**`. Una decisión ya listada en la tabla de wiring cuenta como cubierta por ese gate | Bajos-medios; mitigados por la vía de escape que el repo ya escribe a mano | Bajo; mismo idioma `BASELINE` + `target_date` que los seis gates existentes | Alta; sólo exige tabla con columna de ID y cabecera que la nombre |
| **B. Fichero de registro** | `.sdd/decision-guards.yml` mapeando `DA-N` a rutas de test o a `covered-by` | Bajos en el registro, pero dos ficheros pueden divergir | Medio; artefacto nuevo y sincronización nueva | Muy alta; datos puros, sin depender de la forma del markdown |
| **C. Recorrido por nombre sin ID** | Slug del encabezado de cada decisión; exige un test homónimo | **Altos**; renombrar la prosa rompe el enlace en silencio, y aquí las decisiones son filas con ID, no encabezados | Medio-alto; frágil ante ediciones de prosa | Menor para repos que ya usan tablas con ID |

**Recomendación: opción A.** Formaliza lo que las cabeceras `HARNESS-PROVENANCE`
y la tabla de wiring ya hacen a mano, y se expresa como regla genérica: *toda
fila con ID de la tabla de decisiones debe estar referenciada por su nombre en la
cabecera de un guard test o en la entrada `covered-by` de un gate*.

La opción B queda como vía de mejora para un repositorio cuyo `design.md` no
tenga tabla con ID, o cuando los metadatos de exención necesiten un sitio fuera
de los comentarios de test.

## 5. Universal frente a específico

**Universal — va a la skill `deterministic-quality-harness`:** la forma del gate
de cruce decisión↔guard; el idioma `BASELINE` más `target_date` obligatorio; la
forma de dos tiers del guard test; la regla de conjunto cerrado frente a glob; y
la vía de escape `covered-by` para decisiones ya cubiertas por un gate estructural.

**Específico del repositorio — bloque de configuración, igual que ya hace
`check_layers.py`:** `ROOT_PACKAGE`, los prefijos de módulo, los nombres de capa
hexagonal, los prefijos de ID en uso (`DA-`, `QC-`, `D-`), la raíz de `tests/`, y
la decisión concreta de retirar CRAP, que otro repositorio podría no tomar.

## Áreas afectadas

- `scripts/check_crap.py`, su entrada en el `GATES` de `scripts/quality_report.py`,
  `tests/lanzadera/test_crap_wiring.py`, `tests/fixtures/crap_clean/`,
  `tests/fixtures/crap_violation/` — superficie de retirada.
- `scripts/check_complexity.py` — pasa a ser el único techo de complejidad.
- `.github/workflows/ci.yml` — comentario de la tubería de gates.
- `openspec/changes/lanzadera-mvp/design.md` — origen de la tabla de decisiones y
  plantilla para las futuras.
- Nuevo: `scripts/check_decision_guards.py`, cableado en `GATES`.

## Riesgos

1. **Pérdida de la garantía por función y por PR** para complejidad no ligada a
   una decisión registrada. Real, no hipotética.
2. **Disciplina de etiquetado.** La opción A depende de que la cabecera del guard
   nombre el ID. Sin la vía de escape `covered-by`, los falsos positivos serían
   altos para decisiones cubiertas por un gate estructural.
3. **Techo de complejidad efectivo.** Al retirar CRAP el techo real pasa de 6 a
   15. Si el objetivo era aliviar la fricción, ese es el efecto buscado; conviene
   que sea una decisión consciente y no una consecuencia descubierta después.

## Listo para propuesta

Sí. La pregunta central tiene recomendación respaldada por evidencia, el patrón
de guard tiene plantilla apoyada en código que el repositorio ya escribe, y el
coste de retirar CRAP queda nombrado en lugar de disimulado.
