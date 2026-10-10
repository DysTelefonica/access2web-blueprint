# Gates de CI: contrato, manual y cómo añadir uno

> **Qué es.** El manual de uso de los gates deterministas del repositorio: el contrato global (qué familias hay y cómo llega un PR al merge), qué mide cada `scripts/check_*.py`, cómo se lee su fallo, el contrato de los cinco workflows, el `mutation` semanal y los pasos para añadir un gate nuevo.

> **Qué no es.** No sustituye a `docs/architecture.md` (qué decisiones del sistema protege cada gate, el patrón hexagonal, los core invariants), ni a `CONTRIBUTING.md` (cómo se trabaja aquí, tamaño de los PRs, convenciones), ni a `.github/dependabot.yml` (política de dependencias). Este documento es **el uso de los gates**.

## 1. El contrato global

Los gates se agrupan en tres familias, y cada una corre en un sitio distinto:

| Familia | Qué protege | Dónde corre |
|---|---|---|
| Forma del cambio | nombre de rama, tamaño del PR, contrato del workflow de CI, esquema de los walkthroughs | job `review-budget` (barato: segundos, tope 5 min) |
| Pruebas y calidad del código | formato y lint, capas hexagonales, complejidad, DRY, hashes legacy, retirada legacy, densidad de sitios de mutación | job `quality` (tope 30 min) |
| Contrato del host y del arnés | protección de rama, jobs requeridos, paridad local, política del workflow y matriz de reglas | fase 4 del contrato de adopción: `ci-pattern adoption check --phase 4 .` |

El recorrido de un PR hasta el merge:

```text
PR  (rama `tipo/nº-slug`, cuerpo con `## Chain Context`, ≤400 líneas o excepción declarada)
 │  push a la rama
 ▼
GitHub Actions
 ├── review-budget ──┐
 ├── quality ────────┤
 ├── security ───────┼──► required  (agregador; needs: los cuatro; if: always())
 └── codeql ─────────┘        │
                              │  único contexto requerido por la protección de `main`
                              ▼
                    merge del mantenedor (ruleset `main-maintainers-and-admins-merge`)
```

Cuatro reglas que el diagrama no muestra y que conviene saber:

- **`required` es el único contexto requerido.** Los demás jobs publican su nombre, pero la protección de rama solo exige `required`; por eso un job que no queremos bloqueante no se añade a su lista de `needs`.
- **`required` lleva `if: always()`**: sin él, un dependency saltado (por ejemplo `review-budget` fuera de un `pull_request`) marcaría el propio `required` como *skipped*, y un check requerido saltado no bloquea el merge.
- **`security.yml` y `codeql.yml` son *reusable workflows*** (`workflow_call`): `needs:` no cruza archivos de workflow, así que la única forma de que `required` dependa de sus job ids es invocarlos con `uses:` desde `ci.yml`.
- **Cada job declara `timeout-minutes`**: un gate que se cuelga es un gate que no mide, y el tope convierte el cuelgue en un fallo con diagnóstico.

## 2. Los `check_*.py`, uno por uno

Catorce scripts viven en `scripts/`; su contrato es siempre el mismo: leen datos estructurados (rama, etiquetas, campos de la API, YAML, AST), nunca prosa ni gestos, y publican un veredicto con código de salida.

| Script | Qué mide | Cómo se lee su fallo | Dónde corre |
|---|---|---|---|
| `check_layers.py` | dirección de dependencias entre capas hexagonales, vertical slicing y pureza de capa | nombra cada violación y su clase; publica cuánto comprobó de verdad (un árbol que no pudo clasificar es fallo, no verde) | `quality_report.py` |
| `check_complexity.py` | complejidad ciclomática por función, con techo global absoluto | indica la función y su valor frente al techo | `quality_report.py` |
| `check_mutation_sites.py` | densidad estática de sitios mutables | indica los ficheros con densidad fuera de rango | `ci.yml` + `quality_report.py` |
| `check_dry.py` | bloques duplicados tipo 1 y 2 sobre el AST | nombra los pares de bloques y su ubicación | `ci.yml` + `quality_report.py` |
| `check_legacy_hashes.py` | símbolos criptográficos legacy prohibidos en `platform/src/` | nombra el símbolo y el fichero | `quality_report.py` |
| `check_legacy_retirement.py` | gates de retirada CAP-057..063 que invoca la capa UAT | nombra el gate y el artefacto que sigue vivo | `quality_report.py` |
| `check_branch_name.py` | nombre de rama `tipo/nº-slug`, con lista blanca para ramas de plataforma y de flota | imprime el patrón esperado y el nombre recibido | `review-budget` + `quality_report.py` |
| `check_pr_size.py` | presupuesto de revisión: 400 líneas contadas entre la base y la cabeza | imprime el recuento por fichero, de mayor a menor, y el despachador de la excepción | `review-budget` + `quality_report.py` |
| `check_walkthrough_schema.py` | esquema de los `walkthrough-*.json` (`--strict`) | nombra el fichero y el campo obligatorio que falta | `review-budget` |
| `check_workflows.py` | contrato de los workflows: `timeout-minutes`, puertos de servicios, `uses:` pinneado, preflight de Docker, concurrencia, permisos, aislamiento de runner en PRs y **prohibición de `self-hosted`** | `fichero:línea: NOMBRE-CHECK: ubicación: mensaje` | `review-budget` + `release.yml` + `local-preflight.sh` |
| `check_required_jobs.py` | que todos los jobs exigidos por la protección de rama hayan terminado en success | falla cerrado: una lista vacía de jobs es un fallo de la API, no “cero problemas” | `ci.yml` (con `.github/required-jobs-policy.json` como política) |
| `check_mutation.py` | supervivientes de la sesión semanal, con guarda de sesión degenerada | un árbol sin supervivientes evaluables no es verde | `ci.yml` (semanal) + `Makefile` |
| `check_decision_guards.py` | que cada decisión con guarda tenga su guarda cruzada | — (**no cableado**: el script existe y está documentado, pero ningún job lo invoca todavía) | — |
| `check_test_classification.py` | clasificación de cada `test_*.py` por capa (por defecto en modo seco) | — (**no cableado**: igual que el anterior; la clasificación vigente se mantiene a mano en `docs/testing/testing-strategy.md`) | — |

Los dos últimos se declaran aquí a propósito: un script sin gate es deuda visible, no un gate más. Cuando se cableen, la matriz de `tests/test_ci_workflow.py` lo exigirá.

## 3. Contrato de los workflows

| Workflow | Dispara | Jobs | Runner |
|---|---|---|---|
| `ci.yml` | `pull_request`, `push` a `main`, `schedule` (lunes 03:00 UTC) y `workflow_dispatch` | `review-budget`, `quality`, `mutation`, `security`, `codeql`, `required` | `ubuntu-24.04` (`mutation`: `ubuntu-24.04-arm`) |
| `security.yml` | `workflow_call` desde `ci.yml` | `gitleaks`, `trivy-config`, `pip-audit` | `ubuntu-24.04` |
| `codeql.yml` | `workflow_call` desde `ci.yml` | `codeql` | `ubuntu-24.04` |
| `security-deep.yml` | `schedule` (domingo 03:00 UTC) y manual | `gitleaks-history`, `trivy-image` | `ubuntu-24.04-arm` |
| `operating-doc-drift.yml` | `push` a `main` y `schedule` (lunes 04:00 UTC) | `operating-doc-drift` | `ubuntu-24.04` |
| `release.yml` | `push` de un tag `v*` | `preflight`, `e2e`, `publish`, `verify` | `ubuntu-24.04` (`e2e`/`publish`: `-arm`) |

Dos particularidades con motivo escrito en el propio workflow: el `schedule` existe para lo que no cabe en un PR (mutación semanal, escaneo profundo de historial), y `release.yml` no publica sin el `mutation` sobre los **bytes exactos de la etiqueta** (contrato de `XCUT #703`). El aislamiento de runner está medido por `check_workflows.py`: en este repositorio público **ningún** job usa `self-hosted` (`docs/10-runners.md`).

## 4. El `mutation` semanal

- **Qué detecta:** los tests que pasan sin afirmar nada. Mutar el código y ver que la suite sigue verde delata un assert vacío o un test que no ejercita el camino.
- **Cuándo corre:** lunes 03:00 UTC, en un tag de release y bajo demanda (`workflow_dispatch`). Nunca en un PR: cuesta hasta una hora.
- **Por qué no bloquea:** queda fuera del agregador `required` a propósito, así que su rojo no impide un merge; por eso su seguimiento es por issue y no por protección de rama.
- **Estado (2026-10-10):** el job recibía un runner propio del VPS que no existía —cero runners registrados— y llevaba cuatro semanas en rojo sin seguimiento; el pin se corrigió y el gate que lo prohíbe quedó cableado en #834, y su ejecución real depende de que aterrice la instalación de `cosmic-ray` y su configuración (#799).

## 5. Cómo añadir un gate nuevo

1. **El script**, con docstring que declare qué mide, su QC y la decisión que protege, y sus códigos de salida documentados. Datos estructurados, nunca prosa.
2. **La fila en la tabla** de la §2 de este documento (y en `docs/architecture.md` §CI gates, que mapea qué gate protege qué decisión).
3. **La matriz de `tests/test_ci_workflow.py`**: el nombre entra en `REQUIRED_GATE_ORDER` si es un gate de código. Un script nuevo que no esté en la matriz ni en `KNOWN_PENDING_GATES` pone el test en rojo: así un gate no puede quedarse a medias.
4. **El cableado**: el paso en `ci.yml` con su `timeout-minutes`, y en `scripts/quality_report.py` si es un gate de código (ahí vive el orden, no en el YAML).
5. **La prueba de humo** en `tests/test_gate_smoke.py`: que falle ante una violación real y pase en un árbol limpio. Un gate que no se ha visto fallar no es un gate.
