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


`check_complexity.py`, `check_mutation_sites.py`, `check_mutation.py` y `check_dry.py` son copias del asset del canon `deterministic-quality-harness` 2.2. Los cuatro toman de `quality-policy.json` los sujetos y los techos (BP no sube el techo de complejidad: 10, con fecha de revisión en el propio script) y evalúan su `BASELINE` contra una fecha explícita (`--policy-date`), no contra el reloj. `quality_report.py` propaga esa fecha a los gates que la necesitan (`DATED_GATES`).

Cada copia declara en su cabecera las desviaciones locales que conserva respecto del asset, para que la diferencia sea una decisión y no una deriva silenciosa. Hoy: el techo de complejidad 10 (`check_complexity.py`); la exclusión de docstrings en el contador de sitios (`check_mutation_sites.py`, pendiente de team-skills#999); en `check_mutation.py` el `--shards` de una sesión por tramo de la matriz, el `BASELINE` agregado por módulo en lugar de por mutante (readquirirlo por mutante exige la corrida completa en el runner ARM) y un `--policy-date` **opcional** que cae a la fecha UTC de hoy, porque el canon lo exige y la corrida semanal de `release.yml` todavía no puede pasarlo; y en `check_dry.py` ninguna desviación de mecanismo, solo sus datos: umbrales (`MIN_STATEMENTS = 5`, `MIN_OCCURRENCES = 2`), escaneo por la política, guarda de liveness y sobre ligado son el asset, y su `BASELINE` (cuatro grupos de duplicados tolerados hasta su fecha objetivo) es dato de BP. Un fichero que `check_dry.py` no puede parsear aparece en `subjects.unclassified` en vez de desaparecer del recuento. `check_dry.py` no conserva ninguna desviación de mecanismo, solo sus datos: umbrales (`MIN_STATEMENTS = 5`, `MIN_OCCURRENCES = 2`), escaneo por la política, guarda de liveness y sobre ligado son el asset. Su `BASELINE` (cuatro grupos de duplicados tolerados hasta su fecha objetivo) es dato de BP y así lo declara su cabecera. Un fichero que el gate no puede parsear aparece en `subjects.unclassified` en vez de desaparecer del recuento.


- `check_complexity.py`: solo el techo 10, que es política de BP.
- `check_mutation_sites.py`: ninguna de mecanismo. La exclusión de docstrings la publica el canon desde team-skills#1008, así que su única desviación es el `BASELINE` como dato.
- `check_mutation.py`: `--shards` (una sesión por tramo de la matriz), el `BASELINE` agregado por módulo en lugar de por mutante (readquirirlo por mutante exige la corrida completa en el runner ARM) y un `--policy-date` **opcional** que cae a la fecha UTC de hoy, porque el canon lo exige y la corrida semanal de `release.yml` todavía no puede pasarlo (se retira en BP #1324).
- `check_dry.py`: ninguna de mecanismo —umbrales (`MIN_STATEMENTS = 5`, `MIN_OCCURRENCES = 2`), escaneo por la política, guarda de liveness y sobre ligado son el asset—; su `BASELINE` (cuatro grupos de duplicados tolerados hasta su fecha objetivo) es dato de BP. Un fichero que el gate no puede parsear aparece en `subjects.unclassified` en vez de desaparecer del recuento.

Dos copias no coinciden byte a byte con el asset aunque el mecanismo sea el mismo: el canon no exige `ruff format` y el job `quality` de BP sí, así que el reparto de líneas y algún comentario de contexto son de esta casa. La diferencia que importa —qué mide y con qué política— es la de estas listas.

`quality_report.py` es el agregador y también es copia del asset 2.2: valida el sobre de cada gate (`gate`, `status`, `indicators`, `ceilings`, `findings`, y la ligadura a la política en los gates que la gobiernan), lo vuelve a contrastar con la evidencia que recalcula del árbol, y se niega a copiar al informe una evidencia que ya no cuadra. Conserva tres desviaciones declaradas en su cabecera: su propia lista de gates (los dos caminantes legacy no están en la política y el asset los filtraría), `DATED_GATES` (la fecha solo va a los gates que la aceptan) y `--policy-date` opcional. No tiene `--include-pr-gates`: BP los retiró en #854 y los gates de revisión viven en su propio job.

| Script | Qué mide | Cómo se lee su fallo | Dónde corre |
|---|---|---|---|
| `docs-architecture` (job de `ci.yml`) | dirección de dependencias entre capas hexagonales, independencia entre aplicaciones, pureza de capa y ficheros fuera de las capas, desde el bloque `architecture` de `docs/architecture.md` | nombra `fichero:línea` y la regla incumplida; un fichero que no pertenece a ninguna capa ni a `exempt` es fallo, no verde | `ci.yml` |
| `check_complexity.py` | complejidad ciclomática por función, con techo global absoluto | indica la función y su valor frente al techo | `quality_report.py` |
| `check_mutation_sites.py` | densidad estática de sitios mutables, con techo y sujetos desde `quality-policy.json` | indica los ficheros con densidad fuera de rango | `ci.yml` + `quality_report.py` |
| `check_dry.py` | bloques duplicados tipo 1 y 2 sobre el AST, con techo y sujetos desde `quality-policy.json` | nombra los pares de bloques y su ubicación | `ci.yml` + `quality_report.py` |
| `check_legacy_hashes.py` | símbolos criptográficos legacy prohibidos en `platform/src/` | nombra el símbolo y el fichero | `quality_report.py` |
| `check_legacy_retirement.py` | gates de retirada CAP-057..063 que invoca la capa UAT | nombra el gate y el artefacto que sigue vivo | `quality_report.py` |
| `check_branch_name.py` | nombre de rama `tipo/nº-slug`, con lista blanca para ramas de plataforma y de flota | imprime el patrón esperado y el nombre recibido | `review-budget` + `quality_report.py` |
| `check_pr_size.py` (canon, `assets/pr-size/`) | presupuesto de revisión: 400 líneas contadas entre la base y la cabeza | imprime el recuento por fichero, de mayor a menor, y el despachador de la excepción | `review-budget` + `quality_report.py` |
| `check_walkthrough_schema.py` | esquema de los `walkthrough-*.json` (`--strict`) | nombra el fichero y el campo obligatorio que falta | `review-budget` |
| `check_workflows.py` | contrato de los workflows: `timeout-minutes`, puertos de servicios, `uses:` pinneado, preflight de Docker, concurrencia, permisos, aislamiento de runner en PRs y **prohibición de `self-hosted`** | `fichero:línea: NOMBRE-CHECK: ubicación: mensaje` | `review-budget` + `release.yml` + `local-preflight.sh` |
| `check_required_jobs.py` (canon, `.agents/skills/ci-pattern/assets/required-jobs/check_required_jobs.py`) | que todos los jobs exigidos por la protección de rama hayan terminado en success, con los saltos aceptados por evento que declara la política, y que la política no se relaje sin `relaxation` | falla cerrado: una lista vacía de jobs es un fallo de la API, no “cero problemas”; una clave de `needs` que la política no cubre o un salto no declarado también fallan | `ci.yml` (con `.github/required-jobs-policy.json` como política) |
| `check_mutation.py` | supervivientes de las sesiones de la corrida semanal (una por tramo, `--shards`), con guarda de sesión degenerada y `BASELINE` agregado por módulo | un árbol sin supervivientes evaluables no es verde; una sesión con sujetos fuera de `quality-policy.json` tampoco | `release.yml::mutation` + `Makefile` |
| `check_decision_guards.py` | que cada decisión con guarda tenga su guarda cruzada | — (**no cableado**: el script existe y está documentado, pero ningún job lo invoca todavía) | — |
| `check_test_classification.py` | clasificación de cada `test_*.py` por **directorio de capa** (por defecto en modo seco); es política local de BP, declarada en la cabecera del propio script, no el asset del canon | — (**no cableado**: igual que el anterior; la clasificación vigente se mantiene a mano en `docs/testing/testing-strategy.md`) | — |

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
