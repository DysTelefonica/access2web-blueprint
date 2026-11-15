# Gates de CI: contrato, manual y cómo añadir uno

> **Qué es.** El manual de uso de los gates deterministas del repositorio: el contrato global (qué familias hay y cómo llega un PR al merge), qué mide cada `scripts/check_*.py`, cómo se lee su fallo, el contrato de los cinco workflows, el `mutation` semanal y los pasos para añadir un gate nuevo.

> **Qué no es.** No sustituye a `docs/architecture.md` (qué decisiones del sistema protege cada gate, el patrón hexagonal, los core invariants), ni a `CONTRIBUTING.md` (cómo se trabaja aquí, tamaño de los PRs, convenciones), ni a `.github/dependabot.yml` (política de dependencias). Este documento es **el uso de los gates**.

## 1. El contrato global

Los gates se agrupan en cuatro familias, y cada una corre en un sitio distinto:

| Familia | Qué protege | Dónde corre |
|---|---|---|
| Forma del cambio | tamaño del PR (≤400 líneas), `Closes #N` y `issue-spec`, walkthroughs, dependencias entre cadenas de PRs, carriles de documentación y de superficie | `pr-gates.yml::pr-gate` (gate H1, tope 10 min, `pr-gates.yml:71`) y `pr-gate.yml::pr-gate` (tope 5 min, `pr-gate.yml:112`), llamado desde `ci.yml:93` |
| Calidad y pruebas | formato, lint, tipos, tests, complejidad, DRY, sitios de mutación estáticos, hashes legacy, retirada legacy, suites pobladas | `ci.yml::quality` (tope 30 min, `ci.yml:129`); `ci.yml::integration` (tope 25 min, `ci.yml:380`) y `ci.yml::e2e-smoke` (`ci.yml:894`) |
| Seguridad, gobierno y documentación | secretos, configuración de contenedores, dependencias, CodeQL, sellos de gobierno, imágenes pinneadas por digest, enlaces, comandos y decisiones documentadas | `security.yml`, `pip-audit.yml` y `codeql.yml` (reusable, llamados desde `ci.yml`); `governance-seals` (`ci.yml:652`), `docs-*` (`ci.yml:608` a `:835`) y `ci-images` (`ci.yml:1034`) |
| Contrato del host y del arnés | protección de rama, jobs requeridos, paridad de la política, matriz de reglas | agregador `ci.yml::required` (`ci.yml:1080`) y fase 4 del contrato de adopción: `ci-pattern adoption check --phase 4 .` |

El recorrido de un PR hasta el merge:

```text
PR  (rama `tipo/nº-slug`, cuerpo con `Closes #N` y `## Chain Context`, ≤400 líneas o `size:exception`)
 │  push a la rama
 ▼
GitHub Actions
 ├── pr-gates.yml::pr-gate   pull_request_target: tamaño e issue-spec sobre el catálogo de la base
 ├── ci.yml::pr-gate ──────► pr-gate.yml   workflow_call: walkthroughs, carriles, dependencias
 ├── ci.yml::quality, integration, e2e-smoke, docs-*, governance-seals, ci-images
 ├── ci.yml::security, pip-audit, codeql   ──► security.yml, pip-audit.yml, codeql.yml
 │
 └── ci.yml::required   (agregador; needs: los jobs de ci.yml; if por evento)
              │
              ▼
     .github/host-contract.json → required_checks: `required` y `pr-gate`
              │
              ▼
     merge del mantenedor (ruleset `main-maintainers-and-admins-merge`)
```

Cinco reglas que el diagrama no muestra y que conviene saber:

- **`pr-gate` está partido en dos workflows a propósito.** `pr-gates.yml` corre en `pull_request_target` y ejecuta el catálogo que publica la base, para que la cabeza del PR no pueda alterar el gate que lo mide (`pr-gates.yml:1-13`). `pr-gate.yml` corre en `pull_request` y en `workflow_call`; lo que lee es el árbol y el diff, sin token (`pr-gate.yml:1-16`, `:63-81`).
- **`required` es el agregador.** Su `needs:` (`ci.yml:1115`) reúne los jobs de `ci.yml`, y `pr-gate` y `required` son los dos `required_checks` que declara el contrato del host. Un job que no queremos bloqueante no se añade a `needs`.
- **`required` no usa `always()`.** Su `if:` (`ci.yml:1110`) fija el evento a `pull_request` o `push` y añade `!cancelled()`: un job saltado como dependencia no salta este check, y un check requerido saltado no bloquea el merge. El comentario de `ci.yml:1080-1110` explica por qué no es `always()` literal (HR-30, #1464).
- **`security.yml`, `pip-audit.yml`, `codeql.yml` y `pr-gate.yml` son *reusable workflows*** (`workflow_call`): `needs:` no cruza ficheros de workflow, así que `ci.yml` los invoca con `uses:` (`ci.yml:93`, `:520`, `:540`, `:592`).
- **Cada job declara `timeout-minutes`**: un gate que se cuelga es un gate que no mide, y el tope convierte el cuelgue en un fallo con diagnóstico (lo verifica `check_workflows.py`, §2).

## 2. Los `check_*.py`, uno por uno

Catorce scripts viven en `scripts/`; su contrato es siempre el mismo: leen datos estructurados (rama, etiquetas, campos de la API, YAML, AST), nunca prosa ni gestos, y publican un veredicto con código de salida.


`check_complexity.py`, `check_mutation_sites.py`, `check_mutation.py` y `check_dry.py` son copias del asset del canon `deterministic-quality-harness` 2.2. Los cuatro toman de `quality-policy.json` los sujetos y los techos (BP no sube el techo de complejidad: 10, con fecha de revisión en el propio script) y evalúan su `BASELINE` contra una fecha explícita (`--policy-date`), no contra el reloj. `quality_report.py` propaga esa fecha a los gates que la necesitan (`DATED_GATES`).


Cada copia declara en su cabecera las desviaciones locales que conserva respecto del asset, para que la diferencia sea una decisión y no una deriva silenciosa:

- `check_complexity.py`: solo el techo 10, que es política de BP.
- `check_mutation_sites.py`: ninguna de mecanismo. La exclusión de docstrings la publica el canon desde team-skills#1008, así que su única desviación es el `BASELINE` como dato.
- `check_mutation.py`: `--shards` (una sesión por tramo de la matriz), el `BASELINE` agregado por módulo en lugar de por mutante (readquirirlo por mutante exige la corrida completa en el runner ARM). `--policy-date` dejó de ser opcional en BP #1324: `release.yml::mutation` y `ci.yml::quality` pasan la fecha, y el BASELINE se juzga contra ella, no contra el reloj.
- `check_dry.py`: ninguna de mecanismo —umbrales (`MIN_STATEMENTS = 5`, `MIN_OCCURRENCES = 2`), escaneo por la política, guarda de liveness y sobre ligado son el asset—; su `BASELINE` (cuatro grupos de duplicados tolerados hasta su fecha objetivo) es dato de BP. Un fichero que el gate no puede parsear aparece en `subjects.unclassified` en vez de desaparecer del recuento.

Dos copias no coinciden byte a byte con el asset aunque el mecanismo sea el mismo: el canon no exige `ruff format` y el job `quality` de BP sí, así que el reparto de líneas y algún comentario de contexto son de esta casa. La diferencia que importa —qué mide y con qué política— es la de estas listas.

`quality_report.py` es el agregador y también es copia del asset 2.2: valida el sobre de cada gate (`gate`, `status`, `indicators`, `ceilings`, `findings`, y la ligadura a la política en los gates que la gobiernan), lo vuelve a contrastar con la evidencia que recalcula del árbol, y se niega a copiar al informe una evidencia que ya no cuadra. Conserva cuatro desviaciones declaradas en su cabecera: su propia lista de gates (los dos caminantes legacy no están en la política y el asset los filtraría), `DATED_GATES` (la fecha solo va a los gates que la aceptan, y desde BP #1324 la exige), la ausencia de `--include-pr-gates`, que BP retiró en #854 (los gates de revisión viven en su propio job), y su tabla `INDICATOR_MEANINGS`: el asset no tabula `files_measured` ni `statements_measured` pese a que sus propios gates los emiten, y su valor por defecto los daría por «menor es mejor» cuando son cobertura de medición, así que BP los declara como «mayor es mejor»; las entradas del asset correspondientes a los dos gates que BP no ejecuta están igualmente ausentes, porque esos indicadores no podrían aparecer en este informe y porque DG-12 no permite nombrarlos aquí.

| Script | Qué mide | Cómo se lee su fallo | Dónde corre |
|---|---|---|---|
| `docs-architecture` (job de `ci.yml`) | dirección de dependencias entre capas hexagonales, independencia entre aplicaciones, pureza de capa y ficheros fuera de las capas, desde el bloque `architecture` de `docs/architecture.md` | nombra `fichero:línea` y la regla incumplida; un fichero que no pertenece a ninguna capa ni a `exempt` es fallo, no verde | `ci.yml` |
| `check_complexity.py` | complejidad ciclomática por función, con techo global absoluto | indica la función y su valor frente al techo | `quality_report.py` |
| `check_mutation_sites.py` | densidad estática de sitios mutables, con techo y sujetos desde `quality-policy.json` | indica los ficheros con densidad fuera de rango | `ci.yml` + `quality_report.py` |
| `check_dry.py` | bloques duplicados tipo 1 y 2 sobre el AST, con techo y sujetos desde `quality-policy.json` | nombra los pares de bloques y su ubicación | `ci.yml` + `quality_report.py` |
| `check_legacy_hashes.py` | símbolos criptográficos legacy prohibidos en `platform/src/` | nombra el símbolo y el fichero | `quality_report.py` |
| `check_legacy_retirement.py` | gates de retirada CAP-057..063 que invoca la capa UAT | nombra el gate y el artefacto que sigue vivo | `quality_report.py` |
| `check_suites_populated.py` | que cada suite obligatoria (`unit`, `integration`, `e2e_smoke`) recoja al menos un test | nombra la suite vacía; falla cerrado si `pytest` no recolecta | `ci.yml::quality` (step `suites populated`) |
| `check_branch_name.py` (canon, `assets/branch-name/`) | nombre de rama del patrón gentle-ai (`<tipo>/<slug>`, sin número de issue) | imprime el patrón esperado y el nombre recibido | — (**no cableado**: la copia local de BP se retiró en #859; la regla la impone el ruleset del host) |
| `check_pr_size.py` (canon, `assets/pr-size/`) | presupuesto de revisión: 400 líneas contadas entre la base y la cabeza | imprime el recuento por fichero, de mayor a menor, y el despachador de la excepción | `pr-gates.yml::pr-gate` (gate H1) |
| `check_walkthrough_schema.py` | esquema de los `walkthrough-*.json` (`--strict`) | nombra el fichero y el campo obligatorio que falta | `pr-gate.yml::pr-gate` |
| `check_change_scope.py` | superficie que analiza cada job pesado: `e2e` (`app/`, `e2e/`), `codeql` (`app/`, `scripts/`) y `lockfile` (`app/requirements-dev.lock`, `app/uv.lock`) | no falla: imprime `true` o `false` por superficie | `pr-gate.yml::pr-gate` (publica `e2e_scope_changed`, `codeql_scope_changed` y `lockfile_scope_changed`) |
| `check_docs_only_lane.py` | carril docs-only: `true` solo si el diff es no vacío y cae entero bajo `docs/` | no falla: imprime `true` o `false` | `pr-gate.yml::pr-gate` (publica el marcador `docs_changed`) |
| `check_workflows.py` | contrato de los workflows: `timeout-minutes`, puertos de servicios, `uses:` pinneado, preflight de Docker, concurrencia, permisos, aislamiento de runner en PRs y **prohibición de `self-hosted`** | `fichero:línea: NOMBRE-CHECK: ubicación: mensaje` | `ci.yml::quality` (step `workflow gate`) + `Makefile` (`make check-workflows`) |
| `check_required_jobs.py` (canon, `.agents/skills/ci-pattern/assets/required-jobs/check_required_jobs.py`) | que todos los jobs exigidos por la protección de rama hayan terminado en success, con los saltos aceptados por evento que declara la política, y que la política no se relaje sin `relaxation` | falla cerrado: una lista vacía de jobs es un fallo de la API, no “cero problemas”; una clave de `needs` que la política no cubre o un salto no declarado también fallan | `ci.yml` (con `.github/required-jobs-policy.json` como política) |
| `check_mutation.py` | supervivientes de las sesiones de la corrida semanal (una por tramo, `--shards`), con guarda de sesión degenerada y `BASELINE` agregado por módulo | un árbol sin supervivientes evaluables no es verde; una sesión con sujetos fuera de `quality-policy.json` tampoco | `release.yml::mutation` + `Makefile` |
| `check_decision_guards.py` | que cada decisión con guarda tenga su guarda cruzada | — (**no cableado**: el script existe y está documentado, pero ningún job lo invoca todavía) | — |
| `check_test_classification.py` | clasificación de cada `test_*.py` por **directorio de capa** (por defecto en modo seco); es política local de BP, declarada en la cabecera del propio script, no el asset del canon | — (**no cableado**: igual que el anterior; la clasificación vigente se mantiene a mano en `docs/testing/testing-strategy.md`) | — |
| `check_deadcode.py` | ratchet de código muerto: `vulture app/ --min-confidence 80` frente a `.deadcode-baseline.txt`, y solo se estrecha | nombra los hallazgos nuevos y las entradas caducas del baseline | — (**no cableado**: figura en `KNOWN_PENDING_GATES` de `tests/test_ci_workflow.py`) |

Los tres scripts de `scripts/` que ningún job invoca (`check_decision_guards.py`, `check_test_classification.py` y `check_deadcode.py`) se declaran aquí a propósito: un script sin gate es deuda visible, no un gate más. Cuando se cableen, la matriz de `tests/test_ci_workflow.py` lo exigirá.

## 3. Contrato de los workflows

| Workflow | Dispara | Jobs | Runner |
|---|---|---|---|
| `ci.yml` | `pull_request` (`opened`, `synchronize`, `reopened`), `push` a `main`, `schedule` diario 03:00 UTC (`0 3 * * *`, `ci.yml:68`) y `workflow_dispatch` | `pr-gate`, `quality`, `integration`, `security`, `pip-audit`, `pip-audit-alert`, `codeql`, `docs-architecture`, `governance-seals`, `docs-codeguide`, `docs-links`, `docs-commands`, `docs-decisions`, `e2e-smoke`, `ci-images`, `required` | `ubuntu-24.04` |
| `pr-gate.yml` | `pull_request` (seis tipos, incluidos `edited`, `labeled` y `unlabeled`) y `workflow_call` desde `ci.yml` | `pr-gate` | `ubuntu-24.04` |
| `pr-gates.yml` | `pull_request_target` (`opened`, `edited`, `labeled`, `unlabeled`, `synchronize`, `reopened`) | `canon`, `pr-gate` | `ubuntu-24.04` |
| `security.yml` | `workflow_call` desde `ci.yml` | `gitleaks`, `trivy-config` | `ubuntu-24.04` |
| `pip-audit.yml` | `workflow_call` desde `ci.yml` | `dependency-scan` | `ubuntu-24.04` |
| `codeql.yml` | `workflow_call` desde `ci.yml` | `analyze` | `ubuntu-24.04` |
| `security-deep.yml` | `schedule` (domingo 03:00 UTC, `0 3 * * 0`) y manual | `gitleaks-history`, `trivy-image` | `ubuntu-24.04-arm` |
| `operating-doc-drift.yml` | `schedule` (lunes 04:00 UTC, `0 4 * * 1`) y manual | `operating-doc-drift` | `ubuntu-24.04` |
| `release.yml` | `push` de un tag `v*` | `preflight`, `mutation`, `stage`, `e2e`, `publish`, `verify` | `ubuntu-24.04`; `mutation`, `stage`, `e2e` y `publish`: `ubuntu-24.04-arm` |
| `dysflow-erd.yml` | `workflow_dispatch` | `extract` | runner `self-hosted` Windows del propietario (`dysflow-erd.yml:29`) |

Particularidades con motivo escrito en el propio workflow:

- **`ci.yml` corre una vez por PR y una vez por push a `main`.** El `schedule` nocturno es la red de seguridad sobre la base; no hay cola de merge (`ci.yml:55-68`).
- **La mutación ya no corre en `ci.yml`.** Se retiró del CI de PR (#881) y vive en `release.yml::mutation`; `publish` depende de ella, así que un tag no se publica sin ese gate sobre sus bytes exactos (`release.yml:72`, `:82`).
- **El único job `self-hosted` es `dysflow-erd.yml`**, y solo por `workflow_dispatch`: `check_workflows.py` lo admite únicamente para ese tipo de fichero (`dysflow-erd.yml:1-5`). El resto de runners son `ubuntu-24.04` y `ubuntu-24.04-arm` hospedados por GitHub.

El aislamiento de runner lo mide `check_workflows.py` (§2), que prohíbe `self-hosted` en cualquier otro job.

## 4. El `mutation` en el release

- **Qué detecta:** los tests que pasan sin afirmar nada. Mutar el código y ver que la suite sigue verde delata un assert vacío o un test que no ejercita el camino.
- **Cuándo corre:** solo en un tag `v*` (`release.yml:7-9`). `release.yml` no tiene `schedule` ni `workflow_dispatch`, y no corre en PR ni en `ci.yml`, que lo retiró en #881 (`release.yml:82`). Nunca en un PR: su tope es de 150 minutos (`release.yml:96`).
- **Dónde corre:** el job `mutation` es `ubuntu-24.04-arm` (`release.yml:90`). Reparte cosmic-ray en 10 shards (`release.yml:150-163`) y el ratchet `check_mutation.py` corre sobre las sesiones fusionadas (`release.yml:166-174`).
- **Por qué bloquea un release y no un merge:** `publish` depende de `mutation` (`release.yml:350`), así que un tag no se publica sin este gate sobre sus bytes exactos. Queda fuera del agregador `required` (`ci.yml:1115`), que es el de PR; su rojo no impide un merge.
- **Historia:** el 2026-10-10 el job exigía un runner propio del VPS sin runners registrados y llevó cuatro semanas en rojo sin seguimiento. El pin se corrigió y el gate que lo prohíbe quedó cableado en #834.

## 5. Cómo añadir un gate nuevo

1. **El script**, con docstring que declare qué mide, su QC y la decisión que protege, y sus códigos de salida documentados. Datos estructurados, nunca prosa.
2. **La fila en la tabla** de la §2 de este documento (y en `docs/architecture.md` §CI gates, que mapea qué gate protege qué decisión).
3. **La matriz de `tests/test_ci_workflow.py`**: el nombre entra en `REQUIRED_GATE_ORDER` si es un gate de código. Un script nuevo que no esté en la matriz ni en `KNOWN_PENDING_GATES` pone el test en rojo: así un gate no puede quedarse a medias.
4. **El cableado**: el paso en `ci.yml` con su `timeout-minutes`, y en `scripts/quality_report.py` si es un gate de código (ahí vive el orden, no en el YAML).
5. **La prueba de humo** en `tests/test_gate_smoke.py`: que falle ante una violación real y pase en un árbol limpio. Un gate que no se ha visto fallar no es un gate.
