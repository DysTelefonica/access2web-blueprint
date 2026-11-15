# Veredicto de gates del inventario de CI

Modelo de salida del paso de auditoría (§4 de la skill): cada gate del
inventario recibe un veredicto con la evidencia que lo sostiene. Veredictos
posibles: **se queda**, **se refuerza**, **se duerme** (HR-18, R15), **se
retira** (solo con medición que pruebe que no hay candado que conservar).
Desde HR-32, cada control registra además **el test que lo ve fallar**: la
prueba de que ejecuta su lógica contra un fixture que viola la regla y observa
su veredicto. Un control sin fixture violador no se declara como control: se
clasifica `documented-only` hasta tenerlo.

Estado: auditoría inicial de `ardelperal/APAP_WEB` (2026-09-30); lo no
verificado se marca. Las fricciones citadas se detallan en
`references/fricciones.md`.

| Gate | Veredicto | Test que lo ve fallar | Base |
|---|---|---|---|
| `required` (agregador) | Se queda, reforzado | Suite de paridad del asset (`assets/required-jobs/`) | Falla cerrado con matriz de skips esperados; causa raíz separada de skips en cascada (F-003, ardelperal/APAP_WEB#1118) |
| `pr-size` | Se retira de CI (#1396) | — | El presupuesto lo mide solo A1 de `pr-contract`; ver «Medición en cadete». Antes: excepción como campo de datos del cuerpo, no etiqueta (ardelperal/APAP_WEB#1141; R2) |
| `branch-name` | Aviso, no bloquea (R10 de team-skills#1093) | — | El ruleset del repo aplica `^(feat\|fix\|chore\|docs\|refactor\|test)/[a-z0-9._-]+$` cuando existe. En un host sin rulesets el gate mide y avisa con `AVISO`, sin fallar el PR. |
| `issue-spec` | Se queda, mensaje a mejorar | Sin verificar | Trazabilidad determinista rama + etiquetas + cierre (B4, B11) |
| `ci-images` | Se queda, opt-in | `assets/ci-images/tests/test_ci_images_pinned.py` | Sin `P56_ci_images_gate: on` el gate es inerte; con él mide las imágenes de los workflows del consumer y las excepciones sin motivo o obsoletas son hallazgos (#1264) |
| `workflow-policy` (HR-79) | Se queda, reforzado | `assets/workflow-policy/tests/test_workflow_policy.py` | `needs.<job>.outputs` solo existe para una dependencia DIRECTA: leerla de una transitiva llega vacía y el workflow queda verde con el valor equivocado; el marcador `if:` de #1192 es la excepción (#1277) |
| `integration` (Postgres real) | Se queda | Sin verificar | Verifica comportamiento que el mock no cubre |
| `security` (pip-audit, gitleaks, trivy config) | Se queda, re-agendado | Sin verificar | El resultado de pip-audit es función de la base de avisos del momento, no del diff — siete fallos en cuatro ramas sin relación dentro de 80 minutos (2026-09-29/10-01) —: corre solo en el cron nocturno sobre la rama por defecto (job `dependency-scan` del harness, #1359); no es requerido por PR, tampoco cuando el PR toca manifiestos (D11 lo dejaba por PR con manifiestos; #1359 lo retira); gitleaks y trivy config sí son función del diff y permanecen por PR |

## Registro de identificadores de evidencia

Las citas `(evidencia: R<N>)` de §2 resuelven aquí (auditoría issue→merge de
`ardelperal/APAP_WEB`, épica ardelperal/APAP_WEB#935, 2026-09-29/30):

| ID | Lección que sostiene la regla |
|---|---|
| R2 | Las excepciones de presupuesto se declaran como campo de datos del cuerpo del PR, nunca como etiqueta aplicada tarde. |
| R3 | Todo gate nuevo que toque producción se ejecuta una vez real contra el entorno; ninguna prueba de escritorio lo sustituye. |
| R4 | Un gate que no midió no imprime «OK»: falla en voz alta o no existe. |
| R5 | El job que gatea el merge tiene paridad local: todo lo que corre en CI corre en local con un comando. |
| R6 | La evidencia de deploy se registra sobre el SHA de la revisión desplegada, nunca sobre una variable global ni la punta de la rama por defecto. |
| R9 | Los registros acumulativos de varias sesiones se guardan sin clave de upsert: una observación nueva por entrada. |
| R10 | El mensaje de un fallo describe la causa real y, si el dato no es corregible por el autor, ofrece la vía manual auditable. |
| R12 | Tras corregir el primer paso rojo de un job, se reproducen también todos los pasos posteriores. |
| R14 | Los procesos vivos sondeando CI se sustituyen por mecanismos (auto-merge, workflows programados) y las fricciones se registran con evidencia para automatizar la segunda ocurrencia. |
| R15 | Los gates sin defecto real cazado se duermen tras policy file con transición validada; no se retira el candado construido. |

| `lint` + `ruff`/`mypy` | Se queda | Sin verificar | Con paridad de preflight (ardelperal/APAP_WEB#1145) y actionlint pendiente de añadir |
| Cobertura | Se refuerza | Sin verificar | Solo unit con dobles; 85 % sobre `app` y `migration` |
| e2e de PR | Se refuerza | Sin verificar | Seis tests, ninguno envía formularios aún |
| `mutation` / `security-deep` | Se refuerza (dormible) | Sin verificar | Solo en tag o cron; candidatos al policy file de HR-18 |
| `check_vulture_guard` | Se queda (arreglado) | Sin verificar | Fail-loud desde ardelperal/APAP_WEB#1143 (A9 cerrado); baseline shrink-only (HR-15) |
| `check_mutation_sites` | Se duerme (R15) | Sin verificar | BASELINE 464 y subiendo, sin defecto real cazado; dormido en la era #1167 |
| `check_crap` | Se duerme (R15) | Sin verificar | Informativo sin defecto real cazado; dormido en la era #1167 |
| `check_alantyle` | Se duerme (R15) | Sin verificar | Bloqueante de estilo de docs; falló por una palabra en mayúsculas (ardelperal/APAP_WEB#1133); informativo #1151 |
| Docstrings (balance/cobertura) | Se duerme (R15) | Sin verificar | Informativos, sin defecto real cazado |
| `test_ci_workflow.py` | Se simplifica | Subcadenas del fuente: no cuenta como test (HR-32) | El fichero más editado del repo por cambios de workflow |
| `actionlint` en `lint` | Falta | — | Cazó `services.<almacenamiento>.command`, clave inexistente en Actions (`ci.yml:1251`, caso de APAP_WEB) |
| Job de preflight en CI | Falta | — | El preflight existe (#1145); falta el paso que lo reproduce en el job |

## Medición en cadete: gates de PR web (#1396)

Ventana: `DysTelefonica/cadete`, 2026-10-25 → 2026-10-28, 399 ejecuciones y
47 cabezas de PR (PRs #742–#767 y #604–#609), leídas de la API de Actions.
Se cuenta cada rojo por causa raíz. Los dos únicos defectos reales del periodo
los cazaron las suites: `ext-gd` ausente para composer en CI y la sonda de
arranque de Playwright rota por el login solo JWT (cadete #754).

Coste: los gates de gobierno y lint suman 569 jobs y 387 minutos de runner;
PHPUnit y Playwright, 74 jobs y 140 minutos. Mediana hasta los 8 contextos
requeridos en verde: 10,9 min (p90 28,1), casi todo cola de runner.

| Gate | Rojos y causa | Defectos cazados | Veredicto |
|---|---|---|---|
| PHPUnit, Playwright | 2 causas reales | 2 | Se quedan |
| `pr-size` | 5, todos por falta de `size:exception`, siempre junto a `pr-contract` A1 y con otra cifra (156191 frente a 168630) | 0 | Se retira de CI: A1 de `pr-contract` es el único control de tamaño, con una sola cifra; `check_pr_size.py` queda para uso local |
| `pr-contract` A1 (presupuesto de 400 líneas) | Los mismos 5 | 0 | Se queda bloqueante por decisión del mantenedor (2026-10-28): protege la carga del revisor, que es la regla del equipo, no un detector de defectos. El coste aceptado es la etiqueta |
| `canon` (artefacto del patrón) | 74 fallos en 7 PRs en unas 11 h: la cuota de artefactos de la organización (cadete #752) dejó sin medir todos los contextos requeridos | 0 | Se retira: cada gate toma el patrón de la copia vendorizada de la base, sin token ni artefacto |
| Check-run espejo (`x / x`) | Duplica cada rojo de `pr-gates` en la lista de checks | — | Se queda, documentado: un job de `pull_request_target` cuelga su check-run nativo de la base, y el merge gobernado solo ve el espejo publicado sobre la cabeza |
| `docs-gates` (fase 5) | HR-63/68 ×4, HR-55 ×2; la fase 5 repetía las fases 2 y 3, así que una causa daba dos rojos | 0 | Se queda midiendo solo la fase 5 (`adoption check --only-phase 5`); el nightly sigue con `--phase 5` |
| `governance-seals` (fase 3) | Resellados manuales en #751, #754 y #758; `test_a_pr_cannot_reduce_the_sealing_of_the_base` rojo en PRs ajenos (#756 ×2, #760) | 0 | Se queda por decisión del mantenedor (2026-10-29), justificado por un riesgo de seguridad reproducido y no por un incidente evitado: sin el sello, un PR puede reescribir los ficheros de gobierno que lo miden (hallazgo H1 de la auditoría externa, #1261; escenario E8b reproducido en #1317, un PR que reduce su propia base de sellado). Su coste en CI es pequeño (unos 24 s en un job paralelo en cadete); el coste es el resellado manual, que se reduce automatizándolo, no retirando el gate |
| `env-isolation` | 1, un test que lee `PATH` (#760) | 0 | Se queda; sin defecto cazado en la ventana |
| `issue-spec` | 1 falso positivo: el cuerpo decía «Supersedes the old PRs #133» (#763) | 0 | Se queda; sin defecto cazado en la ventana |
| `branch-name`, `PHP Lint`, `ci-images`, `docs-commands` | Ninguno en 89, 89, 39 y 39 ejecuciones | 0 | Se quedan; sin rojo en la ventana, candidatos a dormirse (R15) si siguen sin cazar nada |
| Frescura del merge gobernado | 12 reruns completos forzados en el tren de la mañana del 2026-10-28; sus únicos rojos fueron sus propios rechazos (#757, #742) | 0 | Se apaga por defecto (`merge_freshness`, #1397); `required` solo donde el deploy compara el árbol fusionado con el revisado |

## Nota sobre R15 (dormir, no retirar)

La corrección de la épica: un gate pasado a informativo sin policy file pierde
el candado construido. El patrón dormant conserva el motor probado y hace de la
re-activación un cambio de datos con review. Regla R15 destilada en la
auditoría de origen (la fuente de trabajo de la skill no se versiona en el
catálogo); forma del policy file en
`assets/parameters.md`; origen del patrón en
`references/benchmark-gentle-ai.md` (idea T1).

## Criterio de decisión

Un gate se retira solo cuando una medición prueba que no cazó ningún defecto
real **y** no conserva candado ninguno que valga la pena dormir. Si duda,
duerma el gate (HR-18): el coste de un gate dormido es una línea de policy; el
coste de reconstruir un gate retirado es un PR entero.
