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
| `pr-size` | Se queda, reforzado | Sin verificar | Excepción como campo de datos del cuerpo, no etiqueta (ardelperal/APAP_WEB#1141; R2) |
| `branch-name` | Se queda | Sin verificar | Deriva la issue del nombre de rama; base del trazado determinista (ardelperal/APAP_WEB#1111) |
| `issue-spec` | Se queda, mensaje a mejorar | Sin verificar | Trazabilidad determinista rama + etiquetas + cierre (B4, B11) |
| `integration` (Postgres real) | Se queda | Sin verificar | Verifica comportamiento que el mock no cubre |
| `security` (pip-audit, gitleaks, trivy config) | Se queda, re-agendado | Sin verificar | El resultado de pip-audit es función de la base de avisos del momento, no del diff — siete fallos en cuatro ramas sin relación dentro de 80 minutos (2026-09-29/10-01) —: corre programado sobre la rama por defecto y en los PRs que tocan manifiestos (HR-39, parámetro 22); gitleaks y trivy config sí son función del diff y permanecen por PR |

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
| `actionlint` en `lint` | Falta | — | Cazó `services.minio.command`, clave inexistente en Actions (`ci.yml:1251`) |
| Job de preflight en CI | Falta | — | El preflight existe (#1145); falta el paso que lo reproduce en el job |

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
