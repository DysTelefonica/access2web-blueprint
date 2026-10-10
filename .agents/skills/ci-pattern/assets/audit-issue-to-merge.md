# Auditoría issue → merge (procedimiento reproducible del patrón web)

Procedimiento específico del patrón web para auditar el ciclo completo
issue → merge de un consumer que ha adoptado `ci-pattern`. Complementa la
auditoría genérica de `repository-delivery-governance` (su §4 «Map the
lifecycle» y su HR-4 poseen el recorrido genérico y el vocabulario de las
cinco capas de enforcement): este asset aporta las consultas concretas, la
plantilla de matriz y las preguntas propias del patrón, y remite a aquella
skill sin copiarla.

Ejecución: solo lecturas (gh API y ficheros del consumer), sin red más allá
de la API de GitHub, sin escrituras. Ejecutado sobre un consumer
`primary_type: web` de `fleet/registry.json`, produce la matriz de
enforcement y las métricas sin ningún paso que no esté escrito aquí.

## Bloque 1 — Recorrido ordenado (issue → evidencia)

Recorra el ciclo de vida en este orden, registrando para cada eslabón su
evidencia (número de issue, SHA de commit, URL de PR o de ejecución):

1. Issue de origen con problema, criterios y duplicate search.
2. Aprobación explícita (etiqueta de aprobación o decisión del mantenedor).
3. Rama de trabajo generada y validada contra el gate del repo.
4. Preflight local completo (paridad local del job que gatea el merge).
5. PR con formulario completo y convención de commits.
6. Etiquetas de tipo y estado aplicadas por humanos.
7. CI: jobs requeridos en verde sobre el diff del PR.
8. Revisión: evidencia estructurada cuando el diff es de alto riesgo.
9. Merge por el pipeline (sin rutas paralelas ni fixes de rama).
10. Evidencia de despliegue o batería registrada sobre el SHA desplegado.

El recorrido genérico (idea → duplicate search → issue → approval → claim →
branch → local validation → PR → labels → review → CI → merge → artifact →
deployment → verification → rollback → closure) lo posee
`repository-delivery-governance` en su paso «Map the lifecycle»: remítase a
él para las transiciones genéricas y use este bloque para los eslabones que
el patrón web añade (preflight de paridad, etiquetas de tipo/estado,
evidencia por SHA).

## Bloque 2 — Matriz de enforcement (plantilla)

Una fila por regla del patrón. La clase de enforcement usa el vocabulario
exacto de la HR-4 de `repository-delivery-governance` (separate policy from
enforcement): `documented-only`, `locally-validated`, `CI-enforced`,
`host-enforced`, `manual`. Toda regla descrita como gate sin evidencia
ejecutable se registra `documented-only`, no como control.

| Regla | Fuente (SKILL.md) | Capa | Failure mode | Recovery path | Evidencia actual |
|---|---|---|---|---|---|
| Paridad del agregador | HR-29 | CI-enforced | job requerido sin paridad | añadir el job o la entrada de política | `assets/required-jobs/` |
| Checks requeridos por evento | HR-30 | CI-enforced | check publicado desde schedule | workflow-policy gate | `assets/workflow-policy/` |
| Verde obsoleto | HR-31 | CI-enforced | gate que lee datos mutables | comparación contra la rama base | `assets/required-jobs/` |
| Baselines/allowlists editables por el PR | HR-15 | CI-enforced | relajación sin motivo | comparación contra la rama base | `assets/ratchet/` |
| Dormant con policy file | HR-18 | CI-enforced | gate dormido sin candado | re-activación por review | `assets/hr-gate-matrix/` |
| Todo control falla ante su violación | HR-32 | CI-enforced | gate sin fixture violador | receta en el meta-test | `assets/hr-gate-matrix/tests/test_gate_violations.py` |
| Exenciones por identidad | HR-33 | CI-enforced | exención sin responsable | escaneo de documentos de gobierno | `assets/exemptions/` |
| Deriva del host | HR-34 | CI-enforced | regla declarada sin readback | drift check de solo lectura | `assets/host-readback/` |
| Evidencia de deploy por SHA | HR-10 | CI-enforced | evidencia sobre rama o variable global | gate de evidencia | `assets/deploy-evidence/` |
| Sustituto de revisión | HR-40 | documented-only | entrada de revisión aproximada | resolver desde lo desplegado | pendiente de mecanizar |
| Salud del presupuesto | HR-38 | documented-only | presupuesto sin métrica | reporte de salud | pendiente de mecanizar |
| Gates función del diff | HR-39 | CI-enforced | rojo externo bloquea PRs | re-agenda programada + manifiestos | `assets/workflow-policy/` |

Complete cada fila con la evidencia vigente en el consumer auditado; una
fila sin evidencia que pruebe la capa declarada es en sí un hallazgo.

## Bloque 3 — Consultas de solo lectura al host

Todas las consultas son GET. Sustituya `<owner>/<repo>` por el consumer
auditado y congele las salidas (JSON) como evidencia:

1. Protección de la rama por defecto:
   `gh api repos/<owner>/<repo>/branches/main/protection` — required checks
   (contexts y checks con app_id), enforce_admins, required_linear_history,
   allow_force_pushes, allow_deletions, required_conversation_resolution.
2. Rulesets: `gh api repos/<owner>/<repo>/rulesets` — nombre y enforcement
   de cada uno; detalle por id.
3. Etiquetas: `gh api repos/<owner>/<repo>/labels?per_page=100` — nombre,
   color y descripción; contraste con las etiquetas declaradas.
4. Métodos de merge:
   `gh api repos/<owner>/<repo>` — allow_merge_commit, allow_squash_merge,
   allow_rebase_merge, allow_update_branch, allow_auto_merge,
   delete_branch_on_merge.
5. Últimos N PRs fusionados (métricas del patrón):
   `gh api "repos/<owner>/<repo>/pulls?state=closed&per_page=100"` filtrando
   `merged_at != null`: autor, quien fusiona (`merged_by`), revisiones
   (`reviews`: aprobaciones y cambios solicitados), tamaño (adiciones y
   borrados), etiquetas.
6. Últimas N ejecuciones de CI:
   `gh api "repos/<owner>/<repo>/actions/runs?per_page=100"` — conclusión,
   evento, rama, y el job que falla vía
   `gh run view <id> --log-failed` (solo lectura).

Congele cada respuesta con su fecha de captura; la auditoría es una foto,
no un estado vivo.

## Bloque 4 — Preguntas del patrón (una por HR comprobable)

Cada pregunta lleva la evidencia que la responde. Las preguntas destiladas
de issues aún no fusionadas se marcan «pendiente» y no se auditan.

1. **¿Paridad del agregador?** (HR-29) — Contraste `required_jobs` de la
   política con los jobs de los workflows: cada job requerido existe como
   job y cada job del workflow está en la política. Evidencia: policy JSON
   y ficheros de workflow.
2. **¿Los checks requeridos se publican solo desde eventos de PR?** (HR-30)
   — Triggers de los workflows que publican los contexts requeridos:
   `pull_request` sí; `schedule`/`workflow_dispatch` no publican el nombre
   requerido. Evidencia: bloques `on:` y `if:` de los jobs.
3. **¿Hay verde obsoleto?** (HR-31) — Gates que leen datos mutables del PR
   o de la issue enlazada (cuerpo, etiquetas) en el momento del veredicto.
   Evidencia: código del agregador y condiciones de los jobs.
4. **¿El mismo PR juzgado puede editar las entradas del gate?** (HR-15) —
   Baselines, allowlists y relaxation fields que el PR toca: cada entrada
   editable debe contrastarse contra la copia de la rama base y llevar su
   campo de datos con motivo y referencia. Evidencia: diff del PR sobre los
   ficheros de datos del gate.
5. **¿Cuál es el alcance de los gates dormidos?** (HR-18) — Gates con
   `enforcement: "dormant"`: policy file presente, snapshot de activación
   inmutable, supresión únicamente de la parte que se durmió, re-activación
   como cambio de datos con review. Evidencia: policy files y su historia.
6. **¿Cada control tiene un test que lo ve fallar?** (HR-32) — Receta de
   violación por gate cuya ejecución sale con código distinto de 0.
   Evidencia: meta-test de violaciones y su salida.
7. **¿Las exenciones llevan identidad verificable?** (HR-33) — Toda
   exención declarada como dato lleva `actor` no vacío o `issue` positivo.
   Evidencia: escaneo de baselines, políticas y matrices.
8. **¿Hay deriva entre la documentación y el host?** (HR-34) — Contrato
   declarado contra la API real (protección, rulesets, etiquetas, métodos
   de merge): toda regla host-enforced presente, toda ausente es hallazgo.
   Evidencia: drift check de solo lectura sobre las consultas del bloque 3.
9. **¿La entrada de revisión previa se resuelve desde lo desplegado?**
   (HR-40) — Los gates de deploy o batería resuelven su entrada de revisión
   desde el estado real del despliegue (salud, SHA desplegado), no de una
   variable global ni de la rama por defecto. Evidencia: scripts de los
   workflows de deploy.
10. **¿El presupuesto de revisión tiene salud medida?** (HR-38) — Métrica
    de salud con ventana, excepciones, tasa y umbral declarado. Evidencia:
    reporte de salud y su baseline.
11. **¿Los gates requeridos por PR son función del diff?** (HR-39) — Gates
    cuyo veredicto depende de estado externo ejecutan programados sobre la
    rama por defecto y por PR solo cuando el diff toca los manifiestos
    declarados; su rojo programado abre issue con la evidencia. Evidencia:
    triggers de los workflows y política de rutas.

## Bloque 5 — Salida

1. **Tabla de hallazgos**: una fila por desvío con la evidencia como
   `fichero:línea` o campo de la API congelado, la regla (HR) que lo
   convierte en hallazgo y la capa declarada frente a la observada.
2. **Afirmaciones no verificadas**: toda afirmación de la documentación que
   el auditador no pudo contrastar con una lectura (falta de acceso, dato
   no congelado, consulta no ejecutada). Cada una es candidata a regla o a
   gate; ninguna se da por cierta.
3. **Matriz de enforcement** del bloque 2 completada con la capa observada
   junto a la declarada: toda discrepancia es un hallazgo de deriva (HR-34).
