# access2web-blueprint

<!-- personal-skills:slice:access2web-blueprint @ vf8c9a61 -->
# slices/partials/web.md

## Manera de trabajar en proyectos web

> Aplica a todo `primary_type: web` del catálogo. Las invariantes de ciclo de vida de PR aquí enunciadas se complementan con las skills universalmente activas — véase `personal-skills/AGENTS.md` raíz para el sistema de propagación, `propagate-team-skills.ps1` para la mecánica de distribución, y el bloque de partials específicos del consumer para las convenciones del proyecto concreto.
>
> Este partial enuncia invariantes. Los procedimientos asociados viven en sus skills respectivas — no se duplican aquí.

### Arranque (invariante)

- **Ningún trabajo empieza sin sincronizar la rama base.** Antes de leer, decidir o tocar un fichero: `git fetch origin && git switch <active_branch> && git pull --ff-only` y, justo después, `ci-pattern checkout-sync`. Un checkout atrasado muestra gobernanza, CI y skills que **ya no rigen**. Si sale 1, seguir la instrucción que imprime (en la base, `git pull --ff-only`; en una rama de trabajo, `git merge origin/<base>`). Si sale 2, no puede medir: no empezar hasta sincronizar (#914). Es un control local y sin red: no es un gate de CI ni bloquea pushes.
- **Señal de alarma, no conclusión.** Si un documento o fichero citado por la gobernanza del consumer **parece no existir**, o **contradice** el modelo vigente, la lectura correcta es "estoy en un checkout atrasado": sincronizar y volver a mirar. Nunca "el repo tiene un hueco" ni "la regla cambió sin avisar".
- **Gobernanza viva, no copiada.** Las issues se leen del tracker (`gh issue view <n>`), nunca de una copia pegada; el `AGENTS.md` y el documento de flujo del consumer se leen **en el SHA sincronizado**.
- **El workspace del runner de CI no es un clon de desarrollo.** Se trabaja en un worktree o clon propio (ver §Worktree y rama remota).
- **Una copia global de skill desactualizada tapa la del repo** (ver §Carga de skills): si una skill del repo contradice a la global, gana la del repo, y la global está pidiendo actualización.

### Skill Registry Protocol (obligatorio)

La tabla de skills del `AGENTS.md` del consumer es un subconjunto curado; el registro generado — `.atl/skill-registry.md`, refrescado por la skill `skill-registry` (`gentle-ai skill-registry refresh`) — es la fuente de verdad de descubrimiento. Una regla que existe y está indexada pero que ninguna sesión carga no gobierna nada: este protocolo convierte el índice en despacho.

1. **leer** — antes de escribir o revisar código, CI o documentación, lea `.atl/skill-registry.md` (o el índice de skills del repo).
2. **contrastar** — contraste la tarea contra la columna `Trigger / description` de las skills candidatas.
3. **cargar** — lea íntegros los `SKILL.md` que matcheen antes de actuar.
4. **delegar** — al delegar a un subagente, pase las rutas exactas de los `SKILL.md` bajo un encabezado `## Skills to load before work` en la tarea. Un nombre en prosa no basta: el subagente no debe tener que adivinar ni redescubrir (eso produce el degradado `fallback-registry`).
5. **GAP** — si ninguna skill matchea y la tarea toca CI/PRs, releases, la flota de runners o la adopción de un repositorio nuevo, es un gap de orquestación: dígalo explícitamente en vez de proceder de memoria.

- **Deriva índice/registro.** La tabla del `AGENTS.md` es la puerta del contribuyente; el registro es la verdad operativa. Que una skill no aparezca en la tabla no autoriza concluir «no aplica»: autoriza buscarla en el registro (`skill-registry` / `gentle-ai skill-registry refresh`).
- **Frescura del registro.** Si `.atl/skill-registry.md` es anterior al último cambio de skills, está stale: refresque el registro antes de delegar.

### Forma del ciclo

- Toda issue es **atómica**: la cambia una persona, la cierra un PR (o varias si encadenadas vía `chained-pr` cuando la diff supera el presupuesto).
- **Toda issue nace con 6 secciones canónicas en el cuerpo** (#432). El contrato de `issue-spec` se compone de los seis encabezados H3, exactos, en este orden:
  1. `### Problema y contexto`
  2. `### Evidencia verificable`
  3. `### Alcance y no objetivos`
  4. `### Criterios de aceptación`
  5. `### Plan de validación`
  6. `### Dependencias y riesgos`

  El contrato se aplica en el momento de crear la issue: la receta canónica es
  `gh issue create --title "<título>" --body-file <cuerpo-de-la-issue.md> --label type:<tipo>`
  (donde `type:<tipo>` sale del host-contract del consumer), o el formulario
  de GitHub `.github/ISSUE_TEMPLATE/<feature>.yml` que ya lleva los seis
  campos como controles. El doc operativo generado por `ci-pattern adoption
  generate-doc` enuncia el comando exacto para el consumer; la rama de
  aprobación por CLI queda cubierta por esa receta. El validador server-side
  que existía en `.github/workflows/issue-spec-open.yml` se retiró en #508
  porque había fallado 6/6 por facturación de runners hospedados (#380); el
  control sigue siendo fail-loud: una issue sin las 6 secciones no se crea
  vía CLI, y un agente que la escriba a mano verá la PR que la cite abrirse
  con el cuerpo incompleto y deberá cerrarla o completarla antes del merge.
- Toda issue aprobada tiene una **rama propia** con el nombre `<tipo>/<nº issue>-<kebab-slug>`, validado por el **ruleset del host** que aplica la regex de `P01_branch_name_pattern`, o por el gate `assets/branch-name/check_branch_name.py` del patrón cuando el host no declara rulesets (mecánica del HR-29 y HR-31).
- Toda rama se desarrolla en un **worktree dedicado** bajo el layout canónico de `worktree-reorg-per-project` v2.0 (sibling-container `<project>-worktrees/<nº issue>-<tema>/`); el comando literal es `git worktree add ../<repo>-worktrees/<nº>-<tema> -b <rama-nueva>` y la limpieza usa la misma ruta.
- Toda PR apunta a `main`. Los repos web no tienen rama de integración: `staging`, `develop` o `release/*` no reciben PR ni se mergean a `main`; lo reutilizable se recupera con `git cherry-pick` en una rama nueva desde `main`. Preproducción y UAT son entornos que despliegan desde `main`, no ramas.

### Presupuesto de revisión

- Toda PR se mantiene bajo el **presupuesto de revisión de 400 líneas** (`additions + deletions`), comprobado por el gate `assets/pr-size/check_pr_size.py` del patrón (HR-33, `SIZE_EXCEPTION_APPROVER`).
- 400 líneas es **techo de revisión, no techo de tamaño**: a partir de esa cifra la revisión pasa de atenta a vistazo. La justificación completa vive en el `CONTRIBUTING.md` de cada consumer; este partial la enuncia sin duplicar.
- **Lectura obligatoria antes de abrir una PR:** el `CONTRIBUTING.md` del consumer (HR-71). Recoge cómo se aplica aquí el flujo del catálogo: issue primero, asistencia de IA declarada en el cuerpo, etiquetas, presupuesto, commits por unidad, comprobaciones previas y automáticas, enlace a la issue. Sin `CONTRIBUTING.md` el gate `docs_manifest` falla cerrado en fase 5.
- La única excepción al presupuesto es la **etiqueta `size:exception`** (decisión del dueño del repo, 2026-10-15). El campo `size-exception-reason:` del cuerpo del PR **no existe**; la identidad del aprobador llega al gate en la variable de entorno `SIZE_EXCEPTION_APPROVER` y `check_pr_size.py` (`override_has_separate_approver`, HR-33) falla cerrado si el aprobador coincide con el autor.
- La etiqueta `size:exception` la aplica **solo el mantenedor**, nunca el agente, y solo cuando partir el PR rompería `main` (no por tests, refactor o cambio menor). Antes de aplicarla, el agente responde las cuatro preguntas de `## Excepción de tamaño` en el cuerpo de la PR; sin las cuatro respuestas la etiqueta no se pide. El merge sigue siendo una decisión humana.
- **Borrar siempre es divisible.** Retirar tests, código muerto o configuración nunca justifica `size:exception`: se borra por trozos de ≤400 líneas, aunque sea dentro del mismo fichero, dejando `main` en verde entre trozos; lo que comparten los trozos (helpers, fixtures) se borra en el último.
- Cuando la diff supera el presupuesto, el orden de escape es: (1) partir por unidad de trabajo, (2) encadenar PRs vía `chained-pr`, (3) `size:exception` (la pone el mantenedor) como último recurso. Si la excepción se vuelve habitual, el problema está en el troceado del issue, no en el presupuesto.

### Evidencia e2e por PR (invariante)

- **Toda PR que toque una superficie de UI añade o actualiza al menos un flujo
  e2e de navegador.** El flujo entra por la navegación real del consumer (no
  por llamadas directas a funciones internas) y vive en el directorio e2e que
  el consumer declare en su política de CI (p. ej. `tests/e2e_ci/`).
- **La exención es explícita y está justificada.** Solo se concede con una
  sección dedicada (`## Exención e2e` por defecto; el matching es tilde-
  tolerant y también acepta `## Exencion e2e`) en el cuerpo de la PR con la
  razón escrita; una sección vacía no exonera nada.
- **El control es detectivo y parametrizado.** El gate `e2e-evidence` del
  patrón (copia instalada en `.agents/skills/ci-pattern/assets/e2e-evidence/`)
  lee como datos del consumer los globs de UI, la lista blanca de rutas no-UI,
  el directorio de flujos y la sección de exención. Se distribuye como asset
  copiado y **nunca invoca la CLI del patrón como subproceso**. Falla cerrado:
  sin UI detectada no aplica; con UI sin flujo y sin exención, rojo.
- Donde el consumer no ejecute el gate en el job de PR, la invariante la
  sostiene el revisor: una PR con UI sin evidencia ni exención justificada no
  está lista para mergear.

### Worktree y rama remota

El ciclo de vida de una feature es único y se apoya siempre en la **rama activa
declarada** para ese consumer (`fleet/registry.json`), nunca en una base impuesta
a mano:

1. **Arranque.** Worktree dedicado y rama propia desde la rama activa
   actualizada, no desde el checkout compartido.
2. **Sincronización.** Al terminar la feature, `git fetch origin` y `git merge
   origin/<active_branch>` **dentro de la rama de feature** —nunca el merge en
   sentido inverso, nunca rebase, `--amend` ni force-push—; resuelva los
   conflictos y repita el preflight y las pruebas aplicables sobre el resultado
   integrado.
3. **Entrega.** Push normal y PR hacia la `<active_branch>` (o actualización del
   PR existente). Si la base avanza durante la revisión, repita la sincronización
   y las comprobaciones: un verde anterior a la integración no es evidencia del
   resultado nuevo.
4. **Merge y limpieza.** El merge lo lanza un mantenedor autorizado. Un agente
   lo lanza solo por orden registrada de un mantenedor, con cuenta de mantenedor
   (`governed-delivery` regla 2), y nunca con `--admin`. Solo después del merge verificado y sin cambios pendientes, retirar el
   **worktree local** y la **rama local** con `safe_cleanup` (ver `ci-pattern`
   `assets/runner-controls/safe_cleanup.py`, invocado desde `worktree-reorg-per-project`
   Phase 3). `safe_cleanup` falla cerrado si la punta no esta preservada (no
   coincide con el remoto, no es ancestro de la base activa, ni coincide con la
   cabeza de una PR mergeada). La **rama remota se conserva siempre**: nunca
   `git push origin --delete <rama>` ni `--delete-branch`.

   **Definición de hecho** (#127, #420). Una unidad de trabajo está terminada
   cuando su PR está mergeada **y**, inmediatamente después, su worktree local
   y su rama local están borrados. El worktree local se borra con
   `git worktree remove <path>` seguido de `git worktree prune` (mecánica en
   `worktree-reorg-per-project` Phase 3). La rama local se borra con
   `git branch -D <rama>` **solo si su punta es la de la PR mergeada o ya está
   en la rama activa**; si quedan commits sin subir, no se borra: se suben a
   una PR o se escalan al responsable. La **rama remota se conserva siempre**,
   aunque la PR ya esté mergeada: nunca `--delete-branch` ni `git push origin
   --delete <rama>`. La limpieza no se aplaza: forma parte de la entrega. Si
   algo no es borrable de forma segura, `safe-merge` (#426) lo deja como
   `CLEANUP PENDING` en vez de forzar el borrado.

El procedimiento detallado —comandos exactos y condiciones de «hecho»— vive en
la skill de día a día del consumer y en su documento operativo generado; este
partial solo enuncia la invariante y no lo duplica.

- La estrategia de merge (`--squash` o `--no-ff`) es decisión del consumer; ambas son válidas. La **invariante** es que la rama remota sobreviva al merge, no la forma concreta del commit en `main`.

### Propiedad de recursos de contenedores (invariante)

La máquina donde trabaja un agente puede ser la misma que sirve producción.
Un agente solo para o borra los contenedores, volúmenes y redes que haya creado
él mismo dentro de la tarea en curso, y los reconoce por su prefijo.

- **El nombre no autoriza el borrado.** `docker rm`, `docker stop` y
  `docker volume rm` se aplican únicamente sobre recursos con el prefijo de la
  tarea. Que un nombre parezca de prueba no es una licencia para borrarlo.
- **Los recursos protegidos son intocables.** Nunca se tocan los que llevan
  etiqueta de protección, los de la plataforma (Coolify o su proxy) ni los de
  los runners.
- **La producción se identifica sola por disciplina del agente.** El gate
  `check_docker_ownership.py` solo enforça la **etiqueta de protección**
  (`apap.protect=true` por defecto, parametrizable por consumer): recorre
  `protect_labels` y rechaza el borrado de cualquier recurso que la lleve,
  sea de producción o no. El **patrón `*-prod-*`** en el nombre y el
  **volumen con nombre** —nunca anónimo— son disciplina del agente que el
  gate no verifica; siguen siendo obligatorios como convención operativa
  (Refs DysTelefonica/team-skills#454).
- **Los barridos automáticos excluyen la etiqueta.** Todo reaper de recursos
  huérfanos deja fuera lo etiquetado como protegido; limpiar no es borrado
  preventivo.

### Disciplina de revisión

- El CI corre **una vez, en la PR**: no hay `strict` ni cola de merge, y la rama base puede avanzar durante la vida del PR sin invalidar los checks. La red de seguridad después del merge es el workflow de `push` a `main` más un nightly diario.
- **Rojo en CI pisa todo el merge.** Regla humana: el revisor no debe pulsar merge con ningún check rojo, ni siquiera si el rojo parece trivial. Complemento técnico: `repository-delivery-governance` HR-7 + `deterministic-quality-harness` HR-1 fail-loud atajan el escenario cuando hay branch protection automatizada.
- **El agente no usa `--admin` ni fuerza el merge.** La decisión de fusionar es humana (HR-50) y la regla anterior es inseparable del bypass: el agente que detecte un merge bloqueado por checks rojos, por la protección del ruleset o por un SHA desactualizado **lo escala al mantenedor** (issue, comentario en el PR o canal acordado), nunca aplica `--admin`, `--force` ni equivalente en CLI; y, por la API, nunca envía `merge_method: admin` ni `sha` distinto al verificado por CI. El job de merge gobernado del consumer (HR-50, `assets/runner-controls/check_governed_merge.py`) aplica `merge` o `squash` con la identidad del bot sobre el SHA verificado, sin bandera admin; si una PR exige admin para entrar, el flujo está roto y se arregla el flujo, no se salta.
- El job `merge-ready` signal-only que consulta `mergeable` y `reviewDecision` está retirado (no cazaba defectos): sin `strict` ni cola de merge, la barrera la aplica el ruleset del host y la red de seguridad post-merge (`push` a `main` + nightly diario), y replicarla en CI duplicaba la señal sin añadir garantía.
- Cuando el CI rojo es por **infra** (runner colgado, red, secret rotado), abrir issue bloqueante de CI y enlazarla desde el PR; no embutir la fix infra en el PR del feature salvo que sea ≤30 LOC y se cierre en el día.

### Anti-slop y atribución

- Anti-slop y anti-sobreingeniería de IA se delegan a la skill del catálogo `ai-slop-discipline` (comprobación de 23 puntos, 3 firmas, límite de delegación para subagentes). El partial no redefine la disciplina — el consumer que adopte la skill la aplica automáticamente antes de cualquier cambio.
- **Sin atribución de IA en commits**: no se añade `Co-Authored-By: ... <AI>` ni equivalente. Esta regla vive también en `personal-skills/AGENTS.md` raíz de la flota; el partial la refleja para que sea visible en el slice de web.
- Mensajes de commit en conventional commits. Castellano peninsular formal (usted) en artefactos documentales raíz; inglés en código, comentarios, mensajes de commit y PR bodies.

### CodeGraph preflight (cuando aplique)

- Si el consumer tiene índice CodeGraph (`.codegraph/` presente), el workflow sigue `engineering-workflow` líneas 64-71: `codegraph init` antes del primer edit, `codegraph_explore` antes de cualquier grep/read/glob amplio. No se reinventa aquí.
- Si el consumer no soporta CodeGraph, el preflight se omite sin romper invariante — la regla es "usar CodeGraph cuando esté disponible", no "requerir CodeGraph siempre".

### `ci-pattern` — a demanda, salvo `governance: ci-pattern` en el registro

- Por defecto, `ci-pattern` — **a demanda, no se carga por defecto.** El patrón de gobierno de CI (gates deterministas, presupuesto de revisión, evidencia por SHA, cadenas de PRs, protocolo de adopción/porting) no llega a este repositorio por propagación ni debe asumirse presente ni vendorizarse. Se invoca por CLI cuando la tarea sea adoptar el patrón en un repositorio nuevo o auditar el CI: `python3 ~/.agents/skills/ci-pattern/assets/bin/ci-pattern verify|status|params validate|manifest show` (contrato: `~/.agents/skills/ci-pattern/references/cli-spec.md`).
- Excepción: si `fleet/registry.json` declara `governance: ci-pattern` para el repo, la skill SÍ llega propagada — partial de gobierno dentro del bloque del slice y copia en `.agents/skills/ci-pattern/`, con las skills declaradas en `fleet/governance-tiers.json`. Vendorizarla a mano en `skills/` sigue prohibido.

### Divergencias documentadas (no son invariantes)

Estas decisiones quedan a la flota / consumer; el partial las registra para que las revisiones no las traten como incumplimientos.

- **Merge strategy.** `--squash` o `--no-ff`, ambos válidos. Invariante compartida: la rama remota se preserva en cualquier caso.
- **Etiqueta `size:exception`.** Solo el mantenedor la aplica cuando partir el PR rompería `main`; el campo `size-exception-reason:` está retirado y el agente nunca se pone esa etiqueta.
- **Rama base única.** Todo aterriza en `main`, antes y después del MVP.
- **Convención multi-app.** Si el consumer migra varias apps, el prefijo de issue/commit es decisión propia; el commit debe identificar el scope de cualquier manera.

### Procedencia (skills que alimentan este partial)

Cada invariante de este partial se ancla a una skill específica del catálogo o upstream. Si la skill referenciada cambia su HR, este partial requiere reauditoría.

- Issue-first atómica → `intake-roadmap-loop` HR-1, HR-14; upstream `engineering-workflow` Step 2.
- Rama `<tipo>/<nº>-<slug>` → `repository-delivery-governance` HR-4 (documentado/CI-enforced).
- Un worktree por issue → `worktree-reorg-per-project` v2.0 (layout + Phase 3 cleanup).
- Rama remota preservada + limpieza segura post-merge → `safe_cleanup` (DysTelefonica/team-skills#436) invocada desde `worktree-reorg-per-project` Phase 3.
- Base = `main` para todo el ciclo de vida del consumer web; sin flip post-MVP → upstream `engineering-workflow` líneas 46-51.
- 400 líneas + `size:exception` → `deterministic-quality-harness` Decision Gate línea 56; upstream `chained-pr` HR-1 línea 16.
- Rojo en CI pisa todo → `repository-delivery-governance` HR-7; upstream `engineering-workflow` línea 79.
- Ciclo de entrega (worktree y rama desde la base activa actualizada → merge de la base en la feature → preflight/pruebas → push → PR → CI sobre cabeza y base actuales → merge del mantenedor → retirada solo del worktree local) → `governed-delivery` regla 1 y el documento operativo generado por `ci-pattern` (#390).
- Sincronización de la rama publicada por `git merge`, nunca rebase → `governed-delivery` regla 1; `ci-pattern` HR-54; #333.
- Rama remota preservada tras el merge, worktree local retirado solo al final → `worktree-reorg-per-project` Phase 3.
- Evidencia e2e por PR (UI exige flujo e2e o exención justificada; gate
  `e2e-evidence` parametrizado, fail-closed) → `ci-pattern` asset `e2e-evidence` (#425).
- Propiedad de recursos de contenedores (solo lo propio con prefijo; producción con `*-prod-*`, volumen con nombre y etiqueta de protección) → DysTelefonica/team-skills#454, por el incidente de APAP_WEB del 2026-10-12: un agente borró la base de producción con `docker rm -fv` porque el contenedor tenía nombre de prueba y volumen anónimo. Invariante de flota sin skill fuente.
- Anti-slop → `ai-slop-discipline` (catálogo, T1 universal).
- CodeGraph preflight → upstream `engineering-workflow` líneas 64-71; `codegraph-usage` HR-1, HR-2.
- Conventional commits + castellano peninsular en artefactos + sin atribución IA → `personal-skills/AGENTS.md` raíz de flota.
- Skill Registry Protocol (`READ`/`MATCH`/`LOAD`/`DELEGATE`/`GAP` + deriva índice/registro + frescura) → upstream `gentle-ai` `internal/assets/skills/_shared/odd-orchestrator-sections.md:183-201`, adaptado vía `gentle-ai-governance-dispatch.md` (adopciones 1-3 del informe de gobernanza 2026-10-02).
- Arranque sin checkout atrasado (sincronizar la rama base antes de leer, decidir o tocar) → `personal-skills/AGENTS.md` raíz de flota (§Lectura obligatoria). Invariante de flota sin skill fuente.

### Cómo auditar este partial usted mismo

Procedimiento de validación periódica (mensual o por release de skill fuente):

- Confirmar que las HRs citadas en §Procedencia siguen existiendo con la misma numeración y redacción en el cuerpo actual de cada skill. Si una skill referenciada cambia su HR, este partial requiere reauditoría.
- Confirmar que el partial no prescribe rebase ni reescritura de historia y que el ciclo de entrega sigue coherente con `governed-delivery` y el documento operativo generado.
- Confirmar que no se haya añadido regla con cuerpo procedural en este partial — los procedimientos viven en skills, no aquí.
- Confirmar que la sección §Divergencias documentadas sigue reflejando las variantes reales de los consumers actuales.

### Antipatrones

- "Vendorizar `ci-pattern` en `skills/ci-pattern/` del consumer (genera drift: 3 PRs de sync en un día) → usarlo por CLI a demanda" — la skill vive fuera del catálogo `skills/`; el propagador solo la copia al tier declarado en `fleet/governance-tiers.json`, nunca a mano (CLI o tier, según `governance` en el registro).
- "El documento no existe / la regla cambió" concluido desde el checkout local sin haber sincronizado la rama base — primero `git pull --ff-only`, después la conclusión.
- "Delegar mencionando la skill por nombre en prosa, sin rutas exactas bajo `## Skills to load before work`" — el subagente redescubre, cae en `fallback-registry` y opera de memoria.
- "No está en la tabla del AGENTS.md, luego no aplica" — la tabla es un subconjunto curado; la ausencia solo autoriza una búsqueda en el registro, nunca una conclusión.
- "Esperar a que CI esté verde para mergear" con la base ya avanzada y sin integrarla (merge) en la rama de feature — el verde contra base obsoleta es stale-green.
- "Sincronizar la rama de feature con la base por rebase, `--amend` o force-push" — la historia publicada se actualiza con `git merge`; reescribirla invalida el PR y el trabajo de quien venía revisando.
- "Borrar la rama remota post-merge porque ya está mergeada" — destruye la granularidad por unidad de trabajo que el flujo pretende crear.
- «Borrar un contenedor o un volumen porque su nombre parece de prueba» — la propiedad se decide por el prefijo de la tarea y por la etiqueta de protección, nunca por la apariencia del nombre; el borrado de un recurso de producción no lo recupera ni el reintentar el deploy (#454).
- "PR con 600 líneas porque el feature lo requiere" — partir primero, encadenar después; `size:exception` es el último recurso, no la primera opción.
- "Mergear con CI rojo aunque el rojo parezca trivial" — la trivialidad la decide el revisor, no el autor.
- "Esperar a que el reviewer apruebe manualmente aunque todos los checks estén verdes" en proyectos con auto-merge standing explícito — revisar la sección de revocación de `merge-workflow.md §15.6` antes de saltarse el gate.

### Carga de skills (orden de resolución)

Al buscar una skill de este catálogo, resuelva en este orden:

1. Por nombre. La ruta física la decide el agente (opencode | claude
   | codex | cursor | pi) y, en el caso de Pi, la variable
   `PI_CODING_AGENT_DIR`. La separación entre los cinco destinos
   compartidos del catálogo y la ruta reservada a skills exclusivas
   de Pi (HR-12) está documentada en la skill `skill-propagation`.
2. Repo: `.agents/skills/` del consumer, materializada desde `.team-skills.yaml` por `scripts/refresh-team-skills.ps1` (vía rama `skill-fleet/<name>`).
3. Catálogo: `skills/<nombre>/SKILL.md` de personal-skills.

Referencias por nombre en prosa; ruta absoluta exacta en la delegación: al pasar una tarea a un subagente, copie la ruta del `SKILL.md` desde el registro bajo `## Skills to load before work` (ver §Skill Registry Protocol). La regla anti-ruta-absoluta gobierna el texto de skills y docs — las rutas varían por máquina — no la delegación, donde la ruta exacta es obligatoria. Nota de precedencia: pi nativo conserva la primera copia encontrada (gana la global), mientras que el registry de gentle-pi prefiere la copia del repo; una copia global desactualizada tapa la del repo en pi, así que las copias globales deben seguir al catálogo. Si una skill citada en este partial no resuelve en ninguno de los tres niveles, repórtelo: falta declararla en `.team-skills.yaml` del consumer.

### Auditoría (Refs #490)

- **2026-10-13:** el texto de *Disciplina de revisión* (linea 108) y la entrada de *Anti-slop* correspondiente (linea 175) ya enuncian la regla con `git merge` y la prohibición de reescritura de historia; el cambio que #490 propone como "Texto propuesto" coincide con el actual. Sin edición de fuente requerida. Si el duplicado que reporta el issue vive en la versión propagada del partial (por ejemplo el `AGENTS.md` del consumer referenciado por el issue), la causa es una propagación pendiente, no un bug del fuente.



## Antes de actuar

**Antes de cualquier cambio, cargue `ai-slop-discipline` y pase su
comprobación** (`.agents/skills/ai-slop-discipline/SKILL.md`).

Este archivo solo orienta: las reglas completas, con su motivo, viven en
las skills. Cargue la que corresponda antes de actuar.

| Si va a… | Skill | Ruta instalada |
| --- | --- | --- |
| cualquier cambio (obligatorio) | `ai-slop-discipline` | `.agents/skills/ai-slop-discipline/SKILL.md` |
| aprobar una issue, actualizar rama, abrir PR, encadenar, pedir merge | `governed-delivery` | `.agents/skills/governed-delivery/SKILL.md` |
| auditar o endurecer la gobernanza del repo (políticas, etiquetas, protección) | `repository-delivery-governance` | `.agents/skills/repository-delivery-governance/SKILL.md` |
| adoptar o auditar el patrón, tocar CI o HRs | `ci-pattern` | `.agents/skills/ci-pattern/SKILL.md` |
| preflight y calidad determinista | `deterministic-quality-harness` | `.agents/skills/deterministic-quality-harness/SKILL.md` |
| trabajar en worktrees | `worktree-reorg-per-project` | `.agents/skills/worktree-reorg-per-project/SKILL.md` |

Invariantes, sin excepción: nada directo a la rama protegida; sin
rebase, `--amend` ni force-push sobre ramas publicadas (HR-54); el merge
lo lanza un mantenedor por el camino gobernado; un agente lo lanza solo
por orden registrada de un mantenedor, con cuenta de mantenedor, y nunca
con `--admin`; sin
nombres propios — solo roles y rutas instaladas. El detective post-push y
el guard de force-push abren el incidente.

<!-- ci-pattern:hr-block:begin -->

<!-- ci-pattern:hr-block:end -->

<!-- ci-pattern-agents:begin -->
## Flujo operativo y mandato de no invención

Bloque gobernado: no lo edite a mano; se regenera desde el patrón.
Antes de actuar, lea la parte del flujo operativo de la situación. Si un paso,
comando, etiqueta o dato que necesita no está en este documento, NO lo invente:
deténgase y pregunte en la issue o al mantenedor.

| Situación | Qué leer antes de actuar |
| --- | --- |
| Crear una issue | Sección 1 (formulario, duplicados, aprobación) |
| Crear una rama | Sección 2 (regex y forma del nombre) |
| Sincronizar la rama de feature con la base | Sección 10, etapa 5 (integración por merge, nunca rebase) |
| Escribir commits | Sección 3 (convención y unidades de trabajo) |
| Abrir un pull request | Sección 4 (cuerpo, etiqueta, presupuesto) |
| Encadenar PRs | Sección 4 (cadena: `chain:partial`, `Refs #N`, punta) |
| Consultar checks | Sección 5 (tabla de checks y recetas de recuperación) |
| Pedir un merge | Sección 6 y la etapa 9 del ciclo de vida |
| Registrar un incidente de producción | Sección 8 (los cuatro niveles: post-mortem, registro operativo, secretos, datos personales) |
| Reportar una fricción | Sección 9 |
| Cualquier paso del ciclo | Sección 10 (ciclo de vida completo) |

**Mandato de no invención:** si un paso, comando, etiqueta o dato que necesita
no está documentado en el flujo operativo, no lo invente: deténgase y
pregunte en la issue o al mantenedor.

## Dónde vive la documentación

Toda la documentación de este proyecto vive en este mismo repositorio, bajo `docs/`
por materia; los post-mortems en `docs/postmortems/<AAAA-MM-DD>-<slug>.md`.
Fuera del repositorio, y sin copia en él, quedan solo los secretos, los
backups, los datos en bruto y el material privado. Nunca ubique documentación
en una ruta absoluta o externa al repositorio (HR-48): el gate de la fase 5 la
rechaza. Detalle: la sección «Dónde vive la documentación» del flujo operativo.
## Dónde trabajar

Antes de actuar, vaya al documento de su situación; no deduzca dónde va el trabajo.

| Si vas a… | Lee… |
| --- | --- |
| Cambiar código o buscar dónde va algo nuevo | `docs/CODEBASE-GUIDE.md` y, desde ella, la página de `docs/codebase/` del área que toque |
| Abrir un pull request | la checklist de su tipo de cambio en `docs/codebase/maintainer-playbook.md` |
| Proponer o cambiar algo que quizá ya se decidió | `docs/decisions/README.md`: consulte antes si ya se rechazó, sustituyó o aplazó |
| Crear un módulo o importar entre módulos | `docs/architecture.md`: árbol de módulos, capas y dependencias permitidas, antes de añadir un módulo o un import entre capas |
| Dudar de si algo existe en el repositorio | las páginas «No existe en este repo» de `docs/codebase/`, antes de inventar código |
| Abrir un pull request | `CONTRIBUTING.md` antes de abrir la PR (issue primero, asistencia de IA, presupuesto, checks) |
<!-- ci-pattern-agents:end -->

# Fragment access2web-blueprint — convenciones operacionales del repo

Este fragment SOBREESCRIBE / AMPLÍA las reglas del partial `web` con convenciones específicas de este consumer. Si una regla entra en conflicto con el partial, prevalece este fragment por ser consumer-specific. Las reglas del partial no se duplican aquí; este fragment solo añade lo que el partial no cubre.

## §F.1 Capas de documentación del repo

La documentación tiene tres capas. Confundirlas es el error más caro de este repo.

| Capa | Ubicación | Quién la mantiene |
|---|---|---|
| Narrativa del producto | `README.md`, `AGENTS.md`, `DOCS.md`, `CODEBASE-GUIDE.md`, `CONTRIBUTING.md`, `CHANGELOG.md`, `docs/` | Una persona, cuando cambia el enfoque |
| Contrato vivo de capacidades | `openspec/specs/` | La fase `archive` lo genera mecánicamente |
| Historia del porqué | `openspec/changes/archive/` | Inmutable. Nadie la escribe. |

**Regla crítica:** nunca redactar a mano en `openspec/specs/`. Esa capa se genera al archivar un change; escribirla directamente rompe la trazabilidad.

## §F.2 Front-door de lectura

Antes de tocar `app/`, `tests/`, `openspec/`, `docs/architecture.md`, o de proponer una decisión arquitectónica nueva (D-<n>), leer en este orden:

1. `AGENTS.md` — alcance del repo y skills disponibles
2. `CODEBASE-GUIDE.md` — overview, ownership, reading path raíz
3. `docs/architecture.md` — fuente de verdad única arquitectónica
4. `CONTRIBUTING.md` — workflow de contribución + convention multi-app + label system
5. `openspec/changes/<change>/design.md` (si el cambio pertenece a un change vivo)
6. `docs/03-aplicaciones/<app>/epic.md` (si toca una app específica)

Saltarse cualquiera deja a la IA operando contra arquitectura obsoleta.

## §F.3 Capas enforced por gate

Las capas hexagonales las impone el job `docs-architecture` (HR-68): compara `docs/architecture.md` con el código en los dos sentidos y ejecuta import-linter con `.importlinter`, derivado del bloque `architecture` de ese documento, que es la única fuente de las capas, de las dependencias permitidas, de la independencia entre aplicaciones y de la pureza de `domain`, `ports` y `application` (DA-1). Mover un adapter a la capa equivocada, o añadir un fichero fuera de las capas sin declararlo en `exempt`, es un gate failure, no un estilo. Esquivar el gate con `|| true` está prohibido por `tests/test_ci_workflow.py`.

## §F.4 Reglas de decisiones arquitectónicas (D-<n>)

- Toda D-<n> tiene estado: `vigente` u `OBSOLETO`. Las obsoletas se reemplazan, no se duplican.
- Buscar en `docs/architecture.md` §Decisiones arquitectónicas D-<n> cross-cutting vigentes antes de proponer una nueva.
- Las decisiones cross-cutting se promueven a `docs/architecture.md` con estado `vigente`. NO se quedan en `app/src/modules/<app>/`.
- Las D-<n> no se inventan. Si la decisión es nueva y cross-cutting, abrir issue `##ABIERTO##` en `openspec/changes/<change>/design.md` y bloquear el código hasta que la design doc declare la D-<n>.

## §F.5 Regeneración de walkthroughs

`docs/03-aplicaciones/<app>/walkthrough-*.json` se regenera con dysflow + codegraph-vba. Cambios "manuales" del JSON quedan desincronizados con la realidad del binario. No editar a mano.

## §F.6 Convención multi-app en commits

El prefijo de commit identifica la app afectada: `fix(expedientes): ...`, `docs(lanzadera): ...`, `chore(platform): ...`. Conventional commit con placeholder `(app)` literal se considera violación de convention (ver `CONTRIBUTING.md` §Convención multi-app). El scope del commit debe ser siempre una sola app o plataforma; los commits cross-cutting van a `platform/` o `chore(platform):`.

## §F.7 Skills locales no distribuidas por catalog

Este repo mantiene dos skills propias que **no están en el catalog** y no se sobrescriben en propagación:

- `skills/architecture-guardrails/` — front-door de lectura, gates de arquitectura, D-<n> management. Se carga antes de cualquier cambio estructural.
- `skills/documentation-alan-style/` — formato documental (Castellano peninsular, sentence case, sin emojis). El catalog tiene su propia `documentation-alan-style` que SOBREESCRIBE esta local; la divergencia histórica entre ambas se cierra tras la primera propagación exitosa.

Si una IA llega al repo y NO ve `architecture-guardrails` cargada, está operando contra arquitectura obsoleta.
<!-- /personal-skills:slice:access2web-blueprint -->

## Guía local del repositorio

El índice de skills, las lecturas obligatorias y las reglas duras del repositorio viven en
[`docs/agents-repo-guide.md`](docs/agents-repo-guide.md). El flujo operativo del CI se genera
en [`docs/ci-pattern-flow.md`](docs/ci-pattern-flow.md) y no se edita a mano.
