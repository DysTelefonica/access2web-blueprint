# access2web-blueprint

<!-- personal-skills:slice:access2web-blueprint @ v1844dd2 -->
# slices/partials/web.md

## Manera de trabajar en proyectos web

> Aplica a todo `primary_type: web` del catálogo. Las invariantes de ciclo de vida de PR aquí enunciadas se complementan con las skills universalmente activas — véase `personal-skills/AGENTS.md` raíz para el sistema de propagación, `propagate-team-skills.ps1` para la mecánica de distribución, y el bloque de partials específicos del consumer para las convenciones del proyecto concreto.
>
> Este partial enuncia invariantes. Los procedimientos asociados viven en sus skills respectivas — no se duplican aquí.

### Arranque (invariante)

- **Ningún trabajo empieza sin sincronizar la rama base.** Antes de leer, decidir o tocar un fichero: `git fetch origin && git switch <active_branch> && git pull --ff-only`. Un checkout atrasado muestra gobernanza, CI y skills que **ya no rigen**.
- **Señal de alarma, no conclusión.** Si un documento o fichero citado por la gobernanza del consumer **parece no existir**, o **contradice** el modelo vigente, la lectura correcta es "estoy en un checkout atrasado": sincronizar y volver a mirar. Nunca "el repo tiene un hueco" ni "la regla cambió sin avisar".
- **Gobernanza viva, no copiada.** Las issues se leen del tracker (`gh issue view <n>`), nunca de una copia pegada; el `AGENTS.md` y el documento de flujo del consumer se leen **en el SHA sincronizado**.
- **El workspace del runner de CI no es un clon de desarrollo.** Se trabaja en un worktree o clon propio (ver §Worktree y rama remota).
- **Una copia global de skill desactualizada tapa la del repo** (ver §Carga de skills): si una skill del repo contradice a la global, gana la del repo, y la global está pidiendo actualización.

### Skill Registry Protocol (MANDATORY)

La tabla de skills del `AGENTS.md` del consumer es un subconjunto curado; el registro generado — `.atl/skill-registry.md`, refrescado por la skill `skill-registry` (`gentle-ai skill-registry refresh`) — es la fuente de verdad de descubrimiento. Una regla que existe y está indexada pero que ninguna sesión carga no gobierna nada: este protocolo convierte el índice en despacho.

1. **READ** — antes de escribir o revisar código, CI o documentación, lea `.atl/skill-registry.md` (o el índice de skills del repo).
2. **MATCH** — contraste la tarea contra la columna `Trigger / description` de las skills candidatas.
3. **LOAD** — lea íntegros los `SKILL.md` que matcheen ANTES de actuar.
4. **DELEGATE** — al delegar a un subagente, pase las rutas exactas de los `SKILL.md` bajo un encabezado `## Skills to load before work` en la tarea. Un nombre en prosa no basta: el subagente no debe tener que adivinar ni redescubrir (eso produce el degradado `fallback-registry`).
5. **GAP** — si ninguna skill matchea y la tarea toca CI/PRs, releases, la flota de runners o la adopción de un repositorio nuevo, es un gap de orquestación: dígalo explícitamente en vez de proceder de memoria.

- **Deriva índice/registro.** La tabla del `AGENTS.md` es la puerta del contribuyente; el registro es la verdad operativa. Que una skill no aparezca en la tabla no autoriza concluir «no aplica»: autoriza buscarla en el registro (`skill-registry` / `gentle-ai skill-registry refresh`).
- **Frescura del registro.** Si `.atl/skill-registry.md` es anterior al último cambio de skills, está stale: refresque el registro antes de delegar.

### Forma del ciclo

- Toda issue es **atómica**: la cambia una persona, la cierra un PR (o varias si encadenadas vía `chained-pr` cuando la diff supera el presupuesto).
- Toda issue aprobada tiene una **rama propia** con el nombre `<tipo>/<nº issue>-<kebab-slug>`, validado por `scripts/check_branch_name.py` del consumer o equivalente.
- Toda rama se desarrolla en un **worktree dedicado** bajo el layout canónico de `worktree-reorg-per-project` v2.0 (sibling-container `<project>-worktrees/<wt-name>/`).
- Toda PR apunta a la `active_branch` declarada en `fleet/registry.json` para el ciclo activo. Pre-MVP single-branch implica `main` por defecto; el flip a `staging` post-MVP sigue la llave de vocabulario documentada en `intake-roadmap-loop` HR-4/HR-8.

### Presupuesto de revisión

- Toda PR se mantiene bajo el **presupuesto de revisión de 400 líneas** (`additions + deletions`), comprobado por `scripts/check_pr_size.py` o equivalente.
- 400 líneas es **techo de revisión, no techo de tamaño**: a partir de esa cifra la revisión pasa de atenta a vistazo. La justificación completa vive en el `CONTRIBUTING.md` de cada consumer; este partial la enuncia sin duplicar.
- La excepción `size:exception` requiere, **obligatoriamente**, en el cuerpo del PR: `size-exception-reason: <por qué>` más un enlace a la evidencia que justifique la superación. La etiqueta `size:exception` queda como mecanismo opcional por consumer (útil para detección CI automática; no es regla invariante).
- Cuando la diff supera el presupuesto, el orden de escape es: (1) partir por unidad de trabajo, (2) encadenar PRs vía `chained-pr`, (3) `size:exception` como último recurso. Si la excepción se vuelve habitual, el problema está en el troceado del issue, no en el presupuesto.

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
4. **Merge y limpieza.** El merge lo lanza un mantenedor autorizado, nunca el
   agente. Solo después del merge verificado y sin cambios pendientes, retirar el
   **worktree local** (`git worktree remove <path>` + `git worktree prune`,
   mecánica en `worktree-reorg-per-project` Phase 3). La **rama remota se
   conserva siempre**: nunca `git push origin --delete <rama>` ni
   `--delete-branch`.

El procedimiento detallado —comandos exactos y condiciones de «hecho»— vive en
la skill de día a día del consumer y en su documento operativo generado; este
partial solo enuncia la invariante y no lo duplica.

- La estrategia de merge (`--squash` o `--no-ff`) es decisión del consumer; ambas son válidas. La **invariante** es que la rama remota sobreviva al merge, no la forma concreta del commit en `main`.

### Disciplina de revisión

- El CI debe estar **verde contra la base actual** antes de pedir revisión. Si la rama base avanzó durante la vida del PR, **integre la base remota en la rama de feature con `git merge`** (nunca rebase) y **reejecute** el preflight y el CI antes de declarar mergeable. El verde contra una base obsoleta es stale-green y corrompe el merge.
- **Rojo en CI pisa todo el merge.** Regla humana: el revisor no debe pulsar merge con ningún check rojo, ni siquiera si el rojo parece trivial. Complemento técnico: `repository-delivery-governance` HR-7 + `deterministic-quality-harness` HR-1 fail-loud atajan el escenario cuando hay branch protection automatizada.
- Donde GitHub Team no está disponible, el consumer replica la barrera con un job `merge-ready` signal-only (ver `access2web-blueprint/ci.yml` como referencia portable) que exit-non-zero si `gh pr view mergeable != true` o `reviewDecision != APPROVED`.
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
- **Etiqueta `size:exception`.** Opcional por consumer. Invariante compartida: el `size-exception-reason:` en el cuerpo del PR es obligatorio.
- **Pre-MVP vs post-MVP base branch.** Mientras no haya flip explícito del usuario con la llave de vocabulario documentada en `intake-roadmap-loop` HR-4/HR-8, todo aterriza en `main`.
- **Convención multi-app.** Si el consumer migra varias apps, el prefijo de issue/commit es decisión propia; el commit debe identificar el scope de cualquier manera.

### Procedencia (skills que alimentan este partial)

Cada invariante de este partial se ancla a una skill específica del catálogo o upstream. Si la skill referenciada cambia su HR, este partial requiere reauditoría.

- Issue-first atómica → `intake-roadmap-loop` HR-1, HR-14; upstream `engineering-workflow` Step 2.
- Rama `<tipo>/<nº>-<slug>` → `repository-delivery-governance` HR-4 (documentado/CI-enforced).
- Un worktree por issue → `worktree-reorg-per-project` v2.0 (layout + Phase 3 cleanup).
- Base pre-MVP = main → upstream `engineering-workflow` líneas 46-51; flip post-MVP vía `intake-roadmap-loop` HR-4/HR-8.
- 400 líneas + `size:exception` → `deterministic-quality-harness` Decision Gate línea 56; upstream `chained-pr` HR-1 línea 16.
- Rojo en CI pisa todo → `repository-delivery-governance` HR-7; upstream `engineering-workflow` línea 79.
- Ciclo de entrega (worktree y rama desde la base activa actualizada → merge de la base en la feature → preflight/pruebas → push → PR → CI sobre cabeza y base actuales → merge del mantenedor → retirada solo del worktree local) → `governed-delivery` regla 1 y el documento operativo generado por `ci-pattern` (#390).
- Sincronización de la rama publicada por `git merge`, nunca rebase → `governed-delivery` regla 1; `ci-pattern` HR-54; #333.
- Rama remota preservada tras el merge, worktree local retirado solo al final → `worktree-reorg-per-project` Phase 3.
- Anti-slop → `ai-slop-discipline` (catálogo, T1 universal).
- CodeGraph preflight → upstream `engineering-workflow` líneas 64-71; `codegraph-usage` HR-1, HR-2.
- Conventional commits + castellano peninsular en artefactos + sin atribución IA → `personal-skills/AGENTS.md` raíz de flota.
- Skill Registry Protocol (READ/MATCH/LOAD/DELEGATE/GAP + deriva índice/registro + frescura) → upstream `gentle-ai` `internal/assets/skills/_shared/odd-orchestrator-sections.md:183-201`, adaptado vía `gentle-ai-governance-dispatch.md` (adopciones 1-3 del informe de gobernanza 2026-10-02).
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
- "PR con 600 líneas porque el feature lo requiere" — partir primero, encadenar después; `size:exception` es el último recurso, no la primera opción.
- "Mergear con CI rojo aunque el rojo parezca trivial" — la trivialidad la decide el revisor, no el autor.
- "Esperar a que el reviewer apruebe manualmente aunque todos los checks estén verdes" en proyectos con auto-merge standing explícito — revisar la sección de revocación de `merge-workflow.md §15.6` antes de saltarse el gate.

### Carga de skills (orden de resolución)

Al buscar una skill de este catálogo, resuelva en este orden:

1. Global: `~/.agents/skills/` (y `~/.pi/agent/skills/`).
2. Repo: `.agents/skills/` del consumer, materializada desde `.team-skills.yaml` por `scripts/refresh-team-skills.ps1` (vía rama `skill-fleet/<name>`).
3. Catálogo: `skills/<nombre>/SKILL.md` de personal-skills.

Referencias por nombre en prosa; ruta absoluta exacta en la delegación: al pasar una tarea a un subagente, copie la ruta del `SKILL.md` desde el registro bajo `## Skills to load before work` (ver §Skill Registry Protocol). La regla anti-ruta-absoluta gobierna el texto de skills y docs — las rutas varían por máquina — no la delegación, donde la ruta exacta es obligatoria. Nota de precedencia: pi nativo conserva la primera copia encontrada (gana la global), mientras que el registry de gentle-pi prefiere la copia del repo; una copia global desactualizada tapa la del repo en pi, así que las copias globales deben seguir al catálogo. Si una skill citada en este partial no resuelve en ninguno de los tres niveles, repórtelo: falta declararla en `.team-skills.yaml` del consumer.



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
lo lanza un mantenedor por el camino gobernado y nunca un agente; sin
nombres propios — solo roles y rutas instaladas. El detective post-push y
el guard de force-push abren el incidente (HR-49, HR-51).

<!-- ci-pattern:hr-block:begin -->
| HR | Resumen | Clase |
| --- | --- | --- |
| HR-48 | Toda la documentación de un proyecto vive en su repositorio, con rutas relativas | gate |
| HR-49 | Toda rama protegida por contrato se audita tras el push y su violación abre issue de incidente | gate |
| HR-51 | Una rama protegida por contrato no se reescribe ni se borra: la violación abre incidente | gate |
| HR-54 | La cadena de PRs se declara como dato y el enlace de la issue no admite mezclas | gate |
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

Toda la documentación de este proyecto vive en ESTE repositorio, bajo `docs/`
por materia; los post-mortems en `docs/postmortems/<AAAA-MM-DD>-<slug>.md`.
Fuera del repositorio, y sin copia en él, quedan solo los secretos, los
backups, los datos en bruto y el material privado. Nunca ubique documentación
en una ruta absoluta o externa al repositorio (HR-48): el gate de la fase 5 la
rechaza. Detalle: la sección «Dónde vive la documentación» del flujo operativo.
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

`scripts/check_layers.py` rechaza imports que violen `ROOT_PACKAGE = "app.src.modules"` con `ALLOWED_IMPORTS` y `PURE_LAYERS = {domain, ports, application}` (DA-1). Mover un adapter a la capa equivocada es un gate failure, no un estilo. Esquivar el gate con `|| true` está prohibido por `tests/test_ci_workflow.py`.

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
