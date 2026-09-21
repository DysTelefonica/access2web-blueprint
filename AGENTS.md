---
description: access2web-blueprint — monorepo del blueprint de migración Access/VBA→web + MVP Lanzadera Python hexagonal.
globs: *
alwaysApply: true
---

[← Back to README](README.md)

# access2web-blueprint — Agent Skills Index

When working on this project, load the relevant skill(s) BEFORE writing any code or docs.

> **Las skills de este proyecto viven en [`skills/`](skills/README.md), versionadas con el código.** Instálelas con `scripts/install-skills.sh` (o `.ps1` en Windows). Fuera de ahí sólo puede asumirse instalado `gentle-ai` y `engram`.
>
> La disciplina documental (tono, nombres, formato, plantillas, anti-patrones) vive en `skills/documentation-alan-style/`. Su contrato es la única referencia; `CONTRIBUTING.md` y los documentos de `docs/` la aplican.

## How to Use

1. Revise la columna «Trigger» para localizar las skills aplicables a la tarea actual.
2. Cargue la skill leyendo el archivo `SKILL.md` en la ruta indicada (mediante la herramienta `read` o el symlink correspondiente).
3. Aplique todos los patrones y reglas de la skill cargada.
4. Cuando varias skills apliquen a la vez, pueden combinarse.

## Mandatory skills

These skills are **mandatory** — load them before any work in their scope:

| Skill | Mandatory for |
|---|---|
| **`branch-pr`** | **Any** commit, PR creation, or merge to `main`. Conventional commits, PR pequeño y reversible. En el merge, **sin** `--delete-branch`: la rama remota se conserva; lo que se limpia es el worktree local (`git worktree remove <path>`). Vive en `~/.config/opencode/skills/branch-pr/SKILL.md` (distribuida por `gentle-ai`). |
| **`worktree-reorg-per-project`** | **Any** worktree op en este repo: `git worktree add`, `git worktree move`, `git worktree remove`, `git worktree prune`, pre-PR cleanup, fresh project setup. El layout canónico es main en la raíz del proyecto + container hermano `<project>-worktrees\` con linked WTs adentro (v2.0). v1 está deprecated — si el proyecto está en v1, migrar a v2 antes de trabajar. Vive en `~/.config/opencode/skills/worktree-reorg-per-project/SKILL.md`. Author: ardelperal. |
| **`architecture-guardrails`** | **Any** cambio estructural en `app/`, `tests/` u `openspec/`; al proponer una decisión arquitectónica nueva (D-<n>); al llegar al repo por primera vez; al extender o contradecir una D-<n> existente. Operacionaliza el front-door a [`docs/architecture.md`](docs/architecture.md) como single source of truth. Skill hermana de `documentation-alan-style`: esa prescribe cómo se escribe, esta prescribe qué se debe saber antes de escribir. Vive en este repo: `skills/architecture-guardrails/SKILL.md`. Author: ardelperal. |
| **`documentation-alan-style`** | **Any** redacción o revisión de `README.md`, `AGENTS.md`, `DOCS.md`, `CODEBASE-GUIDE.md`, `CONTRIBUTING.md`, `CHANGELOG.md`, `epic.md` o `walkthrough-*.json`. Plantillas en `references/templates/`. Vive en este repo: `skills/documentation-alan-style/SKILL.md`. |
| **`repository-delivery-governance`** | **Any** auditoría o cambio de CI/CD, política de issues o PRs, labels, branch protection, rulesets, permisos de merge, artefactos, despliegue o rollback. Vive en `~/.agents/skills/repository-delivery-governance/SKILL.md`, distribuida por `DysTelefonica/team-skills`. Si falta, deténgase y sincronice el catálogo antes de trabajar. La política específica de este repo prevalece sobre su baseline portable. |
| **`gentle-ai-ai-slop-discipline`** | **Any** implementación de PR (código, tests, docs). Carga la skill **antes** de escribir la primera línea y **antes** de pedir `size:exception` o merged. Las 4 preguntas self-check y las 3 firmas de AI-slop son el gate que el AI pasa sobre su propio diff antes de notificar al humano. Sin este gate, los AI-invented tests, los `object` types como placeholder, los duck-typed `hasattr` checks, y los comentarios que duplican la autoridad existente pasan como código permanente. Vive en `~/.agents/skills/gentle-ai-ai-slop-discipline/SKILL.md` (distribuida por `gentle-ai`). |

## Project-context skills

| Skill | Trigger | Path |
|---|---|---|
| `branch-pr` | Crear / abrir / preparar un PR. | `~/.config/opencode/skills/branch-pr/SKILL.md` (gentle-ai) |
| `work-unit-commits` | Partir un cambio en commits revisables; chained PRs; mantener tests+docs con código. | `~/.config/opencode/skills/work-unit-commits/SKILL.md` (gentle-ai) |
| `issue-creation` | Crear, redactar, triar o aprobar un issue del repo actual. | `~/.config/opencode/skills/issue-creation/SKILL.md` (gentle-ai) |
| `chained-pr` | PRs sobre 400 líneas, stacked PRs, review slices. | `~/.config/opencode/skills/chained-pr/SKILL.md` (gentle-ai) |
| `documentation-alan-style` | Redactar o revisar docs raíz, `epic.md`, `walkthrough-*.json`. | `skills/documentation-alan-style/SKILL.md` (este repo) |
| `architecture-guardrails` | Cambios estructurales en `app/`, `tests/`, `openspec/`; decisiones D-<n>. | `skills/architecture-guardrails/SKILL.md` (este repo) |
| `lanzadera-testing-strategy` | Clasificar test nuevo en Categoría 1–5; revisar capa de un test; decidir mock vs. fake; debate sobre `auth_bypass`. | `skills/lanzadera-testing-strategy/SKILL.md` (este repo) |
| `repository-delivery-governance` | Auditar o cambiar CI/CD, issues, PRs, labels, protección de ramas, permisos de merge, artefactos, despliegue o rollback. | `~/.agents/skills/repository-delivery-governance/SKILL.md` (`DysTelefonica/team-skills`) |

## Cross-cutting skills

| Skill | Trigger | Origen |
|---|---|---|
| `telefonica-brand-design` | Cualquier UI Mistica (design tokens, brand, layout). | Gentleman-Programming/mistica |
| `frontend-design` | Diseño UI distintivo (no AI defaults: cream/serif/terracotta; near-black/acid-green; broadsheet/hairline). | Gentleman-Programming |
| `dysflow-usage` | Cualquier uso de dysflow MCP (analyze_form_ui, map_form_behavior, verify_form_bindings, etc.). | Gentleman-Programming |
| `dysflow-arnes` | Hard rules de dysflow (HR-1..HR-14) y anti-patterns (AP-1..AP-12). | Gentleman-Programming |
| `maintainer-prompt-drafter` | Estructurar prompts para el mantenedor de dysflow u otro tool externo. | Gentleman-Programming |
| `access-vba-tdd` | Tests TDD en Access/VBA. | Gentleman-Programming |
| `access-form-ui-builder` | Perceive → act → verify loop para forms Access. | Gentleman-Programming |
| `gentle-ai-ai-slop-discipline` | Cualquier implementación de PR — código, tests, docs. Pre-submission self-review obligatorio antes de pushear o pedir `size:exception`. | gentle-ai |

> Nota: "Alan canonical examples omit YAML; this template includes it for opencode/Claude/Cursor scope filtering. Drop YAML only for byte-exact match with the AGENTS.md de Alan."

## Hard rule — `CONTRIBUTING.md` es vinculante

Lea [`CONTRIBUTING.md`](CONTRIBUTING.md) **antes** de crear un issue, una rama, un commit o un PR. Sus convenciones no son recomendaciones: el workflow `ci` las comprueba y rechaza el PR cuando no se cumplen.

| Gate | Comprueba | Si se ignora |
|---|---|---|
| `scripts/check_branch_name.py` | `<tipo>/<nº issue>-<kebab-slug>` | Hay que rehacer la rama, y el PR con ella |
| `scripts/check_pr_size.py` | 400 líneas (`additions + deletions`) | Hay que partir el PR |
| `scripts/check_workflows.py` | Actions fijadas por SHA de 40 hex; un `concurrency.group` por job | El workflow se rechaza |

Consecuencias operativas:

1. **El issue va primero.** El nombre de la rama necesita su número, así que crearla antes que el issue obliga a rehacerla.
2. **Dimensione el trabajo al redactar el issue, no al abrir el PR.** Estime el alcance en el propio issue y, si no cabe en 400 líneas, deje escrito ahí el troceado: qué entra en cada PR y en qué orden. Un issue que no dice cómo se entrega delega esa decisión al final, que es justo cuando la salida cómoda es la excepción.
3. **Si no cabe en 400 líneas, hay tres salidas y `size:exception` es la última.** Primero partir por unidad de trabajo; si las partes dependen entre sí, encadenar PRs (cada rama parte de la anterior); y sólo para diffs inevitables, `size:exception` con su `size-exception-reason:`. El orden y el motivo están en `CONTRIBUTING.md`.
4. **Al integrar, conserve la rama remota.** Se integra con `--squash` y se limpia el worktree local con `git worktree remove <ruta>`. El motivo está en `CONTRIBUTING.md`.
5. **Verifique en local antes de empujar.** `make` no está disponible en Git Bash en Windows; invoque los scripts con `python scripts/<nombre>.py`.

### Alcance: este repositorio no es un proyecto Access

El producto de este repositorio es una aplicación web en Python. El material Access que contiene —los `*.accdb` de la raíz, `.dysflow/` y los `.cls` bajo `data/` e `inputs/`— es **entrada de migración en sólo lectura**. No es producto y no se mantiene: se lee para reemplazarlo.

Su sola presencia activa los triggers de las skills de dominio Access (`dysflow-*`, `access-*`, `vba-*`), pensadas para los proyectos legacy. En este repositorio:

- **Aplican** únicamente al leer o analizar ese material de entrada.
- **No aplican** al escribir producto: código, tests, documentación o workflows.

Ninguna convención de Access o VBA gobierna `app/`, `docs/`, `scripts/`, `.github/` ni `openspec/`. Ahí manda `CONTRIBUTING.md` y este documento.

### Skills: qué se puede asumir instalado

`gentle-ai` distribuye 26 skills con su instalación, entre ellas `branch-pr`, `chained-pr`, `work-unit-commits`, `issue-creation` y la suite `sdd-*`. Quien tenga el harness las tiene, así que este documento puede referenciarlas.

El resto es configuración personal y **no está garantizada** para todo el equipo:

| Skill | Origen |
|---|---|
| `branch-pr` | `gentle-ai` — garantizada |
| `worktree-reorg-per-project` | personal — no garantizada; además describe el layout de disco de un desarrollador, no una convención del proyecto |
| `estado-planificacion-update` | personal — no garantizada |
| `deterministic-quality-harness` | personal — no garantizada. La ruta `.opencode/skills/` que se cita más abajo **no existe en este repositorio** |
| `repository-delivery-governance` | `DysTelefonica/team-skills` — obligatoria en su scope; si falta, deténgase y sincronice el catálogo antes de trabajar |
| `telefonica-brand-design`, `frontend-design`, `dysflow-*`, `access-*` | externas — no garantizadas |

Regla: una convención obligatoria de este repositorio no puede vivir sólo en una skill personal. Si es obligatoria, su contenido va en `CONTRIBUTING.md`, en `docs/` o en [`skills/`](skills/README.md), versionado dentro del repositorio. `repository-delivery-governance` aporta el método portable de auditoría; no sustituye esas fuentes locales.

Las entradas marcadas como personales se trasladan o se retiran en las entregas siguientes de #167. El contrato de `skills/` y las instrucciones de instalación para un colaborador nuevo están en [`skills/README.md`](skills/README.md).

## Mandatory reads before contributing

Antes de tocar `app/`, `tests/`, `openspec/`, `docs/architecture.md`, o de proponer una decisión arquitectónica nueva, lea en este orden:

1. [`CODEBASE-GUIDE.md`](CODEBASE-GUIDE.md) — overview + ownership + reading path.
2. [`docs/architecture.md`](docs/architecture.md) — fuente de verdad única de la arquitectura (capas hexagonales, decisiones D-<n> cross-cutting vigentes y obsoletas, patrones transversales).
3. [`docs/calidad-de-codigo-y-ci.md`](docs/calidad-de-codigo-y-ci.md) — gates de calidad (los 14 `check_*.py` + los 5 workflows).
4. [`CONTRIBUTING.md`](CONTRIBUTING.md) — workflow + label system + convention multi-app.
5. Si el cambio pertenece a un OpenSpec change vivo, su `openspec/changes/<change>/design.md`.
6. Si toca una app específica, su `docs/03-aplicaciones/<app>/epic.md`.

Este orden lo operacionaliza la skill **`architecture-guardrails`** (§Hard Rules §1). Saltarse cualquier paso deja a la IA operando contra arquitectura obsoleta.

## Hard rule del CI — ningún CI rojo se integra

`main` tiene branch protection desde el 2026-09-09, tras hacer público el
repositorio en GitHub Free. La protección se aplica también a administradores.

GitHub exige un PR actualizado con `main`, todas las conversaciones resueltas
y (tras issue #702) un único check en verde: `required`. Ese job agrega
(`needs:`) `review-budget`, `quality`, `security` y `codeql` — los dos
últimos invocan `security.yml`/`codeql.yml` como jobs `uses:` porque
`needs:` no cruza archivos de workflow — y falla cerrado ante cualquier job
faltante, en `failure`, o en un `skipped` no permitido para ese evento
(`scripts/check_required_jobs.py`). No permite force-push ni borrar la rama.

Todo job alcanzable desde un PR público usa un runner hospedado por GitHub. El
gate `check_workflows.py` rechaza etiquetas propias o dinámicas en esa ruta;
un job que solo invoca un reusable workflow (`uses:`) no declara `runs-on` y
está exento de ese check puntual — el reusable workflow invocado sigue esa
misma regla cuando se lo analiza por su cuenta.

`merge-ready` fue retirado (issue #702): quedaba completamente superado por
`required`, que sí bloquea el merge.

Pasos del revisor antes de mergear:

1. Verifique que `required` pertenece al último SHA del PR.
2. Compruebe que las conversaciones están resueltas.
3. Si un check falla por una dependencia externa, abra un issue y corrija el
   bloqueo antes de integrar. No eluda la protección.
4. Integre con merge commit (no squash, práctica real del historial de `main`) y conserve la rama remota.

Refuerza esta disciplina con `gentle-ai review status --cwd <repo>` antes de mergear.

## Refuerzo

Si una IA escribe código o documentos en este repositorio sin cargar las skills pertinentes, trabaja a ciegas. Las skills son la single source of truth para «cómo se hace X en este proyecto». Cargue primero, escriba después.

### Ciclo de refactor de una app (estado-planificacion-update)

La skill `estado-planificacion-update` cubre el ciclo completo de vida de un refactor y exige mantener el HTML vivo en cada paso. Reglas mínimas:

1. **Estudio inicial** (cargar `access-vba-capability-docs`): tras el primer walkthrough de forms, abrir `estado-planificacion-NEW_<fecha>.html` con la tabla de apps en estudio, RAG 🟢 / 🟡 / 🔴 por app, y los gaps identificados.
2. **Propuesta OpenSpec** (`/sdd-new`): cuando arranca el change, añadir al HTML la sección "Próximos pasos" con la propuesta, los PRs iniciales y el chain strategy decidido.
3. **Durante la planificación** (`sdd-spec`, `sdd-design`, `sdd-tasks`): cada cierre de fase refresca el HTML con la sección correspondiente, los gaps cerrados y los que quedan abiertos, el forecast de carga de revisión y el plan de PRs.
4. **Ejecución** (`sdd-apply`): cada PR cerrado actualiza el Gantt, la columna RAG de la tabla de PRs y el motivo del cambio. La cadena visual stack → PRs → plan debe quedar coherente.
5. **Verificación y archivo** (`sdd-verify`, `sdd-archive`): el HTML final muestra todos los PRs en verde, los gaps cerrados y el cierre del ciclo. Se commitea junto al último PR.

El archivo `estado-planificacion-*.html` es la single source of truth visual del avance para los jefes. Cuando un agente entra al repo, debe leer el HTML más reciente antes de proponer cambios.

## Hard rule — Anti-AI-slop gate antes de cualquier PR

Toda implementación de un PR (código, tests, docs, configuración) pasa por el gate de la skill `gentle-ai-ai-slop-discipline` antes de pushear la rama o pedir `size:exception`. La secuencia es no negociable:

1. **Cargar la skill primero.** Antes de la primera línea del diff, leer `~/.agents/skills/gentle-ai-ai-slop-discipline/SKILL.md` en su versión actual. La skill lleva changelog; saltarse la versión vigente es saltarse el gate.
2. **Pasar las 4 preguntas self-check.** Para cada archivo nuevo o modificado, responder: ¿está en los acceptance scenarios del issue? ¿el humano puede defender cada línea como decisión propia? ¿duplica autoridad existente? ¿se va de scope? Si cualquier respuesta es «no» o «no sé», STOP. Reescribir antes de continuar.
3. **Detectar las 3 firmas.** Prose paralelo que duplica autoridad. Claims amplios sin contrato testeable. Tests que verifican objetos AI-invented en lugar del dominio real. Cada firma detectada es deuda permanente: remover antes de commitear.
4. **Pre-submission self-review.** Releer el diff completo. Para cada línea fuera del scope del issue: ¿por qué está aquí? Si no hay respuesta, va fuera.
5. **AI-slop self-defense con el humano.** Antes de pushear, presentar al humano 1-3 líneas random del diff y preguntar: «¿podés defender esta línea como decisión tuya, no de la IA?». Si no, la línea se elimina.

**Excepción al gate**: edición trivial puramente pasiva (typo fix, formato, link roto). Documentar en el commit por qué no aplica el gate.

**`size:exception` con AI-slop no negociable**: pedir `size:exception` sin haber pasado este gate es rechazo automático. El diff grande no excusa tests AI-invented.

[← Back to README](README.md) · [Next: DOCS.md →](DOCS.md)

<!-- personal-skills:slice:access2web-blueprint @ v93d973f -->
# slices/partials/web.md

## Manera de trabajar en proyectos web

> Aplica a todo `primary_type: web` del catálogo. Las invariantes de ciclo de vida de PR aquí enunciadas se complementan con las skills universalmente activas — véase `personal-skills/AGENTS.md` raíz para el sistema de propagación, `propagate-team-skills.ps1` para la mecánica de distribución, y el bloque de partials específicos del consumer para las convenciones del proyecto concreto.
>
> Este partial enuncia invariantes. Los procedimientos asociados viven en sus skills respectivas — no se duplican aquí.

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

- El **worktree local** se elimina tras el merge, vía `git worktree remove <path>` + `git worktree prune`. Mecánica detallada en `worktree-reorg-per-project` Phase 3.
- La **rama remota** se conserva tras el merge. Nunca `git push origin --delete <rama>`. La granularidad por unidad de trabajo se preserva precisamente porque las ramas quedan referenciables desde el historial de PRs.
- La estrategia de merge (`--squash` o `--no-ff`) es decisión del consumer; ambas son válidas. La **invariante** es que la rama remota sobreviva al merge, no la forma concreta del commit en `main`.

### Disciplina de revisión

- El CI debe estar **verde contra la base actual** antes de pedir revisión. Si la rama base avanzó durante la vida del PR, **rebase + rerun del CI** antes de declarar mergeable. El verde contra una base obsoleta es stale-green y corrompe el merge.
- **Rojo en CI pisa todo el merge.** Regla humana: el revisor no debe pulsar merge con ningún check rojo, ni siquiera si el rojo parece trivial. Complemento técnico: `repository-delivery-governance` HR-7 + `deterministic-quality-harness` HR-1 fail-loud atajan el escenario cuando hay branch protection automatizada.
- Donde GitHub Team no está disponible, el consumer replica la barrera con un job `merge-ready` signal-only (ver `access2web-blueprint/ci.yml` como referencia portable) que exit-non-zero si `gh pr view mergeable != true` o `reviewDecision != APPROVED`.
- Cuando el CI rojo es por **infra** (runner colgado, red, secret rotado), abrir issue bloqueante de CI y enlazarla desde el PR; no embutir la fix infra en el PR del feature salvo que sea ≤30 LOC y se cierre en el día.

### Anti-slop y atribución

- Anti-slop y anti-sobreingeniería de IA se delegan a la skill T1 upstream `gentle-ai-ai-slop-discipline` (4 preguntas, 3 firmas, scope boundary guard para subagentes). El partial no redefine las firmas — el consumer que adopte la skill las aplica automáticamente al revisar PRs.
- **Sin atribución de IA en commits**: no se añade `Co-Authored-By: ... <AI>` ni equivalente. Esta regla vive también en `personal-skills/AGENTS.md` raíz de la flota; el partial la refleja para que sea visible en el slice de web.
- Mensajes de commit en conventional commits. Castellano peninsular formal (usted) en artefactos documentales raíz; inglés en código, comentarios, mensajes de commit y PR bodies.

### CodeGraph preflight (cuando aplique)

- Si el consumer tiene índice CodeGraph (`.codegraph/` presente), el workflow sigue `engineering-workflow` líneas 64-71: `codegraph init` antes del primer edit, `codegraph_explore` antes de cualquier grep/read/glob amplio. No se reinventa aquí.
- Si el consumer no soporta CodeGraph, el preflight se omite sin romper invariante — la regla es "usar CodeGraph cuando esté disponible", no "requerir CodeGraph siempre".

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
- Rama remota preservada, WT local delete post-merge → `worktree-reorg-per-project` Phase 3.
- Anti-slop → upstream `gentle-ai-ai-slop-discipline` (T1 universal).
- CodeGraph preflight → upstream `engineering-workflow` líneas 64-71; `codegraph-usage` HR-1, HR-2.
- Conventional commits + castellano peninsular en artefactos + sin atribución IA → `personal-skills/AGENTS.md` raíz de flota.

### Cómo auditar este partial usted mismo

Procedimiento de validación periódica (mensual o por release de skill fuente):

- Confirmar que las HRs citadas en §Procedencia siguen existiendo con la misma numeración y redacción en el cuerpo actual de cada skill. Si una skill referenciada cambia su HR, este partial requiere reauditoría.
- Confirmar que no se haya añadido regla con cuerpo procedural en este partial — los procedimientos viven en skills, no aquí.
- Confirmar que la sección §Divergencias documentadas sigue reflejando las variantes reales de los consumers actuales.

### Antipatrones

- "Esperar a que CI esté verde para mergear" sin rebasear contra la base actual — el verde contra base obsoleta es stale-green.
- "Borrar la rama remota post-merge porque ya está mergeada" — destruye la granularidad por unidad de trabajo que el flujo pretende crear.
- "PR con 600 líneas porque el feature lo requiere" — partir primero, encadenar después; `size:exception` es el último recurso, no la primera opción.
- "Mergear con CI rojo aunque el rojo parezca trivial" — la trivialidad la decide el revisor, no el autor.
- "Esperar a que el reviewer apruebe manualmente aunque todos los checks estén verdes" en proyectos con auto-merge standing explícito — revisar la sección de revocación de `merge-workflow.md §15.6` antes de saltarse el gate.

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
