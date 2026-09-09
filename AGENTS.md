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
| `telefonica-brand-design`, `frontend-design`, `dysflow-*`, `access-*` | externas — no garantizadas |

Regla: una convención obligatoria de este repositorio no puede vivir sólo en una skill personal. Si es obligatoria, su contenido va en `CONTRIBUTING.md`, en `docs/` o en [`skills/`](skills/README.md), versionado dentro del repositorio.

Las entradas marcadas como personales se trasladan o se retiran en las entregas siguientes de #167. El contrato de `skills/` y las instrucciones de instalación para un colaborador nuevo están en [`skills/README.md`](skills/README.md).

## Mandatory reads before contributing

Antes de tocar `app/`, `tests/`, `openspec/`, `docs/architecture.md`, o de proponer una decisión arquitectónica nueva, lea en este orden:

1. [`CODEBASE-GUIDE.md`](CODEBASE-GUIDE.md) — overview + ownership + reading path.
2. [`docs/architecture.md`](docs/architecture.md) — fuente de verdad única de la arquitectura (capas hexagonales, decisiones D-<n> cross-cutting vigentes y obsoletas, patrones transversales).
3. [`docs/calidad-de-codigo-y-ci.md`](docs/calidad-de-codigo-y-ci.md) — gates de calidad (los 12 `check_*.py` + los 5 workflows).
4. [`CONTRIBUTING.md`](CONTRIBUTING.md) — workflow + label system + convention multi-app.
5. Si el cambio pertenece a un OpenSpec change vivo, su `openspec/changes/<change>/design.md`.
6. Si toca una app específica, su `docs/03-aplicaciones/<app>/epic.md`.

Este orden lo operacionaliza la skill **`architecture-guardrails`** (§Hard Rules §1). Saltarse cualquier paso deja a la IA operando contra arquitectura obsoleta.

## Hard rule del CI — ningún CI rojo se integra

`main` tiene branch protection desde el 2026-09-09, tras hacer público el
repositorio en GitHub Free. La protección se aplica también a administradores.

GitHub exige un PR actualizado con `main`, todas las conversaciones resueltas y
estos checks en verde: `quality`, `review-budget`, `pip-audit`, `gitleaks`,
`trivy-config` y `codeql`. No permite force-push ni borrar la rama.

Todo job alcanzable desde un PR público usa un runner hospedado por GitHub. El
gate `check_workflows.py` rechaza etiquetas propias o dinámicas en esa ruta.

`merge-ready` conserva su función informativa. No es un check protegido porque
no agrega los demás jobs y termina correctamente cuando falta una aprobación.

Pasos del revisor antes de mergear:

1. Verifique que los seis checks protegidos pertenecen al último SHA del PR.
2. Compruebe que las conversaciones están resueltas.
3. Si un check falla por una dependencia externa, abra un issue y corrija el
   bloqueo antes de integrar. No eluda la protección.
4. Integre con `--squash` y conserve la rama remota.

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

[← Back to README](README.md) · [Next: DOCS.md →](DOCS.md)
