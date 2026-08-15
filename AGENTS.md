# access2web-blueprint — Agent Skills Index

When working on this project, load the relevant skill(s) BEFORE writing any code or docs.

> **Las skills de este proyecto viven en [`skills/`](skills/README.md), versionadas
> con el código.** Instálelas con `scripts/install-skills.sh` (o `.ps1` en Windows).
> Fuera de ahí sólo puede asumirse instalado `gentle-ai` y `engram`.
>
> La disciplina documental (tono, nombres, formato, plantillas, anti-patrones) se
> traslada a `skills/documentation-alan-style/` en las entregas siguientes de #167.
> Hasta entonces, la referencia es `CONTRIBUTING.md` y los documentos de `docs/`.

## Cómo usar

1. Revise la columna «trigger» para localizar las skills aplicables a la tarea actual.
2. Cargue la skill leyendo el archivo `SKILL.md` en la ruta indicada (mediante la herramienta `read` o el symlink correspondiente).
3. Aplique todos los patrones y reglas de la skill cargada.
4. Cuando varias skills apliquen a la vez, pueden combinarse.

## Hard rule — `CONTRIBUTING.md` es vinculante

Lea [`CONTRIBUTING.md`](CONTRIBUTING.md) **antes** de crear un issue, una rama,
un commit o un PR. Sus convenciones no son recomendaciones: el workflow `ci` las
comprueba y rechaza el PR cuando no se cumplen.

| Gate | Comprueba | Si se ignora |
|---|---|---|
| `scripts/check_branch_name.py` | `<tipo>/<nº issue>-<kebab-slug>` | Hay que rehacer la rama, y el PR con ella |
| `scripts/check_pr_size.py` | 400 líneas (`additions + deletions`) | Hay que partir el PR |
| `scripts/check_workflows.py` | Actions fijadas por SHA de 40 hex; un `concurrency.group` por job | El workflow se rechaza |

Consecuencias operativas:

1. **El issue va primero.** El nombre de la rama necesita su número, así que
   crearla antes que el issue obliga a rehacerla.
2. **Dimensione el trabajo al redactar el issue, no al abrir el PR.** Estime el
   alcance en el propio issue y, si no cabe en 400 líneas, deje escrito ahí el
   troceado: qué entra en cada PR y en qué orden. Un issue que no dice cómo se
   entrega delega esa decisión al final, que es justo cuando la salida cómoda es
   la excepción.
3. **Si no cabe en 400 líneas, hay tres salidas y `size:exception` es la
   última.** Primero partir por unidad de trabajo; si las partes dependen entre
   sí, encadenar PRs (cada rama parte de la anterior); y sólo para diffs
   inevitables, `size:exception` con su `size-exception-reason:`. El orden y el
   motivo están en `CONTRIBUTING.md`.
4. **Al integrar, conserve la rama remota.** Se integra con `--squash` y se
   limpia el worktree local con `git worktree remove <ruta>`. El motivo está en
   `CONTRIBUTING.md`.
5. **Verifique en local antes de empujar.** `make` no está disponible en Git
   Bash en Windows; invoque los scripts con `python scripts/<nombre>.py`.

### Alcance: este repositorio no es un proyecto Access

El producto de este repositorio es una aplicación web en Python. El material
Access que contiene —los `*.accdb` de la raíz, `.dysflow/` y los `.cls` bajo
`data/` e `inputs/`— es **entrada de migración en sólo lectura**. No es producto
y no se mantiene: se lee para reemplazarlo.

Su sola presencia activa los triggers de las skills de dominio Access
(`dysflow-*`, `access-*`, `vba-*`), pensadas para los proyectos legacy. En este
repositorio:

- **Aplican** únicamente al leer o analizar ese material de entrada.
- **No aplican** al escribir producto: código, tests, documentación o workflows.

Ninguna convención de Access o VBA gobierna `app/`, `docs/`, `scripts/`,
`.github/` ni `openspec/`. Ahí manda `CONTRIBUTING.md` y este documento.

### Skills: qué se puede asumir instalado

`gentle-ai` distribuye 26 skills con su instalación, entre ellas `branch-pr`,
`chained-pr`, `work-unit-commits`, `issue-creation` y la suite `sdd-*`. Quien
tenga el harness las tiene, así que este documento puede referenciarlas.

El resto es configuración personal y **no está garantizada** para todo el equipo:

| Skill | Origen |
|---|---|
| `branch-pr` | `gentle-ai` — garantizada |
| `worktree-reorg-per-project` | personal — no garantizada; además describe el layout de disco de un desarrollador, no una convención del proyecto |
| `estado-planificacion-update` | personal — no garantizada |
| `deterministic-quality-harness` | personal — no garantizada. La ruta `.opencode/skills/` que se cita más abajo **no existe en este repositorio** |
| `telefonica-brand-design`, `frontend-design`, `dysflow-*`, `access-*` | externas — no garantizadas |

Regla: una convención obligatoria de este repositorio no puede vivir sólo en una
skill personal. Si es obligatoria, su contenido va en `CONTRIBUTING.md`, en
`docs/` o en [`skills/`](skills/README.md), versionado dentro del repositorio.

Las entradas marcadas como personales se trasladan o se retiran en las entregas
siguientes de #167. El contrato de `skills/` y las instrucciones de instalación
para un colaborador nuevo están en [`skills/README.md`](skills/README.md).

## Mandatory reads before contributing

Antes de tocar `app/`, `tests/`, `openspec/`, `docs/architecture.md`, o de proponer una decisión arquitectónica nueva, lea en este orden:

1. [`CODEBASE-GUIDE.md`](CODEBASE-GUIDE.md) — overview + ownership + reading path.
2. [`docs/architecture.md`](docs/architecture.md) — fuente de verdad única de la arquitectura (capas hexagonales, decisiones D-<n> cross-cutting vigentes y obsoletas, patrones transversales).
3. [`docs/calidad-de-codigo-y-ci.md`](docs/calidad-de-codigo-y-ci.md) — gates de calidad (los 12 `check_*.py` + los 4 workflows).
4. [`CONTRIBUTING.md`](CONTRIBUTING.md) — workflow + label system + convention multi-app.
5. Si el cambio pertenece a un OpenSpec change vivo, su `openspec/changes/<change>/design.md`.
6. Si toca una app específica, su `docs/03-aplicaciones/<app>/epic.md`.

Este orden lo operacionaliza la skill **`architecture-guardrails`** (§Hard Rules §1). Saltarse cualquier paso deja a la IA operando contra arquitectura obsoleta.

## Skills obligatorias

These skills are **mandatory** — load them before any work in their scope:

| Skill | Obligatorio para |
|---|---|
| **`branch-pr`** | Any commit, PR creation, or merge to `main`. Conventional commits, PR pequeño y reversible. En el merge, **sin** `--delete-branch`: la rama remota se conserva; lo que se limpia es el worktree local (`git worktree remove <path>`). |
| **`worktree-reorg-per-project`** | Any worktree op in `c:\00repos\codigo\`: `git worktree add`, `git worktree move`, `git worktree remove`, `git worktree prune`, pre-PR cleanup, fresh project setup, detected scattered WT or v1 layout (`00_main/` + `.git` FILE). El layout canónico es main en la raíz del proyecto + container hermano `<project>-worktrees\` con linked WTs adentro (v2.0). v1 está deprecated — si el proyecto está en v1, migrar a v2 antes de trabajar. La skill vive en `~/.config/opencode/skills/worktree-reorg-per-project/SKILL.md`. Author: ardelperal. |
| **`estado-planificacion-update`** | Cualquier ciclo de refactor de una app del blueprint, desde el estudio inicial (`access-vba-capability-docs`, walkthrough de forms, propuesta OpenSpec) hasta la entrega (cierre de PRs, archive del change). El ciclo se actualiza con un `estado-planificacion-NEW_YYYY-MM-DD.html` cada vez que se cierra una fase, un PR o un gap. La skill trae `scripts/render_estado.py` (scaffold desde `tasks.md` + JSON config), `assets/style-base.css` y `assets/template.html`. Mantener el archivo vivo es la single source of truth visual del avance para los jefes. Vive en `~/.config/opencode/skills/estado-planificacion-update/SKILL.md`. Author: ardelperal. |
| **`architecture-guardrails`** | Antes de implementar cambios estructurales en `app/`, `tests/` o `openspec/`; al proponer una decisión arquitectónica nueva (D-<n>); al llegar al repo por primera vez; al extender o contradecir una D-<n> existente. Operacionaliza el front-door a [`docs/architecture.md`](docs/architecture.md) como single source of truth. Skill hermana de `documentation-alan-style`: esa prescribe cómo se escribe, esta prescribe qué se debe saber antes de escribir. Vive en este repo: `skills/architecture-guardrails/SKILL.md`. Author: ardelperal. |

## Skills cross-cutting (provienen de otros repos)

| Skill | Trigger | Origen |
|---|---|---|
| `telefonica-brand-design` | Cualquier UI Mistica (design tokens, brand, layout). | Gentleman-Programming/mistica |
| `frontend-design` | Diseño UI distintivo (no AI defaults: cream/serif/terracotta; near-black/acid-green; broadsheet/hairline). | Gentleman-Programming |
| `dysflow-usage` | Cualquier uso de dysflow MCP (analyze_form_ui, map_form_behavior, verify_form_bindings, etc.). | Gentleman-Programming |
| `dysflow-arnes` | Hard rules de dysflow (HR-1..HR-14) y anti-patterns (AP-1..AP-12). | Gentleman-Programming |
| `maintainer-prompt-drafter` | Estructurar prompts para el mantenedor de dysflow u otro tool externo. | Gentleman-Programming |
| `access-vba-tdd` | Tests TDD en Access/VBA. | Gentleman-Programming |
| `access-form-ui-builder` | Perceive → act → verify loop para forms Access. | Gentleman-Programming |
| `deterministic-quality-harness` | Wiring o auditoría de quality gates de plataforma (CI, lint, ratchet baseline, complexity, hexagonal layer gate, secret/dependency scan). Inspirada en `unclebob/swarm-forge` y validada por APAP_WEB. Definida en `.opencode/skills/deterministic-quality-harness/SKILL.md`, symlinkeada desde `~/.config/opencode/skills/`. Author: ardelperal. | este repo |

## Hard rule del CI — ningún CI rojo se mergea

Este repositorio **no tiene branch protection automatizada** en `main` (decisión del 2026-08-10: GitHub Team cuesta $4/user/mes, no aprobado todavía). El CI workflow (`deterministic-quality-harness` v1.4, orquestado en `.github/workflows/ci.yml`) corre en cada PR y los status checks aparecen en línea, pero **no bloquean el merge**.

**Por convención del equipo: el revisor NO debe mergear un PR con CI rojo.** Esto es una regla humana, no técnica. Para reactivarla como automatismo en el futuro, migrar el repo a público (branch protection gratis en Free) o upgrade a GitHub Team.

Pasos del revisor antes de mergear:

1. Verificar que el check `ci / quality` está verde en el PR.
2. Verificar que el check `ci / review-budget` está verde (si el PR es un pull_request).
3. Si el check `ci / security` falla **solo** por vulnerabilidades en paquetes externos no relacionados con el cambio, abrir issue y mergear con un follow-up PR.
4. Si **cualquier** otro check falla, pedir al autor que arregle antes de mergear.

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
