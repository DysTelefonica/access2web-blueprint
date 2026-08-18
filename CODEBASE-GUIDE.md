---
description: CODEBASE-GUIDE — ownership, flows, guardrails (90-second mental model) para mantenedores del monorepo access2web-blueprint.
globs: *
alwaysApply: true
---

[← Back to README](../README.md)

# access2web-blueprint — Codebase Guide

**This guide is for maintainers and contributors who need to understand how blueprint + platform responsibilities coexist in this monorepo, where research, code and documentation live, and where to add new content when the platform grows.**

access2web-blueprint es **monorepo de la plataforma web hexagonal + blueprint del refactor** de las 8 apps legacy Access/VBA. Aloja `docs/` (research + decisiones + épicas) y `app/` (código FastAPI + HTMX + Alembic del MVP de Lanzadera, 2026-08; renombrado desde `platform/` el 2026-08-09 por shadowing del módulo stdlib — ver `app/README.md`). El legacy `.accdb` queda en repos separados hasta el cut-over del ecosistema completo.

> **Sentence that organizes the whole repo**: "Lanzadera es la madre: ahí nacen usuarios, aplicativos y permisos. Las otras 7 apps son consumidoras."

> **Scope del scope**: "Este repo es monorepo de plataforma + blueprint. Desde el MVP de Lanzadera (2026-08), `app/` contiene el código web hexagonal; `docs/` contiene el research y las decisiones; `data/staging/` contiene los binarios legacy extraídos."

## 90-second mental model

```text
Legacy apps (Access/VBA)         Este monorepo                          Apps web (app/src/modules/<app>/)
┌────────────────────┐         ┌──────────────────────────┐         ┌──────────────────────────────┐
│ NoConformidades    │         │   docs/                   │         │  lanzadera/         (MVP)     │
│ Gestion_Riesgos     │         │     03-aplicaciones/      │         │  expedientes/                 │
│ Brass               │  ──►    │     <app>/{epic,walk-     │  ──►    │  hps/                         │
│ HPS                 │ walk-   │     through,capabilities} │  imple- │  hps-solicitudes/             │
│ HPS_Solicitudes     │ through │                          │  menta  │  condor/                      │
│ Condor              │         │   app/  (desde MVP)  │         │  brass/                       │
│ Lanzadera ★        │         │     src/modules/<app>/    │         │  gestion-riesgos/             │
│ Expedientes         │         │   docs/calidad-de-codigo- │         │  no-conformidades/            │
└────────────────────┘         │     y-ci.md (CI + gates)  │         └──────────────────────────────┘
                               │   data/staging/  (R2-pull)│                    ▲
                               └──────────────────────────┘                    │ (corte por módulo)
                                        │ quality gates                        ▼
                                ┌──────────────────────┐             ┌──────────────────────────────┐
                                │ CI: ruff/mypy/pytest/ │             │ PostgreSQL + S3 + (Redis F2) │
                                │ check_layers/security │             │ + Mistica + HTMX + Alpine.js  │
                                └──────────────────────┘             └──────────────────────────────┘
```

★ = Lanzadera es la "madre" porque ahí nacen usuarios, aplicativos y permisos que las otras 7 apps consumen.

## Recommended reading path

| Step | Doc | Read this when... |
|------|-----|---------------------|
| 1    | [README](../README.md) | Necesita entender el producto en 5 minutos |
| 2    | [DOCS](../DOCS.md) | Necesita el technical reference (endpoints, schemas, CLI) |
| 3    | Esta guía | Es usted mantenedor del repositorio o contribuidor nuevo |
| 4    | [docs/AGENT-SETUP](../docs/AGENT-SETUP.md) | Está configurando un agente (Claude, OpenCode) para trabajar aquí |
| 5    | [docs/03-aplicaciones/<app>/epic.md](../docs/03-aplicaciones/) | Necesita entender una app específica |
| 6    | [CONTRIBUTING](../CONTRIBUTING.md) | Vas a hacer un PR o agregar un artefacto nuevo |
| 7    | [docs/prompts/](../docs/prompts/) | Necesita escribir un issue o prompt a un mantenedor de dysflow |
| 8    | [openspec/](../openspec/) | Está planeando una feature SDD o un cambio mayor |
| 9    | [docs/calidad-de-codigo-y-ci](../docs/calidad-de-codigo-y-ci.md) | Está arrancando el MVP de plataforma y necesita configurar los quality gates |

## Quick map: if you need X, read Y

| If you need to...                    | Open first                                              | Then check                                                  |
|--------------------------------------|---------------------------------------------------------|-------------------------------------------------------------|
| Entender el producto en 5 min         | [README](../README.md)                                  | [DOCS](../DOCS.md)                                            |
| Ver el technical reference completo   | [DOCS](../DOCS.md)                                        | [Codebase Guide](CODEBASE-GUIDE.md) (esta guía)              |
| Configurar un agente para trabajar    | [docs/AGENT-SETUP](../docs/AGENT-SETUP.md)                | [AGENTS](../AGENTS.md)                                        |
| Entender una app específica          | `docs/03-aplicaciones/<app>/epic.md`                    | `docs/03-aplicaciones/<app>/walkthrough-*.json`              |
| Agregar una nueva app al blueprint   | Esta guía → "Cómo agregar una nueva app" abajo          | Una epic.md similar a las 8 existentes                       |
| Modificar un walkthrough             | `docs/03-aplicaciones/<app>/walkthrough-*.json`         | Esta guía → "Walkthrough patterns" abajo                     |
| Documentar un bug de dysflow          | [docs/prompts/prompt-ia-mantenedora-dysflow-round-*.md](../docs/prompts/) | [issues filed](#bugs-dysflow-filed) en DysTelefonica/dysflow |
| Diseñar UI/UX de una app             | [docs/design/mockups/](../docs/design/mockups/)         | Skills `telefonica-brand-design` + `frontend-design`         |
| Escribir docs (estilo y formato)     | Skill `documentation-patterns`                          | Skill `docs-index` + `codebase-guide`                         |
| Configurar quality gates del MVP    | [docs/calidad-de-codigo-y-ci](../docs/calidad-de-codigo-y-ci.md) | Esta guía → "Estructura del repo" + `pyproject.toml`          |
| Hacer un PR                          | [CONTRIBUTING](../CONTRIBUTING.md)                       | Esta guía → "Workflow de contribución" abajo                  |

## Estructura del repo

```
raíz/
├── README.md                          ← overview 5 min
├── DOCS.md                            ← technical reference raíz (este archivo's sibling)
├── AGENTS.md                          ← índice de skills para IAs
├── CODEBASE-GUIDE.md                  ← esta guía
├── CONTRIBUTING.md                    ← workflow + label system + conventional commits
├── CHANGELOG.md                       ← cambios por versión
├── docs/                              ← docs por audiencia
│   ├── AGENT-SETUP.md                 ← setup por agente
│   ├── calidad-de-codigo-y-ci.md      ← quality gates del MVP de plataforma
│   └── 03-aplicaciones/               ← 1 carpeta por app legacy
│       ├── <app>/
│       │   ├── epic.md                ← spec de migración
│       │   ├── walkthrough-*.json     ← JSON estructurado por form
│       │   ├── capabilities.md
│       │   ├── forms.md
│       │   ├── data-model.md
│       │   └── security-rules.md
│       └── ...
├── docs/design/                       ← mockups UI
│   └── mockups/                       ← HTML autocontenidos
│       ├── lanzadera-shell.html
│       └── gestion-riesgos-tree.html
├── docs/prompts/                      ← prompts a mantenedores externos (dysflow)
│   ├── prompt-ia-mantenedora-dysflow-round-1-*.md
│   ├── prompt-ia-mantenedora-dysflow-round-2-*.md
│   ├── prompt-ia-mantenedora-dysflow-round-3-*.md
│   └── prompt-ia-mantenedora-dysflow-round-4-*.md
├── openspec/                          ← SDD (Spec-Driven Development) specs
├── scripts/                            ← setup-staging, sync-to-r2, etc.
├── data/staging/                      ← binarios legacy NO commiteados, R2-pulled
│   ├── <app>/frontend/*.accdb        ← NO staging para Lanzadera (D156)
│   └── <app>/backend/*.accdb
└── inputs/                            ← material externo para IAs (automatizaciones-legacy, etc.); versionado liviano
└── .dysflow/project.json              ← dysflow MCP config (8 backends)
```

## Ownership de artefactos

| Artefacto | Ubicación | Owner | Cuándo se mueve |
|---|---|---|---|
| `epic.md` por app | `docs/03-aplicaciones/<app>/` | Research team | Se queda; el código va en `app/src/modules/<app>/`. |
| `walkthrough-*.json` | `docs/03-aplicaciones/<app>/` | Research team | Se queda; el código UI va en `app/src/modules/<app>/delivery/`. |
| `capabilities.md` | `docs/03-aplicaciones/<app>/` | Research team | Se queda como contrato funcional. |
| Mockups UI | `docs/design/mockups/` | UX/Research team | Se queda; la UI real va en `app/src/modules/<app>/ui/`. |
| Bugs dysflow | `docs/prompts/` + issues en DysTelefonica | Mantenedor dysflow | Issues NO se mueven; los prompts son el reporte local. |
| Binarios legacy | `data/staging/<app>/` (R2-pulled) | Infra team | NO se commitean; quedan en R2 hasta el cut-over. |
| Decisiones D1-D82 + QC-1 a QC-9 | `epic.md` per-app + cross-cutting en DOCS + [`docs/calidad-de-codigo-y-ci.md`](calidad-de-codigo-y-ci.md) | Research team + Platform team | Per-app se mantienen; cross-cutting en DOCS; quality gates en su propio doc. |
| Código de plataforma | `app/src/modules/<app>/` | Platform team | Se queda en este monorepo. |
| Migraciones Alembic | `app/migrations/versions/` | Platform team | Se queda en este monorepo (Expand & Contract, D82). |
| Quality gates + CI | [`docs/calidad-de-codigo-y-ci.md`](calidad-de-codigo-y-ci.md) + `.github/workflows/` | Platform team | El doc se mantiene; los workflows viven en `.github/`. |
| Skills de opencode | `C:\Proyectos\skills\skills\` | El user | NO se mueven; trascienden este repo. |

## Walkthrough patterns

3 métodos aplicados según el estado de dysflow:

| Método | Apps | Características | Walkthrough output |
|---|---|---|---|
| **v3** | NC, GR (pre-fix #1407) | `analyze_form_layout` SKIPPED, `form_list_controls` + lint manual de geometry | `tool_warnings: ["map_form_behavior: --json rejected (bug #1408)"]` |
| **v4** | Lanzaderas, Expedientes (post-fix #1407) | `analyze_form_layout` ✅, `map_form_behavior` con `autoFetchCodeGraph:false`, `verify_form_bindings` SKIPPED | `tool_warnings: ["verify_form_bindings: RESULT_CONTRACT_VIOLATION (issue #1412)"]` |

**Schema por form** (ver [skill `documentation-patterns`](../Proyectos/skills/skills/documentation-patterns/SKILL.md) para detalles):

```json
{
  "formName": "Form_X",
  "sourcePath": "...",
  "group": "G<N>",
  "codegraph_summary": {...},
  "ui": {...},
  "geometry": {...},
  "behavior": {...},
  "verify_form_bindings": {...},
  "unattended": false,
  "unattended_evidence": null,
  "tool_warnings": [],
  "method_version": "v4"
}
```

## Cómo agregar una nueva app al blueprint

Si en el futuro se agrega una novena app (no aplica ahora, las 8 están cerradas), los pasos son:

1. **Crear carpeta** `docs/03-aplicaciones/<nueva-app>/`.
2. **Walkthrough**: lanzar 5 sub-agentes en paralelo con método v4 (5 grupos de forms), siguiendo [skill `documentation-patterns`](../Proyectos/skills/skills/documentation-patterns/SKILL.md) y el patrón de las 8 épicas existentes.
3. **Crear `epic.md`** siguiendo la estructura de 7 secciones + 2 anexos + checklist. Ver [skill `docs-index`](../Proyectos/skills/skills/docs-index/SKILL.md).
4. **Crear walkthrough-*.json** (5 archivos, 1 por grupo).
5. **Crear PR** con el título `docs(<app>): add epic - <N>/<N> forms walkthroughed via method v<N>`.
6. **Mergear con `--squash --delete-branch`** (vía `gh pr merge`).
7. **Actualizar este CODEBASE-GUIDE** agregando la fila en la tabla "The 8 Apps" del DOCS.
8. **Crear `app/src/modules/<nueva-app>/`** con su esqueleto hexagonal (`domain/`, `ports/`, `application/`, `adapters/`, `di/`, `delivery/`). Ver [`docs/calidad-de-codigo-y-ci.md`](calidad-de-codigo-y-ci.md) §Hexagonal layer gate para el contrato arquitectónico.

## Workflow de contribución

Ver [CONTRIBUTING](../CONTRIBUTING.md) para:

- Conventional commit format (`docs(app): ...`, `fix(dysflow): ...`, etc.).
- Label system (`type:*`, `status:*`, `priority:*`).
- PR rules (no Co-Authored-By, no mega-commits, etc.).
- Issue-first workflow.

**Reglas específicas de este repo:**

- Las épicas de migración siguen Conventional Commits con `docs(app):` prefix.
- Los walkthrough JSONs NO son docs humanas — son data. Van en commits separados de la epic.md.
- Los issues a dysflow (`docs/prompts/`) son work-in-progress hasta que se mergean upstream — no los merges en main hasta entonces.

## Full technical reference stays in DOCS.md

Esta guía explica ownership, flows, y guardrails. **NO duplica** la API reference completa. Para endpoints, schemas, MCP parameters, y CLI flags, usá [DOCS.md](../DOCS.md).

## What this is

| It is | Evidence in this repo |
|---|---|
| Mapa de ownership, flows y guardrails para mantenedores y contribuidores nuevos. | Las secciones «Recommended reading path», «Quick map inverso» y «Ownership de artefactos» más arriba. |
| Índice de las decisiones operativas del monorepo (rename `platform/` → `app/`, walkthrough v3/v4, scope multi-app). | `docs/architecture.md` §Layout del repo + `CONTRIBUTING.md` §Convención multi-app. |
| Punto de entrada único para agregar una nueva app al blueprint (procedimiento en 8 pasos). | Sección «Cómo agregar una nueva app al blueprint» arriba. |

## What this is not

| It is not | Use this boundary |
|---|---|
| La fuente de verdad arquitectónica. | Las decisiones D-<n> y DA-<n> viven en [`docs/architecture.md`](docs/architecture.md); este doc sólo las referencia. |
| La referencia técnica de endpoints, schemas o CLI flags. | [`DOCS.md`](../DOCS.md) es la technical reference raíz. |
| El manual de uso de los `check_*.py`. | [`docs/calidad-de-codigo-y-ci.md`](docs/calidad-de-codigo-y-ci.md) describe los 12 gates; este doc sólo nombra cuál aplica a qué artefacto. |
| Una guía de estilo de código o convención de naming. | [`CONTRIBUTING.md`](../CONTRIBUTING.md) §Convención multi-app + `pyproject.fragment.toml` + ruff/mypy config. |

## Core invariants

- **Naming multi-app**: cada issue y commit declara a qué app pertenece con el prefijo `[<APP>]` o `[XCUT]` en el título y el scope `(app)` o `(platform)` en el commit. El `app/<slug>` label va cuando afecta una sola app.
- **Ownership de `data/staging/<app>/`**: los binarios legacy viven en R2 y se pull-ean con `scripts/setup-staging.ps1` y `scripts/sync-to-r2.ps1`. NO se commitean nuevas adiciones a `data/`; el staging se mantiene reproducible desde R2.
- **Walkthrough JSONs son data**: los `walkthrough-*.json` viven en commits separados de la `epic.md` correspondiente. Mezclarlos dificulta la revisión y bloquea el ciclo de `dysflow` walkthrough.
- **Issues a dysflow son work-in-progress**: los `docs/prompts/prompt-ia-mantenedora-dysflow-round-*.md` no se mergean en `main` hasta que el issue upstream en DysTelefonica/dysflow se cierre; mantenerlos en working tree.
- **Ramas de MVP se conservan en remoto**: cada PR se integra con `--squash` y la rama remota se queda; la limpieza es local (`git worktree remove <ruta>`). Borrar la rama remota rompe la trazabilidad por unidad de trabajo.

## Existing references

| Doc | Owns | Referenced from this guide |
|---|---|---|
| [`docs/architecture.md`](docs/architecture.md) | Decisiones D-<n> y DA-<n>, layout del repo, gaps conocidos. | §Decisiones metodológicas, §Stack, §Layout. |
| [`DOCS.md`](../DOCS.md) | Technical reference raíz: endpoints, schemas, MCP, CLI, env vars. | §Full technical reference stays in DOCS.md. |
| [`docs/calidad-de-codigo-y-ci.md`](docs/calidad-de-codigo-y-ci.md) | Los 12 `check_*.py` + los 4 workflows de CI. | §CI gates, §Quality gates, §Estructura del repo. |
| [`docs/03-aplicaciones/<app>/epic.md`](docs/03-aplicaciones/) | Spec de migración por app + walkthroughs + capabilities. | §Recommended reading path, §Ownership de artefactos, §Cómo agregar una nueva app. |
| [`docs/prompts/`](docs/prompts/) | Reportes al mantenedor de dysflow (WIP hasta merge upstream). | §Bugs dysflow filed, §Ownership de artefactos. |
| [`CONTRIBUTING.md`](../CONTRIBUTING.md) | Workflow de contribución + label system + convención multi-app. | §Workflow de contribución, §Recommended reading path. |
| [`AGENTS.md`](../AGENTS.md) | Índice de skills + catálogo obligatorio para IAs. | §Quick map inverso (Configurar un agente). |
| [`openspec/`](../openspec/) | Cambios SDD por app: proposal, design, specs, tasks. | §Recommended reading path (cambios mayores). |
| [`skills/documentation-alan-style/SKILL.md`](skills/documentation-alan-style/SKILL.md) | Contrato de tono, formato y anti-patrones de toda la doc. | §Skills de documentación, §Recommended reading path. |

## Contributor checklist

- [ ] El cambio respeta los 5 gates del workflow (`check_branch_name`, `check_pr_size`, `check_workflows`, ruff, mypy).
- [ ] Si toca `app/`, `tests/` u `openspec/`, leyó [`docs/architecture.md`](docs/architecture.md) §Decisiones D-<n> cross-cutting antes de empezar.
- [ ] Si toca `docs/`, leyó `skills/documentation-alan-style/SKILL.md` y aplica tono castellano peninsular formal con usted.
- [ ] El commit usa Conventional Commits con el scope `(app)` o `(platform)` correspondiente.
- [ ] El PR es ≤ 400 líneas (`additions + deletions`); si no, parte por unidad de trabajo o encadena.
- [ ] El worktree local se limpia tras mergear con `git worktree remove <ruta>` (la rama remota se queda).
- [ ] El estado-planificacion HTML más reciente está actualizado si el cambio pertenece a un ciclo de refactor de app.
- [ ] Ningún `Co-Authored-By` ni atribución de IA en los commits.

---

[Next: Mental Model →](../docs/03-aplicaciones/lanzadera/epic.md)
