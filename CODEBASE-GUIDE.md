[← Back to README](../README.md)

# access2web-blueprint — Codebase Guide

**This guide is for maintainers and contributors who need to understand how blueprint + platform responsibilities coexist in this monorepo, where research, code and documentation live, and where to add new content when the platform grows.**

access2web-blueprint es **monorepo de la plataforma web hexagonal + blueprint del refactor** de las 8 apps legacy Access/VBA. Aloja `docs/` (research + decisiones + épicas) y `app/` (código FastAPI + HTMX + Alembic del MVP de Lanzadera, 2026-08). El legacy `.accdb` queda en repos separados hasta el cut-over del ecosistema completo.

> **Sentence that organizes the whole repo**: "Lanzadera es la madre: ahí nacen usuarios, aplicativos y permisos. Las otras 7 apps son consumidoras."

> **Scope del scope**: "Este repo es monorepo de plataforma + blueprint. Desde el MVP de Lanzadera (2026-08), `app/` contiene el código web hexagonal; `docs/` contiene el research y las decisiones; `data/staging/` contiene los binarios legacy extraídos."

## 90-second mental model

```text
Legacy apps (Access/VBA)         Este monorepo                          Apps web (app/modules/<app>/)
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
└── .dysflow/project.json              ← dysflow MCP config (8 backends)
```

## Ownership de artefactos

| Artefacto | Ubicación | Owner | Cuándo se mueve |
|---|---|---|---|
| `epic.md` por app | `docs/03-aplicaciones/<app>/` | Research team | Se queda; el código va en `app/modules/<app>/`. |
| `walkthrough-*.json` | `docs/03-aplicaciones/<app>/` | Research team | Se queda; el código UI va en `app/modules/<app>/delivery/`. |
| `capabilities.md` | `docs/03-aplicaciones/<app>/` | Research team | Se queda como contrato funcional. |
| Mockups UI | `docs/design/mockups/` | UX/Research team | Se queda; la UI real va en `app/modules/<app>/ui/`. |
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
8. **Crear `app/modules/<nueva-app>/`** con su esqueleto hexagonal (`domain/`, `ports/`, `application/`, `adapters/`, `di/`, `delivery/`). Ver [`docs/calidad-de-codigo-y-ci.md`](calidad-de-codigo-y-ci.md) §Hexagonal layer gate para el contrato arquitectónico.

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

---

[Next: Mental Model →](../docs/03-aplicaciones/lanzadera/epic.md)
