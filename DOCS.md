---
description: DOCS — technical reference raíz del monorepo access2web-blueprint (8 apps, stack pinned, decisiones D1-D178, walkthrough patterns).
globs: *
alwaysApply: true
---

[← Back to README](README.md)

# access2web-blueprint — Technical Reference

**Blueprint de migración de 8 apps legacy Access/VBA → web hexagonal (FastAPI + HTMX).**

This is the complete technical reference for the blueprint. For getting started, see the [README](README.md). For per-agent setup, see [Agent Setup](docs/AGENT-SETUP.md).

> **Scope**: Este repo es **monorepo de plataforma + blueprint del refactor**. Aloja la documentación de las 8 apps legadas (Access/VBA → web hexagonal) y, desde el MVP de Lanzadera (2026-08), el código de la plataforma (`app/`: FastAPI + HTMX + Alembic + Docker; renombrado desde `platform/` el 2026-08-09 por shadowing del módulo stdlib — ver `app/README.md`). Las apps legadas viven en repos separados hasta el cut-over del ecosistema completo.

> **Sentence that organizes the whole repo**: "Lanzadera es la madre: ahí nacen usuarios, aplicativos y permisos. Las otras 7 apps son consumidoras."

## What this is

| It is | Evidence in this repo |
|---|---|
| Technical reference raíz: stack pinned, decisiones D1-D178, walkthrough patterns, mocks UI, bugs dysflow filed. | Las secciones «Stack Target», «Decisiones metodológicas», «Migración de binarios», «Bugs dysflow filed», «Walkthrough patterns», «Diseño UI/UX» más abajo. |
| Catálogo del estado de las 8 apps con método walkthrough aplicado + PR de merge por app. | §The 8 Apps: 8/8 cerradas con PR de referencia (#1, #2, #3, #4 o «pre-PR workflow»). |
| Índice de cross-cutting decisions (D8 hexagonal, D14 esquema por módulo, D16 S3, D27 logs, D66-D68 stack, D82 Expand and Contract, D132 XApp, etc.). | §Decisiones metodológicas: tabla D-<n> con apps afectadas y estado. |

## What this is not

| It is not | Use this boundary |
|---|---|
| El mapa de ownership, flows y guardrails para mantenedores. | [`CODEBASE-GUIDE.md`](CODEBASE-GUIDE.md) es la guía de ownership + reading path + cómo agregar una app nueva. |
| La fuente de verdad de las decisiones arquitectónicas D-<n> y DA-<n>. | [`docs/architecture.md`](docs/architecture.md) cataloga decisiones vigentes y obsoletas con paths; este doc sólo las referencia en `§Decisiones metodológicas`. |
| El manual de los 12 `check_*.py` y los 4 workflows de CI. | [`docs/calidad-de-codigo-y-ci.md`](docs/calidad-de-codigo-y-ci.md) describe cada gate; este doc sólo lista dónde se orquestan. |
| El walkthrough por form individual. | Cada `walkthrough-*.json` por app vive en `docs/03-aplicaciones/<app>/walkthrough-*.json`; este doc sólo explica el método v3/v4. |

---

## Quick Navigation

| Section                                          | What you'll find                                                                |
|--------------------------------------------------|---------------------------------------------------------------------------------|
| [Scope del blueprint](#scope-del-blueprint)      | Qué es y qué NO es este repo. Lifecycle research → implementación por app.      |
| [The 8 Apps](#the-8-apps)                         | Estado de cada épica: Lanzadera, HPS, HPS_Solicitudes, Brass, Gestion_Riesgos, NoConformidades, Lanzaderas, Expedientes. |
| [Stack Target](#stack-target)                    | Backend Python/FastAPI/SQLAlchemy · Frontend HTMX/Jinja2/Alpine.js.              |
| [Decisiones metodológicas](#decisiones-metodologicas) | D1-D178 acumuladas. Cross-cutting + per-app.                                  |
| [Migración de binarios](#migracion-de-binarios)  | Cómo staging de binarios .accdb se hace vía R2 (data/staging/*).                |
| [Bugs dysflow filed](#bugs-dysflow-filed)        | Issues abiertos a DysTelefonica/dysflow que afectan el walkthrough.               |
| [Walkthrough patterns](#walkthrough-patterns)    | Método v3/v4 de walkthrough de forms + lint manual de geometry.                  |
| [Diseño UI/UX](#dise%C3%B1o-uiux)                    | Mockups aprobados con Mistica design system + frontend-design skill.             |
| [Skills de documentación](#skills-de-documentacion) | Patrones propios para escribir docs que sirvan a IAs + humanos.                |

For other docs:

| Doc                                                         | Description                                                                                                  |
|-------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------|
| [README](README.md)                                         | Overview de 5 minutos: qué es esto, para quién, cómo empezar.                                               |
| [Agent Setup](docs/AGENT-SETUP.md)                          | Configuración de agentes (Claude, OpenCode, Gemini, Codex) para trabajar en este repo.                      |
| [Codebase Guide](CODEBASE-GUIDE.md)                          | Para mantenedores: 90-second mental model, ownership de artefactos, quick map inverso.                     |
| [Calidad de código y CI](docs/calidad-de-codigo-y-ci.md)     | Quality gates del MVP de plataforma: hexagonal layer gate, ruff, mypy, security scanning, plan día 0-6.     |
| [CONTRIBUTING](CONTRIBUTING.md)                             | Workflow de contribución, conventional commits, label system.                                                |
| [CHANGELOG](CHANGELOG.md)                                   | Cambios por versión del blueprint (cierre de épicas, PRs merged, etc.).                                     |
| [Lanzadera Epic](docs/03-aplicaciones/lanzadera/epic.md)    | Épica de Lanzadera — la madre (users + apps + permissions). 28 forms walkthroughed.                      |
| [NoConformidades Epic](docs/03-aplicaciones/no-conformidades/epic.md) | Épica de NoConformidades. 48 forms walkthroughed.                                                |
| [Gestion_Riesgos Epic](docs/03-aplicaciones/gestion-riesgos/epic.md) | Épica de Gestion_Riesgos. 61 forms walkthroughed.                                                  |
| [Brass Epic](docs/03-aplicaciones/brass/epic.md)            | Épica de Brass. 85 forms walkthroughed.                                                                   |
| [HPS Epic](docs/03-aplicaciones/hps/epic.md)                | Épica de HPS.                                                                                              |
| [HPS_Solicitudes Epic](docs/03-aplicaciones/hps-solicitudes/epic.md) | Épica de HPS_Solicitudes.                                                                          |
| [Expedientes Epic](docs/03-aplicaciones/expedientes/epic.md) | Épica de Expedientes. 46 forms walkthroughed.                                                              |
| [UI Mockups](docs/design/mockups/)                          | Mockups HTML autocontenidos con Mistica design system (Lanzadera shell + Gestion_Riesgos tree).            |

---

## Scope del blueprint

| Qué ES | Qué NO es |
|---|---|
| Blueprint del refactor de las 8 apps legadas (Access/VBA → web hexagonal). | Réplica exacta de los binarios `.accdb` legados. |
| Discovery del modelo de datos + forms + behavior + decisiones arquitectónicas. | Migración one-shot por app (la del MVP es Lanzadera). |
| Walkthroughs de forms legados (JSON estructurado). | Tests E2E de cada app web (viven en `app/tests/`). |
| Hallazgos + decisiones arquitectónicas (D1-D82, QC-1 a QC-9). | Disposición final del legacy (sigue corriendo hasta UAT del ecosistema). |
| Especificación de cada épica de migración (`epic.md` por app). | Specs de producto post-cut-over (viven aquí mismo). |
| Mockups UI base con Mistica design system para validar look & feel. | UI final de cada app (se construye aquí mismo en `app/src/modules/<app>/ui/`). |
| Código de la plataforma web hexagonal (desde MVP Lanzadera, en `app/`). | Código de dysflow (es de Gentleman-Programming). |
| Quality gates y CI del MVP (ver [`docs/calidad-de-codigo-y-ci.md`](calidad-de-codigo-y-ci.md)). | CI/CD de cada app legada. |
| Mantenimiento de issues dysflow (#1407 cerrado, #1408/#1412 abiertos). | — |

**Lifecycle** de un artefacto de este monorepo:

```text
1. Research + discovery  → vive en este repo (docs/03-aplicaciones/<app>/)
2. Plataforma (MVP)      → vive en este repo (app/src/modules/<app>/)
3. Migración por app     → cada app (Lanzadera, Expedientes, ...) se construye aquí mismo
4. Cut-over              → el legacy .accdb queda como referencia histórica; el código web es la fuente de verdad
```

> **Este repo es monorepo desde el MVP de Lanzadera (2026-08). El blueprint describe la dirección; el código la ejecuta. Las apps legadas siguen en repos separados hasta el cut-over del ecosistema completo.**

---

## The 8 Apps

Estado del walkthrough de cada app (8/8 cerradas):

| App | Forms | Método | Estado | PR |
|---|---|---|---|---|
| **Lanzadera** (la madre) | 28 | v4 | ✅ Épica mergeada | [PR #3](https://github.com/DysTelefonica/access2web-blueprint/pull/3) |
| **NoConformidades** | 48 | v3 | ✅ Épica mergeada | [PR #2](https://github.com/DysTelefonica/access2web-blueprint/pull/2) |
| **Gestion_Riesgos** | 61 | - | ✅ Épica mergeada | [PR #1](https://github.com/DysTelefonica/access2web-blueprint/pull/1) |
| **Brass** | 85 | - | ✅ Épica mergeada | (pre-PR workflow) |
| **HPS** | - | - | ✅ Épica mergeada | (pre-PR workflow) |
| **HPS_Solicitudes** | - | - | ✅ Épica mergeada | (pre-PR workflow) |
| **Condor** | - | - | ✅ Épica mergeada | (pre-PR workflow) |
| **Expedientes** | 46 | v4 | ✅ Épica mergeada | [PR #4](https://github.com/DysTelefonica/access2web-blueprint/pull/4) |

**Total forms walkthroughed** (research cerrado): ~268 forms.

Cada `epic.md` por app sigue la misma estructura (7 secciones + 2 anexos + checklist). Ver [Codebase Guide](CODEBASE-GUIDE.md) para detalles.

---

## Stack Target

Definido en cada epic + cross-cutting en todas:

- **Backend**: Python 3.12+ · FastAPI 0.119+ · SQLAlchemy 2.0.x · Alembic 1.13+ · asyncpg 0.30+
- **Frontend**: HTMX 2.0.4 + Jinja2 3.1+ + Alpine.js 3.15+ (server-rendered, NO React)
- **DB**: PostgreSQL compartido con esquema por módulo (D14)
- **Estrategia de migración**: Expand and Contract backward-compatible (D82)
- **Forma destino**: Hexagonal global (D8) · módulo dentro del monolito modular (D68)

Para UI styling: **Mistica design system** (Telefónica brand) con design tokens CSS server-rendered, NO React components. Ver [UI Mockups](docs/design/mockups/) para referencia.

---

## Decisiones metodológicas

Las decisiones D1-D178 se documentan en cada epic.md. Decisiones **cross-cutting** (aplican a múltiples apps):

| ID | Decisión | Apps afectadas |
|---|---|---|
| D8 | Hexagonal global | Todas |
| D14 | Esquema por módulo en PostgreSQL | Todas |
| D16 | Object storage S3-compatible para anexos | NC, GR, Brass, Lanzadera |
| D27 | Logs estructurados (no tablas TbLog) | Cross-cutting |
| D44-D46 | Autorización + capabilities | Todas |
| D66-D67 | Stack: Python/FastAPI/HTMX | Todas |
| D68 | Monolito modular | Todas |
| D82 | Expand and Contract | Todas |
| D86-D87 | Cross-app con `getdbLanzadera()` (NoConformidades + Lanzaderas) | NC + Lanzaderas |
| D88 | 372 getdb() callers en Gestion_Riesgos | GR específicamente |
| D102 | Booleanos Text(2) — cleanup captions | Varias |
| D132 | XApp HTTP/JSON handshake (HPS, Expedientes, Gestion_Riesgos) | Cross-app |
| D140 | Concurrencia con optimistic locking | GR + NC |
| D144 | Timer-driven UX (Lanzadera) → skeleton web | Lanzadera |
| **D146** | **Methodology v3 (no conformidades) — bug dysflow #1407 + #1412 worked around** | **Walkthrough** |
| **D155** | **Methodology v4 (lanzaderas + expedientes) — bug dysflow #1407 RESUELTO en 2.36.2** | **Walkthrough** |
| **D168** | **Methodology v4 aplicada a Expedientes** | **Walkthrough** |

Decisiones **per-app** (D150-D178) viven en cada `epic.md` correspondiente.

---

## Migración de binarios

Los legacy apps son Access/VBA con binarios `.accdb` que no se commitean al repo (auto-contenido, ver D82). Staging se hace vía R2 bucket `access2web-staging-binaries`:

```bash
# Staging inicial de los 8 binarios
pwsh scripts/setup-staging.ps1

# Sync diario (incremental)
pwsh scripts/sync-to-r2.ps1
```

Estado actual del staging:

- `data/staging/lanzaderas/frontend/Lanzadera.accdb` (NO staging — D156)
- `data/staging/no-conformidades/frontend/NoConformidades.accdb` (66.96 MB) · `backend/NoConformidades_Datos.accdb` (32.18 MB) ✅
- `data/staging/gestion-riesgos/frontend/Gestion_Riesgos.accdb` ✅
- `data/staging/brass/frontend/Gestion_Brass_Gestion.accdb` ✅
- `data/staging/expedientes/frontend/Expedientes.accdb` ✅
- `data/staging/hps-solicitudes/frontend/Solicitudes_HPS.accdb` ✅
- `data/staging/condor/frontend/CONDOR.accdb` ✅
- `data/staging/hps/frontend/HPS.accdb` ✅

**Lanzadera**: source-only en este repo (binarios NO staging — D156). Para walkthrough audit de queries/macros se requiere pull de R2.

---

## Bugs dysflow filed

El walkthrough de forms usa **dysflow MCP** + **codegraph-vba** (vía MCP) para descubrir el contenido de cada form. Tres issues filed en [DysTelefonica/dysflow](https://github.com/DysTelefonica/dysflow):

| Issue | Tool | Estado | Hallazgo |
|---|---|---|---|
| [#1407](https://github.com/DysTelefonica/dysflow/issues/1407) | `analyze_form_layout` | ✅ **CERRADO en 2.36.2** | Opaque `RESULT_CONTRACT_VIOLATION` |
| [#1408](https://github.com/DysTelefonica/dysflow/issues/1408) | `map_form_behavior` autoFetchCodeGraph | ❌ OPEN | `--json` flag rejected por codegraph-vba CLI |
| [#1412](https://github.com/DysTelefonica/dysflow/issues/1412) | `verify_form_bindings` | ❌ OPEN | Opaque `RESULT_CONTRACT_VIOLATION` (mismo patrón que #1407) |

**Methodology v3** (NC) y **v4** (Lanzaderas, Expedientes) aplican workaround:
- `analyze_form_layout` → v4 ya funciona (post-2.36.2).
- `map_form_behavior` → usar `autoFetchCodeGraph:false`.
- `verify_form_bindings` → skip con `status:"skipped_tool_broken"`.

---

## Walkthrough patterns

3 métodos aplicados según el estado de dysflow:

| Método | Cuándo | Características |
|---|---|---|
| **v3** | NC, GR (con bug #1407 abierto) | `analyze_form_layout` SKIPPED, `form_list_controls` + lint manual de geometry |
| **v4** | Lanzaderas, Expedientes (post-fix #1407) | `analyze_form_layout` ✅, `map_form_behavior` con `autoFetchCodeGraph:false`, `verify_form_bindings` SKIPPED |

Output por form:

```json
{
  "formName": "Form_FormAuditoriaSeleccion",
  "sourcePath": "C:/.../Form_FormAuditoriaSeleccion.form.txt",
  "group": "G2",
  "codegraph_summary": {
    "handlers_count": 12,
    "source_size_lines": 540,
    "blast_radius_top_5": [...],
    "call_path_top_3": [...],
    "stale_banner_ignored": true
  },
  "ui": { "controls_count": 23, "formEvents": ["Form_Open"], "bindings_count": 5, "warnings": [] },
  "geometry": { ... },
  "behavior": { "controls_with_handler_evidence": 8, "tables_referenced": [...], "codegraph_evidence_status": "empty_bug_1408" },
  "verify_form_bindings": { "status": "skipped_tool_broken", "reason": "verify_form_bindings returns RESULT_CONTRACT_VIOLATION (issue #1412)" },
  "unattended": false,
  "unattended_evidence": null,
  "tool_warnings": [...],
  "method_version": "v4"
}
```

Ver [Codebase Guide — Walkthrough patterns](CODEBASE-GUIDE.md#walkthrough-patterns) para detalles.

---

## Status Visibility Matrix

Cómo se ve el estado del blueprint en cada superficie cuando algo cambia:

| Scenario | Expected reason | Surfaces |
|---|---|---|
| PR abierto en una app | El cambio pertenece a una épica de migración o a un OpenSpec change. | GitHub PR (status: in_progress), `docs/03-aplicaciones/<app>/README.md` enlazando al PR, `docs/uat/estado-planificacion-NEW_<fecha>.html` columna RAG de la tabla de PRs (verde/ámbar/rojo). |
| PR en revisión | El revisor humano aplicó `status:approved` o pidió cambios. | GitHub PR (status: review), `docs/uat/estado-planificacion-NEW_<fecha>.html` actualizado por el autor con motivo del cambio. |
| PR mergeado | La épica o el change se cerraron. | GitHub PR (status: merged), `epic.md` de la app marcada como mergeada en `§The 8 Apps` arriba, `docs/uat/estado-planificacion-NEW_<fecha>.html` columna RAG en verde. |
| CI rojo en un PR | Fallo de un check local o de un workflow; no es regresión de runtime. | GitHub Actions logs en el PR (no bloquean por branch protection), comentario del revisor al autor antes de mergear, `estado-planificacion-*.html` columna RAG en rojo. |
| Issue dysflow filed upstream | El walkthrough encontró un bug en dysflow MCP o codegraph-vba. | `docs/prompts/prompt-ia-mantenedora-dysflow-round-*.md` (WIP), GitHub issue en DysTelefonica/dysflow, §Bugs dysflow filed arriba actualizada con estado. |
| Change archivado | El OpenSpec change cerró: spec implementada, tasks terminadas, apply-progress y verify-report persistidos. | `openspec/changes/<change>/` con `archive-report.md`, `CHANGELOG.md` entrada en `[Unreleased]` o en la versión de release. |

> El `estado-planificacion-*.html` es la single source of truth visual para los jefes. Cuando un agente entra al repo, lo lee primero antes de proponer cambios.

---

## Diseño UI/UX

Mockups HTML autocontenidos con **Mistica design system** + **frontend-design skill** aplicada:

- `docs/design/mockups/lanzadera-shell.html` (61 KB) — Login + shell con sidebar + dashboard con stat cards + task list + app cards.
- `docs/design/mockups/gestion-riesgos-tree.html` (28 KB) — Árbol de riesgos 4 niveles con detail panel.

Para visualizar:

```bash
# Servir localmente
pwsh -Command "python -m http.server 8765 --directory docs/design/mockups"
# Abrir http://localhost:8765/lanzadera-shell.html
```

**Reglas de diseño aplicadas** (Mistica + frontend-design):

- Sin defaults AI-generated (cream/serif/terracotta; near-black/acid-green; broadsheet/hairline).
- Ink teal `#0A2540` como firma Mistica.
- Glass navbar con `backdrop-filter: blur(16px) saturate(180%)`.
- Sparklines SVG en stat cards.
- Mesh gradient en login hero.
- Buttons con radius 32px (Mistica web).

---

## Skills de documentación

3 skills creadas en `C:\Proyectos\skills\skills\` (linkeadas en opencode) que aplican los patrones de Gentleman-Programming adaptados:

- [`documentation-patterns`](../Proyectos/skills/skills/documentation-patterns/SKILL.md) — reglas generales para escribir docs que sirvan a IAs + humanos.
- [`docs-index`](../Proyectos/skills/skills/docs-index/SKILL.md) — patrón `DOCS.md` raíz (este doc).
- [`codebase-guide`](../Proyectos/skills/skills/codebase-guide/SKILL.md) — patrón `CODEBASE-GUIDE.md` (mantenedores).

Las skills siguen el formato de Gentleman-Programming: frontmatter YAML, When to Use con bullets, Critical Rules numeradas, Verification checklist, Cookbook con If/Then/Example, Anti-patterns ❌, Acceptance scenarios ✅.

---

## Reinforcement

Si este `DOCS.md` no se entiende en 5 minutos por un humano nuevo, o si una IA no puede parsearlo para entender el estado del blueprint, **NO publiques**. Revisa primero. La calidad de la documentación define la velocidad del proyecto a 6 meses vista.

## Contributor checklist

- [ ] El cambio mantiene la `Quick Navigation` y la tabla de `Other docs` actualizadas si se agregan o mueven secciones.
- [ ] Las decisiones nuevas se proponen primero en `openspec/changes/<change>/design.md` y luego se referencian aquí; nunca se añade una D-<n> directamente a `§Decisiones metodológicas` sin pasar por el SDD.
- [ ] Los bugs de dysflow se reportan como `docs/prompts/prompt-ia-mantenedora-dysflow-round-N-*.md` y como issue en DysTelefonica/dysflow, no sólo en `§Bugs dysflow filed`.
- [ ] El `Status Visibility Matrix` se actualiza si cambia el flujo de visibilidad (nuevo estado, nueva superficie).
- [ ] El PR respeta los 5 gates del workflow (`check_branch_name`, `check_pr_size`, `check_workflows`, ruff, mypy).
- [ ] Si toca `app/`, `tests/` u `openspec/`, leyó [`docs/architecture.md`](docs/architecture.md) §Decisiones D-<n> cross-cutting antes de empezar.

[Next: CODEBASE-GUIDE → CODEBASE-GUIDE.md]
