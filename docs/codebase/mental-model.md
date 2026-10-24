[← Volver al Codebase Guide](../CODEBASE-GUIDE.md)

# Modelo mental

> **Frase que organiza el repo entero**: «Lanzadera es la madre: ahí nacen usuarios, aplicativos y permisos. Las otras 7 apps son consumidoras.»

Este repo es el **monorepo de la plataforma + blueprint del refactor** de las 8 apps legadas Access/VBA a web hexagonal. Aloja `docs/` (research + decisiones + épicas) y, desde el MVP de Lanzadera (2026-08), `app/` (código FastAPI + HTMX + Alembic). El legacy `.accdb` vive en repos separados hasta el cut-over del ecosistema.

## Qué es este proyecto

| Es | Evidencia en este repo |
|---|---|
| Blueprint del refactor de 8 apps legadas a web hexagonal. | `docs/03-aplicaciones/<app>/{epic.md,walkthrough-*.json,capabilities.md}` para `lanzadera`, `expedientes`, `hps`, `hps-solicitudes`, `brass`, `gestion-riesgos`, `no-conformidades`, `condor`. |
| Plataforma web hexagonal (FastAPI + HTMX + Alembic). | `app/src/main.py` (composition root) + `app/src/modules/<app>/{domain,ports,application,adapters,di,delivery}/`. |
| Catálogo de decisiones arquitectónicas (D1-D178, DA-1..DA-13). | `docs/decisions/README.md` (D-<n>) y `docs/architecture.md` §Decisiones de diseño del MVP (DA-<n>). |
| Suite de quality gates deterministas. | `scripts/check_*.py` + `docs/gates.md` (manual de uso de cada gate). |
| Pipeline de CI por nivel (PR / nightly / release) y agregador de branch protection. | `.github/workflows/ci.yml` (jobs `pr-gate`, `quality`, `integration`, `security`, `codeql`, `required`) y `.github/workflows/pr-gate.yml` (el `pr-gate` en solitario, que escucha los seis eventos de `pull_request`). |
| Suite de tests: unit, integration, e2e, gate-smoke. | `tests/lanzadera/`, `tests/expedientes/`, `e2e/*.spec.ts`, marker `@pytest.mark.e2e_smoke` (gate-smoke subprocess tests). |

## Qué no es este proyecto

| No es | Use esta frontera |
|---|---|
| Réplica exacta de los binarios `.accdb` legados. | `data/staging/<app>/` se R2-pullea; el staging NO se commitea (ver `docs/architecture.md` §Layout). |
| Migración one-shot por app. | Solo Lanzadera (MVP) y Expedientes (F04 scaffold en `app/src/modules/expedientes/delivery/__init__.py`) tienen código hexagonal en este repo. |
| Suite E2E ejecutada en PR. | La suite Playwright en `e2e/` corre en `.github/workflows/release.yml::e2e` contra la imagen publicada. |
| Memoria persistente entre sesiones. | No hay `engram`, `memory.json` ni estado de runtime commiteado; los binarios legacy viven en R2. |
| Cola de merge. | La rama `main` se protege con agregador `required`; `merge_group` desapareció con el modelo gentle-ai (ver `docs/gates.md` §3 y `.github/workflows/ci.yml` comentarios). |

## Modelo de arquitectura en 90 segundos

```text
                    ┌──────────────────────────────┐
                    │  Legacy Access/VBA (.accdb)   │
                    │  en data/staging/<app>/ (R2)  │
                    └──────────────┬───────────────┘
                                   │ dysflow + codegraph-vba
                                   ▼
┌──────────────────────────────────────────────────────────────────┐
│  Este monorepo                                                     │
│                                                                   │
│  docs/03-aplicaciones/<app>/{epic,walkthrough-*.json,capabilities}│
│  ├── Lanzadera ★ (mother, F01 MVP)                                │
│  ├── Expedientes (F04 scaffold en application/, delivery stub)    │
│  └── 6 apps más (research cerrado)                                │
│                                                                   │
│  app/  (Python hexagonal, renombrado desde platform/ 2026-08-09)  │
│  ├── src/main.py                  ← FastAPI composition root      │
│  ├── src/modules/<app>/{domain,ports,application,adapters,        │
│  │                       di,delivery}                              │
│  ├── src/shared/                  ← cross-cutting (cache, audit)  │
│  ├── migrations/versions/         ← Alembic 0001..0006 (D82)       │
│  ├── tests/                       ← unit + integration            │
│  └── pytest_plugin/               ← coverage_gate, etc.           │
│                                                                   │
│  openspec/changes/<change>/{proposal,design,specs/<capability>/spec,tasks}.md    │
│  scripts/check_*.py               ← quality gates (ver docs/gates.md)│
│  e2e/                             ← Playwright suite (release gate)│
│  .github/workflows/{ci,release,security,codeql,security-deep,     │
│                      operating-doc-drift}.yml                      │
│  skills/                          ← copia internalizada de skills  │
└──────────────────────────────────────────────────────────────────┘
                                   │
                                   ▼
                ┌──────────────────────────────────────┐
                │ PostgreSQL (esquema por módulo, D14) │
                │ S3-compatible (anexos, D16)           │
                │ + Redis (F2, detrás de CachePort)    │
                │ + Mistica + HTMX + Alpine.js         │
                └──────────────────────────────────────┘
```

## Invariantes del proyecto

- **Naming multi-app**: cada issue y commit declara a qué app pertenece con el prefijo `[<APP>]` o `[XCUT]` en el título y el scope `(app)` o `(platform)` en el commit. El label etiquetas con prefijo `app` aplica cuando afecta una sola app.
- **Ownership de `data/staging/<app>/`**: los binarios legacy viven en R2 y se pull-ean con `scripts/setup-staging.ps1` y `scripts/sync-to-r2.ps1`. NO se commitean nuevas adiciones a `data/`; el staging se mantiene reproducible desde R2.
- **Walkthrough JSONs son data**: los `walkthrough-*.json` viven en commits separados de la `epic.md` correspondiente. Mezclarlos dificulta la revisión y bloquea el ciclo de dysflow walkthrough.
- **Issues a dysflow son work-in-progress**: los `docs/prompts/prompt-ia-mantenedora-dysflow-round-*.md` NO se mergean en `main` hasta que el issue upstream en DysTelefonica/dysflow se cierre.
- **Ramas de MVP se conservan en remoto**: cada PR se integra con `--squash` y la rama remota se queda; la limpieza es local (`git worktree remove <ruta>`). Borrar la rama remota rompe la trazabilidad por unidad de trabajo.
- **Hexagonalidad enforced por gate**: el job `docs-architecture` rechaza imports que violen las reglas del bloque `architecture` de `docs/architecture.md` (capas, independencia entre aplicaciones y pureza de `domain`, `ports` y `application`). Mover un adapter a la capa equivocada es un gate failure, no un estilo (DA-1, `docs/architecture.md` §Capas y dependencias).
- **Pin test AST rechaza símbolos legacy**: `scripts/check_mutation_sites.py` y `scripts/check_legacy_hashes.py` prohíben `legacy_hash`, `sha256`, `migrate_password` en código de plataforma (DA-13).

## Si necesitas X, lee/toca Y

| Si necesitas... | Empieza por... | Luego revisa... |
|---|---|---|
| Entender qué es Lanzadera y qué consumen las otras apps | Esta página | `docs/08-decisiones-y-preguntas-abiertas.md` |
| Saber en qué capa va un cambio | `docs/codebase/repository-map.md` | `docs/architecture.md` |
| Comprobar si algo existe antes de inventarlo | `docs/codebase/what-is-not-here.md` | `docs/codebase/interfaces.md` |

## Regla de colocación

El código nuevo sigue el orden `domain`, `ports`, `application`, `adapters`, `delivery`, `di`; una capa solo importa de las que su regla permite en el bloque `architecture` de `docs/architecture.md`.

## Lista de comprobación del contribuidor

- [ ] El cambio respeta los 5 gates del workflow (`check_branch_name`, `check_pr_size`, `check_workflows`, ruff, mypy) — ver `docs/gates.md` §2.
- [ ] Si toca `app/`, `tests/` u `openspec/`, leyó `docs/decisions/README.md` antes de empezar.
- [ ] Si toca `docs/`, leyó `skills/documentation-alan-style/SKILL.md` y aplica tono castellano peninsular formal con usted.
- [ ] El commit usa Conventional Commits con el scope `(app)` o `(platform)` correspondiente.
- [ ] El PR es ≤ 400 líneas (`additions + deletions`); si no, parte por unidad de trabajo o encadena.

## Siguiente paso

Continúe con [Mapa del repositorio](repository-map.md) para localizar el paquete que owns el cambio.
