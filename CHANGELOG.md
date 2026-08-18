# Changelog

Todos los cambios relevantes del blueprint se documentan aquí. El formato sigue [Keep a Changelog](https://keepachangelog.com/) y el proyecto respeta [Semantic Versioning](https://semver.org/).

## Where to Find Release Notes

- [GitHub Releases](https://github.com/DysTelefonica/access2web-blueprint/releases) — notas detalladas por tag

## [Unreleased]

## [0.2.0] - 2026-08-18

Aplicación completa de `documentation-alan-style` v2.1 al monorepo (5 root docs, 3 guías de `docs/`, 8 READMs por app, 8 épicas, walkthrough template + validator + migration). El skill prescribe plantillas para README, AGENTS, DOCS, CODEBASE-GUIDE, CHANGELOG, walkthrough.json y epic.md; antes de esta versión, sólo los 5 root docs tenían su plantilla aplicada parcialmente. Esta versión cierra la brecha en todas las superficies documentales del proyecto.

### Added

- **5 root docs** (#326-#330): plantillas v2.1 aplicadas a README.md (hero + badges + Quick start arriba + Documentation table + Next steps + License), AGENTS.md (frontmatter YAML + tabla estandarizada + nota Alan-omits), CODEBASE-GUIDE.md (frontmatter + Core invariants + Existing references + Contributor checklist + Navigation), DOCS.md (What this is/is not + Status Visibility Matrix + Contributor checklist + Back link), CHANGELOG.md (entrada `[0.1.0]` inicial con sub-secciones Keep a Changelog).
- **3 guías de `docs/`** (#334-#336): patrones de sección aplicados a docs/architecture.md (6 invariantes), docs/calidad-de-codigo-y-ci.md (5 invariantes + lista §Los 13 check_*.py), docs/AGENT-SETUP.md (back link + 4 invariantes + 5 checks).
- **8 READMs por app** (#345-#352): back link + Core invariants + Contributor checklist + Navigation por app — Lanzadera (D5/D55/DA-11), NoConformidades (D88/D91/D95), Gestion_Riesgos (D86/D88/D95), Brass (D104), HPS (D92), HPS_Solicitudes (D86/D92), Condor (D93), Expedientes (D86/D87/D94).
- **8 épicas** (#361-#368): sección «Cómo se aplica a access2web-blueprint» + «Lista de comprobación final» del epic.md.tmpl añadidas al final de cada épica, conservando estructura existente (migración híbrida preservadora).
- **walkthrough.json.tmpl ampliado** (#372): clasificación MUST/SHOULD/MAY por campo en `skills/documentation-alan-style/references/templates/walkthrough.json.tmpl`, con Schema history v1.0 → v2.1+. Inclusión de campos evolucionados: `unattended_evidence`, `codegraph_summary` como objeto, `known_working_tools`, `known_degraded_tools`, `known_broken_tools`, `generated_at`, `app`, `project`.
- **scripts/check_walkthrough_schema.py** (#372): nuevo gate análogo a los 12 check_*.py existentes. Itera `docs/03-aplicaciones/*/walkthrough-*.json` y valida que los MUST fields estén presentes. Modo advisory por defecto (exit 0 con findings); `--strict` para enforzar (exit 1).
- **scripts/migrate_walkthrough_schema.py** (#376): one-shot idempotente que añade MUST fields faltantes a los 43 walkthroughs reales. Maneja 4 variantes de schema: camelCase (v2.1), snake_case legacy (form_name, source_path), tercera variante (id, name), sub-form (subfrm en formName).
- **13 quality gates** (#372, #376): tras añadir `check_walkthrough_schema.py`, `scripts/check_*.py` pasa de 12 a 13. El gate corre en `ci.yml` como paso adicional del job `review-budget`, invocado con `--strict` tras la migración.
- **docs/03-aplicaciones/<app>/walkthrough-*.json** (#376): 26 walkthroughs no conformes migrados al schema v2.1 (767 field additions); 17 ya conformes. El validator pasa con `--strict` sobre los 43 archivos.

### Changed

- `.github/workflows/ci.yml` (#372): añade paso `walkthrough schema` al job `review-budget`. Inicialmente en modo advisory (#372); promovido a `--strict` (#376) tras la migración.
- `docs/calidad-de-codigo-y-ci.md` (#372): §Los 12 check_*.py → §Los 13 check_*.py con la entrada `check_walkthrough_schema.py` añadida.
- `docs/architecture.md` (#372): tabla de CI gates actualizada con la entrada del nuevo gate.

### Fixed

- `scripts/check_walkthrough_schema.py` (#372): formato ruff aplicado (`indent=2`, line-length 100). Sin el formato, el CI fallaba en el paso `format`.

## [0.1.0] - 2026-08-18

### Added

- `docs(root): nueva sección de skills obligatorias con columna «Obligatorio para».`
- `docs(repo): añadir docs/architecture.md como fuente de verdad única arquitectónica (#299).`
- `docs(repo): añadir docs/calidad-de-codigo-y-ci.md con los 12 quality gates y los 4 workflows (#304).`
- `docs(repo): formalizar la convención de naming multi-app para issues y commits (#295).`
- `docs(repo): crear CONTRIBUTING.md, CHANGELOG.md y docs/AGENT-SETUP.md como esqueletos iniciales.`
- `docs(repo): internalizar las skills del proyecto bajo skills/ (#167, en cadena).`
- `chore(skills): añadir architecture-guardrails skill con front-door a docs/architecture.md (#307).`
- `chore(skills): actualizar documentation-alan-style a v2.1 (#320).`
- `feat(release): publicar por tags con gate de identidad (#160).`
- `docs(expedientes): publish migration planning status; definir tasks de review; mapear capabilities a delivery issues; especificar security + legacy retirement; E2E exchange; catalogs + query flows; lifecycle + related data; migration scope + architecture (#292, #291, #290, #289, #287, #286, #285, #284).`
- `docs(repo): consolidar en main la cadena de PRs de contribución y skills (#265).`
- `docs(repo): documentar el presupuesto de PR y los PRs encadenados (#164).`

### Changed

- `docs(root): traducir headings y procedimientos de AGENTS.md al Castellano peninsular formal.`
- `fix(ci): set fetch-depth: 0 en los steps de quality y mutation checkout (#314).`
- `docs(openspec): corregir drift platform.* → app.* en lanzadera-mvp/design.md (post-rename 2026-08-09) (#306).`

### Fixed

- `docs(root): reemplazar tuteo y voseo por Castellano peninsular formal con usted en AGENTS.md, CODEBASE-GUIDE.md. (§4 de documentation-alan-style)`
- `docs(root): corregir anchor roto en DOCS.md sección «Diseño UI/UX».`
- `docs(root): agregar navegación «Next» al pie de DOCS.md.`
- `docs(root): corregir code fence del smoke test en README.md de js a bash.`
- `docs(root): reparar tipografía «oContributor» en CODEBASE-GUIDE.md.`
- `fix(docs): corregir drift platform/ → app/ en CODEBASE-GUIDE y DOCS; documentar inputs/ (#303).`
- `test: drop contradictory assertion in test_quality_report_produces_json_envelope (#297).`

### Deprecated

- `docs(workflows): rama `00_main/` + `.git` FILE (layout v1) — reemplazado por main en raíz + `<project>-worktrees/` (v2.0).`

## Breaking Changes

| Versión anterior | Versión nueva | Migración |
|---|---|---|
| `platform/` (directorio pre-2026-08-09) | `app/` | Renombrar imports de `platform.src.modules.<app>` a `app.src.modules.<app>`. El rename ya está mergeado (#303, #306); los cambios futuros usan `app/`. |
| Layout v1 (`00_main/` + `.git` FILE) | Layout v2 (main en raíz + `<project>-worktrees/`) | Migrar worktrees siguiendo la skill `worktree-reorg-per-project`. |
| `walkthrough.json` schema v1 (template mínimo, 12 campos) | Schema v2.1 (MUST/SHOULD/MAY, 14+ campos) | Migrar con `scripts/migrate_walkthrough_schema.py`. El gate `check_walkthrough_schema.py` (en `--strict`) bloquea nuevos walkthroughs no conformes. |

[← Back to README](README.md) · [Next: DOCS.md →](DOCS.md)