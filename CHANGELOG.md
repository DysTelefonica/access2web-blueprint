# Changelog

Todos los cambios relevantes del blueprint se documentan aquí. El formato sigue [Keep a Changelog](https://keepachangelog.com/) y el proyecto respeta [Semantic Versioning](https://semver.org/).

## Where to Find Release Notes

- [GitHub Releases](https://github.com/DysTelefonica/access2web-blueprint/releases) — notas detalladas por tag

## [Unreleased]

### Added

- `docs(root): aplicar las plantillas de documentation-alan-style v2.1 a los 5 docs raíz (README, AGENTS, DOCS, CODEBASE-GUIDE, CHANGELOG). Cierra los issues #325, #321, #322, #324 y #323.`

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

[← Back to README](README.md) · [Next: DOCS.md →](DOCS.md)