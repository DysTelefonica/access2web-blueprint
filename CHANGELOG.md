# Changelog

Todos los cambios relevantes del blueprint se documentan aquí. El formato sigue [Keep a Changelog](https://keepachangelog.com/) y el proyecto respeta [Semantic Versioning](https://semver.org/).

## Where to Find Release Notes

- [GitHub Releases](https://github.com/DysTelefonica/access2web-blueprint/releases) — notas detalladas por tag

## [Unreleased]

Vacío. Las nuevas entregas se documentan aquí antes del siguiente release cut.

### Changed

- **W44 split coverage_gate_helpers.py** (#488): el archivo `app/pytest_plugin/coverage_gate_helpers.py` (127 sitios, BASELINE sobre el ceiling) se rompe en tres módulos cohesivos: `coverage_gate_coverage.py` (3 lookup helpers, 42 sitios), `coverage_gate_messages.py` (4 verdict builders, 26 sitios), `coverage_gate_resolution.py` (3 resolution+eval helpers, 61 sitios). Cada uno queda bien bajo el ceiling de 100 sitios, así que no requiere BASELINE entries propias. `coverage_gate.py` actualiza imports para usar los 3 módulos directamente. `tests/lanzadera/test_coverage_gate_plugin.py` actualiza los `importlib.import_module(...)` dinámicos y los accesos `_file_for_module` / `_module_covered_lines` / etc. al módulo correcto. `coverage_gate_helpers.py` desaparece; la BASELINE entry se cierra. El orchestrator en `coverage_gate.py` sigue en 109 sitios — fuera de scope de W44 (requiere un restructure de la lógica del loop).
- **W45 extract orchestrator to coverage_gate_evaluate** (#490): el orchestrator `_enforce_critical_coverage` (52 líneas) + el helper `_resolve_target` (15 líneas) se mueven a `app/pytest_plugin/coverage_gate_evaluate.py`. `coverage_gate.py` queda con solo el entry point: `_try_import` + `_emit_verdicts` + el wrapper shim + el pytest hook + StashKey + fixture (~50 sitios). `coverage_gate_evaluate.py` tiene el orchestrator real (~62 sitios). El shim preserva la firma original `(session, coverage_data)` para compatibilidad con los tests existentes. BASELINE entry de `coverage_gate.py` se cierra. El `_try_import` se mantiene en `coverage_gate.py` y `coverage_gate_evaluate.py` lo importa lazy para que los monkey-patches de tests sigan funcionando.
- **W46 split platform_user commands** (#492): `delivery/cli/platform_user.py` (231 sitios) se rompe en tres módulos cohesivos + uno compartido: `platform_user_types.py` (47 sitios) con `CommandResult`, `ConfirmableDestructiveCommandError`, los Protocols `_AppRepositoryPort`/`_ProfileRepositoryPort`, y los helpers `_prompt_confirmation`/`_read_password`/`_resolve_user_with_confirmation`; `platform_user_auth.py` (89 sitios) con `cmd_set_password`, `cmd_grant_global_admin`, `cmd_revoke_global_admin`; `platform_user_apps.py` (58 sitios) con `cmd_list_apps`, `cmd_assign_profile`. `platform_user.py` (52 sitios) queda como entry point que re-exporta los command callables. Los split modules importan `_prompt_confirmation` y `_read_password` via `_platform_user.<name>` para que los monkey-patches de tests (`monkeypatch.setattr(platform_user, "_read_password", ...)`) sigan propagando. BASELINE entry de `platform_user.py` se cierra.

## [0.3.0] - 2026-08-24

Cierre de la W-series del Lanzadera MVP (#427..#449): los 8 Postgres adapters se migran de su forma manual `try/except/finally` a los context-managers `AsyncSessionFactory.read_only_session()` y `transaction()`. La migración cierra 27 bloques boilerplate + añade un prelude-cleave pin (#452) que detecta regresiones futuras. El coverage gate (#453..#458) cierra la grieta de los class-methods en `CredentialHasherArgon2id.hash`/`.verify` vía `_resolve_helper` con owner no-None. Los WUs W21..W40 refactorizan preludes duplicados (`_pg_imports.py`, `_imports.py` en `domain/` y `domain/ports/`), bajan el helper de `platform_user.py`, y dejan el repo verde por primera vez: DRY ratchet 6 de 7 BASELINE entries cerradas + 1 permanente; mutation-sites ratchet lockeado al shape actual vía lock-in.

### Added

- **W01..W07 + W09..W20 Postgres adapter migration** (#427..#449): UserRepositoryPg, AppRepositoryPg, ProfileRepositoryPg, AssignmentRepositoryPg, AuditLogPg, GlobalAdminRepositoryPg, ResetTokenRepositoryPg, MailQueueTableAdapter. Cada adapter colapsa su sesión manual `try/except/finally` a uno de los dos context-managers del seam `AsyncSessionFactory`. Los métodos de lectura van a `read_only_session()`; los de escritura a `transaction()` (commit-once + rollback-on-exception). DA-11 (atomicidad) queda enforced por el helper en lugar de por cada adapter individualmente.
- **Async-everywhere en los auth-flow services** (#426): `PasswordHasher`, `UserRepository`, `ResetTokenRepository`, `GlobalAdminRepository`, `NotificationDelivery`, `AuditLog` migran a `async def`. `consume_reset_token` se vuelve async y propaga `await` por toda la cadena. Prereq para W01..W07.
- **Async session factory** (#427, DA-1): `async_session_factory(url)` produce `AsyncEngine` + `AsyncSessionFactoryPort` con `search_path = lanzadera,public`. Es el seam único entre plataforma y Postgres.
- **CI install-step unblock** (#424): el bash-comment multi-línea con backticks en `.github/workflows/ci.yml` se cierra con el `#` que faltaba. El fallback `git show` + `git update-index` + `git checkout-index` para `app/pyproject.toml` aterriza con la corrección.
- **`pytest-asyncio==1.4.0`** (#424): pin en `[project.optional-dependencies].dev`. Pre-condición para que `asyncio_mode = "auto"` del `[tool.pytest.ini_options]` funcione.
- **CI mutation-sites BASELINE entries** (W01..W19): seis nuevas entradas para los adapters Postgres con `target_date="2027-02-13"` consistente con `coverage_gate_helpers`. Las entries se lockean al shape actual vía W29 + W38.
- **W21+22 retire `adapters/repos/` stub** (PR #450): el directorio ``app/src/modules/lanzadera/adapters/repos/__init__.py`` existía sólo como redirect transitorio para enlazar imports durante la W-series. Las redirect reales viven en ``adapters/__init__.py`` desde el PR W02..W05. Una vez consolidada la W-series (W08..W20), el subdirectorio stub queda sin función y se elimina. El bullet list del docstring de ``adapters/__init__.py`` queda sincronizado con la ruta canónica (``.persistence.*``) — antes seguía apuntando a la ruta legacy ``.repos.*`` que ya no existe.
- **W23 prelude-cleave pin** (PR #452): test arquitectónico que barre los 8 Postgres adapters y falla si alguno regresa al shape manual ``session: AsyncSession = self._factory()``. La W-series (W08..W20) pagó el coste de la migración; esta gate evita la regresión. El test pasa hoy y rompe de forma determinista cuando alguien introduce el patrón pre-W08.
- **W24 coverage-gate paths** (PR #453): el plugin CRITICAL_HELPER apuntaba a los nombres ``hash_password`` / ``verify_password`` y a módulos nunca materializados. El plugin actualiza los nombres al nuevo API (`hash`/`verify` en `CredentialHasherArgon2id`, `issue_reset_token`/`consume_reset_token` en `domain/services/`) y los `TARGET_MODULES` a las tres rutas reales. La grieta de class-methods la cerró W28 (#458) vía `_resolve_helper` con owner no-None.
- **W25 coverage-gate helper collapse** (PR #454): `_module_covered_lines` y `_module_total_executable` compartían el mismo file-by-suffix lookup preamble. Extracto de `_file_for_module` que devuelve el path medido; ambos call sites colapsan al post-lookup step solamente.
- **W27 `_file_for_module` contract tests** (PR #456): cuatro casos fijan el short-circuit sin cobertura, el match por suffix, el match por substring y la ausencia de coincidencia del lookup extraído en W25.
- **W32 extract SQLAlchemy prelude** (#467): el bloque de 5 statements que los 7 Postgres adapters duplicaban (`Sequence`, `Any`, `sa`, `select`, `SCHEMA`, `AsyncSessionFactoryPort`) se iza a `app/src/modules/lanzadera/adapters/persistence/repositories/_pg_imports.py`. Cada adapter colapsa a un único statement multi-línea `from _pg_imports import (...)`. Los dialect imports (CITEXT, JSONB, PGUUID) y `from __future__ import annotations` siguen per-adapter / file-scoped.
- **W34 extract domain prelude** (#471): las cinco entidades de dominio (`assignment`, `audit_event`, `profile`, `reset_token`, `session`) compartían el mismo bloque de 4 imports (`Sequence`, `Any`, `UUID`). Mismo patrón de extracción — `app/src/modules/lanzadera/domain/_imports.py`.
- **W36 add StrEnum to domain prelude** (#475): `app.py` y `user.py` (que declaran `AppTopology` y `UserStatus`) extienden el dominio prelude con `StrEnum`.
- **W37 extract ports prelude** (#477): los puertos `AppRepositoryPort` y `AssignmentRepositoryPort` compartían el bloque de 3 imports (`Protocol`, `Sequence`, `UUID`). Mismo patrón — `app/src/modules/lanzadera/domain/ports/_imports.py`.
- **W31 live class-method coverage e2e** (#464): dos casos en `tests/lanzadera/test_coverage_gate_class_methods_e2e.py` ejecutan `pytest -p app.pytest_plugin.coverage_gate --cov=<module>` como subprocess con `CRITICAL_TARGETS` cuyo owner es el nombre de una clase (`Hasher`). Cierra el gap entre `_resolve_helper` mockeado y pytest-cov real para class methods.
- **W33 CHANGELOG doc freshness** (#469): las entradas W24 y W30 describían estados pre-W28 / pre-W32 que ya no aplican; se reescriben sin contradecir las entradas cerradas por W28/W32.

### Changed

- **W08 read-only session helper** (PR #436): `AsyncSessionFactory.read_only_session()` añade el context-manager para paths de lectura. Cierra el boilerplate `try/finally` que se repite 27 veces en los adapters.
- **DRY BASELINE cleanup** (W30, W32, W34, W35, W36, W37, W40): las 7 BASELINE entries originales (`1aa7b9df3f16`, `3d62b08337d3`, `25d63157feb1`, `79d8c5976f0f`, `ce4650821098`, `af8e0d12856d`, `08c289bbf440`) se cierran progresivamente conforme las duplicaciones se eliminan vía W32/W34/W35/W36/W37. La única restante (`110a87c35376`, occurrences=3, dataclass-field shape coincidence en admin/app/user) se vuelve permanente vía W40 — `target=3`, `target_date=2030-01-01`.
- **mutation-sites BASELINE lock-ins** (W29, W38): los cinco Postgres adapters con NOTE activo (W23 prelude-cleave redujo el conteo) y `platform_user.py` (W35 helper extraction) se lockean al shape actual sin relajar la ratchet hacia `target=100` con `target_date=2027-02-13`.
- **`coverage_gate` plugin (#424)**: el paso `install` del job `quality` ahora materializa `app/pyproject.toml` con el triple `git show` + `git update-index` + `git checkout-index`.

### Fixed

- **W26 stale `crap` gate in smoke test** (PR #455): `test_quality_report_produces_json_envelope` pineaba un set de siete gates que incluían `crap`, ausente de `scripts/quality_report.py::GATES` desde la limpieza DA-13. El assertion se actualiza al set real de seis gates (`layers`, `complexity`, `mutation_sites`, `dry`, `legacy_hashes`, `legacy_retirement`).
- **W28 targets exactos de QC-5** (#457): `coverage_gate.py` evalúa `CredentialHasherArgon2id.hash`, `CredentialHasherArgon2id.verify`, `issue_reset_token` y `consume_reset_token` sin producto cartesiano. Los métodos de clase se resuelven sobre la clase mediante `_resolve_helper`; la infracobertura mantiene `session.exitstatus = 1`.
- **W29 + W30 + W32..W38 + W40 + W41 BASELINE cleanups** (#459, #460, #467, #469, #471, #473, #475, #477, #479, #482): ratchets que protegen duplicaciones que ya no existen se cierran o se vuelven permanentes. El check `check_dry.py` y el check `check_mutation_sites.py` corren limpios — sin NOTES ni FAILs.

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