# Changelog

Todos los cambios relevantes del blueprint se documentan aquí. El formato sigue [Keep a Changelog](https://keepachangelog.com/) y el proyecto respeta [Semantic Versioning](https://semver.org/).

## Where to Find Release Notes

- [GitHub Releases](https://github.com/DysTelefonica/access2web-blueprint/releases) — notas detalladas por tag

## [Unreleased]

### Added

- **Async session factory** (#427, DA-1): `async_session_factory(url)` produce `AsyncEngine` + `AsyncSessionFactoryPort` con `search_path = lanzadera,public`. Es el seam único entre plataforma y Postgres.
- **W01 user/app/profile Postgres adapters** (#427): `UserRepositoryPg`, `AppRepositoryPg`, `ProfileRepositoryPg` implementan los Protocols de dominio. Mapean `users`/`apps`/`profiles` (migración 0001) a los dataclasses de dominio.
- **W02 AssignmentRepositoryPg** (#428, DA-12/D22/H11): cuatro métodos (create, list_for_user, list_for_app, effective_permissions). `effective_permissions` JOIN `user_app_assignments` con `profiles.capabilities` JSONB.
- **W03 AuditLogPg** (#429, DA-11/D27/D55): dos métodos (append, list_for_actor). Append-only, sin columnas de telemetría (DA-13 + D55).
- **W04 GlobalAdminRepositoryPg** (#430, DA-1/D21/D42): cuatro métodos (list_all, is_global_admin, grant, revoke). El invariant D42 (al menos un global admin) se enforce dentro de `revoke` antes del DELETE.
- **W05 ResetTokenRepositoryPg** (#432, DA-4/D90): cinco métodos (insert, find_unused, mark_consumed, mark_superseded, purge_expired). Single-use, bounded-TTL, supersedable.
- **W06 MailQueueTableAdapter** (#433, DA-10/D11): un método (send). Encola una fila en `lanzadera.mail_outbox` con `status='pending'`. El dispatcher real es D65 y queda fuera del alcance de esta entrega.
- **W07 Argon2id async API** (#434, DA-2): `CredentialHasherArgon2id.hash/verify` migran a `async def`. El KDF se delega a `asyncio.to_thread` para no bloquear el event loop.
- **W08 read-only session helper** (#436): `AsyncSessionFactory.read_only_session()` añade el context-manager para paths de lectura. Cierra el boilerplate `try/finally` que se repite 27 veces en los adapters.
- **W09..W13 read-only migrations** (#437, #438, #439, #440, #441): los read paths de `AppRepositoryPg`, `ProfileRepositoryPg`, `AssignmentRepositoryPg`, `AuditLogPg`, `GlobalAdminRepositoryPg` migran al helper.
- **W14 ResetTokenRepositoryPg.find_unused** (#442): método de lectura migrado al helper.
- **W15 MailQueueTableAdapter.send → transaction()** (#443): la única vía de escritura migra al context-manager `transaction()` que commit-once + rollback-on-exception.
- **W16 UserRepositoryPg writes → transaction()** (#444): tres métodos (create, update_status, update_password_and_activate). `list_all` migra a `read_only_session()` en el mismo PR.
- **W17 ResetTokenRepositoryPg writes** (#446): cuatro métodos (insert, mark_consumed, mark_superseded, purge_expired) migran a `transaction()`. El INSERT + 2nd SELECT viajan juntos en una sola transacción.
- **W18 AssignmentRepositoryPg.create** (#447): el INSERT + 2nd SELECT viven ahora dentro de una sola transacción (side benefit: cierra una ventana de orden donde el read corría post-commit).
- **W19 GlobalAdminRepositoryPg.grant + revoke** (#448): el SELECT count + DELETE de `revoke` se mantiene atómico dentro de `transaction()`. La verificación D42 sigue siendo previa al DELETE.
- **W20 AuditLog.append + ProfileRepositoryPg.create/set_active** (#449): los últimos 3 bloques `try/except/finally` manuales del W-series se consolidan. La deuda de boilerplate de los 8 Postgres adapters queda cerrada.
- **Async-everywhere en los auth-flow services** (#426): `PasswordHasher`, `UserRepository`, `ResetTokenRepository`, `GlobalAdminRepository`, `NotificationDelivery`, `AuditLog` migran a `async def`. `consume_reset_token` se vuelve async y propaga `await` por toda la cadena. Prereq para W01..W07.
- **CI install-step unblock** (#424): el bash-comment multi-línea con backticks en `.github/workflows/ci.yml` se cierra con el `#` que faltaba. El fallback `git show` + `git update-index` + `git checkout-index` para `app/pyproject.toml` aterriza con la corrección.
- **`pytest-asyncio==1.4.0`** (#424): pin en `[project.optional-dependencies].dev`. Pre-condición para que `asyncio_mode = "auto"` del `[tool.pytest.ini_options]` funcione.
- **CI mutation-sites BASELINE entries** (W01..W19): tres nuevas entradas (`user_repository_pg.py` 148, `app_repository_pg.py` 148, `profile_repository_pg.py` 117, `assignment_repository_pg.py` 204, `audit_log_pg.py` 108, `reset_token_repository_pg.py` 141). `target_date="2027-02-13"` consistente con `coverage_gate_helpers`.
- **DRY BASELINE cleanup**:
  - Entradas marcadas como no-longer-duplicated: `dup:1a1bacf15531`, `dup:25d63157feb1` (tras el post-ruff-format prelude único).
  - BASELINE restantes siguen activas porque el prelude-cleave de los 8 adapters está agendado para W21+: `3d62b08337d3` (W01), `25d63157feb1` con occurrences=3 (W02..W04).
  - **`3d62b08337d3`** en W01: prelude compartido en `UserRepositoryPg`, `AssignmentRepositoryPg`, `ProfileRepositoryPg`.
  - **`25d63157feb1`** con occurrences=3 en W02..W04: prelude compartido en `AssignmentRepositoryPg`, `AuditLogPg`, `GlobalAdminRepositoryPg`.
- **W21+22 retire `adapters/repos/` stub** (PR #450): el directorio ``app/src/modules/lanzadera/adapters/repos/__init__.py`` existía sólo como redirect transitorio para enlazar imports durante la W-series. Las redirect reales viven en ``adapters/__init__.py`` desde el PR W02..W05. Una vez consolidada la W-series (W08..W20), el subdirectorio stub queda sin función y se elimina. El bullet list del docstring de ``adapters/__init__.py`` queda sincronizado con la ruta canónica (``.persistence.*``) — antes seguía apuntando a la ruta legacy ``.repos.*`` que ya no existe.
- **W23 prelude-cleave pin** (PR #452): test arquitectónico que barre los 8 Postgres adapters (UserRepositoryPg, AppRepositoryPg, ProfileRepositoryPg, AssignmentRepositoryPg, AuditLogPg, GlobalAdminRepositoryPg, ResetTokenRepositoryPg, MailQueueTableAdapter) y falla si alguno regresa al shape manual ``session: AsyncSession = self._factory()``. La W-series (W08..W20) pagó el coste de la migración; esta gate evita la regresión. El test pasa hoy y rompe de forma determinista cuando alguien introduce el patrón pre-W08.
- **W24 coverage-gate paths** (PR #453): el plugin CRITICAL_HELPER apuntaba a los nombres ``hash_password`` / ``verify_password`` y a módulos nunca materializados (``application.credential_helpers`` / ``application.auth_reset``). Tras W07 (#434) el adapter Argon2id expone ``hash`` / ``verify``; los dos reset-flow services viven en ``domain/services/``. El plugin actualiza los nombres al nuevo API y los ``TARGET_MODULES`` a las tres rutas reales. El orchestrator quedaba en WARNING porque las ``hash`` / ``verify`` son métodos de clase, no free functions — esa grieta la cerró **W28 (#458)** vía ``_resolve_helper`` con owner no-None, no vía wrappers de módulo ficticios (la rama "expone a nivel de módulo" fue descartada por W28 explícitamente).
- **W25 coverage-gate helper collapse** (PR #454): ``_module_covered_lines`` y ``_module_total_executable`` compartían el mismo file-by-suffix lookup preamble. Extracto de ``_file_for_module`` que devuelve el path medido; ambos call sites colapsan al post-lookup step solamente. Misma entrada → misma salida; los 40 tests contract siguen en PASS.
- **W27 `_file_for_module` contract tests** (PR #456): cuatro casos fijan el short-circuit sin cobertura, el match por suffix, el match por substring y la ausencia de coincidencia del lookup extraído en W25.

### Fixed

- **W26 stale ``crap`` gate in smoke test** (PR #455): ``test_quality_report_produces_json_envelope`` pineaba un set de siete gates que incluían ``crap``, ausente de ``scripts/quality_report.py::GATES`` desde la limpieza DA-13 (el runner CRAP nunca se cableó; ``scripts/check_crap.py`` ya no está en disco). El assertion se actualiza al set real de seis gates (``layers``, ``complexity``, ``mutation_sites``, ``dry``, ``legacy_hashes``, ``legacy_retirement``). **El repo entero queda verde por primera vez desde el rebase**.
- **W28 targets exactos de QC-5** (#457): `coverage_gate.py` evalúa `CredentialHasherArgon2id.hash`, `CredentialHasherArgon2id.verify`, `issue_reset_token` y `consume_reset_token` sin producto cartesiano. Los métodos se resuelven sobre la clase mediante `_resolve_helper`; la infracobertura mantiene `session.exitstatus = 1`.
- **W29 lock-in mutation-sites BASELINE** (#459): tras el prelude-cleave de los 8 adapters (W08..W20, pin en W23) los cinco Postgres adapters con NOTE activo quedan bajo el ratchet: `user_repository_pg.py` 148→145, `assignment_repository_pg.py` 204→203, `audit_log_pg.py` 108→107, `profile_repository_pg.py` 117→115, `reset_token_repository_pg.py` 141→137. `app_repository_pg.py` queda en 148 (sin NOTE, conteo actual). `target=100` y `target_date="2027-02-13"` se preservan para mantener la ratchet al ceiling de mutation-sites. El check corre limpio — sin los cinco NOTES previos.
- **W30 remove stale DRY BASELINE entry** (#460): el grupo `dup:1aa7b9df3f16` (antes `coverage_gate.py` con dos bloques duplicados) ya no aparece en el reporte de `check_dry.py`; la BASELINE entry se retira para no mantener un ratchet sin contenido (Hard Rule 12: ratchet sin destino es rampa a ningún lado). El historial del grupo queda en git blame sobre el archivo. Las entries `3d62b08337d3` y `25d63157feb1` quedaban activas en este punto porque los preludes de los adapters Postgres seguían duplicados — el prelude-cleave pineó el comportamiento futuro pero no eliminó las duplicaciones. Esa deuda la cerró **W32 (#467)** extrayendo el prelude a `_pg_imports.py`, que retiró ambas BASELINE entries en el mismo PR.
- **W31 live class-method coverage e2e** (#464): dos casos en `tests/lanzadera/test_coverage_gate_class_methods_e2e.py` ejecutan `pytest -p app.pytest_plugin.coverage_gate --cov=<module>` como subprocess con `CRITICAL_TARGETS` cuyo owner es el nombre de una clase (`Hasher`). El caso FAIL ejercita sólo la rama feliz de `hash`/`verify` y exige `returncode == 1` + `coverage_gate FAIL` en stderr. El caso SILENT ejercita las cuatro ramas (feliz + empty) y exige `returncode == 0` sin FAIL. Cierra el gap entre `_resolve_helper` mockeado y pytest-cov real para class methods: el live test preexistente (`test_live_pytest_cov_fails_for_undercovered_exact_target`) sólo cubre owner=None.
- **W32 extract SQLAlchemy prelude** (#467): el bloque de 5 statements `from collections.abc import Sequence` + `from typing import Any` + `import sqlalchemy as sa` + `from sqlalchemy import select` + `from async_session_factory import SCHEMA, AsyncSessionFactoryPort` que los 7 Postgres adapters duplicaban se iza a `_pg_imports.py`. Cada adapter colapsa a un único statement multi-línea `from _pg_imports import (...)`. Los dialect imports (CITEXT, JSONB, PGUUID) siguen per-adapter; `from __future__ import annotations` también (Python lo requiere file-scoped). Las dos BASELINE entries del DRY ratchet (`dup:3d62b08337d3` occurrences=2 y `dup:25d63157feb1` occurrences=3) se retiran — Hard Rule 12 dice que un ratchet sin contenido es rampa a ningún lado. El prelude-cleave pin de W23 sigue verde: el patrón runtime que pinea (`session: AsyncSession = self._factory()`) no se ve afectado por el refactor de imports.
- **W33 CHANGELOG doc freshness** (#469): las entradas W24 y W30 describían estados pre-W28 / pre-W32 que ya no aplican. W24 decía "El orchestrator sigue en WARNING hoy porque las ``hash`` / ``verify`` son métodos de clase" — la grieta la cerró W28 vía ``_resolve_helper``. W30 decía "Las entries ``3d62b08337d3`` y ``25d63157feb1`` siguen activas porque los preludes ... siguen duplicados" — la duplicación la eliminó W32. Ambas entradas se reescriben para reflejar la realidad actual sin contradecir W28/W32; el contenido histórico (qué cambió en cada WU) se preserva verbatim.
- **W34 extract domain prelude** (#471): las cinco entidades de dominio (`assignment`, `audit_event`, `profile`, `reset_token`, `session`) compartían el bloque de 4 statements `from __future__ import annotations` + `from dataclasses import dataclass` + `from datetime import datetime` + `from uuid import UUID`. W32 (#467) cerró el patrón análogo en los Postgres adapters extrayendo a `_pg_imports.py`; W34 replica la limpieza en `app/src/modules/lanzadera/domain/_imports.py`. Cada entidad colapsa a un único statement multi-línea. La BASELINE entry DRY `dup:79d8c5976f0f` (occurrences=5) se retira. Las BASELINE restantes (`110a87c35376`, `08c289bbf440`, `ce4650821098`, `af8e0d12856d`) son docstring patterns, no imports — quedan para futuros WUs.
- **W35 extract platform_user CLI helper** (#473): los comandos `cmd_set_password`, `cmd_grant_global_admin` y `cmd_revoke_global_admin` en `app/src/modules/lanzadera/delivery/cli/platform_user.py` repetían el mismo bloque de 5 statements: `email.strip().lower()` + `users.get_by_email(...)` + check user-not-found + prompt confirmation + check cancelled. W32/W34 cerraron el mismo patrón para imports; W35 lo cierra para lógica de comando. Helper privado `_resolve_user_with_confirmation(email, users, *, action, description, cancellation_message, confirmed) -> User | CommandResult` ejecuta los 5 statements comunes. Cada comando colapsa a una llamada al helper + el post-resolve específico (leer password / grant / revoke). Los mensajes de éxito pasan a usar `user.email` (canónico de la DB) en lugar del input normalizado — los tests verifican substrings, no el formato exacto, así que la API observable no rompe. BASELINE entry DRY `dup:ce4650821098` (occurrences=3) se retira.
- **W36 add StrEnum to domain prelude** (#475): `app.py` y `user.py` comparten el mismo bloque de 5 statements que el resto de las entidades (W34 cerró 5 de 7 entidades), pero usan `StrEnum` además de `dataclass/datetime`. W36 añade `StrEnum` al `__all__` de `app/src/modules/lanzadera/domain/_imports.py` y migra `app.py` + `user.py` al mismo patrón multi-línea `from _imports import (...)`. `user.py` también importa `UUID` desde el mismo módulo (ya estaba en `__all__`). BASELINE entry DRY `dup:af8e0d12856d` (occurrences=2) se retira. Las dos BASELINE restantes (`08c289bbf440` ports y `110a87c35376` admin/app/user) son docstring patterns en archivos distintos — quedan para futuros WUs.
- **W37 extract ports prelude** (#477): los puertos `AppRepositoryPort` y `AssignmentRepositoryPort` compartían el bloque de 4 imports `from __future__ import annotations` + `from collections.abc import Sequence` + `from typing import Protocol` + `from uuid import UUID` (más el docstring, que el normaliser del check reescribe a `str` para el hash — la duplicación es estructural aunque los textos difieran). W37 extrae los 3 imports comunes a `app/src/modules/lanzadera/domain/ports/_imports.py`. Cada puerto colapsa a un único statement multi-línea. BASELINE entry DRY `dup:08c289bbf440` (occurrences=2) se retira. La BASELINE restante `dup:110a87c35376` (occurrences=3, dataclass fields en admin/app/user) es field-shape coincidence, no fix estructural viable — queda como BASELINE permanente.
- **W38 lock-in mutation-sites BASELINE for platform_user** (#479): W35 (#473) extrajo el helper de lookup+confirmación; la extracción bajó el conteo de sitios de `platform_user.py` de 234 a 231. BASELINE entry se baja de 234→231 para lock-in. Mismo target=100 y `target_date="2027-02-13"` se preservan para mantener la ratchet al ceiling de mutation-sites. El check corre limpio — sin el NOTE previo.

- **`coverage_gate` plugin (#424)**: el paso `install` del job `quality` ahora materializa `app/pyproject.toml` con el triple `git show` + `git update-index` + `git checkout-index`. La etapa anterior fallaba porque la copia de trabajo del runner self-hosted tenía la ruta en skip-worktree state.

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
