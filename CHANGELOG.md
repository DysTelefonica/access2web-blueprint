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
- **W24 coverage-gate paths** (PR #453): el plugin CRITICAL_HELPER apuntaba a los nombres ``hash_password`` / ``verify_password`` y a módulos nunca materializados (``application.credential_helpers`` / ``application.auth_reset``). Tras W07 (#434) el adapter Argon2id expone ``hash`` / ``verify``; los dos reset-flow services viven en ``domain/services/``. El plugin actualiza los nombres al nuevo API y los ``TARGET_MODULES`` a las tres rutas reales. El orchestrator sigue en WARNING hoy porque las ``hash`` / ``verify`` son métodos de clase, no free functions, y la lookup ``getattr`` no las encuentra. Un WU futuro que las exponga a nivel de módulo cerrará la grieta.
- **W25 coverage-gate helper collapse** (PR #454): ``_module_covered_lines`` y ``_module_total_executable`` compartían el mismo file-by-suffix lookup preamble. Extracto de ``_file_for_module`` que devuelve el path medido; ambos call sites colapsan al post-lookup step solamente. Misma entrada → misma salida; los 40 tests contract siguen en PASS.

### Fixed

- **W26 stale ``crap`` gate in smoke test** (PR #455): ``test_quality_report_produces_json_envelope`` pineaba un set de siete gates que incluían ``crap``, ausente de ``scripts/quality_report.py::GATES`` desde la limpieza DA-13 (el runner CRAP nunca se cableó; ``scripts/check_crap.py`` ya no está en disco). El test fallaba en cada run desde el rebase que aterrizó la limpieza. El assertion se actualiza al set real de seis gates (``layers``, ``complexity``, ``mutation_sites``, ``dry``, ``legacy_hashes``, ``legacy_retirement``). **El repo entero queda verde por primera vez desde el rebase**.

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