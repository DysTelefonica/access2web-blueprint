<!-- Change: lanzadera-mvp · Project: access2web-blueprint · Date: 2026-08-08 -->

# Propuesta: Lanzadera MVP — primer slice de la plataforma hexagonal

> **Frase que organiza**: Lanzadera deja de ser el portal Access/VBA que concentraba identidad, catálogo y lanzamiento y se convierte en el módulo admin de una plataforma web hexagonal donde los usuarios, las aplicaciones y los permisos son la fuente de verdad del resto del ecosistema.

## Contexto

El blueprint `access2web-blueprint` documenta la modernización de las ocho aplicaciones legacy del ecosistema (D1). La pasada de discovery sobre Lanzadera (Lote 1, sesión 2026-08-04) confirma que Lanzadera es la **madre** del ecosistema: ahí nacen los 156 usuarios (`tbUsuarios`), las 20 aplicaciones del catálogo (`TbAplicaciones`) y las 622 asignaciones `usuario × aplicación × rol` (`TbUsuariosAplicacionesPermisos`, D85).

Tras las decisiones D5–D82 consolidadas en `docs/09-arquitectura-objetivo-y-principios.md` y las resoluciones de arquitectura de la sesión del 2026-08-08, el primer slice que sale del blueprint es la **Lanzadera admin en web hexagonal**: persiste identidad, catálogo y permisos en PostgreSQL, retira la mecánica desktop de Lanzadera Access (D51–D55) y mantiene coexistencia cero con el legacy hasta el UAT del ecosistema completo.

Esta propuesta NO describe la migración de los otros siete módulos (Expedientes, Gestion_Riesgos, No_Conformidades, HPS, HPS_Solicitudes, Condor, Brass). Cada uno tendrá su propio `change` SDD cuando llegue su turno.

## Alcance

### Incluye

- El módulo `lanzadera` dentro del monolito modular (`app/src/modules/lanzadera/`).
- Las migraciones 0001–0006 de Alembic: schema core, seed de apps, seed de profiles, migración de users con DNI cifrado, migración de assignments aplicando la regla de mapeo, seed de audit/histórico.
- El script `scripts/migrate_from_access.py` (one-shot) que extrae del `.accdb` con Dysflow y produce fixtures JSON.
- El `BootstrapAdapter` que siembra `global_admins` desde `GLOBAL_ADMIN_EMAILS` (idempotente al arranque del proceso).
- Los 12 quality gates del MVP (QC-1 a QC-9 + derivados), commiteados, wired en `ci.yml` y pinned por test.
- Las decisiones §Criterios de aceptación de esta propuesta.

### No incluye

- Coexistencia runtime con Lanzadera Access legacy. El `.accdb` sigue funcionando hasta el UAT del ecosistema completo (decisión de sesión 2026-08-08). El cut-over es one-shot al go-live.
- Migración de los otros siete módulos. Cada uno tiene su propio `change` SDD.
- Integración con SiteMinder / OCP / JWT (D10, FUTURO). Queda como adapter hexagonal diferido.
- SMTP corporativo real (D11). La cola por tabla actúa como adapter v1 (D13); el SMTP real se decide más adelante (P20).
- Diseño detallado del ciclo UAT (P16, ABIERTO). La gobernanza global-admin-only ya está aprobada (D48).
- Dashboard global de operaciones de notificación (P18, ABIERTO).
- Migración de vídeos, cuestionarios, visionados, ActiveX OCX y los 5 critical findings G4 del walkthrough (D51, D158, D159). Quedan en el legacy archivado, no migran al módulo operativo.
- Replicación automática de datos comunes al NAS oficina (gap P-21, aceptado no bloqueante, ver §Hallazgos y riesgos).
- Topología de despliegue definitiva on-premise / nube corporativa / OCP (D75, ABIERTO).

## Criterios de aceptación

### Funcionales

- [ ] **CA-F1** — El catálogo `apps` queda sembrado con las 20 aplicaciones de `TbAplicaciones` (D85), con `deployment_topology` y `requires_office_presence` correctos.
- [ ] **CA-F2** — Los 156 usuarios de `tbUsuarios` quedan migrados con DNI cifrado en reposo y `legacy_hash` preservado (D36). La verificación legacy funciona en el primer login.
- [ ] **CA-F3** — Las 622 filas de `TbUsuariosAplicacionesPermisos` quedan traducidas a `user_app_assignments` aplicando la regla de mapeo: `Administrador=Sí → ADMIN`, `Calidad=Sí → CALIDAD`, `CalidadAvisos=Sí → CALIDAD_AVISOS`, `Técnico=Sí → TECNICO`, `Economía=Sí → ECONOMIA`, `Secretaría=Sí → SECRETARIA`, `SinAcceso=Sí → SIN_ACCESO` exclusivo, todos No → `DEFAULT`, compuestos → varias filas. La regla `SinAcceso` es exclusiva: si está activa, se ignora el resto.
- [ ] **CA-F4** — El `BootstrapAdapter` siembra `global_admins` desde `GLOBAL_ADMIN_EMAILS` (separados por `;`) de forma idempotente al arranque del proceso; crea los `users` faltantes; registra auditoría; revocación por CLI.
- [ ] **CA-F5** — El menú global web muestra solo los módulos para los que el usuario tiene permiso (D5, D6).
- [ ] **CA-F6** — El lockout funciona según D38–D40: 5 intentos por defecto, 1 hora por defecto, notificación a admins globales en cada lockout, desbloqueo por admin global.
- [ ] **CA-F7** — El rehash transparente al primer login exitoso funciona atómicamente, es idempotente y queda auditado (D37).

### Seguridad

- [ ] **CA-S1** — Los hashes heredados se preservan como `users.legacy_hash` con `algorithm` versionado (D36). Los campos con credenciales en claro no migran como secretos.
- [ ] **CA-S2** — El campo `users.dni_encrypted` cifra el DNI en reposo; la clave vive en el secret manager, no en el repositorio.
- [ ] **CA-S3** — `pip-audit`, `gitleaks` y `trivy` corren en cada PR con imágenes pinned por digest y `exit-code 1` ante cualquier hallazgo (QC-7).
- [ ] **CA-S4** — Ningún secreto aparece en logs ni en argumentos CLI (D25).
- [ ] **CA-S5** — Solo el admin global accede al CLI; las acciones destructivas piden confirmación explícita adicional (D24–D26).

### Calidad de código y arquitectura

- [ ] **CA-Q1** — `scripts/check_layers.py` corre en CI y falla ante cualquier violación de `ALLOWED_IMPORTS`, `PURE_LAYERS` o vertical slicing (QC-2, QC-9).
- [ ] **CA-Q2** — `mypy` strict en `app/` con `enable_error_code = ["ignore-without-code"]` (QC-4).
- [ ] **CA-Q3** — `ruff==0.15.21` con `select = E,F,W,I,UP,B` corre en CI (QC-3).
- [ ] **CA-Q4** — Cobertura global ≥ 85 % y 100 % en `CRITICAL_HELPERS = ["hash_password", "verify_password", "issue_reset_token", "consume_reset_token"]` (QC-5, D88, D90).
- [ ] **CA-Q5** — Las PRs ≤ 400 líneas; las excepciones `size:exception` justifican la razón en el cuerpo (QC-6).
- [ ] **CA-Q6** — Los nombres de rama cumplen `^(feat|fix|refactor|docs|ci|test)/<n>-<slug>$` (QC-6).
- [ ] **CA-Q7** — Las migraciones siguen Expand and Contract (D82): ninguna migración destructiva en una sola release.
- [ ] **CA-Q8** — La complejidad ciclomática del top-10 de funciones no excede `CC ≤ 15` (QC-1 derivado).
- [ ] **CA-Q9** — Los 12 quality gates aparecen commiteados, wired en `ci.yml` y pinned por test (`tests/test_ci_workflow.py`).
- [ ] **CA-F7** — Las credenciales se persisten con Argon2id vía `argon2-cffi==25.1.0` perfil `RFC_9106_LOW_MEMORY` (Argon2id, 64 MiB, 3 iteraciones, 4 hilos). Ningún hash legacy se preserva.
- [ ] **CA-F8** — El reset flow emite tokens one-time de 24 h vía adapter de notificación con cola por tabla; el token se consume atómicamente y deja de ser válido tras el primer uso.
- [ ] **CA-F9** — El primer global admin se aprovisiona con `gentle-ai platform user set-password <email>` (CLI exclusivo de admin global, D25) antes de que el sistema pueda emitir tokens de reset por email.

### Operacionales

- [ ] **CA-O1** — `docker-compose.yml` levanta PostgreSQL + MinIO + el backend en local sin pasos manuales extra (D77).
- [ ] **CA-O2** — El versionado sigue semver por módulo y plataforma (`lanzadera/vX.Y.Z`, `platform/vX.Y.Z`) (D78, D80).
- [ ] **CA-O3** — Conventional Commits alimentan el changelog (D79, trunk-based).
- [ ] **CA-O4** — El catálogo de versiones compatibles entre `lanzadera` y `platform` se declara explícitamente (D80).

## Decisiones arquitectónicas heredadas

Las decisiones D5–D87 y QC-1 a QC-9 son el contrato que este cambio materializa. La tabla resume las más relevantes para el MVP y dónde se aplica cada una. El detalle vive en `docs/08-decisiones-y-preguntas-abiertas.md`, `docs/09-arquitectura-objetivo-y-principios.md` y `docs/calidad-de-codigo-y-ci.md`.

| # | Decisión | Razón | Aplicación en el MVP |
|---|---|---|---|
| D5 | Plataforma única modular permission-aware | Producto único, no microservicios prematuros | Módulo `lanzadera` dentro de `app/` |
| D6 | Navegación app-first anidada | El expediente no es punto de entrada | Menú global web con apps visibles según permisos |
| D7 | Lanzadera = solo administración | Concentrar admin en un módulo | Scope de este `change` |
| D8 | Hexagonal global | Capacidades detrás de puertos, no de mecanismos | `scripts/check_layers.py` enforce la pureza |
| D9 | Auth inicial email/password (adapter) | Reducir superficie de cut-over | Adapter de credenciales con `argon2-cffi==25.1.0`; reset flow con tokens one-time; sin compat con hashes legacy |
| D11 | Notificación unificada como servicio compartido | Una sola integración de canal | Cola por tabla como adapter v1 (D13) |
| D13 | Cola de correo por tabla como adapter transitorio | SMTP corporativo real aún sin decidir (P20) | Adapter v1 listo para sustituir |
| D14 | PostgreSQL compartido con esquemas por módulo | Aislar módulos sin multiplicar bases | Schema `lanzadera` para este módulo |
| D36 | Preservar hashes heredados | No forzar reset masivo al cut-over | **OBSOLETO 2026-08-08.** SHA256 sin salt no cumple estándares modernos (OWASP 2024). D88-D91 lo reemplazan |
| D37 | Rehash transparente al primer login | Política moderna sustituye al hash heredado | **OBSOLETO 2026-08-08.** Reemplazado por reset flow explícito (D90) |
| D88 | Auth usa Argon2id vía `argon2-cffi==25.1.0` con perfil `RFC_9106_LOW_MEMORY` | OWASP 2024 primera opción; memory-hard; resistente a GPU/ASIC | `argon2-cffi==25.1.0` pinned en `pyproject.toml`; `hash_password` y `verify_password` con perfil default |
| D89 | Migración descarta hashes legacy; todo usuario empieza con `password_hash=NULL` + `status='password_reset_required'` | Clean break de credenciales heredadas; sin doble algoritmo | `users.password_hash` NULL para los 156 usuarios; `users.status='password_reset_required'`; sin columna `legacy_hash` |
| D90 | Reset flow con tokens one-time de 24 h vía adapter de notificación (cola por tabla) | Auto-servicio de密码 recovery sin SMTP corporativo real | `issue_reset_token(user_id) -> str`; `consume_reset_token(token, new_password)` atómico y single-use |
| D91 | Primer global admin: `gentle-ai platform user set-password <email>` (CLI exclusivo admin global, D25) | Bootstrap inicial antes de que el sistema pueda emitir tokens por email | CLI subcomando con auditoría; solo accesible por admin global |
| D38–D40 | Lockout configurable + notificación | Defense-in-depth contra fuerza bruta | Service de lockout en `app/src/shared/` |
| D42 | Activación, baja y creación de roles: solo admin global | Separación de scopes | CLI solo accesible por admin global (D25) |
| D51 | Retirar formación/vídeos/cuestionarios | Histórico archivado, no operativo | No migran al esquema `lanzadera` |
| D52 | Retirar mecanismo de lanzamiento Access | Menú web permission-aware | No se incluyen `Shell`, `/cmd` ni UNC |
| D53 | Retirar segmentación oficina / fuera de oficina | UAT moderno sustituye | `EjecucionEnOficina` se ignora |
| D54 | Retirar gestión de rutas y contraseñas de backend | Configuración externalizada | No se migra `TbConfiguracionBackends` a UI |
| D55 | Modernizar auditoría + retirar telemetría SSID/ubicación | Minimización de datos | Solo evento autenticación/apertura |
| D56 | Preservar identidad, catálogo, usuarios, permisos | Suelo mínimo de capacidad de negocio | Núcleo de este `change` |
| D58 | Registro híbrido: deployment técnico + activación global | Idempotencia + activación auditable | `apps.registration_status` |
| D66–D68 | Stack cerrado (Python 3.12+ / FastAPI / HTMX / SQLAlchemy / Alembic) | Versiones verificadas en context7 | `pyproject.toml` con pins exactos |
| D70–D72 | Caché selectiva + ETag + gzip | No cachear contadores ni métricas volátiles | `cachetools.TTLCache` detrás de `CachePort` |
| D77 | Docker desde día uno + Compose para dev local | Sin instalación manual | `Dockerfile` + `docker-compose.yml` |
| D78–D80 | Semver por módulo + catálogo de versiones compatibles | Despliegues seguros | `lanzadera/vX.Y.Z`, `platform/vX.Y.Z` |
| D82 | Migraciones Expand and Contract | Nunca destructiva en una sola release | Las 0001–0006 son aditivas |
| D86–D87 | Forma hexagonal del legacy preservada como referencia | El legacy ya tenía esta forma | Tests VBA → pytest con trazabilidad |

### Quality gates que el MVP referencia explícitamente

| # | Decisión | Aplicación |
|---|---|---|
| QC-1 | Set MVP derivado de APAP_WEB con exclusiones justificadas | 12 gates en CI; exclusiones en `docs/calidad-de-codigo-y-ci.md` |
| QC-2 | Hexagonal layer gate desde el primer slice | `scripts/check_layers.py` corre el día 0 |
| QC-3 | ruff `E,F,W,I,UP,B` desde día uno, sin ratchet | `pyproject.toml` con pin exacto `ruff==0.15.21` |
| QC-4 | mypy strict con `ignore-without-code` | `pyproject.toml` `[tool.mypy]` con código exigido |
| QC-5 | Cobertura 85 % global, 100 % CRITICAL_HELPERS | `hash_password`, `verify_password`, `issue_reset_token`, `consume_reset_token` |
| QC-6 | PR size 400 + branch-name gate | `scripts/check_pr_size.py`, `check_branch_name.py` |
| QC-7 | Security scanning con imágenes pinned por digest | `pip-audit`, `gitleaks`, `trivy` con `@sha256:...` |
| QC-8 | cosmic-ray DIFERIDO a Fase 2 | Sin código que mutar; el gate sería ruido |
| QC-9 | `check_layers.py` adaptado de APAP_WEB | Copia literal; ajusta `ALLOWED_IMPORTS` y `PURE_LAYERS` |

## Hallazgos y riesgos

| # | ID | Severidad | Título | Mitigación |
|---|---|---|---|---|
| H1 | D109 | **Resuelto** | `TbUsuariosAplicaciones.Password` se documenta como plaintext en `data-model.md`, pero D36 asumía hash heredado | **Resuelto 2026-08-08**: SHA256-hex confirmado por `integrations-security.md` (Login compara SHA256-hex del input vs el Password almacenado). Resolución arquitectónica: la columna legacy NO se preserva; D88-D91 sustituyen D36-D37 con Argon2id + reset flow. No se requiere lectura del binario para esta decisión |
| H2 | D156 | Media | Binarios `.accdb` NO están en staging dentro del repo | Pull desde R2 antes de ejecutar el script; documentar hash esperado en `scripts/migrate_from_access.py` y fallar si no coincide |
| H3 | D157 | Baja | `codegraph-vba` no indexa Lanzadera en algunos worktrees (cross-project noise) | **Verificado 2026-08-08**: el worktree `00_LANZADERA\staging` SÍ tiene índice sano (90 archivos, 3087 nodos, 4431 edges, 5.47 MB, status "no source changes detected"). Las queries confirman cobertura de `SHA256`, `Login`, `LoginCorrecto` desde el MCP `codegraph-vba`. H3 aplica solo a OTROS worktrees Lanzadera no indexados (developer branches). Si surge otro worktree Lanzadera sin índice, ejecutar `codegraph init` allí antes del walkthrough. Riesgo de cross-project noise persiste solo si múltiples worktrees apuntan al mismo `.codegraph-vba/` (no es nuestro caso: cada worktree tiene su propio `.codegraph-vba/` por convención global). Slate-clean vía `codegraph uninit` no viable: los MCP servers de los agentes retienen el handle hasta que se cierran los agentes |
| H4 | Gap P-21 | Baja | Sincronización manual de datos comunes (users Lanzadera, expedientes) al NAS oficina | **Aceptado, no bloqueante.** Procedimiento operativo documentado en `docs/operacion/sync-nas-oficina.md` (pendiente de crear en `sdd-apply`). Apps office-only consultan snapshot, no fuente única |
| H5 | D74 / D75 | Baja | Topología de despliegue ABIERTA | Salvaguardada por D74 (no Kubernetes prematuro), D75 (abierta), D76 (PostgreSQL gestionado preferido). No bloquea MVP |
| H6 | D17 / D18 | Baja | Adjuntos sin versionado + papelera 30 días | No aplica al MVP de Lanzadera admin (no hay adjuntos en este slice). Se reabre cuando llegue el módulo de expedientes |
| H7 | D158 | Media | Forms G4 con `MSComctlLib.TreeCtrl.2` + `ImageListCtrl.2` (3 forms) | Quedan en legacy archivado (D51). La migración web no los toca; el admin de Lanzadera no requiere esa UI |
| H8 | D161 / D166 | Media | Global state `m_ObjVideoActivo`, `m_ObjAplicacionActivo` en 7 forms G4 | Quedan en legacy archivado (D51). Pre-requisito para tests del legacy, pero el MVP web no los hereda |
| H9 | D102 | Baja | Booleanos como `Text(2)` (Sí/No) en 20+ columnas | Las columnas equivalentes serán `BOOLEAN` en PostgreSQL; la migración traduce `'Sí'` → `true`, `'No'` → `false`, otros → `NULL` |
| H10 | D110 | Baja | `TbPermisos` con campos dinámicos `F3-F9` (anti-patrón) | No se migra la estructura. Las capabilities viven en `profiles.capabilities JSONB`. Cualquier dato histórico se archiva |
| H11 | Regla de mapeo | Media | `SinAcceso` exclusivo: si una fila legacy tiene `SinAcceso = Sí` Y otros `Sí`, solo se crea la fila `SIN_ACCESO` y se ignoran las demás | Documentado en CA-F3. Verificable con test unitario sobre la matriz legacy → `user_app_assignments`. Caso límite: `SinAcceso = NULL` y al menos un Sí → comportamiento compuesto normal |
| H12 | Topología híbrida | Alta | El flag `apps.requires_office_presence` no se evalúa en el MVP | `LocationPort.is_user_in_office` devuelve `True` por defecto (`assume_in_office`). El refuerzo llegará cuando se introduzcan apps office-only (HPS) en su propio `change` |
| H13 | P16 | Baja | Diseño detallado del ciclo UAT ABIERTO | No bloquea el MVP. La gobernanza admin-global-only ya está aprobada (D48). El ciclo detallado se cierra en un `change` posterior |
| H14 | QC-8 | Baja | cosmic-ray DIFERIDO | Aceptado por QC-8. Se reactiva cuando ≥ 3 módulos tengan cobertura > 70 % y ≥ 5 tests por path real |

## Tickets derivados

Tickets que se crearán al lanzar `sdd-apply`. La numeración es provisional y se consolidará en el PR de apply.

| Ticket | Título | Fase |
|---|---|---|
| TK-LZ-MVP-1 | Crear estructura `app/` con `pyproject.toml`, `Dockerfile`, `docker-compose.yml`, `Makefile`, `.python-version`, `.dockerignore` | Día 0 |
| TK-LZ-MVP-2 | Activar ruff + mypy + pytest + `coverage_gate` + PR size + branch-name en CI (QC-1, QC-3, QC-4, QC-5, QC-6) | Día 1 |
| TK-LZ-MVP-3 | Activar security scanning (`pip-audit`, `gitleaks`, `trivy`) con imágenes pinned por digest (QC-7) | Día 2 |
| TK-LZ-MVP-4 | Implementar `scripts/check_layers.py` con `ALLOWED_IMPORTS`, `PURE_LAYERS`, slicing (QC-2, QC-9) | Día 3 |
| TK-LZ-MVP-5 | Implementar `scripts/check_complexity.py` con `CC ≤ 15` para el top-10 | Día 4 |
| TK-LZ-MVP-6 | Convenciones operativas en `AGENTS.md` raíz (Conf-1 a Conf-8) | Día 5 |
| TK-LZ-MVP-7 | Primer módulo `lanzadera.auth` con TDD; `CRITICAL_HELPERS = ["hash_password", "verify_password", "issue_reset_token", "consume_reset_token"]` | Día 6 |
| TK-LZ-MVP-8 | Migración 0001 — schema core (`users`, `apps`, `profiles`, `user_app_assignments`, `global_admins`, `sessions`, `audit`) | Aplicación |
| TK-LZ-MVP-9 | Migración 0002 — seed de `apps` desde `TbAplicaciones` (20 filas; `deployment_topology`, `requires_office_presence`) | Aplicación |
| TK-LZ-MVP-10 | Migración 0003 — seed de `profiles` por app (`default` + efectivamente usados en legacy) | Aplicación |
| TK-LZ-MVP-11 | Migración 0004 — usuarios desde `tbUsuarios` con DNI cifrado en reposo, `password_hash=NULL`, `status='password_reset_required'` (D88, D89). Sin columna `legacy_hash` | Aplicación |
| TK-LZ-MVP-12 | Migración 0005 — assignments aplicando la regla de mapeo (Administrador/Calidad/CalidadAvisos/Técnico/Economía/Secretaría/SinAcceso/DEFAULT) | Aplicación |
| TK-LZ-MVP-13 | Migración 0006 — seed de `audit` desde `TbConexiones` + `TbAplicacionesAperturas` (D55) | Aplicación |
| TK-LZ-MVP-14 | `scripts/migrate_from_access.py` — extracción one-shot con Dysflow, produce fixtures JSON, ejecutable en local y en CI | Aplicación |
| TK-LZ-MVP-15 | `BootstrapAdapter` — siembra `global_admins` desde `GLOBAL_ADMIN_EMAILS`; idempotente; auditado | Aplicación |
| TK-LZ-MVP-16 | Use cases de admin: alta/baja de usuario, asignación de profiles, listado de catálogo, auditoría | Aplicación |
| TK-LZ-MVP-17 | `CachePort` con `cachetools.TTLCache` para candidatos justificados (catálogo de apps, profiles por app, permisos efectivos) | Aplicación |
| TK-LZ-MVP-18 | `LocationPort.is_user_in_office` con implementación MVP `assume_in_office` | Aplicación |
| TK-LZ-MVP-19 | Adapter v1 de notificación con cola por tabla (`TbCorreosEnviados`) | Aplicación |
| TK-LZ-MVP-20 | Adapter de credenciales: `hash_password(plain) -> str` y `verify_password(plain, hashed) -> bool` con Argon2id `argon2-cffi==25.1.0` perfil `RFC_9106_LOW_MEMORY`. Sin compat con legacy | Aplicación |
| TK-LZ-MVP-21 | Delivery HTMX: rutas FastAPI, templates Jinja async, Alpine.js para interacciones mínimas | Aplicación |
| TK-LZ-MVP-22 | CLI admin global para `bootstrap_global_admins`, `set-password <email>` (D91), `list_apps`, `assign_profile` | Aplicación |
| TK-LZ-MVP-24 | Reset flow: `issue_reset_token(user_id) -> str` y `consume_reset_token(token, new_password)`. Token 24 h single-use; emisión vía cola de notificación por tabla (D90) | Aplicación |
| TK-LZ-MVP-23 | UAT del ecosistema completo (no de este slice aislado): cut-over one-shot desde `.accdb` legacy | Go-live |

## Cómo se aplica a access2web

Esta propuesta abre el flujo SDD sobre `app/`. El árbol de directorios objetivo respeta el monolito modular aprobado en D68 y los límites por paquete del `check_layers.py`.

### Estructura física objetivo

```
access2web-blueprint/
├── app/                       # NUEVO — monolito modular hexagonal
│   ├── src/
│   │   ├── shared/                 # cross-cutting (cache port, audit, auth base)
│   │   ├── modules/
│   │   │   └── lanzadera/          # módulo admin (MVP slice)
│   │   │       ├── domain/         # entities, invariants
│   │   │       ├── ports/          # Protocols
│   │   │       ├── application/    # use cases
│   │   │       ├── adapters/       # impls driven (Postgres, in-memory, ...)
│   │   │       ├── di/             # composition root
│   │   │       └── delivery/       # routes FastAPI + HTMX templates
│   │   └── main.py                 # FastAPI app composition root
│   ├── migrations/                 # Alembic
│   │   └── versions/               # 0001-0006
│   ├── tests/
│   ├── scripts/                    # checkers custom + migrate_from_access.py
│   ├── pyproject.toml
│   └── Dockerfile
├── docker-compose.yml              # postgres + minio dev
├── docs/                           # EXISTENTE
├── openspec/
│   └── changes/
│       └── lanzadera-mvp/          # ESTA PROPUESTA
└── .opencode/skills/               # EXISTENTE
```

### Contrato de capabilities con `sdd-spec`

Esta sección es el contrato entre `proposal` y `sdd-spec`. El sub-agente `sdd-spec` la lee para saber qué archivos `openspec/specs/` producir.

- **New capability**: `lanzadera-admin` — núcleo de administración de plataforma: identidad, catálogo, asignaciones, auditoría, lockout, bootstrap de admins globales, sync NAS manual (P-21).
- **Modified capabilities**: ninguno en el MVP. Las capabilities futuras (expedientes, gestión de riesgos, etc.) abren sus propios `change`.

### Pipeline de datos de un solo paso

```
Lanzadera_Datos.accdb
  └─→ Dysflow read-only (extract select)
       └─→ scripts/migrate_from_access.py (one-shot, fixtures JSON)
            └─→ Alembic 0001 (schema) → 0002 (seed apps) → 0003 (seed profiles)
                 → 0004 (users + DNI cifrado + legacy_hash)
                 → 0005 (assignments aplicando regla de mapeo)
                 → 0006 (audit snapshot)
```

### Plan de rollback

El MVP introduce schema nuevo en PostgreSQL, sin tocar el binario `.accdb` legacy. El rollback es directo:

1. Detener `app/`.
2. `DROP SCHEMA lanzadera CASCADE;` — la base de datos PostgreSQL queda limpia.
3. El `.accdb` legacy sigue intacto y operativo.

Las migraciones 0001–0006 son aditivas (D82, Expand and Contract). Ninguna tira columnas legacy del backend; solo construye el nuevo schema. Para una release de hotfix destructivo, se sigue Contract después de validar uso masivo de la nueva estructura.

## Anexo A — Datos crudos relevantes del walkthrough Lanzadera

> Evidencia consolidada del discovery Lote 1 + segunda pasada de inventario. Los números son **verified-runtime** (Dysflow + `SELECT` agregado). Para el detalle exhaustivo, ver `docs/03-aplicaciones/lanzadera/{README,capabilities,data-model,integrations-security,epic}.md` y los walkthroughs `walkthrough-G1.json` a `walkthrough-G5.json`.

### Volumen del backend autoritativo (`C:\00repos\datos\Lanzadera_Datos.accdb`)

| Tabla | Filas | Notas para migración |
|---|---:|---|
| `tbUsuarios` | 156 | Datos personales: `Matricula_DNI`, `DirCorreo`, `telfijo`, `telmovil`. `SeLogean`, `ParaTareasProgramadas`, `Autorizador` son `YesNo` reales; `EmplazamientoExterno`, `UsuarioDeGestionRiesgos`, `UsuariosI3D` son `Text(2)` (D102) |
| `TbUsuariosAplicacionesPermisos` | 622 | 7 flags de rol como `Text(2)` Sí/No (D102). PK compuesta `CorreoUsuario + IDAplicacion` |
| `TbAplicaciones` | 20 | Catálogo con `Pass` (contraseña de la app) en texto plano ⚠️. Campos `EjecucionEnOficina`, `EnPruebas`, `ConIconoEnLanzadera` como `Text(2)` (D102) |
| `TbUsuariosAplicaciones` | 622 | `Password` en texto plano ⚠️⚠️⚠️. Sistema de caducidad (`FechaProximoCambioContrasenia`, `TieneQueCambiarLaContrasenia`, `PasswordNuncaCaduca`) |
| `TbConexiones` + `TbConexionesRegistro` | TBD | Telemetría SSID/ubicación/coordenadas (D55 retira) |
| `TbAplicacionesAperturas` | TBD | Evento de apertura con máquina/usuario de máquina |
| `TbUsuariosHistoricoContrasenias` | TBD | `PassAntigua` en texto plano (D37 retira, conserva solo evento) |
| Tablas formación | — | `TbCuestionarios`, `TbCuestionaroRespuestas` (typo), `TbVideos*` — **D51 retirar, no migran** |
| `TbPermisos` | TBD | Anti-patrón campos dinámicos `F3-F9` (D110) — no se migra la estructura |

### Forma hexagonal del legacy (D86)

La inspección CodeGraph-VBA confirma que Lanzadera ya tenía forma hexagonal en legacy:

- **Clases de dominio**: `Usuario`, `Aplicacion`, `Conexion`, `AplicacionApertura`, `Menu`, `Video`, `Categoria`, `Visionado`, `UsuarioAplicacionPermisos`.
- **Helpers por dominio**: módulos con `Helper_` por cada clase.
- **Transaccionalidad en `ExpedienteOperaciones`**: unidad de persistencia centralizada.
- **Configuración detrás de `getdb()` + `TbConfiguracionBackends`**: 76 callers internos de `getdb` (verified-static).
- **Permisos centralizados en `UsuarioAplicacionPermisos`**: 7 callers en clases y formularios.

Esta forma es la **referencia** para mapear al nuevo hexagonal web, no una invención del blueprint.

### Tests VBA preservados como referencia (D87)

`tests.vba.json` + `tests.vba.responsable-71.json` registran los `Test_*` por cada `Helper_*`. Ningún test se descarta sin trazabilidad. Cuando el MVP tenga `CRITICAL_HELPERS` en pytest, los tests VBA equivalentes quedan enlazados por `metadata.test_id`.

### Walkthrough 28/28 forms

El walkthrough v4 (2026-08-07, método v4) cubre los 28 forms de Lanzadera en 5 grupos (`walkthrough-G1.json` a `walkthrough-G5.json`). Los 5 critical findings G4 (ActiveX OCX, twin forms, grandparent coupling, zero test coverage, global state) **no migran** al web: la formación se retira (D51) y los forms G4 quedan en el legacy archivado. `codegraph-vba` issues abiertas #1408 y #1412 (`verify_form_bindings` y `map_form_behavior`) se documentaron como `skipped_tool_broken` con workaround aplicado.

### Hallazgos críticos del inventario (Lote 1)

| Hallazgo | Tabla | Acción |
|---|---|---|
| D109 ⚠️⚠️⚠️ | `TbUsuariosAplicaciones.Password`, `TbUsuarioAplicacionesSolicitud.Password` | Migración con hash (ver H1) |
| D109 ⚠️ | `TbUsuariosHistoricoContrasenias.PassAntigua` | Eliminar contraseñas del histórico, conservar evento |
| D109 ⚠️ | `TbAplicaciones.Pass` | Secret manager (no UI) |
| D102 | 20+ columnas como `Text(2)` | Estandarizar a `BOOLEAN` |
| D110 | `TbPermisos.F3-F9` | Migrar capabilities a `profiles.capabilities JSONB` |
| D111 | `TbAplicacionesEstados` con PK `Text(255)` | Agregar ID numérica como PK |
| D112 | `TbConexiones` con telemetría | D55 retira SSID/ubicación/coordenadas |

## Anexo B — Glosario de términos del dominio

| Término | Definición operativa |
|---|---|
| **Administrador global** | Persona con scope plataforma. Única autorizada para alta/baja de usuarios, nombramientos de admin de aplicación, configuración de health-checks, diagnósticos con remediación (D21, D23, D25, D34). Se siembra desde `GLOBAL_ADMIN_EMAILS` |
| **Administrador de aplicación** | Persona con scope limitado a un módulo. Sus capacidades las define el módulo (D22). Solo el admin global los nombra (D23) |
| **App / Aplicación** | Entrada del catálogo `apps` (`TbAplicaciones` en legacy). Cada app tiene `deployment_topology` (`central` o `office-nas`) y `requires_office_presence` |
| **Capability** | Permiso estable y verificable declarado por un módulo. Los `profiles` agrupan capabilities en `capabilities JSONB` (D45) |
| **Profile** | Conjunto de capabilities asociado a una app. `profiles(id, app_id, code, name, capabilities, active)` (D45, D46) |
| **User-app assignment** | Relación N-a-N entre usuarios y profiles dentro de una app. `user_app_assignments(user_id, app_id, profile_id)` |
| **Legacy hash** | Hash preservado del `.accdb` (D36). Se verifica con un adapter versionado y se reemplaza al primer login exitoso (D37) |
| **CRITICAL_HELPERS** | Funciones que el coverage gate marca al 100 % en el MVP: `hash_password`, `verify_password`, `issue_reset_token`, `consume_reset_token` (QC-5, D88, D90) |
| **Expand and Contract** | Estrategia de migración aditiva (D82): añadir estructura nueva sin retirar la vieja en la misma release |
| **LocationPort** | Puerto hexagonal que indica si el usuario está físicamente en la oficina. MVP: `assume_in_office` que devuelve `True` |
| **Deployment topology** | `central` o `office-nas`. Apps `office-nas` requieren NAS en oficina + sync manual P-21 |
| **Rehash transparente** | Política que sustituye el hash heredado por el moderno al primer login exitoso (D37). Atómico, idempotente, auditable |
| **SinAcceso exclusivo** | Regla del MVP: si la fila legacy tiene `SinAcceso = Sí`, solo se crea la fila `user_app_assignments(profile_id = SIN_ACCESO)`. Las demás flags Sí se ignoran |
| **Snapshot de auditoría** | Inserción inicial en `audit` desde `TbConexiones` + `TbAplicacionesAperturas`. D55 retira SSID/ubicación/coordenadas |
| **Pure layer** | Layer sin frameworks importados (`domain`, `ports`, `application`). Enforced por `scripts/check_layers.py` con `PURE_LAYERS` (QC-2, QC-9) |
| **Vertical slicing** | Cada módulo (`lanzadera/`, `expedientes/`, ...) posee su columna hexagonal. `_check_slice` prohíbe reach-across entre módulos salvo core cross-cutting exento |

## Lista de comprobación final

- [ ] Alcance explícito sobre D51–D56 (incluye D56 núcleo; excluye D51–D55 capacidades retiradas).
- [ ] Regla de mapeo `legacy → user_app_assignments` documentada, con `SinAcceso` exclusivo y `DEFAULT` para todas No.
- [ ] Bootstrap de `global_admins` desde `GLOBAL_ADMIN_EMAILS` documentado como idempotente.
- [ ] Topología híbrida documentada: `central` + `office-nas` con sync manual P-21 aceptado.
- [ ] Quality gates QC-1 a QC-9 referenciados con su día de adopción.
- [ ] Decisiones D5–D87 relevantes referenciadas.
- [ ] Hallazgos H1–H14 con `##ABIERTO##` explícito en los gaps que requieren validación antes de `sdd-apply`.
- [ ] Tickets derivados TK-LZ-MVP-1 a TK-LZ-MVP-23 listos para abrir en `sdd-apply`.
- [ ] Plan de rollback explícito: `DROP SCHEMA lanzadera CASCADE`.
- [ ] Anexo A con datos verificados del walkthrough; Anexo B con glosario operativo.
- [ ] Tone: Castellano peninsular formal en el cuerpo narrativo; inglés en nombres de archivo, código y secciones técnicas.
- [ ] No se crean `specs.md`, `design.md` ni `tasks.md` en este `change` (los abren `sdd-spec`, `sdd-design` y `sdd-tasks`).

## Siguiente paso

`sdd-spec` lee esta propuesta y produce `openspec/changes/lanzadera-mvp/specs.md` con los escenarios Given/When/Then por capability. El contrato de capabilities vive en §Cómo se aplica a access2web.
