# 08 · Decisiones y preguntas abiertas

## Propósito

Registro vivo de las decisiones tomadas durante el descubrimiento y de las preguntas que permanecen abiertas, en particular las que bloquean la entrada a un lote o a una fase SDD posterior. No sustituye a la conversación con el usuario; la refleja y la prepara.

## Decisiones (mínimo inicial)

| # | Decisión | Fecha | Origen | Impacto |
|---|---|---|---|---|
| D1 | El alcance del descubrimiento se limita a las **ocho aplicaciones** listadas en `00-alcance-y-evidencia.md`. APAP y APAP_WEB quedan **fuera de alcance**. | 2026-08-04 | Usuario y `exploration.md` | Fija el universo de discovery. |
| D2 | La jerarquía de fuentes de autoridad es: `C:\00repos\documentacion` → checkout `00_main` por aplicación → Dysflow solo lectura → CodeGraph → Engram (solo contexto histórico). | 2026-08-04 | `exploration.md` | Regla transversal para todas las fichas. |
| D3 | La Fase 1 se ejecuta en **9 lotes con aprobación por lote**; la recomendación es comenzar por Lanzadera y Expedientes. | 2026-08-04 | `exploration.md` | Regula el ritmo de descubrimiento. |
| D4 | Las configuraciones Dysflow con `path-mismatch` o ausentes **no se mutan** en esta fase. | 2026-08-04 | `exploration.md` | Evita mutaciones no autorizadas. |

## Decisiones de producto y arquitectura (consolidadas)

Las decisiones **D5–D35** están consolidadas en `09-arquitectura-objetivo-y-principios.md` con su detalle, `topic_key` de Engram y separación entre `APROBADO`, `PROVISIONAL`, `FUTURO` y `ABIERTO`. La tabla resumen se mantiene aquí como índice operativo. Las decisiones **D36–D65** son las adoptadas tras el commit baseline `d521d7b` durante la pasada de Lanzadera (Lote 1) y siguientes; cada una enlaza con su `topic_key`. Las decisiones **D66–D82** cierran el estudio del prompt externo de arquitectura del 2026-08-05 (stack, polling, caché selectiva, despliegue, versionado y releases) tras revisión decisión por decisión con el usuario. Las decisiones **D83–D87** cierran el estudio `p1-p2-p5-resolved-aug2026`. Las decisiones **D88–D91** son las adoptadas en sesión 2026-08-08 al decidir descartar hashes heredados (D36+D37 obsoletos) y usar Argon2id con reset flow explícito.

| # | Tema | Estado | Origen detallado |
|---|---|---|---|
| D5 | Plataforma única modular permission-aware | APROBADO | `architecture/target-platform-shape` |
| D6 | Navegación app-first anidada | APROBADO | `architecture/navigation-model` |
| D7 | Lanzadera = solo administración | APROBADO | `architecture/identity-and-admin-module` |
| D8 | Hexagonal global en plataforma y módulos | APROBADO | `architecture/global-hexagonal-principle` |
| D9 | Autenticación inicial email/password (adaptador) | APROBADO | `architecture/authentication-ports-adapters` |
| D10 | SiteMinder/OCP/JWT unificado | FUTURO | `architecture/authentication-ports-adapters` |
| D11 | Notificación unificada como servicio compartido | APROBADO | `architecture/unified-notification-service` |
| D12 | Notificación v1 solo email | APROBADO | `product/notification-v1-scope` |
| D13 | Cola de correo por tabla como adaptador transitorio | APROBADO | `architecture/notification-delivery-adapter-v1` |
| D14 | PostgreSQL compartido con esquemas por módulo | APROBADO (dirección) | `architecture/target-database-topology` |
| D15 | Rendimiento y caché deliberada como prioridad | APROBADO | `architecture/performance-and-cache` |
| D16 | Adjuntos en object storage S3-compatible detrás de puerto | APROBADO | `architecture/attachment-storage` |
| D17 | Adjuntos sin versionado de contenido | APROBADO | `product/attachment-versioning` |
| D18 | Adjuntos con papelera de retención limitada | APROBADO | `product/attachment-deletion` |
| D19 | Retención de papelera = 30 días | APROBADO | `product/attachment-retention` |
| D20 | Restauración de papelera por borrado o administrador global | APROBADO | `product/attachment-restore-authorization` |
| D21 | Roles: administrador global + admin de aplicación por módulo | APROBADO | `architecture/authorization-roles` |
| D22 | Capacidades de admin de aplicación definidas por módulo | APROBADO | `architecture/module-admin-capabilities` |
| D23 | Solo administrador global nombra admins de aplicación | APROBADO | `architecture/module-admin-assignment` |
| D24 | CLI operable por IA para todas las acciones de aplicación | APROBADO | `architecture/ai-cli-first` |
| D25 | CLI exclusivo para administrador global | APROBADO | `architecture/cli-access-control` |
| D26 | Confirmación explícita en acciones destructivas o masivas | APROBADO | `architecture/cli-destructive-confirmation` |
| D27 | Logs estructurados canónicos con correlación | APROBADO | `architecture/observability-audit-logs` |
| D28 | Retención por niveles configurable con archivo en object storage | APROBADO | `architecture/audit-retention` |
| D29 | Retención provisional 90 días hot + 1 año total | PROVISIONAL | `product/audit-retention-periods` |
| D30 | CLI soporta diagnósticos operativos proactivos | APROBADO | `architecture/ai-operations-cli` |
| D31 | Toda remediación requiere aprobación del administrador global | APROBADO | `architecture/ai-remediation-approval` |
| D32 | Inspecciones de salud bajo demanda + programadas | APROBADO | `architecture/scheduled-health-checks` |
| D33 | Destinatarios de anomalías configurables por aplicación | APROBADO | `architecture/application-operations-settings` |
| D34 | Solo administrador global configura health-checks | APROBADO | `architecture/health-check-configuration-authorization` |
| D35 | Resultado del estudio = roadmap global + roadmap y plan por herramienta | APROBADO | `product/modernization-principles` |
| D36 | Preservar hashes heredados de Lanzadera en migración | OBSOLETO 2026-08-08 | `product/credential-migration` |
| D37 | Rehash transparente al primer login exitoso | OBSOLETO 2026-08-08 | `architecture/opportunistic-password-rehash` |
| D88 | Auth usa Argon2id vía `argon2-cffi==25.1.0` con perfil `RFC_9106_LOW_MEMORY` (Argon2id, 64 MiB, 3 iteraciones, 4 hilos) | APROBADO | `product/lanzadera-auth-modern-crypto` |
| D89 | Migración descarta hashes heredados; todo usuario migrado empieza con `password_hash=NULL` + `status='password_reset_required'`. Sin columna `legacy_hash` | APROBADO | `product/lanzadera-no-legacy-hash` |
| D90 | Reset flow con tokens one-time de 24 h vía adapter de notificación con cola por tabla. `issue_reset_token` y `consume_reset_token` atómicos y single-use | APROBADO | `product/lanzadera-reset-flow` |
| D91 | Primer global admin se aprovisiona con CLI `gentle-ai platform user set-password <email>` (D25) antes de que el sistema pueda emitir tokens por email | APROBADO | `product/lanzadera-bootstrap-cli` |
| D38 | Umbral de lockout configurable, default cinco intentos | APROBADO | `architecture/login-lockout-policy` |
| D39 | Bloqueo durante una hora por defecto; desbloqueo por admin global | APROBADO | `architecture/login-lockout-recovery` |
| D40 | Notificación a administradores globales en cada lockout | APROBADO | `architecture/lockout-notification` |
| D41 | Caducidad de contraseña configurable global + exenciones por usuario | APROBADO | `architecture/password-expiry-policy` |
| D42 | Activación, baja y creación de roles: solo administrador global | APROBADO | `architecture/global-only-identity-actions` |
| D43 | Notificación de cambios de permiso solo por acción del administrador | APROBADO | `architecture/manual-permission-change-notification` |
| D44 | Suplantación para pruebas restringida al administrador global | APROBADO | `architecture/impersonation-authorization` |
| D45 | Capabilities declaradas por módulo + grupos de capabilities | APROBADO | `architecture/capability-driven-ui-variants` |
| D46 | Vista única por defecto; vistas especializadas decididas por módulo | APROBADO | `architecture/per-module-ui-variant-policy` |
| D47 | UAT: asignación explícita de participantes y perfil por ciclo | APROBADO | `product/uat-participant-governance` |
| D48 | UAT: gobernanza reservada al administrador global | APROBADO | `architecture/uat-global-admin-governance` |
| D49 | UAT: excepciones auditadas y visibles en historial de cambios | APROBADO | `architecture/uat-release-exceptions` |
| D50 | UAT: excepciones visibles para usuarios en historial de release | APROBADO | `product/release-exception-transparency` |
| D51 | Lanzadera retira formación/vídeos/cuestionarios | APROBADO | `product/lanzadera-training-disposition` |
| D52 | Lanzadera retira mecanismo de lanzamiento y actualización de ficheros Access | APROBADO | `product/lanzadera-launcher-disposition` |
| D53 | Lanzadera retira segmentación oficina / fuera de oficina | APROBADO | `product/lanzadera-location-visibility` |
| D54 | Lanzadera retira gestión de rutas y contraseñas de backend | APROBADO | `product/lanzadera-backend-config-disposition` |
| D55 | Lanzadera moderniza auditoría; retira telemetría de SSID/ubicación/coordenadas | APROBADO | `product/lanzadera-audit-disposition` |
| D56 | Lanzadera preserva y moderniza identidad, catálogo, usuarios, permisos | APROBADO | `product/lanzadera-core-capabilities` |
| D57 | Lanzadera: UAT detallado y diseño de ciclo ABIERTO | APROBADO (apertura de diseño) | `product/module-uat-lifecycle` |
| D58 | Registro de aplicaciones híbrido: deployment técnico + activación global | APROBADO | `architecture/application-registration` |
| D59 | Scheduler unificado sustituye flags y tareas de aplicación | APROBADO | `architecture/shared-scheduler-operations` |
| D60 | Configuración global de horarios y destinatarios de informes | APROBADO | `architecture/report-job-configuration` |
| D61 | Ejecución manual bajo demanda de cualquier informe configurado | APROBADO | `architecture/manual-report-execution` |
| D62 | Vista previa de informe sin envío | APROBADO | `architecture/report-preview` |
| D63 | Generar y enviar directamente sin vista previa cuando proceda | APROBADO | `architecture/direct-manual-report-send` |
| D64 | Dashboard global de operaciones de notificación (diseño detallado ABIERTO) | APROBADO (dirección) | `architecture/notification-operations-dashboard` |
| D65 | Evidencia legacy: cola por tabla + dispatcher externo cada 5 min | APROBADO | `discovery/legacy-email-queue-flow` |
| D66 | Stack backend: Python 3.12+ con FastAPI 0.119+, Pydantic v2, SQLAlchemy 2.0.x (mixto ORM/Core), Alembic 1.13+, driver asyncpg 0.30+ | APROBADO | `external-prompt-review/stack-versions-verified` |
| D67 | Stack frontend: HTMX 2.0.4 + Jinja2 3.1+ (async) + Alpine.js 3.15+ (SSR puro, sin SPA, sin build pipeline) | APROBADO | `external-prompt-review/stack-versions-verified` |
| D68 | Estructura del repositorio: monorepo `access2web-blueprint/` con monolito modular, límites por paquete y ports por módulo | APROBADO | `external-prompt-review/section-2-resolution` |
| D69 | Actualización de contadores pendientes vía polling HTMX (`hx-trigger="every 30s"`) + botón de refresh manual. Sin SSE, sin WebSockets | APROBADO | `external-prompt-review/section-1-resolution` |
| D70 | Caché de aplicación selectiva y justificada por medición. Candidatos: catálogos, permisos efectivos, diccionarios. NO se cachean contadores ni métricas volátiles | APROBADO | `external-prompt-review/section-3-resolution` |
| D71 | Redis queda como opción detrás del puerto de caché, no dependencia inicial. Pub/Sub se introduce solo cuando se justifique por escala horizontal | APROBADO | `external-prompt-review/section-3-resolution` |
| D72 | Rendimiento HTTP: ETag + 304 Not Modified para fragmentos HTMX, compresión gzip/brotli por defecto, Cache-Control con fingerprint en assets estáticos | APROBADO | `external-prompt-review/section-3-resolution` |
| D73 | No usar Insforge como BaaS. El backend hexagonal es nuestro (FastAPI + adaptadores propios). Insforge puede ser herramienta auxiliar para prototipos, nunca dependencia del backend | APROBADO | `external-prompt-review/section-4-resolution` |
| D74 | No introducir Kubernetes ni OpenShift prematuramente. Para 200 usuarios concurrentes, una instancia de FastAPI + PostgreSQL es suficiente | APROBADO | `external-prompt-review/section-4-resolution` |
| D75 | Topología de despliegue ABIERTA (consistente con P7). Decisión de cloud y orquestador queda pendiente de métricas + IT corporativa | APROBADO (apertura) | `external-prompt-review/section-4-resolution` |
| D76 | PostgreSQL gestionado preferido sobre auto-instalado cuando se decida el cloud. Proveedor concreto se liga a D75 | APROBADO | `external-prompt-review/section-4-resolution` |
| D77 | Contenedores Docker desde el día uno + Docker Compose para dev local con PostgreSQL + MinIO | APROBADO | `external-prompt-review/section-4-resolution` |
| D78 | Versionado semántico por módulo Y plataforma base (`modulo/vX.Y.Z-rc.n` / `modulo/vX.Y.Z` y `platform/vX.Y.Z`) | APROBADO | `external-prompt-review/section-5-resolution` |
| D79 | Trunk-based development + Conventional Commits como entrada al versionado y changelog | APROBADO | `external-prompt-review/section-5-resolution` |
| D80 | Catálogo de versiones compatibles entre módulos y plataforma (evita combinaciones inválidas en despliegues) | APROBADO | `external-prompt-review/section-5-resolution` |
| D81 | Despliegue coexistente estable+RC en UAT DIFERIDO hasta que cadencia y equipo lo justifiquen. UAT y Producción como entornos separados | APROBADO (diferimiento) | `external-prompt-review/section-5-resolution` |
| D82 | Migraciones de BD backward-compatibles con estrategia Expand and Contract: nunca destructivas en una sola release | APROBADO | `external-prompt-review/section-5-resolution` |
| D83 | HPS_Solicitudes se documenta como aplicación **independiente** de HPS: `TbAplicaciones.ID = 22` ("Solicitudes HPS"), con su propio checkout, frontend (`Solicitudes_HPS.accdb`) y backend (`Solicitudes_HPS_datos.accdb`). Comparte nombre conceptual con HPS (ID 17) por razones históricas pero son productos distintos | APROBADO | `blueprint/p1-p2-p5-resolved-aug2026` |
| D84 | Baseline de release de Condor y Brass reside en `C:\00repos\codigo\00_<app>\00_main\` + `C:\00repos\datos\<backend>_datos.accdb`. Staging de Condor en `00_CONDOR\staging\`. Brass requiere clarificación de rama de desarrollo (develop vs release_2026-001) | APROBADO | `blueprint/p1-p2-p5-resolved-aug2026` |
| D85 | Baseline operativo actual: catálogo `TbAplicaciones` con 8 IDs en alcance (5 Riesgos, 6 Brass, 8 No Conformidades, 12 Lanzadera, 17 HPS, 19 Expedientes, 22 Solicitudes HPS, 23 Condor), backends en `C:\00repos\datos\`. Tabla de referencia transversal para Lotes 2 a 9 | APROBADO | `blueprint/p1-p2-p5-resolved-aug2026` |
| D86 | La forma hexagonal del legacy (dominio en clases, helpers por dominio, transaccionalidad en `ExpedienteOperaciones`, config detrás de `getdb()` + `TbConfiguracionBackends`) **se preserva como referencia** para el mapeo a la nueva plataforma. El blueprint hexagonal NO es invención nueva: el legacy ya tenía esta forma | APROBADO | `blueprint/staging-expedientes-hexagonal-shape` |
| D87 | Los tests VBA existentes en staging (`Test_*` por cada `Helper_*` + `tests.vba.json` + `tests.vba.responsable-71.json`) son **evidencia de comportamiento** que se preserva como referencia para los nuevos tests pytest. Ningún `Test_*` se descarta sin trazabilidad | APROBADO | `blueprint/staging-expedientes-hexagonal-shape` |

Reglas operativas de este subregistro:

- Una decisión `APROBADO` cambia solo si el usuario lo revierte.
- Un `PROVISIONAL` se revisa cuando IT o cumplimiento entreguen los criterios definitivos.
- Un `FUTURO` permanece como opción no seleccionada hasta que se apruebe explícitamente.
- Un `ABIERTO` no se rellena aquí; vive en la sección "Preguntas abiertas" hasta obtener respuesta.

## Disposiciones finales sobre Lanzadera

Las decisiones **D51–D56** cierran el debate abierto por capacidad detectada en el Lote 1 (Lanzadera). Se complementan con la lógica de Lanzadera como módulo de administración (D7) y con la política de UAT y ciclo de credencial. Resumen operativo:

- **Preservar y modernizar** identidad, catálogo de aplicaciones, registro de usuarios, asignaciones usuario-aplicación y auditoría de autenticación/apertura.
- **Retirar** formación (vídeos y cuestionarios), mecanismo de lanzamiento Access/Shell, segmentación oficina/fuera y gestión técnica de backend.
- **Reemplazar mecanismo** mediante servicios compartidos: cola de correo (servicio unificado), tasks flags (scheduler unificado), hash legacy (rehash transparente).
- **ABIERTAS** dentro del ciclo UAT: diseño detallado de visibilidad, entorno, aprobaciones y promoción.

El detalle vive en `docs/02-topologia-ecosistema/lanzadera-identidad-permisos.md`, `docs/03-aplicaciones/lanzadera/capabilities.md` y `docs/03-aplicaciones/lanzadera/integrations-security.md`.

## Preguntas abiertas (cinco de mayor valor, tomadas de `exploration.md`)

| # | Pregunta | Bloquea | Lote relacionado |
|---|---|---|---|
| P1 | ¿Cuál es el repositorio/binario principal de **HPS_Solicitudes** y debe documentarse como aplicación independiente de HPS? | Parcial (checkout localizado el 2026-08-05) | Lote 8 |
| P2 | ¿Cuál es la ruta y el nombre del backend vigente de **Condor** y **Brass**, y qué entorno debe considerarse baseline? | Sí | Lotes 6 y 7 |
| P3 | ¿Se autoriza configurar solo lectura los targets Dysflow de **Condor**, **Brass** y **Expedientes**, y diagnosticar el fallo de inventario de **No Conformidades**? | Sí | Lotes 5, 6, 7 y 2 |
| P4 | ¿Qué lote debe priorizarse después de aprobar Lanzadera y Expedientes: **Gestion_Riesgos**, **HPS** o **No Conformidades**? | Parcial | Lote 3, 4 o 5 |
| P5 | ¿Qué catálogo de `TbAplicaciones` y qué versión de backends debe considerarse el **baseline operativo actual**? | Sí | Lote 1 y transversal |

> **Cierre 2026-08-05** (revisión del prompt externo + descubrimiento de rutas):
> - **P1 resuelta** por **D83**: HPS_Solicitudes se documenta como aplicación independiente (ID 22) distinta de HPS (ID 17).
> - **P2 resuelta** por **D84**: baseline de Condor (`00_CONDOR\00_main\CONDOR.accdb` + `condor_datos.accdb`) y Brass (`00_BRASS\00_main\Gestion_Brass_Gestion.accdb` + `Gestion_Brass_Gestion_Datos.accdb`). Brass sin `staging/`; requiere clarificación de rama.
> - **P3 autorizada**: Dysflow read-only configurado contra Condor, Brass y Expedientes; diagnóstico de inventario de No Conformidades en curso.
> - **P4 resuelta**: Lote siguiente = **Gestion_Riesgos** (Lote 3).
> - **P5 resuelta** por **D85**: catálogo baseline con los 8 IDs y rutas en `C:\00repos\datos\`.

## Preguntas abiertas derivadas de las decisiones de arquitectura

Estas preguntas reflejan los puntos que `09-arquitectura-objetivo-y-principios.md` marca como `ABIERTO` o `FUTURO`. No bloquean el discovery legacy; bloquean fases SDD posteriores (selección de stack, despliegue, compliance). Se registran aquí para no perderlas.

| # | Pregunta | Bloquea | Notas |
|---|---|---|---|
| P6 | ¿Qué tecnología concreta de caché se usará detrás del puerto de caché? | Selección de stack | **Caché selectiva y justificada por medición (D70)**. Redis es una opción detrás del puerto (D71), no una selección. La caché NO se introduce proactivamente. |
| P7 | ¿Cuál es la topología de despliegue objetivo (on-premise, nube corporativa, OCP u otro)? | Selección de stack / despliegue | **Kubernetes/OpenShift DIFERIDO (D74)**. Topología ABIERTA (D75). El adaptador de scheduler y los demás puertos sobreviven al cambio de topología sin tocar el dominio. PostgreSQL gestionado preferido (D76). |
| P8 | ¿Cuál es el stack exacto de implementación (framework, librerías, runtime)? | Diseño de módulos | **Resuelta por D66/D67/D68**: Python 3.12+ / FastAPI 0.119+ / Pydantic v2 / SQLAlchemy 2.0.x (mixto ORM/Core) / Alembic 1.13+ / asyncpg 0.30+ / HTMX 2.0.4 / Jinja2 3.1+ (async) / Alpine.js 3.15+. Versiones verificadas en context7 (todas activamente mantenidas). |
| P9 | ¿Cómo se descompone la plataforma (monolito modular vs microservicios vs mixto)? | Diseño y boundaries | Depende de capacidades y equipos. |
| P10 | ¿Cuáles son los periodos definitivos de retención por cumplimiento normativo o política de IT corporativa? | Política de auditoría | Sustituye al baseline provisional 90 días hot + 1 año total. Aplica también a artefactos de informe y evidencia de entrega. |
| P11 | ¿Qué integraciones corporativas se confirman (correo, identidad, monitorización)? | Adaptadores driven | SiteMinder/OCP/JWT, proveedor de email corporativo, etc. |
| P12 | ¿Los hashes de contraseña heredados se migran tal cual o se exige reset? | Adaptador de autenticación | **Resuelto por D88-D91 2026-08-08**: Argon2id vía `argon2-cffi==25.1.0`; D36+D37 marcados OBSOLETO. Reset forzado one-shot para los 156 usuarios; sin columna `legacy_hash`. SHA256 sin salt del legacy no cumple OWASP 2024. |
| P13 | ¿Qué funciones exactas tendrá el "responsable de aplicación" más allá de la administración delegada? | Modelo de roles | Pendiente de descubrimiento por módulo. |
| P14 | ¿Cuál es el catálogo definitivo de health-checks (métricas, umbrales, severidades)? | Operación | Se construye tras Lotes 1–9. |
| P15 | ¿Cuál es la estrategia de migración de datos desde `.accdb` a PostgreSQL? | Migración de datos | No se aborda en esta fase de discovery. |
| P16 | ¿Cuál es el diseño exacto del ciclo UAT (workflow, visibilidad, entorno, aprobaciones, promoción)? | Release governance | Diseño posterior; la gobernanza global-admin-only ya está APROBADA (D48). |
| P17 | ¿Cómo se representa la coexistencia UAT y producción en el menú global (visual, routing, autorización)? | Navegación | Diferida por D28 del Lote 1; **además el despliegue coexistente estable+RC está DIFERIDO por D81** hasta que cadencia y equipo lo justifiquen. No se asume el patrón legacy de IDs duplicados. |
| P18 | ¿Cuál es la UX exacta del dashboard de operaciones de notificación (filtros, contenido sensible, umbrales, acciones operativas)? | Operación de notificaciones | Diseño posterior; la dirección está APROBADA (D64). |
| P19 | ¿Cuál es el catálogo final de variantes de UI por módulo? | Diseño de módulos | Decidir durante cada discovery; Expedientes ya anticipa vistas especializadas. |
| P20 | ¿Cuál es el contrato final del proveedor de email corporativo (host, remitente, entregabilidad)? | Adaptador de notificación | No se prefija; la cola por tabla es el adapter v1. |

## Gaps abiertos en el SDD chain de Lanzadera MVP (2026-08-08)

Detectados durante `sdd-spec` y registrados en design.md §Decisiones pendientes. Todos son **no bloqueantes** para `sdd-tasks`; cada uno se cierra dentro de una tarea acotada.

| # | Gap | Origen | Estado | Cierre propuesto |
|---|---|---|---|---|
| G-1 | Política exacta de normalización del email (lowercase completo vs `local-part`+`domain`, IDN) | `users/spec.md` | **Cerrado** | DA-3 en `design.md`: lowercase completo en `users.email`. |
| G-2 | Set canónico de capabilities por app para los 20 IDs en alcance | `profiles/spec.md` | ABIERTO | `sdd-tasks`: `platform/src/modules/lanzadera/domain/legacy_role_map.py` con `capabilities` mínimos por código. `profiles.capabilities` se siembra en 0003 con JSONB provisional. |
| G-3 | Tabla de campos legacy que NO migran (`Pass`, `Comando`, `URLDIrectorioIconoAplicacion`) | `apps/spec.md` | **Cerrado** | DA-7 en `design.md`: 0002 deja esos campos fuera del schema. |
| G-4 | Severidad y notificación para intentos fallidos de crear admin global (SOC) | `auth-bootstrap/spec.md` | ABIERTO | Suficiente con WARN en log canónico de auditoría; severidad alta se reabre si la política cambia. |
| G-5 | Canal exacto de notificación al usuario cuando se emite un token (P20) | `auth-reset/spec.md` | ABIERTO | Adapter v1 cubre MVP; SMTP corporativo real cuando se cierre P20. |
| G-6 | Política de expiración periódica de contraseña (D41) | `auth-core/spec.md` | ABIERTO | MVP implementa `status='password_reset_required'` sin caducidad periódica; el reset flow cubre la recuperación. |
| G-7 | Peso semántico de `SinAcceso` en la navegación del menú global | `assignments/spec.md` | ABIERTO | Decisión de UI fuera del MVP Lanzadera; reabre con el módulo de menú web. |
| G-8 | Periodos definitivos de retención (P10) | `audit/spec.md` | ABIERTO | Baseline provisional 90 días hot + 1 año total (D29 PROVISIONAL) hasta que IT o cumplimiento entreguen los definitivos. |

## Reglas del registro

- Cada decisión lleva fecha, origen (usuario, exploración, lote) e impacto.
- Cada pregunta abierta lleva su lote relacionado y la condición de desbloqueo.
- Una pregunta se cierra solo cuando el usuario responde o cuando el lote aporta evidencia suficiente y revisada.
- APAP y APAP_WEB no aparecen en decisiones ni en preguntas (su única mención admisible es la exclusión del alcance en `00-alcance-y-evidencia.md`).

## Checklist

- [ ] Cada nueva decisión se añade en el momento de tomarse, no al final del lote.
- [ ] Cada pregunta abierta se cierra o se reasigna a otro lote; nada queda sin dueño.
- [ ] El registro se cruza con `02-topologia-ecosistema/matriz-dependencias.md` cuando aplica.

## Siguiente paso

Resolver P1–P5 con el usuario antes de iniciar el Lote 2; mantener P6–P20 vivos para fases SDD posteriores. **P8 cerrada** tras la revisión del prompt externo (D66–D82). P6, P7 y P17 actualizadas con las salvaguardas D70/D71, D74/D75/D76 y D81 respectivamente.
