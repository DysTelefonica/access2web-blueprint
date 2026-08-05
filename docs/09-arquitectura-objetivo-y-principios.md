# 09 · Arquitectura objetivo y principios

## Propósito

Consolida las decisiones de producto y arquitectura **ya aprobadas** para el estudio de modernización del ecosistema legacy, junto con los baselines provisionales, las opciones futuras que permanecen abiertas y las preguntas que requieren respuesta. Sirve como referencia única para los siguientes lotes de discovery y para futuras fases SDD; **no** propone todavía proposal/spec/design/tasks concretos.

Las decisiones se han extraído de Engram (`project: access2web-blueprint`) con sus `topic_key`. Cada fila del registro consolidado lleva ese identificador para poder auditar el origen.

## Estado del descubrimiento

El discovery legacy (Fase 1 por lotes, ver `openspec/changes/legacy-ecosystem-discovery/exploration.md`) **está incompleto**. La topología general y la matriz inicial de dependencias son provisionales; las capacidades reales se obtendrán al ejecutar los Lotes 1 a 9. Mientras tanto, este documento fija la **forma objetivo** del producto web y sus principios, no la réplica exacta del legacy.

El legacy define el **suelo mínimo de capacidad de negocio**, no el objetivo de UX ni de implementación. La forma moderna debe preservar la intención y adaptarla a patrones web nativos.

## Ruta rápida

1. Plataforma única, modular y permission-aware.
2. Navegación app-first anidada; Lanzadera queda restringida a administración.
3. Hexagonal real en todos los módulos; web y CLI como adaptadores driving; PostgreSQL, cola de correo por tabla, object storage, scheduler, autenticación, caché selectiva (justificada por medición, no proactiva) y almacenamiento de archivo como adaptadores driven.
4. Persistencia objetivo: PostgreSQL compartido con aislamiento por esquema.
5. Notificaciones: servicio unificado; v1 solo email sobre la cola por tabla como adaptador transitorio; la cola real la consume un dispatcher externo cada cinco minutos aprox. en el legacy.
6. Adjuntos: object storage S3-compatible detrás de un puerto; sin versionado de contenido; papelera con retención 30 días; restauración por el borrado o por administrador global.
7. CLI exclusivo para administrador global; reutiliza los mismos casos de uso y autorización que la web; sin secretos en argumentos ni en logs.
8. Logs canónicos industriales; retención por niveles configurable (90 días hot + 1 año total provisional).
9. Scheduler unificado para informes y automatizaciones; governance solo administrador global; configuración por separado del contenido de cada informe (que vive en el módulo).
10. Diagnóstico bajo demanda y programado; remediación siempre con aprobación explícita del administrador global.
11. Ciclo de credencial: hashes heredados preservados con rehash transparente al primer login; lockout configurable (umbral cinco por defecto, una hora por defecto, desbloqueo por admin global); activación y cambios de rol solo por admin global; notificación manual de cambios de permiso.
12. Suplantación solo por administrador global, con doble identidad visible y auditoría completa.
13. UAT gobernado por admin global; asignación explícita por ciclo; excepciones auditadas y visibles para usuarios.
14. Registro de nuevos módulos híbrido: el despliegue crea un registro "pendiente" con metadatos técnicos; el admin global activa y configura la metadatos funcionales.

## Forma de la plataforma

| Decisión | Estado | Origen |
|---|---|---|
| Producto único, modular y permission-aware: cada usuario ve solo los módulos para los que tiene permiso. | APROBADO | `architecture/target-platform-shape` |
| Navegación app-first anidada: barra global con módulos visibles; cada módulo expone su menú interno. El expediente no es punto de entrada. | APROBADO | `architecture/navigation-model` |
| Lanzadera queda restringida a administración de plataforma (usuarios y permisos por aplicación). Los usuarios normales se autentican una vez y solo ven sus módulos autorizados. | APROBADO | `architecture/identity-and-admin-module` |

## Arquitectura hexagonal global

| Decisión | Estado | Origen |
|---|---|---|
| Hexagonal es obligatorio en toda la plataforma y en cada módulo, no solo en autenticación. Las capacidades dependen de puertos, no de mecanismos de infraestructura. | APROBADO | `architecture/global-hexagonal-principle` |
| Adaptadores **driving**: interfaz web y CLI administrativo. | APROBADO | consolidación `ai-cli-first` + navegación |
| Adaptadores **driven** reemplazables: PostgreSQL, cola de correo por tabla, object storage, scheduler, proveedores de autenticación, caché y almacenamiento de archivo. | APROBADO | consolidación `global-hexagonal-principle`, `attachment-storage`, `authentication-ports-adapters`, `notification-delivery-adapter-v1`, `audit-retention`, `scheduled-health-checks`, `performance-and-cache` |
| Autenticación: adaptador inicial email/password migrado desde `Lanzadera_Datos.accdb`. Los hashes heredados se preservan y se verifican con un adapter de verificación legacy versionada. Un adaptador futuro e intercambiable podrá integrarse con SiteMinder y JWT unificado si la plataforma se aloja en OCP corporativo. SiteMinder y OCP son **opciones futuras**, no dependencias iniciales. | APROBADO (adaptador inicial) / FUTURO (SiteMinder/OCP) | `architecture/authentication-ports-adapters` + `product/credential-migration` |
| La cola de correo por tabla del legacy se mantiene como **adaptador transitorio** del servicio unificado de notificaciones; se sustituirá por la integración corporativa cuando IT defina su contrato. La cola real la consume un dispatcher externo cada cinco minutos aprox. en el legacy. | APROBADO | `architecture/notification-delivery-adapter-v1` + `discovery/legacy-email-queue-flow` |
| El CLI nunca duplica lógica de negocio: invoca los mismos casos de uso y pasa por la misma autorización server-side que la web. | APROBADO | `architecture/ai-cli-first` |

## Stack técnico

| Decisión | Estado | Origen |
|---|---|---|
| **Backend**: Python 3.12+ con FastAPI 0.119+, Pydantic v2, SQLAlchemy 2.0.x en modo mixto (ORM predominante + Core para queries complejas, CTEs recursivos y jerarquías), Alembic 1.13+ para migraciones DDL, driver asyncpg 0.30+ (D66). | APROBADO | `external-prompt-review/stack-versions-verified` (D66) |
| **Frontend**: HTMX 2.0.4 + Jinja2 3.1+ en modo async + Alpine.js 3.15+ (D67). SSR puro, sin SPA, sin build pipeline. CSS plano. La API de SSE de HTMX cambió en 2.0 (extensión fuera del core); no la usamos. | APROBADO | `external-prompt-review/stack-versions-verified` (D67) |
| **Estructura del repositorio**: monorepo `access2web-blueprint/` con monolito modular (D68). Límites de módulos por paquete y ports por módulo. Tests con pytest + pytest-asyncio + httpx; Playwright para flujos críticos de UI. | APROBADO | `external-prompt-review/section-2-resolution` (D68) |
| Todas las librerías verificadas como **activamente mantenidas** en context7 (2026-08). Sin librerías zombies. | APROBADO | `external-prompt-review/stack-versions-verified` |

## Persistencia y rendimiento

| Decisión | Estado | Origen |
|---|---|---|
| Persistencia objetivo preferida: una base de datos PostgreSQL compartida; cada módulo dueño de su esquema; servicios comunes en esquema(s) dedicado(s). | APROBADO (preferencia/dirección) | `architecture/target-database-topology` |
| Una base de datos física compartida no implica propiedad compartida: los límites de esquema y las reglas de acceso deben preservar el aislamiento hexagonal y modular. | APROBADO | `architecture/target-database-topology` |
| Rendimiento y caché deliberada como preocupación de producto y arquitectura de primer orden; se elimina la percepción de lentitud del legacy. | APROBADO | `architecture/performance-and-cache` |
| Caché de aplicación **selectiva y justificada por medición** (D70). Candidatos naturales: catálogos estables (tipos, países, provincias, plantillas de informe), permisos efectivos precalculados y diccionarios. **NO** se cachean contadores de pendientes ni métricas de dashboard que mutan con cada acción de usuario. La lentitud legacy es de plataforma, no de datos: PostgreSQL bien indexado resuelve la mayoría sin caché. | APROBADO | `external-prompt-review/section-3-resolution` (D70) |
| Redis queda como **opción detrás del puerto de caché**, no como dependencia inicial (D71). Pub/Sub se introduce solo si la escala horizontal lo justifica. | APROBADO | `external-prompt-review/section-3-resolution` (D71) |
| Actualización de contadores pendientes vía polling HTMX (`hx-trigger="every 30s"`, intervalo configurable) con botón de refresh manual. Sin SSE, sin WebSockets, sin Redis pub/sub (D69). | APROBADO | `external-prompt-review/section-1-resolution` (D69) |
| Rendimiento HTTP: ETag + `304 Not Modified` para fragmentos HTML servidos por HTMX; compresión gzip/brotli por defecto (Starlette); `Cache-Control` correcto en assets estáticos con fingerprint en el nombre (D72). | APROBADO | `external-prompt-review/section-3-resolution` (D72) |

## Infraestructura y despliegue

| Decisión | Estado | Origen |
|---|---|---|
| **Hexagonal primero**: los detalles de infraestructura (cómo se ejecuta, dónde corre, qué cloud) son decisiones de adaptador, no de producto. El dominio y los casos de uso no saben dónde corren. | APROBADO (heredado) | `architecture/global-hexagonal-principle` |
| **Contenedores Docker desde el día uno**: un `Dockerfile` por servicio (backend, scheduler, worker de cola, etc.) y `docker-compose.yml` para desarrollo local con PostgreSQL + MinIO (object storage local) + servicios auxiliares (D77). | APROBADO | `external-prompt-review/section-4-resolution` (D77) |
| **El backend hexagonal es nuestro** (FastAPI + adaptadores propios). **No usar Insforge como BaaS** ni como sustituto del backend (D73). Insforge puede ser herramienta auxiliar para prototipos, nunca dependencia. | APROBADO | `external-prompt-review/section-4-resolution` (D73) |
| **No introducir Kubernetes ni OpenShift prematuramente** (D74). Para 200 usuarios concurrentes, una instancia de FastAPI + PostgreSQL es suficiente. La introducción de orquestador se justifica con métricas reales de carga y con requerimientos de IT corporativa. | APROBADO | `external-prompt-review/section-4-resolution` (D74) |
| **Topología de despliegue ABIERTA** (D75, consistente con P7). Decisión de cloud y orquestador queda pendiente de métricas reales + IT corporativa. Los puertos hexagonales (scheduler, object storage, base de datos) sobreviven al cambio de topología sin tocar el dominio. | APROBADO (apertura) | `external-prompt-review/section-4-resolution` (D75) |
| **PostgreSQL gestionado preferido** sobre auto-instalado cuando se decida el cloud (D76). Proveedor concreto (Cloud SQL, RDS, on-premise) se liga a D75. | APROBADO | `external-prompt-review/section-4-resolution` (D76) |
| **Desarrollo**: local con Docker Compose (preferido) o VPS corporativo con Coolify como panel de despliegue opcional. | APROBADO | `external-prompt-review/section-4-resolution` |

## Notificaciones

| Decisión | Estado | Origen |
|---|---|---|
| Un único servicio unificado de notificaciones compartido por todos los módulos. Los módulos aportan el disparador de negocio, destinatarios y mensaje semántico; canales, reintentos, plantillas, observabilidad e integración con proveedores quedan detrás del puerto. | APROBADO | `architecture/unified-notification-service` |
| Primera versión de release: **solo email** (paridad mínima con el canal de comunicación legacy exigido). Notificaciones in-app y otros canales quedan fuera del alcance inicial. | APROBADO | `product/notification-v1-scope` |
| Proveedor de email corporativo concreto. | ABIERTO (FUTURO considerado) | — |
| Dashboard global de operaciones de notificación (salud de cola, estadísticas de entrega/fallo, reintentos controlados) para el administrador global. UX exacta, filtros, acceso a contenido sensible, retención, umbrales y acciones operativas. | APROBADO (dirección) / ABIERTO (diseño) | `architecture/notification-operations-dashboard` |
| Retención de artefactos de informe y evidencia de entrega. | PROVISIONAL (mismo baseline que auditoría) / ABIERTO (definitivo) | `product/audit-retention-periods` |

## Scheduler y jobs

| Decisión | Estado | Origen |
|---|---|---|
| Scheduler unificado para jobs programados y procesos batch; los flags de tareas de Lanzadera y los mecanismos por aplicación se retiran. El motor concreto (cron-like inicial; reemplazable) vive detrás de un port. | APROBADO | `architecture/shared-scheduler-operations` |
| Solo el administrador global configura y supervisa los jobs: periodicidad, frecuencia, severidad mínima de alerta. | APROBADO | `architecture/health-check-configuration-authorization` |
| Configuración por separado del contenido: los horarios y destinatarios de informes los gestiona el administrador global desde el panel de operaciones; la lógica de negocio del informe vive y se prueba dentro del módulo. | APROBADO | `architecture/report-job-configuration` |
| Ejecución manual bajo demanda: el administrador global puede ejecutar cualquier informe configurado además de su ejecución programada. | APROBADO | `architecture/manual-report-execution` |
| Vista previa sin envío: el administrador global puede generar el artefacto exacto sin seleccionar destinatarios, sin encolar email ni enviar. | APROBADO | `architecture/report-preview` |
| Generar y enviar directamente: cuando la urgencia lo justifique, se salta la vista previa mostrando destinatarios y parámetros antes de la confirmación explícita. | APROBADO | `architecture/direct-manual-report-send` |
| Auditoría completa de cada ejecución: correlación, parámetros, destinatarios, resultado, reintentos, identificadores y referencia al artefacto. | APROBADO | `architecture/observability-audit-logs` |
| Health inspections disponibles **bajo demanda** (CLI) y **programadas** (cron-like). Los jobs programados solo miden calidad y salud; no remedian automáticamente. | APROBADO | `architecture/scheduled-health-checks` |
| Toda remediación requiere aprobación explícita del administrador global, incluso para acciones clasificadas como seguras. | APROBADO | `architecture/ai-remediation-approval` |
| Destinatarios de anomalías configurables por aplicación. La configuración global de health checks la hace solo el administrador global. | APROBADO | `architecture/application-operations-settings` + `architecture/health-check-configuration-authorization` |

## Adjuntos

| Decisión | Estado | Origen |
|---|---|---|
| Almacenamiento en object storage S3-compatible detrás de un puerto de almacenamiento; los módulos no dependen de APIs, semánticas de bucket ni URLs S3. | APROBADO (puerto y compatibilidad S3) / ABIERTO (proveedor final) | `architecture/attachment-storage` |
| Sin historial de versiones de contenido: solo se conserva el fichero activo por referencia de adjunto. | APROBADO | `product/attachment-versioning` |
| Los adjuntos borrados pasan a una papelera con retención limitada antes del borrado permanente. | APROBADO | `product/attachment-deletion` |
| Ventana de recuperación: **30 días**; pasado ese plazo, un proceso programado purga definitivamente el contenido y el estado recuperable. | APROBADO | `product/attachment-retention` |
| Durante la papelera solo pueden restaurar: la persona que borró el adjunto o un **administrador global** de plataforma. Otros usuarios con permiso de edición del registro padre **no** pueden restaurar. | APROBADO | `product/attachment-restore-authorization` |

## Ciclo de credencial y autenticación

| Decisión | Estado | Origen |
|---|---|---|
| Preservar hashes heredados de Lanzadera en la migración; verificación legacy versionada; los campos con credenciales en claro no migran como secretos. | APROBADO | `product/credential-migration` |
| Rehash transparente al primer login exitoso: la política moderna sustituye al hash heredado de forma atómica, idempotente y auditable. | APROBADO | `architecture/opportunistic-password-rehash` |
| Umbral de lockout configurable por el administrador global; valor por defecto cinco intentos. | APROBADO | `architecture/login-lockout-policy` |
| Duración de lockout de una hora por defecto; el administrador global puede desbloquear antes. | APROBADO | `architecture/login-lockout-recovery` |
| Notificación de lockout a los administradores globales por el servicio unificado, con contexto seguro e identificadores de correlación. | APROBADO | `architecture/lockout-notification` |
| Caducidad de contraseña configurable globalmente; posibilidad de eximir a un usuario concreto; "sin caducidad periódica" como valor válido. | APROBADO | `architecture/password-expiry-policy` |

## Activación, permisos y notificación

| Decisión | Estado | Origen |
|---|---|---|
| Solo el administrador global registra usuarios, los da de baja y crea roles por aplicación. Los administradores de aplicación asignan usuarios activos a roles existentes. | APROBADO | `architecture/global-only-identity-actions` |
| La notificación de cambios de permiso la dispara explícitamente un administrador autorizado (no es automática). | APROBADO | `architecture/manual-permission-change-notification` |

## Suplantación para pruebas

| Decisión | Estado | Origen |
|---|---|---|
| Solo el administrador global inicia sesión suplantada; desarrolladores y administradores de aplicación **no** pueden impersonar directamente. Uso normal: pruebas/UAT. | APROBADO | `architecture/impersonation-authorization` |
| La sesión suplantada muestra la doble identidad de forma visible y registra auditoría completa (actor real, persona impersonada, parámetros, marcas temporales, módulo y resultado). | APROBADO | `architecture/impersonation-authorization` |

## Capabilities, políticas y vistas

| Decisión | Estado | Origen |
|---|---|---|
| Cada módulo declara **capacidades estables** (estables, no strings ad-hoc). Los grupos de capabilities componen roles verificables. | APROBADO | `architecture/capability-driven-ui-variants` |
| Las **políticas contextuales** que dependen de estado, recurso o condición de negocio se evalúan en el código del módulo, no en el catálogo. Capabilities y políticas coexisten. | APROBADO | `architecture/capability-driven-ui-variants` |
| El backend es autoritativo: la UI consume el endpoint de capabilities/políticas y el servidor rechaza operaciones no autorizadas aunque la UI las muestre. | APROBADO | `architecture/capability-driven-ui-variants` |
| Cada módulo decide por sí mismo si usa una vista adaptativa o varias especializadas; por defecto, vista única. Las vistas especializadas se justifican por diferencias materiales de flujo. | APROBADO | `architecture/per-module-ui-variant-policy` |

## UAT, releases y excepciones

| Decisión | Estado | Origen |
|---|---|---|
| Cada ciclo UAT declara explícitamente sus participantes y perfil de aplicación. El acceso UAT **no** espeja automáticamente el de producción. | APROBADO | `product/uat-participant-governance` |
| Solo el administrador global define quién participa, qué perfil recibe cada participante y la configuración de acceso del ciclo. Los administradores de aplicación no configuran participación UAT. | APROBADO | `architecture/uat-global-admin-governance` |
| Excepciones auditadas: un administrador global puede liberar con casos UAT fallidos o sin UAT ejecutado, registrando motivo, atribución y marca temporal con la evidencia del release. | APROBADO | `architecture/uat-release-exceptions` |
| Excepciones visibles para usuarios: cuando un release publica con casos UAT fallidos o sin UAT, la excepción y su justificación aparecen en el historial de cambios. | APROBADO | `product/release-exception-transparency` |
| Diseño detallado del ciclo UAT (workflow, visibilidad, entorno, aprobaciones, promoción). | ABIERTO | `product/module-uat-lifecycle` |
| Navegación dual UAT + producción simultánea (representación, routing, autorización, marca visual). | ABIERTO (diferida por D81) | `open/dual-environment-navigation` + `external-prompt-review/section-5-resolution` |

## Versionado y estrategia de releases

| Decisión | Estado | Origen |
|---|---|---|
| **Versionado semántico por módulo Y plataforma base** (D78): formato `modulo/vX.Y.Z-rc.n` para Candidate Releases y `modulo/vX.Y.Z` para estables; la plataforma base usa el mismo formato (`platform/vX.Y.Z`). Cada módulo y la plataforma publican su propio ritmo. | APROBADO | `external-prompt-review/section-5-resolution` (D78) |
| **Branching trunk-based development** (D79): `main` siempre desplegable; feature branches de vida corta; tags en `main`, no en branches. | APROBADO | `external-prompt-review/section-5-resolution` (D79) |
| **Conventional Commits** como entrada al versionado y al changelog automático (D79). | APROBADO | `external-prompt-review/section-5-resolution` (D79) |
| **Catálogo de versiones compatibles** entre módulos y plataforma (D80): la plataforma expone qué versión de cada módulo es compatible con qué versión de plataforma, evitando combinaciones inválidas en despliegues. | APROBADO | `external-prompt-review/section-5-resolution` (D80) |
| **UAT y Producción como entornos separados** en esta fase. El despliegue coexistente estable+RC simultáneo en UAT queda **DIFERIDO** (D81) hasta que cadencia de releases y tamaño del equipo lo justifiquen. La coexistencia UAT + producción simultánea sigue ABIERTA en P17. | APROBADO (diferimiento) | `external-prompt-review/section-5-resolution` (D81) |
| **Migraciones de BD backward-compatibles con estrategia Expand and Contract** (D82): Expand añade estructura nueva sin retirar la vieja; Migrate mueve datos en background o en fase posterior; Contract retira la vieja solo cuando la nueva está en uso. **Nunca** una migración destructiva en una sola release. Alembic soporta el flujo con migraciones forward y backward explícitas. | APROBADO | `external-prompt-review/section-5-resolution` (D82) |

## Registro y activación de aplicaciones

| Decisión | Estado | Origen |
|---|---|---|
| Registro **híbrido** de nuevos módulos: el despliegue crea un registro "pendiente" con metadatos técnicos (identificador estable, versión, rutas, health endpoint, capacidades declaradas). El administrador global revisa y configura metadatos funcionales y activa la visibilidad. El despliegue **nunca** expone un módulo a usuarios por sí mismo. | APROBADO | `architecture/application-registration` |
| El registro debe ser idempotente entre redespliegues; la activación es auditable e independiente de los redespliegues. | APROBADO | `architecture/application-registration` |

## Autorización y roles

| Decisión | Estado | Origen |
|---|---|---|
| Existen dos scopes: **administrador global** de plataforma y **administrador de aplicación** limitado al módulo. Además, cada aplicación tiene una persona responsable. | APROBADO | `architecture/authorization-roles` |
| Las responsabilidades del administrador de aplicación se definen **por módulo**, no por un único contrato universal. Ejemplo: en Gestion_Riesgos el administrador de aplicación asigna qué usuarios trabajan en cada instancia de riesgo; otros módulos exponen permisos delegados distintos y algunos módulos pueden no exponer funciones de admin de aplicación en su versión inicial. | APROBADO | `architecture/module-admin-capabilities` |
| Solo el **administrador global** nombra administradores de aplicación, y solo para los módulos donde aplique administración delegada. Los nombramientos son acciones auditables independientes de las capacidades administrativas del módulo. | APROBADO | `architecture/module-admin-assignment` |
| Para la recuperación de adjuntos, el término "administrador" significa **administrador global** de plataforma. | APROBADO | `architecture/authorization-roles` + `product/attachment-restore-authorization` |

## CLI administrativo para IA

| Decisión | Estado | Origen |
|---|---|---|
| Cada módulo expone un CLI operable por un agente IA para ejecutar **cualquier** acción soportada por el módulo, con credenciales por variables de entorno. | APROBADO | `architecture/ai-cli-first` |
| El CLI es exclusivo del **administrador global** de plataforma. Administradores de aplicación, responsables y usuarios normales no reciben acceso CLI. | APROBADO | `architecture/cli-access-control` |
| Acciones destructivas o de impacto masivo invocadas por CLI requieren una **confirmación explícita adicional** aun con credencial de administrador global autorizada. | APROBADO | `architecture/cli-destructive-confirmation` |
| Los secretos nunca aparecen como argumentos ni se registran en logs; entran por variables de entorno o por adaptadores de proveedor de secretos reemplazables. | APROBADO | `architecture/ai-cli-first` |
| La autorización se aplica **server-side** y se audita en cada acción CLI; el CLI no es un superusuario compartido sin trazabilidad. | APROBADO | `architecture/cli-access-control` |

## Observabilidad y auditoría

| Decisión | Estado | Origen |
|---|---|---|
| Toda la plataforma emite logs estructurados canónicos de calidad industrial, con correlación de acciones de usuario/IA a través de web, CLI, módulos, servicios compartidos, adaptadores, colas y jobs. | APROBADO | `architecture/observability-audit-logs` |
| Eventos canónicos con: ID de correlación/traza, actor, acción, objetivo, resultado, timestamp, módulo y contexto de error seguro; credenciales y cargas sensibles redactadas. | APROBADO | `architecture/observability-audit-logs` |
| Retención por niveles configurable (no hard-coded): recientes en hot consultable; antiguos comprimidos y archivados en object storage S3-compatible a través de un puerto de archivo reemplazable. | APROBADO (mecanismo) | `architecture/audit-retention` |
| Baseline provisional: **90 días en hot consultable + 1 año de retención total**, con archivo en object storage. **Misma política** aplica a artefactos de informe y a evidencia de entrega. | PROVISIONAL | `product/audit-retention-periods` |
| Periodos de retención definitivos por cumplimiento normativo o IT corporativa. | ABIERTO | — |

## Operaciones y salud

| Decisión | Estado | Origen |
|---|---|---|
| El CLI administrativo expone diagnósticos operativos explícitos: inspección de logs canónicos, salud actual, evaluación de corrección, detección de anomalías/degradación, explicación de evidencia y soporte de remediación proactiva. | APROBADO | `architecture/ai-operations-cli` |
| El CLI puede diagnosticar, correlacionar evidencia y proponer acciones correctivas, pero **siempre se detiene y espera aprobación explícita** del administrador global antes de ejecutar cualquier remediación, incluso en acciones clasificadas como seguras. | APROBADO | `architecture/ai-remediation-approval` |
| Inspecciones de salud disponibles **bajo demanda** (CLI) y **programadas** (cron-like); los jobs programados solo miden calidad y salud, no remedian automáticamente. | APROBADO | `architecture/scheduled-health-checks` |
| Cada aplicación expone un área de administración donde se configuran los destinatarios de anomalías de ese módulo. | APROBADO | `architecture/application-operations-settings` |
| Solo el **administrador global** configura health-checks (destinatarios, checks activos, frecuencia, severidad mínima de alerta). Los administradores de aplicación no gestionan esta configuración operativa. | APROBADO | `architecture/health-check-configuration-authorization` |

## Disposiciones sobre Lanzadera

| Decisión | Estado | Origen |
|---|---|---|
| Preservar y modernizar identidad, catálogo de aplicaciones, registro de usuarios, asignaciones usuario-aplicación y auditoría de autenticación / apertura. | APROBADO | `product/lanzadera-core-capabilities` |
| Retirar formación (vídeos, cuestionarios, visionados) y reproducciones ActiveX; el histórico queda archivado, no se migra al módulo operativo. | APROBADO | `product/lanzadera-training-disposition` |
| Retirar mecanismo de lanzamiento Access (`Shell`, `/cmd`, copia/ejecutable, UNC); el menú web permission-aware sustituye el lanzador desktop. | APROBADO | `product/lanzadera-launcher-disposition` |
| Retirar segmentación oficina / fuera de oficina; el control pre-producción será un ciclo UAT moderno. | APROBADO | `product/lanzadera-location-visibility` |
| Retirar gestión de rutas y contraseñas de backend desde UI; la configuración técnica se externaliza a adapters. | APROBADO | `product/lanzadera-backend-config-disposition` |
| Modernizar auditoría de Lanzadera: preservar eventos de autenticación y apertura; retirar telemetría de SSID / ubicación física / coordenadas. | APROBADO | `product/lanzadera-audit-disposition` |
| UAT en Lanzadera: el ciclo detallado (workflow, visibilidad, entorno, aprobaciones, promoción) queda ABIERTO; la gobernanza global-admin-only ya está APROBADA. | APROBADO (apertura de diseño) | `product/module-uat-lifecycle` |

## Resultado del estudio

| Decisión | Estado | Origen |
|---|---|---|
| El estudio termina con un **roadmap global** + un **roadmap detallado y plan de implementación por herramienta**. | APROBADO (resultado objetivo) | `product/modernization-principles` |
| El comportamiento legacy marca el suelo mínimo de capacidad de negocio, no un requisito de paridad exacta de funcionalidad o UI. Las funcionalidades se preservan en intención y se adaptan a flujos web nativos. | APROBADO | `product/modernization-principles` |
| Decisiones de producto e implementación requieren cuestionamiento iterativo y validación explícita del usuario. | APROBADO | `product/modernization-principles` |

## Mapa rápido de decisiones

| # | Tema | Estado | Origen |
|---|---|---|---|
| D5 | Plataforma única modular permission-aware | APROBADO | `architecture/target-platform-shape` |
| D6 | Navegación app-first anidada | APROBADO | `architecture/navigation-model` |
| D7 | Lanzadera = solo administración | APROBADO | `architecture/identity-and-admin-module` |
| D8 | Hexagonal global | APROBADO | `architecture/global-hexagonal-principle` |
| D9 | Autenticación: adaptador inicial email/password | APROBADO | `architecture/authentication-ports-adapters` |
| D10 | SiteMinder/OCP/JWT unificado como adaptador futuro | FUTURO | `architecture/authentication-ports-adapters` |
| D11 | Notificación unificada como servicio compartido | APROBADO | `architecture/unified-notification-service` |
| D12 | Notificación v1 solo email | APROBADO | `product/notification-v1-scope` |
| D13 | Cola de correo por tabla como adaptador transitorio | APROBADO | `architecture/notification-delivery-adapter-v1` |
| D14 | PostgreSQL compartido con esquemas por módulo | APROBADO (dirección) | `architecture/target-database-topology` |
| D15 | Rendimiento y caché deliberada como prioridad de primer orden | APROBADO | `architecture/performance-and-cache` |
| D16 | Adjuntos en object storage S3-compatible detrás de puerto | APROBADO | `architecture/attachment-storage` |
| D17 | Adjuntos sin versionado de contenido | APROBADO | `product/attachment-versioning` |
| D18 | Adjuntos con papelera de retención limitada | APROBADO | `product/attachment-deletion` |
| D19 | Retención de papelera = 30 días | APROBADO | `product/attachment-retention` |
| D20 | Restauración de papelera por borrado o administrador global | APROBADO | `product/attachment-restore-authorization` |
| D21 | Roles: administrador global + administrador de aplicación por módulo | APROBADO | `architecture/authorization-roles` |
| D22 | Capacidades de admin de aplicación definidas por módulo | APROBADO | `architecture/module-admin-capabilities` |
| D23 | Solo administrador global nombra administradores de aplicación | APROBADO | `architecture/module-admin-assignment` |
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
| D36 | Preservar hashes heredados en migración | APROBADO | `product/credential-migration` |
| D37 | Rehash transparente al primer login exitoso | APROBADO | `architecture/opportunistic-password-rehash` |
| D38 | Umbral de lockout configurable, default cinco | APROBADO | `architecture/login-lockout-policy` |
| D39 | Bloqueo una hora por defecto + desbloqueo por admin global | APROBADO | `architecture/login-lockout-recovery` |
| D40 | Notificación de lockout a administradores globales | APROBADO | `architecture/lockout-notification` |
| D41 | Caducidad de contraseña configurable + exenciones por usuario | APROBADO | `architecture/password-expiry-policy` |
| D42 | Activación, baja y creación de roles: solo admin global | APROBADO | `architecture/global-only-identity-actions` |
| D43 | Notificación de cambios de permiso solo por acción del administrador | APROBADO | `architecture/manual-permission-change-notification` |
| D44 | Suplantación restringida al administrador global | APROBADO | `architecture/impersonation-authorization` |
| D45 | Capabilities declaradas por módulo + grupos | APROBADO | `architecture/capability-driven-ui-variants` |
| D46 | Vista única por defecto; especializada por módulo | APROBADO | `architecture/per-module-ui-variant-policy` |
| D47 | UAT: asignación explícita de participantes y perfil por ciclo | APROBADO | `product/uat-participant-governance` |
| D48 | UAT: gobernanza reservada al admin global | APROBADO | `architecture/uat-global-admin-governance` |
| D49 | UAT: excepciones auditadas con la evidencia del release | APROBADO | `architecture/uat-release-exceptions` |
| D50 | UAT: excepciones visibles para usuarios | APROBADO | `product/release-exception-transparency` |
| D51 | Lanzadera retira formación/vídeos/cuestionarios | APROBADO | `product/lanzadera-training-disposition` |
| D52 | Lanzadera retira mecanismo de lanzamiento Access | APROBADO | `product/lanzadera-launcher-disposition` |
| D53 | Lanzadera retira segmentación oficina / fuera de oficina | APROBADO | `product/lanzadera-location-visibility` |
| D54 | Lanzadera retira gestión de rutas y contraseñas de backend | APROBADO | `product/lanzadera-backend-config-disposition` |
| D55 | Lanzadera moderniza auditoría y retira telemetría de ubicación | APROBADO | `product/lanzadera-audit-disposition` |
| D56 | Lanzadera preserva identidad, catálogo, usuarios y permisos | APROBADO | `product/lanzadera-core-capabilities` |
| D57 | Lanzadera: ciclo UAT detallado ABIERTO (gobernanza APROBADA) | APROBADO (apertura) | `product/module-uat-lifecycle` |
| D58 | Registro de aplicaciones híbrido (deployment técnico + activación global) | APROBADO | `architecture/application-registration` |
| D59 | Scheduler unificado sustituye flags y tareas de aplicación | APROBADO | `architecture/shared-scheduler-operations` |
| D60 | Configuración global de horarios y destinatarios de informes | APROBADO | `architecture/report-job-configuration` |
| D61 | Ejecución manual bajo demanda de informes | APROBADO | `architecture/manual-report-execution` |
| D62 | Vista previa de informe sin envío | APROBADO | `architecture/report-preview` |
| D63 | Generar y enviar directamente sin vista previa cuando proceda | APROBADO | `architecture/direct-manual-report-send` |
| D64 | Dashboard global de operaciones de notificación | APROBADO (dirección) | `architecture/notification-operations-dashboard` |
| D65 | Evidencia legacy: cola por tabla + dispatcher externo cada 5 min | APROBADO | `discovery/legacy-email-queue-flow` |
| D66 | Stack backend: Python 3.12+ / FastAPI 0.119+ / Pydantic v2 / SQLAlchemy 2.0.x (mixto ORM/Core) / Alembic 1.13+ / asyncpg 0.30+ | APROBADO | `external-prompt-review/stack-versions-verified` |
| D67 | Stack frontend: HTMX 2.0.4 + Jinja2 3.1+ (async) + Alpine.js 3.15+ (SSR puro, sin SPA) | APROBADO | `external-prompt-review/stack-versions-verified` |
| D68 | Estructura: monorepo + monolito modular, límites por paquete y ports por módulo | APROBADO | `external-prompt-review/section-2-resolution` |
| D69 | Polling HTMX (`hx-trigger="every 30s"`) + botón de refresh manual para contadores pendientes | APROBADO | `external-prompt-review/section-1-resolution` |
| D70 | Caché selectiva justificada por medición. NO contadores ni métricas volátiles | APROBADO | `external-prompt-review/section-3-resolution` |
| D71 | Redis como opción detrás del puerto de caché, no dependencia inicial. Pub/Sub solo si escala horizontal | APROBADO | `external-prompt-review/section-3-resolution` |
| D72 | ETag + 304 + gzip/brotli + Cache-Control con fingerprint para assets estáticos | APROBADO | `external-prompt-review/section-3-resolution` |
| D73 | No usar Insforge como BaaS. El backend hexagonal es nuestro | APROBADO | `external-prompt-review/section-4-resolution` |
| D74 | No introducir Kubernetes ni OpenShift prematuramente | APROBADO | `external-prompt-review/section-4-resolution` |
| D75 | Topología de despliegue ABIERTA (consistente con P7) | APROBADO (apertura) | `external-prompt-review/section-4-resolution` |
| D76 | PostgreSQL gestionado preferido sobre auto-instalado cuando se decida cloud | APROBADO | `external-prompt-review/section-4-resolution` |
| D77 | Contenedores Docker desde día uno + Docker Compose para dev local | APROBADO | `external-prompt-review/section-4-resolution` |
| D78 | Semver por módulo Y plataforma base (`modulo/vX.Y.Z-rc.n` / `modulo/vX.Y.Z`) | APROBADO | `external-prompt-review/section-5-resolution` |
| D79 | Trunk-based development + Conventional Commits | APROBADO | `external-prompt-review/section-5-resolution` |
| D80 | Catálogo de versiones compatibles entre módulos y plataforma | APROBADO | `external-prompt-review/section-5-resolution` |
| D81 | Despliegue coexistente estable+RC en UAT DIFERIDO hasta cadencia/equipo lo justifiquen | APROBADO (diferimiento) | `external-prompt-review/section-5-resolution` |
| D82 | Migraciones backward-compatibles con estrategia Expand and Contract | APROBADO | `external-prompt-review/section-5-resolution` |

## Decisiones aún no tomadas (ABIERTO)

No se han decidido y **no se inventan** en este documento. Las que tienen salvaguarda ya aprobada se marcan con la decisión de origen:

- Tecnología concreta de caché (el "qué" del adaptador de caché) — **salvaguardada por D70/D71**: caché selectiva justificada por medición; Redis queda como opción detrás del puerto, no como dependencia inicial.
- Topología de despliegue (on-premise, nube corporativa, OCP u otro) — **salvaguardada por D74/D75/D76**: Kubernetes/OpenShift NO se introduce prematuramente; topología ABIERTA; PostgreSQL gestionado preferido cuando se decida.
- Stack exacto de implementación — **CERRADO por D66/D67/D68** (Python 3.12+ / FastAPI 0.119+ / Pydantic v2 / SQLAlchemy 2.0.x / Alembic 1.13+ / asyncpg 0.30+ / HTMX 2.0.4 / Jinja2 3.1+ / Alpine.js 3.15+). Se cierra la pregunta P8 del blueprint.
- Descomposición en monolito modular vs microservicios — **parcialmente cerrada por D68**: monolito modular confirmado; la posibilidad de extraer módulos a microservicios queda abierta si la escala lo exige.
- Periodos definitivos de retención por cumplimiento normativo o políticas de IT.
- Integraciones corporativas concretas (correo, identidad, monitorización, etc.).
- Forma exacta del dashboard de operaciones de notificación (UX, filtros, acceso a contenido sensible, umbrales, acciones operativas).
- Diseño detallado del ciclo UAT (workflow, visibilidad, entorno, aprobaciones, promoción).
- Navegación dual UAT + producción simultánea cuando aplique — **salvaguardada por D81**: coexistente estable+RC DIFERIDO.
- Estrategia de migración de datos desde los `.accdb` a PostgreSQL.
- Catálogo definitivo de funciones de admin de aplicación por módulo (más allá del ejemplo de Gestion_Riesgos).
- Catálogo definitivo de health-checks (métricas, umbrales, severidades).

## Reglas del documento

- Este fichero no sustituye a `08-decisiones-y-preguntas-abiertas.md`: lo complementa con el detalle aprobado por el usuario.
- Toda mención a SiteMinder, OCP, Redis, S3, frameworks o colas corporativas se trata como **opción futura** o **abierto**, nunca como decisión final. **Excepción**: cuando una opción se registra como decisión APROBADO con salvaguarda (por ejemplo, Redis como opción detrás de puerto en D71), se admite la mención explícita en la sección correspondiente y en el mapa rápido, pero se mantiene la salvedad de que **no es dependencia inicial**.
- El legacy no es paridad de UX ni de implementación: es suelo mínimo de capacidad de negocio.
- Los IDs de Engram (`topic_key`) son el anclaje autoritativo; cualquier cambio futuro debe actualizar simultáneamente este documento y la observación correspondiente.
- APAP y APAP_WEB no aparecen ni se mencionan en este fichero.
- Documentación en castellano de España, registro técnico-profesional.

## Checklist

- [ ] Cada nueva decisión aprobada por el usuario añade una fila al mapa rápido (D66, D67, …) y actualiza su sección temática.
- [ ] Ningún elemento se mueve de ABIERTO/FUTURO a APROBADO sin confirmación explícita.
- [ ] Los baselines PROVISIONALes se marcan con su condición de revisión.
- [ ] Las referencias a SiteMinder, OCP, Redis, S3, frameworks o colas corporativas se mantienen como FUTURO/ABIERTO.

## Siguiente paso

Documento cruzado con `08-decisiones-y-preguntas-abiertas.md` tras la revisión del prompt externo de arquitectura (D66–D82). El blueprint queda consistente: stack cerrado por D66–D68, topología ABIERTA con salvaguarda D74–D76, versionado y releases consolidados por D78–D82, polling HTMX para contadores por D69, caché selectiva por D70–D72, infraestructura y despliegue por D73–D77.

Próximos pasos operativos:

1. Resolver P1–P5 con el usuario antes de iniciar el Lote 2 (HPS_Solicitudes repo, backends Condor/Brass, autorización Dysflow, orden de prioridad Lote 3–5, baseline operativo).
2. Continuar discovery por aplicación (Lotes 3 a 9) con la matriz legacy → web aprobada.
3. Planificar la fase SDD (proposal → spec → design → tasks → apply → verify → archive) cuando el discovery esté lo bastante maduro y el stack esté validado en un primer esqueleto ejecutable.
