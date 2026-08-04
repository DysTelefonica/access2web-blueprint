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

Las decisiones D5–D35 están consolidadas en `09-arquitectura-objetivo-y-principios.md` con su detalle, `topic_key` de Engram y separación entre `APROBADO`, `PROVISIONAL`, `FUTURO` y `ABIERTO`. La tabla resumen se mantiene aquí como índice operativo.

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

Reglas operativas de este subregistro:

- Una decisión `APROBADO` cambia solo si el usuario lo revierte.
- Un `PROVISIONAL` se revisa cuando IT o cumplimiento entreguen los criterios definitivos.
- Un `FUTURO` permanece como opción no seleccionada hasta que se apruebe explícitamente.
- Un `ABIERTO` no se rellena aquí; vive en la sección "Preguntas abiertas" hasta obtener respuesta.

## Preguntas abiertas (cinco de mayor valor, tomadas de `exploration.md`)

| # | Pregunta | Bloquea | Lote relacionado |
|---|---|---|---|
| P1 | ¿Cuál es el repositorio/binario principal de **HPS_Solicitudes** y debe documentarse como aplicación independiente de HPS? | Sí | Lote 8 |
| P2 | ¿Cuál es la ruta y el nombre del backend vigente de **Condor** y **Brass**, y qué entorno debe considerarse baseline? | Sí | Lotes 6 y 7 |
| P3 | ¿Se autoriza configurar solo lectura los targets Dysflow de **Condor**, **Brass** y **Expedientes**, y diagnosticar el fallo de inventario de **No Conformidades**? | Sí | Lotes 5, 6, 7 y 2 |
| P4 | ¿Qué lote debe priorizarse después de aprobar Lanzadera y Expedientes: **Gestion_Riesgos**, **HPS** o **No Conformidades**? | Parcial | Lote 3, 4 o 5 |
| P5 | ¿Qué catálogo de `TbAplicaciones` y qué versión de backends debe considerarse el **baseline operativo actual**? | Sí | Lote 1 y transversal |

## Preguntas abiertas derivadas de las decisiones de arquitectura

Estas preguntas reflejan los puntos que `09-arquitectura-objetivo-y-principios.md` marca como `ABIERTO` o `FUTURO`. No bloquean el discovery legacy; bloquean fases SDD posteriores (selección de stack, despliegue, compliance). Se registran aquí para no perderlas.

| # | Pregunta | Bloquea | Notas |
|---|---|---|---|
| P6 | ¿Qué tecnología concreta de caché se usará detrás del puerto de caché? | Selección de stack | Redis es una opción, no una selección. |
| P7 | ¿Cuál es la topología de despliegue objetivo (on-premise, nube corporativa, OCP u otro)? | Selección de stack / despliegue | El adaptador de scheduler debe sobrevivir al cambio. |
| P8 | ¿Cuál es el stack exacto de implementación (framework, librerías, runtime)? | Diseño de módulos | No se prefija en esta fase. |
| P9 | ¿Cómo se descompone la plataforma (monolito modular vs microservicios vs mixto)? | Diseño y boundaries | Depende de capacidades y equipos. |
| P10 | ¿Cuáles son los periodos definitivos de retención por cumplimiento normativo o política de IT corporativa? | Política de auditoría | Sustituye al baseline provisional 90 días hot + 1 año total. |
| P11 | ¿Qué integraciones corporativas se confirman (correo, identidad, monitorización)? | Adaptadores driven | SiteMinder/OCP/JWT, proveedor de email corporativo, etc. |
| P12 | ¿Los hashes de contraseña heredados se migran tal cual o se exige reset? | Adaptador de autenticación | Depende de auditoría del algoritmo de hash. |
| P13 | ¿Qué funciones exactas tendrá el "responsable de aplicación" más allá de la administración delegada? | Modelo de roles | Pendiente de descubrimiento por módulo. |
| P14 | ¿Cuál es el catálogo definitivo de health-checks (métricas, umbrales, severidades)? | Operación | Se construye tras Lotes 1–9. |
| P15 | ¿Cuál es la estrategia de migración de datos desde `.accdb` a PostgreSQL? | Migración de datos | No se aborda en esta fase de discovery. |

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

Resolver P1–P5 con el usuario antes de iniciar el Lote 1.
