# 08 · Decisiones y preguntas abiertas

## Propósito

Registro vivo de las decisiones tomadas durante el descubrimiento y de las preguntas que permanecen abiertas, en particular las que bloquean la entrada a un lote o a una fase SDD posterior. No sustituye a la conversación con el usuario; la refleja y la prepara.

## Decisiones

Las decisiones D1 a D91 viven en el [registro de decisiones](decisions/README.md), una por fichero, con `estado`, `fecha`, `decide`, `motivo` y `origen`. Esta página ya no repite su texto: conserva las disposiciones por tema y las preguntas abiertas.

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
