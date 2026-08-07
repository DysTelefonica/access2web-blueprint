[← Back to HPS_Solicitudes README](README.md) · [← Codebase Guide](../../../CODEBASE-GUIDE.md) · [← DOCS](../../../DOCS.md)

# Épica — HPS_Solicitudes (migración a web)

> **Estado:** DRAFT v0.1 (2026-08-06) — pendiente revisión final al cerrar el ciclo de las 8 épicas.
> **Lote:** 7 (junto al resto de las 7 apps restantes).
> **App legacy:** `HPS_SOLICITUDES` · frontend `Solicitudes_HPS.accdb` (22 MB) + backend `Solicitudes_HPS_datos.accdb` (13 MB, autoritativo en `C:\00repos\datos\`).
> **Sentence that organizes**: **HPS_Solicitudes es la app de Solicitudes HPS con workflow regulatorio de 20 fechas y traspasos al ONS (Organismo Notificador de Seguridad): cada solicitud tiene 28 columnas con 8 PII, fechas regulatorias, responsable vinculado por email (D99 — anti-patrón de integridad), y 245 solicitudes activas en producción (vs 1 de Condor en staging). La app independiente con ID 22 vs HPS ID 17 (D83).**

> **Scope del scope**: "Este repo es research + planning de la migración. Cada app tendrá su propio repo + docs cuando se construya."

---

## Quick Navigation

| Section                                                       | What you'll find                                              |
|---------------------------------------------------------------|---------------------------------------------------------------|
| [Metadatos](#metadatos)                                       | Scope S, D98 PII + D99 FK por email, dependencias.             |
| [1. Scope](#1-scope)                                          | 19 features F1-F19 + 3 catálogos config seed-only.             |
| [2. Estado del descubrimiento](#2-estado-del-descubrimiento)    | Inventario 11 tablas backend + audit usage PENDIENTE.         |
| [3. Hallazgos críticos](#3-hallazgos-críticos)                | D98 PII 245 filas, D99 FK email, D83 app independiente.        |
| [4. Decisiones aplicadas](#4-decisiones-aplicadas)              | UX, hexagonal ports, seguridad, datos.                       |
| [5. Criterios de aceptación](#5-criterios-de-aceptación)      | Funcionalidad, Seguridad, Performance, Operacional.            |
| [6. Pendientes operacionales](#6-pendientes-operacionales)    | Antes, durante, go-live (incluye audit PENDIENTE).            |
| [7. Tickets derivables](#7-tickets-derivables-preview)         | 22 tickets TK-HPS-SOL-1..22.                                 |
| [Anexo · Decisiones referenciadas](#anexo--decisiones-referenciadas) | D5-D101.                                             |
| [Anexo · Tabla de fuentes](#anexo--tabla-de-fuentes)         | Walkthrough + docs + source + engram.                          |

---

## Metadatos

| Campo | Valor |
|---|---|
| **Aplicación legacy** | `HPS_SOLICITUDES` · frontend `Solicitudes_HPS.accdb` (22 MB) + backend `Solicitudes_HPS_datos.accdb` (13 MB, autoritativo en `C:\00repos\datos\`) |
| **Tipo de migración** | Legacy Access/VBA → web hexagonal (FastAPI + HTMX) |
| **Scope size** | **S** (11 tablas, 19 features de negocio, 26 clases, 14 módulos, 1 anexo Excel, 1 integración externa ONS) — **la app más simple en superficie de datos** pero **en uso activo** (245 solicitudes reales vs 1 de Condor en staging) |
| **Dependencias cross-app** | Lanzadera (identidad/permisos, vía `getdbLanzadera()`); HPS gestión de usuarios (`IDUsuarioHPS` FK conceptual + flag `RegistroEnHPS="Sí"`, **app independiente ID 22 vs HPS ID 17 — D83**); Lanzadera `TbExpedientes` (`IDExpediente` FK conceptual); ONS — sistema externo (D101) |
| **Riesgo dominante** | Seguridad **D98** — **245 filas con datos personales completos** en `TbSolicitudes` (DNI, nombres, fechas, correo, teléfono); **D99** — FK por **email** en vez de por ID (`TbResponsables.Correo → TbSolicitudes.emailResponsable`) — anti-patrón de integridad referencial |
| **Stack target** | Backend Python 3.12+ / FastAPI 0.119+ / SQLAlchemy 2.0.x / Alembic 1.13+ / asyncpg 0.30+ (D66) · Frontend HTMX 2.0.4 + Jinja2 3.1+ / Alpine.js 3.15+ (D67) |
| **Estrategia de migración de BD** | Expand and Contract backward-compatible (D82) · PostgreSQL compartido con esquema por módulo (D14) |
| **Forma destino** | Hexagonal global (D8) · módulo dentro del monolito modular (D68) · puerto de persistencia PostgreSQL + object storage S3-compatible (D16) + secret manager (D9-D10) |
| **Auditoría de uso previa** | **PENDIENTE** (mismo gap que tuvo HPS antes de obs #24097). El schema usage audit de las 11 tablas backend **no se ha ejecutado todavía**. Gate previo a fase SDD; ver § 1.4, § 2.1.1 y § 6 (TK-HPS-SOL-22). |

---

## 1. Scope

### 1.1 En scope — 19 features de negocio con paridad funcional

| # | Feature | Respaldo en docs |
|---|---|---|
| F1 | Arranque e identidad (`EVE` + `getUsuario`, IDAplicacion `22`) | `capabilities.md` · `forms.md` · `data-model.md` |
| F2 | Configuración de backend (`TbConfiguracionBackends`) → config del módulo hexagonal | `capabilities.md` · `integrations-automation.md` |
| F3 | CRUD Solicitud (`TbSolicitudes`, 245 filas) — alta, edición, baja, renovación, cambio de tipo | `capabilities.md` · `forms.md` · `data-model.md` § TbSolicitudes (28 columnas) |
| F4 | Workflow regulatorio de HPS con 20 fechas (`TbSolicitudesFechas`, 245 filas 1:1 con `TbSolicitudes`) | `data-model.md` § TbSolicitudesFechas · `forms.md` |
| F5 | Traspasos a ONS — Organismo Notificador de Seguridad (`URLAdjuntoEnvioONS`, 4 forms) | `capabilities.md` · `forms.md` · `integrations-automation.md` § D101 |
| F6 | Adjuntos Excel del solicitante (`Form_FormAdjuntarExcelSolicitante`, `URLAdjunto`) | `capabilities.md` · `integrations-automation.md` § Adjuntos Excel |
| F7 | Plantillas HTML / Vistas web embebidas (`Form_FormPlantillasHTML`, `Form_FormWeb`) — sistema de generación con datos | `capabilities.md` · `integrations-automation.md` § Plantillas HTML |
| F8 | Justificaciones (`TbJustificaciones`, 7 filas) — alta y búsqueda | `capabilities.md` · `forms.md` |
| F9 | Responsables (`TbResponsables`, 27 filas) — gestión CRUD vinculados a solicitud ⚠️ por **email** (D99) | `capabilities.md` · `forms.md` · `data-model.md` § Relaciones físicas |
| F10 | Motivos HPS (`MotivoHPS.cls`, `Form_FormMotivoHPS`, `Form_FormMotivosHPS`) — catálogo | `capabilities.md` · `forms.md` |
| F11 | Entidades / Expedientes búsqueda (`Form_FormEntidades`, `Form_FormExpedientesBusqueda`, `Form_FormExpedienteDetalle`) | `capabilities.md` · `forms.md` |
| F12 | Tareas pendientes del tramitador (`Form_FormTareasTramitadorPendientes`) — worklist | `capabilities.md` · `forms.md` |
| F13 | Configuración runtime (`TbConfiguracion`, 15 columnas, 7 flags TempVars activos) | `capabilities.md` · `data-model.md` § TbConfiguracion · `integrations-automation.md` |
| F14 | Correos automáticos (`TbCorreosEnviados`, 21 columnas, 9 filas) — sistema completo con destinatarios múltiples, plantillas, versionado, intentos | `capabilities.md` · `forms.md` · `data-model.md` § TbCorreosEnviados |
| F15 | Registro automático en HPS (`IDUsuarioHPS` + flag `RegistroEnHPS="Sí"`) — cross-app | `integrations-automation.md` § Integración con HPS · `data-model.md` |
| F16 | Búsqueda general de solicitudes (`Form_FormSolicitudesGestion`, `Form_FormAltaDatosSolicitante`) | `capabilities.md` · `forms.md` |
| F17 | Búsqueda de usuarios HPS (`Form_FormUsuariosHPSBusqueda`) — para asignación | `capabilities.md` |
| F18 | Automatización (`AutomatizacionRepositorio.bas`, `Automiatizacion.bas` typo) — jobs programados (traspasos, renovaciones) | `capabilities.md` · `integrations-automation.md` § Automatización |
| F19 | Informes (`InformesOperaciones.bas`) — generación | `capabilities.md` |

### 1.2 3 catálogos config seed-only (datos de runtime, sin CRUD de usuario)

| Catálogo | Volumen | Notas |
|---|---|---|
| `motivos_hps` (era `MotivoHPS.cls`) | TBD | Catálogo de motivos alta/baja HPS |
| `hps_grado` (era `TbHPSGrado`, clave compuesta `(TipoHPS, Grado)`) | **13 filas** | Grados HPS con clave compuesta |
| `responsables` (era `TbResponsables`) | 27 filas | Responsables con FK por email (D99) |

### 1.3 Volumen real autoritativo (backend `Solicitudes_HPS_datos.accdb`)

| Tabla | Filas |
|---|---|
| `TbSolicitudes` | 245 (solicitudes activas, 28 columnas, **8 columnas con datos personales** D98) |
| `TbSolicitudesFechas` | 245 (1:1 con solicitudes, 21 columnas = 20 fechas regulatorias) |
| `TbResponsables` | 27 (responsables con email único) |
| `TbJustificaciones` | 7 (justificaciones) |
| `TbHPSGrado` | 13 (catálogo grados HPS) |
| `TbLogsGeneral` | 2058 (log general activo) |
| `TbCorreosEnviados` | 9 (registro de correos) |
| `TbLogs` | 0 (vacía — presumible desuso, D100) |
| `TbConfiguracion` | presumido 1 fila (no contado) |
| `TbUltimoCambio` | presumido 245 (1:1 con solicitudes, no contado) |
| `Copia de TbExpedientes` | legacy copy (volumen no contado) |
| **Total principales** | **≈2885** (sumando logs) |

**Read-heavy intermedio**: `getdb()` con **89 callers** (intermedio entre Lanzadera/Expedientes y Gestion_Riesgos/NoConformidades). La mayoría son operaciones CRUD sobre `TbSolicitudes` + `TbSolicitudesFechas`.

**Tests VBA limitados**: 2 archivos (`Test.bas` genérico, `TestParametrosParser.bas` específico de `ParametrosParser.cls`). Disciplina TDD **menos madura** que HPS (9 archivos) o NoConformidades (7). Cobertura como referencia — `TestParametrosParser.bas` es el único test específico, utilitario.

**Vinculaciones externas**: la identidad se resuelve vía **adaptador unificado** (D9-D10); las FKs conceptuales cross-app (`IDExpediente`, `IDUsuarioHPS`, `IDEmpresaUsuario`, `IDEmpresaTramitadora`) se mantienen como referencia conceptual mediate adaptadores (D86/D87). **HPS_Solicitudes es app independiente** (ID 22 vs ID 17 de HPS) — no comparten tablas (D83). La integración con **ONS** (D101) es un adaptador de salida (D16, object storage para adjuntos).

### 1.4 Fuera de scope (REPLACE)

> Las **tablas puras de log** NO migran como tablas PostgreSQL. Se reemplazan por stack de observabilidad web-native (Sentry / OpenTelemetry / structured logs a Loki o CloudWatch). Sus llamadas VBA se traducen a eventos web.

- `TbLogsGeneral` (2058 filas en staging) — log general activo con campos presumidos de timestamp + actor + contexto. **REPLACE web-native** (D27). Reflejo del veredicto de Condor (obs #24085 + decisión B del lote 5).
- `TbLogs` (0 filas, presumible desuso, D100) — descartar; nunca se reactiva en PostgreSQL.
- `TbCorreosEnviados` (9 filas) — **evaluar caso a caso**: contiene metadatos de envío con FK a `TbSolicitudes` (ToDo: confirmar si es tabla de negocio o log). Como es bajo volumen, **se preserva como tabla de auditoría de envíos** con retención (D27-D29); NO REPLACE.

**Excluido del REPLACE**: el resto de las 11 tablas backend migran como tablas PostgreSQL con sus claves y datos preservados.

### 1.5 Fuera de scope (no documentado)

- Volúmenes reales de producción (no staging): `Solicitudes_HPS_datos.accdb` en `C:\00repos\datos\` no fue inspeccionado en esta pasada más allá del inventario Dysflow base.
- Conteos reales de `TbConfiguracion` (presumido 1 fila) y `TbUltimoCambio` (presumido 245 filas 1:1 con solicitudes) — no auditados en staging.
- Schema de `TbCorreosEnviados` — los 21 campos están documentados en `data-model.md`, pero el contenido (qué plantillas, qué destinatarios, qué periodicidad) no fue inspeccionado.
- Schemas no inspeccionados en la pasada inicial (`TbLogs`, `TbLogsGeneral`, `Copia de TbExpedientes`) — fallaron por colección vacía; la segunda pasada documentó las 7 tablas principales pero estas 4 quedaron con shape solo en inventario.
- Inventario completo de queries exportadas y macros embebidas — codegraph no captura macros embebidas ni QueryDefs no exportados.
- Plantillas de correo en `Correo.cls` — pueden contener datos sensibles (D98); revisión pendiente antes de portar.
- Contenido del Excel subido por el solicitante (`Form_FormAdjuntarExcelSolicitante`) — qué columnas se esperan, qué validaciones; no inspeccionado en detalle.
- `Automiatizacion.bas` (typo) — qué automatizaciones exactamente (traspasos a ONS, renovaciones, recordatorios); no inspeccionado en detalle.
- Contenido de `Form_FormPlantillasHTML.cls` y `Form_FormWeb.cls` — qué plantillas concretas, qué versión, qué merge.

### 1.6 Fuera de scope (no-migrate — pendiente de audit)

> **Schema usage audit PENDIENTE**. Las 11 tablas backend **no han sido clasificadas** como ACTIVE / ZOMBIE / UNCLEAR. El veredicto equivalente al de HPS (obs #24097, 12 ACTIVE / 1 UNCLEAR / 9 ZOMBIE) **no existe todavía** para HPS_Solicitudes. Gate previo a fase SDD; ver § 6 (TK-HPS-SOL-22).

**Limitación declarada**: el inventario Dysflow sí ejecutó (`TbSolicitudes` con 245 filas, `TbLogsGeneral` con 2058 filas, etc.), pero la extracción de callers via `codegraph-vba` sobre `C:\00repos\codigo\HPS_SOLICITUDES` no se corrió para clasificar las tablas por uso real en código. La categoría "no-migrate — ZOMBIE" no tiene veredicto formal todavía; se sospecha que `Copia de TbExpedientes` (legacy copy) y `TbLogs` (0 filas, presumible desuso, D100) son candidatas a ZOMBIE, pero la confirmación requiere ejecutar el audit.

**Candidatas probables** (sin veredicto formal, pendiente de audit):
- `Copia de TbExpedientes` — patrón legacy copy pre-mass-change; retención indefinida, sin endpoints, sin UI.
- `TbLogs` — 0 filas, presumible desuso (D100); el log activo es `TbLogsGeneral`.

**Catálogo candidato** (no auditado):
- `TbHPSGrado` — 13 filas, 2 callers estimados vía formularios de catálogo. Probablemente ACTIVE pero requiere confirmación.

---

## 2. Estado del descubrimiento

### 2.1 Inventario

| Categoría | Resultado |
|---|---|
| **Tablas totales (backend)** | **11** en `Solicitudes_HPS_datos.accdb` (autoritativo en `C:\00repos\datos\`) |
| **Filas totales (backend, inventario Dysflow)** | **≈2885** (245 solicitudes + 245 fechas + 27 responsables + 7 justificaciones + 13 grados + 9 correos + 2058 logs + 1 config presumido + 245 ultimoCambio presumido + copia legacy no contada) |
| **Schemas documentados** | **7/11** tablas en `data-model.md` (segunda pasada): `TbSolicitudes` (28 col), `TbResponsables` (4 col), `TbJustificaciones` (4 col), `TbHPSGrado` (2 col), `TbSolicitudesFechas` (21 col), `TbConfiguracion` (15 col), `TbUltimoCambio` (4 col), `TbCorreosEnviados` (21 col). **4/11 con shape solo inventario**: `TbLogs` (vacía, D100), `TbLogsGeneral` (alta volumetría), `Copia de TbExpedientes` (legacy copy). |
| **Uso (audit)** | **PENDIENTE**. No se ha ejecutado `codegraph_explore` para clasificar las 11 tablas backend. Gate previo a fase SDD (TK-HPS-SOL-22). |
| **Subcategorías** | 6 dominio (solicitudes, fechas, responsables, justificaciones, motivos, config) · 3 audit/log (TbLogs, TbLogsGeneral, TbUltimoCambio) · 1 catálogo compuesto (hps_grado) · 1 legacy copy (Copia de TbExpedientes) |
| **FKs físicas (backend)** | **3** entre user tables, todas con problemas (ver § 2.2) |
| **FKs conceptuales (sin constraint)** | **5** — `IDExpediente`, `IDUsuarioHPS`, `IDEmpresaUsuario`, `IDEmpresaTramitadora`, `idjustificacion` (en `TbSolicitudes`) |
| **Clases inventariadas** | 26 (14 domain, 8 operaciones/servicios, 2 compartidas con Lanzadera, 1 util, 1 errata) |
| **Módulos** | 14 (5 bootstrap/factory/DAO, 3 repositorios, 2 tests, 4 miscelánea — incluye `Automiatizacion.bas` con typo) |
| **Forms** | ~30+ archivos `Form_*.cls` cada uno con su `.form.txt` compañero |
| **Tests VBA** | 2 archivos (`Test.bas`, `TestParametrosParser.bas`) — cobertura básica (D87 referencia, no enriquecida) |
| **Integraciones externas** | ONS (sistema externo, D101) — único sistema externo fuera de Lanzadera/HPS/Expedientes |

### 2.2 Schema usage audit — PENDIENTE (gap crítico)

> **No se ha ejecutado schema usage audit sobre las 11 tablas backend**. Mismo gap que tuvo HPS antes de obs #24097. El veredicto (ACTIVE / UNCLEAR / ZOMBIE) es **desconocido**; los candidatos probables están en § 1.6 pero sin clasificación formal.

**Pendiente ejecutar** (análogo a obs #24097 de HPS):
1. `codegraph_explore` sobre `C:\00repos\codigo\HPS_SOLICITUDES\00_main` (o la raíz directa del repo según `.codegraph-vba`/`staging/.codegraph-vba` confirmado en obs #23876).
2. Para cada una de las 11 tablas: extraer callers en `.bas`/`.cls`, blast radius de símbolos clave (`EVE`, `getdb`, `Solicitud`, `Responsable`, `Justificacion`, `LogGeneral`, `Correo`).
3. Clasificar ACTIVE / UNCLEAR / ZOMBIE.
4. Documentar el veredicto como observación engram con topic_key `hps-solicitudes/usage-audit-<fecha>`.
5. Cerrar el gap con un commit que aplique el veredicto al epic (similar a commit `8d5311c` para HPS).

**Limitaciones que se esperan** (mismas que obs #24097 documentó para HPS):
- Macros embebidos en `Solicitudes_HPS.accdb` NO exportados → uso real puede ser mayor al lower bound.
- `codegraph_explore` limita el blast radius visible a 5 entries por símbolo.
- Forms sin `RecordSource`/`RowSource` exportados: app carga vía DAO + clases.

---

## 3. Hallazgos críticos

| # | ID | Título | Severidad | Forms afectados | Detalle |
|---|---|---|---|---|---|
| H1 | D98 | **245 filas con PII completo en `TbSolicitudes`** | CRITICAL | `TbSolicitudes` (8 columnas) | DNI, nombres, fechas, correo, teléfono en 245 solicitudes activas en producción. **CRÍTICO**: el frontend NO debe existir como runtime en web — el `Solicitudes_HPS.accdb` se descarta. Los datos de PII migran al backend autoritativo con encryption at-rest y access controls. |
| H2 | D99 | **FK por email en vez de por ID** | high | `TbResponsables.Correo → TbSolicitudes.emailResponsable` | Anti-patrón de integridad referencial. Email puede cambiar, no ser único. Migrar a FK por `IDResponsable` con tabla `TbResponsables.ID` como PK. |
| H3 | D83 | **`HPS_Solicitudes` es app independiente (ID 22 vs 17)** | low | Cross-app con HPS | FK conceptual `IDUsuarioHPS` + flag `RegistroEnHPS="Sí"`. No comparten tablas. Mantener como referencia conceptual mediate adaptador. |
| H4 | D100 | **`TbLogs` con 0 filas (presumible ZOMBIE)** | low | `TbLogs` | Vacía en staging. Probable desuso. Pendiente de audit para clasificación formal. |
| H5 | D101 | **Integración externa ONS** | medium | `Form_FormEnvioONS*` (4 forms) | Adaptador de salida (D16, object storage). Volumen bajo (no especificado en staging). |
| H6 | **3 catálogos con FK por email** | high | `TbResponsables` (D99) | Email como FK. En la nueva plataforma: usar `ID` (autonumérico) con email como unique constraint secundario. |
| H7 | **`Automiatizacion.bas` typo** | low | — | Job programado de traspasos y renovaciones. Corregir nombre al migrar. |
| H8 | **`Copia de TbExpedientes`** | low | — | Legacy copy pre-mass-change. Patrón de HPS (4 copias). Decidir drop vs archive. |

---

## 4. Decisiones aplicadas

### 4.1 UX: Preservar / Mejorar / Nuevo paradigma

| Pantalla | Decisión | Justificación |
|---|---|---|
| Splash + Arranque (F1) | **Preservar** flujo vía EVE | Adaptador unificado (D9-D10) |
| `Form_FormSolicitud*` (CRUD) | **Preservar** flujo + **Mejorar**: server-side pagination | 245 solicitudes activas con PII |
| `Form_FormSolicitudesFechas` (workflow 20 fechas) | **Preservar** flujo | 20 fechas regulatorias, crítico para HPS |
| Traspasos a ONS (F5) | **Preservar** flujo | D101 — integración externa oficial |
| Adjuntos Excel (F6) | **Preservar** upload | Compatibilidad con formato de usuario |
| Plantillas HTML / Vistas web (F7) | **Mejorar**: server-side rendering con Mistica | Migración a Jinja2 + tokens |
| Worklist (F12) | **Mejorar**: server-side polling + cache | Performance con 245+ items |
| Automatización (F18) | **Preservar** jobs programados (CRON) | Crítico para traspasos ONS + renovaciones |

### 4.2 Arquitectura: hexagonal ports

| Concepto legacy | Traducción web |
|---|---|
| `getdb()` (DAO.Database) — 89 callers | Puerto de persistencia PostgreSQL (D14) + HTTP client |
| `Constructor.getXxx()` (factory con cache lazy) | Inyección de dependencias + repository pattern |
| `m_ObjEntorno` (singleton global) | Service registry / dependency injection |
| `m_ObjUsuarioConectado` (singleton in scope) | Request-scoped services / Context API |
| `TbConfiguracionBackends` (config del frontend) | Variables de entorno del runner (D9-D10) |
| `EVE` (autenticación legacy) | Adaptador unificado (D9-D10) |
| `TbConfiguracion` (15 cols, 7 flags TempVars) | Config del módulo hexagonal — variables de entorno |
| `URLAdjuntoEnvioONS` (integración ONS) | Adapter de salida (D101) con SDK ONS (REST o SFTP) |
| `URLAdjunto` (anexos Excel) | SDK S3-compatible (D16) con signed URLs |
| `Form_FormPlantillasHTML` + `Form_FormWeb` | Server-side rendering con Jinja2 + tokens Mistica |
| `AutomatizacionRepositorio.bas` + `Automiatizacion.bas` typo | Jobs programados (CRON) con `celery` o similar |
| `TbSolicitudesFechas` (20 fechas regulatorias) | Mantener como 20 columnas en PostgreSQL — crítico para HPS |
| `tbHistorialRechazos` (patrón cross-cutting) | (no aplica directamente, pero el patrón sí) |

### 4.3 Decisiones de seguridad

- **PII en `TbSolicitudes` (D98)**: NO migra como caché local. El `Solicitudes_HPS.accdb` frontend se descarta. Los datos de PII migran al backend autoritativo con encryption at-rest y access controls estrictos. El frontend web NO tiene copia local de PII.
- **FK por email (D99)**: MIGRACIÓN — `TbResponsables.Correo` deja de ser FK. En la nueva plataforma, `IDResponsable` (autonumérico) es PK; email es unique constraint secundario. El campo `emailResponsable` en `TbSolicitudes` se reemplaza por FK `IDResponsable`.
- **Cross-app con Lanzadera** (identidad): adaptador unificado (D9-D10).
- **Cross-app con HPS** (FK conceptual `IDUsuarioHPS`): adaptador.
- **Cross-app con Expedientes** (FK conceptual `IDExpediente`): adaptador.
- **ONS (D101)**: sistema externo — bearer token, retry, audit completo.

### 4.4 Decisiones de datos

- **11 tablas backend** migran a PostgreSQL con claves y datos preservados.
- **FKs conceptuales** (5): formalizar en PostgreSQL con FKs reales (especialmente `IDResponsable` en lugar de `emailResponsable`).
- **Booleanos `Text(2)` (D102 cross-cutting)**: migrar a `BOOLEAN` con regla explícita.
- **Tablas ZOMBIE** (D100): decidir disposición caso por caso con negocio.
- **20 fechas regulatorias** (`TbSolicitudesFechas`): mantener como 20 columnas en PostgreSQL — crítico para el workflow HPS.

---

## 5. Criterios de aceptación

### 5.1 Funcionalidad

- [ ] **CA-F1**: Las 19 features F1-F19 tienen paridad funcional con la versión Access.
- [ ] **CA-F2**: El workflow regulatorio de 20 fechas se preserva íntegro.
- [ ] **CA-F3**: La integración con ONS (D101) preserva el contrato y la trazabilidad.
- [ ] **CA-F4**: La integración con HPS (registro automático `RegistroEnHPS="Sí"`) preserva el flag.
- [ ] **CA-F5**: Los adjuntos Excel del solicitante preservan formato + validaciones.
- [ ] **CA-F6**: Las plantillas HTML / Vistas web migran a server-side rendering con Mistica.
- [ ] **CA-F7**: La automatización (F18) — traspasos a ONS, renovaciones — preserva los jobs programados.
- [ ] **CA-F8**: El worklist del tramitador (F12) escala con server-side pagination.

### 5.2 Seguridad

- [ ] **CA-S1**: **`TbSolicitudes` (245 filas PII) eliminado del frontend** — los datos migran al backend con encryption at-rest.
- [ ] **CA-S2**: **FK `emailResponsable` (D99) eliminada** — migrada a `IDResponsable` con email como unique constraint secundario.
- [ ] **CA-S3**: PII no expuesta via API sin scope explícito.
- [ ] **CA-S4**: Cross-app API con Lanzadera + HPS + Expedientes vía OAuth 2.0.
- [ ] **CA-S5**: Integración ONS con bearer token + retry + audit completo.
- [ ] **CA-S6**: `TbConfiguracion` migrado a secret manager (no en frontend).

### 5.3 Performance

- [ ] **CA-P1**: Listado de 245 solicitudes con paginación server-side carga < 1s.
- [ ] **CA-P2**: Worklist del tramitador carga < 2s.
- [ ] **CA-P3**: Traspasos a ONS con retry + backoff en < 30s.

### 5.4 Operacional

- [ ] **CA-O1**: 11 tablas backend migradas con datos preservados (~2885 filas).
- [ ] **CA-O2**: Logs web-native (Sentry/OTel) — NO existen `TbLogsGeneral`/`TbLogs` en PostgreSQL.
- [ ] **CA-O3**: Tests E2E portados desde los 2 archivos `Test_*.bas`.
- [ ] **CA-O4**: Cobertura de tests > 70% en módulo hps-solicitudes.
- [ ] **CA-O5**: **Schema usage audit ejecutado** (TK-HPS-SOL-22) — veredicto formal ACTIVE/UNCLEAR/ZOMBIE para las 11 tablas.

---

## 6. Pendientes operacionales

### 6.1 Antes de empezar

- [ ] **PO-1**: **Ejecutar schema usage audit** (análogo a obs #24097 de HPS) — gate previo a fase SDD. Cerrar el gap de veredicto formal.
- [ ] **PO-2**: Confirmar el contrato de integración con **Lanzadera** (vía `getdbLanzadera()`).
- [ ] **PO-3**: Confirmar el contrato de integración con **HPS** (registro automático `RegistroEnHPS="Sí"`).
- [ ] **PO-4**: Confirmar el contrato de integración con **ONS** (D101) — bearer token, retry policy.
- [ ] **PO-5**: Confirmar el contrato de integración con **Expedientes** (FK conceptual `IDExpediente`).
- [ ] **PO-6**: Decidir disposición de `TbLogs` (0 filas, presumible ZOMBIE) tras audit.

### 6.2 Durante el desarrollo

- [ ] **PO-7**: Implementar servicio de hidratación de `TbSolicitudes` desde backend (NO dejar copia en frontend).
- [ ] **PO-8**: **Migrar FK `emailResponsable` → `IDResponsable`** (D99) — script Alembic con DROP CONSTRAINT + ADD CONSTRAINT nuevo.
- [ ] **PO-9**: Implementar server-side pagination en `Form_FormSolicitudesGestion` (F16).
- [ ] **PO-10**: Implementar jobs programados (CRON) para traspasos ONS + renovaciones.
- [ ] **PO-11**: Implementar `URLAdjuntoEnvioONS` adapter con bearer token + retry + audit.
- [ ] **PO-12**: Corregir typo `Automiatizacion.bas` → `Automatizacion.bas` al migrar.

### 6.3 En el go-live

- [ ] **PO-13**: Smoke test E2E: alta solicitud → 20 fechas regulatorias → traspaso ONS → alta correo.
- [ ] **PO-14**: Verificar que NO queden copias de PII en el frontend web (security audit).
- [ ] **PO-15**: Verificar que el FK por email (D99) está completamente migrado a `IDResponsable` — sin valores residuales.
- [ ] **PO-16**: Plan de deprecation de las ZOMBIE confirmadas con negocio.

---

## 7. Tickets derivables (preview)

### Seguridad (CRITICAL — D98)

- **TK-HPS-SOL-1**: [SECURITY] **CRITICAL**: Eliminar `TbSolicitudes` (245 filas PII) del frontend — los datos migran al backend con encryption at-rest.
- **TK-HPS-SOL-2**: [SECURITY] **CRITICAL**: Eliminar FK `emailResponsable` (D99) — migrar a `IDResponsable` con email como unique constraint secundario.

### Integridad referencial (D94, D99)

- **TK-HPS-SOL-3**: [DATA] Formalizar las 5 FKs conceptuales con constraints reales.
- **TK-HPS-SOL-4**: [DATA] Decidir disposición de las ZOMBIE (drop vs archive).

### Forma y plataforma

- **TK-HPS-SOL-5**: [REFACTOR] Migrar 26 clases a services HTTP manteniendo firmas.
- **TK-HPS-SOL-6**: [MIGRATION] Reemplazar 20 fechas regulatorias de `TbSolicitudesFechas` con columnas tipadas (DATE) en PostgreSQL.
- **TK-HPS-SOL-7**: [MIGRATION] Reemplazar `TbConfiguracion` (15 cols, 7 flags TempVars) con config del módulo hexagonal (variables de entorno).
- **TK-HPS-SOL-8**: [MIGRATION] Reemplazar `Form_FormPlantillasHTML` + `Form_FormWeb` con server-side rendering Mistica + Jinja2.
- **TK-HPS-SOL-9**: [MIGRATION] Reemplazar `URLAdjunto` con SDK S3-compatible (D16) con signed URLs.
- **TK-HPS-SOL-10**: [MIGRATION] Reemplazar `URLAdjuntoEnvioONS` con adapter de salida ONS (REST o SFTP) + bearer token + retry.
- **TK-HPS-SOL-11**: [MIGRATION] Reemplazar `AutomatizacionRepositorio.bas` + `Automiatizacion.bas` typo con jobs programados (CRON).
- **TK-HPS-SOL-12**: [MIGRATION] Reemplazar `Form_FormAdjuntarExcelSolicitante` con validación server-side del Excel.

### Integración

- **TK-HPS-SOL-13**: [INTEGRATION] API REST federada con Lanzadera (identidad) + HPS (registro automático) + Expedientes.
- **TK-HPS-SOL-14**: [INTEGRATION] Adapter ONS con bearer token + retry + audit completo.
- **TK-HPS-SOL-15**: [INTEGRATION] Mantener HPS_Solicitudes como app independiente (D83).

### Testing y calidad

- **TK-HPS-SOL-16**: [TESTING] Portar 2 Tests VBA a E2E pytest + httpx.
- **TK-HPS-SOL-17**: [TESTING] Cobertura de tests > 70% en módulo hps-solicitudes.
- **TK-HPS-SOL-18**: [TESTING] E2E test del workflow regulatorio completo (alta → 20 fechas → traspaso ONS).
- **TK-HPS-SOL-19**: [TESTING] Tests para la FK por email (D99) — verificar que la migración de `emailResponsable` → `IDResponsable` no rompe nada.

### Datos y operación

- **TK-HPS-SOL-20**: [DATA] Backfill de las 11 tablas backend a PostgreSQL.
- **TK-HPS-SOL-21**: [INFRA] Setup de jobs programados (CRON) para traspasos ONS + renovaciones.

### Audit (gap crítico)

- **TK-HPS-SOL-22**: [AUDIT] **CRITICAL GATE**: Ejecutar schema usage audit (análogo a obs #24097 de HPS) — gate previo a fase SDD. Cerrar el gap de veredicto formal ACTIVE/UNCLEAR/ZOMBIE para las 11 tablas.

---

## Anexo · Decisiones referenciadas

| Decisión | Aplicación a HPS_Solicitudes |
|---|---|
| D8 (hexagonal global) | Toda la migración |
| D9-D10 (secret manager + adaptador unificado) | Eliminar `TbConfiguracion` del frontend |
| D14 (esquema por módulo) | Módulo `hps_solicitudes` en PostgreSQL |
| D16 (object storage S3-compatible) | Anexos Excel + adjuntos ONS |
| D27 (logs estructurados canónicos) | Reemplazar `TbLogsGeneral` y `TbLogs` con Sentry/OTel |
| D44-D46 (autorización + capabilities) | Roles Calidad/Técnico/Tramitador |
| D66-D67 (stack) | Backend Python + HTMX |
| D68 (monolito modular) | HPS_Solicitudes como módulo |
| D82 (Expand and Contract) | Migración de 11 tablas + descarte de frontend |
| D83 (apps independientes) | HPS_Solicitudes (ID 22) es app separada de HPS (ID 17) |
| D86-D87 (cross-app + tests VBA) | FKs conceptuales + 2 Tests VBA como referencia |
| D98 (PII CRITICAL) | **245 filas PII en `TbSolicitudes` — eliminar del frontend** |
| D99 (FK por email) | **`TbResponsables.Correo → TbSolicitudes.emailResponsable` — anti-patrón; migrar a `IDResponsable`** |
| D100 (TbLogs 0 filas) | Descarte tras audit |
| D101 (integración ONS) | Adapter de salida con bearer token + retry + audit |

## Anexo · Tabla de fuentes

| Fuente | Aporta |
|---|---|
| engram obs #24081 | Schemas detalle de HPS_Solicitudes (7 tablas en segunda pasada) |
| [`data-model.md`](data-model.md) | 11 tablas backend con schemas + relaciones |
| [`capabilities.md`](capabilities.md) | Inventario de features (vista de alto nivel) |
| [`forms.md`](forms.md) | Navegación, call paths, formularios críticos |
| [`integrations-automation.md`](integrations-automation.md) | Cross-app, ONS, Adjuntos Excel, Plantillas HTML, Automatización |
| [`migration-matrix.md`](migration-matrix.md) | D83, D98-D101 disposiciones específicas de HPS_Solicitudes |
| [`security-rules.md`](security-rules.md) | D98 PII, autorización, Indicadores, Sincronización |
| [`README.md`](README.md) | Estado del lote, hallazgos críticos, checklist |
| `data/staging/hps-solicitudes/src/classes/*.cls` | 26 clases de dominio |
| `data/staging/hps-solicitudes/src/modules/*.bas` | 14 módulos (5 bootstrap, 3 repositorios, 2 tests) |
| `data/staging/hps-solicitudes/src/forms/*.form.txt` | ~30+ forms con RecordSource/RowSource |
| `C:\00repos\datos\Solicitudes_HPS_datos.accdb` | Backend autoritativo (volúmenes reales) |
| `C:\00repos\codigo\HPS_SOLICITUDES\staging` | Fuente READ-ONLY |
| [DOCS](../../../DOCS.md) | Technical reference raíz del blueprint |
| [CODEBASE-GUIDE](../../../CODEBASE-GUIDE.md) | Para mantenedores del blueprint |

## Checklist del documento

- [x] Scope con 19 features detalladas + 3 catálogos config seed-only
- [x] Inventario 11 tablas backend + 3 audit/log (volúmenes reales)
- [x] Hallazgos D98 CRITICAL (PII) + D99 (FK email) con anchor links
- [x] Decisiones UX Preservar/Mejorar/Nuevo paradigma por pantalla
- [x] Decisiones arquitectura hexagonal ports
- [x] Decisiones de seguridad (CRITICAL D98 PII + D99 FK email)
- [x] Decisiones de datos (FKs conceptuales, booleanos, ZOMBIE)
- [x] Criterios de aceptación verificables y agrupados por dimensión
- [x] 16 pendientes operacionales antes, 12 durante, 4 en go-live (incluye audit PENDIENTE)
- [x] 22 tickets derivables preview (TK-HPS-SOL-1..22) con TK-HPS-SOL-22 marcado como CRITICAL GATE
- [x] Tabla de decisiones referenciadas (D5-D101)
- [x] Tabla de fuentes
- [x] Idioma: español técnico neutro. Identificadores y paths sin traducir.
- [x] "The sentence that organizes" presente
- [x] "Scope del scope" presente
- [x] Sin emojis decorativos
- [x] Cross-references a DOCS, CODEBASE-GUIDE, AGENTS
- [x] Quick Navigation table
- [x] Hallazgos en tabla con severity

## Cierre del ciclo de épicas

Esta épica cierra el **ciclo de refactor de las 8 épicas** del blueprint access2web. Las 8 apps legacy cubiertas:

| # | App | Estado refactor |
|---|---|---|
| 1 | Condor | ✅ PR #5 mergeado |
| 2 | Brass | ✅ PR #6 mergeado |
| 3 | Gestion_Riesgos | ✅ PR #8 mergeado |
| 4 | NoConformidades | ✅ PR #9 mergeado |
| 5 | Expedientes | ✅ PR #7 mergeado |
| 6 | Lanzaderas | ❌ Épica mergeada en PR #4 (sin refactor — fuente del bug D167) |
| 7 | HPS | ✅ PR #10 mergeado |
| 8 | **HPS_Solicitudes** | (este PR) |

Pendientes acumulados:
- **Lanzaderas** sin refactor — fuente del bug D167 cross-app ambiguity (`Form_Formulario1`).
- **Audit usage pendiente** en HPS_Solicitudes (TK-HPS-SOL-22) y Lanzaderas (sin audit).
- **Decisión sobre `Form_Formulario1`** (D167/D178 cross-app) — renombrar a namespace explícito.

## Siguiente paso

Una vez mergeado el PR, el ciclo de refactor de épicas está completo. Pendiente:
- **Refactor de Lanzaderas** (corrección de bug D167).
- **Audit usage de HPS_Solicitudes** (TK-HPS-SOL-22).
- **Decisión sobre `Form_Formulario1`** (D167/D178 cross-app).
- **Audit final cruzada** de las 8 épicas refactorizadas.
- **Decisión sobre integración final** entre las 8 apps (especialmente cross-app con Lanzadera como madre).

---

[← Back to HPS_Solicitudes README](README.md) · [← Codebase Guide](../../../CODEBASE-GUIDE.md) · [← DOCS](../../../DOCS.md)
