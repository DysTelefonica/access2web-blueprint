# Épica — HPS_Solicitudes (migración a web)

> **Estado:** DRAFT — pendiente revisión.
> **Versión:** v0.1 (2026-08-06).
> **Autor:** placeholder.
> **Lote de discovery:** 7 (junto al resto de las 7 apps restantes).
> **Cross-refs:** `docs/03-aplicaciones/hps-solicitudes/{capabilities,data-model,forms,integrations-automation,migration-matrix,security-rules,README}.md` · engram obs #24081 (schemas detalle) · D83 (independencia con HPS) · D98-D101 disposiciones específicas de HPS_Solicitudes.

## Metadatos

| Campo | Valor |
|---|---|
| **Aplicación legacy** | `HPS_SOLICITUDES` · frontend `Solicitudes_HPS.accdb` (22 MB) + backend `Solicitudes_HPS_datos.accdb` (13 MB, autoritativo en `C:\00repos\datos\`) |
| **Tipo de migración** | Legacy Access/VBA → web hexagonal (FastAPI + HTMX) |
| **Scope size** | **S** (11 tablas, 19 features de negocio, 26 clases, 14 módulos, 1 anexo Excel, 1 integración externa ONS) — **la app más simple en superficie de datos** pero **en uso activo** (245 solicitudes reales vs 1 de Condor en staging) |
| **Dependencias cross-app** | Lanzadera (identidad/permisos, vía `getdbLanzadera()`); HPS gestión de usuarios (`IDUsuarioHPS` FK conceptual + flag `RegistroEnHPS="Sí"`, **app independiente ID 22 vs HPS ID 17 — D83**); Lanzadera `TbExpedientes` (`IDExpediente` FK conceptual); ONS — sistema externo (D101) |
| **Riesgo dominante** | Seguridad D98 — **245 filas con datos personales completos** en `TbSolicitudes` (DNI, nombres, fechas, correo, teléfono); D99 — FK por **email** en vez de por ID (`TbResponsables.Correo → TbSolicitudes.emailResponsable`) — anti-patrón de integridad referencial |
| **Stack target** | Backend Python 3.12+ / FastAPI 0.119+ / SQLAlchemy 2.0.x / Alembic 1.13+ / asyncpg 0.30+ (D66) · Frontend HTMX 2.0.4 + Jinja2 3.1+ / async + Alpine.js 3.15+ (D67) |
| **Estrategia de migración de BD** | Expand and Contract backward-compatible (D82) · PostgreSQL compartido con esquema por módulo (D14) |
| **Forma destino** | Hexagonal global (D8) · módulo dentro del monolito modular (D68) · puerto de persistencia PostgreSQL + object storage S3-compatible (D16) + secret manager (D9-D10) |
| **Auditoría de uso previa** | ⚠️ **PENDIENTE** (mismo gap que tuvo HPS antes de obs #24097). El schema usage audit de las 11 tablas backend **no se ha ejecutado todavía**. Gate previo a fase SDD; ver § 1.4, § 2.1.1 y § 6 (TK-HPS-SOL-22). |

## Forma del documento

1. Scope (en / fuera)
2. Estado del descubrimiento
3. Hallazgos críticos
4. Decisiones aplicadas
5. Criterios de aceptación
6. Pendientes operacionales
7. Tickets derivables (preview)

---

## 1. Scope

### 1.1 En scope

**19 features de negocio** con paridad funcional en la nueva plataforma:

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

**3 catálogos config seed-only** (datos de runtime, sin CRUD de usuario):

| Catálogo | Volumen | Notas |
|---|---|---|
| `motivos_hps` (era `MotivoHPS.cls`) | TBD | Catálogo de motivos alta/baja HPS |
| `hps_grado` (era `TbHPSGrado`, clave compuesta `(TipoHPS, Grado)`) | **13 filas** | Grados HPS con clave compuesta |
| `responsables` (era `TbResponsables`) | 27 filas | Responsables con FK por email (D99) |

**Volumen real autoritativo** (backend `Solicitudes_HPS_datos.accdb` en `C:\00repos\datos\`):

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

### 1.2 Fuera de scope (REPLACE)

> Las **tablas puras de log** NO migran como tablas PostgreSQL. Se reemplazan por stack de observabilidad web-native (Sentry / OpenTelemetry / structured logs a Loki o CloudWatch). Sus llamadas VBA se traducen a eventos web.

- `TbLogsGeneral` (2058 filas en staging) — log general activo con campos presumidos de timestamp + actor + contexto. **REPLACE web-native** (D27). Reflejo del veredicto de Condor (obs #24085 + decisión B del lote 5).
- `TbLogs` (0 filas, presumible desuso, D100) — descartar; nunca se reactiva en PostgreSQL.
- `TbCorreosEnviados` (9 filas) — **evaluar caso a caso**: contiene metadatos de envío con FK a `TbSolicitudes` (ToDo: confirmar si es tabla de negocio o log). Como es bajo volumen, **se preserva como tabla de auditoría de envíos** con retención (D27-D29); NO REPLACE.

**Excluido del REPLACE**: el resto de las 11 tablas backend migran como tablas PostgreSQL con sus claves y datos preservados.

**Persistencia**: engram obs #24081 (schemas detalle de HPS_Solicitudes, 7 tablas documentadas en segunda pasada). Decisión B de logs web-native tomada para Condor (`condor/log-strategy-2026-08-05`) — aplica por simetría cross-cutting.

### 1.3 Fuera de scope (no documentado)

Si aparece algo que no está en los 7 docs de HPS_Solicitudes, se marca como **no documentado** y se acumula en pendientes de discovery para iteración posterior. Ejemplos conocidos:

- Volúmenes reales de producción (no staging): `Solicitudes_HPS_datos.accdb` en `C:\00repos\datos\` no fue inspeccionado en esta pasada más allá del inventario Dysflow base.
- Conteos reales de `TbConfiguracion` (presumido 1 fila) y `TbUltimoCambio` (presumido 245 filas 1:1 con solicitudes) — no auditados en staging.
- Schema de `TbCorreosEnviados` — los 21 campos están documentados en `data-model.md`, pero el contenido (qué plantillas, qué destinatarios, qué periodicidad) no fue inspeccionado.
- Schemas no inspeccionados en la pasada inicial (`TbLogs`, `TbLogsGeneral`, `Copia de TbExpedientes`) — fallaron por colección vacía; la segunda pasada documentó las 7 tablas principales pero estas 4 quedaron con shape solo en inventario.
- Inventario completo de queries exportadas y macros embebidas — codegraph no captura macros embebidas ni QueryDefs no exportados.
- Plantillas de correo en `Correo.cls` — pueden contener datos sensibles (D98); revisión pendiente antes de portar.
- Contenido del Excel subido por el solicitante (`Form_FormAdjuntarExcelSolicitante`) — qué columnas se esperan, qué validaciones; no inspeccionado en detalle.
- `Automiatizacion.bas` (typo) — qué automatizaciones exactamente (traspasos a ONS, renovaciones, recordatorios); no inspeccionado en detalle.
- Contenido de `Form_FormPlantillasHTML.cls` y `Form_FormWeb.cls` — qué plantillas concretas, qué versión, qué merge.

### 1.4 Fuera de scope (no-migrate — pendiente de audit)

> ⚠️ **Schema usage audit PENDIENTE**. Las 11 tablas backend **no han sido clasificadas** como ACTIVE / ZOMBIE / UNCLEAR. El veredicto equivalente al de HPS (obs #24097, 12 ACTIVE / 1 UNCLEAR / 9 ZOMBIE) **no existe todavía** para HPS_Solicitudes. Gate previo a fase SDD; ver § 6 (TK-HPS-SOL-22).

**Limitación declarada**: el inventario Dysflow sí ejecutó (`TbSolicitudes` con 245 filas, `TbLogsGeneral` con 2058 filas, etc.), pero la extracción de callers via `codegraph-vba` sobre `C:\00repos\codigo\HPS_SOLICITUDES` no se corrió para clasificar las tablas por uso real en código (`.bas`/`.cls`). La categoría "no-migrate — ZOMBIE" no tiene veredicto formal todavía; se sospecha que `Copia de TbExpedientes` (legacy copy) y `TbLogs` (0 filas, presumible desuso, D100) son candidatas a ZOMBIE, pero la confirmación requiere ejecutar el audit.

**Candidatas probables** (sin veredicto formal, pendiente de audit obs similar a #24097):

- `Copia de TbExpedientes` — patrón legacy copy pre-mass-change (4× en HPS); retención indefinida, sin endpoints, sin UI.
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
| **Schemas documentados** | **7/11** tablas en `data-model.md` (segunda pasada): `TbSolicitudes` (28 col, completo), `TbResponsables` (4 col), `TbJustificaciones` (4 col), `TbHPSGrado` (2 col), `TbSolicitudesFechas` (21 col), `TbConfiguracion` (15 col), `TbUltimoCambio` (4 col), `TbCorreosEnviados` (21 col). **4/11 con shape solo inventario**: `TbLogs` (vacía, D100), `TbLogsGeneral` (alta volumetría, falló get_schema), `Copia de TbExpedientes` (legacy copy, presumida). |
| **Uso (audit)** | ⚠️ **PENDIENTE**. No se ha ejecutado `codegraph_explore` para clasificar las 11 tablas backend. Gate previo a fase SDD (TK-HPS-SOL-22). Limitación equivalente a la de HPS antes de obs #24097. |
| **Subcategorías** | 6 dominio (solicitudes, fechas, responsables, justificaciones, motivos, config) · 3 audit/log (TbLogs, TbLogsGeneral, TbUltimoCambio) · 1 catálogo compuesto (hps_grado) · 1 legacy copy (Copia de TbExpedientes) |
| **FKs físicas (backend)** | **3** entre user tables, todas con problemas (ver § 2.2) |
| **FKs conceptuales (sin constraint)** | **5** — `IDExpediente`, `IDUsuarioHPS`, `IDEmpresaUsuario`, `IDEmpresaTramitadora`, `idjustificacion` (en `TbSolicitudes`) |
| **Clases inventariadas** | 26 (14 domain, 8 operaciones/servicios, 2 compartidas con Lanzadera, 1 util, 1 errata) |
| **Módulos** | 14 (5 bootstrap/factory/DAO, 3 repositorios, 2 tests, 4 miscelánea — incluye `Automiatizacion.bas` con typo) |
| **Forms** | ~30+ archivos `Form_*.cls` cada uno con su `.form.txt` compañero |
| **Tests VBA** | 2 archivos (`Test.bas`, `TestParametrosParser.bas`) — cobertura básica (D87 referencia, no enriquecida) |
| **Integraciones externas** | ONS (sistema externo, D101) — único sistema externo fuera de Lanzadera/HPS/Expedientes |

### 2.1.1 Schema usage audit — PENDIENTE (gap crítico)

> ⚠️ **No se ha ejecutado schema usage audit sobre las 11 tablas backend**. Mismo gap que tuvo HPS antes de obs #24097. El veredicto (ACTIVE / UNCLEAR / ZOMBIE) es **desconocido**; los candidatos probables están en § 1.4 pero sin clasificación formal.

**Pendiente ejecutar** (análogo a obs #24097 de HPS):

1. `codegraph_explore` sobre `C:\00repos\codigo\HPS_SOLICITUDES\00_main` (o la raíz directa del repo según `.codegraph-vba`/`staging/.codegraph-vba` confirmado en obs #23876).
2. Para cada una de las 11 tablas: extraer callers en `.bas`/`.cls`, blast radius de símbolos clave (`EVE`, `getdb`, `Solicitud`, `Responsable`, `Justificacion`, `LogGeneral`, `Correo`).
3. Clasificar ACTIVE / UNCLEAR / ZOMBIE.
4. Documentar el veredicto como observación engram con topic_key `hps-solicitudes/usage-audit-<fecha>`.
5. Cerrar el gap con un commit que aplique el veredicto al epic (similar a commit `8d5311c` para HPS).

**Limitaciones que se esperan** (mismas que obs #24097 documentó para HPS):

- Macros embebidos en `Solicitudes_HPS.accdb` NO exportados → uso real puede ser mayor al lower bound.
- Forms sin `RecordSource`/`RowSource` exportados: app carga vía DAO + clases.
- `codegraph_explore` limita el blast radius visible a 5 entries por símbolo.
- Posibles gaps del extractor en patrones específicos de esta app (`m_SQL='...'`, `m_Obj*.SetSQL`).

**Gate explícito**: la fase SDD de HPS_Solicitudes **no arranca** hasta que este audit esté committed. Ver § 6 (TK-HPS-SOL-22 como bloqueante) y § 7 (TK-HPS-SOL-22 como ticket primario de seguridad antes de SDD).

### 2.2 Hallazgos estructurales del inventario

**FKs con problemas** (3 FKs físicas entre user tables, todas con inconsistencias, `data-model.md § Relaciones físicas reales`):

1. **`TbResponsables.Correo → TbSolicitudes.emailResponsable`**: FK por **texto email**, no por ID. Si el email cambia en `TbResponsables`, la FK lógica se rompe. **Anti-patrón de integridad referencial** (D99). Migración: agregar columna `idResponsable` (BIGINT) en `TbSolicitudes` y poblar con conversión `emailResponsable → id`. Mantener `emailResponsable` como campo independiente (denormalización para búsqueda).
2. **`TbJustificaciones.idjustificacion → TbSolicitudes.idjustificacion`**: dirección de FK **confusa**. La FK va de `TbJustificaciones` a `TbSolicitudes`, no al revés. La Solicitud tiene su propio `idjustificacion` y la Justificación la referencia de vuelta. Migración: formalizar como FK con la dirección correcta (`TbSolicitudes.idjustificacion → TbJustificaciones.id`).
3. **`TbSolicitudes.IDSolicitud → TbSolicitudesFechas.IDSolicitud`**: 1:N clásico, la única FK limpia del modelo. PK de `TbSolicitudesFechas = (IDSolicitud)` (no `BIGSERIAL`, asume 1:1 estricto).

**FKs conceptuales sin constraint** (5 en `TbSolicitudes`):
- `IDExpediente` (cross-app Lanzadera).
- `IDUsuarioHPS` (cross-app HPS gestión de usuarios, ID 17).
- `IDEmpresaUsuario` (cross-app).
- `IDEmpresaTramitadora` (cross-app).
- `idjustificacion` (intra-app a `TbJustificaciones`, sin FK formal — anomalía D99).

**Booleanos como `Text(2)` 'Sí/No'** (D102 cross-cutting): **4 columnas** confirmadas:
- `TbConfiguracion.AutocancelacionPreMARGA` (Text 2, NULL).
- `TbConfiguracion.AutocancelacionMARGA` (Text 2, NULL).
- `TbConfiguracion.CorreosAutomaticos` (Text 2, NULL).
- `TbCorreosEnviados.Programado` (Text 2, NULL).
- Adicional: `TbSolicitudes.CorreosAutomaticos` (Text 2, NULL) — total **5 columnas booleanas como texto**.

**`TbLogs` vacía — presumible desuso** (D100): 0 filas vs 2058 en `TbLogsGeneral`. La nueva plataforma migra `TbLogsGeneral` como tabla principal de auditoría (o REPLACE web-native según veredicto § 1.2).

**PK compuesta en `TbHPSGrado`**: `(TipoHPS, Grado)` — clave natural compuesta. Migración: `PRIMARY KEY (tipo_hps, grado)` en PostgreSQL.

**`TbSolicitudesFechas` con 20 fechas regulatorias** (D100-bis / crítico): 21 columnas = 1 FK (`IDSolicitud`) + 20 fechas. Esto es un **workflow regulatorio completo de HPS** con eventos: envío/recepción de Excel, recordatorios (4 variantes), desestimación, autocancelación (2 variantes), traspasos a ONS (1 variante), cancelaciones (2 variantes). Decidir si se normaliza (consolidar fechas relacionadas) o se preserva tal cual.

**`TbConfiguracion` con 6 campos de días** (config del workflow regulatorio): `DiasParaRecordatorio*` (4) + `DiasCancelacion*` (2). Migración: config del módulo + tabla de workflow declarativo (similar a D96 de Condor).

**Duplicación `Solicitudes_HPS_datos.accdb`**: backend en `C:\00repos\datos\` (autoritativo) y duplicado en la raíz del repo (`Solicitudes_HPS_datos.accdb` local). Patrón legacy (igual que HPS, Condor). Solo el autoritativo cuenta; el local es legacy.

### 2.3 Herramientas usadas

- **Dysflow MCP** (read-only) sobre `Solicitudes_HPS.accdb` + `Solicitudes_HPS_datos.accdb`: inventario real, conteo de filas, FKs, esquema de las 7 tablas principales en segunda pasada.
- **codegraph-vba** sobre `C:\00repos\codigo\HPS_SOLICITUDES` (índice verificado en obs #23876): extracción de call paths (44 símbolos en 2 archivos para la query inicial). **No se ejecutó `codegraph_explore` de audit de uso** específico para HPS_Solicitudes — pendiente antes de fase SDD.
- **No se ejecutó write-class operation**: setup_project creó `.dysflow/project.json` en `HPS_SOLICITUDES/` con permiso del user, pero el resto del flujo fue read-only.

### 2.4 Pendientes menores de discovery

| Pendiente | Impacto | Iteración |
|---|---|---|
| **Audit de uso específico para HPS_Solicitudes** (codegraph-vba + cruzar con `src/`) | Determinar ACTIVE/ZOMBIE/UNCLEAR de las 11 tablas backend | ⚠️ **PENDIENTE** (gate TK-HPS-SOL-22 antes de SDD) |
| Conteos reales de `TbConfiguracion` y `TbUltimoCambio` | Ajustar criterios de aceptabilidad de migración | Cuando se acceda al backend autoritativo |
| Schemas de `TbLogs`, `TbLogsGeneral`, `Copia de TbExpedientes` (4 tablas con shape solo inventario) | Cerrar el inventario al 100% | Próxima pasada Dysflow con `get_schema` (posible workaround para colección vacía) |
| Contenido de `TbCorreosEnviados` (qué plantillas, qué periodicidad, qué To/CC/BCC típicos) | Evaluar el patrón de notificación para el servicio unificado | Próxima iteración |
| Cadena de uso de `Copia de TbExpedientes` por saved queries / macros | Decidir archivo vs dominio | Inspección manual si negocio requiere |
| Plantillas de correo en `Correo.cls` con datos sensibles (D98) | Política de privacidad antes de portar | Antes de TK-HPS-SOL-19 (notificaciones) |
| Validación de las **20 fechas regulatorias** en `TbSolicitudesFechas`: normalizar vs preservar tal cual | Decisión de modelado (¿workflow declarativo tipo D96 Condor o solo TIMESTAMPs?) | Antes de TK-HPS-SOL-3 (solicitudes + fechas) |
| Auditoría operativa del frontend `Solicitudes_HPS.accdb` en `HPS_SOLICITUDES` (no desde aquí): `git log --all -- Solicitudes_HPS.accdb`, añadir `*.accdb` al `.gitignore` | Riesgo de PII en historial git (D98 cerrado) | TK-HPS-SOL-21 (operativo en repo `HPS_SOLICITUDES`) |
| Inspección de contenido de `Form_FormAdjuntarExcelSolicitante.cls` (qué columnas del Excel se parsean, qué validaciones) | Decidir endpoint de upload + parser server-side (openpyxl/pandas) | Antes de TK-HPS-SOL-6 (adjuntos Excel) |
| Inspección de `Automiatizacion.bas` (typo) y `AutomatizacionRepositorio.bas` — qué jobs hay | Decidir jobs del scheduler unificado (D59) con kill switch (D91) | Antes de TK-HPS-SOL-18 (automatización) |

---

## 3. Hallazgos críticos

Cada hallazgo: descripción corta + impacto + acción + referencia documental.

<a id="hallazgo-D98"></a>
### H1 · D98 — 245 filas de PII en `TbSolicitudes`  ⚠️⚠️⚠️ CRÍTICO

- **Descripción**: `TbSolicitudes` (28 columnas, 245 filas) contiene **datos personales completos** del solicitante: `DNI`, `Nombre`, `Apellido1`, `Apellido2`, `FNacimiento`, `LugarNacimiento`, `email`, `Telefono` — 8 columnas con PII. Mismo patrón que HPS (D92) y Lanzadera. La app es la **más pequeña en superficie pero también tiene datos personales en producción** (245 solicitudes reales).
- **Impacto**:
  1. Si el frontend `Solicitudes_HPS.accdb` (22 MB) o el backend `Solicitudes_HPS_datos.accdb` se versionan, **245 usuarios con datos personales quedan en el historial de git**.
  2. Datos personales embebidos en correo (`TbCorreosEnviados.Destinatarios` con memo) y en adjuntos ONS (`URLAdjuntoEnvioONS`).
  3. Mismas reglas que HPS_D92: frontend se mueve con los datos, copias no cifradas, robo de portátil = exposición.
- **Acción** (combinada, operativa + migración):
  - **Operativo en `HPS_SOLICITUDES`** (no desde aquí): añadir `*.accdb` al `.gitignore`; verificar `git log --all -- Solicitudes_HPS.accdb`; si está versionado, **rotar repo + git-filter-repo**.
  - **Migración**: las 8 columnas PII se preservan en PostgreSQL **sin transformaciones** (mantener paridad funcional) pero se enmascaran en logs y observabilidad (D27); capabilities declaran visibilidad por campo (D45); evaluar encriptación en reposo como follow-up.
  - Plantillas de correo (`Correo.cls`) y adjuntos ONS (`URLAdjuntoEnvioONS`) requieren revisión previa antes de portar al servicio unificado.
- **Detalle completo**: [`migration-matrix.md § D98`](migration-matrix.md#d98--datos-personales-en-tbsolicitudes-245-filas) · [`security-rules.md § D98`](security-rules.md#d98--datos-personales-en-tbsolicitudes-245-filas) · engram obs #24081 (schemas detalle).

<a id="hallazgo-D99"></a>
### H2 · D99 — FK por email en `TbResponsables.Correo → TbSolicitudes.emailResponsable`  ⚠️⚠️ CRÍTICO

- **Descripción**: la FK lógica entre responsables y solicitudes se hace por **texto email**, no por ID numérico (`data-model.md § Relaciones físicas reales`). `TbResponsables.Correo` (VARCHAR NOT NULL) → `TbSolicitudes.emailResponsable` (VARCHAR NULL). Si un responsable cambia su email en `TbResponsables`, la FK lógica se rompe silenciosamente. **Anti-patrón de integridad referencial** que el modelo legacy arrastró.
- **Impacto**:
  1. **Data integrity gap real**: no hay manera en el legacy de detectar la inconsistencia (no hay constraint formal).
  2. Riesgo de **solicitudes huérfanas** al migrar: si el código legacy asume la FK por email y la nueva plataforma la asume por ID, todas las vinculaciones se pierden.
  3. En PostgreSQL hay que formalizar FK numérica; en la migración de datos hay que **poblar el `idResponsable` retroactivamente** desde la conversión del email.
- **Acción**:
  - En PostgreSQL, agregar columna `idResponsable` (BIGINT) en `TbSolicitudes` y poblar con la conversión `emailResponsable → id` durante la migración de datos.
  - Mantener `emailResponsable` como campo independiente (denormalización para búsqueda por texto).
  - Documentar la convención: la FK canónica es por ID; el email es solo campo de referencia para join.
  - Evaluar si el `emailResponsable` puede ser NULL cuando el responsable no tiene email (preservar comportamiento legacy).
- **Detalle completo**: [`migration-matrix.md § D99`](migration-matrix.md#d99--fks-conceptuales-con-data-integrity-gaps) · [`security-rules.md § D99`](security-rules.md#d99--fks-conceptuales-con-data-integrity-gaps).

<a id="hallazgo-D83-hpssolic"></a>
### H3 · D83 — HPS_Solicitudes (ID 22) es app **independiente** de HPS (ID 17)

- **Descripción**: HPS_Solicitudes (`TbAplicaciones.ID = 22`) y HPS (`TbAplicaciones.ID = 17`) comparten **solo el nombre conceptual** por razones históricas; son productos distintos con su propio frontend (`Solicitudes_HPS.accdb` vs `HPS.accdb`), backend (`Solicitudes_HPS_datos.accdb` vs `HPST.accdb`), checkout y permisos. La columna `TbSolicitudes.IDUsuarioHPS` es una **FK conceptual** que referencia a `TbUsuarios` de HPS gestión de usuarios, no comparte tabla. Decisión **D83 APROBADO**.
- **Impacto**:
  1. Las migraciones de HPS y HPS_Solicitudes pueden ocurrir **en orden independiente** sin acoplamiento de schema.
  2. La FK conceptual `IDUsuarioHPS` se mantiene mediate adaptador unificado (D9-D10), no como FK física.
  3. La app tiene su propio IDAplicacion (`"22"`), su propia config (`TbConfiguracion`), su propio log (`TbLogsGeneral`).
  4. **Confusión frecuente**: el nombre "HPS_Solicitudes" lleva a pensar que es un submódulo de HPS — **NO lo es**. Documentación y tickets deben nombrarla explícitamente como "HPS_Solicitudes" o "Solicitudes HPS (ID 22)".
- **Acción**:
  - Tickets y migraciones **no comparten schema** entre las dos apps. Cada módulo hexagonal es independiente.
  - La FK conceptual `IDUsuarioHPS` se mantiene como referencia conceptual mediate adaptador (D86/D87).
  - Documentar en capabilities.md y en el header de cada ticket el `IDAplicacion = "22"` para evitar confusión con HPS ID 17.
- **Detalle completo**: D83 en `docs/08-decisiones-y-preguntas-abiertas.md` · engram obs #24023 (P1 resuelta con filesystem evidence).

<a id="hallazgo-D100"></a>
### H4 · D100 — `TbLogs` vacía, presumible desuso (vs `TbLogsGeneral` con 2058 filas)

- **Descripción**: `TbLogs` tiene **0 filas** mientras que `TbLogsGeneral` tiene **2058 filas** (presumido log general activo). Esto sugiere que `TbLogs` está en desuso y fue reemplazada por `TbLogsGeneral`. **Migración**: REPLACE web-native (D27) — `TbLogsGeneral` se traduce a logs estructurados canónicos (Sentry / OpenTelemetry / Loki), NO migra como tabla PostgreSQL.
- **Impacto**: 2058 filas de log en staging que se traducen a eventos web con correlación (D27). Las queries VBA que insertan en `TbLogsGeneral` se traducen a `logger.info(...)` / `logger.warning(...)` con contexto estructurado.
- **Acción**:
  - `TbLogs` se descarta (0 filas, sin reactivarla).
  - `TbLogsGeneral` se traduce a stack web-native; ninguna tabla `tb_logs_general` aparece en PostgreSQL.
  - El volumen 2058 es el lower bound del inventario; en producción puede ser mayor.
- **Detalle completo**: [`migration-matrix.md § D100`](migration-matrix.md#d100--tblogs-vacía--presumible-desuso).

<a id="hallazgo-D101"></a>
### H5 · D101 — Integración con ONS (Organismo Notificador de Seguridad) ⚠️

- **Descripción**: HPS_Solicitudes tiene un sistema completo de **traspasos a ONS** con adjuntos: `Form_FormAdjuntaTraspasoONS.cls` para adjuntar traspasos, `URLAdjuntoEnvioONS` (Memo en `TbSolicitudes`), y 4 forms específicos de traspaso (`Form_FormSolicitudAltaTraspaso`, `Form_FormSolicitudesAltaTraspasoDatos`, `Form_FormSolicitudesTraspasoFechas`, etc.). La integración con ONS es un **sistema externo** fuera del scope de la migración.
- **Impacto**:
  1. La columna `ExpedienteUnificado = "No"` indica que la unificación de expediente está desactivada (posible feature futuro).
  2. Los adjuntos ONS (`URLAdjuntoEnvioONS`) pueden contener datos sensibles del solicitante (D98).
  3. El workflow regulatorio tiene 20 fechas específicas (`TbSolicitudesFechas` columna `FechaEnvioTraspasoONS` y variantes) — debe preservarse.
- **Acción**:
  - Migrar como **adaptador de salida** (D16, object storage) — serializa la solicitud + adjuntos y los envía al servicio externo ONS vía API o cola.
  - `URLAdjuntoEnvioONS` se reemplaza por URL firmada del object storage que ONS puede descargar.
  - Endpoint "traspaso a ONS" que registre el evento en logs estructurados (D27).
  - Auditar el estado actual de la integración con ONS (¿sigue activa? ¿qué periodicidad? ¿qué campos se sincronizan?) — pendiente.
- **Detalle completo**: [`migration-matrix.md § D101`](migration-matrix.md#d101--integración-con-ons-organismo-notificador-de-seguridad) · [`integrations-automation.md § Integración con ONS (D101)`](integrations-automation.md#integración-con-ons-d101).

### Hallazgos cross-cutting que aplican a HPS_Solicitudes

<a id="hallazgo-D102-hpssolic"></a>
### H6 · D102 — Booleanos como `Text(2)` 'Sí/No' (cross-cutting)

- **Descripción**: **5 columnas** confirmadas con `VARCHAR(2) NULL` y valores 'Sí'/'No':
  - `TbConfiguracion.AutocancelacionPreMARGA`
  - `TbConfiguracion.AutocancelacionMARGA`
  - `TbConfiguracion.CorreosAutomaticos`
  - `TbSolicitudes.CorreosAutomaticos`
  - `TbCorreosEnviados.Programado`
- **Impacto**: queries booleanas se complican; PostgreSQL debe estandarizar a `BOOLEAN`. Cualquier valor fuera de 'Sí'/'No' en origen requiere default explícito. Regla general cross-cutting (D102).
- **Acción**: regla de migración explícita — `'Sí' → TRUE`, `'No' → FALSE`, `NULL → NULL`, otros → `FALSE DEFAULT`. Aplicar en HPS_Solicitudes con el mismo criterio que HPS/Lanzadera/Expedientes/Gestion_Riesgos/NoConformidades.
- **Detalle completo**: [`migration-matrix.md`](migration-matrix.md#d99--fks-conceptuales-con-data-integrity-gaps) · [`data-model.md`](data-model.md) § TbConfiguracion, TbSolicitudes, TbCorreosEnviados.

<a id="hallazgo-D44-hpssolic"></a>
### H7 · D44 — Suplantación (cross-cutting, debe respetarse)

- **Descripción**: el patrón `suplantadoPor` se loguea en logs estructurados (D27). Suplantación restringida al administrador global con doble identidad visible y auditoría completa (D44).
- **Impacto**: HPS_Solicitudes no introduce suplantación propia (no tiene logs VBA que aplicarían el patrón), pero debe respetar la política cross-cutting de plataforma al integrar con adaptadores de identidad.
- **Acción**: capabilities (D45-D46) verifican que solo el administrador global puede iniciar suplantación. Audit completa en logs estructurados (D27).
- **Detalle completo**: D44 en `docs/08-decisiones-y-preguntas-abiertas.md`.

<a id="hallazgo-D86-hpssolic"></a>
### H8 · D86/D87 — Forma hexagonal + tests VBA como evidencia (cross-cutting)

- **Descripción**: HPS_Solicitudes tiene 26 clases (dominio + operaciones + compartidas con Lanzadera) + 14 módulos + 2 archivos `Test_*.bas` con cobertura **básica** (no como HPS con 9 archivos). Indica forma hexagonal previa al blueprint moderno (D86): dominio en clases, helpers por dominio, transaccionalidad en clases de operaciones. Los tests VBA existentes son **evidencia de comportamiento** que se preserva como referencia para los nuevos tests pytest (D87), aunque la cobertura es limitada.
- **Impacto**: ningún `Test_*.bas` se descarta sin trazabilidad. Tests VBA son **referencia base** (no enriquecida) — la nueva plataforma ampliará cobertura con pytest + httpx + Playwright.
- **Acción**: cada test VBA tiene un equivalente pytest en la migración. La cobertura limitada significa que la épica debe **ampliar** con tests nuevos en lugar de solo portar.
- **Detalle completo**: D86/D87 en `docs/08-decisiones-y-preguntas-abiertas.md`.

<a id="hallazgo-D102-fk-intra"></a>
### H9 · D94 — FKs conceptuales sin constraint (HPS_Solicitudes)

- **Descripción**: además de la FK por email de D99, hay 2 inconsistencias adicionales:
  - `TbJustificaciones.idjustificacion → TbSolicitudes.idjustificacion`: dirección de FK confusa (la FK va de Justificaciones hacia Solicitudes, no al revés).
  - 5 FKs conceptuales cross-app (`IDExpediente`, `IDUsuarioHPS`, `IDEmpresaUsuario`, `IDEmpresaTramitadora`, `idjustificacion`) sin constraint.
- **Impacto**: integridad referencial no garantizada en legacy. En PostgreSQL hay que formalizar las FKs intra-app (D94, D99) y mantener las FKs cross-app como referencia conceptual mediate adaptadores (D86/D87).
- **Acción**:
  - **Intra-app**: formalizar con `FOREIGN KEY` real — `TbSolicitudesFechas.IDSolicitud → TbSolicitudes.IDSolicitud` (la única FK limpia), `TbSolicitudes.idResponsable → TbResponsables.IDResponsable` (nueva columna, D99), `TbSolicitudes.idjustificacion → TbJustificaciones.idjustificacion` (corregir dirección).
  - **Cross-app**: mantener como referencia conceptual — `IDExpediente` (Lanzadera), `IDUsuarioHPS` (HPS gestión de usuarios, ID 17, D83), `IDEmpresaUsuario`, `IDEmpresaTramitadora`.
- **Detalle completo**: [`migration-matrix.md § D99`](migration-matrix.md#d99--fks-conceptuales-con-data-integrity-gaps) · [`data-model.md` § Relaciones físicas reales](data-model.md#relaciones-físicas-reales-3-fk-entre-user-tables).

---

## 4. Decisiones aplicadas

> Las decisiones que ya están tomadas y que esta épica respeta sin复议ar.

<a id="dec-d98-pii"></a>
### D-PII · Política de datos personales (D98)

- **Alcance**: las columnas `DNI`, `Nombre`, `Apellido1`, `Apellido2`, `FNacimiento`, `LugarNacimiento`, `email`, `Telefono` en `TbSolicitudes` (y equivalentes en correos/adjuntos) requieren política explícita:
  - **Preservar en PostgreSQL** sin transformaciones para mantener paridad funcional.
  - **Enmascarar en logs y observabilidad** (no loguear valores completos).
  - **Documentar en matriz de capabilities** quién puede ver cada campo (D45).
  - **Evaluar encriptación en reposo** como follow-up de seguridad (no bloqueante para la migración inicial).
  - **Auditar el `.gitignore` del repo** para asegurar que `Solicitudes_HPS.accdb` no está siendo commiteado (regla D92 cross-cutting).
- **Motivación**: 245 filas con datos personales completos; mismo riesgo que HPS (D92) en escala menor.
- **Persistencia**: engram obs #24081 (schemas detalle); [`migration-matrix.md § D98`](migration-matrix.md#d98--datos-personales-en-tbsolicitudes-245-filas).

### D-FK-email · FK por email → FK numérica con columna nueva (D99)

- **Alcance**: la FK `TbResponsables.Correo → TbSolicitudes.emailResponsable` se reemplaza en PostgreSQL por una FK numérica `TbSolicitudes.idResponsable → TbResponsables.IDResponsable`. La migración de datos **pobla retroactivamente** la nueva columna con la conversión `emailResponsable → id`. El campo `emailResponsable` se mantiene como columna independiente (denormalización para búsqueda por texto).
- **Motivación**: data integrity gap real; el modelo legacy no podía detectar la inconsistencia. PostgreSQL debe formalizar la FK numérica.
- **Trade-off**: si un responsable cambia su email, el `emailResponsable` queda desincronizado del `idResponsable` (que es la FK canónica). El comportamiento legacy lo permitía; el nuevo debe documentarlo como regla.
- **Persistencia**: [`migration-matrix.md § D99`](migration-matrix.md#d99--fks-conceptuales-con-data-integrity-gaps).

### D-Independencia-HPS · HPS_Solicitudes NO comparte tabla con HPS (D83)

- **Alcance**: HPS_Solicitudes (ID 22) y HPS (ID 17) son apps independientes. **No comparten tablas, no comparten frontend, no comparten backend**. La FK conceptual `TbSolicitudes.IDUsuarioHPS` se mantiene como referencia conceptual mediate adaptador unificado (D86/D87).
- **Motivación**: comparten nombre conceptual por razones históricas pero son productos distintos (D83). El ticket TK-HPS-SOL-8 refleja esta decisión explícitamente.
- **Implicación**: las migraciones de HPS y HPS_Solicitudes pueden ocurrir en orden independiente sin acoplamiento de schema.
- **Persistencia**: D83 en `docs/08-decisiones-y-preguntas-abiertas.md` · engram obs #24023.

### D-Logs · `TbLogsGeneral` (2058 filas) → REPLACE web-native

- **Alcance**: `TbLogsGeneral` (2058 filas) NO migra como tabla PostgreSQL. Se traduce a stack de observabilidad canónico (Sentry / OpenTelemetry / structured logs a Loki o CloudWatch). Las llamadas VBA `LogGeneral.Registrar` se traducen a `logger.info(...)` con contexto estructurado.
- **Motivación**: tooling moderno de observabilidad con rotación/retención resuelta por el stack; queries de logs vía UI/SQL del backend de logs en vez de PostgreSQL.
- **Trade-off**: se pierde la posibilidad de JOINs SQL directos sobre logs. Las queries que actualmente usan `TbLogsGeneral` para auditoría se traducen a queries sobre el stack de observabilidad.
- **Caso especial `TbLogs`**: 0 filas, presumible desuso (D100); **se descarta** sin migración.
- **Persistencia**: engram topic_key `condor/log-strategy-2026-08-05` (aplicada por simetría cross-cutting a HPS_Solicitudes).

### D-Bool · Booleanos `Text(2)` → `BOOLEAN` con regla explícita (D102)

- **Alcance**: las 5 columnas `Text(2)` confirmadas (`TbConfiguracion.AutocancelacionPreMARGA`, `TbConfiguracion.AutocancelacionMARGA`, `TbConfiguracion.CorreosAutomaticos`, `TbSolicitudes.CorreosAutomaticos`, `TbCorreosEnviados.Programado`) se traducen a `BOOLEAN` en PostgreSQL con regla de migración explícita (`'Sí' → TRUE`, `'No' → FALSE`, `NULL → NULL`, otros → `FALSE DEFAULT`).
- **Motivación**: regla general cross-cutting (D102); queries booleanas se complican con texto. Mismo criterio en HPS, Lanzadera, Expedientes, Gestion_Riesgos, NoConformidades.
- **Trade-off**: si en producción aparecen valores fuera de 'Sí'/'No' se loguean como anomalía y se decide caso a caso.

### D-FK-intra · FKs conceptuales intra-app → formales en PostgreSQL

- **Aplicación de D94/D99 a HPS_Solicitudes**: formalizar las 3 FKs intra-app:
  - `TbSolicitudesFechas.IDSolicitud → TbSolicitudes.IDSolicitud` (única FK limpia del legacy).
  - `TbSolicitudes.idResponsable → TbResponsables.IDResponsable` (nueva columna, D99).
  - `TbSolicitudes.idjustificacion → TbJustificaciones.idjustificacion` (corregir dirección, D99).
- Las 5 FKs conceptuales cross-app (`IDExpediente`, `IDUsuarioHPS`, `IDEmpresaUsuario`, `IDEmpresaTramitadora`, `idjustificacion` legacy) se mantienen como referencia conceptual mediate adaptadores.

### D-Identidad · Vía adaptador unificado (D9-D10)

- **Aplicación**: la identidad se resuelve por el adaptador unificado de autenticación, no por acceso directo a la BD de Lanzadera (`getdbLanzadera()`). `IDAplicacion = "22"` queda como config del módulo.
- **Implicación**: el módulo HPS_Solicitudes no conoce la BD de Lanzadera; el adaptador de identidad expone `getUsuarioConPermisos(usuario, aplicacion)` server-side.

### D-Backend-config · `TbConfiguracionBackends` → config del módulo hexagonal

- **Aplicación**: `BackendActivo`, `EnPruebas`, `IDAplicacion`, `PasswordBackend`, `CacheHabilitada` dejan de ser TempVars o flags de runtime; pasan a **config del módulo hexagonal** (variables de entorno del runner). `PasswordBackend` desaparece (D93 cross-cutting, secret manager D9-D10).
- **Caso especial**: HPS_Solicitudes tiene **7 flags TempVars activos** (más que otras apps) — `EnDesarrollo`, `DatosEnLocal`, `EnPruebas`, `ConCorreoCopiaGestor`, `ActivadoCorreoAutomatico`, `RegistroEnHPS`, `ExpedienteUnificado`. Todos se traducen a config del módulo:
  - `EnDesarrollo`, `DatosEnLocal`, `EnPruebas` → variables de entorno del runner.
  - `ConCorreoCopiaGestor`, `ActivadoCorreoAutomatico`, `RegistroEnHPS`, `ExpedienteUnificado` → config del módulo (boolean).
- **Detalle**: [`integrations-automation.md`](integrations-automation.md#flags-de-operación-tempvars-en-eve).

### D-Configuracion-Dias · `TbConfiguracion.Dias*` → config del módulo + workflow declarativo

- **Aplicación**: las 6 columnas de días en `TbConfiguracion` (`DiasParaRecordatorioExcel1/2`, `DiasCancelacionPreMARGA`, `DiasParaRecordatorioRellenoMarga1/2`, `DiasCancelacionMARGA`) configuran el **workflow regulatorio** de las 20 fechas en `TbSolicitudesFechas`. Migración: traducir a **config del módulo + tabla de workflow declarativo** (similar a D96 de Condor) con FKs explícitas.
- **Implicación**: el workflow de HPS_Solicitudes es **data-driven** (días como variables), no hardcoded. La nueva plataforma preserva este patrón.

### D-Plantillas · Plantillas HTML legacy → Jinja2 server-side

- **Aplicación**: `Form_FormPlantillasHTML.cls` y `Form_FormWeb.cls` (Edge WebView embebido) se traducen a **Jinja2 templates** server-side (D66) con auto-escape para evitar XSS. La nueva plataforma absorbe el patrón: las vistas son páginas nativas, no embebidas en Access.
- **Implicación**: las plantillas legacy que combinan datos de solicitud deben ser reescritas como templates Jinja2 + lógica de merge. **Auditar el contenido** antes de portar — pueden contener PII embebida (D98).

### D-Excel · Adjuntos Excel → endpoint de upload con parsing server-side

- **Aplicación**: `Form_FormAdjuntarExcelSolicitante.cls` permite al solicitante adjuntar un Excel con datos. Migración: **endpoint de upload** con parsing server-side (probablemente `openpyxl` o `pandas.read_excel`). El contenido del Excel puede contener PII parseada — manejo cuidadoso.
- **Implicación**: validar el Excel en server-side, no confiar en validación cliente.

### D-MVVM-traductor · Clases/Servicios/Repositorios → use-cases/adaptadores

- **Aplicación**: la separación MVVM en VBA se traduce uno a uno — clases de dominio → entidades Pydantic, `SolicitudOperaciones.cls` → use case Python con `AsyncSession.begin()` (D66, D82), `RepositorioComun.bas` → adaptador de persistencia detrás del puerto.

### D-Testing · Tests VBA como referencia limitada (D87)

- **Aplicación**: los 2 archivos `Test_*.bas` de HPS_Solicitudes son **evidencia de comportamiento básica** que se preserva como referencia para los nuevos tests pytest. La cobertura limitada significa que la migración debe **ampliar** con tests nuevos (pytest + httpx + Playwright), no solo portar.

---

## 5. Criterios de aceptación

Lista verificable de qué define "épica de HPS_Solicitudes cerrada".

### Funcionales (paridad con legacy)

- [ ] Las 19 features de negocio (F1-F19) tienen CRUD en web con paridad funcional mínima (ver § 1.1).
- [ ] El ciclo de solicitud (alta, edición, baja, renovación, cambio de tipo) preserva el workflow regulatorio completo.
- [ ] Las 20 fechas regulatorias en `TbSolicitudesFechas` migran como `TIMESTAMP NULL` preservando todos los eventos del workflow.
- [ ] El workflow regulatorio se carga desde config del módulo + workflow declarativo (similar a D96 de Condor); los días en `TbConfiguracion` son parámetros runtime.
- [ ] Los traspasos a ONS (D101) tienen endpoint dedicado que serializa la solicitud + adjuntos y los entrega al servicio externo.
- [ ] Los adjuntos Excel (`URLAdjunto`) tienen endpoint de upload con parsing server-side; los datos personales extraídos se enmascaran en logs.
- [ ] Las plantillas HTML legacy se traducen a templates Jinja2 server-side con auto-escape (D66).
- [ ] Las justificaciones (`TbJustificaciones`) tienen CRUD con FK corregida a `TbSolicitudes.idjustificacion` (D99).
- [ ] Los responsables (`TbResponsables`) tienen CRUD; la FK por email se reemplaza por FK numérica con columna `idResponsable` (D99).
- [ ] Los motivos HPS (`MotivoHPS.cls`) tienen CRUD completo + UI HTMX.
- [ ] Las tareas pendientes del tramitador (`Form_FormTareasTramitadorPendientes`) tienen worklist con polling HTMX (D69).
- [ ] La configuración (`TbConfiguracion`, 15 columnas) tiene UI admin + 7 flags traducidos a config del módulo.
- [ ] El correo (`TbCorreosEnviados`) migra con FK a `TbSolicitudes` + 21 campos preservados; integración con servicio unificado de notificaciones (D11-D12).
- [ ] El registro automático en HPS (`IDUsuarioHPS` + flag `RegistroEnHPS`) opera via adaptador unificado (D9-D10) — referencia conceptual, NO FK física (D83).
- [ ] La búsqueda general de solicitudes tiene UI HTMX con filtros + paginación (D69, D72).
- [ ] La búsqueda de usuarios HPS tiene UI HTMX con integración con el módulo HPS (vía adaptador).
- [ ] La automatización (`Automiatizacion.bas`) se traduce a jobs del scheduler unificado (D59) con kill switch (D91).
- [ ] Los informes (`InformesOperaciones.bas`) tienen generación parametrizable con vista previa (D62-D63).

### Config / catálogos

- [ ] Los 3 catálogos seed-only (`motivos_hps`, `hps_grado`, `responsables`) están migrados como seed inicial con sus volúmenes originales (TBD, 13, 27).
- [ ] `TbHPSGrado` mantiene su PK compuesta `(TipoHPS, Grado)` en PostgreSQL.
- [ ] Endpoint admin CRUD sobre cada catálogo con control de capabilities (D45).
- [ ] Los 7 flags TempVars están traducidos a config del módulo / variables de entorno del runner.

### Forma y plataforma

- [ ] Módulo hexagonal dentro del monolito modular (D68); puertos por capacidad.
- [ ] Adaptadores driving: web (HTMX) + CLI admin global (D24).
- [ ] Adaptadores driven: PostgreSQL, object storage S3-compatible (D16), secret manager (D9-D10), notificación unificada (D11).
- [ ] Versión semántica del módulo: `hps_solicitudes/v0.1.0-rc.1` en primer RC, `hps_solicitudes/v1.0.0` en release (D78).
- [ ] Migraciones de BD backward-compatibles con Expand and Contract (D82).

### Seguridad y operación

- [ ] Logs en stack web-native (Sentry/OTel); **NO** existen `tb_logs_general`/`tb_logs` en PostgreSQL (verificación binaria).
- [ ] `TbLogs` (0 filas, D100) **descartada** sin migración a PostgreSQL.
- [ ] `PasswordBackend` (cross-cutting D93/D104) **eliminado** del código; secret en secret manager.
- [ ] Datos personales (DNI, nombres, fechas, correos, teléfonos) enmascarados en logs y observabilidad (D98); capabilities declaran visibilidad por campo (D45).
- [ ] Suplantación acotada al administrador global con doble identidad visible (D44); audit completa (D27).
- [ ] Capabilities declaradas en código; UI consulta endpoint de capabilities; servidor rechaza operaciones no autorizadas (D45).

### Datos

- [ ] D94 + D99 formalizado: las 3 FKs intra-app tienen `FOREIGN KEY` en PostgreSQL; `TbResponsables.Correo → TbSolicitudes.emailResponsable` se reemplaza por FK numérica `idResponsable` con columna nueva poblada retroactivamente.
- [ ] D102 aplicado: las 5 columnas booleanas `Text(2)` migradas a `BOOLEAN` con regla explícita.
- [ ] D83 respetado: HPS_Solicitudes (ID 22) **NO comparte tabla** con HPS (ID 17); la FK conceptual `IDUsuarioHPS` se mantiene como referencia conceptual mediate adaptador.
- [ ] Las 5 FKs conceptuales cross-app (`IDExpediente`, `IDUsuarioHPS`, `IDEmpresaUsuario`, `IDEmpresaTramitadora`, `idjustificacion` legacy) se mantienen como referencia conceptual mediate adaptadores.
- [ ] Migración de datos validada con backfill contra staging + conteos contra `Solicitudes_HPS_datos.accdb` autoritativo (~2885 filas totales).
- [ ] `Copia de TbExpedientes` reside en zona `legacy` del esquema con retención indefinida (post-audit TK-HPS-SOL-22 confirma).
- [ ] **CA-Audit-Pendiente**: el **schema usage audit** de las 11 tablas backend está committed (gate previo a fase SDD, TK-HPS-SOL-22 cerrado antes de cualquier ticket de implementación).

### Testing y calidad

- [ ] Los 2 archivos `Test_*.bas` tienen equivalente pytest que valida el mismo comportamiento (D87).
- [ ] Cobertura ampliada con pytest + httpx + Playwright para flujos críticos (alta, renovación, traspaso a ONS, adjuntos Excel) — la cobertura VBA es **básica**, se debe enriquecer.
- [ ] Tests E2E para ciclo de solicitud completo (alta, edición, baja, renovación, cambio tipo, traspaso a ONS).
- [ ] Tests de capabilities verifican rechazo de operaciones no autorizadas.
- [ ] Tests del workflow declarativo verifican que las 20 fechas regulatorias se calculan correctamente desde la config.
- [ ] Tests del parser server-side del Excel verifican validación + enmascaramiento de PII extraída.
- [ ] Smoke contra staging antes de promover a producción.
- [ ] UAT por admins de aplicación + usuarios de negocio (D47-D48); resultado y excepciones documentadas (D49-D50).

---

## 6. Pendientes operacionales

Acciones manuales que el equipo debe ejecutar antes, durante o después de la migración.

### Antes de empezar

- [ ] **PRIMERO — Ejecutar el schema usage audit** (TK-HPS-SOL-22). Gate previo a fase SDD. Sin este audit cerrado, ningún ticket de implementación puede arrancar. Mismo flujo que HPS obs #24097 / commit `8d5311c`.
- [ ] **Operativo en `HPS_SOLICITUDES`** (no desde aquí): añadir `*.accdb` al `.gitignore` del repo.
- [ ] **Operativo en `HPS_SOLICITUDES`** (no desde aquí): verificar `git log --all -- Solicitudes_HPS.accdb`; si está versionado, **rotar repo + aplicar git-filter-repo** para limpiar el historial.
- [ ] **Auditar `PasswordBackend` en `HPS_SOLICITUDES`** (cross-cutting D93/D104) y rotar la contraseña real del backend si era expuesta.
- [ ] **Mover la contraseña a secret manager** (D9-D10). Actualizar manifests de despliegue y dotenv si aplica.
- [ ] **Auditar git history de `HPS_SOLICITUDES`** por commits que contengan `PasswordBackend`; documentar si fue commiteada. Si sí, **rotación obligatoria** + nota de incidente.
- [ ] **Respaldar `Solicitudes_HPS_datos.accdb` de producción** antes de la migración de datos (snapshot inmutable).
- [ ] **Validar con Natalia / negocio** la disposición de `Copia de TbExpedientes` (legacy copy) una vez el audit TK-HPS-SOL-22 confirme si es ZOMBIE (retención indefinida como histórico) o si se descarta formalmente.
- [ ] **Validar con Natalia / negocio** la disposición de `TbLogs` (0 filas, D100) — confirmar desuso definitivo antes de descartar.
- [ ] **Auditar contenido de `Correo.cls` y `TbCorreosEnviados`** y revisar políticas de privacidad (D98) antes de migrar al servicio unificado de notificaciones.
- [ ] **Inspeccionar contenido de las plantillas HTML** (`Form_FormPlantillasHTML.cls`) — pueden contener PII embebida (D98); revisar antes de portar a Jinja2.
- [ ] **Validar con negocio la normalización de las 20 fechas regulatorias** en `TbSolicitudesFechas`: ¿se preservan tal cual (1:1) o se consolidan en una tabla de eventos (más limpio)?
- [ ] **Auditar la integración con ONS** (D101) — ¿sigue activa? ¿qué periodicidad? ¿qué campos se sincronizan? Documentar antes de TK-HPS-SOL-5 (traspasos a ONS).

### Durante la migración

- [ ] Validar que las **245 filas de `TbSolicitudes`** migran con sus 8 columnas PII sin transformación (verificación binaria de presencia en PostgreSQL).
- [ ] Validar que las **5 columnas booleanas** se traducen con regla explícita (D102), sin valores fuera de 'Sí'/'No' sin default documentado.
- [ ] Validar FKs intra-app corregidas (D94, D99) — específicamente la FK por email que pasa a `idResponsable` (D99) y la dirección corregida de `idjustificacion`. Cualquier inconsistencia debe corregirse en origen antes de promover.
- [ ] Coordinar con migración de HPS (D83 — apps independientes, no comparten tabla; pero la FK conceptual `IDUsuarioHPS` requiere convención documentada mediate adaptador).
- [ ] Validar el workflow declarativo (D96 análogo) con las 6 columnas de días en `TbConfiguracion` — el workflow regulatorio debe producir las 20 fechas esperadas.
- [ ] Portar `Test.bas` y `TestParametrosParser.bas` a pytest antes del go-live; ampliar cobertura significativamente.
- [ ] Validar parser server-side del Excel — el contenido puede tener PII; verificar enmascaramiento en logs.
- [ ] Verificar el veredicto del schema usage audit (TK-HPS-SOL-22) antes de descartar formalmente cualquier tabla.

### Después del go-live

- [ ] Confirmar que `TbLogsGeneral` NO existe en artefactos web (verificación binaria — debe estar en stack de observabilidad).
- [ ] Confirmar que `TbLogs` NO existe en artefactos web (descartada sin migración).
- [ ] Confirmar que la papelera de adjuntos (D19) funciona end-to-end para los adjuntos Excel y los adjuntos ONS.
- [ ] Verificar que la impersonación (D44) solo la inicia el administrador global.
- [ ] Cerrar el gap del audit: una vez emitido el veredicto (obs análogo a #24097), confirmar que el resultado está documentado en `migration-matrix.md` con la categoría "no-migrate" si aplica.
- [ ] Evaluar encriptación en reposo para datos personales (D98 follow-up) — no bloqueante para v1.
- [ ] Verificar que el workflow declarativo de HPS_Solicitudes (20 fechas) se ejecuta correctamente en producción sin drift.

---

## 7. Tickets derivables (preview — NO crear issues todavía)

> Lista de issues que nacerían de esta épica. **NO** se crean todavía; se trata de un preview para alinear con el equipo y empezar el desglose SDD.

### Núcleo de funcionalidad

- [ ] **TK-HPS-SOL-1**: Implementar CRUD Solicitud (`TbSolicitudes`, 245 filas, 28 columnas) en backend web — entidades Pydantic + use cases + adaptadores de persistencia + UI HTMX. Preservar las 8 columnas con datos personales (D98) sin transformación.
- [ ] **TK-HPS-SOL-2**: Workflow regulatorio con 20 fechas (`TbSolicitudesFechas`, 21 columnas, 1:1 con `TbSolicitudes`) en PostgreSQL — PK por `IDSolicitud` + 20 `TIMESTAMP NULL`. Decisión: preservar 1:1 o normalizar como tabla de eventos.
- [ ] **TK-HPS-SOL-3**: Workflow declarativo desde `TbConfiguracion` (D96 análogo) — motor server-side que lee los 6 días (`DiasParaRecordatorioExcel1/2`, `DiasCancelacionPreMARGA`, `DiasParaRecordatorioRellenoMarga1/2`, `DiasCancelacionMARGA`) en runtime + endpoint "fechas calculadas" + validación de capabilities (D45-D46).
- [ ] **TK-HPS-SOL-4**: Seed de los 3 catálogos (`motivos_hps`, `hps_grado`, `responsables`) en migración inicial — scripts Alembic idempotentes con conteos del backend autoritativo (TBD, 13, 27). `hps_grado` con PK compuesta `(tipo_hps, grado)`.
- [ ] **TK-HPS-SOL-5**: Traspasos a ONS (D101) — adaptador de salida que serializa la solicitud + adjuntos y los envía al servicio externo ONS vía API o cola. `URLAdjuntoEnvioONS` reemplazado por URL firmada del object storage (D16); endpoint "traspaso a ONS" que registra el evento en logs estructurados (D27).
- [ ] **TK-HPS-SOL-6**: Adjuntos Excel del solicitante (`URLAdjunto`) — endpoint de upload con parsing server-side (openpyxl/pandas.read_excel); validación + enmascaramiento de PII extraída (D98).
- [ ] **TK-HPS-SOL-7**: Plantillas HTML / Vistas web embebidas (`Form_FormPlantillasHTML`, `Form_FormWeb`) — traducir a templates Jinja2 server-side (D66) con auto-escape; auditar PII embebida antes de portar (D98).
- [ ] **TK-HPS-SOL-8**: Vinculaciones conceptuales cross-app mediate adaptadores unificados (D9-D10, D83) — `IDExpediente` (Lanzadera), `IDUsuarioHPS` (HPS gestión de usuarios, **app independiente ID 17, NO comparte tabla**), `IDEmpresaUsuario`, `IDEmpresaTramitadora` como referencias conceptuales; flag `RegistroEnHPS` como config del módulo.

### Forma y plataforma

- [ ] **TK-HPS-SOL-9**: Sustituir `TbLogsGeneral` (2058 filas) por stack de observabilidad web-native (Sentry/OTel/Loki, D27) — traducir llamadas VBA `LogGeneral.Registrar` a `logger.info(...)` con contexto estructurado; `TbLogs` (0 filas, D100) descartada sin migración.
- [ ] **TK-HPS-SOL-10**: Configuración de backend → config del módulo hexagonal + secret manager (D9-D10) — `BackendActivo`, `EnPruebas`, `IDAplicacion="22"`, `PasswordBackend`, `CacheHabilitada` como variables de entorno del runner; los 7 flags TempVars (`EnDesarrollo`, `DatosEnLocal`, `EnPruebas`, `ConCorreoCopiaGestor`, `ActivadoCorreoAutomatico`, `RegistroEnHPS`, `ExpedienteUnificado`) traducidos a config del módulo / env vars; `PasswordBackend` desaparece.
- [ ] **TK-HPS-SOL-11**: Traducción MVVM → hexagonal — clases de dominio → entidades Pydantic; `SolicitudOperaciones.cls` / `SolicitudFechasOperaciones.cls` / `ResponsableOperaciones.cls` / etc. → use cases Python con `AsyncSession.begin()` (D66, D82); `RepositorioComun.bas` y repositorios específicos → adaptadores de persistencia detrás del puerto.
- [ ] **TK-HPS-SOL-12**: Correos automáticos (`TbCorreosEnviados`, 21 columnas, 9 filas) — traducir al servicio unificado de notificaciones (D11, D12) canal email v1; preservar FK a `TbSolicitudes` + 21 campos; `Programado` migrado a `BOOLEAN` (D102); revisar plantillas con datos sensibles (D98) antes de portar.
- [ ] **TK-HPS-SOL-13**: Justificaciones (`TbJustificaciones`, 7 filas) — CRUD con FK corregida a `TbSolicitudes.idjustificacion` (D99); dirección de FK normalizada en PostgreSQL.
- [ ] **TK-HPS-SOL-14**: Responsables (`TbResponsables`, 27 filas) — CRUD con FK numérica `TbSolicitudes.idResponsable → TbResponsables.IDResponsable` (D99); poblar `idResponsable` retroactivamente desde `emailResponsable`; mantener `emailResponsable` como columna independiente para búsqueda.
- [ ] **TK-HPS-SOL-15**: Motivos HPS (`MotivoHPS.cls`) — CRUD completo + UI HTMX; el catálogo se preserva como seed.
- [ ] **TK-HPS-SOL-16**: Búsqueda general de solicitudes + entidades/expedientes + usuarios HPS (`Form_FormSolicitudesGestion`, `Form_FormEntidades`, `Form_FormExpedientesBusqueda`, `Form_FormUsuariosHPSBusqueda`) — queries parametrizadas + UI HTMX con refresh manual (D69); integración con módulos Lanzadera y HPS mediate adaptadores.
- [ ] **TK-HPS-SOL-17**: Tareas pendientes del tramitador (`Form_FormTareasTramitadorPendientes`) — worklist con polling HTMX (`hx-trigger="every 30s"`, D69) + endpoint "mis tareas".
- [ ] **TK-HPS-SOL-18**: Automatización (`Automiatizacion.bas` typo + `AutomatizacionRepositorio.bas`) — traducir a jobs del scheduler unificado (D59) con kill switch (D91); identificar qué automatizaciones hay (traspasos, renovaciones, recordatorios) — pendiente de inspección.
- [ ] **TK-HPS-SOL-19**: Informes (`InformesOperaciones.bas`) — generación parametrizable con vista previa (D62-D63); UI HTMX con filtros + ejecución bajo demanda.

### Seguridad y cumplimiento

- [ ] **TK-HPS-SOL-20**: Plan D98 — limpieza operativa del frontend con PII (operativo en `HPS_SOLICITUDES`, no desde aquí) — `*.accdb` al `.gitignore`, `git log --all -- Solicitudes_HPS.accdb`, git-filter-repo si aplica; las 8 columnas PII se preservan en PostgreSQL sin transformación, se enmascaran en logs, capabilities declaran visibilidad por campo (D45); cifrado del frontend o extracción de PII a esquema separado como follow-up.
- [ ] **TK-HPS-SOL-21**: Rotar `PasswordBackend` + mover a secret manager + audit git history (cierre cross-cutting D93/D104) — operativo en `HPS_SOLICITUDES`; sin este ticket, la épica no puede avanzar.
- [ ] **TK-HPS-SOL-22**: ⚠️ **Schema usage audit** — ejecutar `codegraph_explore` sobre `C:\00repos\codigo\HPS_SOLICITUDES` para clasificar las 11 tablas backend (ACTIVE / UNCLEAR / ZOMBIE). Documentar el veredicto como observación engram con topic_key `hps-solicitudes/usage-audit-<fecha>`. **Gate previo a fase SDD**: ningún ticket de implementación puede arrancar hasta que este audit esté committed. Mismo flujo que HPS obs #24097 / commit `8d5311c`. Candidatas probables a ZOMBIE: `Copia de TbExpedientes` (legacy copy), `TbLogs` (0 filas, D100).
- [ ] **TK-HPS-SOL-23**: Booleanos `Text(2)` → `BOOLEAN` con regla explícita (D102 cross-cutting) — las 5 columnas (`TbConfiguracion.AutocancelacionPreMARGA`, `TbConfiguracion.AutocancelacionMARGA`, `TbConfiguracion.CorreosAutomaticos`, `TbSolicitudes.CorreosAutomaticos`, `TbCorreosEnviados.Programado`) con regla `'Sí' → TRUE`, `'No' → FALSE`, `NULL → NULL`, otros → `FALSE DEFAULT`.
- [ ] **TK-HPS-SOL-24**: FKs conceptuales (D94 + D99) — formalizar las 3 FKs intra-app (`TbSolicitudesFechas.IDSolicitud`, nueva `TbSolicitudes.idResponsable`, `TbSolicitudes.idjustificacion` corregida); mantener las 5 FKs cross-app como referencia conceptual mediate adaptadores; HPS_Solicitudes (ID 22) **NO comparte tabla** con HPS (ID 17) — D83 explícito.
- [ ] **TK-HPS-SOL-25**: Disposición de `Copia de TbExpedientes` y `TbLogs` (TK-HPS-SOL-22 + D100) — `Copia de TbExpedientes` a zona `legacy` del esquema con retención indefinida, sin endpoints; `TbLogs` descartada sin migración; decisión con Natalia / negocio.
- [ ] **TK-HPS-SOL-26**: Suplantación solo por administrador global (D44 cross-cutting) — doble identidad visible + audit completa; integración con adaptador de identidad (D9-D10).

### Testing y calidad

- [ ] **TK-HPS-SOL-27**: Tests E2E para ciclo de solicitud completo (alta, edición, baja, renovación, cambio tipo, traspaso a ONS) — pytest + httpx + Playwright para flujos críticos de UI (D68); tests del workflow declarativo con las 20 fechas regulatorias.
- [ ] **TK-HPS-SOL-28**: Preservar cobertura de los 2 archivos `Test_*.bas` como referencia del puerto de testing (D87) — `Test.bas` (genérico), `TestParametrosParser.bas` (específico de `ParametrosParser.cls`). Cobertura VBA **básica** — la migración debe **ampliar** significativamente con tests nuevos, no solo portar.
- [ ] **TK-HPS-SOL-29**: Tests de capabilities verifican rechazo de operaciones no autorizadas (D45) — el servidor rechaza aunque la UI las muestre.
- [ ] **TK-HPS-SOL-30**: Tests del parser server-side del Excel — validación + enmascaramiento de PII extraída (D98); cobertura de columnas sensibles esperadas.
- [ ] **TK-HPS-SOL-31**: Migración de datos + smoke contra staging antes de promover a producción — backfill validado con conteos contra `Solicitudes_HPS_datos.accdb` autoritativo (~2885 filas totales); validación de la FK numérica `idResponsable` poblada retroactivamente; validación de las 5 columnas booleanas traducidas.

---

## Anexo · Tabla de decisiones referenciadas

| Decisión | Tema | Estado | Aplica a HPS_Solicitudes |
|---|---|---|---|
| D8 | Hexagonal global | APROBADO | Toda la épica |
| D9-D10 | Autenticación via adaptador | APROBADO | Identidad/permisos (D-Identidad) |
| D11-D13 | Notificación unificada + cola email | APROBADO | `TbCorreosEnviados` + plantillas → email v1 (TK-HPS-SOL-12) |
| D14 | PostgreSQL compartido por esquemas | APROBADO | Persistencia del módulo |
| D16-D20 | Adjuntos object storage + papelera 30 días | APROBADO | Adjuntos Excel (TK-HPS-SOL-6) + adjuntos ONS (TK-HPS-SOL-5) |
| D27-D29 | Logs estructurados + retención | APROBADO | Sustituye `TbLogsGeneral` (TK-HPS-SOL-9) |
| D36-D41 | Credenciales, lockout, caducidad | APROBADO | Integración con Lanzadera (no del módulo) |
| D42-D44 | Autorización + suplantación | APROBADO | Capa de auth + TK-HPS-SOL-26 |
| D45-D46 | Capabilities declaradas | APROBADO | Visibilidad por campo (TK-HPS-SOL-20) |
| D66-D67 | Stack backend + frontend | APROBADO | Toda la épica |
| D68 | Monolito modular | APROBADO | Estructura del módulo |
| D69 | Polling HTMX 30s + refresh manual | APROBADO | Worklist tramitador (TK-HPS-SOL-17) |
| D70-D72 | Caché selectiva + ETag + polling | APROBADO | Tareas pendientes, ETag en fragmentos |
| D78-D82 | Versionado + trunk-based + Expand and Contract | APROBADO | Releases y migraciones |
| D83 | HPS_Solicitudes app independiente (ID 22 vs HPS ID 17) | APROBADO | TK-HPS-SOL-8 (`IDUsuarioHPS` referencia conceptual, NO tabla compartida) — explícito en cada hallazgo |
| D86-D87 | Forma hexagonal + tests VBA | APROBADO | Referencia de mapeo del puerto (TK-HPS-SOL-28 — cobertura básica) |
| D91 | Caché selectivo maduro cross-app | APROBADO | Kill switch para automatización (TK-HPS-SOL-18) |
| D93 | Password hardcoded fallback | PROPUESTO (cross-cutting) | TK-HPS-SOL-21 (cierre cross-cutting) |
| D94 | FKs conceptuales intra-app | PROPUESTO (formalización) | TK-HPS-SOL-24 (3 FKs intra-app) |
| D96 | Workflow declarativo `tbTransiciones` (Condor) | APROBADO | TK-HPS-SOL-3 (workflow regulatorio análogo con 20 fechas) |
| D98 | Datos personales en `TbSolicitudes` (245 filas) | PROPUESTO (cierre en esta épica) | H1, D-PII, TK-HPS-SOL-20 |
| D99 | FK por email en `TbResponsables.Correo` | PROPUESTO (cierre en esta épica) | H2, D-FK-email, TK-HPS-SOL-14, TK-HPS-SOL-24 |
| D100 | `TbLogs` vacía — presumible desuso | PROPUESTO | H4, TK-HPS-SOL-9, TK-HPS-SOL-25 |
| D101 | Integración con ONS | PROPUESTO (cierre en esta épica) | H5, TK-HPS-SOL-5 |
| D102 | Booleanos `Text(2)` cross-cutting | PROPUESTO | H6, TK-HPS-SOL-23 |

---

## Anexo · Tabla de fuentes

| Fuente | Aporta |
|---|---|
| `docs/03-aplicaciones/hps-solicitudes/capabilities.md` | Inventario de features F1-F19, 26 clases, 14 módulos, 2 Test_*.bas, 7 flags TempVars |
| `docs/03-aplicaciones/hps-solicitudes/data-model.md` | 7/11 schemas detallados (segunda pasada): `TbSolicitudes` (28 col), `TbResponsables` (4), `TbJustificaciones` (4), `TbHPSGrado` (2, PK compuesta), `TbSolicitudesFechas` (21 col = 20 fechas), `TbConfiguracion` (15 col), `TbUltimoCambio` (4), `TbCorreosEnviados` (21). 4 con shape inventario. |
| `docs/03-aplicaciones/hps-solicitudes/forms.md` | Navegación completa, ~30+ forms, call paths críticos (alta, renovación, traspaso ONS, justificación, responsables), 26 clases desglosadas |
| `docs/03-aplicaciones/hps-solicitudes/integrations-automation.md` | Lanzadera/HPS/ONS, 7 flags TempVars, integración ONS (D101), plantillas HTML, adjuntos Excel |
| `docs/03-aplicaciones/hps-solicitudes/migration-matrix.md` | D98-D101 disposiciones específicas (PII, FKs conceptuales con gaps, TbLogs vacía, ONS) |
| `docs/03-aplicaciones/hps-solicitudes/security-rules.md` | Autorización, 2 roles (Administrador/Técnico), PII crítica (D98), FK por email (D99), traspasos ONS (D101) |
| `docs/03-aplicaciones/hps-solicitudes/README.md` | Estado del lote 7, 12 hallazgos críticos, 89 callers de getdb(), APAP excluido |
| `docs/03-aplicaciones/hps/epic.md` | Template estructural (commit `8d5311c`) — referencia de patrón |
| `docs/03-aplicaciones/condor/epic.md` | Template original (commit `865ef24`) — patrón base |
| `docs/08-decisiones-y-preguntas-abiertas.md` | Decisiones D5-D87 cross-cutting con detalle (D83 APROBADO aplica explícito) |
| `docs/09-arquitectura-objetivo-y-principios.md` | Decisiones D66-D82 arquitectura consolidada |
| engram obs #24023 | P1 resuelta: HPS_Solicitudes es app independiente (ID 22) vs HPS (ID 17) |
| engram obs #24081 | Schemas detalle de HPS_Solicitudes (segunda pasada, 7 tablas) |
| engram obs #24085 | Audit de uso de Condor (15 tablas) — referencia del patrón aplicado a HPS |
| engram obs #24097 | Audit de uso de las 22 tablas backend de HPS (12 ACTIVE / 1 UNCLEAR / 9 ZOMBIE) — referencia canónica del patrón a aplicar a HPS_Solicitudes |

---

## Checklist del documento

- [x] Estado del descubrimiento sincronizado con obs #24081 (schemas detalle).
- [x] Hallazgos D98, D99, D83, D100, D101, D102 con anchor links.
- [x] Decisión D-PII (D98) citada con referencia a `migration-matrix.md § D98`.
- [x] Decisión D-FK-email (D99) citada con referencia a `migration-matrix.md § D99`.
- [x] 19 features de negocio documentadas (F1-F19).
- [x] 11 tablas backend identificadas explícitamente con disposición por categoría.
- [x] 3 catálogos seed-only identificados.
- [x] Criterios de aceptación verificables y agrupados por dimensión.
- [x] Tickets cubren 19 features + seguridad + testing (TK-HPS-SOL-1..31).
- [x] **⚠️ Schema usage audit PENDIENTE** marcado explícitamente en § 1.4, § 2.1.1, § 5 (CA-Audit-Pendiente), § 6 (TK-HPS-SOL-22 como primer pendiente), § 7 (TK-HPS-SOL-22 como ticket primario gate).
- [x] HPS_Solicitudes (ID 22) marcada como app **independiente** (D83) en todos los lugares relevantes (header, H3, D-Independencia-HPS, TK-HPS-SOL-8, TK-HPS-SOL-24, criterio de aceptación).
- [x] Tabla de decisiones referenciadas (D8-D102).
- [x] Tabla de fuentes.
- [x] Idioma: español técnico neutro. Identificadores y paths sin traducir.
- [x] HPS_Solicitudes es app más pequeña (11 tablas vs 22 de HPS, 15 de Condor) — apunta a ~300-400 líneas (actual: dentro de rango).

## Siguiente paso

Revisión con el equipo. El schema usage audit de HPS_Solicitudes es el **primer paso obligatorio** antes de fase SDD: ejecutar **TK-HPS-SOL-22** (análogo a obs #24097 / commit `8d5311c` de HPS) para clasificar las 11 tablas backend (ACTIVE / UNCLEAR / ZOMBIE) y emitir el veredicto como observación engram con topic_key `hps-solicitudes/usage-audit-<fecha>`. Una vez cerrado el audit, aplicar el veredicto al epic (similar al commit `8d5311c` que actualizó el epic de HPS) y resolver con Natalia / negocio la disposición de las ZOMBIE probables (`Copia de TbExpedientes`, `TbLogs`). Tras la disposición, abrir SDD (`sdd-propose` + `sdd-spec` + `sdd-design` + `sdd-tasks`) para arrancar la implementación por ticket, comenzando por **TK-HPS-SOL-21** (limpieza operativa de `PasswordBackend` cross-cutting D93/D104 — gate previo de seguridad) y **TK-HPS-SOL-20** (limpieza operativa del caché frontend con PII en `HPS_SOLICITUDES`, gate D98).