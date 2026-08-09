[← Back to DOCS](../DOCS.md) · [← Codebase Guide](../CODEBASE-GUIDE.md) · [← README](../README.md)

# Mapa de dependencias entre las 8 apps

> **Estado:** DRAFT v0.1 (2026-08-07) — generado tras refactor de las 8 épicas.
> **Propósito:** consolidar el mapa de dependencias cross-app que estaba disperso en las 8 épicas individuales. Guía el orden de implementación cuando cada app se construya.

> **Sentence that organizes**: **Lanzadera es la madre (define usuarios, aplicativos y permisos) y todas las apps dependen de ella. Expedientes es la fuente documental canónica — la práctica totalidad de las apps la referencian para entidades de procedimiento. La interdependencia NoConformidades ↔ Gestion_Riesgos y HPS ↔ HPS_Solicitudes forman pares acoplados que requieren migración coordinada.**

> **Scope del scope**: "Este repo es research + planning de la migración. Cada app tendrá su propio repo + docs cuando se construya."

---

## Quick Navigation

| Section                                                       | What you'll find                                              |
|---------------------------------------------------------------|---------------------------------------------------------------|
| [1. Matriz de dependencias](#1-matriz-de-dependencias)         | Tabla 8×8 con tipo de dependencia entre cada par de apps.    |
| [2. Tipos de dependencias](#2-tipos-de-dependencias)            | Identidad, FK conceptual, HTTP/JSON handshake, catálogo, etc. |
| [3. Chokepoints](#3-chokepoints)                                | Apps cuya ausencia bloquea a otras (Lanzadera, Expedientes).    |
| [4. Orden de migración recomendado](#4-orden-de-migración-recomendado) | 4 olas basadas en dependencias.                                |
| [5. Decisiones arquitecturales](#5-decisiones-arquitecturales) | Adaptador unificado, OAuth, secret manager, Expand and Contract. |
| [6. Sistemas externos](#6-sistemas-externos)                      | SICA, ONS, AGEDYS, AGEDO, TfE — fuera de las 8 apps.         |
| [7. Patrones de fallo](#7-patrones-de-fallo)                      | Cascada de fallos si una app no migra.                         |
| [Anexo · Tabla de fuentes](#anexo--tabla-de-fuentes)         | Las 8 épicas referenciadas.                                     |

---

## 1. Matriz de dependencias (basada en evidencia ERD)

> Leyenda: ✅ = dependencia confirmada por FK en data-model.md; ⚠️ = dependencia parcial / condicional; 🔗 = acoplamiento fuerte (migración coordinada); — = sin dependencia.

| App (consume ↓ / produce → ) | **Lanzadera** | **Expedientes** | **NoConformidades** | **Gestion_Riesgos** | **HPS** | **HPS_Solicitudes** | **Brass** | **Condor** |
|---|---|---|---|---|---|---|---|---|
| **Lanzadera** (la madre) | — | — | — | — | — | — | — | — |
| **Expedientes** (fuente documental) | 🔗 identidad via `Form_Formulario1` colisión (D167) | — | — | — | — | — | — | — |
| **NoConformidades** | 🔗 `getdbLanzadera()` (D86) | 🔗 `Form_FormExpedientesBusqueda` (D132) | — | 🔗 `TbRiesgosNC.IDNC` (FK lógica) | — | — | 🔗 `Form_FormMotivosNoRequiereControlEficacia` (catálogo compartido) | 🔗 `idNCAsociada` (FK conceptual) |
| **Gestion_Riesgos** | 🔗 `getdbLanzadera()` (D86) | 🔗 `Form_FormExpedientesBusqueda` (D132) | 🔗 `TbRiesgosNC.IDNC` (FK lógica, migración coordinada) | — | — | — | ⚠️ catálogos compartidos presumidos | — |
| **HPS** | 🔗 `getdbLanzadera()` (D86) | ⚠️ FK conceptual `IDExpediente` | — | — | — | 🔗 `IDUsuarioHPS` (D83, app independiente ID 17 vs 22) | — | — |
| **HPS_Solicitudes** | 🔗 `getdbLanzadera()` (D86) | ⚠️ FK conceptual `IDExpediente` | — | — | 🔗 `IDUsuarioHPS` + `RegistroEnHPS="Sí"` (D83) | — | — | — |
| **Brass** | 🔗 `getdbLanzadera()` (D86) | — | 🔗 eventos vinculados | ⚠️ catálogos compartidos presumidos | — | — | — | — |
| **Condor** | 🔗 `getdbLanzadera()` (D86) | ⚠️ FK conceptual `idExpediente` | 🔗 `idNCAsociada` (D95) | — | — | — | — | — |

**Resumen de la matriz**:
- **TODAS (7/7) las apps dependen de Lanzadera** (identidad, permisos).
- **5/7 apps dependen de Expedientes**: Gestion_Riesgos, NoConformidades, HPS, HPS_Solicitudes, Condor.
- **3/7 apps dependen de NoConformidades**: Gestion_Riesgos, Brass, Condor.
- **2/7 apps dependen de Gestion_Riesgos**: NoConformidades, Brass.
- **2/7 apps son interdependientes entre sí**: NoConformidades ↔ Gestion_Riesgos, HPS ↔ HPS_Solicitudes.

---

## 1.1 Evidencia ERD (tablas compartidas)

Análisis de `data-model.md` de las 8 apps (búsqueda de tablas referenciadas):

| Tabla compartida | Lanzadera | Expedientes | NoConformidades | Gestion_Riesgos | HPS | HPS_Solicitudes | Brass | Condor | Apps que la referencian |
|---|---|---|---|---|---|---|---|---|---|
| **`TbUsuarioAplicaciones`** (28 cols, 622 filas) | ✅ dueña (2) | — | FK conceptual | FK conceptual | FK conceptual | FK conceptual | — | FK conceptual | **5/8 apps** la referencian (Lanzadera dueña + 4 FKs conceptuales) |
| **`TbExpedientes`** (procedimientos documentales) | FK conceptual | ✅ dueña (48) | ✅ (2) | ✅ (3 en epic) | ✅ (2) | ✅ (6) | — | ✅ (2) | **6/8 apps** la referencian (Expedientes dueña + 5 FKs conceptuales). Brass NO la menciona. |
| **`TbUsuarios`** (usuarios con PII) | FK conceptual | — | ✅ (1) | ✅ (1) | ✅ dueña (23) | ✅ (3) | — | — | **4/8 apps** |
| **`TbAplicaciones`** (catálogo apps, 22 cols) | ✅ dueña (21) | — | — | — | — | — | — | — | **1/8 app** (Lanzadera) |
| **`TbNoConformidades`** (NC cross-app) | — | — | ✅ dueña (23) | FK lógica (1) | — | — | — | FK conceptual (1) | **3/8 apps** |
| **`TbRiesgos`** (riesgos GR) | — | — | FK conceptual (1) | ✅ dueña (4) | — | — | — | — | **2/8 apps** |
| **`TbConfiguracion`** (config local, runtime flags) | ✅ (2) | — | ✅ (4) | — | — | ✅ (6) | — | ✅ (1) | **4/8 apps** |
| **`TbSolicitudes`** (Solicitudes HPS) | — | — | — | — | ✅ (1) | ✅ dueña (25) | — | — | **2/8 apps** |
| **`TbLogs`** / **`TbLogsGeneral`** (logs locales) | — | — | — | — | — | ✅ (12) / ✅ (7) | — | ✅ (5) / ✅ (3) | **2/8 apps** (HPS_Solicitudes + Condor) |

**Hallazgos del análisis ERD**:

1. **`TbUsuarioAplicaciones` está en 5/8 apps** (Lanzadera dueña + 4 FKs conceptuales). El usuario dijo "todas" — la epic la referencia en 5 de 8 (Lanzadera + NoConformidades + Gestion_Riesgos + HPS + HPS_Solicitudes + Condor). Brass NO la referencia explícitamente.
2. **`TbExpedientes` está en 6/8 apps** (Expedientes dueña + 5 FKs conceptuales). El usuario dijo "casi todas" — confirmado. **Brass NO la referencia**. La única app que NO depende de Expedientes es Brass.
3. **`TbConfiguracion` está en 4/8 apps** — patrón común. Indica que cada app tiene su propio runtime config (no compartido entre apps).
4. **`TbUsuarios` está en 4/8 apps** — patrón de identidad. FKs conceptuales desde NoConformidades, Gestion_Riesgos, HPS_Solicitudes. HPS la dueña (23 menciones = schema detallado).
5. **`TbNoConformidades` está en 3/8 apps** — interdependencia NC ↔ Gestion_Riesgos + FK conceptual en Condor.
6. **`TbRiesgos` está en 2/8 apps** — interdependencia NC ↔ Gestion_Riesgos (FK lógica `TbRiesgosNC.IDNC`).
7. **`TbAplicaciones` está solo en Lanzadera (la dueña)** — confirma que es la única app con su propio catálogo. Las demás apps NO tienen catálogo local de apps; lo consumen vía el adaptador unificado (D9-D10).
8. **Brass es la app más aislada** — NO referencia `TbExpedientes`, `TbUsuarioAplicaciones`, `TbUsuarios`, `TbAplicaciones`, `TbNoConformidades`, `TbRiesgos`. Solo referencia `TbEventos`, `TbMaterial`, `TbCodActividad`, `TbEquipos`, `TbEquiposMedida`, `TbSubcontrataciones` (todas propias de Brass).

**Implicación**: el patrón "Expedientes como fuente documental" se confirma con la evidencia ERD. La app que NO consume Expedientes es **Brass** (aislada de las dependencias cross-app documentales).

---

## 2. Tipos de dependencias

| Tipo | Patrón | Apps involucradas | Implementación web |
|---|---|---|---|
| **Identidad** (D9-D10, D86) | `getdbLanzadera()` (legacy) → adaptador unificado | TODAS (7/7) | API REST con OAuth 2.0/SSO. `IdentityService.get_user_with_permissions(user_id, app_id)`. `IDAplicacion` config del módulo. |
| **Búsqueda cross-app** (D132) | `Form_FormExpedientesBusqueda` (HTTP/JSON handshake) | Gestion_Riesgos, NoConformidades | API REST con paginación. `SearchService.search_expedientes(query)`. |
| **FK conceptual** | `IDExpediente` en HPS/HPS_Solicitudes/Condor, `IDSolicitud` en HPS_Solicitudes, `IDUsuarioHPS` en HPS_Solicitudes, `IDNCAsociada` en NoConformidades | 5 apps con FKs conceptuales | Referencia conceptual mediate adaptador. NO formalizar como FK física en PostgreSQL — mantener el patrón del legacy. |
| **Catálogo compartido** | `Form_FormMotivosNoRequiereControlEficacia` (Lanzadera + Brass + NoConformidades) | Lanzadera, Brass, NoConformidades | Endpoint único `/motivos-no-ce` consumido por ambos workflows. |
| **Vinculación inter-app fuerte** | `TbRiesgosNC.IDNC` (NoConformidades ↔ Gestion_Riesgos) | NoConformidades, Gestion_Riesgos | FK lógica con migración coordinada (orden estricto: NoConformidades primero, luego Condor/GR). |
| **Acoplamiento UI** | `Form_Formulario1` colisión (D167) | Lanzadera, Expedientes | Namespace explícito: `LanzaderaFormulario1` / `ExpedientesFormulario1`. **RESUELTO en refactor de Lanzadera (PR #12)**, pendiente en refactor de Expedientes. |
| **Búsqueda en vivo** | Polling de Expedientes (legacy) | NoConformidades, HPS_Solicitudes | SSE/WebSocket server-push. |

---

## 3. Chokepoints

Apps cuya ausencia bloquea a otras (cuello de botella para la migración).

### 3.1 Lanzadera (D86) — chokepoint CRÍTICO
- **Impacto**: 7/7 apps dependen de Lanzadera para identidad/permisos.
- **Si Lanzadera NO migra primero**: ninguna app puede autenticar. Todas bloqueadas.
- **Acción**: **OLA 1 (GATE)**. Lanzadera es la primera app en construirse.

### 3.2 Expedientes — chokepoint documental
- **Impacto**: 5/7 apps dependen de Expedientes (Gestion_Riesgos, NoConformidades, HPS, HPS_Solicitudes, Condor).
- **Si Expedientes NO migra**: esas 5 apps no pueden referenciar procedimientos/expedientes. Bloqueadas.
- **Acción**: **OLA 1 o 2**. Lanzadera primero, luego Expedientes (o en paralelo si los contratos están claros).

### 3.3 Adaptador unificado (D9-D10) — chokepoint técnico
- **Impacto**: TODAS las apps lo necesitan para identidad.
- **Si NO se implementa**: cada app tendría que reimplementar autenticación contra Lanzadera localmente.
- **Acción**: **Implementar el adaptador unificado como shared library** consumida por todas las apps. Esto es parte de la OLA 1 (Lanzadera) pero con alcance cross-app.

### 3.4 Contratos de integración — chokepoint de coordinación
- **Impacto**: cada par de apps con dependencia requiere contrato explícito (endpoint, schema, versionado).
- **Acción**: definir los contratos en la fase de SDD antes de implementar cada app. Sin contratos, las apps se acoplan implícitamente y luego la migración es dolorosa.

---

## 4. Orden de migración recomendado

Basado en las dependencias, propongo 4 olas:

### OLA 1 (GATE) — Lanzadera + adaptador unificado
- **Lanzadera** (la madre) — sin ella, ninguna app puede autenticarse.
- **Adaptador unificado** (D9-D10) — compartido por todas las apps.
- **Expedientes** (opcional, puede ser OLA 1 o 2) — si se hace en OLA 1, destranca 5 apps.

### OLA 2 — Apps con dependencia fuerte de Expedientes
- **Gestion_Riesgos** — depende de Lanzadera + Expedientes (F54) + NoConformidades (FK lógica). **CRÍTICO**: si Gestion_Riesgos se construye antes que NoConformidades, se usa FK lógica con validación.
- **NoConformidades** — depende de Lanzadera + Expedientes + Brass (catálogo motivos). **CRÍTICO**: si NoConformidades se construye antes que Condor/GR, se respeta el orden estricto (D95).
- **Condor** — depende de Lanzadera + Expedientes + NoConformidades. **CRÍTICO**: NoConformidades debe migrarse antes.

### OLA 3 — Apps con dependencia inter-app
- **HPS** — depende de Lanzadera + Expedientes + HPS_Solicitudes. **CRÍTICO**: HPS_Solicitudes depende del flag `RegistroEnHPS="Sí"` de HPS, así que HPS primero o coordinado.
- **HPS_Solicitudes** — depende de Lanzadera + HPS + Expedientes.

### OLA 4 — Apps con menor dependencia externa
- **Brass** — depende de Lanzadera + NoConformidades + Gestion_Riesgos (catálogos).

### Resumen de orden

| OLA | Apps | Razón |
|---|---|---|
| 1 (GATE) | Lanzadera + Adaptador unificado | Chokepoint crítico (identidad) |
| 2 | Expedientes, Gestion_Riesgos, NoConformidades, Condor | Dependen de Expedientes + Lanzadera |
| 3 | HPS, HPS_Solicitudes | Interdependientes (D83 apps independientes) |
| 4 | Brass | Catálogos compartidos, menor dependencia externa |

**Notas**:
- El orden es **recomendado** pero no estricto. Apps con dependencias de OLA anterior pueden esperar a que esa app esté lista.
- `Gestion_Riesgos` ↔ `NoConformidades` requieren **migración coordinada** (FK lógica `TbRiesgosNC.IDNC`).

---

## 5. Decisiones arquitecturales

| ID | Decisión | Apps afectadas | Estado |
|---|---|---|---|
| D9-D10 | Secret manager + adaptador unificado | TODAS | APLICAR (imprescindible) |
| D86-D87 | FK conceptual cross-app mediate adaptador | 5 apps con FKs conceptuales | APLICAR (mantener patrón legacy) |
| D132 | API REST + OAuth 2.0/SSO para handshake | Gestion_Riesgos, NoConformidades, HPS, HPS_Solicitudes | APLICAR |
| D167 | `Form_Formulario1` colisión Lanzadera/Expedientes — namespace explícito | Lanzadera, Expedientes | **RESUELTO en Lanzadera (PR #12)**, pendiente en Expedientes |
| D83 | HPS (ID 17) y HPS_Solicitudes (ID 22) son apps independientes | HPS, HPS_Solicitudes | APLICAR (no comparten tablas) |
| D94 | FKs conceptuales intra-app en Condor → formales en PostgreSQL | Condor | APLICAR |
| D95 | Orden de migración NoConformidades → Condor (FK lógica) | NoConformidades, Condor | APLICAR (orden estricto) |
| **D177 (new)** | **Migración coordinada NoConformidades ↔ Gestion_Riesgos** (FK lógica `TbRiesgosNC.IDNC`) | NoConformidades, Gestion_Riesgos | APLICAR |

---

## 6. Sistemas externos

Sistemas fuera de las 8 apps que requieren integración (vía adaptadores).

| Sistema | App integradora | Tipo de integración | Estado |
|---|---|---|---|
| **SICA** (sistema de gestión de usuarios externos) | HPS | Vía `TbUsuariosSICA` (referencia externa, sin propiedad) | APLICAR — bearer token, retry, audit |
| **ONS** (Organismo Notificador de Seguridad) | HPS_Solicitudes | Vía `URLAdjuntoEnvioONS` (D101) | APLICAR — bearer token, retry, audit |
| **AGEDYS** | Condor | Vía código compartido (presumido) | PENDIENTE investigar |
| **AGEDO** (registro jurídico TdE) | Gestion_Riesgos | Vía `Edicion.RegistrarEnAGEDO` (D129) | PROPUESTO — API REST o SFTP o manual |
| **TfE** (Telefónica de España) | Gestion_Riesgos | Vía `TbProyectos.Juridica="TdE"` | APLICAR — config del módulo |

---

## 7. Patrones de fallo

Cascadas de fallo si una app no migra en el orden recomendado.

| Si falta... | Falla en cascada |
|---|---|
| **Lanzadera** (OLA 1) | **TODAS las apps bloqueadas** (sin identidad). Migración imposible. |
| **Adaptador unificado** (D9-D10) | Cada app reimplementa autenticación localmente → acoplamiento, code duplication. |
| **Expedientes** (OLA 1 o 2) | 5 apps sin procedimiento de referencia. **Buscan otra fuente para `IDExpediente` o se rediseñan**. |
| **NoConformidades antes que Condor** (D95) | Condor no puede validar FK lógica a `idNCAsociada`. **Orden estricto**. |
| **NoConformidades antes que Gestion_Riesgos** (D177 new) | Gestion_Riesgos no puede vincular NCs en `FormRiesgoNC`. **FK lógica rota**. |
| **HPS antes que HPS_Solicitudes** (D83) | HPS_Solicitudes pierde el flag `RegistroEnHPS="Sí"`. **Migración desincronizada**. |
| **Contratos de integración** | Acoplamiento implícito. Refactor doloroso post-implementación. |

---

## Anexo · Tabla de fuentes

| Fuente | Aporta |
|---|---|
| [`condor/epic.md`](03-aplicaciones/condor/epic.md) | Dependencias cross-app: Lanzadera (D86), Expedientes (FK `idExpediente`), NoConformidades (FK `idNCAsociada`, D95), AGEDYS |
| [`brass/epic.md`](03-aplicaciones/brass/epic.md) | Dependencias cross-app: Lanzadera (D86), NoConformidades (eventos), Gestion_Riesgos (catálogos) |
| [`gestion-riesgos/epic.md`](03-aplicaciones/gestion-riesgos/epic.md) | Dependencias cross-app: Expedientes (F54), NoConformidades (F20, `TbRiesgosNC.IDNC`), Lanzadera (D86). Hallazgo D88 (372 getdb() callers). |
| [`no-conformidades/epic.md`](03-aplicaciones/no-conformidades/epic.md) | Dependencias cross-app: Lanzadera (D86), Expedientes (F8), Gestion_Riesgos (F7), Brass (F6). Hallazgos D146-D154. |
| [`expedientes/epic.md`](03-aplicaciones/expedientes/epic.md) | Dependencias cross-app: HPS (F12), NoConformidades (D144), Gestion_Riesgos (D168). Hallazgos D169-D178. |
| [`hps/epic.md`](03-aplicaciones/hps/epic.md) | Dependencias cross-app: Lanzadera (D86), Expedientes (FK `IDExpediente`), HPS_Solicitudes (FK `IDSolicitud`, D83), SICA (externo), AGEDYS (presumido). Hallazgo D92 PII CRITICAL. |
| [`hps-solicitudes/epic.md`](03-aplicaciones/hps-solicitudes/epic.md) | Dependencias cross-app: Lanzadera (D86), HPS (FK `IDUsuarioHPS`, D83), Expedientes (FK `IDExpediente`), ONS (D101). Hallazgos D98 PII + D99 FK email. |
| [`condor/epic.md`](03-aplicaciones/condor/epic.md) | TBV Condor dependencias |
| [`lanzadera/epic.md`](03-aplicaciones/lanzadera/epic.md) | LA MADRE — genera usuarios, aplicativos, permisos. Hallazgos D155-D167 (D167 RESUELTO en PR #12). |
| [DOCS](../DOCS.md) | Technical reference raíz del blueprint |
| [CODEBASE-GUIDE](../CODEBASE-GUIDE.md) | Para mantenedores del blueprint |

## Checklist del documento

- [x] Matriz de dependencias 8×8 con tipo de dependencia entre cada par
- [x] Tipos de dependencias clasificados (identidad, FK conceptual, HTTP/JSON, catálogo, etc.)
- [x] Chokepoints identificados (Lanzadera, Expedientes, adaptador unificado, contratos)
- [x] Orden de migración recomendado en 4 olas con razones explícitas
- [x] Decisiones arquitecturales aplicables (D9-D10, D86-D87, D132, D167, D83, D94, D95, D177 new)
- [x] Sistemas externos (SICA, ONS, AGEDYS, AGEDO, TfE) con apps integradoras
- [x] Patrones de fallo (cascadas si una app no migra)
- [x] Tabla de fuentes
- [x] Idioma: español técnico neutro. Identificadores y paths sin traducir.
- [x] "Sentence that organizes" presente ("Lanzadera es la madre ... casi todas dependen de Expedientes")
- [x] Cross-references a las 8 épicas
- [x] Sin emojis decorativos

## Siguiente paso

- Aplicar el orden de migración recomendado al construir cada app: **OLA 1 primero** (Lanzadera + adaptador unificado).
- Definir los **contratos de integración** entre cada par de apps (endpoint, schema, versionado) en la fase SDD.
- **Audit usage** de las apps con audit pendiente (HPS_Solicitudes TK-HPS-SOL-22, Lanzaderas sin audit) como gate previo a fase SDD.
- **Refactor de Expedientes** para aplicar el fix D167 simétrico (renombrar `Form_Formulario1` a `ExpedientesFormulario1`).
- Considerar crear un **repositorio shared** con el adaptador unificado (D9-D10) y los modelos de datos compartidos (User, App, Permission).

---

[← Back to DOCS](../DOCS.md) · [← Codebase Guide](../CODEBASE-GUIDE.md) · [← README](../README.md)
