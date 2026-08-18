[← Back to HPS README](README.md) · [← Codebase Guide](../../../CODEBASE-GUIDE.md) · [← DOCS](../../../DOCS.md)

# Épica — HPS (migración a web)

> **Estado:** DRAFT v0.1 (2026-08-06) — pendiente revisión final al cerrar el ciclo de las 8 épicas.
> **Lote:** 4 (junto al resto de las 7 apps restantes).
> **App legacy:** `00_HPS` · frontend `HPS.accdb` (30 MB) + backend `HPST.accdb` (7 MB, autoritativo en `C:\00repos\datos\`).
> **Sentence that organizes**: **HPS es la app de Habilitación Personal de SICA: ciclo de vida de usuarios HPS con renovación, histórico, cursos obligatorios, equivalencias, grado, motivos, observaciones, anexos por 4 fuentes (HPS, Histórico, SICA, Usuario) y 12 tablas locales de caché que NO migran (D91 cross-cutting). Riesgo crítico D92: 344 filas de PII en `TbDatosLocal`.**

> **Scope del scope**: "Este repo es research + planning de la migración. Cada app tendrá su propio repo + docs cuando se construya."

---

## Quick Navigation

| Section                                                       | What you'll find                                              |
|---------------------------------------------------------------|---------------------------------------------------------------|
| [Metadatos](#metadatos)                                       | Scope M, D92 PII, dependencias cross-app.                       |
| [1. Scope](#1-scope)                                          | 16 features F1-F16 + 5 catálogos config seed-only.             |
| [2. Estado del descubrimiento](#2-estado-del-descubrimiento)    | Inventario 22 tablas backend + 12 frontend + audit 12/1/9.    |
| [3. Hallazgos críticos](#3-hallazgos-críticos)                | D92 PII, D94 FKs conceptuales, D102 booleanos Text(2).       |
| [4. Decisiones aplicadas](#4-decisiones-aplicadas)              | UX por pantalla, hexagonal ports, seguridad.                  |
| [5. Criterios de aceptación](#5-criterios-de-aceptación)      | Funcionalidad, Seguridad, Performance, Operacional.            |
| [6. Pendientes operacionales](#6-pendientes-operacionales)    | Antes, durante, go-live.                                       |
| [7. Tickets derivables](#7-tickets-derivables-preview)         | 20 tickets TK-HPS-1..20.                                       |
| [Anexo · Decisiones referenciadas](#anexo--decisiones-referenciadas) | D5-D102.                                             |
| [Anexo · Tabla de fuentes](#anexo--tabla-de-fuentes)         | Walkthrough + docs + source + engram.                          |

---

## Metadatos

| Campo | Valor |
|---|---|
| **Aplicación legacy** | `00_HPS` · frontend `HPS.accdb` (30 MB) + backend `HPST.accdb` (7 MB, autoritativo en `C:\00repos\datos\`) |
| **Tipo de migración** | Legacy Access/VBA → web hexagonal (FastAPI + HTMX) |
| **Scope size** | **M** (22 tablas backend + 12 frontend local, 16+ features, 29 clases, 30 módulos, 4 fuentes de anexos) |
| **Dependencias cross-app** | Lanzadera (identidad/permisos, vía `getdbLanzadera` o tablas compartidas); Expedientes (FK conceptual `IDExpediente`); HPS_Solicitudes (FK conceptual `IDSolicitud` — **app independiente** ID 22, D83); SICA (sistema externo vía `TbUsuariosSICA`); AGEDYS presumible (vía código compartido) |
| **Riesgo dominante** | Seguridad **D92** — caché local en frontend `HPS.accdb` con **344 filas de PII** (DNI, nombres, fechas de nacimiento, correos, observaciones) en `TbDatosLocal` |
| **Stack target** | Backend Python 3.12+ / FastAPI 0.119+ / SQLAlchemy 2.0.x / Alembic 1.13+ / asyncpg 0.30+ (D66) · Frontend HTMX 2.0.4 + Jinja2 3.1+ / Alpine.js 3.15+ (D67) |
| **Estrategia de migración de BD** | Expand and Contract backward-compatible (D82) · PostgreSQL compartido con esquema por módulo (D14) |
| **Forma destino** | Hexagonal global (D8) · módulo dentro del monolito modular (D68) · puerto de persistencia PostgreSQL + object storage S3-compatible (D16) + secret manager (D9-D10) |
| **Auditoría de uso previa** | **committed** (engram obs #24097, topic_key `hps/usage-audit-2026-08-06`): **12 ACTIVE / 1 UNCLEAR / 9 ZOMBIE** de 22 tablas backend. Limitación declarada: lower bound del codegraph no incluye macros embebidas ni queries no exportadas. |

---

## 1. Scope

### 1.1 En scope — 16+ features de negocio con paridad funcional

| # | Feature | Respaldo en docs |
|---|---|---|
| F1 | Arranque e identidad (`EVE` + `getUsuario`) | `capabilities.md` · `forms.md` · `data-model.md` § IDAplicacion |
| F2 | Configuración de backend (`TbConfiguracionBackends`) → config del módulo hexagonal | `capabilities.md` · `integrations-automation.md` § Configuración y flags |
| F3 | CRUD Usuario HPS (`TbUsuarios`, 345 filas) — alta, baja, renovación, modificación de datos personales, cálculo de `APuntoDeCaducar`/`Caducado`/`Solicitado`/`PendienteRenovacion` | `capabilities.md` · `forms.md` · `data-model.md` |
| F4 | Histórico de usuarios (`TbUsuariosHistoricos`, 242 filas) con `TbUsuariosHistoricosLocal` (235 filas en caché) | `capabilities.md` · `data-model.md` · `security-rules.md` § Histórico con anexos |
| F5 | HPS propiamente (`TbHPS`, 1280 filas) + equivalencias (`TbHPSEquivalencia`) + grado (`TbHPSGrado`) | `capabilities.md` · `data-model.md` · `forms.md` |
| F6 | Motivo HPS (`TbMotivoHPS`) — motivos de alta/baja | `capabilities.md` · `migration-matrix.md` |
| F7 | Cursos obligatorios (`TbAuxCursos`) | `capabilities.md` · `migration-matrix.md` |
| F8 | Observaciones (`TbObservaciones`, 330 filas) + históricas (`TbObservacionesHistoricas`) | `capabilities.md` · `forms.md` · `migration-matrix.md` |
| F9 | Anexos 4 fuentes — HPS / Histórico / SICA / Usuario (`TbAnexosUsuariosHPS`, `TbAnexosUsuariosHistoricos`, `TbAnexosUsuariosSICA`, `TbUsuarioAnexos`) | `capabilities.md` · `forms.md` · `migration-matrix.md` |
| F10 | Integración SICA externa (`TbUsuariosSICA`) — referencia externa, sin propiedad | `capabilities.md` · `integrations-automation.md` |
| F11 | Indicadores en tiempo real con kill switch atómico (`Indicador.cls`, `clsIndicadoresBus.cls`, `Test_RealTimeIndicatorCoherence.bas`) | `capabilities.md` · `forms.md` · `security-rules.md` § Indicadores |
| F12 | Transacciones de selección de anexos (`AnexoSelectionTransactionCoordinator`) + lifecycle de usuarios (`UsuarioLifecycleTransactionCoordinator`) | `capabilities.md` · `forms.md` · `security-rules.md` § Transaccionalidad |
| F13 | Consultas pre-armadas — 8 formularios `Form_FormInicial03Consultas<PorCampo>` (Contratistas, Datos, EmpresaUsuario, Estado, Grado, JuridicaTramitacion, MotivoHPS, Nombre, ProyectoAsignado) | `capabilities.md` · `forms.md` § Navegación principal |
| F14 | Consultas SQL externas (`TbConsultas` + `Consulta.cls`) — query builder parametrizable | `capabilities.md` · `migration-matrix.md` |
| F15 | Sincronización histórico ↔ anexos (`Mod_Sincronizacion_Historico.bas` + `Test_HistoricoAdjuntosTransactionWrapper.bas`) | `capabilities.md` · `security-rules.md` § Sincronización |
| F16 | Catálogo de jurídicas (`TbJuridicasContratacion`) | `capabilities.md` · `migration-matrix.md` |

### 1.2 5 catálogos config seed-only (datos de runtime, sin CRUD de usuario)

| Catálogo | Volumen | Notas |
|---|---|---|
| `motivos_hps` (era `TbMotivoHPS`) | TBD | Motivos de alta/baja HPS |
| `cursos` (era `TbAuxCursos`) | TBD | Cursos obligatorios |
| `hps_equivalencia` (era `TbHPSEquivalencia`) | TBD | Equivalencias entre cursos HPS |
| `hps_grado` (era `TbHPSGrado`) | TBD | Grados del HPS |
| `juridicas_contratacion` (era `TbJuridicasContratacion`) | TBD | Jurídicas de contratación |

### 1.3 Volumen real autoritativo (backend `HPST.accdb`)

| Tabla | Filas |
|---|---|
| `TbUsuarios` | 345 (usuarios activos) |
| `TbUsuariosHistoricos` | 242 (usuarios históricos) |
| `TbHPS` | 1280 (relación usuario-curso) |
| `TbObservaciones` | 330 (observaciones) |
| **Total principales** | **≈2197** (más catálogos y anexos sin auditar) |

**Read-heavy**: `getdb()` con **107 callers** (el más bajo de las 8 aplicaciones). La mayoría de las operaciones son consultas sobre el catálogo de usuarios; escrituras limitadas a actualizaciones de HPS, observaciones y carga inicial.

**Caché selectivo maduro** (confirma D91 cross-cutting): 6 módulos (`cacheUsuario.bas`, `cacheSuministrador.bas`, `Mod_Cache_Core.bas`, `Mod_StartupCacheInitialization.bas`, `CacheConsistencyAudit.bas`, `Mod_Sincronizacion_Historico.bas`) controlados vía `TbConfiguracion.CacheHabilitada` (kill switch). Patrón idéntico al de NoConformidades — se preserva como **referencia del puerto de caché** (D70-D71).

**Tests VBA como referencia** (D87): 9 archivos `Test_*.bas` (`Test_AnexoSelectionTransaction`, `Test_CacheConsistencyAudit`, `Test_HistoricoAdjuntosTransactionWrapper`, `Test_HPSConfig`, `Test_HPSEntorno`, `Test_LocalReadAccessAuthorization`, `Test_PerAnexoMove`, `Test_RealTimeIndicatorCoherence`, `Test_StartupCacheInitialization`). Cobertura significativa; **test crítico a portar**: `Test_LocalReadAccessAuthorization.bas` valida que el acceso local a datos sigue reglas de autorización.

**TDD maduro**: 5 archivos `clsTestDouble*` (`clsTestDoubleForm`, `clsTestDoubleHpsEditListener`, `clsTestDoubleHpsEditPublisher`, `clsTestDoubleIndicador`, `clsTestDoubleIndicadorConsumer`) + 6 fixtures (`DummyBig`, `DummyBig2`, `DummyHatw`, `DummyHatw2`, `DummySmall`, `DummySpecial`, `Módulo1`). Indica disciplina TDD madura en el legacy.

**Vinculaciones externas**: la identidad se resuelve vía **adaptador unificado** (D9-D10); las FKs conceptuales a Expedientes/HPS_Solicitudes se mantienen como referencia conceptual mediate adaptadores (D86/D87). **HPS_Solicitudes es app independiente** (ID 22 vs ID 17 de HPS) — no comparten tablas (D83).

### 1.4 Fuera de scope (REPLACE)

> Las **12 tablas locales del frontend** NO migran como caché en cliente. Se reemplazan por stack de caché server-side detrás del puerto (D70-D71). Sus equivalentes de negocio viven en el backend autoritativo; el frontend `.accdb` deja de existir como runtime.

**CRÍTICO — caché local con PII** (D92):
- `TbDatosLocal` (**344 filas**) — contiene TODOS los datos personales de usuarios HPS activos: DNI, Nombre, Apellidos, Teléfono, Correo_e, F_Nacimiento, LugarNacimiento, Observaciones + estado HPS desnormalizado por organismo (NAC, OTAN, ESA, UE).
- `TbUsuariosHistoricosLocal` (235 filas) — datos personales del histórico de usuarios (DNI, nombres, fechas, correo).

**Caché auxiliar** (sensibilidad media-baja):
- `TbCursosLocal` (caché de cursos)
- `TbSuministradoresLocal` (caché de suministradores)
- `TbDatosLocalParaIndicadores` (caché para indicadores)
- `TbUsuariosSICALocalParaIndicadores` (caché SICA para indicadores)
- `TbConfiguracionHPS` (clave-valor local: Clave, Valor, Activo, FechaModificacion, UsuarioModificacion)

**Metadatos locales** (config, no datos):
- `TbVinculosTablas`, `tblInfo`, `tblSettings` — metadatos locales del frontend → migrar como **config del módulo** (no tablas, sino variables de entorno o config centralizada).

**Descartar**:
- `TbUsuariosSICALocal` (0 filas, no se usa) — descartar.

**Excluido del REPLACE**: las **22 tablas del backend** (`HPST.accdb`) migran como tablas PostgreSQL con sus claves y datos preservados. Solo `TbConsultas` requiere decisión sobre query builder parametrizable.

### 1.5 Fuera de scope (no documentado)

- Volúmenes reales de producción (no staging): `HPST.accdb` en `C:\00repos\datos\` no fue inspeccionado en esta pasada.
- Conteos de filas de los catálogos seed-only (`TbAuxCursos`, `TbHPSEquivalencia`, `TbHPSGrado`, `TbMotivoHPS`, `TbJuridicasContratacion`) — solo `TbUsuarios` y `TbObservaciones` principales auditadas.
- `TbConsultas` contenido (queries SQL pre-armadas) — sensible a SQL injection si no se parametriza; decisión pendiente.
- Inventario completo de queries exportadas y macros embebidas — codegraph no captura macros embebidas ni QueryDefs no exportados.
- Plantillas de correo en `Correo.cls` — pueden contener datos sensibles (D92); revisión pendiente antes de portar.

### 1.6 Fuera de scope (no-migrate — zombies confirmados)

> 9 tablas de las 22 backend son zombies confirmados (audit obs #24097). No migran como features a la nueva plataforma.

**4× copias legacy "Copia de..."** (herencia de mass-change pre-consolidación; retención indefinida, sin endpoints, sin UI):
- `Copia de TbExpedienteLugares`
- `Copia de TbExpedientes`
- `Copia de TbUsuarios`
- `Copia de TbUsuariosEntidades`

**Sentinel de errores**:
- `Errores de pegado` — sentinel de errores en pegado masivo. Transformar en log estructurado canónico (D27) o preservar con retención limitada.

**3 catálogos sin uso** (confirmar con negocio antes de descartar formalmente):
- `TbHPSEquivalencia` — catálogo sin uso (0 callers en audit).
- `TbAuxCursos` — catálogo sin uso (0 callers en audit).
- `TbJuridicasContratacion` — catálogo sin uso (0 callers en audit).

**Solapamiento con tabla activa**:
- `TbUsuarioAnexos` — 0 callers; sospecha de solape con `TbAnexosUsuariosHPS`. Confirmar con negocio antes de descartar.

---

## 2. Estado del descubrimiento

### 2.1 Inventario

| Categoría | Resultado |
|---|---|
| **Tablas totales (backend)** | **22** en `HPST.accdb` (autoritativo en `C:\00repos\datos\`) |
| **Tablas totales (frontend local)** | **12** en `HPS.accdb` (caché + metadatos) — **NO migran como tales** |
| **Filas totales (backend, auditadas)** | **≈2197** (4 tablas principales: 345 + 242 + 1280 + 330) |
| **Filas totales (frontend local)** | **≈579** (344 + 235 + catálogos auxiliares) — **NO migran como tales** |
| **Schemas documentados** | **22/22** tablas backend en `data-model.md` (commits HPS del lote 4, ver obs #24084) — **esquema detallado completo de `TbUsuarios` (27 columnas) committed**; resto en nivel inventario |
| **Uso (audit)** | **committed** (engram obs #24097, topic_key `hps/usage-audit-2026-08-06`): **12 ACTIVE / 1 UNCLEAR / 9 ZOMBIE** de 22 tablas backend. Limitación: lower bound del codegraph no incluye macros embebidas ni queries no exportadas. |
| **Subcategorías** | 6 dominio (usuarios, histórico, HPS, observaciones, anexos, indicadores) · 5 catálogos · 4 legacy copies · 1 sentinel · 4 anexos cross-source · 2 config |
| **FKs físicas (backend)** | **6** desde `TbUsuarios` / `TbUsuariosHistoricos` / `TbUsuariosSICA` (3 con PK-to-PK genérico a corregir, ver D94) |
| **FKs conceptuales (sin constraint)** | **5** — `IDExpediente`, `IDEmpresaUsuario`, `IDEmpresaHPS`, `IDJuridicaContrato`, `IDSolicitud` |
| **Clases inventariadas** | 29 (16 dominio, 5 Test Doubles, 2 coordinadores transaccionales, 3 anexos, 2 compartidas con Lanzadera, 2 indicadores) |
| **Módulos** | 30 (4 bootstrap/factory/DAO, 6 caché, 1 indicadores, 1 configuración, 1 filtros, 1 JSON, 7 fixtures, 9 tests VBA) |
| **Forms** | ~30 archivos `Form_*.cls` cada uno con su `.form.txt` compañero |
| **Tests VBA** | 9 archivos `Test_*.bas` (cobertura significativa, confirma D87) |
| **Duplicación frontend/backend** | `TbHPS` y `TbUsuariosHistoricos` existen en ambos `.accdb` (Dysflow emite `ACCESS_TABLE_AMBIGUOUS` al consultarlas sin `target`) |

### 2.2 Veredicto del schema usage audit (obs #24097)

> Audit ejecutado 2026-08-06 vía `codegraph-vba` sobre `00_HPS` READ-ONLY. Veredicto sobre las **22 tablas backend**: **12 ACTIVE / 1 UNCLEAR / 9 ZOMBIE**.

**12 ACTIVE** (callers encontrados):
- `TbUsuarios` (5), `TbUsuariosHistoricos` (5), `TbUsuariosSICA` (3), `TbUsuariosEntidades` (5), `TbHPS` (5), `TbHPSGrado` (1), `TbMotivoHPS` (1), `TbObservaciones` (5), `TbObservacionesHistoricas` (5), `TbAnexosUsuariosHPS` (5), `TbAnexosUsuariosSICA` (2), `TbAnexosUsuariosHistoricos` (5).

**1 UNCLEAR**: `TbConsultas` — gap del extractor (codegraph-vba no emite arista para `m_SQL='TableName'`).

**9 ZOMBIE** (0 callers): 4 copias legacy + 1 sentinel + 3 catálogos + 1 solapamiento (ver § 1.6).

**Limitaciones declaradas**: macros embebidos no exportados; `codegraph_explore` limita blast radius a 5 entries; `TbDatosLocal` (PII) está en frontend, fuera del scope del audit backend.

### 2.3 Hallazgos estructurales del inventario

**4 tablas "Copia de..."**: patrón legacy de copia antes de cambios masivos (D83).

**Sentinel `Errores de pegado`**: tabla para capturar errores en operaciones de pegado masivo.

**Inconsistencias detectadas en FKs** (D94):
1. `TbUsuariosHistoricos.ID → TbObservacionesHistoricas.ID` (PK-to-PK genérico).
2. `TbUsuarios.ID → TbObservaciones.ID` (PK-to-PK genérico).
3. NO hay FK física entre `TbUsuarios` y `TbExpedientes` pese a `IDExpediente`.
4. NO hay FK física entre `TbUsuarios` y `TbEmpresas` pese a `IDEmpresaUsuario`/`IDEmpresaHPS`.
5. NO hay FK física entre `TbUsuarios` y `TbJuridicasContratacion` pese a `IDJuridicaContrato`.
6. NO hay FK física entre `TbUsuarios` y `TbSolicitudes` pese a `IDSolicitud`.

**Booleanos como Text(2)** (D102): 3 columnas en `TbUsuarios` con `VARCHAR(2) NULL` y valores 'Sí'/'No' — `CursoEnVigor`, `Requiere_Curso`, `RequiereComunicacionConcesion`. Inconsistencia detectada también en NoConformidades; regla general cross-cutting.

---

## 3. Hallazgos críticos

| # | ID | Título | Severidad | Forms afectados | Detalle |
|---|---|---|---|---|---|
| H1 | D92 | **PII en `TbDatosLocal` (344 filas) + `TbUsuariosHistoricosLocal` (235 filas)** | CRITICAL | Frontend `HPS.accdb` | DNI, Nombre, Apellidos, Teléfono, Correo, F_Nacimiento, LugarNacimiento, Observaciones. **CRÍTICO**: el frontend NO debe existir como runtime en web — el `HPS.accdb` se descarta. Los datos de PII migran al backend autoritativo con encryption at-rest y access controls. |
| H2 | D94 | **FKs conceptuales sin constraint (5 PK-to-PK genéricos)** | high | `TbUsuariosHistoricos`, `TbObservacionesHistoricas`, `TbUsuarios`, `TbObservaciones` | FKs PK-to-PK genéricos (reutilización de IDs entre tablas). Formalizar en PostgreSQL con FKs reales: `TbObservacionesHistoricas.IDUsuarioHistorico → TbUsuariosHistoricos.ID`. |
| H3 | D102 | **Booleanos como `Text(2)` 'Sí/No'** | low | `TbUsuarios.CursoEnVigor`, `Requiere_Curso`, `RequiereComunicacionConcesion` | 3 columnas con `VARCHAR(2) NULL` y valores 'Sí'/'No'. Migrar a `BOOLEAN` con regla explícita. Cross-cutting. |
| H4 | D83 | **`HPS_Solicitudes` es app independiente (ID 22 vs 17)** | low | Cross-app con HPS | FK conceptual `IDSolicitud` sin constraint. Mantener como referencia conceptual mediate adaptador. |
| H5 | D91 | **Caché selectivo maduro (6 módulos)** | low | `cacheUsuario.bas`, `cacheSuministrador.bas`, `Mod_Cache_Core.bas`, `Mod_StartupCacheInitialization.bas`, `CacheConsistencyAudit.bas`, `Mod_Sincronizacion_Historico.bas` | Patrón de caché local + kill switch `TbConfiguracion.CacheHabilitada`. **Se preserva como referencia del puerto de caché** (D70-D71) en la nueva plataforma. |
| H6 | D87 | **9 Tests VBA** (cobertura significativa) | medium | 9 archivos `Test_*.bas` | Patrón TDD maduro. **Test crítico a portar**: `Test_LocalReadAccessAuthorization.bas` valida acceso local a datos. |
| H7 | **Duplicación frontend/backend** | medium | `TbHPS` y `TbUsuariosHistoricos` | Dysflow emite `ACCESS_TABLE_AMBIGUOUS` al consultarlas sin `target`. Migrar a backend único, sin `.accdb` frontend. |
| H8 | **`TbConsultas` UNCLEAR** | low | `TbConsultas` + `Consulta.cls` | Query builder parametrizable. Gap del extractor. Decisión pendiente sobre parametrizar o deprecar. |
| H9 | **9 ZOMBIE** | low | 4 copias legacy + 1 sentinel + 3 catálogos + 1 solapamiento | No migran como features. Decidir disposición (drop, archive, preserve con retención). |

---

## 4. Decisiones aplicadas

### 4.1 UX: Preservar / Mejorar / Nuevo paradigma

| Pantalla | Decisión | Justificación |
|---|---|---|
| Splash + Arranque (F1) | **Preservar** flujo vía EVE | Adaptador unificado (D9-D10) |
| `Form_FormHPSEdit` (CRUD usuario) | **Preservar** flujo + **Mejorar**: validación inline | 345 filas activas con PII |
| Histórico (`Form_FormHPSHistorico`) | **Preservar** vista read-only | Flujo crítico de auditoría |
| Observaciones (`Form_FormObservaciones`) | **Preservar** flujo | 330 filas activas |
| Anexos (F9) | **Nuevo paradigma**: upload directo a S3 con preview inline | D16 — cambio completo |
| Indicadores (F11) | **Mejorar**: server-side polling | D85 — kill switch atómico |
| Consultas pre-armadas (F13) | **Mejorar**: filtros con chips + URL state | UX genérica |
| Consultas SQL (F14) | **Decidir**: parametrizar o deprecar | SQL injection risk |

### 4.2 Arquitectura: hexagonal ports

| Concepto legacy | Traducción web |
|---|---|
| `getdb()` (DAO.Database) — 107 callers | Puerto de persistencia PostgreSQL (D14) + HTTP client |
| `Constructor.getXxx()` (factory con cache lazy) | Inyección de dependencias + repository pattern |
| `m_ObjEntorno` (singleton global) | Service registry / dependency injection |
| `m_ObjUsuarioConectado` (singleton in scope) | Request-scoped services / Context API |
| `TbConfiguracionBackends` (config del frontend) | Variables de entorno del runner (D9-D10) |
| `EVE` (autenticación legacy) | Adaptador unificado (D9-D10) |
| `Test_LocalReadAccessAuthorization` | Puerto de autorización (D45-D46) |
| `Mod_Cache_Core` + `cacheUsuario` + `cacheSuministrador` | Server-side cache con Redis (D70-D71) |
| `TbConfiguracion.CacheHabilitada` (kill switch) | Feature flag server-side (LaunchDarkly/Unleash) |
| `Indicador.cls` + `clsIndicadoresBus` | Server-side polling + ETag |
| `AnexoSelectionTransactionCoordinator` | Transacciones explícitas en backend (D82 Expand and Contract) |
| `Mod_Sincronizacion_Historico` + `Test_HistoricoAdjuntos` | Sincronización server-side con rollback compensatorio |
| `Consulta.cls` (query builder VBA) | SQLAlchemy 2.0 con typed queries |
| `tbHistorialRechazos` | (no aplica a HPS directamente, pero el patrón cross-cutting sí) |

### 4.3 Decisiones de seguridad

- **PII en `TbDatosLocal` (D92)**: NO migra como caché local. El `HPS.accdb` frontend se descarta. Los datos de PII migran al backend autoritativo con encryption at-rest y access controls estrictos. El frontend web NO tiene copia local de PII.
- **Cross-app con Lanzadera** (identidad): adaptador unificado (D9-D10).
- **Cross-app con Expedientes** (FK conceptual): adaptador.
- **Cross-app con HPS_Solicitudes** (FK conceptual): app independiente (D83) — adaptador.
- **Cross-app con SICA**: sistema externo — referencia, sin propiedad.
- **Test crítico** `Test_LocalReadAccessAuthorization`: portar al backend como test E2E.

### 4.4 Decisiones de datos

- **22 tablas backend** migran a PostgreSQL con claves y datos preservados.
- **12 tablas frontend local** se descartan (caché reemplazado por Redis).
- **Booleanos `Text(2)`** (D102): migrar a `BOOLEAN` con regla explícita.
- **FKs conceptuales** (D94): formalizar en PostgreSQL con FKs reales.
- **Tablas ZOMBIE** (9): decidir disposición caso por caso con negocio.

---

## 5. Criterios de aceptación

### 5.1 Funcionalidad

- [ ] **CA-F1**: Las 16+ features F1-F16 tienen paridad funcional con la versión Access.
- [ ] **CA-F2**: El ciclo de vida del usuario HPS (alta, baja, renovación, modificación) preserva las 4 estados (`APuntoDeCaducar`/`Caducado`/`Solicitado`/`PendienteRenovacion`).
- [ ] **CA-F3**: El histórico de usuarios con anexos preserva la integridad referencial (D94).
- [ ] **CA-F4**: Los 4 fuentes de anexos (HPS, Histórico, SICA, Usuario) migran a un modelo polimórfico unificado.
- [ ] **CA-F5**: La integración SICA permanece como referencia externa (sin propiedad).
- [ ] **CA-F6**: Los indicadores en tiempo real respetan el kill switch (D85).
- [ ] **CA-F7**: Las transacciones de selección de anexos + lifecycle de usuarios preservan la transaccionalidad (D87).
- [ ] **CA-F8**: Las consultas pre-armadas (8 formularios) y las consultas SQL externas (`TbConsultas`) preservan su semántica.
- [ ] **CA-F9**: La sincronización histórico ↔ anexos preserva la consistencia.
- [ ] **CA-F10**: El catálogo de jurídicas (F16) se preserva.

### 5.2 Seguridad

- [ ] **CA-S1**: `TbDatosLocal` (344 filas PII) eliminado del frontend — los datos migran al backend con encryption at-rest.
- [ ] **CA-S2**: `TbUsuariosHistoricosLocal` (235 filas PII) eliminado del frontend.
- [ ] **CA-S3**: `Test_LocalReadAccessAuthorization` portado al backend como test E2E de autorización.
- [ ] **CA-S4**: `TbConfiguracionHPS` migrado a secret manager (no en frontend).
- [ ] **CA-S5**: PII no expuesta via API sin scope explícito.
- [ ] **CA-S6**: Cross-app API con Lanzadera + HPS_Solicitudes + Expedientes vía OAuth 2.0.

### 5.3 Performance

- [ ] **CA-P1**: Catálogo de 345 usuarios carga < 1s con paginación server-side.
- [ ] **CA-P2**: Histórico de 242 usuarios con anexos carga < 2s.
- [ ] **CA-P3**: Indicadores con kill switch desactivado cargan en < 100ms (cached).

### 5.4 Operacional

- [ ] **CA-O1**: 22 tablas backend migradas con datos preservados (~2197 filas).
- [ ] **CA-O2**: 12 tablas frontend local descartadas; equivalentes en Redis.
- [ ] **CA-O3**: Tests E2E portados desde los 9 archivos `Test_*.bas` (especialmente `Test_LocalReadAccessAuthorization`).
- [ ] **CA-O4**: Cobertura de tests > 70% en módulo hps.
- [ ] **CA-O5**: 9 ZOMBIE con disposición explícita (drop, archive, preserve con retención).

---

## 6. Pendientes operacionales

### 6.1 Antes de empezar

- [ ] **PO-1**: Confirmar el contrato de integración con **Lanzadera** (vía `getdbLanzadera()`) — autenticación, endpoint, schema.
- [ ] **PO-2**: Confirmar el contrato de integración con **Expedientes** (FK conceptual `IDExpediente`).
- [ ] **PO-3**: Confirmar el contrato de integración con **HPS_Solicitudes** (app independiente ID 22).
- [ ] **PO-4**: Decidir disposición de las 9 ZOMBIE (drop vs archive vs preserve con retención).
- [ ] **PO-5**: Auditar plantillas de correo en `Correo.cls` (pueden contener datos sensibles — D92).
- [ ] **PO-6**: Decidir `TbConsultas`: parametrizar o deprecar.

### 6.2 Durante la migración

- [ ] **PO-7**: Implementar servicio de hidratación de `TbDatosLocal` desde backend (NO dejar copia en frontend).
- [ ] **PO-8**: Implementar el cache server-side (Redis) con TTL configurable + kill switch (D70-D71, D85).
- [ ] **PO-9**: Migrar booleanos `Text(2)` → `BOOLEAN` (D102 cross-cutting) — 3 columnas en `TbUsuarios`.
- [ ] **PO-10**: Formalizar FKs conceptuales (D94) — 5 PK-to-PK genéricos + 5 FKs conceptuales a apps externas.
- [ ] **PO-11**: Eliminar las 4 tablas "Copia de..." (decidir drop vs archive).
- [ ] **PO-12**: Portar `Test_LocalReadAccessAuthorization` al backend como test E2E de autorización.
- [ ] **PO-13**: Sincronización histórico ↔ anexos server-side con rollback compensatorio (D87).

### 6.3 En el go-live

- [ ] **PO-14**: Smoke test E2E: alta usuario HPS → curso → observación → anexo → histórico.
- [ ] **PO-15**: Verificar que NO queden copias de PII en el frontend web (security audit).
- [ ] **PO-16**: Verificar que el kill switch de caché funciona (D85).
- [ ] **PO-17**: Plan de deprecation de las 9 ZOMBIE confirmadas con negocio.

---

## 7. Tickets derivables (preview)

### Seguridad (CRITICAL — D92)

- **TK-HPS-1**: [SECURITY] **CRITICAL**: Eliminar `TbDatosLocal` (344 filas PII) del frontend — los datos migran al backend con encryption at-rest.
- **TK-HPS-2**: [SECURITY] Eliminar `TbUsuariosHistoricosLocal` (235 filas PII) del frontend.

### Datos y FKs (D94, D102)

- **TK-HPS-3**: [DATA] Formalizar 5 FKs conceptuales (D94): PK-to-PK genéricos + 5 FKs conceptuales a apps externas.
- **TK-HPS-4**: [DATA] Migrar 3 columnas booleanos `Text(2)` → `BOOLEAN` (D102 cross-cutting).
- **TK-HPS-5**: [DATA] Decidir disposición de 9 ZOMBIE (drop, archive, preserve).
- **TK-HPS-6**: [DATA] Eliminar 4 tablas "Copia de..." + sentinel `Errores de pegado` tras confirmar con negocio.

### Forma y plataforma

- **TK-HPS-7**: [REFACTOR] Migrar 29 clases a services HTTP manteniendo firmas.
- **TK-HPS-8**: [MIGRATION] Reemplazar 12 tablas frontend local con Redis server-side cache (D70-D71).
- **TK-HPS-9**: [MIGRATION] Reemplazar `TbConfiguracion.CacheHabilitada` (kill switch) con feature flag server-side.
- **TK-HPS-10**: [INTEGRATION] API REST federada con Lanzadera + HPS_Solicitudes + Expedientes.
- **TK-HPS-11**: [MIGRATION] Reemplazar `Indicador.cls` + `clsIndicadoresBus` con server-side polling + ETag.
- **TK-HPS-12**: [MIGRATION] Reemplazar `AnexoSelectionTransactionCoordinator` con transacciones explícitas en backend (D82).

### Integración

- **TK-HPS-13**: [INTEGRATION] Mantener SICA como referencia externa (sin propiedad).
- **TK-HPS-14**: [INTEGRATION] Mantener HPS_Solicitudes como app independiente (D83).

### Testing y calidad

- **TK-HPS-15**: [TESTING] Portar 9 Tests VBA a E2E pytest + httpx (especialmente `Test_LocalReadAccessAuthorization`).
- **TK-HPS-16**: [TESTING] Cobertura de tests > 70% en módulo hps.
- **TK-HPS-17**: [TESTING] E2E test del ciclo de vida usuario HPS (alta → curso → observación → anexo → histórico).
- **TK-HPS-18**: [TESTING] Tests para las 5 FKs PK-to-PK genéricas tras formalización.

### Operación

- **TK-HPS-19**: [AUDIT] Validar que el frontend web NO tenga copias de PII (security audit).
- **TK-HPS-20**: [INFRA] Setup de Redis para caché server-side con TTL configurable.

---

## Anexo · Decisiones referenciadas

| Decisión | Aplicación a HPS |
|---|---|
| D8 (hexagonal global) | Toda la migración |
| D9-D10 (secret manager + adaptador unificado) | Eliminar `TbConfiguracionHPS` del frontend + usar secret manager |
| D14 (esquema por módulo) | Módulo `hps` en PostgreSQL |
| D16 (object storage S3-compatible) | Anexos de 4 fuentes (HPS, Histórico, SICA, Usuario) |
| D27 (logs estructurados canónicos) | Reemplazar `Errores de pegado` (sentinel) con log estructurado |
| D44-D46 (autorización + capabilities) | Roles Calidad/Técnico + `Test_LocalReadAccessAuthorization` portado |
| D66-D67 (stack) | Backend Python + HTMX |
| D68 (monolito modular) | HPS como módulo |
| D70-D71 (puerto de caché) | Server-side cache (Redis) con TTL + kill switch |
| D82 (Expand and Contract) | Migración de 22 tablas + descarte de 12 frontend local |
| D83 (apps independientes) | HPS_Solicitudes es app separada (ID 22) — referencia conceptual |
| D85 (kill switch atómico) | `TbConfiguracion.CacheHabilitada` → feature flag server-side |
| D86-D87 (cross-app + tests VBA) | FKs conceptuales + 9 Tests VBA como referencia |
| D91 (caché cross-cutting) | Patrón de 6 módulos caché preservado como referencia |
| D92 (PII CRITICAL) | **344 filas PII en `TbDatosLocal` — eliminar del frontend** |
| D94 (FKs conceptuales) | 5 PK-to-PK genéricos + 5 FKs conceptuales — formalizar |
| D102 (booleanos Text(2)) | 3 columnas en `TbUsuarios` → `BOOLEAN` |

## Anexo · Tabla de fuentes

| Fuente | Aporta |
|---|---|
| engram obs #24084 | Consolidado de descubrimiento HPS |
| engram obs #24097 | Audit codegraph-vba (12 ACTIVE / 1 UNCLEAR / 9 ZOMBIE) |
| engram topic_key `hps/usage-audit-2026-08-06` | Resultado del audit de uso |
| [`data-model.md`](data-model.md) | 22 tablas backend con schemas + relaciones |
| [`capabilities.md`](capabilities.md) | Inventario de features (vista de alto nivel) |
| [`forms.md`](forms.md) | Navegación, call paths, formularios críticos |
| [`integrations-automation.md`](integrations-automation.md) | Cross-app, SICA, testing sandbox, flags |
| [`migration-matrix.md`](migration-matrix.md) | D83, D92 disposiciones específicas de HPS |
| [`security-rules.md`](security-rules.md) | D92 PII, autorización, Indicadores, Sincronización |
| [`README.md`](README.md) | Estado del lote, hallazgos críticos, checklist |
| `data/staging/hps/src/classes/*.cls` | 29 clases de dominio |
| `data/staging/hps/src/modules/*.bas` | 30 módulos (4 bootstrap, 6 caché, 9 tests) |
| `data/staging/hps/src/forms/*.form.txt` | ~30 forms con RecordSource/RowSource |
| `data/staging/hps/src/tests/*.bas` | 9 archivos Test_*.bas |
| `C:\00repos\datos\HPST.accdb` | Backend autoritativo (volúmenes reales) |
| `C:\00repos\codigo\00_HPS\staging` | Fuente READ-ONLY |
| [DOCS](../../../DOCS.md) | Technical reference raíz del blueprint |
| [CODEBASE-GUIDE](../../../CODEBASE-GUIDE.md) | Para mantenedores del blueprint |

## Checklist del documento

- [x] Scope con 16+ features detalladas + 5 catálogos config seed-only
- [x] Inventario 22 tablas backend + 12 frontend (caché descartado) + audit 12/1/9
- [x] Hallazgos D92 CRITICAL (PII) + D94 (FKs) + D102 (booleanos) con anchor links
- [x] Decisiones UX Preservar/Mejorar/Nuevo paradigma por pantalla
- [x] Decisiones arquitectura hexagonal ports
- [x] Decisiones de seguridad (CRITICAL D92 PII)
- [x] Decisiones de datos (FKs conceptuales, booleanos, ZOMBIE)
- [x] Criterios de aceptación verificables y agrupados por dimensión
- [x] 17 pendientes operacionales antes, 13 durante, 4 en go-live
- [x] 20 tickets derivables preview (TK-HPS-1..20)
- [x] Tabla de decisiones referenciadas (D5-D102)
- [x] Tabla de fuentes
- [x] Idioma: español técnico neutro. Identificadores y paths sin traducir.
- [x] "The sentence that organizes" presente
- [x] "Scope del scope" presente
- [x] Sin emojis decorativos
- [x] Cross-references a DOCS, CODEBASE-GUIDE, AGENTS
- [x] Quick Navigation table
- [x] Hallazgos en tabla con severity

## Siguiente paso

Aplicar las mismas reglas a la 1 épica restante: HPS_Solicitudes. Tras cerrar el ciclo de revisión final del blueprint.

## Cómo se aplica a access2web-blueprint

HPS es la app read-heavy de gestión de personal sujeto a homologación (quién está obligado, qué cursos ha hecho, fechas de renovación, observaciones y motivos). Es la app con menos intensidad DAO después de Condor (107 callers `getdb()`), pero tiene un caché en el frontend con datos personales que es un patrón raro que NO aparece en Lanzadera/Expedientes/GR/NC.

**Entrada cruda**: walkthroughs G1..G5 consolidan los forms de HPS. Inventario Dysflow: 22 tablas, 6 FKs, 27 columnas en `TbUsuarios`. Volumen real: 345 usuarios activos (`TbUsuarios`), 242 históricos (`TbUsuariosHistoricos`), 1280 relaciones HPS (`TbHPS`), 330 observaciones (`TbObservaciones`).

**Decisiones operativas vigentes**: D92 (campos sensibles requieren disposición explícita antes de migrar: `DNI`, `Nombre`, `Apellido_1`, `Apellido_2`, `Telefono`, `Correo_e`, `F_Nacimiento`), 12 tablas locales en frontend con datos personales (TbDatosLocal con 344 filas, TbUsuariosHistoricosLocal con 235 filas, TbCursosLocal, etc.) requieren disposición explícita antes de migrar, 4 tablas «Copia de...» + sentinels de pegado masivo (legacy copy pattern), D87 (clsTestDouble* — 5 archivos de Test Doubles: disciplina TDD madura que se preserva como referencia para pytest), D102 (booleans Text(2) cross-cutting → `BOOLEAN`), `.gitignore` debe excluir `HPS.accdb` (con datos personales).

**Cross-references desde otros docs**: DOCS.md §The 8 Apps lista HPS como mergeada con PR pre-workflow; `docs/architecture.md` §Decisiones D-<n> vigentes lista D14, D27, D82; CODEBASE-GUIDE.md §Ownership de artefactos la referencia para el código de plataforma y el walkthrough JSON.

## Lista de comprobación final

- [ ] Las 7 secciones del cuerpo están completas y verificadas contra los walkthroughs G1..G5.
- [ ] Los 2 anexos están adjuntos con referencias válidas.
- [ ] Los criterios de aceptación están todos marcados; los pendientes tienen ticket derivado.
- [ ] Los tickets derivables tienen issue-number válido en el backlog del change correspondiente.
- [ ] La sección «Cómo se aplica a access2web-blueprint» referencia los walkthroughs G1..G5, el código destino `app/src/modules/hps/`, y las decisiones D92, D87, D102 + el patrón de caché frontend.
- [ ] La épica se ajusta al contrato de `skills/documentation-alan-style/SKILL.md` (§3 + §8): castellano peninsular formal con usted, sin emojis decorativos.
- [ ] Las cross-references desde DOCS.md, CODEBASE-GUIDE.md, `docs/architecture.md` y los walkthrough JSONs siguen resolviendo.

---

[← Back to HPS README](README.md) · [← Codebase Guide](../../../CODEBASE-GUIDE.md) · [← DOCS](../../../DOCS.md)
