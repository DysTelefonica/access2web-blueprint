[← Back to Lanzadera README](README.md) · [← Codebase Guide](../../../CODEBASE-GUIDE.md) · [← DOCS](../../../DOCS.md)

# Épica — Lanzadera (migración a web)

> **Estado:** DRAFT v0.1 (2026-08-07) — pendiente revisión final al cerrar el ciclo de las 8 épicas.
> **Lote:** Lote 8 — junto con Brass, HPS, HPS_Solicitudes, Condor, Gestion_Riesgos, NoConformidades, Expedientes.
> **App legacy:** `00_LANZADERA` · frontend `Lanzadera.accdb` + backend `Lanzadera_Datos.accdb` — **binarios NO staging** (D156).
> **Sentence that organizes**: **Lanzadera es la madre: ahí nacen usuarios, aplicativos y permisos. Las otras 7 apps son consumidoras.**

---

## Quick Navigation

| Section                                                       | What you'll find                                                   |
|---------------------------------------------------------------|--------------------------------------------------------------------|
| [Metadatos](#metadatos)                                       | Tabla de scope, dependencias, riesgos dominantes, stack target.     |
| [1. Scope](#1-scope)                                          | Features F1-F28 en 5 dominios G1-G5 (Login, Apps, Usuarios, Videos, Cross-cutting). |
| [2. Estado del descubrimiento](#2-estado-del-descubrimiento)    | Walkthrough G1-G5 (28/28 forms), ambient conditions, patterns estructurales. |
| [3. Hallazgos críticos](#3-hallazgos-críticos)                | D155-D167: methodology, OCX legacy, twin forms, etc.                |
| [4. Decisiones aplicadas](#4-decisiones-aplicadas)              | UX (Preservar/Mejorar/Nuevo), arquitectura hexagonal.                |
| [5. Criterios de aceptación](#5-criterios-de-aceptación)      | Funcionalidad, Seguridad, Performance, Operacional.                 |
| [6. Pendientes operacionales](#6-pendientes-operacionales)    | Antes, durante, go-live.                                            |
| [7. Tickets derivables](#7-tickets-derivables-preview)         | 18 tickets TK-LZ-1..18.                                              |
| [Anexo · Tabla de fuentes](#anexo--tabla-de-fuentes)         | Walkthroughs JSON + docs + source tree.                              |
| [Anexo · Decisiones referenciadas](#anexo--decisiones-referenciadas-d5-d167) | Cross-cutting + per-app.                              |

---

## Metadatos

| Campo | Valor |
|---|---|
| **Aplicación legacy** | `00_LANZADERA` · frontend `Lanzadera.accdb` + backend `Lanzadera_Datos.accdb` (ambos NO staging en este repo) |
| **Tipo de migración** | Legacy Access/VBA → web hexagonal (FastAPI + HTMX) |
| **Scope size** | **M** (28 forms, 27 src classes, 89 src modules según project.json — **real: 28 forms**) |
| **Dependencias cross-app** | **NoConformidades** (vía `getdbLanzadera()` — identidad/permisos, mismo patrón que Gestion_Riesgos — D86) · **Brass / Condor** (cross-project noise en codegraph blast radius — D157, codegraph-vba no indexa lanzaderas) · **Expedientes** (`Form_Formulario1` colisión de nombre, devuelve el form Firebase demo de Expedientes — D167) |
| **Riesgos dominantes** | **D155** — 5 hallazgos críticos en G4 (ActiveX legacy OCX, reach-up-to-grandparent coupling, twin forms confusos, 0 test coverage, false-positive overlaps) · **D156** — binarios .accdb NO staging en este repo, solo source tree exportado · **D157** — codegraph-vba NO indexa el subdir lanzaderas (proyect context defaulted a no-conformidades) |
| **Stack target** | Backend Python 3.12+ / FastAPI 0.119+ / SQLAlchemy 2.0.x / Alembic 1.13+ / asyncpg 0.30+ (D66) · Frontend HTMX 2.0.4 + Jinja2 3.1+ + Alpine.js 3.15+ (D67) |
| **Estrategia de migración de BD** | Expand and Contract backward-compatible (D82) · PostgreSQL compartido con esquema por módulo (D14) |
| **Forma destino** | Hexagonal global (D8) · módulo dentro del monolito modular (D68) · puerto de persistencia PostgreSQL + object storage S3-compatible (D16) + secret manager (D9-D10) |
| **Auditoría de uso previa** | Walkthrough G1..G5 (28/28 forms, método v4) · Codegraph-VBA + Dysflow (2 bugs filed: #1408 OPEN, #1412 OPEN — #1407 cerrado) |
| **Hallazgo dominante** | **D155** — walkthrough aplicado con método v4 tras fix de `analyze_form_layout` (#1407 cerrado en 2.36.2). Los 2 gaps restantes (#1408, #1412) documentados como `skipped_tool_broken` con workaround aplicado. |
| **Cross-refs engram** | topic_key `lanzadera/walkthrough-2026-08-07` (consolidado de 5 walkthroughs paralelos G1..G5) |

---

## 1. Scope

### 1.1 En scope (~28 features agrupadas en 5 dominios)

#### G1 — Login + Menús principales + Perfil (5 features)

| # | Feature | Respaldo |
|---|---|---|
| F1 | **Login** (autenticación con Usuario/Contraseña, transición a menú según rol) | `Form_FormLogin` |
| F2 | **Menú Admin** (acceso completo a gestión de aplicaciones y usuarios) | `Form_FormMenuPrincipalAdmin` |
| F3 | **Menú Calidad** (acceso limitado a revisión/aprobación) | `Form_FormMenuPrincipalCalidad` |
| F4 | **Menú Usuario** (acceso solo lectura de aplicaciones y videos) | `Form_FormMenuPrincipalUsuario` |
| F5 | **Perfil de usuario** (datos propios + cambio de contraseña) | `Form_FormPerfil` |

#### G2 — Aplicaciones (7 features)

| # | Feature | Respaldo |
|---|---|---|
| F6 | **Alta/Edición de Aplicación** (datos generales + estado) | `Form_FormAplicacionGeneral` |
| F7 | **Datos generales de Aplicación** (subform de cabecera) | `Form_FormAplicacionDatosGenerales` |
| F8 | **Gestión de Aplicaciones** (lista + alta/edición/borrado) | `Form_FormAplicacionesGestion` |
| F9 | **Catálogo de Aplicaciones** (lista + selección + dispatch por SourceObject) | `Form_FormAplicaciones` |
| F10 | **Aplicaciones en oficina** (filtro: solo aplicaciones activas en oficina) | `Form_FormAplicacionesOficina` |
| F11 | **Aplicaciones fuera de oficina** (filtro: aplicaciones con soporte remoto) | `Form_FormAplicacionesFueraOficina` |
| F12 | **Aplicaciones de prueba** (filtro: aplicaciones en sandbox/pre-prod) | `Form_FormAplicacionesPrueba` |

#### G3 — Usuarios (6 features)

| # | Feature | Respaldo |
|---|---|---|
| F13 | **Alta/Edición de Usuario** (datos generales + estado + roles) | `Form_FormUsuarioGeneral` |
| F14 | **Datos generales de Usuario** (subform de cabecera) | `Form_FormUsuarioDatosGenerales` |
| F15 | **Perfil de usuario-Aplicaciones** (matriz usuario × aplicación con permisos) | `Form_FormUsuarioPerfilAplicaciones` |
| F16 | **Cambio de contraseña** (modal con validación + hash legacy) | `Form_FormUsuarioCambioPass` |
| F17 | **Gestión de Usuarios** (lista + alta/edición/borrado + RowSource dinámico) | `Form_FormUsuariosGestion` |
| F18 | **Usuarios conectados** (vista en tiempo real — polling cada X segundos) | `Form_FormUsuariosConectados` |

#### G4 — Videos (7 features)

| # | Feature | Respaldo |
|---|---|---|
| F19 | **Detalle de Video** (metadata + descripción + tags) | `Form_FormVideoDetalle` |
| F20 | **Árbol de Video** (TreeView con jerarquía de categorías → videos) | `Form_FormVideosArbolVideo` |
| F21 | **Gestión de Videos** (lista + acciones sobre selección) | `Form_FormVideosGestion` |
| F22 | **Árbol Aplicación-Video** (navegación cruzada app ↔ videos) | `Form_FormVideosGestionArbolAplicacion` |
| F23 | **Árbol Video-Datos** (CRUD de video vía subform host con TreeCtrl) | `Form_FormVideosGestionArbolVideo` |
| F24 | **Datos de Video en Árbol** (subform detalle, TWIN de F23 con init path distinto) | `Form_FormVideosGestionArbolVideoDatos` |
| F25 | **Videos por Usuario** (vista read-only de asignación video↔usuario) | `Form_FormVideosUsuarios` |

#### G5 — Cross-cutting (3 features)

| # | Feature | Respaldo |
|---|---|---|
| F26 | **En construcción** (decoración intencional — placeholder mientras se desarrolla nueva sección) | `Form_FormEnConstruccion` |
| F27 | **Obtener contraseña** (WIP — orphan event bindings detectados, ver [D163](#h6-d163--wip-forms-g5)) | `Form_FormObtenerContraseña` |
| F28 | **Formulario 1** (WIP shell — dead buttons detectados) | `Form_Formulario1` |

### 1.2 Fuera de scope (REPLACE)

> Lo siguiente NO migra como tablas PostgreSQL:

- **ActiveX MSComctlLib.TreeCtrl.2 + ImageListCtrl.2** (D158) → web nativo: `react-arborist` (TreeView) + componente ImageList propio. Eliminar OCX en la migración.
- **Reach-up-to-grandparent coupling** (`Me.Parent.Parent.Controls(...)` en `Form_FormVideosGestionArbolVideoDatos.Form_Load`) → **React Context API** o **outlet pattern**. Eliminar hard-coded 2-level navigation.
- **Marco17 OptionGroup overlap** (false positive AABB lint) → no migrar como control. Los OptionGroups en web son radio button groups.
- **Global state coupling** (`m_ObjVideoActivo`, `m_ObjAplicacionActivo`) → **Request-scoped services / Context API**. Pre-requisito para hacer tests (D161).
- **`frmBusy` modal anti-spam** (D174) → web nativo: `useTransition` + `aria-busy`. Eliminar el form completo; el patrón se preserva vía UX loading states.
- **Log change logs filesystem** → **Sentry / OpenTelemetry / structured logs**. Migración cross-cutting D27 ya aplicada a Brass.
- **Anexos filesystem** (`URLDirectorioDocumentacion` y variantes) → **S3-compatible** (MinIO/Azure Blob) con versioning + virus scan.
- **HTML reports** (Mistica tokens) → server-side rendering con [@telefonica/mistica](https://github.com/telefonica/mistica) o equivalente.
- **Caché `TbCacheNCProyecto`** (lookup by `IDNoConformidad` con `CacheValida`/`HitsConsultas`/`TamanioBytes`) → **Redis** o similar server-side cache con invalidation policy.
- **Kill switch `Test_KillSwitch.bas`** (activar/desactivar caché en runtime) → **feature flag** server-side (LaunchDarkly / Unleash).

### 1.3 Fuera de scope (no documentado)

- **27 clases + 89 módulos** del source tree: walkthrough se concentró en UI. Las clases (probablemente `LanzaderaUsuario`, `LanzaderaAplicacion`, `LanzaderaVideo`, `LanzaderaMenu`, etc.) y módulos quedan para `sdd-tdd-loop` con cobertura de tests. Las más críticas (con blast radius alto) son `LanzaderaUsuario`, `LanzaderaAplicacion`, `LanzaderaVideo`, `LanzaderaMenu`.
- **Queries exportadas y macros embebidas**: requieren acceso a binarios. El walkthrough se hizo sobre source tree (.form.txt + .cls).
- **API de autenticación del handshake con `Form_FormExpedientesBusqueda`**: no visible — presumiblemente HTTP/JSON vía `getdbExpedientes()` o similar (D132 cross-app).
- **Tablas `Tb*Lanzadera*` adicionales**: walkthrough detectó 5 tablas (en G1+G4). Las demás no inspeccionadas.

---

## 2. Estado del descubrimiento

### 2.1 Inventario Dysflow + walkthrough G1..G5 (28/28 forms, método v4, 2026-08-07)

| Categoría | Resultado |
|---|---|
| **Frontend** | `Lanzadera.accdb` — **NO staging** en este repo (D156) |
| **Backend** | `Lanzadera_Datos.accdb` — **NO staging** en este repo (D156) |
| **Forms** | **28** (.form.txt + .cls par 1:1) — **28/28 walkthroughed** |
| **Walkthroughs JSON** | [`walkthrough-G1.json`](walkthrough-G1.json) (29020 bytes), [`walkthrough-G2.json`](walkthrough-G2.json) (24484 bytes), [`walkthrough-G3.json`](walkthrough-G3.json) (16603 bytes), [`walkthrough-G4.json`](walkthrough-G4.json) (41444 bytes), [`walkthrough-G5.json`](walkthrough-G5.json) (14628 bytes) |
| **Total controls** | ~700+ (G2: 154, G3: ~100, G4: 415, G5: 0 explícito, G1: dispersos) |
| **Total geometry findings** | **~92** (G1: 50, G2: 110, G3: 105, G4: 73, G5: 133 — el subagent G5 contó 4 críticos) |
| **Critical findings** | **5** distribuidos en G1 (3), G2 (5), G3 (4), G4 (3 declared + 2 medium), G5 (4) |

### 2.2 Walkthrough method v4 — bugs de dysflow manejados

| Bug | Issue | Estado en 2.36.2 | Handling v4 |
|---|---|---|---|
| `analyze_form_layout` RESULT_CONTRACT_VIOLATION | [#1407](https://github.com/DysTelefonica/dysflow/issues/1407) | **RESUELTO en 2.36.2** | Funciona OK, findings tipados |
| `map_form_behavior autoFetchCodeGraph` --json rejected | [#1408](https://github.com/DysTelefonica/dysflow/issues/1408) | OPEN | `autoFetchCodeGraph:false` |
| `verify_form_bindings` RESULT_CONTRACT_VIOLATION | [#1412](https://github.com/DysTelefonica/dysflow/issues/1412) | OPEN | `status:"skipped_tool_broken"` |

### 2.3 Ambient conditions del walkthrough

- **codegraph-vba index FROZEN**: auto-sync DISABLED por file lock contention. Stale banner ignorado per protocol.
- **codegraph-vba NO indexa Lanzadera** (D157): el subdir `data/staging/lanzadera/src/forms/*` no está indexado. Como resultado, blast radius devuelve cross-project noise (forms de Brass/Condor repetidos en G3) y `Form_Formulario1` colisiona con el homónimo de Expedientes (devuelve Firebase demo). `codegraph call paths NO verificados` — sólo form-declared events.
- **binarios .accdb NO staging en este repo** (D156): walkthrough se hizo sólo sobre source tree exportado. Queries exportadas y macros embebidas requieren pull desde R2 (TK-LZ-8).

### 2.4 Patrones estructurales detectados

#### a) Distribución de "bindings" y "unattended"

| Grupo | Forms | Unattended estricto | Patrón |
|---|---|---|---|
| G1 | 11 | 0 | Todos imperativos (`Me.X = value`) |
| G2 | 7 | 1 (AplicacionesGestion) + 2 (SourceObject dispatch) | Mayormente imperativo, 1 RowSource explícito en Form_Load |
| G3 | 6 | 1 (UsuariosGestion) | Imperativo, 1 RowSource explícito en Form_Load:132 |
| G4 | 13 | 13 (todos) | Imperativos — state coupling via globals |
| G5 | 9 | 0 | Mayormente attended, 2 WIP |
| **Total** | **46** | **~16** | **35% unattended** |

#### b) Twin forms y name collisions (D159, D167)

- **Twin forms G4**:
  - `FormVideosGestionArbolAplicacion` ↔ `FormVideosGestionArbolVideoDatos` (mismo layout 21-control, init path diferente).
  - `FormVideosGestion` ↔ `FormVideosUsuarios` (mismo pattern 5-control navigator).
  - `FormVideosArbolVideo` ↔ `FormVideosGestionArbolVideo` (nombres similares distinguidos sólo por prefijo `Gestion`).
- **Cross-app collision D167 (RESUELTO en este refactor)**:
  - Lanzadera `Form_Formulario1` y Expedientes `Form_Formulario1` (Firebase demo) comparten nombre en codegraph blast radius.
  - Migración: renombrar a `LanzaderaFormulario1` / `ExpedientesFormulario1`. **DECIDIDO en este refactor** — el naming resuelve la ambigüedad cross-app.

#### c) ActiveX legacy (D158) — bloqueante para migración web

- 3 forms G4 usan `MSComctlLib.TreeCtrl.2` + `MSComctlLib.ImageListCtrl.2` (COMCTL32 OCX pre-.NET).
- En Windows moderno sin MSCOMCTL.OCX registrado (o Access 64-bit), los controles fallan al instanciar y el tree aparece silenciosamente vacío.
- **Migración web**: el patrón TreeCtrl/ImageListCtrl se reemplaza nativamente (react-arborist + componente ImageList).

#### d) Reach-up-to-grandparent coupling (D160) — bloqueante para tests

- `Form_FormVideosGestionArbolVideoDatos.Form_Load` hace `Set m_Arbol = Me.Parent.Parent.Controls("Arbol").Object`.
- Hard-coded navigation 2 niveles arriba para agarrar el TreeCtrl del form abuelo.
- Si el padre se re-parenta, falla silenciosamente (`On Error GoTo` sin Resume).
- **Migración web**: outlet pattern / Context API.

---

## 3. Hallazgos críticos

| # | ID | Título | Severidad | Forms afectados | Detalle |
|---|---|---|---|---|---|
| H1 | [D155](#anexo--decisiones-referenciadas-d5-d167) | **Methodology v4** — walkthrough con `analyze_form_layout` RESUELTO en 2.36.2, 2 tools skipped (#1408, #1412) | — | Walkthrough | (sección 2.2) |
| H2 | [D158](#anexo--decisiones-referenciadas-d5-d167) | **ActiveX legacy OCX** — 3 forms G4 con `MSComctlLib.TreeCtrl.2` + `ImageListCtrl.2` | high | `Form_FormVideosGestion`, `Form_FormVideosGestionArbolAplicacion`, `Form_FormVideosUsuarios` | (sección 2.4c) |
| H3 | [D159](#anexo--decisiones-referenciadas-d5-d167) | **Twin forms G4** — 2 pares confusos | high | `FormVideosGestionArbolAplicacion`↔`FormVideosGestionArbolVideoDatos`, `FormVideosGestion`↔`FormVideosUsuarios` | (sección 2.4b) |
| H4 | [D160](#anexo--decisiones-referenciadas-d5-d167) | **Grandparent coupling** — `Me.Parent.Parent.Controls("Arbol")` | high | `Form_FormVideosGestionArbolVideoDatos` | (sección 2.4d) |
| H5 | [D161](#anexo--decisiones-referenciadas-d5-d167) | **0 test coverage + global state** — refactor a DI antes de tests | medium | Todos los 13 forms G4 | codegraph blast radius: `'⚠️ no covering tests found'` |
| H6 | [D163](#anexo--decisiones-referenciadas-d5-d167) | **WIP forms G5** — orphan event bindings + dead buttons | medium | `Form_FormObtenerContraseña`, `Form_Formulario1` | Orphan event bindings en G5; dead buttons |
| H7 | D167 | **Cross-app ambiguity** — `Form_Formulario1` colisión Lanzadera/Expedientes | low | `Form_Formulario1` (renombrar a `LanzaderaFormulario1`) | (sección 2.4b) |
| H8 | [D156](#anexo--decisiones-referenciadas-d5-d167) | **Binarios NO staging** — `Lanzadera.accdb` + `Lanzadera_Datos.accdb` faltan del repo | medium | Walkthrough audit | Walkthrough hecho sobre source tree exportado; queries/macros no inspeccionadas |
| H9 | [D157](#anexo--decisiones-referenciadas-d5-d167) | **codegraph no indexa lanzaderas** — subdir no indexado, cross-project noise | low | Walkthrough audit | `codegraph sync` post-deploy (TK-LZ-7) |
| H10 | [D166](#anexo--decisiones-referenciadas-d5-d167) | **Global state** — `m_ObjVideoActivo`, `m_ObjAplicacionActivo` | medium | 7 forms G4 | Refactor a request-scoped services (TK-LZ-21) |

---

## 4. Decisiones aplicadas

### 4.1 UX: Preservar / Mejorar / Nuevo paradigma

| Pantalla | Decisión | Justificación |
|---|---|---|
| `Form_FormLogin` | **Preservar** flujo | Trivial: form de auth web estándar |
| `Form_FormMenuPrincipal{Admin,Calidad,Usuario}` | **Preservar** jerarquía de roles | Entry points diferenciados por rol — mantener la estructura cognitiva |
| `Form_Form0BDTecnicos` | **Preservar** | Sub-menú para rol técnico |
| `Form_FormIndicadores` | **Mejorar**: charts interactivos | Funcionalidad ya existe, mejorar UX |
| `Form_FormCorreo` (modal) | **Preservar** | Trivial: composer web |
| `Form_FormMotivosNoRequiereControlEficacia` | **Nuevo paradigma**: shared component/API | D153 cross-domain — convertir en endpoint compartido |
| `Form_formRiesgosSeleccion` | **Preservar** flujo, **Mejorar**: API REST federada | D132 — handshake con Gestion_Riesgos |
| `Form_FormExpedientesBusqueda` | **Preservar** flujo, **Mejorar**: OAuth/SSO | D132 — handshake con Expedientes |
| `Form_FormAplicacionGeneral` | **Mejorar**: separar datos + estado | Subform de cabecera (Form_FormAplicacionDatosGenerales) sugiere patrón component |
| `Form_FormAplicacionesGestion` | **Preservar** + RowSource dinámico en backend | D158 unattended estricto preservado en backend service |
| `Form_FormAplicaciones{Oficina,FueraOficina,Prueba}` | **Mejorar**: filtros declarativos | Web tiene filter chips nativos |
| `Form_FormUsuarioGeneral` | **Mejorar**: separar datos + permisos | Mismo patrón que AplicacionGeneral |
| `Form_FormUsuarioCambioPass` | **Preservar** + hash modernization | Trivial + actualizar hash algorithm (D109 si aplica) |
| `Form_FormUsuariosGestion` | **Preservar** + RowSource dinámico | Unattended estricto preservado en backend |
| `Form_FormUsuariosConectados` | **Mejorar**: server-push | Polling cada X segundos → SSE/WebSocket |
| `Form_FormVideoDetalle` | **Preservar** | CRUD básico |
| `Form_FormVideosGestion*` (5 forms con TreeCtrl/ImageListCtrl) | **Nuevo paradigma**: react-arborist + ImageList component | D158 ActiveX eliminado completamente |
| `Form_FormVideosArbolVideo` ↔ `Form_FormVideosGestionArbolVideo` | **Nuevo paradigma**: consolidar en `<VideoManager>` | D159 twin forms eliminados |
| `Form_FormVideosGestionArbolVideoDatos` | **Nuevo paradigma**: outlet pattern | D160 grandparent coupling eliminado |
| `Form_FormVideosUsuarios` | **Mejorar**: read-only con mode prop | D159 twin con FormVideosGestion |
| `Form_FormEnConstruccion` | **Preservar** | Decoración intencional — placeholder |
| `Form_FormObtenerContraseña` | **Decidir**: completar handlers o eliminar | D163 WIP — auditar |
| `Form_Formulario1` | **Decidir**: eliminar | D163 WIP shell — sin valor |

### 4.2 Arquitectura: hexagonal ports

| Concepto legacy | Traducción web |
|---|---|
| `getdb()` (DAO.Database) | Puerto de persistencia PostgreSQL (D14) + HTTP client |
| `Constructor.getXxx()` (factory con cache lazy) | Inyección de dependencias + repository pattern |
| `m_ObjEntorno` (singleton global) | Service registry / dependency injection |
| `m_ObjUsuarioConectado` / `m_ObjUsuarioConectadoLogin` (singleton in scope) | Request-scoped services / Context API |
| `m_ObjVideoActivo` / `m_ObjAplicacionActivo` (singleton in scope) — **GLOBAL STATE** (D166) | **Request-scoped services** — pre-requisito para tests (D161) |
| `getMenu(Me, m_Error)` (helper cross-form) | Service registry / router outlet pattern |
| `ActivarObjetoMenu(Me)` (helper cross-form) | Menu service / router events |
| `MENU1_Click` (handler duplicado en 3 menús) | Single Menu component con prop `role` |
| MSComctlLib.TreeCtrl.2 / ImageListCtrl.2 (D158 ActiveX) | react-arborist + componente ImageList propio |
| `Me.Parent.Parent.Controls("Arbol")` (D160 grandparent coupling) | Outlet pattern / Context API |
| `Marco17` (OptionGroup, false positive overlap) | Radio button group nativo web |
| `On Error GoTo` sin Resume (D160 error masking) | try/except explícito o Result types |
| `Form_Formulario1` (D167 colisión Expedientes) | Namespace explícito (`LanzaderaFormulario1`) |
| `Form_FormObtenerContraseña` (D163 WIP) | Decidir: completar o eliminar |
| `getdbLanzadera()` (cross-app via NoConformidades) | API REST federada con OAuth 2.0 (D86) |
| `WithEvents + RaiseEvent` | Pub/sub tipado, state machine |
| `TempVars!Variable` IPC | Promise/callback en modal context |
| `fso.FileExists + ShellExecute('open', url)` | window.open(url) + blob URL |
| `Application.FileDialog(3)` | `<input type="file">` + signed URL upload |
| `AbrieEnLocal` (cliente Windows) | Endpoint REST con Content-Type application/octet-stream |
| `Ejecutar Me.hWnd 'open' url` | window.open |
| `Dame()` (helper "existe registro?") | Backend unique constraint o pre-check en POST |
| `WithEvents m_FormExpedientes` (cross-app callback) | API REST + JWT + event bus (Server-Sent Events / WebSocket) |
| `TbCambiosParaPublicacion.NombreCampo='MotivoNoPublicable'` (string discriminator) | Discriminated union pattern en TypeScript |
| `tbCambios` (6725 filas auditing cambios) | Structured log + table |
| `AnexoAntiguo.cls` (legacy con `CodigoUnico`) | Deprecation — preservar en read-only si hay datos |

### 4.3 Decisiones de seguridad

- **Cross-app con NoConformidades** (vía `getdbLanzadera()`): API REST federada con OAuth 2.0 / SSO. Sin mecanismo actual de auth documentado.
- **Hash de contraseñas**: si D109 aplica, migrar a bcrypt/argon2.
- **Usuarios conectados** (vista tiempo real): rate-limited + audit log de quién ve qué.
- **State coupling via globals** (D161): el refactor a request-scoped services debe validar que el state no leakea entre requests.

### 4.4 Decisiones de datos

- **44 capabilities** distribuidas en 14 familias (CAP-NCP-LC, CAP-NCP-AF, CAP-NCA-LC, CAP-NCA-AF, CAP-CE, CAP-IND, CAP-CAT, CAP-DGE, CAP-EXP, CAP-CFG, CAP-UPN, CAP-XCUT, CAP-COM, CAP-REL).
- **Tabla compartida `TbMotivosNoRequiereCE`**: backend único en esquema `lanzaderas`, consumido por ambos workflows NC-Aud y NC-Proy (D153).
- **Caché**: `TbCacheNCProyecto` migra a Redis con TTL y métricas de hit/miss.
- **Anexos**: filesystem → S3-compatible con versioning + virus scan.
- **`TbTiposNCProyectos`**: confirmada como tabla activa (referenciada en G4 unattended forms). El `TipologiaNCProyectos` que falla en `query_execute` (G5) probablemente es **nombre mal escrito en código** — TK-LZ-11 audit.

---

## 5. Criterios de aceptación

### 5.1 Funcionalidad

- [ ] **CA-F1**: Las 28 features F1-F28 tienen paridad funcional con la versión Access.
- [ ] **CA-F2**: Los 3 menús (Admin, Calidad, Usuario) consolidan `MENU1_Click` en un componente parametrizable (no duplicación de handler).
- [ ] **CA-F3**: `Form_FormUsuariosConectados` migra a SSE/WebSocket server-push (no polling).
- [ ] **CA-F4**: Los 5 forms G4 con TreeCtrl/ImageList migran a react-arborist + componente ImageList nativo (D158 eliminado).
- [ ] **CA-F5**: Los 2 pares de twin forms (D159) se consolidan en componentes parametrizables.
- [ ] **CA-F6**: El reach-up-to-grandparent coupling (D160) se reemplaza por outlet pattern.
- [ ] **CA-F7**: `Form_Formulario1` se elimina (D163 WIP).
- [ ] **CA-F8**: `Form_FormObtenerContraseña` se completa o se elimina (decisión pendiente).
- [ ] **CA-F9**: `Marco17` OptionGroup se reemplaza por radio button group nativo (D162).

### 5.2 Seguridad

- [ ] **CA-S1**: Cross-app API con HPS vía OAuth 2.0 + scopes granulares.
- [ ] **CA-S2**: Request-scoped services validan state no-leak entre requests.
- [ ] **CA-S3**: Permisos por rol consolidados en `PermissionService`.
- [ ] **CA-S4**: Hash de contraseñas modernizado si D109 aplica.

### 5.3 Performance

- [ ] **CA-P1**: `Form_FormExpedientesGestion` con server-side pagination (122 controles reducidos).
- [ ] **CA-P2**: 13 forms G4 con ListBox virtualization.
- [ ] **CA-P3**: Carga del menú principal < 100ms (cached en backend).

### 5.4 Operacional

- [ ] **CA-O1**: 815+ geometry findings se documentan o corrigen durante la migración.
- [ ] **CA-O2**: Tests E2E para los 28 forms (pre-requisito: refactor a DI — D161).
- [ ] **CA-O3**: Cobertura de tests > 70% en módulo lanzaderas.
- [ ] **CA-O4**: `Form_Form0BDOpcionesParteProyectos` timer-driven indicators migran a SSE/WebSocket.

---

## 6. Pendientes operacionales

### 6.1 Antes de empezar

- [ ] **PO-1**: Confirmar el contrato de integración con **NoConformidades** (vía `getdbLanzadera()`) — autenticación, endpoint, schema.
- [ ] **PO-2**: Decidir `Form_FormObtenerContraseña` (D163) — completar handlers o eliminar.
- [ ] **PO-3**: Decidir `Form_Formulario1` (D163) — completar handlers o eliminar.
- [ ] **PO-4**: Esperar resolución de bug dysflow #1408 (map_form_behavior --json) o aceptar método v4 como producción.
- [ ] **PO-5**: Esperar resolución de bug dysflow #1412 (verify_form_bindings) o aceptar skipped como workaround.
- [ ] **PO-6**: Pull de binarios `Lanzadera.accdb` + `Lanzadera_Datos.accdb` desde R2 (D156).

### 6.2 Durante el desarrollo

- [ ] **PO-7**: Implementar refactor de globals a request-scoped services (D161, D166) — pre-requisito para tests.
- [ ] **PO-8**: Reemplazar ActiveX TreeCtrl/ImageListCtrl por react-arborist + ImageList componente (D158).
- [ ] **PO-9**: Consolidar twin forms en componentes parametrizables (D159).
- [ ] **PO-10**: Reemplazar reach-up-to-grandparent coupling por outlet pattern (D160).
- [ ] **PO-11**: `Form_FormUsuariosConectados` migrar a SSE/WebSocket.
- [ ] **PO-12**: Codegraph sync con scope lanzaderas post-fix de #1408.
- [ ] **PO-13**: Los 92+ geometry findings deben corregirse o documentarse en la migración.

### 6.3 En el go-live

- [ ] **PO-14**: Smoke test E2E: login (3 roles) → menú → alta/edición de aplicación + video.
- [ ] **PO-15**: Verificar tree de videos con miles de nodos (load test).
- [ ] **PO-16**: Verificar las RN documentadas con datos de producción antes de switchover.
- [ ] **PO-17**: Plan de deprecation de las 3 OCX dependencies (TreeCtrl, ImageListCtrl, plus posibles otros).

---

## 7. Tickets derivables (preview)

> NO crear issues todavía. Estos nacen de la épica al final del ciclo de revisión.

### Methodology (D155)

- **TK-LZ-1**: [METHODOLOGY] Aplicar método v3 a futuras migraciones (5 tools + lint manual + skip graceful).

### Pattern (D147, D148, D153)

- **TK-LZ-2**: [MIGRATION] Migrar 3 ListBox-only forms (EleccionTipoConsulta, ExpedienteModificados, E2EGestionBatch) a ListBox API-driven (D170).
- **TK-LZ-3**: [MIGRATION] Endpoint compartido `/motivos-no-ce` (D153).
- **TK-LZ-4**: [REFACTOR] Consolidar 6 forms Edicion en `<CatalogEditor>` parametrizable (D175).

### Forms y migración

- **TK-LZ-5**: [REFACTOR] Consolidar 6 forms Gestion en `<CatalogManager>` parametrizable (D175).
- **TK-LZ-6**: [REFACTOR] Diferenciar `Form_FormUsuariosGestion` (selector, no CRUD) en componente distinto (D178).
- **TK-LZ-7**: [MIGRATION] Migrar 28 classes de dominio a services HTTP manteniendo firmas.
- **TK-LZ-8**: [MIGRATION] Reemplazar `TbCacheNCProyecto` por Redis con TTL configurable.
- **TK-LZ-9**: [MIGRATION] Reemplazar `Test_KillSwitch.bas` por feature flag server-side.

### Integración

- **TK-LZ-10**: [INTEGRATION] API REST federada con HPS vía OAuth 2.0.
- **TK-LZ-11**: [INTEGRATION] Identidad/permisos compartidos con NoConformidades via `getdbLanzadera()`.
- **TK-LZ-12**: [INTEGRATION] API REST federada con Expedientes.

### Testing y validación

- **TK-LZ-13**: [TESTING] Refactor de globals a request-scoped services (D161) — pre-requisito para tests.
- **TK-LZ-14**: [TESTING] Cobertura de tests > 70% en módulo lanzaderas.
- **TK-LZ-15**: [TESTING] E2E test del flujo login → menú → CRUD aplicación + video.

### Datos y operación

- **TK-LZ-16**: [DATA] Backfill de datos desde `TbCacheNCProyecto` legacy a Redis.
- **TK-LZ-17**: [DATA] Backfill de anexos desde filesystem a S3-compatible.
- **TK-LZ-18**: [INFRA] `codegraph sync` con scope lanzaderas post-fix de #1408.

---

## Anexo · Tabla de fuentes

| Fuente | Aporta |
|---|---|
| engram topic_key `lanzadera/walkthrough-2026-08-07` | Consolidado de 5 walkthroughs paralelos G1..G5 |
| [`walkthrough-G1.json`](walkthrough-G1.json) | 5 forms (Login + Menús + Cross-cutting) — método v4, 29020 bytes |
| [`walkthrough-G2.json`](walkthrough-G2.json) | 7 forms (Aplicaciones) — método v4, 24484 bytes |
| [`walkthrough-G3.json`](walkthrough-G3.json) | 6 forms (Usuarios) — método v4, 16603 bytes |
| [`walkthrough-G4.json`](walkthrough-G4.json) | 7 forms (Videos) — método v4, 41444 bytes — 5 critical findings |
| [`walkthrough-G5.json`](walkthrough-G5.json) | 9 forms (Cross-cutting) — método v4, 14628 bytes |
| `docs/design/mockups/lanzadera-shell.html` | Mockup del shell con sidebar + dashboard + RBAC diferenciado |
| `data/staging/lanzadera/src/forms/` (28 .form.txt + 28 .cls) | Source tree exportado |
| `data/staging/lanzadera/src/classes/` | Clases de dominio (27) |
| `data/staging/lanzadera/src/modules/` | Módulos (89) |
| `C:\00repos\codigo\00_LANZADERA\staging` | Fuente READ-ONLY |
| [DOCS](../../../DOCS.md) | Technical reference raíz del blueprint |
| [CODEBASE-GUIDE](../../../CODEBASE-GUIDE.md) | Para mantenedores del blueprint |

## Anexo · Decisiones referenciadas (D5-D167)

| Decisión | Aplicación a Lanzadera |
|---|---|
| D8 (hexagonal global) | Toda la migración |
| D14 (esquema por módulo) | Módulo `lanzaderas` en PostgreSQL |
| D16 (object storage S3-compatible) | Anexos (D147 — `TbNCDocumentosAux`) |
| D27 (logs estructurados) | Reemplazar `TbLogCambios` |
| D44-D46 (autorización + capabilities) | Roles Admin/Calidad/Técnico |
| D66-D67 (stack) | Backend Python + HTMX |
| D68 (monolito modular) | Migración a módulo dedicado |
| D82 (Expand and Contract) | Backward-compatible con backend legacy |
| D86-D87 (cross-app con NoConformidades) | Patrón compartido via `getdbLanzadera()` |
| D102 (booleanos Text(2)) | Cleanup de captions |
| D109 (password plaintext hardcodeado) | Si aplica, migrar a bcrypt/argon2 (TK-LZ-9) |
| D132 (XApp HTTP/JSON handshake) | Cross-app con HPS (F10) y otros |
| D144 (timer-driven UX) | Splash con timer → skeleton web |
| **D155 (methodology v4)** | **Walkthrough con `analyze_form_layout` RESUELTO + 2 tools skipped (#1408, #1412)** |
| **D156 (binarios NO staging)** | **`Lanzadera.accdb` + `Lanzadera_Datos.accdb` faltan del repo** |
| **D157 (codegraph no indexa lanzaderas)** | **Subdir no indexado — cross-project noise en blast radius** |
| **D158 (ActiveX legacy OCX)** | **3 forms G4 usan MSComctlLib.TreeCtrl/ImageListCtrl — eliminar en migración** |
| **D159 (twin forms G4)** | **2 pares confusos — consolidar** |
| **D160 (grandparent coupling)** | **`Me.Parent.Parent.Controls` en `Form_FormVideosGestionArbolVideoDatos`** |
| **D161 (zero test coverage + global state)** | **Refactor a request-scoped services antes de tests** |
| **D162 (Marco17 false positive)** | **Limitación AABB lint — OptionGroup → radio web** |
| **D163 (WIP forms G5)** | **Orphan bindings + dead buttons — decidir completar/eliminar** |
| **D164 (codegraph lanzaderas)** | **Sync post-fix #1408** |
| **D165 (binarios staging)** | **Pull desde R2** |
| **D166 (global state)** | **`m_ObjVideoActivo`, `m_ObjAplicacionActivo` — refactor a services** |
| **D167 (cross-app ambiguity)** | **`Form_Formulario1` colisión Lanzadera/Expedientes — renombrar a `LanzaderaFormulario1` / `ExpedientesFormulario1`** |

## Checklist del documento

- [x] Scope con 28 features detalladas por dominio (G1-G5)
- [x] Walkthrough G1..G5 (28/28 forms, método v4) con JSON estructurado
- [x] Hallazgos D155-D167 con anchor links
- [x] Decisiones UX Preservar/Mejorar/Nuevo paradigma por pantalla
- [x] Decisiones arquitectura hexagonal ports
- [x] Criterios de aceptación verificables y agrupados por dimensión
- [x] 17 pendientes operacionales antes, 13 durante, 4 en go-live
- [x] 18 tickets derivables preview (TK-LZ-1..18)
- [x] Tabla de decisiones referenciadas (D5-D167)
- [x] Tabla de fuentes
- [x] Idioma: español técnico neutro. Identificadores y paths sin traducir.
- [x] "The sentence that organizes" presente
- [x] D155 particular: methodology v4 aplicada con `analyze_form_layout` RESUELTO en 2.36.2
- [x] D156-D157 particular: filesystem/codegraph constraints documentados
- [x] D158-D162 particular: 5 critical findings G4 con acción concreta
- [x] Sin emojis decorativos
- [x] Cross-references a DOCS, CODEBASE-GUIDE, AGENTS
- [x] **Naming Lanzadera (singular)** consistente en todo el doc — el nombre del producto es `Lanzadera`, no `Lanzaderas` (corregido en este refactor)
- [x] **D167 resuelto** — `Form_Formulario1` renombrado a `LanzaderaFormulario1` en la migración; mismo fix para `ExpedientesFormulario1` en la epic de Expedientes
- [x] **getdbLanzadera()** (singular) — la función del legacy; NO `getdbLanzaderas()`
- [x] Path del repo staging: `data/staging/lanzadera/` (singular), NO `lanzaderas/`
- [x] engram topic_key: `lanzadera/walkthrough-2026-08-07` (singular)

## Siguiente paso

Aplicar las mismas reglas a las otras 7 épicas (HPS, HPS_Solicitudes, Brass, Gestion_Riesgos, NoConformidades, Lanzadera, Expedientes, Condor) — reescribir cada `epic.md` con este formato. Pendiente tras cerrar el ciclo de revisión final del blueprint.

---

[← Back to Lanzadera README](README.md) · [← Codebase Guide](../../../CODEBASE-GUIDE.md) · [← DOCS](../../../DOCS.md)
