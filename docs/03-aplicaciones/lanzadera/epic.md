# Épica — Lanzaderas (migración a web)

> **Estado:** DRAFT v0.1 (2026-08-07) — pendiente revisión final al cerrar el ciclo de las 8 épicas.
> **App legacy:** Lanzaderas (`C:\00repos\codigo\00_LANZADERA`) · frontend `Lanzadera.accdb` + backend `Lanzadera_Datos.accdb` — **binarios NO staging en este repo** (source-only, ver §2.5).
> **Lote de discovery:** Lote 8 — junto con Brass, HPS, HPS_Solicitudes, Condor, Gestion_Riesgos, NoConformidades, Expedientes.
> **Hallazgo dominante:** **D155 (methodology)** — walkthrough aplicado con método v4 tras fix de `analyze_form_layout` (#1407 cerrado en 2.36.2). Los 2 gaps restantes (#1408, #1412) documentados como `skipped_tool_broken` con workaround aplicado.
> **Cross-refs:** engram topic_key `lanzaderas/walkthrough-2026-08-07` (consolidado de 5 walkthroughs paralelos G1..G5).

## Metadatos

| Campo | Valor |
|---|---|
| **Aplicación legacy** | `00_LANZADERA` · frontend `Lanzadera.accdb` + backend `Lanzadera_Datos.accdb` (ambos NO staging en este repo) |
| **Tipo de migración** | Legacy Access/VBA → web hexagonal (FastAPI + HTMX) |
| **Scope size** | **M** (28 forms, 27 src classes, 89 src modules según project.json — **real: 28 forms**) |
| **Dependencias cross-app** | **NoConformidades** (vía `getdbLanzadera()` — presumiblemente identidad/permisos, mismo patrón que Gestion_Riesgos — D86) · **Brass / Condor** (cross-project noise en codegraph blast radius — síntoma de codegraph-vba no indexar lanzaderas, ver §2.5) · **Expedientes** (`Form_Formulario1` colisión de nombre, devuelve el form Firebase demo de Expedientes — ver §3.7) |
| **Riesgo dominante** | **D155** — 5 hallazgos críticos en G4 (ActiveX legacy OCX, reach-up-to-grandparent coupling, twin forms confusos, 0 test coverage, false-positive overlaps); **D156** — binarios .accdb NO staging en este repo, solo source tree exportado; **D157** — codegraph-vba NO indexa el subdir lanzaderas (proyect context defaulted a no-conformidades) |
| **Stack target** | Backend Python 3.12+ / FastAPI 0.119+ / SQLAlchemy 2.0.x / Alembic 1.13+ / asyncpg 0.30+ (D66) · Frontend HTMX 2.0.4 + Jinja2 3.1+ + Alpine.js 3.15+ (D67) |
| **Estrategia de migración de BD** | Expand and Contract backward-compatible (D82) · PostgreSQL compartido con esquema por módulo (D14) |
| **Forma destino** | Hexagonal global (D8) · módulo dentro del monolito modular (D68) · puerto de persistencia PostgreSQL + object storage S3-compatible (D16) + secret manager (D9-D10) |
| **Auditoría de uso previa** | ✅ Walkthrough G1..G5 (28/28 forms, método v4) · ✅ Codegraph-VBA + Dysflow (2 bugs filed: #1408 OPEN, #1412 OPEN — #1407 cerrado) |

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

### 1.1 En scope (~28 features agrupadas en 5 dominios)

#### G1 — Login + Menús principales + Perfil (~5 features)

| # | Feature | Respaldo |
|---|---|---|
| F1 | **Login** (autenticación con Usuario/Contraseña, transición a menú según rol) | `Form_FormLogin` |
| F2 | **Menú Admin** (acceso completo a gestión de aplicaciones y usuarios) | `Form_FormMenuPrincipalAdmin` |
| F3 | **Menú Calidad** (acceso limitado a revisión/aprobación) | `Form_FormMenuPrincipalCalidad` |
| F4 | **Menú Usuario** (acceso solo lectura de aplicaciones y videos) | `Form_FormMenuPrincipalUsuario` |
| F5 | **Perfil de usuario** (datos propios + cambio de contraseña) | `Form_FormPerfil` |

#### G2 — Aplicaciones (~7 features)

| # | Feature | Respaldo |
|---|---|---|
| F6 | **Alta/Edición de Aplicación** (datos generales + estado) | `Form_FormAplicacionGeneral` |
| F7 | **Datos generales de Aplicación** (subform de cabecera) | `Form_FormAplicacionDatosGenerales` |
| F8 | **Gestión de Aplicaciones** (lista + alta/edición/borrado) | `Form_FormAplicacionesGestion` |
| F9 | **Catálogo de Aplicaciones** (lista + selección + dispatch por SourceObject) | `Form_FormAplicaciones` |
| F10 | **Aplicaciones en oficina** (filtro: solo aplicaciones activas en oficina) | `Form_FormAplicacionesOficina` |
| F11 | **Aplicaciones fuera de oficina** (filtro: aplicaciones con soporte remoto) | `Form_FormAplicacionesFueraOficina` |
| F12 | **Aplicaciones de prueba** (filtro: aplicaciones en sandbox/pre-prod) | `Form_FormAplicacionesPrueba` |

#### G3 — Usuarios (~6 features)

| # | Feature | Respaldo |
|---|---|---|
| F13 | **Alta/Edición de Usuario** (datos generales + estado + roles) | `Form_FormUsuarioGeneral` |
| F14 | **Datos generales de Usuario** (subform de cabecera) | `Form_FormUsuarioDatosGenerales` |
| F15 | **Perfil de usuario-Aplicaciones** (matriz usuario × aplicación con permisos) | `Form_FormUsuarioPerfilAplicaciones` |
| F16 | **Cambio de contraseña** (modal con validación + hash legacy) | `Form_FormUsuarioCambioPass` |
| F17 | **Gestión de Usuarios** (lista + alta/edición/borrado + RowSource dinámico) | `Form_FormUsuariosGestion` |
| F18 | **Usuarios conectados** (vista en tiempo real — polling cada X segundos) | `Form_FormUsuariosConectados` |

#### G4 — Videos (~7 features)

| # | Feature | Respaldo |
|---|---|---|
| F19 | **Detalle de Video** (metadata + descripción + tags) | `Form_FormVideoDetalle` |
| F20 | **Árbol de Video** (TreeView con jerarquía de categorías → videos) | `Form_FormVideosArbolVideo` |
| F21 | **Gestión de Videos** (lista + acciones sobre selección) | `Form_FormVideosGestion` |
| F22 | **Árbol Aplicación-Video** (navegación cruzada app ↔ videos) | `Form_FormVideosGestionArbolAplicacion` |
| F23 | **Árbol Video-Datos** (CRUD de video vía subform host con TreeCtrl) | `Form_FormVideosGestionArbolVideo` |
| F24 | **Datos de Video en Árbol** (subform detalle, TWIN de #23 con init path distinto) | `Form_FormVideosGestionArbolVideoDatos` |
| F25 | **Videos por Usuario** (vista read-only de asignación video↔usuario) | `Form_FormVideosUsuarios` |

#### G5 — Cross-cutting (~3 features)

| # | Feature | Respaldo |
|---|---|---|
| F26 | **En construcción** (decoración intencional — placeholder mientras se desarrolla nueva sección) | `Form_FormEnConstruccion` |
| F27 | **Obtener contraseña** (WIP — orphan event bindings detectados, ver §3.6) | `Form_FormObtenerContraseña` |
| F28 | **Formulario 1** (WIP shell — dead buttons detectados, ver §3.6) | `Form_Formulario1` |

### 1.2 Fuera de scope (REPLACE)

> Lo siguiente NO migra como tablas PostgreSQL:

- **ActiveX MSComctlLib.TreeCtrl.2 + ImageListCtrl.2** (D158) → web nativo: `react-arborist` (TreeView) + componente ImageList propio. Eliminar OCX en la migración.
- **Reach-up-to-grandparent coupling** (`Me.Parent.Parent.Controls(...)` en `Form_FormVideosGestionArbolVideoDatos.Form_Load`) → **React Context API** o **router outlet pattern**. Eliminar hard-coded 2-level navigation.
- **Marco17 OptionGroup overlap** (false positive AABB lint) → no migrar como control. Los OptionGroups en web son radio button groups.
- **Global state coupling** (`m_ObjVideoActivo`, `m_ObjAplicacionActivo`) → **Request-scoped services / Context API**. Pre-requisito para hacer tests (D161).

### 1.3 Fuera de scope (no documentado)

- **27 clases + 89 módulos** del source tree: walkthrough se concentró en UI. Las clases (probablemente `LanzaderaUsuario`, `LanzaderaAplicacion`, `LanzaderaVideo`, `LanzaderaMenu`, etc.) y módulos quedan para `sdd-tdd-loop` con cobertura de tests.
- **Binarios `Lanzadera.accdb` y `Lanzadera_Datos.accdb`** (D156): NO están staging en este repo. Para walkthrough y auditoría de queries/macros se requiere staging de los binarios (proceso pendiente).
- **API de autenticación interna**: presumiblemente vía EVE (D144 cross-app).
- **Queries exportadas y macros embebidas**: no inspeccionadas — requieren acceso a binarios.
- **Handshake cross-app con `Form_FormExpedientesBusqueda`**: no visible — presumiblemente HTTP/JSON vía `getdbExpedientes()` o similar (D132 cross-app).

---

## 2. Estado del descubrimiento

### 2.1 Inventario Dysflow + walkthrough G1..G5 (28/28 forms, método v4, 2026-08-07)

| Categoría | Resultado |
|---|---|
| **Frontend** | `Lanzadera.accdb` — **NO staging en este repo** (D156) |
| **Backend** | `Lanzadera_Datos.accdb` — **NO staging en este repo** (D156) |
| **Forms** | **28** (.form.txt + .cls par 1:1) — **28/28 walkthroughed** |
| **Walkthroughs JSON** | `walkthrough-G1.json` (29020 bytes), `walkthrough-G2.json` (24484 bytes), `walkthrough-G3.json` (16603 bytes), `walkthrough-G4.json` (22849 bytes), `walkthrough-G5.json` (14628 bytes) |
| **Total controls** | ~530+ (154 en G2 + sumandos en G1/G3/G4/G5) |
| **Total geometry findings** | **~387** (G1: 99, G2: 110, G3: 105, G4: 73, G5: ~0 declarado) |
| **Critical findings** | **5** (todos en G4 — ActiveX legacy, grandparent coupling, twin forms, 0 tests, false-positive overlaps) |

### 2.2 Walkthrough method v4 — bugs de dysflow manejados

| Bug | Issue | Estado en 2.36.2 | Handling v4 |
|---|---|---|---|
| `analyze_form_layout` RESULT_CONTRACT_VIOLATION | #1407 | ✅ **RESUELTO en 2.36.2** | Funciona OK, devuelve findings tipados (FORM_LAYOUT_MISSING_GEOMETRY, FORM_LAYOUT_OVERLAP, FORM_LAYOUT_ALIGNMENT, FORM_LAYOUT_TAB_ORDER, FORM_LAYOUT_OFF_SECTION) |
| `map_form_behavior autoFetchCodeGraph` --json rejected | #1408 | ❌ OPEN | `autoFetchCodeGraph:false` en cada llamada |
| `verify_form_bindings` RESULT_CONTRACT_VIOLATION | #1412 | ❌ OPEN | `status:"skipped_tool_broken"` en cada form |

### 2.3 Ambient conditions del walkthrough

- **codegraph-vba index FROZEN**: auto-sync DISABLED por file lock contention (mismo que en walkthrough de NoConformidades).
- **codegraph-vba NO indexa lanzaderas** (D157): el subdir `data/staging/lanzaderas/src/forms/*` no está indexado. Como resultado, blast_radius devuelve **cross-project noise** (forms de Brass/Condor repetidos en los 6 forms de G3) y `Form_Formulario1` colisiona con el homónimo de Expedientes (devuelve Firebase demo). **Codegraph call paths NO verificados** — sólo form-declared events.
- **Stale banner ignorado** per protocol: documentado en cada walkthrough con `stale_banner_ignored:true`.

### 2.4 Patrones estructurales detectados

#### a) Distribución de "bindings" y "unattended"
- **G1** (Login + Menús): 0 unattended. Los 5 usan patrón **imperativo** (`Me.X = value`).
- **G2** (Aplicaciones): 1 unattended **estricto** (Form_FormAplicacionesGestion, RowSource en Form_Load) + 2 por SourceObject dispatch (Form_FormAplicaciones, Form_FormAplicacionGeneral). 154 controles declarados pero 0 DAO calls y 0 inline SELECT — G2 es capa de presentación/navegación sin data layer directo.
- **G3** (Usuarios): 1 unattended estricto (Form_FormUsuariosGestion, `Me.ListaUsuarios.RowSource="ID;Correo;Nombre"` en Form_Load:132).
- **G4** (Videos): 7 unattended (todos imperativos state coupling via globals `m_ObjVideoActivo` / `m_ObjAplicacionActivo`). Twin forms detectados.
- **G5** (Cross-cutting): 0 unattended estricto (FormEnConstruccion es decoración intencional).

#### b) Twin forms y name collisions (D159)
- **FormVideosGestionArbolAplicacion ↔ FormVideosGestionArbolVideoDatos**: mismo layout de 21 controles, mismo event surface, `.cls` diferente sólo en init path (uno para Video nuevo, otro para Video existente).
- **FormVideosGestion ↔ FormVideosUsuarios**: mismo patrón de 5 controles navigator (Tree + ImageList + Splitter + Subform). Uno enruta a edit subforms, otro a read-only.
- **FormVideosArbolVideo ↔ FormVideosGestionArbolVideo**: nombres similares distinguen roles (`ArbolVideo` vs `GestionArbolVideo`).
- **Form_Formulario1 colisión cross-app**: el Lanzaderas `Form_Formulario1` y el Expedientes `Form_Formulario1` (Firebase demo) comparten nombre en codegraph. Resolver en migración.

#### c) ActiveX legacy (D158) — bloqueante para migración
- 3 forms de G4 usan **MSComctlLib.TreeCtrl.2** + **MSComctlLib.ImageListCtrl.2** (COMCTL32 OCX pre-.NET).
- En Windows moderno sin MSCOMCTL.OCX registrado (o Access 64-bit), estos fallan al instanciar — tree aparece silenciosamente vacío.
- **Implicación para migración web**: el patrón TreeCtrl/ImageListCtrl se reemplaza nativamente (react-arborist + componente ImageList).

#### d) Reach-up-to-grandparent coupling (D160) — bloqueante para tests
- `Form_FormVideosGestionArbolVideoDatos.Form_Load` hace `Set m_Arbol = Me.Parent.Parent.Controls("Arbol").Object`.
- Hard-coded navigation 2 niveles arriba para agarrar el TreeCtrl del form abuelo.
- Si el form padre se re-parenta o stack diferente, falla silenciosamente.
- Usa `On Error GoTo` sin `Resume` — failures enmascaradas.
- **Implicación para tests**: sin DI, imposible testear en aislamiento.

### 2.5 Hallazgos filesystem (cross-cutting)

- **D156**: `data/staging/lanzaderas/frontend/Lanzadera.accdb` NO existe — lanzaderas es source-only en este repo.
- **D156**: `data/staging/lanzaderas/backend/Lanzadera_Datos.accdb` NO existe tampoco.
- **D157**: codegraph-vba NO indexa `data/staging/lanzaderas/src/forms/*` (proyect context defaulted a no-conformidades).
- **Encoding issue**: `Form_FormObtenerContraseña` usa Windows-1252 `ñ` (0xF1) en lugar de UTF-8 — dysflow resuelve internamente, PowerShell muestra replacement character (U+FFFD). No es bloqueante.

---

## 3. Hallazgos críticos

### H1 (D158 — ACTIVEX LEGACY) — 3 forms de G4 dependen de MSCOMCTL.OCX

**Síntoma**: `Form_FormVideosGestion`, `Form_FormVideosGestionArbolAplicacion`, `Form_FormVideosUsuarios` usan `MSComctlLib.TreeCtrl.2` + `MSComctlLib.ImageListCtrl.2`. Son OCX pre-.NET (COMCTL32). En Windows moderno sin MSCOMCTL.OCX registrado (o Access 64-bit), los controles fallan al instanciar y el tree aparece silenciosamente vacío.

**Acciones** (TK-LZ-1):
- ✅ Eliminar OCX completamente en migración web.
- ✅ TreeView → `react-arborist` (o equivalente nativo web).
- ✅ ImageList → componente propio con SVG icons.

### H2 (D159 — TWIN FORMS) — 4 forms confusos por nombre/similaridad

**Síntoma**: 2 pares de twin forms detectados en G4:
- `FormVideosGestionArbolAplicacion` ↔ `FormVideosGestionArbolVideoDatos` (mismo layout 21-control, init path diferente).
- `FormVideosGestion` ↔ `FormVideosUsuarios` (mismo pattern navigator 5-control, routing diferente).
- Además: `FormVideosArbolVideo` ↔ `FormVideosGestionArbolVideo` colisión de nombres distinguida sólo por prefijo `Gestion`.

**Acciones** (TK-LZ-2):
- Consolidar twin forms en un componente parametrizable: `<VideoManager mode="editor"|"readonly" route="new"|"edit">`.
- Renombrar en migración para evitar colisión.

### H3 (D160 — GRANDPARENT COUPLING) — Hard-coded 2-level navigation up

**Síntoma**: `Form_FormVideosGestionArbolVideoDatos.Form_Load` ejecuta `Set m_Arbol = Me.Parent.Parent.Controls("Arbol").Object`. Es reach-up-to-grandparent: el subform accede al control del form abuelo vía jerarquía hard-coded. Si el padre se re-parenta, falla silenciosamente (`On Error GoTo` sin Resume).

**Acciones** (TK-LZ-3):
- ❌ NO migrar este patrón a web.
- ✅ Reemplazar con **React Context API** o **outlet pattern** (router prop-drills el TreeCtrl al subform).

### H4 (D161 — ZERO TEST COVERAGE) — 0 tests en todos los forms de G4

**Síntoma**: codegraph blast radius muestra "⚠️ no covering tests found" para todos los `Form_Form*` del módulo lanzaderas. Todos los forms de G4 dependen de **global state** (`m_ObjVideoActivo`, `m_ObjAplicacionActivo`) — imposible testear en aislamiento sin refactor a dependency injection.

**Acciones** (TK-LZ-4):
- Pre-requisito para testing: refactor de globals a `request-scoped services` (D8 hexagonal).
- Después: tests unitarios + integration con mocks de los services.

### H5 (D162 — FALSE POSITIVE OVERLAPS) — Marco17 OptionGroup limita lint AABB

**Síntoma**: `FormVideosGestionArbolAplicacion` y `FormVideosGestionArbolVideoDatos` reportan 14 overlap findings cada uno porque Marco17 (OptionGroup container) geométricamente contiene todos los child controls. `analyze_form_layout` reporta esto como overlap pero es containment intencional. **Limitación conocida del strict AABB lint** — no es bug real.

**Acciones** (TK-LZ-5):
- Documentar en epic: Marco17 OptionGroup es **known false-positive** en AABB lint.
- En la migración web, los OptionGroups se reemplazan por radio button groups nativamente.

### H6 (D163 — WIP FORMS) — 2 forms con bindings rotos en G5

**Síntoma**: `Form_FormObtenerContraseña` tiene **orphan event bindings** (handler declarado sin implementar). `Form_Formulario1` tiene **dead buttons** (botones sin handler o con handler que no hace nada).

**Acciones** (TK-LZ-6):
- Auditar y completar los handlers de FormObtenerContraseña, o eliminar el form si la funcionalidad está duplicada en otro lado.
- Decidir si Formulario1 sigue siendo necesario o se elimina.

### H7 (D164 — CODEGRAPH NO INDEXA LANZADERAS) — Cross-project noise

**Síntoma**: codegraph-vba NO indexa el subdir `data/staging/lanzaderas/src/forms/*`. Como resultado:
- Blast radius devuelve cross-project noise (forms de Brass/Condor repetidos en G3 — "stale index noise, same line each time").
- `Form_Formulario1` en Lanzaderas colisiona con el homónimo de Expedientes (devuelve Firebase demo).

**Acciones** (TK-LZ-7):
- `codegraph sync` con scope lanzaderas (post-fix de #1408).
- O configurar project context por worktree.

### H8 (D165 — BINARIOS NO STAGING) — Lanzaderas source-only

**Síntoma**: `data/staging/lanzaderas/frontend/Lanzadera.accdb` y `data/staging/lanzaderas/backend/Lanzadera_Datos.accdb` NO existen en este repo. Walkthrough se hizo sólo sobre source tree exportado.

**Implicación**: queries exportadas, macros embebidas, y código de eventos no exportado al .cls NO fueron walkthroughed. Se requiere staging de binarios para auditoría completa.

**Acciones** (TK-LZ-8):
- Pull de binarios desde R2 (proceso pendiente de infraestructura).

### H9 (D166 — STATE COUPLING GLOBAL) — m_ObjVideoActivo / m_ObjAplicacionActivo

**Síntoma**: 7 forms de G4 usan globals module-level (`m_ObjVideoActivo`, `m_ObjAplicacionActivo`) como singletons in scope. Imposible testear sin refactor.

**Acciones** (TK-LZ-9):
- Pre-requisito: dependency injection + request-scoped services.

### H10 (D167 — CROSS-APP AMBIGUITY) — Form_Formulario1 colisión con Expedientes

**Síntoma**: Lanzaderas `Form_Formulario1` y Expedientes `Form_Formulario1` (Firebase demo) comparten nombre en codegraph blast radius. El walkthrough de Lanzaderas devolvió el equivocado.

**Acciones** (TK-LZ-10):
- Renombrar en migración: `LanzaderaFormulario1` / `ExpedientesFormulario1` para evitar ambigüedad.

---

## 4. Decisiones aplicadas

### 4.1 UX: Preservar / Mejorar / Nuevo paradigma

| Pantalla | Decisión | Justificación |
|---|---|---|
| `Form_FormLogin` | **Preservar** flujo | Trivial: form de auth web estándar |
| `Form_FormMenuPrincipal{Admin,Calidad,Usuario}` | **Preservar** jerarquía de roles | Entry points diferenciados por rol — mantener la estructura cognitiva |
| `Form_FormPerfil` | **Preservar** | Trivial |
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
| `m_ObjVideoActivo` / `m_ObjAplicacionActivo` (singleton in scope) — **GLOBAL STATE** | **Request-scoped services** (D161 — pre-requisito para tests) |
| `getMenu(Me, m_Error)` (helper cross-form) | Service registry / router outlet pattern |
| `ActivarObjetoMenu(Me)` (helper cross-form) | Menu service / router events |
| `MENU1_Click` (handler duplicado en 3 menús) | Single Menu component con prop `role` |
| MSComctlLib.TreeCtrl.2 / ImageListCtrl.2 (D158 ActiveX) | react-arborist + componente ImageList propio |
| `Me.Parent.Parent.Controls("Arbol")` (D160 grandparent coupling) | Outlet pattern / Context API |
| `Marco17` (OptionGroup, false positive overlap) | Radio button group nativo web |
| `On Error GoTo` sin Resume (D160 error masking) | try/except explícito o Result types |
| `Form_Formulario1` (colisión Expedientes) | Namespace explícito (`LanzaderaFormulario1`) |
| `Form_FormObtenerContraseña` (WIP) | Decidir: completar o eliminar |
| `getdbLanzadera()` (cross-app via NoConformidades) | API REST federada con OAuth 2.0 (D86) |
| `WithEvents + RaiseEvent` | Pub/sub tipado, state machine |
| `TempVars!Variable` IPC | Promise/callback en modal context |
| `fso.FileExists + ShellExecute('open', url)` | window.open(url) + blob URL |

### 4.3 Decisiones de seguridad

- **Cross-app con NoConformidades** (vía `getdbLanzadera()`): API REST federada con OAuth 2.0 / SSO. Sin mecanismo actual de auth documentado.
- **Hash de contraseñas**: si D109 aplica, migrar a bcrypt/argon2.
- **Usuarios conectados** (vista tiempo real): rate-limited + audit log de quién ve qué.
- **State coupling via globals** (D161): el refactor a DI debe validar que el request-scoped service no leakea entre requests.

### 4.4 Decisiones de datos

- **Binarios NO staging** (D165): requieren pull desde R2. Sin esto, queries exportadas y macros NO se pueden auditar.
- **Codegraph-vba index** (D164): requiere `codegraph sync` con scope lanzaderas post-fix de #1408.
- **Twin forms** (D159): consolidar en 1 tabla `videos` + 1 tabla `videos_usuarios` (asignación).
- **ActiveX** (D158): no migrar — eliminar completamente.

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

- [ ] **CA-S1**: Cross-app API con OAuth 2.0 + scopes granulares (NoConformidades via `getdbLanzadera()`).
- [ ] **CA-S2**: Hash de contraseñas modernizado (bcrypt/argon2) si D109 aplica.
- [ ] **CA-S3**: Request-scoped services validan que el state no leakea entre requests (D161).
- [ ] **CA-S4**: Try/except explícito reemplaza `On Error GoTo` sin Resume (D160).

### 5.3 Performance

- [ ] **CA-P1**: `Form_FormUsuariosConectados` con SSE/WebSocket — updates sub-segundo.
- [ ] **CA-P2**: TreeView de G4 con virtualización (react-arborist) — render smooth con miles de nodos.
- [ ] **CA-P3**: Codegraph-vba index lanzaderas refrescado post-deploy (D157, D164).

### 5.4 Operacional

- [ ] **CA-O1**: Binarios `Lanzadera.accdb` + `Lanzadera_Datos.accdb` staging en repo (D165 — pull desde R2).
- [ ] **CA-O2**: Audit de queries exportadas y macros embebidas una vez staging disponible.
- [ ] **CA-O3**: Tests unitarios + integration para los 28 forms (D161 — pre-requisito: dependency injection).
- [ ] **CA-O4**: Cobertura de tests > 70% en módulo lanzaderas.

---

## 6. Pendientes operacionales

### 6.1 Antes de empezar

- [ ] **PO-1**: Confirmar el contrato de integración con **NoConformidades** (vía `getdbLanzadera()`) — autenticación, endpoint, schema.
- [ ] **PO-2**: Decidir `Form_FormObtenerContraseña` (D163) — completar handlers o eliminar.
- [ ] **PO-3**: Decidir `Form_Formulario1` (D163) — completar handlers o eliminar.
- [ ] **PO-4**: Esperar resolución de bug dysflow #1408 (map_form_behavior --json) o aceptar método v4 como producción.
- [ ] **PO-5**: Esperar resolución de bug dysflow #1412 (verify_form_bindings) o aceptar skipped como workaround.
- [ ] **PO-6**: Pull de binarios `Lanzadera.accdb` + `Lanzadera_Datos.accdb` desde R2 (D165).

### 6.2 Durante el desarrollo

- [ ] **PO-7**: Implementar refactor de globals a request-scoped services (D161) — pre-requisito para tests.
- [ ] **PO-8**: Reemplazar ActiveX TreeCtrl/ImageListCtrl por react-arborist + ImageList componente (D158).
- [ ] **PO-9**: Consolidar twin forms en componentes parametrizables (D159).
- [ ] **PO-10**: Reemplazar reach-up-to-grandparent coupling por outlet pattern (D160).
- [ ] **PO-11**: `Form_FormUsuariosConectados` migrar a SSE/WebSocket.
- [ ] **PO-12**: Codegraph sync con scope lanzaderas post-fix de #1408.
- [ ] **PO-13**: Los 387 geometry findings deben corregirse o documentarse en la migración.

### 6.3 En el go-live

- [ ] **PO-14**: Smoke test E2E: login (3 roles) → menú → alta/edición de aplicación + video.
- [ ] **PO-15**: Verificar tree de videos con miles de nodos (load test).
- [ ] **PO-16**: Verificar las RN documentadas con datos de producción antes de switchover.
- [ ] **PO-17**: Plan de deprecation de las 3 OCX dependencies (TreeCtrl, ImageListCtrl, plus posibles otros).

---

## 7. Tickets derivables (preview)

> NO crear issues todavía. Estos nacen de la épica al final del ciclo de revisión.

### ActiveX y Twin forms (D158, D159)

- **TK-LZ-1**: [MIGRATION] Eliminar MSComctlLib.TreeCtrl.2 + ImageListCtrl.2 — reemplazar por react-arborist + ImageList componente.
- **TK-LZ-2**: [REFACTOR] Consolidar 2 pares de twin forms (FormVideosGestionArbolAplicacion↔Datos, FormVideosGestion↔Usuarios) en componentes parametrizables.
- **TK-LZ-3**: [REFACTOR] Reemplazar reach-up-to-grandparent coupling (`Me.Parent.Parent.Controls`) por outlet pattern.

### Testing y DI (D161)

- **TK-LZ-4**: [TESTING] Refactor de globals (`m_ObjVideoActivo`, `m_ObjAplicacionActivo`) a request-scoped services — pre-requisito para tests.
- **TK-LZ-5**: [TESTING] Cobertura de tests > 70% en módulo lanzaderas.
- **TK-LZ-6**: [TESTING] E2E test del flujo login → menú → CRUD aplicación + video.

### WIP forms (D163)

- **TK-LZ-7**: [CLEANUP] Auditar y completar `Form_FormObtenerContraseña` (orphan event bindings).
- **TK-LZ-8**: [CLEANUP] Eliminar `Form_Formulario1` (WIP shell con dead buttons).

### Cross-app y tooling (D164, D165)

- **TK-LZ-9**: [INFRASTRUCTURE] Pull de binarios `Lanzadera.accdb` + `Lanzadera_Datos.accdb` desde R2.
- **TK-LZ-10**: [TOOLING] `codegraph sync` con scope lanzaderas post-fix de #1408.
- **TK-LZ-11**: [NAMING] Renombrar `Form_Formulario1` en migración a `LanzaderaFormulario1` para evitar colisión con Expedientes.

### Forms y migración

- **TK-LZ-12**: [MIGRATION] Migrar 27 clases + 89 módulos de lanzaderas a services HTTP.
- **TK-LZ-13**: [UX] Consolidar `MENU1_Click` (duplicado en 3 menús) en componente `<AppMenu role>`.
- **TK-LZ-14**: [UX] `Form_FormAplicaciones{Oficina,FueraOficina,Prueba}` migrar a filter chips.
- **TK-LZ-15**: [UX] `Form_FormUsuariosConectados` migrar a SSE/WebSocket server-push.

### Datos y operación

- **TK-LZ-16**: [DATA] Migrar binarios staging a S3-compatible con versioning.
- **TK-LZ-17**: [SECURITY] Modernizar hash de contraseñas si D109 aplica.
- **TK-LZ-18**: [MIGRATION] Reemplazar IDs anti-patrón Access (sequences `TbID*`) por IDENTITY/SERIAL.

---

## Anexo · Tabla de fuentes

| Fuente | Aporta |
|---|---|
| engram topic_key `lanzaderas/walkthrough-2026-08-07` | Consolidado de 5 walkthroughs paralelos G1..G5 |
| `docs/03-aplicaciones/lanzadera/walkthrough-G1.json` | 5 forms (Login + Menús + Perfil) — método v4, 29020 bytes |
| `docs/03-aplicaciones/lanzadera/walkthrough-G2.json` | 7 forms (Aplicaciones) — método v4, 24484 bytes |
| `docs/03-aplicaciones/lanzadera/walkthrough-G3.json` | 6 forms (Usuarios) — método v4, 16603 bytes |
| `docs/03-aplicaciones/lanzadera/walkthrough-G4.json` | 7 forms (Videos) — método v4, 22849 bytes — 5 critical findings |
| `docs/03-aplicaciones/lanzadera/walkthrough-G5.json` | 3 forms (Cross-cutting) — método v4, 14628 bytes |
| `data/staging/lanzaderas/src/forms/` (28 .form.txt + 28 .cls) | Source tree exportado |
| `data/staging/lanzaderas/src/classes/` | Clases de dominio (27) |
| `data/staging/lanzaderas/src/modules/` | Módulos (89) |
| `C:\00repos\codigo\00_LANZADERA\staging` | Fuente READ-ONLY |

## Anexo · Decisiones referenciadas (D5-D167)

| Decisión | Aplicación a Lanzaderas |
|---|---|
| D8 (hexagonal global) | Toda la migración |
| D14 (esquema por módulo) | Módulo `lanzaderas` en PostgreSQL |
| D16 (object storage S3-compatible) | Anexos |
| D44-D46 (autorización + capabilities) | Roles Admin/Calidad/Usuario (D158 + 3 menús) |
| D66-D67 (stack) | Backend Python + HTMX |
| D68 (monolito modular) | Migración a módulo dedicado |
| D82 (Expand and Contract) | Backward-compatible con backend legacy |
| D86-D87 (cross-app con NoConformidades) | Patrón compartido via `getdbLanzadera()` |
| D102 (booleanos Text(2)) | Cleanup de captions |
| D109 (password plaintext hardcodeado) | Si aplica, migrar a bcrypt/argon2 (TK-LZ-17) |
| D132 (XApp HTTP/JSON handshake) | Si hay handshake con Expedientes |
| D140 (formato del campo Estado) | Aplica a Form_FormUsuarioGeneral |
| **D155 (methodology v4)** | **Walkthrough con analyze_form_layout RESUELTO + 2 tools skipped (#1408, #1412)** |
| **D156 (binarios NO staging)** | **`Lanzadera.accdb` + `Lanzadera_Datos.accdb` faltan del repo** |
| **D157 (codegraph no indexa lanzaderas)** | **Subdir no indexado — cross-project noise en blast radius** |
| **D158 (ActiveX legacy OCX)** | **3 forms usan MSComctlLib.TreeCtrl/ImageListCtrl — eliminar en migración** |
| **D159 (twin forms G4)** | **2 pares confusos — consolidar** |
| **D160 (grandparent coupling)** | **`Me.Parent.Parent.Controls` en Form_FormVideosGestionArbolVideoDatos** |
| **D161 (zero test coverage + global state)** | **Refactor a request-scoped services antes de tests** |
| **D162 (Marco17 false positive)** | **Limitación AABB lint — OptionGroup → radio web** |
| **D163 (WIP forms G5)** | **Orphan bindings + dead buttons — decidir completar/eliminar** |
| **D164 (codegraph lanzaderas)** | **Sync post-fix #1408** |
| **D165 (binarios staging)** | **Pull desde R2** |
| **D166 (global state)** | **m_ObjVideoActivo, m_ObjAplicacionActivo — refactor a services** |
| **D167 (cross-app ambiguity)** | **Form_Formulario1 colisión Lanzaderas/Expedientes — renombrar** |

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
- [x] D155 particular: metodología v4 aplicada con analyze_form_layout RESUELTO en 2.36.2
- [x] D156-D157 particular: filesystem/codegraph constraints documentados
- [x] D158-D162 particular: 5 critical findings G4 con acción concreta

## Siguiente paso

Revisión con el equipo. Esta épica se revisa junto con las otras 7. El siguiente paso es **Expedientes** (la octava y última) — 302 forms, el set más documental, ~69 archivos en docs/.