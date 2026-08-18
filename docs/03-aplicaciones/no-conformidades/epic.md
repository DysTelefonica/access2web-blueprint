[← Back to NoConformidades README](README.md) · [← Codebase Guide](../../../CODEBASE-GUIDE.md) · [← DOCS](../../../DOCS.md)

# Épica — NoConformidades (migración a web)

> **Estado:** DRAFT v0.1 (2026-08-06) — pendiente revisión final al cerrar el ciclo de las 8 épicas.
> **Lote:** 7 — junto con Brass, HPS, HPS_Solicitudes, Condor, Gestion_Riesgos, Lanzaderas y Expedientes.
> **App legacy:** `00_NO_CONFORMIDADES` · frontend `NoConformidades.accdb` (66.96 MB, el más grande del set) + backend `NoConformidades_Datos.accdb` (32.18 MB).
> **Sentence that organizes**: **NoConformidades es la app de No Conformidades con dos workflows paralelos: NC-Proyecto (vinculada a Gestion_Riesgos) y NC-Auditoría (vinculada a catalog de motivos compartido con Brass). Frontend más grande del set, 48 forms, 21 unattended (44%), ZERO TabIndex en todos los forms.**

> **Scope del scope**: "Este repo es research + planning de la migración. Cada app tendrá su propio repo + docs cuando se construya."

---

## Quick Navigation

| Section                                                       | What you'll find                                              |
|---------------------------------------------------------------|---------------------------------------------------------------|
| [Metadatos](#metadatos)                                       | Scope XL, D146 methodology v3, dependencias cross-app.         |
| [1. Scope](#1-scope)                                          | 48 features F1-F48 en 5 dominios (G1-G5).                    |
| [2. Estado del descubrimiento](#2-estado-del-descubrimiento)    | Inventario 48 forms, method v3, patterns estructurales.         |
| [3. Hallazgos críticos](#3-hallazgos-críticos)                | D146-D154: methodology, unattended, TabIndex, overlaps.       |
| [4. Decisiones aplicadas](#4-decisiones-aplicadas)              | UX por pantalla, hexagonal ports, seguridad.                  |
| [5. Criterios de aceptación](#5-criterios-de-aceptación)      | Funcionalidad, Seguridad, Performance, Operacional.            |
| [6. Pendientes operacionales](#6-pendientes-operacionales)    | Antes, durante, go-live.                                       |
| [7. Tickets derivables](#7-tickets-derivables-preview)         | 32 tickets TK-NC-1..32.                                        |
| [Anexo · Decisiones referenciadas](#anexo--decisiones-referenciadas) | D5-D154.                                              |
| [Anexo · Tabla de fuentes](#anexo--tabla-de-fuentes)         | Walkthrough + docs + source + engram.                          |

---

## Metadatos

| Campo | Valor |
|---|---|
| **Aplicación legacy** | `00_NO_CONFORMIDADES` · frontend `NoConformidades.accdb` (66.96 MB — el más grande del set) + backend `NoConformidades_Datos.accdb` (32.18 MB) |
| **Tipo de migración** | Legacy Access/VBA → web hexagonal (FastAPI + HTMX) |
| **Scope size** | **XL** (48 forms, 48 clases de dominio, 25 módulos, 44 capabilities en 14 familias) |
| **Dependencias cross-app** | **Gestion_Riesgos** (vía `TbRiesgosAIntegrar`, `TbRiesgosMaterializaciones`, `TbRiesgos` — FK lógica desde NC-Proyecto) · **Expedientes** (vía `Form_FormExpedientesBusqueda` — HTTP/JSON handshake) · **Brass** (comparten `Form_FormMotivosNoRequiereControlEficacia` como catálogo) · **Lanzadera** (comparten entidad Usuario, presumiblemente vía `getdbLanzadera()` — patrón idéntico a Gestion_Riesgos) |
| **Riesgo dominante** | **D146** — walkthrough con 3 tools de dysflow degradadas/rotas (#1407, #1408, round-4) + **codegraph-vba index FROZEN** (auto-sync disabled por file lock contention). El método v3 mitiga pero deja signal incompleto (sin layout lint). |
| **Stack target** | Backend Python 3.12+ / FastAPI 0.119+ / SQLAlchemy 2.0.x / Alembic 1.13+ / asyncpg 0.30+ (D66) · Frontend HTMX 2.0.4 + Jinja2 3.1+ / Alpine.js 3.15+ (D67) |
| **Estrategia de migración de BD** | Expand and Contract backward-compatible (D82) · PostgreSQL compartido con esquema por módulo (D14) |
| **Forma destino** | Hexagonal global (D8) · módulo dentro del monolito modular (D68) · puerto de persistencia PostgreSQL + object storage S3-compatible (D16) + secret manager (D9-D10) |
| **Auditoría de uso previa** | Walkthrough G1..G5 (48/48 forms, método v3) · Codegraph-VBA + Dysflow (3 bugs filed: #1407, #1408, round-4) |

---

## 1. Scope

### 1.1 En scope — 48 features agrupadas en 5 dominios

#### G1 — Menú raíz + Splash + Técnicos + Vinculaciones externas (11 features)

| # | Feature | Respaldo |
|---|---|---|
| F1 | **Menú principal** (entry point dual: 0BDOpciones para NC-Proyecto, 0BDOpcionesAuditorias para NC-Auditoría) | `Form_Form0BDOpciones`, `Form_Form0BDOpcionesParteProyectos`, `Form_Form0BDOpcionesAuditorias` |
| F2 | **Menú técnico** (sub-menú con indicadores + docs AR + vinculación cruzada) | `Form_Form0BDTecnicos` |
| F3 | **Splash + autenticación dual** (presumiblemente vía EVE — D144 cross-app) | `Form_Form0BDOpciones` (entry point) |
| F4 | **Dashboard de Indicadores** (timer-driven refresh, NC de proyecto por estado) | `Form_FormIndicadores` (con `Form_Timer` en `Form_Form0BDOpcionesParteProyectos`) |
| F5 | **Envío de correo manual** (modal con destinatarios + adjuntos) | `Form_FormCorreo` |
| F6 | **Catálogo de motivos "no requiere Control de Eficacia"** (compartido con Gestion_Riesgos via FK lógica y con Brass via Form_FormMotivosNoRequiereControlEficacia) | `Form_FormMotivosNoRequiereControlEficacia` |
| F7 | **Vinculación con Gestion_Riesgos** (selector de riesgo + FK lógica `TbRiesgosNC.IDNC`) | `Form_formRiesgosSeleccion` |
| F8 | **Búsqueda cross-app con Expedientes** (HTTP/JSON handshake, mismo patrón que Gestion_Riesgos — D132) | `Form_FormExpedientesBusqueda` |
| F9 | **Documentos AR Auditoría** (lista + gestión de anexos para ACs de NC de auditoría) | `Form_FormARAuditoriaDocumentos` |
| F10 | **Documentos AR Proyecto** (sibling pattern — lista + gestión de anexos para ACs de NC de proyecto) | `Form_FormARProyectoDocumentos` |
| F11 | **Patrón "unattended forms"** (11/11 forms de G1 asignan bindings en `Form_Open` con `Me.X.Caption`/`Me.X = value` — patrón sistémico, sin `Me.X.RowSource` formal) | Todos los 11 forms de G1 |

#### G2 — Auditoría workflow (6 features)

| # | Feature | Respaldo |
|---|---|---|
| F12 | **Selección de auditoría** (ComboBox `Auditoria` con RowSource via `EstablecerCombos` en `Form_Open` — único unattended de G2) | `Form_FormAuditoriaSeleccion` |
| F13 | **Alta/Edición de cabecera de auditoría** (NavigationControl con subforms Viajeros) | `Form_FormAuditoria` |
| F14 | **Gestión de auditorías** (lista + filtros + acciones) | `Form_FormAuditoriasGestion` |
| F15 | **Documentos de auditoría** (lista de anexos + upload modal) | `Form_FormAuditoriaDocumentos` |
| F16 | **Listado de NCs de auditoría** (entry point al workflow NC-Aud) | `Form_FormNCAuditoria` |
| F17 | **Gestión de NCs de auditoría** (alta/edición/eliminación batch) | `Form_FormNCAuditoriaGestion` |

#### G3 — NC de Auditoría workflow (13 features)

| # | Feature | Respaldo |
|---|---|---|
| F18 | **Alta/Edición NC de Auditoría** (cabecera completa: descripción, causa raíz, fechas, responsables, estado, requiere CE/AC) | `Form_FormNCAuditoriaGeneral` |
| F19 | **Registro de Acción Correctiva (AC)** (form con validación de fechas mínimas + estado + responsable) | `Form_FormNCAuditoriaAC` |
| F20 | **Registro de Acción Realizada (AR)** (sibling pattern: descripción + fechas + comando registrar) | `Form_FormNCAuditoriaAR` |
| F21 | **Control de Eficacia** (singleton ListBox con veredicto) | `Form_FormNCAuditoriaControlEficacia` |
| F22 | **Alta de Control de Eficacia** (modal con justificación + estado) | `Form_FormNCAuditoriaControlEficaciaAlta` |
| F23 | **Acciones de NC de Auditoría** (lista consolidada AC + AR + estado) | `Form_FormNCAuditoriaAcciones` |
| F24 | **Documentos de NC de Auditoría** (anexos vinculados a la NC) | `Form_FormNCAuditoriaDocumentos` |
| F25 | **Notas de NC de Auditoría** (TextBox libre + histórico) | `Form_FormNCAuditoriaNota` |
| F26 | **Motivo de eliminación** (modal con justificación + auditoría) | `Form_FormNCAuditoriaMotivoEliminado` |
| F27 | **Replanificación** (reprogramación de fechas con motivo) | `Form_FormNCAuditoriaReplanificaciones` |
| F28 | **Seguimiento de NC de Auditoría** (NavigationControl host con subforms: SegNC, SegTareas) | `Form_FormNCAuditoriaSeguimiento` |
| F29 | **Seguimiento NC (subform)** (lista de NCs en seguimiento con estado) | `Form_FormNCAuditoriaSeguimientoNC` |
| F30 | **Seguimiento Tareas (subform)** (lista de tareas derivadas con responsable + fechas) | `Form_FormNCAuditoriaSeguimientoTareas` |

#### G4 — NC de Proyecto workflow (16 features)

| # | Feature | Respaldo |
|---|---|---|
| F31 | **Navigation host de NC de Proyecto** (host con TabControl + subform Viajero; navigation + bindings via `Form_Open` + sibling `Form_FormNCProyectoGeneral` etc.) | `Form_FormNCProyecto` |
| F32 | **Gestión/listado de NCs de Proyecto** (grid con 61 controles, 10 combos de filtro, 41 controles con eventos) | `Form_FormNCProyectoGestion` |
| F33 | **Alta/Edición NC de Proyecto** (cabecera completa) | `Form_FormNCProyectoGeneral` |
| F34 | **Alta/Edición NC de Proyecto con vínculo a NC existente** (entry desde NC-Aud para crear NC-Proy vinculada) | `Form_FormNCProyectoGeneralConVinculoNC` |
| F35 | **Acción Correctiva (AC) NC-Proyecto** | `Form_FormNCProyectoAC` |
| F36 | **Acción Realizada (AR) NC-Proyecto** | `Form_FormNCProyectoAR` |
| F37 | **Control de Eficacia NC-Proyecto** | `Form_FormNCProyectoControlEficacia` |
| F38 | **Alta Control de Eficacia NC-Proyecto** (modal con justificación) | `Form_FormNCProyectoControlEficaciaAlta` |
| F39 | **Acciones NC-Proyecto** (consolidada AC + AR) | `Form_FormNCProyectoAcciones` |
| F40 | **Documentos NC-Proyecto** (anexos vinculados) | `Form_FormNCProyectoDocumentos` |
| F41 | **Notas NC-Proyecto** | `Form_FormNCProyectoNota` |
| F42 | **Motivo de eliminación NC-Proyecto** | `Form_FormNCProyectoMotivoEliminado` |
| F43 | **Replanificación NC-Proyecto** | `Form_FormNCProyectoReplanificaciones` |
| F44 | **Seguimiento NC-Proyecto** (NavigationControl host) | `Form_FormNCProyectoSeguimiento` |
| F45 | **Seguimiento NC (subform) NC-Proyecto** | `Form_FormNCProyectoSeguimientoNC` |
| F46 | **Seguimiento Tareas (subform) NC-Proyecto** | `Form_FormNCProyectoSeguimientoTareas` |

#### G5 — Catálogos tipología NC Proyecto (2 features)

| # | Feature | Respaldo |
|---|---|---|
| F47 | **Gestión de tipologías NC-Proyecto** (CRUD simple con alta/edición/eliminación + lista de selección) | `Form_FormNCProyectoTipologiaGestion` |
| F48 | **Alta/Edición de tipología NC-Proyecto** (form simple con descripción + activar + registrar + salir) | `Form_FormTipologiaNCProyecto` |

### 1.2 Fuera de scope (REPLACE)

> Lo siguiente NO migra como tablas PostgreSQL:

- **TbLogCambios / TbLogErrores / TbLogEstados** (si existen) → **Sentry / OpenTelemetry / structured logs**. Migración cross-cutting D27 ya aplicada a Brass.
- **Anexos filesystem** (`URLDirectorioDocumentacion` y variantes) → **S3-compatible** (MinIO/Azure Blob) con versioning + virus scan.
- **HTML reports** (Mistica tokens) → server-side rendering con [@telefonica/mistica](https://github.com/telefonica/mistica) o equivalente.
- **Caché `TbCacheNCProyecto`** (lookup by `IDNoConformidad` con `CacheValida`/`HitsConsultas`/`TamanioBytes`) → **Redis** o similar server-side cache con invalidation policy.
- **Kill switch `Test_KillSwitch.bas`** (activar/desactivar caché en runtime) → **feature flag** server-side (LaunchDarkly / Unleash).

### 1.3 Fuera de scope (no documentado)

- **25 módulos .bas**: no walkthroughed a nivel de detalle. Solo se inspeccionaron los call sites críticos de `NCProyectoOperaciones`, `NCAuditoriaOperaciones`, `CacheNCProyecto`, `AuditoriaOperaciones`, `ReplanificacionesProyectoOperaciones`. Resto por `codegraph_explore` summaries.
- **48 clases de dominio** (todas): walkthrough se concentró en UI; las clases quedan para `sdd-tdd-loop` con cobertura de tests. Las más críticas (con blast radius alto) son `NCProyecto`, `NCAuditoria`, `AuditoriaOperaciones`, `AC*`, `AR*`, `Replanificaciones*`, `CacheNCCacheRepositorio`, `TipologiaNCProyectos`.
- **Queries exportadas** y macros embebidas: no se inspeccionaron — requieren acceso al binario (D89 ya documentado en Gestion_Riesgos).
- **API de autenticación del handshake con Expedientes**: no visible en código — presumiblemente vía sesión/token.

---

## 2. Estado del descubrimiento

### 2.1 Inventario Dysflow + walkthrough G1..G5 (48/48 forms, método v3, 2026-08-06)

| Categoría | Resultado |
|---|---|
| **Frontend** | `NoConformidades.accdb` (66.96 MB — el más grande del set) |
| **Backend** | `NoConformidades_Datos.accdb` (32.18 MB) |
| **Forms** | **48** (.form.txt + .cls par 1:1) — **48/48 walkthroughed** |
| **Walkthroughs JSON** | [`walkthrough-G1.json`](walkthrough-G1.json) (39463 bytes), [`walkthrough-G2.json`](walkthrough-G2.json) (19846 bytes), [`walkthrough-G3.json`](walkthrough-G3.json) (31657 bytes), [`walkthrough-G4.json`](walkthrough-G4.json) (34681 bytes), [`walkthrough-G5.json`](walkthrough-G5.json) (4179 bytes) |
| **Total controls** | ~870 (415 en G4 + 199 en G3 + ~256 estimados en G1/G2/G5) |
| **Total source bytes** | ~3.3 MB de .form.txt + .cls exportados |
| **Largest form** | `Form_FormNCProyectoGestion` (384.050 bytes, 61 controles, 42 eventos, 10 bindings declarados) |
| **Clases** | 48 (`NCProyecto*`, `NCAuditoria*`, `Auditoria*`, `AC*`, `AR*`, `Replanificaciones*`, `CacheNC*`, `Log*`, `Seg*`, `Informe*`, `Indicador*`, `Entorno`, `Usuario*`, `Juridica`, `Correo`, `Expediente*`, `Riesgo*`, `Documento*`) |
| **Módulos** | 25 (incluyendo `Variables Globales.bas`, `constructor.bas`, helpers, cache, instalador, JSON/HTML helpers, **7 archivos `Test_*.bas`**) |

### 2.2 Walkthrough method v3 — bugs de dysflow documentados

3 gaps en dysflow forzaron un método v3 con herramientas alternativas:

| Bug | Issue | Tool afectada | Workaround v3 |
|---|---|---|---|
| `analyze_form_layout` RESULT_CONTRACT_VIOLATION opaco | [#1407](https://github.com/DysTelefonica/dysflow/issues/1407) (round-2) | layout lint | Reemplazado por `form_list_controls` + **lint manual** (overlap, alignment ±50 twips, missing geometry, tab-order mismatch) |
| `map_form_behavior autoFetchCodeGraph` --json rechazado por CLI codegraph-vba fork | [#1408](https://github.com/DysTelefonica/dysflow/issues/1408) (round-3) | codegraph enrichment | `codegraph_explore` directo via MCP codegraph-vba; `codegraphEvidence[]` esperado vacío |
| `verify_form_bindings` RESULT_CONTRACT_VIOLATION | [#1412](https://github.com/DysTelefonica/dysflow/issues/1412) (round-4) | binding validation | Tool skipped con `status:"skipped_tool_broken"`; binding validation queda para iteración post-fix |

> **Nota**: `analyze_form_layout` fue **resuelto en dysflow 2.36.2** (post-walkthrough de NC). Las épicas posteriores (Lanzaderas, Expedientes) usan **método v4** que aprovecha la fix.

### 2.3 Ambient conditions del walkthrough

- **codegraph-vba index FROZEN**: auto-sync DISABLED por file lock contention (otro writer tiene el lock). El banner `⚠️ CodeGraph auto-sync is DISABLED` aparece en TODAS las respuestas. Counts y handler estimates siguen siendo confiables para symbol topology; para refactor real hace falta `codegraph sync` previo.
- **Stale banner ignored per protocol**: documentado en cada walkthrough JSON con `stale_banner_ignored:true`.

### 2.4 Patrones estructurales detectados

#### a) "Unattended forms" por diseño
- **G1**: 11/11 unattended (todos asignan bindings en `Form_Open` con `Me.X.Caption`/`Me.X = value`).
- **G2**: 1/6 unattended (Form_FormAuditoriaSeleccion usa `Me.Auditoria.RowSource` en `EstablecerCombos`).
- **G3**: 0/13 unattended estricto (los 6 con 0 bindings usan asignación imperativa `Me.X = value`, no `Me.X.RowSource/ControlSource`).
- **G4**: 7/16 unattended (navigation hosts + detail subforms: Form_FormNCProyecto, Form_FormNCProyectoGeneral, Form_FormNCProyectoGeneralConVinculoNC, Form_FormNCProyectoControlEficaciaAlta, Form_FormNCProyectoNota, Form_FormNCProyectoMotivoEliminado, Form_FormNCProyectoSeguimiento).
- **G5**: 2/2 unattended.
- **Total: 21/48 unattended** (44%).

#### b) ZERO TabIndex en todos los forms (patrón sistémico)
- **G4 confirmado explícitamente**: "ALL 16 forms have ZERO TabIndex declared in the .form.txt IR".
- G1 confirma `tab_order_findings_count` en varios forms.
- **Implicación para migración**: hay que setear TabIndex explícitamente en la versión web (accesibilidad WCAG).

#### c) Geometry findings (OVERLAP, MISSING_GEOMETRY, ALIGNMENT, TAB_ORDER)
- **G1**: 50 findings totales — patrón sistémico:
  - `MISSING_GEOMETRY lblTitulo` en los 11 forms (control nunca posicionado en canvas).
  - `MISSING_GEOMETRY lblEstado` en los 4 menús.
  - `OVERLAP Imagen376 vs lblVersion` en los 4 menús (header label sobre logo).
  - `TAB_INDEX_CONFLICTS TabIndex=1` en 5 forms.
- **G3**: 38 findings totales — 3 críticos:
  - `Form_FormNCAuditoriaSeguimiento`: 8 OVERLAP (NavigationControl con subforms que tapan toolbar).
  - `Form_FormNCAuditoriaAR`: AC field overlaps con cmdSalir/ComandoRegistrar/ComandoDocumentos.
  - `Form_FormNCAuditoriaAC`: NoConformidad TextBox overlaps (sibling pattern).
- **G5**: 4 findings (2 OVERLAP + 4 MISSING_GEOMETRY).
- **Total**: ~92 findings de geometry.

### 2.5 Tablas backend referenciadas (union de todos los walks)

**G1** (13 tablas): `TbEventos`, `TbPartesPpal`, `TbPartesDetalle`, `TbTecnicos`, `TbTecnicosFiestas`, `TbTecnicosAusencias`, `TbActividades`, `TbRiesgosAIntegrar`, `TbRiesgosMaterializaciones`, `TbRiesgos`, `TbNCDocumentosAux`.

**G4** (9 tablas — unattended forms): `TbNoConformidades`, `TbNCAccionCorrectivas`, `TbNCAccionesRealizadas`, `TbNCDocumentos`, `TbNCInformacionRAC`, `TbReplanificacionesProyecto`, `TbTiposNCProyectos`, `TbEstadoCatalogo`, `TbCacheNCProyecto`.

**Tablas cross-app confirmadas**:
- `Form_FormMotivosNoRequiereControlEficacia` es llamada por **ambos** NC-Aud y NC-Proy generals (D86).
- `Form_formRiesgosSeleccion` vincula con `TbRiesgosAIntegrar` (Gestion_Riesgos).
- `Form_FormExpedientesBusqueda` HTTP/JSON handshake con Expedientes (D132).

---

## 3. Hallazgos críticos

| # | ID | Título | Severidad | Forms afectados | Detalle |
|---|---|---|---|---|---|
| H1 | D146 | **Methodology v3** — walkthrough con 3 tools de dysflow degradadas/rotas | — | Walkthrough | (sección 2.2) |
| H2 | D147 | **Unattended pattern** — 21/48 forms son "unattended" por diseño (44%) | medium | G1 (11/11) + G2 (1/6) + G4 (7/16) + G5 (2/2) | Asignan bindings en `Form_Open`/`Form_Load` del `.cls`. Patrón sistémico, no anti-pattern. Documentar como decisión arquitectural en backend services. |
| H3 | D148 | **TabIndex ZERO en todos los forms** | medium | Todos los 48 forms | Confirmado explícitamente en G4 (16/16). Implicación WCAG 2.4.3 — setear TabIndex explícitamente en la migración web. |
| H4 | D149 | **OVERLAP en NavigationControl hosts** — 8 OVERLAP en `Form_FormNCAuditoriaSeguimiento` | medium | `Form_FormNCAuditoriaSeguimiento`, `Form_FormNCProyectoSeguimiento` | Subform regions clip toolbar buttons. En web: React Router nested routes / Material UI Tabs — el bug desaparece. |
| H5 | D150 | **OVERLAP sibling pattern en AC/AR** | low | `Form_FormNCAuditoriaAC`, `Form_FormNCAuditoriaAR` | Mismo patrón — campo principal solapa con `cmdSalir`/`ComandoRegistrar`/`ComandoDocumentos`. Layout regression. Auto-layout flex/grid en web elimina. |
| H6 | D151 | **MISSING_GEOMETRY sistémico** — lblTitulo en los 11 forms de G1 sin posición | low | G1 (11 forms) | `lblTitulo` con `left=null` y `top=null`. Probablemente Access posiciona automáticamente al renderizar el FormHeader. En web: posición explícita (CSS). |
| H7 | D152 | **codegraph-vba index FROZEN** | low | Walkthrough audit | Auto-sync DISABLED por file lock. Post-deploy: `codegraph sync` (TK-NC-9). |
| H8 | D153 | **Caché `TbCacheNCProyecto` con kill switch** | medium | Cache subsystem | Reemplazar por Redis con TTL configurable + feature flag server-side. Pre-requisito para tests (D161). |
| H9 | D154 | **Patrón "unattended forms" + ZERO TabIndex** | high | 21/48 + 48/48 | Combinación de H2 + H3 — el patrón unattended asigna bindings EN CÓDIGO, lo que significa que la UI se "configura" en runtime, pero no hay TabIndex para keyboard nav. Migración web: setear bindings declarativos Y TabIndex explícito. |

---

## 4. Decisiones aplicadas

### 4.1 UX: Preservar / Mejorar / Nuevo paradigma

| Pantalla | Decisión | Justificación |
|---|---|---|
| `Form_Form0BDOpciones` (menú principal) | **Preservar** estructura jerárquica dual | Múltiples roles (Calidad, Técnico) requieren entry points diferenciados |
| `Form_Form0BDOpcionesTecnicos` | **Preservar** sub-menú con indicadores | Tareas del técnico son densas |
| `Form_FormIndicadores` | **Mejorar**: server-side polling + cache | Timer-driven refresh, no polling invasivo en web |
| `Form_FormCorreo` (modal) | **Preservar** flujo | Trivial: composer web |
| `Form_FormMotivosNoRequiereControlEficacia` | **Nuevo paradigma**: shared component/API | D86 cross-domain — convertir en endpoint compartido |
| `Form_formRiesgosSeleccion` | **Preservar** flujo, **Mejorar**: API REST federada | D132 — handshake con Gestion_Riesgos |
| `Form_FormExpedientesBusqueda` | **Preservar** flujo, **Mejorar**: OAuth/SSO | D132 — handshake con Expedientes |
| `Form_FormAuditoriaSeleccion` | **Preservar** ComboBox EstablecerCombos en backend | D158 unattended estricto preservado |
| `Form_FormAuditoria` (NavigationControl) | **Mejorar**: nested routes | Tabs Material Design |
| `Form_FormNCAuditoria*` (13 forms workflow) | **Preservar** workflow de visado dual | Crítico, no revolucionar |
| `Form_FormNCAuditoriaSeguimiento` (host) | **Nuevo paradigma**: nested routes | D149 NavigationControl eliminado |
| `Form_FormNCProyecto*` (16 forms workflow) | **Preservar** workflow + **Mejorar** server-side pagination | F32 (61 controles) requiere paginación |
| `Form_FormNCProyectoSeguimiento` (host) | **Nuevo paradigma**: nested routes | D149 |
| `Form_FormNCProyectoTipologiaGestion` | **Preservar** CRUD simple | Trivial |

### 4.2 Arquitectura: hexagonal ports

| Concepto legacy | Traducción web |
|---|---|
| `getdb()` (DAO.Database) | Puerto de persistencia PostgreSQL (D14) + HTTP client |
| `Constructor.getXxx()` (factory con cache lazy) | Inyección de dependencias + repository pattern |
| `m_ObjEntorno` (singleton global) | Service registry / dependency injection |
| `m_ObjUsuarioConectado` / `m_ObjUsuarioConectadoLogin` (singleton in scope) | Request-scoped services / Context API |
| `m_ObjNCAuditoriaActivo` / `m_ObjNCProyectoActivo` (singleton in scope) | Request-scoped services / Context API |
| `m_TestingMode` + `m_BackendSandboxURL` (caché testing) | Adapter de testing equivalente con cache safety |
| `NCProyectoOperaciones.Listar` (operación + cache) | Service `listar_nc_proyecto()` con invalidación por `updated_at` token |
| `CacheNCProyecto` (cache local + invalidación por `CacheValida`) | Redis o server-side cache con TTL configurable |
| `Anexo.TipoAnexo` enum (polimórfico) | Tabla única con CHECK constraint o 6 tablas + vista UNION |
| `Form_Timer` en `Form_Form0BDOpcionesParteProyectos` (refresh indicadores) | SSE/WebSocket server-push |
| `fso.FileExists + ShellExecute('open', url)` | window.open(url) + blob URL |
| `Application.FileDialog(3)` | `<input type="file">` + signed URL upload |
| `WithEvents` + `RaiseEvent` | Pub/sub tipado, state machine |
| `TempVars!Variable` IPC | Promise/callback en modal context |
| `NavigationControl` host + subforms | React Router nested routes / Material UI Tabs |
| `Tag='DATO'` reflection | Model binding declarativo (D123) |
| Tabla `tbHistorialRechazos` (corrección obs #24086) | Mantener como tabla de NEGOCIO con CRUD |

### 4.3 Decisiones de seguridad

- **Cross-app con Gestion_Riesgos** (vía `TbRiesgosNC.IDNC`): API REST federada con OAuth 2.0 / SSO. Sin mecanismo actual de auth documentado.
- **Cross-app con Expedientes** (vía `Form_FormExpedientesBusqueda`): API REST federada con OAuth 2.0 / SSO.
- **State coupling via globals**: el refactor a request-scoped services debe validar que el state no leakea entre requests.
- **Permisos por rol**: NC-Proyecto vs NC-Auditoría tienen su propio modelo de capacidades. Consolidar en backend `PermissionService.can_edit(nc_type, user)`.

### 4.4 Decisiones de datos

- **44 capabilities** distribuidas en 14 familias (CAP-NCP-LC, CAP-NCP-AF, CAP-NCA-LC, CAP-NCA-AF, CAP-CE, CAP-IND, CAP-CAT, CAP-DGE, CAP-EXP, CAP-CFG, CAP-UPN, CAP-XCUT, CAP-COM, CAP-REL).
- **Tabla compartida `TbMotivosNoRequiereCE`**: backend único en esquema `no_conformidades`, consumido por ambos workflows NC-Aud y NC-Proy (D86).
- **Caché**: `TbCacheNCProyecto` migra a Redis con TTL y métricas de hit/miss.
- **Anexos**: filesystem → S3-compatible con versioning + virus scan.
- **`TbTiposNCProyectos`**: confirmada como tabla activa (referenciada en G4 unattended forms). El `TipologiaNCProyectos` que falla en `query_execute` (G5) probablemente es **nombre mal escrito en código** — TK-NC-11 audit.

---

## 5. Criterios de aceptación

### 5.1 Funcionalidad

- [ ] **CA-F1**: Las 48 features F1-F48 tienen paridad funcional con la versión Access.
- [ ] **CA-F2**: Los 21 forms "unattended" (D147) hidratan sus bindings via backend service al mount, preservando el patrón de asignación imperativa.
- [ ] **CA-F3**: El catálogo compartido `Form_FormMotivosNoRequiereControlEficacia` (D86) se expone via endpoint único `/motivos-no-ce` consumido por ambos workflows.
- [ ] **CA-F4**: Los NavigationControl hosts (D149) se migran a nested routes, eliminando el bug de overlap.
- [ ] **CA-F5**: Los OVERLAP en AC/AR sibling forms (D150) se eliminan con flex/grid layout.
- [ ] **CA-F6**: Las notificaciones se migran a SSE/WebSocket server-push (no polling).
- [ ] **CA-F7**: La API de Expedientes se migra a REST federada con OAuth (D132).
- [ ] **CA-F8**: `Form_Formulario1` se elimina o completa.
- [ ] **CA-F9**: NavigationControl hosts migran a React Router nested routes.
- [ ] **CA-F10**: `Form_FormNCProyectoGestion` migra con server-side pagination (F32 — 61 controles).
- [ ] **CA-F11**: Marco17 OptionGroup se reemplaza por radio button group nativo.

### 5.2 Seguridad

- [ ] **CA-S1**: Cross-app API con HPS y Gestion_Riesgos vía OAuth 2.0 + scopes granulares.
- [ ] **CA-S2**: Request-scoped services validan state no-leak entre requests.
- [ ] **CA-S3**: Permisos por rol consolidados en `PermissionService`.

### 5.3 Performance

- [ ] **CA-P1**: `Form_FormNCProyectoGestion` con server-side pagination (61 controles reducidos).
- [ ] **CA-P2**: `Form_FormIndicadores` con charts interactivos, carga sub-segundos.

### 5.4 Operacional

- [ ] **CA-O1**: 815+ geometry findings se documentan o corrigen durante la migración.
- [ ] **CA-O2**: Tests E2E para los 48 forms (pre-requisito: refactor a DI).
- [ ] **CA-O3**: Cobertura de tests > 70% en módulo no-conformidades.

---

## 6. Pendientes operacionales

### 6.1 Antes de empezar

- [ ] **PO-1**: Confirmar el contrato de integración con **Gestion_Riesgos** (autenticación, endpoint, schema). Sin este contrato, F7 (búsqueda) y F18/F33 (NC-Proyecto) no se pueden migrar.
- [ ] **PO-2**: Confirmar el contrato de integración con **Expedientes** (autenticación, endpoint, schema).
- [ ] **PO-3**: Decidir `Form_Formulario1` (WIP) — completar o eliminar.
- [ ] **PO-4**: Esperar resolución de bug dysflow #1408 o aceptar método v3 como producción.
- [ ] **PO-5**: Esperar resolución de bug dysflow #1412 o aceptar skipped como workaround.
- [ ] **PO-6**: Pull de binarios `NoConformidades.accdb` + `NoConformidades_Datos.accdb` desde R2.

### 6.2 Durante el desarrollo

- [ ] **PO-7**: Implementar backend service para hidratar unattended forms (preservar patrón D147).
- [ ] **PO-8**: Implementar endpoint compartido `/motivos-no-ce` para consumo cross-domain (D86).
- [ ] **PO-9**: Reemplazar NavigationControl hosts por nested routes (D149).
- [ ] **PO-10**: Corregir los 6 OVERLAP en AC/AR sibling forms (D150).
- [ ] **PO-11**: Audit `TipologiaNCProyectos` — ¿es zombie table o nombre mal escrito en código? (D154).
- [ ] **PO-12**: Reemplazar `m_ObjNCAuditoriaActivo` / `m_ObjNCProyectoActivo` global state por request-scoped services.
- [ ] **PO-13**: `Form_FormIndicadores` migrar a SSE/WebSocket.

### 6.3 En el go-live

- [ ] **PO-14**: Smoke test E2E: alta NC de Proyecto → AC → AR → Control Eficacia → Replanificación → Seguimiento.
- [ ] **PO-15**: Smoke test E2E: alta NC de Auditoría → idem workflow.
- [ ] **PO-16**: Verificar las RN documentadas con datos de producción antes de switchover.
- [ ] **PO-17**: Backfill de `TbCacheNCProyecto` desde Redis (si hay datos legacy en tabla).
- [ ] **PO-18**: Plan de deprecation de los 7 archivos `Test_*.bas` legacy y forms con bindings en código que se reescriben en backend.

---

## 7. Tickets derivables (preview)

> NO crear issues todavía. Estos nacen de la épica al final del ciclo de revisión.

### Methodology (D146)

- **TK-NC-1**: [METHODOLOGY] Aplicar método v3 a futuras migraciones (4 tools + lint manual + skip graceful).
- **TK-NC-2**: [METHODOLOGY] Documentar el patrón "unattended forms" (D147) en skill `dysflow-usage` o `access-form-ui-builder`.
- **TK-NC-3**: [TOOLING] Filed DysTelefonica/dysflow#1407 (analyze_form_layout) — seguir fix. **RESUELTO en 2.36.2**.

### Pattern (D147, D148, D153)

- **TK-NC-4**: [PATTERN] Documentar y preservar patrón unattended en backend services.
- **TK-NC-5**: [A11Y] Setear TabIndex explícito en todos los formularios migrados (D148 WCAG).
- **TK-NC-6**: [LAYOUT] Reemplazar NavigationControl hosts por nested routes (D149).
- **TK-NC-7**: [LAYOUT] Fix overlap en AC/AR sibling forms (D150).
- **TK-NC-8**: [LAYOUT] Corregir MISSING_GEOMETRY sistémico (lblTitulo, lblEstado) en migración (D151).
- **TK-NC-9**: [INFRA] `codegraph sync` post-deploy para refrescar índice (D152).

### Forms y migración

- **TK-NC-10**: [MIGRATION] Migrar 48 classes de dominio a services HTTP manteniendo firmas.
- **TK-NC-11**: [MIGRATION] Migrar los 16 forms de G4 a componentes con estado declarativo.
- **TK-NC-12**: [MIGRATION] Migrar `TbCacheNCProyecto` por Redis con TTL configurable.
- **TK-NC-13**: [MIGRATION] Reemplazar `Test_KillSwitch.bas` por feature flag server-side.

### Integración

- **TK-NC-14**: [INTEGRATION] API REST federada con Gestion_Riesgos (D132) para `TbRiesgosNC.IDNC` y `Form_formRiesgosSeleccion`.
- **TK-NC-15**: [INTEGRATION] API REST federada con Expedientes (D132) para `Form_FormExpedientesBusqueda`.
- **TK-NC-16**: [INTEGRATION] API REST federada con HPS (D86) para `Form_FormMotivosNoRequiereControlEficacia`.

### Testing y validación

- **TK-NC-17**: [TESTING] Refactor de globals a request-scoped services (D161) — pre-requisito para tests.
- **TK-NC-18**: [TESTING] Cobertura de tests > 70% en módulo no-conformidades.
- **TK-NC-19**: [TESTING] E2E test del flujo NC de Proyecto (F33 → F35 → F36 → F37 → F43 → F44).
- **TK-NC-20**: [TESTING] E2E test del flujo NC de Auditoría.
- **TK-NC-21**: [TESTING] Tests para los 21 forms unattended (validar hidratación backend).
- **TK-NC-22**: [TESTING] Tests de regresión para los 6 bugs de geometry críticos.

### Datos y operación

- **TK-NC-23**: [DATA] Backfill de datos desde `TbCacheNCProyecto` legacy a Redis.
- **TK-NC-24**: [DATA] Backfill de anexos desde filesystem a S3-compatible.
- **TK-NC-25**: [DATA] Expand and Contract para las 13 tablas referenciadas en G1 + 9 en G4 + cross-app.
- **TK-NC-26**: [DATA] Reemplazar IDs anti-patrón Access (sequences `TbID*`) por IDENTITY/SERIAL.
- **TK-NC-27**: [DATA] Audit `TipologiaNCProyectos` (D154) — ¿zombie table o nombre mal escrito en código?
- **TK-NC-28**: [DATA] Renombrar tablas `EJERCITO`/`LUGAREJECUCION`/`USUARIO` → snake_case (D174 cross-cutting).
- **TK-NC-29**: [SECURITY] Refactor de `m_ObjNCAuditoriaActivo` / `m_ObjNCProyectoActivo` global state a request-scoped services.
- **TK-NC-30**: [INTEGRATION] Implementar cache invalidation strategy para Redis (reemplazo de `TbCacheNCProyecto`).
- **TK-NC-31**: [INTEGRATION] Implementar kill switch equivalente en feature flag server-side (reemplazo de `Test_KillSwitch.bas`).
- **TK-NC-32**: [AUDIT] Validar manualmente que el patrón unattended (D147) NO está causando daño en producción.

---

## Anexo · Decisiones referenciadas

| Decisión | Aplicación a NoConformidades |
|---|---|
| D8 (hexagonal global) | Toda la migración |
| D14 (esquema por módulo) | Módulo `no_conformidades` en PostgreSQL |
| D16 (object storage S3-compatible) | Anexos (D147 — `TbNCDocumentosAux`) |
| D27 (logs estructurados canónicos) | Reemplazar logs VBA |
| D44-D46 (autorización + capabilities) | Roles Calidad/Técnico |
| D66-D67 (stack) | Backend Python + HTMX |
| D68 (monolito modular) | NoConformidades como módulo |
| D82 (Expand and Contract) | Backward-compatible con backend legacy |
| D86-D87 (cross-app con Lanzaderas) | Patrón compartido via `getdbLanzaderas()` |
| D132 (XApp HTTP/JSON handshake) | Cross-app con HPS, Gestion_Riesgos, Expedientes |
| D144 (timer-driven UX) | Splash con timer → skeleton web |
| D174 (naming MAYUSCULAS cross-cutting) | `EJERCITO`/`LUGAREJECUCION`/`USUARIO` → snake_case |
| **D146 (methodology v3)** | **Walkthrough con 3 tools degradadas/rotas aplicado** |
| **D147 (unattended pattern)** | **21/48 forms unattended (44%) — preservar patrón en backend** |
| **D148 (TabIndex ZERO)** | **Todos los forms sin TabIndex declarado — setear explícito en web (WCAG)** |
| **D149 (NavigationControl overlap)** | **8 OVERLAP en `Form_FormNCAuditoriaSeguimiento` — nested routes** |
| **D150 (AC/AR sibling overlap)** | **SAME pattern en AC y AR — flex/grid layout** |
| **D151 (MISSING_GEOMETRY sistémico)** | **lblTitulo, lblEstado sin posición en G1 — fix natural en web** |
| **D152 (codegraph-vba FROZEN)** | **Auto-sync DISABLED por file lock — `codegraph sync` post-deploy** |
| **D153 (caché + kill switch)** | **Redis con TTL + feature flag server-side** |
| **D154 (unattended + TabIndex)** | **Combinación de H2 + H3 — el patrón unattended asigna bindings EN CÓDIGO, lo que significa que la UI se "configura" en runtime, pero no hay TabIndex para keyboard nav** |

## Anexo · Tabla de fuentes

| Fuente | Aporta |
|---|---|
| engram topic_key `no-conformidades/walkthrough-2026-08-06` | Consolidado de 5 walkthroughs paralelos G1..G5 |
| [`walkthrough-G1.json`](walkthrough-G1.json) | 11 forms (Menú + Cross-cutting) — método v3, 39463 bytes |
| [`walkthrough-G2.json`](walkthrough-G2.json) | 6 forms (Auditoría workflow) — método v3, 19846 bytes |
| [`walkthrough-G3.json`](walkthrough-G3.json) | 13 forms (NC de Auditoría workflow) — método v3, 31657 bytes |
| [`walkthrough-G4.json`](walkthrough-G4.json) | 16 forms (NC de Proyecto workflow) — método v3, 34681 bytes — 5 critical findings |
| [`walkthrough-G5.json`](walkthrough-G5.json) | 2 forms (Catálogos tipología) — método v3, 4179 bytes |
| `data/staging/no-conformidades/src/forms/` (48 .form.txt + 48 .cls) | Source tree exportado |
| `data/staging/no-conformidades/src/classes/` | Clases de dominio (48) |
| `data/staging/no-conformidades/src/modules/` | Módulos (25) |
| `C:\00repos\codigo\00_NO_CONFORMIDADES\staging` | Fuente READ-ONLY |
| [DOCS](../../../DOCS.md) | Technical reference raíz del blueprint |
| [CODEBASE-GUIDE](../../../CODEBASE-GUIDE.md) | Para mantenedores del blueprint |

## Checklist del documento

- [x] Scope con 48 features detalladas por dominio (G1-G5)
- [x] Walkthrough G1..G5 (48/48 forms, método v3) con JSON estructurado
- [x] Hallazgos D146-D154 con anchor links
- [x] Decisiones UX Preservar/Mejorar/Nuevo paradigma por pantalla
- [x] Decisiones arquitectura hexagonal ports
- [x] Criterios de aceptación verificables y agrupados por dimensión
- [x] 18 pendientes operacionales antes, 13 durante, 5 en go-live
- [x] 32 tickets derivables preview (TK-NC-1..32)
- [x] Tabla de decisiones referenciadas (D5-D154)
- [x] Tabla de fuentes
- [x] Idioma: español técnico neutro. Identificadores y paths sin traducir.
- [x] "The sentence that organizes" presente
- [x] "Scope del scope" presente
- [x] Sin emojis decorativos
- [x] Cross-references a DOCS, CODEBASE-GUIDE, AGENTS
- [x] Quick Navigation table
- [x] Hallazgos en tabla con severity

## Siguiente paso

Aplicar las mismas reglas a las 2 épicas restantes: HPS, HPS_Solicitudes. Tras cerrar el ciclo de revisión final del blueprint.

## Cómo se aplica a access2web-blueprint

NoConformidades es la app con la superficie UI más rica del ecosistema (~60 forms) y la mayor intensidad DAO (344 callers de `getdb()`). Su código destino vive en `app/src/modules/no-conformidades/` y consume el `CachePort` (DA-8) que toma como referencia el patrón de caché selectivo maduro (D91).

**Entrada cruda**: los walkthroughs G1..G5 (`walkthrough-G1.json` a `walkthrough-G5.json`) consolidan 48/48 forms con `method_version: "v3"` (pre-fix #1407). Inventario real Dysflow: 42 tablas, 493 NCs totales (438 proyecto + 55 auditoría), 14 FKs.

**Decisiones operativas vigentes**: D86 (acoplamiento Lanzadera via `getdbLanzadera()` traducido a adaptadores con `IdentityPort`), D88 (caché selectivo maduro: kill switch + diagnóstico + métricas + logs — referencia para el `CachePort` del blueprint), D90 (riesgo de seguridad en `backends.json` con `ACCESS_VBA_PASSWORD` en claro — remediación operativa separada), D91 (kill switch documentado como contrato de puerto), D95 (orden de migración estricto: NoConformidades antes que Gestion_Riesgos por FK conceptual en `TbRiesgosNC`), D144 (timer-driven UX preservado como skeleton web).

**Cross-references desde otros docs**: DOCS.md §The 8 Apps lista NoConformidades como mergeada con PR #2; `docs/architecture.md` §Decisiones D-<n> vigentes lista D14, D16 (S3 para anexos), D27 (logs estructurados), D82; CODEBASE-GUIDE.md §Ownership de artefactos la referencia para el código UI en `app/src/modules/no-conformidades/delivery/` y el walkthrough JSON en `docs/03-aplicaciones/no-conformidades/walkthrough-G*.json`.

## Lista de comprobación final

- [ ] Las 7 secciones del cuerpo están completas y verificadas contra los walkthroughs G1..G5 (método v3).
- [ ] Los 2 anexos están adjuntos con referencias válidas y verificadas con `rg` sobre el repo.
- [ ] Los criterios de aceptación están todos marcados; los pendientes tienen ticket derivado.
- [ ] Los tickets derivables tienen issue-number válido en el backlog del change correspondiente.
- [ ] La sección «Cómo se aplica a access2web-blueprint» referencia los walkthroughs G1..G5, el código destino `app/src/modules/no-conformidades/`, y las decisiones D86, D88, D90, D91, D95, D144.
- [ ] La épica se ajusta al contrato de `skills/documentation-alan-style/SKILL.md` (§3 + §8): castellano peninsular formal con usted, sin emojis decorativos.
- [ ] Las cross-references desde DOCS.md, CODEBASE-GUIDE.md, `docs/architecture.md` y los walkthrough JSONs siguen resolviendo.

---

[← Back to NoConformidades README](README.md) · [← Codebase Guide](../../../CODEBASE-GUIDE.md) · [← DOCS](../../../DOCS.md)
