# Épica — NoConformidades (migración a web)

> **Estado:** DRAFT v0.1 (2026-08-06) — pendiente revisión final al cerrar el ciclo de las 8 épicas.
> **App legacy:** NoConformidades (`C:\00repos\codigo\00_NO_CONFORMIDADES`) · frontend `NoConformidades.accdb` (66.96 MB, el más grande del set) + backend `NoConformidades_Datos.accdb` (32.18 MB).
> **Lote de discovery:** Lote 7 — junto con Brass, HPS, HPS_Solicitudes, Condor, Gestion_Riesgos, Lanzaderas y Expedientes.
> **Hallazgo dominante:** **D146 (methodology)** — walkthrough profundo requirió método v3 con 2 tools de dysflow saltadas (#1407, #1408) + 1 tool rota (round-4). El método v3 produce JSON estructurado y descubrimiento completo a pesar de los gaps.
> **Cross-refs:** engram topic_key `no-conformidades/walkthrough-2026-08-06` (consolidado de 5 walkthroughs paralelos G1..G5).

## Metadatos

| Campo | Valor |
|---|---|
| **Aplicación legacy** | `00_NO_CONFORMIDADES` · frontend `NoConformidades.accdb` (66.96 MB) + backend `NoConformidades_Datos.accdb` (32.18 MB) |
| **Tipo de migración** | Legacy Access/VBA → web hexagonal (FastAPI + HTMX) |
| **Scope size** | **XL** (48 forms, 48 clases de dominio, 25 módulos, 44 capabilities en 14 familias) |
| **Dependencias cross-app** | **Gestion_Riesgos** (vía `TbRiesgosAIntegrar`, `TbRiesgosMaterializaciones`, `TbRiesgos` — FK lógica desde NC-Proyecto) · **Expedientes** (vía `Form_FormExpedientesBusqueda` — HTTP/JSON handshake) · **Brass** (comparten `Form_FormMotivosNoRequiereControlEficacia` como catálogo) · **Lanzadera** (comparten entidad Usuario, presumiblemente vía `getdbLanzadera()` — patrón idéntico a Gestion_Riesgos) |
| **Riesgo dominante** | **D146** — walkthrough con 3 tools de dysflow degradadas/rotas (#1407, #1408, round-4) + **codegraph-vba index FROZEN** (auto-sync disabled por file lock contention). El método v3 mitiga pero deja signal incompleto (sin layout lint). |
| **Stack target** | Backend Python 3.12+ / FastAPI 0.119+ / SQLAlchemy 2.0.x / Alembic 1.13+ / asyncpg 0.30+ (D66) · Frontend HTMX 2.0.4 + Jinja2 3.1+ + Alpine.js 3.15+ (D67) |
| **Estrategia de migración de BD** | Expand and Contract backward-compatible (D82) · PostgreSQL compartido con esquema por módulo (D14) |
| **Forma destino** | Hexagonal global (D8) · módulo dentro del monolito modular (D68) · puerto de persistencia PostgreSQL + object storage S3-compatible (D16) + secret manager (D9-D10) |
| **Auditoría de uso previa** | ✅ Walkthrough G1..G5 (48/48 forms, método v3) · ✅ Codegraph-VBA + Dysflow (3 bugs filed: #1407, #1408, round-4) |

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

### 1.1 En scope (~48 features agrupadas en 5 dominios)

#### G1 — Menú raíz + Splash + Técnicos + Vinculaciones externas (~11 features)

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

#### G2 — Auditoría workflow (~6 features)

| # | Feature | Respaldo |
|---|---|---|
| F12 | **Selección de auditoría** (ComboBox `Auditoria` con RowSource via `EstablecerCombos` en `Form_Open` — único unattended de G2) | `Form_FormAuditoriaSeleccion` |
| F13 | **Alta/Edición de cabecera de auditoría** (NavigationControl con subforms Viajeros) | `Form_FormAuditoria` |
| F14 | **Gestión de auditorías** (lista + filtros + acciones) | `Form_FormAuditoriasGestion` |
| F15 | **Documentos de auditoría** (lista de anexos + upload modal) | `Form_FormAuditoriaDocumentos` |
| F16 | **Listado de NCs de auditoría** (entry point al workflow NC-Aud) | `Form_FormNCAuditoria` |
| F17 | **Gestión de NCs de auditoría** (alta/edición/eliminación batch) | `Form_FormNCAuditoriaGestion` |

#### G3 — NC de Auditoría workflow (~13 features)

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

#### G4 — NC de Proyecto workflow (~16 features)

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

#### G5 — Catálogos tipología NC Proyecto (~2 features)

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
| **Walkthroughs JSON** | `walkthrough-G1.json` (39463 bytes), `walkthrough-G2.json` (19846 bytes), `walkthrough-G3.json` (31657 bytes), `walkthrough-G4.json` (34681 bytes), `walkthrough-G5.json` (4179 bytes) |
| **Total controls** | ~870 (415 en G4 + 199 en G3 + ~256 estimados en G1/G2/G5) |
| **Total source bytes** | ~3.3 MB de .form.txt + .cls exportados |
| **Largest form** | `Form_FormNCProyectoGestion` (384.050 bytes, 61 controles, 42 eventos, 10 bindings declarados) |
| **Clases** | 48 (`NCProyecto*`, `NCAuditoria*`, `Auditoria*`, `AC*`, `AR*`, `Replanificaciones*`, `CacheNC*`, `Log*`, `Seg*`, `Informe*`, `Indicador*`, `Entorno`, `Usuario*`, `Juridica`, `Correo`, `Expediente*`, `Riesgo*`, `Documento*`) |
| **Módulos** | 25 (incluyendo `Variables Globales.bas`, `constructor.bas`, helpers, cache, instalador, JSON/HTML helpers, **7 archivos `Test_*.bas`**) |

### 2.2 Walkthrough method v3 — bugs de dysflow documentados

3 gaps en dysflow forzaron un método v3 con herramientas alternativas:

| Bug | Issue | Tool afectada | Workaround v3 |
|---|---|---|---|
| `analyze_form_layout` RESULT_CONTRACT_VIOLATION opaco | DysTelefonica/dysflow#1407 (round-2) | layout lint | Reemplazado por `form_list_controls` + **lint manual** (overlap, alignment ±50 twips, missing geometry, tab-order mismatch) |
| `map_form_behavior autoFetchCodeGraph` --json rechazado por CLI codegraph-vba fork | DysTelefonica/dysflow#1408 (round-3) | codegraph enrichment | `codegraph_explore` directo via MCP codegraph-vba; `codegraphEvidence[]` esperado vacío |
| `verify_form_bindings` RESULT_CONTRACT_VIOLATION | round-4 (pendiente) | binding validation | Tool skipped con `status:"skipped_tool_broken"`; binding validation queda para iteración post-fix |

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

### H1 (D146 — METHODOLOGY) — Walkthrough requirió método v3 con 3 tools de dysflow degradadas/rotas

**Síntoma**: el walkthrough planeado (4 tools en orden) no funciona out-of-the-box. `analyze_form_layout` falla con `RESULT_CONTRACT_VIOLATION` opaco (#1407). `map_form_behavior autoFetchCodeGraph:true` falla porque el adapter dysflow invoca `codegraph-vba.cmd explore --json` con flag `--json` que el CLI codegraph-vba fork rechaza (#1408). `verify_form_bindings` falla con `RESULT_CONTRACT_VIOLATION` (round-4 pendiente).

**Acciones** (TK-NC-1, TK-NC-2, TK-NC-3 — issues ya filed):
- ✅ Filed DysTelefonica/dysflow#1407 (round-2).
- ✅ Filed DysTelefonica/dysflow#1408 (round-3).
- ⏳ Round-4 (verify_form_bindings) pendiente de filear.
- ✅ Método v3 implementado y aplicado a los 5 walkthroughs: `codegraph_explore` directo + `analyze_form_ui` + `form_list_controls` (reemplazo de analyze_form_layout) + **lint manual de geometry** (4 kinds) + `map_form_behavior` con `codegraphEvidence:[]` aceptado + `verify_form_bindings` skipped con `status:"skipped_tool_broken"`.

### H2 (D147 — UNATTENDED PATTERN) — 21/48 forms son "unattended" por diseño

**Síntoma**: 44% de los forms no declaran bindings en el `.form.txt` — los asignan en `Form_Open`/`Form_Load` del `.cls`. Patrón sistémico en G1 (11/11), parcial en G2 (1/6), 0 en G3 (los 6 con 0 bindings usan asignación imperativa `Me.X = value`, NO `Me.X.RowSource`), G4 (7/16 navigation hosts + detail subforms), G5 (2/2). 

**Implicación**: la migración 1:1 a SPA debe preservar el patrón "atributos asignados en código al abrir el form". No es un anti-pattern — es diseño deliberado para forms que se configuran según el rol del usuario.

**Acción** (TK-NC-4): documentar el patrón unattended como **decisión arquitectural** en backend (services que hidratan los componentes al mount).

### H3 (D148 — TABINDEX ZERO) — TODOS los forms tienen ZERO TabIndex declarado

**Síntoma**: confirmado explícitamente en G4 (16/16). Implícito en G1/G3/G5. Access usa default visual order (top→bottom, left→right). **Problema de accesibilidad WCAG 2.4.3 (focus order)**.

**Acción** (TK-NC-5): setear TabIndex explícitamente en la migración web. Mantener el default visual en formularios simples; agregar TabIndex manual en NavigationControls y grids (Form_FormNCProyectoGestion con 42 eventos necesita orden explícito).

### H4 (D149 — OVERLAP EN NAVIGATIONCONTROL HOSTS) — 8 OVERLAP en Form_FormNCAuditoriaSeguimiento

**Síntoma**: el host NavigationControl con subforms (Form_FormNCAuditoriaSeguimiento, Form_FormNCProyectoSeguimiento) tiene subform regions que clip toolbar buttons. 8 OVERLAP findings (warning) — `ComandoActualizar`, `ControlDeNavegación0`, `FrmDetalle`, `cmdSalir` todos se solapan entre sí.

**Acción** (TK-NC-6): en la migración web, los NavigationControls se convierten en **React Router nested routes** o tabs Material Design; el problema de overlap desaparece naturalmente.

### H5 (D150 — OVERLAP SIBLING PATTERN EN AC/AR) — Form_FormNCAuditoriaAC y Form_FormNCAuditoriaAR con mismo bug

**Síntoma**: los forms de AC y AR de NC-Auditoría tienen el mismo patrón — el campo principal (NoConformidad TextBox) se solapa con `cmdSalir`, `ComandoRegistrar`, `ComandoDocumentos`. Probable layout regression donde el TextBox es demasiado ancho.

**Acción** (TK-NC-7): corregir en la migración — el patrón web (auto-layout flex/grid) elimina el problema.

### H6 (D151 — MISSING_GEOMETRY SISTÉMICO) — lblTitulo en los 11 forms de G1 sin posición

**Síntoma**: `lblTitulo` (control Label del header) tiene `left=null` y `top=null` en los 11 forms del grupo Menú. Patrón sistémico — probablemente porque Access los posiciona automáticamente al renderizar el FormHeader. Sin embargo, el sub-agente lo marca como MISSING_GEOMETRY (warning).

**Acción** (TK-NC-8): en web, los títulos siempre tienen posición explícita (CSS). El walkthrough confirma que esto se va a corregir naturalmente en la migración.

### H7 (D152 — CODEGRAPH-VBA INDEX FROZEN) — Auto-sync deshabilitado por file lock

**Síntoma**: `codegraph_explore` retorna banner `⚠️ CodeGraph auto-sync is DISABLED` en TODAS las llamadas. El lock está held por otro writer (probablemente el propio agente que está corriendo o un sync en background).

**Acción** (TK-NC-9): post-walkthrough, correr `codegraph sync` una vez para refrescar el índice. Documentar en skill `codegraph-usage` que el stale banner es informativo, no failure.

### H8 (D153 — CATÁLOGO COMPARTIDO) — Form_FormMotivosNoRequiereControlEficacia es cross-domain

**Síntoma**: este catálogo de motivos "no requiere CE" es llamado por **ambos** NC-Aud (`Form_FormNCAuditoriaGeneral`) y NC-Proy (`Form_FormNCProyectoGeneral`). Cross-domain por diseño.

**Acción** (TK-NC-10): en web, este catálogo se convierte en un **shared component** (React/Vue) o endpoint `/motivos-no-ce` consumido por ambos dominios. Backend: tabla única en esquema `no_conformidades`, expuesta vía API a ambos workflows.

### H9 (D154 — POSSIBLE ZOMBIE TABLE) — TipologiaNCProyectos table not found

**Síntoma**: `query_execute` para `TipologiaNCProyectos` (en `Form_FormTipologiaNCProyecto`) retorna `table/query not found`. Posible ZOMBIE table — el form referencia una tabla que no existe en el backend.

**Acción** (TK-NC-11): audit de uso completo de `TipologiaNCProyectos` en `data/staging/no-conformidades/src/classes/` para confirmar si es zombie o si el nombre correcto es `TbTiposNCProyectos` (que SÍ aparece en G4).

---

## 4. Decisiones aplicadas

### 4.1 UX: Preservar / Mejorar / Nuevo paradigma

| Pantalla | Decisión | Justificación |
|---|---|---|
| `Form_Form0BDOpciones` (menú principal) | **Preservar** jerarquía dual NC-Proy/Aud | El usuario navega entre dominios — el menú dual es estructura cognitiva aprendida |
| `Form_Form0BDOpcionesParteProyectos` (timer-driven indicadores) | **Mejorar**: server-side polling + cache | El `Form_Timer` con `PintarIndicadores` indica refresh periódico; en web usar SSE/WebSocket |
| `Form_FormIndicadores` | **Mejorar**: charts interactivos | Funcionalidad ya existe, mejorar UX |
| `Form_FormCorreo` (modal) | **Preservar** flujo | Trivial: composer web |
| `Form_FormMotivosNoRequiereControlEficacia` | **Nuevo paradigma**: shared component/API | D153 cross-domain — convertir en endpoint compartido |
| `Form_formRiesgosSeleccion` | **Preservar** flujo, **Mejorar**: API REST federada | D132 — handshake con Gestion_Riesgos |
| `Form_FormExpedientesBusqueda` | **Preservar** flujo, **Mejorar**: OAuth/SSO | D132 — handshake con Expedientes |
| `Form_FormAuditoriaSeleccion` (unattended) | **Mejorar**: hidratar en backend | D147 — patrón unattended preservado en backend services |
| `Form_FormAuditoria` (NavigationControl) | **Mejorar**: nested routes | Tabs web eliminan overlap (D149) |
| `Form_FormNCAuditoriaGeneral` (alta/edición) | **Preservar** estructura 34 controles | Form denso pero funcional — preservar jerarquía visual |
| `Form_FormNCAuditoriaAC` / `Form_FormNCAuditoriaAR` | **Mejorar**: flex layout | D150 — fix overlap AC/AR sibling pattern |
| `Form_FormNCAuditoriaSeguimiento` (NavigationControl host) | **Nuevo paradigma**: nested routes | D149 — el subform clipping toolbar es bug de layout Access, no del modelo |
| `Form_FormNCProyecto` (navigation host) | **Nuevo paradigma**: layout shell + outlet | NavigationControl no portable; usar shell con outlet pattern |
| `Form_FormNCProyectoGestion` (grid 61 controles, 42 eventos) | **Mejorar**: server-side pagination + chips filtro | El grid carga TODO; en web paginar |
| `Form_FormNCProyectoGeneral*` | **Preservar** estructura + **Mejorar** layout | Mismo patrón que NCAudGeneral — preservar + fix overlap |
| `Form_FormNCProyectoTipologiaGestion` (unattended) | **Mejorar**: hidratar en backend | D147 |

### 4.2 Arquitectura: hexagonal ports

| Concepto legacy | Traducción web |
|---|---|
| `getdb()` (DAO.Database) | Puerto de persistencia PostgreSQL (D14) + HTTP client |
| `Constructor.getXxx()` (factory con cache lazy) | Inyección de dependencias + repository pattern |
| `m_ObjEntorno` (singleton global con ColXxx dictionaries) | Service registry / dependency injection |
| `m_ObjNCAuditoriaActiva` / `m_ObjNCProyectoActiva` (singleton in scope) | Request-scoped services / Context API |
| Unattended form (21/48) — bindings asignados en `Form_Open` con `Me.X = value` | Backend service hidrata componentes al mount (D147 — preservar el patrón) |
| NavigationControl host (Form_FormNCAuditoriaSeguimiento, Form_FormNCProyecto) | React Router nested routes / Material UI Tabs |
| `Form_Timer` en `Form_Form0BDOpcionesParteProyectos` (refresh indicadores) | SSE / WebSocket server-push |
| `Form_FormExpedientesBusqueda` HTTP/JSON handshake | API REST federada + OAuth 2.0 (D132) |
| `Form_formRiesgosSeleccion` FK lógica a Gestion_Riesgos | API REST + cross-module FK en PostgreSQL |
| `WithEvents + RaiseEvent` | Pub/sub tipado, state machine |
| `TempVars!Variable` IPC | Promise/callback en modal context |
| `fso.FileExists + ShellExecute('open', url)` | window.open(url) + blob URL |
| `Application.FileDialog(3)` | `<input type="file">` + signed URL upload |
| `TbCacheNCProyecto` (lookup by `IDNoConformidad` con `CacheValida`/`HitsConsultas`/`TamanioBytes`) | Redis o similar server-side cache |
| `Test_KillSwitch.bas → BuildJsonFail/BuildJsonOk` (activar/desactivar caché) | Feature flag server-side (LaunchDarkly / Unleash) |
| `TbCacheNCProyecto` con `CacheValida`/`HitsConsultas`/`TamanioBytes` | Redis con TTL + métricas |

### 4.3 Decisiones de seguridad

- **Cross-app con Gestion_Riesgos y Expedientes**: API REST federada con OAuth 2.0 / SSO. Sin mecanismo actual de auth documentado — oportunidad para diseño limpio.
- **Concurrencia**: optimistic locking con `updated_at` token en cada edición.
- **Unattended forms en backend**: el service que hidrata los controles debe validar que el usuario tiene permisos antes de poblar datos sensibles (ej. `Form_FormNCAuditoriaGeneral` con datos de auditoría confidencial).

### 4.4 Decisiones de datos

- **44 capabilities** distribuidas en 14 familias (CAP-NCP-LC, CAP-NCP-AF, CAP-NCA-LC, CAP-NCA-AF, CAP-CE, CAP-IND, CAP-CAT, CAP-DGE, CAP-EXP, CAP-CFG, CAP-UPN, CAP-XCUT, CAP-COM, CAP-REL).
- **Tabla compartida `TbMotivosNoRequiereCE`**: backend único en esquema `no_conformidades`, consumido por ambos workflows NC-Aud y NC-Proy (D153).
- **Caché**: `TbCacheNCProyecto` migra a Redis con TTL y métricas de hit/miss.
- **Anexos**: filesystem → S3-compatible con versioning + virus scan.
- **`TbTiposNCProyectos`**: confirmada como tabla activa (referenciada en G4 unattended forms). El `TipologiaNCProyectos` que falla en `query_execute` (G5) probablemente es **nombre mal escrito en el código** — TK-NC-11 audit.

---

## 5. Criterios de aceptación

### 5.1 Funcionalidad

- [ ] **CA-F1**: Las 48 features F1-F48 tienen paridad funcional con la versión Access.
- [ ] **CA-F2**: Los 21 forms "unattended" (D147) hidratan sus bindings via backend service al mount, preservando el patrón de asignación imperativa.
- [ ] **CA-F3**: El catálogo compartido `Form_FormMotivosNoRequiereControlEficacia` (D153) se expone via endpoint único `/motivos-no-ce` consumido por ambos workflows.
- [ ] **CA-F4**: Los NavigationControl hosts (`Form_FormNCAuditoriaSeguimiento`, `Form_FormNCProyecto`, `Form_FormNCProyectoSeguimiento`) se migran a nested routes, eliminando el bug de overlap (D149).
- [ ] **CA-F5**: Los OVERLAP en AC/AR forms (D150) se eliminan con flex/grid layout en web.
- [ ] **CA-F6**: La integración cross-app con Gestion_Riesgos (vía `Form_formRiesgosSeleccion`) se migra a API REST federada con OAuth.
- [ ] **CA-F7**: La integración cross-app con Expedientes (vía `Form_FormExpedientesBusqueda`) se migra a API REST federada con OAuth.
- [ ] **CA-F8**: El cross-app catalog `Form_FormMotivosNoRequiereControlEficacia` (D153) se elimina como form duplicado y se consume via API.
- [ ] **CA-F9**: `TbCacheNCProyecto` migra a Redis con TTL configurable y métricas de hit/miss (reemplaza kill switch legacy).
- [ ] **CA-F10**: Las notificaciones / refresh de indicadores (D144 timer-driven) migran a SSE/WebSocket server-push.

### 5.2 Seguridad

- [ ] **CA-S1**: Cross-app APIs con OAuth 2.0 + scopes granulares (Gestion_Riesgos + Expedientes).
- [ ] **CA-S2**: `updated_at` token en cada edición para optimistic concurrency en NC-Aud y NC-Proy.
- [ ] **CA-S3**: Backend service que hidrata unattended forms valida permisos del usuario antes de poblar datos sensibles.
- [ ] **CA-S4**: Las RN de seguridad documentadas en `legacy.rules.ts` cubren los 44 capabilities.

### 5.3 Performance

- [ ] **CA-P1**: `Form_FormNCProyectoGestion` (61 controles, 42 eventos, 10 combos de filtro) con server-side pagination y chips para filtros.
- [ ] **CA-P2**: `Form_FormIndicadores` con charts interactivos, carga sub-segundos incluso con 5 años de datos.
- [ ] **CA-P3**: `TbCacheNCProyecto` reemplazado por Redis con TTL — hit rate > 90% en producción.
- [ ] **CA-P4**: `codegraph-vba sync` se ejecuta post-deploy para refrescar el índice (mitiga H7 file lock contention).

### 5.4 Operacional

- [ ] **CA-O1**: `Form_Form0BDOpcionesParteProyectos` timer-driven indicators migran a SSE/WebSocket.
- [ ] **CA-O2**: `TbNCDocumentosAux` (anexos) migra a S3-compatible con versioning + virus scan.
- [ ] **CA-O3**: Audit de `TipologiaNCProyectos` (H9 zombie table sospecha) completado — confirmar si es zombie o nombre mal escrito.
- [ ] **CA-O4**: 3 bugs de dysflow tracked (issues #1407, #1408, round-4) con resolución o workaround documentado.

---

## 6. Pendientes operacionales

### 6.1 Antes de empezar

- [ ] **PO-1**: Confirmar el contrato de integración con **Gestion_Riesgos** (autenticación, endpoint, schema) — `Form_formRiesgosSeleccion` y `TbRiesgosNC.IDNC`.
- [ ] **PO-2**: Confirmar el contrato de integración con **Expedientes** — `Form_FormExpedientesBusqueda` HTTP/JSON handshake.
- [ ] **PO-3**: Audit de `TipologiaNCProyectos` (H9) — confirmar si es zombie o nombre mal escrito en código.
- [ ] **PO-4**: Esperar resolución de los 3 bugs de dysflow (round-2 #1407, round-3 #1408, round-4 pendiente) o aceptar método v3 como producción.
- [ ] **PO-5**: Decidir estrategia para los 21 forms "unattended" (D147) — preservar el patrón en backend services vs. forzar bindings declarativos en web.

### 6.2 Durante el desarrollo

- [ ] **PO-6**: Implementar backend service para hidratar unattended forms (preservar patrón D147).
- [ ] **PO-7**: Implementar endpoint compartido `/motivos-no-ce` para consumo cross-domain (D153).
- [ ] **PO-8**: Reemplazar NavigationControl hosts por nested routes / Material Tabs (D149).
- [ ] **PO-9**: Fix layout overlap en AC/AR sibling forms (D150) — flex/grid layout.
- [ ] **PO-10**: Setear TabIndex explícito en formularios migrados (D148 WCAG).
- [ ] **PO-11**: Implementar Redis para reemplazar `TbCacheNCProyecto` con TTL + métricas.
- [ ] **PO-12**: Los 92+ geometry findings (OVERLAP, MISSING_GEOMETRY) deben corregirse o documentarse en la migración.

### 6.3 En el go-live

- [ ] **PO-13**: Smoke test E2E: alta NC de Proyecto → AC → AR → Control Eficacia → Replanificación → Seguimiento.
- [ ] **PO-14**: Smoke test E2E: alta NC de Auditoría → idem workflow.
- [ ] **PO-15**: Verificar las RN documentadas con datos de producción antes de switchover.
- [ ] **PO-16**: Backfill de `TbCacheNCProyecto` desde Redis (si hay datos legacy en tabla).
- [ ] **PO-17**: Plan de deprecation de los 7 archivos `Test_*.bas` legacy y forms con bindings en código que se reescriben en backend.

---

## 7. Tickets derivables (preview)

> NO crear issues todavía. Estos nacen de la épica al final del ciclo de revisión.

### Methodology (D146)

- **TK-NC-1**: [METHODOLOGY] Aplicar método v3 a futuras migraciones (5 tools + lint manual + skip graceful).
- **TK-NC-2**: [METHODOLOGY] Documentar el patrón "unattended forms" (D147) en skill `dysflow-usage` o `access-form-ui-builder`.
- **TK-NC-3**: [TOOLING] Filed DysTelefonica/dysflow#1407 (analyze_form_layout) — seguir fix.

### Pattern (D147, D148, D153)

- **TK-NC-4**: [PATTERN] Documentar y preservar patrón unattended en backend services.
- **TK-NC-5**: [A11Y] Setear TabIndex explícito en todos los formularios migrados (D148 WCAG).
- **TK-NC-6**: [LAYOUT] Reemplazar NavigationControl hosts por nested routes (D149).
- **TK-NC-7**: [LAYOUT] Fix overlap en AC/AR sibling forms (D150).
- **TK-NC-8**: [LAYOUT] Corregir MISSING_GEOMETRY sistémico (lblTitulo, lblEstado) en migración (D151).
- **TK-NC-10**: [INTEGRATION] Endpoint compartido `/motivos-no-ce` (D153).

### Forms y migración

- **TK-NC-11**: [AUDIT] Resolver sospecha zombie table `TipologiaNCProyectos` (D154).
- **TK-NC-12**: [MIGRATION] Migrar 21 unattended forms preservando el patrón (D147).
- **TK-NC-13**: [MIGRATION] Migrar los 4 NavigationControl hosts a nested routes (D149).
- **TK-NC-14**: [MIGRATION] Consolidar `Form_FormARAuditoriaDocumentos` + `Form_FormARProyectoDocumentos` en componente compartido (sibling pattern detectado en G1).
- **TK-NC-15**: [MIGRATION] Migrar 48 classes de dominio a services HTTP manteniendo firmas.
- **TK-NC-16**: [MIGRATION] Reemplazar `TbCacheNCProyecto` por Redis con TTL configurable.
- **TK-NC-17**: [MIGRATION] Reemplazar `Test_KillSwitch.bas` por feature flag server-side.

### Integración

- **TK-NC-18**: [INTEGRATION] API REST federada con Gestion_Riesgos (D132) para `Form_formRiesgosSeleccion` y `TbRiesgosNC.IDNC`.
- **TK-NC-19**: [INTEGRATION] API REST federada con Expedientes (D132) para `Form_FormExpedientesBusqueda`.
- **TK-NC-20**: [INTEGRATION] API REST federada con Brass (D86) si comparten `Form_FormMotivosNoRequiereControlEficacia`.

### Performance y operación

- **TK-NC-21**: [PERFORMANCE] Server-side pagination para `Form_FormNCProyectoGestion` (61 controles).
- **TK-NC-22**: [PERFORMANCE] Charts interactivos para `Form_FormIndicadores`.
- **TK-NC-23**: [OPS] `codegraph sync` post-deploy para refrescar índice (H7).
- **TK-NC-24**: [OPS] SSE/WebSocket para refresh de indicadores (timer-driven legacy).

### Testing y validación

- **TK-NC-25**: [TEST] Tests E2E para workflow completo NC de Proyecto (Alta → AC → AR → CE → Replanif → Seguimiento).
- **TK-NC-26**: [TEST] Tests E2E para workflow completo NC de Auditoría.
- **TK-NC-27**: [TEST] Tests para los 21 unattended forms (validar hidratación backend).
- **TK-NC-28**: [TEST] Tests de regresión para los 3 bugs de dysflow (round-2, round-3, round-4).

### Datos y migración

- **TK-NC-29**: [MIGRATION] Backfill de datos desde `TbCacheNCProyecto` legacy a Redis.
- **TK-NC-30**: [MIGRATION] Backfill de anexos desde filesystem a S3-compatible.
- **TK-NC-31**: [MIGRATION] Expand and Contract para las 13 tablas referenciadas en G1 + 9 en G4 + cross-app.
- **TK-NC-32**: [MIGRATION] Reemplazar IDs anti-patrón Access (sequences `TbID*`) por IDENTITY/SERIAL.

---

## Anexo · Tabla de fuentes

| Fuente | Aporta |
|---|---|
| engram topic_key `no-conformidades/walkthrough-2026-08-06` | Consolidado de 5 walkthroughs paralelos G1..G5 |
| `docs/03-aplicaciones/no-conformidades/walkthrough-G1.json` | 11 forms (Menú + cross-cutting) — método v3, 39463 bytes |
| `docs/03-aplicaciones/no-conformidades/walkthrough-G2.json` | 6 forms (Auditoría workflow) — método v3, 19846 bytes |
| `docs/03-aplicaciones/no-conformidades/walkthrough-G3.json` | 13 forms (NC de Auditoría workflow) — método v3, 31657 bytes |
| `docs/03-aplicaciones/no-conformidades/walkthrough-G4.json` | 16 forms (NC de Proyecto workflow) — método v3, 34681 bytes |
| `docs/03-aplicaciones/no-conformidades/walkthrough-G5.json` | 2 forms (Catálogos tipología) — método v3, 4179 bytes |
| `docs/03-aplicaciones/no-conformidades/forms.md` | Árbol de navegación + call paths |
| `docs/03-aplicaciones/no-conformidades/data-model.md` | Tablas y schema |
| `docs/03-aplicaciones/no-conformidades/capabilities.md` | 14 familias de capabilities |
| `docs/03-aplicaciones/no-conformidades/security-rules.md` | Reglas de seguridad (D89 `list_objects` diagnosticado) |
| `docs/03-aplicaciones/no-conformidades/integrations-automation.md` | Integraciones cross-app |
| `docs/03-aplicaciones/no-conformidades/migration-matrix.md` | Matriz de migración feature × capability |
| `docs/03-aplicaciones/no-conformidades/README.md` | Overview del estudio |
| `data/staging/no-conformidades/src/` (48 cls + 48 form .cls + 48 form.txt + 25 bas) | Source tree exportado |
| `data/staging/no-conformidades/.codegraph-vba/` | Índice codegraph-vba (FROZEN al walkthrough) |
| `data/staging/no-conformidades/docs/` | Documentación preexistente |
| `C:\00repos\codigo\00_NO_CONFORMIDADES\staging` | Fuente READ-ONLY |

## Anexo · Decisiones referenciadas (D5-D154)

| Decisión | Aplicación a NoConformidades |
|---|---|
| D8 (hexagonal global) | Toda la migración |
| D9-D10 (secret manager) | Acceso a `getdbLanzadera` (vía URL de Lanzadera) |
| D14 (esquema por módulo) | Módulo `no_conformidades` en PostgreSQL |
| D16 (object storage S3-compatible) | Anexos (D147 — `TbNCDocumentosAux`) |
| D27 (logs estructurados) | Reemplazar `TbLogCambios` |
| D44-D46 (autorización + capabilities) | Roles Calidad/Técnico |
| D66-D67 (stack) | Backend Python + HTMX |
| D68 (monolito modular) | Migración a módulo dedicado |
| D82 (Expand and Contract) | Backward-compatible con backend legacy durante migración |
| D86-D87 (cross-app con Lanzadera) | Patrón compartido con Gestion_Riesgos vía `getdbLanzadera()` |
| D89 (`list_objects` dysflow diagnosticado) | Referencia histórica |
| D102 (booleanos Text(2)) | Cleanup de captions |
| D132 (XApp Expedientes HTTP/JSON) | `Form_FormExpedientesBusqueda` → API REST federada |
| D144 (timer-driven indicators) | `Form_Timer` en `Form_Form0BDOpcionesParteProyectos` → SSE/WebSocket |
| **D146 (methodology v3)** | **Walkthrough con 3 tools dysflow degradadas/rotas — método v3 aplicado** |
| **D147 (unattended pattern)** | **21/48 forms unattended — preservar patrón en backend services** |
| **D148 (TabIndex ZERO)** | **Todos los forms sin TabIndex declarado — setear explícito en web** |
| **D149 (NavigationControl overlap)** | **8 OVERLAP en Seguimiento — nested routes eliminan el bug** |
| **D150 (AC/AR sibling overlap)** | **SAME pattern en AC y AR — flex/grid layout** |
| **D151 (MISSING_GEOMETRY sistémico)** | **lblTitulo/lblEstado sin posición en G1 — fix natural en web** |
| **D152 (codegraph-vba FROZEN)** | **Auto-sync DISABLED por file lock — `codegraph sync` post-deploy** |
| **D153 (catálogo compartido)** | **`Form_FormMotivosNoRequiereControlEficacia` cross-domain — shared API** |
| **D154 (zombie table sospecha)** | **`TipologiaNCProyectos` no encontrada — audit de uso** |

## Checklist del documento

- [x] Scope con 48 features detalladas por dominio (G1-G5)
- [x] Walkthrough G1..G5 (48/48 forms, método v3) con JSON estructurado
- [x] Hallazgos D146-D154 con anchor links
- [x] Decisiones UX Preservar/Mejorar/Nuevo paradigma por pantalla
- [x] Decisiones arquitectura hexagonal ports
- [x] Criterios de aceptación verificables y agrupados por dimensión
- [x] 17 pendientes operacionales antes, 12 durante, 5 en go-live
- [x] 32 tickets derivables preview (TK-NC-1..32)
- [x] Tabla de decisiones referenciadas (D5-D154)
- [x] Tabla de fuentes
- [x] Idioma: español técnico neutro. Identificadores y paths sin traducir.
- [x] D146 particular: metodología v3 con 3 tools dysflow degradadas/rotas documentadas
- [x] D147 particular: 21/48 forms unattended preservados como decisión arquitectural
- [x] D148 particular: TabIndex ZERO sistémico — setear explícito en web

## Siguiente paso

Revisión con el equipo. Esta épica se revisa junto con la de Brass y las de las otras apps (HPS, HPS_Solicitudes, Condor, Gestion_Riesgos ya existentes). El siguiente lote es una de las 2 apps restantes (Lanzaderas, Expedientes) tras cerrar el ciclo de revisión final de las 8 épicas.
