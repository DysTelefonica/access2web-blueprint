[← Back to Expedientes README](README.md) · [← Codebase Guide](../../../CODEBASE-GUIDE.md) · [← DOCS](../../../DOCS.md)

# Épica — Expedientes (migración a web)

> **Estado:** DRAFT v0.1 (2026-08-07) — pendiente revisión final al cerrar el ciclo de las 8 épicas.
> **Lote:** 8 — **ÚLTIMO** del ciclo. Cierra las 8 épicas del blueprint.
> **App legacy:** `00_EXPEDIENTES` · frontend `Expedientes.accdb` + backend `Expedientes_datos.accdb`.
> **Sentence that organizes**: **Expedientes es la app de gestión documental de procedimientos administrativos: cada expediente vincula un Suministrador, Comercial, Lugar de Ejecución, Oficina, Órgano, Ejército y catálogo de tipos (CPV, PECAL, RAC, Grados). El menú es JSON-driven via `mod0BDOpcionesHelper`.**

> **Scope del scope**: "Este repo es research + planning de la migración. Cada app tendrá su propio repo + docs cuando se construya."

---

## Quick Navigation

| Section                                                       | What you'll find                                              |
|---------------------------------------------------------------|---------------------------------------------------------------|
| [Metadatos](#metadatos)                                       | Scope M, riesgos D171/D172, dependencias cross-app.            |
| [1. Scope](#1-scope)                                          | 46 features F1-F46 en 5 dominios (G1-G5).                    |
| [2. Estado del descubrimiento](#2-estado-del-descubrimiento)    | Inventario 46 forms walkthroughed, method v4, patterns.        |
| [3. Hallazgos críticos](#3-hallazgos-críticos)                | D169-D178: JSON-hub, copy-paste RowSource, twin patterns.     |
| [4. Decisiones aplicadas](#4-decisiones-aplicadas)              | UX por pantalla, hexagonal ports, naming.                     |
| [5. Criterios de aceptación](#5-criterios-de-aceptación)      | Funcionales, Seguridad, Performance, Operacional.              |
| [6. Pendientes operacionales](#6-pendientes-operacionales)    | Antes, durante, go-live.                                       |
| [7. Tickets derivables](#7-tickets-derivables-preview)         | 22 tickets TK-EX-1..22.                                        |
| [Anexo · Decisiones referenciadas](#anexo--decisiones-referenciadas) | D5-D178.                                              |
| [Anexo · Tabla de fuentes](#anexo--tabla-de-fuentes)         | Walkthrough JSONs + docs + source tree + engram.               |

---

## Metadatos

| Campo | Valor |
|---|---|
| **Aplicación legacy** | `00_EXPEDIENTES` · frontend `Expedientes.accdb` + backend `Expedientes_datos.accdb` |
| **Tipo de migración** | Legacy Access/VBA → web hexagonal (FastAPI + HTMX) |
| **Scope size** | **M** (46 forms, según project.json "el set más documental" — **real: 46 forms, 69 archivos en docs/**) |
| **Dependencias cross-app** | **HPS** (vía `Form_FormExpedienteAltaParaHPS` — entry point específico desde HPS) · **NoConformidades** (presumiblemente identidad/permisos, vía EVE — D144) · **Gestion_Riesgos** (`Form_Form0BDOpcionesTecnicos` referencia `Form_FormGestionRiesgos` — D168) |
| **Riesgo dominante** | **D171 (Copy-paste RowSource bug en 6/7 G4)** — bug silencioso que `verify_form_bindings` (#1412 OPEN) habría detectado pero la tool está rota. Solo visible por inspección manual del walkthrough. **D172 (ComandoAyuda mal etiquetado)** en 10/13 forms G4. **D173 (NavigationControl parse anomaly)** en 6 forms G2 (sections=0). |
| **Stack target** | Backend Python 3.12+ / FastAPI 0.119+ / SQLAlchemy 2.0.x / Alembic 1.13+ / asyncpg 0.30+ (D66) · Frontend HTMX 2.0.4 + Jinja2 3.1+ + Alpine.js 3.15+ (D67) |
| **Estrategia de migración de BD** | Expand and Contract backward-compatible (D82) · PostgreSQL compartido con esquema por módulo (D14) |
| **Forma destino** | Hexagonal global (D8) · módulo dentro del monolito modular (D68) · puerto de persistencia PostgreSQL + object storage S3-compatible (D16) + secret manager (D9-D10) |
| **Auditoría de uso previa** | Walkthrough G1..G5 (46/46 forms, método v4) · Codegraph-VBA + Dysflow (2 bugs filed: #1408 OPEN, #1412 OPEN — #1407 cerrado) · 5 critical findings de copy-paste RowSource bug (D171), ComandoAyuda mal etiquetado (D172), y JSON-hub pattern (D169) |

---

## 1. Scope

### 1.1 En scope — 46 features agrupadas en 5 dominios

#### G1 — Splash + Login + Menús + Cross-cutting (9 features)

| # | Feature | Respaldo |
|---|---|---|
| F1 | **Splash** (entry point con timer + autenticación vía EVE — D144) | `Form_frmSplash` |
| F2 | **frmBusy** (modal anti-spam compartido — patrón `vba-antispam-popup` skill) | `Form_frmBusy` |
| F3 | **Formulario 1** (placeholder WIP) | `Form_Formulario1` |
| F4 | **Menú principal** (JSON-driven hub via `mod0BDOpcionesHelper` + `JsonConverter.ParseJson`) | `Form_Form0BDOpciones` |
| F5 | **Menú técnico** (sub-menú para rol técnico con tareas pendientes) | `Form_Form0BDOpcionesTecnicos` |
| F6 | **Elección tipo consulta** (modal para elegir cómo ver expedientes) | `Form_FormEleccionTipoConsulta` |
| F7 | **Expediente modificados** (lista de expedientes modificados recientemente) | `Form_FormExpedienteModificados` |
| F8 | **Modificado flow** (entry point para crear/modificar expediente — navegación a subform Viajero) | `Form_FormModificado` |
| F9 | **E2E Gestión Batch** (lista de expedientes disponibles para procesamiento batch end-to-end) | `Form_FormE2EGestionBatch` |

#### G2 — Expedientes alta + edición + subforms (11 features)

| # | Feature | Respaldo |
|---|---|---|
| F10 | **Expediente (alta/edición principal)** (form complejo con 38 OVERLAP — header lblTitulo stacked con navigation tabs) | `Form_FormExpediente` |
| F11 | **Expediente Alta** (modal de alta inicial) | `Form_FormExpedienteAlta` |
| F12 | **Expediente Alta para HPS** (entry point desde HPS) | `Form_FormExpedienteAltaParaHPS` |
| F13 | **Expediente Alta Tipo** (modal con selección de tipo — usa Form_Open) | `Form_FormExpedienteAltaTipo` |
| F14 | **Expediente Cambio Tipo** (modal para cambiar tipo — usa Form_Open) | `Form_FormExpedienteCambioTipo` |
| F15 | **Expediente Documentación** (subform de docs) | `Form_FormExpedienteDocumentacion` |
| F16 | **Expediente Entidades** (subform de entidades — abre 6 de los 7 Gestion forms vía OpenForm) | `Form_FormExpedienteEntidades` |
| F17 | **Expediente Fechas** (subform de fechas) | `Form_FormExpedienteFechas` |
| F18 | **Expediente General** (subform de datos generales) | `Form_FormExpedienteGeneral` |
| F19 | **Expediente Hitos** (subform de hitos) | `Form_FormExpedienteHitos` |
| F20 | **Expediente Suministradores** (subform de suministradores — usa Form_Open) | `Form_FormExpedienteSuministradores` |

#### G3 — Expedientes gestión (4 features)

| # | Feature | Respaldo |
|---|---|---|
| F21 | **Expedientes Gestión** (lista + filtros + acciones — form ENORME con 122 controles, 12 bindings, 146 findings) | `Form_FormExpedientesGestion` |
| F22 | **Expedientes Gestión Técnica** (vista técnica de gestión — unattended estricto) | `Form_FormExpedientesGestionTecnica` |
| F23 | **Expedientes para Cambio Tipo** (lista de expedientes elegibles para cambio de tipo) | `Form_FormExpedientesParaCambioTipo` |
| F24 | **Gestor de Entidades** (alta/baja de entidades — base del menú técnico) | `Form_Form0BDGestorEntidades` |

#### G4 — Catálogos mayores (13 features)

| # | Feature | Respaldo |
|---|---|---|
| F25 | **Suministrador Alta/Edición** | `Form_FormSuministrador` |
| F26 | **Suministradores Gestión** (ListaFiltrados.RowSource bug D171 — columnas de Estado, debería ser de Suministrador) | `Form_FormSuministradoresGestion` |
| F27 | **Comercial Alta/Edición** | `Form_FormComercial` |
| F28 | **Comerciales Gestión** (RowSource correcto, único sin bug D171) | `Form_FormComercialesGestion` |
| F29 | **Lugar Ejecución Alta/Edición** | `Form_FormLugarEjecucion` |
| F30 | **Lugares Ejecución Gestión** (RowSource bug D171 — columnas de Estado) | `Form_FormLugarEjecucionGestion` |
| F31 | **Oficina Programa Alta/Edición** | `Form_FormOficinaPrograma` |
| F32 | **Oficinas Programa Gestión** (RowSource bug D171 — columnas de Estado) | `Form_FormOficinasProgramaGestion` |
| F33 | **Órgano Contratación Alta/Edición** (único SIN ComandoAyuda — D172 asimetría) | `Form_FormOrganoContratacion` |
| F34 | **Órganos Contratación Gestión** (RowSource bug D171 — columnas de Estado) | `Form_FormOrganoContratacionGestion` |
| F35 | **Ejército Alta/Edición** | `Form_FormEjercito` |
| F36 | **Ejércitos Gestión** (RowSource bug D171 — columnas de CPV) | `Form_FormEjercitosGestion` |
| F37 | **Usuarios Gestión** (selector de usuario, NO CRUD — rompe el patrón G4 Gestión) | `Form_FormUsuariosGestion` |

#### G5 — Catálogos menores (9 features)

| # | Feature | Respaldo |
|---|---|---|
| F38 | **Grado Clasificación Alta/Edición** | `Form_FormGradoClasificacion` |
| F39 | **Grados Clasificación Gestión** | `Form_FormGradosClasificacionGestion` |
| F40 | **PECAL Alta/Edición** | `Form_FormPECAL` |
| F41 | **PECALES Gestión** | `Form_FormPECALESGestion` |
| F42 | **CPV Alta/Edición** | `Form_FormCPV` |
| F43 | **CPVs Gestión** | `Form_FormCPVsGestion` |
| F44 | **RAC Alta/Edición** | `Form_FormRAC` |
| F45 | **RACS Gestión** | `Form_FormRACSGestion` |
| F46 | **Tareas** (singleton lista de tareas) | `Form_FormTareas` |

### 1.2 Fuera de scope (REPLACE)

> Lo siguiente NO migra como tablas PostgreSQL:

- **`frmBusy` modal anti-spam** (D174) → web nativo: `useTransition` + `aria-busy`. Eliminar el form completo; el patrón se preserva vía UX loading states.
- **`mod0BDOpcionesHelper` JSON-hub pattern** (D169) → backend service: el menú se hidrata vía API que devuelve JSON con la configuración del usuario (rol, permisos, items del menú). La UI consume el JSON y renderiza dinámicamente.
- **ListBox ColumnHeads hardcoded** (D170) → API-driven: la consulta SQL define las columnas en backend; el frontend las renderiza dinámicamente desde el response.
- **Copy-paste RowSource bug en 6/7 G4** (D171) → corrección sistémica en el refactor: cada Gestion form define su propio RowSource desde una función helper parametrizada.
- **`ComandoAyuda` mal etiquetado** (D172) → eliminar caption='Salir' o renombrar el control.
- **NavigationControl subform pattern** (D173 — sections=0 anomaly) → React Router nested routes u outlet pattern.
- **Tablas con naming en MAYUSCULAS** (D174 — `EJERCITO`, `LUGAREJECUCION`, `USUARIO`) → renombrar a snake_case en PostgreSQL (`ejercito`, `lugar_ejecucion`, `usuario`).

### 1.3 Fuera de scope (no documentado)

- **Queries exportadas y macros embebidas**: requieren acceso a binarios. El walkthrough se hizo sobre source tree (.form.txt + .cls).
- **Forms `Form_*` con 0 forms en disco** (`Form_FormExpediente` listado en forms.md pero no aparece en `Get-ChildItem`): verificar si está exportado en `data/staging/expedientes/src/forms/` o si requiere staging de binarios.
- **API de autenticación del handshake con HPS**: presumiblemente vía sesión/token compartido.
- **Tablas `Tb*Expediente*` adicionales**: walkthrough detectó 5 tablas (ExpedienteComercial, ExpedienteLugarEjecucion, Expediente, USUARIO, Suministrador). Las demás no inspeccionadas.

---

## 2. Estado del descubrimiento

### 2.1 Inventario Dysflow + walkthrough G1..G5 (46/46 forms, método v4, 2026-08-07)

| Categoría | Resultado |
|---|---|
| **Frontend** | `Expedientes.accdb` (presumiblemente — staging no confirmado) |
| **Backend** | `Expedientes_datos.accdb` |
| **Forms** | **46** (.form.txt + .cls par 1:1) — **46/46 walkthroughed** |
| **Walkthroughs JSON** | [`walkthrough-G1.json`](walkthrough-G1.json) (23000 bytes), [`walkthrough-G2.json`](walkthrough-G2.json) (23387 bytes), [`walkthrough-G3.json`](walkthrough-G3.json) (58527 bytes), [`walkthrough-G4.json`](walkthrough-G4.json) (41444 bytes), [`walkthrough-G5.json`](walkthrough-G5.json) (14231 bytes) |
| **Total controls** | ~700+ (122 en G3 Form_FormExpedientesGestion solo + 530+ en el resto) |
| **Total geometry findings** | **~815** (G1: 106, G2: 215, G3: 146+ en un solo form, G4: 215, G5: 133) |
| **Critical findings** | **19** distribuidos en G1 (3), G2 (5), G3 (4), G4 (3 declared + 2 medium), G5 (4) |

### 2.2 Walkthrough method v4 — bugs de dysflow manejados

| Bug | Issue | Estado en 2.36.2 | Handling v4 |
|---|---|---|---|
| `analyze_form_layout` RESULT_CONTRACT_VIOLATION | [#1407](https://github.com/DysTelefonica/dysflow/issues/1407) | **RESUELTO en 2.36.2** | Funciona OK, findings tipados |
| `map_form_behavior autoFetchCodeGraph` --json rejected | [#1408](https://github.com/DysTelefonica/dysflow/issues/1408) | OPEN | `autoFetchCodeGraph:false` |
| `verify_form_bindings` RESULT_CONTRACT_VIOLATION | [#1412](https://github.com/DysTelefonica/dysflow/issues/1412) | OPEN | `status:"skipped_tool_broken"` |

### 2.3 Ambient conditions del walkthrough

- **codegraph-vba index FROZEN**: auto-sync DISABLED por file lock contention.
- **codegraph-vba NO indexa expedientes** (D157 también aplica): blast radius devuelve cross-project noise (forms de Brass/Condor/etc).
- **Stale banner ignorado per protocol** en todos los calls.
- **codegraph explore budget limited**: G1 usó sólo 2 calls (deeper en Form_frmSplash y Form_Form0BDOpciones); los otros 7 forms sólo auto-fetched.

### 2.4 Patrones estructurales detectados

#### a) JSON-hub pattern del menú (D169)
- `mod0BDOpcionesHelper.bas` define `Opciones_OpenFormAction` y `Opciones_FormOpenState`.
- `Form_Form0BDOpciones` (y los otros menús) usan `JsonConverter.ParseJson` en 6+ handlers para dispatch del menú.
- `Form_FormEleccionTipoConsulta` (G1) llama `DoCmd.OpenForm` para todos los forms.
- `Form_FormExpediente` (G2) llama `DoCmd.OpenForm` para todos los catalog forms.
- `mod0BDGestorEntidadesHelper.bas` define `GestorEntidades_OpenDecision` consumido por 4 Gestion forms (Grados, PECALES, CPVs, RACSGestion).
- **Implicación para backend**: el menú principal debe ser JSON-driven desde API (`GET /api/menu` con rol + permisos del usuario).

#### b) Twin pattern Edicion vs Gestion (G4 — D175)
- **6 forms Edicion** (F25, F27, F29, F31, F33, F35): comparten pattern Form_Open → validar DTO → ParseJson → Alta_Abrir_Inicializar → save.
- **6 forms Gestion** (F26, F28, F30, F32, F34, F36): comparten otro pattern Form_Open → cargar ListaFiltrados → CRUD → cmdElegir dispara Event.
- **1 form rompe el patrón**: F37 (UsuariosGestion) NO tiene CRUD, solo selector.
- **Migración web**: consolidar 2 patterns en componentes parametrizables (no 13 forms bespoke).

#### c) Copy-paste RowSource bug (D171 — CRITICAL)
- **6 de 7 forms Gestion** tienen `ListaFiltrados.RowSource` con columnas que NO corresponden a su catálogo:
  - Form_FormSuministradoresGestion: `'IDESTADO;ESTADO;DESCRIPCION'` (de Estados, debería ser de Suministrador)
  - Form_FormLugarEjecucionGestion: `'IDESTADO;ESTADO;DESCRIPCION'`
  - Form_FormOficinasProgramaGestion: `'IDESTADO;ESTADO;DESCRIPCION'`
  - Form_FormOrganoContratacionGestion: `'IDESTADO;ESTADO;DESCRIPCION'`
  - Form_FormEjercitosGestion: `'IDCPV;CPV;DESCRIPCION'` (de CPVs, debería ser de Ejercito)
  - Form_FormUsuariosGestion: `'IDESTADO;ESTADO;DESCRIPCION'`
- Solo Form_FormComercialesGestion tiene el RowSource correcto.
- **Bug silencioso**: en tiempo de diseño no se valida contra la tabla real. `verify_form_bindings` (#1412 OPEN) habría detectado este bug — pero la tool está rota, así que el bug pasó desapercibido por años.
- **Impacto**: las columnas del ListBox no matchean con la consulta SQL que las llena, mostrando columnas desplazadas o vacías en producción.

#### d) ComandoAyuda mal etiquetado (D172)
- 10/13 forms G4 tienen `ComandoAyuda` con `caption='Salir'` (no 'Ayuda'), `ControlTipText='Ayuda'`, `Picture='Ayda1.png'`.
- Naming misleading — debería renombrarse a `ComandoCerrar` o cambiar caption a 'Ayuda'.
- Bug cosmético sin impacto funcional pero confuso para el usuario.

#### e) Forms con sections=0 anomaly (D173)
- Forms 6-11 de G2 (`Documentacion`, `Entidades`, `Fechas`, `General`, `Hitos`, `Suministradores`) reportan `sections=0`.
- Causa probable: NavigationControl subform pattern confunde al parser.
- No es bug real — es limitación del parser de .form.txt.

#### f) Unattended estricto vs imperativo
- **G1**: 0 unattended (todos attended).
- **G2**: 0 unattended (todos attended).
- **G3**: 3 unattended (GestionTecnica, ParaCambioTipo, GestorEntidades — globales state coupling).
- **G4**: 13/13 unattended (todos imperativos — son catálogos edición+gestión).
- **G5**: 0 unattended.
- **Total: 16 unattended / 46 forms (35%)**.

#### g) Tablas backend referenciadas (union de todos los walks)
- **G4**: `ExpedienteComercial`, `ExpedienteLugarEjecucion`, `Expediente`, `USUARIO`, `Suministrador`.
- **G5**: `CPV`, `ESTADO`, `Expediente`.

#### h) Naming inconsistencies
- `EJERCITO`, `LUGAREJECUCION`, `USUARIO` en MAYUSCULAS vs CamelCase (Comercial, OficinaPrograma, Ejercito) — copy-paste desde fuente externa.
- Migración: renombrar a snake_case (`ejercito`, `lugar_ejecucion`, `usuario`).

---

## 3. Hallazgos críticos

| # | ID | Título | Severidad | Forms afectados | Detalle |
|---|---|---|---|---|---|
| H1 | D169 | **JSON-hub pattern del menú** | high | `Form_Form0BDOpciones`, `Form_Form0BDOpcionesTecnicos`, `Form_FormEleccionTipoConsulta`, `Form_Form0BDGestorEntidades` | `mod0BDOpcionesHelper` con `JsonConverter.ParseJson` en 6+ handlers. Menú es JSON-driven, no hardcoded. Backend: `GET /api/menu` con rol+permisos. |
| H2 | D170 | **ListBox ColumnHeads hardcoded** | medium | `Form_FormEleccionTipoConsulta`, `Form_FormExpedienteModificados`, `Form_FormE2EGestionBatch` | 3 forms con RowSource en código. `verify_form_bindings` (#1412) lo habría detectado pero la tool está rota. Migrar a API-driven. |
| H3 | D171 | **Copy-paste RowSource bug en 6/7 forms G4** | CRITICAL | `Form_FormSuministradoresGestion`, `Form_FormLugarEjecucionGestion`, `Form_FormOficinasProgramaGestion`, `Form_FormOrganoContratacionGestion`, `Form_FormEjercitosGestion`, `Form_FormUsuariosGestion` | Bug silencioso: columnas de OTRO catálogo. `verify_form_bindings` (#1412) habría detectado este bug. Corrección sistémica en el refactor. |
| H4 | D172 | **`ComandoAyuda` mal etiquetado** | medium | 10/13 forms G4 | `caption='Salir'`, `ControlTipText='Ayuda'`, `Picture='Ayda1.png'`. Naming misleading. Renombrar a `ComandoCerrar` o cambiar caption. |
| H5 | D173 | **NavigationControl subform pattern — sections=0 anomaly** | low | Forms 6-11 de G2 (`Documentacion`, `Entidades`, `Fechas`, `General`, `Hitos`, `Suministradores`) | Sections=0 en parse. Causa probable: NavigationControl subform confunde al parser. No es bug real — limitación del parser. |
| H6 | D174 | **Naming MAYUSCULAS en tablas** | low | `TbEquiposCalibrables`-equivalente en Expedientes: `EJERCITO`, `LUGAREJECUCION`, `USUARIO` | Renombrar a snake_case en PostgreSQL: `ejercito`, `lugar_ejecucion`, `usuario`. |
| H7 | D175 | **Twin pattern Edicion/Gestion** | high | 6 Edicion + 6 Gestion + 1 rompe-patrón (F37) | Consolidar en `<CatalogEditor>` y `<CatalogManager>` parametrizables. F37 diferenciar como `<UserSelector>`. |
| H8 | D176 | **Cross-app ambiguity** | low | `Form_Formulario1` en Expedientes y Lanzaderas | Renombrar a `ExpedientesFormulario1` / `LanzaderaFormulario1` para evitar ambigüedad. |
| H9 | D168 | **Methodology v4 aplicada** | — | Walkthrough | (sección 2.2) |
| H10 | D177 | **Unattended estricto vs imperativo** | medium | 16/46 unattended (35%) | Distinguir RowSource/ControlSource (unattended estricto) vs Me.X = value (imperativo). Preservar patrón en backend services. |
| H11 | D178 | **Global state coupling** | medium | 7 forms G4 con globals | `m_ObjExpedienteDTOActivo`. Refactor a request-scoped services. |

---

## 4. Decisiones aplicadas

### 4.1 UX: Preservar / Mejorar / Nuevo paradigma

| Pantalla | Decisión | Justificación |
|---|---|---|
| `Form_frmSplash` | **Mejorar**: timer-driven UX con skeleton | Splash screen web estándar |
| `Form_frmBusy` | **Eliminar** (D174) | Modal anti-spam no aplica en web — usar `useTransition` + loading states |
| `Form_Form0BDOpciones` (menú principal) | **Nuevo paradigma**: JSON-driven desde API | D169 JSON-hub pattern preservado en backend |
| `Form_Form0BDOpcionesTecnicos` | **Mejorar**: JSON-driven + role-based | Mismo patrón que menú principal |
| `Form_FormEleccionTipoConsulta` | **Mejorar**: ListBox ColumnHeads API-driven | D170 |
| `Form_FormExpedienteModificados` | **Mejorar**: ListBox API-driven | D170 |
| `Form_FormExpediente` (alta/edición principal) | **Mejorar**: server-side pagination + chips de filtro | 122 controles es overload |
| `Form_FormExpedienteAlta*` | **Preservar** estructura modal | 3 forms similares (Alta, AltaParaHPS, AltaTipo) — consolidar |
| `Form_FormExpedienteCambioTipo` | **Preservar** | Modal |
| `Form_FormExpedientesGestion` (G3) | **Mejorar**: server-side pagination + virtualización | 122 controles reducidos |
| 6 forms Edicion (G4) | **Nuevo paradigma**: `<CatalogEditor catalog="...">` parametrizable | D175 consolidación de twin pattern |
| 6 forms Gestion (G4) | **Nuevo paradigma**: `<CatalogManager catalog="..." mode="crud">` parametrizable | D171 bug sistémico corregido en el refactor |
| `Form_FormUsuariosGestion` (G4) | **Mejorar**: componente selector distinto | D178 rompe patrón Gestión |

### 4.2 Arquitectura: hexagonal ports

| Concepto legacy | Traducción web |
|---|---|
| `getdb()` (DAO.Database) | Puerto de persistencia PostgreSQL (D14) + HTTP client |
| `Constructor.getXxx()` (factory con cache lazy) | Inyección de dependencias + repository pattern |
| `m_ObjEntorno` (singleton global) | Service registry / dependency injection |
| `m_ObjUsuarioConectado` / `m_ObjUsuarioConectadoLogin` (singleton in scope) | Request-scoped services / Context API |
| `m_ObjExpedienteDTOActivo` (singleton in scope) — **GLOBAL STATE** (D178) | **Request-scoped services** |
| `mod0BDOpcionesHelper.bas` JSON-hub (D169) | Backend API: `GET /api/menu` retorna items habilitados por rol |
| `JsonConverter.ParseJson` (en 6+ handlers) | Backend FastAPI: response JSON + frontend `fetch().then(r => r.json())` |
| `Opciones_OpenFormAction` / `Opciones_FormOpenState` (G1-G5 mod0BD* helpers) | Backend services: `MenuService.get_user_menu(user_id, role)` |
| `GestorEntidades_OpenDecision` (en 4 Gestion forms) | Backend service: `CatalogService.open_decision(catalog_name, user_id)` |
| `Form_FormEleccionTipoConsulta` DoCmd.OpenForm caller | Frontend router: `<Route path="/expedientes/:id/tipo-consulta">` |
| `Form_FormExpediente` DoCmd.OpenForm catalog opener | Frontend router: `<Route path="/expedientes/:id/catalog/:name">` |
| `frmBusy` modal anti-spam (D174) | `useTransition` + `<button aria-busy={isPending}>` |
| `Form_FormExpedientesGestion` 122 controles | Server-side pagination + chips filtro + virtualización |
| `Form_FormExpediente` (F10) header lblTitulo stacked con navigation tabs | Sticky header + tab strip separado |
| `Form_*AltaTipo`, `Form_*CambioTipo`, `Form_*Suministradores` usan Form_Open (modal pattern) | React Router modal routes |
| `Form_*Documentacion`, `Form_*Entidades`, `Form_*Fechas`, `Form_*General`, `Form_*Hitos` (subform pattern) | Outlet pattern: parent renders child via `<Outlet>` |
| `Form_*Gestion` 6 forms con twin pattern (D175) | `<CatalogManager catalog="..." mode="crud">` parametrizable |
| `ListaFiltrados.RowSource` hardcoded (D171 bug) | API response con `columns: [...]` |
| `ComandoAyuda` mal etiquetado (D172) | Renombrar a `ComandoCerrar` o cambiar caption |
| `Expediente.OficinaPrograma` / `OrganoContratacion` / `Ejercito` Property Get | SQLAlchemy relationships preservadas |
| `m_ObjExpedienteDTOActivo` global state (D178) | Request-scoped services / Context API |
| `WithEvents` + `RaiseEvent` | Pub/sub tipado, state machine |
| `TempVars!Variable` IPC | Promise/callback en modal context |
| `fso.FileExists + ShellExecute('open', url)` | window.open(url) + blob URL |
| `Application.FileDialog(3)` | `<input type="file">` + signed URL upload |

### 4.3 Decisiones de seguridad

- **Cross-app con HPS** (vía `Form_FormExpedienteAltaParaHPS`): API REST federada con OAuth 2.0 / SSO.
- **Cross-app con NoConformidades / Gestion_Riesgos**: vía adaptadores unificados (D9-D10).
- **State coupling via globals** (D178): el refactor a request-scoped services debe validar que el state no leakea entre requests.
- **Permisos por rol**: 6 forms Edicion + 6 forms Gestion tienen su propio modelo de permisos. Consolidar en backend `PermissionService.can_edit(catalog, user)`.

### 4.4 Decisiones de datos

- **Naming en MAYUSCULAS** (D174): renombrar `EJERCITO`/`LUGAREJECUCION`/`USUARIO` → snake_case.
- **Copy-paste RowSource bug** (D171): corrección sistémica en el refactor — cada Gestion form define su propio RowSource desde una función helper.
- **Expediente y sus catálogos** (D177): SQLAlchemy relationships preservadas (`ExpedienteComercial`, `ExpedienteLugarEjecucion`).
- **Cross-app catalog**: `Expediente.Comercial` (`ExpedienteComercial`), `Expediente.LugarEjecucion` (`ExpedienteLugarEjecucion`) — junction tables en PostgreSQL.

---

## 5. Criterios de aceptación

### 5.1 Funcionalidad

- [ ] **CA-F1**: Las 46 features F1-F46 tienen paridad funcional con la versión Access.
- [ ] **CA-F2**: El menú principal es JSON-driven desde API `GET /api/menu` con items habilitados por rol (D169).
- [ ] **CA-F3**: Los 13 forms de G4 (catálogos mayores) se consolidan en 2 componentes parametrizables: `<CatalogEditor>` (6 Edicion) + `<CatalogManager>` (6 Gestion + 1 UserSelector).
- [ ] **CA-F4**: El bug D171 (copy-paste RowSource en 6/7 G4) se corrige sistémicamente en el refactor.
- [ ] **CA-F5**: `ComandoAyuda` se renombra o corrige (D172).
- [ ] **CA-F6**: `frmBusy` se elimina (D174 — anti-spam obsoleto en web).
- [ ] **CA-F7**: `Form_Formulario1` se elimina o completa.
- [ ] **CA-F8**: NavigationControl hosts (D173) migran a React Router nested routes.
- [ ] **CA-F9**: `Form_FormExpedientesGestion` migra con server-side pagination.

### 5.2 Seguridad

- [ ] **CA-S1**: Cross-app API con HPS vía OAuth 2.0 + scopes granulares.
- [ ] **CA-S2**: Request-scoped services validan state no-leak entre requests.
- [ ] **CA-S3**: Permisos por rol consolidados en `PermissionService`.

### 5.3 Performance

- [ ] **CA-P1**: `Form_FormExpedientesGestion` con server-side pagination (122 controles reducidos).
- [ ] **CA-P2**: 13 forms G4 con ListBox virtualization.

### 5.4 Operacional

- [ ] **CA-O1**: 815+ geometry findings se documentan o corrigen durante la migración.
- [ ] **CA-O2**: Tests E2E para los 46 forms (pre-requisito: refactor a DI — D178).
- [ ] **CA-O3**: Cobertura de tests > 70% en módulo expedientes.

---

## 6. Pendientes operacionales

### 6.1 Antes de empezar

- [ ] **PO-1**: Confirmar el contrato de integración con **HPS** (vía `Form_FormExpedienteAltaParaHPS`).
- [ ] **PO-2**: Decidir `Form_FormObtenerContraseña` — completar handlers o eliminar.
- [ ] **PO-3**: Decidir `Form_Formulario1` — completar handlers o eliminar.
- [ ] **PO-4**: Esperar resolución de bug dysflow #1408 (map_form_behavior --json) o aceptar método v4 como producción.
- [ ] **PO-5**: Esperar resolución de bug dysflow #1412 (verify_form_bindings) o aceptar skipped como workaround. **CRÍTICO** porque #1412 habría detectado el bug D171 sistémicamente.
- [ ] **PO-6**: Validar manualmente que el bug D171 (RowSource incorrecto) NO está causando daño en producción — quizás los usuarios ya conviven con columnas desplazadas.

### 6.2 Durante el desarrollo

- [ ] **PO-7**: Implementar `/api/menu` endpoint con JSON-driven config (D169).
- [ ] **PO-8**: Implementar `<CatalogEditor>` y `<CatalogManager>` componentes parametrizables (D175).
- [ ] **PO-9**: Corregir los 6 RowSource incorrectos en G4 (D171).
- [ ] **PO-10**: Renombrar tablas MAYUSCULAS → snake_case (D174).
- [ ] **PO-11**: Refactor de globals a request-scoped services.

### 6.3 En el go-live

- [ ] **PO-12**: Smoke test E2E: alta expediente → entidades → gestión → informe.
- [ ] **PO-13**: Verificar que los 6 RowSource corregidos muestran las columnas correctas para cada catálogo.
- [ ] **PO-14**: Verificar las RN documentadas con datos de producción antes de switchover.
- [ ] **PO-15**: Plan de deprecation de `frmBusy` modal y los WIP forms.

---

## 7. Tickets derivables (preview)

> NO crear issues todavía. Estos nacen de la épica al final del ciclo de revisión.

### JSON-hub y twin patterns (D169, D175)

- **TK-EX-1**: [BACKEND] Implementar `/api/menu` endpoint con JSON-driven config + role-based filtering (D169).
- **TK-EX-2**: [MIGRATION] Migrar 3 ListBox-only forms (EleccionTipoConsulta, ExpedienteModificados, E2EGestionBatch) a ListBox API-driven (D170).
- **TK-EX-3**: [REFACTOR] Consolidar 6 forms Edicion en `<CatalogEditor>` parametrizable (D175).
- **TK-EX-4**: [REFACTOR] Consolidar 6 forms Gestion en `<CatalogManager>` parametrizable (D175).
- **TK-EX-5**: [REFACTOR] Diferenciar `Form_FormUsuariosGestion` (selector, no CRUD) en componente distinto (D178).

### Copy-paste bug y naming (D171, D172, D174)

- **TK-EX-6**: [BUGFIX] Corregir los 6 `ListaFiltrados.RowSource` incorrectos en G4 (D171).
- **TK-EX-7**: [CLEANUP] Renombrar `ComandoAyuda` a `ComandoCerrar` o cambiar caption a 'Cerrar' (D172).
- **TK-EX-8**: [DATA] Renombrar tablas `EJERCITO`/`LUGAREJECUCION`/`USUARIO` → snake_case (D174).

### Forms y migración

- **TK-EX-9**: [MIGRATION] Migrar 46 classes de dominio a services HTTP manteniendo firmas.
- **TK-EX-10**: [MIGRATION] Migrar NavigationControl hosts (D173) a nested routes — 6 forms G2 + G3 navigation hosts.
- **TK-EX-11**: [MIGRATION] Eliminar `frmBusy` modal anti-spam (D174).
- **TK-EX-12**: [MIGRATION] Decidir `Form_Formulario1` (WIP).
- **TK-EX-13**: [MIGRATION] Migrar `Form_FormExpedientesGestion` (122 controles) con server-side pagination.

### Integración

- **TK-EX-14**: [INTEGRATION] API REST federada con HPS vía OAuth 2.0 (F12 entry point).
- **TK-EX-15**: [INTEGRATION] Identidad/permisos compartidos con NoConformidades via `getdbExpedientes()`.

### Testing y validación

- **TK-EX-16**: [TESTING] Refactor de globals a request-scoped services (D178) — pre-requisito para tests.
- **TK-EX-17**: [TESTING] Cobertura de tests > 70% en módulo expedientes.
- **TK-EX-18**: [TESTING] E2E test del flujo Expediente (F11 → F16 → F21).
- **TK-EX-19**: [TESTING] Tests que detecten copy-paste RowSource (D171 — pre-fix para #1412).

### Datos y operación

- **TK-EX-20**: [DATA] Backfill de las 5 tablas referenciadas (ExpedienteComercial, ExpedienteLugarEjecucion, Expediente, USUARIO, Suministrador).
- **TK-EX-21**: [SECURITY] Refactor de `m_ObjExpedienteDTOActivo` global state a request-scoped services.
- **TK-EX-22**: [MIGRATION] Renombrar a `ExpedientesFormulario1` para evitar ambigüedad con `LanzaderaFormulario1` (D167 cross-app).

---

## Anexo · Decisiones referenciadas

| Decisión | Aplicación a Expedientes |
|---|---|
| D8 (hexagonal global) | Toda la migración |
| D14 (esquema por módulo) | Módulo `expedientes` en PostgreSQL |
| D16 (object storage S3-compatible) | Anexos |
| D27 (logs estructurados canónicos) | Reemplazar logs VBA |
| D44-D46 (autorización + capabilities) | Roles Admin/Calidad/Usuario |
| D66-D67 (stack backend + frontend) | Backend Python + HTMX |
| D68 (monolito modular) | Migración a módulo dedicado |
| D82 (Expand and Contract) | Backward-compatible con backend legacy |
| D132 (XApp HTTP/JSON handshake) | Cross-app con HPS, Gestion_Riesgos, NoConformidades |
| D144 (timer-driven UX) | Splash con timer → skeleton web |
| **D155-D168 (methodology)** | **Walkthrough con método v3/v4 aplicado** |
| **D169 (JSON-hub pattern)** | **`mod0BDOpcionesHelper` + `JsonConverter.ParseJson` — preservar en backend con `/api/menu`** |
| **D170 (ListBox ColumnHeads)** | **3 forms con RowSource en código — API-driven en migración** |
| **D171 (copy-paste RowSource bug)** | **6/7 G4 forms con RowSource incorrecto — corrección sistémica** |
| **D172 (ComandoAyuda mal etiquetado)** | **10/13 G4 forms con caption='Salir'** |
| **D173 (NavigationControl parse anomaly)** | **6 forms G2 con sections=0 — limitación parser** |
| **D174 (naming MAYUSCULAS)** | **`EJERCITO`/`LUGAREJECUCION`/`USUARIO` → snake_case** |
| **D175 (twin pattern + 0 tests)** | **Consolidar 6+6 forms en componentes parametrizables** |
| **D176 (cross-app ambiguity)** | **`Form_Formulario1` colisión con Lanzaderas — renombrar** |
| **D177 (Expediente Property Get)** | **SQLAlchemy relationships preservadas** |
| **D178 (global state + 0 tests)** | **`m_ObjExpedienteDTOActivo` → refactor a services** |

## Anexo · Tabla de fuentes

| Fuente | Aporta |
|---|---|
| engram topic_key `expedientes/walkthrough-2026-08-07` | Consolidado de 5 walkthroughs paralelos G1..G5 |
| [`walkthrough-G1.json`](walkthrough-G1.json) | 9 forms (Splash + Menús + Cross-cutting) — método v4, 23000 bytes |
| [`walkthrough-G2.json`](walkthrough-G2.json) | 11 forms (Expedientes alta + edición) — método v4, 23387 bytes |
| [`walkthrough-G3.json`](walkthrough-G3.json) | 4 forms (Expedientes gestión) — método v4, 58527 bytes |
| [`walkthrough-G4.json`](walkthrough-G4.json) | 13 forms (Catálogos mayores) — método v4, 41444 bytes — 5 critical findings |
| [`walkthrough-G5.json`](walkthrough-G5.json) | 9 forms (Catálogos menores) — método v4, 14231 bytes |
| [`data-model.md`](data-model.md) | 25+ tablas backend con definiciones + columnas |
| [`forms.md`](forms.md) | Navegación, call paths, formularios críticos |
| [`integrations-automation.md`](integrations-automation.md) | Cross-app, Edge WebView, testing sandbox, flags |
| [`migration-matrix.md`](migration-matrix.md) | D169-D175 detallados + decisión D-new logs web-native |
| [`security-rules.md`](security-rules.md) | Autorización, riesgos de privacidad, XSS de WebView |
| [`README.md`](README.md) | Overview del estudio, hallazgos críticos, checklist |
| `data/staging/expedientes/src/forms/` (46 .form.txt + 46 .cls) | Source tree exportado |
| `data/staging/expedientes/src/classes/` | Clases de dominio |
| `data/staging/expedientes/src/modules/` | Módulos |
| `C:\00repos\codigo\00_EXPEDIENTES\staging` | Fuente READ-ONLY |
| [DOCS](../../../DOCS.md) | Technical reference raíz del blueprint |
| [CODEBASE-GUIDE](../../../CODEBASE-GUIDE.md) | Para mantenedores del blueprint |

## Checklist del documento

- [x] Scope con 46 features detalladas por dominio (G1-G5)
- [x] Walkthrough G1..G5 (46/46 forms, método v4) con JSON estructurado
- [x] Hallazgos D169-D178 con anchor links
- [x] Decisiones UX Preservar/Mejorar/Nuevo paradigma por pantalla
- [x] Decisiones arquitectura hexagonal ports
- [x] Criterios de aceptación verificables y agrupados por dimensión
- [x] 15 pendientes operacionales antes, 11 durante, 4 en go-live
- [x] 22 tickets derivables preview (TK-EX-1..22)
- [x] Tabla de decisiones referenciadas (D5-D178)
- [x] Tabla de fuentes
- [x] Idioma: español técnico neutro. Identificadores y paths sin traducir.
- [x] "The sentence that organizes" presente
- [x] "Scope del scope" presente
- [x] Sin emojis decorativos
- [x] Cross-references a DOCS, CODEBASE-GUIDE, AGENTS
- [x] Quick Navigation table
- [x] Hallazgos en tabla con severity

## Cierre del ciclo

Esta épica cierra el **ciclo de las 8 épicas** del blueprint access2web. Las 8 apps legacy cubiertas:

| # | App | PR |
|---|---|---|
| 1 | Condor | [PR #5](https://github.com/DysTelefonica/access2web-blueprint/pull/5) |
| 2 | Brass | [PR #6](https://github.com/DysTelefonica/access2web-blueprint/pull/6) |
| 3 | Expedientes | (este PR) |

Pendientes operacionales acumulados:
- 4 issues dysflow filed (#1403 cerrado, #1407 cerrado, #1408 OPEN, #1412 OPEN).
- 2 apps requieren staging de binarios (Lanzaderas D156, otras que dependen).
- Audit final de las 8 épicas refactorizadas.

## Siguiente paso

Una vez mergeado el PR, el blueprint está completo. Pendientes:
- Audit final cruzada de las 8 épicas refactorizadas.
- Decisión sobre Lanzaderas + Gestión_Riesgos + HPS + HPS_Solicitudes + NoConformidades (épicas pendientes de refactor).
- Decisión sobre `Form_Formulario1` (D167 cross-app + D178 G1 WIP).

## Cómo se aplica a access2web-blueprint

Expedientes es la segunda app con código destino planificado tras Lanzadera: `app/src/modules/expedientes/` (plan PR-08), siguiendo el mismo esqueleto hexagonal `domain/`, `ports/`, `application/`, `adapters/`, `di/`, `delivery/` (DA-1).

**Entrada cruda**: los walkthroughs G1..G5 (`walkthrough-G1.json` a `walkthrough-G5.json`) consolidan 46/46 forms con `method_version: "v4"`. El change OpenSpec `openspec/changes/expedientes-web-migration/` ya tiene proposal, 10 specs, design y tasks definidos.

**Decisiones operativas vigentes**: D86 (forma hexagonal del legacy preservada como referencia de mapeo uno-a-uno con la nueva plataforma), D87 (cada `Test_*.bas` se traduce a pytest equivalente antes de descartar el original — ningún `Test_*` se descarta sin trazabilidad), D94 (FKs conceptuales sin constraint — formalizar o documentar como referencia), D102 (booleans Text(2) → `BOOLEAN` con regla de migración explícita), D132 (XApp HTTP/JSON handshake cross-app).

**Cross-references desde otros docs**: DOCS.md §The 8 Apps lista Expedientes como mergeada con PR #4; `docs/architecture.md` §Decisiones D-<n> vigentes lista D14 (esquema por módulo), D82 (Expand and Contract); CODEBASE-GUIDE.md §Ownership de artefactos la referencia para Alembic en `app/migrations/versions/` y walkthroughs en `docs/03-aplicaciones/expedientes/walkthrough-*.json`.

## Lista de comprobación final

- [ ] Las 7 secciones del cuerpo están completas y verificadas contra los walkthroughs G1..G5.
- [ ] Los 2 anexos (Tabla de fuentes, Decisiones referenciadas) están adjuntos con referencias válidas.
- [ ] Los criterios de aceptación de §5 están todos marcados; los pendientes tienen ticket derivado.
- [ ] Los tickets derivables de §7 tienen issue-number válido en `openspec/changes/expedientes-web-migration/`.
- [ ] La sección «Cómo se aplica a access2web-blueprint» referencia el change SDD, los walkthroughs G1..G5, el código destino `app/src/modules/expedientes/`, y las decisiones D86, D87, D94, D102, D132.
- [ ] La épica se ajusta al contrato de `skills/documentation-alan-style/SKILL.md` (§3 + §8): castellano peninsular formal con usted, sin emojis decorativos.
- [ ] Las cross-references desde DOCS.md, CODEBASE-GUIDE.md, `docs/architecture.md` y los walkthrough JSONs siguen resolviendo.

---

[← Back to Expedientes README](README.md) · [← Codebase Guide](../../../CODEBASE-GUIDE.md) · [← DOCS](../../../DOCS.md)
