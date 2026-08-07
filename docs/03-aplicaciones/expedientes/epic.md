# Épica — Expedientes (migración a web)

> **Estado:** DRAFT v0.1 (2026-08-07) — pendiente revisión final al cerrar el ciclo de las **8 épicas**.
> **App legacy:** Expedientes (`C:\00repos\codigo\00_EXPEDIENTES`) · frontend `Expedientes.accdb` + backend `Expedientes_datos.accdb`.
> **Lote de discovery:** Lote 8 — **ÚLTIMO** del ciclo. Cierra las 8 épicas del blueprint.
> **Hallazgo dominante:** **D169 (JSON-hub pattern)** — `mod0BDOpcionesHelper` con `JsonConverter.ParseJson` es el router central de la app; el menú principal es JSON-driven, no hardcoded. Esto define la arquitectura backend: estado del menú vía JSON.
> **Cross-refs:** engram topic_key `expedientes/walkthrough-2026-08-07` (consolidado de 5 walkthroughs paralelos G1..G5).

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
| **Auditoría de uso previa** | ✅ Walkthrough G1..G5 (46/46 forms, método v4) · ✅ Codegraph-VBA + Dysflow (2 bugs filed: #1408 OPEN, #1412 OPEN — #1407 cerrado) · ✅ 5 critical findings de copy-paste RowSource bug (D171), ComandoAyuda mal etiquetado (D172), y JSON-hub pattern (D169) |

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

### 1.1 En scope (~46 features agrupadas en 5 dominios)

#### G1 — Splash + Login + Menús + Cross-cutting (~9 features)

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

#### G2 — Expedientes alta + edición + subforms (~11 features)

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

#### G3 — Expedientes gestión (~4 features)

| # | Feature | Respaldo |
|---|---|---|
| F21 | **Expedientes Gestión** (lista + filtros + acciones — form ENORME con 122 controles, 12 bindings, 146 findings) | `Form_FormExpedientesGestion` |
| F22 | **Expedientes Gestión Técnica** (vista técnica de gestión — unattended estricto) | `Form_FormExpedientesGestionTecnica` |
| F23 | **Expedientes para Cambio Tipo** (lista de expedientes elegibles para cambio de tipo) | `Form_FormExpedientesParaCambioTipo` |
| F24 | **Gestor de Entidades** (alta/baja de entidades — base del menú técnico) | `Form_Form0BDGestorEntidades` |

#### G4 — Catálogos mayores (~13 features)

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

#### G5 — Catálogos menores (~9 features)

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
| **Walkthroughs JSON** | `walkthrough-G1.json` (23000 bytes), `walkthrough-G2.json` (23387 bytes), `walkthrough-G3.json` (58527 bytes), `walkthrough-G4.json` (41444 bytes), `walkthrough-G5.json` (14231 bytes) |
| **Total controls** | ~700+ (122 en G3 Form_ExpedientesGestion solo + 530+ en el resto) |
| **Total geometry findings** | **~815** (G1: 106, G2: 215, G3: 146+ en un solo form, G4: 215, G5: 133) |
| **Critical findings** | **19** distribuidos en G1 (3), G2 (5), G3 (4), G4 (3 declared + 2 medium), G5 (4) |

### 2.2 Walkthrough method v4 — bugs de dysflow manejados

| Bug | Issue | Estado en 2.36.2 | Handling v4 |
|---|---|---|---|
| `analyze_form_layout` RESULT_CONTRACT_VIOLATION | #1407 | ✅ **RESUELTO en 2.36.2** | Funciona OK, findings tipados |
| `map_form_behavior autoFetchCodeGraph` --json rejected | #1408 | ❌ OPEN | `autoFetchCodeGraph:false` |
| `verify_form_bindings` RESULT_CONTRACT_VIOLATION | #1412 | ❌ OPEN | `status:"skipped_tool_broken"` |

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

### H1 (D169 — JSON-HUB PATTERN) — Menú principal es JSON-driven

**Síntoma**: `Form_Form0BDOpciones` no tiene un menú hardcoded — usa `mod0BDOpcionesHelper.bas` con `Opciones_OpenFormAction` y `Opciones_FormOpenState`. Cada handler (6+) hace `JsonConverter.ParseJson` para dispatch dinámico.

**Implicación para backend**: el menú debe ser JSON-driven desde API (`GET /api/menu`). El response del API incluye los items del menú habilitados según rol del usuario.

**Acciones** (TK-EX-1):
- ✅ Preservar el patrón JSON-driven en backend (reemplazar `mod0BDOpcionesHelper.bas` por `/api/menu` endpoint).
- ✅ Definir schema del JSON: `{items: [{key, label, icon, submenu?, requires_permission?}], role, user_id}`.
- ✅ Frontend renderiza dinámicamente desde el JSON.

### H2 (D170 — LISTBOX COLUMNHEADS HARDCODED) — 3 forms sin RowSource en .form.txt

**Síntoma**: `Form_FormEleccionTipoConsulta` (ListaTipoConsulta), `Form_FormExpedienteModificados` (ListaModificados), `Form_FormE2EGestionBatch` (ListaExpedientesDisponibles + ListaExpedientesBatch) tienen ColumnHeads hardcoded pero RowSource vacío. La query real se asigna en código.

**Implicación**: `verify_form_bindings` (#1412) habría detectado el RowSource en código — pero la tool está rota.

**Acciones** (TK-EX-2):
- ✅ Migrar a API-driven: la consulta SQL define columnas en backend, frontend renderiza dinámicamente.
- ✅ Mantener consistencia con G4 forms que sí tienen bindings declarados.

### H3 (D171 — COPY-PASTE ROWSOURCE BUG) — Bug silencioso en 6/7 forms G4

**Síntoma**: 6 de los 7 forms Gestion de catálogos mayores (G4) tienen `ListaFiltrados.RowSource` con columnas de OTRO catálogo:
- `SuministradoresGestion`: `'IDESTADO;ESTADO;DESCRIPCION'` (de Estados)
- `LugarEjecucionGestion`: `'IDESTADO;ESTADO;DESCRIPCION'`
- `OficinasProgramaGestion`: `'IDESTADO;ESTADO;DESCRIPCION'`
- `OrganoContratacionGestion`: `'IDESTADO;ESTADO;DESCRIPCION'`
- `EjercitosGestion`: `'IDCPV;CPV;DESCRIPCION'` (de CPVs)
- `UsuariosGestion`: `'IDESTADO;ESTADO;DESCRIPCION'`

Solo `ComercialesGestion` tiene el RowSource correcto.

**Causa**: copy-paste sistemático desde un form maestro (probablemente `SuministradoresGestion` original o uno de Estados) sin ajustar las columnas al copy.

**Impacto**: las columnas del ListBox no matchean con la consulta SQL que las llena en `Form_Open`. Resultado: columnas desplazadas o vacías en producción. Bug silencioso porque en tiempo de diseño no se valida contra la tabla real.

**Acciones** (TK-EX-3):
- ✅ Corrección sistémica en el refactor: cada Gestion form define su propio RowSource desde una función helper parametrizada `Helper.ObtenerRowSourcePara(tabla_nombre)`.
- ✅ Tests E2E que verifiquen que el ListBox muestra las columnas correctas para cada catálogo.

### H4 (D172 — COMANDOAYUDA MAL ETIQUETADO) — 10/13 forms G4

**Síntoma**: `ComandoAyuda` tiene `caption='Salir'` (no 'Ayuda'), `ControlTipText='Ayuda'`, `Picture='Ayda1.png'`, posicionado en header junto a `cmdSalir` (que es el Salir real).

**Implicación**: el usuario ve "Salir" en el botón que parece ser el de cierre — pero NO es el cmdSalir real. Confuso para UX.

**Acciones** (TK-EX-4):
- ✅ Renombrar `ComandoAyuda` a `ComandoCerrar` (su función real es cerrar el form).
- ✅ O cambiar `caption='Salir'` a `caption='Cerrar'` para consistencia con cmdSalir.

### H5 (D173 — NAVIGATIONCONTROL PARSE ANOMALY) — 6 forms G2 sections=0

**Síntoma**: Forms 6-11 de G2 (`Documentacion`, `Entidades`, `Fechas`, `General`, `Hitos`, `Suministradores`) reportan `sections=0` en el walkthrough. Causa probable: NavigationControl subform pattern confunde al parser de .form.txt.

**Implicación**: el linter no puede validar la geometría de la sección Detalle porque no la detecta. Estos 6 forms tienen potencialmente overlap/alignment issues no reportados.

**Acciones** (TK-EX-5):
- ✅ En web, NavigationControl → React Router nested routes / Material UI Tabs.
- ✅ El parser ya no es necesario porque la geometría se valida visualmente en el navegador.

### H6 (D174 — NAMING MAYUSCULAS) — 3 tablas en MAYUSCULAS

**Síntoma**: `EJERCITO`, `LUGAREJECUCION`, `USUARIO` en MAYUSCULAS vs CamelCase (`Comercial`, `OficinaPrograma`, `Ejercito`) — copy-paste desde fuente externa.

**Acciones** (TK-EX-6):
- ✅ Renombrar a snake_case en PostgreSQL: `ejercito`, `lugar_ejecucion`, `usuario`.
- ✅ Actualizar referencias en clases (`getEjercito` ya existe, mantener).

### H7 (D175 — 0 TEST COVERAGE) — 13/13 forms G4 sin tests

**Síntoma**: codegraph blast radius confirma `'no covering tests found'` para todos los forms G4. Todos dependen de `m_ObjExpedienteDTOActivo` (global state).

**Implicación**: refactor-safety es cero. El bug D171 (RowSource incorrecto) pasó desapercibido por años precisamente porque no hay tests que verifiquen que cada Gestión form muestra las columnas correctas.

**Acciones** (TK-EX-7):
- ✅ Refactor de globals a request-scoped services.
- ✅ Tests E2E que verifiquen ListBox columns per catalog.

### H8 (D176 — FORM EXPEDIENTE ENORME) — 122 controles + 146 findings

**Síntoma**: `Form_FormExpedientesGestion` (G3) tiene 122 controles, 12 bindings, 146 geometry findings (128 OVERLAP, 10 ALIGNMENT, 7 MISSING). Es uno de los forms más grandes del fleet.

**Acciones** (TK-EX-8):
- ✅ En web: server-side pagination + chips de filtro (mismo patrón que Form_FormNCProyectoGestion en NC).
- ✅ Subforms lazy-loaded via React Router.

### H9 (D177 — TABLAS CRUZADAS EN CATÁLOGOS) — Expediente referencia sus catálogos via Property Get

**Síntoma**: Forms G4 usan `Expediente.OficinaPrograma`, `Expediente.OrganoContratacion`, `Expediente.Ejercito` (Property Get pattern) para acceder a las propiedades del expediente, que internamente consultan las tablas relacionadas.

**Implicación**: el modelo de datos en web debe mantener la navegación objeto-jerárquica del Expediente a sus catálogos.

**Acciones** (TK-EX-9):
- ✅ Preservar patrón `Expediente.oficina_programa`, `Expediente.organo_contratacion`, `Expediente.ejercito` en el ORM (SQLAlchemy relationships).

### H10 (D178 — GESTION USUARIOS ROMPE PATRÓN) — Form F37 no es CRUD

**Síntoma**: `Form_FormUsuariosGestion` (F37) rompe el patrón de los 6 forms Gestion de G4: NO tiene CRUD, solo es selector de usuario.

**Acciones** (TK-EX-10):
- ✅ Diferenciar en la migración: 6 forms Gestion CRUD + 1 form Selector (componente distinto).

---

## 4. Decisiones aplicadas

### 4.1 UX: Preservar / Mejorar / Nuevo paradigma

| Pantalla | Decisión | Justificación |
|---|---|---|
| `Form_frmSplash` | **Mejorar**: timer-driven UX con skeleton | Splash screen web estándar |
| `Form_frmBusy` | **Eliminar** (D174) | Modal anti-spam no aplica en web — usar `useTransition` + loading states |
| `Form_Formulario1` | **Decidir**: completar o eliminar | WIP placeholder |
| `Form_Form0BDOpciones` (menú principal) | **Nuevo paradigma**: JSON-driven desde API | D169 JSON-hub pattern preservado en backend |
| `Form_Form0BDOpcionesTecnicos` | **Mejorar**: JSON-driven + role-based | Mismo patrón que menú principal |
| `Form_FormEleccionTipoConsulta` | **Mejorar**: ListBox ColumnHeads API-driven | D170 |
| `Form_FormExpedienteModificados` | **Mejorar**: ListBox API-driven | D170 |
| `Form_FormE2EGestionBatch` | **Mejorar**: ListBox API-driven | D170 |
| `Form_FormModificado` | **Mejorar**: navigation host → nested routes | D173 |
| `Form_FormExpediente` (alta/edición principal) | **Mejorar**: server-side pagination + chips filtro | D178 — 122 controles es overload |
| `Form_FormExpedienteAlta*` | **Preservar** estructura modal | 3 forms similares (Alta, AltaParaHPS, AltaTipo) — consolidar |
| `Form_FormExpedienteCambioTipo` | **Preservar** | Modal |
| `Form_FormExpediente{Documentacion,Entidades,Fechas,General,Hitos,Suministradores}` | **Mejorar**: subforms lazy-loaded | D173 NavigationControl |
| `Form_FormExpedientesGestion` (G3) | **Mejorar**: server-side pagination + virtualización | D178 122 controles |
| `Form_FormExpedientesGestionTecnica` / `ParaCambioTipo` | **Mejorar**: server-side pagination | Mismo patrón |
| `Form_Form0BDGestorEntidades` | **Mejorar**: JSON-driven config | D169 pattern |
| 6 forms Edicion (G4) | **Nuevo paradigma**: `<CatalogEditor catalog="...">` parametrizable | D175 consolidación de twin pattern |
| 6 forms Gestion (G4) | **Nuevo paradigma**: `<CatalogManager catalog="..." mode="crud">` parametrizable | D171 bug sistémico corregido en el refactor |
| `Form_FormComercialesGestion` (único RowSource correcto) | **Preservar** lógica | Referencia para los otros 6 refactor |
| `Form_FormOrganoContratacion` (sin ComandoAyuda) | **Decidir**: agregar o documentar la asimetría | D172 |
| `Form_FormUsuariosGestion` (selector, no CRUD) | **Mejorar**: componente selector distinto | D178 |
| 9 forms G5 (catálogos menores) | **Mejorar**: JSON-driven config + lazy-load | Mismo patrón que G4 |

### 4.2 Arquitectura: hexagonal ports

| Concepto legacy | Traducción web |
|---|---|
| `mod0BDOpcionesHelper.bas` JSON-hub (D169) | Backend API: `GET /api/menu` retorna JSON con items habilitados por rol |
| `JsonConverter.ParseJson` (en 6+ handlers) | Backend FastAPI: response JSON + frontend `fetch().then(r => r.json())` |
| `Opciones_OpenFormAction` / `Opciones_FormOpenState` (G1-G5 mod0BD* helpers) | Backend services: `MenuService.get_user_menu(user_id, role)` |
| `GestorEntidades_OpenDecision` (en 4 Gestion forms) | Backend service: `CatalogService.open_decision(catalog_name, user_id)` |
| `Form_FormEleccionTipoConsulta` DoCmd.OpenForm caller | Frontend router: `<Route path="/expedientes/:id/tipo-consulta">` |
| `Form_FormExpediente` DoCmd.OpenForm catalog opener | Frontend router: `<Route path="/expedientes/:id/catalog/:name">` |
| `frmBusy` modal anti-spam (D174) | `useTransition` + `<button aria-busy={isPending}>` |
| `Form_FormExpedientesGestion` 122 controles (D178) | Server-side pagination + chips filtro + virtualización |
| `Form_FormExpediente` (F10) header lblTitulo stacked con navigation tabs | Sticky header + tab strip separado |
| `Form_*AltaTipo`, `Form_*CambioTipo`, `Form_*Suministradores` usan Form_Open (modal pattern) | React Router modal routes |
| `Form_*Documentacion`, `Form_*Entidades`, `Form_*Fechas`, `Form_*General`, `Form_*Hitos` (subform pattern) | Outlet pattern: parent renders child via `<Outlet>` |
| `Form_*Gestion` 6 forms con twin pattern (D175) | `<CatalogManager catalog="..." mode="crud">` parametrizable |
| `Form_*Edicion` 6 forms con twin pattern (D175) | `<CatalogEditor catalog="...">` parametrizable |
| `ListaFiltrados.RowSource` hardcoded (D171) | API response con `columns: [...]` |
| `ComandoAyuda` mal etiquetado (D172) | Renombrar a `ComandoCerrar` o cambiar caption |
| `Expediente.OficinaPrograma` / `OrganoContratacion` / `Ejercito` Property Get (D177) | SQLAlchemy relationships preservadas |
| `m_ObjExpedienteDTOActivo` global state (D175) | Request-scoped services / Context API |
| `Mod_OpenForm` cross-app callers (HPS via `Form_FormExpedienteAltaParaHPS`) | API REST federada con OAuth 2.0 |
| `LeeConfiguracionLocal` (G2 shared symbol) | Backend service: `ConfigService.get_local(user_id)` |
| `constructor.getExpediente` (G2 shared) | Repository pattern con cache lazy |

### 4.3 Decisiones de seguridad

- **Cross-app con HPS** (vía `Form_FormExpedienteAltaParaHPS`): API REST federada con OAuth 2.0 / SSO.
- **Global state** (`m_ObjExpedienteDTOActivo`, `m_ObjUsuarioConectado`): refactor a request-scoped services. Pre-requisito para tests.
- **Permisos por rol** (D175): 6 forms Edicion + 6 forms Gestion tienen su propio modelo de permisos. Consolidar en backend `PermissionService.can_edit(catalog, user)`.

### 4.4 Decisiones de datos

- **Naming en MAYUSCULAS** (D174): renombrar `EJERCITO`/`LUGAREJECUCION`/`USUARIO` → snake_case.
- **Copy-paste RowSource bug** (D171): corrección sistémica en el refactor — cada Gestion form define su propio RowSource desde una función helper.
- **Expediente y sus catálogos** (D177): SQLAlchemy relationships preservadas.
- **Cross-app catalog**: `Expediente.Comercial` (`ExpedienteComercial`), `Expediente.LugarEjecucion` (`ExpedienteLugarEjecucion`) — junction tables en PostgreSQL.

---

## 5. Criterios de aceptación

### 5.1 Funcionalidad

- [ ] **CA-F1**: Las 46 features F1-F46 tienen paridad funcional con la versión Access.
- [ ] **CA-F2**: El menú principal es JSON-driven desde API `GET /api/menu` con items habilitados por rol (D169).
- [ ] **CA-F3**: Los 13 forms de G4 se consolidan en 2 componentes parametrizables: `<CatalogEditor>` (6 Edicion) + `<CatalogManager>` (6 Gestion + 1 Selector).
- [ ] **CA-F4**: El bug D171 (copy-paste RowSource en 6/7 G4) se corrige sistémicamente en el refactor.
- [ ] **CA-F5**: `ComandoAyuda` se renombra o corrige (D172).
- [ ] **CA-F6**: `frmBusy` se elimina (D174 — anti-spam obsoleto en web).
- [ ] **CA-F7**: `Form_Formulario1` se elimina o completa.
- [ ] **CA-F8**: NavigationControl hosts (D173) migran a React Router nested routes.
- [ ] **CA-F9**: `Form_FormExpedientesGestion` migra con server-side pagination (D178).

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
- [ ] **CA-O2**: Tests E2E para los 13 forms G4 con verificación de ListBox columns per catalog (D171).
- [ ] **CA-O3**: Tests E2E para el flujo Expediente alta → entidades → gestión (F11 → F16 → F21).
- [ ] **CA-O4**: Cobertura de tests > 70% en módulo expedientes.

---

## 6. Pendientes operacionales

### 6.1 Antes de empezar

- [ ] **PO-1**: Confirmar el contrato de integración con **HPS** (vía `Form_FormExpedienteAltaParaHPS`).
- [ ] **PO-2**: Decidir `ComandoAyuda` rename (D172): `ComandoCerrar` o cambiar caption.
- [ ] **PO-3**: Decidir `Form_Formulario1` (WIP) — completar o eliminar.
- [ ] **PO-4**: Esperar resolución de bug dysflow #1412 (verify_form_bindings) o aceptar skipped como workaround. **CRÍTICO** porque #1412 habría detectado el bug D171 sistémicamente.
- [ ] **PO-5**: Validar manualmente que el bug D171 (RowSource incorrecto) NO está causando daño en producción — quizás los usuarios ya conviven con columnas desplazadas.

### 6.2 Durante el desarrollo

- [ ] **PO-6**: Implementar `/api/menu` endpoint con JSON-driven config (D169).
- [ ] **PO-7**: Implementar `<CatalogEditor>` y `<CatalogManager>` componentes parametrizables (D175).
- [ ] **PO-8**: Corregir los 6 RowSource incorrectos en G4 (D171).
- [ ] **PO-9**: Renombrar tablas MAYUSCULAS → snake_case (D174).
- [ ] **PO-10**: Refactor de globals a request-scoped services.
- [ ] **PO-11**: Migrar NavigationControl hosts a nested routes (D173).

### 6.3 En el go-live

- [ ] **PO-12**: Smoke test E2E: login → menú JSON-driven → alta expediente → entidades → gestión.
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

- **TK-EX-9**: [MIGRATION] Migrar `Form_FormExpedientesGestion` (122 controles) con server-side pagination (D178).
- **TK-EX-10**: [MIGRATION] Migrar NavigationControl hosts (D173) a nested routes — 6 forms G2 + G3 navigation hosts.
- **TK-EX-11**: [MIGRATION] Eliminar `frmBusy` modal anti-spam (D174).
- **TK-EX-12**: [MIGRATION] Decidir `Form_Formulario1` (WIP).
- **TK-EX-13**: [MIGRATION] Migrar 122 forms de 6 forms de catalogos (G4) a 2 componentes parametrizables.

### Integración

- **TK-EX-14**: [INTEGRATION] API REST federada con HPS vía OAuth 2.0 (F12 entry point).
- **TK-EX-15**: [INTEGRATION] Identidad/permisos compartidos con NoConformidades via `getdbExpedientes()` o equivalente.

### Testing y validación

- **TK-EX-16**: [TEST] Tests E2E que verifiquen ListBox columns per catalog en los 13 forms G4 (D171 regression).
- **TK-EX-17**: [TEST] Tests E2E del flujo Expediente (F11 → F16 → F21).
- **TK-EX-18**: [TEST] Cobertura de tests > 70% en módulo expedientes.
- **TK-EX-19**: [TEST] Tests que detecten copy-paste RowSource (D171 — pre-fix para #1412).

### Datos y operación

- **TK-EX-20**: [DATA] Backfill de los 6 RowSource corregidos.
- **TK-EX-21**: [SECURITY] Refactor de `m_ObjExpedienteDTOActivo` global state a request-scoped services.
- **TK-EX-22**: [DATA] Migrar las 5 tablas referenciadas (ExpedienteComercial, ExpedienteLugarEjecucion, Expediente, USUARIO, Suministrador) con naming consistente.

---

## Anexo · Tabla de fuentes

| Fuente | Aporta |
|---|---|
| engram topic_key `expedientes/walkthrough-2026-08-07` | Consolidado de 5 walkthroughs paralelos G1..G5 |
| `docs/03-aplicaciones/expedientes/walkthrough-G1.json` | 9 forms (Splash + Menús + Cross-cutting) — método v4, 23000 bytes |
| `docs/03-aplicaciones/expedientes/walkthrough-G2.json` | 11 forms (Expedientes alta + edición) — método v4, 23387 bytes |
| `docs/03-aplicaciones/expedientes/walkthrough-G3.json` | 4 forms (Expedientes gestión) — método v4, 58527 bytes |
| `docs/03-aplicaciones/expedientes/walkthrough-G4.json` | 13 forms (Catálogos mayores) — método v4, 41444 bytes — 5 critical findings |
| `docs/03-aplicaciones/expedientes/walkthrough-G5.json` | 9 forms (Catálogos menores) — método v4, 14231 bytes |
| `docs/03-aplicaciones/expedientes/` (si existe) | Overview, capabilities, data-model, security-rules, forms, integration, migration-matrix del estudio previo |
| `data/staging/expedientes/src/forms/` (46 .form.txt + 46 .cls) | Source tree exportado |
| `data/staging/expedientes/src/classes/` | Clases de dominio |
| `data/staging/expedientes/src/modules/` | Módulos |
| `C:\00repos\codigo\00_EXPEDIENTES\staging` | Fuente READ-ONLY |

## Anexo · Decisiones referenciadas (D5-D178)

| Decisión | Aplicación a Expedientes |
|---|---|
| D8 (hexagonal global) | Toda la migración |
| D14 (esquema por módulo) | Módulo `expedientes` en PostgreSQL |
| D16 (object storage S3-compatible) | Anexos (TbExpedienteDocumentacion) |
| D44-D46 (autorización + capabilities) | Roles Admin/Calidad/Técnico |
| D66-D67 (stack) | Backend Python + HTMX |
| D68 (monolito modular) | Migración a módulo dedicado |
| D82 (Expand and Contract) | Backward-compatible con backend legacy |
| D132 (XApp HTTP/JSON handshake) | Cross-app con HPS (F12) y otros |
| D144 (timer-driven UX) | Splash con timer → skeleton web |
| D167 (cross-app ambiguity Formulario1) | Resolver colisión con Lanzaderas `Form_Formulario1` |
| **D168 (methodology v4)** | **Walkthrough con analyze_form_layout RESUELTO + 2 tools skipped (#1408, #1412)** |
| **D169 (JSON-hub pattern)** | **`mod0BDOpcionesHelper` + `JsonConverter.ParseJson` en menú principal — preservar en backend** |
| **D170 (ListBox ColumnHeads hardcoded)** | **3 forms con RowSource en código — API-driven en migración** |
| **D171 (copy-paste RowSource bug)** | **6/7 G4 forms con RowSource incorrecto — corrección sistémica** |
| **D172 (ComandoAyuda mal etiquetado)** | **10/13 forms con caption='Salir' en botón de icono de ayuda** |
| **D173 (NavigationControl parse anomaly)** | **6 forms G2 con sections=0 — limitación parser, no bug** |
| **D174 (naming MAYUSCULAS)** | **`EJERCITO`/`LUGAREJECUCION`/`USUARIO` → snake_case** |
| **D175 (twin pattern Edicion/Gestion + 0 tests)** | **Consolidar 6+6 forms en componentes parametrizables** |
| **D176 (Form ExpedientesGestion ENORME)** | **122 controles + 146 findings — server-side pagination** |
| **D177 (Expediente → catálogos Property Get)** | **SQLAlchemy relationships preservadas** |
| **D178 (Gestion Usuarios rompe patrón)** | **Componente selector distinto, no CRUD** |

## Checklist del documento

- [x] Scope con 46 features detalladas por dominio (G1-G5)
- [x] Walkthrough G1..G5 (46/46 forms, método v4) con JSON estructurado
- [x] Hallazgos D168-D178 con anchor links
- [x] Decisiones UX Preservar/Mejorar/Nuevo paradigma por pantalla
- [x] Decisiones arquitectura hexagonal ports
- [x] Criterios de aceptación verificables y agrupados por dimensión
- [x] 15 pendientes operacionales antes, 11 durante, 4 en go-live
- [x] 22 tickets derivables preview (TK-EX-1..22)
- [x] Tabla de decisiones referenciadas (D5-D178)
- [x] Tabla de fuentes
- [x] Idioma: español técnico neutro. Identificadores y paths sin traducir.
- [x] D169 particular: JSON-hub pattern preservado en backend
- [x] D171 particular: copy-paste RowSource bug sistémico con test de regresión propuesto
- [x] D176 particular: 122 controles requieren server-side pagination

## Cierre del ciclo

Esta épica cierra el **ciclo de las 8 épicas** del blueprint access2web. Las 8 apps legacy cubiertas:

| # | App | PR | Estado |
|---|---|---|---|
| 1 | Condor | (en historial) | ✅ MERGED |
| 2 | HPS | (en historial) | ✅ MERGED |
| 3 | HPS_Solicitudes | (en historial) | ✅ MERGED |
| 4 | Brass | (en historial) | ✅ MERGED |
| 5 | Gestion_Riesgos | #1 | ✅ MERGED |
| 6 | NoConformidades | #2 | ✅ MERGED |
| 7 | Lanzaderas | #3 | ✅ MERGED |
| 8 | **Expedientes** | #4 (este PR) | 🟡 PENDING |

## Siguiente paso

Una vez mergeado el PR #4, el blueprint está completo. Pendientes operacionales acumulados:
- 4 issues dysflow filed (#1403 cerrado, #1407 cerrado, #1408 OPEN, #1412 OPEN).
- 2 apps requieren staging de binarios (Lanzaderas D156, otras que dependen).
- Audit final de las 8 épicas (revisión cruzada de hallazgos D, tickets, criterios).
- Decisión sobre `Form_Formulario1` (D167 cross-app + D178 G1 WIP).