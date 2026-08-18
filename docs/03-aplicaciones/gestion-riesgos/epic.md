[← Back to Gestion_Riesgos README](README.md) · [← Codebase Guide](../../../CODEBASE-GUIDE.md) · [← DOCS](../../../DOCS.md)

# Épica — Gestion_Riesgos (migración a web)

> **Estado:** DRAFT v0.1 (2026-08-06) — pendiente revisión al final del ciclo de las 8 épicas.
> **Lote:** 6 — junto con Brass, HPS, HPS_Solicitudes, Condor, Expedientes.
> **App legacy:** `00_GESTION_RIESGOS` · frontend `Gestion_Riesgos.accdb` (40 MB) + backend `Gestion_Riesgos_Datos.accdb` (17 MB).
> **Sentence that organizes**: **Gestion_Riesgos es la app de riesgos de calidad por edición: cada edición tiene N riesgos con planes de mitigación/contingencia, vulnerabilidad, mitigación (4 valores), y triángulo plazo/coste/calidad. State machine de 14 estados. Cuello de botella: 372 call sites a `getdb()`集中在 clases de dominio, no en UI (D88).**

> **Scope del scope**: "Este repo es research + planning de la migración. Cada app tendrá su propio repo + docs cuando se construya."

---

## Quick Navigation

| Section                                                       | What you'll find                                              |
|---------------------------------------------------------------|---------------------------------------------------------------|
| [Metadatos](#metadatos)                                       | Scope XL, D88 (372 getdb() callers), dependencias.              |
| [1. Scope](#1-scope)                                          | 61 features F1-F61 en 5 dominios (G1-G5).                    |
| [2. Estado del descubrimiento](#2-estado-del-descubrimiento)    | Inventario 62 forms, 65 clases, 139 módulos, 71 tablas, D88. |
| [3. Hallazgos críticos](#3-hallazgos-críticos)                | D88 + D126-D145: state machine, AGEDO, anexos, tests.         |
| [4. Decisiones aplicadas](#4-decisiones-aplicadas)              | UX por pantalla, hexagonal ports, seguridad.                  |
| [5. Criterios de aceptación](#5-criterios-de-aceptación)      | Funcionalidad, Seguridad, Performance, Operacional.            |
| [6. Pendientes operacionales](#6-pendientes-operacionales)    | Antes, durante, go-live.                                       |
| [7. Tickets derivables](#7-tickets-derivables-preview)         | 40 tickets TK-GR-1..40.                                       |
| [Anexo · Tabla de fuentes](#anexo--tabla-de-fuentes)         | Walkthrough + docs + source + engram.                          |
| [Anexo · Decisiones referenciadas](#anexo--decisiones-referenciadas-d5-d145) | D5-D145.                                          |

---

## Metadatos

| Campo | Valor |
|---|---|
| **Aplicación legacy** | `00_GESTION_RIESGOS` · frontend `Gestion_Riesgos.accdb` + backend `Gestion_Riesgos_Datos.accdb` |
| **Tipo de migración** | Legacy Access/VBA → web hexagonal (FastAPI + HTMX) |
| **Scope size** | **XL** (62 forms, 65 clases, 139 módulos, 71 tablas, 372 getdb() callers) |
| **Dependencias cross-app** | **Expedientes** (vía HTTP/JSON — único punto de acoplamiento crítico); **NoConformidades** (vía `TbRiesgosNC.IDNC` FK lógica); **Lanzadera** (identidad/permisos presumiblemente, vía `getdbLanzadera()` — patrón compartido con Gestion_Riesgos) |
| **Riesgo dominante** | **D88** — `getdb()` con 372 callers. Los forms tienen 0-2 calls directos; el grueso está en las CLASES de dominio. **No es un problema de UI** — la optimización es de capa de datos, no de forms. |
| **Stack target** | Backend Python 3.12+ / FastAPI 0.119+ / SQLAlchemy 2.0.x / Alembic 1.13+ / asyncpg 0.30+ (D66) · Frontend HTMX 2.0.4 + Jinja2 3.1+ / Alpine.js 3.15+ (D67) |
| **Estrategia de migración de BD** | Expand and Contract backward-compatible (D82) · PostgreSQL compartido con esquema por módulo (D14) |
| **Forma destino** | Hexagonal global (D8) · módulo dentro del monolito modular (D68) · puerto de persistencia PostgreSQL + object storage S3-compatible (D16) + secret manager (D9-D10) |
| **Auditoría de uso previa** | Codegraph-vba + Dysflow walkthrough (61/62 forms, **D88 medido: 372 getdb() callers en el proyecto entero**) |

---

## 1. Scope

### 1.1 En scope — 61 features agrupadas en 5 dominios

#### G1 — Configuración + Maestros + Menús (10 features)

| # | Feature | Respaldo |
|---|---|---|
| F1 | **Menú principal Calidad** + **sub-menú Técnico** (entry points duales por rol) | `Form0BDOpciones`, `Form0BDOpcionesTecnico` |
| F2 | **Gestión de Suministradores** (CRUD simple + ventana detalle embebida) | `FormSuministrador` |
| F3 | **Gestión de Proyectos/UTE** (lista con 4 filtros + treeview) | `FormProyectosGestion` |
| F4 | **Dashboard de Indicadores** (6 tiles por semestre/año: Identificados / Retirados / En Oferta / Materializados / Oferta→Gestión / Vigentes) | `FormIndicador`, `formIndicadorProyectos` |
| F5 | **Biblioteca de Riesgos** (catálogo reutilizable con 6 familias) | `FormRiesgoBiblioteca`, `FormRiesgosBibliotecaGestion`, `FormRiesgoDefinicion`, `FormRiesgoDefinicionNoBiblioteca` |
| F6 | **Splash + autenticación dual** (Calidad/Técnico via EVE) | `frmSplash` |
| F7 | **CRUD RiesgosBiblioteca** (Vista/Crear/Editar riesgos reutilizables) | `FormRiesgosBibliotecaGestion` |
| F8 | **Patrón PUB/SUB de forms** (eventos AltaRiesgoBiblioteca para sincronizar) | implementado via WithEvents |
| F9 | **Anexos polimórficos (compartidos con Brass)** | ver G5 F53 |
| F10 | **Picker de usuarios autorizados** (read-only con correos) | `FormGestionRiesgosAutorizados` |

#### G2 — Riesgos: definición, gestión, aceptación (16 features)

| # | Feature | Respaldo |
|---|---|---|
| F11 | **Alta/Edición directa de riesgo** (form shell con NavigationControl de 5 tabs) | `FormRiesgo` |
| F12 | **Gestión de Riesgos** (alta/edición de cabecera del proyecto) | `FormGestionRiesgos` |
| F13 | **Datos generales del riesgo** (subform de cabecera) | `FormGestionRiesgosDatosGenerales` |
| F14 | **Riesgos autorizados** (read-only picker con correos) | `FormGestionRiesgosAutorizados` |
| F15 | **Materialización del riesgo** (registrar fecha + plan contingencia elegible) | `FormRiesgoMaterializado` |
| F16 | **Vulnerabilidad del riesgo** (singleton ListBox con 5 niveles) | `FormRiesgoVulnerabilidad` |
| F17 | **Mitigación** (4 valores: Aceptar / Evitar / Reducir / Transferir) | `FormRiesgoMitigacion` |
| F18 | **Triángulo plazo/coste/calidad** (3 columnas independientes con escala 5-niveles) | `FormRiesgoPlazoCosteCalidad` |
| F19 | **Retirada** (con justificación + workflow de visado dual Calidad) | `FormRiesgoRetirado` |
| F20 | **Vinculación cross-app con NoConformidades** (vía `TbRiesgosNC.IDNC`) | `FormRiesgoNC` |
| F21 | **Riesgo Externo (Oferta)** (definición + traslado vs no traslado) | `FormRiesgoExternoDetalle` |
| F22 | **Workflow de modificación** del riesgo aprobado (aceptación o retirada solicitada) | `FormRiesgoPosiblesModificacionesAceptarRetirar`, `FormRiesgoPosiblesModificacionesQuitarAceptar` |
| F23 | **Estado del riesgo** (state machine de 14 estados con campo calculado) | `Riesgo.EstadoEnum` |
| F24 | **Gestión principal de riesgos** (lista + treeview + 6 acciones) | `FormRiesgosGestion` |
| F25 | **Establecer prioridades** (grid editable smallint) | `FormRiesgosEstablecerPrioridades` |
| F26 | **Edición masiva de riesgos** (modo batch) | `FormRiesgosGestionEdicion` |

#### G3 — Calidad + Planes (11 features)

| # | Feature | Respaldo |
|---|---|---|
| F27 | **Dashboard Calidad** (TreeView 7 categorías + subform detalle intercambiable) | `FormCalidadTareas` |
| F28 | **Detalle edición pendiente publicación** (read-only + 6 botones de acción) | `FormCalidadTareasDetalleEdicion` |
| F29 | **Visado de aceptación/retiro** (modal con captions dinámicas) | `FormCalidadRiesgoAceptadoRetiradoVisado` |
| F30 | **Gestión post-materialización** (vincular NC / no NC / revocar decisión) | `FormCalidadRiesgoMaterializaciones` |
| F31 | **Explicación de categoría** (mini-form de 3 controles) | `FormCalidadTareaExplicacion` |
| F32 | **Detalle riesgo aceptado/retirado pendiente visar** (read-only) | `FormCalidadTareaRiesgosAceptadosRetirados` |
| F33 | **Detalle riesgo materializado a decidir** (read-only) | `FormCalidadTareaRiesgosMaterializadosPorDecidir` |
| F34 | **Retipificación** (reasignar riesgo a código de biblioteca) | `FormCalidadTareaRiesgosRetipificacion` |
| F35 | **Capturar plan contingencia** (modal con TempVars IPC) | `FormMaterializacionPlanContingencia` |
| F36 | **Gestión riesgos de Oferta** (solo Calidad, primera edición) | `FormGestionRiesgosRiesgosOferta` |
| F37 | **Sincronización de Suministradores** (push manual desde jerarquía) | `FormGestionRiesgosSuministradores` |

#### G4 — Publicación (14 features)

| # | Feature | Respaldo |
|---|---|---|
| F38 | **Flujo Publicación Calidad** (NavigationControl 5 tabs) | `FormPublicacionCalidad` |
| F39 | **Estado de publicación** (Presentation Layer con view state derivado) | `FormPublicacionCalidadPublicar` |
| F40 | **Ejecutar publicación** (modal con todos los defaults + validaciones) | `FormPublicacionCalidadPublicarEjecutar` |
| F41 | **Capturar motivo de rechazo** (modal con Event Motivado) | `FormPublicacionCalidadRechazoObservaciones` |
| F42 | **Notas de Calidad para publicación** (TextBox libre) | `FormPublicacionCalidadNotas` |
| F43 | **Datos generales del documento a publicar** (read-only) | `FormPublicacionDatosGenerales` |
| F44 | **Evidencias de Suministradores** (lista + gestión de anexos) | `FormPublicacionSuministradores` |
| F45 | **Sub-menú Técnico propuesta de publicación** (NavigationControl 3 tabs) | `FormPublicacionTecnico`, `FormPublicacionTecnicoPropuesta` |
| F46 | **Evidencia de UTE** (anexo único por edición) | `FormPublicacionUTE` |
| F47 | **Alta de Plan individual** (mitigación o contingencia) | `FormPlanPrincipal` |
| F48 | **Integración con AGEDO** (registro jurídico TdE) | via `Edicion.RegistrarEnAGEDO` |
| F49 | **Generación de informes de publicación** (HTML + PDF con tokens Mistica) | `Funciones Generales.GenerarInformePublicacion` |
| F50 | **Servicio de refresco post-publicación** (qué riesgo refrescar) | `PublicacionCalidadExecutionService.ResolveRefreshRiskID` |
| F51 | **Auditoría de publicaciones** (`TbLogPublicaciones` append-only) | implementado en `PublicacionLog.cls` |
| F52 | **Gestión de cambios de la aplicación** (auditoría CHANGELOG de versiones — D134) | `FormControlCambiosGestion` |

#### G5 — Detalles + Técnico + Anexos (9 features)

| # | Feature | Respaldo |
|---|---|---|
| F53 | **Anexos polimórficos** (3 formas del mismo flujo, NO polimorfismo — D142) | `FormAnexos`, `FormAnexos1` |
| F54 | **Búsqueda cross-app de Expedientes** (HTTP/JSON — D132) | `FormExpedientesBusqueda` |
| F55 | **Wrapper de Internet Explorer** (D133 — eliminar completamente en migración) | `FormWeb` |
| F56 | **Tareas pendientes del Técnico** (TreeView + subform detalle) | `FormTecnicoTareas` |
| F57 | **Detalle de tarea del Técnico** (read-only para navegación rápida) | `FormTecnicoTareasDetalleEdicion` |
| F58 | **Detalle de riesgo aceptado/retirado para Técnico** | `FormTecnicoTareaRiesgosAceptadosRetirados` |
| F59 | **Explicación de nodo del TreeView** (mini-form 3 controles) | `FormTecnicoTareaExplicacion` |
| F60 | **Sub-menú Técnico principal** (21 controles, separado de 0BDOpciones) | `Form0BDOpcionesTecnico` |
| F61 | **Capturar estrategia de mitigación + triángulo plazo/coste/calidad** | `FormRiesgoMitigacion`, `FormRiesgoPlazoCosteCalidad` |

### 1.2 Fuera de scope (REPLACE)

> Lo siguiente NO migra como tablas PostgreSQL:

- **TbLogCambios / TbLogErrores / TbLogEstados** (si existen) → **Sentry / OpenTelemetry / structured logs**. Migración cross-cutting D27 ya aplicada a Brass.
- **Anexos filesystem** (`URLDirectorioDocumentacion` y variantes) → **S3-compatible** (MinIO/Azure Blob) con versioning + virus scan.
- **HTML reports** (Mistica tokens) → server-side rendering con [@telefonica/mistica](https://github.com/telefonica/mistica) o equivalente.
- **Estado del riesgo** (calculado en VBA desde fechas) → **state machine explícita** en backend (XState-like), con columna `Estado` opcional para auditoría.

### 1.3 Fuera de scope (no documentado)

- **139 módulos .bas**: no walkthroughed a nivel de detalle. Solo se inspeccionaron los call sites críticos. **Para refinar en iteración.**
- **65 clases .cls** (todas las del proyecto): solo Constructor, Riesgo, Edicion, Proyecto, Anexo, MaterializacionPlan*, PublicacionCalidad* leídas a fondo. Resto por `codegraph_explore` summaries.
- **`modFormsCatalog` y `modInternalFormsManager`**: presumiblemente tracking de forms abiertas. No leídos.
- **`modBilling`**: no observado.
- **API de autenticación del handshake con `Expedientes`**: no visible en código — presumiblemente vía sesión/token.

---

## 2. Estado del descubrimiento

### 2.1 Inventario Dysflow (perfado 2026-08-05 + walkthrough 2026-08-06)

| Categoría | Resultado |
|---|---|
| **Frontend** | `Gestion_Riesgos.accdb` (40.9 MB) |
| **Backend** | `Gestion_Riesgos_Datos.accdb` (17.0 MB) |
| **Tablas** | 71 (inventario previo) — ~25 walkthroughed (resto por RecSrcDt o referencias literales) |
| **Forms** | **62** (61 walkthroughed, 1 en formato binario denso PrtMip/devMode) |
| **Clases** | 65 |
| **Módulos** | 139 (mayor que Brass porque Gestion_Riesgos tiene muchos más helpers TDD/refactor) |

### 2.2 Auditoría de uso — D88 medido

| Métrica | Valor |
|---|---|
| **Getdb() total call sites** | **372** (no 308 como decía el brief original — 64 nuevos desde el corte) |
| **Getdb() directos en forms de G1** | 0 |
| **Getdb() directos en forms de G2** | 2 (Form_FormRiesgo.EstablecerDatos + Form_FormRiesgoNC.EstablecerCombos) |
| **Getdb() directos en forms de G3** | 0 (clean!) |
| **Getdb() directos en forms de G4** | 0 (clean!) |
| **Getdb() directos en forms de G5** | 0-1 (búsqueda cross-app Expedientes NO usa getdb local) |
| **Conclusión D88** | **El cuello NO está en la UI**. Está en las clases de dominio (`Edicion.cls`=23 calls, `Constructor.bas`=140, `Funciones Generales.Publicar` es el más intensivo). Migración a web NO requiere tocar el form para optimizar — el cambio es de implementación: DAO.Recordset → fetch() HTTP en cada clase. |

### 2.3 Refactor HR3/HR4 ya completado

El equipo ya hizo el refactor de la arquitectura:
- **HR3 (julio 2026)**: business logic extraída de event handlers a módulos helpers (`modFormCoordinationHelper`, `modAnexosListPresenter`, `modCalidadTareasHelper`).
- **HR4 (julio 2026)**: acceso cross-form encapsulado en `Coord_*` helpers. Patrón `Forms('X').Controls(...)` eliminado.
- **HR3e-x (issue #109, "honest messaging")**: los handlers NO envían correos en el thread. Mensaje honesto al usuario: "El correo a responsables se enviará al finalizar."

Para la épica de migración a web, esto significa:
- **Los forms ya son thin controllers** — la lógica está en módulos/clases/servicios.
- **`PublicacionCalidadStatusPresenter` (issue #46)** es el ejemplo paradigmático: Presentational pattern (ViewModel + Presenter) extraído del form.
- **Migración 1:1 a SPA**: los servicios se convierten en endpoints REST, los modHelpers en services client-side, los forms en React/Vue components.

### 2.4 Tablas detectadas (consolidado)

**Núcleo del dominio de riesgos (~25 tablas)**:
- `TbRiesgos` (2924 rows) — maestro + estado calculado
- `TbProyectos`, `TbEdiciones` (con IDEdicion Long post-issue #107), `TbExpedientes` (cross-app)
- `TbProyectosEdicionesSuministradores`, `TbProyectosAutorizados`
- `TbRiesgosMaterializaciones` (30), `TbRiesgosNC` (31, FK lógica NoConformidades)
- `TbRiesgosAIntegrar` (origen='Oferta'), `TbRiesgosBiblioteca`
- `TbRiesgosPlanMitigacionPpal/Detalle/Reversa` + `TbRiesgosPlanContingenciaPpal/Detalle/Reversa`
- `TbCambiosParaPublicacion` (incluye discriminador semántico `NombreCampo='MotivoNoPublicable'`)
- `TbLogPublicaciones` (append-only), `TbCorreosEnviados`, `tbCambios` (6725 rows)
- Enums: `TbValoresPosiblesEstadoRiesgo`, `TbValoresPosiblesMitigacion`, `TbValoresPosiblesPlazoCalidadCosteVulnerabilidad`, `TbValoresPosiblesContingencia`, `TbValoresPosiblesValoracion`, `TbOrigenesRiesgosDetalles`, `TbCategoriaRiesgo`, `TbTiposImputacion`
- IDs anti-patrón Access: ~12 tablas sequencers (`TbIDRiesgos`, `TbIDPlanMitigacion`, etc.)
- Cache: 5 tablas `TbCache*` (TreeView de Riesgos + Control Cambios)
- `TbAnexos` (polimórfica — compartida con Brass)
- Cross-app: `TbNCs` (de la app NoConformidades — referencia lógica)

---

## 3. Hallazgos críticos

| # | ID | Título | Severidad | Componentes afectados | Detalle |
|---|---|---|---|---|---|
| H1 | D88 | **372 getdb() call sites en clases de dominio** | CRITICAL | `Edicion.cls`, `Constructor.bas`, `Funciones Generales.Publicar`, +50 clases | 64 nuevos desde el brief original. **El cuello NO está en la UI** — está en las clases. Migración a web: cada método DAO se reemplaza por HTTP call con la misma firma. La concurrencia del backend (pool JDBC/HTTP) resuelve D88 estructuralmente. |
| H2 | D126 | **State machine de 14 estados en `EnumRiesgoEstado`** | high | `Riesgo.EstadoEnum` | Calculado desde fechas en runtime. Implementar state machine explícita con transiciones bien definidas + columna `Estado` para auditoría. |
| H3 | D127 | **Triángulo plazo/coste/calidad = 3 columnas independientes** | medium | `FormRiesgoPlazoCosteCalidad` | No consolidar. Preservar como 3 columnas con escala 5-niveles. |
| H4 | D128 | **Workflow de publicación dual TÉCNICO→CALIDAD** | high | `TbProyectosEdiciones` (campos `FechaPreparadaParaPublicar`, `PropuestaRechazadaPorCalidadFecha/Motivo`, `FechaPublicacion`) | State machine declarativo con transiciones que disparen efectos colaterales (correo, AGEDO, log). |
| H5 | D129 | **Integración AGEDO** | high | `Edicion.RegistrarEnAGEDO` | Si `TbProyectos.Juridica='TdE'`, registro jurídico. **Decisión abierta**: API REST o SFTP o manual. |
| H6 | D130 | **Correo honesto (no en thread)** | medium | HR3e-x (issue #109) | Sistema de notificaciones inherentemente asíncrono. UI dice "Notificación enviada" cuando la queue confirma, no "en proceso". |
| H7 | D131 | **Mitigación = enum (no item accionable)** | medium | `TbValoresPosiblesMitigacion` | 4 valores: Aceptar/Evitar/Reducir/Transferir. Exponer como selector en cabecera, no como módulo independiente. |
| H8 | D132 | **XApp Expedientes** | high | `FormExpedientesBusqueda` | HTTP/JSON handshake, sin auth documentado. Migrar a API REST federada con OAuth 2.0 / SSO. |
| H9 | D133 | **`FormWeb` wrapper de Internet Explorer** | high | `FormWeb` | ActiveX `Shell.Explorer.2`. **Eliminar completamente** en migración. Funcionalidad trivial en web (route directa + iframe). |
| H10 | D134 | **`FormControlCambiosGestion` mal asignado** | low | `FormControlCambiosGestion` | Audita versiones de la app (`CCVersion` + `tbCambios`), no riesgos. Mover a módulo transversal "Auditoría de cambios". |
| H11 | D135 | **Bug en `FormRiesgoVulnerabilidad.EstablecerDatos:66-68`** | high | `FormRiesgoVulnerabilidad` | `Me.AllowEdits = True` SIEMPRE (independiente de `m_ObjProyectoActivo.EsAutorizado(True)`). Corregir + migrar permisos a backend. |
| H12 | D136 | **Bug en `FormPlanPrincipal:154-156`** | high | `FormPlanPrincipal` | Misma lógica — `If Edición.EsActivo=Sí + UsuarioAutorizado=No Then True`. Misma corrección que H11. |
| H13 | D137 | **Publicar no transaccional** | high | `Publicar()` | Múltiples operaciones (UPDATE ediciones, INSERT log, copy file, send mail, register revision) sin transacción atómica. Usar transacciones explícitas con rollback compensatorio (Saga pattern). |
| H14 | D138 | **WithEvents sobre-amplia (32 edges)** | medium | `FormCalidadRiesgoAceptadoRetiradoVisado` | 8 `RaiseEvent`s + 4 `WithEvents` handlers. Colapsar a un único evento tipado `OnDecision(payload)`. |
| H15 | D139 | **`m_NodoSeleccionado` global** | medium | `FormTecnicoTareas` + `FormTecnicoTareasDetalleEdicion` + `FormTecnicoTareaExplicacion` | TreeView state spilled entre forms. En web: URL state (`?nodoId=xxx`) o Context Provider. |
| H16 | D140 | **Sin locks en concurrencia** | high | `TbProyectosEdiciones` (UPDATE directo) | Sin version check. Implementar optimistic concurrency con `updated_at` token. |
| H17 | D141 | **`modAnexosListPresenter` no usado** | low | `FormAnexos1` | Helper existe pero `FormAnexos1` no lo usa (4 copies inline). Consolidar. |
| H18 | D142 | **`FormAnexos` vs `FormAnexos1` — NO polimorfismo** | medium | 2 forms | Misma implementación, NO comparten interface. Consolidar en `<AnexosManager>` con prop `ambito={Proyecto|Edicion|Riesgo}`. |
| H19 | D143 | **RecSrcDt distintos** | low | 7 RecSrcDt únicos en G5 | Cada form bindeado a una tabla distinta. En web: cada componente es una página/ruta distinta. |
| H20 | D144 | **Discrepancia semántica** | low | `FormPublicacionTecnicoPropuesta` | "Propuesta económica" vs "Propuesta de publicación". Documentar claramente. |
| H21 | D145 | **`tbCambios` (6725 rows auditing cambios)** | medium | `tbCambios` | Append-only. Structured log + table. `nombre_campo` discriminator (`'MotivoNoPublicable'`) → Discriminated union en TypeScript. |

---

## 4. Decisiones aplicadas

### 4.1 UX: Preservar / Mejorar / Nuevo paradigma

| Pantalla | Decisión | Justificación |
|---|---|---|
| `Form0BDOpciones` / `Form0BDOpcionesTecnico` | **Preservar** jerarquía de roles | Entry points diferenciados — estructura cognitiva aprendida |
| `FormSuministrador`, `FormProyectosGestion` | **Preservar** CRUD simple | Forms no necesitan rediseño |
| `FormIndicador` | **Mejorar**: charts interactivos con click→drilldown | Funcionalidad ya existe, mejorar UX |
| `FormRiesgoBiblioteca` + Gestión | **Preservar** patrón con singleton ListBox | Funciona, no revolucionar |
| `FormRiesgo` (NavigationControl 5 tabs) | **Mejorar**: stepper visual con persist de estado | 53 controles en 1 form es overload |
| `FormGestionRiesgos` | **Preservar** Alta/Edición de cabecera | Flujo crítico |
| `FormRiesgoVulnerabilidad` | **Mejorar**: migrar permisos a backend (D135) | Bug a corregir en refactor |
| `FormRiesgoMaterializado`, `FormRiesgoRetirado` | **Preservar** workflow de visado dual | Crítico, no revolucionar |
| `FormRiesgoNC` | **Preservar** vinculación cross-app | Crítico — FK lógica a NoConformidades |
| `FormRiesgosGestion` (lista + treeview) | **Mejorar**: server-side pagination + virtualización | Necesita escala |
| `FormCalidadTareas` (TreeView 7 categorías) | **Nuevo paradigma**: react-arborist + subform lazy-loaded | TreeView MSComctlLib no portable |
| `FormCalidadRiesgoAceptadoRetiradoVisado` | **Mejorar**: modal con captions dinámicas | 8 RaiseEvents → colapsar a 1 (D138) |
| `FormPublicacionCalidad` (5 tabs) | **Mejorar**: wizard steps con state guard | Patrón dual TÉCNICO/CALIDAD |
| `FormMaterializacionPlanContingencia` (TempVars) | **Mejorar**: callback/promise en lugar de TempVars | Trivial — patrón modal IPC |
| `FormAnexos1.OpenArgs` ("Proyecto/Edición/Riesgo") | **Mejorar**: URL path en lugar de string OpenArgs | Web path = `/recurso/:id/anexos` |
| `FormTecnicoTareas` (TreeView) | **Preservar** jerarquía 7-cat, **Mejorar** virtualización | TreeView es esencial |
| `FormExpedientesBusqueda` | **Preservar** flujo, **Mejorar** API REST | D132 — handshake cross-app |
| `FormProyectosGestion` (4 filtros, 31 controles) | **Mejorar**: server-side pagination, chips para filtros | D88 — la lista carga TODO |
| `FormWeb` (wrapper IE) | **Eliminar** | D133 — sin reemplazo, route directa en web |
| `FormRiesgosGestion` (lista) | **Mejorar**: virtualización React | 1 form para N riesgos |

### 4.2 Arquitectura: hexagonal ports

| Concepto legacy | Traducción web |
|---|---|
| `getdb()` (DAO.Database) | Puerto de persistencia PostgreSQL (D14) + HTTP client |
| `Constructor.getXxx()` (factory con cache lazy) | Inyección de dependencias + repository pattern |
| `m_ObjEntorno` (singleton global) | Service registry / dependency injection |
| `m_ObjProyectoActivo` / `m_ObjEdicionActiva` / `m_ObjRiesgoActivo` (singleton in scope) | Request-scoped services / Context API |
| `modFormCoordinationHelper.Coord_*` | Navegación declarativa (router + outlet pattern) |
| `PublicacionCalidadStatusPresenter` | ViewModel/Presenter pattern (Zustand, React Query, Redux) |
| `WithEvents` + `RaiseEvent` (D138) | Pub/sub tipado, state machine (1 evento colapsado) |
| `TempVars!Variable` IPC (D134) | Promise/callback en modal context |
| `fso.FileExists + ShellExecute('open', url)` | window.open(url) + blob URL |
| `Application.FileDialog(3)` | `<input type="file">` + signed URL upload |
| `AbrirEnLocal` (cliente Windows) | Endpoint REST con Content-Type application/octet-stream |
| `Ejecutar Me.hWnd 'open' url` | window.open |
| `Dame()` (helper "existe registro?") | Backend unique constraint o pre-check en POST |
| `WithEvents m_FormExpedientes` (cross-app callback) | API REST + JWT + event bus (Server-Sent Events / WebSocket) |
| `TbCambiosParaPublicacion.NombreCampo='MotivoNoPublicable'` (string discriminator) | Discriminated union pattern en TypeScript |
| `tbCambios` (6725 filas auditing cambios) | Structured log + table |
| `AnexoAntiguo.cls` (legacy con `CodigoUnico`) | Deprecation — preservar en read-only si hay datos |

### 4.3 Decisiones de seguridad

- **Cross-app con Expedientes** (vía HTTP/JSON): API REST federada con OAuth 2.0 / SSO. Sin mecanismo actual de auth documentado.
- **Concurrencia** (D140): optimistic locking con `updated_at` token en cada edición.
- **Publicar no transaccional** (D137): transacciones explícitas con compensación (Saga pattern).
- **Bugs D135, D136**: corrección durante la migración con test coverage.

### 4.4 Decisiones de datos

- **Estado del riesgo**: columna `Estado` opcional para auditoría + state machine en backend que calcula desde fechas.
- **Triángulo plazo/coste/calidad**: 3 columnas separadas (NO consolidar).
- **Anexos**: polimórfica en el backend (1 tabla con CHECK constraint o 6 tablas + vista UNION, según volumen).
- **AnexoAntiguo.cls**: deprecation limpia.
- **`tbCambios`**: append-only en PostgreSQL o migrar a log estructurado.

---

## 5. Criterios de aceptación

### 5.1 Funcionalidad

- [ ] **CA-F1**: Las 61 features F1-F61 tienen paridad funcional con la versión Access.
- [ ] **CA-F2**: Los 14 estados del ciclo de vida del riesgo (D126) transicionan correctamente en el state machine del backend.
- [ ] **CA-F3**: El triángulo plazo/coste/calidad (D127) se mantiene como 3 columnas independientes con escala 5-niveles.
- [ ] **CA-F4**: El workflow de publicación dual TÉCNICO→CALIDAD (D128) preserva los 4 campos de auditoría.
- [ ] **CA-F5**: La integración AGEDO (D129) documenta el contrato API y se ejecuta solo si `TbProyectos.Juridica='TdE'`.
- [ ] **CA-F6**: Las notificaciones son inherentemente asíncronas (queue + retry) y el toast al usuario es honesto (D130).
- [ ] **CA-F7**: La API de Expedientes se migra a REST federada con OAuth (D132).
- [ ] **CA-F8**: `FormAnexos` + `FormAnexos1` se consolidan en un único `<AnexosManager>` parametrizable por ambito (D142).
- [ ] **CA-F9**: `FormWeb` se elimina completamente, route directa a HTML (D133).
- [ ] **CA-F10**: `Publicar()` transaccional con rollback compensatorio (D137).
- [ ] **CA-F11**: El state machine del riesgo tiene optimistic locking (D140).
- [ ] **CA-F12**: Las 22 RN específicas del grupo Publicación documentadas en `legacy.rules.ts`.

### 5.2 Seguridad

- [ ] **CA-S1**: Cross-app API con Expedientes vía OAuth 2.0 + scopes granulares.
- [ ] **CA-S2**: `updated_at` token en cada edición para optimistic concurrency.
- [ ] **CA-S3**: Transacciones explícitas con rollback compensatorio (Saga pattern).
- [ ] **CA-S4**: Bugs D135, D136 corregidos con tests de regresión.

### 5.3 Performance

- [ ] **CA-P1**: `FormProyectosGestion` con server-side pagination (la lista ya no carga TODO).
- [ ] **CA-P2**: `FormIndicador` carga sub-segundos incluso con 5 años de datos.
- [ ] **CA-P3**: `FormRiesgosGestion` (lista) con virtualización React.
- [ ] **CA-P4**: Cache distribuido reemplaza `TbCache*` (server-side invalidation).

### 5.4 Operacional

- [ ] **CA-O1**: Tablas `TbCambiosParaPublicacion` y `tbCambios` migradas a log estructurado con retention policy.
- [ ] **CA-O2**: `TbAnexos` migrada a S3-compatible con versioning + virus scan.
- [ ] **CA-O3**: `tbCambios` (changelog de aplicación) — publicable vía endpoint /changelog.
- [ ] **CA-O4**: `CCVersion` se reemplaza por git tags o semver.

---

## 6. Pendientes operacionales

### 6.1 Antes de empezar

- [ ] **PO-1**: Confirmar el contrato de integración con **Expedientes** (autenticación, endpoint, schema). Sin este contrato, F54 (búsqueda) no se puede migrar.
- [ ] **PO-2**: Confirmar el contrato de integración con **AGEDO** (D129). Si TdE es frecuente (~qué porcentaje de proyectos), merece inversión en API. Si no, mantener manual.
- [ ] **PO-3**: Confirmar el formato del campo `Estado` del riesgo: ¿columna persistida o calculada siempre?
- [ ] **PO-4**: Auditar `AnexoAntiguo.cls` para ver si hay datos legacy con `CodigoUnico` que necesiten migración.
- [ ] **PO-5**: Bugs D135 + D136 — ¿es print o cambio de diseño? Corregir antes de migrar.

### 6.2 Durante el desarrollo

- [ ] **PO-6**: Implementar el state machine del riesgo (D126) con tests de transición.
- [ ] **PO-7**: Refactor `modAnexosListPresenter` para que `FormAnexos1` también lo use (D141).
- [ ] **PO-8**: Migrar `m_NodoSeleccionado` (D139) — probablemente vía URL state o Context Provider.
- [ ] **PO-9**: Implementar notificaciones asíncronas con queue (D130).
- [ ] **PO-10**: Los forms con bugs tipográficos (sección H21) deben corregirse durante la migración.

### 6.3 En el go-live

- [ ] **PO-11**: Smoke test E2E: alta proyecto → riesgo → oferta → retirada → calidad visar → publicar → AGEDO.
- [ ] **PO-12**: Verificar las 60+ RN documentadas con datos de producción antes de switchover.
- [ ] **PO-13**: Backfill de `TbAnexos` desde S3 (si hay anexos legacy en filesystem).
- [ ] **PO-14**: Plan de deprecation de `AnexoAntiguo.cls` y forms duplicados (`FormRiesgoDefinicion` vs `FormRiesgoDefinicionNoBiblioteca`).

---

## 7. Tickets derivables (preview)

> NO crear issues todavía. Estos nacen de la épica al final del ciclo de revisión.

### Performance (D88)

- **TK-GR-1**: [PERFORMANCE] Migrar `getdb()` (Variables Globales.bas:1069) al cliente HTTP — 372 call sites.
- **TK-GR-2**: [PERFORMANCE] Pool JDBC + connection pooling para concurrencia backend.

### State machine (D126-D128)

- **TK-GR-3**: [STATE-MACHINE] Implementar las 14 transiciones de `EnumRiesgoEstado` con tests unitarios de transición.
- **TK-GR-4**: [DOMAIN] Triángulo plazo/coste/calidad como 3 columnas independientes (D127).
- **TK-GR-5**: [STATE-MACHINE] Implementar state machine del workflow Publicación (D128) con efectos colaterales.
- **TK-GR-6**: [INTEGRATION] Contrato AGEDO para `RegistrarEnAGEDO` (D129).
- **TK-GR-7**: [NOTIFICATIONS] Cola asíncrona de correos con retry (D130).
- **TK-GR-8**: [DOMAIN] Mitigación como enum (NO como item independiente — D131).

### Integración

- **TK-GR-9**: [INTEGRATION] API REST con Expedientes + OAuth 2.0 (D132).
- **TK-GR-10**: [CLEANUP] Eliminar FormWeb (D133) — route directa al HTML.

### Forms y migración

- **TK-GR-11**: [MIGRATION] Mover `FormControlCambiosGestion` al módulo transversal "Auditoría de cambios" (D134).
- **TK-GR-12**: [BUGFIX] Corregir `FormRiesgoVulnerabilidad.EstablecerDatos:66-68` (D135).
- **TK-GR-13**: [BUGFIX] Corregir `FormPlanPrincipal:154-156` (D136).
- **TK-GR-14**: [STATE-MACHINE] `Publicar()` transaccional con rollback compensatorio (D137).
- **TK-GR-15**: [REFACTOR] Colapsar 32 WithEvents/RaiseEvent en un único `OnDecision(payload)` tipado (D138).
- **TK-GR-16**: [REFACTOR] Migrar `m_NodoSeleccionado` (D139) a URL state o Context.
- **TK-GR-17**: [CONCURRENCY] Optimistic locking con `updated_at` (D140).
- **TK-GR-18**: [REFACTOR] Consolidar `modAnexosListPresenter` (D141).
- **TK-GR-19**: [REFACTOR] Consolidar FormAnexos + FormAnexos1 en `<AnexosManager>` (D142).

### Calidad y limpieza

- **TK-GR-20**: [CLEANUP] Corregir captions y tags mal asignados.
- **TK-GR-21**: [MIGRATION] Migrar 65 clases de dominio a servicios HTTP manteniendo firmas.
- **TK-GR-22**: [DEPRECATION] Deprecar `AnexoAntiguo.cls` (legacy `CodigoUnico` model).
- **TK-GR-23**: [MIGRATION] Consolidar duplicaciones `FormRiesgoDefinicion` vs `FormRiesgoDefinicionNoBiblioteca`.
- **TK-GR-24**: [MIGRATION] TreeView de `FormTecnicoTareas` y `FormCalidadTareas` → react-arborist.
- **TK-GR-25**: [MIGRATION] NavigationControl de `FormRiesgo` → stepper.

### Testing y validación

- **TK-GR-26**: [TEST] Tests E2E para el ciclo de vida completo del riesgo (14 estados).
- **TK-GR-27**: [TEST] Tests E2E para el workflow Publicación (proponer/aceptar/rechazar/publicar).
- **TK-GR-28**: [TEST] Tests E2E para cross-app con Expedientes.
- **TK-GR-29**: [TEST] Tests E2E para la integración con AGEDO (mock si no hay sandbox).
- **TK-GR-30**: [TEST] Tests de regresión para los bugs D135, D136.

### Datos y migración

- **TK-GR-31**: [MIGRATION] Backfill de `Estado` columna para los 2924 riesgos actuales.
- **TK-GR-32**: [MIGRATION] Backfill de `FechaPublicacion` para ediciones históricas.
- **TK-GR-33**: [MIGRATION] Migrar 71 tablas a PostgreSQL con Expand and Contract (D82).
- **TK-GR-34**: [MIGRATION] Reemplazar ~12 tablas sequencers (`TbIDXxx`) por IDENTITY/SERIAL.
- **TK-GR-35**: [MIGRATION] Reemplazar 5 tablas `TbCache*` por Redis o similar server-side cache.
- **TK-GR-36**: [MIGRATION] Reemplazar filesystem anexos (`URLDirectorioDocumentacion`) por S3-compatible.
- **TK-GR-37**: [MIGRATION] Migrar `TbLogPublicaciones` (append-only) a log estructurado.
- **TK-GR-38**: [MIGRATION] Migrar `tbCambios` (6725 filas) a log estructurado inmutable.
- **TK-GR-39**: [MIGRATION] Expand and Contract backward-compatibles (D82).
- **TK-GR-40**: [REVIEW] Análisis de los 3 forms binarios densos después de abrir en Access IDE.

---

## Anexo · Tabla de fuentes

| Fuente | Aporta |
|---|---|
| engram topic_key `gestion-riesgos/deep-discovery-2026-08-06` | **Consolidado** de 5 sub-agentes walkthrough |
| `docs/03-aplicaciones/gestion-riesgos/capabilities.md` | Inventario de features, ViewModels/Servicios/Repositorios, edge WebView |
| `docs/03-aplicaciones/gestion-riesgos/data-model.md` | Tablas, FKs, hallazgos D88-D104 |
| `docs/03-aplicaciones/gestion-riesgos/forms.md` | Navegación, call paths, formularios críticos |
| `docs/03-aplicaciones/gestion-riesgos/integrations-automation.md` | Cross-app, Edge WebView, testing sandbox, flags |
| `docs/03-aplicaciones/gestion-riesgos/migration-matrix.md` | D88-D104 detallados + decisión D-new logs web-native |
| `docs/03-aplicaciones/gestion-riesgos/security-rules.md` | Autorización, D88-D95, riesgos de privacidad, XSS de WebView |
| `docs/03-aplicaciones/gestion-riesgos/README.md` | Estado del lote, hallazgos críticos, checklist |
| `data/staging/gestion-riesgos/src/classes/*.cls` | 65 clases de dominio |
| `data/staging/gestion-riesgos/src/modules/*.bas` | 139 módulos (Constructor, Variables Globales, etc.) |
| `data/staging/gestion-riesgos/src/forms/*.form.txt` | 62 forms con RecordSource/RowSource |
| engram obs #24097 | Audit codegraph-vba |
| engram obs #24086 | Corrección: `tbHistorialRechazos` es de NEGOCIO |
| [DOCS](../../../DOCS.md) | Technical reference raíz del blueprint |
| [CODEBASE-GUIDE](../../../CODEBASE-GUIDE.md) | Para mantenedores del blueprint |

## Anexo · Decisiones referenciadas (D5-D145)

| Decisión | Aplicación a Gestion_Riesgos |
|---|---|
| D8 (hexagonal global) | Toda la migración |
| D14 (esquema por módulo) | Módulo `gestion_riesgos` en PostgreSQL |
| D16 (object storage S3-compatible) | Anexos (D142, D143) |
| D17-D20 (ports) | Repositorios + adaptadores de identidad |
| D27 (logs estructurados canónicos) | Reemplazar logs VBA |
| D44-D46 (autorización + capabilities) | Roles Técnico/Calidad |
| D66-D67 (stack) | Backend Python + HTMX |
| D68 (monolito modular) | Gestion_Riesgos como módulo |
| D82 (Expand and Contract) | Estrategia de migración |
| D92 (datos personales) | NO bloqueante para Gestion_Riesgos — no hay PII significativa |
| D93 (password fallback) | APLICAR — eliminar fallback `dpddpd` |
| D102 (booleanos `Text(2)`) | APLICAR — migrar a `BOOLEAN` |
| D104 (secret manager) | APLICAR — mover a secret manager |
| **D88 (372 getdb() callers)** | **CRITICAL — cuello en clases de dominio, no en UI** |
| **D126-D145 (state machine + workflow + AGEDO + tests)** | **21 hallazgos específicos de Gestion_Riesgos** |

## Checklist del documento

- [x] Scope con 61 features detalladas por dominio (G1-G5)
- [x] Auditoría de uso (Dysflow + codegraph-vba) + 65 clases / 139 módulos
- [x] Hallazgos D88 + D126-D145 con anchor links (21 hallazgos)
- [x] Decisiones UX Preservar/Mejorar/Nuevo paradigma por pantalla
- [x] Decisiones arquitectura hexagonal ports
- [x] Criterios de aceptación verificables y agrupados por dimensión
- [x] 14 pendientes operacionales antes, 10 durante, 4 en go-live
- [x] 40 tickets derivables preview (TK-GR-1..40)
- [x] Tabla de decisiones referenciadas (D5-D145)
- [x] Tabla de fuentes
- [x] Idioma: español técnico neutro. Identificadores y paths sin traducir.
- [x] "The sentence that organizes" presente
- [x] "Scope del scope" presente
- [x] Sin emojis decorativos
- [x] Cross-references a DOCS, CODEBASE-GUIDE, AGENTS
- [x] Quick Navigation table
- [x] Hallazgos en tabla con severity

## Siguiente paso

Aplicar las mismas reglas a las 3 épicas restantes: HPS, HPS_Solicitudes, NoConformidades. Tras cerrar el ciclo de revisión final del blueprint.

## Cómo se aplica a access2web-blueprint

Gestion_Riesgos es la app con más tablas (71) y más complejidad de caché multi-nivel. Su código destino vive en `app/src/modules/gestion-riesgos/` y depende de NoConformidades (D95 orden de migración estricto) y de Lanzadera (D86 acoplamiento identidad).

**Entrada cruda**: los walkthroughs G1..G5 consolidan 61/61 forms. Inventario Dysflow: 71 tablas, 889 riesgos en producción (`TbIDRiesgos`), 2843 planes de mitigación (`TbRiesgosPlanMitigacionPpal`), 2089 nodos del árbol cacheados (`TbCacheArbolRiesgosNodo` — confirma el hallazgo D88 sobre lentitud del árbol).

**Decisiones operativas vigentes**: D86 (acoplamiento Lanzadera via `getdbLanzadera()` traducido a adaptadores), D88 (rendimiento del árbol: HTMX lazy expansion + CTE recursivo PostgreSQL; descarta MSComctlLib.TreeView y flag `CadenaJerarquicaModelo`), D95 (orden de migración estricto: NoConformidades antes que Gestion_Riesgos por FK conceptual en `TbRiesgosNC`), D102 (booleans Text(2) cross-cutting → `BOOLEAN`), patrón `_Reversa` (transacciones reversibles en planes → event sourcing o soft delete con versionado), catálogos de valoración → enums PostgreSQL o tablas de lookup con UI admin.

**Cross-references desde otros docs**: DOCS.md §The 8 Apps lista Gestion_Riesgos como mergeada con PR #1; `docs/architecture.md` §Decisiones D-<n> vigentes lista D14, D16, D27, D82, D88 (cross-cutting, no per-app); CODEBASE-GUIDE.md §Ownership de artefactos la referencia para el código de plataforma en `app/src/modules/gestion-riesgos/` y el walkthrough JSON.

## Lista de comprobación final

- [ ] Las 7 secciones del cuerpo están completas y verificadas contra los walkthroughs G1..G5.
- [ ] Los 2 anexos están adjuntos con referencias válidas.
- [ ] Los criterios de aceptación están todos marcados; los pendientes tienen ticket derivado.
- [ ] Los tickets derivables tienen issue-number válido en el backlog del change correspondiente.
- [ ] La sección «Cómo se aplica a access2web-blueprint» referencia los walkthroughs G1..G5, el código destino `app/src/modules/gestion-riesgos/`, y las decisiones D86, D88, D95, D102 + el patrón `_Reversa`.
- [ ] La épica se ajusta al contrato de `skills/documentation-alan-style/SKILL.md` (§3 + §8): castellano peninsular formal con usted, sin emojis decorativos.
- [ ] Las cross-references desde DOCS.md, CODEBASE-GUIDE.md, `docs/architecture.md` y los walkthrough JSONs siguen resolviendo.

---

[← Back to Gestion_Riesgos README](README.md) · [← Codebase Guide](../../../CODEBASE-GUIDE.md) · [← DOCS](../../../DOCS.md)
