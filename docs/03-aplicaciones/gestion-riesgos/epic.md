# Épica — Gestion_Riesgos (migración a web)

> **Estado:** DRAFT — pendiente revisión al final del ciclo (junto con Brass).
> **Versión:** v0.1 (2026-08-06).
> **App legacy:** Gestion_Riesgos (`C:\00repos\codigo\00_GESTION_RIESGOS`) · frontend `Gestion_Riesgos.accdb` (40 MB) + backend `Gestion_Riesgos_Datos.accdb` (17 MB).
> **Lote de discovery:** Lote 6 — junto con Brass, HPS, HPS_Solicitudes, Condor, Expedientes ya documentados.
> **Hallazgo dominante:** **D88** — `getdb()` con 372 call sites en toda la app (no 308 como decía el brief original — 64 nuevos desde el corte).
> **Cross-refs:** engram topic_key `gestion-riesgos/deep-discovery-2026-08-06` (consolidado de 5 sub-agentes paralelos).

## Metadatos

| Campo | Valor |
|---|---|
| **Aplicación legacy** | `00_GESTION_RIESGOS` · frontend `Gestion_Riesgos.accdb` + backend `Gestion_Riesgos_Datos.accdb` |
| **Tipo de migración** | Legacy Access/VBA → web hexagonal (FastAPI + HTMX) |
| **Scope size** | **XL** (62 forms, 65 clases, 139 módulos, 71 tablas, 372 getdb() callers) |
| **Dependencias cross-app** | **Expedientes** (vía HTTP/JSON — único punto de acoplamiento crítico); **NoConformidades** (vía `TbRiesgosNC.IDNC` FK lógica); **Lanzadera** (identidad/permisos presumiblemente, vía `getdbLanzadera()` — patrón compartido con Gestion_Riesgos) |
| **Riesgo dominante** | **D88** — `getdb()` con 372 callers. Los forms tienen 0-2 calls directos; el grueso está en las CLASES de dominio. **No es un problema de UI** — la optimización es de capa de datos, no de forms. |
| **Stack target** | Backend Python 3.12+ / FastAPI 0.119+ / SQLAlchemy 2.0.x / Alembic 1.13+ / asyncpg 0.30+ (D66) · Frontend HTMX 2.0.4 + Jinja2 3.1+ + Alpine.js 3.15+ (D67) |
| **Estrategia de migración de BD** | Expand and Contract backward-compatible (D82) · PostgreSQL compartido con esquema por módulo (D14) |
| **Forma destino** | Hexagonal global (D8) · módulo dentro del monolito modular (D68) · puerto de persistencia PostgreSQL + object storage S3-compatible (D16) + secret manager (D9-D10) |
| **Auditoría de uso previa** | ✅ Codegraph-vba + Dysflow walkthrough (61/62 forms, **D88 medido: 372 getdb() callers en el proyecto entero**) |

## Forma del documento

1. Scope (en / fuera)
2. Estado del descubrimiento
3. Hallazgos críticos (D88 + D126-D145)
4. Decisiones aplicadas
5. Criterios de aceptación
6. Pendientes operacionales
7. Tickets derivables (preview)

---

## 1. Scope

### 1.1 En scope (~61 features agrupadas en 5 dominios)

#### G1 — Configuración + Maestros + Menús (~10 features)

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

#### G2 — Riesgos: definición, gestión, aceptación (~16 features)

| # | Feature | Respaldo |
|---|---|---|
| F9 | **Alta/Edición directa de riesgo** (form shell con NavigationControl de 5 tabs) | `FormRiesgo` |
| F10 | **Gestión de Riesgos** (alta/edición de cabecera del proyecto) | `FormGestionRiesgos` |
| F11 | **Datos generales del riesgo** (subform de cabecera) | `FormGestionRiesgosDatosGenerales` |
| F12 | **Riesgos autorizados** (read-only picker con correos) | `FormGestionRiesgosAutorizados` |
| F13 | **Materialización del riesgo** (registrar fecha + plan contingencia elegible) | `FormRiesgoMaterializado` |
| F14 | **Vulnerabilidad del riesgo** (singleton ListBox con 5 niveles) | `FormRiesgoVulnerabilidad` |
| F15 | **Mitigación** (4 valores: Aceptar / Evitar / Reducir / Transferir) | `FormRiesgoMitigacion` |
| F16 | **Triángulo plazo/coste/calidad** (3 columnas independientes con escala 5-niveles) | `FormRiesgoPlazoCosteCalidad` |
| F17 | **Retirada** (con justificación + workflow de visado dual Calidad) | `FormRiesgoRetirado` |
| F18 | **Vinculación cross-app con NoConformidades** (vía `TbRiesgosNC.IDNC`) | `FormRiesgoNC` |
| F19 | **Riesgo Externo (Oferta)** (definición + traslado vs no traslado) | `FormRiesgoExternoDetalle` |
| F20 | **Workflow de modificación** del riesgo aprobado (aceptación o retirada solicitada) | `FormRiesgoPosiblesModificacionesAceptarRetirar`, `FormRiesgoPosiblesModificacionesQuitarAceptar` |
| F21 | **Estado del riesgo** (state machine de 14 estados con campo calculado) | `Riesgo.EstadoEnum` |
| F22 | **Gestión principal de riesgos** (lista + treeview + 6 acciones) | `FormRiesgosGestion` |
| F23 | **Establecer prioridades** (grid editable smallint) | `FormRiesgosEstablecerPrioridades` |
| F24 | **Edición masiva de riesgos** (modo batch) | `FormRiesgosGestionEdicion` |
| F25 | **Gestión de cambios de la aplicación** (auditoría CHANGELOG de versiones — D134) | `FormControlCambiosGestion` |

#### G3 — Calidad + Planes (~13 features)

| # | Feature | Respaldo |
|---|---|---|
| F26 | **Dashboard Calidad** (TreeView 7 categorías + subform detalle intercambiable) | `FormCalidadTareas` |
| F27 | **Detalle edición pendiente publicación** (read-only + 6 botones de acción) | `FormCalidadTareasDetalleEdicion` |
| F28 | **Visado de aceptación/retiro** (modal con captions dinámicas) | `FormCalidadRiesgoAceptadoRetiradoVisado` |
| F29 | **Gestión post-materialización** (vincular NC / no NC / revocar decisión) | `FormCalidadRiesgoMaterializaciones` |
| F30 | **Explicación de categoría** (mini-form de 3 controles) | `FormCalidadTareaExplicacion` |
| F31 | **Detalle riesgo aceptado/retirado pendiente visar** (read-only) | `FormCalidadTareaRiesgosAceptadosRetirados` |
| F32 | **Detalle riesgo materializado a decidir** (read-only) | `FormCalidadTareaRiesgosMaterializadosPorDecidir` |
| F33 | **Retipificación** (reasignar riesgo a código de biblioteca) | `FormCalidadTareaRiesgosRetipificacion` |
| F34 | **Capturar plan contingencia** (modal con TempVars IPC) | `FormMaterializacionPlanContingencia` |
| F35 | **Gestión riesgos de Oferta** (solo Calidad, primera edición) | `FormGestionRiesgosRiesgosOferta` |
| F36 | **Sincronización de Suministradores** (push manual desde jerarquía) | `FormGestionRiesgosSuministradores` |
| F37 | **Capturar estrategia de mitigación** (singleton ListBox + sub-form justificación) | `FormRiesgoMitigacion` |
| F38 | **Capturar triángulo plazo/coste/calidad** (3 ListBox sincronizados) | `FormRiesgoPlazoCosteCalidad` |

#### G4 — Publicación (~11 features)

| # | Feature | Respaldo |
|---|---|---|
| F39 | **Flujo Publicación Calidad** (NavigationControl 5 tabs) | `FormPublicacionCalidad` |
| F40 | **Estado de publicación** (Presentation Layer con view state derivado) | `FormPublicacionCalidadPublicar` |
| F41 | **Ejecutar publicación** (modal con todos los defaults + validaciones) | `FormPublicacionCalidadPublicarEjecutar` |
| F42 | **Capturar motivo de rechazo** (modal con Event Motivado) | `FormPublicacionCalidadRechazoObservaciones` |
| F43 | **Notas de Calidad para publicación** (TextBox libre) | `FormPublicacionCalidadNotas` |
| F44 | **Datos generales del documento a publicar** (read-only) | `FormPublicacionDatosGenerales` |
| F45 | **Evidencias de Suministradores** (lista + gestión de anexos) | `FormPublicacionSuministradores` |
| F46 | **Sub-menú Técnico propuesta de publicación** (NavigationControl 3 tabs) | `FormPublicacionTecnico`, `FormPublicacionTecnicoPropuesta` |
| F47 | **Evidencia de UTE** (anexo único por edición) | `FormPublicacionUTE` |
| F48 | **Alta de Plan individual** (mitigación o contingencia) | `FormPlanPrincipal` |
| F49 | **Integración con AGEDO** (registro jurídico TdE) | via `Edicion.RegistrarEnAGEDO` |
| F50 | **Generación de informes de publicación** (HTML + PDF con tokens Mistica) | `Funciones Generales.GenerarInformePublicacion` |
| F51 | **Servicio de refresco post-publicación** (qué riesgo refrescar) | `PublicacionCalidadExecutionService.ResolveRefreshRiskID` |
| F52 | **Auditoría de publicaciones** (`TbLogPublicaciones` append-only) | implementado en `PublicacionLog.cls` |

#### G5 — Detalles + Técnico + Anexos (~11 features)

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
| F61 | **Picker de usuarios autorizados** (read-only con correos) | `FormGestionRiesgosAutorizados` |

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

### H1 (D88 — PERFORMANCE) — 372 getdb() callers

**Síntoma**: `Variables Globales.bas:1069` define `getdb()` y tiene 372 call sites. El cuello NO está en los forms (que delegan en clases) sino en las clases de dominio (`Edicion`, `Anexo`, `Proyecto`, `Riesgo`, etc.).

**Acciones** (TK-GR-1, TK-GR-2):
- **NO** optimizar el form-level getdb — ya está en 0-2 calls.
- Migración a web natural: el `getdb()` se reemplaza por `apiClient.get(...)` HTTP en cada método de clase. Signatura del método se mantiene; implementación cambia.
- Beneficio: la concurrencia del backend (pool JDBC/HTTP) resuelve D88 estructuralmente.

### H2 (D126 — STATE MACHINE) — Ciclo de vida del riesgo

**Síntoma**: `EnumRiesgoEstado` tiene **14 estados** no triviales. El estado actual se **calcula desde fechas** en `Riesgo.EstadoEnum` (no persistido como columna).

```
Detectado(1) → Incompleto(14) → Activo(3)
   ├─ AceptadoSinJustificar(5) → AceptadoSinVisar(6) → Aceptado(8) [terminal Aceptación]
   │   └─ AceptadoRechazado(7)
   └─ RetiradoSinJustificar(9) → RetiradoSinVisar(10) → Retirado(12) [terminal Retiro]
       └─ RetiradoRechazado(11)
Materializado(4) → TbRiesgosMaterializaciones (cross-link to NC)
Cerrado(13) [terminal final]
Planificado(2)
```

**Acción** (TK-GR-3): implementar state machine explícita en el backend con transiciones bien definidas. Columna `Estado` opcional para auditoría (persistir el último estado calculado).

### H3 (D127 — TRIÁNGULO Plazo/Coste/Calidad) — 3 dimensiones independientes

**Síntoma**: son 3 columnas ortogonales independientes con escala 5-niveles (`Plazo`, `Coste`, `Calidad`). NO son sub-formularios relacionados ni trade-offs. Cada dimensión tiene su propia escala semántica (hitos/presupuesto/requisitos).

**Acción** (TK-GR-4): preservarlas como 3 columnas separadas en PostgreSQL. NO consolidarlas en una sola.

### H4 (D128 — Publicación) — State machine dual TÉCNICO→CALIDAD

**Síntoma**: campos persistidos en `TbProyectosEdiciones`:
- `FechaPreparadaParaPublicar` (técnico propuso)
- `PropuestaRechazadaPorCalidadFecha/Motivo` (calidad rechazó)
- `UsuarioProponePublicar` / `UsuarioCalidadRechazaPropuesta` (auditoría dual)
- `FechaPublicacion` (calidad publicó)

**Acción** (TK-GR-5): el state machine debe ser declarativo con transiciones que disparen efectos colaterales (correo, AGEDO, log).

### H5 (D129 — AGEDO) — Integración jurídica crítica

**Síntoma**: si `TbProyectos.Juridica='TdE'`, la publicación se registra en AGEDO (sistema jurídico de Telefónica) vía `RegistrarEnAGEDO` → devuelve `IDDocumentoAGEDO` + `URLAdjunto`. Si no → copia HTML a `URLDirectorioDocumentacion`.

**Acción** (TK-GR-6): definir contrato de AGEDO en la migración (¿API REST? ¿SFTP? ¿exportación manual?). Crítico para cumplimiento normativo.

### H6 (D130 — Correo honesto) — No se envía en el thread

**Síntoma**: refactor HR3e (issue #109) — antes decíamos "Correo enviado" en el thread; ahora "El correo se enviará al finalizar". El envío real es asíncrono.

**Acción** (TK-GR-7): el sistema de notificaciones en web debe ser inherentemente asíncrono (queue + retry). El toast al usuario debe ser honesto: "Notificación enviada" cuando la queue confirma, no "Notificación en proceso".

### H7 (D131 — Mitigación ≠ Item) — Es clasificación, no es acción

**Síntoma**: `Mitigacion` es un campo enum (`Aceptar`/`Evitar`/`Reducir`/`Transferir`) del riesgo, NO un item accionable. "Aceptar" requiere justificación + visado dual.

**Acción** (TK-GR-8): en web, exponer como selector en la cabecera del riesgo, no como módulo independiente.

### H8 (D132 — XApp Expedientes) — HTTP/JSON handshake

**Síntoma**: `FormExpedientesBusqueda` envía HTTP + parsea JSON. Único punto de acoplamiento cross-app con la app Expedientes. Sin mecanismo de autenticación documentado.

**Acción** (TK-GR-9): migrar a API REST federada con OAuth/SSO. Contrato API bien definido (request/response schemas).

### H9 (D133 — FormWeb) — Wrapper de IE

**Síntoma**: `FormWeb` embebe un control ActiveX `Shell.Explorer.2` (Microsoft Web Browser) que navega con `Me.Navegador.Navigate m_URLHTMLActivo`.

**Acción** (TK-GR-10): **eliminar el form completamente** en la migración. La funcionalidad es trivial en web (route directa + iframe si se quiere preservar embed).

### H10 (D134 — FormControlCambiosGestion mal asignado) — No audita riesgos

**Síntoma**: este form está en G2 solo por proximidad física. En realidad audita **versiones de la aplicación** (CCVersion + `tbCambios`), no riesgos.

**Acción** (TK-GR-11): mover este form al módulo transversal "Auditoría de cambios". Para la épica de Gestion_Riesgos, presentarlo como dependencia + ticket de "mover a módulo transversal".

### H11 (D135 — Bug conocido) — `FormRiesgoVulnerabilidad.EstablecerDatos:66-68`

**Síntoma**: lógica de permisos:
```vba
If m_ObjProyectoActivo.EsAutorizado(True) Then
    Me.AllowEdits = True
Else
    Me.AllowEdits = True   ' ← bug: siempre True
End If
```

**Acción** (TK-GR-12): corregir bug. Migrar la lógica de permisos al backend (claims).

### H12 (D136 — Bug potencial) — `FormPlanPrincipal:154-156`

**Síntoma**: misma lógica — `If Edición.EsActivo=Sí + UsuarioAutorizado=No Then True`.

**Acción** (TK-GR-13): aplicar misma corrección que H11.

### H13 (D137 — Publicar no transaccional) — Riesgo de inconsistencia

**Síntoma**: `Publicar()` ejecuta múltiples operaciones (UPDATE ediciones, INSERT log, copy file, send mail, register revision) sin transacción atómica. Si falla a mitad, queda estado inconsistente.

**Acción** (TK-GR-14): en backend usar transacciones explícitas con rollback compensatorio (Saga pattern o compensating transactions).

### H14 (D138 — WithEvents sobre-amplia) — 32 edges de coordinación

**Síntoma**: `FormCalidadRiesgoAceptadoRetiradoVisado` define 8 `RaiseEvent`s (AceptacionAprobada, AceptacionRechazada, AceptacionAprobadaQuitado, AceptacionRechazadaQuitada, + 4 equivalentes Retiro). Los 4 forms padres (`FormCalidadTareaRiesgosAceptadosRetirados`) tienen los 4 `WithEvents` handlers.

**Acción** (TK-GR-15): colapsar a un único evento tipado `OnDecision(payload)` en lugar de 8 strings. Web: state machine + reducer pattern.

### H15 (D139 — `m_NodoSeleccionado` global) — TreeView state spilled

**Síntoma**: variable global module-level en `FormTecnicoTareas` que es consumida por `FormTecnicoTareasDetalleEdicion` y `FormTecnicoTareaExplicacion`.

**Acción** (TK-GR-16): en web, URL state (`?nodoId=xxx`) o Context Provider.

### H16 (D140 — Concurrencia) — Sin locks

**Síntoma**: `TbProyectosEdiciones` se modifica con UPDATE directo (`m_ObjEdicionActiva.SomeProperty = ... .Editar .Update`). Sin version check.

**Acción** (TK-GR-17): implementar optimistic concurrency (`updated_at` token) en el backend.

### H17 (D141 — `modAnexosListPresenter` no usado) — Duplicación

**Síntoma**: existe el helper pero `FormAnexos1` no lo usa (4 copies inline de `EstablecerLista`).

**Acción** (TK-GR-18): consolidar.

### H18 (D142 — FormAnexos vs FormAnexos1) — NO polimorfismo

**Síntoma**: son 2 implementaciones del mismo flujo. NO comparten interface ni factory.

**Acción** (TK-GR-19): consolidar en 1 componente React `<AnexosManager>` con prop `ambito={Proyecto|Edicion|Riesgo}`.

### H19 (D143 — RecSrcDt distintos) — Cada form su tabla

**Síntoma**: 7 RecSrcDt únicos en G5. Cada form bindeado a una tabla distinta. NO comparten recordset.

**Acción**: en web, cada componente es una página/ruta distinta con su propio endpoint. Reflejo natural del patrón legacy.

### H20 (D144 — Discrepancia semántica) — "Propuesta económica" no está aquí

**Síntoma**: `FormPublicacionTecnicoPropuesta` = propuesta de PUBLICACIÓN, no propuesta económica del técnico.

**Acción**: documentar claramente. Si la propuesta económica está en otro flujo (Brass o licitaciones), referenciar.

### H21 (Cross-cutting) — Bug tipográfico observado (replicar del audit)

- `Form0BDOpciones.ComandoAyuda.Caption` = "Salir"
- `ComandoEntrarComo.Caption` = "Salir"
- `lblEstado.ControlTipText` = "Origen del Riesgo" (mismatch con el contenido real)
- `lblEstado.Tag` = "Formulario de control de los pedidos" (mismatch)
- `CombReparadoPor.ControlSource` = "Material.PN" (en otros forms)
- "CerrarAplicacion" comentado en `CmdSalir_Click` (código muerto)
- Hyperlinks `HyperlinkAddress='#'` en labels (lblTareasPendientes, etc.)

**Acción** (TK-GR-20): limpieza de captions y lblEstado durante la migración.

---

## 4. Decisiones aplicadas

### 4.1 UX: Preservar / Mejorar / Nuevo paradigma

| Pantalla | Decisión | Justificación |
|---|---|---|
| FormWeb (wrapper IE) | **Eliminar** | D133 — sin reemplazo, route directa en web. |
| FormGestionRiesgosAutorizados | **Preservar** read-only | Trivial: array.map() en backend. |
| FormTecnicoTareaExplicacion | **Mejorar**: pasar de acoplamiento global a URL state | D139 — fácil. |
| FormAnexos + FormAnexos1 | **Nuevo paradigma**: consolidar en 1 `<AnexosManager>` | D142, D141 — duplicación clara. |
| FormExpedientesBusqueda | **Preservar** flujo; **Mejorar**: API REST | D132 — handshake cross-app. |
| FormProyectosGestion (4 filtros, 31 controles) | **Mejorar**: server-side pagination, chips para filtros | D88 — la lista carga TODO. |
| FormIndicador (6 tiles) | **Mejorar**: charts interactivos con click→drilldown | Funcionalidad ya existe, mejorar UX. |
| FormCalidadTareas (TreeView 7 categorías) | **Nuevo paradigma**: React-arborist + subform lazy-loaded | TreeView MSComctlLib no portable. |
| FormRiesgo (NavigationControl 5 tabs) | **Mejorar**: stepper visual con persist de estado | 53 controles en 1 form es overload. |
| FormPublicacionCalidad (5 tabs) | **Mejorar**: presentar como wizard steps con state guard | Patrón dual TÉCNICO/CALIDAD. |
| FormMaterializacionPlanContingencia (TempVars) | **Mejorar**: callback/promise en lugar de TempVars | Trivial — patrón modal IPC. |
| FormAnexos1.OpenArgs ("Proyecto/Edición/Riesgo") | **Mejorar**: URL path en lugar de string OpenArgs | Web path = /:ambito/:id/anexos. |
| FormTecnicoTareas (TreeView) | **Preservar** jerarquía 7-cat, **Mejorar** virtualización | TreeView es esencial. |

### 4.2 Arquitectura: hexagonal ports

| Concepto legacy | Traducción web |
|---|---|
| `getdb()` (DAO.Database) | Puerto de persistencia PostgreSQL (D14) + HTTP client |
| `Constructor.getXxx()` (factory con cache lazy) | Inyección de dependencias + repository pattern |
| `m_ObjEntorno` (singleton global con ColXxx dictionaries) | Service registry / dependency injection |
| `m_ObjProyectoActivo` / `m_ObjEdicionActiva` / `m_ObjRiesgoActivo` (singleton in scope) | Request-scoped services / Context API |
| `modFormCoordinationHelper.Coord_*` | Navegación declarativa (router + outlet pattern) |
| `PublicacionCalidadStatusPresenter` | ViewModel/Presenter pattern (Zustand, React Query, Redux) |
| `WithEvents + RaiseEvent` | Pub/sub tipado, state machine |
| `TempVars!Variable` IPC | Promise/callback en modal context |
| `fso.FileExists + ShellExecute('open', url)` | window.open(url) + blob URL |
| `Application.FileDialog(3)` | `<input type="file">` + signed URL upload |
| `AbrirEnLocal` (cliente Windows) | Endpoint REST con Content-Type application/octet-stream |
| `Ejecutar Me.hWnd 'open' url` | window.open |
| `Dame()` (helper "existe registro?") | Backend unique constraint o pre-check en POST |
| `WithEvents m_FormExpedientes` (cross-app callback) | API REST + JWT + event bus (Server-Sent Events / WebSocket) |
| `TbCambiosParaPublicacion.NombreCampo='MotivoNoPublicable'` (string discriminator) | Discriminated union pattern en TypeScript |
| `tbCambios` (6725 filas auditing cambios) | Structured log + table |
| `AnexoAntiguo.cls` (legacy con `CodigoUnico`) | Deprecation — preservar en read-only si hay datos |

### 4.3 D88 — Performance

- **NO** tocar el form-level getdb (ya está en 0-2).
- Optimización estructural: backend stateless, pool de conexiones JDBC/HTTP.
- Las 65 clases de dominio son la superficie a portar: cada método DAO se reemplaza por HTTP call con la misma firma.

### 4.4 Decisiones de seguridad

- **Cross-app con Expedientes**: API REST federada con OAuth 2.0 / SSO. Sin mecanismo actual de auth documentado — oportunidad para diseño limpio.
- **Concurrency (D140)**: optimistic locking con `updated_at` token.
- **Publicar no transaccional (D137)**: transacciones explícitas con compensación.
- **Bugs D135, D136**: corrección durante la migración con test coverage.

### 4.5 Decisiones de datos

- **Estado del riesgo**: columna `Estado` opcional para auditoría + state machine en backend que calcula desde fechas.
- **Triángulo plazo/coste/calidad**: 3 columnas separadas (NO consolidar).
- **Anexos**: polimórfica en el backend (1 tabla con CHECK constraint o 6 tablas + vista UNION, según volumen).
- **AnexoAntiguo.cls**: deprecation limpia.

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
- [ ] **CA-F8**: FormAnexos + FormAnexos1 se consolidan en un único `<AnexosManager>` parametrizable por ambito (D142).
- [ ] **CA-F9**: FormWeb se elimina completamente, route directa a HTML (D133).
- [ ] **CA-F10**: Publicar() transaccional con rollback compensatorio (D137).
- [ ] **CA-F11**: El state machine del riesgo tiene optimistic locking (D140).
- [ ] **CA-F12**: Las 22 RN específicas del grupo Publicación documentadas en `legacy.rules.ts`.

### 5.2 Seguridad

- [ ] **CA-S1**: Cross-app API con OAuth 2.0 + scopes granulares.
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

- [ ] **PO-1**: Confirmar el contrato de integración con **Expedientes** (autenticación, endpoint, schema). Sin este contrato, F7 (búsqueda) y F54 (XApp) no se pueden migrar.
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

- [ ] **PO-11**: Smoke test end-to-end: alta proyecto → riesgo → oferta → retirada → calidad visar → publicar → AGEDO.
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
- **TK-GR-9**: [INTEGRATION] API REST con Expedientes + OAuth 2.0 (D132).
- **TK-GR-10**: [CLEANUP] Eliminar FormWeb (D133) — route directa al HTML.

### Forms y migración

- **TK-GR-11**: [MIGRATION] Mover `FormControlCambiosGestion` al módulo transversal "Auditoría de cambios" (D134).
- **TK-GR-12**: [BUGFIX] Corregir `FormRiesgoVulnerabilidad.EstablecerDatos:66-68` (D135).
- **TK-GR-13**: [BUGFIX] Corregir `FormPlanPrincipal:154-156` (D136).
- **TK-GR-14**: [STATE-MACHINE] Publicar() transaccional con rollback compensatorio (D137).
- **TK-GR-15**: [REFACTOR] Colapsar 32 WithEvents/RaiseEvent en un único `OnDecision(payload)` tipado (D138).
- **TK-GR-16**: [REFACTOR] Migrar `m_NodoSeleccionado` (D139) a URL state o Context.
- **TK-GR-17**: [CONCURRENCY] Optimistic locking con `updated_at` (D140).
- **TK-GR-18**: [REFACTOR] Consolidar `modAnexosListPresenter` (D141).
- **TK-GR-19**: [REFACTOR] Consolidar FormAnexos + FormAnexos1 en `<AnexosManager>` (D142).

### Calidad y limpieza

- **TK-GR-20**: [CLEANUP] Corregir captions y tags mal asignados en G1 + G2 (sección H21).
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
- **TK-GR-31**: [PERFORMANCE] Benchmark de la migración de las 372 getdb() calls — antes/después.
- **TK-GR-32**: [PERFORMANCE] Load test del state machine con 100+ usuarios simultáneos (D140).

### Datos y migración

- **TK-GR-33**: [MIGRATION] Backfill de `Estado` columna para los 2924 riesgos actuales.
- **TK-GR-34**: [MIGRATION] Backfill de `FechaPublicacion` para ediciones históricas.
- **TK-GR-35**: [MIGRATION] Migrar 71 tablas a PostgreSQL con Expand and Contract (D82).
- **TK-GR-36**: [MIGRATION] Reemplazar ~12 tablas sequencers (`TbIDXxx`) por IDENTITY/SERIAL.
- **TK-GR-37**: [MIGRATION] Reemplazar 5 tablas `TbCache*` por Redis o similar server-side cache.
- **TK-GR-38**: [MIGRATION] Reemplazar filesystem anexos (`URLDirectorioDocumentacion`) por S3-compatible.
- **TK-GR-39**: [MIGRATION] Migrar `TbLogPublicaciones` (append-only) a log estructurado.
- **TK-GR-40**: [MIGRATION] Migrar `tbCambios` (6725 filas) a log estructurado inmutable.

---

## Anexo · Tabla de fuentes

| Fuente | Aporta |
|---|---|
| engram topic_key `gestion-riesgos/deep-discovery-2026-08-06` | **Consolidado** de 5 sub-agentes walkthrough |
| engram obs #24097 | Audit codegraph-vba previo (372 getdb() callers) |
| engram topic_key `access2web-blueprint/methodology-2026-08-06` | Decisiones metodológicas |
| `data/staging/gestion-riesgos/src/` (65 cls + 139 bas + 62 form .cls) | Source tree exportado |
| `data/staging/gestion-riesgos/.codegraph-vba/` | Índice codegraph-vba regenerado |
| `data/staging/gestion-riesgos/docs/` | Documentación preexistente (Lote 3) |
| `docs/03-aplicaciones/gestion-riesgos/*.md` (7 archivos) | Documentos del estudio (capabilities, data-model, forms, etc.) |
| `C:\00repos\codigo\00_GESTION_RIESGOS\staging` | Fuente READ-ONLY |

## Anexo · Decisiones referenciadas (D5-D140)

| Decisión | Aplicación a Gestion_Riesgos |
|---|---|
| D8 (hexagonal global) | Toda la migración |
| D9-D10 (secret manager) | Acceso a `getdbLanzadera` (vía `URL` de la app Lanzadera) |
| D14 (esquema por módulo) | Módulo `gestion-riesgos` en PostgreSQL |
| D16 (object storage S3-compatible) | Anexos (D142, D143) |
| D27 (logs estructurados) | Reemplazar `tbCambios` + `TbLogPublicaciones` (D137, D139) |
| D44-D46 (autorización + capabilities) | Roles Técnico/Calidad (D139 — permisos diferenciados por tab) |
| D66-D67 (stack) | Backend Python + HTMX |
| D68 (monolito modular) | Migration a módulo dedicado |
| D82 (Expand and Contract) | Estado del riesgo backward-compatible (D126) |
| D88 (308 → 372 getdb() callers) | **CRITICAL PERFORMANCE** — el cuello está en clases, no en forms |
| D102 (booleanos Text(2)) | H21 — cleanup |
| D93-D95 (gestión de secretos) | AGEDO (D129) + cross-app con Expedientes (D132) |

## Checklist del documento

- [x] Scope con ~61 features detalladas por dominio (G1-G5)
- [x] Auditoría de uso + D88 medido a nivel global (372 getdb callers)
- [x] Hallazgos D126-D145 + H1-H21 con anchor links
- [x] Decisiones UX Preservar/Mejorar/Nuevo paradigma por pantalla
- [x] Decisiones arquitectura hexagonal ports
- [x] Criterios de aceptación verificables y agrupados por dimensión
- [x] 14 pendientes operacionales antes, 10 durante, 4 en go-live
- [x] 40 tickets derivables preview (TK-GR-1..40)
- [x] Tabla de decisiones referenciadas (D5-D140)
- [x] Tabla de fuentes
- [x] Idioma: español técnico neutro. Identificadores y paths sin traducir.
- [x] D88 particular: optimizaciones estructural + UI; el grueso está en capa de clases

## Siguiente paso

Revisión con el equipo. Esta épica se revisa junto con la de Brass y las de las otras apps (HPS, HPS_Solicitudes, Condor ya existentes). El siguiente lote es una de las 4 apps restantes (Gestion_Riesgos COMPLETO, NoConformidades, Lanzaderas, Expedientes).
