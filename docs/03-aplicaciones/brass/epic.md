# Épica — Brass (migración a web)

> **Estado:** DRAFT — pendiente revisión al final del ciclo.
> **Versión:** v0.1 (2026-08-06).
> **App legacy:** Brass (`C:\00repos\codigo\00_BRASS`) · frontend+backend `Gestion_Brass_Gestion.accdb` (un solo .accdb, 62 MB).
> **Lote de discovery:** Lote 6 (D104 security CRITICAL).
> **Cross-refs:** `docs/03-aplicaciones/brass/{capabilities,data-model,forms,integrations-automation,migration-matrix,security-rules,README}.md` · engram obs #24097 (audit codegraph-vba) · engram topic_key `brass/deep-discovery-2026-08-06` (consolidado de 5 sub-agentes).

## Metadatos

| Campo | Valor |
|---|---|
| **Aplicación legacy** | `00_BRASS` / `Gestion_Brass_Gestion.accdb` (frontend + backend en un único .accdb) |
| **Tipo de migración** | Legacy Access/VBA → web hexagonal (FastAPI + HTMX) |
| **Scope size** | **L** (88 forms, 27 clases, 16 módulos, 6 tablas de hechos de facturación, anexo polimórfico sobre 6 entidades) |
| **Dependencias cross-app** | Lanzadera (identidad/permisos, `getdbLanzadera`); NoConformidades (eventos vinculados); Gestion_Riesgos (catálogos compartidos presumiblemente) |
| **Riesgo dominante** | **D104** — contraseña `dpddpd` REAL hardcodeada en `Variables Globales.bas:526,560` y `Funciones Generales.bas:1334` (CRITICAL — abre un `.accdb` externo, no es Brass) |
| **Stack target** | Backend Python 3.12+ / FastAPI 0.119+ / SQLAlchemy 2.0.x / Alembic 1.13+ / asyncpg 0.30+ (D66) · Frontend HTMX 2.0.4 + Jinja2 3.1+ + Alpine.js 3.15+ (D67) |
| **Estrategia de migración de BD** | Expand and Contract backward-compatible (D82) · PostgreSQL compartido con esquema por módulo (D14) |
| **Forma destino** | Hexagonal global (D8) · módulo dentro del monolito modular (D68) · puerto de persistencia PostgreSQL + object storage S3-compatible (D16) + secret manager (D9-D10) |
| **Auditoría de uso previa** | ✅ Codegraph-vba + Dysflow walkthrough (85/88 forms, 3 en formato binario denso ilegible) |

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

### 1.1 En scope (47 features de negocio)

#### Grupo G1 — Configuración, Equipos, Calibración (8 features)

| # | Feature | Respaldo |
|---|---|---|
| F1 | **Menú principal** + sub-menú "Otras opciones" (12 opciones de maestros) | `Form0BDOpciones`, `Form0BDOpcionesOtras` |
| F2 | **Gestión de Nodos** (CRUD + borrado bloqueado si hay eventos) | `FormNodoGestion` |
| F3 | **Gestión de BUI** (CRUD + relación N:N con Nodos) | `FormBUIGestion`, `FormNodoBUIGestion` |
| F4 | **Gestión de Subsistemas** (por BUI, con cosecha desde histórico) | `FormSubSistema*` |
| F5 | **Gestión de Ubicaciones** (CRUD bloqueado si en TbActividades) | `FormUbicacionGestion` |
| F6 | **Equipos de Medida** (instrumentos físicos con NS/PN/Marca/Modelo, reflexión sobre Tag='DATO') | `FormEquipoMedida`, `FormEquiposMedidaGestion` |
| F7 | **Calibraciones de Equipos de Medida** (ciclo de calibración + alerta 75 días) | `FormEquipoMedidaCalibracion`, `FormEquipoMedidaCalibracionesGestion` |
| F8 | **Equipos Calibrables** (no son Equipos de Medida; equipos que reciben calibración como servicio) | `FormCalibracion*`, `FormCalibracionEquipo*` |

#### Grupo G2 — Eventos, SLA, Planificación (10 features)

| # | Feature | Respaldo |
|---|---|---|
| F9 | **Eventos (maestro)** — listado con 33 botones, 6 filtros, 3 listas sincronizadas | `FormEventoGestion` |
| F10 | **Alta/Edición de Evento** (53+ controles, validación SLA inline) | `FormEventoAlta`, `FormEventoEdicion` |
| F11 | **Franqueo de Evento** (modal con SLAValidator.ValidarCamposSLA) | `FormEventoFranqueo` |
| F12 | **SLA — Edición exclusiva de los 11 campos SLA** (4 bloques: TRES, TRCM, TRSS, SLA-4) | `FormEventoSLA` |
| F13 | **Copia de Evento** (cabecera, NO copia líneas hijas) | `FormEventoCopia` |
| F14 | **Alta masiva por NS** (batch input, validación previa, reporte éxito/fallo) | `FormEventosAltaMasivaPorNS` |
| F15 | **Equipos de Medida del Evento** (subform embebido WithEvents) | `FormEventoEqMedida` |
| F16 | **Informe SLA corporativo** (donuts HTML Mistica, 4 cards + 4 donuts + tabla expandible) | `FormInformeSLA` |
| F17 | **Planificación preventiva** (alta/cierre/reprogramación/cambio evento/anexos) | `FormPlanificacion*` |
| F18 | **Cuadro de mando de planificación** (11 contadores con drill-down) | `FormPlanificacionGestion` |

#### Grupo G3 — Actividad, Materiales, Anexos (8 features)

| # | Feature | Respaldo |
|---|---|---|
| F19 | **Actividades** (CRUD; regla 7.5h/día por técnico; bloqueado si evento facturado) | `FormActividad*` |
| F20 | **Tipos de Actividad** (catálogo, previamente usados no editables) | `FormActividadTipo*` |
| F21 | **Materiales — Reparación** (item existente + filtro hardcoded EquipoID=538 ⚠️) | `FormMaterialAltaReparacion`, `FormMaterialReparacionEdicion` |
| F22 | **Materiales — Repuesto** (item nuevo, texto libre) | `FormMaterialRepuestoAlta`, `FormMaterialRepuestoEdicion` |
| F23 | **Materiales — Intervención** (cronología de trabajo sobre un material) | `FormMaterialIntervencionAlta`, `FormMaterialIntervencionEdicion`, `FormMaterialSeguimiento` |
| F24 | **Router Material** (decisor Reparación vs Repuesto) | `FormMaterialEleccion` |
| F25 | **Anexos polimórficos** (6 entidades: Evento, Actividad, Material, MaterialSeg, Subcontratación, Gasto) | `FormAnexos`, `FormAnexosNoAlcanzables` |
| F26 | **Catálogo Tipo Acción** (ESTADOREPARABLE) | via `FormMaterial*` |

#### Grupo G4 — Facturación, Gastos, Informes (8 features)

| # | Feature | Respaldo |
|---|---|---|
| F27 | **Alta de Factura** (generador con simulación; valida 6 tablas de hechos) | `FormFacturaAlta` |
| F28 | **Resultado de Factura** (preview 13 precios calculados; modo SIMULACRO o FACTURA) | `FormFacturaResultado` |
| F29 | **Gestión de Facturas** (combo de facturas + acciones: anexar, eliminar, LI, partes) | `FormFacturacionGestion` |
| F30 | **Gastos y Dietas** (autocompleta precio vigente desde `TbGastosImportePorTipo`) | `FormGasto*` |
| F31 | **Informe desde Lista** (staging `TbAuxEventosParaInforme`) | `FormInformesDesdeLista` |
| F32 | **Informe para RAC** (auto-marca `FechaEnInformeRAC` para no duplicados) | `FormInformesParaRAC` |
| F33 | **Informe por Factura** (filtrado por IDFactura) | `FormInformesPorFactura` |
| F34 | **Listado de Intervención (LI)** (3 hojas: SUBCONT_ASISTENCIA_ING, MATERIALES, EVENTOS) | `FormFacturaAlta.ComandoObtenerLI` |

#### Grupo G5 — Técnicos, Subcontratación, Parte, Originador (13 features)

| # | Feature | Respaldo |
|---|---|---|
| F35 | **Técnicos (CRUD)** (alta con validación Alias único + Tipo + fechas) | `FormTecnico`, `FormTecnicoGestion` |
| F36 | **Tipos de Técnico** (catálogo) | `FormTecnicoTipo` |
| F37 | **Precios de Tipo de Técnico** (SCD tipo 2 manual, tramo activo + histórico) | `FormTipoTecnicoPrecio`, `FormTipoTecnicoPreciosGestion` |
| F38 | **Ausencias de Técnico** (rango de fechas; crea un LIbranza por día laborable) | `FormTecnicoAusencias`, `FormTecnicoAusenciasGestion` |
| F39 | **Fiestas** (catálogo global no laborable) | `FormTecnicoFiestas`, `FormTecnicoFiestasGestion` |
| F40 | **Consulta de Horas** (informe de horas laborables/extra por técnico y rango) | `FormTecnicoConsultaHoras` |
| F41 | **Calendario laboral** (weekday + fiestas globales + ausencias del técnico) | `Tecnico.EsDiaLaborable` |
| F42 | **Subcontrataciones externas** (importe único, facturada aparte) | `FormSubContratacion*` |
| F43 | **Partes de trabajo firmados** (1 fila por Centro dentro de una Factura; workflow docx → firma → PDF) | `FormParteGestion` |
| F44 | **Originadores** (catálogo: cliente o depto interno que origina el evento) | `FormOriginadorGestion` |
| F45 | **Jornada laboral** (7.5h/día como constante — D113) | `LIbranza.Registrar` |
| F46 | **Reglas de validación de borrado** (ningún técnico borrable si en TbEventos/TbActividades/TbGastos) | `Tecnico.ValidacionEliminar` |
| F47 | **Bloqueos read-only** (Actividad/Material/Subcontratación read-only si facturado o franqueado) | `lblFacturado`, `lblFranqueado` |

### 1.2 Fuera de scope (REPLACE)

> Las siguientes NO migran como tablas PostgreSQL. Se reemplazan por stack web-native.

- **TbLogCambios, TbLogErrores, TbLogEstados** (si existen) → **Sentry / OpenTelemetry / structured logs**. La web no necesita tablas de log relacionales.
- **Anexos filesystem** (`URLDirectorioAplicaciones\BRASS\ANEXOS\ANEXOS\`) → **S3-compatible** (MinIO / Azure Blob) con versioning habilitado, virus scan, y signed URLs temporales (D114).

### 1.3 Fuera de scope (no documentado)

- 3 forms en formato binario denso (sin .form.txt legible): `FormActividadDeEventoPorLotes`, `FormMaterialIntervencionAlta`, `FormMaterialIntervencionEdicion`. Estructura de UI inferida parcialmente. **Abrir en Access IDE para validación final.**

---

## 2. Estado del descubrimiento

### 2.1 Inventario Dysflow (perfado 2026-08-05)

| Categoría | Resultado |
|---|---|
| **Backend** | `Gestion_Brass_Gestion.accdb` (único .accdb que actúa como frontend + backend) |
| **Tablas** | ~60 (inventario previo; walkthrough identifica 25+ tablas núcleo + 6 hechos facturación + anexos + auxiliares) |
| **Forms** | **88** (85 walkthroughed, 3 en binario denso) |
| **Clases** | 27 |
| **Módulos** | 16 |
| **Reports** | (inferidos via DameHTML + HTMLENTXT; no listados en el walkthrough) |
| **FK relationships** | Conceptual (no constraints físicos); validadas en código via `Dame()` |
| **Volúmenes de muestra** | 5777 eventos, 25639 actividades (según inventario previo) |

### 2.2 Auditoría de uso de las 27 clases + 16 módulos

| Veredicto | Cantidad | Notas |
|---|---|---|
| **ACTIVE** | 27 clases / 16 módulos | 100% activo via codegraph-vba + Dysflow walkthrough |
| **UNCLEAR** | 0 | — |
| **ZOMBIE** | 0 | — |

### 2.3 Tablas detectadas (consolidado de 5 sub-agentes)

**Núcleo (G1+G2+G5)**: `TbBUI`, `TbNodos`, `TbNodoBUI`, `TbSubsistemaBui`, `TbEquipos`, `TbEquipoMedida`, `TbEquipoMedidaCalibracion`, `TbEquiposCalibrables`, `TbEquiposCalibrablesFechas`, `TbUbicacion`, `TbEventos`, `TbPlanificacion`, `TbPlanificacionEquipos`, `TbPlanificacionAñadida`, `TbTecnicos`, `TbTipoTecnico`, `TbTipoTecnicoPrecios` (SCD tipo 2), `TbTecnicosAusencias`, `TbTecnicosFiestas`, `TbLIbranzas`, `TbSubcontrataciones`, `TbOriginador`, `TbActividades`, `TbCodActividad`, `TbUbicacion`, `TbEventosEqMedida`, `TbMaterial`, `TbMaterialSeguimiento`, `TbTipoAccion`, `TbAnexos` (polimórfica 6 FKs)

**Facturación (G4)**: `TbFacturaPrincipal`, `TbFacturaPrincipalPerfiles`, `TbFacturaEventosInvolucrados`, `TbFacturaActividadesInvolucradas`, `TbFacturaMaterialesInvolucrados`, `TbFacturaSubcontratacionesInvolucradas`, `TbFacturaGastosInvolucrados`, `TbGastos`, `TbGastosImportePorTipo`

**Partes (G5)**: `TbPartesPpal`, `TbPartesDetalle`

**Auxiliares (G4)**: `TbAuxEventosParaInforme` (staging), `Tb0FiltroGestion` (histórico de filtros), `TbSellados`, `TbFranqueo`

**Cross-cutting**: `TbAuxPlanificacion` (volcado temporal)

---

## 3. Hallazgos críticos

### H1 (D104 — CRITICAL) — Password REAL hardcodeado

**Síntoma**: `Variables Globales.bas:526,560` y `Funciones Generales.bas:1334` contienen `wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=dpddpd")` y `strPassActual = "dpddpd"`. Es la **contraseña REAL** para abrir un `.accdb` externo (Lanzadera típicamente), no un fallback.

**Impacto**: si la contraseña real es `dpddpd`, rotarla NO sirve (la rotación queda bypassed). Si `dpddpd` era dummy, el código la usa igualmente porque no hay path alternativo.

**Acción inmediata** (TK-BRASS-5):
1. Auditar git history de `00_BRASS` para confirmar si `dpddpd` fue contraseña de producción.
2. Rotar la contraseña real y mover a secret manager (D9-D10).
3. Eliminar `dpddpd` del código, `git filter-repo` para borrar del historial. **Crítico**: hacerlo ANTES de migrar a web (no commitear nada con secretos al repo).

### H2 (D105) — No hay módulo de notificaciones

**Síntoma**: SLA incumplimiento se ve solo en el donut HTML de `FormInformeSLA`. No hay `modNotificaciones`, `modAlertas`, ni equivalente. Los `MsgBox` son el único feedback.

**Impacto**: oportunidad enorme en web (email, webhook, push). El SLA incumple y nadie se entera hasta el informe mensual.

**Acción** (TK-BRASS-9): sistema de notificaciones nativo en web con canales email + webhook; integración con Sentry/Telegram/Teams.

### H3 (D106) — SLA architecture: decoupled 1:1

**Síntoma**: `SLAValidator.bas` (validación) + `SLAReportService.bas` (cálculo + dataset) + `SLAHTMLService.bas` (render). Tres unidades independientes, sin dependencias cruzadas.

**Impacto**: arquitectura a preservar **1:1** en la web. Tres servicios / tres módulos backend.

### H4 (D107) — SLA con 3 estados (Cumple/Inconsistente/No cumple)

**Síntoma**: TRCM y TRSS computan 3 estados. El "Inconsistente" (datos faltantes, fechas invertidas) es **warning**, no error.

**Acción** (TK-BRASS-10): badge diferenciado por estado en el UI web. **No** usar el mismo color rojo para "No cumple" e "Inconsistente".

### H5 (D108) — `modMigracionesSLA` idempotente

**Síntoma**: `MigrarCamposSLA` detecta existencia y solo crea los 12 campos SLA faltantes en `TbEventos`.

**Acción** (TK-BRASS-11): patrón a replicar en web. Cada migración de BD con `IF NOT EXISTS` y log de drift.

### H6 (D109) — `TRSS_USA_DIAS_LABORABLES` boolean switch

**Síntoma**: `SLAReportService.bas:20` — booleano que cambia entre "días corridos" vs "días laborables".

**Acción**: en web, mapear a env var (`TRSS_USA_DIAS_LABORABLES=true|false`) o feature flag configurable. **No** hardcodear.

### H7 (D110) — HTML corporativo con tokens Mistica

**Síntoma**: `SLAHTMLService.GenerarInformeHTML_Corporate` usa tokens Mistica (`--color-highlight #0066FF`, `--color-success #5CB85C`, `--color-error #E66C64`, `--radius-button 32px`).

**Acción** (TK-BRASS-12): el informe web debe ser server-side rendered con los mismos tokens. Email/notificación con el mismo aspect ratio.

### H8 (D111) — Form con 33 botones

**Síntoma**: `FormEventoGestion` — 8 botones `ComandoLimpiar*` + 25 acciones. Overload visible.

**Acción** (TK-BRASS-13): en web, agrupar las `Limpiar*` en un chip "Clear all" + menú kebab para acciones secundarias. Reducción 1:8.

### H9 (D112) — Hardcode `EquipoID = 538`

**Síntoma**: `FormMaterialAltaReparacion.form.txt` filtro `WHERE (TbEventos.EquipoID) =538`. Filtra reparaciones de un solo equipo.

**Acción** (TK-BRASS-14): sustituir por `m_ObjEntorno.EquipoIDPorDefecto` o por `OpenArgs.IDEquipo`. Probablemente un bug histórico a confirmar con Natalia.

### H10 (D113) — Jornada laboral hardcoded 7.5h

**Síntoma**: `LIbranza.Registrar` setea `Me.horas = 7.5` SIEMPRE antes del INSERT.

**Acción** (TK-BRASS-15): parametrizar jornada por técnico o por tipo de técnico. Constante configurable via `TbTecnicos.JornadaLaboral` o `TbTipoTecnico.JornadaDefault`.

### H11 (D114) — Anexos filesystem-based

**Síntoma**: `Anexo.cls` copia `fso.CopyFile LocalOrigen → Destino` en `URLDirectorioAplicaciones\BRASS\ANEXOS\ANEXOS\`. Filesystem localizado, no BLOB.

**Riesgos**: offline sin red → falla; sin antivirus → archivos infectados; sin versioning → sobrescritura pierde el anterior; carpetas compartidas Windows → problemas de permisos NTFS.

**Acción** (TK-BRASS-16): migrar a S3-compatible (MinIO/Azure Blob) con versioning + virus scan + signed URLs temporales. Las URLs firmadas son equivalentes a `fso.FollowHyperlink`.

### H12 (D115) — `TbAuxEventosParaInforme` (staging)

**Síntoma**: tabla temporal que alimenta `FormInformesDesdeLista`. El feeder externo no se identificó en el walkthrough.

**Acción**: identificar el feeder (¿otro proceso que la puebla? ¿es manual?). Si es externa, documentar el contrato.

### H13 (D116) — `Tb0FiltroGestion` (histórico de filtros)

**Síntoma**: `GestorInforme.cls` persiste cada filtro aplicado (Fecha, Datos [pipe-separated], Usuario, Visto=Sí/No).

**Acción** (TK-BRASS-17): en web, esto tiene sentido como log de auditoría. Mantener como tabla relacional o migrar a log nativo.

### H14 (D117) — 6 tablas de hechos de facturación

**Síntoma**: `TbFacturaEventosInvolucrados`, `TbFacturaActividadesInvolucradas`, `TbFacturaMaterialesInvolucrados`, `TbFacturaSubcontratacionesInvolucrados`, `TbFacturaGastosInvolucrados`, `TbFacturaPrincipalPerfiles`. Cada `TbFacturaPrincipal` se desglosa en estas 6 tablas.

**Acción**: las 6 tablas son el corazón de la trazabilidad Técnico → Tipo → Factura. **No** consolidar en una sola — preserva el particionamiento.

### H15 (D118) — Parte = documento por Centro

**Síntoma**: `TbPartesPpal` + `TbPartesDetalle` (1:N eventos). El parte es un documento firmado por cada Centro dentro de una Factura.

**Acción** (TK-BRASS-18): workflow web: generar plantilla docx → cliente (Centro) firma → escanea PDF → upload. Múltiples centros firman partes distintos de la misma factura.

### H16 (D119) — Precios versionados (SCD tipo 2 manual)

**Síntoma**: `TipoTecnicoPrecio.Registrar` al dar de alta un tramo nuevo, hace `UPDATE` del tramo activo poniéndole `FechaFinal=Date`.

**Acción**: el modelo web debe preservar el histórico (probablemente flag `activo` + tabla de histórico o vista materializada).

### H17 (D120) — `TipoImpositivo=21%` y `Recargo=15%` hardcoded

**Síntoma**: `Factura.cls:185-194` y `FormFacturaAlta.Form_Load:221-222`.

**Acción** (TK-BRASS-19): parametrizar via `TbConfiguracion` (IVA España, RecargoEquipos nacionales). Migración a multi-cliente requiere esto configurable.

### H18 (D121) — Código muerto enterrado en `FormInformeSLA`

**Síntoma**: `cmdExportarExcel_Click` (líneas 349-614) está comentado. ~265 líneas de código legacy.

**Acción**: en la web, no migrar. Limpieza natural.

### H19 (D122) — Singleton globals

**Síntoma**: `m_ObjEquipoMedidaActivo`, `m_ObjCalibracionEquipoMedidaActiva`, `m_ObjFacturaActiva`, `m_ObjEntidadParaAnexoActiva`, `m_ColEventosParaInformeDeRAC`, `m_ObjUsuarioConectado`, `m_TextoWin64`. Estado compartido entre forms sin contrato explícito.

**Acción**: en web, traducir a Context API / store global (`/evento-activo`, `/factura-activa`, `/equipo-medida-activo`). El contrato debe ser tipado.

### H20 (D123) — Reflection sobre Tag='DATO'

**Síntoma**: `FormEquipoMedida`), `FormEquipoMedidaCalibracion`: foreach `Me.Controls` con tag='DATO' → `getPropiedad`/`SetPropiedad` sobre la clase de dominio.

**Acción**: en web, **model binding estándar** (HTML forms + backend DTOs). No replicar el patrón Reflection.

### H21 (D124) — WithEvents + RaiseEvent pattern

**Síntoma**: `FormEquipoMedida` ↔ `FormEquipoMedidaCalibracion` ↔ `FormEquipoMedidaCalibracionesGestion` se comunican via `WithEvents` + `RaiseEvent` (sin OpenArgs).

**Acción**: en web, traducir a store global / navegación declarativa. Los eventos encadenados desaparecen con un patrón de composición React/HMX.

### H22 (D125) — `Numero_Actividad` race condition

**Síntoma**: `FormActividadAlta` calcula `IDActividad = DCount("Tabla_Parte_Actividad") + 1`. Race condition potencial en concurrencia.

**Acción**: migrar a IDENTITY (autonumérico). **No** mantener el patrón legacy.

### H23 (Cross-cutting) — D102 booleanos como Text(2)

**Síntoma**: `TbEquiposCalibrables`, `TbTecnicos`, etc. tienen booleanos como `Text(2)` 'Si'/'No'.

**Acción**: en PostgreSQL, `BOOLEAN`. **No** migrar como `Text(2)`.

### H24 (Cross-cutting) — Codificación de fechas

**Síntoma**: `LIbranza.horas = 7.5`, `FormPlanificacionReprogramacion.FechaViernes()`, `FechaFinCalibracion`. Mezcla de `Date` y `String`.

**Acción**: en web, todo `DATE` / `TIMESTAMP` nativo. Sin strings para fechas.

---

## 4. Decisiones aplicadas

### 4.1 UX: Preservar / Mejorar / Nuevo paradigma

| Pantalla | Decisión | Justificación |
|---|---|---|
| FormEventoGestion (33 botones) | **Mejorar**: consolidar `Limpiar*` en un chip. Menú kebab para acciones secundarias. | D111 — overload visible. |
| FormEventoAlta (53+ controles) | **Preservar** cascada BUI/Subsistema/Equipo. **Mejorar**: wizard 2-3 pasos. | Cascada es dominio real. Pero dense form es UX pobre. |
| FormTecnico (alta/edicion) | **Nuevo paradigma**: alta inline en lista (spreadsheet-style). | Modal + tabs es overkill para "agregar un técnico". |
| FormPlanificacionReprogramacion | **Nuevo paradigma**: wizard con date picker visual. **Mantener** "motivo obligatorio". | Reprogramación es operación delicada + auditable. |
| FormInformeSLA | **Preservar** 100% del comportamiento. **Mejorar** donuts interactivos (click → filtrado). | El HTML Mistica es el activo más cuidado de la app. |
| FormAnexos | **Nuevo paradigma**: upload directo a S3 con preview inline. **No** mantener filesystem local. | D114 — cambio de paradigma completo. |
| Form0BDOpciones (menú plano) | **Mejorar**: sidebar persistente con iconos. **Preservar** jerarquía funcional. | Menú plano es UX pobre; sidebar es estándar. |
| Forms catálogo (BUI, Nodo, SubSistema) | **Preservar** validación de unicidad + bloqueo por eventos. **Mejorar**: edición inline. | CRUD simple no necesita alta como form modal. |
| FormFacturaAlta | **Preservar** simulación + validación dura. **Mejorar**: preview por tipo impositivo configurable. | Flujo crítico, no revolucionar. |
| FormParteGestion | **Nuevo paradigma**: generador de docx → cliente escanea → upload PDF. | El workflow actual es arcaico. |
| FormMaterialesGestion (datos hardcoded) | **Rehacer**: SELECT dinámico o eliminar el formulario. | D-05 — es placeholder visual, no funcional. |
| FormTecnicoConsultaHoras (sin .cls) | **Rehacer**: extraer lógica a servicio backend. | D115 — código embebido en form.txt. |

### 4.2 Arquitectura: hexagonal ports

| Concepto legacy | Traducción web |
|---|---|
| `getdb()` (DAO.Database) | Puerto de persistencia PostgreSQL (D14) |
| `Constructor.getXxx()` (factory con cache lazy) | Inyección de dependencias + repository pattern |
| `m_ObjXxxActivo` (singleton global) | Context API / store tipado |
| `fso.CopyFile` (anexos) | SDK S3-compatible (MinIO o Azure Blob) |
| `SLAValidator`, `SLAReportService`, `SLAHTMLService` | 3 servicios / 3 paquetes (D106 — preservar 1:1) |
| `Anexo.TipoAnexo` enum (polimórfico) | Tabla única con CHECK constraint o 6 tablas + vista UNION |
| `TbAuxEventosParaInforme` (staging) | Tabla staging local + job async de cleanup |
| `Tb0FiltroGestion` (histórico filtros) | Log de auditoría (conservar o migrar a log nativo) |
| `m_SQL = "..."` (SQL inline en `*Repositorio.bas`) | SQLAlchemy 2.0 con typed queries |
| `OpenArgs` con sentinel `|` | URL params / query params tipados |
| `Tag='DATO'` reflection | Model binding declarativo (D123) |
| `WithEvents` + `RaiseEvent` | Navegación declarativa (D124) |

### 4.3 Seguridad D104

- **Eliminar `dpddpd` del código**: TK-BRASS-5.
- **Mover a secret manager**: TK-BRASS-5.
- **`git filter-repo` para borrar del historial**: TK-BRASS-5 (urgente, antes de cualquier push).
- **Auditar git history de `00_BRASS`**: TK-BRASS-5.

### 4.4 Datos

- **TipoImpositivo + Recargo** van a `TbConfiguracion` (no hardcoded).
- **Jornada laboral** por técnico o por tipo (no constante 7.5h).
- **Booleano** como `BOOLEAN` en PostgreSQL (no `Text(2)` 'Si'/'No').
- **Fechas** como `DATE`/`TIMESTAMP` (no string).
- **Anexos** como blob en S3-compatible, no filesystem.

---

## 5. Criterios de aceptación

### 5.1 Funcionalidad

- [ ] **CA-F1**: Las 47 features F1-F47 tienen paridad funcional con la versión Access.
- [ ] **CA-F2**: Todos los `MsgBox` del legacy se reemplazan por notificaciones nativas web (toast + center de notificaciones).
- [ ] **CA-F3**: SLA cumple / inconsistente / no cumple con badges de colores diferenciados (verde / naranja / rojo).
- [ ] **CA-F4**: Padre de SLA en TRES, TRCM, TRSS por criticidad (1=1h, 3=3h, 5=8h TRES; 1=1d, 3=5d, 5=15d TRSS).
- [ ] **CA-F5**: Reprogramación de planificación con motivo obligatorio auditable.
- [ ] **CA-F6**: Parte firmado por Centro con workflow docx → firma → PDF.
- [ ] **CA-F7**: Anexos en S3-compatible con versioning, virus scan, signed URLs temporales.
- [ ] **CA-F8**: 6 tablas de facturación preservadas (no consolidar).
- [ ] **CA-F9**: Versionado de precios (SCD tipo 2) preservado.
- [ ] **CA-F10**: Reglas de validación de borrado (ningún maestro borrable si en TbEventos) preservadas.

### 5.2 Seguridad

- [ ] **CA-S1**: `dpddpd` eliminado del código fuente y del historial de git (`git filter-repo`).
- [ ] **CA-S2**: Password real rotada, en secret manager (D9-D10).
- [ ] **CA-S3**: Autenticación vía adaptador unificado (D9-D10, NO acoplamiento directo a Lanzadera).
- [ ] **CA-S4**: PII de eventos/actividades protegida. `TbEventos` no expone DNI, etc. via API sin scope explícito.

### 5.3 Performance

- [ ] **CA-P1**: `FormInformeSLA` con 1 año de datos rinde en < 3s.
- [ ] **CA-P2**: Anexos con upload directo a S3 (no proxy pelo backend).
- [ ] **CA-P3**: Listados (eventos, actividades, materiales) con paginación server-side.

### 5.4 Operacional

- [ ] **CA-O1**: Migración de datos batch con verificación de drift (obs equivalentes a `modMigracionesSLA`).
- [ ] **CA-O2**: Logs operativos en Sentry/OTel (no en PostgreSQL).
- [ ] **CA-O3**: Monitoreo de SLA incumplimiento con alertas a email/webhook.

---

## 6. Pendientes operacionales

### 6.1 Antes de empezar

- [ ] **PO-1**: Auditar git history de `00_BRASS` para confirmar si `dpddpd` fue contraseña de producción.
- [ ] **PO-2**: Rotar la contraseña real y mover a secret manager (D9-D10).
- [ ] **PO-3**: Ejecutar `git filter-repo` para borrar `dpddpd` del historial.
- [ ] **PO-4**: Confirmar con Natalia el hardcode `EquipoID = 538` (D112).
- [ ] **PO-5**: Confirmar TipoImpositivo y Recargo configurables (D120).
- [ ] **PO-6**: Identificar el feeder de `TbAuxEventosParaInforme` (D115).

### 6.2 Durante el desarrollo

- [ ] **PO-7**: Mantener `SLAValidator` / `SLAReportService` / `SLAHTMLService` como 3 servicios backend separados (D106).
- [ ] **PO-8**: Los 3 forms de Informes (`DesdeLista`, `ParaRAC`, `PorFactura`) comparten el mismo motor `Informe.EventosRAC()` — preservar en backend.
- [ ] **PO-9**: Mantener `Datos_Cliente` parametrizable (no hardcoded `BRASS`).
- [ ] **PO-10**: Documentar las 6 FKs polimórficas de `TbAnexos` antes de migrar (decisión: tabla única o 6 tablas + vista).

### 6.3 En el go-live

- [ ] **PO-11**: Validar las 9 reglas documentadas de "ningún maestro borrable si hay eventos" contra el modelo de datos.
- [ ] **PO-12**: Smoke test: 1 evento end-to-end (alta → actividad → material → franqueo → factura → informe → parte).
- [ ] **PO-13**: Backfill de `TbTipoTecnicoPrecios` (tramo activo + histórico).
- [ ] **PO-14**: Verificar que las alertas de SLA incumplimiento funcionen (no `MsgBox`).

---

## 7. Tickets derivables (preview)

> NO crear issues todavía. Estos nacen de la épica al final del ciclo de revisión.

### Seguridad (CRITICAL)

- **TK-BRASS-1**: [SECURITY] Auditar git history de `00_BRASS` para `dpddpd` (raíz desde `Variables Globales.bas:526,560` y `Funciones Generales.bas:1334`).
- **TK-BRASS-2**: [SECURITY] Rotar contraseña real del `.accdb` externo (Lanzadera típicamente) y mover a secret manager.
- **TK-BRASS-3**: [SECURITY] `git filter-repo` para borrar `dpddpd` del historial de `00_BRASS`.
- **TK-BRASS-4**: [SECURITY] Eliminar fallback `dpddpd` de `Variables Globales.bas`; usar `Err.Raise` si falla `GetPasswordDB`.

### Funcionalidad core

- **TK-BRASS-5**: [CRUD] Gestión de Nodos/BUIs/Subsistemas/Ubicaciones (F2-F5) — 4 forms de catálogo con validación de unicidad y bloqueo por eventos.
- **TK-BRASS-6**: [CRUD] Equipos de Medida + Calibraciones (F6, F7) — ciclo de calibración con alerta 75 días.
- **TK-BRASS-7**: [CRUD] Master Evento (F9-F14) — 88 forms → consolidado en: alta, edición, franqueo, copia, masiva por NS, SLA edición exclusiva.
- **TK-BRASS-8**: [ENGINE] SLA — 3 servicios separados (validator, report, html) — preservar arquitectura 1:1.
- **TK-BRASS-9**: [ENGINE] Sistema de notificaciones (push, email, webhook) —填补 D105.
- **TK-BRASS-10**: [UI] Badges SLA 3 estados (cumple/inconsistente/no cumple) — colores diferenciados.
- **TK-BRASS-11**: [MIGRATION] `modMigracionesSLA` idempotente — patrón para todas las migraciones de BD.
- **TK-BRASS-12**: [INTEGRATION] HTML corporativo Mistica — server-side rendering para email/notificación.
- **TK-BRASS-13**: [UX] `FormEventoGestion` 33 botones → consolidated con chip "Clear all" + kebab menu.
- **TK-BRASS-14**: [BUG] Hardcode `EquipoID = 538` en `FormMaterialAltaReparacion` — sustituir por variable.
- **TK-BRASS-15**: [CONFIG] Jornada laboral configurable (no hardcoded 7.5h).
- **TK-BRASS-16**: [STORAGE] Anexos S3-compatible (reemplaza filesystem).
- **TK-BRASS-17**: [AUDIT] `Tb0FiltroGestion` → log de auditoría nativo.
- **TK-BRASS-18**: [WORKFLOW] Parte firmado por Centro — generador docx → upload PDF.
- **TK-BRASS-19**: [CONFIG] `TipoImpositivo` y `Recargo` en `TbConfiguracion` (no hardcoded).
- **TK-BRASS-20**: [CLEANUP] `FormMaterialesGestion` con datos hardcoded — reemplazar por SELECT dinámico o eliminar.
- **TK-BRASS-21**: [REFACTOR] `FormTecnicoConsultaHoras` sin .cls — extraer lógica a servicio backend.
- **TK-BRASS-22**: [BUG] `Numero_Actividad` race condition (D125) — migrar a IDENTITY.
- **TK-BRASS-23**: [DATA] Migrar booleanos `Text(2)` → `BOOLEAN` (D102 cross-cutting).
- **TK-BRASS-24**: [DATA] Migrar fechas string → `DATE`/`TIMESTAMP` (D124 cross-cutting).
- **TK-BRASS-25**: [POLY] Decisión sobre `TbAnexos` polimórfico: 1 tabla con CHECK vs 6 tablas + vista UNION.
- **TK-BRASS-26**: [ENGINE] `TbAuxEventosParaInforme` staging — identificar feeder externo y documentar contrato.
- **TK-BRASS-27**: [INTEGRATION] Feed `TbAuxEventosParaInforme` desde proceso externo (definir interfaz).
- **TK-BRASS-28**: [CRUD] Gestión completa de plancha (F35-F40) — técnicos, tipos, precios, ausencias, fiestas, jornada.
- **TK-BRASS-29**: [CRUD] Subcontrataciones + Partes (F42-F43) — workflow firma por Centro.
- **TK-BRASS-30**: [CRUD] Catálogo Originadores (F44) — simple catálogo.
- **TK-BRASS-31**: [ENGINE] Migración de parte: workflow docx → PDF → upload a S3.
- **TK-BRASS-32**: [REPORT] Informes 3 tipos: Desde Lista, Para RAC, Por Factura — motor `Informe.EventosRAC()` compartido.
- **TK-BRASS-33**: [FACT] `FormFacturaAlta` + `FormFacturaResultado` — flujo SIMULACRO + FACTURA con 6 tablas de hechos.
- **TK-BRASS-34**: [BILL] Gestión de Gastos — autocompletar precio vigente + validación factura.
- **TK-BRASS-35**: [CRUD] Catálogo Tipo Acción / Estado Reparable.
- **TK-BRASS-36**: [INTEGRATION] 6 tablas de hechos facturación (`TbFactura*Involucrados`) — preservar 1:1.
- **TK-BRASS-37**: [STORAGE] S3-compatible para anexos + virus scan + signed URLs.
- **TK-BRASS-38**: [TEST] Tests E2E para los 4 tipos de eventos + workflow completo.
- **TK-BRASS-39**: [MIGRATION] Migrations idempotentes para las 25+ tablas (patrón `modMigracionesSLA`).
- **TK-BRASS-40**: [REVIEW] Análise de los 3 forms binarios densos después de abrir en Access IDE.

---

## Anexo · Tabla de fuentes

| Fuente | Aporta |
|---|---|
| `docs/03-aplicaciones/brass/capabilities.md` | Inventario de features (vista de alto nivel) |
| `docs/03-aplicaciones/brass/data-model.md` | 60 tablas con definiciones + columnas |
| `docs/03-aplicaciones/brass/forms.md` | (walkthrough no incluido — referencia cruzada a epic) |
| `docs/03-aplicaciones/brass/integrations-automation.md` | Cross-app, testing sandbox, flags |
| `docs/03-aplicaciones/brass/migration-matrix.md` | Mapeo legacy → web por campo |
| `docs/03-aplicaciones/brass/security-rules.md` | D104 security finding |
| `docs/03-aplicaciones/brass/README.md` | Estado del lote, hallazgos críticos |
| `data/staging/brass/src/classes/*.cls` | 27 clases de dominio |
| `data/staging/brass/src/modules/*.bas` | 16 módulos (Constructor, Variables Globales, etc.) |
| `data/staging/brass/src/forms/*.form.txt` | 88 forms con RecordSource/RowSource |
| engram obs #24097 | Audit codegraph-vba |
| engram topic_key `brass/deep-discovery-2026-08-06` | **Consolidado** de 5 sub-agentes walkthrough |

## Anexo · Decisiones referenciadas (D5-D104)

| Decisión | Aplicación a Brass |
|---|---|
| D8 (hexagonal global) | Toda la migración |
| D9-D10 (secret manager) | Rotación `dpddpd` |
| D14 (esquema por módulo) | Módulo `brass` en PostgreSQL |
| D16 (object storage S3-compatible) | Anexos |
| D17-D20 (ports) | Repositorios + adaptadores de identidad |
| D27 (logs estructurados canónicos) | Reemplazar logs VBA |
| D44-D46 (autorización + capabilities) | Roles Técnico/Calidad/Administrador |
| D66-D67 (stack backend + frontend) | Toda la migración |
| D68 (monolito modular) | Brass como módulo dentro |
| D82 (Expand and Contract) | Estrategia de migración |
| D92 (datos personales) | NO bloqueante para Brass — no hay PII significativa |
| D93 (password fallback) | APLICAR — eliminar fallback `dpddpd` |
| D102 (booleanos `Text(2)`) | APLICAR — migrar a `BOOLEAN` |
| D104 (secret manager) | CRITICAL — rotar `dpddpd` |

## Checklist del documento

- [x] Scope con 47 features detalladas por grupo
- [x] Auditoría de uso (Dysflow + codegraph-vba) + 27 clases / 16 módulos
- [x] Hallazgos D104-D125 con anchor links
- [x] Decisiones UX Preservar/Mejorar/Nuevo paradigma por pantalla
- [x] Decisiones arquitectura hexagonal ports
- [x] Criterios de aceptación verificables y agrupados por dimensión
- [x] 6 pendientes operacionales antes, 4 durante, 4 en go-live
- [x] 40 tickets derivables preview (TK-BRASS-1..40)
- [x] Tabla de decisiones referenciadas (D5-D104)
- [x] Tabla de fuentes
- [x] Idioma: español técnico neutro. Identificadores y paths sin traducir.

## Siguiente paso

Revisión con el equipo. Una vez validada, abrir SDD (`sdd-propose` + `sdd-spec` + `sdd-design` + `sdd-tasks`) para arrancar la implementación por ticket, comenzando por **TK-BRASS-1** (auditoría de git history para `dpddpd`) antes que cualquier otro cambio de código.
