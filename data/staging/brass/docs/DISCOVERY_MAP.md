# BRASS — Discovery Map

## Estado: EN PROGRESO
**Última actualización:** 2026-03-24
**Última actualización de specs:** 2026-03-24
**Inicio del autodescubrimiento:** 2026-03-11

---

## Vista Global

BRASS es una aplicación Microsoft Access/VBA para **gestión de mantenimiento técnico** en el sector de telecomunicaciones. Maneja:
- Eventos de mantenimiento (correctivo, preventivo)
- Actividades técnicas associateadas a eventos
- Materiales y repuestos
- Planificación preventiva de equipos
- Subcontrataciones
- Facturación por eventos y actividades

---

## Módulos Lógicos Identificados

### 1. Módulo: Gestión de Eventos
**Propósito:** Registro y seguimiento de eventos de mantenimiento

| Componente | Archivo(s) | Tabla(s) BD |
| :--- | :--- | :--- |
| Alta | `Form_FormEventoAlta.cls` | `TbEventos` |
| Edición | `Form_FormEventoEdicion.cls` | `TbEventos` |
| SLA (completo) | `Form_FormEventoSLA.cls` | `TbEventos` |
| Gestión/Listado | `Form_FormEventoGestion.cls` | `TbEventos` |
| Copia | `Form_FormEventoCopia.cls` | `TbEventos` |
| Tipo | `Form_FormEventoTipoGestion.cls` | `TbTipoEvento` |
| Equipos de Medida | `Form_FormEventoEqMedida*.cls` | `TbEventosEquipoMedida` |
| Francqueo | `Form_FormEventoFranqueo.cls` | `TbEventos` |
| Alta Masiva | `Form_FormEventosAltaMasivaPorNS.cls` | `TbEventos` |
| Clase negocio | `Evento.cls` | — |

**Relaciones:**
- `TbEventos.IDEvento` → `TbActividades.IDEvento`
- `TbEventos.IDEquipo` → `TbEquipos.IDEquipo`
- `TbEventos.IDTecnico` → `TbTecnicos.ALIAS`

---

### 2. Módulo: Gestión de Actividades
**Propósito:** Seguimiento de actividades técnicas associateadas a eventos

| Componente | Archivo(s) | Tabla(s) BD |
| :--- | :--- | :--- |
| Alta | `Form_FormActividadAlta.cls` | `TbActividades` |
| Edición | `Form_FormActividadEdicion.cls` | `TbActividades` |
| Gestión/Listado | `Form_FormActividadGestion.cls` | `TbActividades` |
| Por lotes | `Form_FormActividadDeEventoPorLotes.cls` | `TbActividades` |
| Tipos | `Form_FormActividadTipo*.cls` | `TbCodActividad` |
| Clase negocio | `Actividad.cls` | — |

**Relaciones:**
- `TbActividades.IDEvento` → `TbEventos.IDEvento`
- `TbActividades.IDFacturacion` → `TbFacturaPrincipal.IDFactura`

---

### 3. Módulo: Gestión de Técnicos
**Propósito:** Administración de técnicos, ausencias, fiestas y horas

| Componente | Archivo(s) | Tabla(s) BD |
| :--- | :--- | :--- |
| Gestión | `Form_FormTecnicoGestion.cls` | `TbTecnicos` |
| Edición | `Form_FormTecnico.cls` | `TbTecnicos` |
| Tipos/Precios | `Form_FormTipoTecnico*.cls` | `TbTipoTecnico`, `TbTipoTecnicoPrecios` |
| Ausencias | `Form_FormTecnicoAusencias*.cls` | `TbTecnicosAusencias` |
| Fiestas | `Form_FormTecnicoFiestas*.cls` | `TbTecnicosFiestas` |
| Horas consulta | `Form_FormTecnicoConsultaHoras.cls` | `TbTecnicos` |
| Clases negocio | `Tecnico.cls`, `TipoTecnico.cls`, `TipoTecnicoPrecio.cls` | — |

---

### 4. Módulo: Gestión de Equipos
**Propósito:** Catálogo de equipos y equipos calibrables

| Componente | Archivo(s) | Tabla(s) BD |
| :--- | :--- | :--- |
| Gestión | `Form_FormEquipoGestion.cls` | `TbEquipos` |
| Alta | `Form_FormEquipoAlta.cls` | `TbEquipos` |
| Edición | `Form_FormEquipoEdicion.cls` | `TbEquipos` |
| Equipos Medida | `Form_FormEquipoMedida*.cls` | `TbEquiposMedida` |
| Calibraciones | `Form_FormCalibracion*.cls` | `TbEquiposCalibrables*` |
| Equipos calibrables | `Form_FormEquipoMedidaCalibraciones*.cls` | `TbEquiposCalibrables` |
| Clases negocio | `Equipo.cls`, `EquipoMedida.cls`, `EquipoMedidaCalibracion.cls` | — |

---

### 5. Módulo: Gestión de Materiales
**Propósito:** Seguimiento de materiales, repuestos y reparaciones

| Componente | Archivo(s) | Tabla(s) BD |
| :--- | :--- | :--- |
| Gestión | `Form_FormMaterialesGestion.cls` | `TbMaterial` |
| Alta reparación | `Form_FormMaterialAltaReparacion.cls` | `TbMaterial` |
| Alta repuesto | `Form_FormMaterialRepuestoAlta.cls` | `TbMaterial` |
| Alta intervención | `Form_FormMaterialIntervencionAlta.cls` | `TbMaterial` |
| Seguimiento | `Form_FormMaterialSeguimiento.cls` | `TbMaterialSeguimiento` |
| Elección | `Form_FormMaterialEleccion.cls` | `TbMaterial` |
| Clases negocio | `Material.cls`, `MaterialSeguimiento.cls` | — |

---

### 6. Módulo: Gestión de Gastos
**Propósito:** Registro de gastos y dietas

| Componente | Archivo(s) | Tabla(s) BD |
| :--- | :--- | :--- |
| Gestión | `Form_FormGastoGestion.cls` | `TbGastos` |
| Alta | `Form_FormGastoAlta.cls` | `TbGastos` |
| Edición | `Form_FormGastoEdicion.cls` | `TbGastos` |
| Clase negocio | `Gasto.cls` | — |

---

### 7. Módulo: Facturación
**Propósito:** Generación de facturas y perfiles

| Componente | Archivo(s) | Tabla(s) BD |
| :--- | :--- | :--- |
| Gestión | `Form_FormFacturacionGestion.cls` | `TbFacturaPrincipal` |
| Alta | `Form_FormFacturaAlta.cls` | `TbFacturaPrincipal` |
| Resultado | `Form_FormFacturaResultado.cls` | `TbFactura*` |
| Por factura | `Form_FormInformesPorFactura.cls` | `TbFactura*` |
| Para RAC | `Form_FormInformesParaRAC.cls` | `TbFactura*` |
| Desde lista | `Form_FormInformesDesdeLista.cls` | `TbFactura*` |
| Clases negocio | `Factura.cls`, `FacturaPerfiles.cls`, `GestorInforme.cls`, `Informe.cls` | — |

---

### 8. Módulo: Subcontrataciones
**Propósito:** Gestión de trabajos subcontratados

| Componente | Archivo(s) | Tabla(s) BD |
| :--- | :--- | :--- |
| Gestión | `Form_FormSubContratacionGestion.cls` | `TbSubcontrataciones` |
| Alta | `Form_FormSubContratacionAlta.cls` | `TbSubcontrataciones` |
| Edición | `Form_FormSubContratacionEdicion.cls` | `TbSubcontrataciones` |
| Clase negocio | `Subcontratacion.cls` | — |

---

### 9. Módulo: Planificación
**Propósito:** Planificación preventiva de mantenimientos

| Componente | Archivo(s) | Tabla(s) BD |
| :--- | :--- | :--- |
| Gestión | `Form_FormPlanificacionGestion.cls` | `TbPlanificacion` |
| Alta | `Form_FormPlanificacionAlta.cls` | `TbPlanificacion` |
| Equipos | `Form_FormPlanificacionEquipos.cls` | `TbPlanificacionEquipos` |
| Alta equipo | `Form_FormPlanificacionAltaEquipo.cls` | `TbPlanificacionEquipos` |
| Edición equipo | `Form_FormPlanificacionEdicionEquipo.cls` | `TbPlanificacionEquipos` |
| Cierre | `Form_FormPlanificacionCierre.cls` | `TbPlanificacion` |
| Calendarios | `Form_FormPlanificacionCalendarios.cls` | `TbPlanificacion` |
| Anexos | `Form_FormPlanificacionAnexos.cls` | `TbPlanificacionAnexos` |
| Búsqueda | `Form_FormPlanificacionesBusqueda.cls` | `TbPlanificacion` |
| Reprogramación | `Form_FormPlanificacionReprogramacion.cls` | `TbPlanificacion` |
| Cambio ID | `Form_FormPlanificacionCambioIDEvento.cls` | `TbPlanificacion` |

---

### 10. Módulo: Datos Maestros / Catálogos

| Componente | Archivo(s) | Tabla(s) BD |
| :--- | :--- | :--- |
| Nodos | `Form_FormNodoGestion.cls` | `TbNodos` |
| BUI | `Form_FormBUIGestion.cls` | `TbBUI` |
| Nodo-BUI | `Form_FormNodoBUIGestion.cls` | `TbNodoBUI` |
| Subsistemas | `Form_FormSubSistema*.cls` | `TbSubsistemaBui` |
| Ubicaciones | `Form_FormUbicacionGestion.cls` | `TbUbicacion` |
| Originadores | `Form_FormOriginadorGestion.cls` | `TbOriginador` |

---

### 11. Módulo: Anexos y Documentación

| Componente | Archivo(s) | Tabla(s) BD |
| :--- | :--- | :--- |
| Anexos | `Form_FormAnexos.cls` | `TbAnexos` |
| Anexos no alcanzables | `Form_FormAnexosNoAlcanzables.cls` | `TbAnexos` |

---

### 12. Módulo: Partes de Trabajo

| Componente | Archivo(s) | Tabla(s) BD |
| :--- | :--- | :--- |
| Gestión | `Form_FormParteGestion.cls` | `TbPartesPpal`, `TbPartesDetalle` |
| Clase negocio | `Parte.cls` | — |

---

### 13. Módulo: Menú Principal

| Componente | Archivo(s) |
| :--- | :--- |
| Menú principal | `Form_Form0BDOpciones.cls` |
| Informes | `Form_Form0BDInformes.cls` |
| Otras opciones | `Form_Form0BDOpcionesOtras.cls` |

### 13b. Módulo: Informes SLA (Spec-010)

| Componente | Archivo(s) | Tabla(s) BD |
| :--- | :--- | :--- |
| Gestor Informes SLA | `Form_FormInformeSLA.cls` | `TbEventos` |

---

### 14. Módulo: Infraestructura (Identidad + Entorno)

**Propósito:** Autenticación, autorización y configuración de entorno

| Componente | Archivo(s) | Tabla(s) BD (externas) |
| :--- | :--- | :--- |
| Inicialización | `Variables Globales.bas` (EVE) | — |
| Factory usuarios | `Constructor.bas` | — |
| Usuario | `Usuario.cls` | `TbUsuariosAplicaciones` (Lanzadera) |
| Permisos | `UsuarioAplicacionPermisos.cls` | `TbUsuariosAplicacionesPermisos` (Lanzadera) |
| Entorno | `Entorno.cls` | — |

**Relaciones:**
- `Lanzadera_Datos.accdb` → `TbUsuariosAplicaciones.CorreoUsuario` → `Usuario`
- `Lanzadera_Datos.accdb` → `TbUsuariosAplicacionesPermisos.IDAplicacion = 6` → permisos BRASS
- `IDAplicacion` BRASS = `"6"`

**Documentación:** [`PRD/14_Infraestructura_Identidad_Entorno.md`](./PRD/14_Infraestructura_Identidad_Entorno.md)

---

## Módulos Comunes / Utilidades

### src/modules/
| Archivo | Propósito |
| :--- | :--- |
| `Variables Globales.bas` | Constantes, variables globales y función EVE |
| `Funciones Generales.bas` | Funciones helper genéricas |
| `Constructor.bas` | Factory de objetos (usuarios, entidades, colecciones) |
| `Módulo1.bas` | (por investigar) |
| `Pruebas.bas` | Funciones de prueba |
| `HTML.bas` | Generación HTML |
| `JsonConverter.bas` | Conversión JSON |

### src/classes/
| Archivo | Propósito |
| :--- | :--- |
| `Usuario.cls` | Gestión de usuario conectado (documentado en PRD-14) |
| `UsuarioAplicacionPermisos.cls` | Permisos y seguridad (documentado en PRD-14) |
| `Entorno.cls` | Configuración de entorno singleton (documentado en PRD-14) |
| `LIbranza.cls` | Gestión de libranzas |
| `TipoDia.cls` | Tipos de día |
| `TecnicoHoras.cls` | Control de horas |

---

## Modelo de Datos (ERD)

Ubicación: [`ERD/Estructura_Datos.md`](./ERD/Estructura_Datos.md)

**Total de tablas:** 56

**Tablas principales:**
- `TbEventos` — Eventos de mantenimiento
- `TbActividades` — Actividades associateadas a eventos
- `TbEquipos` — Catálogo de equipos
- `TbTecnicos` — Técnicos
- `TbFacturaPrincipal` — Cabecera de facturas

**Tablas auxiliares:**
- `TbAux*` — Tablas de importación/temporales
- `TbFactura*` — Detalles de facturación

---

## PRDs Generados

| PRD | Módulo | Estado | Fecha |
| :--- | :--- | :--- | :--- |
| PRD-01 | Menú Principal (Form_Form0BDOpciones) | ✓ Completado | 2026-03-11 |
| PRD-02 | Gestión de Eventos (Form_FormEventoGestion) | ✓ Completado | 2026-03-11 |
| PRD-03 | Gestión de Actividades (Form_FormActividadGestion) | ✓ Completado | 2026-03-11 |
| PRD-04 | Gestión de Técnicos (Form_FormTecnicoGestion) | ✓ Completado | 2026-03-11 |
| PRD-05 | Gestión de Equipos (Form_FormEquipoGestion) | ✓ Completado | 2026-03-11 |
| PRD-06 | Gestión de Materiales (Form_FormMaterialesGestion) | ✓ Completado | 2026-03-11 |
| PRD-07 | Gestión de Gastos (Form_FormGastoGestion) | ✓ Completado | 2026-03-11 |
| PRD-08 | Facturación (Form_FormFacturacionGestion) | ✓ Completado | 2026-03-11 |
| PRD-09 | Planificación (Form_FormPlanificacionGestion) | ✓ Completado | 2026-03-11 |
| PRD-10 | Subcontrataciones (Form_FormSubContratacionGestion) | ✓ Completado | 2026-03-11 |
| PRD-14 | Infraestructura (Identidad + Entorno) | ✓ Completado | 2026-03-11 |

---

## Specs en Desarrollo

| Spec | Módulo | Estado | Descripción |
| :--- | :--- | :--- | :--- |
| `spec-003` | Eventos/SLA | 🔵 Abierta | Formulario Form_FormEventoSLA con validación SLA completa (código en src/, pendiente de integrar con Spec-004/005) |
| `spec-004` | Eventos/SLA | 🔵 Abierta | Integración cmdSLA en Form_FormEventoAlta — **código en src/ listo, botón por añadir en Access** |
| `spec-005` | Eventos/SLA | 🔵 Abierta | Integración cmdSLA en Form_FormEventoEdicion (depende de Spec-003) |
| `spec-006` | Eventos/SLA | 🔵 Abierta | Informe Trimestral SLA (formulario nuevo + cálculos) |
| `spec-007` | Menú Informes | 🔵 Abierta | Acceso a informe SLA desde Form_Form0BDInformes (depende de Spec-006) |
| `spec-008` | Testing | 🔵 Abierta | Módulo de testing SLA (requiere Specs 001-007 completadas) |
| `spec-009` | Eventos/SLA | 🔵 Abierta | Refactor: extraer ValidarCamposSLA a SLAValidator.bas (depende de Spec-003) |
| `spec-010` | Informes SLA | 🟢 Implementada | Gestor de Informes SLA - filtro por fechas, listado eventos franqueados con SLA, detalle y exportación Excel |

## Specs Completadas

| Spec | Módulo | Descripción |
| :--- | :--- | :--- |
| `spec-001` | Eventos/SLA | Campos SLA en TbEventos (migración) |
| `spec-002` | Eventos/SLA | Propiedades SLA en Evento.cls |

**Campos SLA en `TbEventos`:**
- TRES: `FechaRecepcionNotificacion`, `FechaInicioContactoCliente`
- TRCM: `IncidenciaAveriaOReparacion`, `FechaInicioTiempoAdquisicion`, `FechaFinTiempoAdquisicion`, `TipoReparacion`, `Urgente`
- TRSS: `EventoConServicioAfectado`, `FechaRestablecimientoServicio`
- SLA-4: `TipoRepInsitu`, `TipoRepNoSMT`, `TipoRepValvulas`

**Regla vigente de negocio SLA (alineada a pliego):**
- `Urgente` segmenta TRCM, pero no impone un valor fijo en `TipoReparacion`.

---

## Archivos Pendientes de Clasificación

- `docs/PRD/` — material histórico del repo de código; la fuente documental vigente vive fuera del repo
- `C:\00repos\documentacion\OPENSPEC\00_BRASS` — ubicación canonical de documentación y OpenSpec

---

## Próximos Pasos

1. ~~Exportar código VBA~~ ✓
2. ~~Generar ERD~~ ✓
3. ~~Crear DISCOVERY_MAP~~ ✓
4. Crear `DEUDA_TECNICA.md` en el repositorio documental si sigue siendo necesaria
5. Generar/alinear PRDs en `C:\00repos\documentacion\OPENSPEC\00_BRASS`
6. Continuar cambios SDD/OpenSpec desde el repositorio documental, no desde este repo de código
7. Spec-006: Informe Trimestral SLA
8. Spec-007: Menú Informes SLA
9. Spec-009: Refactor SLAValidator
10. Spec-008: Módulo Testing SLA

---

*Generado automáticamente durante autodescubrimiento*
