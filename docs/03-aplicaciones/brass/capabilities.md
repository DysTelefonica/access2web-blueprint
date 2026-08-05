# Brass — capacidades observadas

## Resultado

La aplicación cubre un agregado de **gestión de eventos regulatorios con equipos de medida calibrados**: alta de evento, asignación de técnico, gestión de equipos de medida, calibración, planificación, partes de trabajo, libranzas, subcontratación, materiales, gastos, facturación, franqueo (cierre) con validación de SLA. Es una app con **dominio regulatorio fuerte** (calibración) y **alto volumen de datos** (25639 actividades, 5777 eventos, 7592 anexos). La paridad futura debe incluir como mínimo las capacidades siguientes; ninguna se marca como retirada.

| Dominio | Capacidades evidenciadas | Evidencia principal |
|---|---|---|
| Arranque e identidad | `EVE` (en `Variables Globales.bas:264`) con `getValorPropiedad`, `TempVars("EnDesarrollo")`, `TempVars("DatosEnLocal")` | `src/modules/Variables Globales.bas:264` (EVE) |
| **Contraseña hardcodeada** | ⚠️ `Variables Globales.bas:560`: `wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")` — **D104 CRÍTICO** | `src/modules/Variables Globales.bas:560` |
| **Eventos** | alta, edición, franqueo (cierre) con validación de SLA, validaciones MotivosNoFranqueable | `Evento.cls` (clase principal), `Form_FormEventoAlta.cls`, `Form_FormEventoEdicion.cls`, `Form_FormEventoEqMedida.cls`, `Form_FormEventoEqMedidaEdicion.cls` |
| **Franqueo con SLA** | `Evento.Franquear(p_CausaFin)` valida SLA + MotivosNoFranqueable + actualiza `FechaFinal`, `HORAFINALEVENTO`, `Franqueado=True`, `CAUSAFIN` | `src/classes/Evento.cls:1255-1313` (`Franquear`) |
| **Equipos de medida con calibración** | gestión de equipos calibrados, calibraciones con fechas, EstadoCalibracion | `EquipoMedida.cls`, `EquipoMedidaCalibracion.cls`, `EventoEquipoMedida.cls`, `EventoEquipoMedidaAntes.cls` |
| **Calibración** (dominio regulatorio) | `EquipoMedidaCalibracion` con 10 callers, gestión de calibraciones con fechas | `src/classes/EquipoMedidaCalibracion.cls` (28 calibraciones en `TbEquiposMedidaCalibraciones`) |
| **Actividades** | alta, edición, con vinculación a Evento y a Material | `Actividad.cls`, `Form_FormActividadAlta.cls` (presumido) |
| **Facturación** | sistema completo con 8 tablas (1 principal + 5 detalles + 1 perfiles + 1 conciliación) | `Factura.cls`, `FacturaPerfiles.cls`, `Form_FormFactura*.cls` (presumido) |
| **Facturas con múltiples involucrados** | `TbFacturaActividadesInvolucradas`, `TbFacturaEventosInvolucrados`, `TbFacturaGastosInvolucrados`, `TbFacturaMaterialesInvolucrados`, `TbFacturaSubcontratacionesInvolucradas` (5 tablas de detalle) | `Factura.cls` |
| **Materiales** | alta, seguimiento, con vinculación a Actividad y a Factura | `Material.cls`, `MaterialSeguimiento.cls` |
| **Subcontratación** | gestión de subcontratación con vinculación a Factura | `Subcontratacion.cls` |
| **Técnicos** | gestión de técnicos, ausencias, fiestas, tipos, precios | `Tecnico.cls`, `TecnicoHoras.cls`, `TipoTecnico.cls`, `TipoTecnicoPrecio.cls` |
| **Partes de trabajo** | `TbPartesPpal` (cabecera) + `TbPartesDetalle` (detalle) | `Parte.cls` |
| **Planificación** | 5 tablas para planificación (Ppal, Anexos, Equipos, Registrada, Aux) | `Planificacion.cls` (presumido) |
| **BUI / Nodos / Subsistemas** | jerarquía organizacional de 4 niveles | `TbBUI`, `TbNodos`, `TbNodoBUI`, `TbSubsistemaBui` |
| **Anexos** | alta, con vinculación a Solicitud/Evento | `Anexo.cls` |
| **Catálogos** | `TbCausaFin`, `TbCriticidad`, `TbCodActividad`, `TbTipoAccion`, `TbTipoAsistencia`, `TbTipoEvento`, `TbRepuestosTipo`, `TbOriginador`, `TbResultadoVerificacion`, `TbUbicacion` | entidades de catálogo |
| **Libranzas** (con typo `LIbranza`) | gestión de pagos | `LIbranza.cls` (typo en el nombre) |
| **Fiestas** | gestión de días festivos | `Fiesta.cls` (presumido) |
| **Informes** | `Informe.cls`, `GestorInforme.cls`, `FacturaPerfiles.cls` | gestión de informes |
| **GuiaConciliaciones** | conciliaciones | `TbGuiaConciliaciones` |
| **Configuración** | `TbConfiguracion` (presumido) | config del sistema |
| **BUI / Subsistemas** | jerarquía BUI | `TbBUI`, `TbSubsistemaBui` |
| **Acoplamiento DAO** | `getdb()` con **155 callers** (intermedio) | `src/modules/Variables Globales.bas:535` (`getdb`) |
| **Acoplamiento con Lanzadera** | `getdbLanzadera()` para identidad (igual que el resto del ecosistema) | `src/modules/Variables Globales.bas` |

## Reglas de conservación

- La **identidad se carga UNA vez en `EVE`** y se pasa por `m_ObjUsuarioConectado`. Patrón análogo al resto del ecosistema.
- **IDAplicacion = "6"** (producción). En la nueva plataforma viene de la configuración (D9).
- **El sistema de calibración es dominio regulatorio CRÍTICO**: las calibraciones tienen implicaciones de calidad/cumplimiento. La nueva plataforma debe **preservar este dominio** con su semántica de fechas, estados y relaciones.
- **El SLA (`TIEMPORESPUESTAEVENTO`)** debe **preservarse** como campo calculado o validado. El workflow valida que el evento se resuelva en un tiempo máximo.
- **El franqueo con validaciones SLA + MotivosNoFranqueable** se traduce a un workflow declarativo en la nueva plataforma (similar a D96 de Condor).
- **`IDEvento` Text(50)**: ⚠️ la nueva plataforma debe decidir si mantiene IDs como string (UUID) o los migra a numéricos (BIGSERIAL). Decisión pendiente.
- **3 tipos de reparación mutuamente excluyentes** (`TipoRepInsitu`, `TipoRepNoSMT`, `TipoRepValvulas` como YesNo): migrar a un enum en PostgreSQL.
- **Sistema de facturación con 8 tablas**: 1 principal + 5 detalles + 1 perfiles + 1 conciliación. Migrar a un modelo normalizado en PostgreSQL con FKs explícitas.
- **El sistema de BUI / Nodos / Subsistemas** es una **jerarquía organizacional** que se traduce a un modelo recursivo en PostgreSQL (CTE o `ltree`).

## Evidencia previa

Se han cosechado PRD, Discovery Map, Architecture Overview, ERD, OpenSpec CAP-001..055, UAT y releases antes de inspeccionar staging. CodeGraph-VBA sobre el repo se consultó primero; el inventario Dysflow real se ejecutó después de `setup_project` (con `projectId: 00-brass-00-main-clean` aplicado) + `register_worktree` + `accessPath` + `backendPath` absolutos: 60 tablas, 27 FKs, 5777 eventos, 25639 actividades, 7592 anexos, 2351 materiales, 806 equipos, 14 equipos de medida, 28 calibraciones, 115 técnicos, 624 planificaciones, 53 facturas. Las afirmaciones divergentes entre esos documentos y el código quedan abiertas, no resueltas por intención.