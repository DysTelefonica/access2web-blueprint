# Capacidades — Brass

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Issue:** #1078 · **Fase:** 5 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Clasificación de las capacidades observadas en `src` frente al estudio previo (`../capabilities.md`, DRAFT). Una capacidad «sin cambios» significa que los objetos de lógica coinciden con el snapshot tras normalizar; no valida que el estudio previo la describiera bien. Los ids `CAP-n` son provisionales de este PRD y no sustituyen a las fichas `CAP-*` previas.

## Clasificación

| ID | Capacidad | Formularios | Objetos de lógica | Frente al estudio previo | Evidencia |
|---|---|---|---|---|---|
| CAP-1 | Arranque, identidad y permisos | — (sin formulario de login propio; `EVE` en el arranque) | `Variables Globales.bas`, `Constructor.bas`, `Usuario.cls`, `UsuarioAplicacionPermisos.cls`, `Entorno.cls` | modificada | `EVE` comprueba al arrancar que el usuario es administrador o técnico y, si no, aborta con «El usuario no tiene permiso para abrir la aplicación» (`Variables Globales.bas:335-347`); ese bloque no existe en el snapshot (diferencia mecánica del README) |
| CAP-2 | Opciones y configuración de la BD | 3 (Form0BDInformes, Form0BDOpciones, Form0BDOpcionesOtras) | `Funciones Generales.bas`, `Entorno.cls` | sin cambios | coincidencia mecánica con el snapshot (README) |
| CAP-3 | Eventos (alta, edición, copia, gestión, SLA, franqueo, alta masiva) | 8 (FormEventoAlta, FormEventoCopia, FormEventoEdicion, FormEventoFranqueo, FormEventoGestion, FormEventoSLA, FormEventoTipoGestion, FormEventosAltaMasivaPorNS) | `Evento.cls`, `EventoBusqueda.cls`, `Constructor.bas`, módulos SLA | sin cambios | `Evento.Franquear` en `Evento.cls:1255`; texto idéntico al del snapshot |
| CAP-4 | Actividades y tipos de actividad | 7 (FormActividadAlta, FormActividadDeEventoPorLotes, FormActividadEdicion, FormActividadGestion, FormActividadTipoAlta, FormActividadTipoEdicion, FormActividadTipoGestion) | `Actividad.cls` | sin cambios | coincidencia mecánica con el snapshot (README) |
| CAP-5 | Equipos, equipos de medida y calibraciones | 15 (FormCalibracionAlta, FormCalibracionEdicion, FormCalibracionEquipoAlta, FormCalibracionEquipoEdicion, FormCalibracionEquipoGestion, FormCalibracionGestion, FormEquipoAlta, FormEquipoEdicion, FormEquipoGestion, FormEquipoMedida, FormEquipoMedidaCalibracion, FormEquipoMedidaCalibracionesGestion, FormEquiposMedidaGestion, FormEventoEqMedida, FormEventoEqMedidaEdicion) | `Equipo.cls`, `EquipoMedida.cls`, `EquipoMedidaCalibracion.cls`, `EventoEquipoMedida.cls`, `EventoEquipoMedidaAntes.cls` | sin cambios | coincidencia mecánica con el snapshot (README) |
| CAP-6 | Planificación y partes | 12 (FormParteGestion, FormPlanificacionAlta, FormPlanificacionAltaEquipo, FormPlanificacionAnexos, FormPlanificacionCalendarios, FormPlanificacionCambioIDEvento, FormPlanificacionCierre, FormPlanificacionEdicionEquipo, FormPlanificacionEquipos, FormPlanificacionGestion, FormPlanificacionReprogramacion, FormPlanificacionesBusqueda) | `Parte.cls`, `Constructor.bas` (no existe `Planificacion.cls`) | sin cambios (clase presumida del estudio previo inexistente) | coincidencia mecánica con el snapshot (README) |
| CAP-7 | Materiales, reparaciones y repuestos | 9 (FormMaterialAltaReparacion, FormMaterialEleccion, FormMaterialIntervencionAlta, FormMaterialIntervencionEdicion, FormMaterialReparacionEdicion, FormMaterialRepuestoAlta, FormMaterialRepuestoEdicion, FormMaterialSeguimiento, FormMaterialesGestion) | `Material.cls`, `MaterialSeguimiento.cls` | sin cambios | coincidencia mecánica con el snapshot (README) |
| CAP-8 | Facturación, gastos y subcontratación | 9 (FormFacturaAlta, FormFacturaResultado, FormFacturacionGestion, FormGastoAlta, FormGastoEdicion, FormGastoGestion, FormSubContratacionAlta, FormSubContratacionEdicion, FormSubContratacionGestion) | `Factura.cls`, `FacturaPerfiles.cls`, `Gasto.cls`, `Subcontratacion.cls`, `LIbranza.cls` | sin cambios | coincidencia mecánica con el snapshot (README) |
| CAP-9 | Técnicos, ausencias, fiestas y precios | 10 (FormTecnico, FormTecnicoAusencias, FormTecnicoAusenciasGestion, FormTecnicoConsultaHoras, FormTecnicoFiestas, FormTecnicoFiestasGestion, FormTecnicoGestion, FormTecnicoTipo, FormTipoTecnicoPrecio, FormTipoTecnicoPreciosGestion) | `Tecnico.cls`, `TecnicoHoras.cls`, `TipoTecnico.cls`, `TipoTecnicoPrecio.cls`, `TipoDia.cls`, `Fiesta.cls` | sin cambios | coincidencia mecánica con el snapshot (README) |
| CAP-10 | Informes (SLA, RAC, desde lista, por factura) | 4 (FormInformeSLA, FormInformesDesdeLista, FormInformesParaRAC, FormInformesPorFactura) | `Informe.cls`, `GestorInforme.cls`, `SLAReportService.bas`, `SLAHTMLService.bas`, `SLAValidator.bas` | sin cambios | coincidencia mecánica con el snapshot (README) |
| CAP-11 | Anexos | 2 (FormAnexos, FormAnexosNoAlcanzables) | `Anexo.cls` | sin cambios | coincidencia mecánica con el snapshot (README) |
| CAP-12 | Estructura y catálogos (BUI, nodos, subsistemas, ubicación, originador) | 8 (FormBUIGestion, FormNodoBUIGestion, FormNodoGestion, FormOriginadorGestion, FormSubSistemaAlta, FormSubSistemaEdicion, FormSubSistemaGestion, FormUbicacionGestion) | `Constructor.bas`, `Funciones Generales.bas` | sin cambios | coincidencia mecánica con el snapshot (README) |

Formularios asignados: 87 de 87; sin asignar: 0.

## Discrepancias con el estudio previo

- **`Planificacion.cls`** y otras clases marcadas «presumido» en el estudio previo: `src/classes` no contiene `Planificacion.cls`; la planificación vive en los formularios `FormPlanificacion*` y en `Constructor.bas`. Estado: a verificar con el propietario si el estudio se refería a otro objeto.
- **Llamadas a `getdb`:** el estudio previo cuenta 155; la medición de este PRD es de unas 180 líneas de producción con `getdb(` (sin pruebas), con otro criterio de recuento. No es una contradicción de comportamiento.
- **Cita de contraseña:** el estudio previo reproduce la contraseña literal de `getdb`; este PRD no la transcribe y registra su existencia en [coexistencia](coexistencia.md).
- **Franqueo:** `Evento.Franquear` sigue en `Evento.cls:1255`, como indica el estudio previo.

## Pendientes

- Capacidades de facturación y calibración: el estudio previo las marca críticas por dominio regulatorio; este PRD no valida la regla regulatoria, solo su implementación en código.
