# Matriz campo a campo — planificación y partes

> Parte de la [matriz campo a campo](../matriz-campos.md). Reglas de mapeo, leyenda y límites de la fuente en ese documento. Commit fijado `3f77d58`.

## TbAuxManteniminetosPreventivosCalendario

9 columnas · 93 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbAuxManteniminetosPreventivosCalendario.BUI` | TEXT(50) → VARCHAR(50) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxManteniminetosPreventivosCalendario.IDEQUIPO` | LONG(4) → INTEGER | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxManteniminetosPreventivosCalendario.NOMBREEQUIPO` | TEXT(255) → VARCHAR(255) | personal | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxManteniminetosPreventivosCalendario.MESPROG` | TEXT(50) → VARCHAR(50) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxManteniminetosPreventivosCalendario.SEMANAPROG` | TEXT(50) → VARCHAR(50) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxManteniminetosPreventivosCalendario.ANIO` | TEXT(50) → VARCHAR(50) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxManteniminetosPreventivosCalendario.FECHASREPARACIONENANIO` | MEMO → TEXT | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxManteniminetosPreventivosCalendario.IDEVENTO` | MEMO → TEXT | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxManteniminetosPreventivosCalendario.PARTESMTOREALIZADOENANIO` | MEMO → TEXT | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbAuxPlanificacion

8 columnas · 81 filas en la extracción del snapshot · citada en 1 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbAuxPlanificacion.IDPlanificacion` | LONG(4) → INTEGER | — | F `Form_FormPlanificacionGestion.cls:589` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxPlanificacion.Equipo` | TEXT(255) → VARCHAR(255) | — | F `Form_FormPlanificacionGestion.cls:590` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxPlanificacion.FechaPrevista` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormPlanificacionGestion.cls:616` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxPlanificacion.DiasRestan` | LONG(4) → INTEGER | — | F `Form_FormPlanificacionGestion.cls:617` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxPlanificacion.BUI` | TEXT(255) → VARCHAR(255) | — | F `Form_FormPlanificacionGestion.cls:594` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxPlanificacion.SUBSISTEMA` | TEXT(255) → VARCHAR(255) | — | F `Form_FormPlanificacionGestion.cls:595` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxPlanificacion.CadenaIDEventosPosibles` | MEMO → TEXT | — | F `Form_FormPlanificacionGestion.cls:620` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxPlanificacion.CadenaIDEventosUsados` | MEMO → TEXT | — | F `Form_FormPlanificacionGestion.cls:621` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbPartesDetalle

14 columnas · 8819 filas en la extracción del snapshot · citada en 3 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbPartesDetalle.IDParte` | TEXT(50) → VARCHAR(50) | — | F `Parte.cls:345` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesDetalle.IDEvento` | TEXT(50) → VARCHAR(50) | — | F `Constructor.bas:288` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesDetalle.FechaAlta` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Parte.cls:377` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesDetalle.HoraAlta` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Parte.cls:378` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesDetalle.Descripcion` | MEMO → TEXT | — | F `Parte.cls:350` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesDetalle.CausaFin` | MEMO → TEXT | — | F `Parte.cls:380` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesDetalle.FechaFin` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Parte.cls:381` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesDetalle.HoraFin` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Parte.cls:382` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesDetalle.AliasCreador` | TEXT(50) → VARCHAR(50) | personal | F `Parte.cls:383` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesDetalle.Originador` | TEXT(255) → VARCHAR(255) | — | F `Parte.cls:384` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesDetalle.TiempoRespuesta` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Parte.cls:385` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesDetalle.SubSistema` | TEXT(255) → VARCHAR(255) | — | F `Parte.cls:386` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesDetalle.Equipo` | TEXT(255) → VARCHAR(255) | — | F `Parte.cls:388` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesDetalle.PMPR` | TEXT(50) → VARCHAR(50) | — | F `Parte.cls:391` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbPartesPpal

8 columnas · 342 filas en la extracción del snapshot · citada en 3 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbPartesPpal.IDParte` | TEXT(50) → VARCHAR(50) | — | F `Parte.cls:345` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesPpal.FechaObtencion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Parte.cls:346` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesPpal.FechaInicial` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Parte.cls:347` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesPpal.FechaFinal` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Constructor.bas:4559` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesPpal.Descripcion` | MEMO → TEXT | — | F `Parte.cls:350` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesPpal.NombreArchivoAdjunto` | TEXT(255) → VARCHAR(255) | personal | F `Form_FormEventoGestion.cls:2659` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesPpal.IDFactura` | LONG(4) → INTEGER | — | F `Form_FormEventoGestion.cls:1241` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPartesPpal.Centro` | TEXT(255) → VARCHAR(255) | — | F `Parte.cls:352` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbPlanificacion

13 columnas · 624 filas en la extracción del snapshot · citada en 8 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbPlanificacion.IDPlanificacion` | LONG(4) → INTEGER | — | F `Form_FormPlanificacionGestion.cls:589` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacion.FechaRegistro` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Generales.bas:2374` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacion.IDEquipo` | LONG(4) → INTEGER | — | F `Form_FormEventoEdicion.cls:1133` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacion.ANIO` | LONG(4) → INTEGER | — | F `Form_FormPlanificacionCierre.cls:68` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacion.Mes` | INT(2) → SMALLINT | — | F `Form_FormPlanificacionCierre.cls:69` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacion.Semana` | INT(2) → SMALLINT | — | F `Form_FormPlanificacionCierre.cls:70` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacion.MotivoCierre` | TEXT(50) → VARCHAR(50) | — | F `Form_FormPlanificacionGestion.cls:870` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacion.IDEvento` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoEdicion.cls:1128` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacion.IDNoEvento` | LONG(4) → INTEGER | — | F `Form_FormPlanificacionCierre.cls:72` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacion.FechaPrevistaCierre` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormPlanificacionCierre.cls:74` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacion.FechaCierre` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormPlanificacionCambioIDEvento.cls:258` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacion.MotivacionReprogramacion` | MEMO → TEXT | — | F `Form_FormPlanificacionGestion.cls:868` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacion.IDNuevaPlanificacion` | LONG(4) → INTEGER | — | F `Form_FormPlanificacionGestion.cls:869` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbPlanificacionAnexos

6 columnas · 5 filas en la extracción del snapshot · citada en 2 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbPlanificacionAnexos.IDAnexoPlanificacion` | LONG(4) → INTEGER | — | F `Funciones Generales.bas:2501` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacionAnexos.IDPlanificacion` | LONG(4) → INTEGER | — | F `Funciones Generales.bas:2373` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacionAnexos.NombreAnexo` | TEXT(255) → VARCHAR(255) | personal | F `Funciones Generales.bas:2503` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacionAnexos.UsuarioAnexa` | TEXT(50) → VARCHAR(50) | personal | F `Funciones Generales.bas:2504` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacionAnexos.FechaAnexo` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Generales.bas:2505` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacionAnexos.URLAnexo` | TEXT(255) → VARCHAR(255) | — | F `Form_FormPlanificacionAnexos.cls:315` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbPlanificacionEquipos

5 columnas · 82 filas en la extracción del snapshot · citada en 6 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbPlanificacionEquipos.IDEquipo` | LONG(4) → INTEGER | — | F `Form_FormPlanificacionEquipos.cls:98` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacionEquipos.Alias` | TEXT(50) → VARCHAR(50) | personal | F `Form_FormPlanificacionEdicionEquipo.cls:225` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacionEquipos.PeriodicidadEnMesesRecomendada` | INT(2) → SMALLINT | — | F `Form_FormPlanificacionEdicionEquipo.cls:231` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacionEquipos.FechaAltaParaPlanificacion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Generales.bas:2826` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacionEquipos.FechabajaParaPlanificacion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormPlanificacionEquipos.cls:783` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbPlanificacionRegistrada

5 columnas · 23 filas en la extracción del snapshot · citada en 2 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbPlanificacionRegistrada.ID` | LONG(4) → INTEGER | — | S `Form_FormPlanificacionCierre.cls:122` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacionRegistrada.IDEquipo` | LONG(4) → INTEGER | — | F `Form_FormPlanificacionCierre.cls:75` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacionRegistrada.FechaMantoRealizado` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Generales.bas:3329` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacionRegistrada.Observaciones` | MEMO → TEXT | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbPlanificacionRegistrada.NombreAnexo` | TEXT(255) → VARCHAR(255) | personal | F `Funciones Generales.bas:2503` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
