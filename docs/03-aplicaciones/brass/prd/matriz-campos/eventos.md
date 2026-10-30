# Matriz campo a campo — eventos, actividades y anexos

> Parte de la [matriz campo a campo](../matriz-campos.md). Reglas de mapeo, leyenda y límites de la fuente en ese documento. Commit fijado `3f77d58`.

## TbActividades

14 columnas · 25639 filas en la extracción del snapshot · citada en 23 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbActividades.IDActividad` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoCopia.cls:196` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbActividades.IDEvento` | TEXT(50) → VARCHAR(50) | — | F `Form_FormActividadAlta.cls:134` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbActividades.ALIASTECNICO` | TEXT(255) → VARCHAR(255) | personal | F `Form_FormActividadAlta.cls:135` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbActividades.Actividad` | TEXT(50) → VARCHAR(50) | — | F `Form_FormActividadAlta.cls:140` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbActividades.FechaAlta` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormActividadAlta.cls:137` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbActividades.HorasLaborables` | DOUBLE(8) → DOUBLE PRECISION | — | F `Form_FormActividadAlta.cls:138` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbActividades.HorasExtras` | DOUBLE(8) → DOUBLE PRECISION | — | F `Form_FormActividadAlta.cls:139` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbActividades.Ubicacion` | MEMO → TEXT | — | F `Form_FormActividadAlta.cls:136` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbActividades.DESCRIPCION` | MEMO → TEXT | — | F `Form_FormActividadAlta.cls:142` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbActividades.IDEVENTOGENERADO` | TEXT(50) → VARCHAR(50) | — | F `Form_FormActividadAlta.cls:141` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbActividades.IDExportacion` | TEXT(50) → VARCHAR(50) | — | F `Funciones Generales.bas:5954` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbActividades.CodImportacion` | TEXT(50) → VARCHAR(50) | — | S `Form_FormActividadGestion.cls:58` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbActividades.IDFacturacion` | LONG(4) → INTEGER | — | S `Form_FormActividadGestion.cls:58` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbActividades.Observaciones` | MEMO → TEXT | — | S `Form_FormMaterialRepuestoAlta.cls:28` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbAnexos

10 columnas · 7583 filas en la extracción del snapshot · citada en 6 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbAnexos.IDAnexo` | LONG(4) → INTEGER | — | F `Form_FormEventoCopia.cls:372` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAnexos.IDEvento` | TEXT(50) → VARCHAR(50) | — | F `Constructor.bas:288` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAnexos.IDActividad` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoCopia.cls:196` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAnexos.IDMaterial` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoCopia.cls:241` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAnexos.IDMaterialSeguimiento` | TEXT(50) → VARCHAR(50) | — | F `Anexo.cls:731` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAnexos.IDSubContratacion` | LONG(4) → INTEGER | — | F `Anexo.cls:734` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAnexos.IDGasto` | LONG(4) → INTEGER | — | F `Anexo.cls:737` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAnexos.Titulo` | TEXT(255) → VARCHAR(255) | — | F `Form_FormEventoCopia.cls:365` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAnexos.Descripcion` | MEMO → TEXT | — | F `Form_FormEventoCopia.cls:254` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAnexos.NombreArchivo` | TEXT(255) → VARCHAR(255) | personal | F `Form_FormEventoCopia.cls:367` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbAuxActividad

14 columnas · 0 filas en la extracción del snapshot · citada en 1 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbAuxActividad.NODO` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3438` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxActividad.BUI` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3439` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxActividad.SUBSISTEMA` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3440` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxActividad.IDEQUIPO` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:2375` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxActividad.IDEVENTO` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3355` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxActividad.TIPOTECNICO` | TEXT(255) → VARCHAR(255) | personal | F `Funciones Generales.bas:1991` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxActividad.ALIASTECNICO` | TEXT(255) → VARCHAR(255) | personal | F `Funciones Generales.bas:3444` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxActividad.TIPOACTIVIDAD` | TEXT(255) → VARCHAR(255) | — | S `Funciones Generales.bas:3428` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxActividad.FECHAACTIVIDAD` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3446` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxActividad.HORASLABORABLES` | TEXT(50) → VARCHAR(50) | — | F `Funciones Generales.bas:3447` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxActividad.HORASEXTRAS` | TEXT(50) → VARCHAR(50) | — | F `Funciones Generales.bas:3448` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxActividad.UBICACION` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3449` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxActividad.DESCRIPCION` | MEMO → TEXT | — | F `Funciones Generales.bas:3450` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxActividad.CODREGISTRODIFERENCIA` | MEMO → TEXT | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbAuxEventos

20 columnas · 0 filas en la extracción del snapshot · citada en 1 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbAuxEventos.NODO` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3438` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.BUI` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3439` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.SUBSISTEMA` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3440` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.EQUIPO` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3441` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.IDEVENTO` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3355` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.PMPR` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3727` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.TIPOEVENTO` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3728` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.CRITICIDAD` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3729` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.TIPOTECNICO` | TEXT(255) → VARCHAR(255) | personal | F `Funciones Generales.bas:1991` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.ORIGINADOR` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3731` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.FECHAALTA` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Generales.bas:826` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.HORAALTA` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Generales.bas:3733` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.TIEMPORESPUESTA` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Generales.bas:3734` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.FECHAFIN` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Generales.bas:3735` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.HORAFIN` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Generales.bas:3736` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.CAUSAFIN` | MEMO → TEXT | — | F `Funciones Generales.bas:3737` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.CONTACTO` | TEXT(255) → VARCHAR(255) | personal | F `Funciones Generales.bas:3738` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.IDEVENTOGENERADO` | MEMO → TEXT | — | F `Funciones Generales.bas:3739` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.DESCRIPCION` | MEMO → TEXT | — | F `Funciones Generales.bas:3450` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxEventos.CODREGISTRODIFERENCIA` | MEMO → TEXT | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbBuiIDEvento

2 columnas · 11 filas en la extracción del snapshot · citada en 1 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbBuiIDEvento.BUI` | TEXT(50) → VARCHAR(50) | — | F `Funciones Generales.bas:3439` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbBuiIDEvento.CODBUIIDVENTO` | TEXT(50) → VARCHAR(50) | — | S `Funciones Generales.bas:6205` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbCausaFin

1 columnas · 4 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbCausaFin.CAUSAFIN` | TEXT(50) → VARCHAR(50) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbCodActividad

2 columnas · 12 filas en la extracción del snapshot · citada en 7 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbCodActividad.CODIGO` | TEXT(50) → VARCHAR(50) | — | F `Form_FormActividadTipoAlta.cls:125` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbCodActividad.ACTIVIDADNOPROGRAMADA` | TEXT(255) → VARCHAR(255) | — | F `Form_FormActividadTipoAlta.cls:126` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbCriticidad

2 columnas · 3 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbCriticidad.CRITICIDAD` | INT(2) → SMALLINT | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbCriticidad.DESCRIPCION` | MEMO → TEXT | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbEventos

38 columnas · 5773 filas en la extracción del snapshot · citada en 38 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbEventos.IDEvento` | TEXT(50) → VARCHAR(50) | — | F `Form_FormActividadEdicion.cls:439` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.NODO` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoAlta.cls:705` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.BUI` | TEXT(50) → VARCHAR(50) | — | F `Form_FormBUIGestion.cls:212` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.SUBSISTEMA` | TEXT(255) → VARCHAR(255) | — | F `Form_FormEventoAlta.cls:441` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.IDEquipo` | LONG(4) → INTEGER | — | F `Form_FormEventoAlta.cls:452` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.PMPR` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoAlta.cls:442` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.TIPOEVENTO` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoAlta.cls:443` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.CRITICIDAD` | INT(2) → SMALLINT | — | F `Form_FormEventoAlta.cls:444` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.ALIASTECNICO` | TEXT(50) → VARCHAR(50) | personal | F `Form_FormActividadEdicion.cls:440` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.ORIGINADOR` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoAlta.cls:446` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.FECHAALTAEVENTO` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormEventoAlta.cls:448` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.HORAINICIALEVENTO` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormEventoAlta.cls:449` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.FechaFinal` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Constructor.bas:4559` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.HORAFINALEVENTO` | SHORT_DATE_TIME(8) → TIMESTAMP | — | S `Form_FormEventoCopia.cls:167` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.TIEMPORESPUESTAEVENTO` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormEventoAlta.cls:450` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.CAUSAFIN` | MEMO → TEXT | — | F `Funciones Generales.bas:3737` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.DESCRIPCION` | MEMO → TEXT | — | F `Form_FormActividadEdicion.cls:447` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.CONTACTO` | TEXT(255) → VARCHAR(255) | personal | F `Form_FormEventoAlta.cls:451` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.FechaRegistroAlta` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormEventoAlta.cls:767` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.FechaRegistroModificacion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormEventoAlta.cls:768` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.Franqueado` | BOOLEAN(1) → BOOLEAN | — | F `Form_FormEventoEdicion.cls:688` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.IDExportacion` | TEXT(50) → VARCHAR(50) | — | F `Funciones Generales.bas:5954` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.CodImportacion` | TEXT(50) → VARCHAR(50) | — | S `Form_FormActividadGestion.cls:58` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.IDParte` | TEXT(255) → VARCHAR(255) | — | F `Parte.cls:345` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.Notas` | MEMO → TEXT | — | F `Form_FormEventoAlta.cls:453` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.FechaEnInformeRAC` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormEventoEdicion.cls:690` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.FechaRecepcionNotificacion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormEventoAlta.cls:729` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.FechaInicioContactoCliente` | SHORT_DATE_TIME(8) → TIMESTAMP | personal | F `Form_FormEventoAlta.cls:735` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.IncidenciaAveriaOReparacion` | BOOLEAN(1) → BOOLEAN | — | F `Form_FormEventoAlta.cls:742` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.FechaInicioTiempoAdquisicion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormEventoSLA.cls:134` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.FechaFinTiempoAdquisicion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormEventoSLA.cls:135` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.TipoReparacion` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoSLA.cls:136` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.Urgente` | BOOLEAN(1) → BOOLEAN | — | F `Form_FormEventoAlta.cls:751` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.EventoConServicioAfectado` | BOOLEAN(1) → BOOLEAN | — | F `Form_FormEventoAlta.cls:760` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.FechaRestablecimientoServicio` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormEventoSLA.cls:167` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.TipoRepInsitu` | BOOLEAN(1) → BOOLEAN | — | F `Form_FormEventoSLA.cls:180` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.TipoRepNoSMT` | BOOLEAN(1) → BOOLEAN | — | F `Form_FormEventoSLA.cls:183` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventos.TipoRepValvulas` | BOOLEAN(1) → BOOLEAN | — | F `Form_FormEventoSLA.cls:186` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbEventosEquipoMedida

4 columnas · 302 filas en la extracción del snapshot · citada en 4 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbEventosEquipoMedida.IDEventoEquipoMedida` | LONG(4) → INTEGER | — | F `Evento.cls:1908` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventosEquipoMedida.IDEvento` | TEXT(255) → VARCHAR(255) | — | F `Constructor.bas:288` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventosEquipoMedida.IDEquipoMedida` | LONG(4) → INTEGER | — | F `EquipoMedida.cls:377` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventosEquipoMedida.IDCalibracion` | LONG(4) → INTEGER | — | F `Evento.cls:1911` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbEventosEquipoMedida_antes

3 columnas · 26 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbEventosEquipoMedida_antes.IDEventoEquipoMedida` | LONG(4) → INTEGER | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventosEquipoMedida_antes.IDEvento` | TEXT(255) → VARCHAR(255) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEventosEquipoMedida_antes.IDEquipoMedida` | LONG(4) → INTEGER | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbOriginador

1 columnas · 7 filas en la extracción del snapshot · citada en 1 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbOriginador.ORIGINADOR` | TEXT(50) → VARCHAR(50) | — | F `Form_FormOriginadorGestion.cls:139` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbResultadoVerificacion

1 columnas · 3 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbResultadoVerificacion.RESULTADOVERIFICACION` | TEXT(50) → VARCHAR(50) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbTipoAccion

1 columnas · 8 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbTipoAccion.ESTADOREPARABLE` | TEXT(255) → VARCHAR(255) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbTipoAsistencia

1 columnas · 4 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbTipoAsistencia.TIPOASISTENCIA` | TEXT(255) → VARCHAR(255) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbTipoEvento

1 columnas · 6 filas en la extracción del snapshot · citada en 1 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbTipoEvento.TIPOEVENTO` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoTipoGestion.cls:142` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
