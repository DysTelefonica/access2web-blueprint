# Matriz campo a campo — solicitud y bloques de datos

> Parte de la [matriz campo a campo](../matriz-campos.md). Reglas de mapeo, leyenda y límites de la fuente en ese documento. Commit fijado `c88347e`.

Las tablas que sostienen la solicitud y sus tres bloques de datos, más las dos tablas retiradas que el volcado del snapshot todavía recoge. `tbDatosPCSUB` y `tbHistorialRechazos` no tienen ninguna cita en el commit fijado: son la traza del tipo de solicitud `PC_SUB` retirado (D-CON-1 del README) y no aparecen en el inventario actual.

## tbSolicitudes

12 columnas · 12 con evidencia en `src` · nombre de tabla citado en 12 ficheros de `src`.

| Origen | Tipo origen → objetivo | Clase de dato | Evidencia en `src` | Correspondencia propuesta | Estado |
|---|---|---|---|---|---|
| `tbSolicitudes.idSolicitud` | LONG(4) → INTEGER | — | P `Solicitud.cls:13` | tabla candidata `condor.tbsolicitudes`; columna sin asignar | candidata a nivel de tabla |
| `tbSolicitudes.idExpediente` | LONG(4) → INTEGER | — | P `Solicitud.cls:14` | tabla candidata `condor.tbsolicitudes`; columna sin asignar | candidata a nivel de tabla |
| `tbSolicitudes.tipoSolicitud` | TEXT(20) → VARCHAR(20) | — | P `Solicitud.cls:16` | tabla candidata `condor.tbsolicitudes`; columna sin asignar | candidata a nivel de tabla |
| `tbSolicitudes.codigoSolicitud` | TEXT(50) → VARCHAR(50) | — | P `Solicitud.cls:18` | tabla candidata `condor.tbsolicitudes`; columna sin asignar | candidata a nivel de tabla |
| `tbSolicitudes.idNCAsociada` | LONG(4) → INTEGER | — | P `Solicitud.cls:19` | tabla candidata `condor.tbsolicitudes`; columna sin asignar | candidata a nivel de tabla |
| `tbSolicitudes.idEstadoInterno` | LONG(4) → INTEGER | — | P `Solicitud.cls:20` | tabla candidata `condor.tbsolicitudes`; columna sin asignar | candidata a nivel de tabla |
| `tbSolicitudes.fechaCreacion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | P `Solicitud.cls:21` | tabla candidata `condor.tbsolicitudes`; columna sin asignar | candidata a nivel de tabla |
| `tbSolicitudes.usuarioCreacion` | TEXT(100) → VARCHAR(100) | personal | P `Solicitud.cls:22` | tabla candidata `condor.tbsolicitudes`; columna sin asignar | candidata a nivel de tabla |
| `tbSolicitudes.fechaModificacion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | P `Solicitud.cls:24` | tabla candidata `condor.tbsolicitudes`; columna sin asignar | candidata a nivel de tabla |
| `tbSolicitudes.usuarioModificacion` | TEXT(100) → VARCHAR(100) | personal | P `Solicitud.cls:25` | tabla candidata `condor.tbsolicitudes`; columna sin asignar | candidata a nivel de tabla |
| `tbSolicitudes.revisionCalidadEstado` | TEXT(20) → VARCHAR(20) | — | P `Solicitud.cls:26` | tabla candidata `condor.tbsolicitudes`; columna sin asignar | candidata a nivel de tabla |
| `tbSolicitudes.revisionCalidadComentarios` | MEMO → TEXT | personal | P `Solicitud.cls:27` | tabla candidata `condor.tbsolicitudes`; columna sin asignar | candidata a nivel de tabla |

## tbDatosPC

41 columnas · 41 con evidencia en `src` · nombre de tabla citado en 4 ficheros de `src`.

| Origen | Tipo origen → objetivo | Clase de dato | Evidencia en `src` | Correspondencia propuesta | Estado |
|---|---|---|---|---|---|
| `tbDatosPC.idDatosPC` | LONG(4) → INTEGER | — | P `DatosPC.cls:14` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.idSolicitud` | LONG(4) → INTEGER | — | P `DatosPC.cls:15` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.refContratoInspeccionOficial` | TEXT(100) → VARCHAR(100) | — | P `DatosPC.cls:16` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.refSuministrador` | TEXT(100) → VARCHAR(100) | — | P `DatosPC.cls:17` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.denominacionContrato` | MEMO → TEXT | — | P `DatosPC.cls:18` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.suministradorNombreDir` | MEMO → TEXT | personal | P `DatosPC.cls:19` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.objetoContrato` | MEMO → TEXT | — | P `DatosPC.cls:20` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.descripcionMaterialAfectado` | MEMO → TEXT | personal | P `DatosPC.cls:22` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.numPlanoEspecificacion` | MEMO → TEXT | — | P `DatosPC.cls:23` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.descripcionPropuestaCambio` | MEMO → TEXT | personal | P `DatosPC.cls:24` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.descripcionPropuestaCambioCont` | MEMO → TEXT | personal | P `DatosPC.cls:25` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.motivoCorregirDeficiencias` | BOOLEAN(1) → BOOLEAN | personal | P `DatosPC.cls:27` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.motivoMejorarCapacidad` | BOOLEAN(1) → BOOLEAN | personal | P `DatosPC.cls:28` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.motivoAumentarNacionalizacion` | BOOLEAN(1) → BOOLEAN | personal | P `DatosPC.cls:29` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.motivoMejorarSeguridad` | BOOLEAN(1) → BOOLEAN | personal | P `DatosPC.cls:30` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.motivoMejorarFiabilidad` | BOOLEAN(1) → BOOLEAN | personal | P `DatosPC.cls:31` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.motivoMejorarCosteEficacia` | BOOLEAN(1) → BOOLEAN | personal | P `DatosPC.cls:32` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.motivoOtros` | BOOLEAN(1) → BOOLEAN | personal | P `DatosPC.cls:33` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.motivoOtrosDetalle` | MEMO → TEXT | personal | P `DatosPC.cls:34` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.incidenciaCoste` | TEXT(50) → VARCHAR(50) | — | P `DatosPC.cls:36` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.incidenciaPlazo` | TEXT(50) → VARCHAR(50) | — | P `DatosPC.cls:38` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.incidenciaSeguridad` | BOOLEAN(1) → BOOLEAN | — | P `DatosPC.cls:40` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.incidenciaFiabilidad` | BOOLEAN(1) → BOOLEAN | — | P `DatosPC.cls:41` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.incidenciaMantenibilidad` | BOOLEAN(1) → BOOLEAN | — | P `DatosPC.cls:42` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.incidenciaIntercambiabilidad` | BOOLEAN(1) → BOOLEAN | — | P `DatosPC.cls:43` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.incidenciaVidaUtilAlmacen` | BOOLEAN(1) → BOOLEAN | — | P `DatosPC.cls:44` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.incidenciaFuncionamientoFuncion` | BOOLEAN(1) → BOOLEAN | — | P `DatosPC.cls:45` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.impactoClasificacion` | TEXT(255) → VARCHAR(255) | — | P `DatosPC.cls:47` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.CambioAfectaAMaterial` | TEXT(255) → VARCHAR(255) | — | P `DatosPC.cls:48` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.firmaOficinaTecnicaNombre` | TEXT(100) → VARCHAR(100) | personal | P `DatosPC.cls:52` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.firmaRepSuministradorNombre` | TEXT(100) → VARCHAR(100) | personal | P `DatosPC.cls:53` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.racCodigo` | TEXT(50) → VARCHAR(50) | — | P `DatosPC.cls:55` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.observacionesRAC` | MEMO → TEXT | personal | P `DatosPC.cls:56` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.racNombre` | TEXT(255) → VARCHAR(255) | personal | P `DatosPC.cls:57` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.racDecision` | TEXT(50) → VARCHAR(50) | — | P `DatosPC.cls:58` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.racRechazoMotivos` | MEMO → TEXT | personal | P `DatosPC.cls:59` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.obsAprobacionAutoridadDiseno` | MEMO → TEXT | — | P `DatosPC.cls:60` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.NombreAutoridadDiseno` | TEXT(100) → VARCHAR(100) | personal | P `DatosPC.cls:61` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.decisionFinal` | TEXT(50) → VARCHAR(50) | — | P `DatosPC.cls:63` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.obsDecisionFinal` | MEMO → TEXT | — | P `DatosPC.cls:64` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosPC.NombreFirmanteFinal` | TEXT(100) → VARCHAR(100) | personal | P `DatosPC.cls:65` | tabla candidata `condor.tbdatospc`; columna sin asignar | candidata a nivel de tabla |

## tbDatosCDCA

40 columnas · 40 con evidencia en `src` · nombre de tabla citado en 4 ficheros de `src`.

| Origen | Tipo origen → objetivo | Clase de dato | Evidencia en `src` | Correspondencia propuesta | Estado |
|---|---|---|---|---|---|
| `tbDatosCDCA.idDatosCDCA` | LONG(4) → INTEGER | — | P `DatosCDCA.cls:14` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.idSolicitud` | LONG(4) → INTEGER | — | P `DatosCDCA.cls:15` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.numContrato` | TEXT(100) → VARCHAR(100) | — | P `DatosCDCA.cls:17` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.refSuministrador` | TEXT(100) → VARCHAR(100) | — | P `DatosCDCA.cls:18` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.SuministradorNombreDir` | MEMO → TEXT | personal | P `DatosCDCA.cls:19` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.refDesviacionesPrevias` | TEXT(100) → VARCHAR(100) | — | P `DatosCDCA.cls:20` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.requiereModificacionContrato` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCA.cls:21` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.identificacionMaterial` | MEMO → TEXT | — | P `DatosCDCA.cls:23` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.numPlanoEspecificacion` | MEMO → TEXT | — | P `DatosCDCA.cls:24` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.cantidadPeriodo` | TEXT(50) → VARCHAR(50) | — | P `DatosCDCA.cls:25` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.numSerieLote` | TEXT(100) → VARCHAR(100) | — | P `DatosCDCA.cls:26` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.causaNC` | MEMO → TEXT | — | P `DatosCDCA.cls:27` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.descripcionImpactoNC` | MEMO → TEXT | personal | P `DatosCDCA.cls:28` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.descripcionImpactoNCCont` | MEMO → TEXT | personal | P `DatosCDCA.cls:29` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.afectaPrestaciones` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCA.cls:31` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.afectaSeguridad` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCA.cls:32` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.afectaFiabilidad` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCA.cls:33` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.afectaVidaUtil` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCA.cls:34` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.afectaMedioambiente` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCA.cls:35` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.afectaIntercambiabilidad` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCA.cls:37` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.afectaMantenibilidad` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCA.cls:36` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.afectaApariencia` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCA.cls:38` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.afectaOtros` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCA.cls:39` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.impactoCoste` | TEXT(50) → VARCHAR(50) | — | P `DatosCDCA.cls:40` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.clasificacionNC` | TEXT(50) → VARCHAR(50) | — | P `DatosCDCA.cls:41` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.esSuministradorAD` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCA.cls:42` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.identificacionAutoridadDiseno` | TEXT(100) → VARCHAR(100) | — | P `DatosCDCA.cls:43` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.efectoFechaEntrega` | MEMO → TEXT | — | P `DatosCDCA.cls:44` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.firmaAprobacionRespIngenieriaNombre` | TEXT(255) → VARCHAR(255) | personal | P `DatosCDCA.cls:47` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.firmaAprobacionRespProduccionNombre` | TEXT(255) → VARCHAR(255) | personal | P `DatosCDCA.cls:48` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.firmaAprobacionRespCalidadNombre` | TEXT(255) → VARCHAR(255) | personal | P `DatosCDCA.cls:49` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.firmaAprobacionRespDisenioNombre` | TEXT(255) → VARCHAR(255) | personal | P `DatosCDCA.cls:50` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.firmaAprobacionRepresentanteSumNombre` | TEXT(255) → VARCHAR(255) | personal | P `DatosCDCA.cls:51` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.racCodigo` | TEXT(50) → VARCHAR(50) | — | P `DatosCDCA.cls:54` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.observacionesRAC` | MEMO → TEXT | personal | P `DatosCDCA.cls:55` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.racNombre` | TEXT(255) → VARCHAR(255) | personal | P `DatosCDCA.cls:56` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.racDecision` | TEXT(50) → VARCHAR(50) | — | P `DatosCDCA.cls:57` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.decisionFinal` | TEXT(50) → VARCHAR(50) | — | P `DatosCDCA.cls:59` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.observacionesFinales` | MEMO → TEXT | personal | P `DatosCDCA.cls:60` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCA.NombreFirmanteFinal` | TEXT(100) → VARCHAR(100) | personal | P `DatosCDCA.cls:61` | tabla candidata `condor.tbdatoscdca`; columna sin asignar | candidata a nivel de tabla |

## tbDatosCDCASUB

44 columnas · 44 con evidencia en `src` · nombre de tabla citado en 5 ficheros de `src`.

| Origen | Tipo origen → objetivo | Clase de dato | Evidencia en `src` | Correspondencia propuesta | Estado |
|---|---|---|---|---|---|
| `tbDatosCDCASUB.idDatosCDCASUB` | LONG(4) → INTEGER | — | P `DatosCDCASUB.cls:19` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.idSolicitud` | LONG(4) → INTEGER | — | P `DatosCDCASUB.cls:20` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.refSuministrador` | TEXT(100) → VARCHAR(100) | — | P `DatosCDCASUB.cls:23` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.refSubSuministrador` | TEXT(100) → VARCHAR(100) | — | P `DatosCDCASUB.cls:24` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.suministradorPrincipalNombreDir` | MEMO → TEXT | personal | P `DatosCDCASUB.cls:25` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.subSuministradorNombreDir` | MEMO → TEXT | personal | P `DatosCDCASUB.cls:26` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.refDesviacionesPrevias` | TEXT(100) → VARCHAR(100) | — | P `DatosCDCASUB.cls:27` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.requiereModificacionContrato` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCASUB.cls:28` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.identificacionMaterial` | MEMO → TEXT | — | P `DatosCDCASUB.cls:31` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.numPlanoEspecificacion` | TEXT(100) → VARCHAR(100) | — | P `DatosCDCASUB.cls:32` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.cantidadPeriodo` | TEXT(50) → VARCHAR(50) | — | P `DatosCDCASUB.cls:33` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.numSerieLote` | TEXT(100) → VARCHAR(100) | — | P `DatosCDCASUB.cls:34` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.causaNC` | MEMO → TEXT | — | P `DatosCDCASUB.cls:35` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.descripcionImpactoNC` | MEMO → TEXT | personal | P `DatosCDCASUB.cls:36` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.descripcionImpactoNCCont` | MEMO → TEXT | personal | P `DatosCDCASUB.cls:37` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.afectaPrestaciones` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCASUB.cls:40` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.afectaSeguridad` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCASUB.cls:41` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.afectaFiabilidad` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCASUB.cls:42` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.afectaVidaUtil` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCASUB.cls:43` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.afectaMedioambiente` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCASUB.cls:44` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.afectaIntercambiabilidad` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCASUB.cls:46` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.afectaMantenibilidad` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCASUB.cls:45` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.afectaApariencia` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCASUB.cls:49` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.afectaOtros` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCASUB.cls:50` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.impactoCoste` | TEXT(50) → VARCHAR(50) | — | P `DatosCDCASUB.cls:51` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.clasificacionNC` | TEXT(50) → VARCHAR(50) | — | P `DatosCDCASUB.cls:52` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.esSubSuministradorAD` | BOOLEAN(1) → BOOLEAN | — | P `DatosCDCASUB.cls:55` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.identificacionAutoridadDiseno` | TEXT(100) → VARCHAR(100) | — | P `DatosCDCASUB.cls:54` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.efectoFechaEntrega` | MEMO → TEXT | — | P `DatosCDCASUB.cls:53` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.firmaAprobacionRespIngenieriaNombre` | TEXT(255) → VARCHAR(255) | personal | P `DatosCDCASUB.cls:58` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.firmaAprobacionRespProduccionNombre` | TEXT(255) → VARCHAR(255) | personal | P `DatosCDCASUB.cls:59` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.firmaAprobacionRespCalidadNombre` | TEXT(255) → VARCHAR(255) | personal | P `DatosCDCASUB.cls:60` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.firmaAprobacionRespDisenioNombre` | TEXT(255) → VARCHAR(255) | personal | P `DatosCDCASUB.cls:61` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.firmaAprobacionRepresentanteSumNombre` | TEXT(255) → VARCHAR(255) | personal | P `DatosCDCASUB.cls:62` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.racCodigo` | TEXT(50) → VARCHAR(50) | — | P `DatosCDCASUB.cls:65` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.observacionesRAC` | MEMO → TEXT | personal | P `DatosCDCASUB.cls:66` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.racNombre` | TEXT(255) → VARCHAR(255) | personal | P `DatosCDCASUB.cls:69` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.racDecision` | TEXT(50) → VARCHAR(50) | — | P `DatosCDCASUB.cls:67` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.racRechazoMotivos` | MEMO → TEXT | personal | P `DatosCDCASUB.cls:68` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.observacionesRACDelegador` | MEMO → TEXT | personal | P `DatosCDCASUB.cls:72` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.racNombreDelegador` | TEXT(255) → VARCHAR(255) | personal | P `DatosCDCASUB.cls:73` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.decisionFinal` | TEXT(50) → VARCHAR(50) | — | P `DatosCDCASUB.cls:76` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.observacionesFinales` | MEMO → TEXT | personal | P `DatosCDCASUB.cls:77` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |
| `tbDatosCDCASUB.NombreFirmanteFinal` | TEXT(100) → VARCHAR(100) | personal | P `DatosCDCASUB.cls:78` | tabla candidata `condor.tbdatoscdcasub`; columna sin asignar | candidata a nivel de tabla |

## tbDatosPCSUB

43 columnas · 0 con evidencia en `src` · nombre de tabla citado en 0 ficheros de `src`.

| Origen | Tipo origen → objetivo | Clase de dato | Evidencia en `src` | Correspondencia propuesta | Estado |
|---|---|---|---|---|---|
| `tbDatosPCSUB.idDatosPCSUB` | LONG(4) → INTEGER | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.idSolicitud` | LONG(4) → INTEGER | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.refContratoInspeccionOficial` | TEXT(100) → VARCHAR(100) | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.refSubSuministrador` | TEXT(100) → VARCHAR(100) | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.denominacionContrato` | MEMO → TEXT | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.SubsuministradorNombreDir` | MEMO → TEXT | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.objetoContrato` | MEMO → TEXT | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.descripcionMaterialAfectado` | MEMO → TEXT | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.numPlanoEspecificacion` | MEMO → TEXT | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.descripcionPropuestaCambio` | MEMO → TEXT | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.descripcionPropuestaCambioCont` | MEMO → TEXT | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.motivoCorregirDeficiencias` | BOOLEAN(1) → BOOLEAN | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.motivoMejorarCapacidad` | BOOLEAN(1) → BOOLEAN | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.motivoAumentarNacionalizacion` | BOOLEAN(1) → BOOLEAN | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.motivoMejorarSeguridad` | BOOLEAN(1) → BOOLEAN | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.motivoMejorarFiabilidad` | BOOLEAN(1) → BOOLEAN | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.motivoMejorarCosteEficacia` | BOOLEAN(1) → BOOLEAN | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.motivoOtros` | BOOLEAN(1) → BOOLEAN | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.motivoOtrosDetalle` | MEMO → TEXT | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.incidenciaCoste` | TEXT(50) → VARCHAR(50) | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.incidenciaPlazo` | TEXT(50) → VARCHAR(50) | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.incidenciaSeguridad` | BOOLEAN(1) → BOOLEAN | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.incidenciaFiabilidad` | BOOLEAN(1) → BOOLEAN | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.incidenciaMantenibilidad` | BOOLEAN(1) → BOOLEAN | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.incidenciaIntercambiabilidad` | BOOLEAN(1) → BOOLEAN | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.incidenciaVidaUtilAlmacen` | BOOLEAN(1) → BOOLEAN | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.incidenciaFuncionamientoFuncion` | BOOLEAN(1) → BOOLEAN | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.impactoClasificacion` | TEXT(255) → VARCHAR(255) | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.CambioAfectaAMaterial` | TEXT(255) → VARCHAR(255) | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.firmaOficinaTecnicaSubSuministradorNombre` | TEXT(100) → VARCHAR(100) | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.firmaRepSubSuministradorNombre` | TEXT(100) → VARCHAR(100) | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.racCodigo` | TEXT(50) → VARCHAR(50) | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.observacionesRAC` | MEMO → TEXT | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.racNombre` | TEXT(255) → VARCHAR(255) | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.racDecision` | TEXT(50) → VARCHAR(50) | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.racRechazoMotivos` | MEMO → TEXT | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.observacionesRACDelegador` | MEMO → TEXT | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.racNombreDelegador` | TEXT(255) → VARCHAR(255) | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.obsAprobacionAutoridadDiseno` | MEMO → TEXT | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.NombreAutoridadDiseno` | TEXT(100) → VARCHAR(100) | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.decisionFinal` | TEXT(50) → VARCHAR(50) | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.obsDecisionFinal` | MEMO → TEXT | — | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbDatosPCSUB.NombreFirmanteFinal` | TEXT(100) → VARCHAR(100) | personal | — | tabla candidata `condor.tbdatospcsub`; columna sin asignar | retirada, sin uso en el commit fijado |

## tbHistorialRechazos

8 columnas · 0 con evidencia en `src` · nombre de tabla citado en 0 ficheros de `src`.

| Origen | Tipo origen → objetivo | Clase de dato | Evidencia en `src` | Correspondencia propuesta | Estado |
|---|---|---|---|---|---|
| `tbHistorialRechazos.ID` | LONG(4) → INTEGER | — | — | tabla candidata `condor.tbhistorialrechazos`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbHistorialRechazos.IdSolicitud` | LONG(4) → INTEGER | — | — | tabla candidata `condor.tbhistorialrechazos`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbHistorialRechazos.FechaRechazo` | SHORT_DATE_TIME(8) → TIMESTAMP | — | — | tabla candidata `condor.tbhistorialrechazos`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbHistorialRechazos.UsuarioCalidad` | TEXT(255) → VARCHAR(255) | personal | — | tabla candidata `condor.tbhistorialrechazos`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbHistorialRechazos.AreaAfectada` | TEXT(255) → VARCHAR(255) | — | — | tabla candidata `condor.tbhistorialrechazos`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbHistorialRechazos.MotivoPrincipal` | MEMO → TEXT | personal | — | tabla candidata `condor.tbhistorialrechazos`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbHistorialRechazos.Comentarios` | MEMO → TEXT | personal | — | tabla candidata `condor.tbhistorialrechazos`; columna sin asignar | retirada, sin uso en el commit fijado |
| `tbHistorialRechazos.EstaResuelto` | BOOLEAN(1) → BOOLEAN | — | — | tabla candidata `condor.tbhistorialrechazos`; columna sin asignar | retirada, sin uso en el commit fijado |
