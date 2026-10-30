# Matriz campo a campo — catálogos, adjuntos y trazabilidad

> Parte de la [matriz campo a campo](../matriz-campos.md). Reglas de mapeo, leyenda y límites de la fuente en ese documento. Commit fijado `c88347e`.

Las tablas de catálogo, de adjuntos y de auditoría. Todas salvo `idRechazo` e `idTransicion` tienen su columna como campo público de la entidad correspondiente; esas dos se escriben solo por SQL y no tienen campo en ninguna entidad (hoja de trazabilidad).

## tbEstados

6 columnas · 6 con evidencia en `src` · nombre de tabla citado en 2 ficheros de `src`.

| Origen | Tipo origen → objetivo | Clase de dato | Evidencia en `src` | Correspondencia propuesta | Estado |
|---|---|---|---|---|---|
| `tbEstados.idEstado` | LONG(4) → INTEGER | — | P `Estado.cls:14` | tabla candidata `condor.tbestados`; columna sin asignar | candidata a nivel de tabla |
| `tbEstados.nombreEstado` | TEXT(50) → VARCHAR(50) | personal | P `Estado.cls:15` | tabla candidata `condor.tbestados`; columna sin asignar | candidata a nivel de tabla |
| `tbEstados.descripcion` | TEXT(255) → VARCHAR(255) | personal | P `Estado.cls:16` | tabla candidata `condor.tbestados`; columna sin asignar | candidata a nivel de tabla |
| `tbEstados.esEstadoInicial` | BOOLEAN(1) → BOOLEAN | — | P `Estado.cls:17` | tabla candidata `condor.tbestados`; columna sin asignar | candidata a nivel de tabla |
| `tbEstados.esEstadoFinal` | BOOLEAN(1) → BOOLEAN | — | P `Estado.cls:18` | tabla candidata `condor.tbestados`; columna sin asignar | candidata a nivel de tabla |
| `tbEstados.orden` | LONG(4) → INTEGER | — | P `Estado.cls:19` | tabla candidata `condor.tbestados`; columna sin asignar | candidata a nivel de tabla |

## tbTransiciones

4 columnas · 3 con evidencia en `src` · nombre de tabla citado en 2 ficheros de `src`.

| Origen | Tipo origen → objetivo | Clase de dato | Evidencia en `src` | Correspondencia propuesta | Estado |
|---|---|---|---|---|---|
| `tbTransiciones.idTransicion` | LONG(4) → INTEGER | — | — | tabla candidata `condor.tbtransiciones`; columna sin asignar | escrita solo por SQL, sin campo de entidad |
| `tbTransiciones.idEstadoOrigen` | LONG(4) → INTEGER | — | S `WorkflowRepositorio.bas:13` | tabla candidata `condor.tbtransiciones`; columna sin asignar | candidata a nivel de tabla |
| `tbTransiciones.idEstadoDestino` | LONG(4) → INTEGER | — | S `WorkflowRepositorio.bas:29` | tabla candidata `condor.tbtransiciones`; columna sin asignar | candidata a nivel de tabla |
| `tbTransiciones.rolRequerido` | TEXT(50) → VARCHAR(50) | — | S `WorkflowRepositorio.bas:14` | tabla candidata `condor.tbtransiciones`; columna sin asignar | candidata a nivel de tabla |

## tbAdjuntos

8 columnas · 8 con evidencia en `src` · nombre de tabla citado en 2 ficheros de `src`.

| Origen | Tipo origen → objetivo | Clase de dato | Evidencia en `src` | Correspondencia propuesta | Estado |
|---|---|---|---|---|---|
| `tbAdjuntos.idAdjunto` | LONG(4) → INTEGER | — | P `Adjunto.cls:17` | tabla candidata `condor.tbadjuntos`; columna sin asignar | candidata a nivel de tabla |
| `tbAdjuntos.idSolicitud` | LONG(4) → INTEGER | — | P `Adjunto.cls:18` | tabla candidata `condor.tbadjuntos`; columna sin asignar | candidata a nivel de tabla |
| `tbAdjuntos.etapaWF` | TEXT(50) → VARCHAR(50) | — | P `Adjunto.cls:19` | tabla candidata `condor.tbadjuntos`; columna sin asignar | candidata a nivel de tabla |
| `tbAdjuntos.nombreArchivo` | TEXT(255) → VARCHAR(255) | personal | P `Adjunto.cls:20` | tabla candidata `condor.tbadjuntos`; columna sin asignar | candidata a nivel de tabla |
| `tbAdjuntos.fechaSubida` | SHORT_DATE_TIME(8) → TIMESTAMP | — | P `Adjunto.cls:21` | tabla candidata `condor.tbadjuntos`; columna sin asignar | candidata a nivel de tabla |
| `tbAdjuntos.usuarioSubida` | TEXT(100) → VARCHAR(100) | personal | P `Adjunto.cls:22` | tabla candidata `condor.tbadjuntos`; columna sin asignar | candidata a nivel de tabla |
| `tbAdjuntos.descripcion` | MEMO → TEXT | personal | P `Adjunto.cls:23` | tabla candidata `condor.tbadjuntos`; columna sin asignar | candidata a nivel de tabla |
| `tbAdjuntos.TipoAccion` | TEXT(255) → VARCHAR(255) | — | P `Adjunto.cls:24` | tabla candidata `condor.tbadjuntos`; columna sin asignar | candidata a nivel de tabla |

## tbMapeoCampos

6 columnas · 6 con evidencia en `src` · nombre de tabla citado en 3 ficheros de `src`.

| Origen | Tipo origen → objetivo | Clase de dato | Evidencia en `src` | Correspondencia propuesta | Estado |
|---|---|---|---|---|---|
| `tbMapeoCampos.idMapeo` | LONG(4) → INTEGER | — | P `MapeoCampos.cls:17` | tabla candidata `condor.tbmapeocampos`; columna sin asignar | candidata a nivel de tabla |
| `tbMapeoCampos.nombrePlantilla` | TEXT(50) → VARCHAR(50) | personal | P `MapeoCampos.cls:18` | tabla candidata `condor.tbmapeocampos`; columna sin asignar | candidata a nivel de tabla |
| `tbMapeoCampos.nombreCampoTabla` | TEXT(100) → VARCHAR(100) | personal | P `MapeoCampos.cls:19` | tabla candidata `condor.tbmapeocampos`; columna sin asignar | candidata a nivel de tabla |
| `tbMapeoCampos.valorAsociado` | TEXT(100) → VARCHAR(100) | — | P `MapeoCampos.cls:20` | tabla candidata `condor.tbmapeocampos`; columna sin asignar | candidata a nivel de tabla |
| `tbMapeoCampos.nombreCampoWord` | TEXT(100) → VARCHAR(100) | personal | P `MapeoCampos.cls:21` | tabla candidata `condor.tbmapeocampos`; columna sin asignar | candidata a nivel de tabla |
| `tbMapeoCampos.numExtensiones` | INT(2) → INT(2) | — | P `MapeoCampos.cls:22` | tabla candidata `condor.tbmapeocampos`; columna sin asignar | candidata a nivel de tabla |

## tbRechazos

10 columnas · 9 con evidencia en `src` · nombre de tabla citado en 2 ficheros de `src`.

| Origen | Tipo origen → objetivo | Clase de dato | Evidencia en `src` | Correspondencia propuesta | Estado |
|---|---|---|---|---|---|
| `tbRechazos.idRechazo` | LONG(4) → INTEGER | — | — | tabla candidata `condor.tbrechazos`; columna sin asignar | escrita solo por SQL, sin campo de entidad |
| `tbRechazos.idSolicitud` | LONG(4) → INTEGER | — | P `Rechazo.cls:13` | tabla candidata `condor.tbrechazos`; columna sin asignar | candidata a nivel de tabla |
| `tbRechazos.fechaRechazo` | SHORT_DATE_TIME(8) → TIMESTAMP | — | P `Rechazo.cls:14` | tabla candidata `condor.tbrechazos`; columna sin asignar | candidata a nivel de tabla |
| `tbRechazos.motivoPrincipal` | TEXT(255) → VARCHAR(255) | personal | P `Rechazo.cls:15` | tabla candidata `condor.tbrechazos`; columna sin asignar | candidata a nivel de tabla |
| `tbRechazos.areaAfectada` | TEXT(255) → VARCHAR(255) | — | P `Rechazo.cls:16` | tabla candidata `condor.tbrechazos`; columna sin asignar | candidata a nivel de tabla |
| `tbRechazos.comentarios` | MEMO → TEXT | personal | P `Rechazo.cls:17` | tabla candidata `condor.tbrechazos`; columna sin asignar | candidata a nivel de tabla |
| `tbRechazos.usuarioRechazo` | TEXT(255) → VARCHAR(255) | personal | P `Rechazo.cls:18` | tabla candidata `condor.tbrechazos`; columna sin asignar | candidata a nivel de tabla |
| `tbRechazos.esActivo` | BOOLEAN(1) → BOOLEAN | — | S `RechazoRepositorio.bas:24` | tabla candidata `condor.tbrechazos`; columna sin asignar | candidata a nivel de tabla |
| `tbRechazos.notasSubsanacion` | MEMO → TEXT | personal | P `Rechazo.cls:19` | tabla candidata `condor.tbrechazos`; columna sin asignar | candidata a nivel de tabla |
| `tbRechazos.CambiosTecnico` | MEMO → TEXT | personal | P `Rechazo.cls:20` | tabla candidata `condor.tbrechazos`; columna sin asignar | candidata a nivel de tabla |

## tbValidacionRevision

10 columnas · 10 con evidencia en `src` · nombre de tabla citado en 2 ficheros de `src`.

| Origen | Tipo origen → objetivo | Clase de dato | Evidencia en `src` | Correspondencia propuesta | Estado |
|---|---|---|---|---|---|
| `tbValidacionRevision.Id` | LONG(4) → INTEGER | — | P `ValidacionRevision.cls:14` | tabla candidata `condor.tbvalidacionrevision`; columna sin asignar | candidata a nivel de tabla |
| `tbValidacionRevision.idSolicitud` | LONG(4) → INTEGER | — | P `ValidacionRevision.cls:15` | tabla candidata `condor.tbvalidacionrevision`; columna sin asignar | candidata a nivel de tabla |
| `tbValidacionRevision.ordinal` | INT(2) → INT(2) | — | P `ValidacionRevision.cls:16` | tabla candidata `condor.tbvalidacionrevision`; columna sin asignar | candidata a nivel de tabla |
| `tbValidacionRevision.idAdjunto` | LONG(4) → INTEGER | — | P `ValidacionRevision.cls:22` | tabla candidata `condor.tbvalidacionrevision`; columna sin asignar | candidata a nivel de tabla |
| `tbValidacionRevision.FechaEnvio` | SHORT_DATE_TIME(8) → TIMESTAMP | — | P `ValidacionRevision.cls:17` | tabla candidata `condor.tbvalidacionrevision`; columna sin asignar | candidata a nivel de tabla |
| `tbValidacionRevision.FechaRecepcion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | P `ValidacionRevision.cls:18` | tabla candidata `condor.tbvalidacionrevision`; columna sin asignar | candidata a nivel de tabla |
| `tbValidacionRevision.Resultado` | TEXT(255) → VARCHAR(255) | — | P `ValidacionRevision.cls:19` | tabla candidata `condor.tbvalidacionrevision`; columna sin asignar | candidata a nivel de tabla |
| `tbValidacionRevision.Comentarios` | MEMO → TEXT | personal | P `ValidacionRevision.cls:20` | tabla candidata `condor.tbvalidacionrevision`; columna sin asignar | candidata a nivel de tabla |
| `tbValidacionRevision.Usuario` | TEXT(255) → VARCHAR(255) | personal | P `ValidacionRevision.cls:21` | tabla candidata `condor.tbvalidacionrevision`; columna sin asignar | candidata a nivel de tabla |
| `tbValidacionRevision.HashDatos` | TEXT(255) → VARCHAR(255) | credencial | P `ValidacionRevision.cls:23` | tabla candidata `condor.tbvalidacionrevision`; columna sin asignar | candidata a nivel de tabla |

## tbLogCambios

10 columnas · 10 con evidencia en `src` · nombre de tabla citado en 5 ficheros de `src`.

| Origen | Tipo origen → objetivo | Clase de dato | Evidencia en `src` | Correspondencia propuesta | Estado |
|---|---|---|---|---|---|
| `tbLogCambios.idLogCambio` | LONG(4) → INTEGER | — | P `LogCambio.cls:18` | tabla candidata `condor.tblogcambios`; columna sin asignar | candidata a nivel de tabla |
| `tbLogCambios.fechaHora` | SHORT_DATE_TIME(8) → TIMESTAMP | — | P `LogCambio.cls:19` | tabla candidata `condor.tblogcambios`; columna sin asignar | candidata a nivel de tabla |
| `tbLogCambios.usuario` | TEXT(100) → VARCHAR(100) | personal | P `LogCambio.cls:20` | tabla candidata `condor.tblogcambios`; columna sin asignar | candidata a nivel de tabla |
| `tbLogCambios.suplantadoPor` | TEXT(255) → VARCHAR(255) | — | P `LogCambio.cls:21` | tabla candidata `condor.tblogcambios`; columna sin asignar | candidata a nivel de tabla |
| `tbLogCambios.tabla` | TEXT(50) → VARCHAR(50) | — | P `LogCambio.cls:22` | tabla candidata `condor.tblogcambios`; columna sin asignar | candidata a nivel de tabla |
| `tbLogCambios.registro` | LONG(4) → INTEGER | — | P `LogCambio.cls:23` | tabla candidata `condor.tblogcambios`; columna sin asignar | candidata a nivel de tabla |
| `tbLogCambios.campo` | TEXT(50) → VARCHAR(50) | — | P `LogCambio.cls:24` | tabla candidata `condor.tblogcambios`; columna sin asignar | candidata a nivel de tabla |
| `tbLogCambios.valorAnterior` | MEMO → TEXT | — | P `LogCambio.cls:25` | tabla candidata `condor.tblogcambios`; columna sin asignar | candidata a nivel de tabla |
| `tbLogCambios.valorNuevo` | MEMO → TEXT | — | P `LogCambio.cls:26` | tabla candidata `condor.tblogcambios`; columna sin asignar | candidata a nivel de tabla |
| `tbLogCambios.tipoOperacion` | TEXT(255) → VARCHAR(255) | — | P `LogCambio.cls:27` | tabla candidata `condor.tblogcambios`; columna sin asignar | candidata a nivel de tabla |

## tbLogErrores

9 columnas · 9 con evidencia en `src` · nombre de tabla citado en 6 ficheros de `src`.

| Origen | Tipo origen → objetivo | Clase de dato | Evidencia en `src` | Correspondencia propuesta | Estado |
|---|---|---|---|---|---|
| `tbLogErrores.idLogError` | LONG(4) → INTEGER | — | P `LogError.cls:17` | tabla candidata `condor.tblogerrores`; columna sin asignar | candidata a nivel de tabla |
| `tbLogErrores.fechaHora` | SHORT_DATE_TIME(8) → TIMESTAMP | — | P `LogError.cls:18` | tabla candidata `condor.tblogerrores`; columna sin asignar | candidata a nivel de tabla |
| `tbLogErrores.usuario` | TEXT(100) → VARCHAR(100) | personal | P `LogError.cls:19` | tabla candidata `condor.tblogerrores`; columna sin asignar | candidata a nivel de tabla |
| `tbLogErrores.suplantadoPor` | TEXT(255) → VARCHAR(255) | — | P `LogError.cls:20` | tabla candidata `condor.tblogerrores`; columna sin asignar | candidata a nivel de tabla |
| `tbLogErrores.modulo` | TEXT(100) → VARCHAR(100) | — | P `LogError.cls:21` | tabla candidata `condor.tblogerrores`; columna sin asignar | candidata a nivel de tabla |
| `tbLogErrores.procedimiento` | TEXT(100) → VARCHAR(100) | — | P `LogError.cls:22` | tabla candidata `condor.tblogerrores`; columna sin asignar | candidata a nivel de tabla |
| `tbLogErrores.numeroError` | LONG(4) → INTEGER | — | P `LogError.cls:23` | tabla candidata `condor.tblogerrores`; columna sin asignar | candidata a nivel de tabla |
| `tbLogErrores.descripcionError` | MEMO → TEXT | personal | P `LogError.cls:24` | tabla candidata `condor.tblogerrores`; columna sin asignar | candidata a nivel de tabla |
| `tbLogErrores.contexto` | MEMO → TEXT | — | P `LogError.cls:25` | tabla candidata `condor.tblogerrores`; columna sin asignar | candidata a nivel de tabla |

## tbLogEstados

6 columnas · 6 con evidencia en `src` · nombre de tabla citado en 3 ficheros de `src`.

| Origen | Tipo origen → objetivo | Clase de dato | Evidencia en `src` | Correspondencia propuesta | Estado |
|---|---|---|---|---|---|
| `tbLogEstados.idLogEstado` | LONG(4) → INTEGER | — | P `LogEstado.cls:20` | tabla candidata `condor.tblogestados`; columna sin asignar | candidata a nivel de tabla |
| `tbLogEstados.idSolicitud` | LONG(4) → INTEGER | — | P `LogEstado.cls:21` | tabla candidata `condor.tblogestados`; columna sin asignar | candidata a nivel de tabla |
| `tbLogEstados.idEstadoAnterior` | LONG(4) → INTEGER | — | P `LogEstado.cls:22` | tabla candidata `condor.tblogestados`; columna sin asignar | candidata a nivel de tabla |
| `tbLogEstados.idEstadoNuevo` | LONG(4) → INTEGER | — | P `LogEstado.cls:23` | tabla candidata `condor.tblogestados`; columna sin asignar | candidata a nivel de tabla |
| `tbLogEstados.fechaTransicion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | P `LogEstado.cls:24` | tabla candidata `condor.tblogestados`; columna sin asignar | candidata a nivel de tabla |
| `tbLogEstados.usuarioTransicion` | TEXT(255) → VARCHAR(255) | personal | P `LogEstado.cls:25` | tabla candidata `condor.tblogestados`; columna sin asignar | candidata a nivel de tabla |
