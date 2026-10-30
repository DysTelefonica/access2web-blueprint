# Matriz campo a campo — conexiones

> Parte de la [matriz campo a campo](../matriz-campos.md). Reglas de mapeo, leyenda y límites de la fuente en ese documento. Commit fijado `1474e8e`.

## TbConexiones

6 columnas · 24 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbConexiones.Usuario` | TEXT(255) → VARCHAR(255) | — | — | tabla candidata `lanzadera.presence_sessions`; columna sin asignar | candidata a nivel de tabla |
| `TbConexiones.UltimaConexion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | — | tabla candidata `lanzadera.presence_sessions`; columna sin asignar | candidata a nivel de tabla |
| `TbConexiones.UltimaDesconexion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | — | tabla candidata `lanzadera.presence_sessions`; columna sin asignar | candidata a nivel de tabla |
| `TbConexiones.InstaladoFW3` | TEXT(2) → VARCHAR(2) | — | — | tabla candidata `lanzadera.presence_sessions`; columna sin asignar | candidata a nivel de tabla |
| `TbConexiones.InstaladoFW4` | TEXT(2) → VARCHAR(2) | — | — | tabla candidata `lanzadera.presence_sessions`; columna sin asignar | candidata a nivel de tabla |
| `TbConexiones.Exitoso` | TEXT(2) → VARCHAR(2) | — | — | tabla candidata `lanzadera.presence_sessions`; columna sin asignar | candidata a nivel de tabla |

## TbConexionesRegistro

9 columnas · 10360 filas en la extracción del snapshot · citada en 4 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbConexionesRegistro.IDConexion` | LONG(4) → INTEGER | — | F `Conexion.cls:120` | tabla candidata `lanzadera.presence_sessions`; columna sin asignar | candidata a nivel de tabla |
| `TbConexionesRegistro.Usuario` | TEXT(255) → VARCHAR(255) | — | F `Conexion.cls:121` | tabla candidata `lanzadera.presence_sessions`; columna sin asignar | candidata a nivel de tabla |
| `TbConexionesRegistro.FechaConexion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Conexion.cls:122` | tabla candidata `lanzadera.presence_sessions`; columna sin asignar | candidata a nivel de tabla |
| `TbConexionesRegistro.FechaCierre` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Genreales.bas:3565` | tabla candidata `lanzadera.presence_sessions`; columna sin asignar | candidata a nivel de tabla |
| `TbConexionesRegistro.ConContraseña` | BOOLEAN(1) → BOOLEAN | credencial | F `Conexion.cls:123` | tabla candidata `lanzadera.presence_sessions`; columna sin asignar | candidata a nivel de tabla |
| `TbConexionesRegistro.UsuarioSSID` | TEXT(255) → VARCHAR(255) | personal | F `Conexion.cls:124` | tabla candidata `lanzadera.presence_sessions`; columna sin asignar | candidata a nivel de tabla |
| `TbConexionesRegistro.EnOficina` | BOOLEAN(1) → BOOLEAN | — | F `Conexion.cls:125` | tabla candidata `lanzadera.presence_sessions`; columna sin asignar | candidata a nivel de tabla |
| `TbConexionesRegistro.Vertical` | LONG(4) → INTEGER | — | S `Conexion.cls:20` | tabla candidata `lanzadera.presence_sessions`; columna sin asignar | candidata a nivel de tabla |
| `TbConexionesRegistro.Horizontal` | LONG(4) → INTEGER | — | S `Conexion.cls:21` | tabla candidata `lanzadera.presence_sessions`; columna sin asignar | candidata a nivel de tabla |

## TbConexionUltimaAppAbierta

2 columnas · 1469 filas en la extracción del snapshot · citada en 1 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbConexionUltimaAppAbierta.IDConexion` | LONG(4) → INTEGER | — | S `Constructor.bas:602` | sin correspondencia en el esquema web actual | pendiente |
| `TbConexionUltimaAppAbierta.IDUltimaAplicacionAbierta` | LONG(4) → INTEGER | — | S `Constructor.bas:1099` | sin correspondencia en el esquema web actual | pendiente |

## TbUsuariosCorreosEnvio

9 columnas · 1042 filas en la extracción del snapshot · citada en 3 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbUsuariosCorreosEnvio.IDCorreo` | LONG(4) → INTEGER | personal | F `Correo.cls:90` | tabla candidata `lanzadera.mail_outbox`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosCorreosEnvio.Destinatarios` | MEMO → TEXT | — | F `Correo.cls:97` | tabla candidata `lanzadera.mail_outbox`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosCorreosEnvio.DestinatariosConCopia` | MEMO → TEXT | — | F `Correo.cls:100` | tabla candidata `lanzadera.mail_outbox`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosCorreosEnvio.DestinatariosConCopiaOculta` | MEMO → TEXT | — | F `Correo.cls:103` | tabla candidata `lanzadera.mail_outbox`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosCorreosEnvio.Asunto` | TEXT(255) → VARCHAR(255) | — | F `Correo.cls:105` | tabla candidata `lanzadera.mail_outbox`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosCorreosEnvio.Cuerpo` | MEMO → TEXT | — | F `Correo.cls:107` | tabla candidata `lanzadera.mail_outbox`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosCorreosEnvio.FechaEnvio` | SHORT_DATE_TIME(8) → TIMESTAMP | — | S `Correo.cls:19` | tabla candidata `lanzadera.mail_outbox`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosCorreosEnvio.FechaCreado` | SHORT_DATE_TIME(8) → TIMESTAMP | — | — | tabla candidata `lanzadera.mail_outbox`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosCorreosEnvio.URLAdjunto` | TEXT(255) → VARCHAR(255) | — | S `Correo.cls:21` | tabla candidata `lanzadera.mail_outbox`; columna sin asignar | candidata a nivel de tabla |

## TbCuestionarios

6 columnas · 0 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbCuestionarios.IDCuestionario` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbCuestionarios.FechaRealizado` | SHORT_DATE_TIME(8) → TIMESTAMP | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbCuestionarios.IDUsuarioRealiza` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbCuestionarios.IDAplicacion` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbCuestionarios.IDRespuestaCorrecta` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbCuestionarios.Observaciones` | MEMO → TEXT | — | — | sin correspondencia en el esquema web actual | pendiente |

## TbCuestionarioPreguntas

3 columnas · 0 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbCuestionarioPreguntas.IDPregunta` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbCuestionarioPreguntas.IDCuestionario` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbCuestionarioPreguntas.Texto` | MEMO → TEXT | — | — | sin correspondencia en el esquema web actual | pendiente |

## TbCuestionaroRespuestas

4 columnas · 0 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbCuestionaroRespuestas.IDRespuesta` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbCuestionaroRespuestas.IDPregunta` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbCuestionaroRespuestas.Letra` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbCuestionaroRespuestas.Texto` | MEMO → TEXT | — | — | sin correspondencia en el esquema web actual | pendiente |

## TbVideos

11 columnas · 16 filas en la extracción del snapshot · citada en 2 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbVideos.IDVideo` | LONG(4) → INTEGER | — | F `Video.cls:526` | sin correspondencia en el esquema web actual | pendiente |
| `TbVideos.Titulo` | TEXT(255) → VARCHAR(255) | — | F `Video.cls:529` | sin correspondencia en el esquema web actual | pendiente |
| `TbVideos.NombreArchivo` | TEXT(255) → VARCHAR(255) | — | F `Video.cls:528` | sin correspondencia en el esquema web actual | pendiente |
| `TbVideos.IDAplicacion` | LONG(4) → INTEGER | — | F `Video.cls:527` | sin correspondencia en el esquema web actual | pendiente |
| `TbVideos.Observaciones` | MEMO → TEXT | — | F `Video.cls:532` | sin correspondencia en el esquema web actual | pendiente |
| `TbVideos.Descripcion` | MEMO → TEXT | — | F `Video.cls:530` | sin correspondencia en el esquema web actual | pendiente |
| `TbVideos.SubidoPor` | TEXT(255) → VARCHAR(255) | — | F `Video.cls:534` | sin correspondencia en el esquema web actual | pendiente |
| `TbVideos.FechaSubido` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Video.cls:535` | sin correspondencia en el esquema web actual | pendiente |
| `TbVideos.ParaCalidad` | TEXT(2) → VARCHAR(2) | — | F `Video.cls:537` | sin correspondencia en el esquema web actual | pendiente |
| `TbVideos.ParaAdministrador` | TEXT(2) → VARCHAR(2) | — | F `Video.cls:536` | sin correspondencia en el esquema web actual | pendiente |
| `TbVideos.ParaTecnicos` | TEXT(2) → VARCHAR(2) | — | F `Video.cls:538` | sin correspondencia en el esquema web actual | pendiente |

## TbVideosCategorias

3 columnas · 15 filas en la extracción del snapshot · citada en 2 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbVideosCategorias.IDCategoriaVideo` | LONG(4) → INTEGER | — | F `Video.cls:578` | sin correspondencia en el esquema web actual | pendiente |
| `TbVideosCategorias.IDCategoria` | LONG(4) → INTEGER | — | F `Video.cls:579` | sin correspondencia en el esquema web actual | pendiente |
| `TbVideosCategorias.IDVideo` | LONG(4) → INTEGER | — | F `Video.cls:526` | sin correspondencia en el esquema web actual | pendiente |

## TbVideosCuestionario

6 columnas · 0 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbVideosCuestionario.IDVideoCuestionario` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbVideosCuestionario.IDCuestionario` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbVideosCuestionario.IDVideo` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbVideosCuestionario.FechaRealizado` | SHORT_DATE_TIME(8) → TIMESTAMP | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbVideosCuestionario.UsuarioRealiza` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbVideosCuestionario.Observaciones` | MEMO → TEXT | — | — | sin correspondencia en el esquema web actual | pendiente |

## TbVideosVisionados

6 columnas · 138 filas en la extracción del snapshot · citada en 1 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbVideosVisionados.IDVisionado` | LONG(4) → INTEGER | — | F `Visionado.cls:262` | sin correspondencia en el esquema web actual | pendiente |
| `TbVideosVisionados.IDVideo` | LONG(4) → INTEGER | — | F `Visionado.cls:263` | sin correspondencia en el esquema web actual | pendiente |
| `TbVideosVisionados.TiempoVisionado` | LONG(4) → INTEGER | — | F `Visionado.cls:264` | sin correspondencia en el esquema web actual | pendiente |
| `TbVideosVisionados.TiempoVideo` | LONG(4) → INTEGER | — | F `Visionado.cls:267` | sin correspondencia en el esquema web actual | pendiente |
| `TbVideosVisionados.IDUsuario` | LONG(4) → INTEGER | — | F `Visionado.cls:265` | sin correspondencia en el esquema web actual | pendiente |
| `TbVideosVisionados.FechaVisionado` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Visionado.cls:266` | sin correspondencia en el esquema web actual | pendiente |
