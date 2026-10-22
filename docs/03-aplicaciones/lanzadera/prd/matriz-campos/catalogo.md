# Matriz campo a campo — catalogo

> Parte de la [matriz campo a campo](../matriz-campos.md). Reglas de mapeo, leyenda y límites de la fuente en ese documento. Commit fijado `1474e8e`.

## TbAplicaciones

22 columnas · 20 filas en la extracción del snapshot · citada en 3 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbAplicaciones.IDAplicacion` | LONG(4) → INTEGER | — | F `Funciones Genreales.bas:3127` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.NombreAplicacion` | TEXT(255) → VARCHAR(255) | — | F `Aplicacion.cls:1884` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.NombreCorto` | TEXT(255) → VARCHAR(255) | — | F `Aplicacion.cls:1887` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.NombreEjecutable` | TEXT(255) → VARCHAR(255) | — | F `Aplicacion.cls:1890` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.NombreArchivoDatos` | TEXT(255) → VARCHAR(255) | — | F `Aplicacion.cls:1893` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.Pass` | TEXT(255) → VARCHAR(255) | credencial | F `Aplicacion.cls:1896` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.NombreCarpeta` | TEXT(255) → VARCHAR(255) | — | F `Aplicacion.cls:1902` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.NombreFuncionPublicacion` | TEXT(255) → VARCHAR(255) | — | S `Aplicacion.cls:34` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.NombreCarpetaTemporal` | TEXT(255) → VARCHAR(255) | — | F `Aplicacion.cls:1899` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.TituloAplicacion` | TEXT(255) → VARCHAR(255) | — | S `Aplicacion.cls:35` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.NombreIconoParaArbol` | TEXT(255) → VARCHAR(255) | — | F `Aplicacion.cls:1911` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.NombreIcono` | TEXT(255) → VARCHAR(255) | — | F `Aplicacion.cls:1905` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.NombreIconoLanzadera` | TEXT(255) → VARCHAR(255) | — | F `Aplicacion.cls:1908` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.EjecucionEnOficina` | TEXT(2) → VARCHAR(2) | — | F `Aplicacion.cls:1914` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.NombreCarpetaDocumentacion` | TEXT(255) → VARCHAR(255) | — | F `Aplicacion.cls:1917` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.NombreDirectorioIconos` | TEXT(255) → VARCHAR(255) | — | F `Aplicacion.cls:1920` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.NombreDirectorioAyuda` | TEXT(255) → VARCHAR(255) | — | F `Aplicacion.cls:1923` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.NombreDirectorioRecursos` | TEXT(255) → VARCHAR(255) | — | F `Aplicacion.cls:1926` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.URLDIrectorioIconoAplicacion` | TEXT(255) → VARCHAR(255) | — | S `Aplicacion.cls:36` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.EnPruebas` | TEXT(2) → VARCHAR(2) | — | F `Aplicacion.cls:1929` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.ConIconoEnLanzadera` | TEXT(2) → VARCHAR(2) | — | F `Aplicacion.cls:1932` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicaciones.Comando` | MEMO → TEXT | — | S `Aplicacion.cls:33` | tabla candidata `lanzadera.apps`; columna sin asignar | candidata a nivel de tabla |

## TbAplicacionesAperturas

15 columnas · 1697 filas en la extracción del snapshot · citada en 2 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbAplicacionesAperturas.IDApertura` | LONG(4) → INTEGER | — | F `Funciones Genreales.bas:3126` | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesAperturas.IDAplicacion` | LONG(4) → INTEGER | — | F `Funciones Genreales.bas:3127` | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesAperturas.NombreUsuario` | TEXT(255) → VARCHAR(255) | — | F `Funciones Genreales.bas:3128` | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesAperturas.FechaApertura` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Genreales.bas:3129` | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesAperturas.HoraApertura` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Genreales.bas:3130` | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesAperturas.FechaCierre` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Genreales.bas:3565` | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesAperturas.HoraCierre` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Genreales.bas:3566` | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesAperturas.NombreAplicacion` | TEXT(255) → VARCHAR(255) | — | F `Funciones Genreales.bas:3131` | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesAperturas.FechaEnvioCorreoAdministrador` | SHORT_DATE_TIME(8) → TIMESTAMP | personal | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesAperturas.EnOficina` | TEXT(2) → VARCHAR(2) | — | F `Funciones Genreales.bas:3132` | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesAperturas.UsuarioConectadoMaquina` | TEXT(255) → VARCHAR(255) | personal | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesAperturas.VersionAplicacion` | TEXT(255) → VARCHAR(255) | — | F `Funciones Genreales.bas:3135` | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesAperturas.NombreMaquina` | TEXT(255) → VARCHAR(255) | personal | F `Funciones Genreales.bas:3134` | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesAperturas.UsuarioMaquina` | TEXT(255) → VARCHAR(255) | personal | F `Funciones Genreales.bas:3133` | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesAperturas.Observaciones` | MEMO → TEXT | — | F `Funciones Genreales.bas:3567` | sin correspondencia en el esquema web actual | pendiente |

## TbAplicacionesEdiciones

5 columnas · 1262 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbAplicacionesEdiciones.IDAplicacion` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesEdiciones.IDVersion` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesEdiciones.Version` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesEdiciones.FechaPublicacion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesEdiciones.ParaInforme` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |

## TbAplicacionesEdicionesCambios

5 columnas · 1364 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbAplicacionesEdicionesCambios.IDCambio` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesEdicionesCambios.IDVersion` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesEdicionesCambios.Cambio` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesEdicionesCambios.FechaCambio` | SHORT_DATE_TIME(8) → TIMESTAMP | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesEdicionesCambios.DescripcionCambio` | MEMO → TEXT | — | — | sin correspondencia en el esquema web actual | pendiente |

## TbAplicacionesEstados

2 columnas · 6 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbAplicacionesEstados.PerfilAplicacion` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesEstados.PerfilAplicacionEncriptado` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |

## TbAplicacionesParametros

3 columnas · 60 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbAplicacionesParametros.IDAplicacion` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesParametros.IDParametro` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesParametros.Valor` | MEMO → TEXT | — | — | sin correspondencia en el esquema web actual | pendiente |

## TbAplicacionesPerfiles

2 columnas · 34 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbAplicacionesPerfiles.IDAplicacion` | LONG(4) → INTEGER | — | — | tabla candidata `lanzadera.profiles`; columna sin asignar | candidata a nivel de tabla |
| `TbAplicacionesPerfiles.Perfil` | TEXT(255) → VARCHAR(255) | — | — | tabla candidata `lanzadera.profiles`; columna sin asignar | candidata a nivel de tabla |

## TbAplicacionesVideos

9 columnas · 0 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbAplicacionesVideos.IDAplicacionVideo` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesVideos.IDVideo` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesVideos.IDAplicacion` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesVideos.Descripcion` | MEMO → TEXT | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesVideos.NombreArchivo` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesVideos.FechaCreacion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesVideos.UsuarioCrea` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesVideos.FechaModificacion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbAplicacionesVideos.UsuarioModifica` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |

## TbCategorias

2 columnas · 2 filas en la extracción del snapshot · citada en 2 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbCategorias.IDCategoria` | LONG(4) → INTEGER | — | S `Categoria.cls:13` | sin correspondencia en el esquema web actual | pendiente |
| `TbCategorias.NombreCategoria` | TEXT(255) → VARCHAR(255) | — | S `Categoria.cls:14` | sin correspondencia en el esquema web actual | pendiente |

## TbDetalleVersiones

4 columnas · 28 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbDetalleVersiones.IDAplicacion` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbDetalleVersiones.IDVersion` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbDetalleVersiones.IDDetalle` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbDetalleVersiones.Detalle` | MEMO → TEXT | — | — | sin correspondencia en el esquema web actual | pendiente |

## TbParametros

2 columnas · 53 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbParametros.IDParametro` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbParametros.Parametro` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |

## TbUbicaciones

3 columnas · 8 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbUbicaciones.NombreUbicacion` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbUbicaciones.Sirdee` | TEXT(2) → VARCHAR(2) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbUbicaciones.Ubicacion` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |

## TbTablasAVincular

4 columnas · 411 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbTablasAVincular.IDBBDD` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbTablasAVincular.IDAplicacion` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbTablasAVincular.NombreTabla` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbTablasAVincular.NombreTablaEnLocal` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |

## Tb0HerramientaDocAyuda

2 columnas · 4 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `Tb0HerramientaDocAyuda.NombreFormulario` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `Tb0HerramientaDocAyuda.NombreArchivoAyuda` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |

## TbConfiguracion

5 columnas · 1 filas en la extracción del snapshot · citada en 1 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbConfiguracion.Clave` | TEXT(50) → VARCHAR(50) | credencial | S `Funciones Genreales.bas:3742` | sin correspondencia en el esquema web actual | pendiente |
| `TbConfiguracion.Valor` | TEXT(255) → VARCHAR(255) | — | S `Funciones Genreales.bas:3742` | sin correspondencia en el esquema web actual | pendiente |
| `TbConfiguracion.TipoDato` | TEXT(20) → VARCHAR(20) | — | S `Funciones Genreales.bas:3742` | sin correspondencia en el esquema web actual | pendiente |
| `TbConfiguracion.Descripcion` | MEMO → TEXT | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbConfiguracion.FechaModificacion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | — | sin correspondencia en el esquema web actual | pendiente |
