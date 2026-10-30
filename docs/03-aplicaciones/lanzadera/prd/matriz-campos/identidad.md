# Matriz campo a campo — identidad

> Parte de la [matriz campo a campo](../matriz-campos.md). Reglas de mapeo, leyenda y límites de la fuente en ese documento. Commit fijado `1474e8e`.

## tbUsuarios

18 columnas · 156 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `tbUsuarios.Id` | LONG(4) → INTEGER | — | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `tbUsuarios.Nombre` | TEXT(50) → VARCHAR(50) | personal | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `tbUsuarios.UsuarioRed` | TEXT(50) → VARCHAR(50) | personal | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `tbUsuarios.DirCorreo` | TEXT(255) → VARCHAR(255) | personal | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `tbUsuarios.Matricula_DNI` | TEXT(50) → VARCHAR(50) | personal | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `tbUsuarios.Cargo` | TEXT(50) → VARCHAR(50) | — | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `tbUsuarios.telfijo` | LONG(4) → INTEGER | personal | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `tbUsuarios.telmovil` | LONG(4) → INTEGER | personal | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `tbUsuarios.JefeDelUsuario` | TEXT(50) → VARCHAR(50) | — | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `tbUsuarios.FechaAlta` | SHORT_DATE_TIME(8) → TIMESTAMP | — | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `tbUsuarios.FechaBaja` | SHORT_DATE_TIME(8) → TIMESTAMP | — | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `tbUsuarios.EmplazamientoExterno` | TEXT(2) → VARCHAR(2) | — | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `tbUsuarios.SeLogean` | BOOLEAN(1) → BOOLEAN | — | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `tbUsuarios.ParaTareasProgramadas` | BOOLEAN(1) → BOOLEAN | — | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `tbUsuarios.Autorizador` | BOOLEAN(1) → BOOLEAN | — | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `tbUsuarios.DiaEnvioTareas` | BYTE(1) → SMALLINT | — | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `tbUsuarios.UsuarioDeGestionRiesgos` | TEXT(2) → VARCHAR(2) | — | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `tbUsuarios.UsuariosI3D` | TEXT(2) → VARCHAR(2) | — | — | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |

## TbUsuariosAplicaciones

26 columnas · 186 filas en la extracción del snapshot · citada en 3 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbUsuariosAplicaciones.CorreoUsuario` | TEXT(255) → VARCHAR(255) | personal | F `Usuario.cls:1686` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.Password` | TEXT(255) → VARCHAR(255) | credencial | F `Usuario.cls:758` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.UsuarioRed` | TEXT(255) → VARCHAR(255) | personal | F `Usuario.cls:1692` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.Nombre` | TEXT(255) → VARCHAR(255) | personal | F `Usuario.cls:1688` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.Matricula` | TEXT(255) → VARCHAR(255) | personal | F `Usuario.cls:1690` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.FechaAlta` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Usuario.cls:1693` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.Activado` | BOOLEAN(1) → BOOLEAN | — | F `Usuario.cls:1015` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.FechaProximoCambioContrasenia` | SHORT_DATE_TIME(8) → TIMESTAMP | credencial | F `Usuario.cls:926` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.FechaUltimaConexion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | S `Usuario.cls:38` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.TieneQueCambiarLaContrasenia` | BOOLEAN(1) → BOOLEAN | credencial | F `Usuario.cls:760` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.Telefono` | TEXT(255) → VARCHAR(255) | personal | F `Usuario.cls:1704` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.Movil` | TEXT(255) → VARCHAR(255) | personal | F `Usuario.cls:1707` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.Observaciones` | MEMO → TEXT | — | F `Usuario.cls:1710` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.UsuarioImborrable` | BOOLEAN(1) → BOOLEAN | — | F `Usuario.cls:1717` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.EsAdministrador` | TEXT(2) → VARCHAR(2) | — | F `Usuario.cls:1718` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.PermisosAsignados` | BOOLEAN(1) → BOOLEAN | — | S `Usuario.cls:27` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.FechaBaja` | SHORT_DATE_TIME(8) → TIMESTAMP | — | S `Usuario.cls:28` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.PasswordNuncaCaduca` | BOOLEAN(1) → BOOLEAN | credencial | F `Usuario.cls:1702` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.MantenerLanzaderaAbierta` | BOOLEAN(1) → BOOLEAN | — | F `Usuario.cls:1719` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.PassIncialPlana` | TEXT(255) → VARCHAR(255) | credencial | F `Usuario.cls:763` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.UsuarioSSID` | TEXT(255) → VARCHAR(255) | personal | S `Usuario.cls:32` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.Id` | INT(2) → SMALLINT | — | F `Usuario.cls:1685` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.JefeDelUsuario` | TEXT(50) → VARCHAR(50) | — | S `Usuario.cls:34` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.PermisoPruebas` | TEXT(2) → VARCHAR(2) | — | S `Usuario.cls:35` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.ParaTareasProgramadas` | BOOLEAN(1) → BOOLEAN | — | S `Usuario.cls:36` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicaciones.FechaBloqueo` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Usuario.cls:1697` | tabla candidata `lanzadera.users`; columna sin asignar | candidata a nivel de tabla |

## TbUsuariosAplicacionesPermisos

9 columnas · 622 filas en la extracción del snapshot · citada en 3 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbUsuariosAplicacionesPermisos.CorreoUsuario` | TEXT(255) → VARCHAR(255) | personal | F `Usuario.cls:1686` | tabla candidata `lanzadera.user_app_assignments`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicacionesPermisos.IDAplicacion` | LONG(4) → INTEGER | — | F `UsuarioAplicacionPermisos.cls:307` | tabla candidata `lanzadera.user_app_assignments`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicacionesPermisos.EsUsuarioAdministrador` | TEXT(2) → VARCHAR(2) | — | F `UsuarioAplicacionPermisos.cls:362` | tabla candidata `lanzadera.user_app_assignments`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicacionesPermisos.EsUsuarioCalidad` | TEXT(2) → VARCHAR(2) | — | F `UsuarioAplicacionPermisos.cls:363` | tabla candidata `lanzadera.user_app_assignments`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicacionesPermisos.EsUsuarioEconomia` | TEXT(2) → VARCHAR(2) | — | F `UsuarioAplicacionPermisos.cls:364` | tabla candidata `lanzadera.user_app_assignments`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicacionesPermisos.EsUsuarioSecretaria` | TEXT(2) → VARCHAR(2) | — | F `UsuarioAplicacionPermisos.cls:365` | tabla candidata `lanzadera.user_app_assignments`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicacionesPermisos.EsUsuarioTecnico` | TEXT(2) → VARCHAR(2) | — | F `UsuarioAplicacionPermisos.cls:366` | tabla candidata `lanzadera.user_app_assignments`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicacionesPermisos.EsUsuarioSinAcceso` | TEXT(2) → VARCHAR(2) | — | F `UsuarioAplicacionPermisos.cls:369` | tabla candidata `lanzadera.user_app_assignments`; columna sin asignar | candidata a nivel de tabla |
| `TbUsuariosAplicacionesPermisos.EsUsuarioCalidadAvisos` | TEXT(2) → VARCHAR(2) | — | F `UsuarioAplicacionPermisos.cls:367` | tabla candidata `lanzadera.user_app_assignments`; columna sin asignar | candidata a nivel de tabla |

## TbUsuariosHistoricoContrasenias

3 columnas · 455 filas en la extracción del snapshot · citada en 3 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbUsuariosHistoricoContrasenias.Usuario` | TEXT(255) → VARCHAR(255) | — | F `HistoricoPass.cls:46` | sin correspondencia en el esquema web actual | pendiente |
| `TbUsuariosHistoricoContrasenias.PassAntigua` | TEXT(255) → VARCHAR(255) | credencial | F `HistoricoPass.cls:47` | sin correspondencia en el esquema web actual | pendiente |
| `TbUsuariosHistoricoContrasenias.FechaPass` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `HistoricoPass.cls:48` | sin correspondencia en el esquema web actual | pendiente |

## TbUsuarioConfiguracion

2 columnas · 14 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbUsuarioConfiguracion.UsuarioDeRed` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbUsuarioConfiguracion.MantenerLanzaderaAbierta` | TEXT(2) → VARCHAR(2) | — | — | sin correspondencia en el esquema web actual | pendiente |

## TbUsuariosAplicacionesTareas

5 columnas · 7 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbUsuariosAplicacionesTareas.CorreoUsuario` | TEXT(255) → VARCHAR(255) | personal | — | sin correspondencia en el esquema web actual | pendiente |
| `TbUsuariosAplicacionesTareas.EsAdministrador` | TEXT(2) → VARCHAR(2) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbUsuariosAplicacionesTareas.EsTecnico` | TEXT(2) → VARCHAR(2) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbUsuariosAplicacionesTareas.EsCalidad` | TEXT(2) → VARCHAR(2) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbUsuariosAplicacionesTareas.EsEconomia` | TEXT(2) → VARCHAR(2) | — | — | sin correspondencia en el esquema web actual | pendiente |

## TbUsuariosTareasDiarias

1 columnas · 8 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbUsuariosTareasDiarias.UsuarioRed` | TEXT(255) → VARCHAR(255) | personal | — | sin correspondencia en el esquema web actual | pendiente |

## TbUsuarioAplicacionesSolicitud

7 columnas · 0 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbUsuarioAplicacionesSolicitud.CorreoUsuario` | TEXT(255) → VARCHAR(255) | personal | — | sin correspondencia en el esquema web actual | pendiente |
| `TbUsuarioAplicacionesSolicitud.Password` | TEXT(255) → VARCHAR(255) | credencial | — | sin correspondencia en el esquema web actual | pendiente |
| `TbUsuarioAplicacionesSolicitud.Nombre` | TEXT(255) → VARCHAR(255) | personal | — | sin correspondencia en el esquema web actual | pendiente |
| `TbUsuarioAplicacionesSolicitud.Matricula` | TEXT(255) → VARCHAR(255) | personal | — | sin correspondencia en el esquema web actual | pendiente |
| `TbUsuarioAplicacionesSolicitud.Telefono` | TEXT(255) → VARCHAR(255) | personal | — | sin correspondencia en el esquema web actual | pendiente |
| `TbUsuarioAplicacionesSolicitud.Movil` | TEXT(255) → VARCHAR(255) | personal | — | sin correspondencia en el esquema web actual | pendiente |
| `TbUsuarioAplicacionesSolicitud.FechaSolicitud` | SHORT_DATE_TIME(8) → TIMESTAMP | — | — | sin correspondencia en el esquema web actual | pendiente |

## TbPermisos

9 columnas · 205 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Correspondencia en la web actual | Estado |
|---|---|---|---|---|---|
| `TbPermisos.IDAplicacion` | LONG(4) → INTEGER | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbPermisos.Usuario` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbPermisos.F3` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbPermisos.F4` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbPermisos.F5` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbPermisos.F6` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbPermisos.F7` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbPermisos.F8` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
| `TbPermisos.F9` | TEXT(255) → VARCHAR(255) | — | — | sin correspondencia en el esquema web actual | pendiente |
