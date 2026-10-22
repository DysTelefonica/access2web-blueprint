# Clases de trazabilidad: error, estados y registros de log

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/classes/LogServicio.cls` (225 líneas), `CondorError.cls` (103), `LogCambio.cls` (58), `LogError.cls` (55), `LogEstado.cls` (52), `Estado.cls` (51), `EstadoServicio.cls` (36), `LogCambioServicio.cls` (33) · **PRD:** [README](../README.md)

La infraestructura de auditoría de Condor, que va desde la excepción tipada hasta las tres tablas de log. `CondorError` es la pieza transversal: cualquier fallo de la aplicación se envuelve en ella, se registra sola y se vuelve a lanzar con la pila de llamadas en la descripción. `LogServicio` escribe y lee los registros de cambio y de error; los repositorios correspondientes ya seGiron en la hoja de módulos de trazabilidad.

Abreviaturas: `LS`, `classes/LogServicio.cls`; `CE`, `classes/CondorError.cls`; `LC`, `classes/LogCambio.cls`; `LE`, `classes/LogError.cls`; `LEE`, `classes/LogEstado.cls`; `EST`, `classes/Estado.cls`; `ES`, `classes/EstadoServicio.cls`; `LCS`, `classes/LogCambioServicio.cls`; `LR`, `modules/LogCambioRepositorio.bas`; `LER`, `modules/LogErrorRepositorio.bas`; `VG`, `modules/Variables Globales.bas`.

## Superficie pública

| Objeto | Tipo | Procedimientos y campos públicos | Qué resuelve |
|---|---|---|---|
| `CondorError` | clase de error | `Create`, `AddToCallStack`, `Raise`, `CallStackToString` (`CE:50-103`) | Excepción propia de la aplicación, con contexto acumulado |
| `LogServicio` | clase de servicio | `RegistrarCambio`, `RegistrarError`, `getLogs`, `getErrores`, `getLogCambioPorID`, `getLogErrorPorID` (`LS:18-165`) | Escribir y leer los registros de cambio y de error |
| `LogCambio` | clase de entidad | `idLogCambio`, `fechaHora`, `usuario`, `suplantadoPor`, `tabla`, `registro`, `campo`, `valorAnterior`, `valorNuevo`, `tipoOperacion` (`LC:18-49`) | Fila de `tbLogCambios` |
| `LogError` | clase de entidad | `idLogError`, `fechaHora`, `usuario`, `suplantadoPor`, `modulo`, `procedimiento`, `numeroError`, `descripcionError`, `contexto` (`LE:17-46`) | Fila de `tbLogErrores` |
| `LogEstado` | clase de entidad | `idLogEstado`, `idSolicitud`, `idEstadoAnterior`, `idEstadoNuevo`, `fechaTransicion`, `usuarioTransicion` (`LEE:20-44`) | Fila de `tbLogEstados` |
| `Estado` | clase de entidad | `idEstado`, `nombreEstado`, `descripcion`, `esEstadoInicial`, `esEstadoFinal`, `orden` (`EST:14-42`) | Fila de `tbEstados` |
| `EstadoServicio` | clase de servicio | `getEstadoPorID` (`ES:17-36`) | Resolver un estado por su identificador |
| `LogCambioServicio` | clase de servicio | `LimpiarHistorialCambios` (`LCS:19-33`) | Borrar el historial de cambios de una solicitud |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| `CE` | `Create` | Guarda número, descripción y origen del error | — | Sin efectos | `CE:50-58` |
| `CE` | `AddToCallStack` | Añade el procedimiento al contexto | — | Sin efectos | `CE:59-63` |
| `CE` | `Raise` | Guarda el error como último error, lo registra y lo relanza con la pila en la descripción | `ErrorLogger.LogToTable`, `Err.Raise` | `INSERT` en `tbLogErrores` | `CE:64-75`; registro en `CE:69`, relanzamiento en `CE:74` |
| `CE` | `CallStackToString` | Convierte la pila en texto para persistencia | — | Sin efectos | `CE:78-103` |
| `LS` | `RegistrarCambio` | Valida y escribe un registro de cambio con usuario y suplantación | `ValidarCambio`, `LogCambioRepositorio` | `INSERT` en `tbLogCambios` | `LS:18-49`; usuario en `LS:26`, suplantación en `LS:27` |
| `LS` | `RegistrarError` | Valida y escribe un registro de error | `ValidarError`, `LogErrorRepositorio` | `INSERT` en `tbLogErrores` | `LS:50-91` |
| `LS` | `getLogs`, `getErrores` | Filtran los registros por usuario, tabla, registro, módulo o procedimiento | `getLogsSQL`, `getErroresSQL` | Lectura | `LS:92-134`, `LS:165-194` |
| `LS` | `getLogCambioPorID`, `getLogErrorPorID` | Devuelven un registro por su clave | repositorios | Lectura | `LS:135-164` |
| `LS` | `ValidarCambio`, `ValidarError` | Aplican las reglas mínimas de un registro de log | — | Sin escritura | `LS:195-225`; objeto nulo en `LS:198`, usuario en `LS:199`, tabla en `LS:200`, registro en `LS:201`, campo en `LS:202` |
| `ES` | `getEstadoPorID` | Resuelve un estado | `EstadoRepositorio` | Lectura | `ES:17-36` |
| `LCS` | `LimpiarHistorialCambios` | Borra el historial de una solicitud | `LogCambioRepositorio` | `DELETE` | `LCS:19-33` |

## Flujos de trabajo

- **Normal — error:** una excepción de VBA se envuelve en `CondorError` con su número y su origen (`CE:50`), el servicio añade su contexto (`CE:59`) y al lanzar se registra sola (`CE:69`) antes de volver a lanzar el error original con la pila de llamadas dentro de la descripción (`CE:74`). Así ninguna ruta de error se queda sin traza, ni siquiera las que no pasan por el manejador del formulario.
- **Normal — cambio:** `RegistrarCambio` fecha el registro con la hora actual, toma el usuario activo y, si hay suplantación en curso, guarda también quién la inició (`LS:25-27`), valida (`LS:195-202`) y escribe.
- **Alternativos:** la suplantación solo aparece en `tbLogCambios` y `tbLogErrores` (`LC:21`, `LE:20`): `tbLogEstados` no la registra.
- **Errores:** `RegistrarCambio` valida con el error 513 (`LS:198-202`), que el presentador clasifica como validación de negocio y no como fallo técnico.

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Toda excepción de la aplicación queda registrada al lanzarse | `CondorError.Raise` | `CE:64-75` |
| El contexto se acumula por niveles y se escribe en la descripción | `AddToCallStack`, `Raise` | `CE:59`, `CE:74` |
| Un registro de log sin usuario, tabla, registro o campo no se escribe | `ValidarCambio` | `LS:199-202` |
| Bajo suplantación, el log guarda quién suplantó a quién | `RegistrarCambio` | `LS:26-27` |
| El historial de estados no registra la suplantación | entidad `LogEstado` | `LEE:20-44` |

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Cambios | `tbLogCambios` | `Recordset` y SQL | La que pase el servicio | `LS:18-49`, `LS:135-164` |
| Errores | `tbLogErrores` | `Recordset` y SQL | La que pase el servicio | `CE:69`, `LS:50-91` |
| Estados | `tbEstados` | SQL parametrizada | Ninguna | `ES:17-36` |
| Historial de estados | `tbLogEstados` | `LogEstadoRepositorio` | La que pase el servicio | hoja de repositorios de catálogos |

## Coexistencia

Dos decisiones. La primera es el formato del error: hoy el número de error de VBA viaja dentro del código de la excepción y la pila de llamadas viaja dentro de la descripción (`CE:74`), de modo que el destino necesita una estructura de error con esos dos niveles, o una forma de conservarlos sin depender del motor de VBA. La segunda es la suplantación: `RegistrarCambio` guarda el usuario activo y el real cuando hay suplantación (`LS:26-27`), un rasgo que en la versión web debe pasar a ser un atributo de la sesión y del registro de auditoría, no dos columnas que se rellenan según una variable global.

## Comportamiento actual · Posibles defects · Propuestas web pendientes de decisión

**Comportamiento actual.** Una excepción propia que se autoregistra al lanzarse, un servicio de auditoría con validación y filtrado, y cinco entidades sin comportamiento.

**Posibles defectos, a verificar.**
1. `CondorError.Raise` escribe en la base desde el manejador del error (`CE:69`): si el fallo original es de conexión, el registro también falla y la traza se pierde.
2. `LogServicio.RegistrarCambio` depende del usuario activo global (`LS:26`): fuera de una sesión iniciada escribe un valor vacío y la validación lo rechaza.
3. La suplantación se registra solo en dos de los tres logs (`LC:21`, `LE:20`, `LEE:20-44`).
4. `getLogs` y `getErrores` filtran por texto libre (`LS:92-134`, `LS:165-194`).

**Propuestas web pendientes de decisión.** Que la auditoría sea un servicio del destino al que toda operación sensible se suscriba, en lugar de llamadas explícitas. Que la suplantación sea un atributo de la sesión propagado a toda traza. Que el error de la API tenga una estructura estable con código, origen y contexto.

## Criterios de aceptación

1. Toda excepción de la aplicación deja registro con usuario, módulo, procedimiento y contexto. Trazado a `CE:64-75`.
2. Un registro de log sin usuario, tabla, registro o campo se rechaza con error de validación. Trazado a `LS:195-202`.
3. Bajo suplantación, el log identifica al usuario real y al activo. Trazado a `LS:26-27`.
4. El filtrado de cambios y de errores devuelve solo los registros que cumplen los criterios. Trazado a `LS:92-134`.

## Pendientes

- Confirmar contra el binario si `tbLogEstados` tiene columna de suplantación que el código no rellena (`LEE:20-44`).
- Confirmar si existe alguna ruta que lance `CondorError` sin pasar por `Raise`, con lo que quedaría sin registrar (`CE:64`).
- Confirmar si `EstadoServicio.getEstadoPorID` tiene consumidores vivos o si el estado se resuelve siempre desde el diccionario del catálogo (`ES:17-36`).
- Confirmar la retención de las tres tablas de log, porque condiciona la matriz de campos y la política de auditoría.