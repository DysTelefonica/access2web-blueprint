# Módulos de repositorio de trazabilidad y validación

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/modules/LogCambioRepositorio.bas` (123 líneas), `LogErrorRepositorio.bas` (106), `CorreoRepositorio.bas` (90), `NoConformidadRepositorio.bas` (233), `ValidacionRevisionRepositorio.bas` (606) · **PRD:** [README](../README.md)

Cinco repositorios que registran lo que pasa alrededor de una solicitud: qué se cambió, qué falló, qué correo salió, qué no conformidad se le asoció y qué validación de revisión se emitió. Todos escriben con `Recordset` sobre la tabla y calculan su clave con `Max(id) + 1`; solo `NoConformidadRepositorio` y `ValidacionRevisionRepositorio` aceptan una conexión externa como parámetro, y es justamente lo que permite que una no conformidad viva en la base de No Conformidades y no en la de Condor.

Abreviaturas: `LOG`, `modules/LogCambioRepositorio.bas`; `ERR`, `modules/LogErrorRepositorio.bas`; `COR`, `modules/CorreoRepositorio.bas`; `NC`, `modules/NoConformidadRepositorio.bas`; `VR`, `modules/ValidacionRevisionRepositorio.bas`; `CE`, `classes/CondorError.cls`; `RC`, `modules/RepositorioComun.bas`.

## Superficie pública

| Objeto | Tipo | Procedimientos públicos | Qué resuelve |
|---|---|---|---|
| `LogCambioRepositorio` | módulo | `GuardarLog`, `getLogCambioPorID`, `EliminarPorSolicitud`, `GetCambiosPorIntervalo` (`LOG:9-123`) | Auditoría de cambios, incluido el rango de fechas |
| `LogErrorRepositorio` | módulo | `GuardarError`, `getLogErrorPorID`, `GuardarDesdeCondorError` (`ERR:12-106`) | Auditoría de errores y traducción de `CondorError` a fila de log |
| `CorreoRepositorio` | módulo | `GuardarCorreo` (`COR:40-90`) | Registro del correo enviado |
| `NoConformidadRepositorio` | módulo | `getTodas`, `getPorExpediente`, `getPorCodigoCondor`, `existeEnBaseDatosExternaa`, `actualizarCodigoConcesion`, `limpiarCodigoConcesion`, `getPorID`, `LimpiarCodConcesionAsociada` (`NC:10-233`) | No conformidades de **otra base de datos** y su vínculo con el código de concesión de Condor |
| `ValidacionRevisionRepositorio` | módulo | `GetUltimoOrdinal`, `GetUltimoCicloCerrado`, `GetHistorialPorSolicitud`, `EliminarPorSolicitud`, `ActualizarHashRevision`, `GetUltimoPendienteId`, `GetCicloAbierto`, `GetAdjuntoIdCicloActual`, `GuardarHitoEnvio`, `GuardarHitoBorrador`, `GetHashVersionAnterior`, `CerrarHitoRespuesta`, `ActualizarRespuestaPorId`, `GetVersionActualBorrador`, `GetUltimoEstadoValidacion`, `CrearNuevoCicloValidacion` (`VR:6-606`) | Ciclo de vida completo de la validación de revisión:Ordinal, envío, borrador, respuesta, cierre y versionado del hash de datos |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| `LOG` | `GuardarLog` | Inserta el registro de cambio con clave calculada | `getSiguienteIDLogCambios` | `INSERT` en `tbLogCambios` | `LOG:9-32`; `Recordset` en `LOG:19`, `Max` en `LOG:39` |
| `LOG` | `getLogCambioPorID` | Hidrata un cambio por clave | `RC` | `SELECT` sobre `tbLogCambios` | `LOG:55-72`; SQL en `LOG:61` |
| `LOG` | `EliminarPorSolicitud` | Borra los cambios de una solicitud, identificando la tabla con un literal | `RC` | `DELETE` sobre `tbLogCambios` | `LOG:74-94`; literal de tabla en `LOG:81` |
| `LOG` | `GetCambiosPorIntervalo` | Devuelve los cambios desde una fecha, como `Recordset` abierto | `getdb()` | `SELECT` sobre `tbLogCambios` | `LOG:96-123`; SQL en `LOG:108-113` |
| `ERR` | `GuardarError`, `getLogErrorPorID` | Escribe y lee el log de errores con el mismo patrón | `getSiguienteIDLogErrores` | `INSERT` y `SELECT` sobre `tbLogErrores` | `ERR:12-85`; `Recordset` en `ERR:23`, `Max` en `ERR:49`, SQL en `ERR:73` |
| `ERR` | `GuardarDesdeCondorError` | Convierte una excepción de la aplicación en fila de log | `Now()`, `GuardarError` | `INSERT` en `tbLogErrores` | `ERR:87-106`; marca temporal en `ERR:91` |
| `COR` | `GuardarCorreo` | Inserta el correo enviado con clave calculada | `getSiguienteIDCorreo` | `INSERT` en `tbCorreosEnviados` | `COR:40-90`; `Recordset` en `COR:50`, `Max` en `COR:18` |
| `NC` | `getTodas`, `getPorExpediente` | Lista no conformidades, con lista de columnas explícita | `RepositorioComun`, `getdb()` | `SELECT` sobre `TbNoConformidades` | `NC:10-58`; SQL en `NC:17` y `NC:40` |
| `NC` | `getPorCodigoCondor` | Devuelve la no conformidad que ya tiene una concesión asociada | `getdb()` | `SELECT` sobre `TbNoConformidades` | `NC:59-94`; SQL en `NC:69`, hidratación con `Nz` en `NC:81-84` |
| `NC` | `existeEnBaseDatosExternaa` | Comprueba la existencia por clave con un `QueryDef` y un parámetro posicional | `CreateQueryDef` | Sin escritura | `NC:96-124`; `COUNT(*)` en `NC:104` |
| `NC` | `actualizarCodigoConcesion`, `limpiarCodigoConcesion` | Enlaza o desenlaza la concesión en la base externa | `RepositorioComun`, conexión externa | `UPDATE` sobre `TbNoConformidades` | `NC:125-233`; `UPDATE` en `NC:132` |
| `VR` | `GetUltimoOrdinal` | Devuelve el ordinal más alto de la solicitud | `RC` | `SELECT` sobre `tbValidacionRevision` | `VR:6-46`; SQL en `VR:21`, `Nz` en `VR:29` |
| `VR` | `GuardarHitoEnvio`, `GuardarHitoBorrador`, `CrearNuevoCicloValidacion` | Abre un ciclo con su ordinal y registra los hitos de envío y de borrador con su hash | `RC` | `INSERT` y `UPDATE` sobre `tbValidacionRevision` | `VR:274-379`, `VR:568-606`; alta en `VR:274` y `VR:330` |
| `VR` | `GetCicloAbierto`, `GetUltimoPendienteId`, `GetUltimoEstadoValidacion` | Localizan el ciclo en curso o el último pendiente de la solicitud | `RC` | `SELECT` sobre `tbValidacionRevision` | `VR:176-273`, `VR:522-567` |
| `VR` | `CerrarHitoRespuesta`, `ActualizarRespuestaPorId`, `GetVersionActualBorrador`, `GetHashVersionAnterior` | Cierran el ciclo con su resultado y versionan el hash de datos | `RC` | `UPDATE` sobre `tbValidacionRevision` | `VR:380-521`; cierre en `VR:417`, hash anterior en `VR:380` |
| `VR` | `GetUltimoCicloCerrado` | Devuelve el ciclo más reciente con su resultado y su hash | `RC` | `SELECT` sobre `tbValidacionRevision` | `VR:48-99`; `TOP 1` en `VR:64`, campos en `VR:76-80` |
| `VR` | `GetHistorialPorSolicitud` | Devuelve el historial completo como `Recordset` | `RC` | `SELECT` sobre `tbValidacionRevision` | `VR:101-127`; SQL en `VR:115` |
| `VR` | `ActualizarHashRevision` | Fija el hash de datos de un ciclo concreto | `RC` | `UPDATE` sobre `tbValidacionRevision` | `VR:152-606`; SQL en `VR:163` |

## Flujos de trabajo

- **Normal:** un cambio de estado o de bloque escribe primero el log y luego el dato (`LOG:19`, `ERR:23`). La lectura del historial filtra por tabla, registro y fecha (`LOG:81`, `LOG:108-113`).
- **Alternativos:** `ValidacionRevision` numera sus ciclos por ordinal, no por autonumérico (`VR:21`, `VR:29`), de modo que el mismo ciclo se identifica por la pareja solicitud y ordinal (`VR:163`). `GuardarDesdeCondorError` es la vía que usa el manejador de errores de los formularios para dejar traza (`ERR:87-106`).
- **Errores:** los cinco módulos repiten el patrón `On Error GoTo Errores` con `CondorError` y contexto (`LOG:117-121`, `ERR:100-104`, `VR` en su etiqueta `Errores`). `NoConformidadRepositorio` propaga el fallo sin envolver cuando la conexión externa es la que falla, porque su firma ya recibe la conexión.

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| La clave de las tablas de log es `Max(id) + 1` | `getSiguienteID*` de los tres | `LOG:39`, `ERR:49`, `COR:18` |
| El borrado de cambios se limita a una tabla concreta | literal `tabla = 'tbSolicitudes'` | `LOG:81` |
| El ciclo de revisión se identifica por solicitud y ordinal | `WHERE idSolicitud = … AND ordinal = …` | `VR:163` |
| La no conformidad sevacía con `Nz` al hidratarse | `getPorCodigoCondor` | `NC:81-84` |
| El hash se guarda junto al resultado del ciclo | `GetUltimoCicloCerrado` | `VR:79` |

## Lecturas y escrituras

| Operación | Tabla | Base | Tipo | Transacción | Evidencia |
|---|---|---|---|---|---|
| Cambios | `tbLogCambios` | Condor | `Recordset` para escribir, SQL parametrizada para leer y borrar | Ninguna | `LOG:19`, `LOG:61`, `LOG:81` |
| Errores | `tbLogErrores` | Condor | `Recordset` y SQL parametrizada | Ninguna | `ERR:23`, `ERR:73` |
| Correos enviados | `tbCorreosEnviados` | Condor | `Recordset` | Ninguna | `COR:18`, `COR:50` |
| No conformidades | `TbNoConformidades` | No Conformidades | SQL concatenada con lista de columnas y `QueryDef` posicional | Ninguna | `NC:17`, `NC:40`, `NC:69`, `NC:104`, `NC:132` |
| Validación de revisión | `tbValidacionRevision` | Condor | SQL parametrizada | Ninguna | `VR:21`, `VR:64`, `VR:115`, `VR:163` |

## Coexistencia

Cuatro puntos. `Max(id) + 1` en las tres tablas de auditoría (`LOG:39`, `ERR:49`, `COR:18`) debe pasar a secuencia o identidad, igual que en los repositorios de datos. `TOP 1 … ORDER BY … DESC` (`VR:64`, `LOG:110` de la hoja de catálogos) se traduce a `LIMIT 1` conservando el orden. El `COUNT(*)` con parámetro posicional `?` (`NC:104`) es un `QueryDef` en memoria, sin equivalente en el destino. Y `NoConformidadRepositorio` opera sobre una base externa con conexión propia (`NC:125-233`): en el esquema destino esa tabla debe existir como tabla enlazada o replicada, y esa decisión condiciona la coherencia entre No Conformidades y Condor.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual.** Los tres repositorios de auditoría escriben con `Recordset` y calculan la clave a mano. Los dos repositorios de dominio externo —no conformidades y validación de revisión— leen catálogos y escriben con SQL parametrizada, salvo la actualización del código de concesión, que concatena la lista de columnas.

**Posibles defectos, a verificar.**
1. `Max + 1` en las tres tablas de auditoría no es seguro con dos sesiones concurrentes (`LOG:39`, `ERR:49`, `COR:18`).
2. `GetCambiosPorIntervalo` y `GetHistorialPorSolicitud` devuelven un `DAO.Recordset` abierto (`LOG:113`, `VR:101`).
3. El borrado de cambios depende de un literal de tabla embebido en la SQL (`LOG:81`): el repositorio está atado a `tbSolicitudes`.
4. `existeEnBaseDatosExternaa` usa un parámetro posicional `?` en lugar del diccionario de parámetros del resto del módulo (`NC:104`).
5. La no conformidad se escribe en la base externa con una conexión que puede no ser la transaccional del servicio (`NC:125-233`).

**Propuestas web pendientes de decisión.** Que las tablas de auditoría compartan generador de clave. Que el historial se devuelva como lista y no como cursor. Que la limpieza de cambios por solicitud sea una operación por dominio y no por nombre de tabla. Que el vínculo con No Conformidades se resuelva por replication o por consulta remota.

## Criterios de aceptación

1. Un cambio de estado queda registrado con usuario, tabla y registro. Trazado a `LOG:9-32`.
2. El historial de cambios de una solicitud se recupera acotado por fecha. Trazado a `LOG:96-123`.
3. Un `CondorError` se convierte en una fila de log con su marca temporal. Trazado a `ERR:87-106`.
4. El ciclo de validación de revisión más reciente se obtiene con su ordinal y su hash. Trazado a `VR:48-99`.
5. Vincular y desvincular una concesión de una no conformidad deja el valor a nulo en el segundo caso. Trazado a `NC:125-233`.

## Pendientes

- Confirmar contra el binario qué base de datos usa la conexión externa de `NoConformidadRepositorio` y si es la misma que usa el formulario de alta de solicitudes al elegir la no conformidad.
- Confirmar si `hashDatos` se valida en lectura o solo se guarda (`VR:79`, `VR:163`).
- Confirmar qué consumidores recorren los `Recordset` de `LOG` y `VR` en lugar de copiar las filas.
- **Confirmar si `ValidacionRevisionRepositorio` deja procedimientos públicos sin uso desde algún formulario: los dieciséis están inventariados arriba, pero solo cinco aparecen en las hojas de formularios ya escritas.**