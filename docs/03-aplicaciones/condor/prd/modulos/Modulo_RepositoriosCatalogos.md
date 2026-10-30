# Módulos de repositorio de catálogos y de identidad

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/modules/UsuarioRepositorio.bas` (66 líneas), `AplicacionRepositorio.bas` (60), `EstadoRepositorio.bas` (27), `WorkflowRepositorio.bas` (58), `SuministradorRepositorio.bas` (28), `MapeoRepositorio.bas` (26), `LogEstadoRepositorio.bas` (134) · **PRD:** [README](../README.md)

Siete módulos pequeños que solo leen catálogos, identidad y metadatos. Dos de ellos —`UsuarioRepositorio` y `AplicacionRepositorio`— no leen del backend de Condor sino del de Lanzadera (`getdbLanzadera()`), que es el vínculo entre las dos aplicaciones y el punto donde las reglas compartidas hay que decidir. Los otros cinco se quedan en el backend propio; dos de ellos, `MapeoRepositorio` y `SuministradorRepositorio`, no aceptan conexión: siempre usan `getdb()`.

Abreviaturas: `USR`, `modules/UsuarioRepositorio.bas`; `APP`, `modules/AplicacionRepositorio.bas`; `EST`, `modules/EstadoRepositorio.bas`; `WF`, `modules/WorkflowRepositorio.bas`; `SUM`, `modules/SuministradorRepositorio.bas`; `MAP`, `modules/MapeoRepositorio.bas`; `LOG`, `modules/LogEstadoRepositorio.bas`; `R`, `modules/RepositorioComun.bas`; `UAP`, `classes/UsuarioAplicacionPermisos.cls`; `VG`, `modules/Variables Globales.bas`.

## Superficie pública

| Objeto | Tipo | Procedimientos públicos | Qué resuelve |
|---|---|---|---|
| `UsuarioRepositorio` | módulo | `getUsuario`, `getUsuarioConectadoPorMaquina` (`USR:9-66`) | Usuario de la tabla de usuarios de **Lanzadera**, con cuatro criterios de búsqueda |
| `AplicacionRepositorio` | módulo | `getAplicacionesPermisos` (`APP:9-60`) | Aplicaciones y permisos de un usuario, leídos de **Lanzadera** |
| `EstadoRepositorio` | módulo | `getTodosLosEstados` (`EST:10-27`) | Catálogo de estados del flujo, ordenado |
| `WorkflowRepositorio` | módulo | `getPosiblesDestinos` (`WF:13-58`) | Transiciones posibles desde un estado, con su rol requerido |
| `SuministradorRepositorio` | módulo | `getSuministradorPorID` (`SUM:10-28`) | Suministrador por identificador |
| `MapeoRepositorio` | módulo | `getMapeoParaPlantilla` (`MAP:10-26`) | Mapeo de campos asociado a una plantilla |
| `LogEstadoRepositorio` | módulo | `Guardar`, `getHistorialPorIdSolicitud`, `EliminarPorIdSolicitud`, `getUltimoEstadoAnterior` (`LOG:10-134`) | Historial de transiciones de estado de una solicitud |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| `USR` | `getUsuario` | Devuelve el primer usuario que cumple los criterios no vacíos (identificador, usuario de red, nombre o correo) | `R.HidratarEntidadDesdeSQL`, `getdbLanzadera()` | `SELECT` sobre `TbUsuariosAplicaciones` | `USR:9-50`; filtros en `USR:22-37`, SQL en `USR:41`, hidratación en `USR:42` |
| `USR` | `getUsuarioConectadoPorMaquina` | Traduce el usuario de Windows en un registro de usuario | `CreateObject("Wscript.Network")`, `getUsuario` | Sin escritura | `USR:52-66`; lectura del entorno en `USR:56-57` |
| `APP` | `getAplicacionesPermisos` | Devuelve un diccionario aplicación → permisos | `R.EjecutarConsulta`, `getdbLanzadera()`, `UAP.SetPropiedad` | `SELECT` sobre `TbUsuariosAplicacionesPermisos` | `APP:9-60`; SQL en `APP:24-27`, hidratación reflectiva en `APP:34-38` |
| `EST` | `getTodosLosEstados` | Lista el catálogo de estados como diccionario | `R.EjecutarConsulta` | `SELECT` sobre `tbEstados` | `EST:10-27`; SQL en `EST:18` |
| `WF` | `getPosiblesDestinos` | Devuelve los destinos posibles desde un estado con su rol | `R.EjecutarConsulta` | `SELECT` sobre `tbTransiciones` | `WF:13-58`; SQL en `WF:29` |
| `SUM` | `getSuministradorPorID` | Devuelve el suministrador por clave | `R.HidratarEntidadDesdeSQL`, `getdb()` | `SELECT` sobre `TbSuministradores` | `SUM:10-28`; SQL en `SUM:16` |
| `MAP` | `getMapeoParaPlantilla` | Devuelve el mapeo de campos de una plantilla | `R.HidratarColeccionDesdeSQL` | `SELECT` sobre `tbMapeoCampos` | `MAP:10-26`; SQL en `MAP:15`, hidratación en `MAP:19` |
| `LOG` | `Guardar` | Inserta la transición de estado | `OpenRecordset` en `dbOpenDynaset` sobre la tabla | `INSERT` en `tbLogEstados` | `LOG:10-44`; conexión en `LOG:17`, apertura en `LOG:22` |
| `LOG` | `getHistorialPorIdSolicitud` | Devuelve el historial como diccionario | `R.EjecutarConsulta` | `SELECT` sobre `tbLogEstados` | `LOG:46-74`; SQL en `LOG:56` |
| `LOG` | `getUltimoEstadoAnterior` | Devuelve el estado del que se salió en la última transición, o 0 | `R.EjecutarConsulta` | Sin escritura | `LOG:96-134`; `TOP 1` en `LOG:110-111`, `Nz` en `LOG:119`, valor cero en `LOG:122` |
| `LOG` | `EliminarPorIdSolicitud` | Borra el historial de una solicitud | `R.EjecutarAccion` | `DELETE` sobre `tbLogEstados` | `LOG:76-94`; SQL en `LOG:82` |

## Flujos de trabajo

- **Normal:** el servicio resuelve primero los catálogos (`EST:18`, `WF:29`) para pintar la interfaz, y luego escribe cada transición en `tbLogEstados` (`LOG:22`). La lectura del historial ordena por fecha ascendente (`LOG:56`) y la del estado anterior, por fecha descendente y se queda con la primera (`LOG:110-111`).
- **Alternativos:** `getUsuario` admite cuatro criterios y solo combina los que llegan informed; si ninguno viene informado, sale sin devolver nada (`USR:39`). Si no hay historial, `getUltimoEstadoAnterior` devuelve 0 en lugar de error (`LOG:120-122`).
- **Errores:** cada módulo envuelve el fallo en `CondorError` con el nombre del procedimiento (`USR:46-49`, `APP` en su etiqueta `Errores`, `LOG` en su etiqueta `Errores`). `getUsuarioConectadoPorMaquina` depende de un objeto COM del sistema y no captura el caso de estación sin ese componente más allá del `On Error`.

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Los permisos se indexan por aplicación y se ignoran duplicados | `getAplicacionesPermisos` | `APP:29`, `APP:40-41` |
| La comparación de claves es insensible a mayúsculas | `colPermisos.CompareMode = TextCompare` | `APP:29` |
| Los nulos se convierten en cadena vacía al hidratar por reflexión | `SetPropiedad` con `Nz` | `APP:37` |
| El historial se ordena por fecha de transición, no por clave | `getHistorialPorIdSolicitud` | `LOG:56` |
| Sin historial, el estado anterior es 0 | `getUltimoEstadoAnterior` | `LOG:118-122` |

## Lecturas y escrituras

| Operación | Tabla | Base de datos | Tipo | Transacción | Evidencia |
|---|---|---|---|---|---|
| Usuario de aplicación | `TbUsuariosAplicaciones` | Lanzadera | SQL con `WHERE` construido por unión de condiciones | Ninguna | `USR:41-42` |
| Usuario de red del puesto | — | — | Variable de entorno de Windows | — | `USR:56-57` |
| Aplicaciones y permisos | `TbUsuariosAplicacionesPermisos` | Lanzadera | SQL parametrizada | Ninguna | `APP:24-27` |
| Catálogos de estados y transiciones | `tbEstados`, `tbTransiciones` | Condor | SQL parametrizada | Ninguna | `EST:18`, `WF:29` |
| Suministrador y mapeo de plantilla | `TbSuministradores`, `tbMapeoCampos` | Condor | SQL parametrizada | Ninguna | `SUM:16`, `MAP:15` |
| Historial de estados | `tbLogEstados` | Condor | `Recordset` para escribir, SQL parametrizada para leer y borrar | Ninguna | `LOG:22`, `LOG:56`, `LOG:82` |

## Coexistencia

Estos siete módulos concentran tres decisiones que el resto del PRD no puede cerrar por su cuenta. La primera es el vínculo entre aplicaciones: `USR` y `APP` leen de la base de **Lanzadera**, no de la de Condor (`USR:42`, `APP:27`), así que el esquema destino debe decidir si las tablas de identidad se replican, se consultan en línea o se sustituyen por un proveedor de identidad común. La segunda es la identidad: `getUsuarioConectadoPorMaquina` toma el usuario de Windows mediante COM (`USR:56`), un acoplamiento que en web se resuelve con la sesión del navegador, nunca con el nombre de la estación. La tercera es la construcción dinámica del `WHERE`: `USR:41` concatena las condiciones con `Join`, de modo que el filtro depende de un diccionario en memoria y no de parámetros ligados; hay que decidir la lista cerrada de criterios que la API web aceptará.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual.** Los siete son de solo lectura salvo `LogEstadoRepositorio`, que escribe. Las tablas de identidad viven fuera del backend de Condor; los catálogos del flujo, dentro. `getdb()` y `getdbLanzadera()` conviven en el mismo repositorio de la aplicación.

**Posibles defectos, a verificar.**
1. `getUsuario` construye el `WHERE` por concatenación de un diccionario (`USR:41`): la seguridad depende de que ningún criterio llegue de fuera sin control.
2. `getSuministradorPorID` y `getMapeoParaPlantilla` no aceptan conexión (`SUM:10`, `MAP:10`): no pueden participar en la transacción del servicio.
3. `getUsuarioConectadoPorMaquina` depende de un objeto COM del sistema (`USR:56`).
4. La hidratación reflectiva de permisos (`APP:34-38`) puede fallar en silencio si el conjunto de columnas del recordset y el de `ColCampos` divergen.

**Propuestas web pendientes de decisión.** Que la identidad deixe de resolverse contra tablas enlazadas y pase al mecanismo de sesión que se apruebe para el conjunto del roadmap. Que el filtro de usuario sea una lista cerrada de criterios de API. Que los catálogos de estados y transiciones se lean desde una tabla del esquema de destino y dejen de duplicarse en el backend de cada aplicación.

## Criterios de aceptación

1. Los permisos de un usuario se resuelven con el mismo resultado que la consulta actual y sin distinguir mayúsculas. Trazado a `APP:29-42`.
2. El historial de una solicitud sale ordenado por fecha de transición ascendente. Trazado a `LOG:56`.
3. Una solicitud sin historial devuelve 0 como estado anterior, no un error. Trazado a `LOG:118-122`.
4. El usuario de la sesión se resuelve sin depender del nombre de la estación. Trazado a `USR:52-58`.
5. Borrar una solicitud elimina también su historial de estados. Trazado a `LOG:76-94`.

## Pendientes

- Confirmar contra el binario si `getdbLanzadera()` y `getdb()` apuntan a ficheros distintos y qué tablas comparten, más allá de `TbUsuariosAplicaciones` y `TbUsuariosAplicacionesPermisos` (`USR:41`, `APP:24`).
- Confirmar qué consumidores dependen de la reflexión `SetPropiedad` y `ColCampos` (`APP:34-37`), porque fija el contrato de la clase de permisos.
- Confirmar si `MapeoRepositorio` tiene consumidor vivo: el snapshot lo conservería aunque el flujo ya no lo use.