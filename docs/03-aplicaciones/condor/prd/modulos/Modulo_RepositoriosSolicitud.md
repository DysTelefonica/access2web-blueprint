# Módulos de repositorio de solicitud, expediente, adjunto y rechazo

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/modules/SolicitudRepositorio.bas` (217 líneas), `src/modules/ExpedienteRepositorio.bas` (107), `src/modules/AdjuntoRepositorio.bas` (221), `src/modules/RechazoRepositorio.bas` (326) · **PRD:** [README](../README.md)

Cuatro repositorios de apoyo que no guardan bloques de datos de la solicitud, sino sus objetos transversal: la solicitud, su expediente, los adjuntos por etapa del flujo y el rechazo. Comparten el patrón de los repositorios de datos (conexión opcional del servicio, `CondorError` en la etiqueta de errores, SQL parametrizada en unos y concatenada en otros) pero divergen en cómo calculan identificadores y cómo abren transacción: solo `RechazoRepositorio` abre la suya.

Abreviaturas: `SOL`, `modules/SolicitudRepositorio.bas`; `EXP`, `modules/ExpedienteRepositorio.bas`; `ADJ`, `modules/AdjuntoRepositorio.bas`; `RECH`, `modules/RechazoRepositorio.bas`; `R`, `modules/RepositorioComun.bas`; `CE`, `classes/CondorError.cls`; `WVC`, `modules/Variables Globales.bas`.

## Superficie pública

| Objeto | Tipo | Procedimientos públicos | Qué resuelve |
|---|---|---|---|
| `SolicitudRepositorio` | módulo | `getSolicitudPorID`, `getSolicitudPorCodigo`, `GuardarSolicitud`, `getSiguienteOrdinalParaClave`, `ExisteCodigo`, `ActualizarEstado`, `Eliminar` (`SOL:9-217`) | Alta, lectura y transición de `tbSolicitudes` |
| `ExpedienteRepositorio` | módulo | `getExpedientePorID`, `getSuministradoresPorExpediente`, `GetRacAsignadoPorExpediente` (`EXP:10-107`) | Lectura del expediente **de otra base de datos** (Lanzadera) |
| `AdjuntoRepositorio` | módulo | `Guardar`, `ExisteAdjuntoEtapa`, `getPorEtapa`, `getPorIdSolicitud`, `getPorID`, `Eliminar`, `EliminarPorIdSolicitud` (`ADJ:6-221`) | Adjuntos de una solicitud, con una fila por etapa del flujo |
| `RechazoRepositorio` | módulo | `GetUltimoRechazoActivo`, `GetUltimoRechazo`, `GetFechaUltimoRechazo`, `GuardarRechazo`, `DesactivarRechazosPrevios`, `EliminarRechazosPorSolicitud`, `ActualizarComentarioRechazoActivo`, `ActualizarJSON`, `GetHistorialPorSolicitud` (`RECH:11-326`) | Rechazos: historial, baja del anterior y alta del nuevo |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| `SOL` | `getSolicitudPorID`, `getSolicitudPorCodigo` | Hidratan la solicitud por clave o por código | `R.EjecutarAccion` o `QueryDef` | `SELECT` sobre `tbSolicitudes` | `SOL:9-52`; SQL parametrizada en `SOL:18` y `SOL:39` |
| `SOL` | `GuardarSolicitud` | Inserta o actualiza la solicitud por `idSolicitud` | `CreateQueryDef` en el modo lectura previa | `INSERT`/`UPDATE` sobre `tbSolicitudes` | `SOL:54-94`; `QueryDef` en `SOL:65-67`, `Recordset` de tabla en `SOL:75` |
| `SOL` | `getSiguienteOrdinalParaClave` | Calcula el ordinal siguiente a partir del sufijo del código | `CreateQueryDef` | Sin escritura | `SOL:96-128`; SQL en `SOL:107-109`, `Nz` en `SOL:114` |
| `SOL` | `ExisteCodigo` | Cuenta las filas con ese código | `CreateQueryDef` | Sin escritura | `SOL:130-160`; `COUNT(*)` en `SOL:141-145` |
| `SOL` | `ActualizarEstado` | Actualiza estado interno, fecha y usuario | `R.EjecutarAccion` | `UPDATE` sobre `tbSolicitudes` | `SOL:162-195`; SQL en `SOL:173-177`, ejecución en `SOL:186` |
| `EXP` | `getExpedientePorID` | Lee la cabecera del expediente con lista de columnas explícita | `getdb()` | `SELECT` sobre `tbExpedientes` | `EXP:10-50`; SQL construida por concatenación en `EXP:27` |
| `EXP` | `getSuministradoresPorExpediente` | Devuelve un diccionario de suministradores por expediente | `getdb()` | `SELECT` con `JOIN` a `tbExpedientesSuministradores` | `EXP:52-77`; SQL en `EXP:58` |
| `EXP` | `GetRacAsignadoPorExpediente` | Devuelve el nombre del RAC asignado | `getdb()` | `SELECT FIRST(TbRACS.RAC)` sobre `tbExpedientesRACs` | `EXP:79-107`; `FIRST` en `EXP:87-91` |
| `ADJ` | `Guardar` | Inserta la fila de adjunto con clave calculada | `getSiguienteIDAdjunto`, `R.EjecutarAccion` | `INSERT` sobre `tbAdjuntos` | `ADJ:6-35`; clave en `ADJ:12`, `INSERT` parametrizada en `ADJ:14-27` |
| `ADJ` | `ExisteAdjuntoEtapa`, `getPorEtapa` | Comprueba y devuelve el adjunto de una etapa | `getdb()` | `COUNT(*)` y `SELECT TOP 1` sobre `tbAdjuntos` | `ADJ:61-113`; `COUNT` en `ADJ:69`, `TOP 1` en `ADJ:99` |
| `ADJ` | `getPorIdSolicitud`, `getPorID` | Devuelven una `Collection` y una entidad | `R.HidratarEntidadDesdeSQL` | `SELECT` sobre `tbAdjuntos` | `ADJ:114-176`; SQL en `ADJ:128` y `ADJ:162` |
| `ADJ` | `Eliminar`, `EliminarPorIdSolicitud` | Borran por clave o por solicitud | `R.EjecutarAccion` | `DELETE` sobre `tbAdjuntos` | `ADJ:177-221`; SQL en `ADJ:183` y `ADJ:204` |
| `RECH` | `GetUltimoRechazoActivo`, `GetUltimoRechazo`, `GetFechaUltimoRechazo` | Leen el rechazo más reciente, con y sin filtro de activo | `OpenRecordset` en `dbOpenSnapshot` | `SELECT TOP 1` sobre `tbRechazos` | `RECH:11-132`; `TOP 1` en `RECH:23` y `RECH:69`, hidratación con `Nz` en `RECH:32-37` |
| `RECH` | `GuardarRechazo` | Da de baja los rechazos anteriores y guarda el nuevo | `DesactivarRechazosPrevios` | `UPDATE` e `INSERT` sobre `tbRechazos` | `RECH:134-168`; llamada en `RECH:146` |
| `RECH` | `DesactivarRechazosPrevios` | Marca como inactivos los rechazos de la solicitud | `db.Execute` | `UPDATE … SET esActivo = False` | `RECH:170-188`; SQL concatenada en `RECH:179`, ejecución en `RECH:180` |
| `RECH` | `ActualizarComentarioRechazoActivo` | Escribe el comentario sin truncar a 255 caracteres | `getWorkspace`, `ws.BeginTrans` | `UPDATE` del campo `comentarios` (Memo) | `RECH:209-252`; transacción propia en `RECH:219-223`, escritura en `RECH:229-233`, `CommitTrans` en `RECH:238` |
| `RECH` | `ActualizarJSON` | Guarda el JSON del rechazo con el mismo patrón transaccional | `getWorkspace`, `ws.BeginTrans` | `UPDATE` de un campo de texto largo | `RECH:254-298`; `BeginTrans` en `RECH:268`, `CommitTrans` en `RECH:283` |
| `RECH` | `GetHistorialPorSolicitud` | Devuelve un `DAO.Recordset` abierto al llamador | `getdb()` | `SELECT` sobre `tbRechazos` | `RECH:300-326` |

## Flujos de trabajo

- **Normal — alta de una solicitud:** el servicio pide el ordinal (`SOL:96`), compone el código, comprueba que no exista (`SOL:130`) y guarda (`SOL:54`). El ordinal se extrae del propio código almacenado con `Max(Val(Mid(...)))` sobre la parte posterior al último guion (`SOL:107-109`).
- **Normal — rechazo:** `GuardarRechazo` desactiva los rechazos previos de la misma solicitud (`RECH:146`) y guarda el nuevo. La lectura posterior filtra por `esActivo` (`RECH:23`).
- **Alternativos — escritura sin servicio:** `ActualizarComentarioRechazoActivo` y `ActualizarJSON` detectan que no se les pasó conexión, abren una `Workspace`, reabren el backend con su propia cadena de conexión y gestionan la transacción ellos mismos (`RECH:219-223`, `RECH:268`). El resto de procedimientos del módulo, si no reciben conexión, caen en `getdb()` sin transacción.
- **Errores:** todos los procedimientos lanzan `CondorError` con el nombre del procedimiento (`RECH:51`, `SOL:193`, `ADJ:33`). Los dos procedimientos transaccionales revierten con `ws.Rollback` antes de relanzar (`RECH:248`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| La clave de adjunto es `MAX(idAdjunto) + 1` | `getSiguienteIDAdjunto` | `ADJ:45-47` |
| El ordinal sale del sufijo numérico del código | `getSiguienteOrdinalParaClave` | `SOL:107-114` |
| Solo hay un rechazo activo por solicitud | `DesactivarRechazosPrevios` antes de insertar | `RECH:146`, `RECH:179` |
| Un adjunto por etapa, el último gana | `getPorEtapa` ordena por fecha de subida | `ADJ:99` |
| El comentario se escribe con DAO, no con SQL, para no truncar | `ActualizarComentarioRechazoActivo` | `RECH:210-212`, `RECH:232` |

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Alta y baja de solicitudes | `tbSolicitudes` | `QueryDef` parametrizado y `Recordset` de tabla | Ninguna en el módulo | `SOL:65-75`, `SOL:173-186`, `SOL:203` |
| Lectura del expediente y sus suministradores | `tbExpedientes`, `tbExpedientesSuministradores` | SQL concatenada | Ninguna | `EXP:27`, `EXP:58` |
| Lectura del RAC asignado | `tbExpedientesRACs`, `TbRACS` | SQL concatenada con `FIRST` | Ninguna | `EXP:87-91` |
| Adjuntos | `tbAdjuntos` | `INSERT`/`SELECT`/`DELETE` parametrizados, `MAX` en SQL | Ninguna | `ADJ:14-27`, `ADJ:45`, `ADJ:183`, `ADJ:204` |
| Rechazos | `tbRechazos` | `TOP 1` concatenado en lectura; `UPDATE`/`DELETE` concatenados en escritura | Propia en `ActualizarComentarioRechazoActivo` y `ActualizarJSON` | `RECH:23`, `RECH:179`, `RECH:198`, `RECH:223`, `RECH:268` |

## Coexistencia

Siete adaptaciones. La cadena de conexión del backend se construye con la contraseña en claro y el tipo de proveedor `MS Access` (`RECH:222`, `RECH:267`): con PostgreSQL desaparece el parámetro de contraseña y la apertura pasa a ser una cadena de conexión propia del vínculo. `Max(Val(Mid(codigoSolicitud, InStrRev(codigoSolicitud, '-') + 1)))` (`SOL:107-109`) depende de `InStrRev`, `Val`, `Mid` y de `Max` sobre texto, que es ACE: en PostgreSQL hay que decidir si el ordinal pasa a ser una columna numérica con su propio índice. `FIRST` (`EXP:87`), `TOP 1` (`ADJ:99`, `RECH:23`, `RECH:69`) y `esActivo = False` en SQL de Access (`RECH:179`) se traducen a `LIMIT`, `limit` y a un literal booleano respectively, o a una columna de fecha de baja, si se decide. `Nz` en la hidratación (`RECH:32-37`) convierte nulos en cadena vacía al construir el objeto: es una diferencia semántica entre `NULL` y `''` que hay que decidir por columna. La transacción por `Workspace` (`RECH:221-223`) no tiene equivalente: en PostgreSQL la transacción la abre la capa de servicio. Por último, `GetHistorialPorSolicitud` devuelve un `Recordset` de DAO abierto (`RECH:300`), un tipo que no existe en el destino.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual.** Los cuatro módulos resuelven una consulta o una acción por procedimiento, sin estado propio. La conexión la aporta el servicio; cuando no la hay, cada módulo decide por su cuenta.

**Posibles defectos, a verificar.**
1. `MAX + 1` en adjuntos (`ADJ:45-47`) y el ordinal derivado del código (`SOL:107-114`) no son seguros con dos sesiones concurrentes.
2. `DesactivarRechazosPrevios` concatena el identificador (`RECH:179`) y `EliminarRechazosPorSolicitud` también (`RECH:198`): es SQL dinámico, hoy acotado por el tipo `Long`.
3. La contraseña del backend se compone en una cadena de conexión en el propio repositorio (`RECH:222`, `RECH:267`).
4. `GetHistorialPorSolicitud` expone un `DAO.Recordset` al llamador (`RECH:300`): el contrato del repositorio incluye el tipo de la base de datos.
5. `ActualizarComentarioRechazoActivo` y `ActualizarJSON` abren su propia transacción (`RECH:219-223`, `RECH:268`) mientras el resto del módulo no lo hace: la atomicidad depende de quién llama.
6. `ExpedienteRepositorio` no acepta conexión como parámetro (`EXP:10`, `EXP:52`, `EXP:79`): siempre usa `getdb()`, y por tanto no participa en la transacción del servicio aunque el expediente se modifique en la misma unidad de trabajo.

**Propuestas web pendientes de decisión.** Que el ordinal de la solicitud sea una columna numérica con restricción de unicidad por expediente. Que el rechazo activo se modele con fecha de baja en lugar de un booleano. Que el historial se devuelva como una lista de objetos y no como un cursor. Que la conexión al backend deje de construirse en el repositorio y pase al servicio o a un proveedor único.

## Criterios de aceptación

1. El código de una solicitud es único por expediente y el ordinal avanza sin huecos bajo concurrencia. Trazado a `SOL:96-128` y `SOL:130-160`.
2. Un rechazo nuevo invalida el anterior y ambos quedan en el historial. Trazado a `RECH:146` y `RECH:179`.
3. Guardar un comentario largo no lo trunca. Trazado a `RECH:229-233`.
4. La lectura del expediente devuelve cabecera, suministradores y RAC sin dejar cursores abiertos. Trazado a `EXP:10-107`.
5. Un fallo en una escritura deja la solicitud sin cambios parciales. Trazado a `SOL:191-195` y a la transacción del servicio.

## Pendientes

- Confirmar contra el binario qué base de datos usa `getdb()` en cada punto y cuál es la ruta: `EXP` lee de otro backend (el expediente vive en Lanzadera según la hoja de coexistencias de la aplicación) y el código no lo distingue.
- Confirmar qué consumidores de `GetHistorialPorSolicitud` usan el `Recordset` directamente en lugar de copiarlo.
- Confirmar si `GuardarSolicitud` mantiene la rama `QueryDef` (`SOL:65-67`) o si es código muerto tras el refactor a `Recordset`.
- Confirmar si `getUltimoRechazo` (sin filtro de activo) se usa en algún flujo: la pareja activa/no activa sugiere una transición que podría quedar a medias.