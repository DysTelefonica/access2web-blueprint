# Módulos de repositorio de datos de solicitud

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/modules/DatosPCRepositorio.bas` (545 líneas), `src/modules/DatosCDCARepositorio.bas` (455), `src/modules/DatosCDCASUBRepositorio.bas` (410) · **PRD:** [README](../README.md)

Tres repositorios gemelos que guardan los tres bloques de datos de una solicitud: Propuesta de Cambio (`tbDatosPC`), Concesión o Desviación (`tbDatosCDCA`) y concesión o desviación de subsuministrador (`tbDatosCDCASUB`). Los tres aplican el mismo patrón —abrir un `Recordset` por `idSolicitud`, insertar o actualizar, cerrar— y ninguno abre transacción: la reciben por parámetro opcional del servicio. Sus diferencias no son de estilo, sino de superficie (columnas que limpian, nombres de procedimiento) y están documentadas abajo.

Abreviaturas: `R`, `modules/RepositorioComun.bas`; `PC`, `DatosPCRepositorio.bas`; `CDCA`, `DatosCDCARepositorio.bas`; `SUB`, `DatosCDCASUBRepositorio.bas`; `CE`, `classes/CondorError.cls`; `Tab`, `Form_` + el nombre del contenedor, en `src/forms/`.

## Superficie pública

| Objeto | Tipo | Procedimientos públicos | Qué resuelve |
|---|---|---|---|
| `DatosPCRepositorio` | módulo | `Guardar`, `getPorIdSolicitud`, `ActualizarDatosGenerales`, `ActualizarPropuesta`, `ActualizarImpacto`, `ActualizarAprobacionSuministrador`, `ActualizarDictamenRAC`, `ActualizarDecisionFinal`, `EliminarPorIdSolicitud`, `LimpiarDictamenRAC`, `LimpiarAprobacionSuministrador`, `LimpiarDecisionFinal` (`PC:10-545`) | Persistencia del bloque `PC` de una solicitud |
| `DatosCDCARepositorio` | módulo | los mismos doce, con prefijo `idDatosCDCA` (`CDCA:13-455`) | Persistencia del bloque `CD_CA` |
| `DatosCDCASUBRepositorio` | módulo | `Guardar`, `getPorIdSolicitud`, `ActualizarDatosGenerales`, `ActualizarPropuesta`, `ActualizarImpacto`, `ActualizarAprobacion` (sin «Suministrador»), `ActualizarDictamenRAC`, `ActualizarDecisionFinal`, `EliminarPorIdSolicitud`, `LimpiarDictamenRAC`, `LimpiarAprobacionSuministrador`, `EliminarDecisionFinal` (`SUB:6-410`) | Persistencia del bloque `CD_CA_SUB` |
| — | privado | `getSiguienteIDDatosPC` (`PC:75`), `getSiguienteIDDatosCDCA` (`CDCA:13`), `getSiguienteIDDatosCDCASUB` (`SUB:6`) | Asignar la clave primaria por `Max + 1` |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| `PC` | `Guardar` | Upsert del objeto completo con `RellenarRecordsetDesdeObjeto`; omite la clave primaria en el modo actualización | `getdb()`, `R.RellenarRecordsetDesdeObjeto`, `getSiguienteIDDatosPC` | `INSERT` o `UPDATE` sobre `tbDatosPC` | `PC:10-50`; inserción `PC:25-32`, actualización `PC:35-37`, `rcd.Update` en `PC:40` |
| `PC` | `getPorIdSolicitud` | Hidrata la entidad desde una SQL parametrizada | `R.HidratarEntidadDesdeSQL` | `SELECT` sobre `tbDatosPC` | `PC:53-71`; SQL con marcador `[p_IDSol]` en `PC:60`, parámetros en `PC:61-62`, hidratación en `PC:64` |
| `PC` | `getSiguienteIDDatosPC` | Calcula la siguiente clave como máximo más uno | `OpenRecordset` sobre `Max(idDatosPC)` | Sin escritura | `PC:75-92`; `Max` en `PC:80`, `Nz` en `PC:83`, incremento en `PC:86` |
| `PC` | `ActualizarDatosGenerales` | Inserta o actualiza escribiendo campo a campo | `getdb()`, `R.RellenarRecordsetDesdeObjeto` | `INSERT`/`UPDATE` de las columnas de cabecera | `PC:105-161`; SQL concatenada en `PC:115-116`, escritura explícita en modo actualización en `PC:137-141` |
| `PC` | `ActualizarPropuesta`, `ActualizarDictamenRAC` | Igual que el anterior pero con `QueryDef` temporal parametrizado | `CreateQueryDef`, `qdf.Parameters` | `INSERT`/`UPDATE` | `PC:163-221` y `PC:338-383`; `QueryDef` en `PC:180-184` y `PC:396-398` |
| `PC` | `ActualizarImpacto`, `ActualizarAprobacionSuministrador`, `ActualizarDecisionFinal` | Variantes con SQL concatenada y sin `db` opcional en dos de ellas | `getdb()` | `INSERT`/`UPDATE` | `PC:222-292`, `PC:293-337`, `PC:384-424`; SQL en `PC:236-237`, `PC:302-303`, `PC:353-354` |
| `PC` | `EliminarPorIdSolicitud` | Borra la fila por `idSolicitud` con SQL parametrizada | `R.EjecutarAccion` | `DELETE` sobre `tbDatosPC` | `PC:425-447`; SQL en `PC:431`, ejecución en `PC:437` |
| `CDCA` | mismo conjunto | Réplica con `tbDatosCDCA` y con las firmas de aprobación propias del bloque | — | — | `CDCA:38-455`; `getPorIdSolicitud` en `CDCA:264-292`, `Guardar` en `CDCA:316-356` |
| `SUB` | mismo conjunto | Réplica con `tbDatosCDCASUB` | — | — | `SUB:28-410`; `Guardar` en `SUB:28-60`, `getPorIdSolicitud` en `SUB:62-78` |

## Flujos de trabajo

- **Normal:** el servicio abre la transacción, resuelve la conexión y la pasa como `db` opcional. `Guardar` usa esa conexión (`PC:18`: `If db Is Nothing Then Set dbConexion = getdb() Else Set dbConexion = db`), abre el `Recordset`, decide inserción o actualización por `rcd.EOF` (`PC:23`), vuelca el objeto y confirma (`PC:40`). Si el procedimiento se invoca sin conexión, `getdb()` abre una nueva y el cambio queda fuera de la transacción del servicio: el propio código lo documenta como «conexión dual» (`PC:391-392`, `CDCA:326`, `SUB:32`).
- **Alternativos:** `Guardar` de `CDCA` se conserva solo para el servicio de rechazo transaccional y el propio módulo advierte que no debe usarse desde el formulario de edición (`CDCA:318-320`). `SUB.Guardar` llama a `RellenarRecordsetDesdeObjeto` en los dos modos (`SUB:44` y `SUB:48`), mientras que `PC.Guardar` solo lo hace en inserción (`PC:32`) y en actualización omite la clave (`PC:37`).
- **Errores:** cada procedimiento salta a una etiqueta `Errores` que construye un `CondorError` con el número y la descripción de VBA, añade contexto con `AddToCallStack` y lo vuelve a lanzar (`PC:46-50`, `PC:67-71`, `PC:447`). `ActualizarPropuesta` de `PC` genera además un error propio 513 cuando no existe fila para la solicitud (`PC:187-190`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| La clave primaria es el máximo de la tabla más uno | `getSiguienteID*` de los tres | `PC:80-86`, `CDCA:18`, `SUB:11` |
| Solo se asigna clave si el objeto llega a cero | `Guardar` | `PC:28-30` |
| En actualización, los campos se escriben con `Nz(valor, "")` | `Actualizar*` | `PC:137-141` |
| Los procedimientos de limpieza ponen `Null`, no cadena vacía | `Limpiar*` / `EliminarDecisionFinal` | `PC:464-470`, `PC:500-501`, `PC:529-531`; `CDCA:372-377`, `CDCA:407-411`, `CDCA:439-441`; `SUB:341-344`, `SUB:369-373`, `SUB:397-399` |
| La conexión es transaccional si el servicio la pasa |parameter `db` opcional | `PC:18`, `CDCA:96`, `SUB:32` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Upsert del bloque completo | `tbDatosPC` / `tbDatosCDCA` / `tbDatosCDCASUB` | SQL concatenada sobre `Recordset` | Ninguna en el módulo | `PC:20-21`, `CDCA:329-330`, `SUB:34-35` |
| Lectura por `idSolicitud` | las mismas tres | SQL parametrizada (`QueryDef`) o diccionario de parámetros | Ninguna | `PC:60-64`, `PC:177-184`, `CDCA:278-282`, `SUB:68-72` |
| Cálculo de clave | las mismas tres | `Max()` en SQL | Ninguna | `PC:80`, `CDCA:18`, `SUB:11` |
| Borrado por `idSolicitud` | las mismas tres | `DELETE` parametrizado | Ninguna (la del servicio) | `PC:431-437`, `CDCA:298-302`, `SUB:321-324` |

## Coexistencia

Cinco puntos a adaptar en PostgreSQL. Primero, `Max(id) + 1` (`PC:80-86`) no es seguro con dos sesiones concurrentes: la clave debe pasar a ser una secuencia o una identidad, y el valor debe dejar de calcularse en el cliente. Segundo, el SQL concatenado con el identificador (`PC:20`, `PC:115`, `PC:236`, `PC:302`, `PC:353`) debe pasar a parámetro ligado; hoy el tipo `Long` hace que el riesgo sea bajo, pero el patrón no es portable tal cual. Tercero, `Nz(valor, "")` al escribir (`PC:137-141`) convierte `Null` en cadena vacía, mientras que las limpiezas escriben `Null` (`PC:464-470`): en PostgreSQL esa diferencia entre `''` y `NULL` debe ser deliberada. Cuarto, `CreateQueryDef("", sql)` (`PC:180`, `PC:396`) crea una consulta efímera en memoria: no hay equivalente en el esquema de destino, así que la lectura parametrizada debe reescribirse con parámetros de verdad. Quinto, `Option Compare Database` y los nombres de tabla con minúsculas (`tbDatosPC`) obligan a decidir si el esquema destino conserva esa convención o la pasa a `snake_case`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual.** El bloque de datos se identifica por `idSolicitud`, no por su clave propia: hay como máximo una fila por solicitud y bloque, pero la tabla lo permite (`PC:20` filtra y ordena por `idDatosPC`). Los doce procedimientos por módulo son públicos salvo el cálculo de clave. La lectura y el borrado usan parámetros; la escritura concatena.

**Posibles defectos, a verificar.**
1. `Max + 1` puede devolver una clave ya usada si dos sesiones insertan a la vez (`PC:80-86`).
2. La conexión dual permite que una escritura quede fuera de la transacción del servicio si el llamador omite `db` (`PC:18`).
3. `Nz(valor, "")` al escribir y `Null` al limpiar producen el mismo campo con dos valores distintos según el camino (`PC:137-141` frente a `PC:529-531`).
4. `SUB` llama a `EliminarDecisionFinal` donde `PC` y `CDCA` llaman a `LimpiarDecisionFinal` (`SUB:387` frente a `PC:515` y `CDCA:425`): la superficie no es idéntica entre los tres módulos y conviene confirmarlo contra el binario.
5. Los comentarios del módulo hablan de acciones pendientes de refactor (`PC:6-8`, `CDCA:316-320`): el fichero mezcla código y nota de trabajo.

**Propuestas web pendientes de decisión.** Que el identificador de la solicitud sea la clave del agregado en el esquema destino. Que la limpieza de un bloque sea una operación idempotente y explícita en lugar de un `UPDATE` a `Null` campo a campo. Que un mismo servicio atenga los tres bloques con una tabla de destino común o con tres tablas: es una decisión de modelo, no de código.

## Criterios de aceptación

1. Guardar un bloque nuevo crea exactamente una fila y devuelve su clave. Trazado a `PC:23-40`.
2. Guardar un bloque existente no crea fila nueva ni cambia su clave. Trazado a `PC:35-40`.
3. Leer por `idSolicitud` devuelve el bloque y no el de otra solicitud. Trazado a `PC:60-64`.
4. Limpiar un bloque deja todos sus campos en el estado nulo declarado, sin cadena vacía. Trazado a `PC:464-470` y `PC:529-531`.
5. Un fallo al guardar propaga el error con el contexto de la solicitud y no deja la fila a medias. Trazado a `PC:46-50` y al control de errores del servicio.

## Pendientes

- Confirmar contra el binario si las tres tablas declaran autonumérico real en `tbDatosPC.idDatosPC`, `tbDatosCDCA.idDatosCDCA` y `tbDatosCDCASUB.idDatosCDCASUB`. El código calcula la clave a mano, lo que sugiere que no lo hay, pero los metadatos no están en el repositorio (D-CON-5).
- Confirmar qué consumidores invocan `Guardar` sin `db` (`CDCA:318-320` lo desaconseja explícitamente).
- Confirmar si `ActualizarPropuesta` y `ActualizarDictamenRAC` de `PC` se usan en producción: la ruta con `QueryDef` es la única que no concatena (`PC:177-184`, `PC:394-398`).