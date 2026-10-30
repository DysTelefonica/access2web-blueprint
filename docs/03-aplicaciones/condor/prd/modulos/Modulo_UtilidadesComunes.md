# Módulos de utilidad común, enumeradores y variables globales

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/modules/RepositorioComun.bas` (201 líneas), `Variables Globales.bas` (177), `HashHelper.bas` (82), `ChecklistHelper.bas` (70), `Factoria.bas` (65), `modActualizaciones.bas` (58), `modEnumeradores.bas` (55), `ResultadoImportacionRAC.bas` (10), `Módulo1.bas` (4) · **PRD:** [README](../README.md)

La capa que el resto de los repositorios da por supuesta: ejecución de SQL con parámetros, hidratación de objetos, instanciación por nombre, hashing, catálogos de estados en código y el estado global de la sesión (conexiones y usuarios). `RepositorioComun` es la pieza de la que dependen casi todas las escrituras del PRD: todos los `INSERT`, `UPDATE`, `DELETE` y todas las hidrataciones terminan aquí.

Abreviaturas: `RC`, `modules/RepositorioComun.bas`; `VG`, `modules/Variables Globales.bas`; `HH`, `modules/HashHelper.bas`; `CH`, `modules/ChecklistHelper.bas`; `FAC`, `modules/Factoria.bas`; `MA`, `modules/modActualizaciones.bas`; `ENUM`, `modules/modEnumeradores.bas`; `RIR`, `modules/ResultadoImportacionRAC.bas`; `CE`, `classes/CondorError.cls`.

## Superficie pública

| Objeto | Tipo | Procedimientos públicos | Qué resuelve |
|---|---|---|---|
| `RepositorioComun` | módulo | `EjecutarAccion`, `EjecutarConsulta`, `HidratarColeccionDesdeSQL`, `HidratarEntidadDesdeSQL`, `RellenarRecordsetDesdeObjeto` (`RC:7-201`) | Infraestructura común de acceso a datos: ejecutar, leer, hidratar y rellenar |
| `Factoria` | módulo | `CreateEntity` (`FAC:11-65`) | Instanciar una entidad por su nombre de clase |
| `HashHelper` | módulo | `CalcularSHA256` (`HH:39-82`) | Huella SHA-256 de un texto con la API criptográfica de Windows |
| `ChecklistHelper` | módulo | `GetItemsChecklistPorEstado` (`CH:36-70`), con cuatro funciones privadas por bloque (`CH:6-34`) | Lista de comprobación que se muestra en cada estado del flujo |
| `modEnumeradores` | módulo | seis `Enum`: `rol`, `EnumSiNo`, `EnumDecision`, `TipoGrafico`, `enumEstados`, `enumBloqueFormulario` (`ENUM:11-54`) | Vocabulario cerrado de la aplicación, con los valores fijados en código |
| `Variables Globales` | módulo | tres procedimientos y el estado global (`VG:12-177`) | Conexiones, usuarios activo y real, y constantes de la aplicación |
| `modActualizaciones` | módulo | `EjecutarActualizacion_Ambito`, `EjecutarActualizacion_Spec080_MapeoCDCASUB` (`MA:7-58`) | Scripts de actualización de esquema ejecutados a mano |
| `ResultadoImportacionRAC` | módulo con campos públicos (`RIR:6-10`) | — | Contenedor del resultado de una importación |
| `Módulo1` | módulo vacío (`Módulo1:1-4`) | — | Sin contenido |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| `RC` | `EjecutarAccion` | Ejecuta una sentencia de acción con parámetros ligados | `getdb()`, `CreateQueryDef`, `Execute` | `INSERT`, `UPDATE` o `DELETE` | `RC:7-37`; conexión en `RC:13`, error 513 si falta en `RC:14-18`, `QueryDef` en `RC:20` |
| `RC` | `EjecutarConsulta` | Ejecuta una consulta con parámetros y devuelve el `Recordset` | `getdb()`, `CreateQueryDef` | `SELECT` | `RC:39-67`; control de conexión en `RC:47` |
| `RC` | `HidratarColeccionDesdeSQL` | Rellena una `Collection` de objetos del tipo indicado | `RellenarObjetoDesdeRecordset` | Sin escritura | `RC:69-111` |
| `RC` | `HidratarEntidadDesdeSQL` | Devuelve una instancia de la clase indicada, o `Nothing` | `RellenarObjetoDesdeRecordset` | Sin escritura | `RC:113-143` |
| `RC` | `RellenarRecordsetDesdeObjeto` | Vuelca un objeto sobre un `Recordset`, con opción de omitir la clave primaria | — | `INSERT`/`UPDATE` pendiente | `RC:173-201` |
| `RC` | `RellenarObjetoDesdeRecordset` | Vuelca un `Recordset` sobre un objeto por reflexión | — | Sin escritura | `RC:145-171` |
| `FAC` | `CreateEntity` | Devuelve una instancia nueva según el nombre recibido | `New` sobre cada clase | Sin efectos | `FAC:11-65`; `On Error Resume Next` en `FAC:14`, `Select Case` en `FAC:16` |
| `HH` | `CalcularSHA256` | Devuelve la huella SHA-256 en hexadecimal | `CryptAcquireContext`, `CryptCreateHash`, `CryptHashData`, `CryptGetHashParam` | Sin efectos | `HH:39-82`; cadena vacía en `HH:49-52`, contexto en `HH:58`, hash en `HH:59` |
| `CH` | `GetItemsChecklistPorEstado` | Elige el texto de la lista de comprobación según estado y tipo de solicitud | cuatro funciones privadas | Sin efectos | `CH:36-70`; primera rama en `CH:45-46` |
| `MA` | `EjecutarActualizacion_*` | Ejecuta `ALTER TABLE … ADD COLUMN` contra la base de desarrollo | `getdb()`, `db.Execute` | Cambia el esquema | `MA:7-58`; sentencia en `MA:16` |
| `VG` | — | Expone el estado global: usuarios y conexiones | — | — | `VG:26-35` |

## Flujos de trabajo

- **Normal:** un repositorio llama a `RC.EjecutarAccion` con su SQL y su diccionario de parámetros; `RC` usa la conexión que le pasan o, si no, `getdb()` (`RC:13`). Si la conexión no está inicializada, lanza un error 513 con mensaje propio (`RC:14-18`).
- **Alternativos:** la hidratación admite forma de colección o de entidad única (`RC:69`, `RC:113`) y `RellenarRecordsetDesdeObjeto` admite omitir la clave primaria (`RC:173`), que es como los repositorios evitan reescribir el autonumérico en una actualización.
- **Errores:** `RC` captura el fallo con un `CondorError` que incluye la SQL y la clave de la conexión (`RC:33`, `RC:64`). `CreateEntity` hace lo contrario: `On Error Resume Next` y devuelve `Nothing` si el nombre no existe (`FAC:14`), de modo que un nombre mal escrito no se detecta aquí sino más tarde.

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Los roles tienen valores fijos 1, 2 y 3 | `Enum rol` | `ENUM:11-15` |
| Los estados del flujo empiezan en 1 y siguen un orden fijo | `Enum enumEstados` | `ENUM:36-…` |
| Sin conexión no se ejecuta nada: error 513 | `EjecutarAccion` y `EjecutarConsulta` | `RC:14-18`, `RC:47` |
| Un texto vacío no produce huella | `CalcularSHA256` | `HH:49-52` |
| El tipo de solicitud decide el texto de la lista de comprobación | `GetItemsChecklistPorEstado` | `CH:36-50` |
| La aplicación se identifica por una constante de nombre | `PROJECT_NAME` | `VG:12` |

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Cualquier acción de escritura de la aplicación | la que indique la SQL del llamador | SQL con `QueryDef` y parámetros ligados | La de la conexión recibida | `RC:7-37` |
| Cualquier lectura de la aplicación | la que indique la SQL del llamador | SQL con parámetros ligados | Ninguna | `RC:39-67` |
| Scripts de actualización de esquema | `tbSolicitudes` y otra | `ALTER TABLE … ADD COLUMN` | Ninguna | `MA:16`, `MA:38-58` |

## Coexistencia

`RC` es el cuello de botella de la migración: sus dos primitivas (`QueryDef` en memoria y parámetros ligados por diccionario) son un patrón de ACE que en el esquema destino se traduce a consultas parametrizadas reales. `HH` depende de la API criptográfica de Windows (`HH:58-68`): el hash que hoy identifica la versión de los datos de una revisión debe pasar a una función equivalente en el lenguaje de destino, conservando el algoritmo y la codificación, porque los hashes ya almacenados se comparan con los nuevos. `ENUM` fija valores de dominio en código (`ENUM:11-48`): el destino necesita tablas de catálogo o una decisión explícita de conservar los enteros. `VG` mantiene conexiones globales y el usuario activo y real en memoria (`VG:26-35`): ese estado no existe en web y su sustituto es la sesión. `MA` ejecuta DDL contra la base (`MA:16`): las migraciones de esquema dejan de ser código de aplicación.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual.** `RC` centraliza el acceso a datos por `QueryDef` con parámetros; los repositorios solo aportan SQL y valores. `FAC` instancia clases por nombre. `CH` devuelve textos de interfaz codificados en el módulo. `ENUM` y `VG` fijan el vocabulario y el estado global.

**Posibles defectos, a verificar.**
1. `VG:23` declara una constante con un prefijo de red de oficina: es configuración de entorno dentro del código.
2. `CreateEntity` silencia el error si la clase no existe (`FAC:14`) y devuelve `Nothing`.
3. `RC.EjecutarAccion` crea un `QueryDef` por llamada (`RC:20`): el patrón no existe en el destino y hay que elegir su equivalente.
4. `modActualizaciones` ejecuta cambios de esquema desde la aplicación (`MA:16`), y su propio comentario dice que fallar al añadir una columna ya existente es el resultado esperado (`MA:14-15`).
5. `Módulo1` son cuatro líneas sin procedimiento (`Módulo1:1-4`) y `ResultadoImportacionRAC` declara campos públicos en un `.bas`, no en una clase (`RIR:6-10`).

**Propuestas web pendientes de decisión.** Que el acceso a datos se centralice en una capa equivalente con consultas parametrizadas del destino. Que el hash de la revisión se calcule con una implementación del lenguaje de destino y se documente el algoritmo. Que los catálogos de estados, roles y decisiones sean datos y no código. Que la lista de comprobación de cada estado sea contenido traducible, no literales en un módulo.

## Criterios de aceptación

1. Una escritura con parámetros falla si no hay conexión inicializada y lo dice con el error propio. Trazado a `RC:14-18`.
2. Hidratar una entidad inexistente devuelve un valor vacío y no un error. Trazado a `RC:113-143`.
3. El hash de un texto es estable y el de un texto vacío es la cadena vacía. Trazado a `HH:49-52`.
4. La lista de comprobación cambia según el estado interno y el tipo de solicitud. Trazado a `CH:36-50`.
5. Los roles se conservan con sus valores actuales al migrar. Trazado a `ENUM:11-15`.

## Pendientes

- Confirmar contra el binario qué valor toma `SubRedOficina` en producción y si algún flujo lo usa para filtrar por ubicación (`VG:23`).
- Confirmar si `CreateEntity` recibe algún nombre desde fuera de los `Select Case` (por ejemplo desde configuración o desde una tabla): si fuera así, la desactivación de errores sería un problema de seguridad (`FAC:14`).
- Confirmar si `modActualizaciones` se ejecuta alguna vez en producción o solo en desarrollo (`MA:7-58`).
- Confirmar qué clase usa realmente el resultado de importación a RAC, ya que el contenedor está declarado en un módulo y no en `src/classes` (`RIR:6-10`).