# Clases del bloque de datos de concesión o desviación

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/classes/DatosCDCAServicio.cls` (909 líneas), `DatosCDCA.cls` (130), `DatosCDCAViewModel.cls` (113) · **PRD:** [README](../README.md)

El segundo bloque de datos, para concesiones y desviaciones (`CD_CA`). Replica la arquitectura del bloque de propuesta de cambio: entidad, modelo de vista y servicio con guardado por bloques y comprobación de completitud. La diferencia relevante está en la firma de aprobación, que distingue unidades de producción, diseño y representación, y en el nombre del procedimiento que la guarda.

Abreviaturas: `DCS`, `classes/DatosCDCAServicio.cls`; `DC`, `classes/DatosCDCA.cls`; `DVM`, `classes/DatosCDCAViewModel.cls`; `REPO`, `modules/DatosCDCARepositorio.bas`; `WF`, `classes/WorkflowServicio.cls`.

## Superficie pública

| Objeto | Tipo | Procedimientos y campos públicos | Qué resuelve |
|---|---|---|---|
| `DatosCDCA` | clase de entidad | treinta y nueve campos públicos (`DC:14-…`), `getPropiedad` y `SetPropiedad` (`DC:121-130`) | Fila de `tbDatosCDCA` |
| `DatosCDCAViewModel` | clase de presentación | entidad, solicitud, expediente, estado y permisos por bloque (`DVM:19-…`), `getPropiedad` y `SetPropiedad` (`DVM:100-113`) | Estado de edición del bloque |
| `DatosCDCAServicio` | clase de servicio | veintitrés procedimientos públicos (`DCS:16-909`) | Valores iniciales, guardado por bloques, completitud y limpieza |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| `DCS` | `ActualizarCamposDependientesDeExpediente` | Copia al bloque los datos que vienen del expediente | — | Modifica la entidad en memoria | `DCS:16-45` |
| `DCS` | `GetDatosCDCA` | Devuelve la entidad del bloque | `REPO.getPorIdSolicitud` | Lectura | `DCS:93-115` |
| `DCS` | `GetDatosGeneralesIniciales` | Prepara el bloque con los valores por defecto | repositorio y expediente | Sin escritura | `DCS:116-198` |
| `DCS` | `GetNombreSuministradorDefault` | Resuelve el suministrador por defecto del expediente | repositorio de expediente | Lectura | `DCS:199-262` |
| `DCS` | `GetDatosAprobacionIniciales` | Prepara el bloque de aprobación | — | Sin escritura | `DCS:263-287` |
| `DCS` | `GuardarDatosGenerales`, `GuardarPropuesta`, `GuardarImpacto`, `GuardarDictamenRAC`, `GuardarDecisionFinal` | Guardan un bloque concreto | `REPO` | Escritura por bloque | `DCS:288-350`, `DCS:351-419`, `DCS:420-507`, `DCS:547-590`, `DCS:591-610` |
| `DCS` | `GuardarAprobacionSuministrador` | Guarda el bloque de aprobación según las unidades que lo firman | `REPO` | Escritura | `DCS:508-546`; firmas por unidad en `DCS:508` |
| `DCS` | `PurgaTecnica` | Vacía el bloque | `REPO` | Borrado o puesta a nulo | `DCS:46-92` |
| `DCS` | `EliminarDictamenRAC`, `EliminarAprobacionSuministrador`, `EliminarDecisionFinal` | Limpian un bloque poniéndolo a nulo | `REPO.Limpiar*` | `UPDATE` a nulo | `DCS:774-816` |
| `DCS` | `EsDatosGeneralesCompleta`, `EsParteTecnicaCompleta`, `EsDetalleCompleto`, `EsMotivosCompleto`, `EsDictamenRACCompleto`, `EsAprobacionSuministradorCompleta`, `EsDecisionFinalCompleta` | Comprueban la completitud por bloque | — | Sin escritura | `DCS:611-773` |
| `DC`, `DVM` | `getPropiedad`, `SetPropiedad` | Acceso reflectivo por nombre | — | Sin efectos | `DC:121-130`, `DVM:100-113` |

## Flujos de trabajo

- **Normal — carga y guardado:** el contenedor pide los valores iniciales (`DCS:116-262`), el usuario edita y el servicio guarda el bloque correspondiente (`DCS:288-610`); el modelo de vista lleva los permisos de guardado por bloque (`DVM:19-…`).
- **Alternativos:** `GuardarAprobacionSuministrador` recibe tres indicadores, uno por unidad que firma (`DCS:508`), de modo que el bloque aprobado puede tener solo algunas firmas rellenadas.
- **Errores:** la completitud se informa por valor de retorno (`DCS:611-773`); los guardados lanzan `CondorError` y el manejador del formulario los presenta.

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Cada bloque se guarda por separado | seis procedimientos de guardado | `DCS:288-610` |
| La aprobación se guarda según las unidades que firman | tres indicadores de entrada | `DCS:508` |
| La completitud se comprueba bloque a bloque, sin excepciones | siete procedimientos | `DCS:611-773` |
| El limpiado pone a nulo y no borra la fila | `Eliminar*` | `DCS:774-816` |
| El suministrador por defecto viene del expediente | `GetNombreSuministradorDefault` | `DCS:199-262` |

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura del bloque | `tbDatosCDCA` | SQL parametrizada | Ninguna | `DCS:93-115` |
| Escritura por bloque | `tbDatosCDCA` | `Recordset` y SQL parametrizada | La que pase el servicio | `DCS:288-610` |
| Valores iniciales | expediente, solicitud y bloque | lectura | Ninguna | `DCS:116-262` |

## Coexistencia

El mismo esquema que para la propuesta de cambio: entidad ancha de treinta y nueve columnas (`DC:14-…`), guardado por bloques y completitud por código. La diferencia que importa para el destino es la firma de aprobación: hoy se guarda en el mismo bloque, con una firma por unidad (`DCS:508`), y las hojas de formularios muestran que esas firmas se pueden vaciar de forma independiente (`REPO` de la hoja de repositorios de datos). En el destino conviene decidir si las firmas siguen siendo columnas o pasan a una entidad de firmas con su propio ciclo de vida.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual.** Entidad, modelo de vista y servicio con la misma división en bloques que la propuesta de cambio. La firma de aprobación se guarda por unidad.

**Posibles defectos, a verificar.**
1. La firma de aprobación depende de tres booleanos de entrada (`DCS:508`): el estado de cada firma depende de quién llama.
2. `GetNombreSuministradorDefault` (`DCS:199-262`) es, junto con `GetDatosGeneralesIniciales`, la parte más larga del servicio para un cálculo de valores por defecto.
3. Los siete procedimientos de completitud repiten forma y pueden divergir (`DCS:611-773`).

**Propuestas web pendientes de decisión.** Que las firmas sean una entidad propia con estado y fecha, no columnas del bloque. Que la completitud se declare por bloque. Que los valores iniciales se calculen en el servidor y no en la carga de la vista.

## Criterios de aceptación

1. Los valores iniciales del bloque provienen del expediente y de la solicitud. Trazado a `DCS:116-262`.
2. Guardar un bloque no modifica los demás. Trazado a `DCS:288-610`.
3. El bloque de aprobación guarda solo las firmas de las unidades que firman. Trazado a `DCS:508-546`.
4. Un bloque incompleto se detecta sin lanzar error. Trazado a `DCS:611-773`.

## Pendientes

- Confirmar contra el binario la lista completa de columnas de `tbDatosCDCA` y su correspondencia con las de `tbDatosPC`, porque comparten nombre en varios campos (`DC:14-…`).
- Confirmar si los siete bloques de completitud se usan todos (`DCS:611-773`).
- Confirmar qué ocurre con las firmas de un bloque que se limpia a posteriori (`DCS:774-816`).