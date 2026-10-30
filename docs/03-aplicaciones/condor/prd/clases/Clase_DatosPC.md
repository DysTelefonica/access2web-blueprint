# Clases del bloque de datos de propuesta de cambio

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/classes/DatosPCServicio.cls` (887 líneas), `DatosPC.cls` (129), `DatosPCViewModel.cls` (104) · **PRD:** [README](../README.md)

La Propuesta de Cambio es uno de los tres bloques de datos de una solicitud, y el único cuya entidad tiene una hundredada de campos con nombres de negocio explícitos. `DatosPCServicio` reparte el trabajo: carga valores iniciales desde el expediente, guarda cada bloque por separado, comprueba si cada bloque está completo y registra el rechazo que nace en la formalización.

Abreviaturas: `DPS`, `classes/DatosPCServicio.cls`; `DP`, `classes/DatosPC.cls`; `DVM`, `classes/DatosPCViewModel.cls`; `DPREPO`, `modules/DatosPCRepositorio.bas`; `WF`, `classes/WorkflowServicio.cls`; `FU`, `modules/FUNCIONES UTILES.bas`; `SOL`, `classes/SolicitudServicio.cls`.

## Superficie pública

| Objeto | Tipo | Procedimientos y campos públicos | Qué resuelve |
|---|---|---|---|
| `DatosPC` | clase de entidad | campos del bloque: cabecera (`DP:14-20`), propuesta (`DP:22-25`), motivos (`DP:27-34`), impactos (`DP:36-48`), firmas y RAC (`DP:52-…`), `getPropiedad` y `SetPropiedad` (`DP:121-129`) | Fila de `tbDatosPC` |
| `DatosPCViewModel` | clase de presentación | entidad, solicitud y expediente, más los indicadores de permiso por bloque (`DVM:19-…`), `getPropiedad` y `SetPropiedad` (`DVM:93-104`) | Estado de edición de la propuesta: qué se ha cargado, en qué estado está y qué bloques se pueden guardar |
| `DatosPCServicio` | clase de servicio | veintidós procedimientos públicos (`DPS:15-887`) | Valores iniciales, guardado por bloques, completitud, purga y rechazo en formalización |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| `DPS` | `ActualizarCamposDependientesDeExpediente` | Copia al bloque los datos que vienen del expediente | — | Modifica la entidad en memoria | `DPS:15-53` |
| `DPS` | `GetDatosPCIniciales` | Prepara el bloque con los valores por defecto | `SOL.ProponerCodigoSolicitud` y expediente | Sin escritura | `DPS:138-233` |
| `DPS` | `GetDatosAprobacionSuministradorIniciales` | Prepara el bloque de aprobación de forma aislada | — | Sin escritura | `DPS:101-137` |
| `DPS` | `GuardarDatosGenerales` | Guarda la cabecera del bloque | `DPREPO.ActualizarDatosGenerales` | `INSERT`/`UPDATE` en `tbDatosPC` | `DPS:234-292` |
| `DPS` | `GuardarPropuesta`, `GuardarImpacto`, `GuardarAprobacionSuministrador`, `GuardarDictamenRAC`, `GuardarDecisionFinal` | Guarda un bloque concreto | `DPREPO` | Escritura por bloque | `DPS:293-365`, `DPS:366-492`, `DPS:493-517`, `DPS:518-561`, `DPS:562-575` |
| `DPS` | `PurgaTecnica` | Vacía el bloque de la solicitud | `DPREPO` | Borrado o puesta a nulo | `DPS:54-100` |
| `DPS` | `EliminarDictamenRAC`, `EliminarAprobacionSuministrador`, `EliminarDecisionFinal` | Limpian un bloque en lugar de borrarlo | `DPREPO.Limpiar*` | `UPDATE` a nulo | `DPS:576-603` |
| `DPS` | `getDatosPC` | Devuelve la entidad del bloque | `DPREPO.getPorIdSolicitud` | Lectura | `DPS:664-674` |
| `DPS` | `EsDatosGeneralesCompleta`, `EsParteTecnicaCompleta`, `EsDetalleCompleto`, `EsMotivosCompleto`, `EsDictamenRACCompleto`, `EsAprobacionSuministradorCompleta`, `EsDecisionFinalCompleta` | Comprueban la completitud de cada bloque | — | Sin escritura | `DPS:675-840` |
| `DPS` | `EsSolicitudEnValidacion` | Indica si la solicitud está en la fase de validación | — | Sin escritura | `DPS:841-870` |
| `DPS` | `RegistrarRechazoDesdeFormalizacion` | Registra el rechazo que nace al formalizar y revierte la fase | `WF`, `WF.BeginTrans` | Escritura en varias tablas | `DPS:604-663`; apertura de transacción en `DPS:613-615` |
| `DPS` | `AgregarSufijoRol` | Añade el rol entre paréntesis al nombre de un firmante | — | Sin efectos | `DPS:871-887`; no duplica el sufijo en `DPS:878-880` |
| `DP`, `DVM` | `getPropiedad`, `SetPropiedad` | Acceso reflectivo por nombre | — | Sin efectos | `DP:121-129`, `DVM:93-104` |

## Flujos de trabajo

- **Normal — carga:** el contenedor carga los valores iniciales del bloque (`DPS:138-233`), el usuario edita y guarda un bloque concreto (`DPS:293-575`); el modelo de vista lleva qué bloques puede guardar (`DVM:34-…`).
- **Normal — formalización:** al formalizar, si se produce un rechazo, `RegistrarRechazoDesdeFormalizacion` abre su propia transacción, lee el bloque, avisa al servicio de flujo y revierte (`DPS:604-663`).
- **Alternativos:** `GuardarImpacto` es, con diferencia, el bloque más extenso (`DPS:366-492`): recoge los motivos y los impactos, que son la parte más amplia de la entidad (`DP:27-48`).
- **Errores:** los siete procedimientos de completitud no lanzan errores: devuelven un booleano (`DPS:675-840`), de modo que quien decide es el flujo.

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Cada bloque se guarda por separado, no la entidad entera | siete procedimientos de guardado | `DPS:293-575` |
| La completitud se comprueba bloque a bloque | siete procedimientos `Es…Completa` | `DPS:675-840` |
| Limpiar un bloque es ponerlo a nulo, no borrarlo | `Eliminar*` frente a `PurgaTecnica` | `DPS:576-603` |
| El nombre del firmante lleva el rol entre paréntesis, sin duplicarlo | `AgregarSufijoRol` | `DPS:871-887` |
| La fase de validación condiciona qué se puede escribir | `EsSolicitudEnValidacion` | `DPS:841-870` |

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura y escritura del bloque | `tbDatosPC` | `Recordset` y SQL parametrizada | La que pase el servicio | `DPS:664-674`, `DPS:293-575` |
| Rechazo en formalización | `tbDatosPC`, rechazos y estado | varias, con transacción propia | Propia si el llamador no pasa la suya | `DPS:604-663` |
| Valores iniciales | expediente y solicitud | lectura | Ninguna | `DPS:101-233` |

## Coexistencia

La entidad `DatosPC` es la prueba más clara de por qué la migración necesita una matriz de campos: más de cien columnas con nombres de negocio explícitos (`DP:14-…`) que en el destino serán columnas o relaciones. La completitud por bloque (`DPS:675-840`) es lógica de negocio que debe sobrevivir tal cual, pero puede expresarse como reglas declarativas en el destino en lugar de código. Y `RegistrarRechazoDesdeFormalizacion` abre su propia transacción sobre el backend (`DPS:613-615`), exactamente el patrón que en el destino no puede existir: la unidad transaccional la decide el servidor.

## Comportamiento actual · Posibles defects · Propuestas web pendientes de decisión

**Comportamiento actual.** Una entidad ancha, un modelo de vista que transporta permisos de guardado y un servicio que carga, guarda por bloques, comprueba completitud y registra el rechazo de formalización.

**Posibles defectos, a verificar.**
1. `RegistrarRechazoDesdeFormalizacion` abre su propia transacción (`DPS:613-615`) y puede quedar fuera de la del servicio que lo invoca.
2. Los siete procedimientos de completitud repiten la misma forma y podrían divergir al cambiar una regla (`DPS:675-840`).
3. La entidad mezcla cabecera, propuesta, motivos, impactos, firmas y RAC en una sola fila (`DP:14-…`), lo que dificulta cualquier consulta parcial.

**Propuestas web pendientes de decisión.** Que la completitud sea una declaración por bloque y no una serie de funciones. Que la unidad transaccional del rechazo la imponga el servidor. Que la entidad se divida en el destino si el número de columnas de la tabla bloquea las consultas que el producto necesita.

## Criterios de aceptación

1. Los valores iniciales del bloque se cargan con los datos del expediente y de la solicitud. Trazado a `DPS:138-233`.
2. Guardar un bloque no modifica los demás. Trazado a `DPS:293-575`.
3. Un bloque incompleto se detecta sin lanzar error. Trazado a `DPS:675-840`.
4. Limpiar un bloque lo deja a nulo sin borrar la fila. Trazado a `DPS:576-603`.
5. El rechazo de formalización revierte la fase y deja el registro. Trazado a `DPS:604-663`.

## Pendientes

- Confirmar contra el binario la lista completa de columnas de `tbDatosPC`, que condiciona la matriz de campos (`DP:14-…`).
- Confirmar si los siete bloques de completitud se usan todos o si alguno quedó sin consumidor (`DPS:675-840`).
- Confirmar qué invocadores pasan conexión a `RegistrarRechazoDesdeFormalizacion` y cuáles no (`DPS:604-663`).
- Confirmar el formato de las columnas de motivos e impactos, que hoy son booleanos y cadenas (`DP:27-48`).