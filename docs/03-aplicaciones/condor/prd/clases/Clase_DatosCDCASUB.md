# Clases del bloque de datos de concesión o desviación de subsuministrador

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/classes/DatosCDCASUBServicio.cls` (730 líneas), `DatosCDCASUB.cls` (149), `DatosCDCASUBViewModel.cls` (100) · **PRD:** [README](../README.md)

El tercer bloque de datos: concesión o desviación de subsuministrador (`CD_CA_SUB`). Mantiene la misma forma que los otros dos —entidad, modelo de vista y servicio por bloques— pero añade el rechazo en formalización, que en la propuesta de cambio abre su propia transacción, y expone `EsSolicitudEnValidacion` como privado en lugar de privado también en el bloque de propuesta de cambio, señal de que la copia no fue exacta.

Abreviaturas: `SSRV`, `classes/DatosCDCASUBServicio.cls`; `SUB`, `classes/DatosCDCASUB.cls`; `SVM`, `classes/DatosCDCASUBViewModel.cls`; `REPO`, `modules/DatosCDCASUBRepositorio.bas`; `WF`, `classes/WorkflowServicio.cls`.

## Superficie pública

| Objeto | Tipo | Procedimientos y campos públicos | Qué resuelve |
|---|---|---|---|
| `DatosCDCASUB` | clase de entidad | cuarenta y tres campos públicos (`SUB:14-…`), `getPropiedad` y `SetPropiedad` (`SUB:141-149`) | Fila de `tbDatosCDCASUB` |
| `DatosCDCASUBViewModel` | clase de presentación | entidad, solicitud, expediente, estado y permisos por bloque (`SVM:19-…`), `getPropiedad` y `SetPropiedad` (`SVM:89-100`) | Estado de edición del bloque |
| `DatosCDCASUBServicio` | clase de servicio | veintitrés procedimientos públicos (`SSRV:14-730`) | Valores iniciales, guardado por bloques, completitud, limpieza y rechazo en formalización |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| `SSRV` | `getDatosCDCASUB` | Devuelve la entidad del bloque | `REPO.getPorIdSolicitud` | Lectura | `SSRV:14-23` |
| `SSRV` | `GetDatosGeneralesIniciales` | Prepara el bloque con los valores por defecto | repositorios de solicitud, expediente y bloque | Sin escritura | `SSRV:24-113` |
| `SSRV` | `GetDatosAprobacionIniciales` | Prepara el bloque de aprobación | — | Sin escritura | `SSRV:114-132` |
| `SSRV` | `GuardarDatosGenerales`, `GuardarPropuesta`, `GuardarImpacto`, `GuardarAprobacionSuministrador`, `GuardarDictamenRAC`, `GuardarDecisionFinal` | Guardan un bloque concreto | `REPO` | Escritura por bloque | `SSRV:133-172`, `SSRV:173-237`, `SSRV:238-316`, `SSRV:317-326`, `SSRV:327-364`, `SSRV:365-376` |
| `SSRV` | `EsDatosGeneralesCompleta`, `EsParteTecnicaCompleta`, `EsDetalleCompleto`, `EsMotivosCompleto`, `EsDictamenRACCompleto`, `EsAprobacionSuministradorCompleta`, `EsDecisionFinalCompleta` | Comprueban la completitud por bloque | — | Sin escritura | `SSRV:377-541` |
| `SSRV` | `EliminarDictamenRAC`, `EliminarAprobacionSuministrador`, `EliminarDecisionFinal` | Limpian un bloque poniéndolo a nulo | `REPO` | `UPDATE` a nulo | `SSRV:542-559` |
| `SSRV` | `ActualizarCamposDependientesDeExpediente` | Copia al bloque los datos del expediente | — | Modifica la entidad en memoria | `SSRV:560-585` |
| `SSRV` | `PurgaTecnica` | Vacía el bloque | `REPO` | Borrado o puesta a nulo | `SSRV:586-626` |
| `SSRV` | `RegistrarRechazoDesdeFormalizacion` | Registra el rechazo que nace al formalizar y revierte la fase | `WF` | Escritura en varias tablas | `SSRV:627-683` |
| `SSRV` | `EsSolicitudEnValidacion`, `AgregarSufijoRol` | Reglas internas del servicio | — | Sin efectos | `SSRV:684-730` |
| `SUB`, `SVM` | `getPropiedad`, `SetPropiedad` | Acceso reflectivo por nombre | — | Sin efectos | `SUB:141-149`, `SVM:89-100` |

## Flujos de trabajo

- **Normal — carga y guardado:** el contenedor pide los valores iniciales (`SSRV:24-132`), el usuario edita y el servicio guarda el bloque (`SSRV:133-376`); la completitud se comprueba por bloques (`SSRV:377-541`).
- **Alternativos:** el rechazo en formalización (`SSRV:627-683`) sigue el mismo recorrido que en la propuesta de cambio: obtener el bloque, avisar al servicio de flujo y revertir.
- **Errores:** los guardados lanzan `CondorError`; los siete procedimientos de completitud devuelven booleano (`SSRV:377-541`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Cada bloque se guarda por separado | seis procedimientos | `SSRV:133-376` |
| La completitud se comprueba bloque a bloque | siete procedimientos | `SSRV:377-541` |
| El limpiado pone a nulo y no borra la fila | `Eliminar*` | `SSRV:542-559` |
| La fase de validación condiciona la escritura | `EsSolicitudEnValidacion` | `SSRV:684-713` |
| El nombre del firmante lleva el rol entre paréntesis | `AgregarSufijoRol` | `SSRV:714-730` |

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura del bloque | `tbDatosCDCASUB` | SQL parametrizada | Ninguna | `SSRV:14-23` |
| Escritura por bloque | `tbDatosCDCASUB` | `Recordset` y SQL parametrizada | La que pase el servicio | `SSRV:133-376` |
| Rechazo en formalización | `tbDatosCDCASUB`, rechazos y estado | varias | La del propio flujo | `SSRV:627-683` |

## Coexistencia

Los tres bloques de datos comparten problema de destino: tabla ancha, guardado por bloques y completitud en código. La diferencia entre ellos se reduce a columnas, y esa diferencia es la que exige una matriz de campos con trazabilidad columna a columna. Lo que este bloque añade es el rechazo en formalización (`SSRV:627-683`), cuyo comportamiento debe coincidir con el del bloque de propuesta de cambio para que la interfaz se comporte igual ante un rechazo, independientemente del tipo de solicitud.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual.** Entidad, modelo de vista y servicio con la misma división en bloques que los otros dos tipos, más el rechazo en formalización.

**Posibles defectos, a verificar.**
1. `EsSolicitudEnValidacion` es privado aquí (`SSRV:684-713`) y también en el bloque de propuesta de cambio, pero el resto de los servicios no lo expone: la regla de fase puede estar replicada con nombres distintos.
2. La superficie de los tres servicios no es idéntica: aquí no hay procedimiento de campos dependientes del expediente en la misma posición ni nombres de parámetro homogéneos (`SSRV:560-585`).
3. Los siete procedimientos de completitud repiten forma (`SSRV:377-541`).

**Propuestas web pendientes de decisión.** Que la completitud de los tres tipos sea una declaración compartida y no tres juegos de funciones. Que la regla de fase valida se evalúe en un único punto del servidor. Que la matriz de campos documente las columnas comunes a las tres tablas y las específicas de cada una.

## Criterios de aceptación

1. Los valores iniciales del bloque provienen del expediente y de la solicitud. Trazado a `SSRV:24-132`.
2. Guardar un bloque no modifica los demás. Trazado a `SSRV:133-376`.
3. Un bloque incompleto se detecta sin lanzar error. Trazado a `SSRV:377-541`.
4. El rechazo en formalización revierte la fase y deja el registro. Trazado a `SSRV:627-683`.

## Pendientes

- Confirmar contra el binario la lista completa de columnas de `tbDatosCDCASUB` y su correspondencia con `tbDatosPC` y `tbDatosCDCA` (`SUB:14-…`).
- Confirmar si el rechazo en formalización de este bloque usa la misma ruta transaccional que el de la propuesta de cambio (`SSRV:627-683`).
- Confirmar si los siete bloques de completitud se usan todos (`SSRV:377-541`).