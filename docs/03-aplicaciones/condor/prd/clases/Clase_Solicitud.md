# Clases de solicitud, filtro de búsqueda y sus modelos de vista

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/classes/Solicitud.cls` (64 líneas), `SolicitudServicio.cls` (741), `SolicitudBusquedaViewModel.cls` (65), `SolicitudViewModel.cls` (54), `FiltrosSolicitud.cls` (34) · **PRD:** [README](../README.md)

El agregado central de Condor y sus tres acompañantes. `Solicitud` es la entidad con campos públicos; `SolicitudServicio` es el único lugar donde se crea, valida y transiciona; los dos modelos de vista transportan a la interfaz lo que la entidad no debe exponer (correos, nombres de estado, tiempo de ciclo), y `FiltrosSolicitud` describe el criterio de búsqueda activa.

Abreviaturas: `SOL`, `classes/Solicitud.cls`; `SS`, `classes/SolicitudServicio.cls`; `SVM`, `classes/SolicitudViewModel.cls`; `SBVM`, `classes/SolicitudBusquedaViewModel.cls`; `FS`, `classes/FiltrosSolicitud.cls`; `SR`, `modules/SolicitudRepositorio.bas`; `SH`, `modules/SnapshotHelper.bas`; `FU`, `modules/FUNCIONES UTILES.bas`.

## Superficie pública

| Objeto | Tipo | Procedimientos y campos públicos | Qué resuelve |
|---|---|---|---|
| `Solicitud` | clase de entidad | campos públicos `idSolicitud`, `idExpediente`, `tipoSolicitud`, `codigoSolicitud`, `idNCAsociada`, `idEstadoInterno`, `fechaCreacion`, `usuarioCreacion` y el resto del bloque (`SOL:13-55`); `getPropiedad` y `SetPropiedad` (`SOL:56-64`) | Representación en memoria de una fila de `tbSolicitudes` |
| `SolicitudServicio` | clase de servicio | dieciséis procedimientos públicos (`SS:26-741`) | Alta con código propuesto, validación, búsqueda, historial de estados y borrado completo |
| `SolicitudViewModel` | clase de presentación | doce campos públicos (`SVM:16-39`), `getPropiedad` y `SetPropiedad` (`SVM:41-54`) | Fila de listado con datos de expediente, estado y responsables |
| `SolicitudBusquedaViewModel` | clase de presentación | campos públicos (`SBVM:21-57`), `getPropiedad` y `SetPropiedad` (`SBVM:59-65`) | Fila del resultado de búsqueda con datos del expediente y del RAC |
| `FiltrosSolicitud` | clase de filtro | `HayFiltrosActivos` (`FS:24-34`) y sus campos de criterio | Estado de los filtros de la pantalla de búsqueda |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| `SS` | `GuardarNuevaSolicitud` | Crea la entidad, la valida y la guarda | `CrearNuevaSolicitud`, `Validar`, `SR.GuardarSolicitud` | `INSERT` en `tbSolicitudes` | `SS:42-66`; secuencia en `SS:49-51`, conexión en `SS:47` |
| `SS` | `CrearNuevaSolicitud` | Construye la entidad con los valores iniciales | — | Sin escritura | `SS:316-347` |
| `SS` | `Validar` | Aplica las reglas de validación de la solicitud | — | Sin escritura | `SS:348-396` |
| `SS` | `ProponerCodigoSolicitud` | Propone el código siguiente del expediente | `ExpedienteServicio`, `SR.getSiguienteOrdinalParaClave` | Sin escritura | `SS:397-425`; expediente sin identificador en `SS:404-407` |
| `SS` | `getSolicitudesViewModel` | Devuelve el listado de solicitudes con datos de expediente y estado | repositorios de solicitud, expediente y estado | Lectura | `SS:67-172` |
| `SS` | `getSolicitudViewModelPorID` | Devuelve una fila concreta ya compuesta | varios repositorios | Lectura | `SS:173-302` |
| `SS` | `getSolicitudPorID`, `Guardar`, `ExisteCodigo` | Acceso directo a la entidad y su código | `SolicitudRepositorio` | Lectura y escritura | `SS:303-315`, `SS:661-671`, `SS:672-686` |
| `SS` | `getHistorialDeEstados` | Devuelve el historial de transiciones de la solicitud | `LogEstadoRepositorio` | Lectura | `SS:426-454` |
| `SS` | `EliminarSolicitudCompleta` | Borra la solicitud y todo lo que cuelga de ella | varios repositorios | Borrado en cascada manual | `SS:455-633` |
| `SS` | `GuardarConVinculacionNC` | Guarda la solicitud enlazando la no conformidad | repositorio y servicio de no conformidad | Escritura en dos bases | `SS:707-741` |
| `SS` | `CalcularHashFaseModificacion` | Calcula la huella de la fase de modificación | `SH.CalcularHashFaseValidacion` | Sin escritura | `SS:687-706` |
| `SS` | `getNombreAmigableTipoSolicitud` | Devuelve la etiqueta de un tipo de solicitud | — | Sin efectos | `SS:634-660` |
| `SS` | `SetForceNCCleanupFailure`, `GetNCCleanupFailureRaised` | Interruptor de prueba para forzar el fallo de limpieza | — | Sin efectos | `SS:26-41` |
| `SOL`, `SVM`, `SBVM` | `getPropiedad`, `SetPropiedad` | Acceso reflectivo por nombre de propiedad, con `Select Case` | — | Sin efectos | `SOL:56-64`, `SVM:41-54`, `SBVM:59-65` |
| `FS` | `HayFiltrosActivos` | Indica si hay algún criterio informado | — | Sin efectos | `FS:24-34` |

## Flujos de trabajo

- **Normal — alta:** el formulario pide un código propuesto (`SS:397`), el usuario lo acepta o lo corrige, se crea la entidad (`SS:316`), se valida (`SS:348`) y se guarda (`SS:661`). `GuardarNuevaSolicitud` encapsula los tres últimos pasos y fija la conexión (`SS:42-66`).
- **Normal — listado:** `getSolicitudesViewModel` compone cada fila con el expediente, el estado y los responsables (`SS:67-172`); por eso los modelos de vista llevan correos y nombres en lugar de identificadores.
- **Alternativos:** `getSolicitudViewModelPorID` devuelve una fila única (`SS:173`); si el expediente no tiene identificador, la propuesta de código sale vacía en lugar de fallar (`SS:404-407`).
- **Errores:** cada procedimiento público captura el fallo y lanza `CondorError` con el contexto; `SetForceNCCleanupFailure` permite simular el fallo de limpieza para probar el camino de error (`SS:26-41`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| El código de la solicitud se compone por expediente más ordinal | `ProponerCodigoSolicitud` | `SS:397-425` |
| Crear, validar y guardar es una secuencia obligatoria | `GuardarNuevaSolicitud` | `SS:49-51` |
| La entidad se valida antes de escribir | `Validar` | `SS:348-396` |
| La fila de listado no expone identificadores de estado, sino nombres | campos de `SVM` | `SVM:19-20` |
| Los modelos de vista no tienen comportamiento de negocio | solo acceso reflectivo | `SVM:41-54`, `SBVM:59-65` |

## Lecturas y escrituras

| Operación | Tabla o recurso | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Alta de solicitud | `tbSolicitudes` | vía `SolicitudRepositorio` | La que pase el servicio | `SS:51` |
| Lectura de listados | `tbSolicitudes`, `tbExpedientes`, `tbEstados`, `tbUsuariosAplicaciones` | repositorios | Ninguna | `SS:67-302` |
| Historial de estados | `tbLogEstados` | `LogEstadoRepositorio` | Ninguna | `SS:426-454` |
| Borrado completo | solicitud, bloques de datos, adjuntos, rechazos, logs | varios repositorios | La que pase el servicio | `SS:455-633` |
| Vinculación de no conformidad | `tbSolicitudes` y `TbNoConformidades` | dos conexiones | La que pase el servicio | `SS:707-741` |

## Coexistencia

La entidad y el servicio trasladan casi tal cual; el punto delicado es `getSolicitudesViewModel` y `getSolicitudViewModelPorID` (`SS:67-302`), que hoy componen la fila en memoria cruzando varias consultas. En el destino, la misma información debería venir de una consulta con las tablas unidas, y eso cambia el número de viajes a la base y el momento en que un cambio se ve reflejado. El segundo punto es `getPropiedad` y `SetPropiedad` (`SOL:56-64`, `SVM:41-54`, `SBVM:59-65`): son el mecanismo de hidratación por reflexión que usa `RepositorioComun`, y en el destino desaparecen porque la deserialización no necesita un contrato de propiedades por nombre. El tercero es el borrado completo (`SS:455-633`), que hoy es una secuencia explícita de borrados: en el destino debe decidir si esa secuencia sigue siendo explícita o pasa a una restricción de integridad.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual.** Una entidad de campos públicos, un servicio con toda la orquestación y tres clases de presentación sin comportamiento. La conexión la pasa el servicio; el repositorio la usa.

**Posibles defectos, a verificar.**
1. El borrado completo es una secuencia manual (`SS:455-633`): si un paso falla a mitad, quedan filas huérfanas salvo que la transacción cubra toda la secuencia.
2. `SetForceNCCleanupFailure` es un interruptor de fallo en el servicio de producción (`SS:26-41`).
3. `getSolicitudesViewModel` compone la fila consultando varias tablas por separado (`SS:67-172`): el coste crece con el número de filas.
4. La hidratación depende de que el nombre de la propiedad coincida exactamente con el nombre de la columna (`SOL:56-64`).

**Propuestas web pendientes de decisión.** Que el listado se resuelva con una consulta de lectura única y proyección explícita. Que el borrado completo sea una operación transaccional declarada, no una secuencia de llamadas. Que el contrato de la API exponga identificadores y nombres por separado, para no replicar el modelo de vista tal cual.

## Criterios de aceptación

1. El alta de una solicitud valida antes de escribir y no deja filas parciales. Trazado a `SS:42-66`.
2. El código propuesto para un expediente es el siguiente ordinal disponible. Trazado a `SS:397-425`.
3. El listado muestra nombre de estado y responsables sin exponer sus identificadores. Trazado a `SVM:19-28`.
4. El borrado completo elimina solicitud, bloques, adjuntos y trazas en una sola unidad. Trazado a `SS:455-633`.
5. Un expediente sin identificador produce un código vacío, no un error. Trazado a `SS:404-407`.

## Pendientes

- Confirmar contra el binario el orden exacto de los borrados de `EliminarSolicitudCompleta` y si todos comparten transacción (`SS:455-633`).
- Confirmar quién activa el interruptor de fallo y si sigue siendo necesario en producción (`SS:26-41`).
- Confirmar si `FiltrosSolicitud` se persiste entre sesiones o solo vive en memoria (`FS:24-34`).
- Confirmar el número de consultas que emite `getSolicitudesViewModel` para una página de resultados, porque condiciona la decisión sobre la consulta única (`SS:67-172`).