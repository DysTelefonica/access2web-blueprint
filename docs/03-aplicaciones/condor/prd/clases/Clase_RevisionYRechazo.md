# Clases de revisión, rechazo y no conformidad

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/classes/RevisionServicio.cls` (161 líneas), `NoConformidadServicio.cls` (105), `ValidacionRevisionServicio.cls` (99), `RechazoServicio.cls` (72), `ValidacionRevision.cls` (54), `NoConformidad.cls` (46), `Rechazo.cls` (20), `ResultadoImportacionRAC.cls` (18) · **PRD:** [README](../README.md)

Los cuatro objetos que rodean la revisión de calidad y el rechazo: la decisión de revisión con su historial, la validación de revisión con su versionado por ordinal, el rechazo con su motivo, y la no conformidad, que vive en otra base de datos y se vincula a la solicitud por su código de concesión.

Abreviaturas: `RS`, `classes/RevisionServicio.cls`; `VRS`, `classes/ValidacionRevisionServicio.cls`; `RCS`, `classes/RechazoServicio.cls`; `NCS`, `classes/NoConformidadServicio.cls`; `VR`, `classes/ValidacionRevision.cls`; `REC`, `classes/Rechazo.cls`; `NC`, `classes/NoConformidad.cls`; `RIR`, `classes/ResultadoImportacionRAC.cls`; `VRR`, `modules/ValidacionRevisionRepositorio.bas`; `RRE`, `modules/RechazoRepositorio.bas`; `NCR`, `modules/NoConformidadRepositorio.bas`; `SOL`, `classes/SolicitudServicio.cls`.

## Superficie pública

| Objeto | Tipo | Procedimientos y campos públicos | Qué resuelve |
|---|---|---|---|
| `RevisionServicio` | clase de servicio | `GuardarDecisionRevision`, `ActualizarMotivoRechazoActivo`, `ObtenerRechazoActivo` (`RS:22-161`) | Registrar la decisión de revisión de calidad y su motivo |
| `ValidacionRevisionServicio` | clase de servicio | seis procedimientos (`VRS:19-99`) | Ciclo de validación de revisión por ordinal, con su hash |
| `RechazoServicio` | clase de servicio | `LimpiarHistorialRechazos`, `GetUltimoRechazoActivo`, `DesactivarRechazosPrevios` (`RCS:19-72`) | Historial de rechazos de una solicitud |
| `NoConformidadServicio` | clase de servicio | ocho procedimientos (`NCS:14-105`) | No conformidades de la base externa y su vínculo con el código de concesión |
| `ValidacionRevision` | clase de entidad | `id`, `idSolicitud`, `ordinal`, `fechaEnvio`, `fechaRecepcion`, `resultado`, `comentarios`, `usuario`, `idAdjunto`, `hashDatos` (`VR:14-23`) | Fila de `tbValidacionRevision` |
| `Rechazo` | clase de entidad | `idSolicitud`, `fechaRechazo`, `motivoPrincipal`, `areaAfectada`, `comentarios`, `usuarioRechazo`, `notasSubsanacion`, `CambiosTecnico` (`REC:13-20`) | Fila de `tbRechazos` |
| `NoConformidad` | clase de entidad | `IDNoConformidad`, `CodigoNoConformidad`, `CodConcesionAsociada`, `descripcion`, `idExpediente` (`NC:17-21`) | Fila de `TbNoConformidades` |
| `ResultadoImportacionRAC` | clase de entidad | `Success` y los campos del resultado de importación (`RIR:14-18`) | Resultado de importar un RAC |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| `RS` | `GuardarDecisionRevision` | Guarda la decisión de calidad, deja traza y actualiza la solicitud | `SOL`, `LogServicio`, `RechazoServicio` | Escritura en varias tablas, con transacción propia si no la recibe | `RS:22-100`; variables de transacción en `RS:28-31` |
| `RS` | `ActualizarMotivoRechazoActivo` | Reescribe el motivo del rechazo activo | `RechazoServicio` | `UPDATE` en `tbRechazos` | `RS:101-149` |
| `RS` | `ObtenerRechazoActivo` | Devuelve el rechazo activo | `RCS.GetUltimoRechazoActivo` | Lectura | `RS:150-161` |
| `VRS` | `GetUltimoOrdinal` | Devuelve el ordinal del último ciclo | `VRR.GetUltimoOrdinal` | Lectura | `VRS:19-28` |
| `VRS` | `RegistrarEnvioRAC`, `RegistrarBorrador`, `LimpiarPorSolicitud`, `GetUltimoPendienteId` | Gestionan los hitos del ciclo de validación | `VRR` | Escritura en `tbValidacionRevision` | `VRS:29-38`, `VRS:70-79`, `VRS:80-88`, `VRS:89-99` |
| `VRS` | `RegistrarRespuestaRAC` | Cierra un ciclo con su resultado, resolviendo el hito pendiente si no se le pasa identificador | `VRR.GetUltimoPendienteId` | Escritura en `tbValidacionRevision` | `VRS:39-59`; resolución del hito en `VRS:43-47` |
| `VRS` | `RegistrarRespuestaRACPorId` | Cierra un ciclo por identificador de hito | `VRR` | Escritura | `VRS:60-69` |
| `RCS` | `GetUltimoRechazoActivo`, `DesactivarRechazosPrevios`, `LimpiarHistorialRechazos` | Leen, invalidan y borran el historial de rechazos | `RRE` | Lectura y escritura | `RCS:19-72` |
| `NCS` | `getNoConformidades`, `getNoConformidadesPorExpediente`, `getNoConformidadPorCodigoCondor`, `getNoConformidadPorID` | Buscan no conformidades | `NCR` | Lectura de la base externa | `NCS:14-62` |
| `NCS` | `estaRegistradaEnBaseDatosExternaa` | Comprueba si la no conformidad existe fuera | `NCR.existeEnBaseDatosExternaa` | Lectura | `NCS:63-70` |
| `NCS` | `actualizarCodigoConcesionEnTransaccion`, `limpiarCodigoConcesionEnTransaccion` | Enlaza y desenlaza la concesión dentro de la transacción del servicio | `NCR` | `UPDATE` en la base externa | `NCS:71-105` |
| `VR`, `REC`, `NC`, `RIR` | campos y acceso reflectivo | Representación en memoria de cada fila | — | Sin efectos | `VR:14-23`, `REC:13-20`, `NC:17-21`, `RIR:14-18` |

## Flujos de trabajo

- **Normal — revisión:** calidad decide sobre la solicitud; `GuardarDecisionRevision` deja traza del cambio, guarda la decisión y, si procede, registra el rechazo (`RS:22-100`). El motivo del rechazo activo puede reescribirse después (`RS:101-149`).
- **Normal — validación de revisión:** el ciclo avanza por ordinal; el envío y el borrador se registran como hitos (`VRS:29-38`, `VRS:70-79`) y la respuesta cierra el ciclo, ya sea sobre el hito que se le pasa o sobre el último pendiente (`VRS:43-47`).
- **Alternativos:** `RegistrarRespuestaRAC` admite las dos formas (`VRS:39-59`), lo que evita que el llamador tenga que consultar el ordinal antes de responder.
- **Errores:** las operaciones que escriben en la base externa reciben la conexión del servicio (`NCS:71-105`) para que el vínculo con la concesión quede dentro de la misma unidad de trabajo.

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Un ciclo de revisión se identifica por solicitud y ordinal | `ValidacionRevision` | `VR:14-23` |
| El ciclo conserva el hash de los datos que se enviaron | campo `hashDatos` | `VR:23` |
| Solo hay un rechazo activo por solicitud | `DesactivarRechazosPrevios` | `RCS:55-72` |
| La no conformidad se vincula al código de concesión, no a la clave de la solicitud | `CodConcesionAsociada` | `NC:19` |
| El resultado de la importación del RAC es un objeto con resultado y mensaje | `ResultadoImportacionRAC` | `RIR:14-18` |

## Lecturas y escrituras

| Operación | Tabla | Base | Tipo | Transacción | Evidencia |
|---|---|---|---|---|---|
| Decisión de revisión | solicitud, logs y rechazos | Condor | varias | Propia si el llamador no pasa la suya | `RS:22-100` |
| Ciclos de validación | `tbValidacionRevision` | Condor | SQL parametrizada | La que pase el servicio | `VRS:19-99` |
| Rechazos | `tbRechazos` | Condor | `Recordset` y SQL | La que pase el servicio | `RCS:19-72` |
| No conformidades | `TbNoConformidades` | No Conformidades | SQL con lista de columnas | La del servicio | `NCS:14-105` |

## Coexistencia

Dos decisiones de frontera. La primera es la no conformidad: `NoConformidadServicio` opera sobre la base de No Conformidades con conexión propia (`NCS:14-105`), y Condor guarda en ella el código de la concesión asociada (`NC:19`); en el destino, ese vínculo necesita una clave estable que sobreviva al cambio de código de la solicitud, porque hoy depende de una cadena de texto (`NCR` de la hoja de trazabilidad). La segunda es el ciclo de validación: la pareja solicitud y ordinal más el hash de datos (`VR:14-23`) es el mecanismo de versionado, y debe conservarse con el mismo algoritmo para que los ciclos abiertos en el momento de la migración sigan siendo verificables.

## Comportamiento actual · Posibles defects · Propuestas web pendientes de decisión

**Comportamiento actual.** Cuatro entidades sin comportamiento y cuatro servicios: decisión de revisión con su motivo, ciclo de validación por ordinal con hash, historial de rechazos y catálogo de no conformidades de otra base.

**Posibles defectos, a verificar.**
1. `GuardarDecisionRevision` abre su propia transacción si nadie le pasa conexión (`RS:28-31`).
2. La vinculación con la no conformidad depende del código de concesión, que cambia si cambia el expediente (`NC:19`).
3. `RegistrarRespuestaRAC` admite dos formas de identificar el hito (`VRS:43-47`): la forma implícita depende de cuál sea el último pendiente.
4. `ResultadoImportacionRAC` está declarado como clase y también como módulo (`RIR:14-18`, `RIR` de la hoja de utilidades): hay dos definiciones del mismo concepto.

**Propuestas web pendientes de decisión.** Que el vínculo con la no conformidad use una clave propia en lugar del código de concesión. Que el ciclo de validación conserve ordinal y hash con el mismo algoritmo. Que la unidad transaccional de la decisión de revisión la imponga el servidor.

## Criterios de aceptación

1. La decisión de revisión deja traza del cambio y actualiza la solicitud en una sola unidad. Trazado a `RS:22-100`.
2. El motivo del rechazo activo puede corregirse sin crear un rechazo nuevo. Trazado a `RS:101-149`.
3. Cada ciclo de validación conserva ordinal, resultado y hash. Trazado a `VRS:19-99`, `VR:14-23`.
4. Solo un rechazo queda activo por solicitud. Trazado a `RCS:55-72`.
5. Vincular una concesión a una no conformidad queda dentro de la transacción del servicio. Trazado a `NCS:71-105`.

## Pendientes

- Confirmar contra el binario el catálogo de resultados de la validación de revisión, porque la interfaz los muestra (`VR:19`).
- Confirmar si la no conformidad asociada se desvincula al eliminar una solicitud o queda huérfana (`NCS:71-105`).
- Confirmar cuál de las dos definiciones de `ResultadoImportacionRAC` se usa en ejecución (`RIR:14-18`).
- Confirmar la tabla completa de `tbValidacionRevision` para la matriz de campos (`VR:14-23`).