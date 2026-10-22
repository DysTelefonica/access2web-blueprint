# Clases de adjuntos y su servicio

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/classes/AdjuntosServicio.cls` (541 líneas), `Adjunto.cls` (52), `AdjuntoViewModel.cls` (39) · **PRD:** [README](../README.md)

Los adjuntos son la única parte de Condor que toca el sistema de ficheros. `AdjuntosServicio` no solo guarda filas: copia el fichero a un directorio de destino, decide si el adjunto es del tipo que exige firma, y al eliminar borra también el fichero. Eso convierte al servicio en el punto donde hay que decidir qué pasa con el almacenamiento de ficheros en la versión web.

Abreviaturas: `AS`, `classes/AdjuntosServicio.cls`; `ADJ`, `classes/Adjunto.cls`; `AVM`, `classes/AdjuntoViewModel.cls`; `AR`, `modules/AdjuntoRepositorio.bas`; `FU`, `modules/FUNCIONES UTILES.bas`; `CE`, `classes/CondorError.cls`.

## Superficie pública

| Objeto | Tipo | Procedimientos y campos públicos | Qué resuelve |
|---|---|---|---|
| `Adjunto` | clase de entidad | campos `idAdjunto`, `idSolicitud`, `etapaWF`, `nombreArchivo`, `fechaSubida`, `usuarioSubida`, `descripcion`, `tipoAccion` (`ADJ:17-24`); `ColCampos` (`ADJ:27`); `getPropiedad` y `SetPropiedad` (`ADJ:45-52`) | Fila de `tbAdjuntos` con la etapa del flujo a la que pertenece |
| `AdjuntoViewModel` | clase de presentación | los mismos campos más `etapaNombre` (`AVM:18-27`), `getPropiedad` y `SetPropiedad` (`AVM:29-39`) | Adjunto con el nombre legible de su etapa |
| `AdjuntosServicio` | clase de servicio | trece procedimientos públicos (`AS:17-541`) | Guardar un fichero como adjunto, actualizarlo, listarlo por solicitud o etapa y eliminarlo |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| `AS` | `GuardarAdjuntoDesdeArchivo` | Valida, copia el fichero al destino y crea la fila del adjunto | `Validar`, `AR.Guardar`, sistema de ficheros | `INSERT` en `tbAdjuntos` y escritura de fichero | `AS:17-131`; firma en `AS:17-23`, objetos de trabajo en `AS:24-26` |
| `AS` | `ActualizarFicheroAdjunto` | Sustituye el fichero de un adjunto existente | `AR` y sistema de ficheros | `UPDATE` y escritura de fichero | `AS:132-196` |
| `AS` | `Validar` | Aplica las reglas de la entidad y lanza el error 513 si alguna falla | — | Sin escritura | `AS:277-309`; objeto nulo en `AS:283`, identificador en `AS:286`, etapa vacía en `AS:289`, nombre en `AS:291` |
| `AS` | `ExisteAdjuntoEtapa`, `GetIdAdjuntoPorEtapa` | Resuelven el adjunto de una etapa por su nombre | `AR` | Lectura | `AS:197-221` |
| `AS` | `ExisteAdjuntoGeneradoCierre`, `GeneradoPDFParaFirma` | Indican si existe el adjunto de cierre y si es un PDF para firma | `AR` | Lectura | `AS:222-266` |
| `AS` | `getAdjuntoPorID`, `getAdjuntosPorSolicitud`, `getAdjuntosPorSolicitudTx` | Leen adjuntos por clave o por solicitud, con y sin conexión del servicio | `AR` | Lectura | `AS:267-276`, `AS:310-354` |
| `AS` | `getAdjuntosViewModelPorSolicitud` | Compone las filas con el nombre de la etapa | los anteriores | Lectura | `AS:355-412` |
| `AS` | `EliminarAdjunto` | Borra la fila y el fichero, con transacción propia si no la recibe | `AR`, `EtapaAdjuntoCoincide`, sistema de ficheros | `DELETE` y borrado de fichero | `AS:484-541`; identificador inválido en `AS:491-493`, transacción local en `AS:495-496` |
| `AS` | `EliminarAdjuntosPorEtapas` | Borra todos los adjuntos de las etapas indicadas | `AR` y sistema de ficheros | `DELETE` y borrado de ficheros | `AS:413-483`; coincidencia por etapa en `AS:455-483` |
| `ADJ`, `AVM` | `getPropiedad`, `SetPropiedad` | Acceso reflectivo por nombre de propiedad | — | Sin efectos | `ADJ:45-52`, `AVM:29-39` |

## Flujos de trabajo

- **Normal — alta:** el usuario elige un fichero en una etapa del flujo; el servicio valida la entidad (`AS:277-291`), copia el fichero al directorio de destino (`AS:17-26`), inserta la fila (`AS:17-131`) y devuelve el identificador nuevo mediante el parámetro de salida.
- **Normal — borrado:** `EliminarAdjunto` borra la fila y el fichero; si el llamador no pasó conexión, abre su propia transacción (`AS:495-496`), igual que hacen los repositorios de rechazo.
- **Alternativos:** `getAdjuntosPorSolicitudTx` usa la conexión transaccional del servicio y su gemelo sin `Tx` usa la conexión global (`AS:310-354`): la diferencia importa dentro de una transacción abierta.
- **Errores:** la validación lanza el error 513 de la aplicación, que el manejador de formularios presenta como validación y no como fallo técnico (`AS:283-291`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Un adjunto siempre pertenece a una etapa del flujo | `Validar`, regla 2 | `AS:289` |
| El nombre del fichero es obligatorio y tiene longitud acotada | `Validar`, regla 3 | `AS:291` |
| La etapa se almacena como código y se muestra como nombre | entidad y modelo de vista | `ADJ:19`, `AVM:26` |
| La etapa buscada se resuelve por nombre, no por código | `ExisteAdjuntoEtapa`, `GetIdAdjuntoPorEtapa` | `AS:197-221` |
| El cierre genera un adjunto de tipo PDF para firma | `GeneradoPDFParaFirma` | `AS:252-266` |

## Lecturas y escrituras

| Operación | Recurso | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Fila del adjunto | `tbAdjuntos` | `Recordset` y SQL parametrizada | La que pase el servicio | `AS:17-131`, `AS:310-354` |
| Fichero adjunto | directorio de adjuntos de la aplicación | escritura y borrado | Fuera de la transacción | `AS:17-131`, `AS:484-541` |
| Listado por solicitud | `tbAdjuntos` | lectura | Ninguna | `AS:355-412` |

## Coexistencia

La decisión de fondo es que hoy hay dos recursos que se guardan juntos —una fila y un fichero— sin transacción que los abarque. Si la copia del fichero falla después de insertar la fila, queda una fila sin fichero; si el borrado del fichero falla tras borrar la fila, queda un fichero huérfano. En el destino hay que decidir si el fichero pasa a un almacenamiento de objetos con su propia clave y su propia vida, y si la fila pasa a guardar esa clave en lugar de una ruta. La segunda decisión es `etapaWF` (`ADJ:19`): es el vínculo entre el adjunto y el estado del flujo, y las hojas de formularios ya muestran que se almacena como código mientras la interfaz lo muestra como nombre (`AVM:26`). Esa columna tiene que seguir siendo un catálogo, no un texto libre.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual.** El servicio valida la entidad, copia el fichero, guarda la fila y devuelve el identificador. Al eliminar, borra fila y fichero, abriendo transacción propia si nadie se la pasó. La etapa se resuelve por nombre en las consultas y se almacena como código.

**Posibles defectos, a verificar.**
1. La escritura del fichero no está dentro de la transacción de la base (`AS:17-131`): un fallo intermedio deja la fila sin fichero.
2. El borrado del fichero ocurre fuera de la transacción y sin reintento (`AS:484-541`).
3. `EliminarAdjuntosPorEtapas` borra varios ficheros en bucle (`AS:413-483`): un fallo intermedio deja el resto sin borrar.
4. La etapa se busca por nombre en unas consultas y por código en otras (`AS:197-221` frente a `AR:61-113`).

**Propuestas web pendientes de decisión.** Que el adjunto sea una entidad con su fichero en almacenamiento de objetos y una clave estable. Que la eliminación sea idempotente y que el fichero se elimine por collectors de huérfanos y no en la misma operación. Que la etapa sea un identificador de catálogo validado, no un nombre.

## Criterios de aceptación

1. Guardar un adjunto deja fila y fichero, o no deja ninguno. Trazado a `AS:17-131`.
2. Un adjunto sin etapa o sin nombre se rechaza con error de validación y no escribe nada. Trazado a `AS:283-291`.
3. El listado de una solicitud muestra el nombre de la etapa de cada adjunto. Trazado a `AS:355-412`.
4. Eliminar un adjunto borra fila y fichero, y repetir la operación no falla. Trazado a `AS:484-541`.
5. El cierre genera un adjunto PDF para firma y se detecta como tal. Trazado a `AS:222-266`.

## Pendientes

- Confirmar contra el binario el directorio de destino de los adjuntos y quién lo limpia (`AS:17-26`).
- Confirmar si existe una tarea programada que borre adjuntos huérfanos.
- Confirmar si los adjuntos de etapas ya formalizadas se conservan tras el cierre (`AS:413-483`).
- Confirmar el formato de la columna `etapaWF` y su catálogo, porque la matriz de campos lo necesita.