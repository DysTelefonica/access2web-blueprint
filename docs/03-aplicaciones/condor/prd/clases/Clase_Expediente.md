# Clases de expediente, sus documentos y sus proveedores

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/classes/DocumentoServicio.cls` (1095 líneas), `ExpedienteServicio.cls` (307), `Expediente.cls` (67), `ExpedienteSuministrador.cls` (51), `ExpedienteViewModel.cls` (36) · **PRD:** [README](../README.md)

El expediente no es de Condor: vive en la base de Lanzadera y esta aplicación lo consume. `ExpedienteServicio` lo lee y compone, y `Expediente` lo representa. `DocumentoServicio` es la pieza de mayor tamaño del grupo y trabaja con Microsoft Word: genera los documentos de la solicitud y sus PDF, e importa datos de un RAC en formato Word.

Abreviaturas: `ES`, `classes/ExpedienteServicio.cls`; `EXP`, `classes/Expediente.cls`; `ESUM`, `classes/ExpedienteSuministrador.cls`; `EVM`, `classes/ExpedienteViewModel.cls`; `DOC`, `classes/DocumentoServicio.cls`; `ER`, `modules/ExpedienteRepositorio.bas`; `SOL`, `classes/SolicitudServicio.cls`.

## Superficie pública

| Objeto | Tipo | Procedimientos y campos públicos | Qué resuelve |
|---|---|---|---|
| `Expediente` | clase de entidad | campos públicos del expediente (`EXP:15-66`), `getPropiedad` y `SetPropiedad` | Fila de expediente de la base de Lanzadera |
| `ExpedienteSuministrador` | clase de entidad | campos del vínculo (`ESUM:22-43`), `getPropiedad` y `SetPropiedad` (`ESUM:44-51`) | Relación entre expediente y suministro |
| `ExpedienteViewModel` | clase de presentación | `Expediente`, `NombresContratistas`, `NombresSubcontratistas` (`EVM:21-26`), `getPropiedad` y `SetPropiedad` (`EVM:28-36`) | Expediente con sus contratistas en una sola fila |
| `ExpedienteServicio` | clase de servicio | siete procedimientos (`ES:18-307`) | Búsqueda, lectura y composición del expediente, y la clave que da nombre al código de solicitud |
| `DocumentoServicio` | clase de servicio | doce procedimientos públicos (`DOC:26-1095`) | Generación del documento Word de cada tipo de solicitud, exportación a PDF e importación de RAC |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| `ES` | `getExpedientes` | Busca expedientes por palabra clave y otros criterios | `ER` | Lectura de la base de Lanzadera | `ES:18-88` |
| `ES` | `getExpedientePorID` | Devuelve un expediente | `ER.getExpedientePorID` | Lectura | `ES:89-105` |
| `ES` | `getClaveParaCodigoSolicitud` | Deriva la clave que compone el código de la solicitud | — | Sin escritura | `ES:106-161`; prioridad al código de actividad en `ES:110-113`, análisis del código corto en `ES:115-117` |
| `ES` | `getNombresSuministradores` | Devuelve los nombres de los suministros de un expediente | `ER` y consultas | Lectura | `ES:162-229` |
| `ES` | `getSuministradoresPorExpediente` | Devuelve los vínculos de suministro | `ER.getSuministradoresPorExpediente` | Lectura | `ES:230-251` |
| `ES` | `getExpedienteViewModelPorID` | Compone expediente, contratistas y subcontratistas | los anteriores | Lectura | `ES:252-292` |
| `ES` | `GetRacAsignadoPorExpediente` | Devuelve el RAC del expediente | `ER.GetRacAsignadoPorExpediente` | Lectura | `ES:293-307` |
| `DOC` | `GenerarDocumentoGenerico` | Rellena una plantilla de Word con los datos de la solicitud | Word por automatización | Escribe el documento | `DOC:291-471` |
| `DOC` | `GenerarDocumentoPC`, `GenerarDocumentoCDCA`, `GenerarDocumentoCDCASUB` | Eligen plantilla y datos según el tipo de solicitud | `GenerarDocumentoGenerico` | Escribe el documento | `DOC:275-277`, `DOC:279-281`, `DOC:283-290` |
| `DOC` | `GenerarDocumentoPDF_CDCASUB`, `GenerarDocumentoPDF_PC`, `GenerarDocumentoParaSolicitud` | Exporta a PDF | Word por automatización | Escribe el PDF | `DOC:26-83`, `DOC:539-598`, `DOC:599-655`; formato PDF en `DOC:50` |
| `DOC` | `ImportarDatosRACDesdeWord`, `ImportarDatosRACDesdeWord_CDCA` | Lee un documento Word y devuelve los datos del RAC | Word por automatización, en solo lectura | Sin escritura | `DOC:140-252`, `DOC:656-765`; apertura en `DOC:157` y `DOC:674` |
| `DOC` | `GetFormFieldValue`, `SetFormFieldValue` | Lee y escribe campos de formulario de Word | Word por automatización | Modifica el documento | `DOC:253-274`, `DOC:472-538` |
| `EXP`, `ESUM`, `EVM` | `getPropiedad`, `SetPropiedad` | Acceso reflectivo por nombre de propiedad | — | Sin efectos | `EXP:56-67`, `ESUM:44-51`, `EVM:28-36` |

## Flujos de trabajo

- **Normal — alta de solicitud:** el formulario elige un expediente y el servicio deriva de él la clave que compone el código (`ES:106-161`). La regla tiene prioridad absoluta para el código de actividad y, si falta, analiza el código corto (`ES:110-117`).
- **Normal — documento:** al cerrar una solicitud se genera el documento del tipo correspondiente (`DOC:275-290`), que es una plantilla de Word rellenada campo a campo (`DOC:291-471`), y se exporta a PDF (`DOC:539-598`).
- **Alternativos — importación:** una plantilla de RAC en Word se lee como documento de solo lectura y se convierte en datos (`DOC:140-252`, `DOC:656-765`), que luego alimentan el bloque de datos correspondiente.
- **Errores:** cada procedimiento de Word se apoya en el objeto de automatización; si Word no está disponible o el documento está bloqueado, el error sube como `CondorError`.

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| El código de actividad manda sobre el código corto para componer la clave | `getClaveParaCodigoSolicitud` | `ES:110-113` |
| Si no hay código de actividad, se analiza el formato del código corto | `getClaveParaCodigoSolicitud` | `ES:115-117` |
| Cada tipo de solicitud tiene su plantilla de documento | tres procedimientos de generación | `DOC:275-290` |
| La importación de RAC abre el documento en solo lectura | `ImportarDatosRAC*` | `DOC:157`, `DOC:674` |
| El PDF se genera con la exportación nativa de Word | `ExportAsFixedFormat` | `DOC:50`, `DOC:566` |

## Lecturas y escrituras

| Operación | Recurso | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura de expedientes y suministros | base de Lanzadera | SQL | Ninguna | `ES:18-251` |
| Generación de documentos | plantillas de Word en disco | escritura de fichero | — | `DOC:291-471` |
| Exportación a PDF | fichero PDF | escritura | — | `DOC:50`, `DOC:566` |
| Importación de RAC | documento de Word en disco | lectura | — | `DOC:140-252` |

## Coexistencia

Este grupo concentra dos decisiones grandes. La primera es el vínculo: `Expediente` se lee de la base de Lanzadera (`ER:27`, `ES:18-307`), de modo que el esquema destino tiene que decidir si el expediente se replica, se enlaza o se consulta por servicio. La segunda es Word: `DocumentoServicio` automatiza Microsoft Word para generar y para leer documentos (`DOC:44`, `DOC:155`, `DOC:347`, `DOC:557`, `DOC:672`), y esa automatización no tiene equivalente en un servicio web. Lo que debe sobrevivir es el **resultado**: la plantilla rellenada y el PDF. Las dos rutas posibles son convertir la plantilla a un formato de documento del destino y generarlo allí, o mantener la generación en un servicio de escritorio; la elección es del propietario. Lo mismo aplica a la importación: leer un Word desde un servidor exige una biblioteca de documentos y un contrato para los campos de formulario (`DOC:253-274`).

## Comportamiento actual · Posibles defects · Propuestas web pendientes de decisión

**Comportamiento actual.** Cuatro entidades sin comportamiento y dos servicios: uno lee y compone expedientes de otra base, el otro trabaja con Word para generar documentos e importar RAC.

**Posibles defectos, a verificar.**
1. Cinco instancias distintas de automatización de Word en el mismo servicio (`DOC:44`, `DOC:155`, `DOC:347`, `DOC:557`, `DOC:672`): no hay reutilización ni cierre garantizado.
2. La clave del código de solicitud depende de analizar texto del expediente (`ES:115-117`): cualquier cambio de formato en Lanzadera altera los códigos ya emitidos.
3. El expediente se lee de otra base sin caché ni invalidación (`ES:18-307`).
4. Los documentos y PDF se escriben en rutas locales, sin que el código indique quién los borra (`DOC:291-471`, `DOC:539-598`).

**Propuestas web pendientes de decisión.** Que el expediente viva en un esquema compartido con las demás aplicaciones, o que se consuma por servicio con caché declarado. Que la plantilla de documento se migre a un formato del destino y que la generación ocurra allí. Que la importación de RAC acepte un formato estructurado y deje de depender de campos de formulario de Word.

## Criterios de aceptación

1. El código de una solicitud se compone a partir del expediente con la regla de prioridad vigente. Trazado a `ES:106-161`.
2. Cada tipo de solicitud genera su documento con los datos de esa solicitud. Trazado a `DOC:275-290`.
3. El PDF generado corresponde al documento Word generado en la misma operación. Trazado a `DOC:539-598`.
4. La importación de un RAC devuelve los datos del bloque correspondiente sin modificar el documento de origen. Trazado a `DOC:140-252`.
5. El expediente se muestra con sus contratistas y subcontratistas en una sola vista. Trazado a `ES:252-292`.

## Pendientes

- Confirmar contra el binario dónde viven las plantillas de Word y si están versionadas con el código (`DOC:291-471`).
- Confirmar si la automatización de Word deja la aplicación colgada al no cerrar el proceso (`DOC:44`, `DOC:155`, `DOC:347`, `DOC:557`, `DOC:672`).
- Confirmar el formato exacto de los campos de formulario de la plantilla de RAC, porque condiciona la importación (`DOC:253-274`).
- Confirmar si `Expediente` se lee alguna vez de la base de Condor y no de la de Lanzadera.