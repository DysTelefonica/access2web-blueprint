# Clases de entorno, mapeo, instantánea y visor web

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/classes/Entorno.cls` (645 líneas), `WebVisorCacheServicio.cls` (180), `SnapshotServicio.cls` (103), `MapeoServicio.cls` (88), `MapeoCampos.cls` (49) · **PRD:** [README](../README.md)

Cinco clases de apoyo al formulario de gestión y al visor de HTML. `Entorno` es el objeto que vive en la sesión y mantiene catálogos y datos de contexto; `WebVisorCacheServicio` decide si el HTML que se muestra se genera o se lee de una caché en disco; `SnapshotServicio` produce la instantánea que alimenta el hash de la fase de validación; y el par de mapeo traduce entre los nombres de campo de la base y los nombres de las plantillas de documento.

Abreviaturas: `ENT`, `classes/Entorno.cls`; `WVC`, `classes/WebVisorCacheServicio.cls`; `SNAP`, `classes/SnapshotServicio.cls`; `MS`, `classes/MapeoServicio.cls`; `MC`, `classes/MapeoCampos.cls`; `SH`, `modules/SnapshotHelper.bas`; `MR`, `modules/MapeoRepositorio.bas`; `VG`, `modules/Variables Globales.bas`.

## Superficie pública

| Objeto | Tipo | Procedimientos y campos públicos | Qué resuelve |
|---|---|---|---|
| `Entorno` | clase de sesión | cuatro procedimientos de refresco (`ENT:176-195`) y el resto de la superficie hasta `ENT:632` | Estado de contexto de la sesión: catálogos, responsables, entorno de trabajo y datos auxiliares |
| `WebVisorCacheServicio` | clase de servicio | `ObtenerHTMLCacheable`, `ExisteCache` (`WVC:36-180`), inicialización y terminación propias | Decide si el HTML del visor se genera o se lee de la caché en disco |
| `SnapshotServicio` | clase de servicio | `GenerarSnapshotGlobal`, `CalcularHashSnapshot` (`SNAP:19-57`) | Instantánea JSON de una solicitud y su huella |
| `MapeoServicio` | clase de servicio | `getMapeoPC`, `getMapeoCDCA`, `getMapeoCDCASUB`, `SembrarBookmarksDesdePlantilla` (`MS:18-88`) | Mapeo de campos por tipo de solicitud y siembra de marcadores de una plantilla |
| `MapeoCampos` | clase de entidad | campos del mapeo (`MC:…`), `getPropiedad` y `SetPropiedad` (`MC:41-49`) | Fila de `tbMapeoCampos` |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| `ENT` | `RefreshEstados`, `RefreshResponsablesTecnicos`, `RefreshResponsablesCalidad`, `RefreshSuministradores` | Recargan catálogos dentro del objeto de sesión | servicios y repositorios de catálogo | Sin escritura | `ENT:176-195` |
| `ENT` | superficie restante hasta `getPropiedad` | Acceso reflectivo y estado auxiliar | — | Sin efectos | `ENT:632` |
| `WVC` | `Class_Initialize`, `Class_Terminate` | Preparan y liberan la caché al crear y destruir el objeto | — | Sin efectos | `WVC:23-35` |
| `WVC` | `ObtenerHTMLCacheable` | Devuelve el HTML de la caché si el modo de carga lo permite y hay contenido | `ObtenerModoCarga`, `LeerHTMLCacheado` | Lectura de disco | `WVC:36-65`; valor por defecto en `WVC:39`, modo en `WVC:42-46`, acierto de caché en `WVC:48-52` |
| `WVC` | `LeerHTMLCacheado`, `GuardarHTMLEnCache`, `ObtenerRutaHTML`, `ObtenerDirectorioHTML` | Resuelven la ruta del fichero y su contenido | — | Lectura y escritura en disco | `WVC:78-129`, `WVC:103-129`, `WVC:146-180` |
| `WVC` | `ExisteCache` | Indica si hay HTML cacheado para un tipo | — | Sin escritura | `WVC:130-145` |
| `SNAP` | `GenerarSnapshotGlobal` | Devuelve el JSON con los datos de todas las pestañas de la solicitud, excluyendo campos de auditoría | `ObtenerDatosTabla` | Lectura | `SNAP:19-51`; cabecera en `SNAP:29-30`, bloque PC en `SNAP:32-33` |
| `SNAP` | `CalcularHashSnapshot` | Devuelve la huella de la instantánea | `SH.CalcularHashFaseValidacion` | Sin escritura | `SNAP:52-57` |
| `SNAP` | `ObtenerDatosTabla`, `EsCampoExcluido` | Leen una tabla por clave y filtran campos | — | Lectura | `SNAP:58-103` |
| `MS` | `getMapeoPC`, `getMapeoCDCA`, `getMapeoCDCASUB` | Devuelven el mapeo de campos del tipo correspondiente | `MR.getMapeoParaPlantilla` | Lectura | `MS:18-50` |
| `MS` | `SembrarBookmarksDesdePlantilla` | Inserta marcadores en la plantilla a partir del mapeo | Word por automatización | Modifica la plantilla | `MS:51-88` |
| `MC` | `getPropiedad`, `SetPropiedad` | Acceso reflectivo por nombre | — | Sin efectos | `MC:41-49` |

## Flujos de trabajo

- **Normal — visor:** `WorkflowServicio` genera el HTML del visor y `ObtenerHTMLCacheable` decide si se muestra el generado o el de la caché (`WVC:36-52`); el modo de carga se consulta antes (`WVC:41-46`).
- **Normal — validación:** la instantánea se genera excluyendo campos de auditoría (`SNAP:21`, `SNAP:91-103`) y su huella se guarda en el ciclo de revisión (`SNAP:52-57`).
- **Alternativos:** `ObtenerHTMLCacheable` devuelve siempre el HTML generado si el modo no permite caché (`WVC:39`, `WVC:44-46`), de modo que la caché es una optimización y no un requisito.
- **Errores:** `ObtenerHTMLCacheable` captura el fallo y devuelve el HTML generado (`WVC:36-39`): la caché nunca impide mostrar el visor.

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| La caché es una optimización, nunca un requisito | `ObtenerHTMLCacheable` | `WVC:39`, `WVC:44-46` |
| La instantánea excluye los campos de auditoría | `EsCampoExcluido` | `SNAP:21`, `SNAP:91-103` |
| Hay un mapeo de campos por tipo de solicitud | tres procedimientos de mapeo | `MS:18-50` |
| Los catálogos de la sesión se recargan a demanda | cuatro `Refresh*` | `ENT:176-195` |
| La huella se calcula sobre la instantánea completa | `CalcularHashSnapshot` | `SNAP:52-57` |

## Lecturas y escrituras

| Operación | Recurso | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Instantánea | `tbSolicitudes`, `tbDatosPC`, `tbDatosCDCA`, `tbDatosCDCASUB` | lectura | Ninguna | `SNAP:29-33`, `SNAP:58-103` |
| HTML del visor | ficheros de caché | lectura y escritura en disco | — | `WVC:78-180` |
| Mapeo de campos | `tbMapeoCampos` | lectura | Ninguna | `MS:18-50` |
| Plantilla de documento | fichero de plantilla | escritura de marcadores | — | `MS:51-88` |

## Coexistencia

La caché del visor (`WVC:78-180`) y la siembra de marcadores en la plantilla (`MS:51-88`) son operaciones sobre ficheros locales del cliente, y no existe equivalente razonable en un servicio: la caché del destino tendrá que vivir en el servidor o en el navegador, y la plantilla tendrá que ser un documento que el servidor pueda modificar. La instantánea (`SNAP:19-103`), en cambio, es pura lectura y debe sobrevivir tal cual, porque su formato de serialización es el que alimenta el hash de la fase de validación: si cambia, los hashes ya almacenados dejan de ser comparables. `Entorno` (`ENT:176-195`) es el estado de sesión que en web se sustituye por la sesión del servidor más una caché de catálogos.

## Comportamiento actual · Posibles defects · Propuestas web pendientes de decisión

**Comportamiento actual.** Un objeto de sesión con catálogos, una caché de HTML en disco, un generador de instantáneas con exclusión de campos de auditoría y un servicio de mapeo por tipo de solicitud.

**Posibles defectos, a verificar.**
1. La caché del visor vive en un directorio del cliente (`WVC:146-180`): dos usuarios de la misma estación comparten caché.
2. `SnapshotServicio` y `SnapshotHelper` generan instantáneas de la misma solicitud (`SNAP:19-51`, hoja de `SH:478-521`): son dos rutas que deben producir el mismo JSON.
3. `Entorno` mantiene catálogos en memoria sin invalidación declarada (`ENT:176-195`).
4. `SembrarBookmarksDesdePlantilla` modifica la plantilla en el cliente (`MS:51-88`).

**Propuestas web pendientes de decisión.** Que la caché del visor viva en el servidor con su política de invalidación. Que exista una única implementación de la instantánea. Que el mapeo de campos sea un recurso del servidor y que la plantilla se procese allí.

## Criterios de aceptación

1. El visor muestra el HTML generado cuando el modo no permite caché. Trazado a `WVC:39`, `WVC:44-46`.
2. La instantánea no contiene campos de auditoría. Trazado a `SNAP:21`, `SNAP:91-103`.
3. La huella de la instantánea coincide con la que se guarda en el ciclo. Trazado a `SNAP:52-57`.
4. Cada tipo de solicitud tiene su propio mapeo de campos. Trazado a `MS:18-50`.

## Pendientes

- Confirmar contra el binario si `SnapshotServicio` y `SnapshotHelper` están ambos en uso y si sus instantáneas coinciden (`SNAP:19-51`, `SH:478-521`).
- Confirmar el directorio de caché del visor y su política de limpieza (`WVC:146-180`).
- Confirmar qué plantilla se siembra y si se modifica en el cliente (`MS:51-88`).
- Confirmar la lista completa de campos excluidos de la instantánea, porque determina qué entra en el hash (`SNAP:91-103`).