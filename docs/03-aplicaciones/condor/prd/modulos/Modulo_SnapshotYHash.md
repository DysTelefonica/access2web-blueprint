# Módulo de instantáneas y hash de la fase de validación

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/modules/SnapshotHelper.bas` (701 líneas, 13 procedimientos) · **PRD:** [README](../README.md)

`SnapshotHelper` construye una representación JSON del estado de una solicitud, la convierte en una huella y guarda ambas cosas. Es el módulo que da soporte al versionado de la validación de revisión: el hash que se almacena en `tbValidacionRevision.HashDatos` sale de aquí, de modo que el módulo participa del comportamiento observable aunque su nombre suene a diagnóstico.

Abreviaturas: `SH`, `modules/SnapshotHelper.bas`; `HH`, `modules/HashHelper.bas`; `FU`, `modules/FUNCIONES UTILES.bas`; `ENUM`, `modules/modEnumeradores.bas`.

## Superficie pública

| Objeto | Tipo | Procedimientos públicos | Qué resuelve |
|---|---|---|---|
| `SnapshotHelper` | módulo | `CalcularHashFaseValidacion`, `ExportarSnapshotParaDebug`, `GenerarSnapshotGlobal`, `ObtenerHashDesdeTablaRevision`, `PruebaDebugLog` (`SH:154-701`) | Serializar el estado de una solicitud, calcular su huella y volcarlo para depuración |

## Procedimientos y llamadas

| Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|
| `CalcularHashFaseValidacion` | Genera la instantánea de la solicitud y devuelve su huella | `GenerarSnapshotGlobal`, `HH.CalcularSHA256` | Sin escritura | `SH:154-175`; vacío si no hay datos en `SH:161-164` |
| `GenerarSnapshotGlobal` | Serializa todas las tablas que aplican al tipo de solicitud | `GenerarSnapshotDesdeTabla`, `ObtenerJsonDesdeRecordset` | Lectura de las tablas de la solicitud | `SH:478-521` |
| `ObtenerHashDesdeTablaRevision` | Recalcula la huella desde la tabla de revisión | `HH.CalcularSHA256` | Sin escritura | `SH:664-696` |
| `ExportarSnapshotParaDebug` | Escribe la instantánea a disco para diagnóstico | `GenerarSnapshotGlobal` | Escribe un fichero | `SH:176-259` |
| `InicializarDicts`, `ObtenerDictCampos` | Construyen el diccionario de campos por tipo de solicitud | — | Sin efectos | `SH:11-141`, `SH:142-153` |
| `ObtenerJsonDesdeRecordset`, `ObtenerJsonCamposEspecificos`, `EscapeJsonString` | Serialización interna | — | Sin efectos | `SH:319-399`, `SH:522-633`, `SH:634-663` |
| `PruebaDebugLog` | Prueba del registro de depuración | — | Sin efectos | `SH:697-701` |

## Flujos de trabajo

- **Normal:** el servicio de validación pide la huella de la fase (`SH:154`), que internamente serializa la solicitud completa (`SH:478`) y la resume. La huella se guarda en el ciclo de revisión con `ActualizarHashRevision` (hoja de repositorios de trazabilidad).
- **Alternativos:** si la solicitud no tiene datos, la huella sale vacía en lugar de la huella de una cadena vacía (`SH:161-164`). Existe una segunda vía de cálculo, directa desde la tabla de revisión (`SH:664`), cuyo resultado debería coincidir con la primera.
- **Errores:** cada procedimiento privado y público salta a una etiqueta `Errores` con `CondorError`; el fallo deja la huella vacía en lugar de propagar el error al flujo de negocio.

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| El conjunto de campos depende del tipo de solicitud | `ObtenerDictCampos` | `SH:142-153` |
| Sin datos no hay huella | `CalcularHashFaseValidacion` | `SH:161-164` |
| El escapado de cadenas es propio, no el del conversor JSON general | `EscapeJsonString` | `SH:634-663` |

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura del estado de la solicitud | tablas `tbDatosPC`, `tbDatosCDCA`, `tbDatosCDCASUB` y `tbSolicitudes` | `Recordset` serializado a JSON | La del servicio, si se pasa | `SH:478-521` |
| Recálculo desde la revisión | `tbValidacionRevision` | lectura | Ninguna | `SH:664-696` |
| Volcado para depuración | fichero local | escritura a disco | — | `SH:176-259` |

## Coexistencia

El hash de la fase de validación es el punto más delicado de este módulo. Se calcula sobre una serialización JSON del estado de la solicitud, y ese JSON depende del orden de los campos, del formato de las fechas y de la representación de los nulos. En el destino hay que decidir si se conserva el mismo algoritmo y la misma serialización —de modo que los hashes ya almacenados sigan siendo comparables— o si se acepta que el versionado empiece de cero con el primer ciclo migrado. `ExportarSnapshotParaDebug` escribe la instantánea en disco (`SH:176-259`), algo que en web se resuelve con un registro de diagnóstico, y que además condiciona qué datos personales podrían quedar expuestos en esos ficheros.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual.** El módulo serializa el estado de una solicitud a JSON, calcula su huella SHA-256 y puede volcar la instantánea a disco. El diccionario de campos se resuelve por tipo de solicitud.

**Posibles defectos, a verificar.**
1. Hay dos vías de cálculo de la huella (`SH:154-175` y `SH:664-696`): si divergen en el formato de entrada, producen valores distintos para el mismo estado.
2. El escapado de cadenas es una implementación propia (`SH:634-663`) en lugar del conversor general: son dos reglas de escapado que pueden divergir.
3. `ExportarSnapshotParaDebug` vuelca el estado completo de una solicitud a disco (`SH:176-259`), incluidos los campos de texto largo.
4. Ante un error, la huella queda vacía (`SH:154-175`): un fallo y una ausencia de datos son indistinguibles para quien la consume.

**Propuestas web pendientes de decisión.** Que el hash se calcule en el destino con un algoritmo documentado y que se decida explícitamente si los ciclos ya existentes se recomputan o se heredan. Que la instantánea sea un recurso de diagnóstico con permisos y retención, no un fichero suelto. Que la ausencia de huella por fallo se distinga de la ausencia por datos.

## Criterios de aceptación

1. Dos cálculos de la huella sobre el mismo estado devuelven el mismo valor. Trazado a `SH:154-175` y `SH:664-696`.
2. Una solicitud sin datos no produce huella. Trazado a `SH:161-164`.
3. El conjunto de campos del hash depende del tipo de solicitud. Trazado a `SH:142-153`.
4. El hash guardado en el ciclo de revisión coincide con el recalculado. Trazado a `SH:664-696`.

## Pendientes

- Confirmar contra el binario si las dos vías de cálculo producen hoy el mismo valor para una solicitud real.
- Confirmar si `ExportarSnapshotParaDebug` se invoca desde producción y con qué control de acceso (`SH:176-259`).
- Confirmar qué columnas de `tbValidacionRevision.HashDatos` se comparan al reabrir un ciclo, porque fija el formato que debe conservarse.