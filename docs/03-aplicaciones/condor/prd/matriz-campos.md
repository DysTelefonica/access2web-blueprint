# Matriz campo a campo — Condor

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Issue:** #1075 · **Fase:** 3 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Cada columna de las 15 tablas del backend de Condor, con la evidencia de su uso en `src` y una correspondencia propuesta con el esquema de destino. Ninguna fila está aprobada: la topología PostgreSQL global sigue abierta y el módulo web de Condor todavía no existe en `app/src/modules`.

## Fuente y límites

- **Columnas de origen:** [`ERD/schema.sql`](../ERD/schema.sql), volcado del backend `Condor_Datos.accdb` con 25 tablas declaradas, de las que 15 están volcadas y 9 aparecen como enlazadas a ficheros externos que no estaban disponibles en la máquina de extracción. **El volcado es anterior al commit fijado**: incluye `tbDatosPCSUB`, tabla del tipo de solicitud retirado, de modo que no refleja el estado de `c88347e` (D-CON-1 del README). El ERD de origen pendiente del propietario (D-CON-5) es lo que cerrará esa diferencia.
- **Qué falta en esa fuente:** solo columna y tipo. No hay claves primarias, índices, relaciones, nulabilidad ni valores por defecto.
- **Evidencia en `src`:** para cada columna se busca su campo público en la entidad que la refleja, y si no aparece, una cita de la columna en un fichero de producción que nombre la tabla. Lectura estática sobre el clon de solo lectura del commit fijado.
- **Sin valores:** la matriz no contiene datos de ninguna tabla, solo estructura.

## Reglas de mapeo de tipos (propuesta)

| Tipo de origen | Tipo objetivo propuesto | Decisión abierta |
|---|---|---|
| `LONG(4)` | `INTEGER` | Identidad: la clave se calcula hoy con `Max(id) + 1` (hoja de repositorios de datos) y en el destino debería ser una secuencia; la correspondencia de claves entre Access y web no está definida |
| `TEXT(n)` | `VARCHAR(n)` | Los indicadores «Sí»/«No» en texto deben decidir si pasan a `BOOLEAN` |
| `MEMO` | `TEXT` | Los campos de texto largo que hoy se truncan al concatenar en SQL se escriben con `Recordset`; el destino no tiene ese problema |
| `BOOLEAN(1)` | `BOOLEAN` | — |
| `SHORT_DATE_TIME(8)` | `TIMESTAMP` | Fecha o fecha y hora, y zona horaria |

## Clase de dato

La columna «Clase de dato» es una **heurística por nombre de columna**, sin revisar contenido: `credencial` (contraseñas, hashes y claves) o `personal` (nombres, correos, firmas, descripciones, motivos y comentarios) o «—». Sirve para priorizar la revisión de privacidad y de cifrado y no sustituye a una clasificación formal.

## Correspondencia con el destino

No hay módulo web de Condor en `app/src/modules` (solo `expedientes` y `lanzadera`), de modo que la columna «Correspondencia» nombra la **tabla candidata** en el esquema propuesto `condor` y deja la columna sin asignar. La asignación columna a columna llegará con el ERD global y con la decisión sobre identidad compartida con Lanzadera (hoja de coexistencia).

## Leyenda de la columna «Evidencia en `src`»

- `P`: la columna es un campo público de la entidad que la refleja; es la vía por la que `RepositorioComun` hidrata y rellena el `Recordset`.
- `S`: el nombre aparece en un fichero de producción que cita la tabla, sin encontrar un campo de entidad; coincidencia de nombre.
- `—`: sin cita en los ficheros de producción. En las dos tablas retiradas es la señal de que la tabla ya no existe en el commit fijado.

## Resumen por tabla

| Tabla | Columnas | Con evidencia | Ficheros de `src` que citan la tabla | Familia | Detalle |
|---|---|---|---|---|---|
| `tbAdjuntos` | 8 | 8 | 2 | catálogos y trazabilidad | [catalogos-y-trazabilidad](matriz-campos/catalogos-y-trazabilidad.md) |
| `tbDatosCDCA` | 40 | 40 | 4 | solicitud y bloques de datos | [solicitud-y-bloques](matriz-campos/solicitud-y-bloques.md) |
| `tbDatosCDCASUB` | 44 | 44 | 5 | solicitud y bloques de datos | [solicitud-y-bloques](matriz-campos/solicitud-y-bloques.md) |
| `tbDatosPC` | 41 | 41 | 4 | solicitud y bloques de datos | [solicitud-y-bloques](matriz-campos/solicitud-y-bloques.md) |
| `tbDatosPCSUB` | 43 | 0 | 0 | retirada en el commit fijado | [solicitud-y-bloques](matriz-campos/solicitud-y-bloques.md) |
| `tbEstados` | 6 | 6 | 2 | catálogos y trazabilidad | [catalogos-y-trazabilidad](matriz-campos/catalogos-y-trazabilidad.md) |
| `tbHistorialRechazos` | 8 | 0 | 0 | retirada en el commit fijado | [solicitud-y-bloques](matriz-campos/solicitud-y-bloques.md) |
| `tbLogCambios` | 10 | 10 | 5 | catálogos y trazabilidad | [catalogos-y-trazabilidad](matriz-campos/catalogos-y-trazabilidad.md) |
| `tbLogErrores` | 9 | 9 | 6 | catálogos y trazabilidad | [catalogos-y-trazabilidad](matriz-campos/catalogos-y-trazabilidad.md) |
| `tbLogEstados` | 6 | 6 | 3 | catálogos y trazabilidad | [catalogos-y-trazabilidad](matriz-campos/catalogos-y-trazabilidad.md) |
| `tbMapeoCampos` | 6 | 6 | 3 | catálogos y trazabilidad | [catalogos-y-trazabilidad](matriz-campos/catalogos-y-trazabilidad.md) |
| `tbRechazos` | 10 | 9 | 2 | catálogos y trazabilidad | [catalogos-y-trazabilidad](matriz-campos/catalogos-y-trazabilidad.md) |
| `tbSolicitudes` | 12 | 12 | 12 | solicitud y bloques de datos | [solicitud-y-bloques](matriz-campos/solicitud-y-bloques.md) |
| `tbTransiciones` | 4 | 3 | 2 | catálogos y trazabilidad | [catalogos-y-trazabilidad](matriz-campos/catalogos-y-trazabilidad.md) |
| `tbValidacionRevision` | 10 | 10 | 2 | catálogos y trazabilidad | [catalogos-y-trazabilidad](matriz-campos/catalogos-y-trazabilidad.md) |

## Totales y reconciliación

- **Columnas volcadas:** 257 en 15 tablas. Con evidencia: 204 (campo de entidad `P`: 200; coincidencia de nombre `S`: 4; sin cita: 53).
- **Tablas del volcado sin ninguna cita en `src`:** `tbDatosPCSUB` (43 columnas) y `tbHistorialRechazos` (8 columnas). Son las dos retiradas con el tipo `PC_SUB` (D-CON-1); sus columnas se listan sin destino y quedan a la espera del ERD de origen.
- **Columnas sin campo de entidad dentro de tablas vigentes:** `tbRechazos.idRechazo` y `tbTransiciones.idTransicion`. Se escriben solo por SQL y no aparecen en ninguna entidad, de modo que su destino debe decidir si conserva la clave primaria.
- **Tablas enlazadas:** nueve, declaradas en el volcado y ausentes de él porque los ficheros externos no estaban en la máquina de extracción. Las nueve tienen cita en `src`; se detallan en la parte de tablas enlazadas.
- **Pendiente de origen:** el ERD de origen pendiente del propietario (D-CON-5) debe confirmar claves, índices, nulabilidad, valores por defecto y relaciones. Hasta entonces, todas las filas son **candidatas a nivel de tabla**, no asignaciones.
