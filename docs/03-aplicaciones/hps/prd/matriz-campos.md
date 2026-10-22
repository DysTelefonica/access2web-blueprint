# Matriz campo a campo — HPS

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Issue:** #1076 · **Fase:** 3 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Cada columna de origen de las 23 tablas del backend de HPST, con la evidencia de su uso en `src`. **Ninguna fila está aprobada**: la topología PostgreSQL global y el modelo de identidad siguen abiertos. En total **298 columnas**, repartidas en seis familias; las que van llegando en las siguientes PRs de la cadena.

## Fuente y límites

- **Columnas de origen:** [`ERD/schema.sql`](../ERD/schema.sql), volcado de `HPST.accdb` del snapshot (23 tablas, 298 columnas).
- **Qué falta en esa fuente:** el volcado da **solo columna y tipo**. No hay claves primarias, índices, relaciones, nulabilidad ni valores por defecto. Sin eso la matriz describe estructura, no modelo.
- **Evidencia en `src`:** se busca cada columna en los ficheros de producción (excluidos `Test*`, `clsTestDouble*` y `Dummy*`) que citan la tabla. Lectura estática.
- **Sin valores:** la matriz no contiene datos de ninguna tabla, solo estructura.
- **Contraste con las fichas:** cinco tablas no las cita ningún fichero de producción; se escriben por DAO con `Fields(...)` dinámico o no se usan desde `src`.

## Reglas de mapeo de tipos (propuesta)

| Tipo de origen | Tipo objetivo propuesto | Decisión abierta |
|---|---|---|
| `LONG` | `INTEGER` | Identidad: el módulo web usa otros identificadores y la correspondencia de claves entre Access y web **no está definida** |
| `TEXT(n)` | `VARCHAR(n)` | `Sí`/`No` en texto: conservar hasta decidir `BOOLEAN` |
| `MEMO` | `TEXT` | — |
| `SHORT_DATE_TIME` | `TIMESTAMP` | Fecha o fecha y hora; zona horaria |
| `BOOLEAN` | `BOOLEAN` | — |

## Clase de dato

Heurística por nombre de columna, sin revisar contenido: `credencial` (contraseñas e históricos), `personal` (nombre, correo, DNI, teléfonos, nacimiento, usuario) o «—». Sirve para priorizar la revisión de privacidad (RGPD) y de cifrado. Las credenciales se migran, si se migran, con tratamiento propio y **nunca en claro**; esa decisión es del propietario.

## Correspondencia con la web actual

El único módulo que hoy toca HPS es [`app/src/modules/expedientes/application/hps`](../../../../app/src/modules/expedientes/application/hps), que sirve el alta de usuario hacia HPS, y **no tiene esquema propio**. La web no tiene tabla para la mayor parte del legacy, así que la correspondencia queda sin decidir y es trabajo de este PRD.

## Leyenda de la columna «Evidencia»

- `F`: la columna aparece como `Fields("columna")` en un fichero que cita la tabla.
- `S`: el nombre aparece en un fichero que cita la tabla, sin `Fields(...)`; coincidencia de nombre, poco fiable para nombres comunes.
- `—`: sin cita en los ficheros de producción.

## Familias

| Familia | Tablas | Estado |
|---|---|---|
| [identidad, usuarios y credenciales](matriz-campos/identidad.md) | 6 | llega en las siguientes PRs de la cadena |
| [entidades y relación con el expediente](matriz-campos/entidades.md) | 2 | llega en las siguientes PRs de la cadena |
| [los cuatro tipos de HPS](matriz-campos/hps.md) | 4 | [detalle](matriz-campos/hps.md) |
| [anexos y observaciones](matriz-campos/anexos.md) | 6 | llega en las siguientes PRs de la cadena |
| [expedientes y catalogación de contratación](matriz-campos/expedientes.md) | 3 | [detalle](matriz-campos/expedientes.md) |
| [catálogo y soporte](matriz-campos/catalogo.md) | 2 | [detalle](matriz-campos/catalogo.md) |

## Resumen por tabla

| Tabla | Columnas | Filas (snapshot) | Ficheros de `src` | Familia | Detalle |
|---|---|---|---|---|---|
| Copia de TbExpedienteLugares | 4 | 87 | 0 | expedientes y catalogación de contratación | [detalle](matriz-campos/expedientes.md#copia-de-tbexpedientelugares) |
| Copia de TbExpedientes | 7 | 70 | 0 | expedientes y catalogación de contratación | [detalle](matriz-campos/expedientes.md#copia-de-tbexpedientes) |
| Copia de TbUsuarios | 18 | 335 | 0 | identidad, usuarios y credenciales | — |
| Copia de TbUsuariosEntidades | 78 | 331 | 0 | entidades y relación con el expediente | — |
| Errores de pegado | 18 | 1 | 0 | identidad, usuarios y credenciales | — |
| TbAnexosUsuariosHPS | 5 | 2183 | 4 | anexos y observaciones | — |
| TbAnexosUsuariosHistoricos | 4 | 1395 | 3 | anexos y observaciones | — |
| TbAnexosUsuariosSICA | 4 | 9 | 2 | anexos y observaciones | — |
| TbAuxCursos | 1 | 90 | 0 | catálogo y soporte | [detalle](matriz-campos/catalogo.md#tbauxcursos) |
| TbConfiguracionUsuariosAccesoLocal | 4 | 4 | 1 | identidad, usuarios y credenciales | — |
| TbConsultas | 2 | 8 | 1 | catálogo y soporte | [detalle](matriz-campos/catalogo.md#tbconsultas) |
| TbHPS | 5 | 1360 | 8 | los cuatro tipos de HPS | [detalle](matriz-campos/hps.md#tbhps) |
| TbHPSEquivalencia | 4 | 3 | 0 | los cuatro tipos de HPS | [detalle](matriz-campos/hps.md#tbhpsequivalencia) |
| TbHPSGrado | 2 | 13 | 1 | los cuatro tipos de HPS | [detalle](matriz-campos/hps.md#tbhpsgrado) |
| TbJuridicasContratacion | 1 | 3 | 0 | expedientes y catalogación de contratación | [detalle](matriz-campos/expedientes.md#tbjuridicascontratacion) |
| TbMotivoHPS | 1 | 58 | 2 | los cuatro tipos de HPS | [detalle](matriz-campos/hps.md#tbmotivohps) |
| TbObservaciones | 4 | 340 | 4 | anexos y observaciones | — |
| TbObservacionesHistoricas | 4 | 365 | 3 | anexos y observaciones | — |
| TbUsuarioAnexos | 4 | 385 | 1 | anexos y observaciones | — |
| TbUsuarios | 19 | 368 | 24 | identidad, usuarios y credenciales | — |
| TbUsuariosEntidades | 78 | 363 | 9 | entidades y relación con el expediente | — |
| TbUsuariosHistoricos | 19 | 247 | 12 | identidad, usuarios y credenciales | — |
| TbUsuariosSICA | 12 | 50 | 9 | identidad, usuarios y credenciales | — |

**Total: 23 tablas, 298 columnas.** Ficheros de producción que citan alguna tabla: 84.
