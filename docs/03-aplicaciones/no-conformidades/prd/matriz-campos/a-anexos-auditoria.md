# Matriz campo a campo: anexos y auditorías

Esta parte de la matriz cubre siete tablas del esquema previo: `Copia de TbNCARAvisos`, `TbAnexos`, `TbAnexosAuditoria`, `TbAnexosNCAuditorias`, `TbAuditoriaLog`, `TbAuditorias` y `TbAuxPuntoNorma`. La línea base es el código de `main` (producción) del clon fijado en el commit `1d7a2f8`; las citas `fichero:línea` apuntan a `src/forms`, `src/modules` y `src/classes` de ese clon. El esquema previo (`ERD/schema.sql`) aporta solo nombres y tipos de columna, sin claves, relaciones ni nulabilidad, y no es autoridad. Por eso las relaciones se marcan «inferida por código» con su cita, o «sin evidencia». El backend es `NoConformidades_Datos.accdb`, cuyo ERD de origen real sigue pendiente del propietario; todos los destinos `no_conformidades.<tabla>.<columna>` y los tipos PostgreSQL son orientativos y quedan como propuesta hasta que se apruebe el ERD global. Las columnas sin uso en `main` se dejan como «pendiente» con su motivo. El resumen general de formularios está en [el README del PRD](../README.md).

Búsqueda común a todas las tablas: coincidencia literal de cada nombre de tabla sobre todo `src` (`forms`, `modules`, `classes`, `queries`). En `src` no aparece ninguna instrucción `BeginTrans`, `CommitTrans` ni `Rollback` sobre estas tablas; los únicos `CommitTransaccionCache` y `RollbackTransaccionCache` están en `CacheNCProyecto.bas:646` y `CacheNCProyecto.bas:689` y no tocan estas tablas (a verificar si el motor aplica alguna transacción implícita).

## Copia de TbNCARAvisos

**Uso en `main`.** Ninguno. La búsqueda literal de `Copia de TbNCARAvisos` y de `NCARAvisos` en `src` no devuelve coincidencias, ni en código ni en `queries`. La tabla parece una copia de seguridad de otra tabla de avisos de acciones realizadas (a verificar con el propietario).

**Claves y relaciones.** Sin evidencia. Por nombre, `IDAR` sugiere un vínculo con acciones realizadas, pero ningún código lo confirma.

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | Copia de TbNCARAvisos | ID | LONG(4) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | Copia de TbNCARAvisos | IDAR | LONG(4) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | Copia de TbNCARAvisos | IDCorreo15 | LONG(4) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | Copia de TbNCARAvisos | IDCorreo7 | LONG(4) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | Copia de TbNCARAvisos | IDCorreo0 | LONG(4) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | Copia de TbNCARAvisos | Fecha | SHORT_DATE_TIME(8) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |

## TbAnexos

**Uso en `main`.** Ninguno. La búsqueda literal de `TbAnexos` en `src` solo encuentra los nombres de las dos tablas hermanas `TbAnexosAuditoria` y `TbAnexosNCAuditorias` (`Instalador.bas:126`, `Instalador.bas:170`); no hay referencia a `TbAnexos` como tabla propia. Los documentos de una no conformidad de auditoría se leen hoy de `TbDocumentosAuditorias` por `IDNoConformidad` (`constructor.bas:4924-4927`); para proyectos, no hay evidencia en estas búsquedas (a verificar con las fichas de [documentos de NC de proyecto](../formularios/Form_FormNCProyectoDocumentos.md)).

**Claves y relaciones.** Sin evidencia. Por nombre, `IDNoConformidad` apunta a una no conformidad, sin cita que lo confirme.

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbAnexos | IDAnexo | LONG(4) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbAnexos | IDNoConformidad | LONG(4) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbAnexos | TituloAnexo | TEXT(255) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbAnexos | DescripcionAnexo | MEMO | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbAnexos | URLInicial | TEXT(255) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbAnexos | NombreArchivoFinalAnexo | TEXT(255) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbAnexos | FechaAnexo | SHORT_DATE_TIME(8) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |

## TbAnexosAuditoria

**Quién la lee y escribe.** Solo la lee `RellenarAnexosAuditorias`, mediante DAO sobre la tabla completa (`Instalador.bas:126-127`), recorriendo todos los registros (`Instalador.bas:129`, `Instalador.bas:161`). Por cada registro con `IDAuditoria` numérico (`Instalador.bas:130-131`) toma `NombreArchivo` (`Instalador.bas:132`), le quita la extensión (`Instalador.bas:133`) y busca en `TbDocumentosAuditorias` con SQL dinámico concatenado (`Instalador.bas:134-137`). Si no existe, inserta allí un documento mediante DAO con `IDDocumento`, `IDAuditoria`, `Documento` y `NombreAnexo` (`Instalador.bas:150-156`). Nunca escribe en `TbAnexosAuditoria`. No hay transacción. La rutina es una carga de migración: la invoca `InstalarProyecto` (`Instalador.bas:100`) y no se encuentra ningún llamador de `InstalarProyecto` en `src` (búsqueda literal), así que no forma parte del flujo de usuario de [las fichas de formularios](../README.md) (a verificar si se ejecutó en producción).

**Claves y relaciones.**
- `IDAuditoria` referencia a `TbAuditorias.IDAuditoria`: inferida por código (`Instalador.bas:130-136`, donde el mismo valor filtra `TbDocumentosAuditorias.IDAuditoria`).
- Clave primaria: sin evidencia (`IDAnexo` no se lee en `src`).

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbAnexosAuditoria | IDAnexo | LONG(4) | sin destino propuesto | — | — | Sin evidencia: no se lee en src | Sin evidencia | pendiente: sin uso de la columna en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbAnexosAuditoria | IDAuditoria | LONG(4) | no_conformidades.documentos_auditoria.id_auditoria (candidata, si no se descarta la tabla; ver nota) | integer | Directa; el contenido ya se vuelca a `TbDocumentosAuditorias` (`Instalador.bas:151-152`) | Se tolera vacío: `Nz(..., "")` y se omite si no es numérico (`Instalador.bas:130-131`); duplicados: sin evidencia | Debe ser numérico (`Instalador.bas:131`); existencia en `TbAuditorias`: sin evidencia | propuesta |
| NoConformidades_Datos.accdb | TbAnexosAuditoria | URLInicial | TEXT(255) | sin destino propuesto | — | — | Sin evidencia: no se lee en src | Sin evidencia | pendiente: sin uso de la columna en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbAnexosAuditoria | NombreArchivo | TEXT(255) | no_conformidades.documentos_auditoria.documento (candidata, si no se descarta la tabla; ver nota) | varchar(255) | Se conserva el nombre sin extensión (`Instalador.bas:133`, `Instalador.bas:154`); `NombreAnexo` se recalcula con `getURLAnexoFinal` (`Instalador.bas:145-146`) | Sin `Nz`: un valor nulo provocaría error en la asignación (`Instalador.bas:132`, a verificar); duplicados: se evita insertar si ya hay fila con mismo `IDAuditoria` y `Documento` (`Instalador.bas:134-139`) | Sin evidencia | propuesta |
| NoConformidades_Datos.accdb | TbAnexosAuditoria | FechaAnexo | SHORT_DATE_TIME(8) | sin destino propuesto | — | — | Sin evidencia: no se lee en src | Sin evidencia | pendiente: sin uso de la columna en main, decisión del propietario |

Nota: dado que la única lectura vuelca el contenido a `TbDocumentosAuditorias`, la propuesta es no crear una tabla equivalente y migrar solo lo que ese volcado aún no haya cubierto; es decisión del propietario (a verificar).

## TbAnexosNCAuditorias

**Quién la lee y escribe.** Mismo patrón que la tabla anterior, en la segunda mitad de `RellenarAnexosAuditorias`: lectura DAO de la tabla completa (`Instalador.bas:170-171`), uso de `IDNoConformidad` (`Instalador.bas:174-175`) y `NombreArchivo` (`Instalador.bas:176-177`), búsqueda en `TbDocumentosAuditorias` con SQL dinámico concatenado por `IDNoConformidad` y `Documento` (`Instalador.bas:178-181`) e inserción DAO condicional (`Instalador.bas:194-200`). Nunca escribe en `TbAnexosNCAuditorias`. No hay transacción. Misma salvedad sobre el llamador (`Instalador.bas:100`).

**Claves y relaciones.**
- `IDNoConformidad` referencia a la no conformidad de auditoría: inferida por código (`Instalador.bas:187`, `Instalador.bas:196`, donde se escribe como `IDNoConformidad` de un documento de auditoría; el nombre de la tabla de NC de auditoría no figura en esas líneas).
- Clave primaria: sin evidencia.

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbAnexosNCAuditorias | IDAnexo | LONG(4) | sin destino propuesto | — | — | Sin evidencia: no se lee en src | Sin evidencia | pendiente: sin uso de la columna en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbAnexosNCAuditorias | IDNoConformidad | LONG(4) | no_conformidades.documentos_auditoria.id_no_conformidad (candidata, si no se descarta la tabla; ver nota) | integer | Directa (`Instalador.bas:196`) | Se tolera vacío: `Nz(..., "")` y se omite si no es numérico (`Instalador.bas:174-175`); duplicados: sin evidencia | Debe ser numérico (`Instalador.bas:175`); existencia de la NC: sin evidencia | propuesta |
| NoConformidades_Datos.accdb | TbAnexosNCAuditorias | URLInicial | TEXT(255) | sin destino propuesto | — | — | Sin evidencia: no se lee en src | Sin evidencia | pendiente: sin uso de la columna en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbAnexosNCAuditorias | NombreArchivo | TEXT(255) | no_conformidades.documentos_auditoria.documento (candidata, si no se descarta la tabla; ver nota) | varchar(255) | Se conserva el nombre sin extensión (`Instalador.bas:177`, `Instalador.bas:198`); `NombreAnexo` se recalcula (`Instalador.bas:189-190`) | Sin `Nz`: un valor nulo provocaría error (`Instalador.bas:176`, a verificar); duplicados: se evita insertar si ya hay fila con mismo `IDNoConformidad` y `Documento` (`Instalador.bas:178-183`) | Sin evidencia | propuesta |
| NoConformidades_Datos.accdb | TbAnexosNCAuditorias | FechaAnexo | SHORT_DATE_TIME(8) | sin destino propuesto | — | — | Sin evidencia: no se lee en src | Sin evidencia | pendiente: sin uso de la columna en main, decisión del propietario |

Nota: igual que en `TbAnexosAuditoria`, la propuesta es absorber el contenido en la tabla de documentos de auditoría; decisión del propietario.

## TbAuditoriaLog

**Uso en `main`.** Ninguno con este nombre. La búsqueda literal de `TbAuditoriaLog` y de `AuditoriaLog` en `src` no devuelve coincidencias. La bitácora de auditoría se escribe y se lee en `TbLogAuditoria`, otra tabla del mismo esquema con las mismas ocho columnas y tipos: alta DAO en `LogNCAuditoria.cls:215-241`, cálculo de `IDLog` en `LogNCAuditoria.cls:137` y lectura por `IDNC`, `IDAC` o `IDAR` en `constructor.bas:1296-1298`. Que `TbAuditoriaLog` sea una copia o un resto de `TbLogAuditoria` es una hipótesis a verificar con el propietario; las columnas de `TbLogAuditoria` se tratarán en su propia parte de la matriz.

**Claves y relaciones.** Sin evidencia para `TbAuditoriaLog`.

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbAuditoriaLog | IDLog | LONG(4) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbAuditoriaLog | IDNC | LONG(4) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbAuditoriaLog | IDAC | LONG(4) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbAuditoriaLog | IDAR | LONG(4) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbAuditoriaLog | Usuario | TEXT(255) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbAuditoriaLog | Fecha | SHORT_DATE_TIME(8) | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbAuditoriaLog | Titulo | MEMO | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbAuditoriaLog | Linea | MEMO | sin destino propuesto | — | — | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |

## TbAuditorias

**Quién la lee y escribe.** Fichas relacionadas: [FormAuditoria](../formularios/Form_FormAuditoria.md) (alta y edición, `Form_FormAuditoria.cls:76-78`) y [FormAuditoriasGestion](../formularios/Form_FormAuditoriasGestion.md) (`Form_FormAuditoriasGestion.cls:227-235`, que instancia `AuditoriaOperaciones`).

| Operación | Dónde | Tipo de acceso |
|---|---|---|
| Alta | `AuditoriaOperaciones.cls:77-99` | DAO sobre la tabla (`AddNew`/`Update`) |
| Edición | `AuditoriaOperaciones.cls:83-84`, `AuditoriaOperaciones.cls:102-107` | SQL dinámico concatenado por `IDAuditoria` y DAO (`Edit`/`Update`) |
| Baja | `AuditoriaOperaciones.cls:173-175` | SQL dinámico concatenado (`DELETE`) con `Execute` |
| Lectura de una auditoría | `constructor.bas:4815-4819` | SQL dinámico concatenado por `IDAuditoria` |
| Lectura de todas | `constructor.bas:4861-4863` | SQL con `ORDER BY IDAuditoria DESC` |
| Cálculo del siguiente identificador | `Auditoria.cls:157`, `constructor.bas:1165-1178` | `Max(IDAuditoria)` más uno, sin secuencia |
| Etiqueta «año_tipo» | `constructor.bas:4150-4155`, `Variables Globales.bas:465-471` | Uniones de SQL dinámico concatenado con `TbNoConformidadesAuditoria` |

No hay transacción: el alta de la auditoría y el volcado de sus anexos son operaciones separadas (`Form_FormAuditoria.cls:78`, `Form_FormAuditoria.cls:83`), y la baja borra primero los documentos y después la fila sin transacción (`AuditoriaOperaciones.cls:160-175`). La baja se bloquea si la auditoría tiene NC asociadas (`AuditoriaOperaciones.cls:139-142`). Tras alta, edición o baja se invalidan las cachés de auditorías (`AuditoriaOperaciones.cls:113-114`, `AuditoriaOperaciones.cls:178-179`). La lista de columnas que se hidrata es la de `Auditoria.ColCampos` (`Auditoria.cls:213-216`).

**Claves y relaciones.**
- Clave primaria `IDAuditoria`: inferida por código (identificador único calculado con `Max + 1` en `constructor.bas:1165-1178` y consulta puntual por él en `constructor.bas:4815-4817`; diccionario sin repetidos en `constructor.bas:4883-4884`). No hay metadatos declarados.
- `TbNoConformidadesAuditoria.IDAuditoria` referencia a `TbAuditorias.IDAuditoria`: inferida por código (`constructor.bas:4154-4155`, `Variables Globales.bas:471`).
- `TbDocumentosAuditorias.IDAuditoria` referencia a `TbAuditorias.IDAuditoria`: inferida por código (`constructor.bas:4920-4923`).
- A verificar: `constructor.bas:4284`, `constructor.bas:4351`, `constructor.bas:4419` y `constructor.bas:4488` unen `TbAuditorias.IDAuditoria` con `TbNoConformidadesAuditoria.ID` y no con `IDAuditoria`; puede ser un defecto del legado que afecte a los resultados de esas consultas.

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbAuditorias | IDAuditoria | LONG(4) | no_conformidades.auditorias.id_auditoria | integer, clave primaria | Se conserva el valor para no romper las referencias; la secuencia destino debe arrancar por encima del máximo migrado | Obligatorio en la práctica (se asigna siempre en el alta, `AuditoriaOperaciones.cls:78`, `AuditoriaOperaciones.cls:95`); unicidad solo por código, no demostrada en el backend | Numérico; se genera como máximo más uno (`constructor.bas:1178`) | propuesta |
| NoConformidades_Datos.accdb | TbAuditorias | Tipo | TEXT(255) | no_conformidades.auditorias.tipo | varchar(20) con restricción de valores | Directa; valores admitidos «Interna» y «Externa» | Se exige informado en el alta y la edición (`AuditoriaOperaciones.cls:37-40`); al leer se convierte nulo en vacío (`constructor.bas:4828`, `constructor.bas:4874`); existencia de valores fuera del catálogo en datos históricos: a verificar | Debe ser «Interna» o «Externa» (`AuditoriaOperaciones.cls:37-40`) | propuesta |
| NoConformidades_Datos.accdb | TbAuditorias | FechaInicio | SHORT_DATE_TIME(8) | no_conformidades.auditorias.fecha_inicio | timestamp, o date si se confirma que no hay hora | Directa; el año alimenta la etiqueta «año_tipo» (`constructor.bas:4151`) | Obligatoria al guardar (`AuditoriaOperaciones.cls:25-28`); históricos nulos: a verificar | Debe ser fecha válida (`AuditoriaOperaciones.cls:25`) | propuesta |
| NoConformidades_Datos.accdb | TbAuditorias | FechaFin | SHORT_DATE_TIME(8) | no_conformidades.auditorias.fecha_fin | timestamp, o date si se confirma que no hay hora | Directa | Obligatoria al guardar (`AuditoriaOperaciones.cls:29-32`); históricos nulos: a verificar | Fecha válida y no anterior a la inicial (`AuditoriaOperaciones.cls:29-36`) | propuesta |

## TbAuxPuntoNorma

**Quién la lee y escribe.** Es un catálogo auxiliar de puntos de la norma que se autoalimenta. Se rellena desde el guardado de una NC de auditoría ([FormNCAuditoriaGeneral](../formularios/Form_FormNCAuditoriaGeneral.md)): `NCaUDITORIAOperaciones.cls:254-256` llama a `PuntoNormaAuditoriaRegistrar`, que busca el valor con SQL dinámico concatenado (`Funciones Generales.bas:3323-3326`) y, si no existe, lo inserta con DAO (`Funciones Generales.bas:3327-3331`); si se ha insertado, invalida la caché (`NCaUDITORIAOperaciones.cls:257-259`). Se lee con DAO sobre la tabla completa (`constructor.bas:5958-5959`) para cargar el diccionario `getColPuntosNormaAuditoria` (`constructor.bas:5967-5975`), que cachea `Entorno.cls:2285-2297` y que llena el desplegable (`Funciones Generales.bas:3116-3126`, invocado desde `Form_FormNCAuditoriaGeneral.cls:243`). No hay baja ni edición. No hay transacción.

**Claves y relaciones.**
- Clave primaria `PuntoNorma`: inferida por código (el valor se trata como único: búsqueda previa a insertar en `Funciones Generales.bas:3323-3327` y control `Exists` en `constructor.bas:5973`).
- Relación con las NC: el texto elegido se guarda como valor libre `PuntoNorma` de la NC de auditoría (`NCaUDITORIAOperaciones.cls:167`, `NCaUDITORIAOperaciones.cls:206`); no hay identificador de enlace ni integridad referencial (sin evidencia de clave foránea).

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbAuxPuntoNorma | PuntoNorma | TEXT(255) | no_conformidades.aux_punto_norma.punto_norma | varchar(255), clave primaria o única | Directa; recortar espacios es opcional (el desplegable ya omite vacíos, `Funciones Generales.bas:3124`) | Un valor vacío no se inserta (`Funciones Generales.bas:3320-3322`); duplicados evitados por código en alta (`Funciones Generales.bas:3327`) y en carga (`constructor.bas:5973`); la comparación del diccionario no distingue mayúsculas (`constructor.bas:5971`), por lo que variantes de caja pueden colapsar (a verificar) | El valor se concatena sin escapar en SQL (`Funciones Generales.bas:3325`): un apóstrofo rompería la consulta (a verificar); en destino usar parámetros | propuesta |

## Pendientes

- Decisión del propietario sobre `Copia de TbNCARAvisos`, `TbAnexos` y `TbAuditoriaLog` (sin uso en `main`): descartar, archivar o migrar.
- Confirmar si `InstalarProyecto` se ejecutó alguna vez en producción y, por tanto, si `TbAnexosAuditoria` y `TbAnexosNCAuditorias` tienen contenido pendiente de volcar a `TbDocumentosAuditorias`.
- Confirmar con el ERD de origen real (extracción pendiente del propietario) claves primarias, nulabilidad y relaciones, y la hipótesis de que `TbAuditoriaLog` duplica `TbLogAuditoria`.
- Revisar las uniones por `TbNoConformidadesAuditoria.ID` en `constructor.bas:4284`, `constructor.bas:4351`, `constructor.bas:4419` y `constructor.bas:4488`.
- Confirmar si `FechaInicio` y `FechaFin` de `TbAuditorias` llevan hora.
