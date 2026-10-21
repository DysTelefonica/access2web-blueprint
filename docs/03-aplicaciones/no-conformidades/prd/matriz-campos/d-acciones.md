# Matriz campo a campo: acciones, documentos y avisos (No Conformidades)

Parte «d-acciones» de la matriz campo a campo del PRD de No Conformidades (#1080). Cubre las tablas del esquema previo `TbNCAccionCorrectivas`, `TbNCAccionesRealizadas`, `TbNCARAvisos`, `TbNCAuditoriaAccionCorrectivas`, `TbNCAuditoriaAccionesRealizadas`, `TbNCDocumentos`, `TbNCInformacionRAC` y `TbNCInformacionRACAnexos`. Índice de formularios y capacidades en [el PRD de la aplicación](../README.md).

## Método

La línea base es el código de `main` (producción), commit `1d7a2f826a1e3c582874e4e29886a6183efb9067`, directorios `src/classes`, `src/modules`, `src/forms` y `src/queries`. Los tipos de origen proceden del DDL de `ERD/schema.sql` (volcado de 2026-08-06), que solo declara tipos: no declara claves, relaciones ni nulabilidad, y no hay backend legible en esta fase. Por eso toda clave o relación se marca «inferida por código» con su cita, o «sin evidencia». Todo destino es una propuesta sujeta a la aprobación del ERD global; los tipos de destino son orientativos para PostgreSQL. Las citas `fichero:línea` apuntan a la línea del `FROM`, `JOIN`, `OpenRecordset` o asignación de campo que respalda la afirmación; las lecturas de `constructor.bas` se citan por muestra representativa, no de forma exhaustiva. Lo que el código sugiere pero no demuestra se marca «a verificar». No se incluyen datos de negocio del backend.

Clasificación de acceso usada en los apartados: **DAO** (`OpenRecordset` de tabla o `SELECT *` seguido de `AddNew`/`Edit`/`Update`), **SQL guardada** (consulta de `src/queries`) y **SQL dinámico concatenado** (cadena VBA construida con variables y ejecutada con `Execute` u `OpenRecordset`). Las cadenas VBA fijas sin variables se indican como «SQL embebido fijo».

Hallazgos transversales a las cinco tablas con uso:

- Ninguna operación de escritura usa transacciones: no hay `BeginTrans`, `CommitTrans` ni `Rollback` en `src` salvo `CommitTransaccionCache` (`CacheNCProyecto.bas:646`), que no envuelve escrituras de estas tablas. Cada alta, edición o borrado va seguido de recálculos en escrituras separadas (p. ej. `ACProyectoOperaciones.cls:241`, `ARProyectoOperaciones.cls:304-310`), de modo que un fallo intermedio deja datos derivados desactualizados (a verificar).
- Las claves se generan con «máximo más uno» (`constructor.bas:1165-1166`), sin bloqueo visible: riesgo de colisión con dos usuarios simultáneos (a verificar). En destino debe sustituirse por identidad o secuencia.
- Los textos de acción se sanean al registrar: se eliminan comillas simples y `;` pasa a `:` (`ACProyectoOperaciones.cls:133-134`, `ARProyectoOperaciones.cls:146-147`, `ACAuditoriaOperaciones.cls:128-129`, `ARAuditoriaOperaciones.cls:138-139`). Es una mitigación de inyección en SQL concatenado; en destino se usarán consultas parametrizadas y el texto debería conservarse sin alterar (los datos históricos ya saneados no se pueden recuperar).
- Las columnas `ESTADO`, `FechaInicialMinima`, `FechaFinalUltima` y `FechaFinPrevistaUltima` de las acciones correctivas, y `ESTADO` de las acciones realizadas, son datos derivados que el código recalcula y regraba; en destino son candidatas a vista o columna generada en lugar de dato almacenado (decisión del propietario).
- El esquema previo no incluye `IDNoConformidadCalculada`, que `Instalador.bas:301` y `Instalador.bas:329` usan sobre `TbNCDocumentos`: el volcado está desactualizado respecto a `main` o la columna ya no existe en el backend (a verificar con el ERD de origen).

## TbNCAccionCorrectivas

### Uso en `main`

- **Alta**: DAO. `m_SQL = "TbNCAccionCorrectivas"` y `OpenRecordset` + `AddNew` en `ACProyectoOperaciones.cls:137`, `ACProyectoOperaciones.cls:157`, campos en `ACProyectoOperaciones.cls:168-201`. El identificador se calcula con `getID` (`ACProyecto.cls:678`).
- **Edición**: DAO (`SELECT *` por `IDAccionCorrectiva`, `Edit`/`Update`) en `ACProyectoOperaciones.cls:211-228`; `NAccion` en `ACProyectoOperaciones.cls:284-289` y `Responsable` en `ACProyectoOperaciones.cls:473-478`.
- **Baja**: SQL dinámico concatenado, `DELETE` por `IDAccionCorrectiva` en `ACProyectoOperaciones.cls:331-334`, tras eliminar antes las acciones realizadas dependientes (cascada hecha por el código, no por el esquema).
- **Recálculo de derivados**: SQL dinámico concatenado (`UPDATE` por `IDAccionCorrectiva`) en `ACProyecto.cls:242-249` (`ESTADO`), `ACProyecto.cls:558-565`, `ACProyecto.cls:593-600`, `ACProyecto.cls:627-634`.
- **Lecturas** (SQL embebido y dinámico, uniones con la NC y con las acciones realizadas): `constructor.bas:460-462`, `constructor.bas:795`, `constructor.bas:858`, `constructor.bas:1892-1897`, `constructor.bas:3162-3169`, `constructor.bas:3225-3229`, `constructor.bas:3755-3756`, `constructor.bas:3885-3888`, `constructor.bas:5381-5385`; `CacheNCProyecto.bas:807-809`; `Variables Globales.bas:445-456`.
- **Mantenimiento puntual**: `Instalador.bas:373` (`NAccion`), `Instalador.bas:443-451` (sustitución de identificadores de `Responsable`), `Instalador.bas:931` y `Instalador.bas:1038` (listados de acciones sin estado); `Consulta2.sql:1-3` es una consulta guardada ad hoc con identificador literal, sin uso funcional.
- **Transacción**: ninguna (véase hallazgos transversales). Alta y recálculos son escrituras independientes (`ACProyectoOperaciones.cls:241`).
- **Formularios**: [Form_FormNCProyectoAC](../formularios/Form_FormNCProyectoAC.md), [Form_FormNCProyectoAcciones](../formularios/Form_FormNCProyectoAcciones.md).

### Claves y relaciones

- Clave primaria: `IDAccionCorrectiva`, inferida por código (`ACProyectoOperaciones.cls:154-155` y `ACProyectoOperaciones.cls:331-334` acotan siempre por ese campo).
- `IDNoConformidad` → `TbNoConformidades.IDNoConformidad`: inferida por código (`constructor.bas:460-462`). Sin integridad referencial declarada: sin evidencia.
- `IDAccionCorrectiva` ← `TbNCAccionesRealizadas.IDAccionCorrectiva`: inferida por código (`constructor.bas:2135-2139`).
- Unicidad de `(NC, AccionCorrectiva)`: la impone solo el código (`ACProyectoOperaciones.cls:95`), no el esquema.

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbNCAccionCorrectivas | IDAccionCorrectiva | LONG(4) | no_conformidades.accion_correctiva.id | bigint (PK, identidad) | Se descarta el «máximo más uno»; se conservan los valores históricos y se reinicia la secuencia por encima del máximo. | Siempre informado en el alta (`ACProyectoOperaciones.cls:138`). Duplicados: sin restricción declarada; posible colisión por concurrencia (a verificar). | Sin validación propia; clave de todas las operaciones (`ACProyectoOperaciones.cls:211`). | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionCorrectivas | IDNoConformidad | LONG(4) | no_conformidades.accion_correctiva.no_conformidad_id | bigint (FK) | Copia directa; FK a `no_conformidad`. | Obligatorio por código: «No se conoce la NC/Obs» (`ACProyectoOperaciones.cls:66`). Se reasigna en edición (`ACProyectoOperaciones.cls:212`). | Debe existir la NC; relación inferida, sin FK en origen. | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionCorrectivas | NAccion | LONG(4) | no_conformidades.accion_correctiva.n_accion | integer | Copia directa. Se calcula como máximo+1 por NC (`Funciones Generales.bas:3821-3823`) y solo se graba si es numérico (`ACProyectoOperaciones.cls:174-176`). | Puede quedar nulo en registros antiguos (hay rutinas de relleno: `Instalador.bas:373`, `Instalador.bas:1501-1560`). Duplicados dentro de la misma NC: sin restricción (a verificar). | Numérico; correlativo por NC. En destino, restricción única `(nc, n_accion)` a validar con el propietario. | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionCorrectivas | AccionCorrectiva | MEMO | no_conformidades.accion_correctiva.accion_correctiva | text | Copia directa; el saneamiento de `'` y `;` (`ACProyectoOperaciones.cls:133-134`) no se replica. | Obligatorio por código (`ACProyectoOperaciones.cls:71`). Duplicado por NC prohibido por código: «Existe otra acción correctiva con el mismo nombre…» (`ACProyectoOperaciones.cls:95`). | No vacío; única por NC (regla de aplicación, comparación exacta de cadena: a verificar si distingue mayúsculas). | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionCorrectivas | FechaAccionCorrectiva | SHORT_DATE_TIME(8) | no_conformidades.accion_correctiva.fecha_accion_correctiva | timestamp without time zone | Copia directa; en origen la fecha de alta es `Now()` y no es editable. | Se fija con `Now()` en el alta (`ACProyectoOperaciones.cls:172, 187`); no se modifica en edición (`ACProyectoOperaciones.cls:211-228`). | Sin validación; es marca de creación (candidata a `created_at`; a decidir con el propietario). | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionCorrectivas | ESTADO | TEXT(255) | no_conformidades.accion_correctiva.estado | varchar(32) con CHECK candidato | Valor derivado: `ACTIVA`, `SINACCIONES`, `FINALIZADA`, `PTEREPLANIFICAR`, `PTEREREGULARIZAR`, `REGISTRADA` (`Entorno.cls:263-268`). Propuesta: vista calculada o columna regenerada por disparador. | Nulo permitido: se graba `Null` si el estado calculado está vacío (`ACProyecto.cls:242-249`). Existen rutinas para ACs sin estado (`Instalador.bas:931`). | Catálogo cerrado en código; el texto largo de pantalla está en `Entorno.cls:321-326` (no se almacena). | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionCorrectivas | FechaInicialMinima | SHORT_DATE_TIME(8) | no_conformidades.accion_correctiva.fecha_inicial_minima | date | Derivado: mínimo de `FechaInicio` de las acciones realizadas (`Funciones Generales.bas:3645-3647`). Se graba con formato `mm/dd/yyyy` en SQL concatenado (`ACProyecto.cls:558`). | Nulo si no hay acciones realizadas con fecha (`UPDATE … Null`, mismo rango). Sin duplicados aplicables. | Derivado; en destino preferible calcularlo en consulta. | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionCorrectivas | FechaFinalUltima | SHORT_DATE_TIME(8) | no_conformidades.accion_correctiva.fecha_final_ultima | date | Derivado: máximo de `FechaFinReal` de las acciones realizadas (`Funciones Generales.bas:3680-3682`). | Nulo si ninguna acción realizada está cerrada (`ACProyecto.cls:593-600`). | Derivado. | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionCorrectivas | Notas | MEMO | no_conformidades.accion_correctiva.notas | text | Copia directa. | Solo se graba si no está vacía en el alta; en edición se pone `Null` si queda vacía (`ACProyectoOperaciones.cls:215-217`). | Sin validación. | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionCorrectivas | Responsable | TEXT(255) | no_conformidades.accion_correctiva.responsable | varchar(255) | Copia directa del identificador de usuario de red; se propone resolver a la entidad de usuarios del ERD global. | Opcional; vacío se graba `Null` en edición (`ACProyectoOperaciones.cls:220-222`). `RegistrarResponsable` lo copia de la NC (`ACProyectoOperaciones.cls:449-478`). | Si se informa, debe existir en la lista de usuarios (`ACProyectoOperaciones.cls:84-86`); relación a usuarios inferida, sin FK en origen. | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionCorrectivas | FechaFinPrevistaUltima | SHORT_DATE_TIME(8) | no_conformidades.accion_correctiva.fecha_fin_prevista_ultima | date | Derivado: máximo de `FechaFinPrevista` de las acciones realizadas (`Funciones Generales.bas:3715-3717`). | Nulo si ninguna acción tiene fecha prevista (`ACProyecto.cls:627-634`). | Derivado. | propuesta |

## TbNCAccionesRealizadas

### Uso en `main`

- **Alta**: DAO (`OpenRecordset` de tabla + `AddNew`), campos en `ARProyectoOperaciones.cls:179-220`; identificador con `getID` (`ARProyecto.cls:312`).
- **Edición**: DAO (`SELECT *` por `IDAccionRealizada`, `Edit`/`Update`) en `ARProyectoOperaciones.cls:232-265`; replanificación (`FechaFinPrevista`) en `ARProyectoOperaciones.cls:499-506`.
- **Baja**: SQL dinámico concatenado, `DELETE` por `IDAccionRealizada` en `ARProyectoOperaciones.cls:351-354`. No borra sus documentos: el borrado de documentos es una operación aparte (`DocumentoProyectoOperaciones.cls:46-49`).
- **Estado**: SQL dinámico concatenado (`UPDATE … ESTADO`) en `ARProyecto.cls:206-213`.
- **Lecturas** (SQL embebido y dinámico, uniones con la acción correctiva y la NC): `constructor.bas:919`, `constructor.bas:969-972`, `constructor.bas:1019`, `constructor.bas:1097`, `constructor.bas:1892-1897`, `constructor.bas:2135-2139`, `constructor.bas:3225-3229`, `constructor.bas:3885-3888`; `CacheNCProyecto.bas:863-867`; `Variables Globales.bas:445-456`.
- **Transacción**: ninguna. Tras el alta se recalculan la acción correctiva y la NC en escrituras separadas (`ARProyectoOperaciones.cls:304-310`).
- **Formularios**: [Form_FormNCProyectoAR](../formularios/Form_FormNCProyectoAR.md), [Form_FormNCProyectoSeguimientoTareas](../formularios/Form_FormNCProyectoSeguimientoTareas.md).

### Claves y relaciones

- Clave primaria: `IDAccionRealizada`, inferida por código (`ARProyectoOperaciones.cls:167-168` y `ARProyectoOperaciones.cls:351-354`).
- `IDAccionCorrectiva` → `TbNCAccionCorrectivas.IDAccionCorrectiva`: inferida por código (`constructor.bas:2135-2139`). Sin FK declarada: sin evidencia.
- Documentos asociados: `TbNCDocumentos.IDAccionRealizada`: inferida por código (`constructor.bas:737-741`). 

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbNCAccionesRealizadas | IDAccionRealizada | LONG(4) | no_conformidades.accion_realizada.id | bigint (PK, identidad) | Se descarta el «máximo más uno»; se conservan los valores históricos. | Siempre informado en el alta (`ARProyectoOperaciones.cls:180`). Posible colisión por concurrencia (a verificar). | Clave de todas las operaciones. | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionesRealizadas | IDAccionCorrectiva | LONG(4) | no_conformidades.accion_realizada.accion_correctiva_id | bigint (FK) | Copia directa; FK a la acción correctiva. | Obligatorio en el alta (`ARProyectoOperaciones.cls:181`); no se modifica en edición (`ARProyectoOperaciones.cls:232-265`). | Debe existir la acción correctiva; relación inferida. | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionesRealizadas | NAccion | LONG(4) | no_conformidades.accion_realizada.n_accion | integer | Copia directa; se calcula con `NAccionARCalculado` (`ACProyecto.cls:80-112`). | Siempre informado en el alta (`ARProyectoOperaciones.cls:182`). Rutinas de relleno para registros sin número: `Instalador.bas:1676`, `Instalador.bas:2320`. | Correlativo por acción correctiva; restricción única `(accion_correctiva, n_accion)` a validar. | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionesRealizadas | AccionRealizada | MEMO | no_conformidades.accion_realizada.accion_realizada | text | Copia directa; el saneamiento de `'` y `;` (`ARProyectoOperaciones.cls:146-147`) no se replica. | Obligatorio por código (se graba siempre). Duplicado por acción correctiva prohibido (`ARProyectoOperaciones.cls:94`). | No vacío; única por acción correctiva (regla de aplicación). | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionesRealizadas | FechaAccionRealizada | SHORT_DATE_TIME(8) | no_conformidades.accion_realizada.fecha_accion_realizada | timestamp without time zone | Copia directa. Semántica compuesta: se fija con `Now()` en el alta y se sobrescribe con `FechaFinReal` al cerrar (`ARProyectoOperaciones.cls:197-199`). | Se pone `Null` en edición cuando se reabre (`ARProyectoOperaciones.cls:250`). | Dato redundante con `FechaFinReal` tras el cierre; a verificar con el propietario si se conserva. | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionesRealizadas | FechaInicio | SHORT_DATE_TIME(8) | no_conformidades.accion_realizada.fecha_inicio | date | Copia directa. | Se graba `Null` si no es fecha válida (`ARProyectoOperaciones.cls:187-189`). | `IsDate`; `BuenaFecha` (`ARProyectoOperaciones.cls:62-65`) y coherencia fin/inicio (`ARProyectoOperaciones.cls:56`). | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionesRealizadas | FechaFinPrevista | SHORT_DATE_TIME(8) | no_conformidades.accion_realizada.fecha_fin_prevista | date | Copia directa; cada cambio genera un registro de replanificación aparte (`ARProyectoOperaciones.cls:270-297`). | Se graba `Null` si no es fecha válida (`ARProyectoOperaciones.cls:192-194`). Replanificación: `ARProyectoOperaciones.cls:499-506`. | `IsDate`; la replanificación no puede ser anterior a hoy en proyecto (`ACProyectoOperaciones.cls:408-410`). | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionesRealizadas | FechaFinReal | SHORT_DATE_TIME(8) | no_conformidades.accion_realizada.fecha_fin_real | date | Copia directa; su presencia define la acción como cerrada. | `Null` mientras esté abierta (`ARProyectoOperaciones.cls:201`). Cierre exige evidencia (documento). | Para cerrar hay que adjuntar una evidencia (`ARProyectoOperaciones.cls:71`, `ARProyectoOperaciones.cls:82`). | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionesRealizadas | ESTADO | TEXT(255) | no_conformidades.accion_realizada.estado | varchar(32) con CHECK candidato | Derivado: `ACTIVA`, `FINALIZADA`, `PTEREPLANIFICAR`, `IRREGULAR`, `REGISTRADA` (`Entorno.cls:180-184`). Propuesta: vista calculada. | Se graba `Null` si el estado calculado está vacío (`ARProyecto.cls:206-213`); en alta y edición se recalcula (`ARProyectoOperaciones.cls:213-216, 262-263`). | Catálogo cerrado en código. | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionesRealizadas | Notas | MEMO | no_conformidades.accion_realizada.notas | text | Copia directa. | En el alta solo si no está vacía; en edición `Null` si queda vacía (`ARProyectoOperaciones.cls:253-255`). | Sin validación. | propuesta |
| NoConformidades_Datos.accdb | TbNCAccionesRealizadas | Responsable | TEXT(255) | no_conformidades.accion_realizada.responsable | varchar(255) | Copia directa del identificador de usuario de red; resolver a usuarios del ERD global. | Si está vacío hereda el de la acción correctiva (`ARProyectoOperaciones.cls:148-150`); en edición vacío ⇒ `Null` (`ARProyectoOperaciones.cls:258-260`). | Relación a usuarios inferida, sin FK en origen. | propuesta |

## TbNCARAvisos

### Uso en `main`

Sin uso en `main`. Búsqueda: `grep -rnwI` del nombre `TbNCARAvisos` y del nombre del DDL `Copia de TbNCARAvisos` sobre `src/classes`, `src/modules`, `src/forms` y `src/queries` del commit fijado, sin resultados; y búsqueda sin distinguir mayúsculas de `avisos` sobre `src`, cuyo único resultado relacionado es el indicador de permiso `EsUsuarioCalidadAvisos` (`UsuarioAplicacionPermisos.cls:21`), que no accede a la tabla (a verificar). Obsérvese que el DDL de `schema.sql` declara la tabla como `Copia de TbNCARAvisos` (línea 4), lo que apunta a una copia de seguridad o tabla abandonada.

### Claves y relaciones

Sin evidencia. Por nombre, `IDAR` podría referirse a una acción realizada y `IDCorreo15`, `IDCorreo7` e `IDCorreo0` a avisos de correo escalonados; es una hipótesis a verificar, sin cita de código.

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbNCARAvisos (DDL: «Copia de TbNCARAvisos») | ID | LONG(4) | no_conformidades.ar_aviso.id | bigint (PK candidata) | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNCARAvisos (DDL: «Copia de TbNCARAvisos») | IDAR | LONG(4) | no_conformidades.ar_aviso.accion_realizada_id | bigint (FK candidata) | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNCARAvisos (DDL: «Copia de TbNCARAvisos») | IDCorreo15 | LONG(4) | no_conformidades.ar_aviso.correo_15_id | bigint | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNCARAvisos (DDL: «Copia de TbNCARAvisos») | IDCorreo7 | LONG(4) | no_conformidades.ar_aviso.correo_7_id | bigint | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNCARAvisos (DDL: «Copia de TbNCARAvisos») | IDCorreo0 | LONG(4) | no_conformidades.ar_aviso.correo_0_id | bigint | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNCARAvisos (DDL: «Copia de TbNCARAvisos») | Fecha | SHORT_DATE_TIME(8) | no_conformidades.ar_aviso.fecha | timestamp without time zone | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |

## TbNCAuditoriaAccionCorrectivas

### Uso en `main`

- **Alta**: DAO. `m_SQL = "TbNCAuditoriaAccionCorrectivas"` y `OpenRecordset` + `AddNew` en `ACAuditoriaOperaciones.cls:132`, `ACAuditoriaOperaciones.cls:152`, campos en `ACAuditoriaOperaciones.cls:163-190`. El identificador se calcula con `getID` (`ACAuditoria.cls:635`).
- **Edición**: DAO (`SELECT *` por `IDAccionCorrectiva`, `Edit`/`Update`) en `ACAuditoriaOperaciones.cls:196-215`; `NAccion` en `ACAuditoriaOperaciones.cls:266-271` y `Responsable` en `ACAuditoriaOperaciones.cls:452-457`.
- **Baja**: SQL dinámico concatenado, `DELETE` por `IDAccionCorrectiva` en `ACAuditoriaOperaciones.cls:314-317`, tras eliminar antes las acciones realizadas dependientes (cascada hecha por el código, no por el esquema).
- **Recálculo de derivados**: SQL dinámico concatenado (`UPDATE` por `IDAccionCorrectiva`) en `ACAuditoria.cls:212-219` (`ESTADO`), `ACAuditoria.cls:454-461`, `ACAuditoria.cls:519-526`, `ACAuditoria.cls:584-591`.
- **Lecturas** (SQL embebido y dinámico, uniones con la NC y con las acciones realizadas): `constructor.bas:1955-1960`, `constructor.bas:4156-4158`, `constructor.bas:4222`, `constructor.bas:4285-4290`, `constructor.bas:5023-5027`, `constructor.bas:5082`, `constructor.bas:5147`, `constructor.bas:5321-5325`, `constructor.bas:5504-5506`, `constructor.bas:5841-5848`; `Variables Globales.bas:459-470`.
- **Transacción**: ninguna (véase hallazgos transversales). Alta y recálculos son escrituras independientes (`ACAuditoriaOperaciones.cls:224`).
- **Formularios**: [Form_FormNCAuditoriaAC](../formularios/Form_FormNCAuditoriaAC.md), [Form_FormNCAuditoriaAcciones](../formularios/Form_FormNCAuditoriaAcciones.md).

### Claves y relaciones

- Clave primaria: `IDAccionCorrectiva`, inferida por código (`ACAuditoriaOperaciones.cls:149-150` y `ACAuditoriaOperaciones.cls:314-317` acotan siempre por ese campo).
- `ID` → `TbNoConformidadesAuditoria.ID`: inferida por código (`constructor.bas:5504-5505`). Sin integridad referencial declarada: sin evidencia.
- `IDAccionCorrectiva` ← `TbNCAuditoriaAccionesRealizadas.IDAccionCorrectiva`: inferida por código (`constructor.bas:5206-5210`).
- Unicidad de `(NC, AccionCorrectiva)`: la impone solo el código (`ACAuditoriaOperaciones.cls:93`), no el esquema.

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionCorrectivas | IDAccionCorrectiva | LONG(4) | no_conformidades.auditoria_accion_correctiva.id | bigint (PK, identidad) | Se descarta el «máximo más uno»; se conservan los valores históricos y se reinicia la secuencia por encima del máximo. | Siempre informado en el alta (`ACAuditoriaOperaciones.cls:133`). Duplicados: sin restricción declarada; posible colisión por concurrencia (a verificar). | Sin validación propia; clave de todas las operaciones (`ACAuditoriaOperaciones.cls:196`). | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionCorrectivas | ID | LONG(4) | no_conformidades.auditoria_accion_correctiva.no_conformidad_auditoria_id | bigint (FK) | Copia directa; FK a `no_conformidad_auditoria`. | Obligatorio por código (`ACAuditoriaOperaciones.cls:66`). No se reasigna en edición (`ACAuditoriaOperaciones.cls:196-215`). El nombre de columna es `ID`, no `IDNoConformidad`. | Debe existir la NC; relación inferida, sin FK en origen. | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionCorrectivas | NAccion | LONG(4) | no_conformidades.auditoria_accion_correctiva.n_accion | integer | Copia directa. Se calcula como máximo+1 por NC (`Funciones Generales.bas:3825-3827`) y solo se graba si es numérico (`ACAuditoriaOperaciones.cls:167`). | Puede quedar nulo en registros antiguos (hay rutinas de relleno: `Instalador.bas:373`, `Instalador.bas:1501-1560`). Duplicados dentro de la misma NC: sin restricción (a verificar). | Numérico; correlativo por NC. En destino, restricción única `(nc, n_accion)` a validar con el propietario. | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionCorrectivas | AccionCorrectiva | MEMO | no_conformidades.auditoria_accion_correctiva.accion_correctiva | text | Copia directa; el saneamiento de `'` y `;` (`ACAuditoriaOperaciones.cls:128-129`) no se replica. | Obligatorio por código (`ACAuditoriaOperaciones.cls:71`). Duplicado por NC prohibido por código: «Existe otra acción correctiva con el mismo nombre…» (`ACAuditoriaOperaciones.cls:93`). | No vacío; única por NC (regla de aplicación, comparación exacta de cadena: a verificar si distingue mayúsculas). | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionCorrectivas | FechaAccionCorrectiva | SHORT_DATE_TIME(8) | no_conformidades.auditoria_accion_correctiva.fecha_accion_correctiva | timestamp without time zone | Copia directa; en origen la fecha de alta es `Now()` y no es editable. | En el alta de proyecto se fija con `Now()` (`ACAuditoriaOperaciones.cls:175`); en auditoría solo se graba si viene informada (`ACAuditoriaOperaciones.cls:175`). Nulo posible en auditoría. | Sin validación; es marca de creación (candidata a `created_at`; a decidir con el propietario). | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionCorrectivas | ESTADO | TEXT(255) | no_conformidades.auditoria_accion_correctiva.estado | varchar(32) con CHECK candidato | Valor derivado: `ACTIVA`, `SINACCIONES`, `FINALIZADA`, `PTEREPLANIFICAR`, `PTEREREGULARIZAR`, `REGISTRADA` (`Entorno.cls:263-268`). Propuesta: vista calculada o columna regenerada por disparador. | Nulo permitido: se graba `Null` si el estado calculado está vacío (`ACAuditoria.cls:212-219`). Existen rutinas para ACs sin estado (`Instalador.bas:931`). | Catálogo cerrado en código; el texto largo de pantalla está en `Entorno.cls:321-326` (no se almacena). | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionCorrectivas | FechaInicialMinima | SHORT_DATE_TIME(8) | no_conformidades.auditoria_accion_correctiva.fecha_inicial_minima | date | Derivado: mínimo de `FechaInicio` de las acciones realizadas (`Funciones Generales.bas:3645-3647`). Se graba con formato `mm/dd/yyyy` en SQL concatenado (`ACAuditoria.cls:454`). | Nulo si no hay acciones realizadas con fecha (`UPDATE … Null`, mismo rango). Sin duplicados aplicables. | Derivado; en destino preferible calcularlo en consulta. | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionCorrectivas | FechaFinalUltima | SHORT_DATE_TIME(8) | no_conformidades.auditoria_accion_correctiva.fecha_final_ultima | date | Derivado: máximo de `FechaFinReal` de las acciones realizadas (`Funciones Generales.bas:3680-3682`). | Nulo si ninguna acción realizada está cerrada (`ACAuditoria.cls:519-526`). | Derivado. | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionCorrectivas | Notas | MEMO | no_conformidades.auditoria_accion_correctiva.notas | text | Copia directa. | Solo se graba si no está vacía en el alta; en edición se pone `Null` si queda vacía (`ACAuditoriaOperaciones.cls:200-202`). | Sin validación. | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionCorrectivas | Responsable | TEXT(255) | no_conformidades.auditoria_accion_correctiva.responsable | varchar(255) | Copia directa del identificador de usuario de red; se propone resolver a la entidad de usuarios del ERD global. | Opcional; vacío se graba `Null` en edición (`ACAuditoriaOperaciones.cls:205-207`).  | Sin comprobación de existencia del responsable en `MotivoNoOK` (`ACAuditoriaOperaciones.cls:56-107`, a verificar); relación a usuarios inferida, sin FK en origen. | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionCorrectivas | FechaFinPrevistaUltima | SHORT_DATE_TIME(8) | no_conformidades.auditoria_accion_correctiva.fecha_fin_prevista_ultima | date | Derivado: máximo de `FechaFinPrevista` de las acciones realizadas (`Funciones Generales.bas:3715-3717`). | Nulo si ninguna acción tiene fecha prevista (`ACAuditoria.cls:584-591`). | Derivado. | propuesta |

## TbNCAuditoriaAccionesRealizadas

### Uso en `main`

- **Alta**: DAO (`OpenRecordset` de tabla + `AddNew`), campos en `ARAuditoriaOperaciones.cls:171-212`; identificador con `getID` (`ARAuditoria.cls:308`).
- **Edición**: DAO (`SELECT *` por `IDAccionRealizada`, `Edit`/`Update`) en `ARAuditoriaOperaciones.cls:220-253`; replanificación (`FechaFinPrevista`) en `ARAuditoriaOperaciones.cls:480-487`.
- **`NAccion` y `Responsable`**: DAO en `ARAuditoriaOperaciones.cls:545-555` y `ARAuditoriaOperaciones.cls:588-598`.
- **Baja**: SQL dinámico concatenado, `DELETE` por `IDAccionRealizada` en `ARAuditoriaOperaciones.cls:334-337`.
- **Estado**: SQL dinámico concatenado (`UPDATE … ESTADO`) en `ARAuditoria.cls:205-212`.
- **Lecturas** (SQL embebido y dinámico, uniones con la acción correctiva y la NC): `constructor.bas:3479`, `constructor.bas:3596`, `constructor.bas:3659`, `constructor.bas:4287`, `constructor.bas:4553`, `constructor.bas:4610`, `constructor.bas:5206-5210`, `constructor.bas:5571`, `constructor.bas:5635`; `Variables Globales.bas:459-470`.
- **Transacción**: ninguna. Tras el alta se recalculan la acción correctiva y la NC en escrituras separadas (`ARAuditoriaOperaciones.cls:288-294`).
- **Formularios**: [Form_FormNCAuditoriaAR](../formularios/Form_FormNCAuditoriaAR.md), [Form_FormNCAuditoriaSeguimientoTareas](../formularios/Form_FormNCAuditoriaSeguimientoTareas.md).

### Claves y relaciones

- Clave primaria: `IDAccionRealizada`, inferida por código (`ARAuditoriaOperaciones.cls:159-160` y `ARAuditoriaOperaciones.cls:334-337`).
- `IDAccionCorrectiva` → `TbNCAuditoriaAccionCorrectivas.IDAccionCorrectiva`: inferida por código (`constructor.bas:5206-5210`). Sin FK declarada: sin evidencia.
- Documentos asociados: `TbDocumentosAuditorias.IDAccionRealizada`: inferida por código (`constructor.bas:5023-5027`). `TbDocumentosAuditorias` queda fuera de este fichero.

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionesRealizadas | IDAccionRealizada | LONG(4) | no_conformidades.auditoria_accion_realizada.id | bigint (PK, identidad) | Se descarta el «máximo más uno»; se conservan los valores históricos. | Siempre informado en el alta (`ARAuditoriaOperaciones.cls:172`). Posible colisión por concurrencia (a verificar). | Clave de todas las operaciones. | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionesRealizadas | IDAccionCorrectiva | LONG(4) | no_conformidades.auditoria_accion_realizada.auditoria_accion_correctiva_id | bigint (FK) | Copia directa; FK a la acción correctiva. | Obligatorio en el alta (`ARAuditoriaOperaciones.cls:173`); no se modifica en edición (`ARAuditoriaOperaciones.cls:220-253`). | Debe existir la acción correctiva; relación inferida. | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionesRealizadas | NAccion | LONG(4) | no_conformidades.auditoria_accion_realizada.n_accion | integer | Copia directa; se completa en `ARAuditoriaOperaciones.cls:545-555`. | Siempre informado en el alta (`ARAuditoriaOperaciones.cls:174`). Rutinas de relleno para registros sin número: `Instalador.bas:1676`, `Instalador.bas:2320`. | Correlativo por acción correctiva; restricción única `(accion_correctiva, n_accion)` a validar. | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionesRealizadas | AccionRealizada | MEMO | no_conformidades.auditoria_accion_realizada.accion_realizada | text | Copia directa; el saneamiento de `'` y `;` (`ARAuditoriaOperaciones.cls:138-139`) no se replica. | Obligatorio por código (se graba siempre). La comprobación de repetidas existe en `AccionRepetida` (a verificar el alcance en auditoría). | No vacío; única por acción correctiva (regla de aplicación). | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionesRealizadas | FechaAccionRealizada | SHORT_DATE_TIME(8) | no_conformidades.auditoria_accion_realizada.fecha_accion_realizada | timestamp without time zone | Copia directa. Semántica compuesta: se fija con `Now()` en el alta y se sobrescribe con `FechaFinReal` al cerrar (`ARAuditoriaOperaciones.cls:189-191`). | Se pone `Null` en edición cuando se reabre (`ARAuditoriaOperaciones.cls:238`). | Dato redundante con `FechaFinReal` tras el cierre; a verificar con el propietario si se conserva. | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionesRealizadas | FechaInicio | SHORT_DATE_TIME(8) | no_conformidades.auditoria_accion_realizada.fecha_inicio | date | Copia directa. | Se graba `Null` si no es fecha válida (`ARAuditoriaOperaciones.cls:179-181`). | `IsDate`; sin comprobación de fechas en `MotivoNoOK` (`ARAuditoriaOperaciones.cls:16-60`, a verificar). | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionesRealizadas | FechaFinPrevista | SHORT_DATE_TIME(8) | no_conformidades.auditoria_accion_realizada.fecha_fin_prevista | date | Copia directa. | Se graba `Null` si no es fecha válida (`ARAuditoriaOperaciones.cls:184-186`). Replanificación: `ARAuditoriaOperaciones.cls:480-487`. | `IsDate`; la replanificación no puede ser anterior a hoy en proyecto (`ACProyectoOperaciones.cls:408-410`). | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionesRealizadas | FechaFinReal | SHORT_DATE_TIME(8) | no_conformidades.auditoria_accion_realizada.fecha_fin_real | date | Copia directa; su presencia define la acción como cerrada. | `Null` mientras esté abierta (`ARAuditoriaOperaciones.cls:193`). | En auditoría no se encontró la regla de evidencia en esta clase (a verificar). | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionesRealizadas | ESTADO | TEXT(255) | no_conformidades.auditoria_accion_realizada.estado | varchar(32) con CHECK candidato | Derivado: `ACTIVA`, `FINALIZADA`, `PTEREPLANIFICAR`, `IRREGULAR`, `REGISTRADA` (`Entorno.cls:180-184`). Propuesta: vista calculada. | Se graba `Null` si el estado calculado está vacío (`ARAuditoria.cls:205-212`); en alta y edición se recalcula (`ARAuditoriaOperaciones.cls:205-208, 251`). | Catálogo cerrado en código. | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionesRealizadas | Notas | MEMO | no_conformidades.auditoria_accion_realizada.notas | text | Copia directa. | En el alta solo si no está vacía; en edición `Null` si queda vacía (`ARAuditoriaOperaciones.cls:241-243`). | Sin validación. | propuesta |
| NoConformidades_Datos.accdb | TbNCAuditoriaAccionesRealizadas | Responsable | TEXT(255) | no_conformidades.auditoria_accion_realizada.responsable | varchar(255) | Copia directa del identificador de usuario de red; resolver a usuarios del ERD global. | Si está vacío hereda el de la acción correctiva (`ARAuditoriaOperaciones.cls:140-142`); en edición vacío ⇒ `Null` (`ARAuditoriaOperaciones.cls:246-248`). | Relación a usuarios inferida, sin FK en origen. | propuesta |

## TbNCDocumentos

### Uso en `main`

- **Alta**: DAO. `m_SQL = "TbNCDocumentos"`, `OpenRecordset` + `AddNew` en `DocumentoProyectoOperaciones.cls:220-243`; el identificador usa `getID` (`DocumentoProyecto.cls:64`). Tras grabar la fila copia el fichero al directorio de documentación (`DocumentoProyectoOperaciones.cls:256`): la fila y el fichero no son atómicos.
- **Cambio de nombre lógico**: SQL dinámico concatenado, `UPDATE … SET Documento = '…'` por `IDDocumento`, sin escapar el nombre (`DocumentoProyectoOperaciones.cls:302-304`).
- **Baja**: SQL dinámico concatenado, `DELETE` por `IDDocumento` (`DocumentoProyectoOperaciones.cls:46-49`) y después borrado del fichero (`DocumentoProyectoOperaciones.cls:51`); impide borrar el último documento de una acción cerrada (`DocumentoProyectoOperaciones.cls:35-42`).
- **Comprobación de repetidos**: SQL dinámico concatenado, `SELECT *` por `(IDNoConformidad, Documento)` o `(IDAccionRealizada, Documento)` (`DocumentoProyectoOperaciones.cls:98-107`).
- **Lecturas**: ``constructor.bas:519-528``, ``constructor.bas:642-646``, ``constructor.bas:708-712``, ``constructor.bas:737-741`` (SQL dinámico concatenado y embebido).
- **Mantenimiento**: `Instalador.bas:300-329` rellena `IDNoConformidadCalculada`, columna ausente del DDL (véase hallazgos transversales).
- **Transacción**: ninguna. El orden es fila, recálculo de log y copia de fichero (`DocumentoProyectoOperaciones.cls:243-256`).
- **Alcance**: solo proyecto. Los documentos de auditoría viven en `TbDocumentosAuditorias` (`constructor.bas:5026`), fuera de este fichero.
- **Formularios**: [Form_FormNCProyectoDocumentos](../formularios/Form_FormNCProyectoDocumentos.md), [Form_FormARProyectoDocumentos](../formularios/Form_FormARProyectoDocumentos.md).

### Claves y relaciones

- Clave primaria: `IDDocumento`, inferida por código (`DocumentoProyectoOperaciones.cls:46-49` y `DocumentoProyectoOperaciones.cls:302-304`).
- `IDNoConformidad` → `TbNoConformidades.IDNoConformidad`: inferida por código (`DocumentoProyectoOperaciones.cls:98-100`).
- `IDAccionRealizada` → `TbNCAccionesRealizadas.IDAccionRealizada`: inferida por código (`constructor.bas:737-741`).
- Exclusión mutua: un documento pertenece a una NC o a una acción realizada (`DocumentoProyectoOperaciones.cls:142`). No hay restricción en el esquema.

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbNCDocumentos | IDDocumento | LONG(4) | no_conformidades.documento.id | bigint (PK, identidad) | Se descarta el «máximo más uno» (`DocumentoProyecto.cls:64`). | Siempre informado en el alta (`DocumentoProyectoOperaciones.cls:229`). | Clave de baja y renombrado. | propuesta |
| NoConformidades_Datos.accdb | TbNCDocumentos | IDNoConformidad | LONG(4) | no_conformidades.documento.no_conformidad_id | bigint (FK, nullable) | Copia directa. | Solo se graba si viene informado (`DocumentoProyectoOperaciones.cls:230-232`); nulo cuando el documento es de una acción realizada. | Exclusión con `IDAccionRealizada` por código (`DocumentoProyectoOperaciones.cls:142`); CHECK candidato. | propuesta |
| NoConformidades_Datos.accdb | TbNCDocumentos | Documento | TEXT(255) | no_conformidades.documento.nombre | varchar(255) | Copia directa; nombre lógico elegido por el usuario. | Obligatorio («Falta Documento», `DocumentoProyectoOperaciones.cls:89-91`). Duplicado prohibido por NC o por acción realizada (`DocumentoProyectoOperaciones.cls:98-107`, `DocumentoProyectoOperaciones.cls:289-299`). | Único por `(no_conformidad_id)` o `(accion_realizada_id)`; sin escape en el `UPDATE` (`DocumentoProyectoOperaciones.cls:302`), por lo que el destino debe parametrizar. | propuesta |
| NoConformidades_Datos.accdb | TbNCDocumentos | NombreAnexo | TEXT(255) | no_conformidades.documento.nombre_fichero | varchar(255) | Copia directa del nombre del fichero físico (`DocumentoProyectoOperaciones.cls:219`). El contenido está fuera de la base de datos, en un directorio de documentación configurado (`DocumentoProyecto.cls:87`): migrar a almacenamiento de objetos con clave derivada. | Siempre informado en el alta. Duplicados: nombre calculado con ordinal (`DocumentoProyecto.cls:191-215`, a verificar). | El fichero debe existir en el alta (`DocumentoProyectoOperaciones.cls:151-152`, a verificar la línea de comprobación); integridad fila-fichero no garantizada. | propuesta |
| NoConformidades_Datos.accdb | TbNCDocumentos | IDAccionRealizada | LONG(4) | no_conformidades.documento.accion_realizada_id | bigint (FK, nullable) | Copia directa. | Solo si viene informado (`DocumentoProyectoOperaciones.cls:233-235`); nulo para documentos de NC. | Exclusión con `IDNoConformidad`; para cerrar una acción hace falta al menos un documento (`DocumentoProyectoOperaciones.cls:35-42`). | propuesta |
| NoConformidades_Datos.accdb | TbNCDocumentos | IDNoConformidadResultante | LONG(4) | no_conformidades.documento.no_conformidad_resultante_id | bigint (FK, nullable) | Copia directa. Se calcula: el propio `IDNoConformidad` o el de la NC de la acción (`DocumentoProyecto.cls:41-45`). Es redundante: derivable por unión. | Siempre informado en el alta (`DocumentoProyectoOperaciones.cls:222`, `DocumentoProyectoOperaciones.cls:236-238`); registros antiguos pueden carecer de él (a verificar). | Candidata a eliminar en destino si se deriva por unión (decisión del propietario). | propuesta |

## TbNCInformacionRAC

### Uso en `main`

Sin uso en `main`. Búsqueda: `grep -rnwI` del nombre `TbNCInformacionRAC` sobre `src/classes`, `src/modules`, `src/forms` y `src/queries` del commit fijado, sin resultados; y búsqueda sin distinguir mayúsculas de `InformacionRAC` y `AnexosRAC` sobre `src`, sin resultados. Sin formularios asociados.

### Claves y relaciones

Sin evidencia por código. Por nombre, `IDNoConformidad` podría referirse a `TbNoConformidades` y `UsuarioCrea`/`UltimoUsuarioEdita` a usuarios; es una hipótesis a verificar. El significado de «RAC» no consta en `src` (a verificar con el propietario).

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbNCInformacionRAC | IDInformacionRAC | LONG(4) | no_conformidades.informacion_rac.id | bigint (PK candidata) | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNCInformacionRAC | IDNoConformidad | LONG(4) | no_conformidades.informacion_rac.no_conformidad_id | bigint (FK candidata) | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNCInformacionRAC | Informacion | MEMO | no_conformidades.informacion_rac.informacion | text | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNCInformacionRAC | FechaInformacion | SHORT_DATE_TIME(8) | no_conformidades.informacion_rac.fecha_informacion | timestamp without time zone | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNCInformacionRAC | FechaCreacion | SHORT_DATE_TIME(8) | no_conformidades.informacion_rac.fecha_creacion | timestamp without time zone | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNCInformacionRAC | FechaEdicion | SHORT_DATE_TIME(8) | no_conformidades.informacion_rac.fecha_edicion | timestamp without time zone | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNCInformacionRAC | UsuarioCrea | TEXT(255) | no_conformidades.informacion_rac.usuario_crea | varchar(255) | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNCInformacionRAC | UltimoUsuarioEdita | TEXT(255) | no_conformidades.informacion_rac.ultimo_usuario_edita | varchar(255) | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |

## TbNCInformacionRACAnexos

### Uso en `main`

Sin uso en `main`. Búsqueda: `grep -rnwI` del nombre `TbNCInformacionRACAnexos` sobre `src/classes`, `src/modules`, `src/forms` y `src/queries` del commit fijado, sin resultados; y búsqueda sin distinguir mayúsculas de `InformacionRAC` y `AnexosRAC` sobre `src`, sin resultados. Sin formularios asociados.

### Claves y relaciones

Sin evidencia por código. Por nombre, `IDInformacionRAC` podría referirse a `TbNCInformacionRAC.IDInformacionRAC`; hipótesis a verificar. No consta dónde residirían los ficheros de `NombreArchivo`.

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbNCInformacionRACAnexos | IDAnexoInformacionRAC | LONG(4) | no_conformidades.informacion_rac_anexo.id | bigint (PK candidata) | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNCInformacionRACAnexos | IDInformacionRAC | LONG(4) | no_conformidades.informacion_rac_anexo.informacion_rac_id | bigint (FK candidata) | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNCInformacionRACAnexos | NombreArchivo | TEXT(255) | no_conformidades.informacion_rac_anexo.nombre_archivo | varchar(255) | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNCInformacionRACAnexos | FechaAnexo | SHORT_DATE_TIME(8) | no_conformidades.informacion_rac_anexo.fecha_anexo | timestamp without time zone | Copia directa si se decide migrar. | Sin evidencia. | Sin evidencia. | pendiente: sin uso en main, decisión del propietario |

## Pendientes

- ERD de origen real (claves, relaciones, nulabilidad): pendiente del propietario; hasta entonces las relaciones son «inferidas por código».
- Confirmar si `IDNoConformidadCalculada` existe en `TbNCDocumentos` del backend actual (el DDL de 2026-08-06 no la trae).
- Decidir si `ESTADO` y las fechas derivadas de las acciones correctivas, y `ESTADO` de las acciones realizadas, se almacenan o se calculan en destino.
- Decidir el destino de `TbNCARAvisos`, `TbNCInformacionRAC` y `TbNCInformacionRACAnexos` (sin uso en `main`) y el significado de «RAC».
- Definir el almacenamiento de los ficheros de `TbNCDocumentos` y la atomicidad fila-fichero; la coexistencia Access y web de los documentos de auditoría (`TbDocumentosAuditorias`) se trata en otra parte de la matriz.
- Revisar la política de saneamiento de comillas y `;` en textos de acción antes de migrar históricos.
