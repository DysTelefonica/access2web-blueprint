# Matriz campo a campo: lotes, replanificaciones y catálogos auxiliares

Esta parte de la matriz del PRD de No Conformidades (#1080) cubre ocho tablas del esquema previo (`ERD/schema.sql`, volcado del snapshot de 2026-08-06): tres de ingreso por lotes, dos de replanificaciones, `TbTareasExplicaciones`, `TbTipologia` y `TbTiposNCProyectos`. El método es: buscar cada tabla por nombre en `src/forms`, `src/modules`, `src/classes` y `src/queries` del commit fijado de `main` (producción), registrar quién la lee o escribe, de qué tipo es el acceso (DAO, SQL guardada o SQL dinámico concatenado) y si hay transacción, y proponer un destino en el esquema `no_conformidades`. El backend es `NoConformidades_Datos.accdb`. No hay metadatos declarados (el esquema previo solo trae tipos), por lo que toda clave o relación se marca «inferida por código» con su cita, o «sin evidencia». Los tipos de destino son orientativos para PostgreSQL y todo destino es propuesta hasta que se apruebe el ERD global. Las citas son `fichero:línea` dentro de `src/classes` o `src/modules`, salvo indicación. Fichas de formulario relacionadas: [FormNCProyectoReplanificaciones](../formularios/Form_FormNCProyectoReplanificaciones.md), [FormNCAuditoriaReplanificaciones](../formularios/Form_FormNCAuditoriaReplanificaciones.md), [FormNCProyectoTipologiaGestion](../formularios/Form_FormNCProyectoTipologiaGestion.md) y [FormTipologiaNCProyecto](../formularios/Form_FormTipologiaNCProyecto.md).

Hallazgo transversal: en todo `src` no hay `BeginTrans`, `CommitTrans` ni `Rollback`, de modo que ninguna de las escrituras descritas es transaccional.

## TbNoConformidadesIngresoPorLotesDetalle

### Uso en `main`

Sin uso en `src` de `main`. Búsqueda realizada: `grep -rniE "IngresoPorLotes|IDLoteExcel" src` sobre `forms`, `modules`, `classes` y `queries` del clon fijado, sin ninguna coincidencia. La única coincidencia cercana por significado es la propiedad `EsLote` de `Expediente.cls:27` y `Expediente.cls:382`, que no accede a estas tablas (a verificar si guarda relación).

### Claves y relaciones

Sin evidencia de claves ni relaciones. Por nombre, `IDLoteExcel` apuntaría a `TbNoConformidadesIngresoPorLotesPrincipal.IDLoteExcel`, pero no hay código que lo confirme (a verificar con el propietario).

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | IDLoteExcel | LONG(4) | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.id_lote_excel` | integer | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | CodigoNoConformidad | TEXT(255) | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.codigo_no_conformidad` | varchar(255) | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | EsNoConformidad | BOOLEAN(1) | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.es_no_conformidad` | boolean | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | DocumentoDeReferencia | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.documento_de_referencia` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | EXPEDIENTE | TEXT(255) | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.expediente` | varchar(255) | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | PROYECTO | TEXT(255) | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.proyecto` | varchar(255) | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | VEHICULO | TEXT(255) | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.vehiculo` | varchar(255) | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | DESCRIPCION | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.descripcion` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | CAUSA | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.causa` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | ACCIONCORRECTIVA | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.accion_correctiva` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | ACCIONREALIZADA | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.accion_realizada` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | ENTIDADRESPONSABLE | TEXT(50) | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.entidad_responsable` | varchar(50) | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | RESPONSABLETELEFONICA | TEXT(50) | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.responsable_telefonica` | varchar(50) | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | FECHAAPERTURA | SHORT_DATE_TIME(8) | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.fecha_apertura` | date | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | FECHAPREVISTACIERRE | SHORT_DATE_TIME(8) | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.fecha_prevista_cierre` | date | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | FECHACIERRE | SHORT_DATE_TIME(8) | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.fecha_cierre` | date | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | TIPO | TEXT(255) | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.tipo` | varchar(255) | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | NOTAS | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.notas` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesDetalle | DatosDelRegistro | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_detalle.datos_del_registro` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |

## TbNoConformidadesIngresoPorLotesPrincipal

### Uso en `main`

Sin uso en `src` de `main`. Búsqueda realizada: `grep -rniE "IngresoPorLotes|IDLoteExcel" src` sobre `forms`, `modules`, `classes` y `queries` del clon fijado, sin ninguna coincidencia.

### Claves y relaciones

Sin evidencia de claves ni relaciones.

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesPrincipal | IDLoteExcel | LONG(4) | `no_conformidades.no_conformidades_ingreso_por_lotes_principal.id_lote_excel` | integer | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesPrincipal | FechaRegistro | SHORT_DATE_TIME(8) | `no_conformidades.no_conformidades_ingreso_por_lotes_principal.fecha_registro` | timestamp | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesPrincipal | URLCompletaOrigen | TEXT(255) | `no_conformidades.no_conformidades_ingreso_por_lotes_principal.url_completa_origen` | varchar(255) | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesPrincipal | NombreArchivoExcel | TEXT(255) | `no_conformidades.no_conformidades_ingreso_por_lotes_principal.nombre_archivo_excel` | varchar(255) | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesPrincipal | FilaDatosInicial | INT(2) | `no_conformidades.no_conformidades_ingreso_por_lotes_principal.fila_datos_inicial` | smallint | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesPrincipal | Observaciones | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_principal.observaciones` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |

## TbNoConformidadesIngresoPorLotesTemporal

### Uso en `main`

Sin uso en `src` de `main`. Búsqueda realizada: `grep -rniE "IngresoPorLotes|ValidacionDatos" src` sobre `forms`, `modules`, `classes` y `queries` del clon fijado, sin ninguna coincidencia.

### Claves y relaciones

Sin evidencia de claves ni relaciones. Por su nombre es una tabla de trabajo temporal; su migración probablemente no procede (decisión del propietario).

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | IDTmp | LONG(4) | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.id_tmp` | integer | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | EsNoConformidadTmp | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.es_no_conformidad_tmp` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | DocumentoDeReferenciaTmp | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.documento_de_referencia_tmp` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | EXPEDIENTETmp | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.expediente_tmp` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | PROYECTOTmp | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.proyecto_tmp` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | VEHICULOTmp | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.vehiculo_tmp` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | DESCRIPCIONTmp | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.descripcion_tmp` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | CAUSATmp | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.causa_tmp` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | ACCIONCORRECTIVATmp | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.accion_correctiva_tmp` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | ACCIONREALIZADATmp | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.accion_realizada_tmp` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | ENTIDADRESPONSABLETmp | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.entidad_responsable_tmp` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | RESPONSABLETELEFONICATmp | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.responsable_telefonica_tmp` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | FECHAAPERTURATmp | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.fecha_apertura_tmp` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | FECHAPREVISTACIERRETmp | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.fecha_prevista_cierre_tmp` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | FECHACIERRETmp | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.fecha_cierre_tmp` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | TIPOTmp | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.tipo_tmp` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | NOTASTmp | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.notas_tmp` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | blnPasaCriterioParaGrabar | BOOLEAN(1) | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.pasa_criterio_para_grabar` | boolean | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | DatosDelRegistroTmp | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.datos_del_registro_tmp` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbNoConformidadesIngresoPorLotesTemporal | ValidacionDatos | MEMO | `no_conformidades.no_conformidades_ingreso_por_lotes_temporal.validacion_datos` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |

## TbReplanificacionesAuditoria

### Uso en `main`

- **Escritura, alta (DAO)**: `ReplanificacionesAuditoriaOperaciones.cls:126-142` abre la tabla con `OpenRecordset`, `AddNew` y `Update`. Lo invocan `ARAuditoriaOperaciones.cls:269-272` y `ARAuditoriaOperaciones.cls:500-503`.
- **Escritura, baja (SQL dinámico concatenado)**: `DELETE` por `IDReplanificacion` en `ReplanificacionesAuditoriaOperaciones.cls:26-29`.
- **Lectura, una fila (SQL dinámico concatenado)**: `SELECT *` por `IDReplanificacion` en `constructor.bas:2943-2945`.
- **Lectura, por NC (SQL dinámico concatenado)**: `SELECT *` por `IDNoConformidad` en `constructor.bas:3053-3055`; la lanza `NCAuditoria.cls:1152`.
- **Lectura, comprobación de repetida (SQL dinámico concatenado)**: `ReplanificacionesAuditoriaOperaciones.cls:174-177`.
- **Cálculo de identificador**: `constructor.getID` (máximo+1) desde `ReplanificacionesAuditoria.cls:41`.
- **Transacción**: ninguna. El alta no se revierte si falla el registro de log posterior (ReplanificacionesAuditoriaOperaciones.cls:143-146).


### Claves y relaciones

- `IDNoConformidad` → `TbNoConformidadesAuditoria.ID`: **inferida por código** (`ReplanificacionesAuditoriaOperaciones.cls:130`; `NCAuditoria.cls:1152`).
- `IDAccionRealizada` → `TbNCAuditoriaAccionesRealizadas.IDAccionRealizada`: **inferida por código** (`ReplanificacionesAuditoriaOperaciones.cls:132`; `Variables Globales.bas:459`).
- Clave primaria `IDReplanificacion`: **inferida por código** (`constructor.bas:2944-2945`, `ReplanificacionesAuditoria.cls:41`).
- Unicidad lógica (`IDAccionRealizada`, `FechaPrevistaAlInicio`): **inferida por código** (`ReplanificacionesAuditoriaOperaciones.cls:174-177`).
- Sin evidencia de índices ni de borrado en cascada.

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbReplanificacionesAuditoria | IDReplanificacion | LONG(4) | `no_conformidades.replanificaciones_auditoria.id_replanificacion` | integer | Se calcula como máximo+1 (constructor.bas:1165-1178) y se escribe en el alta (ReplanificacionesAuditoriaOperaciones.cls:128-129). En destino: identidad generada por la base de datos | Sin nulos esperados. Duplicado posible por concurrencia (cálculo no atómico, a verificar) | Clave de lectura única (constructor.bas:2944-2945) | propuesta (PK candidata) |
| NoConformidades_Datos.accdb | TbReplanificacionesAuditoria | IDNoConformidad | LONG(4) | `no_conformidades.replanificaciones_auditoria.id_no_conformidad` | integer | Se copia del identificador de la NC (ReplanificacionesAuditoriaOperaciones.cls:130-131) | Sin nulos esperados; se filtra por este campo (constructor.bas:3053-3055) | FK inferida por código hacia `TbNoConformidadesAuditoria.ID` (ReplanificacionesAuditoriaOperaciones.cls:130; Variables Globales.bas:460) | propuesta (FK candidata) |
| NoConformidades_Datos.accdb | TbReplanificacionesAuditoria | IDAccionRealizada | LONG(4) | `no_conformidades.replanificaciones_auditoria.id_accion_realizada` | integer | Se copia de la AR (ReplanificacionesAuditoriaOperaciones.cls:132-133) | Sin nulos esperados. Con `FechaPrevistaAlInicio` no se admiten duplicados (ReplanificacionesAuditoriaOperaciones.cls:174-177) | FK inferida por código hacia `TbNCAuditoriaAccionesRealizadas.IDAccionRealizada` (Variables Globales.bas:459) | propuesta (FK candidata) |
| NoConformidades_Datos.accdb | TbReplanificacionesAuditoria | FechaReprogramacion | SHORT_DATE_TIME(8) | `no_conformidades.replanificaciones_auditoria.fecha_reprogramacion` | date | Se fija con la fecha del día en el alta (ReplanificacionesAuditoriaOperaciones.cls:124, 134); sin hora | Sin nulos esperados | Ninguna adicional en código | propuesta |
| NoConformidades_Datos.accdb | TbReplanificacionesAuditoria | FechaPrevistaAlInicio | SHORT_DATE_TIME(8) | `no_conformidades.replanificaciones_auditoria.fecha_prevista_al_inicio` | date | Si llega vacía se toma de la fecha fin prevista de la AR (ReplanificacionesAuditoriaOperaciones.cls:70-76); en consultas se formatea `mm/dd/yyyy` (ReplanificacionesAuditoriaOperaciones.cls:174-177) | Sin nulos esperados; duplicado (AR, fecha) rechazado: «Hay registrada al menos una replanificación con la misma fecha inicial prevista» | Obligatoria (ReplanificacionesAuditoriaOperaciones.cls:70-76) | propuesta (única candidata con `id_accion_realizada`) |
| NoConformidades_Datos.accdb | TbReplanificacionesAuditoria | FechaPrevistaReplanificada | SHORT_DATE_TIME(8) | `no_conformidades.replanificaciones_auditoria.fecha_prevista_replanificada` | date | Se copia de la fecha fin prevista de la AR (ARAuditoriaOperaciones.cls:266) | Sin nulos esperados | Debe ser fecha válida (`IsDate`, ReplanificacionesAuditoriaOperaciones.cls:77-80) | propuesta |
| NoConformidades_Datos.accdb | TbReplanificacionesAuditoria | Observaciones | MEMO | `no_conformidades.replanificaciones_auditoria.observaciones` | text | Solo se escribe si no está vacía (ReplanificacionesAuditoriaOperaciones.cls:138-139); al leer, nulo se convierte en cadena vacía (constructor.bas:2956) | Nulo posible en datos antiguos (a verificar); sin restricción de duplicados | Obligatoria al registrar: «Se ha de registrar el motivo de la reprogramación» (ReplanificacionesAuditoriaOperaciones.cls:89-92) | propuesta (NOT NULL a decidir) |

## TbReplanificacionesProyecto

### Uso en `main`

- **Escritura, alta (DAO)**: `ReplanificacionesProyectoOperaciones.cls:129-145` abre la tabla con `OpenRecordset`, `AddNew` y `Update`. Lo invocan `ARProyectoOperaciones.cls:281-284` y `ARProyectoOperaciones.cls:519-522`.
- **Escritura, baja (SQL dinámico concatenado)**: `DELETE` por `IDReplanificacion` en `ReplanificacionesProyectoOperaciones.cls:26-29`.
- **Lectura, una fila (SQL dinámico concatenado)**: `SELECT *` por `IDReplanificacion` en `constructor.bas:2896-2898`.
- **Lectura, por NC (SQL dinámico concatenado)**: `SELECT *` por `IDNoConformidad` en `constructor.bas:2992-2994`; además `CacheNCProyecto.bas:933-935` (con `ORDER BY`) y `IndicadorRepositorio.bas:25` (INNER JOIN por periodo); la lanza `NCProyecto.cls:690`.
- **Lectura, comprobación de repetida (SQL dinámico concatenado)**: `ReplanificacionesProyectoOperaciones.cls:180-183`.
- **Cálculo de identificador**: `constructor.getID` (máximo+1) desde `ReplanificacionesProyecto.cls:41`.
- **Transacción**: ninguna. El alta no se revierte si falla el registro de log posterior (ReplanificacionesProyectoOperaciones.cls:149-152) ni la invalidación de caché (`CacheNCProyecto.InvalidarCache`, línea 147).


### Claves y relaciones

- `IDNoConformidad` → `TbNoConformidades.IDNoConformidad`: **inferida por código** (`IndicadorRepositorio.bas:25`; `ReplanificacionesProyectoOperaciones.cls:133`).
- `IDAccionRealizada` → `TbNCAccionesRealizadas.IDAccionRealizada`: **inferida por código** (`ReplanificacionesProyectoOperaciones.cls:135`; `Variables Globales.bas:445`).
- Clave primaria `IDReplanificacion`: **inferida por código** (`constructor.bas:2897-2898`, `ReplanificacionesProyecto.cls:41`).
- Unicidad lógica (`IDAccionRealizada`, `FechaPrevistaAlInicio`): **inferida por código** (`ReplanificacionesProyectoOperaciones.cls:180-183`).
- Sin evidencia de índices ni de borrado en cascada.

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbReplanificacionesProyecto | IDReplanificacion | LONG(4) | `no_conformidades.replanificaciones_proyecto.id_replanificacion` | integer | Se calcula como máximo+1 (constructor.bas:1165-1178) y se escribe en el alta (ReplanificacionesProyectoOperaciones.cls:131-132). En destino: identidad generada por la base de datos | Sin nulos esperados. Duplicado posible por concurrencia (cálculo no atómico, a verificar) | Clave de lectura única (constructor.bas:2897-2898) | propuesta (PK candidata) |
| NoConformidades_Datos.accdb | TbReplanificacionesProyecto | IDNoConformidad | LONG(4) | `no_conformidades.replanificaciones_proyecto.id_no_conformidad` | integer | Se copia del identificador de la NC (ReplanificacionesProyectoOperaciones.cls:133-134) | Sin nulos esperados; se filtra por este campo (constructor.bas:2992-2994) | FK inferida por código hacia `TbNoConformidades.IDNoConformidad` (constructor.bas:2992-2994; IndicadorRepositorio.bas:25) | propuesta (FK candidata) |
| NoConformidades_Datos.accdb | TbReplanificacionesProyecto | IDAccionRealizada | LONG(4) | `no_conformidades.replanificaciones_proyecto.id_accion_realizada` | integer | Se copia de la AR (ReplanificacionesProyectoOperaciones.cls:135-136) | Sin nulos esperados. Con `FechaPrevistaAlInicio` no se admiten duplicados (ReplanificacionesProyectoOperaciones.cls:180-183) | FK inferida por código hacia `TbNCAccionesRealizadas.IDAccionRealizada` (Variables Globales.bas:445) | propuesta (FK candidata) |
| NoConformidades_Datos.accdb | TbReplanificacionesProyecto | FechaReprogramacion | SHORT_DATE_TIME(8) | `no_conformidades.replanificaciones_proyecto.fecha_reprogramacion` | date | Se fija con la fecha del día en el alta (ReplanificacionesProyectoOperaciones.cls:127, 137); sin hora | Sin nulos esperados; filtro de periodo en IndicadorRepositorio.bas:27 | Ninguna adicional en código | propuesta |
| NoConformidades_Datos.accdb | TbReplanificacionesProyecto | FechaPrevistaAlInicio | SHORT_DATE_TIME(8) | `no_conformidades.replanificaciones_proyecto.fecha_prevista_al_inicio` | date | Si llega vacía se toma de la fecha fin prevista de la AR (ReplanificacionesProyectoOperaciones.cls:73-79); en consultas se formatea `mm/dd/yyyy` (ReplanificacionesProyectoOperaciones.cls:180-183) | Sin nulos esperados; duplicado (AR, fecha) rechazado: «Hay registrada al menos una replanificación con la misma fecha inicial prevista» | Obligatoria (ReplanificacionesProyectoOperaciones.cls:73-79) | propuesta (única candidata con `id_accion_realizada`) |
| NoConformidades_Datos.accdb | TbReplanificacionesProyecto | FechaPrevistaReplanificada | SHORT_DATE_TIME(8) | `no_conformidades.replanificaciones_proyecto.fecha_prevista_replanificada` | date | Se copia de la fecha fin prevista de la AR (ARProyectoOperaciones.cls:278) | Sin nulos esperados | Debe ser fecha válida (`IsDate`, ReplanificacionesProyectoOperaciones.cls:80-83) | propuesta |
| NoConformidades_Datos.accdb | TbReplanificacionesProyecto | Observaciones | MEMO | `no_conformidades.replanificaciones_proyecto.observaciones` | text | Solo se escribe si no está vacía (ReplanificacionesProyectoOperaciones.cls:141-142); al leer, nulo se convierte en cadena vacía (constructor.bas:2909) | Nulo posible en datos antiguos (a verificar); sin restricción de duplicados | Obligatoria al registrar: «Se ha de registrar el motivo de la reprogramación» (ReplanificacionesProyectoOperaciones.cls:92-95) | propuesta (NOT NULL a decidir) |

## TbTareasExplicaciones

### Uso en `main`

Sin uso en `src` de `main`. Búsqueda realizada: `grep -rniE "TareasExplicaciones|NodoTarea" src` sobre `forms`, `modules`, `classes` y `queries` del clon fijado, sin ninguna coincidencia.

### Claves y relaciones

Sin evidencia de claves ni relaciones.

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbTareasExplicaciones | NodoTarea | TEXT(255) | `no_conformidades.tareas_explicaciones.nodo_tarea` | varchar(255) | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbTareasExplicaciones | TituloTarea | TEXT(255) | `no_conformidades.tareas_explicaciones.titulo_tarea` | varchar(255) | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbTareasExplicaciones | Explicacion | MEMO | `no_conformidades.tareas_explicaciones.explicacion` | text | Sin evidencia de uso; transformación no definida | Sin evidencia | Sin evidencia | pendiente: sin uso en main, decisión del propietario |

## TbTipologia

### Uso en `main`

- **Lectura (SQL dinámico concatenado, JOIN con `TbNoConformidades.TIPO`)**: `Instalador.bas:667-668` (`RellenarTablaTipologia`) y `Instalador.bas:727-728` (`RellenarTipologia`, un `UPDATE` sobre `TbNoConformidades.IDTipo`).
- **Escritura**: ninguna en `src` (no hay alta, edición ni baja de `TbTipologia`).
- **Quién lo llama**: `RellenarTipologia` se invoca desde `Instalador.bas:80`. Es un paso de la rutina de instalación que rellena `TbTiposNCProyectos` a partir del código antiguo; en ejecución normal ningún formulario lee esta tabla (búsqueda en `src/forms` sin coincidencias).
- **Transacción**: ninguna.


### Claves y relaciones

- `CodTipologia` ↔ `TbNoConformidades.TIPO`: **inferida por código** (`Instalador.bas:668`).
- `Tipologia` ↔ `TbTiposNCProyectos.Tipologia` (unión por texto): **inferida por código** (`Instalador.bas:728`).
- Clave primaria: sin evidencia.

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbTipologia | CodTipologia | TEXT(2) | `no_conformidades.tipologia.cod_tipologia` | varchar(2) | Sin transformación; en la migración histórica se cruza con `TbNoConformidades.TIPO` (Instalador.bas:668) | Duplicados no validados por código (sin evidencia) | Sin evidencia | propuesta (catálogo de origen; ver Pendientes) |
| NoConformidades_Datos.accdb | TbTipologia | Tipologia | TEXT(255) | `no_conformidades.tipologia.tipologia` | varchar(255) | Se vuelca como texto a `TbTiposNCProyectos.Tipologia` (Instalador.bas:667-703) | Sin evidencia de nulos; duplicados absorbidos por `DISTINCT` y por el diccionario (Instalador.bas:667, 677-679) | Sin evidencia | propuesta |

## TbTiposNCProyectos

### Uso en `main`

- **Lectura por identificador o por texto (SQL dinámico concatenado)**: `constructor.bas:1357-1366`; listado ordenado por `Tipologia` en `constructor.bas:1408-1410`; lectura `SELECT *` en `TipologiaNCProyectos.cls:154-155`.
- **Escritura, alta (DAO)**: `TipologiaNCProyectos.cls:151-170` (`AddNew`/`Update`).
- **Escritura, edición (DAO)**: `TipologiaNCProyectos.cls:154-155, 176-179` (`Edit`/`Update`).
- **Escritura, baja (SQL dinámico concatenado)**: `DELETE` por `IDTipo` en `TipologiaNCProyectos.cls:223-226`.
- **Escritura, carga inicial (DAO vía `Registrar`)**: `Instalador.bas:690-703`; y `UPDATE` de `TbNoConformidades.IDTipo` en `Instalador.bas:727-730`.
- **Lectura en consultas de seguimiento (SQL concatenado, LEFT JOIN)**: `Variables Globales.bas:449, 457`.
- **Transacción**: ninguna; el log de alta o edición se registra después de la escritura (`TipologiaNCProyectos.cls:171-174, 180-183`) y un fallo del log no la revierte.
- **Formularios**: [FormNCProyectoTipologiaGestion](../formularios/Form_FormNCProyectoTipologiaGestion.md) y [FormTipologiaNCProyecto](../formularios/Form_FormTipologiaNCProyecto.md).


### Claves y relaciones

- `TbNoConformidades.IDTipo` → `IDTipo`: **inferida por código** (`constructor.bas:1588-1589`; `Variables Globales.bas:457`; `Instalador.bas:729`).
- Clave primaria `IDTipo`: **inferida por código** (`TipologiaNCProyectos.cls:154-155`, `constructor.bas:1360`).
- Unicidad de `Tipologia`: **inferida por código** (`TipologiaNCProyectos.cls:60-68`).
- Las NC de auditoría no usan esta tabla: `TbNoConformidadesAuditoria.Tipo` se lee como texto (`Variables Globales.bas:462`), a verificar.

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbTiposNCProyectos | IDTipo | LONG(4) | `no_conformidades.tipos_nc_proyectos.id_tipo` | integer | Se calcula máximo+1 en el alta (TipologiaNCProyectos.cls:27, 152, 168; constructor.bas:1165-1178). En destino: identidad generada | Sin nulos esperados; clave de lectura (TipologiaNCProyectos.cls:154-155) | Se rechaza el borrado si hay NC que lo usan (TipologiaNCProyectos.cls:108-113) | propuesta (PK candidata) |
| NoConformidades_Datos.accdb | TbTiposNCProyectos | Tipologia | TEXT(255) | `no_conformidades.tipos_nc_proyectos.tipologia` | varchar(255) | Sin transformación; búsqueda por igualdad de texto (constructor.bas:1363-1365) | Sin nulos esperados; duplicados rechazados: «La tipología está repetida» (TipologiaNCProyectos.cls:60-68, 71-80) | Única por código (TipologiaNCProyectos.cls:66); apóstrofos sin escapar en el SQL concatenado (constructor.bas:1365, a verificar) | propuesta (única candidata) |

## Pendientes

- Las tres tablas de ingreso por lotes y `TbTareasExplicaciones` no tienen uso en `main`: el propietario decide si se migran, se archivan o se descartan. `TbNoConformidadesIngresoPorLotesTemporal` parece de trabajo temporal.
- Confirmar si `EsLote` de `Expediente.cls:27` está relacionado con la antigua carga por lotes (a verificar).
- `TbTipologia` solo interviene en la rutina de instalación; decidir si se conserva como catálogo histórico de códigos o se descarta tras fijar `TbTiposNCProyectos`.
- Generación de `IDReplanificacion` y `IDTipo` por máximo+1 (`constructor.bas:1165-1178`): en destino, identidad de base de datos; confirmar que no se conservan los identificadores históricos con significado.
- Las claves y relaciones son inferidas por código; el ERD de origen real está pendiente de la extracción del propietario.
- Nulabilidad y valores por defecto de las tablas de replanificaciones y de tipologías: sin evidencia en el esquema previo.
