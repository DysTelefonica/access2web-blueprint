# Matriz campo a campo: caché de indicadores y de listados (No Conformidades)

Parte «b-cache-indicadores» de la matriz campo a campo del PRD de No Conformidades (#1080). Cubre las tablas del esquema previo `TbCacheIndicadoresAuditoriaDetalle`, `TbCacheIndicadoresAuditoriaHeader`, `TbCacheIndicadoresConfig`, `TbCacheIndicadoresProyectoDetalle`, `TbCacheIndicadoresProyectoHeader` y `TbCacheListadoNC`. Índice de formularios y capacidades en [el PRD de la aplicación](../README.md); las fichas de formulario están en `../formularios/`.

## Método

La línea base es el código de `main` (producción), commit `1d7a2f826a1e3c582874e4e29886a6183efb9067`, directorios `src/classes`, `src/modules`, `src/forms` y `src/queries`. Los tipos de origen proceden del DDL de `ERD/schema.sql` (volcado de 2026-08-06), que solo declara tipos: no declara claves, relaciones ni nulabilidad, y no hay backend legible en esta fase. Por eso toda clave o relación se marcaría «inferida por código» con su cita, y se escribe «sin evidencia» cuando el código no la muestra. El backend de todas las filas es `NoConformidades_Datos.accdb`. Los destinos (`no_conformidades.<tabla>.<columna>`) y los tipos PostgreSQL son orientativos y quedan como propuesta hasta que se apruebe el ERD global.

Búsqueda realizada: coincidencia insensible a mayúsculas, sobre todo el contenido de `src/classes`, `src/modules`, `src/forms` y `src/queries`, de cada nombre de tabla y de las cadenas `TbCacheIndicadores`, `TbCacheListado`, `CacheIndicadores`, `CacheListadoNC`, `IndicadoresAuditoria`, `IndicadoresProyecto`, y de columnas características (`IDCacheConfig`, `IDCacheIndicadorProyecto`, `FechaSnapshot`, `VersionRegla`, `ClaveEntidad`). Resultado: **0 coincidencias** en todas ellas. Tampoco hay construcción dinámica del nombre: la única tabla de caché con prefijo `TbCache` que aparece en `src` es `TbCacheNCProyecto` (declarada en `CacheNCProyecto.bas:34` y manejada en `InicializadorCache.bas:117`), que no pertenece a esta parte.

Coincide con el inventario del PRD, que registra la capa de caché de indicadores y de listados como retirada frente al estudio previo (véase [Objetos retirados frente al estudio previo](../README.md)): `ModuloCacheIndicadores.bas`, `ModuloCacheIndicadoresIssue18.bas`, `NCAuditoriaListadoCache.bas` y las clases `CacheNCCacheRepositorio.cls`, `CacheNCService.cls` y `CacheNCCrud.cls` no existen en `main`. Esos nombres se citan en texto plano porque ya no hay fichero que enlazar. En consecuencia, no se analiza el diseño de estas tablas: cada una queda «pendiente: sin uso en main, decisión del propietario» (si se conserva, se migra, o se descarta con su capa de caché).

Hallazgo transversal: el aplicativo en `main` no consume ninguna de estas tablas, por lo que no hay lecturas, escrituras ni transacciones que describir. Si el backend de producción las conserva, son datos huérfanos del código de `main`; confirmarlo exige el ERD de origen real, pendiente del propietario.

## TbCacheIndicadoresAuditoriaDetalle

### Uso en `main`

- **Lectura y escritura**: ninguna. Búsqueda descrita en el apartado «Método»: 0 coincidencias del nombre `TbCacheIndicadoresAuditoriaDetalle` en `src/classes`, `src/modules`, `src/forms` y `src/queries`.
- **Transacción**: no aplica (sin uso).
- **Acceso**: ni DAO, ni SQL guardada, ni SQL dinámico concatenado.

### Claves y relaciones

- Clave primaria: sin evidencia.
- Relaciones: sin evidencia (el código de `main` no referencia la tabla; ninguna se infiere por código).

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | ID | LONG(4) | no_conformidades.cache_indicadores_auditoria_detalle.id | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | IDCacheIndicadorProyecto | LONG(4) | no_conformidades.cache_indicadores_auditoria_detalle.id_cache_indicador_proyecto | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | IDCacheConfig | LONG(4) | no_conformidades.cache_indicadores_auditoria_detalle.id_cache_config | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | Dominio | TEXT(32) | no_conformidades.cache_indicadores_auditoria_detalle.dominio | varchar(32) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | Bucket | TEXT(50) | no_conformidades.cache_indicadores_auditoria_detalle.bucket | varchar(50) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | TipoFila | TEXT(50) | no_conformidades.cache_indicadores_auditoria_detalle.tipo_fila | varchar(50) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | IDEntidad | LONG(4) | no_conformidades.cache_indicadores_auditoria_detalle.id_entidad | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | IDNoConformidad | LONG(4) | no_conformidades.cache_indicadores_auditoria_detalle.id_no_conformidad | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | IDAccionCorrectiva | LONG(4) | no_conformidades.cache_indicadores_auditoria_detalle.id_accion_correctiva | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | IDAccionRealizada | LONG(4) | no_conformidades.cache_indicadores_auditoria_detalle.id_accion_realizada | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | IDTarea | LONG(4) | no_conformidades.cache_indicadores_auditoria_detalle.id_tarea | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | ResponsableCalidad | TEXT(255) | no_conformidades.cache_indicadores_auditoria_detalle.responsable_calidad | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | ResponsableUsuarioRed | TEXT(255) | no_conformidades.cache_indicadores_auditoria_detalle.responsable_usuario_red | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | DisplayTitulo | TEXT(255) | no_conformidades.cache_indicadores_auditoria_detalle.display_titulo | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | DisplaySubtitulo | MEMO | no_conformidades.cache_indicadores_auditoria_detalle.display_subtitulo | text | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | FechaSnapshot | SHORT_DATE_TIME(8) | no_conformidades.cache_indicadores_auditoria_detalle.fecha_snapshot | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | ClaveEntidad | TEXT(255) | no_conformidades.cache_indicadores_auditoria_detalle.clave_entidad | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | OrigenTabla | TEXT(255) | no_conformidades.cache_indicadores_auditoria_detalle.origen_tabla | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaDetalle | VersionRegla | TEXT(64) | no_conformidades.cache_indicadores_auditoria_detalle.version_regla | varchar(64) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |

## TbCacheIndicadoresAuditoriaHeader

### Uso en `main`

- **Lectura y escritura**: ninguna. Búsqueda descrita en el apartado «Método»: 0 coincidencias del nombre `TbCacheIndicadoresAuditoriaHeader` en `src/classes`, `src/modules`, `src/forms` y `src/queries`.
- **Transacción**: no aplica (sin uso).
- **Acceso**: ni DAO, ni SQL guardada, ni SQL dinámico concatenado.

### Claves y relaciones

- Clave primaria: sin evidencia.
- Relaciones: sin evidencia (el código de `main` no referencia la tabla; ninguna se infiere por código).

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaHeader | ID | LONG(4) | no_conformidades.cache_indicadores_auditoria_header.id | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaHeader | IDCacheConfig | LONG(4) | no_conformidades.cache_indicadores_auditoria_header.id_cache_config | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaHeader | Dominio | TEXT(32) | no_conformidades.cache_indicadores_auditoria_header.dominio | varchar(32) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaHeader | FechaSincronizacion | SHORT_DATE_TIME(8) | no_conformidades.cache_indicadores_auditoria_header.fecha_sincronizacion | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaHeader | UsuarioSincronizacion | TEXT(255) | no_conformidades.cache_indicadores_auditoria_header.usuario_sincronizacion | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaHeader | Estado | TEXT(50) | no_conformidades.cache_indicadores_auditoria_header.estado | varchar(50) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresAuditoriaHeader | ErrorUltimaSincronizacion | MEMO | no_conformidades.cache_indicadores_auditoria_header.error_ultima_sincronizacion | text | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |

## TbCacheIndicadoresConfig

### Uso en `main`

- **Lectura y escritura**: ninguna. Búsqueda descrita en el apartado «Método»: 0 coincidencias del nombre `TbCacheIndicadoresConfig` en `src/classes`, `src/modules`, `src/forms` y `src/queries`.
- **Transacción**: no aplica (sin uso).
- **Acceso**: ni DAO, ni SQL guardada, ni SQL dinámico concatenado.

### Claves y relaciones

- Clave primaria: sin evidencia.
- Relaciones: sin evidencia (el código de `main` no referencia la tabla; ninguna se infiere por código).

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbCacheIndicadoresConfig | IDCacheConfig | LONG(4) | no_conformidades.cache_indicadores_config.id_cache_config | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresConfig | Dominio | TEXT(32) | no_conformidades.cache_indicadores_config.dominio | varchar(32) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresConfig | Activo | BOOLEAN(1) | no_conformidades.cache_indicadores_config.activo | boolean | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresConfig | VersionRegla | TEXT(64) | no_conformidades.cache_indicadores_config.version_regla | varchar(64) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresConfig | FechaConfiguracion | SHORT_DATE_TIME(8) | no_conformidades.cache_indicadores_config.fecha_configuracion | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresConfig | UsuarioConfiguracion | TEXT(255) | no_conformidades.cache_indicadores_config.usuario_configuracion | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |

## TbCacheIndicadoresProyectoDetalle

### Uso en `main`

- **Lectura y escritura**: ninguna. Búsqueda descrita en el apartado «Método»: 0 coincidencias del nombre `TbCacheIndicadoresProyectoDetalle` en `src/classes`, `src/modules`, `src/forms` y `src/queries`.
- **Transacción**: no aplica (sin uso).
- **Acceso**: ni DAO, ni SQL guardada, ni SQL dinámico concatenado.

### Claves y relaciones

- Clave primaria: sin evidencia.
- Relaciones: sin evidencia (el código de `main` no referencia la tabla; ninguna se infiere por código).

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | IDCacheDetalle | LONG(4) | no_conformidades.cache_indicadores_proyecto_detalle.id_cache_detalle | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | IDCacheIndicadorProyecto | LONG(4) | no_conformidades.cache_indicadores_proyecto_detalle.id_cache_indicador_proyecto | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | Bucket | TEXT(64) | no_conformidades.cache_indicadores_proyecto_detalle.bucket | varchar(64) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | TipoFila | TEXT(16) | no_conformidades.cache_indicadores_proyecto_detalle.tipo_fila | varchar(16) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | IDEntidad | LONG(4) | no_conformidades.cache_indicadores_proyecto_detalle.id_entidad | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | IDNoConformidad | LONG(4) | no_conformidades.cache_indicadores_proyecto_detalle.id_no_conformidad | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | IDAccionCorrectiva | LONG(4) | no_conformidades.cache_indicadores_proyecto_detalle.id_accion_correctiva | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | IDAccionRealizada | LONG(4) | no_conformidades.cache_indicadores_proyecto_detalle.id_accion_realizada | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | ResponsableCalidad | TEXT(255) | no_conformidades.cache_indicadores_proyecto_detalle.responsable_calidad | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | CodigoNoConformidad | TEXT(255) | no_conformidades.cache_indicadores_proyecto_detalle.codigo_no_conformidad | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | Descripcion | MEMO | no_conformidades.cache_indicadores_proyecto_detalle.descripcion | text | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | Nemotecnico | TEXT(255) | no_conformidades.cache_indicadores_proyecto_detalle.nemotecnico | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | Tarea | MEMO | no_conformidades.cache_indicadores_proyecto_detalle.tarea | text | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | Estado | TEXT(255) | no_conformidades.cache_indicadores_proyecto_detalle.estado | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | Tecnico | TEXT(255) | no_conformidades.cache_indicadores_proyecto_detalle.tecnico | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | TipoNC | TEXT(255) | no_conformidades.cache_indicadores_proyecto_detalle.tipo_nc | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | IDExpediente | LONG(4) | no_conformidades.cache_indicadores_proyecto_detalle.id_expediente | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | NAccion | TEXT(50) | no_conformidades.cache_indicadores_proyecto_detalle.n_accion | varchar(50) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | FechaInicio | SHORT_DATE_TIME(8) | no_conformidades.cache_indicadores_proyecto_detalle.fecha_inicio | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | FechaFinPrevista | SHORT_DATE_TIME(8) | no_conformidades.cache_indicadores_proyecto_detalle.fecha_fin_prevista | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | FechaFinReal | SHORT_DATE_TIME(8) | no_conformidades.cache_indicadores_proyecto_detalle.fecha_fin_real | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | FechaCierre | SHORT_DATE_TIME(8) | no_conformidades.cache_indicadores_proyecto_detalle.fecha_cierre | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | RequiereControlEficacia | TEXT(50) | no_conformidades.cache_indicadores_proyecto_detalle.requiere_control_eficacia | varchar(50) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | ResultadoControlEficacia | TEXT(255) | no_conformidades.cache_indicadores_proyecto_detalle.resultado_control_eficacia | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | FechaSnapshot | SHORT_DATE_TIME(8) | no_conformidades.cache_indicadores_proyecto_detalle.fecha_snapshot | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | IDCacheConfig | LONG(4) | no_conformidades.cache_indicadores_proyecto_detalle.id_cache_config | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | Dominio | TEXT(32) | no_conformidades.cache_indicadores_proyecto_detalle.dominio | varchar(32) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | ClaveEntidad | TEXT(128) | no_conformidades.cache_indicadores_proyecto_detalle.clave_entidad | varchar(128) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | IDTarea | LONG(4) | no_conformidades.cache_indicadores_proyecto_detalle.id_tarea | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | OrigenTabla | TEXT(64) | no_conformidades.cache_indicadores_proyecto_detalle.origen_tabla | varchar(64) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | ResponsableUsuarioRed | TEXT(255) | no_conformidades.cache_indicadores_proyecto_detalle.responsable_usuario_red | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | DisplayTitulo | TEXT(255) | no_conformidades.cache_indicadores_proyecto_detalle.display_titulo | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | DisplaySubtitulo | MEMO | no_conformidades.cache_indicadores_proyecto_detalle.display_subtitulo | text | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | FechaActualizacionEntidad | SHORT_DATE_TIME(8) | no_conformidades.cache_indicadores_proyecto_detalle.fecha_actualizacion_entidad | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoDetalle | VersionRegla | TEXT(64) | no_conformidades.cache_indicadores_proyecto_detalle.version_regla | varchar(64) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |

## TbCacheIndicadoresProyectoHeader

### Uso en `main`

- **Lectura y escritura**: ninguna. Búsqueda descrita en el apartado «Método»: 0 coincidencias del nombre `TbCacheIndicadoresProyectoHeader` en `src/classes`, `src/modules`, `src/forms` y `src/queries`.
- **Transacción**: no aplica (sin uso).
- **Acceso**: ni DAO, ni SQL guardada, ni SQL dinámico concatenado.

### Claves y relaciones

- Clave primaria: sin evidencia.
- Relaciones: sin evidencia (el código de `main` no referencia la tabla; ninguna se infiere por código).

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoHeader | IDCacheIndicadorProyecto | LONG(4) | no_conformidades.cache_indicadores_proyecto_header.id_cache_indicador_proyecto | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoHeader | FechaSincronizacion | SHORT_DATE_TIME(8) | no_conformidades.cache_indicadores_proyecto_header.fecha_sincronizacion | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoHeader | UsuarioSincronizacion | TEXT(255) | no_conformidades.cache_indicadores_proyecto_header.usuario_sincronizacion | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoHeader | Estado | TEXT(50) | no_conformidades.cache_indicadores_proyecto_header.estado | varchar(50) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoHeader | ErrorUltimaSincronizacion | MEMO | no_conformidades.cache_indicadores_proyecto_header.error_ultima_sincronizacion | text | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoHeader | IDCacheConfig | LONG(4) | no_conformidades.cache_indicadores_proyecto_header.id_cache_config | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoHeader | Dominio | TEXT(32) | no_conformidades.cache_indicadores_proyecto_header.dominio | varchar(32) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoHeader | VersionRegla | TEXT(64) | no_conformidades.cache_indicadores_proyecto_header.version_regla | varchar(64) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoHeader | MotivoSincronizacion | TEXT(64) | no_conformidades.cache_indicadores_proyecto_header.motivo_sincronizacion | varchar(64) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoHeader | IDNoConformidadUltimaSync | LONG(4) | no_conformidades.cache_indicadores_proyecto_header.id_no_conformidad_ultima_sync | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoHeader | FechaUltimaSincronizacionNC | SHORT_DATE_TIME(8) | no_conformidades.cache_indicadores_proyecto_header.fecha_ultima_sincronizacion_nc | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheIndicadoresProyectoHeader | OperadorSync | TEXT(64) | no_conformidades.cache_indicadores_proyecto_header.operador_sync | varchar(64) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |

## TbCacheListadoNC

### Uso en `main`

- **Lectura y escritura**: ninguna. Búsqueda descrita en el apartado «Método»: 0 coincidencias del nombre `TbCacheListadoNC` en `src/classes`, `src/modules`, `src/forms` y `src/queries`.
- **Transacción**: no aplica (sin uso).
- **Acceso**: ni DAO, ni SQL guardada, ni SQL dinámico concatenado.

### Claves y relaciones

- Clave primaria: sin evidencia.
- Relaciones: sin evidencia (el código de `main` no referencia la tabla; ninguna se infiere por código).

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbCacheListadoNC | IDNoConformidad | LONG(4) | no_conformidades.cache_listado_nc.id_no_conformidad | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | Version | LONG(4) | no_conformidades.cache_listado_nc.version | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | CodigoNoConformidad | TEXT(255) | no_conformidades.cache_listado_nc.codigo_no_conformidad | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | IDExpediente | LONG(4) | no_conformidades.cache_listado_nc.id_expediente | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | Nemotecnico | TEXT(255) | no_conformidades.cache_listado_nc.nemotecnico | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | CodExp | TEXT(255) | no_conformidades.cache_listado_nc.cod_exp | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | JuridicaExp | TEXT(255) | no_conformidades.cache_listado_nc.juridica_exp | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | IDTipo | LONG(4) | no_conformidades.cache_listado_nc.id_tipo | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | Descripcion | MEMO | no_conformidades.cache_listado_nc.descripcion | text | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | Notas | MEMO | no_conformidades.cache_listado_nc.notas | text | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | Estado | TEXT(100) | no_conformidades.cache_listado_nc.estado | varchar(100) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | FechaApertura | SHORT_DATE_TIME(8) | no_conformidades.cache_listado_nc.fecha_apertura | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | FechaCierre | SHORT_DATE_TIME(8) | no_conformidades.cache_listado_nc.fecha_cierre | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | RequiereControlEficacia | TEXT(10) | no_conformidades.cache_listado_nc.requiere_control_eficacia | varchar(10) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | ControlEficacia | TEXT(255) | no_conformidades.cache_listado_nc.control_eficacia | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | ResponsableTelefonica | TEXT(255) | no_conformidades.cache_listado_nc.responsable_telefonica | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | RESPONSABLECALIDAD | TEXT(255) | no_conformidades.cache_listado_nc.responsablecalidad | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | ACR | TEXT(255) | no_conformidades.cache_listado_nc.acr | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | Cerrada | TEXT(10) | no_conformidades.cache_listado_nc.cerrada | varchar(10) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | FechaCache | SHORT_DATE_TIME(8) | no_conformidades.cache_listado_nc.fecha_cache | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNC | CacheValida | BOOLEAN(1) | no_conformidades.cache_listado_nc.cache_valida | boolean | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |

## Pendientes

- Decisión del propietario sobre las seis tablas: conservar, migrar o descartar junto con la capa de caché retirada. Sin esa decisión no se propone destino definitivo.
- Confirmar con el ERD de origen real si las tablas siguen existiendo en el backend de producción y si contienen datos que haya que conservar (extracción pendiente del propietario).
- Si se decidiera reactivar la caché, esta parte habría que rehacerla desde el código que entonces se libere; el análisis del código no liberado queda fuera de la línea base.
