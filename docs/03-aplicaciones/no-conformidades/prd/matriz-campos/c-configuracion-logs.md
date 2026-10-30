# Matriz campo a campo: configuración, caché por NC, logs, documentos y catálogos (No Conformidades)

Parte «c-configuracion-logs» de la matriz campo a campo del PRD de No Conformidades (#1080). Cubre las tablas del esquema previo `TbCacheListadoNCAuditoria`, `TbCacheNCProyecto`, `TbConexiones`, `TbConfiguracion`, `TbConsultasPorFechas`, `TbDocumentosAuditorias`, `TbEstadoCatalogo`, `TbHerramientaDocAyuda`, `TbLog`, `TbLogAuditoria` y `TbLogCache`. Índice de formularios y capacidades en [el PRD de la aplicación](../README.md); las fichas de formulario están en el [inventario de formularios](../README.md#formularios).

## Método

La línea base es el código de `main` (producción), commit `1d7a2f826a1e3c582874e4e29886a6183efb9067`, directorios `src/classes`, `src/modules`, `src/forms` y `src/queries`. Los tipos de origen proceden del DDL de `ERD/schema.sql` (volcado de 2026-08-06), que solo declara tipos: no declara claves, relaciones ni nulabilidad, y no hay backend legible en esta fase. Por eso toda clave o relación se marca «inferida por código» con su cita, y se escribe «sin evidencia» cuando el código no la muestra. El backend de todas las filas es `NoConformidades_Datos.accdb`. Los destinos (`no_conformidades.<tabla>.<columna>`) y los tipos PostgreSQL son orientativos y quedan como propuesta hasta que se apruebe el ERD global.

Búsqueda realizada: coincidencia sobre todo el contenido de `src/classes`, `src/modules`, `src/forms` y `src/queries` (incluidos `queries.json` y los `.form.txt`) del nombre de cada tabla y, en las tablas sin uso, de sus columnas características. Resultado: `TbCacheListadoNCAuditoria`, `TbConexiones`, `TbConfiguracion`, `TbConsultasPorFechas` y `TbEstadoCatalogo` dan **0 coincidencias**, tampoco con el nombre sin el prefijo `Tb` ni con columnas propias (`AccionesCorrectivasConcatenadas`, `UltimaDesconexion`, `InstaladoFW`, `CacheHabilitada`, `MotivoCambioCache`, `TipoEntidad`, `EnumValor`, `UsuarioBootstrap`). No se detecta construcción dinámica de esos nombres. `TbHerramientaDocAyuda` solo aparece en una función sin llamadores. Las demás tablas se usan y se detallan abajo.

Precisiones sobre la configuración y la caché en `main`:

- La configuración de backend que conserva `main` se lee de otra tabla, `TbConfiguracionBackends` (`Variables Globales.bas:540-541`, vía `CurrentDb`), que no figura en el DDL del esquema previo; queda fuera de esta parte. `TbConfiguracion` (interruptor de caché) no es esa tabla y no se lee en `main`.
- La caché por NC (`TbCacheNCProyecto`) se conserva: `CacheNCProyecto.bas` e `InicializadorCache.bas`. Las cachés de listados e indicadores están retiradas (véase la parte `b-cache-indicadores`).
- El interruptor efectivo de la caché es la variable global `AplicarCache`, fijada a `False` en `Variables Globales.bas:286` y `Variables Globales.bas:328` (la asignación a `True` está comentada en `Variables Globales.bas:329`). Solo se consulta en `constructor.bas:441` y `NCProyectoOperaciones.cls:438`. Por tanto, en `main` la lectura de la caché y su regeneración tras guardar una NC no se ejecutan; en cambio, las invalidaciones (`CacheNCProyecto.InvalidarCache`) se invocan sin consultar esa variable en las líneas citadas, de modo que siguen escribiendo en `TbCacheNCProyecto` y `TbLogCache`.
- No existe transacción DAO en ninguna de estas tablas: la búsqueda de `BeginTrans`/`CommitTrans` en `src` solo encuentra nombres de funciones de la caché (`CacheNCProyecto.bas:646`), que gestionan un diccionario en memoria (`CacheNCProyecto.bas:633`) y no una transacción del motor.

Estado «propuesta»: el destino es candidato hasta que el propietario entregue el ERD de origen real.

## TbCacheListadoNCAuditoria

### Uso en `main`

- **Lectura y escritura**: ninguna. Búsqueda descrita en el apartado «Método»: 0 coincidencias de `TbCacheListadoNCAuditoria` en `src/classes`, `src/modules`, `src/forms` y `src/queries`. Es la caché de listados de NC de auditoría, retirada con el resto de cachés de listados e indicadores.
- **Transacción**: no aplica (sin uso).
- **Acceso**: ni DAO, ni SQL guardada, ni SQL dinámico concatenado.

### Claves y relaciones

- Clave primaria: sin evidencia.
- Relaciones: sin evidencia (el código de `main` no referencia la tabla; ninguna se infiere por código).

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | ID | LONG(4) | no_conformidades.cache_listado_nc_auditoria.id | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | IDAuditoria | LONG(4) | no_conformidades.cache_listado_nc_auditoria.id_auditoria | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | Tipo | TEXT(255) | no_conformidades.cache_listado_nc_auditoria.tipo | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | Numero | TEXT(255) | no_conformidades.cache_listado_nc_auditoria.numero | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | Descripcion | MEMO | no_conformidades.cache_listado_nc_auditoria.descripcion | text | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | CAUSARAIZ | MEMO | no_conformidades.cache_listado_nc_auditoria.causaraiz | text | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | RESPONSABLEIMPLANTACION | TEXT(255) | no_conformidades.cache_listado_nc_auditoria.responsableimplantacion | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | Estado | TEXT(255) | no_conformidades.cache_listado_nc_auditoria.estado | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | FechaApertura | SHORT_DATE_TIME(8) | no_conformidades.cache_listado_nc_auditoria.fecha_apertura | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | FECHACIERRE | SHORT_DATE_TIME(8) | no_conformidades.cache_listado_nc_auditoria.fechacierre | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | RequiereControlEficacia | TEXT(25) | no_conformidades.cache_listado_nc_auditoria.requiere_control_eficacia | varchar(25) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | ControlEficacia | MEMO | no_conformidades.cache_listado_nc_auditoria.control_eficacia | text | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | Notas | MEMO | no_conformidades.cache_listado_nc_auditoria.notas | text | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | Cerrada | TEXT(10) | no_conformidades.cache_listado_nc_auditoria.cerrada | varchar(10) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | Borrado | BOOLEAN(1) | no_conformidades.cache_listado_nc_auditoria.borrado | boolean | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | AccionesCorrectivasConcatenadas | MEMO | no_conformidades.cache_listado_nc_auditoria.acciones_correctivas_concatenadas | text | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | AccionesRealizadasConcatenadas | MEMO | no_conformidades.cache_listado_nc_auditoria.acciones_realizadas_concatenadas | text | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | FechaCache | SHORT_DATE_TIME(8) | no_conformidades.cache_listado_nc_auditoria.fecha_cache | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | CacheValida | BOOLEAN(1) | no_conformidades.cache_listado_nc_auditoria.cache_valida | boolean | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbCacheListadoNCAuditoria | Version | LONG(4) | no_conformidades.cache_listado_nc_auditoria.version | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |

## TbCacheNCProyecto

### Uso en `main`

- **Escritura (DAO)**: `GenerarCacheCompleto` abre un `Recordset` dinámico filtrado por NC y hace `AddNew` (versión 1) o `Edit` (versión + 1) (`CacheNCProyecto.bas:309-339`). Sin transacción DAO.
- **Escritura (SQL con parámetro)**: `InvalidarCache` ejecuta `UPDATE ... SET CacheValida=False` y `BorrarCache` ejecuta `DELETE` mediante un `QueryDef` temporal con parámetro, no una SQL guardada (`CacheNCProyecto.bas:383-386`, `CacheNCProyecto.bas:416-419`). Al revertir una transacción lógica de caché se vuelve a marcar `CacheValida=True` (`CacheNCProyecto.bas:717-719`).
- **Lectura (SQL dinámico concatenado)**: carga de la NC filtrando por `IDNoConformidad` y `CacheValida=True` (`CacheNCProyecto.bas:119-122`), existencia (`CacheNCProyecto.bas:211-214`) y validez (`CacheNCProyecto.bas:234-237`). La carga solo se invoca desde `constructor.bas:441-442` cuando `AplicarCache` es verdadero, es decir, nunca en `main` (véase «Método»).
- **Invalidaciones desde el dominio** (siempre ejecutadas): `NCProyecto.cls:1528`, `NCProyecto.cls:1562`, `NCProyecto.cls:1596`, `NCProyectoOperaciones.cls:882`, `NCProyectoOperaciones.cls:949`, `NCProyectoOperaciones.cls:1011`, `NCProyectoOperaciones.cls:1093`, `NCProyectoOperaciones.cls:1163`, `NCProyectoOperaciones.cls:1266`, `NCProyectoOperaciones.cls:1323`, `ACProyectoOperaciones.cls:238`, `ACProyectoOperaciones.cls:336`, `ARProyectoOperaciones.cls:301`, `ARProyectoOperaciones.cls:356`, `ReplanificacionesProyectoOperaciones.cls:31`, `ReplanificacionesProyectoOperaciones.cls:147`, `Funciones Generales.bas:2751` y, desde la selección de riesgos, `Form_formRiesgosSeleccion.cls:167` (ficha: [Form_formRiesgosSeleccion](../formularios/Form_formRiesgosSeleccion.md)).
- **Mantenimiento (SQL dinámico concatenado)**: `InicializadorCache.bas` contiene rutinas públicas de carga masiva, limpieza, verificación de integridad y estadísticas: borrado de cachés obsoletos (`InicializadorCache.bas:117-120`), comprobación de huérfanos y de JSON vacío (`InicializadorCache.bas:463-466`, `InicializadorCache.bas:513-516`), reparación (`InicializadorCache.bas:583-587`), invalidación y borrado por antigüedad (`InicializadorCache.bas:660-675`, `InicializadorCache.bas:703-706`) y estadísticas (`InicializadorCache.bas:734-737`, `InicializadorCache.bas:750-751`, `InicializadorCache.bas:766-767`). La búsqueda de esas rutinas en `src/classes`, `src/forms` y el resto de módulos no encuentra llamadores: se ejecutarían a mano (a verificar), y solo `TestCacheNC.bas:48-98` ejercita el módulo de caché.
- **Transacción**: ninguna del motor (véase «Método»).

### Claves y relaciones

- Una fila por NC: **inferida por código** (`CacheNCProyecto.bas:309-318`: si no hay fila para el `IDNoConformidad` se inserta, si la hay se edita). Unicidad declarada: sin evidencia.
- Clave primaria `IDCache`: **sin evidencia**; el código solo la lee (`InicializadorCache.bas:463`, `InicializadorCache.bas:797`) y no la asigna al insertar (`CacheNCProyecto.bas:317-339`).
- Relación con `TbNoConformidades` por `IDNoConformidad`: **inferida por código** (`InicializadorCache.bas:463-465`, `InicializadorCache.bas:489-491`). Integridad referencial declarada: sin evidencia.

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbCacheNCProyecto | IDNoConformidad | LONG(4) | no_conformidades.cache_nc_proyecto.id_no_conformidad | integer | entero tal cual; relación con la NC por `IDNoConformidad` inferida por código (`InicializadorCache.bas:489-491`) | no se admite nulo en la práctica: toda lectura filtra por él (`CacheNCProyecto.bas:119-120`); una fila por NC inferida por código (`CacheNCProyecto.bas:309-318`); nulabilidad declarada: sin evidencia | numérico antes de escribir (`CacheNCProyecto.bas:274-278`) | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbCacheNCProyecto | IDCache | LONG(4) | no_conformidades.cache_nc_proyecto.id_cache | integer | entero tal cual; en destino, identidad generada (propuesta); el código no la asigna (`CacheNCProyecto.bas:317-339`) | sin evidencia (si es autonumérico: a verificar con el ERD de origen) | sin evidencia | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbCacheNCProyecto | Version | INT(2) | no_conformidades.cache_nc_proyecto.version | smallint | entero; 1 al insertar y +1 en cada regeneración (`CacheNCProyecto.bas:319`, `CacheNCProyecto.bas:323`) | sin evidencia de nulabilidad; el código la lee sin `Nz` (`CacheNCProyecto.bas:323`) | sin evidencia de tope; el tipo de origen admite hasta 32767 regeneraciones (a verificar) | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbCacheNCProyecto | FechaCache | SHORT_DATE_TIME(8) | no_conformidades.cache_nc_proyecto.fecha_cache | timestamp | fecha y hora de la aplicación (`Now()`) en cada regeneración (`CacheNCProyecto.bas:327`); sin zona horaria declarada: a verificar | sin evidencia; sustituye a `FechaUltimoUso` cuando esta falta (`InicializadorCache.bas:660-663`) | sin evidencia | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbCacheNCProyecto | DatosNC | MEMO | no_conformidades.cache_nc_proyecto.datos_nc | text | texto largo con JSON generado por `GenerarJSONNC` (`CacheNCProyecto.bas:285`, `CacheNCProyecto.bas:332`); `jsonb` si se conserva, a verificar | se lee con `Nz(...,"")` (`CacheNCProyecto.bas:135`); nulo o vacío cuenta como caché inválida | la verificación de integridad rechaza nulo o vacío (`InicializadorCache.bas:513-516`) | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbCacheNCProyecto | DatosACs | MEMO | no_conformidades.cache_nc_proyecto.datos_a_cs | text | texto largo con JSON de acciones correctivas (`CacheNCProyecto.bas:288`, `CacheNCProyecto.bas:333`); `jsonb` si se conserva, a verificar | se lee con `Nz(...,"")` (`CacheNCProyecto.bas:146`); nulo o vacío cuenta como caché inválida | la verificación de integridad rechaza nulo o vacío (`InicializadorCache.bas:513-516`) | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbCacheNCProyecto | DatosARs | MEMO | no_conformidades.cache_nc_proyecto.datos_a_rs | text | texto largo con JSON de acciones realizadas (`CacheNCProyecto.bas:334`); `jsonb` si se conserva, a verificar | se lee con `Nz(...,"")` (`CacheNCProyecto.bas:157`); nulo o vacío cuenta como caché inválida | la verificación de integridad rechaza nulo o vacío (`InicializadorCache.bas:513-516`) | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbCacheNCProyecto | DatosReplanificaciones | MEMO | no_conformidades.cache_nc_proyecto.datos_replanificaciones | text | texto largo con JSON de replanificaciones (`CacheNCProyecto.bas:335`); `jsonb` si se conserva, a verificar | se lee con `Nz(...,"")` (`CacheNCProyecto.bas:169`); no entra en la verificación de integridad | sin evidencia de validación | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbCacheNCProyecto | DatosRiesgos | MEMO | no_conformidades.cache_nc_proyecto.datos_riesgos | text | texto largo con JSON de riesgos asociados (`CacheNCProyecto.bas:336`); `jsonb` si se conserva, a verificar | se lee con `Nz(...,"")` (`CacheNCProyecto.bas:180`); no entra en la verificación de integridad | sin evidencia de validación | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbCacheNCProyecto | UsuarioCache | TEXT(50) | no_conformidades.cache_nc_proyecto.usuario_cache | varchar(50) | usuario de red (`Environ("USERNAME")`, `CacheNCProyecto.bas:1280-1284`) recortado a 50 caracteres (`CacheNCProyecto.bas:281-282`); dato personal: tratamiento a decidir | sin evidencia | longitud máxima 50 aplicada por el código (`CacheNCProyecto.bas:282`) | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbCacheNCProyecto | CacheValida | BOOLEAN(1) | no_conformidades.cache_nc_proyecto.cache_valida | boolean | booleano de Access (verdadero/falso) a `boolean`; verdadero al regenerar (`CacheNCProyecto.bas:329`), falso al invalidar (`CacheNCProyecto.bas:384`), verdadero al revertir (`CacheNCProyecto.bas:717`) | sin evidencia de nulabilidad; la lectura exige verdadero (`CacheNCProyecto.bas:119-120`) | sin evidencia de validación adicional | propuesta: pendiente de ERD de origen y aprobación del ERD global |

**Columnas citadas en el código y ausentes del DDL del esquema previo** (a verificar contra el ERD de origen real: si no existen en el backend, las rutinas de mantenimiento que las usan fallan y el error se traga con `Debug.Print`, `InicializadorCache.bas:683-684`). Ninguna se asigna en la escritura de `GenerarCacheCompleto` (`CacheNCProyecto.bas:317-339`).

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbCacheNCProyecto | FechaUltimoUso | sin evidencia (no figura en el DDL) | no_conformidades.cache_nc_proyecto.fecha_ultimo_uso | timestamp | sin evidencia; se lee en `InicializadorCache.bas:117-118` y `InicializadorCache.bas:662-663` | se trata como posible nulo (`Is Null`, `InicializadorCache.bas:117-118`) | sin evidencia | a verificar: columna citada por el código y ausente del DDL |
| NoConformidades_Datos.accdb | TbCacheNCProyecto | HitsConsultas | sin evidencia (no figura en el DDL) | no_conformidades.cache_nc_proyecto.hits_consultas | integer | sin evidencia; se agrega en estadísticas (`InicializadorCache.bas:734-737`) | sin evidencia | sin evidencia | a verificar: columna citada por el código y ausente del DDL |
| NoConformidades_Datos.accdb | TbCacheNCProyecto | TamanioBytes | sin evidencia (no figura en el DDL) | no_conformidades.cache_nc_proyecto.tamanio_bytes | integer | sin evidencia; se lee en `InicializadorCache.bas:766-767` | sin evidencia | sin evidencia | a verificar: columna citada por el código y ausente del DDL |

**Decisión abierta**: la caché es un dato derivado y regenerable (`CacheNCProyecto.bas:254-339`); migrar sus filas o regenerarlas en destino es una decisión del propietario (a verificar), y con `AplicarCache` fijado a falso la caché no se consulta en `main`.

## TbConexiones

### Uso en `main`

- **Lectura y escritura**: ninguna. Búsqueda descrita en el apartado «Método»: 0 coincidencias de `TbConexiones` en `src/classes`, `src/modules`, `src/forms` y `src/queries`. Las únicas coincidencias parciales son la propiedad `FechaUltimaConexion` de la clase `Usuario` (`Usuario.cls:21`, `Usuario.cls:293`), que pertenece a la lista de campos de usuario y no a esta tabla (a verificar).
- **Transacción**: no aplica (sin uso).
- **Acceso**: ni DAO, ni SQL guardada, ni SQL dinámico concatenado.

### Claves y relaciones

- Clave primaria: sin evidencia.
- Relaciones: sin evidencia (el código de `main` no referencia la tabla; ninguna se infiere por código).

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbConexiones | Usuario | TEXT(255) | no_conformidades.conexiones.usuario | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbConexiones | UltimaConexion | SHORT_DATE_TIME(8) | no_conformidades.conexiones.ultima_conexion | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbConexiones | UltimaDesconexion | SHORT_DATE_TIME(8) | no_conformidades.conexiones.ultima_desconexion | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbConexiones | InstaladoFW3 | TEXT(2) | no_conformidades.conexiones.instalado_fw3 | varchar(2) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbConexiones | InstaladoFW4 | TEXT(2) | no_conformidades.conexiones.instalado_fw4 | varchar(2) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbConexiones | Exitoso | TEXT(2) | no_conformidades.conexiones.exitoso | varchar(2) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |

## TbConfiguracion

### Uso en `main`

- **Lectura y escritura**: ninguna. Búsqueda descrita en el apartado «Método»: 0 coincidencias de `TbConfiguracion` en `src/classes`, `src/modules`, `src/forms` y `src/queries`. Las coincidencias con el prefijo `TbConfiguracion` son `TbConfiguracionBackends` (`Variables Globales.bas:178`, `Variables Globales.bas:526`, `Variables Globales.bas:540`), otra tabla. Ninguna de las columnas de caché (`CacheHabilitada` y similares) se referencia: el interruptor real es la variable `AplicarCache` (`Variables Globales.bas:328`).
- **Transacción**: no aplica (sin uso).
- **Acceso**: ni DAO, ni SQL guardada, ni SQL dinámico concatenado.

### Claves y relaciones

- Clave primaria: sin evidencia.
- Relaciones: sin evidencia (el código de `main` no referencia la tabla; ninguna se infiere por código).

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbConfiguracion | ID | LONG(4) | no_conformidades.configuracion.id | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbConfiguracion | CacheHabilitada | BOOLEAN(1) | no_conformidades.configuracion.cache_habilitada | boolean | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbConfiguracion | FechaCambioCache | SHORT_DATE_TIME(8) | no_conformidades.configuracion.fecha_cambio_cache | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbConfiguracion | UsuarioCambioCache | TEXT(255) | no_conformidades.configuracion.usuario_cambio_cache | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbConfiguracion | MotivoCambioCache | MEMO | no_conformidades.configuracion.motivo_cambio_cache | text | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |

## TbConsultasPorFechas

### Uso en `main`

- **Lectura y escritura**: ninguna. Búsqueda descrita en el apartado «Método»: 0 coincidencias de `TbConsultasPorFechas` en `src/classes`, `src/modules`, `src/forms` y `src/queries`.
- **Transacción**: no aplica (sin uso).
- **Acceso**: ni DAO, ni SQL guardada, ni SQL dinámico concatenado.

### Claves y relaciones

- Clave primaria: sin evidencia.
- Relaciones: sin evidencia (el código de `main` no referencia la tabla; ninguna se infiere por código).

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbConsultasPorFechas | Descripcion | TEXT(255) | no_conformidades.consultas_por_fechas.descripcion | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |

## TbDocumentosAuditorias

### Uso en `main`

- **Escritura (DAO, alta)**: `Registrar` obtiene el identificador, abre la tabla completa y hace `AddNew` con los campos informados (`DocumentoAuditoriaOperaciones.cls:56-84`), y después copia el fichero al destino (`DocumentoAuditoriaOperaciones.cls:91`). Sin transacción: el alta de la fila y la copia del fichero no son atómicas.
- **Escritura (SQL dinámico concatenado)**: borrado por `IDDocumento` (`DocumentoAuditoriaOperaciones.cls:186-188`), seguido del borrado del fichero (`DocumentoAuditoriaOperaciones.cls:191-193`); cambio de nombre (`UPDATE ... SET Documento = '...'` con el texto concatenado sin duplicar comillas: a verificar, `DocumentoAuditoriaOperaciones.cls:238-240`).
- **Escritura (carga inicial)**: `RellenarAnexosAuditorias` crea filas a partir de `TbAnexosAuditoria` y `TbAnexosNCAuditorias` (`Instalador.bas:135-137`, `Instalador.bas:150-156`, `Instalador.bas:179-181`, `Instalador.bas:194`); su única llamadora es `InstalarProyecto` (`Instalador.bas:100`), que a su vez no tiene llamadores en `src` (búsqueda realizada).
- **Lectura (SQL dinámico concatenado)**: por `IDDocumento` o por `Documento` + `IDAuditoriaResultante` (`constructor.bas:587-593`); por `IDAuditoria`, `IDNoConformidad` o `IDAccionRealizada` (`constructor.bas:4922-4932`); por `IDNoConformidad` (`constructor.bas:4992-4994`); y por unión con `TbNCAuditoriaAccionesRealizadas` y `TbNCAuditoriaAccionCorrectivas` (`constructor.bas:5022-5027`). Comprobación de nombre repetido en `DocumentoAuditoriaOperaciones.cls:275-290`.
- **Consumidores**: `Auditoria.cls:181`, `NCAuditoria.cls:964`, `NCAuditoria.cls:993`, `ARAuditoria.cls:76`, `AuditoriaOperaciones.cls:160`, `AuditoriaOperaciones.cls:206`, `NCaUDITORIAOperaciones.cls:341`, `NCaUDITORIAOperaciones.cls:362`, `Form_FormAuditoriaDocumentos.cls:283` y `Form_FormNCAuditoriaDocumentos.cls:237`. Fichas: [Form_FormAuditoriaDocumentos](../formularios/Form_FormAuditoriaDocumentos.md) y [Form_FormNCAuditoriaDocumentos](../formularios/Form_FormNCAuditoriaDocumentos.md).
- **Transacción**: ninguna (véase «Método»).

### Claves y relaciones

- Clave `IDDocumento`: **inferida por código** como único (se calcula con máximo + 1, `DocumentoAuditoria.cls:158`, `constructor.bas:1165-1178`, y se consulta por él para borrar y renombrar, `DocumentoAuditoriaOperaciones.cls:186-188`). Unicidad declarada: sin evidencia. El máximo + 1 no es seguro con altas concurrentes (a verificar).
- `IDAuditoria`: **inferida por código** como relación con la auditoría (`constructor.bas:4922-4926`, `Auditoria.cls:181`).
- `IDNoConformidad`: **inferida por código** como relación con la NC de auditoría (`NCAuditoria.cls:964`, `constructor.bas:5022-5027` la compara con `TbNCAuditoriaAccionCorrectivas.ID`).
- `IDAccionRealizada`: **inferida por código** como relación con la acción realizada de auditoría (`ARAuditoria.cls:76`, `constructor.bas:5024-5026`).
- `IDAuditoriaResultante`: **inferida por código** como auditoría resultante (`constructor.bas:587-593`, `DocumentoAuditoria.cls:36-50`); su semántica frente a `IDAuditoria` es parcialmente solapada (`DocumentoAuditoriaOperaciones.cls:229` pasa la resultante como identificador de auditoría): a verificar.
- Regla de negocio de exclusividad: un documento debe pertenecer a una NC, una AR o una auditoría (`DocumentoAuditoriaOperaciones.cls:114-117`), y el nombre no puede repetirse dentro del mismo elemento (`DocumentoAuditoriaOperaciones.cls:253-290`). Integridad referencial declarada: sin evidencia.

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbDocumentosAuditorias | IDDocumento | LONG(4) | no_conformidades.documentos_auditorias.id_documento | integer | entero; en destino, identidad generada (propuesta) en lugar de máximo + 1 (`DocumentoAuditoria.cls:158`) | sin nulos en la práctica (siempre se asigna en el alta, `DocumentoAuditoriaOperaciones.cls:67`); sin duplicados por unicidad inferida | sin evidencia de validación adicional | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbDocumentosAuditorias | IDNoConformidad | LONG(4) | no_conformidades.documentos_auditorias.id_no_conformidad | integer | entero; solo se escribe si viene informado (`DocumentoAuditoriaOperaciones.cls:75-77`) | nulo admitido por código; el alta exige NC, AR o auditoría (`DocumentoAuditoriaOperaciones.cls:114-117`) | sin evidencia de existencia de la NC referenciada | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbDocumentosAuditorias | Documento | TEXT(255) | no_conformidades.documentos_auditorias.documento | varchar(255) | texto; nombre lógico del documento, sin extensión en la carga inicial (`Instalador.bas:154`) | no vacío (`DocumentoAuditoriaOperaciones.cls:120-122`); no se repite por elemento (`DocumentoAuditoriaOperaciones.cls:130-136`) | unicidad por elemento comprobada por código (`DocumentoAuditoriaOperaciones.cls:253-290`) | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbDocumentosAuditorias | NombreAnexo | TEXT(255) | no_conformidades.documentos_auditorias.nombre_anexo | varchar(255) | texto; nombre de fichero del anexo calculado al registrar (`DocumentoAuditoriaOperaciones.cls:52`, `DocumentoAuditoriaOperaciones.cls:61`) | se exige para componer la ruta del fichero (`DocumentoAuditoria.cls:132`) | sin evidencia de validación adicional; la ruta física del anexo no se almacena en la tabla | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbDocumentosAuditorias | IDAccionRealizada | LONG(4) | no_conformidades.documentos_auditorias.id_accion_realizada | integer | entero; solo si viene informado (`DocumentoAuditoriaOperaciones.cls:72-74`) | nulo admitido por código | sin evidencia de existencia de la AR referenciada | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbDocumentosAuditorias | IDAuditoria | LONG(4) | no_conformidades.documentos_auditorias.id_auditoria | integer | entero; solo si viene informado (`DocumentoAuditoriaOperaciones.cls:68-70`) | nulo admitido por código | sin evidencia de existencia de la auditoría referenciada | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbDocumentosAuditorias | IDAuditoriaResultante | LONG(4) | no_conformidades.documentos_auditorias.id_auditoria_resultante | integer | entero; solo si viene informado (`DocumentoAuditoriaOperaciones.cls:78-80`); la carga inicial no la rellena (línea comentada en `Instalador.bas:153`) | nulo admitido por código; las filas de la carga inicial no la informan | sin evidencia de existencia de la auditoría referenciada | propuesta: pendiente de ERD de origen y aprobación del ERD global |

## TbEstadoCatalogo

### Uso en `main`

- **Lectura y escritura**: ninguna. Búsqueda descrita en el apartado «Método»: 0 coincidencias de `TbEstadoCatalogo` en `src/classes`, `src/modules`, `src/forms` y `src/queries`.
- **Transacción**: no aplica (sin uso).
- **Acceso**: ni DAO, ni SQL guardada, ni SQL dinámico concatenado.

### Claves y relaciones

- Clave primaria: sin evidencia.
- Relaciones: sin evidencia (el código de `main` no referencia la tabla; ninguna se infiere por código).

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbEstadoCatalogo | TipoEntidad | TEXT(10) | no_conformidades.estado_catalogo.tipo_entidad | varchar(10) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbEstadoCatalogo | EnumValor | LONG(4) | no_conformidades.estado_catalogo.enum_valor | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbEstadoCatalogo | Codigo | TEXT(100) | no_conformidades.estado_catalogo.codigo | varchar(100) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbEstadoCatalogo | Titulo | TEXT(255) | no_conformidades.estado_catalogo.titulo | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbEstadoCatalogo | Texto | TEXT(100) | no_conformidades.estado_catalogo.texto | varchar(100) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbEstadoCatalogo | Version | LONG(4) | no_conformidades.estado_catalogo.version | integer | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbEstadoCatalogo | Activo | BOOLEAN(1) | no_conformidades.estado_catalogo.activo | boolean | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbEstadoCatalogo | FechaBootstrap | SHORT_DATE_TIME(8) | no_conformidades.estado_catalogo.fecha_bootstrap | timestamp | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbEstadoCatalogo | UsuarioBootstrap | TEXT(255) | no_conformidades.estado_catalogo.usuario_bootstrap | varchar(255) | sin evidencia | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |

## TbHerramientaDocAyuda

### Uso en `main`

- **Lectura (DAO, tabla completa)**: `getDocumentosAyuda` abre la tabla por nombre y construye un diccionario `NombreFormulario` → `NombreArchivoAyuda`, insensible a mayúsculas (`constructor.bas:257-258`, `constructor.bas:267-272`).
- **Escritura**: ninguna en `src` (búsqueda del nombre de la tabla); la carga de datos es externa al código (a verificar).
- **Llamadores**: ninguno. `getDocumentosAyuda` solo aparece en su definición (`constructor.bas:249-285`) en `src/classes`, `src/forms` y `src/modules`. La ayuda contextual que sí existe compone el nombre del PDF con el nombre del formulario activo (`Funciones Generales.bas:745-746`), sin consultar la tabla (a verificar).
- **Transacción**: no aplica.
- **Acceso**: DAO sobre la tabla.

### Claves y relaciones

- Clave por `NombreFormulario`: **inferida por código** (si ya existe, la primera fila gana y las repetidas se ignoran, `constructor.bas:271-273`). Unicidad declarada: sin evidencia.
- Relaciones: sin evidencia (los nombres de formulario no se enlazan a otra tabla).

### Columnas

| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbHerramientaDocAyuda | NombreFormulario | TEXT(255) | no_conformidades.herramienta_doc_ayuda.nombre_formulario | varchar(255) | texto tal cual; clave del diccionario (`constructor.bas:271-272`) | se lee sin `Nz` (`constructor.bas:271`): un nulo rompería la lectura; duplicados: se ignoran los posteriores (`constructor.bas:271`) | comparación insensible a mayúsculas (`constructor.bas:269`) | pendiente: uso solo en función sin llamadores en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbHerramientaDocAyuda | NombreArchivoAyuda | TEXT(255) | no_conformidades.herramienta_doc_ayuda.nombre_archivo_ayuda | varchar(255) | texto tal cual; valor del diccionario (`constructor.bas:272`) | se lee sin `Nz` (`constructor.bas:272`); sin evidencia de nulabilidad declarada | sin evidencia de validación de existencia del fichero de ayuda | pendiente: uso solo en función sin llamadores en main, decisión del propietario |

## TbLog

### Uso en `main`

- **Escritura (DAO)**: `LogNCProyecto.Alta` calcula `IDLog` con máximo + 1, comprueba que no exista (`SELECT * ... WHERE IDLog=`) y hace `AddNew` + `Update` (`LogNCProyecto.cls:186-241`, `constructor.bas:1165-1178`). Sin transacción; el máximo + 1 no es seguro con altas concurrentes (a verificar).
- **Quién escribe**: `RegistrarLogProyecto` (`Funciones Generales.bas:993-1095`), invocado desde numerosas clases de operaciones, p. ej. `ReplanificacionesProyectoOperaciones.cls:33`, `ReplanificacionesProyectoOperaciones.cls:149`, `RiesgoServicio.cls:47` y `DocumentoProyectoOperaciones.cls:54`; y las clases `NCProyectoOperaciones.cls:952` y `NCProyectoOperaciones.cls:1014`, que instancian `LogNCProyecto`. `NCaUDITORIAOperaciones.cls:459` también instancia `LogNCProyecto` dentro de operaciones de auditoría: a verificar si escribe en `TbLog`.
- **Lectura (SQL dinámico concatenado)**: `SELECT TbLog.* ... WHERE TbLog.<IDNC|IDAC|IDAR> = <valor> ORDER BY Fecha` (`constructor.bas:1219-1222`); se muestra como HTML (`Funciones Generales.bas:1573`).
- **Transacción**: ninguna (véase «Método»).

### Claves y relaciones

- Clave `IDLog`: **inferida por código** como única (se comprueba antes del alta, `LogNCProyecto.cls:214-221`; el diccionario de lectura se indexa por ella, `constructor.bas:1244-1245`). Unicidad declarada: sin evidencia.
- `IDNC`, `IDAC`, `IDAR`: **inferidas por código** como relaciones con la NC, la acción correctiva y la acción realizada de proyecto, según el tipo del objeto (`constructor.bas:1209-1217`). Integridad referencial declarada: sin evidencia.

### Columnas
| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbLog | IDLog | LONG(4) | no_conformidades.log.id_log | integer | entero; en destino, identidad generada (propuesta) en lugar de máximo + 1 (`LogNCProyecto.cls:137`, `constructor.bas:1165-1178`) | siempre asignado; duplicado rechazado por el alta (`LogNCProyecto.cls:214-221`) | sin evidencia de validación adicional | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLog | IDNC | LONG(4) | no_conformidades.log.id_nc | integer | entero; solo se escribe si es numérico (`LogNCProyecto.cls:232-234`) | nulo admitido por código; en lectura el nulo pasa a cadena vacía (`constructor.bas:1233-1235`) | sin evidencia de existencia de la NC referenciada | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLog | IDAC | LONG(4) | no_conformidades.log.id_ac | integer | entero; solo se escribe si es numérico (`LogNCProyecto.cls:235-237`) | nulo admitido por código (una fila referencia NC, AC o AR, según el objeto, `constructor.bas:1209-1217`) | sin evidencia de existencia de la AC referenciada | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLog | IDAR | LONG(4) | no_conformidades.log.id_ar | integer | entero; solo se escribe si es numérico (`LogNCProyecto.cls:238-240`) | nulo admitido por código | sin evidencia de existencia de la AR referenciada | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLog | Usuario | TEXT(255) | no_conformidades.log.usuario | varchar(255) | usuario de red de la sesión (`Funciones Generales.bas:1022`, `Funciones Generales.bas:1084`); dato personal: tratamiento a decidir | sin evidencia; siempre informado por el registro | sin evidencia de longitud máxima (origen 255) | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLog | Fecha | SHORT_DATE_TIME(8) | no_conformidades.log.fecha | timestamp | fecha y hora de la aplicación (`Now()`, `LogNCProyecto.cls:202`) a `timestamp`; en destino `timestamptz` con valor del servidor, a verificar | siempre asignada en el alta; lectura ordenada por ella (`constructor.bas:1222`) | sin evidencia de validación adicional | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLog | Titulo | MEMO | no_conformidades.log.titulo | text | texto largo tal cual (`LogNCProyecto.cls:225`) | obligatorio: el alta se aborta si está vacío (`LogNCProyecto.cls:196-199`) | no vacío (`LogNCProyecto.cls:196-199`) | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLog | Linea | MEMO | no_conformidades.log.linea | text | texto largo; une la línea libre y las entradas de cambios separadas por salto de línea (`LogNCProyecto.cls:161-175`); formato «Campo_Inicial:valor|Campo_Final:valor» (`Funciones Generales.bas:1066-1070`) | solo se escribe si no está vacía (`LogNCProyecto.cls:226-228`): nulo admitido | sin evidencia de validación; puede contener valores de campos de negocio: tratar como dato sensible al migrar | propuesta: pendiente de ERD de origen y aprobación del ERD global |

## TbLogAuditoria

### Uso en `main`

- **Escritura (DAO)**: `LogNCAuditoria.Alta`, con la misma mecánica que `TbLog` (`LogNCAuditoria.cls:186-241`, `LogNCAuditoria.cls:137`). Sin transacción.
- **Quién escribe**: `RegistrarLogAuditoria` (`Funciones Generales.bas:1098-1190`) y `NCaUDITORIAOperaciones.cls:400`, que instancia `LogNCAuditoria`.
- **Lectura (SQL dinámico concatenado)**: `SELECT * FROM TbLogAuditoria WHERE TbLogAuditoria.<IDNC|IDAC|IDAR> = <valor> ORDER BY Fecha` (`constructor.bas:1295-1298`); se muestra como HTML (`Funciones Generales.bas:1644`).
- **Transacción**: ninguna.

### Claves y relaciones

- Clave `IDLog`: **inferida por código** como única (`LogNCAuditoria.cls:214-221`). Unicidad declarada: sin evidencia.
- `IDNC`, `IDAC`, `IDAR`: **inferidas por código** como relaciones con la NC, la acción correctiva y la acción realizada de auditoría (`constructor.bas:1285-1294`). Integridad referencial declarada: sin evidencia.

### Columnas
| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbLogAuditoria | IDLog | LONG(4) | no_conformidades.log_auditoria.id_log | integer | entero; en destino, identidad generada (propuesta) en lugar de máximo + 1 (`LogNCAuditoria.cls:137`, `constructor.bas:1165-1178`) | siempre asignado; duplicado rechazado por el alta (`LogNCAuditoria.cls:214-221`) | sin evidencia de validación adicional | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLogAuditoria | IDNC | LONG(4) | no_conformidades.log_auditoria.id_nc | integer | entero; solo se escribe si es numérico (`LogNCAuditoria.cls:232-234`) | nulo admitido por código; en lectura el nulo pasa a cadena vacía (`constructor.bas:1309-1311`) | sin evidencia de existencia de la NC referenciada | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLogAuditoria | IDAC | LONG(4) | no_conformidades.log_auditoria.id_ac | integer | entero; solo se escribe si es numérico (`LogNCAuditoria.cls:235-237`) | nulo admitido por código (una fila referencia NC, AC o AR, según el objeto, `constructor.bas:1285-1294`) | sin evidencia de existencia de la AC referenciada | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLogAuditoria | IDAR | LONG(4) | no_conformidades.log_auditoria.id_ar | integer | entero; solo se escribe si es numérico (`LogNCAuditoria.cls:238-240`) | nulo admitido por código | sin evidencia de existencia de la AR referenciada | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLogAuditoria | Usuario | TEXT(255) | no_conformidades.log_auditoria.usuario | varchar(255) | usuario de red de la sesión (`Funciones Generales.bas:1127`, `Funciones Generales.bas:1188`); dato personal: tratamiento a decidir | sin evidencia; siempre informado por el registro | sin evidencia de longitud máxima (origen 255) | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLogAuditoria | Fecha | SHORT_DATE_TIME(8) | no_conformidades.log_auditoria.fecha | timestamp | fecha y hora de la aplicación (`Now()`, `LogNCAuditoria.cls:202`) a `timestamp`; en destino `timestamptz` con valor del servidor, a verificar | siempre asignada en el alta; lectura ordenada por ella (`constructor.bas:1298`) | sin evidencia de validación adicional | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLogAuditoria | Titulo | MEMO | no_conformidades.log_auditoria.titulo | text | texto largo tal cual (`LogNCAuditoria.cls:225`) | obligatorio: el alta se aborta si está vacío (`LogNCAuditoria.cls:196-199`) | no vacío (`LogNCAuditoria.cls:196-199`) | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLogAuditoria | Linea | MEMO | no_conformidades.log_auditoria.linea | text | texto largo; une la línea libre y las entradas de cambios separadas por salto de línea (`LogNCAuditoria.cls:161-175`); formato «Campo_Inicial:valor|Campo_Final:valor» (`Funciones Generales.bas:1066-1070`) | solo se escribe si no está vacía (`LogNCAuditoria.cls:226-228`): nulo admitido | sin evidencia de validación; puede contener valores de campos de negocio: tratar como dato sensible al migrar | propuesta: pendiente de ERD de origen y aprobación del ERD global |

## TbLogCache

### Uso en `main`

- **Escritura (SQL dinámico concatenado)**: `LogCacheOperacion` inserta una fila por operación de caché con `getdb().Execute`, con `On Error Resume Next` (`CacheNCProyecto.bas:1286-1304`). Lo invocan la carga, generación, invalidación, borrado y reversión de la caché (`CacheNCProyecto.bas:69`, `CacheNCProyecto.bas:348`, `CacheNCProyecto.bas:390`, `CacheNCProyecto.bas:423`, `CacheNCProyecto.bas:725`), de modo que cada invalidación desde el dominio escribe aquí aunque `AplicarCache` sea falso (véase «Método»).
- **Borrado (SQL dinámico concatenado)**: `LimpiarLogsAntiguos` purga por antigüedad general y por errores con `dbFailOnError` (`InicializadorCache.bas:977-985`); sin llamadores en `src` (búsqueda realizada).
- **Lectura**: solo `DCount` de recuento en la limpieza (`InicializadorCache.bas:977`, `InicializadorCache.bas:985`).
- **Transacción**: ninguna (véase «Método»).

### Claves y relaciones

- Clave primaria `IDLog`: **sin evidencia** (el `INSERT` no la informa, `CacheNCProyecto.bas:1299-1304`).
- Relación con la NC por `IDNoConformidad`: **inferida por código** (`CacheNCProyecto.bas:1300-1301`); integridad referencial declarada: sin evidencia.

### Columnas
| Backend | Tabla | Columna | Tipo origen | Destino (no_conformidades.<tabla>.<columna>, propuesta) | Tipo destino | Transformación | Nulos y duplicados | Validación | Estado |
|---|---|---|---|---|---|---|---|---|---|
| NoConformidades_Datos.accdb | TbLogCache | IDLog | LONG(4) | no_conformidades.log_cache.id_log | integer | entero; el `INSERT` no lo informa (`CacheNCProyecto.bas:1299-1304`): si es autonumérico, a verificar; en destino, identidad generada (propuesta) | sin evidencia (si es autonumérico: a verificar con el ERD de origen) | sin evidencia | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLogCache | IDNoConformidad | LONG(4) | no_conformidades.log_cache.id_no_conformidad | integer | entero; se concatena sin comillas en el `INSERT` (`CacheNCProyecto.bas:1301`) | sin evidencia; la inserción ignora errores (`On Error Resume Next`, `CacheNCProyecto.bas:1297`), por lo que una fila fallida se pierde sin aviso | sin evidencia de existencia de la NC referenciada | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLogCache | TipoOperacion | TEXT(50) | no_conformidades.log_cache.tipo_operacion | varchar(50) | texto; valores escritos: `Generar`, `Error`, `Consultar`, `Generar Completo`, `Invalidar`, `Borrar`, `Rollback` (`CacheNCProyecto.bas:69`, `CacheNCProyecto.bas:71`, `CacheNCProyecto.bas:89`, `CacheNCProyecto.bas:348`, `CacheNCProyecto.bas:390`, `CacheNCProyecto.bas:423`, `CacheNCProyecto.bas:725`) | sin evidencia | la limpieza trata `Error` de forma especial (`InicializadorCache.bas:983`); sin lista cerrada declarada | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLogCache | SeccionCache | TEXT(50) | no_conformidades.log_cache.seccion_cache | varchar(50) | sin evidencia: el código no la escribe (`CacheNCProyecto.bas:1299-1304`) ni la lee | sin evidencia | sin evidencia | pendiente: sin uso en main, decisión del propietario |
| NoConformidades_Datos.accdb | TbLogCache | Detalles | MEMO | no_conformidades.log_cache.detalles | text | texto largo; el apóstrofo se duplica antes de insertar (`CacheNCProyecto.bas:1301`); en `Error` contiene el mensaje de error | sin evidencia | sin evidencia; puede contener mensajes técnicos: no migrar sin revisión | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLogCache | FechaOperacion | SHORT_DATE_TIME(8) | no_conformidades.log_cache.fecha_operacion | timestamp | fecha y hora de la aplicación (`Now()`, `CacheNCProyecto.bas:1302`) | siempre asignada por el `INSERT` | base de la purga por antigüedad (`InicializadorCache.bas:980`, `InicializadorCache.bas:983`) | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLogCache | Usuario | TEXT(50) | no_conformidades.log_cache.usuario | varchar(50) | texto; `Sistema` o usuario de red (`CacheNCProyecto.bas:69`, `CacheNCProyecto.bas:390`); dato personal: tratamiento a decidir | siempre informado | el origen declara 50 caracteres y el `INSERT` no recorta (`CacheNCProyecto.bas:1302`): a verificar | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLogCache | DuracionMs | LONG(4) | no_conformidades.log_cache.duracion_ms | integer | entero; 0 por defecto (`CacheNCProyecto.bas:1292`) | siempre informado (0 si no se mide) | sin evidencia de validación | propuesta: pendiente de ERD de origen y aprobación del ERD global |
| NoConformidades_Datos.accdb | TbLogCache | Exito | BOOLEAN(1) | no_conformidades.log_cache.exito | boolean | booleano de Access a `boolean`; se escribe literal `True`/`False` (`CacheNCProyecto.bas:1302`) | siempre informado | sin evidencia de validación adicional | propuesta: pendiente de ERD de origen y aprobación del ERD global |

## Pendientes

- ERD de origen real (claves, relaciones, nulabilidad, autonumeración de `IDCache` e `IDLog` de `TbLogCache`): pendiente del propietario.
- Confirmar si `FechaUltimoUso`, `HitsConsultas` y `TamanioBytes` existen en el backend de producción (el DDL del esquema previo no las trae).
- Decidir con el propietario el destino de `TbCacheListadoNCAuditoria`, `TbConexiones`, `TbConfiguracion`, `TbConsultasPorFechas` y `TbEstadoCatalogo` (sin uso en `main`) y de `TbHerramientaDocAyuda` (función sin llamadores).
- Decidir si se migran los datos de `TbCacheNCProyecto` y `TbLogCache` (derivados y regenerables) o solo la estructura.
- Verificar si `NCaUDITORIAOperaciones.cls:459` escribe en `TbLog` desde operaciones de auditoría.
- Tratamiento de los campos `Usuario`, `UsuarioCache` y `Linea`/`Detalles` (usuarios de red y texto libre) al migrar.
