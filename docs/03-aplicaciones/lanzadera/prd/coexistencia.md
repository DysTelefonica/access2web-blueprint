# Coexistencia Access y web — Lanzadera

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Issue:** #1074 · **Fase:** 4 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Qué debe adaptarse, preservarse o decidirse para que Lanzadera opere sobre tablas PostgreSQL mientras la web la sustituye. Todo es lectura estática de `src` (los recuentos excluyen los módulos `Test*`): no se ha ejecutado nada ni comprobado compatibilidad real con PostgreSQL. Las propuestas no son requisitos hasta que se aprueben. La arquitectura de datos de la coexistencia (P-EXP-3) está abierta y en estudio en el PRD de Expedientes (`docs/03-aplicaciones/expedientes/prd/coexistencia.md`, apartados 11 y 12, en la cadena de #1073); este documento mide lo mismo para Lanzadera sin recomendar ninguna vía.

## 1. Cómo accede hoy Lanzadera a los datos

| Hecho | Evidencia |
|---|---|
| El acceso a datos pasa por `getdb()`: 89 llamadas en 13 ficheros. Abre el backend por ruta con `DBEngine.Workspaces(0).OpenDatabase(ruta, False, False, "MS Access;PWD=…")` y cachea la conexión. | `Variables Globales.bas:253, 309-310` |
| La ruta y la clave del backend salen de la tabla local `TbConfiguracionBackends`, leída con `CurrentDb` (`TOP 1`); si la clave está vacía se toma de una variable de entorno. | `Variables Globales.bas:368-391` |
| No hay cadenas ODBC ni `.Connect` ni consultas paso-a-través. | 0 apariciones en `src` |

## 2. Reglas del propietario, verificadas en Lanzadera

- **Tabla local única: `TbConfiguracionBackends`** (se queda en Access, decisión del propietario). Columnas que lee el código: `BackendActivo`, `BackendProduccion`, `BackendSandbox`, la columna de clave, `IDAplicacion` y `EnPruebas` (`Variables Globales.bas:368-426`). Excepción respecto a Expedientes: aquí **el código de producción también la escribe**: `Config_BackendHelper.bas:28-91` (`ConfigurarTbConfiguracionBackends`, con `AddNew`/`Edit` sobre `CurrentDb`). Ningún módulo ajeno a las pruebas la invoca; quién la usa en la práctica queda pendiente.
- **Ninguna otra tabla local.** `CurrentDb` aparece solo en esas dos lecturas y escrituras de configuración (`Variables Globales.bas:379`, `Config_BackendHelper.bas:57`). El resto de usos de datos recibe una base de datos de `getdb()`. Que el frontend vincule o no otras tablas no es comprobable desde `src` (el frontend es binario).
- **Formularios desconectados: confirmado.** 0 `RecordSource` y 0 `ControlSource` en los 28 `.form.txt`; los 19 listados y combos con tipo de origen declarado son `Value List`.
- **Tablas temporales o de configuración por puesto** (`TbUsuarioConfiguracion`, `TbConfiguracion`, `TbTablasAVincular`) están en el backend del snapshot (`ERD/schema.sql`); con la regla del propietario vivirían en el backend, no en el frontend. Su disposición queda por decidir.

## 3. Lo específico de Lanzadera: gestiona los frontends de las demás aplicaciones

Además del acceso a datos, Lanzadera abre **ficheros `.accdb` de otras aplicaciones** en modo lectura y escritura con su clave para ajustar propiedades (`AlterarPropriedad`, p. ej. el icono): `Aplicacion.cls:1432-1458`, `Funciones Genreales.bas:1946, 2093, 2242, 2961`. También copia ejecutables, iconos y ficheros `.ini` entre la ubicación remota y la local (`Aplicacion.cls:1425-1431`) y lanza procesos del sistema (`Funciones Genreales.bas:165, 346, 398`). Eso es despliegue y lanzamiento de escritorio, no acceso a datos: no pasa por PostgreSQL y qué significa «lanzar» en la web es una decisión de producto (pendiente).

## 4. Qué se adapta y qué se preserva

| Construcción | Ocurrencias | Primer ejemplo | Riesgo con PostgreSQL |
|---|---|---|---|
| SQL por concatenación de cadenas | 76 líneas en 11 ficheros | `Form_FormAplicacionesGestion.cls:142` | Escapado de comillas (hallazgo de apóstrofo en `Constructor.bas:922`); parametrizar |
| `DELETE *` | 1 | `Visionado.cls:287` | No es SQL estándar |
| Literales de fecha con `#` y `Format(…, "mm/dd/yyyy")` | 5 líneas en 3 ficheros | `Constructor.bas:174, 255, 263, 268`; `Usuario.cls:688` | Literal de Access, orden mes/día |
| `Date()` en SQL | 2 | `Constructor.bas:335`; `Funciones Genreales.bas:3162` | Función de Access |
| `LIKE '*texto*'` | 2 | `Constructor.bas:846, 922` | Comodín `*` frente a `%` |
| `SELECT TOP n` | 2, ambas sobre la tabla local | `Variables Globales.bas:380`; `Config_BackendHelper.bas:58` | Tabla local: no migra |
| `LEFT JOIN` | 13 líneas en 1 fichero | `Constructor.bas:189` | Paréntesis anidados de ACE |
| Texto `Sí`/`No` | 183 líneas en 13 ficheros | `Form_FormPerfil.cls:46` | Conservar como texto o decidir `BOOLEAN` |
| `Option Compare Database` | 57 de 59 ficheros de código | — | Comparación de texto insensible a mayúsculas |

Se preserva el comportamiento observable: valores `Sí`/`No`, orden de listados, mensajes de validación y el flujo de apertura de aplicaciones.

## 5. Recordsets, identidad, transacciones y tipos

- **Recordsets:** `OpenRecordset` 59 líneas en 13 ficheros; `AddNew` 15, `Edit` 10, `Update` 24 y `Execute` 14. Sin `dbOpenTable`, `.Seek`, `.Index`, `FindFirst`, `.Bookmark`, `.Filter` ni `.Sort`. Sin `@@IDENTITY` ni `.LastModified`. Con vínculos ODBC, un recordset actualizable exige clave única declarada; el `schema.sql` no contiene claves (pendiente del ERD, D-LAN-3).
- **Identificadores:** `DameID` (`Max + 1`, `Funciones Genreales.bas:473`), 12 llamadas en 7 ficheros. Colisionaría con identificadores generados por la web.
- **Lectura de estructura:** 9 clases leen `getdb().TableDefs("TbX").Fields` para enumerar columnas (p. ej. `Aplicacion.cls:2093`, `Usuario.cls:1972`): la enumeración debe devolver las mismas columnas que hoy.
- **Transacciones: ninguna.** 0 `BeginTrans`, `CommitTrans` o `Rollback`. Cada escritura es independiente; ninguna operación de varias tablas es atómica hoy.
- **Tipos:** 0 columnas `BOOLEAN` en el `schema.sql` de 35 tablas y 230 columnas. Conviene revisar las fechas (`SHORT_DATE_TIME`) y las columnas `MEMO` con la matriz.

## 6. Identidad, permisos y credenciales

- **Lanzadera es propietaria de la identidad que usan otras aplicaciones**: Expedientes, HPS, Condor y No Conformidades vinculan `TbUsuariosAplicaciones` y `TbUsuariosAplicacionesPermisos` de este backend, y Gestión de Riesgos también `TbAplicaciones` (ver la [matriz](matriz-campos.md#tablas-de-lanzadera-que-consumen-otras-aplicaciones)). Un cambio de estructura o de claves en esas tablas afecta a cinco de las ocho aplicaciones durante toda la coexistencia.
- **Credenciales:** la contraseña se procesa con la API criptográfica de Windows (`Critografia.bas:12-31`) y existen columnas de contraseña, de clave inicial y de histórico (nombres en `ERD/schema.sql`). El esquema web ya tiene `password_hash` y `reset_tokens` (migración `0001`). La estrategia de migración o de reinicio de credenciales es una decisión del propietario y no se documenta aquí ningún valor ni parámetro.
- **Defectos de seguridad** del flujo de autenticación: ver los hallazgos H-LAN-1 a H-LAN-4, H-LAN-9, H-LAN-14 y H-LAN-16 del [README](README.md#hallazgos-del-legacy) (acceso sin contraseña para administradores de Windows, bloqueo que no caduca, contraseña de un solo uso en claro, cambio de contraseña sin pedir la actual, contraseña de aplicación sin máscara, alta de usuario sin clave inicial).
- **Presencia y conexiones:** `TbConexiones` y `TbConexionesRegistro` (datos de red del puesto) tienen correspondencia candidata con `presence_sessions` de la web (migración `0008`), sin asignar columna a columna.

## 7. Concurrencia y auditoría

Sin bloqueo explícito (0 usos de `dbPessimistic` u `dbOptimistic`) y sin versión de fila: el último que guarda gana. No hay registro de último cambio equivalente al de otras aplicaciones: 0 usos de `RegistrarUltimoCambio` ni de `TbCambios`; la trazabilidad la dan las tablas de aperturas, ediciones y conexiones. La web tiene `audit`. Qué debe auditarse en la coexistencia queda por definir.

## 8. Pruebas de aceptación bidireccionales (propuestas)

A ejecutar después, en entorno aislado y nunca sobre producción:

1. Un usuario creado en la web inicia sesión en Lanzadera de Access, y al revés, con el mismo resultado de permisos.
2. Un permiso cambiado en un lado se ve en el otro y en una aplicación que vincule `TbUsuariosAplicacionesPermisos`.
3. Un bloqueo por intentos fallidos se aplica igual en ambos lados.
4. Un cambio de contraseña en un lado no deja al otro con una credencial obsoleta.
5. Dos altas simultáneas (una por interfaz) no obtienen el mismo identificador.
6. Alta y baja de una aplicación en el catálogo se reflejan en ambos lados y en el listado de aperturas.
7. Una caída de conexión a mitad de una escritura no deja datos a medias (hoy sin transacciones).

## 9. Corte de la aplicación

Lanzadera es la entrada de las demás aplicaciones: al activar la web de otra aplicación se deshabilita **su** botón en esta Lanzadera. El corte de la propia Lanzadera es el último y depende de que las otras cuatro aplicaciones que vinculan sus tablas de identidad hayan cambiado de origen. Deshabilitar la entrada no borra binarios ni datos, ni cierra sesiones abiertas o accesos directos. La reversión está por definir.

## Pendientes

- Definir qué significa «lanzar» una aplicación en la web y qué hace con los frontends de escritorio (apartado 3).
- Extraer claves, índices y relaciones reales del backend (D-LAN-3, pendiente del propietario).
- Decidir la estrategia de credenciales y de identificadores compartidos con las demás aplicaciones.
- Trazar quién invoca `ConfigurarTbConfiguracionBackends` fuera de las pruebas.
- Comprobar el comportamiento de `TableDefs` y de los recordsets actualizables sobre tablas vinculadas.
