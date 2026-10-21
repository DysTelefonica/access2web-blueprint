# Coexistencia Access y web — Expedientes

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Issue:** #1073 · **Fase:** 4 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Qué debe adaptarse, preservarse o decidirse para que la versión de Access de Expedientes, durante la transición, opere sobre tablas PostgreSQL vinculadas que la web también usa. Todo lo que sigue es lectura estática de `src` (los recuentos excluyen los módulos `Test*`); no se ha ejecutado nada ni se ha comprobado compatibilidad real con PostgreSQL. Las propuestas no son requisitos hasta que se aprueben.

## 1. Cómo accede hoy Expedientes a los datos

| Hecho | Evidencia |
|---|---|
| El acceso a datos pasa casi siempre por `getdb()`: 303 llamadas en 62 ficheros. | `Variables Globales.bas:607` |
| `getdb()` abre el fichero del backend con `DBEngine.Workspaces(0).OpenDatabase(ruta, False, False, ";pwd=…")` y cachea la conexión. No usa tablas vinculadas del frontend. | `Variables Globales.bas:642-645` |
| `CurrentDb` (las tablas vinculadas del propio frontend) aparece solo 5 veces en 4 ficheros. | `FUNCIONES UTILES.bas:2849` |
| No hay cadenas `ODBC;` ni `.Connect` en `src`. | búsqueda en `src` (0 apariciones) |
| La ruta y la clave del backend salen de una tabla local del frontend, `TbConfiguracionBackends`, con tres destinos (producción, sandbox y test) resueltos por `BackendResolver`. | `Variables Globales.bas:521`, `BackendResolver.bas:1-40`, `ModFixtureF45CacheE2E.bas:43` |

**Consecuencia.** El modelo de vínculos del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) (Access sobre tablas PostgreSQL vinculadas por ODBC) puede realizarse de más de una forma y la elección es la P-EXP-3, **abierta y escalada al propietario**. Este documento no la resuelve ni recomienda: (a) sustituir `getdb()` para que devuelva `CurrentDb` con vínculos en el frontend; (b) mantener un backend `.accdb` sincronizado con PostgreSQL, que el roadmap descarta; (c) la hipótesis del propietario, que mantiene `getdb()` y vincula las tablas por ODBC dentro del backend. La vía (c) se mide en el [apartado 11](#11-p-exp-3-verificación-de-la-hipótesis-del-propietario).

## 2. Qué se adapta y qué se preserva

### SQL en dialecto Jet/ACE

| Construcción | Ocurrencias | Primer ejemplo | Riesgo con PostgreSQL |
|---|---|---|---|
| SQL montado por concatenación de cadenas | 391 líneas en 43 ficheros | `Form_FormE2EGestionEnlaces.cls:25` | Inyección y escapado de comillas (varios hallazgos de apóstrofo en las fichas); parametrizar |
| `DELETE *` | 49 en 15 ficheros | `CPVOperaciones.cls:120` | No es SQL estándar; `DELETE FROM` |
| Literales de fecha con `#` | 11 en 5 ficheros | `ExpedienteOperaciones.cls:1789` | PostgreSQL usa literales entrecomillados; ODBC reescribe a veces, no siempre |
| `LIKE` con `*` o `?` | 20 en 3 ficheros | `FUNCIONES UTILES.bas:292` | Comodines de ACE frente a `%` y `_` de PostgreSQL |
| `SELECT TOP n` | 8 en 6 ficheros | `Expediente.cls:209` | `LIMIT`; el orden requiere `ORDER BY` |
| `IIf(`, `Nz(`, `Format(` en SQL o en proyección | 30, 587 y 29 apariciones (incluye VBA) | `Form_FormExpedienteEntidades.cls:1602` | Funciones de Access sin equivalente directo; dependen de si ODBC las traduce |
| `JOIN` (incluido autojoin con alias) | 68 en 11 ficheros | `constructor.bas:2291-2293` (`TbExpedientes AS TbExpedientes_1`) | Los paréntesis anidados de ACE no se aceptan en PostgreSQL |
| Texto `Sí` / `No` como valor | 369 líneas en 29 ficheros | `Form_Form0BDGestorEntidades.cls:125` | Columnas `VARCHAR(2)` hoy; pasarlas a `BOOLEAN` rompería las comparaciones de VBA |

**Se preserva** el comportamiento observable de la fuente: valores `Sí`/`No` en texto, formato de fechas visible en pantalla, orden de listados y mensajes de error de validación.

### Escritura por `Recordset`

`AddNew` (32 en 18 ficheros), `Edit` (24 en 17), `Update` (53 en 19) y `Execute` (113 en 25; con `dbFailOnError` solo 46 en 10) escriben fila a fila sobre `Recordset` DAO. Con tablas vinculadas, un `Recordset` actualizable exige una clave única declarada (o un índice único elegido al vincular); sin ella la tabla se abre en solo lectura. El `schema.sql` del snapshot no contiene claves (ver [matriz campo a campo](matriz-campos.md#fuente-y-límites)); **comprobarlo es previo** a cualquier afirmación de compatibilidad.

## 3. Identidad y generación de claves

- Los identificadores nuevos se calculan en VBA como máximo más uno con `DameID` (30 llamadas en 30 ficheros; definición `FUNCIONES UTILES.bas:2836-2864`), sin bloqueo ni transacción. No hay `DMax`/`DLookup`.
- Las tablas E2E nuevas usan autonumérico (`dbAutoIncrField`, `M_E2EMigration.bas:280-285`).
- **Colisión web–Access.** Si la web inserta con una secuencia de PostgreSQL y Access con `Max + 1`, ambos pueden obtener el mismo valor. Propuesta pendiente de aprobación: una única fuente de identificadores (secuencia de PostgreSQL) y retirar `DameID` en la versión de Access que coexista.
- Los identificadores de usuario son texto en `TbExpedientes` (`IDUsuarioCreacion`, `TEXT(255)`) y numéricos en las tablas E2E (`M_E2EMigration.bas:81-82, 120`): hay que unificar antes de enlazar auditoría entre Access y web.

## 4. Fechas, decimales y nulos

- Las fechas son `SHORT_DATE_TIME` (tipo de fecha y hora de Access); `Date`/`Now` se usan en 27 líneas de 11 ficheros (`CorreoOperaciones.cls:43`). Pendiente decidir si se guarda solo la fecha o fecha y hora (`TIMESTAMP`) y la zona horaria; Access no guarda zona.
- Los importes son `DOUBLE`. La comparación de `HaHabidoCambios` y el redondeo dependen de ello. `NUMERIC` frente a `DOUBLE PRECISION` está sin decidir.
- Casi todas las columnas de texto se guardan como `Null` cuando el valor está vacío (por ejemplo, `ExpedienteOperaciones.cls:99-381`), y la lectura usa `Nz` 587 veces. PostgreSQL distingue cadena vacía de `NULL`: la web debe normalizar igual o las dos interfaces divergirán.

## 5. Transacciones

Solo hay 12 `BeginTrans` en 5 ficheros, todos con el patrón `DBEngine.Workspaces(0)` + `getdb()`:

| Fichero | Uso | Evidencia |
|---|---|---|
| `ExpedienteOperaciones.cls` | `Registrar` (guardado de expediente) y `Eliminar` | `:47`, `:2093` |
| `ExpedienteSuministradorServicio.cls` | Cuatro operaciones del árbol de suministradores, cada una en su transacción | `:67, 123, 202, 258` |
| `E2EEnlacesService.bas` | Alta, baja y desvinculación de enlaces | `:102, 192, 259` |
| `ExpedienteJsonExporter.bas` | Exportación | `:327, 382` |
| `Instalador.bas` | Instalador | `:653` |

Las demás operaciones (catálogos, hitos, modificados, anexos, anualidades, responsables) **no son transaccionales**: varias fichas lo recogen como hallazgo. Una transacción de espacio de trabajo DAO sobre tablas ODBC vinculadas se comporta distinto que sobre un `.accdb`; está sin comprobar. Fuera de la transacción de `Registrar` quedan además el gancho E2E, `PintarTareas` y `RegistrarUltimoCambio` (`ExpedienteOperaciones.cls:523-554`).

## 6. Reglas que hoy solo aplica VBA

La web no puede omitirlas, y PostgreSQL no las aplica si no se declaran como restricciones. Cada una necesita decisión (replicar en servicio, restricción compartida o descartar) y la propuesta de restricción compartida está **pendiente de aprobación**:

- Validación de guardado de expediente `MotivoNoOK` (`ExpedienteOperaciones.cls:800`) y sus recálculos previos (`ESTADO`, `FechaFinGarantia`, `TIpo`, `TipoInforme`, `Ordinal`; `:57-83`).
- Invariante de responsables (`ExpedienteOperaciones.cls:1080`; ver [Form_FormExpedientesResponsables](formularios/Form_FormExpedientesResponsables.md)).
- Unicidad por nombre, código o CIF comprobada con consultas previas, no con índice (por ejemplo, `ComercialOperaciones.cls:28-49`).
- Comprobaciones de «en uso» antes de borrar, sin clave foránea que las respalde (por ejemplo, `CPVOperaciones.cls:102-123`).
- Protección del registro «N/A» por su nombre literal (`ComercialOperaciones.cls:115`).
- Permisos: `EsAdministrador`, `m_AbiertoParaEditar` y etiqueta `EJECUTIVO` en controles (ver [Form_FormExpediente](formularios/Form_FormExpediente.md)).
- Cachés en memoria por sesión (`Entorno`): una escritura de la web no las invalida; Access puede mostrar datos obsoletos hasta reiniciar o repintar.

Recíprocamente, la web debe replicar los **efectos laterales** de `Registrar`: recálculo de la tabla denormalizada `TbExpedientesConEntidades`, `PintarTareas`, `RegistrarUltimoCambio` y el encolado E2E.

## 7. Concurrencia

- No hay bloqueo pesimista ni optimista explícito (0 apariciones de `dbPessimistic`, `dbOptimistic` o `LockEdits`): se usa el modo por defecto de DAO.
- Las ediciones cargan el expediente en memoria (`m_ObjExpedienteDTOActivo`) y lo escriben al guardar, sin comprobar versión de fila: el último que guarda gana, y una edición web intermedia se pisaría.
- `Requery`/`Refresh` aparecen en 16 líneas de 3 ficheros; el resto de pantallas no refresca por su cuenta.
- **Política de conflictos propuesta, pendiente de decisión:** columna de versión o marca de tiempo por fila y rechazo del guardado obsoleto en ambos lados.

## 8. Conexión, credenciales y auditoría

- Hoy el acceso es por fichero con clave guardada en la tabla local de configuración (`Variables Globales.bas:643`). Con ODBC hará falta un origen de datos con credenciales gestionadas fuera del frontend. **No se publica ni modifica ninguna conexión en este estudio.**
- Permisos mínimos propuestos: lectura y escritura solo sobre las tablas vinculadas que Access usa; sin DDL. Pendiente de aprobación.
- Auditoría: `RegistrarUltimoCambio` (27 llamadas en 4 ficheros) delega en `UltimoCambioOperaciones.Registrar` y además refresca una etiqueta de `FormExpedientesGestion` desde la capa de negocio (`FUNCIONES UTILES.bas:189-215`). Cómo se determina el actor no se ha trazado (pendiente). `Environ` solo se usa para el nombre del equipo en correos y para el usuario del sistema en el JSON E2E (`FUNCIONES UTILES.bas:3002`, `ExpedienteJsonExporter.bas:1395`). La web debe distinguir actor y cliente de origen (Access o web) en el mismo registro.

## 9. Pruebas de aceptación bidireccionales (propuestas)

A ejecutar después, en un entorno aislado y nunca sobre datos de producción:

1. Un expediente creado en la web aparece en `FormExpedientesGestion` de Access con los mismos valores, sin reiniciar la aplicación.
2. Un expediente creado en Access aparece en la web con los mismos importes, fechas y textos `Sí`/`No`.
3. Dos altas simultáneas (una por interfaz) no obtienen el mismo identificador.
4. La validación de `MotivoNoOK` rechaza lo mismo en ambos lados, con el mismo mensaje.
5. Un guardado concurrente de un mismo expediente provoca conflicto explícito, no pérdida silenciosa.
6. Borrar un catálogo en uso se rechaza igual en ambos lados.
7. Tras una escritura web, `TbExpedientesConEntidades`, las tareas y el último cambio coinciden con los que habría dejado `Registrar`.
8. Una pérdida de conexión a mitad de un guardado deja la base consistente (transacción íntegra o nada).
9. La auditoría registra actor y cliente de origen en ambos sentidos.

## 10. Corte de la aplicación

El roadmap fija el orden: carga inicial, reconciliación, cambio de vínculos, activación web, deshabilitación de la entrada en la Lanzadera, comprobación de que las demás aplicaciones siguen disponibles y reversión por definir. Lo que `src` de Expedientes aporta: la Lanzadera no aparece como dependencia técnica (solo una preferencia de usuario, `MantenerLanzaderaAbierta`, `Usuario.cls:32`), de modo que la entrada vive fuera de este repositorio y se estudia en Lanzadera. Deshabilitar la entrada no borra binarios ni datos, ni cierra sesiones abiertas o accesos directos.

## 11. P-EXP-3: verificación de la hipótesis del propietario

Estado: **en estudio, sin recomendación**. Esta verificación mide, con lectura estática de los 135 ficheros de producción de `src` (sin los módulos `Test*`), qué usos de RC1 impedirían o encarecerían la hipótesis. No se ha ejecutado nada contra PostgreSQL ni contra un controlador ODBC: «comprobar» significa que solo una prueba real lo resuelve.

### Hipótesis, tal como la formula el propietario

> Tablas idénticas en PostgreSQL, vinculadas por ODBC dentro de cada backend `.accdb`, y el front sin cambios. Como `getdb()` hace `OpenDatabase` sobre el backend, los recordsets sobre tablas vinculadas de ese backend deberían seguir funcionando.
>
> Precisión: `getdb()` sigue abriendo el **mismo** `.accdb` del backend actual (misma ruta); dentro, las tablas locales se sustituyen por vínculos ODBC a PostgreSQL con el mismo nombre.

Que el backend ya contenga vínculos no es nuevo: el `schema.sql` registra 24 tablas vinculadas a otros `.accdb`, y 10 de ellas se usan en `src` (`TbCorreosEnviados`, `TbNoConformidades`, `TbProyectos_AGEDYS`, `TbProyectos_Riesgos`, `TbProyectosEdiciones_Riesgos`, `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos`, `TbUsuarios_HPS`, `TbUsuariosHistoricos_HPS` y `TbExpedientes1`; primera cita de cada una en `Correo.cls:45`, `FUNCIONES UTILES.bas:2544, 2338, 2380, 6108`, `ExpedienteJsonExporter.bas:487`, `constructor.bas:133`, `SuministradorOperaciones.cls:267, 333` e `Instalador.bas:713`). Esos vínculos seguirían siendo de Access a Access.

### Decisiones y reglas del propietario que se registran

1. **Tabla local única.** La única tabla local es la de configuración y se queda en Access: no migra a PostgreSQL.
2. **El frontend no tiene tablas propias**: todo dato va al backend vía `getdb()`.
3. **Formularios desconectados**: sin `RecordSource` ligado a tablas; los datos van por código.

### Verificación de la regla «el frontend no tiene más tablas propias»

- **Confirmado: una sola tabla local, `TbConfiguracionBackends`.** El propio código la llama «tabla local del front» (`ModFixtureF45CacheE2E.bas:11, 28`). La lee `LeeConfiguracionLocal` con `CurrentDb` (`Variables Globales.bas:541-542`), invocada en el arranque (`Variables Globales.bas:309`, función `EVE`) y desde `getdb()` cuando no hay ruta activa (`:622`). Columnas leídas en producción: `BackendActivo`, `BackendProduccion`, `BackendSandbox`, `BackendTest`, `IDAplicacion`, `RutaDirectorioAplicacion_PROD` y `RutaDirectorioAplicacion_LOCAL` (`Variables Globales.bas:549-576`); el fixture lee además `BackendSandbox` y la columna de clave (`ModFixtureF45CacheE2E.bas:43`). Los módulos de prueba la escriben (`Test_ModFixtureF45CacheE2E.bas:90, 104`); ningún código de producción la escribe.
- **Ninguna otra tabla local en producción.** `CurrentDb` aparece en producción solo en `Variables Globales.bas:542` (lectura de la configuración), `:660` (nombre del fichero para la sección del `.ini`), `FUNCIONES UTILES.bas:2849` (respaldo de `DameID`, no recorrido) y `M_E2EMigration.bas:38` (valor por defecto sin `p_Db`). Todo otro acceso a datos usa una base de datos obtenida con `getdb()`. Las tablas que `src` cita y no están en el `schema.sql` son `TbE2ECola` y `TbE2EEnlaces` (se crean en el backend con `getdb()`), `TbExpedientesDatosAGEDYS` (`ExpedienteAGEDYS.cls:59`, también vía `getdb()`; su correspondencia con `TbExpAgedys` del snapshot está sin comprobar) y los alias `TbExpedientes_1` y `TbExpedientesRACs`.
- **Vínculos del frontend a otros orígenes: no comprobable desde `src`.** El frontend es binario y `export-manifest.json` no lista tablas. Hay indicios de que el frontend no vincula las tablas del backend: «el frontend puede no tener las tablas linkeadas» (`Test_M_E2EMigration_TbE2EColaVerificacion.bas:19-21`; `ModFixtureF45CacheE2E.bas:8-9`).
- **Excepciones a la regla (hallazgos):**
  - `RunE2EMigration`, sin el parámetro `p_Db`, trabaja sobre `CurrentDb` (`M_E2EMigration.bas:34-38`), y su cabecera la invoca así desde la ventana inmediata (`:9`): en esa llamada crearía `TbE2EEnlaces` y `TbE2ECola` en el frontend, no en el backend. Ningún módulo de producción la llama con `getdb()`.
  - `DameID` cae a `CurrentDb()` si no recibe base de datos (`FUNCIONES UTILES.bas:2849`); sus 30 llamadas en producción pasan una base de datos, así que el camino no se recorre.
  - `ListaCola.RowSource` se resuelve con SQL contra el motor del frontend (ver verificación de formularios).
  - Estado local por puesto que no es una tabla: `Expedientes.ini` junto al frontend (`Variables Globales.bas:655-666`, `Entorno.cls:1767`), carpeta `%APPDATA%\Aplicaciones DYSN\Expedientes` (`Entorno.cls:2141-2146`), ficheros temporales (`E2EColaHtmlSnapshot.bas:183`, `FUNCIONES UTILES.bas:4226, 4231`) y 33 usos de `TempVars` en memoria.
- **Tablas del backend de tipo temporal o de configuración por puesto.** En el snapshot son `TbE2EExportSeleccionTemp`, `TbE2EJsonDestinationUserConfig`, `TbAuxNemotecnico`, `TbAusExpPostAGEDO` y `TbAuxEstadosMartina`, y ninguna se cita en producción de RC1. `TbConfMostrarEstado` sí se cita (3 ficheros). Con la regla del propietario, si existen viven en el backend; su disposición (PostgreSQL o local dentro del backend) depende de la P-EXP-12.

### Verificación de la regla «formularios desconectados»

- **Confirmado en el diseño:** 0 `RecordSource` y 0 `ControlSource` en los 45 `.form.txt`. De los 82 listados y combos con nombre, 81 declaran `RowSourceType = "Value List"`.
- **Excepción 1, `ListaCola` de `Form_FormE2ECola`:** `RowSourceType = "Table/View/StoredProc"` (`Form_FormE2ECola.form.txt:288`) y `RowSource` asignado desde código con SQL sobre `TbE2ECola` (`Form_FormE2ECola.cls:137`, SQL en `modE2EColaPanelHelper.bas:74-86`). Esa SQL la resuelve el motor del frontend, no `getdb()`. Es el único caso, y el formulario ya no se abre desde el menú (opción 12 abre el panel HTML; `Form0BDGestorEntidadesLogic.bas:55-57`).
- **Subformularios:** `FrmDetalle.SourceObject` se asigna por código a formularios, no a tablas (`Form_FormExpediente.cls:146`).
- **No ligados pero con SQL:** los demás listados reciben `AddItem` o cadenas de lista de valores construidas desde datos leídos por `getdb()` (por ejemplo `Form_FormExpedienteEntidades.cls:1967`), de modo que los datos viajan por el camino de `getdb()`.

### Qué rompería o encarecería la hipótesis

| Categoría | Evidencia | ¿Rompe? | Coste (sitios) |
|---|---|---|---|
| `dbOpenTable`, `.Seek`, `.Index`, `FindFirst`, `FindNext`, `.Bookmark`, `.Filter`, `.Sort`, `.LastModified` | 0 usos en 135 ficheros | No | 0 |
| `OpenRecordset` por nombre de tabla (`m_SQL = "TbX"` y `OpenRecordset(m_SQL)` sin tipo) | 32 líneas en 15 ficheros; por ejemplo `ExpedienteOperaciones.cls:76, 88`, `CPVOperaciones.cls:50` | Cambia el tipo por defecto (de tabla a dynaset), sin dependencias de `Seek`; la escritura requiere clave única en la tabla vinculada | Comprobar 32 sitios; clave pendiente del ERD |
| Altas con `Max + 1` (`DameID`) | 30 llamadas en 30 ficheros, 27 pares tabla-campo, una sola función (`FUNCIONES UTILES.bas:2836-2864`); una sobre tabla vinculada a otro backend (`TbCorreosEnviados`/`IDCorreo`) y otra sobre una tabla solo de `src` (`TbExpedientesDatosAGEDYS`) | **Sí**: colisión con las altas de la web | 1 función + 4 `Max` directos independientes (`E2EEnlacesService.bas:110`, `E2EFamilyExportHelper.bas:303`, `FUNCIONES UTILES.bas:1730, 1746`) |
| Dialecto Jet en SQL (literales) | `DELETE *`: 49 líneas en 15 ficheros (`CPVOperaciones.cls:120`); `TOP n`: 4 líneas (`Expediente.cls:209`, `ExpedienteCompleto.cls:575`, `ManualE2EExport.bas:250`); `#` de fecha: 7 líneas (`ExpedienteOperaciones.cls:1789, 1855`); `Date()`/`Now()`: 5 líneas (`E2EEnlacesService.bas:120, 199, 261`, `constructor.bas:1586`); `DateDiff`: `constructor.bas:1586, 1641`; `Nz` en SQL: 4 líneas (`ExpedienteSuministradorRepositorio.bas:24`); `[campo]`: 6 líneas (`E2EFamilyExportHelper.bas:101, 199, 277`); `LIKE` con `*`: `FUNCIONES UTILES.bas:403, 412, 1228`, `constructor.bas:2680`; `InStr`: `FormE2EGestionEnlacesLogic.bas:480-481` | A comprobar con el controlador (ACE puede evaluar localmente lo que no traduzca, con coste de rendimiento) | Unas 80 líneas de las ≈1 000 con SQL (40 ficheros) |
| Consultas paso-a-través | 0 (`dbSQLPassThrough`, `.Connect`, `ReturnsRecords`: sin usos) | No aplica | 0 |
| Consultas guardadas | 1 exportada (`src/queries/Consulta1.sql`, sobre `TbExpedientes`); 5 `CreateQueryDef("")` temporales con `PARAMETERS` (`E2EFamilyExportHelper.bas:89, 99, 197, 274`, `ExpedienteJsonExporter.bas:599`) | Sintaxis Jet de parámetros; a comprobar | 6 |
| Transacciones | 12 `BeginTrans` en 5 ficheros (tabla siguiente) | A comprobar: una transacción de `Workspaces(0)` sobre vínculos ODBC; mezcla de motores si participa una tabla local o vinculada a otro `.accdb` | 12 sitios |
| DDL por DAO sobre tablas | `M_E2EMigration.bas` (35 líneas: `CreateTableDef`, `Fields.Append`, `Indexes.Append`, `Required`) | Rompe para tablas ya vinculadas: DAO no modifica la estructura de una tabla vinculada (a confirmar con el controlador) | 1 módulo |
| `Recordset.RecordCount`/`MoveLast` | `ExpedienteSuministradorServicio.cls:395-396` | No | 1 |
| Tipo `Sí`/`No` | 0 columnas `BOOLEAN` en el `schema.sql`: se guardan como `TEXT(2)`; 369 líneas comparan el texto | No, mientras se conserve como texto | 0 si se conserva |
| `MEMO` | 24 columnas en las tablas activas (por ejemplo `Titulo`, `ObjetoContrato`, `Observaciones`) | A comprobar: ACE ordena y agrupa un memo solo por sus primeros 255 caracteres | Revisar `ORDER BY` y `GROUP BY` sobre esas 24 |
| Fecha y hora | 29 columnas `SHORT_DATE_TIME`; formatos de salida con `Format` en 29 líneas | A comprobar: precisión de fracciones de segundo al comparar con un `timestamp` | 29 columnas |
| Autonumérico | `IDEnlace` e `IDCola` (`M_E2EMigration.bas:76, 112`); `INSERT` que omite `IDEnlace` (`E2EEnlacesService.bas:117-118`); 0 usos de `@@IDENTITY` o `.LastModified` | A comprobar: valor por defecto de la columna en PostgreSQL; el resto de autonuméricos depende del ERD | 2 tablas conocidas |

Transacciones, por procedimiento (tablas citadas dentro del procedimiento; las operaciones que llama escriben más):

| Procedimiento | Tablas citadas | Observación |
|---|---|---|
| `ExpedienteOperaciones.Registrar` (`:47`) | `TbExpedientes`, `TbExpedientesConEntidades` | Llama a las operaciones de anualidades, comerciales, CPV, PECAL, RAC, anexos, responsables, lugares y suministradores, y a `ActualizarNCs` (`FUNCIONES UTILES.bas:6317`); `TbNoConformidades`, vinculada a otro `.accdb`, se lee en `getNCParaExpediente` (`FUNCIONES UTILES.bas:2544`), sin comprobar si participa en la transacción |
| `ExpedienteOperaciones.Eliminar` (`:2093`) | 13 tablas del expediente | La más amplia |
| `ExpedienteSuministradorServicio` (`:67, 123, 202, 258`) | por el repositorio | Cuatro operaciones del árbol |
| `E2EEnlacesService` (`:102, 192, 259`) | `TbE2EEnlaces`, `TbExpedientes` | Combinan el `Max` del ordinal con la escritura en dos tablas |
| `ExpedienteJsonExporter` (`:327, 382`) | ninguna directa | Exportación |
| `Instalador.RegistraEnlazadoE2EEnEntidades` (`:653`) | `TbExpedientesConEntidades` | Utilidad de instalación |

### Coste agregado, sin priorizar

- **Concentrado en un punto:** `DameID` (30 llamadas, 1 función).
- **Por sitios:** 4 `Max` directos, 32 aperturas por nombre de tabla a comprobar, 12 transacciones, 6 consultas con sintaxis Jet de parámetros, 1 módulo de DDL y las construcciones Jet de SQL.
- **Dependiente del ERD de origen (D-EXP-4):** claves únicas para recordsets actualizables, autonuméricos, columnas `NOT NULL` y la lista final de tablas del backend.
- **Fuera de `src`:** la prueba real de las filas «a comprobar» exige un backend de RC1 y un controlador ODBC; ver P-EXP-1.

## Pendientes

- Resolver la P-EXP-3 (sección 1 y apartado 11); está abierta y escalada al propietario.
- Extraer claves, índices y relaciones reales de las tablas vinculables (D-EXP-4 en el README).
- Comprobar el comportamiento de transacciones DAO y de `Recordset` actualizables sobre tablas ODBC.
- Aprobar la fuente única de identificadores y la política de conflictos.
- Trazar cómo se determina el actor del último cambio y de la auditoría.
- Inventariar qué consultas guardadas (`src/queries/`) deben recrearse como vistas.
