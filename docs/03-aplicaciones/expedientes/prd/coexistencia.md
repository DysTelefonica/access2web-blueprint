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

**Consecuencia principal.** El modelo de vínculos PostgreSQL del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) (Access sobre tablas vinculadas por ODBC) no es un cambio de configuración: `getdb()` abre un `.accdb` por ruta y las 303 llamadas esperan un `DAO.Database` de ese fichero. Hay dos caminos y la decisión está **abierta**: (a) sustituir `getdb()` para que devuelva `CurrentDb` con las tablas vinculadas, o (b) mantener un backend `.accdb` y sincronizarlo con PostgreSQL, que el roadmap descarta («no exige sincronizar dos copias de datos»). Con (a), cada fragmento de SQL, de `Recordset` y de transacción de las secciones siguientes se ejecutaría contra el controlador ODBC.

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

## Pendientes

- Decidir entre sustituir `getdb()` por `CurrentDb` sobre vínculos o conservar otra vía (sección 1).
- Extraer claves, índices y relaciones reales de las tablas vinculables (D-EXP-4 en el README).
- Comprobar el comportamiento de transacciones DAO y de `Recordset` actualizables sobre tablas ODBC.
- Aprobar la fuente única de identificadores y la política de conflictos.
- Trazar cómo se determina el actor del último cambio y de la auditoría.
- Inventariar qué consultas guardadas (`src/queries/`) deben recrearse como vistas.
