# Módulo de funciones útiles

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/modules/FUNCIONES UTILES.bas` (1035 líneas, 40 procedimientos) · **PRD:** [README](../README.md)

El módulo más grande y más transversal de Condor: abre las conexiones a los cinco backends, reconecta, implementa el manejador de errores que usan todos los formularios, y reúne utilidades de interfaz (foco, estilos de bloqueo, combos,barra de progreso) y de soporte a la aplicación (búsqueda de identificadores, existencia de registros, duración). Es, con diferencia, el punto donde más decisiones de coexistencia se concentran: aquí se decide a qué base se abre cada cosa y en qué circumstances.

Abreviaturas: `FU`, `src/modules/FUNCIONES UTILES.bas`; `VG`, `modules/Variables Globales.bas`; `RC`, `modules/RepositorioComun.bas`; `CE`, `classes/CondorError.cls`.

## Superficie pública

| Objeto | Tipo | Procedimientos públicos | Qué resuelve |
|---|---|---|---|
| `FUNCIONES UTILES` | módulo | cuarenta procedimientos: conexiones, interfaz y soporte (`FU:56-1035`) | Apertura y reconexión de los cinco backends, manejo de errores de formulario y utilidades compartidas |

## Procedimientos y llamadas

| Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|
| `getWorkspace` | Devuelve el `Workspace` aislado de la aplicación | — | Sin efectos | `FU:56-63` |
| `getdb` | Devuelve la conexión global de Condor, revalidándola y reconectando si hace falta | `CheckAndReconnect` | Abre `g_dbCondor` | `FU:84-124`; revalidación en `FU:91-103`, reconexión en `FU:105-108`, error 3045 en `FU:111-114` |
| `getdbLanzadera`, `getdbExpedientes`, `getdbNoConformidades`, `getdbCorreo` | Devuelven la conexión de cada backend | `CheckAndReconnect` | Abren la respective `g_db*` | `FU:64-83`, `FU:136-210` |
| `CheckAndReconnect` | Decide la ruta de la conexión y la abre | `getWorkspace`, `GetPasswordDB` | Abre la base elegida | `FU:211-263`; elección por tipo en `FU:233-247`, apertura en `FU:252` |
| `GetPasswordDB` | Devuelve la contraseña con la que se abren los backends | — | Sin efectos | `FU:126-135` |
| `CerrarConexionesGlobales` | Cierra las conexiones globales de la aplicación | — | Cierra bases | `FU:264-276` |
| `ManejadorErroresFormulario` | Traduce un error de VBA a mensaje de usuario y deja traza | `LogErrorUI`, `ExtraerDescripcionLimpia`, `GetBaseErrorNumber` | Inserta en el log de errores y muestra un `MsgBox` | `FU:588-679`; captura inmediata en `FU:596-598` |
| `GetBaseErrorNumber` | Desenvuelve el número de error de un `CondorError` | — | Sin efectos | `FU:680-688`; la resta del offset de error de objeto en `FU:682-683` |
| `RegistroExiste` | Comprueba si existe una fila por tabla, campo y valor | `OpenRecordset` | Sin escritura | `FU:827-846`; SQL construida por concatenación en `FU:832` |
| `DameID` | Lee el valor de un campo de identidad de una tabla | `getdb` | Sin escritura | `FU:465-517` |
| `LlenarComboDesdeDiccionario` | Rellena un combo desde un diccionario, con opción «Todos» opcional | — | Sin efectos | `FU:538-587` |
| `TieneControl`, `QuitarFocoControl`, `TryFocusAnyControl`, `AplicarEstiloBloqueo`, `HabilitarControlesSubformulario`, `Ajustar`, `FormularioAbierto`, `CommitEditsOnForm`, `CommitEditsOnActiveSubform` | Utilidades de interfaz y de confirmación de edición | varias | Sin efectos sobre la base | `FU:277-346`, `FU:775-782`, `FU:847-897`, `FU:898-980` |
| `CalcularDuracion` | Devuelve la duración entre dos fechas con formato corto | — | Sin efectos | `FU:981-1004`; caso no válido en `FU:986-987` |
| `Avance`, `AvanceCerrar`, `ResetAvanceCounter`, `GestionarRibbon`, `AbrirEnLocal`, `AbrirAyuda`, `JoinCollection`, `EliminarContenidoCarpeta`, `TablaExiste`, `ReemplazarComillas`, `ConstruirContexto`, `ExtraerDescripcionLimpia`, `FormatearDescripcionValidacion`, `LogErrorUI`, `ControlPuedeRecibirFoco`, `RellenarObjetoDesdeRecordset` (privada en `RC`) | Utilidades varias | — | Sin efectos salvo `EliminarContenidoCarpeta` (`FU:518-537`) | `FU:346-464`, `FU:1005-1035`, `FU:689-826` |

## Flujos de trabajo

- **Normal:** cualquier repositorio pide su conexión con `getdb()`; el módulo comprueba que el objeto global siga vivo,leyendo `.Name` para forzar la validación real porque el recolector de VBA puede liberar el objeto aunque quede una referencia (`FU:88-103`). Si la conexión falta o está cerrada, `CheckAndReconnect` la reconstruye (`FU:105-108`, `FU:219-231`).
- **Alternativos:** la ruta de cada backend depende de dos variables temporales de la aplicación: si `DatosEnLocal` vale «Sí», se abre la copia local en lugar de la remota (`FU:235`, `FU:238`, `FU:240`, `FU:242`, `FU:244`), y para Condor, si `EnPruebas` vale «Sí», se abre la base de pruebas en lugar de la de producción (`FU:237`). Un tipo de base desconocido lanza un error 513 con el valor recibido (`FU:246`).
- **Errores:** `ManejadorErroresFormulario` captura número, descripción y origen del error antes de cualquier otra instrucción (`FU:596-598`) y distingue los errores propios de la aplicación de los de Access desenvolviendo el offset (`FU:680-688`). Si el backend principal ya está abierto en otra sesión, el error 3045 se traduce a un mensaje y la aplicación se cierra (`FU:111-114`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Hay cinco backends y cada uno tiene su rama de ruta | `CheckAndReconnect` | `FU:233-247` |
| El modo local y el de pruebas se activan con variables temporales de la aplicación | `Application.TempVars` | `FU:235`, `FU:237` |
| La contraseña se compone al abrir, no en la cadena de la aplicación | `GetPasswordDB` | `FU:126-135`, `FU:252` |
| Los errores propios se leen restando el offset de error de objeto | `GetBaseErrorNumber` | `FU:682-683` |
| Una duración inválida devuelve `0m` en lugar de un valor vacío | `CalcularDuracion` | `FU:986-987` |

## Lecturas y escrituras

| Operación | Recurso | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Apertura de los cinco backends | ficheros `.accdb` de cada aplicación | conexiones DAO con contraseña en la cadena | Ninguna | `FU:233-252` |
| Existencia de una fila | tabla indicada por el llamador | SQL concatenada | Ninguna | `FU:832-833` |
| Identidad de una tabla | tabla indicada por el llamador | lectura directa | Ninguna | `FU:465-517` |
| Registro de errores | tabla de log de errores | inserción desde el manejador | Ninguna | `FU:588-679` |

## Coexistencia

Este módulo concentra la decisión de coexistencia más delicada del PRD. Hoy la aplicación abre **cinco** ficheros `.accdb` por ruta, con la contraseña en la cadena de conexión (`FU:252`), y elige entre copia local y remota según variables temporales (`FU:235-244`). Con PostgreSQL vinculado, las cinco conexiones se convierten en un enlace por aplicación —o en una sola conexión si el esquema se unifica— y las variables de modo local y de pruebas dejan de existir: el destino es único. `ElimiarContenidoCarpeta` (`FU:518-537`) existe para vaciar la copia local de datos, lo que solo tiene sentido con ese modelo de trabajo. `RegistroExiste` concatena el nombre de tabla y de campo en la SQL (`FU:832`): en el destino hay que decidir si se admite SQL dinámico o si se restringe a un catálogo cerrado de tablas. `AplicarEstiloBloqueo` y `CommitEditsOnForm` (`FU:847-897`, `FU:931-980`) modelan el concepto de «bloquear el formulario mientras guarda», que en web es un estado del componente, no un control de Access. Y `ManejadorErroresFormulario` decide, por número de error, si el mensaje es de validación, operativo o técnico (`FU:680-714`): esa clasificación debe sobrevivir como contrato de la API.

## Comportamiento actual · Posibles defects · Propuestas web pendientes de decisión

**Comportamiento actual.** Un único módulo concentrations conexiones, manejo de errores y utilidades. La conexión es global por aplicación y se revalida en cada llamada. El destino de cada backend depende del modo local y del modo de pruebas.

**Posibles defectos, a verificar.**
1. `RegistroExiste` construye la SQL con el nombre de tabla y de campo concatenados (`FU:832`): cualquier llamador con dato externo convierte esto en inyección.
2. La contraseña del backend se compone dentro de la cadena de conexión (`FU:252`, `FU:126-135`).
3. La revalidación de la conexión global se hace leyendo `.Name` de un objeto que puede haber sido liberado (`FU:91-103`): es un apaño documentado, no un diseño.
4. El modo local abre una copia de datos por aplicación y existe una rutina para vaciarla (`FU:235-244`, `FU:518-537`): son datos fuera del control del repositorio.
5. La aplicación se cierra sola si la base principal está abierta en otra sesión (`FU:111-114`).

**Propuestas web pendientes de decisión.** Que exista un único enlace por base de datos en el destino, o uno solo para todo el roadmap, y que la decisión quede documentada. Que el modo de pruebas sea un entorno con sus propias credenciales y no una variable de la aplicación. Que la clasificación de errores por número sea un contrato de la API y no un `Select Case` de VBA. Que la limpieza de datos sea una operación de soporte con permisos, no una rutina invocable desde la interfaz.

## Criterios de aceptación

1. La aplicación abre sus backends una sola vez y reutiliza la conexión. Trazado a `FU:105-108`.
2. Un error de validación y un error operativo producen mensajes distintos y ambos dejan traza. Trazado a `FU:588-679` y `FU:680-714`.
3. Si el backend principal está ocupado, el mensaje es el propio y no un error técnico. Trazado a `FU:111-114`.
4. Una duración con fecha final anterior a la inicial devuelve `0m`. Trazado a `FU:986-987`.
5. La comprobación de existencia de registro funciona con tabla y campo variables sin construir SQL con literales. Trazado a `FU:832-833`.

## Pendientes

- Confirmar contra el binario el valor de las constantes de ruta local y remota y quién las mantiene: son configuración de entorno dentro del repositorio (`FU:235-244`).
- Confirmar si el modo local de `Application.TempVars("DatosEnLocal")` sigue activo en producción o es solo del entorno de pruebas (`FU:235`).
- Confirmar qué formularios dependen de `GestionarRibbon` (`FU:1024-1035`), porque la cinta de opciones no existe en web.
- Confirmar el destino de `AbrirEnLocal` y `AbrirAyuda` (`FU:404-464`): son dos procesos que abre el cliente de escritorio.
- Confirmar qué consumidores pasan datos externos a `RegistroExiste` (`FU:827`).