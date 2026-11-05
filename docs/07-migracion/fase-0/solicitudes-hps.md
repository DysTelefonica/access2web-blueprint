# Fase 0: Solicitudes HPS (`Solicitudes_HPS` / `Solicitudes_HPS_datos`)

Épica [`docs/07-migracion/epica-coexistencia-postgresql.md`](../epica-coexistencia-postgresql.md) (fase 0). Seguimiento: issue #1581 (descubrimiento global de todas las herramientas); no existe todavía una issue propia de la ficha de Solicitudes HPS.
Base de análisis: rama `main`, SHA `cfa1937ab05a9e40150ea9236243898de96f84da`, repositorio `00_HPS_SOLICITUDES` (`C:\00repos\codigo\00_HPS_SOLICITUDES`). El árbol de trabajo coincide con el blob publicado: `git hash-object` de `Solicitudes_HPS.accdb` y `Solicitudes_HPS_datos.accdb` devuelve los mismos objetos que `main:`, y `git status --short` no muestra ninguno de los dos como modificado.
Fuentes: ERD del front y del backend con `scripts/dysflow-erd/extract-erd.ps1` (solo lectura, sobre copia temporal) y trazado del código con CodeGraph sobre el índice del propio repositorio (108 ficheros, 5653 nodos), completado con búsqueda de texto sobre `src/` para las cadenas de SQL y de conexión. CodeGraph sí estaba operativo en este entorno; no hubo que declararlo.

Las rutas de código de esta ficha son relativas al repositorio `00_HPS_SOLICITUDES` en el SHA indicado.

## Base de análisis

| Hecho | Evidencia |
|---|---|
| No hay rama RC1 ni de release: solo `main` y `origin/develop` | `git branch -avv` en `00_HPS_SOLICITUDES`: `main` en `cfa1937`, `origin/develop` en `0070da9`, `origin/HEAD` → `origin/main` |
| `develop` no es una RC: no cambia el código | `git diff --stat main origin/develop` devuelve una sola línea: `AGENTS.md \| 301 +++---` (209 inserciones, 92 borrados). `src/` y los dos `.accdb` son idénticos en ambas ramas |
| `develop` es un volcado de trabajo sin ciclo de release detrás | `git log --format='%h %ad %s' --date=iso origin/develop`: `0070da9 2026-05-20 chore(sync): snapshot local changes` y `10f3467 2026-04-19 cambios en AGENTS.md`. El único commit de código es `cfa1937 2026-04-16`, ya en `main` |
| No hay documento de UAT en el repositorio | El árbol `docs/` solo contiene `DISCOVERY_MAP.md` y `Informes/Informe_Correccion_ID256_2026-001.md`. Cero coincidencias de «UAT» en el repositorio |
| La única evidencia de ciclo de release es un informe de corrección | `docs/Informes/Informe_Correccion_ID256_2026-001.md:4` («Versión corregida: 2026-001»), `:16` (fecha 16/04/2026), `:20` («Estado: Solucionado»), `:47-51` (validación: «La compilación en Access ha finalizado sin errores» y el formulario muestra la solicitud 256) |
| El cambio OpenSpec de esa corrección sigue sin archivar | `C:\00repos\documentacion\OPENSPEC\00_HPS_SOLICITUDES\changes\fix-tareas-pendientes-tramitador\` contiene `proposal.md`, `SPEC.md` y `TASKS.md`, sin carpeta `archive`. Los criterios de éxito de `proposal.md` («Success Criteria») siguen sin marcar |
| La versión declarada por el front no coincide con la del informe | `Solicitudes_HPS.ini:2` → `VersionAplicacion=2025-042`, y el informe declara la versión corregida `2026-001`. La versión se lee del `.ini` en `src/classes/Entorno.cls:2200` (`m_sVersionAplicacion = LeerIni("VersionAplicacion", "#ERR#")`) |
| La lanzadera registra la aplicación, pero sin ediciones ni aperturas | Consulta de solo lectura sobre una copia de `C:\00repos\datos\Lanzadera_Datos.accdb`: `TbAplicaciones` tiene `IDAplicacion=22`, `NombreAplicacion=SOLICITUDES HPS`, `NombreCorto=SOLICITUDES_HPS`. `TbAplicacionesEdiciones` no tiene ninguna fila para `IDAplicacion=22` y `TbAplicacionesAperturas` no tiene ninguna fila para `SOLICITUDES%` |
| El inventario del blueprint está desfasado en la ruta y en las ramas | `docs/01-inventario-aplicaciones.md:15` apunta a `C:\00repos\codigo\HPS_SOLICITUDES` «(sin prefijo `00_`; fallback `main` por ausencia de checkout `00_HPS_SOLICITUDES/`)», y `docs/03-aplicaciones/hps-solicitudes/README.md:13` afirma «sin rama staging, sin rama develop, sin rama release; solo `main`». Los dos clones existen hoy y son idénticos en contenido: mismo árbol `src/` (108 ficheros) y mismos SHA-256 de los dos binarios (`022c81e3…` el front, `df0c380e…` el backend). `origin/develop` sí existe |

Consecuencia: la regla del propietario («si hay RC1 con UAT pendiente, se estudia esa RC1») no se activa en Solicitudes HPS: no hay RC1 y no hay UAT. La base de este inventario es `main` en `cfa1937`, cuyo contenido es idéntico al de `develop`. La ausencia de evidencia de UAT es un dato, no un vacío: el repositorio no registra ningún ciclo de pruebas del sistema.

Hay dos clones del mismo repositorio en disco. El estudio se ha hecho sobre `00_HPS_SOLICITUDES`, el que lleva el prefijo canónico; el otro (`HPS_SOLICITUDES`) tiene una rama local `audit/usuarios-entidades-consumers` (`b5ac650`, «chore: wip snapshot before removing local worktree») que solo añade `.codegraph/.gitignore` y `.engram/config.json`: no aporta evidencia sobre los consumidores de `TbUsuariosEntidades`.

## Estructura del front

| Hecho | Valor |
|---|---|
| Tablas del front | 1 propia y 0 vinculadas |
| Tabla propia del front | `TbCorreosEnviadosParaPruebas` (0 filas, 18 campos, clave primaria `IDCorreo`) |
| Vínculos en el front | Ninguno. El front no hereda ninguna tabla: todo se abre por código |
| Consultas guardadas del front | 0 |
| Formularios | 34 en el binario y 34 en `src/forms/` (34 `.form.txt` y 34 `.cls`), con los mismos nombres salvo el prefijo `Form_` que añade el export |
| Formularios con `RecordSource` de diseño | 0. Cero apariciones de `RecordSource` en todo `src/` |
| Controles con `RowSourceType = Table/Query` o `Table/View/StoredProc` | 0 |
| Controles con `RowSourceType = "Value List"` | 50, de los que 39 llevan valor de `RowSource`; ninguno contiene `SELECT` |
| Controles con `ControlSource` con valor | 0. La única aparición es `src/forms/Form_FormEnvioCorreo.form.txt:1624-1628`, que deja el control sin origen («Unbound para permitir edición») |
| Asignaciones a `.RowSource` en tiempo de ejecución | 64, todas con lista literal de columnas o cadena vacía; ninguna contiene SQL |
| Asignaciones a `.RecordSource` o `.ControlSource` en tiempo de ejecución | 0 con SQL; 1 que pone el `ControlSource` a vacío |

La regla del propietario se confirma por medición en esta aplicación: los formularios son desconectados. No hay un solo control que ata la interfaz a un esquema, ni en diseño ni en tiempo de ejecución. El SQL vive en el código, en cadenas construidas a mano, y es lo que hay que Trasladar a PostgreSQL.

Detalle de los 39 `RowSource` con valor, todos listas de valores literales, en el [anexo](#anexo-controles-con-origen-de-datos).

### Macros

Cero coincidencias de `AutoExec`, `AutoOpen`, `AutoClose` o `Macro` en los 108 ficheros de `src/`. No se pudo confirmar por otra vía que el front no tenga macros embebidas: la enumeración por COM de `Access.Application.AllMacros` sobre una copia del binario se quedó colgada en este entorno y se cerró el proceso. Queda declarado en «Lo que no se pudo verificar».

## Aperturas directas

Cinco aperturas, todas con la misma credencial literal en la línea de `OpenDatabase` (`"MS Access;PWD=***"`). Ninguna usa `DoCmd.OpenTable`, `DoCmd.OpenQuery` ni `TransferDatabase`: cero coincidencias en `src/`.

| Apertura | Código | Ruta |
|---|---|---|
| Backend propio `Solicitudes_HPS_datos.accdb` | `src/modules/Variables Globales.bas:458` (`getdb()`), apertura en `:478` | Local `…\000datoslocal\Solicitudes_HPS_datos.accdb` (`:467`) o remota `\\datoste\aplicaciones_dys\Aplicaciones PpD\SOLICITUDES HPS\Solicitudes_HPS_datos.accdb` (`:277` + `:278`) |
| `HPST.accdb` | `src/modules/Variables Globales.bas:522` (`getdbHPS()`), apertura en `:542` | `…\000datoslocal\HPST.accdb` (`:531`) o `…\HPS\HPST.accdb` (`:533`) |
| `Expedientes_datos.accdb` | `src/modules/Variables Globales.bas:580` (`getdbExpedientes()`), apertura en `:599` | `…\000datoslocal\Expedientes_datos.accdb` (`:588`) o `…\EXPEDIENTES\Expedientes_datos.accdb` (`:590`) |
| `Lanzadera_Datos.accdb` | `src/modules/Variables Globales.bas:430` (`getdbLanzadera()`), apertura en `:449` | `…\000datoslocal\Lanzadera_Datos.accdb` (`:438`) o `…\0Lanzadera\Lanzadera_Datos.accdb` (`:440`) |
| `Correos_datos.accdb` | `src/modules/Variables Globales.bas:488` (`getdbCorreo()`), apertura en `:512` | `…\000datoslocal\Correos_datos.accdb` (`:497`) o `…\00Recursos\Correos_datos.accdb` (`:499`) |

La raíz remota se fija en `src/modules/Variables Globales.bas:277` (`\\datoste\aplicaciones_dys\Aplicaciones PpD\`) y la carpeta de la aplicación en `:278` (`SOLICITUDES HPS\`). El resto de backends cuelga de la raíz común: `HPS\`, `EXPEDIENTES\`, `0Lanzadera\`, `00Recursos\`.

`CurrentDb()` aparece dos veces (`src/modules/FUNCIONES UTILES.bas:773` y `src/modules/Variables Globales.bas:613`), pero ninguna de las dos es un acceso a tabla: la primera es el valor por defecto de `DameID` y la segunda lee el nombre del fichero para el `.ini`. La primera sí produce un defecto; está en la propuesta S11.

## Tablas de otros backends que consume el front

El backend `Solicitudes_HPS_datos` vincula 10 tablas de otros backends, pero el front no usa ninguna por vínculo: las abre por código. La tabla recoge, por backend de origen, qué tablas lee y cuáles escribe.

| Tabla | Backend dueño | Lectura | Escritura |
|---|---|---|---|
| `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` | `Lanzadera_Datos` | `src/modules/constructor.bas:87` (`getdbLanzadera()` en `:91`), `:118-123`, `:173-178`, `:228-233`, `:331` | Ninguna |
| `TbExpedientes` | `Expedientes_datos` | `src/modules/constructor.bas:2306` (`:2309`), `:2526` y `:2529` (`:2537`), `:3869` y `:3873` (`:3876`); los tres por `getdbExpedientes()` | Ninguna |
| `TbSuministradores`, `TbExpedientesSuministradores` | `Expedientes_datos` | `src/modules/constructor.bas:14-16` (`getdbExpedientes()` en `:21`), `:2353`, `:2356`, `:2404`, `:2406` | Ninguna |
| `TbHPS` | `HPST` | `src/modules/constructor.bas:625` (`getdbHPS()` en `:631`) | Ninguna |
| `TbUsuarios` | `HPST` | `src/modules/constructor.bas:2205`, `:3126`, `:3144`, `:3229`, `:3494`, `:3545`, `:3599`, `:3762`, `:3818`; `src/classes/SolicitudOperaciones.cls:1178` | **Sí**: `src/modules/FUNCIONES UTILES.bas:2449` (`UPDATE TbUsuarios SET IDSolicitud`), `src/classes/MotivoHPS.cls:208` (`UPDATE TbUsuarios SET Motivo_HPS`, vía `getdb()` y el vínculo) |
| `TbUsuariosEntidades` | `HPST` | Sin lecturas directas | **Sí**: `src/modules/FUNCIONES UTILES.bas:18`, `:23`, `:27` (renombrado de empresa) y `:2452` (`UPDATE … SET IDSolicitud`), los cuatro con `getdbHPS().Execute` |
| `TbUsuariosHistoricos` | `HPST` | `src/modules/constructor.bas:3162`, `:3180`, `:3248` | **Sí**: `src/classes/MotivoHPS.cls:211` (`UPDATE TbUsuariosHistoricos SET Motivo_HPS`, vía `getdb()`) |
| `TbMotivoHPS` | `HPST` | `src/modules/constructor.bas:2111`, `:2163`; `src/classes/MotivoHPS.cls:74-76`; `src/modules/FUNCIONES UTILES.bas:1457-1459` | **Sí**: `src/classes/MotivoHPS.cls:41` (`UPDATE`, vía `getdb()`), `src/modules/FUNCIONES UTILES.bas:1430` (`INSERT`, vía `getdbHPS()`), `src/modules/FUNCIONES UTILES.bas:1466-1470` (`AddNew` + `Update`, vía `getdbHPS()`) |
| `TbCorreosEnviados` | `Correos_datos` | — | **Sí**: `src/modules/FUNCIONES UTILES.bas:652-663` inserta el aviso de error con `Aplicacion = "HPS_SOLICITUD"` en la tabla de `Correos_datos` |

Dos conclusiones de esta tabla:

1. **El nombre `TbCorreosEnviados` designa dos tablas distintas en dos backends.** La del backend propio (`Solicitudes_HPS_datos`) es la que usa la aplicación para su listado de correo (`src/modules/CorreoRepositorio.bas:21` con `getdb()` en `:47`, `src/modules/InformesOperaciones.bas:112`, `:182`, `:304`, `src/modules/constructor.bas:2462-2469`, `:2787`, `:2830`). La de `Correos_datos` es donde se registra el correo enviado (`src/modules/FUNCIONES UTILES.bas:652`).
2. **El front escribe en el backend de otra aplicación.** Cuatro tablas de `HPST` (`TbUsuarios`, `TbUsuariosEntidades`, `TbUsuariosHistoricos`, `TbMotivoHPS`) y una de `Correos_datos` reciben escrituras desde esta aplicación, sin pasar por su propio backend. Es una excepción a R-1: el front consulta y escribe en el backend de otra aplicación.

## Backend `Solicitudes_HPS_datos`

Medido sobre el binario del repositorio (`Solicitudes_HPS_datos.accdb`, 3 952 640 bytes, SHA-256 `df0c380e…`). El backend autoritativo del entorno de datos (`C:\00repos\datos\Solicitudes_HPS_datos.accdb`, 13 631 488 bytes) se midió aparte y se contrasta en el apartado siguiente.

| Hecho | Valor |
|---|---|
| Tablas | 21: 11 propias y 10 vinculadas |
| Propias | `Copia de TbExpedientes` (70 filas), `TbConfiguracion` (1), `TbCorreosEnviados` (9), `TbHPSGrado` (13), `TbJustificaciones` (7), `TbLogs` (0), `TbLogsGeneral` (1855), `TbResponsables` (26), `TbSolicitudes` (220), `TbSolicitudesFechas` (220), `TbUltimoCambio` (4) |
| Vinculadas a `HPST` | `TbHPS`, `TbHPSEquivalencia`, `TbHPSGrado1` (origen `TbHPSGrado`), `TbMotivoHPS`, `TbUsuarios`, `TbUsuariosEntidades`, `TbUsuariosHistoricos`, `TbUsuarioSolicitud` |
| Vinculadas a `Expedientes_datos` | `TbExpedientes`, `TbExpedientesLugaresEjecucion` |
| Relaciones | 4: 3 de negocio (`TbJustificaciones`→`TbSolicitudes` por `idjustificacion`, `TbResponsables`→`TbSolicitudes` por `Correo`/`emailResponsable`, `TbSolicitudes`→`TbSolicitudesFechas` por `IDSolicitud`) y 1 del panel de navegación de Access (`MSysNavPaneGroupCategories`→`MSysNavPaneGroups`), que no es una relación de negocio |
| Consultas guardadas | 3: `Consulta1`, `Consulta2` y `Solicitudes_alta` |
| Tablas sin clave primaria | Ninguna: las 11 tablas propias tienen clave primaria declarada |
| Campos multivalor, adjuntos u objeto OLE | Ninguno. Los tipos presentes son 1 (booleano), 3 (entero), 4 (entero largo), 8 (fecha), 10 (texto) y 12 (texto largo); ningún campo es complejo |
| Capitalización de la ruta de los vínculos | Mixta: 9 vínculos apuntan a `\\datoste\Aplicaciones_dys\…` y uno a `\\datoste\aplicaciones_dys\…` (`TbExpedientesLugaresEjecucion`). El código usa la variante en minúsculas |

El SQL de las tres consultas guardadas sí es legible por `QueryDef.SQL` (la tabla de sistema `MSysQueries` no lo es: «no tiene permiso para READ»):

| Consulta | SQL | Tablas |
|---|---|---|
| `Consulta1` | `SELECT TbSolicitudes.Tipo, TbSolicitudes.IDSolicitud, TbSolicitudes.DNI, TbSolicitudesFechas.FechaEnvioONS, TbHPS.F_Concesion, TbHPS.F_Caducidad FROM (TbSolicitudes INNER JOIN TbSolicitudesFechas ON …) INNER JOIN (TbUsuarios INNER JOIN TbHPS ON TbUsuarios.ID = TbHPS.IDUsuario) ON TbSolicitudes.DNI = TbUsuarios.DNI WHERE (((TbHPS.[TipoHPS])='Nacional'));` | Propias (`TbSolicitudes`, `TbSolicitudesFechas`) y de `HPST` (`TbUsuarios`, `TbHPS`) |
| `Consulta2` | `SELECT TbUsuarioSolicitud.*, TbUsuarioSolicitud.IDUsuarioSolicitud FROM TbUsuarioSolicitud WHERE (((TbUsuarioSolicitud.IDUsuario)=1) AND ((TbUsuarioSolicitud.IDSolicitud)=1));` | `TbUsuarioSolicitud`, que no existe en `HPST` |
| `Solicitudes_alta` | `SELECT TbSolicitudesFechas.IDSolicitud, TbSolicitudes.Tipo, TbSolicitudes.DNI, TbSolicitudesFechas.FechaEnvioONS, TbUsuariosEntidades.HPS_NAC_F_Concesion, TbUsuariosEntidades.HPS_NAC_F_Caducidad FROM (TbSolicitudes INNER JOIN TbUsuariosEntidades ON TbSolicitudes.DNI = TbUsuariosEntidades.DNI) INNER JOIN TbSolicitudesFechas ON …;` | Propias (`TbSolicitudes`, `TbSolicitudesFechas`) y de `HPST` (`TbUsuariosEntidades`) |

`Consulta1` y `Solicitudes_alta` justifican los vínculos con `HPST`: son consultas del backend que cruzan datos propios con datos de HPS. El front no nombra ninguna consulta guardada por su nombre (cero coincidencias de `Consulta1`, `Consulta2` y `Solicitudes_alta` en `src/`): su uso, si lo hay, es de otro consumidor.

### El backend del repositorio y el backend autoritativo no son el mismo fichero

| Hecho | ERD del repositorio (`00_HPS_SOLICITUDES`) | ERD del autoritativo (`C:\00repos\datos`) |
|---|---|---|
| Tamaño y fecha | 3 952 640 bytes, 14/04/2026 (commit `cfa1937`) | 13 631 488 bytes, 12/10/2026 |
| Tablas propias | 11 | 11 |
| Tablas vinculadas | 10 | 9: **falta `TbUsuarioSolicitud`** |
| Filas de `TbSolicitudes` | 220 | 245 |
| Filas de `TbLogsGeneral` | 1855 | 2058 |
| Filas de `TbResponsables` | 26 | 27 |
| Filas de `TbUltimoCambio` | 4 | 5 |
| Rutas de los vínculos | `\\datoste\…` (inalcanzables en este entorno) | `C:\00repos\datos\…` (alcanzables) |

El inventario de este documento es el del repositorio, que es el que declara la ficha. Las diferencias con el autoritativo son de fecha de corte, no de esquema, salvo en el vínculo `TbUsuarioSolicitud`, que solo existe en la copia del repositorio. Por eso el ERD de esta copia no puede leer los campos de las tablas vinculadas: los diez vínculos dan error de vínculo por ruta no alcanzable. Los campos de esas tablas se han leído del ERD del autoritativo, donde los vínculos apuntan a `C:\00repos\datos\`.

### `TbHPSGrado`: una copia idéntica de la tabla de HPS

El backend tiene su propia `TbHPSGrado` (13 filas, clave primaria `TipoHPS` + `Grado`) y además el vínculo `TbHPSGrado1`, que apunta a la `TbHPSGrado` de `HPST`. Las dos tablas tienen las mismas dos columnas, la misma clave primaria y las mismas 13 filas: la consulta de solo lectura sobre `C:\00repos\datos\HPST.accdb` devuelve exactamente los mismos 13 pares de valores que la del backend propio. El front lee la copia propia (`src/modules/constructor.bas:2884`, `getdb()` en `:2885`) y nunca el vínculo. El ERD del autoritativo confirma que `HPST` es el dueño de `TbHPSGrado` (13 filas).

## Relación con HPS

Las dos aplicaciones están acopladas en los dos sentidos, y no con el mismo peso: HPS depende de Solicitudes HPS para leer, y Solicitudes HPS depende de HPS para leer y para escribir.

| Sentido | Mecanismo | Evidencia |
|---|---|---|
| HPS → Solicitudes HPS | `HPST` vincula `TbSolicitudes` y `TbSolicitudesFechas` de `Solicitudes_HPS_datos`; el front de HPS las lee con `getdbSolicitudHPS()` | `docs/07-migracion/fase-0/hps.md:104` (los dos vínculos) y `:74` (`src/modules/VariablesEntorno.bas:645-671`, apertura en `:664`); `docs/07-migracion/fase-0/grafo-dependencias.md:16` (`HPST` toma 2 tablas de `Solicitudes_HPS_datos`) |
| Solicitudes HPS → HPS, lectura | El backend de esta aplicación vincula 8 tablas de `HPST` y el front lee cuatro de ellas por `getdbHPS()` | `TbHPS` (`constructor.bas:625`), `TbUsuarios` (`constructor.bas:2205`), `TbUsuariosHistoricos` (`constructor.bas:3162`), `TbMotivoHPS` (`MotivoHPS.cls:74-76`) |
| Solicitudes HPS → HPS, escritura | Cuatro tablas de `HPST` se actualizan desde esta aplicación | `TbUsuarios` (`FUNCIONES UTILES.bas:2449`), `TbUsuariosEntidades` (`FUNCIONES UTILES.bas:18`, `:23`, `:27`, `:2452`), `TbUsuariosHistoricos` (`MotivoHPS.cls:211`), `TbMotivoHPS` (`MotivoHPS.cls:41`, `FUNCIONES UTILES.bas:1430`, `:1466-1470`) |
| HPS → Solicitudes HPS, escritura | Ninguna medida: el front de HPS no escribe en `TbSolicitudes` ni en `TbSolicitudesFechas` | La ficha de HPS marca `TbSolicitudes` como «Propia y compartida», con 2 referencias del front |

Dos consecuencias para el resto del inventario:

- **`TbSolicitudesFechas` está bien ubicada donde está.** La ficha de HPS la marca como «Mal ubicada (candidata)» porque `HPST` la vincula sin que el front de HPS la use (`docs/07-migracion/fase-0/hps.md:141`, propuesta H8). La medición de esta ficha dice que la tabla la usa su dueña: 115 referencias en 4 ficheros del front de Solicitudes HPS (`src/classes/SolicitudFechasOperaciones.cls:215-218`, `:428`; `src/classes/SolicitudOperaciones.cls:314`, `:1255`, `:1337-1420`, `:1522-2190`, `:2598`; `src/modules/InformesOperaciones.bas:235`; `src/modules/constructor.bas:452`, `:685-1993`). Lo que sobra es el vínculo que la expone desde `HPST`, no la tabla.
- **La precondición de H2 no se cumple.** La propuesta H2 de la ficha de HPS («Retirar la tabla `TbUsuariosEntidades` del backend (352 filas) […] el código ya no las usa desde #33») condiciona la retirada a «que la aplicación hermana `HPS_SOLICITUDES` ya no escriba su tabla equivalente» (`docs/07-migracion/fase-0/hps.md:190`). La medición dice lo contrario: el front de Solicitudes HPS escribe esa tabla en cuatro sentencias (`FUNCIONES UTILES.bas:18`, `:23`, `:27`, `:2452`). H2 no debe aplicarse antes de decidir con las dos aplicaciones.

## Veredicto de propiedad por tabla

Referencias medidas en el código de producción del front (`src/` completo, 108 ficheros). Cada referencia se ha atribuido a la base de datos que abre la sentencia, no solo al nombre de la tabla.

### Tablas del backend

| Tabla | Clase | Referencias del front | Evidencia |
|---|---|---|---|
| `TbSolicitudes` | Propia | 166 en 13 ficheros | `src/classes/SolicitudOperaciones.cls:72-75`, `:318`, `:1292`; `src/classes/Solicitud.cls:1205`, `:1267`; `src/modules/constructor.bas:495`, `:684` |
| `TbSolicitudesFechas` | Propia | 115 en 4 ficheros | `src/classes/SolicitudFechasOperaciones.cls:215-218`; `src/modules/constructor.bas:685-1993` |
| `TbCorreosEnviados` | Propia | 25 en 10 ficheros | `src/modules/CorreoRepositorio.bas:21`; `src/modules/InformesOperaciones.bas:112`; `src/classes/CorreoOperaciones.cls:28`, `:141-144`, `:268-271` |
| `TbUsuarios` | Vinculada de `HPST` | 17 en 4 ficheros | `src/modules/constructor.bas:2205`, `:3126`, `:3494`, `:3762`; `src/classes/SolicitudOperaciones.cls:1178` |
| `TbMotivoHPS` | Vinculada de `HPST` | 15 en 3 ficheros | `src/classes/MotivoHPS.cls:41`, `:74-76`, `:150-152`; `src/modules/FUNCIONES UTILES.bas:1430-1432`, `:1457-1459` |
| `TbLogsGeneral` | Propia | 11 en 4 ficheros | `src/classes/LogGeneralOperaciones.cls:36`; `src/modules/Instalador.bas:18`, `:70-88`; `src/modules/constructor.bas:2692`, `:2736` |
| `TbResponsables` | Propia | 9 en 3 ficheros | `src/classes/ResponsableOperaciones.cls:56`, `:59`, `:153`; `src/modules/constructor.bas:2583-2586` |
| `TbUsuariosEntidades` | Vinculada de `HPST` | 7 en 1 fichero (4 de ellas escrituras) | `src/modules/FUNCIONES UTILES.bas:18-28`, `:2452` |
| `TbExpedientes` | Vinculada de `Expedientes_datos` | 5 en 1 fichero | `src/modules/constructor.bas:2306`, `:2526`, `:2529`, `:3869`, `:3873` |
| `TbJustificaciones` | Propia | 5 en 1 fichero | `src/modules/JustificacionRepositorio.bas:13`, `:21`, `:41`, `:65`, `:90` |
| `TbUltimoCambio` | Propia | 4 en 3 ficheros | `src/classes/UltimoCambioOperaciones.cls:33`; `src/modules/constructor.bas:3398`, `:3403` |
| `TbUsuariosHistoricos` | Vinculada de `HPST` | 4 en 2 ficheros | `src/classes/MotivoHPS.cls:211`; `src/modules/constructor.bas:3162`, `:3180`, `:3248` |
| `TbConfiguracion` | Propia | 3 en 3 ficheros | `src/classes/Configuracion.cls:45`; `src/classes/ConfiguracionOperaciones.cls:141`; `src/modules/constructor.bas:282` |
| `TbHPSGrado` | Propia | 1 en 1 fichero | `src/modules/constructor.bas:2884` |
| `TbHPS` | Vinculada de `HPST` | 1 en 1 fichero | `src/modules/constructor.bas:625` |
| `TbHPSGrado1` | Vinculada de `HPST` | 0 | Cero coincidencias en `src/`. El front lee la copia propia `TbHPSGrado` |
| `TbHPSEquivalencia` | Vinculada de `HPST` | 0 | Cero coincidencias en `src/`. Su dueño es la aplicación de HPS |
| `TbExpedientesLugaresEjecucion` | Vinculada de `Expedientes_datos` | 0 | Cero coincidencias en `src/`. Su dueño es la aplicación de Expedientes |
| `TbUsuarioSolicitud` | Vinculada de `HPST` | 0 | Cero coincidencias en `src/`, y **`HPST` no tiene esa tabla** (18 propias y 9 vinculadas, ninguna con ese nombre): el vínculo está roto y la consulta `Consulta2` que lo usa no puede ejecutarse |
| `Copia de TbExpedientes` | Propia | 0 | 70 filas, 9 campos, clave primaria `IDExpediente`. Copia de la tabla de `Expedientes_datos` que sí usa el front (`constructor.bas:2306`) |
| `TbLogs` | Propia | 0 | 0 filas, 6 campos, clave primaria `IDLog`. La tabla viva es `TbLogsGeneral` |

### Tablas del front

| Tabla | Clase | Referencias del front | Evidencia |
|---|---|---|---|
| `TbCorreosEnviadosParaPruebas` | Propia del front | 0 | 0 filas, 18 campos, clave primaria `IDCorreo`. Copia recortada del esquema de `TbCorreosEnviados` (21 campos). Cero coincidencias en `src/`, incluidos los formularios |

### Tablas de otros backends que el front usa sin vínculo en su backend

| Tabla | Backend dueño | Evidencia |
|---|---|---|
| `TbSuministradores`, `TbExpedientesSuministradores` | `Expedientes_datos` | `src/modules/constructor.bas:14-16`, `:2353`, `:2404`. El backend de esta aplicación no las vincula |
| `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` | `Lanzadera_Datos` | `src/modules/constructor.bas:87`, `:118-123`, `:331`. El backend de esta aplicación no vincula la lanzadera |

### Referencias a tablas que no existen

| Referencia | Código | Lectura |
|---|---|---|
| `Motivo_HPS` como columna de `TbMotivoHPS` | `src/modules/FUNCIONES UTILES.bas:1430-1432` | `INSERT INTO TbMotivoHPS ( Motivo_HPS ) SELECT distinct TbMotivoHPS.Motivo_HPS FROM TbMotivoHPS;` con `getdbHPS().Execute`. La tabla de `HPST` tiene una sola columna, `MotivoHPS` (verificado en el ERD del autoritativo y en el ERD del proyecto del 16/04/2026). La sentencia no puede ejecutarse. La función `RegistrarTbMotivoHPS` no tiene ningún llamante |
| `TbConfiguracion` en `CurrentDb()` | `src/classes/Configuracion.cls:45` | `DameID("TbConfiguracion", "ID", , m_Error)` omite la base de datos, y `DameID` cae en `CurrentDb()` (`src/modules/FUNCIONES UTILES.bas:772-774`), que es el front: allí la tabla no existe. Afecta al alta de la fila de configuración (`src/classes/ConfiguracionOperaciones.cls:141-146`), que solo se ejecuta cuando no hay ninguna fila |

### Patrones de riesgo en el VBA del front

| Patrón | Medida |
|---|---|
| Generadores `Max + 1` | `DameID` (`src/modules/FUNCIONES UTILES.bas:760-796`, `SELECT Max(tabla.campo) … FROM tabla`) con 8 puntos de llamada sobre 7 tablas: `TbConfiguracion` (`Configuracion.cls:45`), `TbCorreosEnviados` (`Correo.cls:136`, `FUNCIONES UTILES.bas:652`), `TbLogsGeneral` (`LogGeneral.cls:126`), `TbResponsables` (`Responsable.cls:32`), `TbSolicitudes` (`Solicitud.cls:1205`), `TbUsuarios` (`SolicitudOperaciones.cls:1184`), `TbUltimoCambio` (`UltimoCambio.cls:33`) |
| `LIKE '*…'` en SQL | 9 apariciones: `src/modules/constructor.bas:2357`, `:2530-2532`, `:2588`, `:3763-3765`, `:3819` |
| Fechas escritas como literal `#…#` | 13 sentencias con el formato `mm/dd/yyyy` incrustado: `src/modules/constructor.bas:628`, `:802`, `:814`, `:877`, `:889`, `:1173`, `:1186`, `:1249`, `:1262`; `src/classes/CorreoOperaciones.cls:268`; `src/classes/SolicitudOperaciones.cls:1255`, `:1523`, `:1688` |
| `Date()` dentro de SQL | Ninguna: la fecha se compone con `Format(Date, "mm/dd/yyyy")` antes de concatenarla |
| `Seek`, `dbOpenTable` | Ninguno |
| `dbSeeChanges` | Ninguno: ningún recordset se abre con actualización en vivo |
| `Option Compare Database` | En 72 de los 74 ficheros `.bas`/`.cls`; falta en `src/classes/SolicitudFechas.cls` y `src/modules/JsonConverter.bas` |
| Escrituras en el backend de otra aplicación | 10 sentencias sobre tablas de `HPST` y 1 sobre `Correos_datos` (ver «Tablas de otros backends que consume el front») |

## Propuestas para después del descubrimiento global

Ninguna se aplica: la puerta de descubrimiento global sigue cerrada (#1576). El propietario borra en el legacy; aquí solo hay evidencia, comprobación previa y reversión. Los grupos son los de la fase 1 de la épica: (a) imprescindible antes de migrar, (b) recomendable si es barato, (c) aplazar a la refactorización final.

| Id | Propuesta | Grupo (fase 1) | Comprobación previa | Reversión |
|---|---|---|---|---|
| S1 | Poner en cuarentena las dos tablas propias del backend sin ninguna referencia en el front: `TbLogs` (0 filas) y `Copia de TbExpedientes` (70 filas). Renombrar a `zzBAJA_<tabla>`. | (a) | Cruzar los vínculos de todos los backends (tarea T6) y confirmar en la copia de producción que ninguna consulta guardada las usa. Medir que ninguna otra herramienta (Excel, informe programado, web) las lee. | Quitar el prefijo `zzBAJA_`. |
| S2 | Retirar el vínculo roto `TbUsuarioSolicitud` y la consulta guardada `Consulta2` que lo usa con los valores fijos `1` y `1`. | (a) | Confirmar en `HPST` que la tabla no existe (medido: 18 tablas propias y 9 vinculadas, ninguna con ese nombre) y que ningún informe externo usa `Consulta2`. | Recrear el vínculo y la consulta con su definición original. |
| S3 | Resolver la duplicidad de `TbHPSGrado`: el backend tiene una copia de la tabla de `HPST` con las mismas 13 filas y la misma clave primaria, y además un vínculo `TbHPSGrado1` que nadie usa. Propuesta: dejar una sola fuente de verdad en `HPST` y que el front lea de allí. | (a) | Confirmar con el dueño de HPS que `TbHPSGrado` no va a cambiar de backend, y medir que las 13 filas de ambas copias coinciden (medido hoy: coinciden). | Volver a la copia local del backend y restaurar el vínculo `TbHPSGrado1`. |
| S4 | Declarar la propiedad de las cuatro tablas de `HPST` que esta aplicación escribe (`TbUsuarios`, `TbUsuariosEntidades`, `TbUsuariosHistoricos`, `TbMotivoHPS`). Hoy son de HPS y las escribe Solicitudes HPS, lo que impide cerrar H2 de la ficha de HPS. | (a) | Decidir con el propietario y con el responsable de HPS quién es el dueño de cada tabla y qué aplicación escribe cuál. Hasta esa decisión, H2 de la ficha de HPS no se aplica. | Sin cambios: es una decisión, no una modificación. |
| S5 | Retirar la credencial literal de las cinco cadenas `OpenDatabase` (`src/modules/Variables Globales.bas:449`, `:478`, `:512`, `:542`, `:599`) y leerla de la configuración. | (a) | La misma credencial abre los cinco backends y los binarios del ecosistema. Comprobar que no queda ninguna otra copia del literal en `src/` y si es la misma que usan otros fronts. | Restaurar las líneas originales. |
| S6 | Revisar los 10 vínculos del backend, que apuntan a la ruta de producción `\\datoste\…`: son una vía de escritura sobre producción desde el front y desde cualquier proceso que use el backend. | (a) | Contar los errores de vínculo de los 10 con la misma comprobación que se hizo con los de Brass, y confirmar en runtime que ninguno se usa. | Restituir cada vínculo con su definición original (`TableDef.Connect`). |
| S7 | Retirar del backend los dos vínculos que ninguna aplicación usa: `TbExpedientesLugaresEjecucion` (de `Expedientes_datos`) y `TbHPSEquivalencia` (de `HPST`). Por R-2 la tabla pertenece a la aplicación que la usa. No se borra nada: solo se deja de exponer desde el backend que no la usa. | (b) | Grafo global (tarea T6) para confirmar que ningún proceso lee esas tablas a través de este backend: un vínculo no usado por el front puede seguir usándose desde la web o desde procesos programados. | Recrear el vínculo con su definición original. |
| S8 | Añadir al backend los vínculos que el front usa por código, o mover esas lecturas al backend: `TbSuministradores` y `TbExpedientesSuministradores` de `Expedientes_datos`, `TbUsuariosAplicaciones` y `TbUsuariosAplicacionesPermisos` de `Lanzadera_Datos`. | (b) | Confirmar que las cuatro tablas existen en sus backends con clave primaria y que ningún vínculo del backend las aporta ya. La lectura de `TbUsuariosAplicaciones` es la autoridad de permisos: no puede quedar en un solo front. | Eliminar los vínculos añadidos. |
| S9 | Separar los dos `TbCorreosEnviados`: el propio del backend, que la aplicación usa para su listado, y el de `Correos_datos`, donde se registra el correo enviado con `Aplicacion = "HPS_SOLICITUD"`. Proposta: que el registro del correo salga por el mismo camino que el resto de las aplicaciones. | (b) | Verificar en `Correos_datos` si el resto de aplicaciones escriben en la misma tabla con su identificador de aplicación, y qué columnas son obligatorias. | Volver a la ruta actual de escritura. |
| S10 | Corregir `RegistrarTbMotivoHPS` (`src/modules/FUNCIONES UTILES.bas:1422-1440`): inserta y selecciona la columna `Motivo_HPS`, que no existe en `TbMotivoHPS` (la tabla tiene una sola columna, `MotivoHPS`). La función no tiene llamantes, así que el fallo no se ve. | (b) | Reproducir en copia la llamada a la función y comprobar el error 3075 (columna no encontrada). | Restaurar la línea original. |
| S11 | Corregir la generación del identificador de `TbConfiguracion`: `DameID` sin base de datos (`src/classes/Configuracion.cls:45`) cae en `CurrentDb()`, que es el front, donde la tabla no existe. Afecta al alta de la fila de configuración. | (b) | En copia, borrar la fila de `TbConfiguracion` y arrancar el alta desde `ConfiguracionOperaciones.cls:141-146` para reproducir el fallo. | Restaurar la llamada original. |
| S12 | Retirar del front la tabla `TbCorreosEnviadosParaPruebas`: 0 filas, 0 referencias en `src/`, 18 campos que son una copia recortada del esquema de `TbCorreosEnviados`. | (b) | Abrir el front en copia y comprobar que ni el arranque ni el formulario de correo de prueba la leen. Revisar el SQL de las consultas guardadas del front (no hay ninguna). | Reponer el binario del front anterior. |
| S13 | Fijar una sola versión declarada: el `.ini` de `main` declara `2025-042` (`Solicitudes_HPS.ini:2`) y el informe de corrección declara la versión `2026-001` (`docs/Informes/Informe_Correccion_ID256_2026-001.md:4`). La lanzadera no registra ediciones para `IDAplicacion=22`, así que no hay una tercera versión que contrastar. | (b) | Confirmar con el propietario qué versión está desplegada en el servidor y corregir una de las dos declaraciones. | Restaurar el `.ini` anterior. |
| S14 | Revisar los datos personales incrustados en el front: 39 listas de valores de diseño con nombres de personas, empresas y correos, y un destinatario literal con dos buzones personales en el aviso de error (`src/modules/FUNCIONES UTILES.bas:649`). | (b) | Medir qué formularios dependen de esas listas y qué se rompe al vaciarlas; revisar la política de datos personales aplicable (D92) con el propietario. | Reponer el binario del front anterior. |

Lo que **no** se toca ahora: los 10 vínculos del backend más allá de la propuesta S6, las 9 tablas propias del backend que el front usa, los 8 vínculos con `HPST` más allá de las propuestas S3 y S7, las 3 relaciones de negocio del backend y las 3 consultas guardadas.

## Anexo: controles con origen de datos

Los 39 controles con `RowSource` con valor, todos `RowSourceType = "Value List"`. Ninguno contiene SQL y ninguno consulta una tabla: son listas literales que el runtime sustituye por valores. Se listan porque son el inventario de lo que el front tiene escrito dentro, y porque varias de ellas son instantáneas de datos de negocio y de personas.

| Formulario | Control | Valor de `RowSource` (extracto) |
|---|---|---|
| `Form_FormAltaDatosSolicitante` | `IDEmpresaUsuario`, `IDEmpresaTramitadora`, `CorreosAutomaticos`, `TIPO` | Niveles de clasificación; `Sí`/`No`; `Alta`/`Traspaso` |
| `Form_FormCambioTipo` | `TipoFuturo` | `Alta`/`Renovación`/`Elevación Grado` |
| `Form_FormConfiguracion` | `AutocancelacionPreMARGA`, `AutocancelacionMARGA`, `CorreosAutomaticos` | `Sí`/`No` |
| `Form_FormCorreosGestion` | `ListaFiltrados`, `Estado`, `solicitud` | Columnas de correo; estados; listado de solicitudes con personas |
| `Form_FormEntidades` | `ListaEntidades` | Catálogo de entidades con nombres de empresa |
| `Form_FormEnvioCorreo` | `Programado` | `Sí`/`No` |
| `Form_FormExpedientesBusqueda` | `ListaFiltrados` | Columnas de expediente con datos de un expediente real |
| `Form_FormPlantillasHTML` | `ListaPlantillas` | Catálogo de plantillas |
| `Form_FormResponsablesGestion` | `ListaFiltrados` | Columnas de responsable |
| `Form_FormSolicitudAlta` | `TIPO` | Tipos de solicitud |
| `Form_FormSolicitudRenovacionDatos` | `IDEmpresaUsuario`, `IDEmpresaTramitadora`, `CorreosAutomaticos`, `Gestor`, `Motivo_HPS` | Empresas, `Sí`/`No`, nombres de tramitadores, 58 motivos de HPS |
| `Form_FormSolicitudesAltaTraspasoDatos` | `IDEmpresaUsuario`, `IDEmpresaTramitadora`, `CorreosAutomaticos`, `Gestor`, `Motivo_HPS`, `emailResponsable` | Empresas, `Sí`/`No`, tramitadores, motivos, correos de responsables |
| `Form_FormSolicitudesGestion` | `ListaFiltrados`, `TIPO`, `Estado`, `Gestor`, `IDEmpresaUsuario`, `IDExpediente` | Columnas, tipos, estados, tramitadores, empresas, expedientes |
| `Form_FormTareasTramitadorPendientes` | `ListaFiltrados`, `Gestor` | Columnas y tramitadores |
| `Form_FormUsuariosHPSBusqueda` | `ListaFiltrados` | Columnas de usuario |
| `Form_frmBuscarJustificacion` | `lstResultados`, `cboFiltroEstado` | Columnas y filtro de estado |

Las 50 declaraciones de diseño de `RowSourceType = "Value List"` son 39 con valor y 11 sin él. Las otras dos apariciones de `RowSource =` en los exports están en el bloque de código, no en el diseño. En tiempo de ejecución hay 64 asignaciones a `.RowSource`: todas con lista de columnas o cadena vacía (`src/forms/Form_FormExpedientesBusqueda.cls:183`, `:218`; `src/forms/Form_FormResponsablesGestion.cls:42`), ninguna con SQL.

## Lo que no se pudo verificar

1. **Macros embebidas en el front.** No se pudo enumerar `Access.Application.AllMacros`: la automatización por COM sobre una copia del binario se quedó colgada en este entorno y se cerró el proceso. La evidencia disponible es que los 108 ficheros de `src/` no contienen ninguna referencia a macros y que el export cubre los 34 formularios del binario.
2. **Los vínculos en producción.** Los diez vínculos del backend apuntan a `\\datoste\…`, que no es alcanzable desde este entorno: el ERD del repositorio no puede leer sus campos. Los campos se han leído del ERD del autoritativo, donde los vínculos apuntan a `C:\00repos\datos\`. La definición real en producción no se ha leído.
3. **El estado de la UAT.** No hay documento de UAT en el repositorio ni en el repositorio OpenSpec externo, y la copia local de `Lanzadera_Datos` no registra ediciones ni aperturas para `SOLICITUDES HPS`. La evidencia de release es el informe de corrección, no un ciclo de pruebas.
4. **El binario desplegado.** El `.accdb` del repositorio es el blob de `main`, pero no se comprobó que sea el mismo fichero que el instalado en el servidor de aplicaciones.
5. **El runtime.** No se ejecutó la aplicación. La traza de uso de tablas es estática sobre `src/` y no cubre referencias construidas en tiempo de ejecución fuera de las cadenas de SQL revisadas, ni el SQL de las funciones de la lanzadera y de la web.
6. **Los consumidores ajenos a Access.** No se ha revisado ningún consumidor externo (Excel, informes programados, la web en construcción) de las tablas propuestas en cuarentena, ni de `TbSolicitudesFechas`, `TbHPSGrado` ni `TbConfiguracion`.
7. **Los recuentos de filas.** Proceden de copias en disco, no de producción: el backend del repositorio va 25 solicitudes y 203 líneas de log por detrás del autoritativo.
