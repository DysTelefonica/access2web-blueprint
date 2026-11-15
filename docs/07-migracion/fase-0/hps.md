# Fase 0: HPS (`HPS` / `HPST`)

Épica [`docs/07-migracion/epica-coexistencia-postgresql.md`](../epica-coexistencia-postgresql.md) (fase 0). Seguimiento: issue #1581 (descubrimiento global de todas las herramientas); no existe todavía una issue propia de la ficha de HPS.
Base de análisis: rama `main`, SHA `dc194f4` (`origin/main`), repositorio `00_HPS`. Su árbol es exactamente el de la RC1 liberada (tag `2026-005`), según se detalla más abajo.
Fuentes: ERD del front y del backend con dysflow (solo lectura, sobre copia) y trazado del código con CodeGraph sobre el índice del worktree, completado con búsqueda de texto sobre `src/` para las cadenas de SQL y de conexión.

Las rutas de código de esta ficha son relativas al repositorio `00_HPS` en el SHA indicado.

## Base de análisis

| Hecho | Evidencia |
|---|---|
| La RC1 ya está aprobada y liberada: no hay una RC1 con UAT pendiente | `docs/ciclos/RELEASE-2026-005.md:17` («Rama integrada \| `fix/2026-005-rc1`, tras la UAT aprobada el 2026-10-26») y `:87` («Escenarios UAT de la rc1: 24, aprobados el 2026-10-26») |
| La RC1 está integrada en `main` con su binario, y `main` es la base | `docs/ciclos/RELEASE-2026-005.md:24` («toma exactamente el árbol de `fix/2026-005-rc1`, binario incluido»); PR #47 «release: 2026-005», fusionada el 2026-10-26; tag `2026-005` en `4761aef`; `origin/fix/2026-005-rc1` en `4467369` |
| El contenido de `main` desde 2026-004 no forma parte de esta versión | `docs/ciclos/RELEASE-2026-005.md:26-29`: los commits previos de `main` «quedan en el historial, pero su contenido no forma parte de esta versión» |
| El worktree `00_HPS-worktrees/2026-005-rc1` está huérfano | `00_HPS-worktrees/2026-005-rc1/.git:1` apunta a `…/00_HPS/.git/worktrees/2026-005-rc1`, directorio que no existe; `git worktree list` en `00_HPS` solo devuelve `main` y un `candidate-views` |
| Su `src/` es el de `main`, salvo una línea espuria | Comparación blob a blob de los 136 ficheros de `src/` entre `dc194f4` y el worktree: solo difiere `src/forms/Form_FormInicial.form.txt`, por la línea `egc_live_…` añadida al final (línea 1335 del worktree; el fichero en `main` tiene 1334) |
| Su binario `HPS.accdb` es una copia local modificada, no el publicado | SHA-256 `ef9674e9…` en el worktree frente a `1212c01e…` del blob publicado (`git show 2026-005:HPS.accdb`); el `HPS.accdb` de `main` también está modificado en el árbol de trabajo (`git status`: `M HPS.accdb`) |
| `origin/fix/2026-005-rc1-uat-gaps` no es base de estudio | Rama `c718589` (2026-10-26 08:23), anterior a la fusión de la release (13:28) y no ancestro de `main`; su `src/` difiere de `main` en 38 ficheros y no tiene PR abierta |

Consecuencia: la base de este inventario es `main` en `dc194f4`, cuyo árbol es el de la RC1 aprobada, más un commit de documentación (`dc194f4`). La regla del propietario («si hay RC1 con UAT pendiente, se estudia la RC1») no se activa en HPS: la UAT de la RC1 está aprobada y su contenido está en `main`. El inventario de abajo describe ese árbol; el ERD del front se midió sobre la copia de trabajo del binario en el worktree huérfano, que no coincide con el blob publicado (dato declarado en «Lo que no se pudo verificar»).

## Estructura del front

| Hecho | Valor |
|---|---|
| Tablas del front | 15: 10 propias y 5 vinculadas |
| Tablas propias del front | `Copia de TbUsuariosEntidades`, `TbCursosLocal`, `TbDatosLocal`, `TbDatosLocalParaIndicadores`, `tblInfo`, `tblSettings`, `TbSuministradoresLocal`, `TbUsuariosHistoricosLocal`, `TbUsuariosSICALocal`, `TbUsuariosSICALocalParaIndicadores` |
| Destino de los 5 vínculos | 2 a `\\datoste\APLICACIONES_DYS\Aplicaciones PpD\EXPEDIENTES\Expedientes_datos.accdb` (`TbExpedientes`, `TbSuministradores`) y 3 a `\\datoste\APLICACIONES_DYS\Aplicaciones PpD\HPS\HPST.accdb` (`TbHPS`, `TbUsuariosEntidades`, `TbUsuariosHistoricos`). Los cinco dan error de vínculo en este entorno: la ruta de producción no es alcanzable |
| Vínculos a `Lanzadera_Datos` en el front | Ninguno. La identidad y los permisos se leen abriendo la base directamente (apartado «Aperturas directas») |
| Formularios | 36, con su código en `src/forms/` (36 `.cls` y 36 `.form.txt`) |
| Formularios con `RecordSource` de diseño | 3, y los tres leen cachés locales: `src/forms/Form_FormInicial03ConsultasDatos.form.txt:289` (`SELECT TbDatosLocal.* FROM TbDatosLocal`), `src/forms/Form_FormInicial07UsuariosSICADatos.form.txt:68` (`TbUsuariosSICALocal`) y `src/forms/Form_FormInicial08UsuariosHistoricosDatos.form.txt:83` (`TbUsuariosHistoricosLocal`). Hay 3 asignaciones más en runtime: `Form_FormInicial03ConsultasDatos.form.txt:3987`, `Form_FormInicial07UsuariosSICA.form.txt:1602` y `Form_FormInicial08UsuariosHistoricos.form.txt:1532` |
| Controles con `RowSourceType = Table/Query` | Ninguno. Todos los `RowSourceType` son listas de valores literales (p. ej. `src/forms/Form_FormInicial03ConsultasGrado.form.txt:339`) |
| Controles con `ControlSource` | 4 formularios: `Form_FormInicial03ConsultasDatos.form.txt` (103), `Form_FormInicial07UsuariosSICA.form.txt` (1), `Form_FormInicial07UsuariosSICADatos.form.txt` (14) y `Form_FormInicial08UsuariosHistoricosDatos.form.txt` (20). Todos cuelgan del `RecordSource` local anterior |

El resto de formularios está desconectado, como es la norma salvo Brass: no tienen `RecordSource` de diseño y se rellenan en tiempo de ejecución. Los procedimientos de carga son los que mandan, no las propiedades de diseño:

| Qué se carga | Procedimiento | Evidencia |
|---|---|---|
| Listado de datos de usuario | `Form_Load` fija el `RecordSource` sobre la caché local | `src/forms/Form_FormInicial03ConsultasDatos.cls:27-33` |
| Listado de SICA | Filtro aplicado por código al subformulario | `src/forms/Form_FormInicial07UsuariosSICA.cls:117-128` |
| Listado de históricos | Filtro aplicado por código al subformulario | `src/forms/Form_FormInicial08UsuariosHistoricos.cls:125-140` |
| Pestañas de consultas | Consultas al backend por `Constructor` y relleno de combos con `AddItem` | `src/forms/Form_FormInicial03ConsultasEmpresaUsuario.cls:130-172` |
| Caché de usuarios | `cache_usuarios_regenerar` recorre `TbUsuarios` del backend y escribe `TbDatosLocal` | `src/modules/cacheUsuario.bas:17`, `:36`, `:113` |
| Caché de históricos | Reconstrucción completa desde el backend | `src/modules/Mod_Sincronizacion_Historico.bas:131-136` |
| Caché de suministradores | Reconstrucción completa desde `Expedientes_datos` | `src/modules/cacheSuministrador.bas:17`, `:26`, `:89` |

### Cachés locales del front y su coherencia con el backend

El plan de arranque declara el contrato de las cachés locales en `src/modules/Mod_StartupCacheInitialization.bas:19-23`: solo cuatro tienen regenerador y cobertura (`TbDatosLocal`, `TbUsuariosSICALocal`, `TbUsuariosHistoricosLocal`, `TbSuministradoresLocal`). Las otras seis tablas locales del front quedan fuera de ese contrato.

| Tabla local del front | Origen real | Filas (front en disco) | Filas de la tabla de origen (backend) | Estado |
|---|---|---|---|---|
| `TbDatosLocal` | `TbUsuarios` de `HPST` | 348 | 349 (`TbUsuarios`) | Caché viva, regenerada; la diferencia de una fila está sin explicar |
| `TbUsuariosSICALocal` | `TbUsuariosSICA` de `HPST` | 50 | 50 | Caché viva, regenerada |
| `TbUsuariosHistoricosLocal` | `TbUsuariosHistoricos` de `HPST` | 249 | 248 (`TbUsuariosHistoricos`) | Caché viva, regenerada; la diferencia de una fila está sin explicar |
| `TbSuministradoresLocal` | `TbSuministradores` de `Expedientes_datos` | 73 | No medible: el vínculo no se resuelve en este entorno | Caché viva, regenerada |
| `TbDatosLocalParaIndicadores` | Retirada en #32 | 344 | — | El código ya no la alimenta (`Mod_StartupCacheInitialization.bas:19-23` no la cubre) |
| `TbUsuariosSICALocalParaIndicadores` | Retirada en #32 | 50 | — | Igual que la anterior |
| `Copia de TbUsuariosEntidades` | Retirada en #33 | 99 | 352 (`TbUsuariosEntidades` en `HPST`) | Copia obsoleta: los mismos 102 campos, 99 filas frente a 352 |
| `TbCursosLocal` | — | 0 | — | Sin uso |
| `tblInfo` | — | 1 | — | Sin uso |
| `tblSettings` | — | 9 | — | Sin uso |

Sobre la pregunta de partida: las diez tablas locales del front **no** replican usuarios de la lanzadera. La identidad y los permisos viven solo en `Lanzadera_Datos` y el front los lee por apertura directa (`src/modules/VariablesEntorno.bas:545-570`), sin copia local. Lo que el front replica son datos de su propio backend y de `Expedientes_datos`, en cuatro cachés con regenerador declarado, y en seis tablas sin contrato: dos retiradas del ciclo RC1 (#32 y #33) y cuatro con cero referencias en todo el repositorio. La coherencia que debe quedar fijada es: una fuente de verdad por tabla (el backend para los datos de HPS, la lanzadera para la identidad, `Expedientes_datos` para los suministros), el front como caché declarada y regenerable, y ninguna tabla local sin contrato.

## Aperturas directas

| Apertura | Código | Observación |
|---|---|---|
| Backend propio `HPST.accdb` | `src/modules/VariablesEntorno.bas:576-604` (`getdb()`), apertura en `:597` | Ruta local `…\000datoslocal\HPST.accdb` (`:586`) o remota `\\datoste\aplicaciones_dys\Aplicaciones PpD\HPS\HPST.accdb` (`:588`). La credencial va literal en el código (`MS Access;PWD=***`) |
| `Lanzadera_Datos.accdb` | `src/modules/VariablesEntorno.bas:545-570` (`getdbLanzadera()`), apertura en `:564` | Rutas `…\000datoslocal\Lanzadera_Datos.accdb` (`:553`) y `…\0Lanzadera\Lanzadera_Datos.accdb` (`:555`) |
| `Correos_datos.accdb` | `src/modules/VariablesEntorno.bas:609-638` (`getdbCorreo()`), apertura en `:633` | Rutas `…\000datoslocal\Correos_datos.accdb` (`:618`) y `…\00Recursos\Correos_datos.accdb` (`:620`) |
| `Solicitudes_HPS_datos.accdb` | `src/modules/VariablesEntorno.bas:645-671` (`getdbSolicitudHPS()`), apertura en `:664` | Rutas `…\000datoslocal\Solicitudes_HPS_datos.accdb` (`:653`) y `…\SOLICITUDES HPS\Solicitudes_HPS_datos.accdb` (`:655`) |
| `Expedientes_datos.accdb` | `src/modules/VariablesEntorno.bas:674-699` (`getdbExpedientes()`), apertura en `:693` | Rutas `…\000datoslocal\Expedientes_datos.accdb` (`:682`) y `…\EXPEDIENTES\Expedientes_datos.accdb` (`:684`) |
| `DoCmd.OpenTable`, `TransferDatabase`, `DoCmd.OpenQuery` | Ninguno | Cero coincidencias en `src/`, incluidos los módulos de test |
| Macros `AutoExec` / `AutoOpen` / `AutoClose` | Ninguna | Cero coincidencias en `src/` |

Las cinco aperturas comparten una única credencial literal en la línea de `OpenDatabase` (`src/modules/VariablesEntorno.bas:564`, `:597`, `:633`, `:664`, `:693`).

El front también escribe DDL en el backend: `src/classes/Entorno.cls:274` comprueba si `TbConfiguracionUsuariosAccesoLocal` existe en `HPST` y, si no, la crea (`Entorno.cls:345`) y la siembra (`Entorno.cls:417`). La tabla está en el backend, pero la crea el front.

## Tablas de otros backends que consume el front

| Tabla | Backend | Código | Enlace en `HPST` |
|---|---|---|---|
| `TbUsuariosAplicaciones` | `Lanzadera_Datos` | `src/modules/Constructor.bas:1430`, `:1468`; `src/classes/Usuario.cls:186` | Sí, vinculado en `HPST` |
| `TbUsuariosAplicacionesPermisos` | `Lanzadera_Datos` | `src/modules/Constructor.bas:1625`; `src/classes/UsuarioAplicacionPermisos.cls:263` | Sí, vinculado en `HPST` |
| `TbCorreosEnviados` | `Correos_datos` | `src/classes/Correo.cls:42`, `:154` | Sí, vinculado en `HPST` |
| `TbSolicitudes` | `Solicitudes_HPS_datos` | `src/modules/Constructor.bas:3124`, `:3170-3174` | Sí, vinculado en `HPST` |
| `TbExpedientes` | `Expedientes_datos` | `src/modules/Constructor.bas:2804`, `:3218` | Sí, vinculado en `HPST` |
| `TbExpedientesConEntidades` | `Expedientes_datos` | `src/modules/Constructor.bas:196` | Sí, vinculado en `HPST` |
| `TbSuministradores` | `Expedientes_datos` | `src/modules/cacheSuministrador.bas:17`; `src/forms/Form_FormInicial03ConsultasEmpresaUsuario.cls:154-158` | Sí, vinculado en `HPST` |
| `TbExpedientesSuministradores` | `Expedientes_datos` | `src/modules/Constructor.bas:83-90` | **No**: no figura entre los 9 vínculos de `HPST` |

`TbExpedientesSuministradores` es la única tabla de otro backend que el front lee sin pasar por su backend: la consulta se hace con `getdbExpedientes()` (`src/modules/Constructor.bas:90`) sobre una tabla que `HPST` no vincula.

## Backend `HPST.accdb`

| Hecho | Valor |
|---|---|
| Tablas | 27: 18 propias y 9 vinculadas |
| Propias | `TbAnexosUsuariosHistoricos`, `TbAnexosUsuariosHPS`, `TbAnexosUsuariosSICA`, `TbAuxCursos`, `TbConfiguracionUsuariosAccesoLocal`, `TbConsultas`, `TbHPS`, `TbHPSEquivalencia`, `TbHPSGrado`, `TbJuridicasContratacion`, `TbMotivoHPS`, `TbObservaciones`, `TbObservacionesHistoricas`, `TbUsuarioAnexos`, `TbUsuarios`, `TbUsuariosEntidades`, `TbUsuariosHistoricos`, `TbUsuariosSICA` |
| Vínculos | 4 a `Expedientes_datos.accdb` (`TbExpedientes`, `TbExpedientesConEntidades`, `TbExpedientesLugaresEjecucion`, `TbSuministradores`), 2 a `Solicitudes_HPS_datos.accdb` (`TbSolicitudes`, `TbSolicitudesFechas`), 2 a `Lanzadera_Datos.accdb` (`TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos`), 1 a `Correos_datos.accdb` (`TbCorreosEnviados`) |
| Relaciones | 8 en el ERD: 7 de negocio (`TbUsuarios`→`TbHPS`, `TbUsuarios`→`TbObservaciones`, `TbUsuarios`→`TbAnexosUsuariosHPS`, `TbUsuarios`→`TbUsuarioAnexos`, `TbUsuariosHistoricos`→`TbObservacionesHistoricas`, `TbUsuariosHistoricos`→`TbAnexosUsuariosHistoricos`, `TbUsuariosSICA`→`TbAnexosUsuariosSICA`) y 1 del panel de navegación de Access (`MSysNavPaneGroupCategories`→`MSysNavPaneGroups`), que no es una relación de negocio |
| Consultas guardadas | 8: `Consulta1` a `Consulta7` y `ConsultaDatosRealesNoCacheados` |
| Tablas sin clave primaria | 1: `TbCorreosEnviados`, que además es vinculada |
| Campos multivalor, adjuntos u objeto OLE | Ninguno. Los tipos presentes son 1 (booleano), 3 (entero), 4 (entero largo), 7 (doble), 8 (fecha), 10 (texto) y 12 (texto largo) |

El ERD del backend se midió sobre la copia `C:\00repos\datos\HPST.accdb`, cuyos vínculos apuntan a copias locales de los otros backends (`C:\00repos\datos\…`) y no a la ruta de producción. La definición de los vínculos en producción no se pudo leer en este entorno.

El SQL de las 8 consultas guardadas no se pudo leer: la tabla de sistema `MSysQueries` devuelve «no tiene permiso para READ». El front no nombra ninguna consulta guardada por su nombre (cero coincidencias de `Consulta1`…`Consulta7` y `ConsultaDatosRealesNoCacheados` en `src/`), de modo que su uso, si lo hay, es de otro consumidor.

## Veredicto de propiedad por tabla

Referencias medidas en el código de producción del front (ficheros de test excluidos). Es el coste de cambio de la fase 1 para cada tabla.

### Tablas del backend

| Tabla | Clase | Referencias del front | Evidencia |
|---|---|---|---|
| `TbUsuarios` | Propia | 61 en 15 ficheros | `src/modules/cacheUsuario.bas:17`; `src/modules/Constructor.bas:834-845` |
| `TbHPS` | Propia | 57 en 8 ficheros | `src/classes/HPS.cls:562`; `src/modules/Mod_Indicadores_Backend.bas:30-56` |
| `TbUsuariosHistoricos` | Propia | 45 en 13 ficheros | `src/modules/Constructor.bas:1018-1020`; `src/classes/UsuarioHistorico.cls:646` |
| `TbUsuariosSICA` | Propia | 35 en 6 ficheros | `src/classes/UsuarioSICA.cls:332`; `src/modules/Constructor.bas:1240-1296` |
| `TbObservaciones` | Propia | 19 en 4 ficheros | `src/classes/Observacion.cls:185`; `src/modules/Constructor.bas:649` |
| `TbObservacionesHistoricas` | Propia | 17 en 4 ficheros | `src/classes/ObservacionHistorica.cls:177`, `:184`; `src/modules/Constructor.bas:2599` |
| `TbAnexosUsuariosHPS` | Propia | 25 en 5 ficheros | `src/classes/AnexoUsuarioHPS.cls:252`; `src/modules/AnexoNombreUnico.bas:29` |
| `TbAnexosUsuariosHistoricos` | Propia | 15 en 4 ficheros | `src/classes/AnexoUsuarioHistorico.cls:245`; `src/classes/UsuarioLifecycleTransactionCoordinator.cls:265` |
| `TbAnexosUsuariosSICA` | Propia | 6 en 2 ficheros | `src/classes/AnexoUsuarioSICA.cls:306`; `src/modules/Constructor.bas:1949` |
| `TbUsuarioAnexos` | Propia | 3 en 1 fichero | `src/modules/Constructor.bas:1919-1920` |
| `TbConsultas` | Propia | 1 en 1 fichero | `src/modules/Constructor.bas:3067-3068` |
| `TbMotivoHPS` | Propia | 2 en 2 ficheros | `src/modules/Constructor.bas:241-243`; `src/modules/Funciones Generales.bas:3993` |
| `TbHPSGrado` | Propia | 1 en 1 fichero | `src/modules/Constructor.bas:2401-2403` |
| `TbConfiguracionUsuariosAccesoLocal` | Propia | 13 en 1 fichero | `src/classes/Entorno.cls:274`, `:297`; la crea el front si falta (`Entorno.cls:345`) |
| `TbUsuariosEntidades` | Huérfana (candidata; retirada del código en #33) | 0 | Las únicas referencias son tests que exigen su ausencia: `src/modules/Test_CacheConsistencyAudit.bas:250`, `src/modules/Test_IndicadoresContadores.bas:551`. El retiro está documentado en `odd/tasks/2026-005-rc1-retirada-tbusuariosentidades.md` y en `docs/ciclos/RELEASE-2026-005.md:72` y `:112` |
| `TbHPSEquivalencia` | Huérfana (candidata) | 0 | Sin ninguna coincidencia en `src/` de `main` ni de la RC1 |
| `TbJuridicasContratacion` | Huérfana (candidata) | 0 | Igual que la anterior |
| `TbAuxCursos` | Huérfana (candidata) | 0 | Igual que las anteriores |
| `TbExpedientesLugaresEjecucion` | Mal ubicada (candidata) | 0 | Vinculada en `HPST` desde `Expedientes_datos.accdb` y sin ninguna referencia en el front: su dueño es la aplicación de Expedientes |
| `TbSolicitudesFechas` | Mal ubicada (candidata) | 0 | Vinculada en `HPST` desde `Solicitudes_HPS_datos.accdb` y sin ninguna referencia en el front: su dueño es la aplicación de HPS Solicitudes |
| `TbExpedientes`, `TbExpedientesConEntidades`, `TbSuministradores` | Propia y compartida | 33, 1 y 31 | Vinculadas desde `Expedientes_datos.accdb`; el front las usa a través de `getdbExpedientes()` y el backend las expone a sus propias uniones (`src/modules/Constructor.bas:839-843`) |
| `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` | Propia y compartida | 4 y 3 | Vinculadas desde `Lanzadera_Datos.accdb`; el front las lee por `getdbLanzadera()` |
| `TbSolicitudes` | Propia y compartida | 2 | Vinculada desde `Solicitudes_HPS_datos.accdb`; el front la lee por `getdbSolicitudHPS()` |
| `TbCorreosEnviados` | Propia y compartida | 2 | Vinculada desde `Correos_datos.accdb`; el front la lee por `getdbCorreo()` |

### Tablas del front

| Tabla | Clase | Referencias del front | Evidencia |
|---|---|---|---|
| `TbDatosLocal` | Caché del front | 51 en 12 ficheros | `src/modules/cacheUsuario.bas:36`, `:113`; `src/forms/Form_FormInicial03ConsultasDatos.form.txt:289` |
| `TbUsuariosSICALocal` | Caché del front | 20 en 8 ficheros | `src/forms/Form_FormInicial07UsuariosSICA.cls:124-125`; `src/modules/CacheConsistencyAudit.bas:139`; `src/modules/Funciones Generales.bas:3390` |
| `TbUsuariosHistoricosLocal` | Caché del front | 18 en 10 ficheros | `src/modules/Mod_Sincronizacion_Historico.bas:131-136`; `src/forms/Form_FormInicial08UsuariosHistoricos.cls:125-126` |
| `TbSuministradoresLocal` | Caché del front | 17 en 4 ficheros | `src/modules/cacheSuministrador.bas:26`, `:89`; `src/classes/Entorno.cls:2291` |
| `TbDatosLocalParaIndicadores` | Huérfana (candidata; retirada en #32) | 0 | Fuera del plan de cachés (`src/modules/Mod_StartupCacheInitialization.bas:19-23`); los tests exigen su ausencia (`src/modules/Test_CacheConsistencyAudit.bas:159`) |
| `TbUsuariosSICALocalParaIndicadores` | Huérfana (candidata; retirada en #32) | 0 | Igual que la anterior |
| `Copia de TbUsuariosEntidades` | Huérfana (candidata) | 0 | Cero coincidencias en todo el repositorio, no solo en `src/` |
| `TbCursosLocal` | Huérfana (candidata) | 0 | Cero coincidencias en todo el repositorio |
| `tblInfo` | Huérfana (candidata) | 0 | Cero coincidencias en todo el repositorio |
| `tblSettings` | Huérfana (candidata) | 0 | Cero coincidencias en todo el repositorio |

### Referencias a tablas que no existen

| Referencia | Código | Lectura |
|---|---|---|
| `TbObservacionesLocal` | `src/modules/Constructor.bas:1858-1859`, `:2067-2070` (7 apariciones) | La tabla no está ni en el front ni en `HPST`. La búsqueda global aborta en ese punto y el error lo recoge el manejador, de modo que las secciones «Observaciones» y «Observaciones Históricas» de la búsqueda no llegan a ejecutarse |
| `TbObservacionesHistoricasLocal` | `src/modules/Constructor.bas:1889-1890` | Misma causa |
| `TbAnexosHistoricos`, `TbUsuarioHistoricosAnexos` | `src/modules/Constructor.bas:2538`, `:2662`, `:2684` | Ramas muertas de `DameIDAnexo`: no hay ningún llamante que pase esos nombres, y las tablas no existen |

### Patrones de riesgo en el VBA del front

| Patrón | Medida |
|---|---|
| Generadores `Max + 1` | 11 puntos: `src/modules/Funciones Generales.bas:2794` (`DameID`), `src/modules/Constructor.bas:2546`, `:2585` (`TbUsuarios`), `:2603` (`TbUsuariosHistoricos`), `:2650` (`TbUsuarioAnexos`), `:2666` y `:2688` (ramas muertas), `:2732` (`TbObservaciones`), `:2750` (`TbObservacionesHistoricas`) y `src/classes/UsuarioSICA.cls:331` (`Max(Right([ID],4))` sobre un identificador de texto) |
| `LIKE '*…'` en SQL | 7 apariciones: `src/forms/Form_FormInicial07UsuariosSICA.cls:122`, `src/forms/Form_FormInicial08UsuariosHistoricos.cls:134`, `src/modules/Constructor.bas:44`, `:2795-2797`, `src/modules/Filtro.bas:136` |
| Fechas escritas como literal `#…#` con formato fijo | `src/classes/HPS.cls:645`, `src/classes/UsuarioHPS.cls:1450` y `:1456` (formato `mm/dd/yyyy` incrustado en el SQL) |
| `Date()` dentro de SQL | Ninguna |
| `Seek`, `dbOpenTable` | Ninguno |
| `dbSeeChanges` | Ninguna aparición: ningún recordset se abre con actualización en vivo |
| `Option Compare Database` | En los 132 ficheros de `src/`, es decir, la comparación es insensible a mayúsculas en todo el código |
| Escrituras en el backend desde el front | `src/classes/Entorno.cls:345` (crea `TbConfiguracionUsuariosAccesoLocal` en `HPST`) |

## Propuestas para después del descubrimiento global

Ninguna se aplica: la puerta de descubrimiento global sigue cerrada (#1576). El propietario borra en el legacy; aquí solo hay evidencia, comprobación previa y reversión. Los grupos son los de la fase 1 de la épica: (a) imprescindible antes de migrar, (b) recomendable si es barato, (c) aplazar a la refactorización final.

| Id | Propuesta | Grupo (fase 1) | Comprobación previa | Reversión |
|---|---|---|---|---|
| H1 | Poner en cuarentena las 3 tablas propias del backend sin ninguna referencia en el front: `TbAuxCursos`, `TbHPSEquivalencia` y `TbJuridicasContratacion`. Renombrar a `zzBAJA_<tabla>`. | (a) | Cruzar los vínculos de todos los backends (tarea T6) y leer el SQL de las 8 consultas guardadas del backend, que en este entorno no es legible. Comprobar en la copia de producción el recuento y que ninguna consulta guardada las usa. | Quitar el prefijo `zzBAJA_`. |
| H2 | Retirar la tabla `TbUsuariosEntidades` del backend (352 filas) y su copia `Copia de TbUsuariosEntidades` del front (99 filas). El código ya no las usa desde #33 y los tests exigen que no se creen. | (a) | Confirmar en el grafo global que ningún otro front ni la web las leen (tarea T6), y que la aplicación hermana `HPS_SOLICITUDES` ya no escribe su tabla equivalente. | Restaurar la tabla desde la copia de seguridad previa. |
| H3 | Retirar del front las dos cachés de indicadores retiradas en #32: `TbDatosLocalParaIndicadores` (344 filas) y `TbUsuariosSICALocalParaIndicadores` (50 filas). | (a) | Ejecutar en copia el arranque, los indicadores y el ciclo de vida de un usuario, comprobando que ninguna de las dos tablas se recrea (contrato ya validado en `odd/tasks/2026-005-rc1-retirada-cache-indicadores.md`, S3). | Reponer el binario del front anterior a la retirada. |
| H4 | Retirar del front las 4 tablas locales sin ninguna referencia en el repositorio: `Copia de TbUsuariosEntidades`, `TbCursosLocal`, `tblInfo` y `tblSettings`. | (b) | Abrir el front en copia y comprobar que ni el arranque ni la exportación a Excel las leen (no hay macros ni `AutoExec`), y revisar el SQL de la única consulta guardada del front (`Consulta1`), que en este entorno no es legible. | Reponer el binario del front anterior. |
| H5 | Fijar por escrito el contrato de las 4 cachés vivas del front (`TbDatosLocal`, `TbUsuariosSICALocal`, `TbUsuariosHistoricosLocal`, `TbSuministradoresLocal`): origen único en el backend y regenerador declarado en el plan de arranque (`src/modules/Mod_StartupCacheInitialization.bas:19-23`). | (b) | Explicar antes la diferencia de una fila entre `TbDatosLocal` (348) y `TbUsuarios` (349), y entre `TbUsuariosHistoricosLocal` (249) y `TbUsuariosHistoricos` (248), comparando el contenido tras un arranque completo en copia. | Restituir la cobertura anterior del plan `LOCAL_CACHE_COVERAGE`. |
| H6 | Retirar los 5 vínculos del front que apuntan a la ruta de producción (`\\datoste\…`): 2 a `Expedientes_datos.accdb` y 3 a `HPST.accdb`. El código no los consulta, pero un vínculo es una vía de escritura sobre producción desde el front. | (a) | Contar los errores de vínculo de los 5 vínculos del front con la misma comprobación que se hizo con los 57 de Brass, y confirmar en runtime que ningún formulario abierto los usa (no hay `DoCmd.OpenTable` ni macros). | Restituir cada vínculo con su definición original (`TableDef.Connect`). |
| H7 | Añadir a `HPST` el vínculo a `TbExpedientesSuministradores` de `Expedientes_datos.accdb`, o mover esa lectura al backend: hoy el front la consulta con `getdbExpedientes()` (`src/modules/Constructor.bas:83-90`) sobre una tabla que su backend no vincula. | (b) | Confirmar que la tabla existe en `Expedientes_datos.accdb` con clave primaria y que ningún otro vínculo de `HPST` la aporta ya. | Eliminar el vínculo añadido. |
| H8 | Retirar de `HPST` los 2 vínculos que ninguna aplicación usa: `TbExpedientesLugaresEjecucion` (de `Expedientes_datos.accdb`) y `TbSolicitudesFechas` (de `Solicitudes_HPS_datos.accdb`). Por R-2 la tabla pertenece a la aplicación que la usa; no se borra nada, solo se deja de exponer desde el backend que no la usa. | (b) | Grafo global (tarea T6) para confirmar que ningún otro proceso lee esas tablas a través de `HPST`: un vínculo no usado por el front puede seguir usándose desde la web o desde procesos programados. | Recrear el vínculo con su definición original. |
| H9 | Retirar la credencial literal de las cinco cadenas de conexión (`src/modules/VariablesEntorno.bas:564`, `:597`, `:633`, `:664`, `:693`) y leerla de la configuración. | (a) | La misma credencial abre los cinco backends del front. Comprobar que no queda ninguna otra copia del literal en el front y si es la misma que usan otros fronts. | Restaurar las líneas originales. |
| H10 | Sustituir los generadores `Max + 1` por autonumérico, empezando por `DameID` (`src/modules/Funciones Generales.bas:2794`) y por el identificador de texto de SICA (`src/classes/UsuarioSICA.cls:331`, `:346`). | (a) | Contar duplicados por tabla afectada antes de fijar la clave (`TbUsuarios`, `TbUsuariosHistoricos`, `TbObservaciones`, `TbObservacionesHistoricas`, `TbAnexosUsuariosHPS`, `TbAnexosUsuariosHistoricos`, `TbAnexosUsuariosSICA`, `TbUsuarioAnexos`) y revisar el tipo de las claves foráneas que las apuntan (`IDUsuario`, `IDAnexo`, `IDObservacion`). | Volver al generador actual. |
| H11 | Eliminar las ramas muertas de `DameIDAnexo` que apuntan a `TbAnexosHistoricos` y `TbUsuarioHistoricosAnexos` (`src/modules/Constructor.bas:2538`, `:2662`, `:2684`): no hay llamantes y las tablas no existen. | (b) | Confirmar que ningún consumidor externo llama a `Constructor.getID` con esos nombres (la función es pública). | Restituir las ramas. |
| H12 | Revisar las referencias a tablas inexistentes `TbObservacionesLocal` y `TbObservacionesHistoricasLocal` (`src/modules/Constructor.bas:1858-1859`, `:1889-1890`, `:2067-2070`): o se crean en el front, o se retiran las secciones de búsqueda que hoy nunca se ejecutan. | (b) | Reproducir en runtime la búsqueda global y comprobar qué secciones se ejecutan hoy y qué resultado devuelve. | Restituir las secciones de búsqueda. |

Lo que **no** se toca ahora: los 5 vínculos del front más allá de la propuesta H6, las 14 tablas propias del backend que el front usa, las 4 cachés vivas del front y las 7 relaciones de negocio del backend.

## Lo que no se pudo verificar

1. El inventario de tablas del binario publicado: el ERD del front se midió sobre la copia de trabajo del `HPS.accdb` del worktree huérfano (SHA-256 `ef9674e9…`), que no coincide con el blob del tag `2026-005` (`1212c01e…`). No se comparó el inventario de tablas de ambos binarios.
2. El SQL de las 8 consultas guardadas del backend y el de la única consulta guardada del front: `MSysQueries` no tiene permiso de lectura en este entorno. No se puede afirmar qué tablas usan ni si algún consumidor las usa.
3. Los vínculos en producción: el ERD del backend se midió sobre una copia local cuyos vínculos apuntan a `C:\00repos\datos\…`. La definición real en `\\datoste\…` no se pudo leer.
4. Los recuentos de filas: proceden de las copias en disco (front en el worktree, backend en `C:\00repos\datos\HPST.accdb`), no de producción. Las diferencias de una fila entre cachés y origen pueden ser deriva o pueden ser efecto de un arranque parcial; no se midió en producción.
5. El estado de la UAT en el tracker: la evidencia es el documento de release del repo (`docs/ciclos/RELEASE-2026-005.md`) y el historial de git, no el issue de UAT. El HTML de `docs/uat/uat-staging-2026-005-rc1-2026-10-22.html` es el formulario con los 24 escenarios, sin los resultados rellenados.
6. Consumidores fuera de Access (Excel, informes, procesos programados) de las tablas propuestas en cuarentena: no se ha revisado ninguno.
7. El trabajo de `origin/fix/2026-005-rc1-uat-gaps` (`c718589`) no se comparó módulo a módulo con `main`; solo se comprobó que no es ancestro, que no tiene PR y que su `src/` difiere en 38 ficheros.
8. El front abierto en runtime: no se ejecutó la aplicación. La traza de uso de tablas es estática, sobre el `src/` exportado, y no cubre posibles referencias construidas en tiempo de ejecución fuera de las cadenas de SQL revisadas.
