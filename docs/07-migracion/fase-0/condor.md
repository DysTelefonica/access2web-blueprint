# Fase 0: Condor (`CONDOR`)

Épica [`docs/07-migracion/epica-coexistencia-postgresql.md`](../epica-coexistencia-postgresql.md) (fase 0). Sin issue propia: la decisión del propietario (#1576) deja la fase 0 en catálogo y sin aplicación, y estas propuestas se revisan con el catálogo conjunto al cerrar el descubrimiento.
Base de análisis: rama `release/2026-007.rc1`, SHA `09e432ac48786bc2b3b88764239791ced5bf0d27`, worktree `00_CONDOR-worktrees/2026-007.rc1`.
Fuentes: ERD del front y del backend con `scripts/dysflow-erd/extract-erd.ps1` (solo lectura, sobre copia), trazado del código con CodeGraph (CLI 2.1.1 sobre el índice del worktree RC1: 303 ficheros, 11.133 nodos) y `read`/`grep` sobre `src/`.
Rutas: las que empiezan por `src/`, `docs/`, `openspec/`, `odd/` son del repositorio de la aplicación; las que empiezan por `docs/07-migracion/` son de este repositorio.

## Base de análisis

| Hecho | Evidencia |
|---|---|
| La rama del ciclo es `release/2026-007.rc1`, cortada de `staging` | `docs/condor/roadmap.md:9` («Rama \| `release/2026-007.rc1` (corte de `staging` `aebcef6`)») |
| `main` es la línea de producción de referencia, no la base | `docs/condor/roadmap.md:10` («Producción de referencia \| tag `2026-006` en `main`»); `main` está en `97b98732501545727c91e5807d55af8613c65800` |
| La UAT del ciclo no está superada | Contrato de aceptación en «Revisión 1 (borrador)», que «se congela cuando Calidad firme los criterios» (`openspec/changes/release-2026-007-calidad-acceptance/specs/calidad-acceptance/spec.md:9`); auditoría del rc1 del 2026-11-03 con REQ-01..06 pendientes, REQ-07 sin criterio y REQ-08 y REQ-11 parciales (`:15-27`); el roadmap deja los once REQ sin cerrar (`docs/condor/roadmap.md:18-28`); no hay registro de UAT firmado en el repositorio |
| El encargo de UAT es posterior a esa auditoría | `odd/tasks/uat-2026-007-requirements-audit.md:3` («para que luego las pruebas UAT sean contra eso que nos pidieron») |
| La única web de aceptación del repositorio es de otra feature | `docs/uat/uat-staging-2026-07-13.html` (`UAT_META.featureId = "tipo-solicitud-pc-subcontratista"`), sin registro descargado |
| El trabajo local de la RC1 va dos commits por delante del remoto | `git log origin/release/2026-007.rc1..HEAD` → `09e432a`, `c7f8cd7`; el remoto está en `8ea212b` |
| `main` y la línea RC1 no comparten ancestro en este clon | `git merge-base main origin/staging` no devuelve nada: `main` tiene 4 commits con raíz `c88347e` y `origin/staging` 7 con raíz `43087ad`. No hay diff posible entre `main` y la RC1 |
| El binario del worktree tiene cambios sin commitear | `git status --short` → `M CONDOR.accdb`, `M .atl/skill-registry.md`. El ERD del front refleja el binario del worktree, no el del commit `09e432a` |

Consecuencia: el inventario de abajo describe la RC1, no `main`. Las diferencias con `main` no son un error del inventario.

## Estructura del front

| Hecho | Valor |
|---|---|
| Tablas del front | 2, ambas propias, 0 vínculos |
| Tablas propias del front | `TbConfiguracionBackends` (1 registro), `tbSandboxLog` (179 registros) |
| Formularios | 45 en el binario; 45 `*.form.txt` en `src/forms/` |
| Consultas guardadas | 0 |
| Macros | Ninguna: el inventario de objetos del front no lista módulos de macro |
| Formularios con `RecordSource` | 0: ninguna coincidencia de `RecordSource` en los 45 `*.form.txt` |
| `RowSourceType` en diseño | Solo «Value List», en todos los combos y listas |
| `RowSourceType` en runtime | «Table/Query» en `src/forms/frmVerLogCambios.cls:15,19` y `src/forms/frmVerLogErrores.cls:16,21` |
| `ControlSource` | Solo en `src/forms/frmDetalleLogCambio.form.txt` y `src/forms/frmDetalleLogError.form.txt`, sobre tablas sin `RecordSource`: inertes en diseño |

La excepción de Brass no se aplica aquí: **los formularios de Condor están desconectados**. El diseño no dice nada del uso real, y el relleno ocurre en runtime:

| Procedimiento | Qué rellena | Evidencia |
|---|---|---|
| `frmDetalleLogCambio` | Los campos de un registro de `tbLogCambios`, localizado por `OpenArgs` | `src/forms/frmDetalleLogCambio.cls:17,28-35` |
| `frmBuscarSolicitudes` | La lista de resultados, como lista de valores | `src/forms/frmBuscarSolicitudes.cls:84,103` |
| `frmBuscarExpediente` | La lista de resultados | `src/forms/frmBuscarExpediente.cls:131,151` |
| `frmAltaSolicitud` | El combo de no conformidades, con `TbNoConformidades` | `src/forms/frmAltaSolicitud.cls:337,362,367` |
| `frmGestionAdjuntos` | La lista de adjuntos | `src/forms/frmGestionAdjuntos.cls:204,212` |
| `frm0Ppal` | El combo de técnicos, desde un diccionario | `src/forms/frm0Ppal.cls:160` |

Los valores de diseño están además caducos: `src/forms/frm0Ppal.form.txt:21737` guarda una lista literal de usuarios que el runtime sustituye (`:160`). Los dos combos de filtro de log son los únicos que consultan tablas por SQL en runtime, y lo hacen contra `tbLogCambios` y `tbLogErrores` (`src/forms/frmVerLogCambios.cls:16`, `src/forms/frmVerLogErrores.cls:17,22`).

## Aperturas directas

| Apertura | Definición | Ruta resuelta | Llamadores en la RC1 |
|---|---|---|---|
| Backend propio | `getdb()`, `src/modules/FUNCIONES UTILES.bas:83` | Desde `TempVars` (`BackendActivo`): `:272-284` | 20 (CodeGraph) |
| Lanzadera | `getdbLanzadera()`, `src/modules/FUNCIONES UTILES.bas:63` | `:271` (`…\0Lanzadera\Lanzadera_Datos.accdb` o `…\000datoslocal\…`) | 0 |
| Correos | `getdbCorreo()`, `src/modules/FUNCIONES UTILES.bas:167` | `:286` (`…\00Recursos\Correos_datos.accdb`) | 0 |
| Expedientes | `getdbExpedientes()`, `src/modules/FUNCIONES UTILES.bas:187` | `:288` (`…\EXPEDIENTES\Expedientes_datos.accdb`) | 0 |
| No Conformidades | `getdbNoConformidades()`, `src/modules/FUNCIONES UTILES.bas:206` | `:290` (`…\No Conformidades\NoConformidades_Datos.accdb`) | 0 |
| `DoCmd.OpenTable`, `TransferDatabase`, `TransferText` | — | — | 0 |
| `CurrentDb` | `src/modules/SchemaInspector.bas:25,86,181`, `src/modules/TestHelper.bas:341`, `src/modules/Variables Globales.bas:248`, `src/modules/modDevTools.bas:35` | El propio front | 6: cinco para leer la configuración local y uno en la herramienta de diagnóstico |

La contraseña de todas las aperturas sale de `TbConfiguracionBackends.PasswordBackend` (`src/modules/FUNCIONES UTILES.bas:146,295,303`); no hay literales en el código. La comprobación de los cero llamadores es de CodeGraph (`codegraph callers getdbLanzadera|getdbExpedientes|getdbCorreo|getdbNoConformidades`: «No callers found»), y el grafo estático no prueba ausencia de invocación dinámica.

En `main` la situación era la contraria: 27 llamadas de negocio en los cuatro abridores —4 en `getdbLanzadera`, 8 en `getdbExpedientes`, 11 en `getdbNoConformidades` y 4 en `getdbCorreo`— (`src/classes/UsuarioServicio.cls:141`, `src/modules/AplicacionRepositorio.bas:27`, `src/modules/UsuarioRepositorio.bas:42`, `src/modules/Variables Globales.bas:94`, `src/classes/ExpedienteServicio.cls:79`, `src/classes/SuministradorServicio.cls:34`, `src/classes/UsuarioServicio.cls:121`, `src/modules/ExpedienteRepositorio.bas:40,68,86`, `src/modules/SuministradorRepositorio.bas:20`, `src/classes/SolicitudServicio.cls:571,713`, `src/modules/NoConformidadRepositorio.bas:15,36,67,102,130,154,180,213`, `src/classes/NotificacionServicio.cls:44,101,168`, `src/modules/CorreoRepositorio.bas:48`). El [inventario de aperturas directas](aperturas-directas-lanzadera.md#resumen-por-front) mide `main` (SHA `97b9873`) y por eso cuenta 4 llamadas y 4 procedimientos de `getdbLanzadera`: en la RC1 son cero. La regresión es de la RC1, no de `main`.

## Tablas de otros backends que consume el front

En la RC1 el front **no** abre ninguna de estas bases: llega a las diez por los vínculos de su propio backend. La tabla «vía» es, por tanto, la única.

| Tabla | Backend de origen | Cómo llega el front | Evidencia |
|---|---|---|---|
| `TbUsuariosAplicaciones` | `Lanzadera_Datos` | Vínculo en `condor_datos` | `src/classes/UsuarioServicio.cls:135-141`, `src/classes/ExpedienteServicio.cls:35` |
| `TbUsuariosAplicacionesPermisos` | `Lanzadera_Datos` | Vínculo en `condor_datos` | `src/classes/UsuarioServicio.cls:136-140`, `src/modules/AplicacionRepositorio.bas:25` |
| `TbExpedientes` | `Expedientes_datos` | Vínculo en `condor_datos` | `src/classes/ExpedienteServicio.cls:34`, `src/classes/SolicitudServicio.cls:103` |
| `TbExpedientesResponsables` | `Expedientes_datos` | Vínculo en `condor_datos` | `src/classes/ExpedienteServicio.cls:37,66`, `src/classes/SolicitudServicio.cls:85,123` |
| `TbExpedientesRACS` | `Expedientes_datos` | Vínculo en `condor_datos` | `src/classes/ExpedienteServicio.cls:40` |
| `TbRACS` | `Expedientes_datos` | Vínculo en `condor_datos` | `src/classes/ExpedienteServicio.cls:40` |
| `TbExpedientesSuministradores` | `Expedientes_datos` | Vínculo en `condor_datos` | `src/classes/SolicitudServicio.cls:148`, `src/modules/ExpedienteRepositorio.bas:62` |
| `TbSuministradores` | `Expedientes_datos` | Vínculo en `condor_datos` | `src/classes/SuministradorServicio.cls:25`, `src/modules/SuministradorRepositorio.bas:20` |
| `TbNoConformidades` | `NoConformidades_Datos` | Vínculo en `condor_datos` | `src/modules/NoConformidadRepositorio.bas:16,39,68,102,130,154,181,213` |
| `TbCorreosEnviados` | `Correos_datos` | Vínculo en `condor_datos`, con escritura | `src/modules/CorreoRepositorio.bas:18,50-51,76` |

Los diez vínculos están justificados: el VBA del front propietario los necesita. El único problema es de forma: `TbCorreosEnviados` se escribe desde el front (`src/modules/CorreoRepositorio.bas:50-51,76`) y no tiene clave primaria.

## Backend `condor_datos`

Fichero: `C:\00repos\datos\condor_datos.accdb`, el que declara el propio proyecto (`openspec/changes/release-2026-007-calidad-acceptance/specs/calidad-acceptance/spec.md:13`, `odd/tasks/uat-2026-007-requirements-audit.md:33`).

| Hecho | Valor |
|---|---|
| Tablas | 25: 15 propias y 10 vinculadas |
| Relaciones | 5: cuatro reales (`tbSolicitudes` → `tbAdjuntos`, `tbDatosCDCA`, `tbDatosCDCASUB`, `tbDatosPC`) y una del sistema de navegación (`MSysNavPaneGroupCategories` → `MSysNavPaneGroups`) |
| Consultas guardadas | 2: `Consulta1` y `Consulta2` (tipo 0, sin SQL en el ERD) |
| Tablas sin clave primaria | 1: `TbCorreosEnviados` (vinculada a `Correos_datos`) |
| Campos multivalor, adjuntos u OLE | Ninguno |
| Tipos presentes | Boolean, Integer, Long, Currency, Date, Text y Memo. Los dos campos `Currency` (`ImporteLicitacion` e `ImporteContratacion`) están en la tabla vinculada `TbExpedientes`, es decir, en el backend de Expedientes |

Vínculos: 6 a `Expedientes_datos.accdb` (`MS Access;PWD=***;DATABASE=…`), 2 a `Lanzadera_Datos.accdb`, 1 a `NoConformidades_Datos.accdb` y 1 a `Correos_datos.accdb`.

El ERD declara relación con cuatro tablas de detalle (`tbAdjuntos`, `tbDatosCDCA`, `tbDatosCDCASUB` y `tbDatosPC`); el código une por `idSolicitud` con otras cuatro que no la tienen declarada: `tbDatosPCSUB`, `tbRechazos` (`src/modules/RechazoRepositorio.bas:197`), `tbLogEstados` (`src/modules/LogEstadoRepositorio.bas:56`) y `tbValidacionRevision` (`src/modules/SnapshotHelper.bas:466`).

El ERD del repositorio (`docs/ERD/condor_datos.md`, generado el 2026-05-08) describe 22 tablas; el backend vivo tiene 25. La diferencia son `tbDatosPCSUB`, `TbCorreosEnviados` y `TbUsuariosAplicacionesPermisos`: el ERD del repositorio está atrasado y no sirve como inventario.

## Veredicto de propiedad

Referencias: líneas de `src/` que mencionan el nombre de la tabla como palabra entera, en código de producción (excluidos `Test_*`, `modBattery_Canonical.bas` y `Pruebas.bas`); «SQL» son las que además contienen una cláusula `FROM`, `INTO`, `UPDATE` o `JOIN`. Es el coste del cambio, no el número de llamadas.

| Tabla | Tipo | ¿La usa el VBA del front? | Refs | SQL | Clase |
|---|---|---|---|---|---|
| `tbSolicitudes` | Propia | Sí | 34 | 14 | Propia |
| `tbDatosPC` | Propia | Sí | 25 | 16 | Propia |
| `tbDatosPCSUB` | Propia | Sí | 29 | 16 | Propia |
| `tbDatosCDCASUB` | Propia | Sí | 22 | 18 | Propia |
| `tbDatosCDCA` | Propia | Sí | 19 | 16 | Propia |
| `tbValidacionRevision` | Propia | Sí | 19 | 19 | Propia |
| `tbLogErrores` | Propia | Sí | 13 | 7 | Propia |
| `tbRechazos` | Propia | Sí | 13 | 8 | Propia |
| `tbAdjuntos` | Propia | Sí | 13 | 9 | Propia |
| `tbLogCambios` | Propia | Sí | 10 | 7 | Propia |
| `tbMapeoCampos` | Propia | Sí | 6 | 4 | Propia |
| `tbLogEstados` | Propia | Sí | 6 | 3 | Propia |
| `tbEstados` | Propia | Sí | 4 | 2 | Propia |
| `tbTransiciones` | Propia | Sí | 2 | 1 | Propia |
| `tbHistorialRechazos` | Propia | **No** | 0 | 0 | **Huérfana, provisional** |
| `TbExpedientes` | Vinculada (`Expedientes_datos`) | Sí | 8 | 5 | De la app de Expedientes |
| `TbExpedientesResponsables` | Vinculada (`Expedientes_datos`) | Sí | 7 | 7 | De la app de Expedientes |
| `TbExpedientesRACS` | Vinculada (`Expedientes_datos`) | Sí | 4 | 3 | De la app de Expedientes |
| `TbRACS` | Vinculada (`Expedientes_datos`) | Sí | 4 | 3 | De la app de Expedientes |
| `TbExpedientesSuministradores` | Vinculada (`Expedientes_datos`) | Sí | 2 | 2 | De la app de Expedientes |
| `TbSuministradores` | Vinculada (`Expedientes_datos`) | Sí | 2 | 2 | De la app de Expedientes |
| `TbNoConformidades` | Vinculada (`NoConformidades_Datos`) | Sí | 8 | 8 | De la app de No Conformidades |
| `TbCorreosEnviados` | Vinculada (`Correos_datos`) | Sí, escribe | 2 | 1 | De la app de Correo |
| `TbUsuariosAplicaciones` | Vinculada (`Lanzadera_Datos`) | Sí | 13 | 9 | De la Lanzadera |
| `TbUsuariosAplicacionesPermisos` | Vinculada (`Lanzadera_Datos`) | Sí | 5 | 2 | De la Lanzadera |
| `TbConfiguracionBackends` | Del front | Sí | 11 | 6 | Configuración del front |
| `tbSandboxLog` | Del front | **No** | 0 | 0 | **Huérfana, provisional** |

Tres ajustes sobre el veredicto:

- **`tbHistorialRechazos` es huérfana provisional.** Cero referencias en el `src/` de la RC1 (CodeGraph: «No results found» para `tbHistorialRechazos`; el único símbolo con ese nombre es el procedimiento `RechazoServicio.LimpiarHistorialRechazos`, `src/classes/RechazoServicio.cls:18`, que borra de `tbRechazos`, no de `tbHistorialRechazos`: `src/modules/RechazoRepositorio.bas:197`). La tabla tiene 0 registros y un esquema casi idéntico al de `tbRechazos` (`FechaRechazo`, `MotivoPrincipal`, `AreaAfectada`, `Comentarios`): es la predecesora de `tbRechazos`. La clasificación es provisional hasta el grafo de vínculos de todos los backends.
- **`tbSandboxLog` es huérfana provisional en el front.** Cero referencias en el `src/` de la RC1 (CodeGraph: «No results found» para `tbSandboxLog`); el módulo `SandboxLogger.bas` y la tabla se eliminaron del slice el 2026-06-15 (`docs/testing/access-vba-tdd-v2_4_2-debt.md:25`). Quedan 179 registros. En `main` el módulo todavía existe (`src/modules/SandboxLogger.bas`, `src/classes/SandboxGestor.cls:194,722-744`), de modo que la tabla solo es huérfana en la RC1.
- **Las dos tablas del front están fuera de la regla «las tablas viven en el backend».** `TbConfiguracionBackends` la lee el front con `CurrentDb` antes de poder abrir ningún backend (`src/modules/Variables Globales.bas:247-248`, `src/modules/SchemaInspector.bas:86-91`, `src/modules/TestHelper.bas:341-343`): contiene la contraseña y las rutas del backend, y moverla al backend es circular. `tbSandboxLog` es residuo del sandbox retirado.

### Verificación de R-1 y R-2

- **R-1 (cada front consulta solo su backend): se confirma en la RC1.** El único `getdb()` en uso devuelve `condor_datos` (`src/modules/FUNCIONES UTILES.bas:83,272-284`) y las diez tablas externas llegan por sus vínculos en ese backend. La excepción son los cuatro abridores externos, que siguen en el código sin llamadores: son deuda, no camino de datos.
- **R-2 (una tabla vive en el backend de la aplicación que la usa): se confirma para las quince tablas propias**, con la salvedad de `tbHistorialRechazos`. Las diez vínculos no están mal ubicadas: su dueña es la aplicación del backend de origen (Expedientes, No Conformidades, Correo y Lanzadera), y el vínculo en `condor_datos` es el mecanismo correcto de coexistencia.

## Patrones de riesgo en VBA

Medidos sobre el código de producción de la RC1 (143 módulos).

| Patrón | Medida | Dónde |
|---|---|---|
| Generadores `Max + 1` | 9 puntos | `src/modules/FUNCIONES UTILES.bas:537` (helper `DameID`, `:516`, llamado en `src/classes/SolicitudServicio.cls:321`), `src/modules/AdjuntoRepositorio.bas:44`, `src/modules/CorreoRepositorio.bas:18`, `src/modules/DatosCDCARepositorio.bas:17`, `src/modules/DatosCDCASUBRepositorio.bas:10`, `src/modules/DatosPCRepositorio.bas:79`, `src/modules/DatosPCSUBRepositorio.bas:80`, `src/modules/LogCambioRepositorio.bas:38`, `src/modules/LogErrorRepositorio.bas:48` y `src/modules/ErrorLogger.bas:46` |
| `LIKE` o `Like` con comodín | 3 puntos | `src/classes/DocumentoServicio.cls:854,1034` (`Like "rac*"`, sensible a mayúsculas por `Option Compare Database`); el resto son `LIKE [p_PalabraClave]` parametrizado |
| SQL contra tablas del sistema | 3 puntos | `src/forms/frmVerLogCambios.cls:16`, `src/forms/frmVerLogErrores.cls:17,22` (`UNION SELECT … FROM MSysObjects`) |
| `Option Compare Database` | 139 de 143 módulos | Comparación de cadenas insensible a mayúsculas en todo el código |
| Literales de fecha `#…#` en SQL | 0 | — |
| `Date()` en SQL | 0 | — |
| `Seek` o `dbOpenTable` | 0 | — |
| Recordsets sin `dbSeeChanges` | 0 | Ninguna llamada a `OpenRecordset` pasa argumento de opciones; 92 llamadas en 27 módulos, con `dbOpenDynaset` o `dbOpenSnapshot` |

## Propuestas para después del descubrimiento global

Ninguna se aplica: la puerta de descubrimiento global sigue cerrada (#1576). Los grupos son los de la fase 1 de la épica: (a) imprescindible antes de migrar, (b) recomendable si es barato, (c) aplazar a la refactorización final.

| Id | Propuesta | Grupo | Comprobación previa | Reversión |
|---|---|---|---|---|
| C1 | Poner en cuarentena `tbHistorialRechazos`: renombrar a `zzBAJA_tbHistorialRechazos` tras una copia de seguridad. | (a) | Confirmar que ningún otro front ni la web la leen: exige el grafo de vínculos de todos los backends. Leer el SQL de `Consulta1` y `Consulta2`, que el ERD no trae, para descartar que la usen. Verificar en la copia de producción que los 0 registros son reales. | Quitar el prefijo y restaurar la tabla desde la copia. |
| C2 | Declarar `TbConfiguracionBackends` como configuración propia del front, exenta de la regla «las tablas viven en el backend», con su motivo: se lee con `CurrentDb` antes de poder abrir ningún backend. | (a) (declaración, no movimiento) | Confirmar con el propietario que la tabla no puede residir en el backend: contiene la contraseña y la ruta del backend (`src/modules/Variables Globales.bas:247-248`). | Sin cambio en el legacy; si el propietario decide otra cosa, se planifica aparte. |
| C3 | Poner en cuarentena `tbSandboxLog` en el front: renombrar a `zzBAJA_tbSandboxLog`. | (b) | Confirmar que ninguna otra copia del front la usa en runtime: `main` todavía tiene `SandboxLogger.bas` y `SandboxGestor.cls`. Decidir si el front se apaga sin más o se conserva. | Quitar el prefijo. |
| C4 | Retirar los cuatro abridores externos sin llamadores: `getdbLanzadera`, `getdbExpedientes`, `getdbCorreo`, `getdbNoConformidades` (`src/modules/FUNCIONES UTILES.bas:63,167,187,206`). | (a) | Ejecutar el guard estático del cambio `backend-getdb-single-entrypoint` sobre la RC1 y sobre el binario, no solo sobre `src/`: CodeGraph no ve invocación dinámica. Comprobar en el binario que el recuento de llamadas es cero. | Restaurar las definiciones de `src/modules/FUNCIONES UTILES.bas` y las variables `g_db*`. |
| C5 | Declarar clave primaria en `TbCorreosEnviados`, en `Correos_datos.accdb` (backend de la aplicación de correo, no de Condor). | (a) | Verificar duplicados de `IDCorreo` en la copia del backend de correo. Medir el coste sobre el escritor de Condor (`src/modules/CorreoRepositorio.bas:50-51,76`), que es el único consumidor conocido. Comprobar que la aplicación de correo no inserta en esa tabla con su propia clave. | Quitar la clave primaria. |
| C6 | Sustituir los 9 generadores `Max + 1` por autonumérico o identidad. | (a) | Coste por tabla según la tabla de referencias anterior: `tbSolicitudes` (34), `tbDatosPCSUB` (29), `tbDatosPC` (25), `tbDatosCDCASUB` (22), `tbDatosCDCA` (19), `tbValidacionRevision` (19), `tbLogErrores` (13), `tbRechazos` (13), `tbAdjuntos` (13), `tbLogCambios` (10), `tbMapeoCampos` (6), `tbLogEstados` (6), `tbEstados` (4), `tbTransiciones` (2). Comprobar que no hay duplicados ni huecos antes de cambiar el esquema. | Volver al generador `Max + 1` y a los valores calculados. |
| C7 | Declarar las cuatro relaciones que faltan en el backend: `tbSolicitudes` → `tbDatosPCSUB`, `tbRechazos`, `tbLogEstados` y `tbValidacionRevision`. | (b) | Comprobar que no hay filas huérfanas en las cuatro tablas antes de declarar la integridad referencial. | Quitar las relaciones. |
| C8 | Reescribir el filtro de usuarios de los dos combos de log sin `MSysObjects` (`src/forms/frmVerLogCambios.cls:16`, `src/forms/frmVerLogErrores.cls:17,22`): `MSysObjects` no existe en PostgreSQL. | (b) | Comprobar que el «(Todos)» de la lista se puede emitir como literal en la misma consulta. | Restaurar las tres líneas. |
| C9 | Decidir el destino de la contraseña del backend, hoy en claro en `TbConfiguracionBackends.PasswordBackend` y en las cadenas de conexión de todos los repositorios. | (b) | Con el propietario: durante la coexistencia la web necesita abrir las mismas tablas. Medir el alcance: la clave sale de `GetPasswordDB()` (`src/modules/FUNCIONES UTILES.bas:146`) y se concatena en 42 aperturas de `src/` (15 módulos). | Dejar la columna y la configuración como están. |

Lo que **no** se toca ahora: los diez vínculos del backend, las quince tablas propias, las cuatro relaciones reales existentes, la tabla de configuración del front y el binario del front.

## Lo que no se pudo verificar

1. Si `tbHistorialRechazos` o `tbSandboxLog` las consume otro aplicativo o la web: exige el grafo de vínculos de todos los backends, pendiente del descubrimiento global (#1576).
2. El SQL de `Consulta1` y `Consulta2`: el ERD solo devuelve nombre y tipo. No se ha abierto el backend para leerlos; puede ocultar referencias a `tbHistorialRechazos`.
3. El estado de la UAT en el tracker: la evidencia es el repositorio (contrato de aceptación en borrador y roadmap sin cerrar), no la issue de UAT.
4. Las propiedades de diseño de los formularios se han leído del export `src/forms/*.form.txt`, no del binario en vivo: la extracción del ERD no incluye propiedades de formulario. La paridad de formularios entre binario y `src/` (45 y 45) y la reconciliación del 2026-11-03 (`docs/condor/roadmap.md:36`) son la evidencia indirecta.
5. Que el binario de la RC1 no tenga llamadas a los cuatro abridores externos que `src/` no tiene: el grafo de CodeGraph y el `grep` cubren `src/`; el binario del worktree tiene cambios sin commitear respecto a `09e432a`.
6. Las diferencias entre `main` y la RC1 módulo a módulo: en este clon ambas líneas no comparten ancestro, así que no hay diff posible. La comparación se ha hecho por existencia y por contadores sobre `src/`, no sobre el historial.
7. Si `TbExpedientes*`, `TbNoConformidades`, `TbCorreosEnviados` y `TbUsuariosAplicaciones*` las usan además otras aplicaciones que no son sus dueñas: no se ha revisado el `src/` de los demás.fronts.
