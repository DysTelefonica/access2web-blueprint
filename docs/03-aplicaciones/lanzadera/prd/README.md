# PRD — Lanzadera

> **Estado:** en curso · fases 0 (fijación) y 1 (inventario) cerradas; fase 2 (formularios) con 13 de 28.

| Campo | Valor |
|---|---|
| Repositorio | DysTelefonica/Lanzadera |
| Commit fijado | `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` (`main`, 2026-07-18, «chore(agents): delegate dysflow config to global») |
| Índice CodeGraph-VBA | 88 ficheros indexados, 4 473 nodos, versión 2.1.1, sobre clon de solo lectura de ese commit |
| Estudio previo | [epic.md](../epic.md) (DRAFT), [forms.md](../forms.md), [capabilities.md](../capabilities.md), [data-model.md](../data-model.md), [integrations-security.md](../integrations-security.md), [walkthroughs G1 a G5](../walkthrough-G1.json), [ERD/schema.sql](../ERD/schema.sql) |
| Snapshot comparado | `data/staging/lanzaderas/src` (2026-08-06) |
| Web existente | módulo `app/src/modules/lanzadera` (MVP); este PRD lo reconcilia con `src` sin reescribirlo |
| Issue | #1074 (método y plantilla de #1069) |

## Alcance actual

Fija el commit, concilia el inventario y deja todos los formularios como **pendientes**. No contiene fichas, ERD, matriz, coexistencia ni capacidades: llegan en las tandas siguientes de la cadena. Los recuentos de líneas, controles y procedimientos son mecánicos (ver [Método del recuento](#método-del-recuento)) y no sustituyen al análisis.

## Inventario reconciliado

### Recuentos por tipo

| Tipo | `1474e8e` | Snapshot 2026-08-06 | Diferencia |
|---|---|---|---|
| Formularios (`src/forms/*.cls`, con su `.form.txt`) | 28 | 28 | 0 (mismos nombres) |
| Módulos estándar (`src/modules/*.bas`) | 17 (9 de pruebas) | 17 | 0 |
| Clases (`src/classes/*.cls`) | 14 | 14 | 0 |
| Consultas guardadas | sin carpeta `src/queries/` | sin recuento | — |

Tras quitar las líneas `Attribute VB_`, los comentarios, las líneas en blanco y las diferencias de espacios y mayúsculas, **los 59 ficheros de código (formularios, módulos y clases) son idénticos a los del snapshot**. Diferencia textual cero no demuestra que el estudio previo sea correcto: solo que el código no ha cambiado desde entonces.

### Formularios

Líneas: `.cls` del formulario. Controles: bloques `Begin` del `.form.txt`. Procedimientos: `Sub` y `Function` del `.cls`. «Frente al snapshot»: **igual** si el `.cls` coincide con el del snapshot tras la normalización anterior.

| Objeto | Tipo | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha o motivo |
|---|---|---|---|---|---|---|---|---|
| Form_FormAplicacionDatosGenerales | formulario | src/forms/Form_FormAplicacionDatosGenerales.cls | 268 | 48 | 4 | igual | analizado | [formularios/Form_FormAplicacionDatosGenerales.md](formularios/Form_FormAplicacionDatosGenerales.md) |
| Form_FormAplicacionGeneral | formulario | src/forms/Form_FormAplicacionGeneral.cls | 129 | 22 | 12 | igual | analizado | [formularios/Form_FormAplicacionGeneral.md](formularios/Form_FormAplicacionGeneral.md) |
| Form_FormAplicaciones | formulario | src/forms/Form_FormAplicaciones.cls | 156 | 24 | 15 | igual | analizado | [formularios/Form_FormAplicaciones.md](formularios/Form_FormAplicaciones.md) |
| Form_FormAplicacionesFueraOficina | formulario | src/forms/Form_FormAplicacionesFueraOficina.cls | 493 | 47 | 13 | igual | pendiente | sin ficha todavía |
| Form_FormAplicacionesGestion | formulario | src/forms/Form_FormAplicacionesGestion.cls | 149 | 21 | 6 | igual | analizado | [formularios/Form_FormAplicacionesGestion.md](formularios/Form_FormAplicacionesGestion.md) |
| Form_FormAplicacionesOficina | formulario | src/forms/Form_FormAplicacionesOficina.cls | 371 | 29 | 9 | igual | pendiente | sin ficha todavía |
| Form_FormAplicacionesPrueba | formulario | src/forms/Form_FormAplicacionesPrueba.cls | 293 | 16 | 6 | igual | pendiente | sin ficha todavía |
| Form_FormEnConstruccion | formulario | src/forms/Form_FormEnConstruccion.cls | 2 | 12 | 0 | igual | analizado | [formularios/Form_FormEnConstruccion.md](formularios/Form_FormEnConstruccion.md) |
| Form_FormLogin | formulario | src/forms/Form_FormLogin.cls | 311 | 23 | 7 | igual | analizado | [formularios/Form_FormLogin.md](formularios/Form_FormLogin.md) |
| Form_FormMenuPrincipalAdmin | formulario | src/forms/Form_FormMenuPrincipalAdmin.cls | 323 | 40 | 33 | igual | analizado | [formularios/Form_FormMenuPrincipalAdmin.md](formularios/Form_FormMenuPrincipalAdmin.md) |
| Form_FormMenuPrincipalCalidad | formulario | src/forms/Form_FormMenuPrincipalCalidad.cls | 375 | 39 | 27 | igual | analizado | [formularios/Form_FormMenuPrincipalCalidad.md](formularios/Form_FormMenuPrincipalCalidad.md) |
| Form_FormMenuPrincipalUsuario | formulario | src/forms/Form_FormMenuPrincipalUsuario.cls | 373 | 33 | 23 | igual | analizado | [formularios/Form_FormMenuPrincipalUsuario.md](formularios/Form_FormMenuPrincipalUsuario.md) |
| Form_FormObtenerContraseña | formulario | src/forms/Form_FormObtenerContraseña.cls | 98 | 20 | 3 | igual | analizado | [formularios/Form_FormObtenerContraseña.md](formularios/Form_FormObtenerContraseña.md) |
| Form_FormPerfil | formulario | src/forms/Form_FormPerfil.cls | 220 | 39 | 4 | igual | analizado | [formularios/Form_FormPerfil.md](formularios/Form_FormPerfil.md) |
| Form_FormUsuarioCambioPass | formulario | src/forms/Form_FormUsuarioCambioPass.cls | 111 | 26 | 4 | igual | analizado | [formularios/Form_FormUsuarioCambioPass.md](formularios/Form_FormUsuarioCambioPass.md) |
| Form_FormUsuarioDatosGenerales | formulario | src/forms/Form_FormUsuarioDatosGenerales.cls | 393 | 47 | 6 | igual | pendiente | sin ficha todavía |
| Form_FormUsuarioGeneral | formulario | src/forms/Form_FormUsuarioGeneral.cls | 181 | 25 | 16 | igual | pendiente | sin ficha todavía |
| Form_FormUsuarioPerfilAplicaciones | formulario | src/forms/Form_FormUsuarioPerfilAplicaciones.cls | 385 | 31 | 6 | igual | pendiente | sin ficha todavía |
| Form_FormUsuariosConectados | formulario | src/forms/Form_FormUsuariosConectados.cls | 27 | 11 | 1 | igual | pendiente | sin ficha todavía |
| Form_FormUsuariosGestion | formulario | src/forms/Form_FormUsuariosGestion.cls | 194 | 20 | 7 | igual | pendiente | sin ficha todavía |
| Form_FormVideoDetalle | formulario | src/forms/Form_FormVideoDetalle.cls | 81 | 18 | 3 | igual | pendiente | sin ficha todavía |
| Form_FormVideosArbolVideo | formulario | src/forms/Form_FormVideosArbolVideo.cls | 58 | 21 | 6 | igual | pendiente | sin ficha todavía |
| Form_FormVideosGestion | formulario | src/forms/Form_FormVideosGestion.cls | 203 | 16 | 7 | igual | pendiente | sin ficha todavía |
| Form_FormVideosGestionArbolAplicacion | formulario | src/forms/Form_FormVideosGestionArbolAplicacion.cls | 417 | 33 | 8 | igual | pendiente | sin ficha todavía |
| Form_FormVideosGestionArbolVideo | formulario | src/forms/Form_FormVideosGestionArbolVideo.cls | 51 | 21 | 6 | igual | pendiente | sin ficha todavía |
| Form_FormVideosGestionArbolVideoDatos | formulario | src/forms/Form_FormVideosGestionArbolVideoDatos.cls | 408 | 33 | 8 | igual | pendiente | sin ficha todavía |
| Form_FormVideosUsuarios | formulario | src/forms/Form_FormVideosUsuarios.cls | 171 | 16 | 6 | igual | pendiente | sin ficha todavía |
| Form_Formulario1 | formulario | src/forms/Form_Formulario1.cls | 5 | 9 | 1 | igual | analizado | [formularios/Form_Formulario1.md](formularios/Form_Formulario1.md) |

Totales (formularios): 13 analizados · 0 no aplicables · 15 pendientes.
Totales (controles `Begin`): 740 · procedimientos: 247 · líneas de `.cls`: 6,246.

### Módulos y clases

Pendiente de análisis. Módulos: `Config_BackendHelper.bas`, `Constructor.bas`, `Critografia.bas`, `Funciones Genreales.bas`, `JsonConverter.bas`, `Módulo1.bas`, `Variables Globales.bas`, `mdlCursor.bas`. Pruebas (9): `Test.bas`, `Test_BackendBootstrap.bas`, `Test_BackendConfigHelper.bas`, `Test_ConfigFixtures.bas`, `Test_Helper.bas`, `Test_LanzaderaSuite.bas`, `Test_LocalEveProvisioning.bas`, `Test_ManifestContracts.bas`, `Test_TestingHarness.bas`. Clases: `Aplicacion.cls`, `AplicacionApertura.cls`, `Categoria.cls`, `ComprobacionAccesoVideos.cls`, `Conexion.cls`, `Correo.cls`, `Entorno.cls`, `HTML.cls`, `HistoricoPass.cls`, `Menu.cls`, `Usuario.cls`, `UsuarioAplicacionPermisos.cls`, `Video.cls`, `Visionado.cls`.

## Delta frente al estudio previo

En el código, **cero**: los formularios, módulos y clases coinciden con el snapshot. El único avance de `main` respecto a la rama `staging` (`63ba5e0`, baseline del estudio previo) son 4 commits de documentación de agentes (`9d1f6f4`, `e00269f`, `9bb143c`, `1474e8e`; 2026-06-08 a 2026-07-18). La clasificación por capacidades queda para `capacidades.md`.

## Capacidades

Pendiente. Las capacidades del [estudio previo](../capabilities.md) son contexto, no evidencia.

## Discrepancias abiertas

| Id | Discrepancia | Evidencia | Estado |
|---|---|---|---|
| D-LAN-1 | #1074 cita como fecha del commit de partida 2026-10-21; el commit `1474e8e` es del 2026-07-18 (el 2026-10-21 es el día del recuento). | `git log -1 1474e8e` | Resuelta: se cita la fecha real. |
| D-LAN-2 | El baseline del estudio previo es `63ba5e0` (rama `staging`) y el commit de partida es `main`. | `git rev-list --count origin/staging..HEAD` = 4, solo documentación de agentes | Resuelta en lo mecánico: sin cambio de código entre ambos. |
| D-LAN-3 | El ERD de origen no se ha extraído: `dysflow` no está disponible en el entorno de trabajo y el backend `Lanzadera_Datos.accdb`, presente en el clon, no es legible con herramientas de solo lectura (`mdb-tables` lo rechaza como base de Access; probable cifrado por contraseña). | `Lanzadera_Datos.accdb` en la raíz del clon | Abierta, **pendiente del propietario**: extraer en su máquina Windows, con `dysflow` y en solo lectura, los metadatos (tablas, claves, índices, relaciones, nulabilidad). |

## Hallazgos del legacy

Comportamientos sospechosos detectados al analizar el código. Son hallazgos del estudio: no se corrigen, no generan issues en el repositorio de Lanzadera y cada uno queda **a verificar** en ejecución. Los que afectan a autenticación se marcan para el propietario por su sensibilidad.

| Id | Hallazgo | Evidencia | Ficha | Estado |
|---|---|---|---|---|
| H-LAN-1 | Si el usuario de Windows es administrador, cualquier correo registrado entra sin contraseña; la rama de contraseña caducada es inalcanzable porque `Login` devuelve antes de evaluar la caducidad. | `Form_FormLogin.cls:117-118`; `Funciones Genreales.bas:1409-1437` | [Form_FormLogin](formularios/Form_FormLogin.md) | a verificar |
| H-LAN-2 | El bloqueo usa `DateDiff("m", Now(), FechaBloqueo) < 20` (meses, orden invertido): no caduca nunca; `NumeroVecesIntentado` nunca se incrementa y el control de 5 intentos no actúa. | `Funciones Genreales.bas:1385`; `Form_FormLogin.cls:6, 71` | [Form_FormLogin](formularios/Form_FormLogin.md) | a verificar |
| H-LAN-3 | La contraseña de un solo uso se guarda y compara en claro; cualquiera puede resetear la contraseña de otro usuario conociendo solo su correo; `ObtenerContraseña` descarta el resultado de `FormatoCorreoCorrecto`. | `Usuario.cls:763`; `Funciones Genreales.bas:1418`; `Form_FormLogin.cls:22-31`; `Form_FormObtenerContraseña.cls:20, 25-34` | [Form_FormObtenerContraseña](formularios/Form_FormObtenerContraseña.md) | a verificar |
| H-LAN-4 | El cambio de contraseña no pide la actual; el mensaje habla de «las últimas 10 veces» pero `PassRepetida` busca en todo el histórico; el manejador de `LoginCorrecto` compara `Err.Number <> ""` (tipos incompatibles). | `Form_FormUsuarioCambioPass.cls:17-28`; `Usuario.cls:911, 971-974`; `Variables Globales.bas:739` | [Form_FormUsuarioCambioPass](formularios/Form_FormUsuarioCambioPass.md) | a verificar |
| H-LAN-5 | `ActDesActMenu` con `p_EsActivar = No` usa `i` (siempre 0) en vez de `p_Numero`, de modo que los menús nunca se desactivan; `Form_Load` deshace la pestaña de «última aplicación» y fuerza siempre «fuera de la oficina». | `Menu.cls:83, 89`; `Form_FormMenuPrincipalAdmin.cls:52-70`; `Form_FormMenuPrincipalCalidad.cls:44-64` | [Form_FormMenuPrincipalAdmin](formularios/Form_FormMenuPrincipalAdmin.md) | a verificar |
| H-LAN-6 | El menú «Perfil» de Calidad no tiene `OnClick`: pulsarlo no hace nada; el cambio de contraseña cierra los menús de administrador y de usuario pero no el de Calidad. | `Form_FormMenuPrincipalCalidad.form.txt:615-650, 1425`; `Form_FormUsuarioCambioPass.cls:44-49` | [Form_FormMenuPrincipalCalidad](formularios/Form_FormMenuPrincipalCalidad.md) | a verificar |
| H-LAN-7 | `On Error Resume Next` en toda la navegación del menú de usuario oculta los fallos; «Salir» usa `acQuitSaveNone` en el de administrador y `acQuitSaveAll` en los otros dos. | `Form_FormMenuPrincipalUsuario.cls:152, 366`; `Form_FormMenuPrincipalAdmin.cls:117`; `Form_FormMenuPrincipalCalidad.cls:159` | [Form_FormMenuPrincipalUsuario](formularios/Form_FormMenuPrincipalUsuario.md) | a verificar |
| H-LAN-8 | El alta de una aplicación parece inalcanzable (el botón Registrar queda deshabilitado sin aplicación activa); si se guardara, `Registrar` abre la tabla entera y sobrescribe la primera fila con `.Edit`, y el ID calculado no se asigna al campo. | `Form_FormAplicacionDatosGenerales.cls:90-93`; `Aplicacion.cls:1867-1878, 1936-1937` | [Form_FormAplicacionDatosGenerales](formularios/Form_FormAplicacionDatosGenerales.md) | a verificar |
| H-LAN-9 | La contraseña de la aplicación se muestra sin máscara y se guarda tal cual; `HaHabidoCambios` compara `NombreIconoLanzadera` dos veces y omite `ConIconoEnLanzadera` y `NombreIconoParaArbol`; un nombre con apóstrofo rompe el SQL del filtro. | `Form_FormAplicacionDatosGenerales.form.txt:1201-1214`; `Aplicacion.cls:1895-1896`; `Form_FormAplicacionDatosGenerales.cls:36, 48, 201, 204`; `Constructor.bas:922` | [Form_FormAplicacionDatosGenerales](formularios/Form_FormAplicacionDatosGenerales.md) | a verificar |
| H-LAN-10 | `FormEnConstruccion` y `Formulario1` están huérfanos: ningún código de `src` los abre. | `Form_FormEnConstruccion.cls:3`; `Form_Formulario1.cls:4-6` | [Form_FormEnConstruccion](formularios/Form_FormEnConstruccion.md) | a verificar |
| H-LAN-11 | `Aplicacion.Lanzar` comprueba el icono cuando debería comprobar el BMP; `getVersion` pierde los errores de lectura; al fallar un lanzamiento salen dos mensajes; `LimpiarCache` ignora la colección que recibe. | `Aplicacion.cls:1407`; `Funciones Genreales.bas:274-275, 1838, 2998` | [Form_FormAplicaciones](formularios/Form_FormAplicaciones.md) | a verificar |

## Método del recuento

- El clon vive en `.atl/sources/lanzadera-1474e8e/` (ruta ignorada, solo lectura) y está en `HEAD` desacoplado en el commit fijado.
- Controles: líneas `^\s*Begin ` del `.form.txt` anteriores al código embebido. Procedimientos: líneas `(Private |Public )?(Sub|Function) ` del `.cls`.
- Comparación con el snapshot: normalización descrita arriba, sobre `.cls` y `.bas`. No se comparan los `.form.txt`.
- Recuento de CodeGraph: tabla `nodes` de `.codegraph/codegraph.db` del clon.
