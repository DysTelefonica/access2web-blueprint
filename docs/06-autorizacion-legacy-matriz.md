# Matriz actual de autorización legacy

## Alcance y baseline

Exploración estática y de esquema agregado, no diseño futuro. Baseline funcional: `staging`; `main` solo referencia publicada. No se ejecutaron tests ni se afirma comportamiento passing. Confianzas: **Verified-static** (source/CodeGraph), **Verified-runtime-schema/aggregate** (Dysflow read-only y agregados), **Intended** (documentación), **Likely** (inferencia acotada), **Divergent** (staging y main o catálogo/código no alineados).

| Aplicación | Baseline inspeccionada | `main` referencia | ID catálogo / backend de autorización | Entrada de identidad y resolución |
|---|---|---|---|---|
| Lanzadera | `staging` `63ba5e01617fdda857503d43151f06bb7bc11829` | `main` `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` | 12 / `C:\00repos\datos\Lanzadera_Datos.accdb` | Login por correo + contraseña; bypass de contraseña para administrador de máquina; `Usuario` y `UsuarioAplicacionPermisos`. **Verified-static** |
| Gestion_Riesgos | `staging` `2767bb0224169b2386d9d84508212270de9f8791` | `main` `36fcf60161fa5ec84c3956b1308537ddf2579256` | 5 / Lanzadera vía `getdbLanzadera` | `VBA.Command`/red según arranque; `Usuario`; flags globales `EsTecnico`, `EsCalidad`, `EsAdministrador`. **Verified-static; Divergent** |
| No_Conformidades | `staging` `ab4abd16603d2311feaa1b494eca9486265c48e6` | `main` `a7e998493e90fae6b50b50b94955847ddca6c5c2` | 8 / Lanzadera vía `getdb`/constructor | usuario de red; `UsuarioAplicacionPermisos`; administración/calidad/técnico. **Verified-static; Divergent** |
| Condor | `staging` `43087ad9868c94ad7c9b8f11692df67be595a600` | `main` `a1e79fe05f17462ec2d3c1d8ba839d12cfa2f4ed` | 23 / `getdbLanzadera` + backend Condor | usuario de red; `rolUsuario` enum (`Tecnico`, `Administrador`), con impersonación. **Verified-static; Divergent** |
| HPS_Solicitudes | `main` `cfa1937ab05a9e40150ea9236243898de96f84da` (**sin staging local; fallback explícito**) | misma referencia | 22 / `C:\00repos\datos\Solicitudes_HPS_datos.accdb` localizado; configuración Dysflow ausente | `VBA.Command` correo o `Wscript.Network`; `UsuarioAplicacionPermisos`. **Verified-static; Likely runtime target** |
| HPS | `staging` `8fcedf31b44a7a4d599a9269c9131f8cc42897c0` | `main` mismo commit | 17 / Lanzadera vía `getdbLanzadera` | `VBA.Command` o usuario de red; administrador/técnico desde permiso por aplicación. **Verified-static** |
| Brass | `main` `91bd8c66c7bd5ec869c2417dc659f7ecf84f566f` (**sin staging local; fallback explícito**) | misma referencia | 6 / `C:\00repos\datos\Gestion_Brass_Gestion_Datos.accdb` localizado; configuración Dysflow ausente | usuario de red; `UsuarioAplicacionPermisos`; admin visible en menú. **Verified-static; Likely runtime target** |
| Expedientes | `staging` `0946b6a0a40acf4fb88eb62e435ec3f76f8a234d` | `main` `535c38a04da5f40aedfd0f1fe6a8464199807918` | 19 / Lanzadera vía `getdb` | usuario de red; administrador/técnico y permisos por aplicación. **Verified-static; Divergent** |

## Modelo común observado

`TbUsuariosAplicaciones` identifica la persona (`CorreoUsuario`, `UsuarioRed`, `ID`, baja/activo y, según app, `EsAdministrador`). `TbUsuariosAplicacionesPermisos` asigna por pareja `CorreoUsuario + IDAplicacion` flags de texto `Sí/No`: `EsUsuarioAdministrador`, `EsUsuarioTecnico`, `EsUsuarioCalidad`, `EsUsuarioEconomia`, `EsUsuarioSecretaria`, `EsUsuarioSinAcceso` y `EsUsuarioCalidadAvisos`. `Constructor.getAplicacionesPermisos` hidrata un diccionario indexado por `IDAplicacion`; `Usuario.Permisos` obtiene la entrada de la aplicación global `IDAplicacion`. **Verified-static**.

La cardinalidad efectiva no es “un rol por usuario”: puede haber varias filas por usuario en aplicaciones distintas y varias flags simultáneas dentro de una fila. La clase convierte texto a enums (`EnumSiNo`); los valores vacíos/null no equivalen de forma demostrada a `No`. **Verified-static; Divergent**.

## Matriz por aplicación: enforcement actual

| Aplicación | Roles/perfiles y asignación | Admin global/aplicación y denegación | UI/route/control | Non-UI y condiciones de recurso |
|---|---|---|---|---|
| Lanzadera | Siete flags; asignación administrable en formularios de usuarios/aplicaciones. | `EsAdministrador` global habilita mantenimiento; usuario sin registro o credenciales inválidas no entra. Sin permiso de app: botones de aplicación no quedan utilizables; no se probó un deny server-side uniforme. | `getBoton`/`Tag=IDAplicacion`; `AllowEdits` y `Enabled` en formularios de administración; filtro de vídeos por admin/calidad/técnico. | `Login` valida hash, bloqueo, caducidad y cambio inicial; `Aplicacion.Lanzar` valida aplicación y `EjecucionEnOficina`. **Verified-static** |
| Gestion_Riesgos | Administrador, técnico, calidad; también economía/avisos en catálogo. `getListaUsuarios` filtra por flag + `IDAplicacion`. | Menús completos para admin/calidad; vista técnica separada. Falta permiso/usuario produce error de arranque en `EVE`; deny explícito no trazado en cada caso de uso. | `Form_Form0BDOpciones` habilita menú; formularios de anexos, tareas y acciones alteran `Enabled`/visibilidad por técnico/calidad y estado. | Filtros y responsables de proyecto/riesgo; `FechaBaja` para activos. La mayoría de reglas de escritura observadas son UI. **Verified-static** |
| No_Conformidades | Administrador, calidad, técnico, economía, secretaria y avisos disponibles en clase/catálogo. | `EsAdministrador` controla edición/menú; calidad y técnico separan auditoría/parte-proyecto. Ausencia de permiso suele dejar navegación/control deshabilitado; no hay prueba de guardia DAO universal. | `Form_Form0BDOpciones*`; `Funciones Generales` habilita navegación por perfil, estado y formulario. | Condiciones por proyecto/auditoría, responsable, estado y excepción `Nemotecnico`; DAO usa `IDAplicacion` al listar usuarios. **Verified-static** |
| Condor | Enum local `rolUsuario`: técnico/admin; catálogo Lanzadera conserva flags más amplios. `AplicacionRepositorio` carga permisos por aplicación. | Admin ve `cmdAdminOpciones`; técnico usa menú simplificado y puede volver a sesión admin si hay impersonación. | `Form_frm0Ppal` y `Form_frm0PpalTecnico`; muchos controles `Enabled` dependen además del estado de solicitud. | Workflow, responsable, estado, expediente obligatorio y NC asociada; controles de alta/editado son evidencia UI, no autorización DAO demostrada. **Verified-static** |
| HPS_Solicitudes | Admin/técnico y helpers de permisos; consultas explícitas por `EsUsuarioAdministrador` y `EsUsuarioTecnico` + `IDAplicacion`. | Menú y opciones se adaptan a técnico; edición puede quedar bloqueada. Sin staging/config Dysflow, backend efectivo queda pendiente. | `Form_Form0BDOpciones` y formularios de solicitudes alternan `AllowEdits`, `Enabled`, búsqueda y filtros; técnico limita opciones. | responsable técnico/calidad, estado, solicitante y expediente; filtros `soloMisSolicitudes` son selección de datos, no sustituto de deny. **Verified-static; Likely** |
| HPS | Administrador y técnico derivados de fila global/permisos de HPS. | `Form_FormInicial` ofrece acceso lectura/escritura según permiso; consultas de usuarios bloquean alta/baja/edición para no admin. | Menú lateral y formularios de consulta controlan botones y `AllowEdits`; vista lectura explícita. | usuario HPS/SICA, expediente, solicitud y estado; no se identificó una capa de autorización no-UI común. **Verified-static** |
| Brass | Permisos por aplicación heredados; código visible rastrea principalmente admin. | Admin habilita acciones de mantenimiento; no se vio perfil técnico equivalente en el menú principal. | `Form_Form0BDOpciones`; CRUD de actividades condiciona botones también por estado/facturación/anexos. | evento, actividad, equipo/material, facturación y responsables; condiciones de negocio locales pueden ocultar acciones sin ser permisos. **Verified-static; Likely** |
| Expedientes | Admin/técnico; helpers de formulario consultan `EsAdministrador` y `EsTecnico`. | Admin habilita alta, gestor de entidades y buscador técnico; técnico recibe buscadores/vistas especializadas. | `Form_Form0BDOpciones` y CRUD de catálogos usan `Enabled`, `Visible`, `AllowEdits`; helpers JSON deciden algunos controles. | responsable por rol, expediente/estado, entidades, suministradores y ownership operativo; no se demostró enforcement uniforme fuera de UI. **Verified-static** |

## Agregados runtime seguros

Consulta read-only agregada a `C:\00repos\datos\Lanzadera_Datos.accdb`, sin usuarios ni correos: `TbAplicaciones` = 20 filas; `TbUsuariosAplicaciones` = 186; `TbUsuariosAplicacionesPermisos` = 622. Para las ocho aplicaciones: `ID 12 Lanzadera` 42 (admin 3/técnico 32/calidad 6/economía 3), `ID 5 Riesgos` 74, `ID 8 No Conformidades` 66, `ID 23 Condor` 69, `ID 22 Solicitudes HPS` 7, `ID 17 HPS` 9, `ID 6 Brass` 14, `ID 19 Expedientes` 63. Los conteos de flags se solapan y no son personas únicas. **Verified-runtime-schema/aggregate**.

## Cross-ecosystem synthesis

- **Contrato Lanzadera compartido:** identidad `VBA.Command`/red → `TbUsuariosAplicaciones` → `TbUsuariosAplicacionesPermisos` → diccionario `UsuarioAplicacionPermisos` → flags/enums. Se reutiliza conceptualmente en las ocho; la implementación y el `IDAplicacion` son locales. **Verified-static**.
- **Extensiones reales:** Condor introduce `rolUsuario` e impersonación; HPS separa lectura/escritura; Riesgos, No Conformidades y Solicitudes HPS tienen vistas técnico/calidad y filtros de responsable; Expedientes agrega administración de catálogos y responsables por rol; Lanzadera tiene administración del catálogo y lanzamiento. **Verified-static**.
- **No hay deny uniforme probado:** la evidencia dominante es ocultar/deshabilitar menús, controles y edición. Login, `EVE`, `Login`, `Aplicacion.Lanzar` y algunos filtros/consultas sí contienen checks explícitos; no se demostró una política central que proteja todos los helpers DAO. **Verified-static**.
- **Catálogo frente a código:** el catálogo contiene flags y aplicaciones adicionales no necesariamente visibles en cada implementación; hay null/vacíos y combinaciones solapadas. Esto es una divergencia de datos/código, no una decisión de normalización. **Divergent**.
- **Staging/main:** Lanzadera staging ya documenta cambios operativos; HPS coincide en commit; los demás staging muestran divergencias versionadas importantes o no tienen staging local. Ninguna delta fue interpretada como autorización vigente sin revisar su código. **Divergent**.

## Arranque: obtención de roles y selección del perfil efectivo

La siguiente tabla normaliza únicamente la cadena observada al iniciar cada aplicación: identidad → usuario → `IDAplicacion` → permisos → normalización → precedencia → perfil efectivo → primera vista → rechazo/error. **No es una propuesta de diseño.**

| Aplicación | Cadena de arranque observada y perfil efectivo | Primera vista y denegación/fallback |
|---|---|---|
| Lanzadera | `Form_FormLogin.Form_Load` → `EVEMin`; correo de login o SSID precargado → `Constructor.getUsuario` por `CorreoUsuario`; login normal (`Login`) salvo administrador de máquina. `Usuario.Permisos` indexa `ColAplicacionesPermisos` por `IDAplicacion`; `Constructor.getAplicacionesPermisos` carga el diccionario y conserva la primera fila por ID. `LoginCorrecto`: Administrador > Calidad > Usuario. | `FormMenuPrincipalAdmin`, `FormMenuPrincipalCalidad` o `FormMenuPrincipalUsuario`; usuario inexistente cierra la aplicación, credenciales/cuenta bloqueada/caducada derivan a error, cambio de contraseña o salida. La disponibilidad de botones usa `Tag=IDAplicacion`; no quedó probado un deny uniforme para cada aplicación. |
| Gestion_Riesgos | `Variables Globales.EVE` limpia estado, lee `TbConfiguracionBackends.IDAplicacion` (fallback técnico `1` si no hay fila), obtiene correo de `VBA.Command` o `Wscript.Network.UserName`, y `Constructor.getUsuario`. `EsAdministrador` → no técnico; si no, `EsCalidad` → no técnico; si no, se asigna Técnico. `EVE(p_CorreoUsuario)` permite impersonar y `Usuario.Permisos` resuelve por ID. | `Form0BDOpciones` o `Form0BDOpcionesTecnico`; el técnico puede volver al perfil inicial. Usuario ausente, baja o error de recursos aborta por `p_Error`; falta de permiso específico se manifiesta principalmente como controles deshabilitados. |
| No_Conformidades | `Variables Globales.EVE` lee configuración y selecciona `IDAplicacion` configurado; en pruebas usa `81`, en normal `8`. Identidad por `VBA.Command`, si falta por red; `constructor.getUsuario`; Administrador tiene precedencia sobre Calidad/Técnico y se fuerza Técnico=No. El diccionario de permisos se indexa por `IDAplicacion`. | `Form0BDOpciones`; en pruebas cambia título/Ribbon y conserva navegación. Falta de usuario/configuración/backend produce error fail-fast; varias restricciones de Técnico están comentadas o se aplican en formularios posteriores, por lo que la vista inicial no equivale a deny efectivo. |
| Condor | `Form_frmSplash.Form_Timer` → `EVE`; `VBA.Command` (correo) o `UsuarioServicio.getUsuarioConPermisos` por usuario de red; `AplicacionRepositorio` indexa permisos por `IDAplicacion`, pero el perfil arrancado es `Usuario.rol`. Se copia a `rolUsuario` y `rolUsuarioReal`; no se observa combinación de flags para resolver un perfil. | `rol.Administrador`/`Calidad` → `frm0Ppal`; `rol.Tecnico` → `frm0PpalTecnico`; cualquier otro valor lanza error 513 y muestra error crítico. Impersonación cambia `m_ObjUsuarioActivo`, `rolUsuario` y `g_blnImpersonando`; es sesión, no permiso persistido. |
| HPS_Solicitudes | `Variables Globales.EVE` fija `IDAplicacion=22`, toma correo de `VBA.Command` o `Wscript.Network`, y `constructor.getUsuario`. Admin se calcula primero; solo si no es admin se calcula Técnico; si ninguno, error “Usuario no autorizado”. `UsuarioAplicacionPermisos` normaliza texto `Sí/No`; la colección se indexa por ID. | `Form_Form0BDOpciones` (desde `frmSplash`); Técnico queda en solo lectura y no admin puede usar escritura. Usuario ausente o sin perfil aborta; el `EnPruebas`/local cambia entorno y vistas, no concede por sí mismo. No hay chequeo explícito de usuario inactivo en esta cadena. |
| HPS | `VariablesEntorno.EVE` obtiene correo de `VBA.Command` o red, carga `TbUsuariosAplicaciones`, fija `IDAplicacion=17` y exige Administrador o Técnico. La precedencia efectiva es Administrador > Técnico; permisos se consultan mediante `Usuario.Permisos`/diccionario por ID. | `Form_FormInicial` sincroniza y luego `FormInicial00Principal`; lectura fuera de oficina exige `PermisoEntradaLocalEnCasa`, escritura se deshabilita en modo local. Usuario inexistente o sin admin/técnico aborta; no se vio baja/inactivo comprobada en EVE. |
| Brass | `Variables Globales.EVE` fija `IDAplicacion=6`, toma `VBA.Command` o red y `Constructor.getUsuario`; carga la colección de permisos por ID, pero el arranque consume explícitamente `EsAdministradorCalculado`. No hay selector de perfil Técnico equivalente en la ruta observada: efectivo es Admin o no-Admin. | `Form_Form0BDOpciones.Form_Open`; Admin controla Ribbon en desarrollo y acciones de mantenimiento. Usuario ausente/error de entorno muestra error y sale; permiso ausente queda como no-admin, no como rechazo uniforme. Baseline es `main` por ausencia de staging. |
| Expedientes | `Variables Globales.EVE` resetea, lee configuración; `IDAplicacion` procede de `TbConfiguracionBackends` y queda vacío si no hay fila. Identidad por `VBA.Command` o red; `constructor.getUsuario`; Administrador > Calidad > Técnico, y los otros dos se fuerzan a No cuando uno gana. | `Form_Form0BDOpciones.Form_Open`; helper JSON decide alta/gestión/entidades/buscador técnico usando los flags. Usuario ausente o error de configuración aborta; sin permiso explícito no se demostró conversión a `Sin Acceso` ni deny DAO uniforme. |

### Reglas y peligros explícitos de la cadena

- `Constructor.getUsuario` devuelve `Nothing` ante cero filas; `getAplicacionesPermisos` devuelve colección vacía si no hay filas y descarta duplicados del mismo `IDAplicacion` conservando la primera. En `Usuario.Permisos`, aplicación sin clave devuelve `Nothing`; varias propiedades calculadas convierten ese caso a `No`, pero no todos los consumidores comprueban el error. **Peligro: ausencia silenciosa puede acabar en UI parcialmente habilitada.**
- La combinación de flags no es uniformemente multirrol: Riesgos, No Conformidades y Expedientes imponen precedencia; HPS_Solicitudes hace Admin > Técnico; Lanzadera hace Admin > Calidad > Usuario; Condor usa enum; Brass deja Admin/no-Admin. No se observó una regla común de `Sin Acceso` que gane siempre.
- No se observó en estos arranques que “permiso ausente” se convierta literalmente en Técnico; sí existe un **fallback implícito a Técnico** en EVE de Riesgos, No Conformidades y Expedientes cuando el usuario no resulta Admin ni Calidad. Debe distinguirse de una fila explícita `EsUsuarioSinAcceso=Sí`.
- Hay acoplamiento directo nombre/flag: `EsAdministradorCalculado`, `EsUsuarioTecnicoCalculado`, `EsUsuarioCalidadCalculado` deciden tanto perfil global como ruta de formulario. Los `Enabled`, `Visible`, `AllowEdits`, filtros de responsable y `SourceObject` son routing/UI; solo login, `EVE`, validación de recursos, permiso de entrada local y algunas consultas constituyen enforcement no-UI demostrado.
- **Contradicciones/gaps:** IDs de prueba (`81` en No_Conformidades), IDs configurables o ausentes (Riesgos/Expedientes), `IDAplicacion` fijo en otras apps, flags nulos/vacíos, baja no comprobada en HPS/HPS_Solicitudes, y permisos cargados pero no siempre exigidos antes de abrir la vista inicial. Staging/main se mantuvo separado: HPS coincide; HPS_Solicitudes y Brass son fallback `main`; las demás diferencias quedan **Divergent**, no resueltas.

**Símbolos inspeccionados:** `EVE`, `EVEMin`, `LoginCorrecto`, `Constructor.getUsuario`, `Constructor.getAplicacionesPermisos`, `Usuario.Permisos`, `UsuarioAplicacionPermisos.*Calculado`, `Form_FormLogin`, `Form_Form0BDOpciones` (y variantes), `Form_frmSplash`, `Form_frm0Ppal`, `Form_frm0PpalTecnico`, `VariablesEntorno`, `AplicacionRepositorio`, `rolUsuario` y helpers JSON de Expedientes. Evidencia **Verified-static**; no se ejecutaron formularios ni tests.

## Evidencia, límites y gaps

CodeGraph-VBA se llamó primero en cada worktree staging disponible y en los dos fallbacks main; se reutilizaron los índices propios existentes y no se inicializó, sincronizó ni reconstruyó ningún índice en esta pasada. Después se hicieron exploraciones enfocadas y referencias de source/forms. Dysflow `get_capabilities({cwd})` se ejecutó para los ocho targets y devolvió `dysflow.result/v1`, adapter `2.35.3`, 94 tools; se usaron únicamente lecturas/consultas agregadas, sin tests, sync, writes ni cambios de configuración. Los backends se contrastaron bajo `C:\00repos\datos`; HPS_Solicitudes y Brass quedaron sin target Dysflow configurado aunque sus `.accdb` fueron localizados.

Gaps: no se pudieron validar por runtime las rutas/configuraciones de HPS_Solicitudes y Brass; no se ejecutaron formularios ni casos negativos; no se inspeccionaron datos personales; no se resolvieron todos los bordes dinámicos (`TempVars`, DAO, nombres construidos, handlers). La matriz no autoriza una solución futura ni afirma seguridad efectiva.

## Disposiciones de arquitectura posteriores a la matriz

La matriz anterior describe **cómo funciona hoy** la cadena identidad → usuario → permisos → primera vista en cada aplicación. Esta subsección cruza esa evidencia con las decisiones APROBADAS que ya están en `docs/09-arquitectura-objetivo-y-principios.md`. La matriz **no autoriza** la solución futura; estas disposiciones sí la **orientan** sin cambiar lo anterior.

### Jerarquía de administración (dirección)

| Ámbito | Quién | Disposición |
|---|---|---|
| Cross-platform | **Administrador global** | Nombra administradores de aplicación, opera el CLI, configura health checks, gobierna UAT y libera releases con o sin UAT. |
| Limitado al módulo | **Administrador de aplicación** (donde aplique) | Solo asigna usuarios activos a roles existentes en su módulo. **No** registra/baja usuarios ni crea roles. Sus capacidades concretas las define cada módulo. |
| Identidad de la aplicación | Persona responsable | Figura separada; no acumula capacidades por defecto. |

### Capabilities y políticas (dirección)

- **Capabilities declaradas por módulo**: cada módulo declara sus capacidades estables (p. ej. `risk.assign`, `nc.sign-off`, `expediente.manage`). Los **grupos** componen roles verificables.
- **Políticas contextuales en código**: la autorización que depende de estado/recurso/condición de negocio se evalúa en el código del módulo, no en strings del catálogo. Capabilities (estables) y políticas (contextuales) coexisten.
- **Backend autoritativo, UI derivada**: la UI consume el endpoint de capabilities/políticas; el backend rechaza lo no autorizado aunque la UI lo muestre.
- **Vistas por módulo**: cada módulo decide vista única o variantes especializadas. Por defecto, una sola vista; especializada cuando la complejidad del flujo lo justifique (p. ej. Expedientes). La selección se guía por capabilities/políticas efectivas.

### Suplantación para pruebas

- **Hoy**: la matriz muestra que **solo Condor** implementa un flujo de impersonación observable, mutando `m_ObjUsuarioActivo`/`rolUsuario`/`g_blnImpersonando`. Es sesión local, no permiso persistido.
- **Disposición objetivo**: la suplantación es una capacidad **reservada al administrador global**, con doble identidad visible y auditoría completa. Los desarrolladores usan **personas sintéticas** o piden al administrador global que la ejecute. No se comparte ninguna credencial de usuario.

### UAT, releases y excepciones

- **UAT no espeja acceso de producción**: cada ciclo declara explícitamente sus participantes y perfil de aplicación.
- **Gobernanza solo administrador global**: configuración de participación, perfiles y acceso al ciclo.
- **Excepciones auditadas y visibles**: un release puede publicarse con casos UAT fallidos o sin UAT; la excepción se conserva con su evidencia y aparece en el historial de cambios para los usuarios.

### Dirección de sustitución por aplicación

| Aplicación | Modelo legacy observado | Dirección objetivo |
|---|---|---|
| Lanzadera | Catálogo + permisos + admin por UI; enforcer UI + helpers parciales; telemetría de ubicación. | Lanzadera queda como **módulo de administración de plataforma**. Identidad/permisos/auditoría migradas al núcleo. Telemetría de ubicación **retirada**. Configuración técnica externalizada. |
| Gestion_Riesgos | Permiso por flag + `EVE`; admin/calidad/técnico; ausencia de técnico = error en arranque. | Capabilities estables + políticas contextuales (admin asigna usuarios a riesgos); admin de aplicación con alcance delegado. |
| No_Conformidades | Banderas + precedencia; técnico con varias restricciones diferidas a formularios posteriores. | Mismas reglas, pero el evaluador server-side valida cada operación. Las políticas "técnico solo lectura" deben ser código, no flags de formulario. |
| Condor | Enum `rolUsuario` + impersonación local. | Suplantación restringida al administrador global; capacidades estables por catálogo. |
| HPS_Solicitudes | `IDAplicacion=22`; admin > técnico; sin check de inactivo en `EVE`. | Debe añadirse comprobación de usuario inactivo en arranque y server-side. UAT gobernado por admin global. |
| HPS | `IDAplicacion=17`; admin/técnico; `PermisoEntradaLocalEnCasa` para entrada local. | Permiso de entrada local se modela como capability/política; admin de aplicación solo asigna. |
| Brass | `EsAdministradorCalculado`; ausencia de admin = no-admin. | Igual modelo, con capabilities estables y admin global como único que nombra administradores de aplicación. |
| Expedientes | `IDAplicacion` configurable; helper JSON decide alta/gestión. | Helper JSON se sustituye por capabilities/políticas evaluadas server-side. Variante UI especializada justificada (gestión + vista técnica). |

Las columnas de la matriz original **no se reinterpretan** a la luz de estas disposiciones; siguen describiendo el legacy. La correspondencia con el modelo objetivo se trabaja en cada lote de descubrimiento por aplicación.
