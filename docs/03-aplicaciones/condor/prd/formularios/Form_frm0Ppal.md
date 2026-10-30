# Form_frm0Ppal

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frm0Ppal.cls` (322 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Menú principal de los roles Administrador y Calidad: buscar solicitudes, ver tareas pendientes, crear solicitudes, suplantar a un técnico y abrir las opciones de administración. Es **no enlazado**. Las citas son `Form_frm0Ppal.cls:línea` salvo indicación; `FU.bas` abrevia `modules/FUNCIONES UTILES.bas`, `VG.bas`, `modules/Variables Globales.bas`, y `SolSrv`, `classes/SolicitudServicio.cls`.

## Entrada y salida

- **Se abre desde:** Form_frmSplash por rol (`Form_frmSplash.cls:58-59`) y Form_frm0PpalTecnico al dejar de suplantar (`Form_frm0PpalTecnico.cls:127`).
- **Abre:** Form_frm0OtrosAdmin (`:40`), frmBuscarSolicitudes (`:55, 91`), frmAltaSolicitud (`:111`), frmGestionSolicitud (`:289`), Form_frm0PpalTecnico (`:201`); manual y vídeo externos (`:20, 224`).
- **Refresco externo:** `RefrescarTablerosPrincipales` y frmGestionSolicitud llaman a `ActualizarContadorTareas` (`VG.bas:170-171`; `Form_frmGestionSolicitud.cls:2401-2402`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` | Título, versión, visibilidad de `cmdAdminOpciones`, combo de técnicos y contador | `:149-172`, `.form.txt:221` |
| `cmdBuscarSolicitudes`, `shpSolicitudes` | botón, forma | `OnClick` | Abre frmBuscarSolicitudes sin argumentos | `:48-65, 316-318`, `.form.txt:19986, 20064-20066` |
| `cmdTareasPendientes`, `shpTareas`, `lblContadorTareas` | botón, forma, etiqueta | `OnClick` | Abre frmBuscarSolicitudes con `OpenArgs = "PENDIENTES"` | `:85-101, 266-268, 320-322`, `.form.txt:19001, 19023-19025, 21602-21605` |
| `cmdNuevaSolicitud`, `shpNueva` | botón, forma | `OnClick` | Abre frmAltaSolicitud y escucha sus eventos | `:103-120, 312-314`, `.form.txt:20041, 20966-20968` |
| `cboUsuariosTecnicos` | combo (lista de valores) | — | Se rellena en `Form_Load` con `Entorno.ResponsablesTecnicos` | `:161`, `.form.txt:21737-21739` |
| `cmdImpersonar` | botón | `OnClick` | Suplanta al técnico elegido | `:174-214`, `.form.txt:21771-21773` |
| `cmdAdminOpciones` | botón | `OnClick` | Abre Form_frm0OtrosAdmin; visible solo para Administrador | `:38-46, 159`, `.form.txt:21879-21881` |
| `cmdAbrirAyuda` («Ayuda Workflow»), `ComandoAyuda` | botones | `OnClick` | Ambos abren `Entorno.URLManual` si el fichero existe | `:8-32, 123-147`, `.form.txt:18870-18872, 18911-18913` |
| `cmdSalir` | botón | `OnClick` | Confirma, cierra conexiones y sale sin guardar | `:67-83`, `.form.txt:21639-21641` |
| `lblBloqueo` | etiqueta oculta | — | Aviso de espera durante operaciones | `:50-51`, `.form.txt:21559-21572` |
| `lblVersion` | etiqueta | — | `Entorno.VersionAplicacion` (clave INI) | `:160`; `classes/Entorno.cls:214-220` |

## Flujos

### Normal

1. `Form_Load`: aviso «Cargando entorno...» (`:152-155`); título `TituloUsuarioConectado` (`:158`; `classes/Entorno.cls:565-590`); `cmdAdminOpciones` visible si `rolUsuario = Administrador` (`:159`); combo con jefes de proyecto de expedientes PECAL sin fecha de baja (`:161`; `classes/UsuarioServicio.cls:116-121`; `FU.bas:538-569`).
2. `ActualizarContadorTareas` pide `getSolicitudesViewModel(soloPendientes:=True)` y muestra el número si es mayor que 0 (`:238-265`). Para estos roles, pendientes = estado interno distinto de «desarrollo técnico» (3) y «aprobada» (8) (`SolSrv:113-116`; `modules/modEnumeradores.bas:39, 44`).
3. **Nueva solicitud:** si frmAltaSolicitud ya está abierto le da el foco y vuelve a abrirlo (`:106-112`). Al recibir `SolicitudCreadaConExito` (`Form_frmAltaSolicitud.cls:505`), refresca el contador, muestra «Solicitud creada correctamente.», abre frmGestionSolicitud con el identificador en `OpenArgs` y llama a `AbrirFormularioDeDatos` (`:271-297`; `Form_frmGestionSolicitud.cls:2328-2333`).
4. **Suplantar:** exige técnico seleccionado (`:180-184`); carga el usuario con permisos por identificador (`:191`; `classes/UsuarioServicio.cls:17-40`); fija usuario activo, `rolUsuario = Tecnico` y `g_blnImpersonando = True` (`:197-199`); abre el menú técnico y se cierra (`:201-202`). El usuario real se conserva en `m_ObjUsuarioReal` (`VG.bas:27, 115`).

### Alternativos

- Calidad no ve `cmdAdminOpciones` (`:159`); el resto de botones es igual para ambos roles.
- Ayuda: si el fichero no existe, «No se pudo obtener el manual de uso.» (`:22`) o «...el documento de ayuda.» (`:137`).
- Salir: si no se confirma, no hace nada (`:71-74`).

### Errores

Los handlers usan `ManejadorErroresFormulario`: oculta `lblBloqueo`, muestra el mensaje según el número (513 validación, 514 operativo, resto técnico) y registra en `tbLogErrores` (`FU.bas:588-676, 783-798`). Un error en `Form_Load` cierra Access con `acQuitSaveAll` (`:169-170`). Técnico inexistente: error 513 «No se pudo encontrar al usuario seleccionado.» (`:193-195`). Un fallo del contador solo oculta la etiqueta (`:262-264`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Contador de pendientes | `tbSolicitudes`, `tbEstados`, `TbExpedientes`, `TbExpedientesResponsables`, `TbUsuariosAplicaciones` | SQL dinámico con `QueryDef` | `SolSrv:86-164`; `modules/RepositorioComun.bas:80` |
| Técnicos del combo | `TbExpedientes`, `TbExpedientesResponsables`, `TbUsuariosAplicaciones` (base de expedientes) | SQL fijo | `classes/UsuarioServicio.cls:116-121` |
| Usuario suplantado | `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` (Lanzadera) | `QueryDef` parametrizada | `modules/UsuarioRepositorio.bas:22-42`; `modules/AplicacionRepositorio.bas:24-27` |
| Registro de errores | `tbLogErrores` | `INSERT` dinámico | `FU.bas:792-798` |

Sin transacciones.

## Coexistencia

La consulta de pendientes usa `IIf`, `FIRST()` y comodín `*` en `LIKE` (`SolSrv:86, 94, 135`), propios de Jet/ACE. La subconsulta compara `EsJefeProyecto = 'SÍ'` (`SolSrv:88`) y la de técnicos `'Sí'` (`classes/UsuarioServicio.cls:119`): Jet no distingue mayúsculas; PostgreSQL sí. `tbLogErrores` evalúa `Now()` y `Environ('USERNAME')` dentro del SQL (`FU.bas:793`). El combo mezcla dos bases (expedientes y Lanzadera).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- «Ayuda Workflow» abre el manual, no `URLAyudaWorkflow` (`:13`; `classes/Entorno.cls:458-467`).
- `lblVideo_Click` (`:216-233`) no tiene control `lblVideo` en el diseño.
- Al suplantar, el título muestra la sesión real con rol «Técnico», porque el colectivo se calcula con `rolUsuario` ya cambiado (`:198`; `classes/Entorno.cls:576-587`).
- `tbLogErrores` guarda `suplantadoPor` siempre `Null` (`FU.bas:793`).
- Con frmAltaSolicitud abierto, `SetFocus` no sale y repite `OpenForm` (`:106-112`).
- Doble aviso de alta: frmAltaSolicitud ya muestra un mensaje (`Form_frmAltaSolicitud.cls:501-504`) y este formulario otro (`:283-285`).
- El contador carga todas las solicitudes pendientes para contarlas (`:250-253`).
- `Resume LimpiarYSalir` tras `Application.Quit` es inalcanzable (`:170-171`).
- El diseño guarda una lista de valores con nombres reales, sobrescrita al cargar (`.form.txt:21739`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: suplantación auditada con identidad real y suplantada; contador servido por un recuento en base de datos.

## Criterios de aceptación

Propuestos, no aprobados:

1. Administrador y Calidad ven el menú; solo Administrador ve las opciones de administración.
2. El contador coincide con las solicitudes en estado distinto de desarrollo técnico y aprobada, y se oculta si es 0.
3. Tareas pendientes abre la búsqueda filtrada por pendientes; Solicitudes, sin filtro.
4. Tras crear una solicitud se abre su gestión con la ficha de datos.
5. Suplantar sin técnico seleccionado muestra el aviso; con técnico, abre el menú técnico como ese usuario.
6. Salir pide confirmación.

## Pendientes

- Contenido de frmAltaSolicitud, frmGestionSolicitud y frmBuscarSolicitudes en sus fichas.
- Código de `EjecutarConsulta` frente a tablas vinculadas: paso de parámetros y tipo de `Recordset` (`modules/RepositorioComun.bas:39-52`).
- Formato exacto de `TituloUsuarioConectado` sin suplantación (`classes/Entorno.cls:590`).
