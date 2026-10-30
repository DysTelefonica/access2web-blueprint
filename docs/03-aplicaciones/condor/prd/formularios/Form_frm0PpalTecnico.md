# Form_frm0PpalTecnico

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frm0PpalTecnico.cls` (195 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Menú principal del rol Técnico, también usado cuando un administrador o calidad suplanta a un técnico: tareas pendientes, solicitudes propias, ayuda y vuelta a la sesión real. Es **no enlazado**. Las citas son `Form_frm0PpalTecnico.cls:línea` salvo indicación; `FU.bas` abrevia `modules/FUNCIONES UTILES.bas`, `VG.bas`, `modules/Variables Globales.bas`, y `SolSrv`, `classes/SolicitudServicio.cls`.

## Entrada y salida

- **Se abre desde:** Form_frmSplash con rol Técnico (`Form_frmSplash.cls:60-61`) y Form_frm0Ppal al suplantar (`Form_frm0Ppal.cls:201`).
- **Abre:** frmBuscarSolicitudes (`:19, 86`) y Form_frm0Ppal al dejar de suplantar (`:127`); manual externo (`:44`).
- **Refresco externo:** `RefrescarTablerosPrincipales` y frmGestionSolicitud llaman a `ActualizarContadorTareas` (`VG.bas:174-175`; `Form_frmGestionSolicitud.cls:2403-2404`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` | Título, versión, botones según suplantación y contador | `:58-74`, `.form.txt:221` |
| `cmdTareasPendientes`, `shpTareas`, `lblContadorTareas` | botón, forma, etiqueta | `OnClick` | Abre frmBuscarSolicitudes con `OpenArgs = "PENDIENTES"` | `:9-30, 186-188, 194-196`, `.form.txt:974-998, 2981-2984` |
| `cmdBuscarSolicitudes`, `shpSolicitudes` | botón, forma | `OnClick` | Abre frmBuscarSolicitudes e inyecta el filtro del técnico activo | `:77-111, 190-192`, `.form.txt:1959-1960, 2013-2015` |
| `ComandoAyuda` | botón | `OnClick` | Abre `Entorno.URLManual` si el fichero existe | `:32-56`, `.form.txt:884-886` |
| `cmdVolverASesionAdmin` («Volver a Sesión Ant.») | botón | `OnClick` | Restaura la sesión real; visible solo al suplantar | `:68, 113-135`, `.form.txt:3107-3109` |
| `cmdSalir` | botón | `OnClick` | Cierra conexiones y sale sin guardar; oculto al suplantar | `:69, 137-147`, `.form.txt:3018-3020` |
| `lblBloqueo`, `lblVersion` | etiquetas | — | Aviso de espera; versión desde el INI | `:12, 65`, `.form.txt:858, 2938-2951` |

## Flujos

### Normal

1. `Form_Load`: `Ajustar Me`, título `TituloUsuarioConectado` y versión (`:61-65`; `classes/Entorno.cls:565-590, 214-220`); muestra `cmdVolverASesionAdmin` u `cmdSalir` según `g_blnImpersonando` (`:68-69`).
2. `ActualizarContadorTareas` pide `getSolicitudesViewModel(soloPendientes:=True, soloMisSolicitudes:=True)` (`:167`). Con rol Técnico filtra estado interno 3 («desarrollo técnico») y expedientes donde el usuario activo figura en `TbExpedientesResponsables` (`SolSrv:110-112, 124-126`; `modules/modEnumeradores.bas:39`). Muestra el número si es mayor que 0 (`:171-177`).
3. **Solicitudes:** abre frmBuscarSolicitudes, cuyo `Form_Load` ya marca «solo mis solicitudes» y busca (`Form_frmBuscarSolicitudes.cls:37-55`); después asigna `FiltroInicial.idResponsableTecnico` y repite la búsqueda (`:86-99`).
4. **Tareas pendientes:** la búsqueda abre con «solo pendientes» marcado (`:19`; `Form_frmBuscarSolicitudes.cls:52-53`).

### Alternativos

- **Suplantación activa:** «Volver a Sesión Ant.» restaura `m_ObjUsuarioActivo = m_ObjUsuarioReal`, `g_blnImpersonando = False` y `rolUsuario = rolUsuarioReal`, abre Form_frm0Ppal y se cierra (`:117-128`). No hay botón de salida (`:69`).
- Ayuda sin fichero: «No se pudo obtener el documento de ayuda.» (`:46`).

### Errores

Los handlers usan `ManejadorErroresFormulario` (`FU.bas:588-676`), que registra en `tbLogErrores` (`FU.bas:792-798`). Un error en `Form_Load` muestra el mensaje y deja el formulario abierto (`:72-73`). Un fallo del contador solo oculta la etiqueta (`:181-183`).

## Reglas, validaciones y cálculos

- La salida directa no pide confirmación, a diferencia de Form_frm0Ppal (`:137-140`; `Form_frm0Ppal.cls:71`).
- El rol del contador es el global `rolUsuario`, que la suplantación fija a Técnico (`Form_frm0Ppal.cls:198`; `SolSrv:110, 124`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Contador de pendientes | `tbSolicitudes`, `tbEstados`, `TbExpedientes`, `TbExpedientesResponsables`, `TbUsuariosAplicaciones` | SQL dinámico con `QueryDef` y parámetro `p_IDUsuarioActivo` | `SolSrv:86-164` |
| Registro de errores | `tbLogErrores` | `INSERT` dinámico | `FU.bas:792-798` |

Sin transacciones.

## Coexistencia

Mismas adaptaciones que la consulta de Form_frm0Ppal: `IIf`, `FIRST()`, `'SÍ'` sensible a mayúsculas en PostgreSQL (`SolSrv:86-94`). El `EXISTS` correlacionado con parámetro DAO (`SolSrv:125`) debe verificarse con tablas vinculadas.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El filtro inyectado no tiene efecto: con «solo mis solicitudes» marcado se ignora `idResponsableTecnico` (`SolSrv:139`), y la búsqueda se ejecuta dos veces (`Form_frmBuscarSolicitudes.cls:55`; `:99`).
- «Mis solicitudes» incluye cualquier responsable del expediente, no solo el jefe de proyecto (`SolSrv:125`).
- Las acciones hechas al suplantar se registran con el usuario de Windows y `suplantadoPor = Null` (`FU.bas:793`).
- El manejador de `ComandoAyuda` se identifica como `cmdAbrirAyuda_Click` (`:54`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: indicar de forma permanente la suplantación y auditar la identidad real.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un técnico ve el contador de sus solicitudes en desarrollo técnico; oculto si es 0.
2. Tareas pendientes y Solicitudes abren la búsqueda limitada a sus expedientes.
3. Durante la suplantación se ve «Volver a Sesión Ant.» y no «Salir»; al volver se recupera el rol real.

## Pendientes

- Definición de «mis solicitudes» para el técnico: cualquier responsable o solo jefe de proyecto (`SolSrv:88` frente a `:125`).
- `src/forms/frm0PpalTecnico.form.txt` duplica el diseño y el código con otro nombre; falta decidir cuál rige.
