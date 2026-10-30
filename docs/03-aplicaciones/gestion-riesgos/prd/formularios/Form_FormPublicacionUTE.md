# Form_FormPublicacionUTE

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormPublicacionUTE.cls` (236 líneas) y `src/forms/FormPublicacionUTE.form.txt` (1373 líneas) · **PRD:** [README](../README.md)

Pantalla del anexo de evidencia de acuerdo con la UTE. Permite adjuntar a la edición activa un documento firmado con los miembros de la Unidad de Trámite, quitarlo y abrir el fichero ya registrado. Las citas son `Form_FormPublicacionUTE.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormPublicacionCalidad` | Pestaña de navegación lateral: el diseño del contenedor declara `NavigationTargetName = "FormPublicacionUTE"` | `FormPublicacionCalidad.form.txt:534` |
| `FormPublicacionTecnico` | Pestaña de navegación lateral, mismo mecanismo | `FormPublicacionTecnico.form.txt:1052` |

No hay ninguna llamada a `DoCmd.OpenForm` con este formulario en el clon: siempre llega como subformulario anidado, y la referencia de navegación vive solo en el diseño de los contenedores, así que quién navega a esta pestaña y con qué condición **no se puede confirmar leyendo código**. No abre ningún formulario: lanza el fichero del anexo en el equipo del usuario (`AbrirEnLocal`, `:208`) y escribe a través de la clase `Anexo` (`:45`, `:94`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Ajusta el tamaño, valida la edición activa y refresca | `:125-163`, `.form.txt:41` |
| `ImagenAnexo` | imagen | `OnClick` → `ImagenAnexo_Click` | Abre el fichero del anexo en local | `:192-225`, `.form.txt:923-934` |
| `ComandoEvidenciaDeAcuerdo` | botón | `OnClick` → `ComandoEvidenciaDeAcuerdo_Click` | Selecciona un fichero y lo registra como evidencia UTE | `:15-74`, `.form.txt:990-999` |
| `ComandoEvidenciaDeAcuerdoQuitar` | botón | `OnClick` → `ComandoEvidenciaDeAcuerdoQuitar_Click` | Elimina el anexo registrado, con confirmación | `:76-123`, `.form.txt:1053-1064` |
| `lbl1` | etiqueta | — | Rótulo del bloque; el diseño lo conserva sin código asociado | `.form.txt:1118` |

`ComandoEvidenciaDeAcuerdo` llega al diseño con `Enabled = NotDefault` (`.form.txt:990`) y `ImagenAnexo` y el botón de quitar arrancan ocultos: los muestra u oculta `EstablecerDatos` (`:178-185`).

## Flujos

### Apertura (`Form_Open`, `:125-163`)

1. Reloj de arena y `m_Error = ""` (`:129-132`).
2. `AjustarTamaño Me` (`:133`; `Funciones Generales.bas:1043`).
3. Si `m_ObjEdicionActiva Is Nothing`, escribe «No hay ninguna edición activa» y lanza el error controlado (`:136-138`).
4. `EstablecerDatos m_Error` y, si vuelve con error, lo relanza (`:144-146`).
5. Quita el reloj (`:149-151`).

### Normal: registrar la evidencia (`ComandoEvidenciaDeAcuerdo_Click`, `:15-74`)

1. Reloj de arena y `m_Error = ""` (`:20-23`).
2. Si `m_ObjEdicionActiva.Proyecto.EnUTE = "No"`, aborta con «No es un proyecto en UTE» (`:24-26`).
3. Abre el selector de fichero: `Seleccionar(True, "Seleccione el archivo a anexar", m_Error)` (`:29`; `Funciones Generales.bas:1101`).
4. Si el usuario cancela, quita el reloj y sale sin crear nada (`:33-37`).
5. Crea el anexo y lo rellena: `IDEdicion` (`:41`), `EvidenciaUTE = "Sí"` (`:42`), `EvidenciaSuministrador = "No"` (`:43`) y el título (`:44`).
6. `Registrar m_ObjEdicionActiva, m_URL, m_Error` (`:45`; `Anexo.cls:347`).
7. Recarga el anexo con `Constructor.getAnexo` (`:50`; `Constructor.bas:2617`) y lo enlaza a la edición con `Set m_ObjEdicionActiva.AnexoEvidenciaUTE = m_Anexo` (`:54`, `Edicion.cls:1127`).
8. Refresca la pantalla (`:55`) y quita el reloj (`:59-61`).

### Alternativo: quitar la evidencia (`ComandoEvidenciaDeAcuerdoQuitar_Click`, `:76-123`)

1. Si `m_ObjAnexoEvicenciaUTE Is Nothing`, aborta con «No parece tener registrado un anexo de evidencia de acuerdo con la UTE» (`:80-82`).
2. Pide confirmación con `vbYesNo` y **botón por defecto el no** (`:85-86`); si no es `vbYes`, sale sin borrar (`:87-88`).
3. `EliminarAnexo` (`:94`; `Anexo.cls:483`), recarga la edición con `Constructor.getEdicion` (`:98`; `Constructor.bas:24`) y vuelve a enlazar el anexo (`:102`).
4. Devuelve el foco al botón de adjuntar (`:106`) y refresca (`:107`).

### Alternativo: abrir el documento (`ImagenAnexo_Click`, `:192-225`)

1. Si `Not fso.FileExists(m_ObjAnexoEvicenciaUTE.URLAnexo)`, avisa con «El archivo no es alcanzable en estos momentos» y la ruta (`:200-202`); si `FicheroAbierto(URL)`, avisa «Ya tiene el fichero abierto» (`:204-206`; `Funciones Generales.bas:796`).
2. `AbrirEnLocal m_ObjAnexoEvicenciaUTE.URLAnexo, Me.hWnd, m_Error` (`:208`; `Funciones Generales.bas:854`) y, si vuelve con error, lanza el error (`:209-210`).

### De error

Los cuatro manejadores con `On Error` siguen el patrón común: quitar el reloj y, si `Err.Number <> 1000`, componer el mensaje, escribirlo al administrador con `CorreoAlAdministrador` (`Funciones Generales.bas:964`) y avisar con diálogo crítico; si es 1000, solo advertencia (`:63-72`, `:114-121`, `:153-160`, `:216-223`). Ninguno cierra el formulario. El único manejador del cierre (`ComandoEvidenciaDeAcuerdo_Click`, `:72`) llama además `AvanceCerrar`.

## Reglas, validaciones y cálculos

- **Solo en proyectos UTE.** Sin `Proyecto.EnUTE` no se puede adjuntar nada (`:24-26`), pero **no se oculta el botón**: el error sale al pulsarlo.
- **El permiso es global.** `EstablecerDatos` habilita el botón de adjuntar según `blnPermitidoEscribir` (`Variables Globales.bas:121`) (`:172-177`), y muestra el botón de quitar solo si además hay anexo (`:183-185`).
- **Cada edición tiene un único anexo de este tipo.** El objeto global `m_ObjAnexoEvicenciaUTE` (`Variables Globales.bas:122`) es el único puntero, y `Edicion.AnexoEvidenciaUTE` (`Edicion.cls:1103`) es una propiedad única: adjuntar un segundo fichero **no** comprueba antes si ya había uno.
- **El tipo del anexo se fija al crear** (`EvidenciaUTE = "Sí"` y `EvidenciaSuministrador = "No"`, `:42-43`), lo que lo separa del anexo de evidencias de suministradores, y **no hay acceso directo a datos**: sin `getdb()`, `Recordset`, `Open`, `Execute`, SQL, `DLookup`, `Nz` ni transacciones, todo pasa por `Anexo` y `Constructor`.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Edición y proyecto sobre los que se opera | `m_ObjEdicionActiva` (`Variables Globales.bas:79`) | `:24`, `:41`, `:45`, `:98`, `:137` |
| Anexo de evidencia UTE | `m_ObjAnexoEvicenciaUTE` (`Variables Globales.bas:122`), `Edicion.AnexoEvidenciaUTE` (`Edicion.cls:1103`, `Edicion.cls:1127`) | `:54`, `:81`, `:94`, `:178`, `:200-208` |
| Ruta del fichero | `m_ObjAnexoEvicenciaUTE.URLAnexo` | `:200-208` |
| Permiso de escritura | `blnPermitidoEscribir` (`Variables Globales.bas:121`) | `:172`, `:183` |
| Alta del anexo | `Anexo.Registrar` (`Anexo.cls:347`) | `:45` |
| Baja del anexo | `Anexo.EliminarAnexo` (`Anexo.cls:483`) | `:94` |
| Recarga de anexo y edición | `Constructor.getAnexo` (`Constructor.bas:2617`), `Constructor.getEdicion` (`Constructor.bas:24`) | `:50`, `:98` |
| Objeto de sistema de ficheros | `fso` (`Variables Globales.bas:71`) | `:200` |

Escritura efectiva: una alta o una baja de anexo por la clase `Anexo`; la vista no escribe por sí misma.

## Coexistencia

Pendiente de `../coexistencia.md`. Aquí pesa el almacenamiento de ficheros: hoy se guarda una ruta y se abre en local, patrón que en web exige decidir si el anexo vive como recurso con descarga, con antivirus y con control de acceso, en lugar de una ruta que el cliente resuelve.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-109.** `ImagenAnexo_Click` lanza `Err.Raise 100` (`:210`) donde el resto de la pantalla usa `Err.Raise 1000` (`:26`, `:31`, `:47`, `:52`, `:57`, `:83`, `:96`, `:100`, `:139`, `:146`, `:202`, `:206`). El manejador solo distingue el error controlado por su número (`:217`), así que **un fallo al abrir el documento se trata como error inesperado**: compone el mensaje interno, escribe al administrador y muestra un diálogo crítico en vez del aviso controlado.
- **H-GR-110.** `EstablecerDatos` instala `On Error Resume Next` (`:171`) **sin etiqueta de manejador** y no avisa ni registra nada. Es `Private Function` sin tipo de retorno y su resultado se descarta: cualquier fallo al habilitar, mostrar u ocultar controles se traga en silencio y la pantalla queda inconsistente sin que nadie lo sepa.
- **H-GR-111.** El mensaje de fichero no alcanzable concatena la **ruta almacenada en el anexo** al texto que ve el usuario final (`:201`). Exponer una ruta interna en un aviso de usuario es información que no debería salir de la aplicación.
- **H-GR-112.** `ComandoEvidenciaDeAcuerdoQuitar_Click` vacía la global `m_URLInforme` (`:103`), que no aparece en ningún otro punto de esta pantalla. Es un efecto lateral fuera del alcance de la función, sin explicación en el código.
- **H-GR-113.** Tras borrar, el refresco se llama **sin argumento de error**: `EstablecerDatos` a secas (`:107`), al contrario que las otras dos llamadas (`:55`, `:144`). Un fallo de refresco no se detecta y la pantalla puede mostrar el anexo que ya no existe.
- **H-GR-114.** El título del anexo se compone por concatenación con el número de edición: `"ED " & m_ObjEdicionActiva.Edicion & " Acuerdo tratamiento riesgos con miembros de UTE"` (`:44`). Es un nombre derivado en el momento del alta; en web conviene que el anexo tenga nombre propio editable y metadatos estables.
- **H-GR-115.** El alta ocurre antes de la recarga: se registra el anexo (`:45`) y solo después se pide a `Constructor.getAnexo` (`:50-52`) y se enlaza a la edición (`:54`). Si cualquiera de esos dos pasos falla, el anexo **queda escrito en la base de datos** sin registro en la edición y sin que la pantalla lo muestre.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: subida del fichero con almacenamiento real y antivirus en lugar de una ruta local; un único recurso de evidencia UTE por edición, con suPUT y su borrado atómicos; descarga autenticada en vez de `AbrirEnLocal`; mensaje de error sin exponer rutas; refresco con manejo de error uniforme.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir la pantalla con una edición activa muestra la evidencia UTE si la hay y el botón de adjuntar si el usuario puede escribir.
2. En un proyecto que no está en UTE no se puede adjuntar, y se avisa de por qué.
3. Adjuntar un fichero lo registra y lo muestra sin recargar la pantalla.
4. Quitar la evidencia pide confirmación, la elimina y la pantalla deja de mostrarla.
5. Un documento no alcanzable avisa sin exponer la ruta interna.
6. Un fallo al registrar no deja el anexo a medias.

## Pendientes

- Qué tablas toca `Anexo.Registrar` (`Anexo.cls:347`) y `EliminarAnexo` (`Anexo.cls:483`), y si el borrado es físico o lógico (**H-GR-115**).
- Qué contenedores navegan a esta pestaña y bajo qué condición (la referencia solo está en el diseño), y qué hace `AbrirEnLocal` (`Funciones Generales.bas:854`) con rutas de red o locales (**H-GR-111**).
- Si existe control de concurrencia sobre `Edicion.AnexoEvidenciaUTE` cuando dos usuarios adjuntan a la vez (**H-GR-114**).