# Form_FormSolicitudAltaTraspaso

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormSolicitudAltaTraspaso.cls` (142 líneas) y `.form.txt` (25 controles `Begin`) · **PRD:** [README](../README.md)

Contenedor de la solicitud de alta y de traspaso: es un **casco** con dos pestañas de navegación que alojan subformularios —datos de la solicitud y fechas del trámite— y no contiene lógica de negocio propia; toda la persistencia ocurre en el subformulario de datos. Es un formulario **no enlazado** cuyo parámetro de entrada es `OpenArgs`, y cuyo estado de trabajo vive en las globales de sesión. Las citas son `Form_FormSolicitudAltaTraspaso.cls:línea` salvo indicación; `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`, `classes/` = `src/classes/`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormSolicitudesGestion` | `DoCmd.OpenForm "FormSolicitudAltaTraspaso", OpenArgs:=` «Alta» o «Traspaso» según la fila seleccionada | `Form_FormSolicitudesGestion.cls:524-537` |
| `FormSolicitudAlta` (selector de tipo) | Abre el mismo formulario; el lanzador queda abierto detrás | `Form_FormSolicitudAlta.cls:60, 73` |
| `FormTareasTramitadorPendientes` | Mismo patrón con `OpenArgs` | `Form_FormTareasTramitadorPendientes.cls:178-187` |
| `FormAltaDatosSolicitante` | Abre el formulario **sin** `OpenArgs` | `Form_FormAltaDatosSolicitante.cls:372-375` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `FormWeb` | `InformeDelRegistro` genera un HTML con el registro de la solicitud y abre ese visor | `:32`; `FU.bas:1721-1745` |
| Explorador del sistema | `AbrirAyuda` abre el PDF de ayuda del formulario activo | `:13`; `FU.bas:475-503` |
| Padre (`FormSolicitudesGestion`) | `Form_FormSolicitudesGestion.Filtrar` al releer datos, invocado desde el subformulario | `Form_FormSolicitudesAltaTraspasoDatos.cls:468-470` |

**Entrada por `OpenArgs`:** `Form_Load` solo admite «Alta» o «Traspaso»; cualquier otro valor aborta con el aviso «El formulario se ha abierto con parámetros insuficientes» (`:108-112`). Estado compartido: `m_ObjSolicitudActiva`, `m_ObjSolicitudAlInicio`, `m_ObjUsuarioConectado`, `m_ObjEntorno` y `m_TituloFormulario` (`VG.bas:188-204`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` → `Form_Load` | Ajusta, valida `OpenArgs`, monta las pestañas y elige el subformulario de fechas | `:94-138`, `.form.txt:225` |
| `ControlDeNavegación0` | control de navegación | — | Aloja `NavDatos` y `NavFechas`; su subformulario asociado es `FrmDetalle` | `.form.txt:677, 693, 705, 757` |
| `NavDatos` | pestaña | — | Destino por defecto: el subformulario de datos | `.form.txt:720`, este cls `:117` |
| `NavFechas` | pestaña | — | Destino según el tipo: fechas de alta o de traspaso; se deshabilita mientras no exista solicitud guardada | `.form.txt:774`, este cls `:119-123, 55, 67, 80` |
| `FrmDetalle` | subformulario | — | Fuente del contenido de la pestaña activa | `.form.txt:862`, este cls `:118` |
| `lblTitulo` | etiqueta | — | Encabezado contextual: alta, tipo y titular | `:54, 73, 81`, `.form.txt:636` |
| `lblModificacion` | etiqueta | — | Quién y cuándo hizo el último cambio | `:72`, `.form.txt:1022` |
| `ComandoVerLogs` | botón | `OnClick` → `ComandoVerLogs_Click` | Abre el registro de la solicitud activa en `FormWeb` | `:27-44`, `.form.txt:985` |
| `ComandoAyuda` | botón | `OnClick` → `ComandoAyuda_Click` | Abre el PDF de ayuda | `:9-25`, `.form.txt:561` |
| `cmdSalir` | botón | `OnClick` → `cmdSalir_Click` | Cierra sin guardar | `:140-142`, `.form.txt:900` |

## Flujos

### Apertura (`Form_Load`, `:94-138`)

1. `Ajustar Me` y `Me.Caption = m_TituloFormulario` (`:104, 107`); el rótulo de diseño queda sustituido por «SOLICITUDES HPS …» con versión y usuario (`VG.bas:388-400`).
2. Lee `OpenArgs` en `m_Tipo` y exige que sea «Alta» o «Traspaso» (`:108-112`).
3. `EstablecerDatos m_Error` (`:113-115`), cuyo resultado decide título, visibilidad de la etiqueta de modificación y disponibilidad de la pestaña de fechas.
4. Fija `NavDatos.NavigationTargetName` y `FrmDetalle.SourceObject` al subformulario de datos, y `NavFechas.NavigationTargetName` al de alta o al de traspaso (`:117-123`).

### Armado del encabezado (`EstablecerDatos`, `:46-91`)

1. **Sin solicitud en sesión:** crea una instancia nueva de `solicitud`, le asigna el tipo recibido, oculta la etiqueta de modificación, deshabilita la pestaña de fechas y rotula «SOLICITUD DE <TIPO>» (`:76-81`). Es la vía del alta nueva.
2. **Con solicitud ya persistida** (`IDSolicitud <> ""`): la relee con `constructor.getSolicitud` y activa la pestaña de fechas (`:58-67`; `constructor.bas:483-495`). Si la relectura devuelve nada, aborta con «El formulario se ha abierto con parámetros suficientes» invertidos: el aviso es que los parámetros eran insuficientes (`:62-65`).
3. **Alta en curso** (`IDSolicitud = ""` dentro de una solicitud existente): deja el título «ALTA DE SOLICITUD DE HPS» y deshabilita la pestaña de fechas (`:51-56`), pero la ejecución cae después en las líneas comunes y sobrescribe ambos ajustes (`:71-73`).
4. En los casos de solicitud existente, rotula «SOLICITUD DE <TIPO> PARA <NOMBRE COMPLETO>» usando el tipo recibido por `OpenArgs`, no el tipo almacenado, y muestra la fecha de última modificación (`:71-73`; `Solicitud.cls:244-278`).

### Alternativos

- Consultar la trazabilidad: `ComandoVerLogs_Click` genera el HTML del registro y abre `FormWeb`; si la sesión no tiene solicitud, falla antes de generar nada (`:27-44`).
- Volver atrás sin guardar: `cmdSalir_Click` cierra sin confirmar cambios pendientes (`:140-142`).

### Errores

`Form_Load` es el único manejador que **cierra el formulario** tras avisar (`:137`); los handlers de botón dejan el formulario abierto. Los errores con número distinto de 1000 se notifican al administrador (`:34-37, 52-57, 116-120`) y `pregunta` recoge el resultado de cada `MsgBox` (`VG.bas:182`).

## Reglas, validaciones y cálculos

- **Contrato de apertura:** el tipo viaja en `OpenArgs` y se valida por comparación literal (`:108-112`); es el único dato de control de la pantalla.
- **Disponibilidad de la pestaña de fechas:** depende de que la solicitud tenga identificador, no de que tenga fechas (`:53-55, 62-67, 80`).
- **Título:** compone el tipo recibido y el nombre completo del titular en mayúsculas (`:73, 81`); `NombreCompleto` se calcula en `Solicitud.cls` (`Solicitud.cls:70-71`).
- **Sin validación propia:** este contenedor no comprueba obligatorios ni campos de negocio; toda la validación vive en `SolicitudOperaciones.MotivoNoOK` al invocar «Guardar» desde el subformulario (`SolicitudOperaciones.cls:343-446`).
- **Sin control de permisos:** se llega desde tres pantallas distintas sin que este contenedor filtre por rol ni por modo de oficina (a verificar).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Relectura de la solicitud al abrir | `TbSolicitudes` | DAO, `SELECT *` por identificador | `:58`; `constructor.bas:483-495` |
| Registro de la solicitud (trazabilidad) | fichero HTML generado y `FormWeb` | sistema de ficheros + formulario | `:32`; `FU.bas:1721-1745` |
| Ayuda | PDF en el directorio de documentación | sistema de ficheros | `:13`; `FU.bas:492-499` |
| Escrituras propias | ninguna | — | — |

## Coexistencia

Pendiente de `../coexistencia.md`. Puntos previsibles: el transporte del parámetro por `OpenArgs` no tiene equivalente en una aplicación web, donde el tipo debería formar parte de la ruta; el contrato literal «Alta»/«Traspaso» es una cadena mágica replicada en tres apertores (`Form_FormSolicitudesGestion.cls:531, 536`, `Form_FormTareasTramitadorPendientes.cls:181, 186`); la composición del título con mayúsculas es lógica de presentación dentro del contenedor y no debería sobrevivir a una vista. La lógica está **duplicada** en el `.cls` y en el bloque `CodeBehindForm` del `.form.txt` (`.form.txt:1045-1190`), y este último contiene además las mismas dos referencias `Err.Numbr` que el `.cls`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **Caída de flujo en la rama de alta:** tras dejar el título «ALTA DE SOLICITUD DE HPS» y ocultar la etiqueta de modificación (`:53-54`), las líneas comunes la vuelven a mostrar con la fecha de última modificación y sobrescriben el título con el tipo recibido (`:71-73`). En una alta en curso el encabezado no dice «ALTA» sino «SOLICITUD DE <TIPO> PARA» con el titular vacío.
- **El tipo del título no es el tipo de la solicitud:** se compone con `m_Tipo` (el `OpenArgs`) y no con `m_ObjSolicitudActiva.TIPO` (`:73, 81`); si el usuario cambia el tipo desde el diálogo de cambio, el encabezado puede quedar desalineado hasta la siguiente reapertura.
- **Estado global heredado entre aperturas:** la rama de solicitud nueva (`:76-81`) no pone `m_ObjSolicitudAlInicio` a nada, a diferencia del lanzador (`Form_FormSolicitudAlta.cls:59`). Si quedó una solicitud anterior en esa global, el subformulario toma la rama de actualización en lugar de la de inserción (`SolicitudOperaciones.cls:73-77, 151-154`).
- **Apertura sin `OpenArgs` rota:** `FormAltaDatosSolicitante` abre este formulario sin parámetro (`Form_FormAltaDatosSolicitante.cls:375`), por lo que `Form_Load` aborta con «parámetros insuficientes», avisa y cierra (`:109-112, 137`). La cadena de alta del solicitante quedaría interrumpida siempre.
- **Mensajes de error que no corresponden al handler:** `ComandoVerLogs_Click` notifica el fallo como si fuera del handler de ayuda (`:40`, y el mismo literal en `:21`).
- **Referencia a un miembro inexistente:** los dos manejadores de botón componen el mensaje con `Err.Numbr` en lugar del número de error (`:21, 40`). Tal como está, el módulo no compila; hay seis apariciones en el proyecto (este par, más `Form_FormSolicitudesAltaTraspasoDatos.cls:224`, `Form_FormSolicitudRenovacionDatos.cls:289`, `Form_FormSolicitudRenovacion.cls:19`, `Form_FormAltaDatosSolicitante.cls:383`). A verificar si el binario difiere del export.
- **Línea inválida:** un apóstrofo suelto tras el cierre de `EstablecerDatos` (`:93`) impide compilar el módulo por el mismo motivo.
- **Eventos públicos muertos:** declara `Alta` y `Editado` (`:4-5`) y no los emite nunca; ningún suscriptor puede recibir nada.
- **`cmdSalir` descarta sin avisar** (`:140-142`): los cambios no guardados del subformulario se pierden en silencio.
- **`ComandoVerLogs` con la sesión vacía:** si `m_ObjSolicitudActiva` es nada, la lectura del identificador falla y el aviso no explica que no hay solicitud seleccionada (`:32`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: tipo y variante de trámite como parámetro de ruta en lugar de `OpenArgs`; un único formulario de solicitud con pestañas en lugar de un contenedor que reparte entre dos subformularios; encabezado derivado del modelo y no de literales; evento de guardado único y con significado (`guardada` en vez de `Alta`/`Editado`); redirección tras guardar en lugar de abrir un visor HTML generado en local.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir con tipo «Alta» monta la pestaña de datos y deja la de fechas deshabilitada; con una solicitud ya guardada, la de fechas queda disponible.
2. Abrir sin un tipo válido avisa y no monta nada.
3. El encabezado nombra el tipo de la solicitud y su titular.
4. «Ver logs» abre el registro de la solicitud activa; sin solicitud activa explica la causa.
5. «Salir» cierra sin guardar.

## Pendientes

- Por qué `FormAltaDatosSolicitante` abre sin `OpenArgs` (`Form_FormAltaDatosSolicitante.cls:375`): si la intención era abrir el editor con la solicitud recién creada.
- Si el subformulario de datos escribe alguna vez en `m_ObjSolicitudAlInicio` al terminar un alta (`Form_FormSolicitudesAltaTraspasoDatos.cls:817-819`).
- Qué estado hereda `Form_FormSolicitudesGestion` al fijar `m_ObjSolicitudActiva` (`Form_FormSolicitudesGestion.cls:524`) y si lo limpia al cerrar.
- Contenido real de `constructor.getSolicitud` (`constructor.bas:483-495`) y su coste con `TbSolicitudes` de tamaño real.
- Sincronía entre el `.cls` y la copia del bloque `CodeBehindForm` del `.form.txt` (`.form.txt:1045`).