# Form_FormSolicitudRenovacion

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormSolicitudRenovacion.cls` (150 líneas) y `.form.txt` (21 controles `Begin`) · **PRD:** [README](../README.md)

Contenedor de la renovación (y de la elevación de grado) de una solicitud de HPS: valida el tipo con el que se abre, muestra el rótulo del usuario y encaja los dos subformularios —datos y fechas— en un control de navegación. Es un **formulario no enlazado** (sin `RecordSource`): no escribe en la base de datos, solo delega. Las citas son `Form_FormSolicitudRenovacion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormSolicitudesGestion` | `DoCmd.OpenForm "FormSolicitudRenovacion", OpenArgs:=.TIPO`, cerrando antes cualquier instancia abierta | `Form_FormSolicitudesGestion.cls:542-546` |
| `FormUsuariosHPSBusqueda` | `DoCmd.OpenForm "FormSolicitudRenovacion", OpenArgs:=m_Tipo` | `Form_FormUsuariosHPSBusqueda.cls:150-155` |
| `FormTareasTramitadorPendientes` | `DoCmd.OpenForm "FormSolicitudRenovacion", OpenArgs:=.TIPO` | `Form_FormTareasTramitadorPendientes.cls:192-196` |

El tipo viaja por `OpenArgs` y se cachea en `m_Tipo` (`:69`). Los tres llamantes son el menú de gestión, la búsqueda de usuario HPS y la lista de tareas, y todos pasan el `TIPO` de la fila seleccionada. No expone eventos a su contenedor.

Estado compartido: `m_ObjSolicitudActiva`, `m_ObjSolicitudAlInicio`, `m_ObjUsuarioHPSActivo` (`VG.bas:197-205`). No recibe parámetros más allá de `OpenArgs`.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `NavDatos` | botón de navegación | — | Engancha `FormSolicitudRenovacionDatos` como origen del control de navegación | `:82`, `.form.txt:427-428` |
| `NavFechas` | botón de navegación | — | Engancha `FormSolicitudRenovacionFechas`; se deshabilita si la solicitud aún no tiene `IDSolicitud` | `:110, 121, 135`, `.form.txt:479-497` |
| `FrmDetalle` | subformulario | — | Aloja el formulario activo de la pestaña | `:83`, `.form.txt:585` |
| `lblTitulo` | etiqueta | — | Encabezado visible; lo compone `EstablecerDatos` | `:109, 137`, `.form.txt:286-287` |
| `lblModificacion` | etiqueta | — | Fecha y autor de la última modificación de la solicitud | `:126, 129`, `.form.txt:707-708` |
| `ComandoVerLogs` | botón | `OnClick` | Genera el informe HTML de la solicitud y lo abre en `FormWeb` | `:25-53`, `.form.txt:736-738` |
| `ComandoAyuda` | botón | `OnClick` | Abre la ayuda (`AbrirAyuda`) | `:8-23`, `.form.txt:313-315` |
| `cmdSalir` | botón | `OnClick` | Cierra el formulario sin guardar | `:149-151`, `.form.txt:623-625` |
| `BotónDeNavegación1` | barra de nuevo registro | — | Barra de navegación estándar; no hay código asociado | `.form.txt:534-535` |

## Flujos

### Apertura (`Form_Load`, `:58-99`)

1. Pasa el reloj de sistema, llama a `Ajustar Me` (`:65`) y deja el formulario en solo lectura (`:67`).
2. Exige que `OpenArgs` sea `Renovación` o `Elevación Grado` (`:70-72`); si no, aborta con «parámetros insuficientes».
3. Exige que `m_ObjUsuarioHPSActivo` esté informado (`:74-76`); sin usuario HPS seleccionado no entra.
4. Carga el contexto con `EstablecerDatos` (`:78`) y, si todo va bien, engancha `FormSolicitudRenovacionDatos` en `FrmDetalle` y en `NavDatos` (`:82-83`).

### Contexto (`EstablecerDatos`, `:101-147`)

1. Si hay `m_ObjSolicitudActiva` con `IDSolicitud` vacío, toma la rama de **alta en frío**: la solicitud aún no está en la base de datos, se deshabilita `NavFechas` (`:110`) y no se relee nada (`:107-111`).
2. Si hay `IDSolicitud`, relee la solicitud con `constructor.getSolicitud` y la guarda en `m_ObjSolicitudAlInicio` como foto de comparación (`:112-118`), y habilita `NavFechas` (`:121`).
3. Si no hay ninguna solicitud activa, **crea una** en memoria con el `TIPO` recibido y el `id` del usuario HPS activo (`:130-134`) y deja `NavFechas` deshabilitado (`:135`).
4. En ambos casos pinta el rótulo de última modificación (`:126-127`) y recompone el título con el nombre completo del usuario HPS en mayúsculas (`:137`).

### Ver logs (`ComandoVerLogs_Click`, `:25-53`)

1. Bloquea la edición (`:31`) y pide a `InformeDelRegistro` el informe de la solicitud activa (`:35`; `FU.bas:1721-1740`), que compone el HTML, lo guarda en disco y abre `FormWeb`.
2. Cierra el indicador de progreso (`:39`), restaura `AllowEdits` (`:53`) y avisa al administrador ante errores no internos (`:45-48`).

### Errores

`Form_Load` es terminal: ante error compone el aviso, lo envía al administrador si no es el código interno 1000, muestra el `MsgBox` **y cierra el formulario** (`:90-98`). `EstablecerDatos`, en cambio, no propaga el error: solo escribe en `p_Error` (`:141-143`), de modo que el cierre depende de que el llamante reenvíe.

## Reglas, validaciones y cálculos

- **Tipo obligatorio:** solo se admiten `Renovación` y `Elevación Grado`; cualquier otro valor aborta la apertura (`:70-72`).
- **Usuario HPS obligatorio:** sin `m_ObjUsuarioHPSActivo` no hay a quién renovar (`:74-76`).
- **La pestaña de fechas depende del identificador:** con `IDSolicitud` vacío no hay registro de fechas que mostrar, así que `NavFechas` se deshabilita en las dos ramas de alta (`:110, 135`) y se habilita solo tras releer la solicitud (`:121`).
- **Título derivado:** el encabezado se compone con el nombre completo del usuario HPS en mayúsculas (`:137`), no con el tipo recibido.
- **Sin cálculo propio:** este formulario no calcula fechas, estados ni importes; toda la lógica vive en `SolicitudOperaciones` y en los subformularios.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Lectura de la solicitud activa | `TbSolicitudes` (vía `constructor.getSolicitud`) | DAO | `Form_FormSolicitudRenovacion.cls:112` |
| Informe HTML de la solicitud | `TbLogGeneral` y anexos (vía `HTMLLog`) | DAO + fichero | `FU.bas:1721-1740` |

Sin escrituras propias: `AllowEdits` solo protege la interfaz.

## Coexistencia

Pendiente del capítulo `coexistencia.md`, que todavía no existe en este PRD. Lo que ya se ve: la ficha depende por completo de una variable global (`m_ObjSolicitudActiva`) y de un parámetro de texto (`OpenArgs`), un patrón sin equivalente directo en una ruta web; el subformulario embebido y el `NavigationControl` se translating a un enrutador o a pestañas de servidor. El informe HTML se genera en disco local y se muestra en otro formulario, lo que en web es una vista o un fichero descargable.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `EstablecerDatos` desactiva el rótulo de modificación en la rama de alta en frío (`:108`) y tres líneas después lo vuelve a activar sin condición (`:126`): en un alta la etiqueta queda visible con el texto de última modificación de una solicitud que todavía no existe.
- El rótulo de «alta de solicitud» que se fija en `:109` se sobrescribe incondicionalmente en `:137` con el título de renovación: nunca llega a verse.
- `ComandoAyuda_Click` concatena `Err.Numbr` (`:19`), que no es una propiedad de `Err`: la propia rama de errores revienta al construir el mensaje y el `MsgBox` final no se muestra.
- El botón de ayuda lleva en diseño el rótulo de «Salir» (`.form.txt:314`), el mismo que el botón real de salida (`.form.txt:624`): dos botones idénticos en pantalla con comportamientos opuestos.
- `cmdSalir_Click` cierra con `acSaveNo` sin preguntar (`:149-151`): los cambios pendientes del subformulario de datos se pierden en silencio.
- `EstablecerDatos` declara tipo de retorno `String` y nunca le asigna valor (`:101, 144`); la señal de error viaja solo por `p_Error`, que el formulario no comprueba en `:78` más allá de la cadena devuelta.
- `Form_Load` deja `AllowEdits = False` (`:67`) y no lo restaura en la rama de error (`:90-98`); el formulario se cierra igualmente, pero el cierre depende de que `DoCmd.Close` llegue a ejecutarse.
- `m_ObjSolicitudAlInicio` se rellena en `:112` pero no se usa en este `.cls`: lo consume el subformulario de datos, que lo vuelve a pedir tras guardar.
- El `Caption` de diseño (`.form.txt:41`) fija versión, nombre de una persona y marca de entorno; es metadato residual que reaparece igual en los formularios de datos (`.form.txt:39`) y de fechas (`.form.txt:38`). Los valores no se reproducen aquí.
- `ComandoAyuda` no comprueba que exista ayuda asociada al formulario antes de invocar `AbrirAyuda` (`:11`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: la solicitud activa como identificador de ruta en lugar de global; pestañas de servidor en lugar de `NavigationControl`; confirmación de salida con aviso de cambios pendientes; informe de auditoría como vista consultable en lugar de HTML generado en disco.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir el formulario sin `OpenArgs` o con un tipo no admitido muestra el aviso de parámetros insuficientes y no abre el formulario.
2. Abrirlo sin usuario HPS seleccionado avisa y no abre el formulario.
3. Con una solicitud existente, la pestaña de fechas está habilitada y el subformulario de datos muestra los datos guardados.
4. Con una solicitud nueva (alta en frío), la pestaña de fechas aparece deshabilitada.
5. El encabezado identifica al usuario HPS con su nombre completo en mayúsculas.
6. «Ver Logs» genera el informe y lo muestra en el visor; si la solicitud no tiene historial, el error se comunica sin cerrar el formulario.
7. «Salir» cierra el formulario y descarta los cambios no guardados.

## Pendientes

- Contenido real de `AbrirAyuda` y si admite formulario como parámetro (`Form_FormSolicitudRenovacion.cls:11`).
- `UltimaModificacionTexto` de `Solicitud`: qué columnas de `TbSolicitudes` compone y de dónde saca el nombre del autor (`Form_FormSolicitudRenovacion.cls:126`).
- Si `Elevación Grado` tiene un tratamiento de datos y de fechas distinto de `Renovación`, más allá del rótulo.
- Si hay algún otro punto de apertura de este formulario que no cierre antes la instancia previa: los tres llamantes conocidos sí lo hacen (`Form_FormSolicitudesGestion.cls:542`, `Form_FormUsuariosHPSBusqueda.cls:150`, `Form_FormTareasTramitadorPendientes.cls:192`), pero el barrido no es exhaustivo.