# Form_FormNCProyecto

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCProyecto.cls` (212 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ficha de una NC de proyecto, en alta o en edición: un título, un control de navegación con siete pestañas (General, Acciones, Eficacia, Notas, Documentos, Replanificaciones y Motivo borrado) que cargan en el subformulario `FrmDetalle` el formulario correspondiente, y los botones «Informe» y «Salir». Es un formulario **no enlazado** (sin `RecordSource` visible en el `.form.txt`); la NC se toma de la variable global `m_ObjNCProyectoActiva` y los datos los editan los formularios de cada pestaña. Las citas son `Form_FormNCProyecto.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas` y `VG.bas`, `modules/Variables Globales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_Form0BDOpcionesParteProyectos | Alta: pone `m_ObjNCProyectoActiva = Nothing` y abre el formulario | `Form_Form0BDOpcionesParteProyectos.cls:41-45` |
| Form_FormNCProyectoGestion, Form_FormNCProyectoSeguimientoNC, Form_FormNCProyectoSeguimientoTareas | Edición: abren el formulario con una NC activa ya fijada | `Form_FormNCProyectoGestion.cls:299, 556`; `Form_FormNCProyectoSeguimientoNC.cls:193`; `Form_FormNCProyectoSeguimientoTareas.cls:241` |

Se cierra con «Salir» (`:45-47`, sin guardar). Al cerrar (`Form_Close`, `:50-90`) puede enviar el correo de alta a Calidad. Los eventos públicos `NCNuevo` y `NCEditado` están declarados (`:5-6`) y este fichero no los levanta.

## Controles y eventos

`.form.txt` declara 24 bloques `Begin`: el formulario (`:5`), 7 de valores por defecto (`:742-839`), 3 secciones (`EncabezadoDelFormulario` `:873`, `Detalle` `:915`, `PieDelFormulario` `:1412`) y los 12 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | «ALTA DE NO CONFORMIDAD DE PROYECTO» o «EDICIÓN … <código>[ --> <código asociado>]» | `:151, 177-180`, `.form.txt:894` |
| `ControlDeNavegación0` | control de navegación | — | Contenedor de las pestañas | `.form.txt:935` |
| `NavGeneral` («General») | pestaña | — | Destino `FormNCProyectoGeneral`, fijado por código | `:198-199`, `.form.txt:963` |
| `NavAcciones` («Acciones») | pestaña | — | Destino `FormNCProyectoAcciones` | `.form.txt:1015, 1032` |
| `NavControlEficacia` («Eficacia») | pestaña | — | Destino `FormNCProyectoControlEficacia` | `.form.txt:1070, 1087` |
| `NavNotas` («Notas») | pestaña | — | Destino `FormNCProyectoNota` | `.form.txt:1124, 1141` |
| `NavDocumentos` («Documentos») | pestaña | — | Destino `FormNCProyectoDocumentos` | `.form.txt:1178, 1195` |
| `NavReplanificaciones` («Replanificaciones») | pestaña | — | Destino `FormNCProyectoReplanificaciones` | `.form.txt:1232, 1249` |
| `NavMotivoBorrado` («Motivo borrado») | pestaña | — | Destino `FormNCProyectoMotivoEliminado` | `.form.txt:1287, 1304` |
| `BotónDeNavegación1` («[Agregar nuevo]») | pestaña | — | Marcador del asistente de Access, sin destino | `.form.txt:1341-1342` |
| `FrmDetalle` | subformulario | — | Hospeda la pestaña activa; el código lo fija a `FormNCProyectoGeneral` | `:199`, `.form.txt:1392` |
| `ComandoSalir` | botón (rótulo «Comando47») | `OnClick` | Cierra sin guardar | `:45-47`, `.form.txt:1428-1430` |
| `ComandoInforme` («Informe») | botón | `OnClick` | Genera el informe Word de la NC activa | `:18-43`, `.form.txt:1498-1500` |

## Flujos

### Carga (`Form_Load` → `EstablecerDatos`, `:92-123, 124-212`)

1. `Ajustar Me`, `blnRequiereCorreoAlta = False` y `EstablecerDatos` (`:103-105`).
2. Título de ventana con `TituloUsuarioConectado` (`:135`). Si `EsTecnico = Sí`: `m_PermitidoEditar = No` y `AllowEdits = False`; si no, ambos habilitados (`:136-146`).
3. **Alta** (`m_ObjNCProyectoActiva Is Nothing`): título de alta, pestañas Acciones, Eficacia, Notas, Documentos, Replanificaciones y Motivo borrado deshabilitadas, NC nueva, `blnRequiereCorreoAlta = True` e «Informe» oculto (`:150-164`).
4. **Edición**: «Informe» visible, recarga la NC como `m_NCAlInicio` con `constructor.getNCProyecto` (error «Parece que no se ha podido encontrar la NC de proyecto registrada» si no existe), título con código (y NC asociada si la hay), y si la NC está borrada ⇒ solo lectura (`:167-185`). Si no es técnico, `EstablecerColorPestañasNCProyecto` colorea y habilita pestañas (`:190-195`; `FG.bas:2287`).
5. Fija `NavGeneral` y `FrmDetalle` a `FormNCProyectoGeneral`; en alta da el foco a `PalabraClave` del subformulario (`:198-202`).

### Informe (`ComandoInforme_Click`, `:18-43`)

Llama a `GenerarWordNoConformidades` con la NC activa (`:26`; `FG.bas:3978`). Error 1000 ⇒ advertencia; otro ⇒ `CorreoAlAdministrador` y mensaje crítico (`:36-42`).

### Cierre (`Form_Close`, `:50-90`)

Si el usuario es técnico, no hace nada (`:57-59`). Si la NC ya tiene id y `blnRequiereCorreoAlta`: sin acciones ⇒ `PintarIndicadores` reiniciando y aviso «No se va a mandar el correo a Calidad al no haber acciones registradas» (`:66-70`); con acciones ⇒ `EnviarCorreoAltaNCProyecto` (`:72`; `FG.bas:3344`).

## Reglas, validaciones y cálculos

- Solo se permite editar si el usuario no es técnico y la NC no está borrada (`:136-146, 182-185`). El comentario que bloqueaba por fecha de cierre está desactivado (`:186-189`).
- El correo de alta se envía una sola vez al cerrar, cuando se creó la NC en esa sesión y tiene acciones (`:66-72, 163`).
- El título refleja la NC asociada (`NCProyectoAsociada`) cuando existe (`:176-180`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Recargar la NC en edición | Tablas de NC de proyecto vía `constructor.getNCProyecto` | DAO en `constructor`, fuera de esta ficha | `:168` |
| Escrituras | Ninguna directa; las hacen los formularios de pestaña | — | `:1-212` |
| Correo de alta y Word | Salida a correo y a documento | Funciones de `FG.bas` | `FG.bas:3344, 3978` |

## Coexistencia

Sin SQL propio. Hay que revisar `getNCProyecto` para PostgreSQL vinculado. El control de navegación, `AllowEdits`, `Ajustar` y el cursor de espera son de interfaz Access; el correo y el Word dependen de Outlook y Word de escritorio (a verificar en `FG.bas`).

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual**

- Una sola vista con siete pestañas, de las que solo General está disponible en el alta (`:150-160`).
- El rol técnico lo deja en solo lectura y sin correo de alta (`:136-139, 57-59`).

**Posibles defectos (a verificar)**

- La rama `Else` de `:142-146` comprueba otra vez `EsTecnico <> Sí` y su `AllowEdits = False` es inalcanzable.
- `Form_Close` pone el cursor de espera y, ante error 1000 por falta de acciones, no lo restablece en la ruta de `errores` (`:50-90`).
- `blnRequiereCorreoAlta` depende de que la NC se haya creado en esta sesión: si se guarda y se cierra sin acciones, nunca se envía el correo (`:66-70, 163`).
- El rótulo de `ComandoSalir` es «Comando47» (`.form.txt:1429`).
- Los eventos `NCNuevo` y `NCEditado` no se levantan en este fichero (`:5-6`).

**Propuestas web pendientes de decisión**

- Pestañas con carga bajo demanda y estado de validación por pestaña.
- Envío del correo de alta al guardar con acciones, no al cerrar.

## Criterios de aceptación

1. En alta, solo la pestaña General está disponible y «Informe» está oculto (`:150-164`).
2. En edición, el título muestra el código de la NC y la asociada si existe (`:177-180`).
3. Un técnico no puede editar ni dispara el correo de alta (`:136-139, 57-59`).
4. Una NC borrada se abre en solo lectura (`:182-185`).
5. «Informe» genera el documento de la NC activa (`:26`).
6. Al cerrar una NC nueva con acciones se envía el correo a Calidad; sin acciones, se avisa y no se envía (`:66-72`).
7. «Salir» cierra sin guardar (`:46`).

## Pendientes

- Estructura y reglas de color de `EstablecerColorPestañasNCProyecto` (`FG.bas:2287-2340`), a detallar en la ficha del módulo.
- Origen de `m_ObjNCProyectoActiva` al editar (fijado por `FG.bas:2753`, función no identificada con precisión).
- Contenido de `EnviarCorreoAltaNCProyecto` y `GenerarWordNoConformidades` (dependencias de Outlook/Word).
