# Form_Form0BDOpciones

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_Form0BDOpciones.cls` (357 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Menú principal de la aplicación para administradores y usuarios de calidad: da acceso a la búsqueda de expedientes, al alta, al gestor de entidades, a la vista de técnico, a las tareas pendientes, a la ayuda y a la salida. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) y no escribe en la base de datos. Las citas son `Form_Form0BDOpciones.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_frmSplash` | Tras `EVE`, abre `Form0BDOpciones` si `EsTecnico <> Sí`; los técnicos van a `Form0BDOpcionesTecnicos` | `Form_frmSplash.cls:38-50` |
| `Form_frmSplash` (error) | Si falla el arranque, abre `Form0BDOpciones` «de todas formas» | `Form_frmSplash.cls:57-61` |

`EVE` fija el rol: `EsAdministrador` desde `Usuario.EsAdministradorCalculado`; si no es administrador ni de calidad, `EsTecnico = Sí` (`Variables Globales.bas:383-392`); los roles son excluyentes (`Variables Globales.bas:394-406`). También compone `m_TituloFormulario`, con «(Sólo lectura)» para quien no es administrador (`Variables Globales.bas:448-463`).

Salida: `cmdSalir` cierra Access entero (`:26`). El formulario no se cierra al abrir los demás.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ComandoGestionExpedientes` (rótulo «Comando36», imagen) | botón | `OnClick` | Abre `FormExpedientesGestion`, o `FormExpedientesGestionTecnica` si `EsTecnico = Sí`; cierra antes la instancia abierta | `:109-128`, `.form.txt:1311-1313` |
| `lblGestionExpedientes` («Búsqueda de Expedientes») | etiqueta | `OnClick` | Delega en el botón si está habilitado | `:340-344`, `.form.txt:1248-1251` |
| `ComandoAlta` | botón | `OnClick` | Abre `FormExpedienteAltaTipo` con `OpenArgs = "DesdeInicio"` | `:38-48`, `.form.txt:1099-1100` |
| `lblAlta` («Alta AM/Lote/Exp») | etiqueta | `OnClick` | Delega en `ComandoAlta` si está habilitado | `:346-351`, `.form.txt:1280-1283` |
| `ComandoGestorEntidades` | botón | `OnClick` | Cierra y abre `Form0BDGestorEntidades` | `:196-210`, `.form.txt:2722-2723` |
| `lblGestorEntidades` («Entidades») | etiqueta | `OnClick` | Delega en el botón si está habilitado | `:352-357`, `.form.txt:2934-2937` |
| `ComandoBuscadorTecnico` («Abrir como técnico») | botón | `OnClick` | Cierra y abre `Form0BDOpcionesTecnicos` | `:83-98`, `.form.txt:3125-3127` |
| `ImagenTareasON`, `ImagenTareasOFF` | imágenes (campana) | `OnClick` | Abren `FormTareas` (sin cerrar la instancia previa) | `:286-302, 313-329`, `.form.txt:594-595, 658-659` |
| `lblNumeroTareas` | etiqueta | — | Número total de tareas; lo escribe `PintarTareas` | `FUNCIONES UTILES.bas:4032-4045` |
| `ComandoAyuda` (rótulo «Salir», consejo «Ayuda») | botón | `OnClick` | `AbrirAyuda` | `:59-72`, `.form.txt:268-271` |
| `cmdSalir` («Salir») | botón | `OnClick` | Confirma y ejecuta `DoCmd.Quit acQuitSaveAll` | `:8-27`, `.form.txt:3004-3006` |
| `lblVersion` | etiqueta | — | «Versión: » + `m_ObjEntorno.VersionAplicacion` | `:251` |
| `lblEstado` | etiqueta oculta | — | Se oculta en la apertura y tras abrir tareas | `:265, 297, 324` |
| `ComandoGestionE2EEnlaces` | **no existe** en el `.form.txt` | `OnClick` (huérfano) | Abriría `FormE2EGestionEnlaces` | `:151-176` |

## Flujos

### Apertura (`Form_Open`, `:222-283`)

1. `Ajustar Me` ajusta el tamaño de la ventana (`:230`; `FUNCIONES UTILES.bas:2703`).
2. Si `m_ObjEntorno` es `Nothing`, vuelve a ejecutar `EVE` (`:232-237`).
3. `PintarTareas EnumSiNo.Sí` recalcula las tareas y pinta la campana y el contador (`:238-241`; detalle en Reglas).
4. Oculta la cinta salvo con `TempVars("EnDesarrollo") = "Sí"` (`:242-246`); `EVE` la deja siempre en «No» (`Variables Globales.bas:312`).
5. Título = `m_TituloFormulario`; versión en `lblVersion` (`:249-251`).
6. Permisos por rol (`:248-262`; ver Reglas). Cierra la barra de avance y oculta `lblEstado` (`:264-265`).

### Salida (`CmdSalir_Click`, `:8-36`)

1. Pregunta «¿Desea realmente salir de la aplicación» con Sí/No; con No no hace nada (`:13-16`).
2. Con Sí, `DoCmd.Quit acQuitSaveAll` (`:26`).

### Errores

- Los handlers de navegación acumulan el error en `m_Error`, envían `CorreoAlAdministrador` salvo el error interno 1000 y muestran `MsgBox` crítico (p. ej. `:49-56`, `:99-106`).
- `Form_Open`: mismo patrón y, fuera de desarrollo, **cierra la aplicación** con `acQuitSaveAll` (`:271-282`).
- `CmdSalir_Click`: ante cualquier error muestra el mensaje y cierra la aplicación igualmente (`:28-35`); no envía correo.

## Reglas, validaciones y cálculos

- **Permisos:** `ComandoAlta` parte deshabilitado y solo se habilita con `EsAdministrador = Sí` (`:248, 254-255`); `ComandoBuscadorTecnico` solo es visible para administradores (`:257-259`). `ComandoGestionExpedientes` se habilita siempre (`:252`). `ComandoGestorEntidades` se habilita para administradores (`:256`), pero no se deshabilita para el resto (ver Posibles defectos).
- **Gate E2E:** `AplicarRoleGateAdmin` muestra el control solo a administradores y no hace nada si el control no existe (`:181-193`, llamado en `:262`).
- **Destino de búsqueda:** `EsTecnico = Sí` ⇒ `FormExpedientesGestionTecnica`; si no, `FormExpedientesGestion` (`:116-120`).
- **Contador de tareas:** `PintarTareas` vacía las seis colecciones de tareas del entorno cuando se le pide actualizar (`FUNCIONES UTILES.bas:4000-4006`), suma sus tamaños (`FUNCIONES UTILES.bas:4009-4030`) y, con total > 0, muestra `ImagenTareasON` y el número; con 0, `ImagenTareasOFF` y «0» (`FUNCIONES UTILES.bas:4032-4045`). Las colecciones se recargan perezosamente desde `constructor` (p. ej. `classes/Entorno.cls:1251-1260`).
- **Alta desde inicio:** `"DesdeInicio"` hace que `FormExpedienteAltaTipo` arranque sin expediente padre (`Form_FormExpedienteAltaTipo.cls:276-279`).
- **Ayuda:** `AbrirAyuda` siempre abre el mismo PDF general; la ruta por formulario se calcula y se sobrescribe en la línea siguiente (`FUNCIONES UTILES.bas:3094-3095`). Sin fichero, error «No se ha generado la ayuda para este formulario aún» (`FUNCIONES UTILES.bas:3096-3098`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Colecciones de tareas (seis tipos) | vía `constructor.getExpedientesEstadoDesconocido` y análogos | no trazado aquí | `classes/Entorno.cls:1260` |
| Entorno y usuario (si se relanza `EVE`) | configuración local, usuario, propiedades de `Entorno` | no trazado aquí | `Variables Globales.bas:309, 381` |

Sin escrituras propias. `acQuitSaveAll` guarda objetos de diseño al salir, no datos de negocio.

## Coexistencia

Pendiente de `../coexistencia.md`. El menú no accede a tablas directamente; su coste está en `PintarTareas`, que recalcula seis consultas de tareas en cada apertura.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComandoGestorEntidades` no tiene `Enabled` en el `.form.txt` (`.form.txt:2711-2726`) y `Form_Open` solo lo habilita para administradores (`:256`): un usuario de calidad también puede abrir el gestor de entidades, y `Form0BDGestorEntidades` no aplica gate de rol en su apertura (`Form_Form0BDGestorEntidades.cls:115-133`).
- `ComandoGestionE2EEnlaces` no existe en el `.form.txt`: su handler (`:151-176`) y el gate (`:262`) no tienen efecto; el comentario del propio código lo prevé como «acción del humano al compilar» (`:142-145`). Caso análogo al retirado en `Test_ControlPanel.bas:854-857`.
- La rama técnica de `ComandoGestionExpedientes` (`:118-119`) no se alcanza por el arranque normal, que envía a los técnicos a otro menú (`Form_frmSplash.cls:44-50`); solo tras un fallo de arranque (`Form_frmSplash.cls:61`).
- Tras un fallo de `EVE` en el arranque, el menú se abre igualmente (`Form_frmSplash.cls:61`) y, si `m_ObjEntorno` ya existía, no reintenta `EVE` (`:232`): el rol podría quedar sin asignar.
- `ComandoAyuda` rotula «Salir» en el `.form.txt` (`.form.txt:268-269`); `ComandoGestionExpedientes` rotula «Comando36» (`.form.txt:1312`). Ambos se ven como imagen, pero el rótulo es el nombre accesible.
- `CmdSalir_Click`: el mensaje de confirmación no cierra la interrogación (`:13`) y activa y desactiva el reloj de arena sin trabajo entre medias (`:17-25`).
- Restos sin efecto: `Debug.Print Me.lblAlta.ForeColor` (`:263`); `Me.AllowEdits = True` en un formulario no enlazado (`:300, 327`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: menú único con opciones filtradas por rol en servidor; contador de tareas cargado de forma asíncrona; cerrar sesión en lugar de cerrar la aplicación.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un administrador ve y puede usar Búsqueda, Alta, Entidades y «Abrir como técnico».
2. Un usuario no administrador ve el alta deshabilitada y no ve «Abrir como técnico».
3. Un técnico es dirigido a su propio menú y su búsqueda abre la vista técnica.
4. Con tareas pendientes, el icono de tareas muestra el total; sin ellas, muestra 0.
5. Pulsar el icono de tareas abre la lista de tareas.
6. Salir pide confirmación; al rechazarla no ocurre nada.

## Pendientes

- Decisión de negocio sobre si un usuario de calidad debe acceder al gestor de entidades.
- Consultas concretas de las seis colecciones de tareas (`constructor.getExpedientes…`).
- Si `Ajustar` y `AvanceCerrar` (`FUNCIONES UTILES.bas:2703, 2770`) tienen efecto visible en este formulario.
- Si el control `ComandoGestionE2EEnlaces` existe en el binario aunque falte en el `.form.txt` exportado.
