# Form_FormResponsablesGestion

> **Commit fijado:** `cfa1937a` · **Fuente:** `src/forms/Form_FormResponsablesGestion.cls` (392 líneas) y `.form.txt` (1218 líneas) · **PRD:** [README](../README.md)

Listado de responsables con alta, edición y baja. Es la pestaña «Responsables» del mantenimiento de la aplicación. Recibe filtros por nombre y abre `Form_FormResponsable` para crear o editar; tras el alta o la edición refresca su lista capturando los eventos `Alta` y `Editado` del formulario de detalle. Las citas son `Form_FormResponsablesGestion.cls:línea` salvo que se indique el `.form.txt`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Menú principal | `DoCmd.OpenForm "FormResponsablesGestion"` | global `m_TituloFormulario` (`Variables Globales.bas:196, 388-401`) |
| `Form_FormResponsable` (alta) | `RaiseEvent Alta` ⇒ `m_Form_Alta` ⇒ `Filtrar` | `:5-6`, `:380-382`, `Form_FormResponsable.cls:5, 86` |
| `Form_FormResponsable` (edición) | `RaiseEvent Editado` ⇒ `m_Form_Editado` ⇒ `Filtrar` | `:5-6`, `:385-387`, `Form_FormResponsable.cls:6, 87` |

Estado compartido: la global `m_ObjResposableActivo` (`Variables Globales.bas:204`), escrita desde `ComandoAlta_Click` (`:129-131`, `Nothing`) y desde `ComandoEditar_Click` (`:160-162`, objeto en edición). El detalle la lee en `EstablecerDatos` (`Form_FormResponsable.cls:147-150`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | Rótulo «RESPONSABLES» | `.form.txt:208-209` |
| `Nombre` | texto | `BeforeUpdate` ⇒ `Filtrar` | Filtra por coincidencia parcial | `:392`, `.form.txt:610-611` |
| `ListaFiltrados` | lista 4 columnas (ID oculto), valor *value list* | `OnClick` y `OnDblClick` | Muestra ID, Nombre, Correo, Descripción; doble clic abre editar | `:182-217`, `:323-330`, `.form.txt:412, 422-428` |
| `lblListaDeResultados` | etiqueta | — | Recuento «Registros filtrados (n)»; heredado del diseño con valor inicial «(1)» | `:60, 72`, `.form.txt:458-459` |
| `ComandoAlta` | botón | `OnClick` | Abre `FormResponsable` con `m_ObjResposableActivo = Nothing` | `:111-141`, `.form.txt:478-482` |
| `ComandoEditar` | botón | `OnClick` | Abre `FormResponsable` con la selección actual (deshabilitado sin selección) | `:144-176`, `.form.txt:536-540` |
| `ComandoEliminar` | botón | `OnClick` | Pide confirmación y llama a `ResponsableOperaciones.Eliminar` | `:179-217`, `.form.txt:641-645` |
| `ComandoActualizar` | botón | `OnClick` | Relanza `Filtrar` | `:91-107`, `.form.txt:320-322` |
| `ComandoAyuda` | botón | `OnClick` | Llama a `AbrirAyuda` | `:62-87`, `.form.txt:227-231` |
| `cmdSalir` | botón | `OnClick` | Cierra el formulario (`acSaveNo`) | `:13-16`, `.form.txt:724-726` |
| `lblEstado` | etiqueta | — | Rótulo heredado «URLDirectorioAplicaciones», no se usa | `.form.txt:703-704` |

## Flujos

### Apertura (`Form_Open`, `:251-289`)

1. `Ajustar Me` y asigna `Me.Caption = m_TituloFormulario` (`:255-257`).
2. `Me.Caption = Me.lblTitulo.Caption` (`:260`) sobreescribe con «RESPONSABLES» (señalado en defectos).
3. Llama a `Filtrar` (`:262`), pinta el contador (`:60, 72`) y cierra el progreso.

### Filtrado (`Filtrar`, `:20-78`)

1. Deselecciona todas las filas y reinicia la lista con cabecera `IDResponsable;Nombre;Correo;Descripción` (`:25-46`).
2. `ComandoEditar.Enabled = False` y `ComandoEliminar.Enabled = False` (`:47-48`).
3. `constructor.getResponsables(p_Nombre:=Nz(Me.Nombre, ""))` carga el diccionario (`constructor.bas:2572-2623`); si devuelve `Nothing`, sale sin pintar (`:50-56`).
4. Añade filas `<IDResponsable>;<Nombre>;<Correo>;<Descripcion>` (`:57-65`); actualiza el contador (`:71`); si hay un solo resultado, lo preselecciona y dispara `ListaFiltrados_Click` (`:72-76`).

### Selección (`ListaFiltrados_Click`, `:333-387`)

1. Si la columna 0 está vacía (cabecera), `Set m_ResponsableSeleccionado = Nothing` (`:350-358`).
2. Si no hay `m_ResponsableSeleccionado` o el ID cambió, recarga con `constructor.getResponsable(p_IDResponsable:=…)` (`constructor.bas:3431-3474`) (`:359-373`).
3. Habilita `ComandoEditar` y `ComandoEliminar` (`:382-383`).
4. Doble clic delega en `ComandoEditar_Click` si está activo (`:323-330`).

### Alta, edición, baja

- **Alta** (`ComandoAlta_Click`, `:111-141`): cierra un `FormResponsable` previo, deja `m_ObjResposableActivo = Nothing` y abre el detalle. El alta del detalle dispara `m_Form_Alta` ⇒ `Filtrar` (`:380-382`).
- **Edición** (`ComandoEditar_Click`, `:144-176`): cierra detalle previo, asigna `m_ObjResposableActivo = m_ResponsableSeleccionado` y abre `FormResponsable`. La edición del detalle dispara `m_Form_Editado` ⇒ `Filtrar` (`:385-387`).
- **Baja** (`ComandoEliminar_Click`, `:179-217`): si no hay selección, aborta (`:192-196`); pide confirmación Sí/No (`:199-200`); instancia `ResponsableOperaciones.Eliminar` (`:202-211`); al terminar, relanza `ComandoActualizar_Click` y un `Filtrar` adicional (`:212-217`).

### Errores

Cada handler acumula en `m_Error`, envía `CorreoAlAdministrador` (si procede) y muestra `MsgBox vbCritical`; `Form_Open` cierra el formulario tras el aviso (`:280-289`). `Filtrar` sólo devuelve el texto en `p_Error` (`:73-76`).

## Reglas, validaciones y cálculos

- **Filtro por nombre**: `LIKE '*<texto>*'` sobre `Nombre` (concatenado en `constructor.bas:2597-2599`; construcción del filtro a verificar). Sin `ORDER BY`: el orden lo decide el motor.
- **Eliminar**: `ResponsableOperaciones.ValidacionEliminar` (`ResponsableOperaciones.cls:124-160`) impide borrar un responsable cuyo correo esté asignado a alguna solicitud; si lo está, devuelve «Ese correo de Responsable ya está registrado» con la lista de `IDSolicitud` separadas por salto de línea.
- **Permisos**: este formulario no comprueba rol; se llega desde el menú de gestión. La ficha no tiene `EsTecnico`/`EsAdministrador` en su lógica, al contrario que el detalle (`Form_FormResponsable.cls:212-218`).
- **Global reutilizada**: `m_ObjResposableActivo` está escrita desde la gestión y leída desde el detalle; su errata ortográfica («Resposable» en lugar de «Responsable») está consolidada en `Variables Globales.bas:204` y a lo largo de este formulario: `:131, 160`.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Evidencia |
|---|---|---|---|
| Lista por nombre | `TbResponsables` | SQL dinámico concatenado por `Nombre` | `constructor.bas:2572-2623` |
| Responsable por ID | `TbResponsables` | SQL dinámico concatenado | `constructor.bas:3431-3474` |
| Eliminar responsable | `TbResponsables` | `DELETE *` concatenado | `ResponsableOperaciones.cls:167-176` |
| Insertar/actualizar responsable | `TbResponsables` | `AddNew/Edit` DAO; sin `INSERT/UPDATE` parametrizado | `ResponsableOperaciones.cls:73-110` |
| Validación de borrado | `TbSolicitudes` (vía `getSolicitudesConResponsable`) | SQL dinámico | `constructor.bas:2626-2677` |

Tras `Eliminar`, `ResponsableOperaciones.Registrar` limpia `m_ObjEntorno.Responsables = Nothing` (`ResponsableOperaciones.cls:114`) para forzar recarga del caché; lo mismo al editar con éxito.

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se observa: la concatenación `WHERE`/`DELETE` con IDs numéricos (`ResponsableOperaciones.cls:81, 171`) y la concatenación de `LIKE '*texto*'` (`constructor.bas:2597-2599`) son patrones Jet/ACE; en PostgreSQL exigirán parametrización e `ILIKE '%texto%'`. La validación se hace por correo (`ValidacionEliminar`), lo que obliga a modelar unicidad en la capa de datos.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Me.Caption = m_TituloFormulario` es reemplazado a continuación por `Me.Caption = Me.lblTitulo.Caption` (`:257, 260`); el título resuelto por `Variables Globales.bas:388-401` con versión y nombre del usuario queda invisible al usuario.
- `lblTitulo.Caption = "RESPONSABLES"` pero la propiedad `Caption` del formulario quedó como «GESTIÓN DE LAS SOLICITUDES» (`.form.txt:38`): rótulo heredado del diseño de la pestaña que reutiliza el formulario.
- `ListaFiltrados` se crea con `RowSource = "IDSolicitud;DNI;Nombre;Estado"` (`.form.txt:424`) y `ColumnCount = 4` (`.form.txt:412`), aunque `Filtrar` reescribe con «IDResponsable;Nombre;Correo;Descripción» (`:46, 57-62`): el binario lleva un esquema de columnas que no se usa nunca.
- `ComandoEliminar` habilita `Me.AllowEdits = True` antes de mostrar la confirmación (`:188-189, 197`), contrariamente al resto de handlers; tras la confirmación se reabre y cierra (`:194-196`).
- `lblEstado` mantiene el rótulo heredado «URLDirectorioAplicaciones» (`.form.txt:703-704`).
- `Set m_ObjResposableActivo = Nothing` desde `ComandoAlta_Click` (`:131`) cierra cualquier edición previa sin avisar; `ComandoEditar_Click` no cierra formularios abiertos que no sean `FormResponsable` (puede haber otros con la misma selección).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: cambiar `DoCmd.OpenForm` por navegación declarativa; mover el alta/baja a un modal de mantenedor compartido con `FormSolicitud…`; «caché de responsables» se vuelve estado de servidor invalidado por evento y no global Access.

## Criterios de aceptación

Propuestos, no aprobados:

1. Escribir parte del nombre lista los responsables cuyo nombre lo contiene.
2. La lista muestra siempre cuatro columnas: ID oculto, Nombre, Correo y Descripción.
3. Sin selección, los botones Editar y Eliminar están deshabilitados.
4. Pulsar Alta abre el detalle con rótulo «ALTA DE RESPONSABLE» y al guardar, la lista se refresca y la fila nueva queda seleccionada.
5. Pulsar Editar abre el detalle con rótulo «EDICIÓN DE RESPONSABLE» y, tras guardar, la lista se refresca.
6. Eliminar pide confirmación; si el correo está referenciado por alguna solicitud, no se elimina y muestra los `IDSolicitud` afectados.
7. Cancelar la confirmación no toca la base.

## Pendientes

- Sentido real del filtro por nombre en `Constructor.bas:2597-2599` (LIKE exacto, ¿case-sensitive?).
- Si la validación de borrado aplica a renovaciones además de altas/traspasos: lo verifica `SolicitudOperaciones` y debería estar descrito en su ficha.
- Cuándo se invoca `Entorno.Responsables = Nothing` por otras vías de mutación de responsables.
- Si `lblEstado` se usa para mostrar el directorio activo, cuál era la intención original del control.
