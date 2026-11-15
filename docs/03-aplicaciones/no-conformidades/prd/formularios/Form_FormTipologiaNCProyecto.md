# Form_FormTipologiaNCProyecto

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormTipologiaNCProyecto.cls` (245 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Diálogo de alta y edición de una tipología de NC de proyecto: un único campo de texto («Tipología») y los botones «Registrar» y «Salir». Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); los datos viajan por el objeto global `m_ObjTipologiaProyectoActiva` (`Variables Globales.bas:122`) y se persisten en la tabla `TbTiposNCProyectos` mediante la clase `TipologiaNCProyectos`. Las citas son `Form_FormTipologiaNCProyecto.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`, `VG.bas`, `modules/Variables Globales.bas` y `TipologiaNCProyectos.cls`, `classes/TipologiaNCProyectos.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_FormNCProyectoTipologiaGestion (alta) | Cierra el formulario si estaba abierto, pone `m_ObjTipologiaProyectoActiva = Nothing` y lo abre con `DoCmd.OpenForm`; si el usuario es técnico, avisa y no lo abre | `Form_FormNCProyectoTipologiaGestion.cls:22-37` |
| Form_FormNCProyectoTipologiaGestion (edición) | Asigna la tipología seleccionada a `m_ObjTipologiaProyectoActiva` y lo abre; sin selección avisa «Seleccione un elemento de la lista» | `Form_FormNCProyectoTipologiaGestion.cls:60-78` |

Sin `OpenArgs`. Devuelve el resultado al formulario llamante con los eventos `AltaTipo` y `EdicionTipo` (`:5-6, 144-146, 195-197`), que el llamante escucha con `WithEvents` (`Form_FormNCProyectoTipologiaGestion.cls:6`). Se cierra siempre con `acSaveNo` (`:113, 149, 202`).

## Controles y eventos

`.form.txt` declara 19 bloques `Begin`: el formulario (`.form.txt:4`), 10 de valores por defecto (`.form.txt:132-194`), 3 secciones (`EncabezadoDelFormulario` `.form.txt:201`, `Detalle` `.form.txt:240`, `PieDelFormulario` `.form.txt:301`) y los 5 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo1` | etiqueta | — | Su `Caption` y el del formulario se fijan a «Alta de tipología» o «Edición de tipología» | `:42, 52, 55-56`, `.form.txt:211-234` |
| `Tipologia` | cuadro de texto sin origen de datos | — | Único dato editable; se lee con `Nz(Me.Tipologia, "")` | `:70, 97`, `.form.txt:247-277` |
| `Etiqueta61` («Tipología») | etiqueta | — | Rótulo del campo | `.form.txt:278-300` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` | Valida, guarda y cierra | `:165-214`, `.form.txt:311-322` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra; si hay cambios pregunta si guardarlos | `:108-162`, `.form.txt:358-370` |

## Flujos

### Apertura (`Form_Open`, `:215-244`)

1. Reloj de arena, `Ajustar Me` (`:218-221`; `FG.bas:615`) y `EstablecerDatos` (`:225`).
2. `EstablecerDatos` (`:35-86`): con `m_ObjTipologiaProyectoActiva` informado, título «Edición de tipología» y recarga la tipología de la base por `IDTipo` en `m_TipologiaAlInicio` (`:41-43`); si no existe, «No parece estar registrada la tipología» (`:47-50`). Sin objeto, título «Alta de tipología» y `m_TipologiaAlInicio` queda vacío (`:52`).
3. En edición copia el texto al campo si no está vacío (`:67-72`).
4. Si `EsTecnico = Sí` (`VG.bas:98`), `AllowEdits = False` y «Registrar» deshabilitado; en otro caso ambos habilitados (`:57-64, 73-79`).
5. Cualquier error cierra el formulario (`:234-243`).

### Registrar (`ComandoRegistrar_Click`, `:165-214`)

Rechaza al técnico («No tiene autorización para esa acción», `:173-176`) y la ausencia de cambios («No hay cambios que guardar», `:181-184`); vuelca el campo en el objeto (`:185, 88-98`), llama a `Registrar` (`:190`), lanza `AltaTipo` o `EdicionTipo` según `m_TipologiaAlInicio` (`:194-198`) y cierra.

### Salir (`cmdSalir_Click`, `:108-162`)

Cierra sin preguntar si «Registrar» está deshabilitado, el usuario es técnico o es un alta (`:112-123`). En edición con cambios pregunta «¿Desea guardar los cambios antes de salir» (`:128-129`): «No» cierra sin guardar (`:130-133`); «Sí» registra y lanza `EdicionTipo` (`:134-147`).

### Errores

Error 1000 ⇒ aviso con `m_Error`; otro ⇒ mensaje con número y descripción, `CorreoAlAdministrador` y `MsgBox` crítico (`:153-161, 204-212, 234-242`).

## Reglas, validaciones y cálculos

- **Cambios:** solo se compara `Tipologia` con el valor inicial mediante `<>` (`:22`).
- **Tipología obligatoria:** «No se conoce Tipologia» si está vacía (`TipologiaNCProyectos.cls:55-58`).
- **Duplicados:** en alta, «La tipología está repetida» si existe otra con el mismo texto; en edición, solo si el texto cambia (`TipologiaNCProyectos.cls:59-82`).
- **Identificador:** `IDTipo` = máximo existente + 1 (`TipologiaNCProyectos.cls:152, 27`; `constructor.bas:1165-1178`).
- **Rol:** el técnico solo consulta (`:57-64, 173-176`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Cargar tipología por `IDTipo` | `TbTiposNCProyectos` | SQL dinámico concatenado (`WHERE IDTipo=` + valor) | `constructor.bas:1338-1370` |
| Comprobar duplicado por texto | `TbTiposNCProyectos` | SQL dinámico concatenado (`WHERE Tipologia='` + texto + `'`) | `constructor.bas:1338-1370`, `TipologiaNCProyectos.cls:60, 72` |
| Calcular `IDTipo` | `TbTiposNCProyectos` | SQL dinámico concatenado (`SELECT Max(...)`) | `constructor.bas:1143-1178` |
| Alta | `TbTiposNCProyectos` | DAO `AddNew/Update` | `TipologiaNCProyectos.cls:151, 158, 167-170` |
| Edición | `TbTiposNCProyectos` | SQL dinámico concatenado + DAO `Edit/Update` | `TipologiaNCProyectos.cls:154-155, 176-179` |
| Log de la operación | tabla de log de NC de proyecto vía `RegistrarLogProyecto` | DAO (`LogNCProyecto.Alta`) | `TipologiaNCProyectos.cls:171, 180`; `FG.bas:993-1023` |

Sin transacciones: el registro y su log son operaciones separadas.

## Coexistencia

Pendiente de `../coexistencia.md`. Puntos a adaptar: identificador `Max + 1` sin secuencia (`constructor.bas:1165-1178`), SQL concatenado con el texto sin escapar (`constructor.bas:1338-1370`), `AddNew/Update` y `Edit/Update` de DAO sobre tabla vinculada, y ausencia de transacción entre dato y log.

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual**

- Alta y edición de un único campo; los técnicos lo ven bloqueado.
- El código no normaliza el texto antes de comprobar duplicados (`TipologiaNCProyectos.cls:55-82`).

**Posibles defectos (a verificar)**

- `TipologiaNCProyectos.cls:72` y `constructor.bas:1338-1370` concatenan el texto en la SQL: un apóstrofo en la tipología rompería la consulta.
- En `cmdSalir_Click`, tras `RellenarObjetoConDatosFormulario` con error se lanza el error 100, no el 1000 (`:136`), con lo que se trataría como error inesperado (mensaje crítico y correo al administrador, `:155-158`).
- En `cmdSalir_Click` la rama «Sí» comprueba `m_TipologiaAlInicio Is Nothing` después de `Registrar` (`:143`), pero esa rama solo es alcanzable en edición (`:120-123`): el evento `AltaTipo` nunca se lanza desde ahí.
- `TipologiaNCProyectos.cls:152` calcula `IDTipo` con `Max + 1`: dos altas simultáneas pueden colisionar.

**Propuestas web pendientes de decisión**

- Sustituir el diálogo por un formulario modal con el mismo campo y validación de duplicado sin distinción de mayúsculas.
- Registrar el dato y el log en una sola transacción.

## Criterios de aceptación

1. Al abrir sin tipología activa, el título es «Alta de tipología» y el campo está vacío (`:52, 55-56`).
2. Al abrir con una tipología activa, el título es «Edición de tipología» y el campo muestra el texto guardado (`:42, 70`).
3. Un usuario técnico ve el campo bloqueado y «Registrar» deshabilitado (`:57-64`), y «Salir» cierra sin preguntar (`:112-119`).
4. «Registrar» sin cambios muestra «No hay cambios que guardar» y no cierra (`:181-184`).
5. Una tipología vacía se rechaza con «No se conoce Tipologia» (`TipologiaNCProyectos.cls:55-58`).
6. Una tipología repetida se rechaza con «La tipología está repetida» (`TipologiaNCProyectos.cls:65-67, 77-79`).
7. Un alta correcta crea la fila con el identificador siguiente, anota «Alta de Tipo: …» en el log, lanza `AltaTipo` y cierra (`TipologiaNCProyectos.cls:152, 167-171`; `:194-195`).
8. Una edición correcta actualiza la fila, anota «Edición de Tipo: …», lanza `EdicionTipo` y cierra (`TipologiaNCProyectos.cls:176-180`; `:196-197`).
9. «Salir» con cambios en edición pregunta si guardar; «No» no escribe nada (`:128-133`).

## Pendientes

- Esquema físico de `TbTiposNCProyectos` y de la tabla de log: no está en `src`; a resolver con `../erd-origen.md`.
- `Eliminar` de la clase (`TipologiaNCProyectos.cls:199-226`) no se invoca desde este formulario; lo invoca Form_FormNCProyectoTipologiaGestion (`Form_FormNCProyectoTipologiaGestion.cls:126`).
