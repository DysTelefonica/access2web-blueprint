# Form_FormNCProyectoNota

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCProyectoNota.cls` (270 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña de la ficha de una NC de proyecto para ver, guardar y borrar la nota libre de la NC (`TbNoConformidades.Notas`). Es un formulario **no enlazado** (sin `RecordSource` ni `ControlSource`): trabaja sobre la NC activa en memoria (`m_ObjNCProyectoActiva`) y, si esta tiene una NC asociada (vinculada), muestra la nota de la asociada en solo lectura. Las citas son `Form_FormNCProyectoNota.cls:línea` salvo indicación; `NCProyectoOperaciones.cls` abrevia `classes/NCProyectoOperaciones.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Ficha de NC de proyecto (`FormNCProyecto`) | Destino de navegación del control de pestañas | `Form_FormNCProyecto.form.txt:1141` |

No abre otros formularios. Usa `m_PermitidoEditar` de `FormNCProyecto` (`Form_FormNCProyecto.cls:8`) y `m_NCAlInicio` (`:12`).

## Controles y eventos

`.form.txt` declara 14 bloques `Begin`: el formulario (`:5`), 8 bloques de valores por defecto (`:724-861`), la sección `Detalle` (`:882`) y los 4 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Notas` | cuadro de texto, `Tag="DATO"` | — | Texto de la nota; se vacía y rellena en `EstablecerDatos` | `:102, 135-140`, `.form.txt:894-908` |
| `Etiqueta50` («Nota») | etiqueta | — | Rótulo | `.form.txt:924-936` |
| `ComandoGrabar` («Grabar») | botón | `OnClick` | Guarda la nota | `:184-244`, `.form.txt:955-966` |
| `ComandoEliminar` («Eliminar») | botón | `OnClick` | Borra la nota tras confirmar | `:10-61`, `.form.txt:991-1003` |

## Flujos

### Carga (`Form_Load`, `:63-89`; `EstablecerDatos`, `:92-159`)

1. Vacía `Notas`, deshabilita ambos botones y activa `AllowEdits` (`:102-105`). Sin NC activa: «El formulario se ha abierto con parámetros insuficentes» (`:106-109`).
2. Elige la NC de la que leer: la asociada si existe, la activa si no (`:110-114`).
3. Solo si la NC activa **no** tiene asociada: con `m_PermitidoEditar = Sí` y usuario no técnico habilita «Grabar»; técnico o edición no permitida ⇒ `AllowEdits = False` (`:117-129`).
4. Si la NC elegida tiene nota, la muestra; sin asociada, habilita `Notas` a no técnicos y «Eliminar» a no técnicos con edición permitida (`:133-148`).

### Grabar (`:184-244`)

Exige no ser técnico («No tiene autorización para esa acción», `:195-198`) y cambios respecto a `m_NCAlInicio.Notas` («No hay cambios que guardar», `:199-206`). Copia los controles «DATO» a la NC activa (`:207, 245-270`) y llama a `RegistrarNota` (`:211-220`). Muestra «Nota guardada correctamente», actualiza `m_NCAlInicio.Notas` y recarga (`:224-229`).

### Eliminar (`:10-61`)

No técnico (`:16-19`); confirma «¿Desea borrar la nota asociada?» con «No» por defecto (`:20-23`); pone `Notas` a vacío en la NC activa y llama a `EliminarNota` (`:27-37`); muestra «Nota borrada correctamente», vacía `m_NCAlInicio.Notas` y recarga (`:41-46`).

### Errores

Error 1000 ⇒ aviso con `m_Error`; otro ⇒ `CorreoAlAdministrador` y mensaje crítico (`:54-60, 237-243`).

## Reglas, validaciones y cálculos

- **Rol:** técnicos no graban ni eliminan (`:16-19, 195-198`).
- **Nota obligatoria al grabar:** `RegistrarNota` rechaza NC sin identificador y nota vacía («No hay dato de la nota») (`NCProyectoOperaciones.cls:1342-1362`); para vaciar la nota hay que usar «Eliminar».
- **NC vinculada:** la nota mostrada es la de la NC asociada y los botones quedan deshabilitados (`:110-114, 117-132`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Evidencia |
|---|---|---|---|
| Guardar nota | `TbNoConformidades` | SQL dinámico concatenado (`Notas = '…'`), `Execute`, sin transacción | `NCProyectoOperaciones.cls:1363-1367` |
| Borrar nota | `TbNoConformidades` | SQL dinámico (`Notas=Null`), `Execute`, sin transacción | `NCProyectoOperaciones.cls:1403-1408` |
| Historial «Registrar Notas» / «Registrar Eliminar Notas» | tabla de log de NC de proyecto | DAO vía `RegistrarLogProyecto` | `NCProyectoOperaciones.cls:1369, 1409`, `Funciones Generales.bas:993-1020` |

## Coexistencia

La nota se concatena entre comillas simples sin escapar (`NCProyectoOperaciones.cls:1363-1364`): debe parametrizarse. Sin transacción entre actualización y log. Resto, pendiente de `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Una nota con comilla simple rompe el `UPDATE` (`NCProyectoOperaciones.cls:1363`).
- Con NC vinculada, `AllowEdits` permanece en `True` (`:105`) y solo se desactiva en la rama sin asociada (`:123, 126`): el texto podría escribirse aunque no pueda guardarse.
- En `RegistrarNota` el error de `RegistrarLogProyecto` se comprueba (`NCProyectoOperaciones.cls:1369-1372`), pero en `EliminarNota` no (`NCProyectoOperaciones.cls:1409`); un fallo de log al borrar no avisaría.
- `ComandoGrabar_Click` lanza `Err.Raise 100` (no 1000) si falla el volcado (`:208-210`).
- El texto de confirmación dice «nota asociada» aunque se borra la nota de la propia NC (`:20`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: permitir guardar nota vacía en lugar de exigir «Eliminar»; mostrar de forma explícita que la nota proviene de la NC vinculada.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario no técnico con edición permitida guarda una nota no vacía y se confirma «Nota guardada correctamente».
2. Guardar una nota vacía devuelve «No hay dato de la nota».
3. «Eliminar» pide confirmación; con «No» no cambia nada; con «Sí» deja la nota vacía.
4. Un técnico recibe «No tiene autorización para esa acción» y no puede editar.
5. En una NC con asociada se ve la nota de la asociada y no se puede grabar ni eliminar.
6. Cada guardado o borrado deja entrada en el historial de la NC.

## Pendientes

- Cómo `FormNCProyecto` carga la pestaña y calcula `m_PermitidoEditar` (`Form_FormNCProyecto.form.txt:1141`): se ficha en `FormNCProyecto`.
- Efecto real de `AllowEdits = True` frente a `Notas.Enabled` en una NC vinculada (`:105, 139`).
