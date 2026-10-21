# Form_FormNCAuditoriaNota

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCAuditoriaNota.cls` (241 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Formulario de una nota de texto libre asociada a una no conformidad (NC) de auditoría: permite guardarla o borrarla. Es **no enlazado** (sin `RecordSource`) y trabaja con la NC activa global `m_ObjNCAuditoriaActiva` (`VG.bas:127`) y la copia inicial `Form_FormNCAuditoria.m_NCAlInicio`. Las citas son `Form_FormNCAuditoriaNota.cls:línea` salvo indicación; `VG.bas` abrevia `modules/Variables Globales.bas` y `Ops`, `classes/NCaUDITORIAOperaciones.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_FormNCAuditoria | Botón de navegación con destino `FormNCAuditoriaNota` | `Form_FormNCAuditoria.form.txt:1143` |

No abre otros formularios. No tiene botón de cierre.

## Controles y eventos

`.form.txt` declara 14 bloques `Begin`: el formulario (`.form.txt:5`), 8 de valores por defecto (`:141-278`), la sección `Detalle` (`:299`) y 4 controles. Evento `OnLoad` (`.form.txt:83`).

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Notas` | cuadro de texto, `Tag = "DATO"` | — | Texto de la nota | `.form.txt:323-325` |
| Etiqueta «Nota» | etiqueta | — | Rótulo | `.form.txt:353` |
| `ComandoGrabar` («Grabar») | botón | `OnClick` | Guarda la nota | `:155-215`; `.form.txt:381-383` |
| `ComandoEliminar` («Eliminar») | botón | `OnClick` | Borra la nota previa confirmación | `:10-61`; `.form.txt:418-420` |

## Flujos

### Carga (`Form_Load` → `EstablecerDatos`, `:63-130`)
Habilita edición, vacía `Notas` y deshabilita ambos botones (`:102-106`); con NC activa habilita «Grabar» (`:107-112`) y, si la NC tiene nota, la muestra y habilita «Eliminar» (`:113-119`). Sin NC activa: «El formulario se ha abierto con parámetros insuficentes» (`:107-110`).

### Grabar (`ComandoGrabar_Click`, `:155-215`)
1. Rechaza a técnicos (`:166-169`).
2. Compara con `m_NCAlInicio.Notas`; sin diferencias, «No hay cambios que guardar» (`:170-177`, `:132-153`).
3. Copia `Notas` a la NC activa con `SetPropiedad` (`:178, 216-241`).
4. `Ops.RegistrarNota` (`:182-191`), aviso «Nota guardada correctamente» (`:195`), actualiza `m_NCAlInicio.Notas` y recarga (`:199-200`).

### Eliminar (`ComandoEliminar_Click`, `:10-61`)
Rechaza a técnicos (`:16-19`); confirma «¿Desea borrar la nota asociada?» con «No» por defecto (`:20-23`); pone `Notas = ""` en la NC activa y llama a `Ops.EliminarNota` (`:27-37`); avisa «Nota borrada correctamente» y vacía `m_NCAlInicio.Notas` (`:41-46`).

### Errores
Error 1000 ⇒ aviso con `m_Error`; otro ⇒ número y descripción, `CorreoAlAdministrador` y `MsgBox` crítico (`:52-60, 80-88, 206-214`).

## Reglas, validaciones y cálculos

- Solo no técnicos graban o borran (`:16-19, 166-169`).
- `RegistrarNota` exige NC registrada y nota no vacía («No hay dato de la nota») (`Ops:732-738`); `EliminarNota` exige NC registrada (`Ops:775-777`).
- Vaciar el cuadro y pulsar «Grabar» no borra la nota: falla con «No hay dato de la nota» (`Ops:736-738`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Evidencia |
|---|---|---|---|
| Guarda la nota | `TbNoConformidadesAuditoria.Notas` por `ID` | SQL dinámico concatenado con el texto (`Execute`) | `Ops:740-743` |
| Borra la nota (`Notas=Null`) | `TbNoConformidadesAuditoria` | SQL (`Execute`) | `Ops:779-783` |
| Log | `RegistrarLogAuditoria "Registrar Notas"` / `"Registrar Eliminar Notas"` | utilidad | `Ops:745, 784` |

Sin transacciones.

## Coexistencia

La nota se concatena sin escapar comillas (`Ops:740-741`): falla con comillas simples y exige parámetros en PostgreSQL. El campo es de texto largo; revisar longitud y tipo en el vínculo. Pendiente de `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- En ambas operaciones el `UPDATE` se ejecuta antes del log: si el log falla, la nota ya está modificada (`Ops:743-745, 783-784`).
- `Ops.EliminarNota` no valida que haya nota que borrar (`Ops:775-784`).
- `ComandoGrabar_Click` ignora una nota que solo difiere en espacios finales o en nulo frente a vacío (`:141`).
- `RellenarDatosObjeto` crea una NC nueva si no hay activa (`:222-224`), aunque la carga ya lo impide.
- El formulario reutiliza `m_Error` de módulo y `Err.Raise 100` para un error de relleno (`:179-181`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: permitir borrar la nota dejando el texto vacío, sin botón separado.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin NC activa informa de parámetros insuficientes (`:107-110`).
2. Con nota existente la muestra y habilita «Eliminar» (`:113-119`).
3. Un técnico no puede grabar ni eliminar (`:16-19, 166-169`).
4. «Grabar» sin cambios avisa y no escribe (`:174-177`).
5. «Eliminar» exige confirmación; con «No» no cambia nada (`:20-23`).
6. Tras grabar o eliminar se avisa y el formulario refleja el estado nuevo (`:195-200, 41-46`).

## Pendientes

- Contenido de `NCAuditoria.SetPropiedad` para la propiedad `Notas`: no revisado.
- Cómo se cierra el formulario (no hay botón; depende de la navegación del formulario padre).
