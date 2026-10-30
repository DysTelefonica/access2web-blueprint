# Form_frmElegirEtapaAdjunto

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frmElegirEtapaAdjunto.cls` (67 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Diálogo que pide la descripción de un adjunto antes de seleccionar el fichero. Es **no enlazado** (sin `RecordSource`) y no lee ni escribe datos. Pese al nombre y al comentario de cabecera, que habla de preseleccionar la etapa (`:21`), no ofrece ninguna etapa: las propiedades `EtapaSeleccionada` y `EtapaNombreSeleccionado` devuelven siempre el literal «Otro» (`:6-12`). Las citas son `Form_frmElegirEtapaAdjunto.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_frmGestionAdjuntos | «Añadir Adjunto»: `OpenForm` en `acDialog` con el identificador de la solicitud en `OpenArgs` | `Form_frmGestionAdjuntos.cls:86` |

El diálogo no se cierra a sí mismo al aceptar o cancelar: se oculta y deja el resultado en `Tag` (`:48-49, 59-60`). El llamador lee `Tag` y las tres propiedades y lo cierra (`Form_frmGestionAdjuntos.cls:88-98`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` | `Ajustar Me`, fija el título y exige `OpenArgs` distinto de 0 | `:24-40`, `.form.txt:222` |
| `txtDescripcion` | texto | — | Descripción del adjunto; se devuelve recortada | `:14-16`, `.form.txt:698` |
| `cmdAceptar` | botón | `OnClick` | Valida la descripción, pone `Tag = "OK"` y oculta | `:42-55`, `.form.txt:735-737` |
| `cmdCancelar` | botón | `OnClick` | Pone `Tag = "Cancel"` y oculta | `:57-66`, `.form.txt:785-787` |

## Flujos

- **Normal:** el usuario escribe la descripción y pulsa «Aceptar» (`:44-49`); el llamador recibe etapa «Otro», nombre de etapa «Otro» y la descripción (`Form_frmGestionAdjuntos.cls:90-92`).
- **Alternativo:** «Cancelar» pone `Tag = "Cancel"` (`:59`); el llamador cierra el diálogo sin subir nada (`Form_frmGestionAdjuntos.cls:97-98`).
- **Errores:** descripción vacía o solo espacios ⇒ mensaje «Debe indicar una descripcion para el adjunto.» y el diálogo sigue abierto (`:44-46`). Sin `OpenArgs` ⇒ error 513, `ManejadorErroresFormulario` y cierre del formulario (`:32-33, 37-39`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Descripción obligatoria (recortada) | `cmdAceptar_Click` | `:44` |
| Etapa fija «Otro» | propiedades públicas | `:6-12` |
| `OpenArgs` obligatorio, pero no se usa después | `Form_Load` | `:32-33` |

## Lecturas y escrituras

Ninguna. Sin transacciones.

## Coexistencia

Sin acceso a datos; no requiere adaptación para PostgreSQL vinculado. Debe preservarse que la etapa «Otro» y la descripción lleguen al guardado de `tbAdjuntos` (ver Form_frmGestionAdjuntos).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Si el diálogo se cierra sin pasar por «Aceptar» o «Cancelar», o se cierra por error en `Form_Load` (`:39`), el llamador referencia `Forms("frmElegirEtapaAdjunto")` sobre un formulario ya cerrado (`Form_frmGestionAdjuntos.cls:88`); cabría esperar el error de formulario no abierto.
- El comentario anuncia una preselección de etapa que el código no hace (`:21` frente a `:6-12`).
- Existe además `src/forms/frmElegirEtapaAdjunto.form.txt` con otro orden de controles y la etiqueta «Descripcion» sin tilde (`frmElegirEtapaAdjunto.form.txt:632-673`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: integrar la descripción en el propio diálogo de subida.

## Criterios de aceptación

Propuestos, no aprobados:

1. Aceptar con descripción vacía muestra el aviso y no inicia la subida.
2. Aceptar con descripción devuelve etapa «Otro» y la descripción recortada.
3. Cancelar no inicia la subida.

## Pendientes

- Cuál de los dos `.form.txt` corresponde al binario en uso.
- Si la elección de etapa se retiró de forma deliberada o queda pendiente.
