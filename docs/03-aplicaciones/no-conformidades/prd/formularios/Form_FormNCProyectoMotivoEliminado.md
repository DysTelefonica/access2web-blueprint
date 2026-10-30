# Form_FormNCProyectoMotivoEliminado

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCProyectoMotivoEliminado.cls` (223 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña de la ficha de una NC de proyecto que muestra y permite modificar el motivo con el que se marcó como borrada. Es un formulario **no enlazado** (sin `RecordSource` ni `ControlSource` en el `.form.txt`): lee de la NC activa en memoria (`m_ObjNCProyectoActiva`) y de la copia al abrir (`Form_FormNCProyecto.m_NCAlInicio`). Las citas son `Form_FormNCProyectoMotivoEliminado.cls:línea` salvo indicación; `NCProyectoOperaciones.cls` abrevia `classes/NCProyectoOperaciones.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Ficha de NC de proyecto (`FormNCProyecto`) | Destino de navegación del control de pestañas, por nombre de formulario | `Form_FormNCProyecto.form.txt:1304` |

No abre otros formularios. Depende de variables públicas de `FormNCProyecto` (`m_NCAlInicio`, `m_PermitidoEditar`; `Form_FormNCProyecto.cls:8, 12`) y del global `m_ObjNCProyectoActiva` (`Variables Globales.bas:117`).

## Controles y eventos

`.form.txt` declara 13 bloques `Begin`: el formulario (`.form.txt:5`), 8 bloques de valores por defecto (`:723-860`), la sección `Detalle` (`:881`) y los 3 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `MotivoBorrado` | cuadro de texto, `Tag="DATO"`, `Locked` en diseño | `AfterUpdate` | Muestra el motivo; al cambiar recalcula el color del borde si se puede editar | `:9-16`, `:70-72`, `:117-124`, `.form.txt:893-909` |
| `Etiqueta50` («Motivo») | etiqueta | — | Rótulo | `.form.txt:925-937` |
| `ComandoGrabar` («Grabar») | botón, deshabilitado en diseño | `OnClick` | Guarda el motivo | `:171-224`, `.form.txt:956-968` |

## Flujos

### Carga (`Form_Load`, `:9-35`)

1. Reloj de arena y `EstablecerDatos` (`:12-16`).
2. `EstablecerDatos` exige `m_ObjNCProyectoActiva` con `IDNoConformidad` y `m_NCAlInicio`; si no, avisa «El formulario se ha abierto con parámetros insuficientes» (`:44-56`).
3. Si `m_PermitidoEditar = No`, deshabilita «Grabar» (`:58-60`); si no, lo habilita solo cuando la NC al inicio está borrada y el usuario no es técnico (`:62-66`).
4. Vuelca el motivo guardado si no está vacío (`:68-75`) y pinta el borde (`:76`).

### Grabar (`:171-224`)

Comprueba `EsTecnico` («No tiene autorización para esa acción», `:182-185`), que haya cambios respecto a `m_NCAlInicio.MotivoBorrado` («No hay cambios que guardar», `:186-193`), copia los controles con `Tag` «DATO» a la NC activa (`:194, 147-170`) y llama a `ModificarMotivoBorrado` (`:198-206`). Si va bien, actualiza `m_NCAlInicio.MotivoBorrado` y recarga (`:208-209`). No hay mensaje de éxito.

### Errores

Error 1000 ⇒ aviso con `m_Error`; otro ⇒ `CorreoAlAdministrador` y mensaje crítico (`:26-34, 215-223`).

## Reglas, validaciones y cálculos

- **Rol:** un técnico no puede grabar (`:182-185`); `ModificarMotivoBorrado` no repite la comprobación (`NCProyectoOperaciones.cls:1457-1495`).
- **Motivo obligatorio y NC borrada:** `ModificarMotivoBorrado` rechaza NC sin identificador, motivo vacío («Se ha de indicar el motivo») o NC no borrada («No aparece como borrado») (`NCProyectoOperaciones.cls:1457-1480`).
- **Color de borde (solo lectura visual):** con la NC borrada y motivo vacío, o con la NC no borrada y motivo informado, el borde cambia al color de «no relleno» (`:98-110`). Constantes `COLOR_BORDE_CAMPO_*` (origen no verificado, ver Pendientes).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Evidencia |
|---|---|---|---|
| Actualiza `MotivoBorrado` | `TbNoConformidades` (`WHERE IDNoConformidad`) | SQL dinámico concatenado, `Execute`, sin transacción | `NCProyectoOperaciones.cls:1480-1484` |
| Registro en el historial «Registrar Motivo Borrado» | tabla de log de NC de proyecto vía `RegistrarLogProyecto` | DAO (clase `LogNCProyecto.Alta`) | `NCProyectoOperaciones.cls:1485`, `Funciones Generales.bas:993-1020` |

El formulario no lee tablas directamente.

## Coexistencia

La actualización concatena el motivo entre comillas simples sin escapar (`NCProyectoOperaciones.cls:1480-1481`): en la versión web debe parametrizarse. Sin transacción entre `UPDATE` y registro de log. Resto, pendiente de `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `MotivoBorrado` está `Locked` en diseño (`.form.txt:894`) y el código no lo desbloquea (no hay `Locked` en el `.cls`); «Grabar» podría quedar habilitado sin poder editar el motivo.
- Un motivo con comilla simple rompe el `UPDATE` (`NCProyectoOperaciones.cls:1480`).
- `ComandoGrabar_Click` lanza `Err.Raise 100` (no 1000) si falla el volcado de datos (`:195-197`), por lo que mostraría el mensaje genérico con aviso al administrador.
- Si falta `IDNoConformidad`, `EstablecerDatos` escribe el mensaje en `m_Error` y no en `p_Error` (`:49-51`); como `m_Error` es la variable de módulo que `Form_Load` pasa por referencia (`:16`), el aviso sale igualmente, pero el patrón difiere del resto de ramas (`:45, 54`).
- No se invoca el refresco de la lista de gestión tras grabar.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: integrar el motivo en el diálogo de borrado de la NC.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con una NC borrada y un usuario no técnico con edición permitida, se puede guardar un motivo no vacío y distinto del guardado.
2. Con la NC no borrada, guardar devuelve «No aparece como borrado».
3. Un técnico recibe «No tiene autorización para esa acción».
4. Sin cambios, devuelve «No hay cambios que guardar».
5. Un motivo con comillas se guarda sin error.

## Pendientes

- Definición de `COLOR_BORDE_CAMPO_RELLENO`, `COLOR_BORDE_CAMPO_NORELLENO` y anchos (usadas en `:89-107`): no localizadas en esta ficha.
- Cómo `FormNCProyecto` carga este formulario en la navegación (`Form_FormNCProyecto.form.txt:1304`): se ficha en `FormNCProyecto`.
- Si el campo `MotivoBorrado` es editable en ejecución con `Locked` en diseño.
