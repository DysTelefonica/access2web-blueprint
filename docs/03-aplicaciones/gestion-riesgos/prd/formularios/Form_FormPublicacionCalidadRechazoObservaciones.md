# Form_FormPublicacionCalidadRechazoObservaciones

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormPublicacionCalidadRechazoObservaciones.cls` (197 líneas) y `src/forms/FormPublicacionCalidadRechazoObservaciones.form.txt` · **PRD:** [README](../README.md)

Diálogo modal para registrar el motivo por el que calidad rechaza la propuesta de publicación de una edición. Muestra las fechas de propuesta y de rechazo, admite el texto del motivo solo cuando procede, y al grabar emite el evento `Motivado` con la observación escrita. El formulario no escribe en base de datos: la persistencia la hace el padre. Las citas son `Form_FormPublicacionCalidadRechazoObservaciones.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormPublicacionCalidadPublicar` | `ComandoRechazar_Click`: cierra la instancia anterior, abre esta y se engancha al evento `Motivado` con `WithEvents` | `Form_FormPublicacionCalidadPublicar.cls:20, 239-244` |
| `Form_FormPublicacionTecnicoPropuesta` | `DoCmd.OpenForm` directo, sin enganchar eventos | `Form_FormPublicacionTecnicoPropuesta.cls:127-130` |

No acepta parámetros. Toma la edición de los objetos de sesión y, en vez de fiarse de ellos, **rehidrata una copia propia**: `Constructor.getEdicion(m_ObjEdicionActiva.IDEdicion)` (`:124`), que es la que usa para todas las comparaciones posteriores. La salida habitual es cerrarse (`:66`); el padre recibe el motivo por `RaiseEvent Motivado` (`:64`) y es quien escribe en la edición y avisa a los responsables.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Tamaño, título, copia de la edición y llamada a `EstablecerDatos` | `:107-147`, `.form.txt:1389` |
| (formulario) | — | `EstablecerDatos` (procedimiento privado) | Rellena fechas y decide si el motivo y el botón de grabar están habilitados | `:149-194`, `.form.txt:1442-1453` |
| (formulario) | — | `ComandoGrabar_Click` | Valida el motivo, emite `Motivado` y cierra | `:46-80`, `.form.txt:1171` |
| (formulario) | — | `ComandoSalir_Click` | Cierra; avisa si se estaba dejando vacío un rechazo ya existente | `:83-104`, `.form.txt:1206` |
| (formulario) | — | `ComandoAyuda_Click` | Abre la ayuda del formulario | `:18-43`, `.form.txt:924` |
| `PropuestaRechazadaPorCalidadMotivo` | texto | — | Motivo del rechazo; único campo editable | `:171-183`, `.form.txt:1056` |
| `FechaPreparadaParaPublicar` | texto | — | Fecha en que calidad propuso la publicación | `:164-166`, `.form.txt:1100` |
| `PropuestaRechazadaPorCalidadFecha` | texto | — | Fecha del rechazo registrado | `:168-170`, `.form.txt:1131` |
| `ComandoGrabar` | botón | — | Se habilita solo cuando el motivo es editable y el usuario no es técnico | `:173, 178, 182`, `.form.txt:1023` |

## Flujos

### Apertura (`Form_Load`, `:107-147`)

1. `AjustarTamaño Me` y fija el título con el usuario conectado (`:113-117`).
2. Si `m_ObjEdicionActiva Is Nothing`, avisa «El formulario se ha abierto con parámetros insuficientes» y lanza el error (`:119-122`).
3. Rehidrata `m_EdicionAlInicio` con `Constructor.getEdicion` (`:124-127`).
4. Llama a `EstablecerDatos` (`:128-131`).

### Preparación (`EstablecerDatos`, `:149-194`)

1. Si **ni** la fecha de rechazo **ni** la de propuesta son fechas, falla con «La edición no aparece como propuesta para publicación» (`:157-160`).
2. Copia cada fecha solo si es válida, formateada como día/mes/año (`:162-170`).
3. Ramifica según el estado (`:171-183`):
   - **Ya hay motivo registrado:** lo muestra y deja el campo deshabilitado, el botón de grabar deshabilitado, pero el campo **no bloqueado** (`:172-176`).
   - **No hay motivo y el usuario no es técnico:** campo habilitado, editable y botón de grabar habilitado (`:177-180`).
   - **No hay motivo y el usuario es técnico:** campo deshabilitado y botón deshabilitado (`:181-184`).

### Grabar (`ComandoGrabar_Click`, `:46-80`)

1. Si el motivo está vacío, falla con «No se puede quedar vacío el campo del motivo del rechazo» (`:55-58`).
2. Si el motivo es idéntico al de `m_EdicionAlInicio`, falla con «Ningún cambio por grabar» (`:59-62`).
3. `RaiseEvent Motivado(Me.PropuestaRechazadaPorCalidadMotivo)` (`:64`) y cierre sin guardar (`:66`).

### Alternativo (`ComandoSalir_Click`, `:83-104`)

Si el campo ha quedado vacío y en la edición sí había un motivo, muestra «Si desea quitar el rechazo, basta con que publique» (`:87-91`) y cierra igualmente. El aviso es puramente informativo: cancelar **no** propaga ningún evento y el rechazo anterior permanece intacto.

### Error

Los cuatro manejadores comparten patrón: número `1000` distinto de los demás ⇒ mensaje con número y descripción, `CorreoAlAdministrador` y `MsgBox` crítico; número `1000` ⇒ solo el texto que dejó el propio formulario. Los de `Form_Load` y `EstablecerDatos` cierran además la ventana (`:140-147`).

## Reglas, validaciones y cálculos

- **Quien puede rechazar no es quien publica.** La habilitación depende de `EsTecnico <> EnumSiNo.Sí` (`:175`), es decir, solo calidad; un técnico que llega por `FormPublicacionTecnicoPropuesta` abre el mismo diálogo en modo lectura.
- **Un rechazo existente se muestra, no se edita.** El campo se deshabilita pero se deja `Locked = False` (`:174`): el texto queda visible y, aun así, editable por otros medios.
- **El botón depende del texto inicial, no del actual.** `ComandoGrabar.Enabled` se fija una sola vez en la apertura (`:173, 178, 182`) y no vuelve a calcularse; si el motivo se vacía a mano, el botón sigue activo y el error salta en el clic.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Copia de la edición | `Constructor.getEdicion` | `:124` |
| Fechas de propuesta y de rechazo, motivo vigente | Propiedades de `Edicion` | `:157-172` |
| Motivo del usuario | Se propaga al padre por el evento `Motivado` | `:64` |
| Persistencia del motivo | La hace el padre: `Edicion.RechazoPropuestaParaPublicacion` | `Form_FormPublicacionCalidadPublicar.cls:409-413`; `src/classes/Edicion.cls:3296` |

Este formulario no ejecuta ninguna escritura: escribe en base de datos quien recibe el evento.

## Coexistencia

Pendiente de `../coexistencia.md`. El patrón a trasplantar es la rehidratación de la entidad al abrir (`:124`) y la comparación contra el valor inicial para detectar cambios reales (`:59`); el patrón a evitar es el estado repartido entre el hijo, que emite, y el padre, que persiste.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **El campo se declara deshabilitado pero no bloqueado.** En la rama de rechazo existente, `Enabled = False` va acompañado de `Locked = False` (`:173-174`), y lo mismo en las otras dos ramas (`:176, 180-181, 183`): `Locked` se fija a `False` en los tres caminos y nunca a `True`. El control queda protegido solo por el estado de habilitación.
- **«Ningún cambio por grabar» se compara contra la copia de apertura.** La comparación de `:59` usa `m_EdicionAlInicio`, rehidratada en `Form_Load` (`:124`). Si otra ventana modifica el motivo mientras este diálogo está abierto, el usuario sí verá un cambio real y aun así será rechazado como «sin cambios».
- **El aviso de `ComandoSalir_Click` promete algo que este formulario no hace.** El texto afirma que basta con publicar para quitar el rechazo (`:88`), pero cancelar no emite ningún evento y el padre no llega a enterarse de que el campo se vació; el borrado real ocurre en otro camino de la aplicación.
- **La condición de apertura excluye un estado válido.** `EstablecerDatos` solo acepta ediciones cuya fecha de rechazo **o** fecha de propuesta sean válidas (`:157`). Una edición marcada como propuesta mediante un indicador distinto de esas dos fechas se rechaza con un mensaje que no explica cuál de las dos falta.
- **`Form_Load` comprueba el contexto después de usarlo.** El título se fija desde `m_ObjEntorno` en `:117` y la comprobación de `m_ObjEdicionActiva Is Nothing` llega en `:119`; si la edición falta pero el entorno no, el aviso sale correcto, y al revés el error se produce antes de poder avisar.
- **Un técnico ve el mismo diálogo que calidad sin diferencia visible.** La única señal de solo lectura es el estado deshabilitado del campo y del botón (`:181-184`); no hay rótulo ni título que explique por qué no puede escribir.
- **`Err.Raise 1000` como control de flujo compartido.** Los manejadores deciden entre «error técnico a notificar» y «aviso de negocio» comparando `Err.Number` con `1000` (`:36-41, 71-79, 96-103, 139-146, 188-192`); cualquier reasignación del número convierte un aviso de negocio en un aviso al administrador.
- **Dos aperturas con contratos distintos.** `FormPublicacionCalidadPublicar` se engancha al evento `Motivado` y recibe el motivo; `FormPublicacionTecnicoPropuesta` abre la misma ventana sin engancharlo (`:127-130`), de modo que desde ese camino el botón de grabar —deshabilitado, pero presente en el diseño— no tiene consumidor.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: devolver el motivo como resultado de la operación de rechazo en lugar de propagarlo por evento; comparar contra el estado persistido en el servidor en el momento de guardar, no contra una copia tomada al abrir; mostrar una pantalla de solo lectura específica para el técnico en vez de un formulario deshabilitado; y derivar el permiso de rechazo del rol resuelto en el servidor, no de un `If` en la carga del formulario.

## Criterios de aceptación

Propuestos, no aprobados:

1. Solo calidad puede escribir el motivo; un técnico ve las observaciones sin poder modificarlas.
2. Un rechazo ya registrado se muestra íntegro y no se puede sobrescribir desde aquí.
3. No se puede grabar un motivo vacío ni repetir el mismo motivo sin cambios.
4. Al grabar, el rechazo queda registrado, los responsables del proyecto reciben el aviso y la cadena de formularios se cierra.
5. Al cancelar, no se modifica nada y se explica que el rechazo se retira publicando.

## Pendientes

- Qué hace exactamente `RechazoPropuestaParaPublicacion` (`src/classes/Edicion.cls:3296`): si escribe fecha además del motivo y a quién avisa.
- Si el diálogo abierto desde `FormPublicacionTecnicoPropuesta` debería quedar visible para el técnico o ser una consulta de solo lectura.
- Si existe un camino de la aplicación que retire el rechazo sin publicar, dado que el mensaje de `:88` lo sugiere como alternativo.
- El motivo por el que `Locked` se escribe siempre a `False`: si es un valor heredado del diseño o una decisión consciente.