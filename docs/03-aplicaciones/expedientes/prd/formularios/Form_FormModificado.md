# Form_FormModificado

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormModificado.cls` (212 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ventana emergente de **edición** de un modificado ya registrado de un expediente (título de diseño «EDICIÓN DE MODIFICADO/HITO», `.form.txt:38, 225`). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): lee el modificado de la base al abrir y lo guarda con `ExpedienteModificadoOperaciones.Registrar`. No da de alta modificados: el alta se hace en la propia pestaña [Form_FormExpedienteModificados](Form_FormExpedienteModificados.md) (`Form_FormExpedienteModificados.cls:61-110`). Las citas son `Form_FormModificado.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormExpedienteModificados` (`ComandoEditar`) | Exige `EsAdministrador = Sí`, toma el modificado seleccionado en `ListaModificados`, lo deja en la global `m_ObjModificadoActivo`, cierra `FormModificado` si estaba abierto y lo abre sin `OpenArgs`; escucha su evento `Editado` con `WithEvents` | `Form_FormExpedienteModificados.cls:5, 22-47` |

Estado compartido: `m_ObjModificadoActivo` (`Variables Globales.bas:268`), `EsAdministrador` y `m_TituloFormulario` (`Variables Globales.bas:273`). El formulario no recibe el modificado por parámetro.

Salidas: `CmdSalir_Click` cierra sin guardar (`:10-13`); `ComandoRegistrar_Click` lanza `Editado` y cierra tras guardar (`:79-83`). Al recibir `Editado`, la pestaña invalida `Expediente.Modificados`, recarga `ColModificados` y la lista (`Form_FormExpedienteModificados.cls:221-234`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `DESCRIPCION` (rótulo «DESCRIP.») | texto multilínea | — | Descripción del modificado; obligatoria al guardar | `.form.txt:351, 381`; `:156, 64` |
| `NModificado` (rótulo «NMod.») | texto, `Format = General Number` | — | Número de modificado; debe ser numérico al guardar | `.form.txt:519-520, 551`; `:159, 65` |
| `FechaFirmaModificado` (rótulo «F.Firma») | texto, `Format = Short Date` | — | Fecha de firma; opcional | `.form.txt:442-443, 473`; `:162, 66` |
| `FechaFinModificado` (rótulo «F.Fin») | texto, `Format = Short Date` | — | Fecha de fin; opcional | `.form.txt:406-407, 493`; `:165, 67` |
| `ComandoRegistrar` (rótulo «Editar») | botón | `OnClick` → `ComandoRegistrar_Click` | Comprueba cambios y guarda | `.form.txt:583-585`; `:42-93` |
| `ComandoAyuda` (rótulo «Salir» en `.form.txt`) | botón | `OnClick` → `ComandoAyuda_Click` | Abre la ayuda con `AbrirAyuda` | `.form.txt:244-246`; `:15-40` |
| `cmdSalir` | botón | `OnClick` → `CmdSalir_Click` | Cierra sin guardar | `.form.txt:686-688`; `:10-13` |
| `lblTitulo` | etiqueta | — | Su rótulo pasa a ser el título de la ventana | `:140` |

## Flujos

### Apertura (`Form_Open`, `:178-212`)

1. `Ajustar Me` ajusta el tamaño (`:186`; `FUNCIONES UTILES.bas:2703`).
2. `ComandoRegistrar` deshabilitado y `AllowEdits = False`; ambos se habilitan solo con `EsAdministrador = Sí` (`:187-193`).
3. `EstablecerDatos` (`:131-176`): sin `m_ObjModificadoActivo`, error «El formulario se ha abierto con parámetros insuficientes» (`:135-139`). Relee el modificado de la base por `IDExpedienteModificado` y lo guarda como `m_ModificadoAlInicio` (`:142-147`); si no existe, error «No se ha podido encontrar el Modificado registrado» (`:148-151`). Copia a los controles los cuatro campos no vacíos (`:154-167`).
4. `AvanceCerrar` (`:199`; `FUNCIONES UTILES.bas:2770`).

### Guardado (`ComandoRegistrar_Click`, `:42-93`)

1. `HaHabidoCambios` compara los cuatro controles (con `Nz(…, "")`) con `m_ModificadoAlInicio`; sin cambios aborta con «No hay cambios que guardar» (`:54-61, 95-129`).
2. Copia los cuatro controles a `m_ObjModificadoActivo` (`:63-69`).
3. `ExpedienteModificadoOperaciones.Registrar` valida y escribe (`:71-78`; reglas abajo).
4. Éxito: `RaiseEvent Editado` y cierre (`:79-83`).

### Errores

Los handlers acumulan el error en `m_Error`, envían `CorreoAlAdministrador` salvo el error interno 1000 y muestran `MsgBox` crítico (`:85-92`, `:32-39`). Si falla la apertura, además cierran el formulario (`:204-211`). Un fallo de validación en el guardado deja el formulario abierto con los datos tecleados (`:85-92`).

## Reglas, validaciones y cálculos

Todas en `ExpedienteModificadoOperaciones.Registrar` (`classes/ExpedienteModificadoOperaciones.cls`):

- **Descripción obligatoria:** vacía ⇒ «La Descripción es obligatoria» (`:26-29`).
- **Número numérico:** `IsNumeric(NModificado)` falso ⇒ «Se ha de indicar el número de modificado» (`:30-33`). No se comprueba que el número sea único en el expediente.
- **Descripción única por expediente:** solo si cambia frente al inicio, busca otro modificado del mismo `IDExpediente` con esa descripción ⇒ «Ya existe un modificado para este expediente con la misma descripción» (`:34-43`; búsqueda en `constructor.bas:5645-5647`).
- **Fecha de firma única por expediente:** solo si `FechaFinModificado` es fecha, la firma inicial era fecha y la firma cambia ⇒ «Ya existe un Modificado para este expediente con la misma fecha de firma» (`:48-60`; búsqueda en `constructor.bas:5640-5643`).
- **Fecha de firma obligatoria:** regla comentada, no se aplica (`:44-47`).
- **Fechas:** se guardan si `IsDate`; si no, `Null` (`:70-81`). No hay regla entre firma y fin.
- **Permiso:** solo `EsAdministrador` (`Form_FormModificado.cls:190-193`; también en el que abre, `Form_FormExpedienteModificados.cls:22-25`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Carga del modificado | `TbExpedientesModificados` por `IDExpedienteModificado` (6 campos de `ColCampos`) | SQL dinámico, DAO | No | `constructor.bas:5635-5638, 5660-5666`; `classes/ExpedienteModificado.cls:88-96` |
| Duplicado por descripción | `TbExpedientesModificados` por `IDExpediente` + `DESCRIPCION` | SQL dinámico concatenado | No | `constructor.bas:5645-5647` |
| Duplicado por fecha de firma | `TbExpedientesModificados` por `IDExpediente` + `FechaFirmaModificado` (`#mm/dd/yyyy#`) | SQL dinámico concatenado | No | `constructor.bas:5640-5643` |
| Guardado | `Edit` + `Update` de `TbExpedientesModificados` (`DESCRIPCION`, `NModificado`, `FechaFirmaModificado`, `FechaFinModificado`) | DAO sobre SQL dinámico | No | `classes/ExpedienteModificadoOperaciones.cls:61-86` |
| Expediente del modificado | `constructor.getExpediente` si no estaba en caché | DAO | No | `classes/ExpedienteModificado.cls:29-42` |
| Último cambio | `UltimoCambioOperaciones.Registrar`; refresca `lblUltimaModificacion` de `FormExpedientesGestion` | DAO (no trazado aquí) | No | `classes/ExpedienteModificadoOperaciones.cls:87`; `FUNCIONES UTILES.bas:189-216` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: literales de fecha Jet `#mm/dd/yyyy#` (`constructor.bas:5643`); descripción entre comillas simples sin escapar (`constructor.bas:5647`); fechas asignadas como texto al campo (`classes/ExpedienteModificadoOperaciones.cls:71, 78`), dependientes de la configuración regional; guardado sin transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La unicidad de la fecha de firma se condiciona a que `FechaFinModificado` sea fecha y a que la firma inicial lo fuera (`classes/ExpedienteModificadoOperaciones.cls:48-49`): informar la firma por primera vez, o hacerlo sin fecha de fin, no se comprueba.
- Descripción concatenada en SQL sin escapar (`constructor.bas:5647`): un apóstrofo rompe la consulta.
- `Registrar` no comprueba `EOF` antes de `Edit` (`classes/ExpedienteModificadoOperaciones.cls:64-66`): si el registro se borró, falla con error de DAO.
- Sin transacción: si `RegistrarUltimoCambio` falla tras el `Update` (`:83-87`), se muestra error pero el modificado ya está guardado y el formulario sigue abierto.
- No hay control de unicidad de `NModificado` ni de orden firma ≤ fin (`classes/ExpedienteModificadoOperaciones.cls:26-60`).
- `HaHabidoCambios` compara el valor del control (fecha o número) con texto (`:104-119`): un formato distinto podría dar falsos cambios.
- La rama «sin modificado inicial» de `HaHabidoCambios` (`:98-101`) es inalcanzable (`EstablecerDatos` aborta, `:148-151`) y `Registrar` desreferencia `p_ModificadoAlInicio` sin comprobarlo (`classes/ExpedienteModificadoOperaciones.cls:34`).
- El permiso de edición ignora el modo solo lectura del expediente contenedor: solo mira `EsAdministrador` (`:190-193`).
- `Me.Caption = m_TituloFormulario` (`:189`) queda sobrescrito por `lblTitulo.Caption` (`:140`).
- `ComandoAyuda` tiene el rótulo «Salir» (`.form.txt:244-246`), igual que `cmdSalir`.
- En `Form_Open`, `CorreoAlAdministrador` no va protegido por `On Error Resume Next` (`:207`), a diferencia de los demás handlers (`:35-36, 88-89`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: un único formulario de alta y edición con las mismas validaciones; unicidad de número, descripción y fecha de firma como restricciones en base de datos; validar fin ≥ firma; guardado y registro de último cambio en una sola transacción.

## Criterios de aceptación

Propuestos, no aprobados:

1. Solo un administrador puede modificar los campos y guardar.
2. Guardar sin cambios muestra «No hay cambios que guardar» y no escribe.
3. Guardar con descripción vacía o número no numérico muestra el mensaje correspondiente y no escribe.
4. Cambiar la descripción a una ya usada por otro modificado del mismo expediente se rechaza.
5. Las fechas vacías se guardan como nulas.
6. Tras guardar, la lista de modificados del expediente muestra los nuevos valores y se actualiza el último cambio.

## Pendientes

- Si Access permite asignar valores por código con `AllowEdits = False` en controles no enlazados (`:188, 154-167`).
- Contenido de `UltimoCambioOperaciones.Registrar` y tablas que escribe (`FUNCIONES UTILES.bas:208`).
- Formato con que `SetPropiedad` guarda las fechas leídas (`constructor.bas:5662`), necesario para cerrar la comparación de `HaHabidoCambios`.
- Si `Form_FormExpedienteModificados` restringe `ComandoEditar` en modo solo lectura del expediente (no visto en `Form_FormExpedienteModificados.cls:9-59`).
