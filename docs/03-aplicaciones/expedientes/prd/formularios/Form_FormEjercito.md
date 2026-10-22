# Form_FormEjercito

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormEjercito.cls` (199 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta y edición de un registro del catálogo de ejércitos (`TbEjercitos`): nombre y descripción. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); el modo se decide por la global `m_ObjEjercitoActivo` (`Variables Globales.bas:253`): vacía, alta; informada, edición (`:130-135`). Las citas son `Form_FormEjercito.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEjercitosGestion` (Alta) | Vacía la global, cierra el formulario si estaba abierto y lo abre; escucha `Alta`/`Editado` | `Form_FormEjercitosGestion.cls:73-81` |
| `Form_FormEjercitosGestion` (Editar) | Asigna a la global el ejército seleccionado y lo abre | `Form_FormEjercitosGestion.cls:222-230` |
| `Form_FormExpedienteGeneral` (`ComandoNuevoEjercito`) | Vacía la global y lo abre en alta; escucha solo `Alta` | `Form_FormExpedienteGeneral.cls:103-117, 631` |

Eventos hacia quien lo abrió: `Alta` (`:5`, lanzado en `:79`) y `Editado` (`:6`, lanzado en `:81`). Tras guardar, cierra el formulario (`:86`); `cmdSalir` cierra sin guardar (`:10-13`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `EJERCITO` | texto | — | Nombre del ejército (obligatorio al guardar) | `.form.txt:619`; `:66, 149` |
| `DESCRIPCION` | texto | — | Descripción opcional | `.form.txt:654`; `:67, 152` |
| `ComandoRegistrar` (rótulo «Editar») | botón, deshabilitado en diseño | `OnClick` | Guarda el registro | `.form.txt:831-844`; `:42-96` |
| `ComandoAyuda` (rótulo «Salir», tip «Ayuda») | botón | `OnClick` | `AbrirAyuda` | `.form.txt:513-517`; `:15-40` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | `.form.txt:735-737`; `:10-13` |
| `lblTitulo` | etiqueta | — | «ALTA DE EJÉRCITO» o «EDICIÓN DE …» | `:131, 135, 145` |

## Flujos

### Apertura (`Form_Open`, `:166-199`)

1. `Ajustar Me` redimensiona la ventana (`:174`; `FUNCIONES UTILES.bas:2703`). Parte con `ComandoRegistrar` deshabilitado y `AllowEdits = False`; solo con `EsAdministrador = Sí` los habilita (`:175-181`).
2. `EstablecerDatos`: sin ejército activo, título de alta y fin (`:130-134`). Con él, relee el registro con `constructor.getEjercito` por `IDEjercito` y lo guarda en `m_EjercitoAlInicio` (`:137`); si no existe, error «No se ha podido encontrar el Ejercito registrado» (`:141-144`). Copia nombre y descripción no vacíos a los controles (`:147-154`).
3. Cualquier error: correo al administrador, `MsgBox` crítico y cierre (`:191-198`).

### Guardado (`ComandoRegistrar_Click`, `:42-96`)

1. `HaHabidoCambios`: en alta siempre hay cambios; en edición compara nombre y descripción con `m_EjercitoAlInicio` (`:98-116`). Sin cambios, «No hay cambios que guardar» (`:58-61`).
2. Crea el objeto si la global está vacía y le copia los dos controles con `Nz` (`:62-68`).
3. `EjercitoOperaciones.Registrar` con el estado inicial (`:70-77`; detalle en Reglas).
4. Lanza `Alta` o `Editado` según hubiera estado inicial y cierra (`:78-86`). En alta, la global queda con el `IDEjercito` asignado por `Registrar` (`EjercitoOperaciones.cls:52`), que `Form_FormExpedienteGeneral` usa para seleccionar el nuevo valor (`Form_FormExpedienteGeneral.cls:644-645`).

### Errores

Errores distintos del interno 1000 se envían con `CorreoAlAdministrador` y se muestran con `MsgBox` crítico (`:88-95`); el formulario queda abierto con los datos tecleados.

## Reglas, validaciones y cálculos

- **Permiso:** solo administradores pueden editar y guardar (`:178-181`). No hay otro control de rol.
- **Nombre obligatorio:** vacío ⇒ «Falta el campo Ejercito» (`EjercitoOperaciones.cls:29-32`).
- **Nombre único:** en alta, o si el nombre cambia, busca por nombre con `getEjercito` y rechaza con «Ese nombre de Ejercito ya está registrado» (`EjercitoOperaciones.cls:33-49`; `constructor.bas:446-448`).
- **Identificador:** en alta, `IDEjercitoCalculado` = máximo de `TbEjercitos.IDEjercito` + 1 (`EjercitoOperaciones.cls:52`; `Ejercito.cls:30`; `FUNCIONES UTILES.bas:2851-2864`).
- **Descripción:** en alta solo se escribe si no está vacía; en edición, vacía se guarda como `Null` (`EjercitoOperaciones.cls:65-76`).
- **Caché y resincronización:** tras escribir invalida `m_ObjEntorno.Ejercitos` (`EjercitoOperaciones.cls:82`); si en una edición cambió el nombre, `ResyncCatalogo "ejercito"` relee cada expediente de `TbExpedientes` con ese `IDEjercito` y ejecuta `SincronizarDerivados` con ámbito `Todo`: cadena PECAL y caché `RegistrarExpEntidades` (`EjercitoOperaciones.cls:86-93`; `ExpedienteOperaciones.cls:698-742, 776-786`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Relectura del registro | `TbEjercitos` por `IDEjercito` | SQL dinámico concatenado | no | `constructor.bas:442-444` |
| Unicidad del nombre | `TbEjercitos` por `Ejercito` | SQL dinámico, nombre entre comillas | no | `constructor.bas:446-448` |
| Nuevo identificador | `Max(IDEjercito)` de `TbEjercitos` | SQL dinámico | no | `FUNCIONES UTILES.bas:2851-2853` |
| Alta / edición | `TbEjercitos` (`AddNew`/`Edit` + `Update`) | DAO | no | `EjercitoOperaciones.cls:60-78` |
| Resincronización tras renombrar | `TbExpedientes` por `IDEjercito`, luego derivados de cada expediente | SQL dinámico + DAO | no | `ExpedienteOperaciones.cls:712-742` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: identificador por `Max + 1` en cliente en lugar de autonumérico o secuencia; búsqueda de unicidad con `=` sobre texto, que en Jet no distingue mayúsculas y en PostgreSQL sí; alta, edición y resincronización sin transacción común.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Un error de `DameID` queda en `Ejercito.Error` y no se consulta: el alta seguiría con `IDEjercito` vacío (`Ejercito.cls:30-41`; `EjercitoOperaciones.cls:52`).
- `Max + 1` sin bloqueo: dos altas simultáneas obtendrían el mismo identificador (`FUNCIONES UTILES.bas:2851-2864`).
- Un apóstrofo en el nombre rompe la consulta de unicidad (`constructor.bas:448`).
- Si la resincronización falla, el renombrado ya está escrito: no hay transacción (`EjercitoOperaciones.cls:70-93`).
- En edición, el formulario modifica el mismo objeto que tiene seleccionado `Form_FormEjercitosGestion` antes de validar (`:65-68`; `Form_FormEjercitosGestion.cls:222`): un guardado rechazado deja ese objeto alterado.
- Título inconsistente: «EDICIÓN DE EJÉRCITO» se sobrescribe con «EDICIÓN DE Ejercito» (`:135, 145`). Restos de diseño: `Caption` «Edición de Suministrador» (`.form.txt:35`, sustituido en `:177`) y `ComandoAyuda` con rótulo «Salir» (`.form.txt:514`).
- `Form_Open` llama a `CorreoAlAdministrador` sin `On Error Resume Next`, a diferencia del resto de handlers (`:194`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: identificador generado por la base; unicidad del nombre como restricción con criterio de mayúsculas decidido; renombrado y resincronización en una única transacción.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario no administrador ve el registro en solo lectura y no puede guardar.
2. Guardar sin nombre muestra «Falta el campo Ejercito» y no escribe nada.
3. Guardar un nombre ya existente se rechaza; guardar el mismo registro sin renombrar no se rechaza.
4. Guardar sin cambios muestra «No hay cambios que guardar».
5. Vaciar la descripción en edición la deja nula.
6. Renombrar un ejército actualiza los expedientes que lo referencian.
7. Crear un ejército desde la ficha de expediente lo deja seleccionado en su combo.

## Pendientes

- Detalle de `ActualizarCadenaPECAL` y `RegistrarExpEntidades` dentro de la resincronización (`ExpedienteOperaciones.cls:777, 783`).
- Si `AllowEdits = False` bloquea la escritura en controles no enlazados (`:176`).
- Contenido de `AbrirAyuda` y `CorreoAlAdministrador`.
