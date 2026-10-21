# Form_FormOrganoContratacion

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormOrganoContratacion.cls` (206 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta y edición de un registro del catálogo de órganos de contratación (`TbOrganosContratacion`): siglas y descripción. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); el modo lo decide la global `m_ObjOrganoContratacionActivo` (`Variables Globales.bas:261`): vacía, alta; informada, edición (`:130-135`). Repite el patrón de [Form_FormEjercito](Form_FormEjercito.md); esta ficha recoge las diferencias. Las citas son `Form_FormOrganoContratacion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormOrganoContratacionGestion` (Alta) | Vacía la global, cierra el formulario si estaba abierto y lo abre; escucha `Alta`/`Editado` | `Form_FormOrganoContratacionGestion.cls:74-82` |
| `Form_FormOrganoContratacionGestion` (Editar) | Asigna a la global el órgano seleccionado y lo abre | `Form_FormOrganoContratacionGestion.cls:225-233` |
| `Form_FormExpedienteGeneral` (`ComandoNuevoOC`) | Vacía la global y lo abre en alta; escucha solo `Alta` | `Form_FormExpedienteGeneral.cls:140-147, 663` |

Eventos hacia quien lo abrió: `Alta` (`:5`, lanzado en `:79`) y `Editado` (`:6`, lanzado en `:81`). Tras guardar, cierra (`:86`); `cmdSalir` cierra sin guardar (`:10-13`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `OrganoContratacion` (etiqueta «OC.») | texto | — | Siglas del órgano (obligatorio al guardar) | `.form.txt:926, 950`; `:66, 148` |
| `DESCRIPCION` | texto | — | Descripción opcional | `.form.txt:1005`; `:67, 151` |
| `ComandoRegistrar` (rótulo «Editar») | botón, deshabilitado en diseño | `OnClick` | Guarda el registro | `.form.txt:1142-1155`; `:42-96` |
| `ComandoAyuda` (rótulo «Salir», tip «Ayuda») | botón | `OnClick` | `AbrirAyuda` | `.form.txt:824-828`; `:15-40` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | `.form.txt:1046-1048`; `:10-13` |
| `lblTitulo` | etiqueta | — | «ALTA DE ÓRGANO DE CONTRATACIÓN» o «EDICIÓN DE ÓRGANO DE CONTRATACIÓN» | `:131, 135` |

## Flujos

### Apertura (`Form_Open`, `:165-206`)

1. `Ajustar Me` y título de ventana desde `m_TituloFormulario` (`:174, 177`). Con `EsAdministrador = Sí` habilita `ComandoRegistrar` y los dos campos; si no, deshabilita los campos (`:176-185`).
2. `EstablecerDatos`: sin órgano activo, título de alta y fin (`:130-134`). Con él, relee el registro con `constructor.getOrganoContratacion` por ID (`:137`; `constructor.bas:736-739`); si no existe, «No se ha podido encontrar EL OrganoContratacion registrado» (`:141-144`). Copia los valores no vacíos (`:146-153`).
3. Al terminar, `AllowEdits = True` para cualquier usuario (`:195`): a diferencia de `FormEjercito`, el solo lectura lo da `Enabled`, no `AllowEdits`.
4. Error: correo, `MsgBox` y cierre (`:197-205`).

### Guardado (`ComandoRegistrar_Click`, `:42-96`)

1. `HaHabidoCambios` compara siglas y descripción con `m_OrganoContratacionAlInicio`; en alta siempre hay cambios (`:98-116`). Sin cambios: «No hay cambios que guardar» (`:58-61`).
2. Copia los controles con `Nz` al objeto global (`:62-68`) y llama a `OrganoContratacionOperaciones.Registrar` (`:70-77`).
3. Lanza `Alta` o `Editado` y cierra (`:78-86`). En alta, la global queda con el ID nuevo (`OrganoContratacionOperaciones.cls:52`), que `Form_FormExpedienteGeneral` selecciona en su combo (`Form_FormExpedienteGeneral.cls:676-677`).

## Reglas, validaciones y cálculos

- **Permiso:** solo administradores editan y guardan (`:178-185`).
- **Siglas obligatorias:** vacías ⇒ «Falta el campo OrganoContratacion» (`OrganoContratacionOperaciones.cls:29-32`).
- **Siglas únicas:** en alta o si cambian, busca por nombre y rechaza con «Ese nombre de OrganoContratacion ya está registrado» (`OrganoContratacionOperaciones.cls:33-49`; `constructor.bas:741-743`).
- **Identificador:** en alta, `IDOrganoContratacionCalculado` = `DameID` (máximo + 1) (`OrganoContratacionOperaciones.cls:52`; `OrganoContratacion.cls:30`; `FUNCIONES UTILES.bas:2851-2852`).
- **Descripción:** en alta solo se escribe si no está vacía; en edición, vacía pasa a `Null` (`OrganoContratacionOperaciones.cls:65-76`).
- **Caché y resincronización:** invalida `m_ObjEntorno.ColOrganosContrataciones` (`OrganoContratacionOperaciones.cls:82`), la que usa el combo de la ficha de expediente (`Form_FormExpedienteGeneral.cls:449`); no invalida `OrganosContratacion`, la del gestor. Si una edición cambió las siglas, `ResyncCatalogo "organo"` recalcula los derivados de cada expediente con ese `IDOrganoContratacion` (`OrganoContratacionOperaciones.cls:86-93`; `ExpedienteOperaciones.cls:692-694, 712-742`). Cubierto por `Test_Issue237_RenombrarOrganoContratacion_CacheActualizaExpedientes` (`Test_Issue237CacheAltasSueltas.bas:2314`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Relectura | `TbOrganosContratacion` por ID | SQL dinámico concatenado | no | `constructor.bas:736-739` |
| Unicidad | `TbOrganosContratacion` por siglas, entre comillas | SQL dinámico | no | `constructor.bas:741-743` |
| Nuevo ID | `Max(IDOrganoContratacion)` | SQL dinámico | no | `FUNCIONES UTILES.bas:2851-2852` |
| Alta / edición | `TbOrganosContratacion` (`AddNew`/`Edit` + `Update`) | DAO | no | `OrganoContratacionOperaciones.cls:60-78` |
| Resincronización | `TbExpedientes` por `IDOrganoContratacion` y derivados | SQL dinámico + DAO | no | `ExpedienteOperaciones.cls:712-742` |

## Coexistencia

Pendiente de `../coexistencia.md`. Igual que el resto de catálogos: ID por `Max + 1` en cliente, unicidad con `=` sobre texto (sensibilidad a mayúsculas distinta en Jet y PostgreSQL) y escritura más resincronización sin transacción común.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Comunes a la familia y ya descritos en [Form_FormEjercito](Form_FormEjercito.md): error de `DameID` no consultado (`OrganoContratacion.cls:30-41`), `Max + 1` sin bloqueo, apóstrofo que rompe la unicidad (`constructor.bas:743`), renombrado sin transacción (`OrganoContratacionOperaciones.cls:60-93`), objeto del gestor alterado antes de validar (`:65-68`; `Form_FormOrganoContratacionGestion.cls:225`) y `CorreoAlAdministrador` sin `On Error Resume Next` en `Form_Open` (`:200`).
- Restos de diseño: `Caption` «Edición de Suministrador» (`.form.txt:38`, sustituido en `:177`) y `ComandoAyuda` rotulado «Salir» (`.form.txt:825`).

**Propuestas web pendientes de decisión:** ninguna aprobada; las candidatas son las de [Form_FormEjercito](Form_FormEjercito.md).

## Criterios de aceptación

Propuestos, no aprobados:

1. Un no administrador ve el registro con los campos deshabilitados y no puede guardar.
2. Guardar sin siglas muestra «Falta el campo OrganoContratacion» y no escribe.
3. Unas siglas ya registradas se rechazan; guardar el mismo registro sin renombrar, no.
4. Vaciar la descripción en edición la deja nula.
5. Renombrar un órgano actualiza los derivados de los expedientes que lo referencian.
6. Un órgano creado desde la ficha de expediente aparece seleccionado en su combo.

## Pendientes

- Si `Enabled = False` basta para impedir edición con `AllowEdits = True` en todos los controles (`:183-184, 195`).
- Contenido de `SincronizarDerivados` en la resincronización (`ExpedienteOperaciones.cls:735`).
