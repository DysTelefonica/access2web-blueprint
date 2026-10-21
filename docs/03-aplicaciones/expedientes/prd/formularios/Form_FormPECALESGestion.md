# Form_FormPECALESGestion

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormPECALESGestion.cls` (424 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado del catálogo de PECAL con filtro por código, alta, edición y borrado. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): la lista se rellena por código desde la caché `m_ObjEntorno.PECALES`. Sigue el patrón común de los `*Gestion` de catálogos descrito en [Form_FormCPVsGestion](Form_FormCPVsGestion.md); esta ficha solo detalla lo propio de la PECAL. Las citas son `Form_FormPECALESGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| [Form_Form0BDGestorEntidades](Form_Form0BDGestorEntidades.md) | opción «8» ⇒ `FormPECALESGestion`; `DoCmd.OpenForm` | `Form0BDGestorEntidadesLogic.bas:27-28`; `Form_Form0BDGestorEntidades.cls:58-73`; `classes/Entorno.cls:737` |

Abre [Form_FormPECAL](Form_FormPECAL.md) en alta o edición y escucha sus eventos con `WithEvents` (`:7, 73-81, 222-230`). Declara el evento `Eliminar` (`:8`, lanzado en `:182`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `PECAL` | texto de filtro | `OnExit` → `PECAL_Exit` | Al salir del campo lanza `ComandoBuscar_Click` | `:247-269`, `.form.txt:941-956` |
| `ListaFiltrados` | lista 3 columnas (ID oculto, PECAL, descripción), con cabecera | `OnClick`, `OnDblClick` | Selecciona la PECAL; doble clic edita | `:311-379`, `.form.txt:898-921` |
| `ComandoBuscar` (rótulo «Comando36») | botón | `OnClick` | `Filtrar` | `:121-145`, `.form.txt:996-998` |
| `ComandoLimpiar` (rótulo «Eliminar» en `.form.txt`) | botón | `OnClick` | Vacía el filtro y busca | `:199-202`, `.form.txt:1060-1062` |
| `ComandoAlta` | botón | `OnClick` | Abre `FormPECAL` en alta | `:64-95`, `.form.txt:1270-1272` |
| `ComandoEditar` | botón | `OnClick` | Abre `FormPECAL` con la seleccionada | `:204-244`, `.form.txt:1165-1167` |
| `ComandoEliminar` | botón | `OnClick` | Borra la seleccionada tras confirmar | `:147-197`, `.form.txt:1370-1372` |
| `ComandoAyuda` (rótulo «Salir»), `cmdSalir` | botones | `OnClick` | Ayuda (`:97-119`); cierre sin guardar (`:12-15`) | `.form.txt:808-810, 1560-1562` |

## Flujos

### Apertura y filtrado

1. `Form_Open`: `Alta`, `Editar` y `Eliminar` deshabilitados; `Alta` se habilita y recibe el foco con `EsAdministrador = Sí` (`:282-289`). Llama a `Filtrar` (`:290`); un error cierra el formulario (`:300-307`).
2. `Filtrar` reinicia la selección, deshabilita `Editar`/`Eliminar` y deja solo la cabecera `IDPECAL;PECAL;DESCRIPCIÓN` (`:27-32`). Recorre `m_ObjEntorno.PECALES` (`:35`; `classes/Entorno.cls:956-978`), descarta las que no contienen el filtro en el código (`InStr`, `:45-49`) y añade ID, código y descripción con `;` sustituido por `:` (`:50`). Selecciona la primera fila y simula el clic (`:54-55`).
3. `ListaFiltrados_Click` relee la PECAL con `constructor.getPecal` si cambia la fila (`:329-343`); habilita `Editar` siempre y `Eliminar` solo a administradores (`:350-353`).

### Alta, edición y retorno

1. `ComandoAlta` vacía `m_ObjPECALActiva` y abre `FormPECAL` (`:73-81`); `ComandoEditar` pasa la seleccionada (`:212-230`). Sin selección: «Se ha de seleccionar un Elemento de la lista» (`:217-220`).
2. Al recibir `Alta` o `Editado`, invalida la caché y vuelve a buscar con el filtro vigente (`:381-425`).

### Borrado (`ComandoEliminar_Click`, `:147-197`)

1. Pide confirmación «¿Desea realmente borrar la PECAL seleccionada?» con «No» por defecto (`:168-171`).
2. `PECALOperaciones.Eliminar` (`:175-181`), lanza `Eliminar` y vuelve a buscar (`:182-184`).

## Reglas, validaciones y cálculos

- **PECAL estructural:** el código «N/A» no se borra: «Es una opción Estructural que otras herramientas necesitan.No se puede borrar» (`PECALOperaciones.cls:112-115`).
- **PECAL en uso:** si algún expediente la usa (`TbExpedientesPECAL`), el error lista sus `CodExp` y no borra (`PECALOperaciones.cls:116-132`; `constructor.bas:3830-3832`).
- **Orden:** la caché se carga `ORDER BY PECAL` (`constructor.bas:3166-3168`) en un diccionario por `IDPECAL` (`constructor.bas:3194-3195`).
- **Filtro:** solo por código, por subcadena, con `Option Compare Database` (`:2, 46`); la descripción se muestra pero no filtra.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Listado | `TbPECAL` | DAO, `SELECT *` cacheado en `Entorno` | — | `constructor.bas:3166-3168` |
| Detalle | `TbPECAL` | SQL dinámico por `IDPECAL` | — | `constructor.bas:786-789` |
| Uso en expedientes | `TbExpedientes` ⋈ `TbExpedientesPECAL` | SQL dinámico | — | `constructor.bas:3830-3833` |
| Borrado | `TbPECAL` | `DELETE` concatenado con `Execute` | no | `PECALOperaciones.cls:133-136` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `DELETE * FROM` (sintaxis Jet) y comprobación de uso separada del borrado, sin transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `PECAL_Exit` relanza la búsqueda cada vez que el foco sale del filtro (`:247-255`), aunque el texto no haya cambiado y también al pulsar `Salir`; cada búsqueda reinicia la selección a la primera fila (`:28, 54-55`). Los demás `*Gestion` buscan con Intro (`Form_FormCPVsGestion.cls:249`).
- Comunes a la familia `*Gestion` (ver [Form_FormCPVsGestion](Form_FormCPVsGestion.md)): `Filtrar` mira `p_Error` (`:35-38`) aunque el fallo de la caché queda en `Entorno.Error` (`classes/Entorno.cls:975-977`); `lst.Selected(1)` con la lista vacía (`:54`); comprobación de uso y `DELETE` sin transacción (`PECALOperaciones.cls:116-136`); nadie escucha `Eliminar` (`:8, 182`).
- Restos de diseño: `RowSource` «IDESTADO;ESTADO;DESCRIPCIÓN» (`.form.txt:917`, sobrescrito en `:32`), título de ventana «GESTIÓN DE LOS PEDIDOS» (`.form.txt:36`, sobrescrito en `:285`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtrar también por descripción; búsqueda explícita en lugar de al salir del campo; borrado protegido por clave foránea.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un no administrador puede filtrar y abrir una PECAL, pero no crear ni borrar.
2. Filtrar por un fragmento del código muestra solo las PECAL que lo contienen, ordenadas por código, con su descripción.
3. Borrar una PECAL en uso muestra los códigos de los expedientes afectados y no borra.
4. La PECAL «N/A» no se puede borrar.
5. Tras crear o editar una PECAL, la lista se refresca con el filtro vigente.

## Pendientes

- Orden real de `Exit` del filtro frente al `Click` del botón pulsado (`:247-255`).
- Si `TbExpedientesPECAL.IDPECAL` tiene relación declarada con `TbPECAL` (ERD de origen).
- Si `ComandoBuscar` y `ComandoLimpiar` muestran imagen en lugar de rótulo (`.form.txt:997, 1061`).
