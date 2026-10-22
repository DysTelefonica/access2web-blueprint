# Form_FormCPVsGestion

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormCPVsGestion.cls` (416 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado del catálogo de CPV con filtro por código, alta, edición y borrado. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): la lista se rellena por código desde la caché `m_ObjEntorno.CPVs`. Sigue el patrón común de los formularios `*Gestion` de catálogos (misma estructura `Filtrar` con `lst.Selected(1)`, p. ej. `Form_FormEjercitosGestion.cls:35, 54`); esta ficha solo describe lo propio del CPV. Las citas son `Form_FormCPVsGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| [Form_Form0BDGestorEntidades](Form_Form0BDGestorEntidades.md) | opción «2» ⇒ `FormCPVsGestion`; `DoCmd.OpenForm` | `Form0BDGestorEntidadesLogic.bas:15-16`; `Form_Form0BDGestorEntidades.cls:58-73` |

Abre [Form_FormCPV](Form_FormCPV.md) en alta o edición y escucha sus eventos con `WithEvents` (`:7, 73-81, 222-230`). Declara el evento `Eliminar` (`:8`, lanzado en `:182`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `CPV` | texto de filtro | `OnKeyDown` | Intro lanza `ComandoBuscar_Click` si está habilitado | `:249-262`, `.form.txt:921-923` |
| `ListaFiltrados` | lista 2 columnas (ID oculto), con cabecera | `OnClick`, `OnDblClick` | Selecciona el CPV; doble clic edita | `:303-371`, `.form.txt:943-966` |
| `ComandoBuscar` | botón | `OnClick` | `Filtrar` | `:121-145`, `.form.txt:996-998` |
| `ComandoLimpiar` (rótulo «Eliminar» en `.form.txt`) | botón | `OnClick` | Vacía el filtro y busca | `:199-202`, `.form.txt:1060-1062` |
| `ComandoAlta` | botón | `OnClick` | Abre `FormCPV` en alta | `:64-95` |
| `ComandoEditar` | botón | `OnClick` | Abre `FormCPV` con el seleccionado | `:204-244` |
| `ComandoEliminar` | botón | `OnClick` | Borra el seleccionado tras confirmar | `:147-197` |
| `ComandoAyuda` (rótulo «Salir»), `cmdSalir` | botones | `OnClick` | Ayuda (`:97-119`); cierre sin guardar (`:12-15`) | `.form.txt:807-809, 1560-1562` |

## Flujos

### Apertura y filtrado

1. `Form_Open`: `Alta`, `Editar` y `Eliminar` deshabilitados; `Alta` se habilita con `EsAdministrador = Sí` (`:273-281`). Llama a `Filtrar` (`:282`); un error cierra el formulario (`:292-299`).
2. `Filtrar` reinicia la selección, deshabilita `Editar`/`Eliminar` y deja solo la cabecera `IDCPV;CPV` (`:27-32`). Recorre `m_ObjEntorno.CPVs` (`:35`; `classes/Entorno.cls:865-879`), descarta los que no contienen el texto del filtro en el código (`InStr`, `:45-49`) y añade `IDCPV;CPV` con `;` sustituido por `:` (`:50`). Selecciona la primera fila y simula el clic (`:54-55`).
3. `ListaFiltrados_Click` relee el CPV con `constructor.getCPV` si cambia la fila (`:321-335`); habilita `Editar` siempre y `Eliminar` solo a administradores (`:342-345`).

### Alta, edición y retorno

1. `ComandoAlta` vacía `m_ObjCPVActivo` y abre `FormCPV` (`:73-81`); `ComandoEditar` pasa el seleccionado (`:212-230`). Sin selección: «Se ha de seleccionar un Elemento de la lista» (`:217-220`).
2. Al recibir `Alta` o `Editado`, invalida la caché y vuelve a buscar con el filtro vigente (`:373-417`).

### Borrado (`ComandoEliminar_Click`, `:147-197`)

1. Pide confirmación «¿Desea realmente borrar al CPV seleccionado?» con «No» por defecto (`:168-171`).
2. `CPVOperaciones.Eliminar` (`:175-183`), lanza `Eliminar` y vuelve a buscar (`:182-184`).

## Reglas, validaciones y cálculos

- **CPV estructural:** el código «N/A» no se borra (`CPVOperaciones.cls:98-101`).
- **CPV en uso:** si algún expediente lo usa (`TbExpedientesCPVs`), el error lista sus `CodExp` y no borra (`CPVOperaciones.cls:102-118`; `constructor.bas:2733-2735`).
- **Orden:** la caché se carga `ORDER BY CPV` (`constructor.bas:2785-2786`) en un diccionario por `IDCPV` (`constructor.bas:2808-2814`).
- **Filtro:** solo por código, por subcadena, con `Option Compare Database` (`:2, 46`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Listado | `TbCPV` | DAO, `SELECT *` cacheado en `Entorno` | — | `constructor.bas:2785-2792` |
| Detalle | `TbCPV` | SQL dinámico por `IDCPV` | — | `constructor.bas:391-394` |
| Uso en expedientes | `TbExpedientes` ⋈ `TbExpedientesCPVs` | SQL dinámico | — | `constructor.bas:2733-2736` |
| Borrado | `TbCPV` | `DELETE` concatenado con `Execute` | no | `CPVOperaciones.cls:120-123` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `DELETE * FROM` (sintaxis Jet) y comprobación de uso separada del borrado, sin transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Filtrar` comprueba `p_Error` tras leer `m_ObjEntorno.CPVs` (`:35-38`), pero la propiedad deja su error en `Entorno.Error` (`classes/Entorno.cls:883-886`): un fallo de carga mostraría una lista vacía sin aviso.
- `lst.Selected(1) = True` con la lista sin filas (`:54`): sin verificar si Access da error cuando el filtro no encuentra nada. Patrón compartido por los diez `*Gestion` de catálogos.
- Comprobación de uso y `DELETE` sin transacción (`CPVOperaciones.cls:102-123`): un expediente que reciba el CPV entre ambos pasos quedaría con referencia huérfana si no hay integridad referencial declarada.
- Nadie escucha `Eliminar` (`:8, 182`): no hay `WithEvents` de `Form_FormCPVsGestion` en `src`.
- Invalidación duplicada de la caché: `CPVOperaciones.Registrar` ya la hace (`CPVOperaciones.cls:78`) y los handlers la repiten (`:380, 403`).
- El `RowSource` de diseño declara una tercera columna `DESCRIPCIÓN` (`.form.txt:962`) que el código no rellena (`:32, 50`); título de ventana de diseño «GESTIÓN DE LOS PEDIDOS» (`.form.txt:36`, sobrescrito en `:277`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtrar también por descripción y mostrarla; borrado protegido por clave foránea en lugar de comprobación previa.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un no administrador puede filtrar y abrir un CPV, pero no crear ni borrar.
2. Filtrar por un fragmento del código muestra solo los CPV que lo contienen, ordenados por código.
3. Borrar un CPV en uso muestra los códigos de los expedientes afectados y no borra.
4. El CPV «N/A» no se puede borrar.
5. Tras crear o editar un CPV, la lista se refresca con el filtro vigente.

## Pendientes

- Comportamiento de `Selected(1)` con la lista vacía (`:54`).
- Si `TbExpedientesCPVs.IDCPV` tiene relación declarada con `TbCPV` (ERD de origen).
- Si `ComandoBuscar` y `ComandoLimpiar` muestran imagen en lugar de rótulo (`.form.txt:997, 1061`).
