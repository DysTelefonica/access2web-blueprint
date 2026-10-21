# Form_FormRACSGestion

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormRACSGestion.cls` (414 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado del catálogo de RAC con filtro por nombre, alta, edición y borrado. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): la lista se rellena por código desde la caché `m_ObjEntorno.RACs`. Sigue el patrón común de los `*Gestion` de catálogos descrito en [Form_FormCPVsGestion](Form_FormCPVsGestion.md); esta ficha solo describe lo propio del RAC. Las citas son `Form_FormRACSGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| [Form_Form0BDGestorEntidades](Form_Form0BDGestorEntidades.md) | opción «9» ⇒ `FormRACSGestion`; `DoCmd.OpenForm` | `Form0BDGestorEntidadesLogic.bas:29-30`; `classes/Entorno.cls:738`; `Form_Form0BDGestorEntidades.cls:71-73` |

Abre [Form_FormRAC](Form_FormRAC.md) en alta o edición y escucha sus eventos con `WithEvents` (`:7, 72-80, 221-229`). Declara el evento `Eliminar` (`:8`, lanzado en `:181`), que nadie escucha en `src`.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `RAC` | texto de filtro | `OnKeyDown` | Intro lanza `ComandoBuscar_Click` si está habilitado | `:402-415`, `.form.txt:953-955` |
| `ListaFiltrados` | lista 2 columnas (ID oculto), con cabecera | `OnClick`, `OnDblClick` | Selecciona el RAC; doble clic edita | `:286-354`, `.form.txt:898-920` |
| `ComandoBuscar` | botón con imagen | `OnClick` | `Filtrar` | `:120-144`, `.form.txt:993-998` |
| `ComandoLimpiar` (rótulo «Eliminar», imagen de papelera) | botón | `OnClick` | Vacía el filtro y busca | `:198-201`, `.form.txt:1057-1062` |
| `ComandoAlta` | botón | `OnClick` | Abre `FormRAC` en alta | `:63-94`, `.form.txt:1267-1269` |
| `ComandoEditar` | botón | `OnClick` | Abre `FormRAC` con el seleccionado | `:203-243`, `.form.txt:1162-1164` |
| `ComandoEliminar` | botón | `OnClick` | Borra el seleccionado tras confirmar | `:146-196`, `.form.txt:1367-1369` |
| `ComandoAyuda` (rótulo «Salir»), `cmdSalir` | botones | `OnClick` | Ayuda (`:96-118`); cierre sin guardar (`:12-15`) | `.form.txt:806-808, 1557-1559` |

## Flujos

### Apertura y filtrado

1. `Form_Open`: `Alta`, `Editar` y `Eliminar` deshabilitados; `Alta` se habilita y recibe el foco con `EsAdministrador = Sí` (`:257-264`). Llama a `Filtrar` (`:265`); un error cierra el formulario (`:275-282`).
2. `Filtrar` reinicia la selección, deshabilita `Editar`/`Eliminar` y deja solo la cabecera `IDRAC;RAC` (`:27-32`). Recorre `m_ObjEntorno.RACs` (`:35`; `classes/Entorno.cls:987-1009`), descarta los que no contienen el texto del filtro en el nombre (`InStr`, `:44-48`) y añade `IDRAC;RAC` con `;` sustituido por `:` (`:49`). Selecciona la primera fila y simula el clic (`:53-54`).
3. `ListaFiltrados_Click` relee el RAC con `constructor.getRAC` si cambia la fila (`:304-318`); habilita `Editar` siempre y `Eliminar` solo a administradores (`:325-328`).

### Alta, edición y retorno

1. `ComandoAlta` vacía `m_ObjRACActivo` y abre `FormRAC` (`:72-80`); `ComandoEditar` pasa el seleccionado (`:211-229`). Sin selección: «Se ha de seleccionar un Elemento de la lista» (`:216-219`).
2. Al recibir `Alta` o `Editado` solo vuelve a buscar con el filtro vigente (`:356-400`); la caché ya la invalidó `RACOperaciones.Registrar` (`RACOperaciones.cls:90`).

### Borrado (`ComandoEliminar_Click`, `:146-196`)

1. Pide confirmación «¿Desea realmente borrar el RAC seleccionado?» con «No» por defecto (`:167-170`).
2. `RACOperaciones.Eliminar` (`:174-180`), lanza `Eliminar` y vuelve a buscar (`:181-183`).

## Reglas, validaciones y cálculos

- **RAC estructural:** el nombre «N/A» no se borra (`RACOperaciones.cls:122-125`).
- **RAC en uso:** si algún expediente lo usa (`TbExpedientesRACS`), el error lista sus `CodExp` y no borra (`RACOperaciones.cls:126-142`; `constructor.bas:3776-3778`).
- **Orden:** la caché se carga `ORDER BY RAC` (`constructor.bas:2957-2958`) en un diccionario por `IDRAC` (`constructor.bas:2980-2986`).
- **Filtro:** solo por nombre, por subcadena, con `Option Compare Database` (`:2, 45`); correo y descripción no se muestran ni filtran.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Listado | `TbRACS` | DAO, `SELECT *` cacheado en `Entorno` | — | `constructor.bas:2957-2958` |
| Detalle | `TbRACS` | SQL dinámico por `IDRAC` | — | `constructor.bas:836-838` |
| Uso en expedientes | `TbExpedientes` ⋈ `TbExpedientesRACS` | SQL dinámico | — | `constructor.bas:3776-3778` |
| Borrado | `TbRACS` | `DELETE` concatenado con `Execute` | no | `RACOperaciones.cls:143-146` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `DELETE * FROM` (sintaxis Jet) y comprobación de uso separada del borrado, sin transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Patrones compartidos con los `*Gestion` de catálogos y ya registrados en [Form_FormCPVsGestion](Form_FormCPVsGestion.md): `Filtrar` comprueba `p_Error` tras leer la caché, pero `Entorno.RACs` deja su error en `Entorno.Error` (`:35-38`; `classes/Entorno.cls:1004-1008`); `lst.Selected(1)` con la lista vacía (`:53`); comprobación de uso y `DELETE` sin transacción (`RACOperaciones.cls:126-146`); evento `Eliminar` sin oyentes (`:8, 181`).
- `ComandoEditar` está habilitado también para no administradores (`:325`) y abre `FormRAC`, donde los campos son editables pero no se puede guardar (`Form_FormRAC.cls:184-188`).
- `ComandoEditar` pasa el mismo objeto `m_RACSeleccionado` como global (`:221`); `FormRAC` lo modifica antes de guardar (`Form_FormRAC.cls:65-69`), así que un guardado fallido deja la selección del listado alterada hasta el siguiente `Filtrar` (`:28`).
- Rótulos de diseño incoherentes: ventana «GESTIÓN DE LOS PEDIDOS» (`.form.txt:36`, sobrescrito en `:260`); `RowSource` de diseño con columnas de estados (`.form.txt:916`), sustituido en `:32`; `ComandoLimpiar` rotulado «Eliminar» con imagen de papelera, igual que el borrado real (`.form.txt:1058-1062, 1368-1372`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: mostrar y filtrar también por correo; borrado protegido por clave foránea; distinguir visualmente «limpiar filtro» de «eliminar».

## Criterios de aceptación

Propuestos, no aprobados:

1. Un no administrador puede filtrar y abrir un RAC, pero no crear ni borrar.
2. Filtrar por un fragmento del nombre muestra solo los RAC que lo contienen, ordenados por nombre.
3. Borrar un RAC en uso muestra los códigos de los expedientes afectados y no borra.
4. El RAC «N/A» no se puede borrar.
5. Tras crear o editar un RAC, la lista se refresca con el filtro vigente.

## Pendientes

- Comportamiento de `Selected(1)` con la lista vacía (`:53`).
- Si `TbExpedientesRACS.IDRAC` tiene relación declarada con `TbRACS` (ERD de origen).
- Quién asigna `m_TituloFormulario` antes de abrir el formulario (`:260`).
