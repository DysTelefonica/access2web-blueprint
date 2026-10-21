# Form_FormOrganoContratacionGestion

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormOrganoContratacionGestion.cls` (428 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado del catálogo de órganos de contratación con filtro por siglas, alta, edición y borrado. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): la lista se rellena desde la caché `m_ObjEntorno.OrganosContratacion`. Alta y edición se delegan en [Form_FormOrganoContratacion](Form_FormOrganoContratacion.md). Sigue el patrón de gestor de catálogo descrito en [Form_FormEjercitosGestion](Form_FormEjercitosGestion.md); esta ficha recoge solo lo propio. Las citas son `Form_FormOrganoContratacionGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDGestorEntidades` (opción 7) | `ResolverFormDestino` devuelve `FormOrganoContratacionGestion` y se abre sin `OpenArgs` | `Form0BDGestorEntidadesLogic.bas:25-26`; `Form_Form0BDGestorEntidades.cls:58-71` |

Abre `FormOrganoContratacion` en alta (`:74-82`) o edición (`:225-233`) y escucha `Alta` y `Editado` con `WithEvents` (`:7, 384-428`). Declara `Eliminar` (`:8`), lanzado tras un borrado (`:183`). `cmdSalir` cierra (`:12-15`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `OrganoContratacion` | texto (filtro) | `OnExit` | Al salir del campo ejecuta la búsqueda (no hay `OnKeyDown`, a diferencia de ejércitos) | `.form.txt:649-650`; `:250-272` |
| `ComandoBuscar` (tip «Ejecutar la búsqueda») | botón | `OnClick` | `Filtrar` | `.form.txt:689-693`; `:122-146` |
| `ComandoLimpiar` (tip «Limpiar campo Búsqueda») | botón | `OnClick` | Vacía el filtro y busca | `.form.txt:753-757`; `:200-203` |
| `ListaFiltrados` | lista 3 columnas (ID oculto), con encabezados | `OnClick`, `OnDblClick` | Selecciona; doble clic equivale a Editar | `.form.txt:593-616`; `:314-382` |
| `ComandoEditar` | botón, deshabilitado en diseño | `OnClick` | Abre `FormOrganoContratacion` con el seleccionado | `.form.txt:848-862`; `:205-247` |
| `ComandoAlta` | botón | `OnClick` | Abre `FormOrganoContratacion` en alta | `.form.txt:963-967`; `:65-96` |
| `ComandoEliminar` | botón, deshabilitado en diseño | `OnClick` | Borra el seleccionado tras confirmar | `.form.txt:1052-1067`; `:148-198` |
| `ComandoAyuda`, `cmdSalir` | botones | `OnClick` | Ayuda y cierre | `:98-120, 12-15` |
| `lblTitulo` | etiqueta | — | No se asigna en el `.cls`: muestra el rótulo de diseño | `.form.txt:483-484` |

## Flujos

- **Apertura** (`Form_Open`, `:274-311`): deshabilita Alta, Editar y Eliminar; con `EsAdministrador = Sí` habilita Alta (`:285-292`). `Filtrar` carga la lista (`:293`); un error cierra el formulario (`:303-310`).
- **Filtrado** (`:18-63`): encabezados en `RowSource` (`:32`), recorre la caché (`:35`; `Entorno.cls:1104-1126`, que carga `TbOrganosContratacion` sin `ORDER BY`, `constructor.bas:3611-3612`), descarta las siglas que no contienen el filtro (`InStr`, `:45-49`), cambia `;` por `:` (`:50-51`) y selecciona la primera fila (`:55-56`).
- **Selección** (`:314-372`): relee el órgano si cambia el ID (`:333-347`); habilita Editar para cualquiera y Eliminar solo para administradores (`:354-357`).
- **Alta y edición:** al recibir `Alta` o `Editado`, invalida `OrganosContratacion` y vuelve a buscar (`:391-392, 414-415`).
- **Borrado** (`:148-198`): exige selección (`:156-165`), confirma con «No» por defecto (`:169-172`) y llama a `OrganoContratacionOperaciones.Eliminar` (`:176-184`): rechaza «N/A» como estructural (`OrganoContratacionOperaciones.cls:110-113`), consulta la caché de órganos en uso (`:114-119`) y ejecuta el `DELETE` (`:120-123`). Después lanza `Eliminar` y vuelve a buscar sin invalidar la caché (`:183-185`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Listado | `TbOrganosContratacion` (caché de `Entorno`) | SQL dinámico | no | `constructor.bas:3611-3612` |
| Selección | `TbOrganosContratacion` por ID | SQL dinámico concatenado | no | `constructor.bas:736-739` |
| Uso en expedientes | `TbExpedientes` ⋈ `TbOrganosContratacion` por texto | SQL dinámico | no | `constructor.bas:3664-3667` |
| Borrado | `DELETE` en `TbOrganosContratacion` | SQL dinámico (`Execute`) | no | `OrganoContratacionOperaciones.cls:120-123` |

## Coexistencia

Pendiente de `../coexistencia.md`. La integridad con `TbExpedientes.IDOrganoContratacion` (columna que escribe `ExpedienteOperaciones.cls:197-200`) no la garantiza el código (ver defectos); a confirmar si existe relación declarada en `../erd-origen.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **La comprobación de uso nunca bloquea:** `Entorno.ColOrganoContratacioneSEnExpedientes` pasa `m_Error` como primer argumento, que es `p_IDOrganoContratacion` (`Entorno.cls:1202`; `constructor.bas:3649-3652`); con ID vacío la función devuelve `Nothing` (`constructor.bas:3661-3663`) y `Eliminar` salta la comprobación (`OrganoContratacionOperaciones.cls:114`). Se podría borrar un órgano usado por expedientes.
- La consulta de uso, además, tiene un paréntesis sobrante y cruza por `TbExpedientes.OrganoContratacion` (texto), no por `IDOrganoContratacion` (`constructor.bas:3666-3667`).
- `Eliminar` no invalida ninguna caché (`OrganoContratacionOperaciones.cls:104-134`) y el gestor tampoco antes de buscar (`:185`): el órgano borrado seguiría en la lista y en el combo de la ficha de expediente hasta recargar.
- `lblTitulo` muestra «GESTIÓN DE LAS OFICINAS DE PROGRAMA» (`.form.txt:484`); `Caption` «GESTIÓN DE LOS PEDIDOS» (`.form.txt:36`, sustituido en `:288`); `RowSource` de estados en la lista (`.form.txt:612`, sustituido en `:32`); `ComandoLimpiar` rotulado «Eliminar» (`.form.txt:754`).
- Comunes a la familia, ya descritos en [Form_FormEjercitosGestion](Form_FormEjercitosGestion.md): error de carga de la caché no comprobado (`:35-38`; `Entorno.cls:1125`) y `Selected(1)` con filtro sin resultados (`:55`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: clave foránea `expedientes → organos_contratacion` con borrado restringido; las demás, como en [Form_FormEjercitosGestion](Form_FormEjercitosGestion.md).

## Criterios de aceptación

Propuestos, no aprobados:

1. Solo un administrador puede dar de alta o eliminar; cualquier usuario puede abrir un registro.
2. Al salir del filtro se listan solo los órganos cuyas siglas lo contienen; Limpiar los muestra todos.
3. Eliminar pide confirmación con «No» por defecto; «N/A» no se puede eliminar.
4. Un órgano usado por algún expediente no se puede eliminar.
5. Tras alta, edición o borrado, el listado y el combo de la ficha de expediente reflejan el cambio.

## Pendientes

- Si `TbExpedientes` tiene columna `OrganoContratacion` de texto (`constructor.bas:3666`) y si existe relación declarada con `TbOrganosContratacion`.
- Quién escucha el evento `Eliminar` (`:8`): no se ha encontrado `WithEvents` sobre este formulario.
- Si Access dispara `OnExit` del filtro al pulsar Intro en el último control (`:250`).
