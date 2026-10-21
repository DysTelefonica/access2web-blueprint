# Form_FormEjercitosGestion

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormEjercitosGestion.cls` (414 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado del catálogo de ejércitos con filtro por nombre, alta, edición y borrado. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): la lista se rellena desde la caché `m_ObjEntorno.Ejercitos`. El alta y la edición se delegan en [Form_FormEjercito](Form_FormEjercito.md). Sigue el mismo patrón que otros gestores de catálogo (p. ej. `Form_FormCPVsGestion.cls:54, 380`); esta ficha describe solo el caso de ejércitos. Las citas son `Form_FormEjercitosGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDGestorEntidades` (opción 3) | `ResolverFormDestino` devuelve `FormEjercitosGestion` y se abre con `DoCmd.OpenForm` sin `OpenArgs` | `Form_Form0BDGestorEntidades.cls:58-71`; `Form0BDGestorEntidadesLogic.bas:17-18` |

Abre `FormEjercito` en alta (`:73-81`) o edición (`:222-230`) y escucha sus eventos `Alta` y `Editado` con `WithEvents` (`:7, 371-415`). Declara el evento `Eliminar` (`:8`), lanzado tras un borrado (`:182`). `cmdSalir` cierra el formulario (`:12-15`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Ejercito` | texto (filtro) | `OnKeyDown` | Con Intro ejecuta la búsqueda y anula la tecla | `.form.txt:953-955`; `:247-260` |
| `ComandoBuscar` (tip «Ejecutar la búsqueda») | botón | `OnClick` | `Filtrar` | `.form.txt:993-997`; `:121-145` |
| `ComandoLimpiar` (tip «Limpiar campo Búsqueda») | botón | `OnClick` | Vacía el filtro y busca | `.form.txt:1058-1062`; `:199-202` |
| `ListaFiltrados` | lista 3 columnas (ID oculto), con encabezados | `OnClick`, `OnDblClick` | Selecciona; doble clic equivale a Editar | `.form.txt:898-920`; `:301-369` |
| `ComandoEditar` (tip «Editar/Ver Registro») | botón, deshabilitado en diseño | `OnClick` | Abre `FormEjercito` con el seleccionado | `.form.txt:1153-1167`; `:204-244` |
| `ComandoAlta` (tip «Añadir Registro») | botón | `OnClick` | Abre `FormEjercito` en alta | `.form.txt:1268-1272`; `:64-95` |
| `ComandoEliminar` (tip «Eliminar registro») | botón, deshabilitado en diseño | `OnClick` | Borra el seleccionado tras confirmar | `.form.txt:1357-1372`; `:147-197` |
| `ComandoAyuda`, `cmdSalir` | botones | `OnClick` | Ayuda y cierre | `:97-119, 12-15` |

## Flujos

### Apertura (`Form_Open`, `:262-298`)

1. Deshabilita Alta, Editar y Eliminar; solo con `EsAdministrador = Sí` habilita Alta (`:272-279`).
2. `Filtrar` carga la lista (`:280`). Un error cierra el formulario tras correo y `MsgBox` (`:290-297`).

### Filtrado (`Filtrar`, `:18-62`)

1. Vacía la selección, deshabilita Editar y Eliminar y deja en `RowSource` solo la fila de encabezados (`:28-32`).
2. Recorre `m_ObjEntorno.Ejercitos` (`:35`; `Entorno.cls:895-917`, que carga `constructor.getEjercitos` sin `ORDER BY`, `constructor.bas:2841-2842`).
3. Con filtro informado, descarta los nombres que no lo contienen (`InStr`) (`:45-49`). Sustituye `;` por `:` en nombre y descripción antes de añadir la fila (`:50`).
4. Selecciona la primera fila de datos y ejecuta `ListaFiltrados_Click` (`:54-55`).

### Selección (`ListaFiltrados_Click`, `:301-359`)

Sin ID en la fila, limpia la selección (`:312-318`). Si cambia el ID, relee el ejército con `constructor.getEjercito` (`:319-333`). Con ejército válido habilita Editar para cualquier usuario y Eliminar solo para administradores (`:340-343`).

### Alta y edición

Alta vacía `m_ObjEjercitoActivo`; Editar le asigna el seleccionado; ambos cierran `FormEjercito` si estaba abierto y lo abren (`:73-81, 222-230`). Al recibir `Alta` o `Editado`, invalida la caché y vuelve a buscar (`:378-379, 401-402`). La edición en solo lectura la controla `FormEjercito` (`Form_FormEjercito.cls:175-181`).

### Borrado (`ComandoEliminar_Click`, `:147-197`)

1. Sin selección, «Se ha de seleccionar un Elemento de la lista» (`:155-164`).
2. Confirmación con «No» por defecto (`:168-171`).
3. `EjercitoOperaciones.Eliminar` (`:175-183`): rechaza el registro «N/A» como estructural (`EjercitoOperaciones.cls:112-115`); si algún expediente lo usa, lista sus `CodExp` y rechaza (`:116-132`); si no, `DELETE` e invalida la caché (`:134-139`).
4. Lanza `Eliminar` y vuelve a buscar (`:182-184`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Listado | `TbEjercitos` (vía caché de `Entorno`) | SQL dinámico | no | `constructor.bas:2841-2842` |
| Selección | `TbEjercitos` por `IDEjercito` | SQL dinámico concatenado | no | `constructor.bas:442-444` |
| Uso en expedientes | `TbExpedientes` por `IDEjercito` | SQL dinámico concatenado | no | `constructor.bas:3395-3397` |
| Borrado | `DELETE` en `TbEjercitos` | SQL dinámico (`Execute`) | no | `EjercitoOperaciones.cls:134-137` |

Altas y ediciones: ver [Form_FormEjercito](Form_FormEjercito.md#lecturas-y-escrituras).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: la comprobación de uso y el `DELETE` no comparten transacción; la integridad referencial con `TbExpedientes` la aplica el código, no consta como relación (a confirmar en `../erd-origen.md`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Si la carga de la caché falla, el error queda en `m_ObjEntorno.Error` y `Filtrar` comprueba `p_Error`, que sigue vacío: la lista sale vacía sin aviso (`:35-41`; `Entorno.cls:905-916`).
- `lst.Selected(1) = True` con la lista sin datos (filtro sin resultados) podría producir error (`:54`).
- El mensaje de borrado rechazado dice «Ese RAC forma parte de estos expedientes» (`EjercitoOperaciones.cls:130`), copiado de `RACOperaciones.cls:140`.
- La comprobación de uso solo mira `TbExpedientes.IDEjercito` (`constructor.bas:3395-3397`); no se ha comprobado si otras tablas lo referencian.
- Restos de diseño sustituidos en ejecución: `Caption` «GESTIÓN DE LOS PEDIDOS» (`.form.txt:36`) y `RowSource` de CPV en la lista (`.form.txt:916`, sustituido en `:32`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: orden alfabético explícito del listado; borrado protegido por clave foránea en lugar de comprobación previa; marcar «N/A» como registro de sistema en datos y no por nombre.

## Criterios de aceptación

Propuestos, no aprobados:

1. Solo un administrador puede dar de alta o eliminar; cualquier usuario puede abrir un registro.
2. Escribir un texto en el filtro y pulsar Intro muestra solo los ejércitos cuyo nombre lo contiene; Limpiar vuelve a mostrarlos todos.
3. Eliminar pide confirmación con «No» como opción por defecto.
4. No se puede eliminar el registro «N/A» ni un ejército usado por expedientes; en este caso se listan sus códigos.
5. Tras un alta, edición o borrado, el listado refleja el cambio.

## Pendientes

- Comportamiento de `Selected(1)` con lista vacía y encabezados activos (`:54`).
- Quién escucha el evento `Eliminar` (`:8`): no se ha encontrado ningún `WithEvents` sobre este formulario.
- Origen del permiso `EsAdministrador` (ver [Form_FormExpediente](Form_FormExpediente.md#reglas-validaciones-y-cálculos)).
