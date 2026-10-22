# Form_FormGradosClasificacionGestion

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormGradosClasificacionGestion.cls` (425 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado del catálogo de grados de clasificación con búsqueda por nombre, alta, edición y borrado. Sigue el patrón común de los formularios `*Gestion` de catálogos: el código es casi idéntico al de `Form_FormEjercitosGestion.cls` salvo nombres y el disparo de la búsqueda (aquí `OnExit` del campo, allí `KeyDown`). Formulario no enlazado: la lista se rellena desde la caché `m_ObjEntorno.GradosClasificaciones`. Las citas son `Form_FormGradosClasificacionGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| [Form_Form0BDGestorEntidades](Form_Form0BDGestorEntidades.md) | Opción «10» → `ResolverFormDestino` → `DoCmd.OpenForm` sin `OpenArgs` | `Form0BDGestorEntidadesLogic.bas:31-32`, `Form_Form0BDGestorEntidades.cls:58-71` |
| Catálogo de opciones del entorno | `EnumOtrasOpciones.GradosClasificacion` → nombre del formulario | `classes/Entorno.cls:732` |

Abre [Form_FormGradoClasificacion](Form_FormGradoClasificacion.md) y escucha sus eventos `Alta` y `Editado` (`:7, 382-426`). Declara el evento `Eliminar` (`:8`, lanzado en `:184`); no se ha encontrado quién lo escucha. `CmdSalir_Click` cierra (`:12-15`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `GradoClasificacion` | texto | `OnExit` → `GradoClasificacion_Exit` | Lanza la búsqueda al salir del campo | `:249-271`, `.form.txt:953-954` |
| `ComandoBuscar` (rótulo «Comando36») | botón | `OnClick` | `Filtrar` | `:123-147`, `.form.txt:993-995` |
| `ComandoLimpiar` (rótulo «Eliminar», ayuda «Limpiar campo Búsqueda») | botón | `OnClick` | Vacía el filtro y busca | `:201-204`, `.form.txt:1058-1062` |
| `ListaFiltrados` | lista 3 columnas (ID oculto, encabezados) | `OnClick`, `OnDblClick` | Selecciona y habilita botones; doble clic edita | `:312-380`, `.form.txt:898, 914-920` |
| `ComandoAlta` | botón | `OnClick` | Abre la ficha vacía | `:66-97`, `.form.txt:1268-1270` |
| `ComandoEditar` | botón | `OnClick` | Abre la ficha con el seleccionado | `:206-246`, `.form.txt:1163-1165` |
| `ComandoEliminar` | botón | `OnClick` | Confirma y borra | `:149-199`, `.form.txt:1368-1370` |
| `ComandoAyuda` (rótulo «Salir»), `cmdSalir` | botones | `OnClick` | Ayuda; cerrar | `:99-121, 12-15`, `.form.txt:806-808, 1558-1560` |

## Flujos

### Apertura y búsqueda

1. `Form_Open` deshabilita los tres botones de acción, sustituye el `Caption` de diseño por `m_TituloFormulario` y habilita `ComandoAlta` solo a administradores (`:282-290`); después `Filtrar` (`:291`). Un error cierra el formulario (`:301-308`).
2. `Filtrar` reinicia la selección, sustituye el `RowSource` de diseño por la fila de encabezados (`:28-32`; `.form.txt:916`) y añade cada grado de la caché cuyo nombre contenga el texto buscado con `InStr` (`:35-55`). Cambia `;` por `:` en los textos (`:50-52`). Selecciona la primera fila de datos y simula el clic (`:56-57`).
3. `ListaFiltrados_Click` carga el grado con `constructor.getGradoClasificacion` si cambia la fila (`:330-344`); habilita `ComandoEditar` para todos y `ComandoEliminar` solo a administradores (`:351-354`).

### Alta y edición

`ComandoAlta` vacía `m_ObjGradoClasificacionActivo` (`:75`); `ComandoEditar` le asigna el seleccionado (`:224`). Ambos cierran la ficha si estaba abierta y la abren (`:77-83, 226-232`). Al recibir `Alta` o `Editado`, invalidan la caché y vuelven a buscar (`:389-390, 412-413`).

### Borrado (`ComandoEliminar_Click`, `:149-199`)

1. Sin selección, intenta cargar la fila; si no hay, «Se ha de seleccionar un Elemento de la lista» (`:157-166`).
2. Confirma con «¿Desea realmente borrar al GradoClasificacion seleccionado?», con «No» por defecto (`:170-173`).
3. `GradoClasificacionOperaciones.Eliminar` (`:177-183`), lanza `Eliminar` y refresca (`:184-186`).

## Reglas, validaciones y cálculos

- El filtro solo busca en el nombre, no en la descripción (`:45-48`). Sin `ORDER BY`: el orden es el de lectura (`constructor.bas:2898-2899`).
- `Eliminar` rechaza el grado «N/A» con «Es una opción Estructural que otras herramientas necesitan.No se puede borrar» (`GradoClasificacionOperaciones.cls:111-114`).
- `Eliminar` rechaza un grado en uso, «Al menos un expediente usa este grado de clasificación», según la caché `ColGradosClasificacionesEnExpedientes` (`GradoClasificacionOperaciones.cls:115-120`), que cruza `TbExpedientes.GradoClasificacion` con el **nombre** del grado (`constructor.bas:2421-2423`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Listado (caché) | `TbGradosClasificacion` completa | SQL dinámico | `constructor.bas:2898-2905`, `classes/Entorno.cls:1134-1146` |
| Grados en uso (caché) | `TbExpedientes` ⋈ `TbGradosClasificacion` por nombre | SQL dinámico | `constructor.bas:2421-2424`, `classes/Entorno.cls:1637-1650` |
| Borrado | `DELETE` en `TbGradosClasificacion` por ID | SQL dinámico, `Execute` sin transacción | `GradoClasificacionOperaciones.cls:121-124` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `DELETE *` es sintaxis Jet (`GradoClasificacionOperaciones.cls:121`); la comprobación de uso depende de una caché en memoria y no de una clave foránea.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La comprobación de uso antes de borrar cruza por `TbExpedientes.GradoClasificacion` (nombre) (`constructor.bas:2423`), pero el expediente guarda `IDGradoClasificacion` (`ExpedienteOperaciones.cls:191-194`): podría no detectar grados en uso.
- Esa caché no se invalida en ningún punto del código (no hay asignación a `Nothing`), y si su carga falla devuelve `Nothing` y el borrado sigue (`GradoClasificacionOperaciones.cls:115`; `classes/Entorno.cls:1652-1657`).
- `Filtrar` comprueba `p_Error` tras leer la caché (`:35-38`), pero el error de `Entorno` queda en `.Error`: un fallo de lectura muestra una lista vacía sin aviso.
- Con la lista sin datos, `lst.Selected(1)` (`:56`) podría producir error al abrir.
- Rótulos de diseño incoherentes: «Eliminar» en `ComandoLimpiar`, «Salir» en `ComandoAyuda` y título «GESTIÓN DE LOS PEDIDOS» (`.form.txt:36, 807, 1059`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: integridad referencial con clave foránea en lugar de caché; búsqueda también por descripción; listado ordenado.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, la lista muestra todos los grados; escribir un texto y salir del campo filtra por nombre.
2. Solo un administrador ve habilitados «Alta» y «Eliminar».
3. Borrar el grado «N/A» se rechaza con su mensaje.
4. Borrar un grado usado por un expediente se rechaza con «Al menos un expediente usa este grado de clasificación».
5. Tras un alta o una edición, la lista refleja el cambio sin reabrir el formulario.

## Pendientes

- Si `TbExpedientes` conserva una columna `GradoClasificacion` y cómo se rellena (ERD de origen).
- Quién escucha el evento `Eliminar`.
- Permisos para abrir `Form0BDGestorEntidades` (ver su ficha).
