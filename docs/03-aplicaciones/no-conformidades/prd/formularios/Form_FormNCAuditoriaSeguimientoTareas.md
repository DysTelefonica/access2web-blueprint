# Form_FormNCAuditoriaSeguimientoTareas

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCAuditoriaSeguimientoTareas.cls` (363 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Tareas» del seguimiento de auditoría: lista las acciones realizadas (tareas) de NC de auditoría en estado activo, pendiente de replanificar o irregular, filtrables por responsable, estado y auditoría, y abre la NC de la tarea elegida en `FormNCAuditoria`. Se aloja en `FrmDetalle` de [Form_FormNCAuditoriaSeguimiento](Form_FormNCAuditoriaSeguimiento.md). Es un formulario **no enlazado** (sin `RecordSource`); la lista sale de las colecciones `ColSegsTareasAuditoria…` del entorno. Es gemelo de [Form_FormNCAuditoriaSeguimientoNC](Form_FormNCAuditoriaSeguimientoNC.md). Las citas son `Form_FormNCAuditoriaSeguimientoTareas.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_FormNCAuditoriaSeguimiento | Botón de navegación `NavTareas` carga este formulario en `FrmDetalle` | `Form_FormNCAuditoriaSeguimiento.form.txt:1009` |
| `PintarIndicadores` | Llama a `Filtrar` si `FrmDetalle.SourceObject` es este formulario | `FG.bas:335-336` |

Sale hacia `FormNCAuditoria`: fija `m_ObjNCAuditoriaActiva` con la NC de la tarea (`m_ARAuditoriaSeleccionada.AC.nc`), cierra la instancia abierta sin guardar y la abre (`:51-56`). Solo se cierra a sí mismo ante un error en `Form_Load` (`:133`).

## Controles y eventos

`.form.txt` declara 24 bloques `Begin`: el formulario (`:4`), 9 de valores por defecto (`:713-772`), la sección `Detalle` (`:778-780`) y los 13 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaFiltrados` | lista (valores) | `OnClick`, `OnDblClick` | Una fila por tarea (9 columnas, `:200`); el clic carga la tarea y habilita «Detalle»; el doble clic abre el detalle | `:243-303`, `.form.txt:810-815` |
| `lblListaDeResultados` | etiqueta | — | Contador «Registros filtrados (n)» | `:199, 230`, `.form.txt:844-845` |
| `ComandoVerDetalle` («Detalle») | botón | `OnClick` | Abre `FormNCAuditoria` con la NC de la tarea | `:34-68`, `.form.txt:986-988` |
| `Responsable` + `Etiqueta57` («Responsable implant.») | combinado (valores) + etiqueta | `AfterUpdate` | Filtra por responsable | `:307-333`, `.form.txt:895-899` |
| `Estado` + `Etiqueta261` («Estado tarea») | combinado (valores) + etiqueta | `AfterUpdate` | Filtra por estado (3 valores) | `:70-96, 165-182`, `.form.txt:955-958` |
| `Auditoria` + `Etiqueta86` («Auditoría») | combinado (valores) + etiqueta | `AfterUpdate` | Filtra por auditoría | `:335-361`, `.form.txt:1413-1416, 1628` |
| `ComandoLimpiarResposable`, `ComandoLimpiarEstado`, `ComandoLimpiarAuditoria` | botones | `OnClick` | Ponen a Null su filtro y refiltran | `:12-30`, `.form.txt:1036, 1219, 1446` |
| `Etiqueta237` | etiqueta | — | Decorativa | `.form.txt:787` |

## Flujos

### Carga (`Form_Load`, `:98-134`)

1. `Ajustar Me` y cabecera de la lista (`:104-109`).
2. `EstablecerCombos`: responsables (`FG.bas:3075`), estados (`:165-182`) y auditorías (`FG.bas:3137`) (`:136-163`).
3. Si el usuario conectado está en `ColUsuariosCalidad`, preselecciona su nombre (`:115-117`).
4. Preselecciona el estado de acción «Activa» y filtra (`:118-119`).

### Filtrado (`Filtrar`, `:183-241`)

Llama a `constructor.getARsDeAuditoriaBusqueda` con auditoría (`Nz(Me.Auditoria, "")`), responsable y `Estado.Column(0)` (`:203-204`), y añade una fila por tarea (`:224-226`). Con el estado «ACTIVA» muestra el estado propio de cada tarea y no el literal (`:214-218`). Después actualiza el contador, selecciona la fila 1 y dispara el clic (`:230-232`).

### Selección y detalle

`ListaFiltrados_Click` carga con `constructor.getARAuditoria` la tarea de `Column(0)` si cambia (`:262-274`) y habilita «Detalle» (`:281`). «Detalle» sin selección avisa «Seleccione un elemento de la lista» (`:41-49`).

### Errores

Error 1000 ⇒ advertencia con `m_Error`; otro ⇒ `CorreoAlAdministrador` y mensaje crítico (`:61-67`, `:126-133`).

## Reglas, validaciones y cálculos

- Estados ofrecidos: ACTIVA, PTEREPLANIFICAR e IRREGULAR de `EnumEstadoAR` (`:173-175`).
- `getARsDeAuditoriaBusqueda` elige la colección comparando el texto con «ACTIVA», «PENDIENTE DE REPLANIFICAR» y «ESTADO IRREGULAR»; en otro caso, `ColSegsTareasAuditoriaTotales`. Filtra por nombre de auditoría y de responsable (`constructor.bas:2612-2640`).
- `;` de la tarea se sustituye por `:` y los números se muestran con dos dígitos (`:224-225`).
- No hay comprobación de rol.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Listado y combos | Colecciones del entorno (`ColSegsTareasAuditoria…`, `ColUsuariosCalidad`, `ColAuditorias`) | Memoria; carga en `Entorno`, fuera de esta ficha | `constructor.bas:2612-2623`, `FG.bas:3087` |
| Cargar la tarea seleccionada | `TbNCAuditoriaAccionesRealizadas` | SQL dinámico concatenado con `p_IDAR` | `constructor.bas:5629-5631` |
| Escrituras | — | Ninguna | `:1-363` |

## Coexistencia

Revisar en `getARAuditoria` la concatenación de `p_IDAR` (tipo numérico del id al pasar a PostgreSQL vinculado). `Format(…, "00")` y `Nz` son de VBA. La comparación por textos literales de estado («ACTIVA», «PENDIENTE DE REPLANIFICAR»…) depende de los títulos del entorno.

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual**

- Filtra por auditoría y responsable sobre la colección elegida por estado (`constructor.bas:2612-2640`).
- Con el combo de estado vacío, lista todas las tareas (`constructor.bas:2619-2621`).

**Posibles defectos (a verificar)**

- La cabecera de la lista se define con 6 columnas en `Form_Load` y con 9 en `Filtrar` (`:109`, `:200`); el diseño trae una tercera variante (`.form.txt` de `ListaFiltrados`).
- `Filtrar` compara `Auditoria` contra el nombre (`m_AR.Auditoria`) pero la combo guarda id y nombre (`FG.bas:3137-3160`); depende de la columna enlazada (`:203`).
- `Estado.Column(0)` guarda el título y `getARsDeAuditoriaBusqueda` espera literales distintos de los títulos, tal vez coincidentes (`:204`, `constructor.bas:2612-2616`).
- `lst.Selected(1) = True` con la lista vacía y contador `ListCount - 1` (`:230-231`).
- Salida de `ListaFiltrados_Click` con tarea `Nothing` vuelve a activar el cursor de espera (`:275-279`).
- Mensaje de error de `ComandoVerDetalle_Click` dice «ComandoEditar_Click» (`:62`).

**Propuestas web pendientes de decisión**

- Tabla paginada con filtros en servidor y navegación directa a la NC.
- Unificar esta pestaña y la de Rev. Estados con un solo componente de lista.

## Criterios de aceptación

1. Al abrir, el responsable es el usuario si es de Calidad y el estado «Activa» (`:115-118`).
2. Cambiar cualquier filtro refiltra y actualiza el contador (`:70-96, 307-361`).
3. Las tres «X» vacían su filtro (`:12-30`).
4. «Detalle» (o doble clic) abre `FormNCAuditoria` con la NC de la tarea; sin selección avisa (`:41-56, 298-303`).
5. Error 1000 ⇒ advertencia; otro ⇒ mensaje crítico y aviso al administrador (`:61-67`).

## Pendientes

- Carga de `ColSegsTareasAuditoria…` en `Entorno` (referencia dinámica no resuelta; `constructor.bas:2612-2623`).
- `BoundColumn` y `ColumnCount` de `Auditoria` y `Estado` en el binario.
