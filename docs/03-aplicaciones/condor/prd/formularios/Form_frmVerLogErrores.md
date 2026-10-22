# Form_frmVerLogErrores

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frmVerLogErrores.cls` (183 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Visor del registro de errores (`tbLogErrores`) con filtros por usuario, módulo y procedimiento. Misma construcción que Form_frmVerLogCambios: no enlazado, lista `Value List` en memoria. Las citas son `Form_frmVerLogErrores.cls:línea` salvo indicación; `FU.bas` abrevia `modules/FUNCIONES UTILES.bas` y `RC.bas`, `modules/RepositorioComun.bas`.

## Entrada y salida

- **Se abre desde:** Form_frm0OtrosAdmin, botón `cmdVerLogErrores` (`Form_frm0OtrosAdmin.cls:46-49`).
- **Abre:** Form_frmDetalleLogError (`:126`).
- **Permisos:** sin comprobación de rol propia.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` | Rellena los combos y aplica filtros | `:10-35`, `.form.txt:212` |
| `cboFiltroUsuario` | combo | `AfterUpdate` | «(Todos)» + `DISTINCT usuario`; refiltra | `:17-19, 136-138`, `.form.txt:611-632` |
| `cboFiltroModulo` | combo | `AfterUpdate` | «(Todos)» + `DISTINCT modulo`; refiltra | `:22-24, 140-142`, `.form.txt:683-705` |
| `txtFiltroProcedimiento` («Proc.») | texto | `KeyDown` | Con Intro refiltra | `:144-155`, `.form.txt:737, 767-769` |
| `cmdLimpiarFiltros` | botón | `OnClick` | Restablece filtros y refiltra | `:157-171`, `.form.txt:793-795` |
| `lstResultados` | lista, 6 columnas, ID oculto | `OnDblClick` | Abre el detalle | `:97-99`, `.form.txt:978-1000` |
| `cmdVerDetalle` | botón | `OnClick` | Abre el detalle | `:101-103`, `.form.txt:1115-1117` |
| `cmdSalir` | botón | `OnClick` | Cierra el formulario | `:173-181`, `.form.txt:1039-1041` |
| `lblTitulo` | etiqueta | — | Texto fijo «Visor de Log de Cambios» | `.form.txt:569-570` |

## Flujos

1. **Normal:** combos a «(Todos)» (`:17-24`) y `AplicarFiltros` (`:27`): aviso y reloj (`:46-48`), «(Todos)» ⇒ vacío (`:56-57`), `LogServicio.getErrores` (`:60`) y lista con cabecera «ID;Fecha y Hora;Usuario;Módulo;Nº Error;Descripción» (`:62-79`).
2. **Filtrado:** igualdad por `usuario` y `modulo`; `procedimiento LIKE [p_Procedimiento]` con el texto envuelto en `*` (contiene); orden `fechaHora DESC` (`LogServicio.cls:121-126, 179-191`).
3. **Detalle:** igual que el visor de cambios, con el mismo aviso de selección (`:111-126`).
4. **Errores:** `ManejadorErroresFormulario` (`:90-92`; `FU.bas:588-676`); en la carga cierra el formulario (`:31-34`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Valores de los combos | `MSysObjects` ∪ `tbLogErrores` | SQL dinámico como `RowSource` | `:18, 23` |
| Lista filtrada | `tbLogErrores` | DAO `QueryDef` parametrizado, snapshot | `LogServicio.cls:183-190`, `RC.bas:52-60, 80-96` |

## Coexistencia

- `MSysObjects` en los combos (`:18, 23`), igual que en el visor de cambios.
- El comodín `*` de Jet (`LogServicio.cls:124`) debe conservar la semántica «contiene» frente a `tbLogErrores` vinculada (en PostgreSQL, `%`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** el descrito en Flujos.

**Posibles defectos (a verificar):**

- La lista no muestra el procedimiento aunque se filtra por él (`:62`).
- Título visible «Visor de Log de Cambios» en el visor de errores (`.form.txt:570`).
- Combos con 2 columnas declaradas y una devuelta (`.form.txt:617, 632`; `:18`).
- Un `;` en `descripcionError` desplaza columnas; límite de longitud de la `Value List` (`:68-74`).
- El manejador de `AbrirDetalle` sale con `GoTo` en lugar de `Resume` (`:131`) y deja el error activo.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: columna de procedimiento; filtro por fechas y número de error; paginación.

## Criterios de aceptación (propuestos, no aprobados)

1. Al abrir, se listan los errores por fecha descendente con el título del visor de errores.
2. El filtro de procedimiento encuentra coincidencias parciales.
3. Doble clic o «Ver detalle» abren el detalle; sin selección se muestra el aviso.

## Pendientes

- Volumen de `tbLogErrores` y política de retención.
