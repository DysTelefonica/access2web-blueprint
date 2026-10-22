# Form_frmVerLogCambios

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frmVerLogCambios.cls` (187 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Visor de la auditoría de cambios (`tbLogCambios`) con filtros por usuario, tabla y registro, y acceso al detalle. Es **no enlazado**: la lista se construye como `Value List` en memoria. Las citas son `Form_frmVerLogCambios.cls:línea` salvo indicación; `FU.bas` abrevia `modules/FUNCIONES UTILES.bas` y `RC.bas`, `modules/RepositorioComun.bas`.

## Entrada y salida

- **Se abre desde:** Form_frm0OtrosAdmin, botón `cmdVerLogCambios` (`Form_frm0OtrosAdmin.cls:35-38`).
- **Abre:** Form_frmDetalleLogCambio (`:128`).
- **Permisos:** sin comprobación de rol propia; depende del acceso al menú de administración.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` | Rellena los combos y aplica filtros | `:10-32`, `.form.txt:209` |
| `cboFiltroUsuario` | combo | `AfterUpdate` | «(Todos)» + `DISTINCT usuario`; refiltra | `:16-18, 138-140`, `.form.txt:608-629` |
| `cboFiltroTabla` | combo | `AfterUpdate` | «(Todos)» + `DISTINCT tabla`; refiltra | `:20-22, 142-144`, `.form.txt:680-702` |
| `txtFiltroRegistro` (etiqueta «Texto») | texto | `KeyDown` | Con Intro refiltra | `:146-159`, `.form.txt:734, 764-766` |
| `cmdLimpiarFiltros` | botón | `OnClick` | Restablece filtros y refiltra | `:161-175`, `.form.txt:790-792` |
| `lstResultados` | lista, 9 columnas, ID oculto | `OnDblClick` | Abre el detalle | `:98-100`, `.form.txt:987-1009` |
| `cmdVerDetalle` | botón | `OnClick` | Abre el detalle | `:102-104`, `.form.txt:1124-1126` |
| `cmdSalir` | botón | `OnClick` | Cierra el formulario | `:177-185`, `.form.txt:1048-1050` |
| `lblBloqueo` | etiqueta | — | Visible mientras se filtra | `:44, 84`, `.form.txt:536-548` |

## Flujos

1. **Normal:** al cargar, combos a «(Todos)» (`:16-22`) y `AplicarFiltros` (`:24`): muestra el aviso y el reloj (`:44-46`), traduce «(Todos)» a vacío (`:54-55`), llama a `LogServicio.getLogs` (`:58`) y compone la lista con cabecera «ID;Fecha y Hora;Usuario;Tabla;ID Registro;Campo;Valor Anterior;Valor Nuevo;Operación» y fecha `dd/mm/yyyy hh:nn:ss` (`:60-80`).
2. **Filtrado:** `getLogsSQL` añade `usuario = [p_Usuario]`, `tabla = [p_Tabla]` y `registro = [p_RegistroID]` solo si el registro es numérico; orden `fechaHora DESC` (`LogServicio.cls:165-177`); parámetros en `LogServicio.cls:100-105`.
3. **Detalle:** toma `Column(0)`; si es ≤ 0, «Por favor, seleccione un registro de la lista para ver su detalle.» (`:112-117`); si el detalle ya está abierto con el mismo identificador le da el foco, si no lo cierra y abre el nuevo (`:118-128`).
4. **Errores:** `ManejadorErroresFormulario` oculta el aviso, muestra el mensaje y registra (`:91-93`; `FU.bas:609-612, 658, 672`); en la carga cierra el formulario (`:28-31`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Valores de los combos | `MSysObjects` ∪ `tbLogCambios` | SQL dinámico como `RowSource` | `:17, 21` |
| Lista filtrada | `tbLogCambios` | DAO `QueryDef` parametrizado, snapshot | `LogServicio.cls:169-176`, `RC.bas:52-60, 80-96` |

## Coexistencia

- Los combos unen la tabla de sistema local `MSysObjects` con `tbLogCambios` (`:17, 21`): consulta heterogénea que Access resuelve localmente; no existe equivalente en PostgreSQL.
- La lista carga todas las filas que cumplen el filtro, sin límite (`RC.bas:84-96`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** el descrito en Flujos.

**Posibles defectos (a verificar):**

- Los combos declaran 2 columnas con la primera de ancho 0 (`.form.txt:614, 629`), pero la consulta devuelve una sola (`:17`): podría no verse ningún valor.
- La etiqueta dice «Texto», pero un valor no numérico se descarta sin aviso (`LogServicio.cls:103, 174`).
- Un `;` dentro de `valorAnterior` o `valorNuevo` desplaza las columnas de la lista (`:66-75`); la `Value List` tiene límite de longitud y podría truncar historiales grandes.
- El comentario anuncia «modo diálogo», pero `OpenForm` no lo usa (`:127-128`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: paginación en servidor; filtro por rango de fechas; búsqueda textual en valores.

## Criterios de aceptación (propuestos, no aprobados)

1. Al abrir, se listan todos los cambios por fecha descendente con las ocho columnas visibles.
2. Cada filtro restringe por igualdad y «Limpiar» restablece la lista completa.
3. Doble clic o «Ver detalle» abren el detalle de la fila; sin selección se muestra el aviso.
4. Valores con `;` se muestran íntegros en su columna.

## Pendientes

- Volumen real de `tbLogCambios` y quién puede llegar a Form_frm0OtrosAdmin.
