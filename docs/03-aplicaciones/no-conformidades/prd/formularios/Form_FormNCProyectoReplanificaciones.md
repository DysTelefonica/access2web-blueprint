# Form_FormNCProyectoReplanificaciones

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCProyectoReplanificaciones.cls` (181 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña de solo consulta de la ficha de una NC de proyecto: lista las replanificaciones de fechas de las tareas (acciones realizadas) de la NC y, al seleccionar una, muestra la acción correctiva, la tarea y el motivo. Es un formulario **no enlazado** (sin `RecordSource` ni `ControlSource`); la lista es de tipo «Value List» alimentada por código. No escribe datos. Las citas son `Form_FormNCProyectoReplanificaciones.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Ficha de NC de proyecto (`FormNCProyecto`) | Destino de navegación del control de pestañas | `Form_FormNCProyecto.form.txt:1249` |

No abre otros formularios. Usa el global `m_ObjNCProyectoActiva` (`Variables Globales.bas:117`) y expone `m_RepSeleccionada` (`:5`).

## Controles y eventos

`.form.txt` declara 18 bloques `Begin`: el formulario (`:5`), 8 bloques de valores por defecto (`:141-278`), la sección `Detalle` (`:299`) y los 8 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Etiqueta50` («Replanificaciones de tareas») | etiqueta | — | Título de la lista | `.form.txt:311-323` |
| `ListaReplanificaciones` | lista de valores, 7 columnas, con cabecera, columna 0 oculta | `OnClick` | Una fila por replanificación; al pulsar carga el detalle | `:69-117, 119-181`, `.form.txt:342-362` |
| `AC` | cuadro de texto bloqueado | — | Acción correctiva de la replanificación seleccionada | `:155-159`, `.form.txt:381-395` |
| `AR` | cuadro de texto bloqueado | — | Acción realizada (tarea) | `:150-154`, `.form.txt:411-425` |
| `Observaciones` | cuadro de texto bloqueado | — | Motivo de la replanificación | `:160-164`, `.form.txt:503-517` |
| `Etiqueta70` («Acción correctiva»), `Etiqueta71` («Tarea»), `Etiqueta73` («Motivo») | etiquetas | — | Rótulos | `.form.txt:441-453, 472-484, 533-545` |

Columnas de la lista: `ID` (oculta), `NºAC`, `NºAR`, `Tarea`, `F.Replanif.`, `F.antes`, `F.después` (`:84`, `.form.txt:360`).

## Flujos

### Carga (`Form_Load` `:11-37`; `EstablecerDatos` `:38-65`; `EstablecerLista` `:69-117`)

1. Sin `m_ObjNCProyectoActiva`: «El formulario se ha abierto con parámetros insuficentes» (`:48-51`).
2. Vacía `AC` y `AR`, fija la cabecera de la lista como primera fila (`:80-84`).
3. Toma la NC asociada si existe, la activa si no (`:86-90`) y recorre su colección de replanificaciones (`:91-105`); sin colección, deja solo la cabecera (`:92-94`). Cada fila: id, nº de AC, nº de AR, acción realizada (con `;` sustituido por `:`), fecha de replanificación, fecha prevista al inicio y fecha prevista replanificada (`:99-101`).
4. Con una sola fila (la cabecera) termina; si no, selecciona la primera y ejecuta el clic (`:106-111`).

### Selección (`ListaReplanificaciones_Click`, `:119-181`)

Limpia `AC` y `AR`; si la columna 0 está vacía descarta la selección (`:126-135`). Si cambia la replanificación seleccionada la lee por id (`:136-148`) y muestra la acción realizada, la acción correctiva y las observaciones, o vacío si no existen (`:149-165`).

### Errores

Error 1000 ⇒ aviso con `m_Error`; otro ⇒ `CorreoAlAdministrador` y mensaje crítico (`:30-36, 174-180`).

## Reglas, validaciones y cálculos

- No hay validación de rol: el formulario es de solo lectura (`AC`, `AR`, `Observaciones` bloqueados: `.form.txt:382, 412, 504`).
- Las replanificaciones mostradas son las de la NC asociada cuando la hay (`:86-91`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Evidencia |
|---|---|---|---|
| Replanificaciones de la NC | `TbReplanificacionesProyecto` (`WHERE IDNoConformidad`) | SQL dinámico concatenado, `OpenRecordset` | `constructor.bas:2973-2992` (vía `NCProyecto.Replanificaciones`, `NCProyecto.cls:677-699`) |
| Detalle de una replanificación | `TbReplanificacionesProyecto` (`WHERE IDReplanificacion`) | SQL dinámico concatenado, `OpenRecordset` | `constructor.bas:2878-2905` |
| Acciones realizadas y correctivas asociadas | a través de las propiedades `AR` y `AC` del objeto | no verificado | `:99-100, 150-159` |

Sin escrituras ni transacciones.

## Coexistencia

Las consultas concatenan identificadores numéricos (`constructor.bas:2891-2893`); en PostgreSQL deben parametrizarse. Las fechas se concatenan como texto en las filas de la lista (`:100-101`): en la web conviene dar formato explícito. Resto, pendiente de `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La fila usa `.AR.AC.NAccion` y `.AR.NAccion` sin comprobar que `AR` exista (`:99`), mientras que el clic sí lo comprueba (`:150`); una replanificación sin AR haría fallar la carga.
- El texto de la tarea solo sustituye `;` (`:100`); cualquier otro separador no afecta, pero la fila depende de ese carácter.
- Si hay una única replanificación, el criterio `ListCount = 1` solo detecta lista vacía (`:106`); la selección de la fila `1` se hace con el índice de la primera fila de datos (`:110`).
- La lista incluye un valor de ejemplo en el diseño (`.form.txt:359`) que se sustituye en `EstablecerLista` (`:84`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: tabla ordenable por fecha; mostrar la procedencia cuando se leen las de la NC asociada.

## Criterios de aceptación

Propuestos, no aprobados:

1. Una NC con replanificaciones muestra una fila por cada una con las siete columnas.
2. Una NC sin replanificaciones muestra la lista vacía y `AC`, `AR` y `Observaciones` vacíos.
3. Al seleccionar una fila se ven acción correctiva, tarea y motivo; los campos son de solo lectura.
4. En una NC con asociada se listan las de la asociada.
5. Abrir el formulario sin NC activa muestra el aviso de parámetros insuficientes.

## Pendientes

- Estructura de `TbReplanificacionesProyecto` y propiedades `AR`/`AC` de `ReplanificacionesProyecto` (clase no leída en esta ficha).
- Orden de las filas: la colección procede de un `SELECT *` sin `ORDER BY` (`constructor.bas:2989-2991`).
