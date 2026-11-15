# Form_FormNCAuditoriaReplanificaciones

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCAuditoriaReplanificaciones.cls` (177 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Consulta de solo lectura de las replanificaciones de tareas de la no conformidad (NC) de auditoría activa: una lista con todas las replanificaciones y, al seleccionar una, la acción correctiva, la tarea y el motivo. Es **no enlazado** (sin `RecordSource`) y trabaja sobre `m_ObjNCAuditoriaActiva` (`VG.bas:127`). Las citas son `Form_FormNCAuditoriaReplanificaciones.cls:línea` salvo indicación; `VG.bas` abrevia `modules/Variables Globales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_FormNCAuditoria | Botón de navegación con destino `FormNCAuditoriaReplanificaciones` | `Form_FormNCAuditoria.form.txt:1251` |

No abre otros formularios ni tiene botón de cierre.

## Controles y eventos

`.form.txt` declara 18 bloques `Begin`: el formulario (`.form.txt:5`), 8 de valores por defecto (`:141-278`), la sección `Detalle` (`:299`) y 8 controles. Evento `OnLoad` (`.form.txt:83`).

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Etiqueta «Replanificaciones de Tareas» | etiqueta | — | Título | `.form.txt:323` |
| `ListaReplanificaciones` | lista de valores, 7 columnas | `OnClick` | Selecciona una replanificación y muestra su detalle | `:115-177`; `.form.txt:342-362` |
| `AC` («Acción correctiva») | cuadro de texto bloqueado | — | Texto de la acción correctiva | `:151-153`; `.form.txt:382-395, 453` |
| `AR` («Tarea») | cuadro de texto bloqueado | — | Texto de la acción realizada | `:146-148`; `.form.txt:412-425, 484` |
| `Observaciones` («Motivo») | cuadro de texto bloqueado | — | Observaciones de la replanificación | `:156-160`; `.form.txt:504-517, 545` |
| Etiquetas de `AC`, `AR` y `Observaciones` | 3 etiquetas | — | Rótulos | `.form.txt:453, 484, 545` |

## Flujos

### Carga (`Form_Load` → `EstablecerDatos` → `EstablecerLista`, `:11-113`)
1. Exige NC activa: «El formulario se ha abierto con parámetros insuficentes» (`:48-51`).
2. Vacía `AC` y `AR` (`:80-81`) y fija las columnas de la lista: `ID;NºAC;NºAR;Tarea;F.Replanif.;F.antes;F.después` (`:84`; `ColumnCount =7`, `.form.txt:348`; la columna 0 oculta, `.form.txt:360`).
3. Recorre `NCAuditoria.Replanificaciones` (`:87-101`) y añade por fila: identificador, número de acción correctiva, número de tarea, tarea (con «;» cambiado por «:»), fecha de replanificación, fecha prevista inicial y fecha replanificada (`:95-97`).
4. Con solo la fila de cabecera, no selecciona nada (`:102-104`); si hay datos selecciona la primera y ejecuta el clic (`:106-107`).

### Selección (`ListaReplanificaciones_Click`, `:115-177`)
Vacía `AC` y `AR` (`:122-123`); si no hay identificador, descarta la selección (`:124-131`); si cambia, recarga la replanificación con `constructor.getReplanificacionAuditoria` (`:132-144`) y muestra `AR.AccionRealizada`, `AC.AccionCorrectiva` y `Observaciones` (nulo si vacíos) (`:145-160`).

### Errores
Error 1000 ⇒ aviso con `m_Error`; otro ⇒ número y descripción, `CorreoAlAdministrador` y `MsgBox` crítico (`:30-36, 169-176`).

## Reglas, validaciones y cálculos

- Solo lectura: los tres cuadros están bloqueados (`.form.txt:382, 412, 504`) y no hay escritura ni control de rol.
- Las filas se ordenan según el orden del diccionario de replanificaciones (`:92-101`); no hay ordenación propia.
- La fila ya seleccionada no se recarga si el identificador coincide (`:138`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Evidencia |
|---|---|---|---|
| Replanificaciones de la NC | `TbReplanificacionesAuditoria` por `IDNoConformidad` | SQL dinámico concatenado (vía `NCAuditoria.Replanificaciones`) | `constructor.bas:3053-3055`; `NCAuditoria.cls:1139-1161` |
| Una replanificación | `TbReplanificacionesAuditoria` por `IDReplanificacion` | SQL dinámico concatenado | `constructor.bas:2943-2945` |

Solo lecturas; sin transacciones.

## Coexistencia

Consultas `SELECT *` por clave numérica sin dialecto Jet propio (`constructor.bas:2943-2945, 3053-3055`); válidas con PostgreSQL vinculado. Las fechas se concatenan como texto en la lista (`:96-97`): el formato depende de la configuración regional. Pendiente de la coexistencia (documento aún no redactado).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `.AR.AC.NAccion` y `.AR.NAccion` se leen sin comprobar que `AR` exista (`:95`), mientras que el clic sí lo comprueba (`:146`).
- Con una sola fila de datos, `ListCount = 1` equivale a cabecera: no se distingue de una lista vacía y la primera fila queda sin seleccionar (`:102-106`).
- `EstablecerDatos` asigna `m_Error` y no `p_Error` (`:49`).
- `m_RepSeleccionada` es pública y persiste entre cargas (`:5`), pero se reutiliza solo comparando identificadores (`:138`).
- La lista se rellena con texto separado por «;»: un valor con ese carácter distinto de la tarea rompería la fila (`:95-97`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: tabla ordenable por fecha; mostrar acción correctiva y tarea en la propia fila.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin NC activa informa de parámetros insuficientes (`:48-51`).
2. La lista muestra una fila por replanificación con las 6 columnas visibles (`:84-97`).
3. Al seleccionar una fila se muestran acción correctiva, tarea y motivo; los vacíos aparecen vacíos (`:145-160`).
4. Ninguno de los tres campos es editable (`.form.txt:382, 412, 504`).
5. Una NC sin replanificaciones muestra la lista vacía sin error (`:88-90`).

## Pendientes

- Estructura de `ReplanificacionesAuditoria` (`IDReplanificacion`, `AR`, `AC`, fechas): no revisada.
- Cómo se cierra el formulario (no hay botón; depende de la navegación del formulario padre).
