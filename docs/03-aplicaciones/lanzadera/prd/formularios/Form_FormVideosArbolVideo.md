# Form_FormVideosArbolVideo

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormVideosArbolVideo.cls` (58 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Contenedor del panel derecho de la consulta de vídeos: una cabecera con una sola opción de menú, «Video», y el subformulario `FormDetalle`, que aloja [Form_FormVideoDetalle](Form_FormVideoDetalle.md). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); no lee ni escribe datos. Las citas son `Form_FormVideosArbolVideo.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormVideosUsuarios` | Subformulario de diseño `FormDetalle` y reasignado en `Arbol_NodeClick` | `Form_FormVideosUsuarios.form.txt:2823`, `Form_FormVideosUsuarios.cls:55-57` |

Aloja `FormVideoDetalle` en `FormDetalle` (`.form.txt:415-416`; `:25-27`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `MENU1` | rectángulo | `OnClick` → `MENU1_Click` | Muestra `IND1` y carga `FormVideoDetalle` si no está cargado | `:20-29`, `.form.txt:321-323` |
| `MENU1` | rectángulo | `OnMouseMove` → `MENU1_Mousemove` | Muestra `IND1` y pasa el fondo al color de menú activado | `:36-48` |
| `lblMenu1` | etiqueta «Video» | `OnClick`, `OnMouseMove` | Delegan en `MENU1_Click` y `MENU1_Mousemove` | `:30-32, 49-52`, `.form.txt:364-367` |
| `IND1` | etiqueta indicadora (oculta en diseño) | — | Marca la opción activa | `.form.txt:334, 343` |
| `EncabezadoDelFormulario` | sección | `OnMouseMove` | Oculta `IND1` y devuelve `MENU1` al color desactivado | `:5-16`, `.form.txt:302-303` |
| `FormDetalle` | subformulario | — | Aloja `FormVideoDetalle` | `.form.txt:415-416` |
| `Imagen1` | imagen | — | Decorativa, sin evento | `.form.txt:383` |

## Flujos

- **Normal:** el formulario no tiene `OnLoad`; el subformulario de diseño ya es `FormVideoDetalle` (`.form.txt:416`). Al pulsar «Video», `MENU1_Click` lo recarga solo si `SourceObject` difiere (`:25-27`).
- **Resaltado:** al pasar el ratón sobre `MENU1` o `lblMenu1` se activa el color (`m_ObjEntorno.ColorMenuActivado`, `:43-45`); al pasar sobre la cabecera se restablece (`:11-13`).
- **Errores:** ningún handler tiene manejo de errores (`:5-57`); un fallo de `m_ObjEntorno` mostraría el error estándar de Access.

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Solo cambia el subformulario si no es ya `FormVideoDetalle` | `MENU1_Click` | `:25` |
| Colores de menú del entorno | `MENU1_Mousemove`, cabecera | `:11-12, 43-44` |

## Lecturas y escrituras

Ninguna. El formulario solo manipula la presentación.

## Coexistencia

Sin acceso a datos. Nada que adaptar para PostgreSQL. Véase `coexistencia.md` (tanda posterior).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `IND1_Mousemove` existe (`:54-57`), pero `IND1` no declara `OnMouseMove` en el `.form.txt` (`.form.txt:334-355`): el handler no se ejecuta nunca.
- `SourceObject` se compara con `"FormVideoDetalle"` (`:25`), mientras el diseño guarda `"Form.FormVideoDetalle"` (`.form.txt:416`); si Access devuelve el valor con prefijo, la comparación siempre difiere y cada clic recarga el subformulario.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: suprimir esta capa en la web; con una sola opción de menú, el detalle puede ir directamente junto al árbol.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al seleccionar un vídeo en el árbol se ve su detalle sin pasos intermedios.
2. La opción «Video» aparece resaltada al pasar el puntero.

## Pendientes

- Valor que devuelve `SourceObject` en ejecución (con o sin prefijo `Form.`), para confirmar o descartar la recarga en cada clic (`:25`).
