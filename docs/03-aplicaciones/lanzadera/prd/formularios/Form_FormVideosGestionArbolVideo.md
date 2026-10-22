# Form_FormVideosGestionArbolVideo

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormVideosGestionArbolVideo.cls` (51 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Contenedor intermedio de la edición de un vídeo: una barra de menú con una sola opción («Datos») y un subformulario `FormDetalle` que aloja [Form_FormVideosGestionArbolVideoDatos](Form_FormVideosGestionArbolVideoDatos.md). Es un formulario **no enlazado**, sin lecturas ni escrituras propias. Las citas son `Form_FormVideosGestionArbolVideo.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| [Form_FormVideosGestion](Form_FormVideosGestion.md) | Al pulsar un nodo de vídeo: `FormDetalle.SourceObject = "FormVideosGestionArbolVideo"` y `MENU1_Click` | `Form_FormVideosGestion.cls:61-63` |

Abre `FormVideosGestionArbolVideoDatos` en su `FormDetalle` y llama a su `CargarDatos` (`:18-23`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `MENU1` | rectángulo | `OnClick` → `MENU1_Click` (público) | Carga la pestaña de datos si no está ya cargada | `:18-23`, `.form.txt:317-319` |
| `lblMenu1` (rótulo «Datos») | etiqueta | `OnClick`, `OnMouseMove` | Delegan en `MENU1` | `:24-26, 41-44`, `.form.txt:357-360` |
| `MENU1`, `IND1` | rectángulo, etiqueta | `OnMouseMove` | Muestra el indicador y aplica `ColorMenuActivado` | `:29-40, 46-49`, `.form.txt:336-337` |
| `EncabezadoDelFormulario` | sección | `OnMouseMove` | Oculta el indicador y aplica `ColorMenuDesactivado` | `:5-16`, `.form.txt:299-300` |
| `FormDetalle` | subformulario | — | Sin `SourceObject` de diseño | `.form.txt:416-428` |

## Flujos

- **Normal:** el contenedor carga este formulario y llama a `MENU1_Click`; como `FormDetalle` está vacío, fija `FormVideosGestionArbolVideoDatos` y llama a su `CargarDatos` (`:19-21`).
- **Alternativo:** pulsar «Datos» con la pestaña ya cargada no hace nada; no recarga los datos (`:19`).
- **Errores:** no hay manejador de errores; `CargarDatos` se llama sin `p_Error`, así que sus errores no se muestran (`:21`; `Form_FormVideosGestionArbolVideoDatos.cls:17, 100-102`).

## Reglas, validaciones y cálculos

- Colores del menú desde `m_ObjEntorno.ColorMenuActivado` y `ColorMenuDesactivado` (`:11-12, 36-37`).
- Sin reglas de negocio ni permisos propios: hereda los del contenedor ([Form_FormVideosGestion](Form_FormVideosGestion.md)).

## Lecturas y escrituras

Ninguna.

## Coexistencia

Sin impacto propio. Ver `../coexistencia.md` cuando exista.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `CargarDatos` sin `p_Error` (`:21`): un fallo al cargar el vídeo deja el panel vacío o con datos de otro vídeo sin aviso.
- `m_Error` declarado y sin uso (`:3`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: suprimir este nivel intermedio de una sola pestaña y mostrar la edición directamente.

## Criterios de aceptación

Propuestos, no aprobados:

1. Seleccionar un vídeo en el árbol muestra sus datos sin pulsar «Datos».
2. Un fallo al cargar los datos del vídeo se comunica al usuario.

## Pendientes

- Si Access reconoce el valor localizado `[Procedimiento de evento]` en `IND1.OnMouseMove` (`.form.txt:337`) igual que `[Event Procedure]`.
- Si `Form_FormVideosGestionArbolVideoDatos.CargarDatos` (`:21`) actúa sobre la instancia del subformulario.
