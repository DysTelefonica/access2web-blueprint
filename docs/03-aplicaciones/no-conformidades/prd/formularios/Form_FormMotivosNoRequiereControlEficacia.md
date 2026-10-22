# Form_FormMotivosNoRequiereControlEficacia

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormMotivosNoRequiereControlEficacia.cls` (28 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ventana auxiliar para capturar o consultar el texto libre «motivo por el que la NC no requiere control de eficacia». La abren las pantallas de NC de auditoría (`Form_FormNCAuditoriaGeneral.cls:553`) y de NC de proyecto (`Form_FormNCProyectoGeneral.cls:1174`), que le pasan el motivo actual como argumento de apertura y reciben el nuevo mediante un evento. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) y no lee ni escribe tablas: el único almacén es el objeto de NC activo del formulario que lo abre. Las citas son `Form_FormMotivosNoRequiereControlEficacia.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_FormNCAuditoriaGeneral | `DoCmd.OpenForm` con `OpenArgs` = motivo actual del objeto de NC activo | `Form_FormNCAuditoriaGeneral.cls:553` |
| Form_FormNCProyectoGeneral | Ídem con el objeto de NC de proyecto activo | `Form_FormNCProyectoGeneral.cls:1174` |

Entrada: `OpenArgs` opcional (`:11-15`). Salida: evento público `MotivoRegistrado(m_Motivo As String)` (`:6`), disparado por «Registrar» (`:20`) y recogido por `m_FormMotivos_MotivoRegistrado` en los dos llamantes (`Form_FormNCAuditoriaGeneral.cls:410-418`, `Form_FormNCProyectoGeneral.cls:792-801`). No abre otros formularios.

## Controles y eventos

`.form.txt` declara 19 bloques `Begin`: el formulario (`.form.txt:5`), 11 bloques de valores por defecto (`.form.txt:409-480`), 3 secciones (`EncabezadoDelFormulario` `.form.txt:485`, `Detalle` `.form.txt:514`, `PieDelFormulario` `.form.txt:558`) y los 4 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` («MOTIVOS NO CE») | etiqueta | — | Título; su texto se copia a la barra de título | `.form.txt:495-504`, `:10` |
| `MotivoNoRequiereControlEficacia` | cuadro de texto, sin origen de datos | — | Recoge el texto del motivo; se precarga con `OpenArgs` | `.form.txt:522-536`, `:15` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` | Dispara `MotivoRegistrado` con el texto (`Nz` a cadena vacía) y cierra sin guardar | `:18-23`, `.form.txt:672-684` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar y sin disparar el evento | `:25-28`, `.form.txt:568-579` |
| Formulario | — | `OnLoad` | Fija el título de ventana y precarga el motivo | `:8-16`, `.form.txt:223` |

## Flujos

1. **Alta o consulta:** el llamante abre el formulario con el motivo vigente; `Form_Load` lo vuelca en el cuadro de texto si no está vacío (`:11-15`).
2. **Registrar:** dispara el evento y cierra (`:20-22`). El llamante guarda el texto en su objeto de NC, cambia el rótulo de su botón a «Ver Motivos No CE» o «Meter Motivos No CE» y muestra u oculta `ImagenMotivosNoCE` según haya texto (`Form_FormNCAuditoriaGeneral.cls:411-418`).
3. **Salir:** cierra descartando lo escrito (`:27`).
4. **Errores:** los tres handlers usan `On Error Resume Next` (`:9, 19, 26`); no hay mensajes ni aviso al administrador.

## Reglas, validaciones y cálculos

- Sin validación: se acepta cualquier texto, incluido vacío; un motivo vacío limpia el motivo en el llamante (`:20`; `Form_FormNCAuditoriaGeneral.cls:411-417`).
- No limita la longitud del texto en el código del formulario; el límite real lo marca el campo del objeto de NC (a verificar).
- El formulario no conoce el rol ni la NC; el control de edición lo ejerce el llamante.

## Lecturas y escrituras

Ninguna con base de datos. Únicamente el llamante asigna `MotivoNoRequiereControlEficacia` a su objeto de NC en memoria (`Form_FormNCAuditoriaGeneral.cls:411`); su persistencia depende de la ficha de esa pantalla. Sin transacciones.

## Coexistencia

Sin SQL ni DAO propios, no requiere adaptación a PostgreSQL. Lo único a trasladar es el campo de texto del motivo en la tabla de NC, cuyo tipo y longitud se resuelven en `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El comentario del llamante habla de «modal dialog» y de leer el valor tras cerrar (`Form_FormNCAuditoriaGeneral.cls:554-556`), pero `OpenForm` no usa `acDialog` (`Form_FormNCAuditoriaGeneral.cls:553`): el llamante sigue ejecutándose y depende del evento.
- Si el formulario no está abierto en `Form_Load`, `Resume Next` oculta el fallo y la suscripción `m_FormMotivos` puede no establecerse (`Form_FormNCAuditoriaGeneral.cls:558-561`; `:9`).
- `Resume Next` en `ComandoRegistrar_Click` puede cerrar la ventana aunque el evento falle y se pierda el texto (`:19-22`).
- La propiedad `Caption` del `.form.txt` dice «Edición de Suministrador» (`.form.txt:42`), residuo de otro formulario; la sobrescribe `Form_Load` (`:10`).
- Salir con texto sin registrar lo descarta sin confirmación (`:27`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: sustituir la ventana por un campo o diálogo modal integrado en la ficha de NC, sin evento intermedio.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir desde una NC con motivo, el cuadro muestra ese motivo (`:11-15`).
2. «Registrar» devuelve el texto al llamante y cierra la ventana (`:20-22`).
3. «Salir» cierra sin modificar el motivo de la NC (`:27`).
4. Registrar un texto vacío deja la NC sin motivo y el rótulo vuelve a «Meter Motivos No CE» (`Form_FormNCAuditoriaGeneral.cls:415-417`).
5. La ventana se titula «MOTIVOS NO CE» (`:10`).

## Pendientes

- Tipo y longitud del campo de motivo en la tabla de NC: pendiente de `../erd-origen.md`.
- Persistencia efectiva del motivo en el llamante: ver fichas de `Form_FormNCAuditoriaGeneral` y `Form_FormNCProyectoGeneral`.
