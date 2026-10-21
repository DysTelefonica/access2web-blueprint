# Form_FormExpedienteModificados

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormExpedienteModificados.cls` (344 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Modificados» de la ficha del expediente: lista los modificados del contrato y permite darlos de alta, eliminarlos y abrirlos para edición en `FormModificado`. Formulario sin `RecordSource`: la lista es una lista de valores rellenada por código. Las citas son `Form_FormExpedienteModificados.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormExpediente` | Botón de navegación «Modificados» (`BotónDeNavegación69`) con `NavigationTargetName = "FormExpedienteModificados"`; no vuelca el DTO | `Form_FormExpediente.form.txt:1066` |

Abre `FormModificado` (`:41-47`) y escucha su evento `Editado` (`:5`, `:221-246`; declarado y lanzado en `Form_FormModificado.cls:6, 79`). Estado compartido: `m_ObjExpedienteDTOActivo` (expediente y caché `ColModificados`, `ExpedienteDTO.cls:27`), `m_ObjModificadoActivo` (`Variables Globales.bas:268`), `EsAdministrador` y `Form_FormExpediente.m_AbiertoParaEditar`.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaModificados` | ListBox, lista de valores, 5 columnas (ID oculto, Nº, Descripción, Firma, F.Fin), con encabezados | sin evento | Muestra los modificados del expediente | `.form.txt:590-609`, `.cls:248-281` |
| `Descripcion`, `FechaFirmaModificado`, `FechaFinModificado`, `NModificado` | TextBox independientes (fechas `Short Date`, número `General Number`) | — | Datos del modificado a dar de alta; `NModificado` se precarga con el siguiente número | `.form.txt:639, 668-669, 698-699, 1006-1007`, `.cls:330` |
| `ComandoRegistrar` (icono, ayuda «Alta», `Tag = EJECUTIVO`) | botón | `OnClick` → `ComandoRegistrar_Click` | Alta (o sobrescritura, ver defectos) de un modificado | `.cls:61-122`, `.form.txt:726-731` |
| `ComandoEliminar` (icono papelera, `Tag = EJECUTIVO`) | botón | `OnClick` → `ComandoEliminar_Click` | Borra el modificado seleccionado | `.cls:126-184`, `.form.txt:818-824` |
| `ComandoEditar` (ayuda «Alta», `Tag = EJECUTIVO`) | botón | `OnClick` → `ComandoEditar_Click` | Abre `FormModificado` con el modificado seleccionado | `.cls:9-59`, `.form.txt:1064-1070` |
| Formulario | — | `OnLoad` → `Form_Load` | Permisos y carga | `.cls:188-220`, `.form.txt:223` |

## Flujos

### Apertura (`Form_Load`, `.cls:188-220`)

1. `AllowEdits = True` solo si `Form_FormExpediente.m_AbiertoParaEditar` y `EsAdministrador = Sí`; en otro caso `False` (`:196-203`).
2. `EstablecerDatos` exige `m_ObjExpedienteDTOActivo.Expediente`; si falta, «El formulario se ha abierto con parámetros insuficientes» (`:312-315`). Habilita los controles con `Tag = "EJECUTIVO"` con la misma regla de permisos (`:316-328`).
3. `RellenarListas` → `RellenarListaModificados`: toma `ColModificados` del DTO o, si está vacía, `Expediente.Modificados` (`:258-261`); añade una fila por modificado, sustituyendo `;` por `:` en la descripción, y selecciona la fila 0 (`:264-273`).
4. Precarga `NModificado` con `Expediente.NModificadoCalculado` = máximo `NModificado` numérico + 1, o «1» si no hay modificados (`Expediente.cls:271-303`).

### Alta (`ComandoRegistrar_Click`, `.cls:61-122`)

1. Construye un `ExpedienteModificado` con descripción, fechas y número del formulario, sin `IDExpedienteModificado` (`:79-85`).
2. `ExpedienteOperaciones.RegistrarModificado` (`ExpedienteOperaciones.cls:1875-1943`) valida descripción no vacía y `NModificado` numérico (`:1892-1893`); busca por `IDExpediente` + `NModificado` (`:1899`); si no existe hace `AddNew` con `IDExpedienteModificadoCalculado` (`DameID`, máximo + 1, `ExpedienteModificado.cls:65`, `FUNCIONES UTILES.bas:2851`), y si existe hace `Edit` (`:1904-1911`). Fechas no válidas se guardan `Null` (`:1914-1923`). Llama a `RegistrarUltimoCambio` (`:1935`, `FUNCIONES UTILES.bas:189`) y devuelve «1» (`:1931`).
3. Con «1», invalida la caché, recarga la lista (`.cls:95-102`); siempre selecciona la última fila, vacía descripción y fechas y recalcula `NModificado` (`:105-110`).

### Eliminación (`ComandoEliminar_Click`, `.cls:126-184`)

1. Exige `EsAdministrador = Sí` («Operacion no autorizada», `:138-141`) y un modificado seleccionado, que obtiene de `ListaModificados.Column(0)` solo si `m_MOdificadoSeleccionado` está vacío («Selecccione un elemento de la lista», `:143-155`).
2. `EliminarModificado` ejecuta `DELETE` por `IDExpedienteModificado` y `RegistrarUltimoCambio` (`ExpedienteOperaciones.cls:1957-1963`). Sin confirmación previa.
3. Recarga la lista y `NModificado` (`.cls:165-170`).

### Edición (`ComandoEditar_Click`, `.cls:9-59`)

1. Mismas comprobaciones de permiso y selección que la eliminación (`:22-39`).
2. Copia el seleccionado a `m_ObjModificadoActivo`, cierra `FormModificado` si estaba abierto y lo abre (`:40-47`).
3. Al recibir `Editado`, invalida la caché, recarga la lista y vacía `m_MOdificadoSeleccionado` (`:221-234`).

### Errores

Los errores ajenos a 1000 se envían por correo al administrador (`CorreoAlAdministrador`) y todos se muestran con `MsgBox` crítico (`:52-58, 115-121, 177-183, 213-219, 239-245`). No hay transacción en alta ni en baja.

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Editar y eliminar exigen `EsAdministrador = Sí` en el handler; el alta solo depende de que el botón esté habilitado | `.cls` | `:22-25, 138-141, 316-328` |
| Descripción obligatoria; `NModificado` numérico | `ExpedienteOperaciones.cls` | `:1892-1893` |
| Número propuesto = máximo numérico + 1 | `Expediente.cls` | `:296-303` |
| Clave natural de alta: `IDExpediente` + `NModificado` | `ExpedienteOperaciones.cls` | `:1899` |
| Fecha no reconocida por `IsDate` → `Null` | `ExpedienteOperaciones.cls` | `:1914-1923` |

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de modificados | `TbExpedientesModificados` por `IDExpediente`, sin `ORDER BY` | SQL dinámico DAO | no | `constructor.bas:5679-5715` |
| Modificado seleccionado | `TbExpedientesModificados` por `IDExpedienteModificado` | SQL dinámico DAO | no | `constructor.bas:5615-5662` |
| Alta o sobrescritura | `TbExpedientesModificados` (`AddNew`/`Edit`) | Recordset DAO sobre SQL dinámico | no | `ExpedienteOperaciones.cls:1896-1927` |
| Baja | `DELETE * FROM TbExpedientesModificados` | SQL dinámico `Execute` | no | `ExpedienteOperaciones.cls:1959-1960` |
| Último cambio | vía `RegistrarUltimoCambio` | — | no | `ExpedienteOperaciones.cls:1935, 1963` |

## Coexistencia

Pendiente de `coexistencia.md`. Ya visible: `DELETE *` y literales concatenados son sintaxis Jet; `NModificado` se compara como texto entrecomillado (`:1899`); el ID se genera con `Max + 1` en VBA (`FUNCIONES UTILES.bas:2851`), con riesgo de colisión concurrente.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Las validaciones de `RegistrarModificado` y `EliminarModificado` lanzan `Err.Raise 1000` con texto, pero el manejador solo rellena `p_Error` si el número no es 1000 (`ExpedienteOperaciones.cls:1940-1942, 1967-1969`): el mensaje se pierde, el formulario no avisa y aun así vacía los campos (`.cls:105-110`).
- Dar de alta un número ya existente edita y sobrescribe ese modificado en lugar de rechazarlo (`ExpedienteOperaciones.cls:1899, 1909-1911`).
- `m_MOdificadoSeleccionado` se cachea en el primer Editar/Eliminar y no se refresca al cambiar la selección de la lista; Eliminar no lo vacía, así que un segundo borrado o edición actúa sobre el registro anterior (`.cls:143-155, 165-170`).
- Tras el alta selecciona `ListCount - 1`, que no es necesariamente el nuevo registro porque la consulta no ordena (`.cls:105`; `constructor.bas:5690-5692`); con lista vacía el índice sería -1.
- `RellenarListaModificados` sustituye `RowSource` por la fila de encabezados y luego `Selected(0)` con `ColumnHeads` activo; no se ha comprobado qué fila queda seleccionada (`.cls:257, 273`).
- Los errores de `RellenarListas` en `EstablecerDatos` y de `RellenarListaModificados` en `m_Form_Editado` no se recogen: se llaman sin pasar el error (`.cls:229-233, 329`).
- `ComandoEditar` y `ComandoRegistrar` comparten la ayuda «Alta» (`.form.txt:731, 1069`); `m_ModExpOp` se declara y no se usa (`.cls:10`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: confirmación antes de borrar; unicidad de `NModificado` por expediente en base de datos; orden explícito por número.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin permiso de edición o sin rol administrador, alta, edición y baja no están disponibles.
2. El número propuesto es el máximo numérico existente + 1, o 1 si no hay modificados.
3. Un alta sin descripción o con número no numérico se rechaza con mensaje visible y sin escribir.
4. Eliminar o editar actúa siempre sobre la fila seleccionada en ese momento.
5. Fechas vacías o no válidas se guardan como nulas.

## Pendientes

- Tipo de `NModificado` y de las fechas en `TbExpedientesModificados` (ERD origen).
- Comportamiento de `Selected(0)` con encabezados y de `FechaFirmaModificado` (fecha regional) al convertirse a texto.
- Contenido de `RegistrarUltimoCambio` y de `FormModificado` (su ficha).
- Origen de `EsAdministrador` (ver `Form_FormExpediente.md`).
