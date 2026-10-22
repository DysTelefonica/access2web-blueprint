# Form_FormAuditoriasGestion

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormAuditoriasGestion.cls` (741 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pantalla de gestión del catálogo de auditorías: lista las auditorías filtradas por año y tipo (interna o externa) y, sobre la fila elegida, permite cambiar fechas, gestionar anexos, eliminarla, ver sus NC o dar de alta una NC de auditoría. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); la lista es una «Lista de valores» que se rellena desde el diccionario `m_ObjEntorno.ColAuditorias`. Las citas son `Form_FormAuditoriasGestion.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas` y `AuditoriaOperaciones.cls`, `classes/AuditoriaOperaciones.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_Form0BDOpcionesAuditorias | `DoCmd.OpenForm "FormAuditoriasGestion"` | `Form_Form0BDOpcionesAuditorias.cls:103` |

Abre, tras cerrar una instancia previa si existe: `FormNCAuditoria` (`:86`), `FormAuditoriaDocumentos` (`:125`), `FormAuditoria` (`:170`) y `FormNCAuditoriaGestion` (`:420`). Deja en globales `m_ObjAuditoriaActiva` (`:66, 128, 166`) y, para el alta, una `NCAuditoria` nueva con `IDAuditoria` (`:84-85`). Se suscribe a los eventos de `FormAuditoria` (`m_Form_Alta`, `m_Form_Editada`: `:172, 604-670`) y de `FormAuditoriaDocumentos` (`m_FormAnexos_Alta`, `m_FormAnexos_Eliminado`: `:126-128, 673-716`).

## Controles y eventos

`.form.txt` declara 30 bloques `Begin`: el formulario (`.form.txt:4`), 11 bloques de valores por defecto (`.form.txt:131-203`), 3 secciones (`EncabezadoDelFormulario` `.form.txt:215`, `Detalle` `.form.txt:254`, `PieDelFormulario` `.form.txt:1083`) y los 15 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo1` («GESTIÓN DE LAS AUDITORÍAS») | etiqueta | — | Título; se copia a la barra de título | `.form.txt:235-236`, `:457` |
| `Año` (etiqueta `Etiqueta86`) | cuadro de texto | `OnKeyDown` | Con Intro o Tab aplica el filtro (el foco pasa a `ComandoLimpiarAño`) | `:15-45`, `.form.txt:915-917` |
| `ComandoLimpiarAño` | botón | `OnClick` | Vacía el año y refiltra | `:339-366`, `.form.txt:310-312` |
| `Tipo` (etiqueta `Etiqueta92`) | cuadro combinado «Interna;Externa» | `AfterUpdate` | Refiltra | `:718-741`, `.form.txt:604-608` |
| `ComandoLimpiarTipo` | botón | `OnClick` | Vacía el tipo y refiltra | `:368-395`, `.form.txt:634-636` |
| `ListaFiltrados` | lista de valores, 4 columnas (`0;2268;1278;1278`) | `OnClick`, `OnDblClick` | Elige la auditoría; el doble clic ejecuta «Cambiar fechas» si está habilitado | `:527-600`, `.form.txt:261-284` |
| `ComandoEditar` («Cambiar fechas») | botón, nace deshabilitado | `OnClick` | Abre `FormAuditoria` con la auditoría activa; técnicos no | `:145-187`, `.form.txt:495-497` |
| `ComandoEliminar` («Eliminar») | botón, nace deshabilitado | `OnClick` | Elimina la auditoría; técnicos no | `:189-263`, `.form.txt:545-547` |
| `ComandoAnexos` («Anexos») | botón, nace deshabilitado | `OnClick` | Abre `FormAuditoriaDocumentos` | `:104-143`, `.form.txt:819-821` |
| `ComandoVerNC` («Ver registros») | botón, nace deshabilitado | `OnClick` | Abre `FormNCAuditoriaGestion` filtrada por la auditoría | `:397-442`, `.form.txt:944-946` |
| `ComandoAltaNCAuditoria` («Alta NC») | botón, nace deshabilitado | `OnClick` | Abre `FormNCAuditoria` con una NC nueva; técnicos no | `:54-102`, `.form.txt:1043-1045` |
| `ImagenAnexo` | imagen, oculta | `OnClick` | Abre el anexo si la auditoría tiene exactamente uno | `:473-525`, `.form.txt:989-990` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | `:47-52`, `.form.txt:1102-1104` |
| Formulario | — | `OnLoad` | Ajusta tamaño, filtra y fija el título | `:444-471`, `.form.txt:82` |

## Flujos

### Apertura y filtrado (`Form_Load` `:444-471`; `Filtrar` `:267-335`)

1. `Ajustar Me` (`:452`; `FG.bas:615`) y `Filtrar`.
2. `Filtrar` borra la selección y las globales (`:276-277`), deshabilita los cinco botones de fila y oculta `ImagenAnexo` (`:280-285`), reinicia el `RowSource` con la cabecera `ID;Nombre;FInicial;FFinal` (`:287`) y recorre `ColAuditorias` (`:291, 302`).
3. Excluye las que no coinciden con el año (`Year(FechaInicio)` frente a `Trim(Año)`, `:305-309`) o con el tipo (`:310-314`); añade `ID;Nombre;FechaInicio;FechaFin` (`:316`).
4. Con alguna fila, selecciona la primera y ejecuta `ListaFiltrados_Click` (`:323-327`).

### Selección (`ListaFiltrados_Click` `:527-591`)

Recarga la auditoría con `constructor.getAuditoria` si cambia el identificador de la columna 0 (`:551-563`), muestra `ImagenAnexo` si tiene documentos (`:570-572`) y habilita los cinco botones (`:573-577`). Sin identificador, deja la selección vacía (`:544-546`).

### Acciones sobre la fila

- **Cambiar fechas:** abre `FormAuditoria`; al volver `Alta` o `Editada` refiltra y, si está abierto `FormNCAuditoriaGestion`, vacía la caché y recarga su combo (`:604-670`; `FG.bas:3137-3168`).
- **Anexos:** abre `FormAuditoriaDocumentos`; sus eventos de alta o baja relanzan la selección (`:673-716`).
- **Ver registros:** cierra y abre `FormNCAuditoriaGestion`, pone `Estado` a `Null` y `Auditoria` al identificador, y llama a `ActualizarLista` (`:417-424`).
- **Alta NC:** crea `NCAuditoria` con la auditoría activa y abre `FormNCAuditoria` (`:84-86`).
- **Eliminar:** tras confirmar «¿Desea eliminar definitivamente del sistema la auditoría seleccionada?» (por defecto «No», `:218-220`), llama a `AuditoriaOperaciones.Eliminar` (`:227-235`), invalida la caché (`:238`, `AuditoriaOperaciones.cls:178-179`), recarga el combo de `FormNCAuditoriaGestion` si está abierto (`:237-245`) y refiltra (`:246`).
- **Anexo (imagen):** con un solo documento lo abre en local con `AbrirEnLocal`; con varios no hace nada (`:498-508`).

### Errores

Error 1000 ⇒ `MsgBox` de advertencia; otro ⇒ mensaje con número y detalle, `CorreoAlAdministrador` y `MsgBox` crítico. Patrón en todos los handlers (p. ej. `:66-77`).

## Reglas, validaciones y cálculos

- **Rol:** «Cambiar fechas», «Eliminar» y «Alta NC» devuelven «No tiene autorización para esa acción» a los técnicos (`:63, 152, 199`). «Anexos» y «Ver registros» no comprueban rol.
- **Eliminación** (`AuditoriaOperaciones.cls:138-175`): rechazada si la auditoría tiene NC asociadas («Al menos hay una no conformidad de auditoría asociada»); o si algún anexo existente en disco está abierto; borra primero los anexos (`DocumentoAuditoriaOperaciones.Eliminar`) y después la fila.
- **Filtro de año:** comparación textual con el año de `FechaInicio`; no se valida que `Año` sea numérico (`:306`).
- **Orden de la lista:** el del diccionario, por identificador descendente (`constructor.bas:4861-4862`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Lista de auditorías (caché del entorno) | `TbAuditorias` | SQL dinámico concatenado, DAO `OpenRecordset` | `constructor.bas:4861-4863` |
| Auditoría seleccionada | `TbAuditorias` por `IDAuditoria` | SQL dinámico concatenado | `constructor.bas:4815-4819` |
| Borrado de la auditoría | `TbAuditorias` | SQL dinámico concatenado, `DELETE *` por `Execute` | `AuditoriaOperaciones.cls:173-175` |
| Borrado de anexos (fichero y fila) | tabla de documentos | delegado en `DocumentoAuditoriaOperaciones.Eliminar` | `AuditoriaOperaciones.cls:162-172` (detalle: ver ficha de `FormAuditoriaDocumentos`) |
| Registro de errores no controlados | `TbCorreosEnviados` | DAO `AddNew/Update` | `FG.bas:1291-1305` |

Sin transacciones: anexos y fila se borran en pasos separados.

## Coexistencia

- `DELETE * FROM ...` (`AuditoriaOperaciones.cls:173`) es sintaxis Jet/ACE; en PostgreSQL es `DELETE FROM`.
- Identificadores concatenados en el SQL (`constructor.bas:4817`; `AuditoriaOperaciones.cls:174`): usar parámetros.
- Las fechas se pasan a la lista por conversión implícita de `Date` a cadena (`:316`), dependiente de la configuración regional; con PostgreSQL conviene tipo `date` y formato explícito.
- Los anexos tienen soporte en el sistema de ficheros (`AuditoriaOperaciones.cls:151`); la eliminación mezcla ficheros y filas sin transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Err.Raise 100` en `ComandoEliminar_Click` (`:232`) no coincide con el 1000 de los avisos: un rechazo de `Eliminar` (p. ej., NC asociadas) se trataría como error inesperado, con aviso crítico y correo al administrador.
- Los anexos se borran antes que la fila de auditoría y sin transacción: un fallo posterior deja la auditoría sin documentos (`AuditoriaOperaciones.cls:162-175`).
- `m_Form_Alta` está suscrito a `FormAuditoria`, pero este formulario solo abre `FormAuditoria` desde «Cambiar fechas» (`:170-173`): no hay botón de alta de auditoría en esta pantalla.
- `Año` se compara como texto: «2025 » o valores no numéricos no producen error, solo lista vacía (`:305-306`).
- Tras `Filtrar`, `ListaFiltrados_Click` selecciona siempre la primera fila (`:326-327`), cambiando la auditoría activa sin acción del usuario.
- El cuadro de diálogo de confirmación usa el título «Eliminación/Deseliminacion» (`:219`).
- Las etiquetas de `ComandoLimpiarAño` y `ComandoLimpiarTipo` tienen rótulo «Comando36» (`.form.txt:311, 635`); probablemente se muestran con imagen.
- `.form.txt:39` conserva el título «GESTIÓN DE LAS NO CONFORMIDADES / OBSERVACIONES»; `Form_Load` lo sustituye (`:457`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtro de año numérico; eliminación en una sola transacción con baja lógica; una tabla con ordenación y paginación en lugar de lista de valores.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir se listan todas las auditorías, de mayor a menor identificador, con la primera seleccionada (`:287-327`).
2. Filtrar por año y tipo reduce la lista; los botones «Limpiar» restauran cada filtro (`:305-314, 339-395`).
3. Sin fila elegida los botones de fila están deshabilitados (`:280-285`).
4. Un técnico recibe «No tiene autorización para esa acción» en «Cambiar fechas», «Eliminar» y «Alta NC» (`:63, 152, 199`).
5. No se puede eliminar una auditoría con NC asociadas; el mensaje lo indica y la auditoría permanece (`AuditoriaOperaciones.cls:139-141`).
6. «Eliminar» exige confirmación, con «No» por defecto (`:218-220`).
7. «Ver registros» abre las NC de la auditoría elegida (`:420-424`).
8. «Alta NC» abre `FormNCAuditoria` con una NC nueva de esa auditoría (`:84-86`).

## Pendientes

- Definición de `EsTecnico` (`:63, 152, 199`): no localizada en `src/modules`; confirmar su origen.
- Procedimiento `ActualizarLista` de `Form_FormNCAuditoriaGestion` (`:424`) y comportamiento de `FormAuditoria` y `FormAuditoriaDocumentos`: ver sus fichas.
- Efecto exacto de `DocumentoAuditoriaOperaciones.Eliminar` sobre el fichero del anexo (`AuditoriaOperaciones.cls:166`).
- Imagen de los botones «Limpiar» (`.form.txt:311, 635`).
