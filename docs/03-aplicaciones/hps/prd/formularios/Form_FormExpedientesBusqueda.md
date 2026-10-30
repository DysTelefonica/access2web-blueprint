# Form_FormExpedientesBusqueda

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormExpedientesBusqueda.cls` (279 líneas) y `.form.txt` (1128 líneas) · **PRD:** [README](../README.md)

Formulario modal de búsqueda de expedientes. Lee de `TbExpedientes` (backend de expedientes vía `getdbExpedientes()`), aplica un filtro por `PalabraClave` con `Constructor.getExpedientesBusqueda` y devuelve el resultado **por evento**: `RaiseEvent Seleccionado(m_Exp As Expediente)` consume un `Expediente` y luego cierra el formulario (`ComandoElegir_Click`, `:48-83`). La rejilla es un `ListBox` con cuatro columnas (`IDExp`, `Cod_Exp`, `Nemotécnico`, `Título`), no un datasheet. Tiene un segundo comportamiento, `ComandoActualizar_Click`, que descarta los caches `m_PalabraClaveUltima` y `m_IDResponsableCalidadUltimo` (campos que **no se usan en este formulario** — ver §Comportamiento actual) y vuelve a filtrar. Las citas son `Form_FormExpedientesBusqueda.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Cualquier formulario que necesite un expediente | Lo abre con `DoCmd.OpenForm "FormExpedientesBusqueda"` y se suscribe a `Seleccionado` mientras esté abierto | ver `clsExpedienteBusquedaConsumer` (no incluido en esta ficha) |
| `Constructor.getExpedientesBusqueda(p_PalabraClave, p_Error)` | Recibe el `PalabraClave` y devuelve un `Scripting.Dictionary` (clave `IDExpediente`, valor `Expediente`) | `Constructor.bas:3340-3395` |
| `Constructor.getExpediente(p_IDExpediente, p_Error)` | Hidrata el `m_ExpSeleccionado` cuando el usuario hace click en una fila o pulsa Elegir | `:203, 209`, `Constructor.bas:568-616` |

Acepta un único elemento (no multiselección); el `RaiseEvent Seleccionado` entrega un `Expediente`, no un array, y cierra el formulario.

## Controles y eventos

`Caption ="EXPEDIENTES REGISTRADOS"` (`.form.txt:40`), `OnOpen ="[Event Procedure]"` (`.form.txt:41`). El formulario no tiene `RecordSource`: el resultado se pinta sobre un `ListBox` con `RowSourceType ="Value List"`.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnOpen` → `Form_Open` | Ajusta tamaño, fija `Caption` y columnas del `ListBox` | `:104-132`, `.form.txt:41` |
| `PalabraClave` | caja de texto | `OnExit` → `PalabraClave_Exit` | Filtra la lista por `PalabraClave` al perder foco | `:246-277`, `.form.txt:461` |
| `ComandoActualizar` | botón | `Click` → `ComandoActualizar_Click` | Resetea los `m_…Ultimo` y reejecuta `Filtrar` | `:14-45`, `.form.txt:339` |
| `ComandoLimpiarPalabraClave` | botón | `Click` → `ComandoLimpiarPalabraClave_Click` | Pone `PalabraClave = Null` y re-filtra si ya había valor | `:86-93`, `.form.txt:488` |
| `ComandoElegir` | botón | `Click` → `ComandoElegir_Click` | Dispara `RaiseEvent Seleccionado` y cierra | `:48-83`, `.form.txt:731` |
| `ComandoSalir` | botón | `Click` → `ComandoSalir_Click` | Cierra sin elegir | `:96-99`, `.form.txt:777` |
| `ListaFiltrados` | `ListBox` (`RowSourceType ="Value List"`) | `OnClick` → `ListaFiltrados_Click` | Resuelve el `Expediente` de la fila y actualiza `ComandoElegir` + `Titulo` | `:180-242`, `.form.txt:425` |
| `Titulo` | caja de texto | — | Muestra el título del expediente seleccionado; solo lectura | `:192, 223`, `.form.txt:672` |
| `lblTitulo1` | etiqueta | — | Sirve para componer el `Caption`; rótulo interno del formulario | `.form.txt:320` |
| `Etiqueta29` | etiqueta | — | «Palabra Clave» | `.form.txt:701` |

## Flujos

### Apertura (`Form_Open`, `:104-132`)

1. `AjustarTamaño Me` (`:111`).
2. `Me.Caption = Me.lblTitulo1.Caption` (`:115`): el título de la ventana replica el rótulo interno.
3. `Me.ListaFiltrados.RowSource = "IDExp;Cod_Exp;Nemotécnico;Título"` (`:116`): fija las **cuatro cabeceras de columna** del `ListBox` en formato `Value List`.
4. En error: `CorreoAlAdministrador` salvo 1000 y `MsgBox`; el formulario se cierra (`DoCmd.Close acForm, Me.Name, acSaveNo`, `:127, 131`).

### Filtrar (`Filtrar`, `:135-178`)

Función `Public` que también puede invocarse desde fuera:

1. `Me.Titulo = Null` (`:147`): cada filtro limpia el título visible.
2. Limpia las filas previas: `lst.RowSource = "IDExp;Cod_Exp;Nemotécnico;Título"` (`:150`).
3. `Set m_Col = Constructor.getExpedientesBusqueda(p_PalabraClave:=Nz(Me.PalabraClave, ""), p_Error:=p_Error)` (`:153`; `Constructor.bas:3340-3395`).
4. Si vuelve error, `Err.Raise 1000` (`:155`). Si vuelve `Nothing`, sale (`Exit Function`, `:158`).
5. Recorre `m_Col` y hace `lst.AddItem .IDExpediente & ";" & .CodExp & ";" & .Nemotecnico & ";" & Replace(.Titulo, ";", ":")` (`:163-166`): **sustituye `;` por `:`** dentro del título para no romper el separador de columnas.
6. `lst.Selected(1) = True` y dispara `ListaFiltrados_Click` (`:168-169`) para que la primera fila quede resuelta con su `Expediente` ya cargado.

### Cambio de selección (`ListaFiltrados_Click`, `:180-242`)

1. `Application.Echo False` y deshabilita `ComandoElegir` (`:190-191`).
2. `Me.Titulo = Null` (`:192`).
3. Si la columna 0 está vacía, sale (`:194-201`).
4. Si `m_ExpSeleccionado Is Nothing` (`:202`) o cambió de fila (`:208`), lo rehidrata con `Constructor.getExpediente(p_IDExpediente:=Nz(Me.ListaFiltrados.Column(0), ""), p_Error:=m_Error)` (`:203, 209`; `Constructor.bas:568-616`).
5. Si vuelve error, `Err.Raise 1000`.
6. Si vuelve `Nothing`, sale sin habilitar el botón Elegir (`:215-221`).
7. `Me.Titulo = m_ExpSeleccionado.Titulo` (`:223`) y `Me.ComandoElegir.Enabled = True` (`:224`).
8. `Application.Echo True` antes de devolver (`:230`).
9. En error: el `errores` distingue 1000 (`MsgBox vbExclamation`) de los demás (`MsgBox vbCritical` + `CorreoAlAdministrador`, `:235-241`).

### Elegir (`ComandoElegir_Click`, `:48-83`)

1. Si `m_ExpSeleccionado Is Nothing`, lo hidrata con la fila seleccionada (`:57`); si la hidratación falla, `Err.Raise 1000`.
2. Si sigue `Nothing`, eleva 1000 con «Se ha de seleccionar un elemento de la lista» (`:62-63`).
3. `RaiseEvent Seleccionado(m_ExpSeleccionado)` (`:68`): publica el evento.
4. `DoCmd.Close acForm, Me.Name, acSaveNo` (`:69`): cierra inmediatamente, sin esperar al handler.
5. En error: misma lógica de `vbExclamation`/`vbCritical` (`:75-82`).

### Cambio de filtro en la caja (`PalabraClave_Exit`, `:246-277`)

1. Si el valor no cambió respecto a `m_PalabraClaveUltima`, sale sin refiltrar (`:253-257`).
2. `Filtrar m_Error` y, si vuelve con error, `Err.Raise 1000` (`:259-261`).
3. `m_PalabraClaveUltima = Nz(Me.PalabraClave, "")` (`:263`) — guarda la última palabra aplicada.

### Botón limpiar (`ComandoLimpiarPalabraClave_Click`, `:86-93`)

`On Error Resume Next`; `Me.PalabraClave = Null` (`:88`). Si `m_PalabraClaveUltima <> ""`, llama a `Filtrar` (`:89-91`). **No reescribe** `m_PalabraClaveUltima`, así que la siguiente pérdida de foco en la caja verá «igual que la anterior» y no refiltrará hasta que se vuelva a salir de la caja tras escribir de nuevo.

### Actualizar (`ComandoActualizar_Click`, `:14-45`)

Resetea `m_PalabraClaveUltima = ""` (`:22`) y `m_IDResponsableCalidadUltimo = ""` (`:23`) antes de `Filtrar` (`:24`). `m_IDResponsableCalidadUltimo` **no se usa en este `.cls`**: queda como residuo de una versión anterior en la que se filtraba por responsable de calidad.

### Salir (`ComandoSalir_Click`, `:96-99`)

`On Error Resume Next` + `DoCmd.Close acForm, Me.Name, acSaveNo`. Si el cierre falla, el error se descarta.

## Reglas, validaciones y cálculos

- **Búsqueda con dos modos por la misma cadena.** `getExpedientesBusqueda` (`Constructor.bas:3340-3395`) carga **todos** los expedientes de `TbExpedientes` (`Constructor.bas:3351-3352`) y luego aplica el filtro en memoria sobre cada fila con `p_PalabraClave <> m_Expediente.IDExpediente And InStr(1, m_Expediente.Nemotecnico, p_PalabraClave) = 0 And InStr(1, m_Expediente.Titulo, p_PalabraClave) = 0` (`Constructor.bas:3366-3371`): **el `PalabraClave` se compara contra `IDExpediente` exacto, contra `Nemotecnico` parcial y contra `Titulo` parcial**.
- **El filtro se ejecuta en memoria, no en SQL.** No hay `WHERE` en `getExpedientesBusqueda` (`Constructor.bas:3351-3352`): todos los expedientes viajan al frontend y se descartan uno a uno. El tamaño de la tabla marca el tiempo de respuesta.
- **`Value List` con `;` como separador.** Cada fila se compone con `lst.AddItem` (`:163`) y los `;` del título se escapan con `Replace(.Titulo, ";", ":")` (`:165`).
- **El formulario publica y se va.** `ComandoElegir_Click` (`:48-83`) hace `RaiseEvent` y cierra **antes** de que el consumidor pueda responder (`:68-69`). El handler del evento debe capturar el `Expediente` y guardarlo antes del `DoCmd.Close`.
- **`m_ExpSeleccionado` se rehidrata al cambiar de fila.** `ListaFiltrados_Click` (`:202-213`) gestiona caché: si la fila cambia, descarta el `Expediente` y vuelve a llamar a `Constructor.getExpediente`. Si la fila no cambia, el `Expediente` se reutiliza entre `OnClick` y `ComandoElegir_Click`.
- **`m_PalabraClaveUltima` se reescribe solo en `PalabraClave_Exit`.** El botón «limpiar» lo deja obsoleto; el siguiente `OnExit` verá `Me.PalabraClave = Nz(Null, "") = ""` y `m_PalabraClaveUltima = <viejo>`, así que **no refiltra**.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Catálogo completo de expedientes | `TbExpedientes` (vía `getdbExpedientes()`) | `SELECT * FROM TbExpedientes` | `Constructor.bas:3351-3352` |
| Hidratación de la fila | `TbExpedientes` (vía `Constructor.getExpediente`) | `SELECT * FROM TbExpedientes WHERE IDExpediente=<valor>` | `:203, 209`, `Constructor.bas:581-583` |
| Resultado filtrado en memoria | comparación sobre cada `Expediente` | filtro en código | `Constructor.bas:3366-3371` |
| Título mostrado | `m_ExpSeleccionado.Titulo` | propiedad del objeto | `:223` |

No hay `INSERT`/`UPDATE`/`DELETE` directos en este `.cls`.

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Ya se ve: **acoplamiento por `RaiseEvent Seleccionado(m_Exp)`** (`:68`): quien abra el formulario debe suscribirse al evento `WithEvents` y guardar el `Expediente` antes de que el `DoCmd.Close` lo desmonte; **concatenación SQL** `WHERE IDExpediente=<valor>` (`Constructor.bas:581-583`) sin sanitizar — aquí siempre viene de `Nz(Me.ListaFiltrados.Column(0), "")`, lo que blinda parcialmente la entrada; **`ListBox` con `Value List`** en lugar de consulta, con la sustitución `;` → `:`; `m_PalabraClaveUltima` y `m_IDResponsableCalidadUltimo` declarados a nivel módulo (`:8-9`) son **estado mutable entre llamadas** que mezcla este formulario con su(s) consumidor(es), el segundo campo sin uso en `src/`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Filtrar` (`:135-178`) **no aplica** `PalabraClave = ""` como filtro universal — `getExpedientesBusqueda` (`Constructor.bas:3366-3371`) devuelve lo que cumpla la condición de omisión: si la palabra no aparece en ningún expediente, devuelve un diccionario vacío. No hay un aviso de «sin resultados»; la rejilla se queda con las cabeceras y `ComandoElegir.Enabled = False`.
- `getExpedientesBusqueda` (`Constructor.bas:3340-3395`) hace `SELECT *` sobre `TbExpedientes` **sin filtrar en SQL** y aplica el filtro en memoria.
- `ListaFiltrados.Selected(1) = True` (`:168`) asume que la rejilla tiene al menos un elemento tras `Filtrar`. Si el filtro devuelve colección vacía, `lst.Selected(1)` falla.
- `ComandoLimpiarPalabraClave_Click` (`:86-93`) **no** pone `m_PalabraClaveUltima = ""`, así que el siguiente `OnExit` con `PalabraClave = ""` no refiltra porque `Nz(Me.PalabraClave, "") = m_PalabraClaveUltima`.
- `m_IDResponsableCalidadUltimo` (`:9`) no se usa en este `.cls`: solo se inicializa a `""` en `ComandoActualizar_Click` (`:23`). El campo es residuo de una versión con búsqueda por responsable de calidad.
- El evento `Seleccionado` se lanza antes de cerrar el formulario (`:68-69`).
- `ComandoSalir_Click` con `On Error Resume Next` (`:97`) descarta cualquier error al cerrar.
- La concatenación SQL `WHERE IDExpediente=<valor>` (`Constructor.bas:581-583`) acepta cualquier cadena; aunque `Nz(Me.ListaFiltrados.Column(0), "")` blinda aquí, una nueva ruta que pase el `IDExpediente` por UI puede abrir una inyección.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: mover el filtro a la capa de servicio con `WHERE` parametrizado (`Nemotecnico LIKE :p OR Titulo LIKE :p OR IDExpediente = :p`); sustituir el `ListBox` con `Value List` por una tabla virtual bindada a una colección de DTOs; eliminar `m_IDResponsableCalidadUltimo` si la búsqueda por responsable ya no aplica; y un componente que entregue el `Expediente` vía `Promise` o `Observable` en lugar de `RaiseEvent`.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, la rejilla muestra las cabeceras `IDExp;Cod_Exp;Nemotécnico;Título` con el `Caption` «EXPEDIENTES REGISTRADOS».
2. Escribir una `PalabraClave` y perder foco aplica el filtro (`IDExpediente` exacto, `Nemotecnico` parcial, `Titulo` parcial).
3. `ComandoLimpiarPalabraClave` vacía la caja y, si ya había filtro, reejecuta el filtro con cadena vacía.
4. Seleccionar una fila hidrata `m_ExpSeleccionado`, muestra el título en `Titulo` y habilita `ComandoElegir`.
5. Pulsar `ComandoElegir` lanza el evento `Seleccionado(m_Exp)` y cierra el formulario.
6. `ComandoSalir` cierra sin elegir expediente.
7. Los errores que no sean 1000 disparan `CorreoAlAdministrador` y un `MsgBox` crítico; los 1000 muestran `MsgBox vbExclamation`.

## Pendientes

- Documentar el consumidor `clsExpedienteBusquedaConsumer` (no incluido en esta ficha).
- Confirmar en el binario si `m_IDResponsableCalidadUltimo` se usa en alguna ruta residual.
- Contenido de `clsExpediente.colCampos` (la base de la hidratación por `Constructor.getExpediente`).
- Si `ListaFiltrados.Selected(1) = True` cuenta desde 0 o 1 en esta versión de Access (la rama actual asume 1-indexado, `:168`).
