# Form_FormGastoGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormGastoGestion.cls` (683 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado, filtrado y gestión de gastos y dietas (`TbGastos`): filtra por técnico, tipo de gasto y rango de fechas, y desde la lista da de alta, edita, anexa, elimina y exporta a Excel. Es un formulario **no enlazado** (sin `RecordSource`); la lista `ListaFiltrados` recibe el SQL por código. Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`; `Constructor.bas` = `modules/Constructor.bas`. Las citas son `Form_FormGastoGestion.cls:línea` salvo indicación; «`.form.txt`» cita `Form_FormGastoGestion.form.txt`.

## Entrada y salida

| Desde / hacia | Cómo | Evidencia |
|---|---|---|
| Desde `Form_Form0BDOpciones` | Etiqueta de menú `lblGastosYDietas`: `DoCmd.OpenForm "FormGastoGestion"` sin argumentos | `Form_Form0BDOpciones.cls:96-99` |
| Hacia `FormGastoAlta` | Cierra la instancia abierta y la abre | `:45-48` |
| Hacia `FormGastoEdicion` | Cierra la instancia abierta y la abre con `OpenArgs` = `ID\|URLAnexo\|Sí/No` | `:118-122` |
| Hacia `FormAnexos` | Cierra la abierta, publica el gasto en `m_ObjEntidadParaAnexoActiva` y la abre | `:77-82` |
| Retorno desde `FormGastoAlta` | Llama a `Form_FormGastoGestion.ActualizarLista` | `Form_FormGastoAlta.cls:54-55` |
| Retorno desde `FormGastoEdicion` | `ListaFiltrados.Requery` (no reconstruye el SQL) | `Form_FormGastoEdicion.cls:80-81, 252-253` |

Controles: 25 (25 `Begin` en el `.form.txt` dentro de la sección Detalle; se excluyen la propia sección y las plantillas de valores por defecto, `Form_FormGastoGestion.form.txt:752-884`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ListaFiltrados` | lista, 7 columnas, con cabecera | `OnClick` → `ListaFiltrados_Click` | Resultado del filtro. Columnas visibles: ID oculta, gasto, importe, fecha, técnico; `IDFactura` y `NombreArchivo` ocultas | CAP-8 | `Form_FormGastoGestion.form.txt:954-999`; `:425-483` |
| `ComboTecnicoBusqueda` | combo (alias, nombre) | `OnChange`, `OnExit` | Refresca la lista si cambia el valor | CAP-8 | `Form_FormGastoGestion.form.txt:1188-1225`; `:246-267, 335-356` |
| `ComboGastoBusqueda` | combo de tipos de gasto distintos | `OnChange`, `OnExit` | Ídem | CAP-8 | `Form_FormGastoGestion.form.txt:1582-1621`; `:223-244, 313-334` |
| `FechaInicialBusqueda`, `FechaFinalBusqueda` | texto, formato fecha corta | `OnExit` | Refrescan la lista si cambia el valor | CAP-8 | `Form_FormGastoGestion.form.txt:1333-1394`; `:269-312` |
| `ComandoActualizarTodo` | botón | `OnClick` | Limpia técnico, fechas y gasto y recarga | CAP-8 | `:12-32` |
| `ComandoLimpiarTecnico`, `ComandoLimpiarGasto`, `ComandoLimpiarFechas` | 3 botones | `OnClick` | Vacían su filtro y recargan si tenía valor; `On Error Resume Next` | CAP-8 | `:207-221, 617-647` |
| `ComandoAlta` | botón | `OnClick` | Abre `FormGastoAlta` | CAP-8 | `:33-59` |
| `ComandoEditar` | botón | `OnClick` | Abre `FormGastoEdicion` con el gasto seleccionado | CAP-8 | `:95-133` |
| `ComandoEliminar` | botón | `OnClick` | Confirma y elimina el gasto | CAP-8 | `:134-183` |
| `ComandoAnexar` (rótulo «Adjuntar») | botón | `OnClick` | Abre `FormAnexos` sobre el gasto | CAP-11 | `:60-93` |
| `ComandoExportarExcel` | botón | `OnClick` | Exporta el SQL actual de la lista a Excel | CAP-10 | `:184-206` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar; `On Error Resume Next` | CAP-8 | `:385-388` |
| `ImagenAnexo` | imagen | `OnClick` | Abre el anexo o, si hay varios, `FormAnexos` | CAP-11 | `:389-424` |
| `lblFacturado` | etiqueta | — | Visible solo si el gasto tiene factura | CAP-8 | `:450, 454` |
| `lblTitulo1`, `lblTitulo2`, `Etiqueta210`, `Etiqueta230`, `Etiqueta231`, `Etiqueta241` | 6 etiquetas | — | Títulos y rótulos de filtros; `lblTitulo1` da el `Caption` del formulario | CAP-8 | `:369`; `Form_FormGastoGestion.form.txt:1632-1654` |
| `CuadroSuperior`, `Cuadro2` | 2 rectángulos | — | Decoración | sin asignar | `Form_FormGastoGestion.form.txt:908-953` |

## Flujos

**Apertura (`Form_Open`).** Ajusta tamaño con `AjustarTamaño` (`FG.bas:261-288`), fija el título y llama a `ActualizarLista` (`:363-370`). Con los cuatro filtros vacíos el SQL lleva `IDGasto=-1`, así que la lista abre **vacía** (`:513-518`). Error: `MsgBox` y cierre del formulario (`:377-383`).

**Filtrado.** Cada cambio llama a `ActualizarLista` (`:555-616`): deshabilita Editar/Eliminar/Anexar y oculta `ImagenAnexo` y `lblFacturado` (`:565-569`), construye el SQL con `Filtrar` (`:575-577`) y lo asigna a `RowSource` con `Requery` (`:583-585`). Sin filas, deshabilita Exportar (`:588-590`); con una sola, la selecciona y ejecuta `ListaFiltrados_Click` (`:591-594`); con varias, solo reactiva la selección previa si la hay (`:596-602`).

**Selección (`ListaFiltrados_Click`, `:425-483`).** Sin ID, desactiva los botones (`:433-443`). Con ID habilita Editar y Anexar (`:444-445`); si la columna 5 (`IDFactura`) está vacía habilita Eliminar, y si no la deshabilita y muestra `lblFacturado` (`:446-455`). Lee el gasto con `Constructor.getGasto` y guarda `URLAnexos` en `strURLAnexo`; la imagen se muestra si hay anexos (`:456-466`).

**Eliminación (`:134-183`).** Exige selección (`:140-146`); pregunta con botón «No» por defecto, con texto distinto si `strURLAnexo` no está vacío (`:150-158`); llama a `Eliminar` del formulario (`:166`), que verifica el ID numérico, carga el gasto y llama a `Gasto.Eliminar` (`:658-673`); recarga la lista (`:172`).

**Anexos.** `ImagenAnexo_Click` sin URL no hace nada (`:393-395`); con `;` en la URL delega en `ComandoAnexar_Click` (`:399-402`); con una sola llama a `AbrirEnLocal` (`FG.bas:218-260`), que comprueba, copia al directorio local configurado y abre el fichero (`:408-413`).

**Exportación.** `GenerarConsultasSQL` (`FG.bas:1584-1688`) ejecuta el SQL actual de la lista, vuelca cabeceras y filas a un `.xlsx` en el directorio local configurado y lo abre (`:190-194`; `FG.bas:1619-1670`).

**Errores.** Todos los handlers muestran `MsgBox` crítico con `strTextoError`, salvo los dos `On Error Resume Next` (`:207-208, 385-386`). El error interno 1000 conserva el texto ya cargado (`:27-31`).

## Reglas, validaciones y cálculos

- **Selección obligatoria** para Editar, Eliminar y Anexar (`:69-72, 103-106, 141-146`).
- **Importe total** = `ImporteUnitario` × `NumeroUnidades` formateado `#,##0.00 €`, o `Null` si alguno no es numérico (`:507`).
- **Fechas:** inicio vacío o no válido ⇒ `01/01/1900`; fin vacío o no válido ⇒ `01/01/2100`; se formatean `mm/dd/yyyy` (`:519-528, 539`). Filtro `Between` sobre `FechaImputacionGasto`, orden descendente (`:539-544`).
- **Gasto/técnico vacío** ⇒ `Like '*' Or Is Null`; informado ⇒ igualdad con el texto entre comillas (`:529-538`).
- **No eliminable si está facturado:** el botón se deshabilita por `IDFactura` (`:449, 453`) y `Gasto.Eliminar` lo rechaza con «Está facturado» (`Gasto.cls:302-305`).
- **Eliminación en cascada de anexos** antes del `DELETE` (`Gasto.cls:298, 306-309`); se aborta si algún anexo está abierto (`Gasto.cls:268-271`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista filtrada | `TbGastos`, `TbFacturaGastosInvolucrados`, `TbTecnicos`, `TbAnexos` | SQL dinámico concatenado (`LEFT JOIN`, `IIf`, `Format`, `Like '*'`, literales `#fecha#`) asignado a `RowSource` | — | `:506-545, 584` |
| Combo técnico | `TbTecnicos` | SQL guardada en `RowSource` | — | `Form_FormGastoGestion.form.txt:1208` |
| Combo gasto | `TbGastos` | `SELECT DISTINCT ... ORDER BY` en `RowSource` | — | `Form_FormGastoGestion.form.txt:1603-1604` |
| Carga del gasto | `TbGastos` | SQL concatenado `WHERE IDGasto=` | — | `Constructor.bas:1869-1871` |
| Exportación | las de la lista | DAO `OpenRecordset` sobre el SQL | — | `FG.bas:1619` |
| Baja del gasto | `TbGastos` | `Execute` de `DELETE` concatenado | no | `Gasto.cls:306-309` |
| Baja de anexos | `TbAnexos` | `Anexo.Eliminar` por anexo | no | `Gasto.cls:272-279` |

## Coexistencia

Adaptar a PostgreSQL vinculado: `IIf`, `IsNumeric`, `CDbl` y `Format(..., '#,##0.00 €')` dentro del SQL (`:507`); comodín `Like '*'` (`:530, 535`); literales de fecha `#...#` en `mm/dd/yyyy` (`:539`); `DELETE *` (`Gasto.cls:306`); `LEFT JOIN` anidado con paréntesis (`:509-512`). Baja de anexos y del gasto sin transacción (`Gasto.cls:272-309`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Gasto.Eliminar` borra los anexos antes de comprobar si está facturado (`Gasto.cls:298` frente a `:302`); si el botón se habilitara con un gasto facturado quedarían anexos eliminados y el gasto vivo. Sin transacción.
- El `LEFT JOIN` con `TbAnexos` (`:512`) y con `TbFacturaGastosInvolucrados` (`:509`) puede repetir un gasto por cada anexo o factura.
- `strGasto` y `strAlias` se concatenan entre comillas simples (`:532, 537`): un apóstrofo rompe la consulta.
- `strURLAnexo` es variable de módulo (`:4`) y no se limpia al deseleccionar o recargar; el texto de confirmación de borrado depende de la última selección con anexos (`:150-158`).
- `m_FormularioQueAbre` se declara y se pasa como argumento a `FormAnexos` sin asignarse (`:62, 82`).
- `ComandoEliminar_Click` repite la misma comprobación anidada (`:141-142`) y descarta el valor devuelto por `Eliminar` (`:166`).
- Con varias filas, `ActualizarLista` no selecciona ninguna y `strIDGastoSeleccionadoActualmente` se asigna sin usarse (`:592, 598`).
- `ComboGastoBusqueda`/`ComboTecnicoBusqueda` refrescan en `Change` y en `Exit` (`:223-244, 313-334`): consultas repetidas mientras se escribe.
- Rótulos de diseño sin sentido (`Comando227`, `Comando46`, `ComandoLimpiarEquipo`): `Form_FormGastoGestion.form.txt:1237, 1675, 1818, 1998, 2178`.
- No se observa comprobación de rol en este formulario ni en su apertura (`Form_Form0BDOpciones.cls:96-99`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtros parametrizados en servidor, paginación en lugar de lista vacía inicial, baja transaccional gasto + anexos.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, la lista aparece vacía hasta aplicar algún filtro (`:513-518`).
2. Filtrar por técnico, tipo y fechas devuelve solo los gastos coincidentes, ordenados por fecha descendente (`:539-544`).
3. Un gasto facturado muestra «Facturado» y no permite eliminar (`:450-454`).
4. Eliminar sin selección muestra «Seleccione el gasto que pretende eliminar» (`:143`).
5. Eliminar con confirmación borra el gasto y sus anexos y recarga la lista (`:166-172`).
6. Exportar sin filas no está disponible; con filas genera un Excel con las columnas de la lista (`:588-590, 604`).
7. El clic en el icono abre el anexo único o `FormAnexos` si hay varios (`:399-413`).

## Pendientes

- Contenido de `Anexo.Eliminar`, `Anexo.URLAnexo` y `Gasto.ColAnexos`/`AnexosCerrados` (llamados en `Gasto.cls:94, 103-107, 263, 274`).
- Función `Ejecutar` y propiedad `Entorno.URLDirectorioLocal` (`FG.bas:246, 1670`; `Entorno.cls:191`).
- Variables globales sin `Dim` local `pregunta`, `flag`, `dato`, `m_SQL` (`:31, 166, 410, 190`): declaración en `Variables Globales.bas` no verificada.
- Estructura, índices y tipos de `TbGastos`, `TbAnexos` y `TbFacturaGastosInvolucrados` (ERD de origen).
