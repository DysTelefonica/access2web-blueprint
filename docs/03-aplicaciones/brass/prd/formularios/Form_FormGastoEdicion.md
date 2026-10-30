# Form_FormGastoEdicion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormGastoEdicion.cls` (558 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Edición de un gasto existente de `TbGastos`. Es un formulario **no enlazado** (sin `RecordSource`): `Form_Open` recibe en `OpenArgs` el ID, la ruta del anexo y el estado de facturación (`Form_FormGastoEdicion.cls:277-291`), `EstablecerDatos` carga el registro y `EditarGasto` (`:469-555`) lo actualiza con DAO. Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`; `VG.bas` = `modules/Variables Globales.bas`; `:N` sin fichero se refiere a `Form_FormGastoEdicion.cls`. Capacidad de todos los controles: CAP-8 Facturación, gastos y subcontratación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormGastoGestion` (`ComandoEditar_Click`) | Lee ID (columna 0) y estado de facturación (columna 5) de `ListaFiltrados`; el estado se normaliza a «No» si está vacío y «Sí» en otro caso; cierra el formulario si estaba abierto y lo abre con `OpenArgs` = `ID\|URLAnexo\|Facturado` | `Form_FormGastoGestion.cls:95-122` |

Al guardar con éxito, si `FormGastoGestion` está abierto hace `Requery` de su `ListaFiltrados` (`:252-254`, `:80-82`) y se cierra sin guardar el diseño (`:255`). No devuelve valores ni lanza eventos.

## Controles y eventos

Controles: 26 (35 `Begin` en el `.form.txt`: 1 formulario, 7 plantillas de control por defecto en `Form_FormGastoEdicion.form.txt:145-248`, 1 sección `Detalle` y 26 controles con nombre).

| Control | Tipo | Evento y handler | Qué hace / origen de datos | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComboTipoGasto` | combo | **sin evento enlazado** (existen `_Change` y `_Exit`) | Lista `SELECT DISTINCT TipoGasto FROM TbGastos`; no limita a la lista | CAP-8 | `Form_FormGastoEdicion.form.txt:713-715`, `:144-221` |
| `ImporteUnitario` | texto no enlazado | **sin evento enlazado** (existe `_Exit`) | Formato moneda | CAP-8 | `Form_FormGastoEdicion.form.txt:852-853`, `:451-459` |
| `NUnidades` | texto no enlazado | **sin evento enlazado** (existe `_Exit`) | Formato numérico general | CAP-8 | `Form_FormGastoEdicion.form.txt:909-910`, `:460-468` |
| `FechaImputacion` | texto no enlazado | — | Formato fecha corta; obligatoria al guardar | CAP-8 | `Form_FormGastoEdicion.form.txt:678-679`, `:510-513` |
| `ComboAlias` | combo (lista de valores, 3 columnas, limita a la lista) | — | Se rellena en `EstablecerDatos` | CAP-9 | `Form_FormGastoEdicion.form.txt:928-948`, `:380` |
| `Observaciones` | texto no enlazado | — | Texto libre | CAP-8 | `Form_FormGastoEdicion.form.txt:796` |
| `ImporteTotal` | texto bloqueado | — | Se calcula solo al abrir (importe × unidades de la base) | CAP-8 | `Form_FormGastoEdicion.form.txt:1018-1033`, `:365-367` |
| `ComandoRegistrar` (rótulo «Registrar») | botón | `OnClick` → `ComandoRegistrar_Click` | Guarda y cierra; se deshabilita si el gasto está facturado | CAP-8 | `Form_FormGastoEdicion.form.txt:334-336`, `:222-266, 368-374` |
| `cmdSalir` | botón | `OnClick` | Cierra; pregunta si guardar si hay cambios y no está facturado | CAP-8 | `Form_FormGastoEdicion.form.txt:372-374`, `:55-96` |
| `lblFacturado` | etiqueta, oculta por defecto | — | Visible si el gasto está facturado | CAP-8 | `Form_FormGastoEdicion.form.txt:449-461`, `:368-374` |
| `ImagenAnexo` | imagen | `OnClick` → `ImagenAnexo_Click` | Visible si hay ruta de anexo; copia el anexo a la carpeta local y lo abre | CAP-11 | `Form_FormGastoEdicion.form.txt:489-490`, `:375-379, 396-423` |
| `lblTitulo1`, `lblTitulo2` | etiquetas «Edición de Gasto» | — | `Me.Caption` toma el rótulo de `lblTitulo1` | CAP-8 | `Form_FormGastoEdicion.form.txt:413-436`, `:276` |
| `Etiqueta142`, `Etiqueta198`, `Etiqueta201`, `Etiqueta219`, `NUnidadeslbl`, `Etiqueta229`, `Etiqueta230` | 7 etiquetas de campo | — | Rótulos | CAP-8 | `Form_FormGastoEdicion.form.txt:650-651, 741-742, 768-769, 823-824, 880-881, 974-975, 1001-1002` |
| `Etiqueta1`, `Etiqueta2`, `Etiqueta3`, `Etiqueta4` | 4 etiquetas «*» | — | Marca visual de obligatorio | CAP-8 | `Form_FormGastoEdicion.form.txt:1057-1124` |
| `CuadroSuperior`, `Cuadro2` | 2 rectángulos | — | Decoración | CAP-8 | `Form_FormGastoEdicion.form.txt:289, 312` |

## Flujos

### Apertura (`Form_Open`, `:267-320`)

1. `AjustarTamaño Me`, título desde `lblTitulo1` (`:274-276`).
2. `OpenArgs` debe contener `|` y partirse en exactamente tres partes; el ID debe ser numérico; si no, «No se han cargado todos los datos» (`:278-295`).
3. `EstablecerDatos` (`:321-395`): abre `TbGastos` por ID (`:329-332`); si no existe, error (`:334-337`). Guarda los valores iniciales (`:338-343`), los copia a los controles (`:347-364`) y calcula el total (`:365-367`). Si `strFacturado = "Sí"`, deshabilita `ComandoRegistrar` y muestra `lblFacturado` (`:368-374`). Muestra `ImagenAnexo` solo si hay ruta (`:375-379`). Rellena `ComboAlias` con `RellenarComboTecnicos ..., EnumSino.Sí, , strTextoError` (`:380`; `FG.bas:7224`).
4. Cualquier error: `MsgBox` crítico y cierre del formulario (`:313-319`).

### Guardado (`ComandoRegistrar_Click`, `:222-266`)

1. Sin cambios (`HaHabidoCambios`, `:7-54`) ⇒ «No ha habido cambios que guardar» (`:229-235`).
2. `Registrar` (`:97-143`) exige `strIDGasto` (`:114-117`), lee los controles con `Nz(..., "#")` (`:118-123`) y llama a `EditarGasto` (`:128-129`).
3. En éxito: `Requery` de la lista de Gestión y cierre (`:252-255`). En error: `MsgBox` crítico y el formulario queda abierto (`:260-265`).

### Salida (`cmdSalir_Click`, `:55-96`)

Si hay cambios **y** `strFacturado = "No"`, pregunta «¿Desea guardar los cambios producidos?»; con «Sí» ejecuta `Registrar` (`:59-75`). Cierra siempre (`:88`). Si el gasto está facturado, cierra sin preguntar aunque haya cambios.

### Anexo (`ImagenAnexo_Click`, `:396-423`)

`AbrirEnLocal(strURLAnexo, Me.hWnd)` (`:407`; `FG.bas:218-258`): comprueba que exista la ruta de red configurada, copia el fichero a la carpeta local del entorno (borrando antes una copia previa) y lo abre con el programa asociado.

## Reglas, validaciones y cálculos

- **`EditarGasto`:** ID numérico; tipo ni vacío ni «#»; importe y unidades numéricos; fecha válida (`:494-513`).
- **Marcador «#»:** un alias u observaciones con valor «#» (control vacío) se guarda como `Null`; vacío se ignora; otro valor se escribe (`:528-541`).
- **Facturado:** el comentario de cabecera dice que no se puede cambiar importe, unidades ni fecha (`:483`), pero el código de `EditarGasto` no lo comprueba; el bloqueo es solo la inhabilitación de `ComandoRegistrar` (`:368-370`).
- **Detección de cambios:** compara seis valores contra los iniciales; fecha como `dd/mm/yyyy` e importe como `CStr(CDbl(...))` (`:35-42`); ante error devuelve «hay cambios» (`:51-53`).
- **Total:** misma fórmula que en el alta (`:430-439`); solo se ejecuta si se llama a `CalculoTotal`.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de tipos | `TbGastos` | SQL guardada en `RowSource` (`SELECT DISTINCT`) | — | `Form_FormGastoEdicion.form.txt:715` |
| Lectura del gasto | `TbGastos` | SQL dinámico con el ID concatenado | — | `:329-332` |
| Importe vigente por tipo | `TbGastosImportePorTipo` | SQL dinámico con el tipo entre comillas (handlers no enlazados) | — | `:151-154, 190-193` |
| Edición | `TbGastos` | DAO `Edit`/`Update` | no | `:514-543` |
| Técnicos del combo | tabla de técnicos | vía `Constructor.getTecnicos` | — | `FG.bas:7239` |
| Copia del anexo | sistema de ficheros | `fso.CopyFile`/`DeleteFile` | — | `FG.bas:244-247` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: comillas simples concatenadas (`:153`); la fecha se escribe como texto tal cual la devuelve el control (`:527`) en un campo de fecha; importes con formato regional «€» (`:351, 366`); `Edit`/`Update` sin transacción; el anexo depende de una ruta de red configurada y de una carpeta local de Windows (`FG.bas:226-247`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Los handlers `ComboTipoGasto_Change` (`:144`), `ComboTipoGasto_Exit` (`:183`), `ImporteUnitario_Exit` (`:451`) y `NUnidades_Exit` (`:460`) no tienen `OnChange`/`OnExit` en el `.form.txt` (solo hay `OnOpen`, `OnClick` ×3: `Form_FormGastoEdicion.form.txt:43, 336, 374, 490`). De ser así, nunca se ejecutan: al cambiar tipo, importe o unidades no se actualiza el importe vigente ni el total, y `ImporteTotal` queda desfasado (a diferencia de `Form_FormGastoAlta`).
- `ComboTipoGasto_Change` y `_Exit` son duplicados exactos, y sin tipo vacío buscan con `''` (`:150-153`) y ponen `ImporteUnitario = Null` (`:162-166`); en el alta se evita con una salida temprana.
- Gasto facturado: solo se deshabilita el botón; los campos siguen editables y `cmdSalir` descarta cambios sin avisar (`:59, 368-370`). `EditarGasto` no verifica la facturación (`:483, 494-543`).
- `strFacturado` se deriva de la columna 5 de la lista: cualquier valor no vacío cuenta como «Sí» (`Form_FormGastoGestion.cls:107-111`).
- La fecha se guarda sin formatear (`:527`), mientras el alta la formatea `dd/mm/yyyy` (`Form_FormGastoAlta.cls:229`): posible inconsistencia según la configuración regional.
- Mensaje «El método Eliminar ha devuelto un error» en la apertura (`:304`) y «salir del gasto» en los handlers del tipo (`:174, 213`): textos copiados de otros formularios.
- `AbrirEnLocal` devuelve «El archivo debe estar abierto» cuando `FicheroAbierto` es cierto (`FG.bas:231-233`); el mensaje parece lo contrario de lo que ocurre.
- `m_SQL`, `flag`, `dato` y `pregunta` son globales compartidas (`VG.bas:203-206`).
- `Option Compare Database` (`:2`) y comparación de cadenas en `HaHabidoCambios` (`:41-42`) podrían no distinguir mayúsculas.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: bloquear campos y servidor cuando el gasto esté facturado; recálculo del total en el cliente; fecha como tipo nativo; apertura del anexo mediante descarga en lugar de copia local.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin `OpenArgs` completos o con ID no numérico muestra «No se han cargado todos los datos» y cierra (`:278-295`).
2. Abrir un gasto muestra tipo, importe, unidades, fecha, alias, observaciones y total (`:347-367`).
3. Un gasto facturado muestra «Facturado» y no permite guardar (`:368-370`).
4. `ImagenAnexo` solo es visible si hay anexo y al pulsarla lo abre (`:375-379, 407`).
5. Guardar sin cambios muestra «No ha habido cambios que guardar» (`:229-235`).
6. Guardar sin tipo, con importe/unidades no numéricos o fecha no válida muestra el mensaje y no escribe (`:498-513`).
7. Vaciar alias u observaciones y guardar los deja a `Null` (`:528-541`).
8. Salir con cambios en un gasto no facturado pregunta si guardar (`:59-60`).
9. Un guardado correcto refresca la lista de Gestión si está abierta (`:252-254`).

## Pendientes

- Tipos de columna e índices de `TbGastos` y `TbGastosImportePorTipo` (ERD de origen).
- Confirmar en el binario si los eventos de `ComboTipoGasto`, `ImporteUnitario` y `NUnidades` están enlazados (`Form_FormGastoEdicion.form.txt:713-715, 852-853, 909-910`).
- Contenido de `AjustarTamaño` (`:274`), `Ejecutar` (`FG.bas:247`) y `Constructor.getTecnicos` con `EnumSino.Sí` (`:380`).
- Si `FormGastoGestion.ListaFiltrados` conserva la selección tras `Requery` (`:253`).
