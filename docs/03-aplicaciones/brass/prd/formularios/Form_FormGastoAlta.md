# Form_FormGastoAlta

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormGastoAlta.cls` (250 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta de un gasto en `TbGastos`. Es un formulario **no enlazado** (sin `RecordSource` en `Form_FormGastoAlta.form.txt`): recoge tipo, importe unitario, unidades, fecha, alias y observaciones, y la función pública `Alta` (`Form_FormGastoAlta.cls:165-249`) inserta el registro con DAO. Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`; `VG.bas` = `modules/Variables Globales.bas`; `:N` sin fichero se refiere a `Form_FormGastoAlta.cls`. Capacidad de todos los controles: CAP-8 Facturación, gastos y subcontratación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormGastoGestion` (`ComandoAlta_Click`) | Cierra `FormGastoAlta` si ya estaba abierto y lo abre sin argumentos | `Form_FormGastoGestion.cls:33`, `Form_FormGastoGestion.cls:45-48` |

Al terminar con éxito, si `FormGastoGestion` está abierto llama a `Form_FormGastoGestion.ActualizarLista` (`:54-56`, definida en `Form_FormGastoGestion.cls:555`) y se cierra sin guardar el diseño (`:57`). `cmdSalir` cierra sin preguntar (`:70-73`). No devuelve valores ni lanza eventos.

## Controles y eventos

Controles: 24 (33 `Begin` en el `.form.txt`: 1 formulario, 7 plantillas de control por defecto en `:144-247`, 1 sección `Detalle` y 24 controles con nombre).

| Control | Tipo | Evento y handler | Qué hace / origen de datos | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComboTipoGasto` | combo | `OnChange` → `ComboTipoGasto_Change` | Lista `SELECT DISTINCT TipoGasto FROM TbGastos`; no limita a la lista. Al cambiar, busca el importe vigente del tipo y lo copia a `ImporteUnitario` | CAP-8 | `Form_FormGastoAlta.form.txt:476-480`, `:74-117` |
| `ImporteUnitario` | texto no enlazado | `OnExit` → `ImporteUnitario_Exit` | Formato moneda; al salir recalcula el total | CAP-8 | `Form_FormGastoAlta.form.txt:616-618`, `:147-155` |
| `NUnidades` | texto no enlazado | `OnExit` → `NUnidades_Exit` | Formato numérico general; al salir recalcula el total | CAP-8 | `Form_FormGastoAlta.form.txt:674-676`, `:156-164` |
| `FechaImputacion` | texto no enlazado | — | Formato fecha corta; opcional (si no es fecha se usa `Now()`) | CAP-8 | `Form_FormGastoAlta.form.txt:441-442`, `:200-202` |
| `ComboAlias` | combo (lista de valores, 3 columnas, limita a la lista) | — | Se rellena en `Form_Load` con `RellenarComboTecnicos` (alias, nombre, tipo); opcional | CAP-9 | `Form_FormGastoAlta.form.txt:694-714`, `:128` |
| `Observaciones` | texto no enlazado | — | Texto libre opcional | CAP-8 | `Form_FormGastoAlta.form.txt:560`, `:233-235` |
| `ImporteTotal` | texto bloqueado, fuera del orden de tabulación | — | Calculado: importe × unidades | CAP-8 | `Form_FormGastoAlta.form.txt:784-799`, `:4-23` |
| `ComandoAlta` (rótulo «Registrar») | botón | `OnClick` → `ComandoAlta_Click` | Valida y da de alta | CAP-8 | `Form_FormGastoAlta.form.txt:333-335`, `:24-68` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | CAP-8 | `Form_FormGastoAlta.form.txt:371-373`, `:70-73` |
| `lblTitulo1`, `lblTitulo2` | etiquetas «Alta de Gasto» | — | `Me.Caption` toma el rótulo de `lblTitulo1` | CAP-8 | `Form_FormGastoAlta.form.txt:825-848`, `:127` |
| `Etiqueta142`, `Etiqueta198`, `Etiqueta201`, `Etiqueta219`, `NUnidadeslbl`, `Etiqueta229`, `Etiqueta230` | 7 etiquetas de campo | — | Rótulos | CAP-8 | `Form_FormGastoAlta.form.txt:413-414, 505-506, 532-533, 587-588, 645-646, 740-741, 767-768` |
| `Etiqueta1`, `Etiqueta2`, `Etiqueta3`, `Etiqueta4` | 4 etiquetas «*» | — | Marca visual de obligatorio | CAP-8 | `Form_FormGastoAlta.form.txt:868-935` |
| `CuadroSuperior`, `Cuadro2` | 2 rectángulos | — | Decoración | CAP-8 | `Form_FormGastoAlta.form.txt:288, 311` |

## Flujos

### Apertura (`Form_Load`, `:118-146`)

1. `AjustarTamaño Me` (`FG.bas:261`), título desde `lblTitulo1` y `RellenarComboTecnicos Me.ComboAlias, EnumSino.No, EnumSino.Sí, strTextoError` (`:126-128`; definida en `FG.bas:7224`, usa `Constructor.getTecnicos`, `FG.bas:7239`).
2. Si devuelve error, lo eleva como 1000, muestra `MsgBox` crítico y cierra el formulario (`:129-131, 139-145`).

### Cambio de tipo (`ComboTipoGasto_Change`, `:74-117`)

Tipo vacío ⇒ `ImporteUnitario = Null` y sale (`:77-81`). Si no, abre `TbGastosImportePorTipo` con `FechaFinalImporte Is Null` para ese tipo (`:85-89`); si hay fila numérica, escribe el importe con formato «#,##0.00 €», si no deja `Null` (`:97-101`), y recalcula el total (`:102`).

### Alta (`ComandoAlta_Click`, `:24-68`)

1. Lee los seis campos con `Nz` (`:32-37`) y llama a `Alta` (`:43`).
2. `Alta` valida (reglas abajo), obtiene el ID con `DameID("TbGastos", "IDGasto")` (`:208`), comprueba que no exista (`:215-223`) e inserta con `AddNew`/`Update` (`:224-236`). Devuelve el ID o `#ERR|texto` (`:238, 248`).
3. Si el resultado contiene `|`, muestra el error (`:44-48`); en éxito actualiza la lista de Gestión y cierra (`:54-57`).
4. Cualquier error: `MsgBox` crítico y el formulario queda abierto (`:62-67`).

## Reglas, validaciones y cálculos

- **Obligatorios:** tipo no vacío, importe unitario numérico, unidades numéricas (`:188-199`).
- **Fecha:** si no es fecha válida se sustituye por `Now()` (`:200-202`) y se guarda como texto `dd/mm/yyyy` (`:229`).
- **Alias y observaciones:** solo se escriben si no están vacíos (`:230-235`).
- **Identificador:** `DameID` recorre todos los `IDGasto` de la tabla y devuelve el máximo numérico + 1 (`FG.bas:409-428`).
- **Total:** `ImporteUnitario × NUnidades` si ambos son numéricos; si el resultado es 0, `Null`; si no, texto con formato «#,##0.00 €» (`:10-19`). `CalculoTotal` usa `On Error Resume Next` (`:6`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de tipos | `TbGastos` | SQL guardada en `RowSource` (`SELECT DISTINCT`) | — | `Form_FormGastoAlta.form.txt:478` |
| Importe vigente por tipo | `TbGastosImportePorTipo` | SQL dinámico con el tipo concatenado entre comillas | — | `:85-89` |
| Máximo ID | `TbGastos` | `SELECT` del campo ID y bucle en VBA | no | `FG.bas:408-428` |
| Comprobación de ID | `TbGastos` | SQL dinámico con el ID concatenado | — | `:215-218` |
| Alta | `TbGastos` | DAO `AddNew`/`Update` (campos `IDGasto`, `TipoGasto`, `ImporteUnitario`, `NumeroUnidades`, `FechaImputacionGasto`, `ALIAS`, `Observaciones`) | no | `:224-236` |
| Técnicos del combo | tabla de técnicos | vía `Constructor.getTecnicos` | — | `FG.bas:7239` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: comillas simples concatenadas en el tipo de gasto (`:87`); ID calculado en VBA con máximo + 1 fuera de transacción; fecha escrita como texto `dd/mm/yyyy` en un campo `FechaImputacionGasto` (`:229`); importes convertidos con `CDbl` y formato regional «€» (`:98, 227`); `AddNew`/`Update` sobre recordset sin transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `DameID` (máximo + 1) sin transacción ni bloqueo (`FG.bas:408-429`; `:208`): dos altas simultáneas podrían calcular el mismo ID; la comprobación posterior (`:220-223`) no es atómica.
- El tipo de gasto se concatena entre comillas simples (`:87`): un valor con apóstrofo rompería la consulta. `ComboTipoGasto` admite texto libre (sin `LimitToList` en `Form_FormGastoAlta.form.txt:476-480`).
- Los mensajes de error de `Alta` dicen «El método Eliminar» (`:46`), y el comentario de cabecera cita `FormGastoAlta.ComandoRegistrar_Click` (`:180`) cuando el handler es `ComandoAlta_Click` (`:24`).
- `CalculoTotal` con `On Error Resume Next` (`:6`) oculta errores de conversión; con configuración regional distinta de la española, `CDbl` sobre un texto con «€» (`:11`) podría fallar o interpretar mal los separadores.
- `m_SQL`, `flag`, `dato` y `pregunta` son globales compartidas (`VG.bas:203-206`), usadas por este formulario sin declararlas localmente (`:43-47, 85, 208`).
- La fecha se guarda formateada como texto `dd/mm/yyyy` (`:229`); según la configuración regional podría invertirse día y mes.
- Si el usuario cambia el importe a mano tras elegir el tipo, no hay tope ni aviso; el total solo se recalcula al salir del campo (`:147-164`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: identificador generado por la base de datos; lista cerrada de tipos de gasto; fecha como tipo fecha nativo; recálculo del total en el cliente.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, `ComboAlias` muestra los técnicos activos con alias, nombre y tipo (`:128`).
2. Elegir un tipo con importe vigente rellena `ImporteUnitario` y `ImporteTotal`; elegir un tipo sin importe vigente deja el importe vacío (`:85-102`).
3. Vaciar el tipo deja `ImporteUnitario` vacío (`:77-81`).
4. Registrar sin tipo, sin importe numérico o sin unidades numéricas muestra el mensaje correspondiente y no escribe (`:188-199`).
5. Registrar sin fecha guarda la fecha y hora actuales (`:200-202`).
6. Un alta correcta crea una fila en `TbGastos` con ID = máximo + 1, cierra el formulario y refresca la lista de Gestión si está abierta (`:208-236, 54-57`).
7. Un alta con alias y observaciones vacíos no escribe esos campos (`:230-235`).
8. `ImporteTotal` no es editable (`Form_FormGastoAlta.form.txt:784`).

## Pendientes

- Tipos de columna, valores por defecto e índices de `TbGastos` y `TbGastosImportePorTipo` (ERD de origen).
- Contenido de `AjustarTamaño` (`:126`) y de `Constructor.getTecnicos` con los parámetros `EnumSino.No`/`EnumSino.Sí` (`:128`).
- Si `ImporteUnitario` con formato de moneda devuelve texto numérico utilizable por `CDbl` en todas las configuraciones regionales (`:11, 227`).
