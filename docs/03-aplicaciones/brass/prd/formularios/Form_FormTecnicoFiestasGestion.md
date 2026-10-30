# Form_FormTecnicoFiestasGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormTecnicoFiestasGestion.cls` (315 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado de las fiestas (festivos) del calendario de técnicos, con filtro por año, alta, baja y exportación a Excel. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): la lista `ListaFiestas` es una lista de valores que el código rellena con `Constructor.getFiestas` sobre `TbTecnicosFiestas` (`Constructor.bas:4002-4062`). Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`; `VG.bas` = `modules/Variables Globales.bas`. Las citas sin fichero son `Form_FormTecnicoFiestasGestion.cls:línea`. Capacidad de todos los controles: CAP-9 Técnicos, ausencias, fiestas y precios.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormTecnicoGestion` | Abre el formulario con `DoCmd.OpenForm` | `Form_FormTecnicoGestion.cls:120` |

Hacia `Form_FormTecnicoFiestas` (alta): `ComandoAlta_Click` lo abre y se suscribe a su evento `Alta` mediante `m_Form` (`:94-95`, `:7`); al recibirlo, `m_Form_Alta` vuelve a filtrar (`:298-307`; el evento se declara en `Form_FormTecnicoFiestas.cls:7` y se lanza en `:28`). `cmdSalir` cierra sin guardar (`:255-258`).

## Controles y eventos

Controles: 19 (21 `Begin` en el `.form.txt`: 19 controles más `Form` y `Section`).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ListaFiestas` | lista de valores, 2 columnas con cabecera | `OnClick` → `ListaFiestas_Click`; `OnDblClick` declarado | Lista `Fecha;Observaciones`; al hacer clic carga la fiesta seleccionada y habilita `ComandoEliminar` | CAP-9 | `:259-296`, `.form.txt:429-466` |
| `ComboAño` | combo de valores | `AfterUpdate` → `ComboAño_AfterUpdate` | Años distintos con fiestas; filtra la lista | CAP-9 | `:199-221`, `.form.txt:495-529` |
| `ComandoLimpiarAño` | botón | `OnClick` | Vacía el combo y muestra todas las fiestas | CAP-9 | `:173-197`, `.form.txt:574-585` |
| `ComandoAlta` | botón «Alta» | `OnClick` | Abre `FormTecnicoFiestas` | CAP-9 | `:92-97`, `.form.txt:306-316` |
| `ComandoEliminar` | botón «Eliminar» | `OnClick` | Borra la fiesta seleccionada tras confirmar | CAP-9 | `:99-144`, `.form.txt:359-372` |
| `ComandoExportarExcel` | botón | `OnClick` | Exporta el contenido de la lista a Excel | CAP-9 | `:145-170`, `.form.txt:742-752` |
| `cmdSalir` | botón «Salir» | `OnClick` | Cierra el formulario | CAP-9 | `:255-258`, `.form.txt:332-342` |
| `ComandoAyudaFormulario` | botón con imagen | — (sin evento) | Sin manejador | CAP-9 | `.form.txt:391-401` |
| `lblTitulo1`, `lblTitulo2`, `Etiqueta192` | 3 etiquetas | — | Títulos y rótulo «AÑO»; `Form_Open` copia `lblTitulo1` a `Caption` | CAP-9 | `:231`, `.form.txt:468-570` |
| `CuadroSuperior`, `Cuadro2` | 2 rectángulos | — | Decoración | CAP-9 | `.form.txt:260-304` |
| Definiciones por defecto del formulario | 6 `Begin` (etiqueta, rectángulo, botón, grupo de opciones, lista, combo) | — | Plantillas de estilo sin `Name` | sin asignar | `.form.txt:141-248` |

## Flujos

### Apertura (`Form_Open`, `:222-254`)

1. Cursor de espera, `AjustarTamaño Me` (`FG.bas:261`) y título desde `lblTitulo1` (`:226-231`).
2. `RellenarCombo` añade los años distintos de `TbTecnicosFiestas` (`:9-38`); sin filas sale sin error (`:21-25`).
3. `Filtrar` sin año carga todas las fiestas, de la más reciente a la más antigua (`:237`; `Constructor.bas:4023-4031`).
4. Error: `MsgBox` crítico y cierre sin guardar (`:247-253`).

### Filtrado (`Filtrar`, `:39-90`)

1. Deshabilita `ComandoEliminar` y reinicia la lista a la cabecera (`:53-56`). Pide el diccionario a `getFiestas`; si es `Nothing` termina con la lista vacía (`:57-63`).
2. Añade una fila `FechaFiesta;Observaciones` por fiesta (`:64-68`).
3. Con una sola fiesta (`ListCount = 2` por la cabecera) la selecciona y simula el clic (`:70-72`); si ya había selección, repite el clic (`:74-81`).

`ComboAño_AfterUpdate` filtra con `Column(0)` del combo (`:206`); `ComandoLimpiarAño_Click` pone el combo a `Null` y filtra sin año (`:181-182`).

### Baja (`ComandoEliminar_Click`, `:99-144`)

1. Si no hay objeto seleccionado, lo carga con `Column(0)` de la lista o avisa «Ha de seleccionar un elemento de la lista» (`:104-117`).
2. Pregunta «¿Desea realmente eliminar la fiesta seleccionada?» con «No» por defecto (`:119-122`).
3. `Fiesta.Borrar` y nuevo `Filtrar` con el año del combo (`:124-134`). Error: `MsgBox` crítico (`:138-143`).

### Alta y exportación

Alta: la validación vive en `Form_FormTecnicoFiestas` (`Fiesta.Registrar`, `Fiesta.cls:68-106`). Exportación: `GenerarConsultas("FiltroFiestas", Me.hWnd, , Me.ListaFiestas, "No")` (`:154`; `FG.bas:1384`); un resultado con «|» se trata como error (`:155-159`).

## Reglas, validaciones y cálculos

- **Filtro por año:** solo si `p_Año` es numérico; en otro caso, todas las fiestas (`Constructor.bas:4021-4031`).
- **Una fiesta por fecha:** el diccionario se indexa por `FechaFiesta` e ignora duplicados (`Constructor.bas:4048-4050`).
- **Fecha obligatoria, observaciones obligatorias y fecha no repetida** al registrar (`Fiesta.cls:30-49`; mensajes en `:31, :35, :47`).
- **Borrado por fecha:** exige fecha válida (`Fiesta.cls:117-120`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Años del combo | `TbTecnicosFiestas` | SQL fija con `Year()` (`DISTINCT`) | no | `:18-19` |
| Listado de fiestas | `TbTecnicosFiestas` | SQL dinámica concatenando el año numérico | no | `Constructor.bas:4021-4033` |
| Lectura de una fiesta | `TbTecnicosFiestas` | SQL dinámica con literal de fecha `#mm/dd/yyyy#` | no | `Constructor.bas:3971-3977` |
| Baja | `TbTecnicosFiestas` | `DELETE` dinámico con literal de fecha, `CurrentDb().Execute` | no | `Fiesta.cls:121-123` |
| Alta (desde el otro formulario) | `TbTecnicosFiestas` | DAO `AddNew`/`Update`; duplicados con `Count` dinámico | no | `Fiesta.cls:38-41, 90-97` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: literales de fecha `#mm/dd/yyyy#` y `Year()` de Jet/ACE (`Constructor.bas:3976, 4023`; `:18`), `DELETE *` (`Fiesta.cls:121`), `DISTINCT` sobre expresión de año, apertura de la tabla completa para insertar (`Fiesta.cls:90-91`) y `FechaFiesta` tratada como cadena en VBA (`Fiesta.cls:14`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ListaFiestas` declara `OnDblClick` (`.form.txt:449`) y no hay procedimiento `ListaFiestas_DblClick` en el `.cls`; el doble clic no haría nada.
- `RellenarCombo` solo se ejecuta en `Form_Open` (`:232`): un año nuevo creado con el alta no aparece en el combo hasta reabrir.
- `Fiesta.Borrar` usa `m_SQL` sin declararlo en la clase (`Fiesta.cls:121`); resuelve a la global pública `m_SQL` (`VG.bas:203`), compartida con el resto de la aplicación.
- `Filtrar` concatena `p_Año` en el SQL tras `IsNumeric` (`Constructor.bas:4021-4024`); `IsNumeric` admite valores como `1e3`.
- `ListaFiestas_Click` usa `Err.Raise 100` (`:286`), distinto del 1000 interno, de modo que el mensaje del error se sobrescribe en `:292`.
- Rótulos de diseño sin sentido: `ComandoAyudaFormulario` «Comando46», `ComandoLimpiarAño` «Comando36», `ComandoExportarExcel` «Comando227» (`.form.txt:400, 584, 751`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: clave primaria natural por fecha con restricción única, edición de observaciones (hoy solo alta/baja) y refresco de años tras el alta.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir se listan todas las fiestas de la más reciente a la más antigua y el combo ofrece los años existentes (`:231-237`).
2. Elegir un año filtra la lista; «limpiar» vuelve a mostrar todas (`:206, :181-182`).
3. `Eliminar` está deshabilitado hasta seleccionar una fila (`:53, :288`).
4. Eliminar pide confirmación y, si se acepta, la fiesta desaparece del listado (`:119-134`).
5. Tras un alta correcta en el otro formulario, la lista se refresca conservando el año (`:298-307`).
6. Exportar a Excel genera un fichero con las filas visibles (`:154`).
7. Una fecha ya registrada o unas observaciones vacías se rechazan al dar de alta (`Fiesta.cls:30-49`).

## Pendientes

- Ubicación y formato de la exportación de `GenerarConsultas` con `strRemota = "No"` (`:154`).
- Contenido de `AjustarTamaño` más allá del tamaño (`FG.bas:261`; `DoCmd.Restore`).
- Índice único de `TbTecnicosFiestas.FechaFiesta` y tipo real de la columna (ERD de origen).
