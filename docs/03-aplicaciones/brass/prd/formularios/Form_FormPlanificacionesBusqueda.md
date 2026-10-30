# Form_FormPlanificacionesBusqueda

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormPlanificacionesBusqueda.cls` (1058 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pantalla «GESTIÓN DE ACTUACIONES PLANIFICADAS» (`.form.txt:2078-2118`): busca planificaciones (`TbPlanificacion`, unidas a `TbPlanificacionEquipos` y `TbEquipos`) por BUI, subsistema, equipo y fechas, las lista en `ListaFiltrados` y, según el estado de la planificación seleccionada, abre los formularios de alta, cierre, reprogramación, anexos, evento o actividad que la cierra. Es un formulario **sin `RecordSource`**: la lista se rellena asignando `RowSource` en código (`Form_FormPlanificacionesBusqueda.cls:36-39`). Las citas sin fichero son `Form_FormPlanificacionesBusqueda.cls:línea`; `FG.bas` = `modules/Funciones Generales.bas`; `VG.bas` = `modules/Variables Globales.bas` (globales `m_SQL`, `flag`, `dato`, `pregunta`, `VG.bas:203-206`). Capacidad de todo el formulario: CAP-6 Planificación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormPlanificacionGestion` | Abre el formulario sin argumentos | `Form_FormPlanificacionGestion.cls:73` |
| `Form_FormPlanificacionEquipos` | Lo cierra si está abierto y lo abre con `OpenArgs` = `BUI\|SUBSISTEMA\|equipo` | `Form_FormPlanificacionEquipos.cls:480-490` |
| `Form_FormPlanificacionAlta`, `…Reprogramacion`, `…CambioIDEvento` | No lo abren: tras guardar hacen `Requery` de `ListaFiltrados` si el formulario está abierto | `Form_FormPlanificacionAlta.cls:112-113`; `Form_FormPlanificacionReprogramacion.cls:109-110`; `Form_FormPlanificacionCambioIDEvento.cls:32-33` |

Salida: abre `FormPlanificacionAlta`, `FormPlanificacionCierre`, `FormPlanificacionReprogramacion`, `FormPlanificacionAnexos`, `formPlanificacionCambioIDEvento`, `FormEventoGestion`, `FormActividadGestion` y `FormAnexos`, y genera un Excel (detalle abajo). No devuelve eventos ni valores al llamante.

## Controles y eventos

Controles: 51 (51 `Begin` en el `.form.txt`): 1 `Form` (`:5`) + 8 plantillas de formato por defecto (`:762-886`) + 1 `Section` (`:907`) + 41 controles con nombre, desglosados abajo.

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ListaFiltrados` | lista, 11 columnas con cabecera | `OnClick` → `ListaFiltrados_Click` | Muestra el resultado; al elegir fila fija el ID, lee el estado y habilita botones (columnas 0 ID, 1 Equipo, 2 Fecha Prevista, 3 FCierre, 4 MotivoCierre, 5 IDNueva, 6 Días Restan, 7 BUI, 8 SUBSISTEMA, 9 IDEquipo, 10 IDEvento) | CAP-6 | `.form.txt:987-1031`; `:902-999`; `:228-232` |
| `ComboBUIBusqueda` | combo (BUI) | `OnChange` | Si cambia el valor: `ActualizarLista`, `CargarSubSistema`, `CargarIDEquipo` | CAP-6 | `:627-650`; `.form.txt:1304-1339` |
| `ComboSubsistemaBusqueda` | combo | `OnChange` | Si cambia: `ActualizarLista` y `CargarIDEquipo` | CAP-6 | `:717-739`; `.form.txt:1264-1303` |
| `ComboEquipoBusqueda` | combo (3 columnas) | `OnChange` | Si cambia: `ActualizarLista` | CAP-6 | `:673-694`; `.form.txt:1394-1436` |
| `ComboCriterioPlanificacion` | combo de lista de valores (Realizadas, No Realizadas, Fuera de Plazo, Cumplen Próximamente, En Plazo) | `OnChange` | Si cambia: `ActualizarLista` (ver defecto sobre su efecto) | CAP-6 | `:651-672`; `.form.txt:1464-1501` |
| `ComboIDMantenimiento` | combo de `IDPlanificacion` descendente | `OnChange` | Si cambia: `ActualizarLista` (su valor no se lee) | CAP-6 | `:695-716`; `.form.txt:2010-2049` |
| `FechaInicialBusqueda` | texto, formato fecha corta | — | Límite inferior de `FechaPrevistaCierre`; sin evento | CAP-6 | `.form.txt:1668-1697` |
| `FechaFinalBusqueda` | texto, formato fecha corta | `OnExit` | Si cambia al salir: `ActualizarLista` | CAP-6 | `:740-761`; `.form.txt:1725-1755` |
| `ComandoAltaPlanificacion` («Planifi. Nuevo») | botón | `OnClick` | Exige planificación seleccionada y abre `FormPlanificacionAlta` con `BUI\|SUBSISTEMA\|IDEquipo\|Equipo` | CAP-6 | `:78-109`; `.form.txt:1529-1555` |
| `ComandoReprogramar` | botón | `OnClick` | Abre `FormPlanificacionReprogramacion` con el ID | CAP-6 | `:424-442`; `.form.txt:1556-1583` |
| `ComandoMarcarComoRealizada` | botón | `OnClick` | Abre `FormPlanificacionCierre` con el ID | CAP-6 | `:404-422`; `.form.txt:1584-1611` |
| `ComandoVerDetalle` («Ver») | botón | `OnClick` | Abre también `FormPlanificacionCierre` con el ID | CAP-6 | `:597-626`; `.form.txt:1612-1639` |
| `ComandoTrabajoQueLaCierra` | botón | `OnClick` | Abre `FormEventoGestion` o `FormActividadGestion` posicionado en el trabajo que cerró la planificación | CAP-3, CAP-4, CAP-6 | `:443-535`; `.form.txt:1640-1667` |
| `ComandoCambiarEventoDeCierre` | botón | `OnClick` | Abre `formPlanificacionCambioIDEvento` con `ID\|IDEvento\|IDEquipo\|FechaPrevista` | CAP-6 | `:255-275`; `.form.txt:2050-2077` |
| `ComandoAdjuntos` («Adjuntar») | botón | `OnClick` | Abre `FormPlanificacionAnexos` con el ID | CAP-11 | `:110-139`; `.form.txt:3298-3330` |
| `ImagenDocAnexado` | imagen | `OnClick` | Abre el anexo de la planificación o `FormPlanificacionAnexos` si hay varios | CAP-11 | `:854-901`; `.form.txt:1105-1125` |
| `ImagenAnexoEvento` | imagen | `OnClick` | Abre el anexo del evento que cierra o `FormAnexos` si hay varios | CAP-11 | `:808-853`; `.form.txt:1783-1800` |
| `ComandoInformeEnExcel` | botón | `OnClick` | Exporta a Excel el resultado actual | CAP-10 | `:276-304`; `.form.txt:1942-1982` |
| `ComandoLimpiarBUI`, `…SubSistema`, `…Equipo`, `…ID`, `…Fechas`, `…Criterios` | 6 botones | `OnClick` | Vacían su filtro y refrescan (detalle en Flujos) | CAP-6 | `:307-402`; `.form.txt:2119, 2287, 3130, 2455, 2623, 2962` |
| `ComandoActualizarTodo` | botón | `OnClick` | Invoca en cadena los seis «Limpiar» | CAP-6 | `:1033-1056`; `.form.txt:2791-2830` |
| `cmdSalir` («Salir») | botón | `OnClick` | `DoCmd.Close` sin guardar | CAP-6 | `:74-77`; `.form.txt:1032-1058` |
| `ComandoAyudaFormulario` | cuadro de texto (no botón) | `OnClick` declarado en el `.form.txt` | Sin procedimiento en el `.cls` | sin asignar | `.form.txt:2922-2961` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título; `Form_Open` copia `lblTitulo1` a `Me.Caption` | CAP-6 | `:769`; `.form.txt:2078, 2099` |
| `Etiqueta1`, `Etiqueta201`, `Etiqueta267`, `Etiqueta268`, `Etiqueta273`, `Etiqueta276`, `Etiqueta278` | 7 etiquetas | — | Rótulos «ID.», «Planif.», «BUI», «SUBS.», «Equipo.», «ENTRE», «Y» | CAP-6 | `.form.txt:1983, 1502, 1340, 1367, 1437, 1698, 1756` |
| `CuadroFormulario`, `CuadroSuperior`, `Cuadro2`, `Marco120`, `Marco189` | 3 rectángulos y 2 grupos de opciones | — | Decoración sin código | sin asignar | `.form.txt:918, 941, 963, 1059, 1082` |

## Flujos

**Apertura** (`Form_Open`, `:762-807`): reloj de arena, `AjustarTamaño` (`FG.bas:261`), título desde `lblTitulo1`. Si `OpenArgs` contiene `|` exige exactamente tres partes (si no, «Se ha abierto el formulario con datos insuficientes», `:772-777`) y, por cada parte no vacía, asigna el combo y llama a su `_Change` (`:782-793`). Sin argumentos no se llama a `ActualizarLista`: la lista queda con el `RowSource` de diseño (`.form.txt:1004-1013`). Cualquier error muestra `MsgBox` y cierra el formulario (`:800-806`).

**Filtrar** (`ActualizarLista`, `:8-65`): deshabilita los siete botones dependientes de la selección (`:16-22`), llama a `Filtrar` (`:29-30`), asigna el SQL a `ListaFiltrados.RowSource` y hace `Requery` (`:36-39`). Si queda una única fila de datos la selecciona y ejecuta `ListaFiltrados_Click` (`:42-45`); si no, relanza el clic sobre la primera fila marcada (`:47-53`). Un `|` en el resultado de `Filtrar` se trata como error (`:31-35`).

**Selección** (`ListaFiltrados_Click`, `:902-999`): `EstadoPlanificacion` (`FG.bas:2024-2095`) devuelve «Abierta» (sin `FechaCierre`), «Cerrada por trabajo realizado» (`MotivoCierre` «Trabajo Realizado») o «Cerrada por reprogramación» (`MotivoCierre` «Reprogramación»); cualquier otro motivo es error. «Abierta» habilita `ComandoMarcarComoRealizada` y `ComandoReprogramar`; las dos cerradas habilitan `ComandoTrabajoQueLaCierra` (`:938-949`). `ComandoAdjuntos` y `ComandoVerDetalle` se habilitan con cualquier selección válida (`:922-923`). `DameCadenaAnexos` (`FG.bas:2850`) decide la visibilidad de `ImagenDocAnexado` (`:960-971`); si la fila tiene `IDEvento`, `Constructor.getEvento` carga el evento, se habilita `ComandoCambiarEventoDeCierre` y `ImagenAnexoEvento` se muestra si el evento tiene `URLAnexos` (`:972-988`).

**Limpiar** (`:307-402`): cada botón hace `LimpiarUltimos`, **sale sin refrescar** si su filtro ya está vacío, pone el combo a `Null` y llama a `ActualizarLista`. `ComandoLimpiarBUI` no recarga subsistemas ni equipos. `ComandoActualizarTodo` invoca los seis (`:1039-1044`).

**Trabajo que la cierra** (`:443-535`): relee el ID de la fila, exige estado «Cerrada por trabajo realizado» (`:468-471`), lee `IDEvento` e `IDNoEvento` con `Dame` y abre `FormEventoGestion` (asigna `ComboIDEventoBusqueda` y llama a `ComboIDEventoBusqueda_AfterUpdate`, `:506-511`) o, si no hay evento, `FormActividadGestion` (`ComboIDActividadBusqueda_Change`, `:518-523`); sin ninguno de los dos, error (`:496-499`).

**Excel** (`GenerarConsultaEnExcel`, `:1000-1032`): toma `ListaFiltrados.RowSource`, lo parte por `"Where"` (`:1012-1014`), le antepone otra `SELECT` (añade BUI, SUBSISTEMA y «Evento Cierra», `:1005-1011`) y llama a `GenerarConsultasSQL` (`FG.bas:1584`), que abre un `OpenRecordset`, rechaza consulta vacía (`FG.bas:1599-1601`) o sin filas (`FG.bas:1620-1623`), vuelca a un `.xlsx` temporal en el directorio local configurado por Automatización de Excel y lo abre con `Ejecutar` (`FG.bas:1606-1608, 1625, 1670`).

**Errores:** patrón común `Err.Raise 1000` con texto propio o «error nº… Detalle…», `MsgBox` crítico y reloj de arena apagado (p. ej. `:59-65`). Los «Limpiar» y `ComandoCambiarEventoDeCierre` usan `On Error Resume Next` y no informan (`:257, 308, 324, 340, 356, 373, 389`).

## Reglas, validaciones y cálculos

- **Fechas:** `IsDate` y `Format(..., "mm/dd/yyyy")` en literales `#…#`; sin fecha inicial usa 01/01/1900 y sin final 01/01/2200 (`:171-181`). Filtra sobre `FechaPrevistaCierre`.
- **Textos de filtro:** BUI, SUBSISTEMA y Equipo se comparan con `=` entre comillas si hay valor; si no, `Like '*' Or … Is Null` (`:183-197`).
- **Días Restan:** `IIf(IsNull(FechaCierre), DateDiff('d', Now(), FechaPrevistaCierre), 0)` (`:231, 209`).
- **Criterios (rama no alcanzable desde el formulario, ver defectos):** Realizadas (`MotivoCierre` no nulo), No Realizadas, Fuera de Plazo (`Días Restan` < 0), Cumplen Próximamente (entre 0 y `DiasParaAvisoDePlanificaciones`), En Plazo (> 0), Totales (`:203-227`).
- **Selección obligatoria:** `IsNumeric(strIDPlanificacionSeleccionada)` en alta, reprogramar, marcar, ver, adjuntos y trabajo (`:85, 116, 407, 427, 603, 450`); «Planifi. Nuevo» solo «a partir de otra existente» (`:86`).
- **Combos dependientes:** subsistemas por BUI (`-1` si vacío); equipos con `FechabajaParaPlanificacion Is Null`, filtrados por BUI y subsistema solo si ambos tienen valor (`:544-547, 571-581`).
- **Anexos:** más de uno si la cadena contiene `;` (`:815, 864`).

## Lecturas y escrituras

El formulario **no escribe** en datos; solo lee y abre otros formularios.

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista principal | `TbPlanificacion`, `TbPlanificacionEquipos`, `TbEquipos` | SQL dinámico concatenado (fechas `#…#`, textos entre comillas, ID sin comillas) en `RowSource` | — | `:181-243` |
| Combo subsistema | `TbEquipos`, `TbPlanificacionEquipos` | SQL dinámico concatenado con BUI entre comillas | — | `:545-549` |
| Combo equipo | ídem | SQL dinámico concatenado | — | `:572-583` |
| Combos BUI, ID y criterios | `TbBUI`, `TbPlanificacion`, lista de valores | `RowSource` guardado en el diseño | — | `.form.txt:1323, 2032, 1484` |
| Estado de la planificación | `TbPlanificacion` | DAO `OpenRecordset` con ID concatenado | — | `FG.bas:2056-2062` |
| Cadena de anexos | `TbPlanificacionAnexos` | DAO con ID concatenado | — | `FG.bas:2873-2877` |
| `IDEvento`, `IDNoEvento` | `TbPlanificacion` | `Dame`: SQL concatenado, reintento con comillas si falla | — | `FG.bas:300-310`; `:477, 489` |
| Informe | resultado de la lista | DAO `OpenRecordset` sobre SQL derivado + Excel | — | `FG.bas:1618-1619` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: literales de fecha `#mm/dd/yyyy#`, `Now()`, `IIf`, `IsNull`, `DateDiff('d', …)`, `Like '*'` y corchetes (`:181, 209, 231`) son dialecto Jet/ACE y requieren traducción a PostgreSQL; la lista se construye por concatenación con ID numérico sin comillas (`:199`) y textos entre comillas (`:184, 189, 194`). Generación del `.xlsx` mediante Automatización de Excel (`FG.bas:1625-1670`) no tiene equivalente directo en servidor.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ActualizarLista` pasa `ComboEquipoBusqueda` dos veces a `Filtrar` (`:29`), de modo que el segundo parámetro de entrada (`strID`) recibe el equipo y se concatena como `IDPlanificacion=<equipo>` (`:198-199`) además de `Equipo='…'` (`:194`): con un equipo seleccionado la condición sobre el ID podría dar error o ningún resultado.
- `Filtrar` no recibe `ComboCriterioPlanificacion` (`:29-30`): `strEstadoPlanificacion` queda en «Todas» (`:167-169`), que no figura entre las ramas (`:203-227`), por lo que `strParteWhereMotivoCierre` y `strParteWhereDifFechas` quedan vacías y el `WHERE` concatenado (`:235-242`) tendría `AND` consecutivos; además `ComboCriterioPlanificacion` y `ComboIDMantenimiento` no influyen en el resultado. Esto implicaría que la lista no puede refrescarse correctamente.
- `GenerarConsultaEnExcel` localiza con `InStr` (`:1012`, insensible a mayúsculas con `Option Compare Database`, `:2`) pero separa con `Split` por `"Where"` (`:1013`), mientras el SQL generado usa `WHERE` (`:235`): `dato(1)` podría no existir.
- `ImagenDocAnexado_Click`, con un solo anexo, abre `strURLAnexoEvento` en lugar de `strCadenaURLAnexos` (`:870`).
- Estado «Cerrada por reprogramación» habilita `ComandoTrabajoQueLaCierra` (`:946-947`) pero el clic solo admite «Cerrada por trabajo realizado» (`:468-471`).
- `ComandoVerDetalle` abre `FormPlanificacionCierre`, igual que `ComandoMarcarComoRealizada` (`:615, 415`), y su mensaje de error habla de «marcar como realizada» (`:622`).
- `ComandoLimpiarID` y `ComandoLimpiarEquipo` limpian el mismo combo (`:375-381, 342-348`); ningún botón limpia `ComboIDMantenimiento`.
- `ComandoActualizarTodo` no refresca si todos los filtros están vacíos (cada «Limpiar» sale antes, p. ej. `:310-312`).
- `ComboBUIBusqueda_Change` refresca la lista **antes** de que `CargarSubSistema` y `CargarIDEquipo` pongan a `Null` subsistema y equipo (`:635-637, 542-543, 568`), y `ComboSubsistemaBusqueda_Change` hace lo mismo con el equipo (`:725-726`): la lista podría quedar filtrada por valores ya borrados.
- `FechaInicialBusqueda` no tiene evento: cambiarla no refresca hasta salir de `FechaFinalBusqueda` (`.form.txt:1668-1697`; `:740-753`).
- `strIDPlanificacionSeleccionadoActualmente` se asigna pero no se usa para reseleccionar (`:23, 43, 49`).
- Los `RowSource` de diseño de `ListaFiltrados` y de los combos de subsistema y equipo usan `TbIDEquipos`, `EquipoID` y un valor literal de ejemplo (`.form.txt:1004-1013, 1285-1287, 1416-1420`), distintos de `TbEquipos` e `IDEquipo` del código (`:233-234, 546`).
- `ComandoAyudaFormulario` es un cuadro de texto con `OnClick` sin procedimiento (`.form.txt:2922-2961`).
- `ComandoCambiarEventoDeCierre_Click` usa `On Error Resume Next` (`:257`) y deja el reloj de arena encendido si algo falla.
- `Dame` y `EstadoPlanificacion` concatenan el ID sin comillas, aunque validado con `IsNumeric` (`FG.bas:2056-2062`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtro único por parámetros enlazados (sin SQL concatenado); incorporar el criterio de estado al filtro; un solo botón de limpiar; descarga de Excel generada en servidor.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir con `BUI|SUBSISTEMA|equipo` deja los tres filtros asignados (`:782-793`) y se muestra la lista filtrada.
2. Abrir con un argumento que contenga `|` y no tenga tres partes muestra «datos insuficientes» y cierra el formulario (`:774-777, 806`).
3. Elegir una planificación abierta habilita «Marcar como Realizada» y «Reprogramar» y deshabilita «Trabajo que la cierra» (`:938-941`).
4. Elegir una cerrada por trabajo realizado habilita «Trabajo que la cierra» y deshabilita las otras dos (`:942-945`).
5. «Planifi. Nuevo» sin selección muestra «Desde aquí sólo se puede generar una planificación a partir de otra existente» (`:85-88`).
6. Una fila con anexos muestra `ImagenDocAnexado`; una fila con evento cuyo `URLAnexos` no está vacío muestra `ImagenAnexoEvento` (`:967-987`).
7. Sin fechas, el filtro abarca de 1900 a 2200 (`:171-180`).
8. Informe en Excel con filtro sin resultados muestra «La consulta no contiene registros» (`FG.bas:1620-1623`).

## Pendientes

- Referencias dinámicas: nombres de formulario abiertos por texto (`:98, 128, 271, 415, 435, 509, 521, 840`) y controles de `FormEventoGestion` y `FormActividadGestion` por `Controls("…")` (`:510, 522`); contenido de `AjustarTamaño` (`FG.bas:261`), `FicheroAbierto`, `Ejecutar` y `Constructor.getEvento`.
- Si `Filtrar` con «Todas» produce SQL válido en Access (`:167-169, 235-242`).
- Tipo de dato y valores reales de `MotivoCierre` y de `IDEquipo` (ERD de origen).
- Valor de `DiasParaAvisoDePlanificaciones` en el entorno (`:218`) y ruta del directorio local del Excel (`FG.bas:1608`).
- Quién escribe `FechabajaParaPlanificacion` (`:574`).
