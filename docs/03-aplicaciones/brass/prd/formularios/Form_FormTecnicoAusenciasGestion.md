# Form_FormTecnicoAusenciasGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormTecnicoAusenciasGestion.cls` (490 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado de gestión de las ausencias («libranzas») de los técnicos, con filtro por año y por técnico y acciones de alta, edición, eliminación y exportación a Excel. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): el listado `ListaAusencias` es una lista de valores que se rellena con `AddItem` desde objetos `LIbranza` obtenidos de `TbTecnicosAusencias`. Abreviaturas: `Constructor.bas` = `modules/Constructor.bas`; `FG.bas` = `modules/Funciones Generales.bas`; `VG.bas` = `modules/Variables Globales.bas`. Las citas sin fichero son `Form_FormTecnicoAusenciasGestion.cls:línea`. Capacidad: CAP-9 Técnicos, ausencias, fiestas y precios.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormTecnicoGestion` (`ComandoLibranzas_Click`) | Asigna `m_ObjTecnicoActivo = m_ObjTecnicoSeleccionado`, cierra el formulario si estaba abierto y lo abre | `Form_FormTecnicoGestion.cls:251`, `Form_FormTecnicoGestion.cls:270-274` |

Entrada: la global `m_ObjTecnicoActivo` (`VG.bas:245`); si existe, preselecciona su alias y el año actual (`:390-393`); si no, abre sin filtro (`:395`). Salida: abre `FormTecnicoAusencias` (alta o edición) a través de `m_ObjLibranzaActiva` (`VG.bas:246`) y escucha su evento `Alta` para refrescar el listado (`:7`, `:469-491`). `cmdSalir` cierra sin guardar (`:415-418`).

## Controles y eventos

Controles: 16 (25 `Begin` en el `.form.txt`: 1 `Form`, 7 plantillas por defecto de tipo en la cabecera —líneas 141-249—, 1 `Section` y 16 controles).

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaAusencias` | cuadro de lista, 4 columnas, con encabezados, multiselección 2 (extendida) | `OnClick` → `ListaAusencias_Click` | Carga el `LIbranza` seleccionado y habilita Editar/Eliminar; con 0 o más de 1 elementos o sin ID, descarta la selección y deja ambos deshabilitados | `.form.txt:327-364`; `:419-465` |
| `ComandoAlta` | botón | `OnClick` | Resuelve el técnico (fila de la lista y, si no, `ComboTecnico`); abre `FormTecnicoAusencias` | `.form.txt:365`; `:57-92` |
| `ComandoEdicion` | botón (inicia deshabilitado) | `OnClick` | Abre `FormTecnicoAusencias` con la libranza seleccionada | `.form.txt:403-415`; `:94-128` |
| `ComandoEliminar` | botón (inicia deshabilitado) | `OnClick` | Pide confirmación y borra la libranza | `.form.txt:443-456`; `:210-259` |
| `ComboAño` | combo (lista de valores) | `AfterUpdate`, `Enter` | Filtra; `Enter` hace `Requery` | `.form.txt:514-533`; `:315-344` |
| `ComboTecnico` | combo (lista de valores) | `AfterUpdate` (pública), `Enter` | Filtra; `Enter` hace `Requery` | `.form.txt:573-592`; `:346-375` |
| `ComandoLimpiarAño` / `ComandoLimpiarTecnico` | 2 botones | `OnClick` | Ponen su combo a `Null` y refiltran | `.form.txt:738`, `.form.txt:918`; `:262-313` |
| `ComandoExportarExcel` | botón | `OnClick` | Exporta el contenido de la lista a Excel con `GenerarConsultas("FiltroAusencias", …)` | `.form.txt:606-616`; `:129-154` |
| `CmdSalir` | botón | `OnClick` → `cmdSalir_Click` | Cierra el formulario sin guardar | `.form.txt:657-666`; `:415-418` |
| `lblTitulo1`, `lblTitulo2` | 2 etiquetas «GESTIÓN DE LAS LIBRANZAS» | — | `lblTitulo1` da el título de ventana | `.form.txt:695-737`; `:385` |
| `Etiqueta142`, `Etiqueta4` | 2 etiquetas «AÑO» y «TÉCNICO» | — | Rótulos de los filtros | `.form.txt:487-499`, `.form.txt:546-558` |
| `CuadroSuperior`, `Cuadro2` | 2 rectángulos | — | Decoración | `.form.txt:281-290`, `.form.txt:303-313` |

## Flujos

### Apertura (`Form_Open`, `:376-414`)

1. `AjustarTamaño Me` (`FG.bas:261`) y título desde `lblTitulo1` (`:383-385`).
2. `RellenarCombos` (`:10-54`): `ComboAño` con los años distintos de `FechaLibranza` y `ComboTecnico` con los alias distintos ordenados (`:19`, `:35`). Si la tabla está vacía sale sin error y sin rellenar el segundo combo (`:22-26`).
3. Con técnico activo: fija `ComboTecnico` y `ComboAño = Year(Date)` y filtra (`:390-393`); sin técnico activo, `Filtrar , ,` (`:395`).
4. Cualquier error: `MsgBox` crítico y cierre del formulario (`:407-413`).

### Filtrado (`Filtrar`, `:156-209`)

Deshabilita Editar/Eliminar, fija las cabeceras `IDAusencia;Alias;Ausencia;Motivo` (`:170-174`), obtiene `Constructor.getLibranzas(año, técnico)` (`:175`) y añade una fila por libranza (`:182-187`). Si queda exactamente una fila de datos (`ListCount = 2`, con la cabecera) la selecciona y simula el clic (`:189-191`); si no, repite el clic sobre la primera fila seleccionada (`:193-199`). El objeto `m_Form` (`FormTecnicoAusencias`) refresca el listado al lanzar `Alta` (`:469-480`).

### Alta y edición (`:57-128`)

- Alta: `m_ObjLibranzaActiva = Nothing`; técnico = columna 1 de la fila seleccionada o, si no hay, el de `ComboTecnico`; si ninguno, «No se conoce el técnico activo» (`:63-77`). Cierra `FormTecnicoAusencias` si estaba abierto, lo abre y engancha `m_Form` (`:78-83`).
- Edición: exige fila seleccionada («Ha de seleccionar un elemento de la lista») y la carga con `getLibranza` (`:99-113`); mismo cierre/apertura (`:115-120`).

### Eliminación (`:210-259`)

Carga la libranza si no estaba (`:216-229`), pregunta «¿Desea realmente eliminar la libranza seleccionada?» con «No» por defecto (`:233-236`), llama a `Borrar` y refiltra con los combos actuales (`:240-244`).

### Exportación (`:129-154`)

`GenerarConsultas` rechaza lista vacía («La consulta no contiene registros») y exige exactamente uno de SQL o lista (`FG.bas:1384-1440`); vuelca cada celda de la lista (con `Nz(…, "null")`) a un libro `.xlsx` en el directorio local configurado (`FG.bas:1447-1448`, `FG.bas:1533-1535`). Un error devuelto como `#ERR|texto` se muestra tal cual (`:139-143`).

### Errores

Todos los handlers acumulan el texto en `m_Error` y lo muestran con `MsgBox` crítico (`:86-91`, `:253-258`); el error interno 1000 conserva el mensaje previo.

## Reglas, validaciones y cálculos

- **Selección única:** Editar/Eliminar solo se habilitan con exactamente una fila válida (`:427-434`, `:456-457`).
- **Confirmación** de borrado con «No» por defecto (`:233`).
- **Filtro:** año + técnico, solo técnico, solo año o ninguno (`Constructor.bas:4135-4161`; ver defectos).
- **Orden:** `FechaLibranza DESC` (`Constructor.bas:4141`, `Constructor.bas:4148`).
- **Duplicados:** el diccionario de libranzas ignora un `IDAusencia` repetido (`Constructor.bas:4177-4184`).
- **Permisos:** no hay comprobación de rol en este formulario; cualquier usuario que lo abra puede borrar (a verificar fuera de él).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Años de los filtros | `TbTecnicosAusencias` | SQL `SELECT DISTINCT Year(...)` fija, `CurrentDb()` | — | `:19-20` |
| Alias de los filtros | `TbTecnicosAusencias` | SQL `SELECT DISTINCT` fija | — | `:35-36` |
| Listado | `TbTecnicosAusencias` | SQL dinámico concatenado (alias entre comillas, año sin comillas) vía `getdb()` | — | `Constructor.bas:4135-4162` |
| Lectura de una libranza | `TbTecnicosAusencias` | SQL dinámico concatenado por `IDAusencia` | — | `Constructor.bas:4086-4093` |
| Lectura del técnico | `TbTecnicos` | SQL dinámico con alias entre comillas | — | `Constructor.bas:1950-1951` |
| Borrado | `TbTecnicosAusencias` | `DELETE … WHERE IDAusencia=` concatenado, `CurrentDb().Execute` | no | `LIbranza.cls:264-283` |

La base `getdb()` es la del backend vinculado y abre con una contraseña literal en el código (`VG.bas:548-575`), mientras que los combos usan `CurrentDb()` (`:20`, `:36`).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL concatenado con fechas/alias; `Year([FechaLibranza])` y `DISTINCT` (en PostgreSQL, `EXTRACT(YEAR FROM …)`); `DELETE tabla.campo FROM` es dialecto Jet (`LIbranza.cls:274-276`); el filtrado por año compara con un literal numérico sin conversión de tipo; `getdb()` y `CurrentDb()` pueden apuntar a orígenes distintos con tablas vinculadas.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `getLibranzas` no define `m_SQL` cuando año y técnico llegan vacíos: las condiciones no cubren ese caso (`Constructor.bas:4135-4161`). `Form_Open` sin técnico activo invoca ese caso (`:395`) y `OpenRecordset` recibiría una cadena vacía.
- La rama «sin técnico, con año» (`Constructor.bas:4149-4155`) concatena `Alias =''` y carece de `AND`; además la cuarta condición (`Constructor.bas:4156`) repite la tercera y es inalcanzable. `ComandoLimpiarTecnico` con año elegido llega a esa rama (`:298`).
- `ComandoLimpiarAño` y `ComboAño_AfterUpdate` pasan `ComboAño.Column(0)`; tras `Null`, `Nz` devuelve cadena vacía (`:270-271`).
- `ComboAño` define un `RowSource` residual en el `.form.txt` (`.form.txt:531`) que `RellenarCombos` vacía (`:18`).
- Años sin `ORDER BY` en `RellenarCombos` (`:19`), por lo que el orden de la lista no está garantizado.
- `ComandoEliminar_Click` tras borrar lanza `Err.Raise 100` si `Filtrar` devuelve error, sin pasar por el mensaje de 1000 (`:244-247`); y el borrado deja `m_ObjLibranzaSeleccionada` apuntando a la libranza ya eliminada (`:240-244`), que solo se reinicia en `ListaAusencias_Click` si cambia la fila (`:439-450`).
- `ComandoAlta_Click` usa `On Error Resume Next` antes de enganchar `m_Form` (`:82-83`); si falla, el refresco tras el alta no se produciría.
- El rótulo de `ComandoExportarExcel` es «Comando227» y el de ambos botones de limpiar «Comando36» (`.form.txt:615`, `.form.txt:748`, `.form.txt:928`); la variable `pregunta` y otras no se declaran localmente, pues se asumen globales (`:91`, `:138`).
- `ListaAusencias` admite multiselección extendida (`.form.txt:327-364`) pero el formulario solo trabaja con una fila.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtrado con parámetros en servidor; borrado con confirmación y comprobación de permisos; exportación generada en servidor.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir desde la gestión de técnicos con un técnico seleccionado, se muestra su alias y el año actual y solo sus libranzas de ese año (`:390-393`).
2. Editar y Eliminar permanecen deshabilitados hasta que haya una única fila seleccionada (`:424-457`).
3. Alta sin fila seleccionada ni técnico elegido muestra «No se conoce el técnico activo» (`:74-77`).
4. Eliminar pide confirmación y, si se acepta, la fila desaparece del listado (`:233-244`).
5. Tras guardar un alta en `FormTecnicoAusencias`, el listado se refresca con los filtros vigentes (`:477`).
6. Exportar con la lista vacía informa «La consulta no contiene registros» (`FG.bas:1384-1440`).
7. Filtrar solo por año, y quitar el técnico con un año elegido, devuelven resultados sin error (ver defectos).

## Pendientes

- Campos y tipos de `TbTecnicosAusencias` (`ColCampos` de `LIbranza`) y cómo persiste `FormTecnicoAusencias` (ficha propia).
- Contenido de `AjustarTamaño` más allá de `FG.bas:261-275`.
- Si el `RowSource` residual de `ComboAño` (`.form.txt:531`) se usa antes de `RellenarCombos`.
- Código duplicado en el `.form.txt` (`.form.txt:1238-1250`): confirmar que no es una segunda copia ejecutable.
