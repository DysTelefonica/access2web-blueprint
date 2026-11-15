# Form_FormIndicadores

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormIndicadores.cls` (155 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Panel «INDICADORES» de NC de proyecto: para un semestre y un año calcula cuatro recuentos sobre `TbNoConformidades` (abiertas en el periodo, replanificadas, stock activo y con riesgo asociado) y exporta el detalle de cada uno a Excel. Es un formulario **no enlazado** (sin `RecordSource`, vista única `DefaultView =0`, `.form.txt:14`); la lógica vive en la clase `IndicadorServicio` y el módulo `IndicadorRepositorio`. Las citas son `Form_FormIndicadores.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`, `IS.cls` abrevia `classes/IndicadorServicio.cls` e `IR.bas`, `modules/IndicadorRepositorio.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_Form0BDOpcionesParteProyectos | Botón «Indicadores» (`DoCmd.OpenForm "FormIndicadores"`) | `Form_Form0BDOpcionesParteProyectos.cls:156-170` |

No abre otros formularios. Produce ficheros Excel en el directorio local de informes y los abre con `ShellExecute` si el usuario acepta (`:123, 130-131`; `FG.bas:4399-4400`). «Salir» cierra el formulario (`:10-12`).

## Controles y eventos

`.form.txt` declara 40 bloques `Begin`: el formulario (`.form.txt:4`), 11 de valores por defecto (`.form.txt:406-478`), 3 secciones (`EncabezadoDelFormulario` `.form.txt:490`, `Detalle` `.form.txt:568`, `PieDelFormulario` `.form.txt:1836`) y los 25 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `cmbSemestre` | combo, lista de valores `S1;S2` | — | Semestre; en la carga, `S1` hasta junio y `S2` desde julio | `:57-61`, `.form.txt:662-664` |
| `txtAnio` | cuadro de texto, `DefaultValue =Year([Date])` | — | Año; la carga lo fija al año en curso | `:56`, `.form.txt:691-692` |
| `btnCalcular` | botón | `OnClick` | Calcula el periodo y escribe los cuatro recuentos | `:74-91`, `.form.txt:609-611` |
| `btnExcelAbiertas`, `btnExcelReplan`, `btnExcelStock`, `btnExcelRiesgo` | botones (texto «Calcular», tip «Aplicar filtro») | `OnClick` | `ExportarDirecto` con `ABIERTAS`, `REPLAN`, `STOCK` o `RIESGO` | `:141-155`, `.form.txt:886-888, 1136-1138, 1386-1388, 1636-1638` |
| `lblResAbiertas`, `lblResReplan`, `lblResStock`, `lblResRiesgo` | etiquetas | — | Resultado de cada recuento; `-` tras la carga | `:67-72, 85-88`, `.form.txt:864, 1114, 1364, 1614` |
| `cmdAyuda` | botón (tip «Grabar la decisión») | `OnClick` | Copia la ayuda HTML del servidor al directorio local y la abre | `:14-35`, `.form.txt:537-541` |
| `btnSalir` | botón (tip «Grabar la decisión») | `OnClick` | `DoCmd.Close` sin guardar | `:10-12`, `.form.txt:1856-1860` |
| `lblTitulo` | etiqueta | — | Título «INDICADORES», copiado a la ventana | `:53-54`, `.form.txt:510` |
| `lblTit1`-`lblTit4`, `Etiqueta59` («Semestre»), `Etiqueta121` («Año») | etiquetas | — | Rótulos fijos | `.form.txt:715, 735, 777, 1093, 1343, 1593` |
| `shpParametros`, `boxFondo1`-`boxFondo4` | rectángulos | — | Decorativos | `.form.txt:586, 756, 798, 820, 842` |

## Flujos

### Carga (`Form_Load`, `:50-65`)

Crea `IndicadorServicio`, `Ajustar Me`, fija título, año en curso y semestre por mes actual, y pone `-` en los cuatro resultados (`:51-64`).

### Cálculo (`btnCalcular_Click`, `:74-96`)

1. Periodo: `S1` = 1 de enero-30 de junio; cualquier otro valor = 1 de julio-31 de diciembre (`IS.cls:26-36`). Sin año, el año en curso (`:82`).
2. Cada recuento envuelve la SQL del indicador en `SELECT Count(*) AS Total FROM (…) AS T` (`IS.cls:55-63`); si falla devuelve `-1` (`IS.cls:72-74`).
3. Escribe los cuatro valores en las etiquetas (`:85-88`).

### Exportación (`ExportarDirecto`, `:100-139`)

1. Recalcula el periodo (`:111`) y obtiene la colección de objetos `NCProyecto` del indicador (`:114`; `IR.bas:66-117`).
2. Si no hay datos, lanza el error 1000 «No hay datos.» (`:117-118`).
3. `GestionNCProyectosAExcel` genera `Informe_<TIPO>_<fecha-hora>.xlsx` en el directorio local de informes (`:123`; `FG.bas:4367, 4399-4400`); pregunta «Informe generado. ¿Abrir?» y lo abre (`:130-131`).

### Ayuda (`cmdAyuda_Click`, `:14-48`)

Comprueba que existe la plantilla HTML de ayuda en la ruta de recursos configurada (`:23-27`; `Entorno.cls:849-870`), la copia al directorio local sobrescribiendo (`:28-29`) y la abre (`:35`). Si no existe, «No se encuentra la ayuda en el servidor».

### Errores

Cálculo: `MsgBox "Error al calcular: …"` sin correo al administrador (`:93-95`). Exportación: los errores distintos de 1000 se muestran como información (`:136-138`); los 1000 no se muestran. Ayuda: error 1000 o no controlado ⇒ `MsgBox` crítico; el no controlado además llama a `CorreoAlAdministrador` (`:40-47`).

## Reglas, validaciones y cálculos

Las cuatro consultas exigen `Borrado = False` (solo NC no borradas):

- **NC abiertas:** `FECHAAPERTURA` entre inicio y fin, ambos incluidos (`IR.bas:11-15`).
- **Replanificadas:** NC distintas con alguna `TbReplanificacionesProyecto.FechaReprogramacion` en el periodo (`IR.bas:21-27`).
- **Stock activo:** `FECHAAPERTURA` anterior al inicio y sin `FECHACIERRE` o cerrada a partir del inicio; la fecha fin no interviene (`IR.bas:33-38`).
- **Con riesgo asociado:** NC cuya `Max(FechaDecison)` en `TbRiesgosNC` cae en el periodo (`IR.bas:44-58`).
- El parámetro de identificadores de proyecto se pasa con el semestre y el servicio lo ignora (`:80, 85-88`; `IS.cls:80-104`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Recuentos | `TbNoConformidades`, `TbReplanificacionesProyecto`, `TbRiesgosNC` | SQL dinámico concatenado con literales `#mm/dd/yyyy#` envuelto en `Count(*)` | `IR.bas:13-58`; `IS.cls:59-61` |
| Detalle para Excel | las mismas | SQL dinámico concatenado, hidratación de `NCProyecto` | `IR.bas:75-113` |
| Informe Excel | fichero local `.xlsx` | automatización COM de Excel | `FG.bas:4385, 4399-4400` |
| Ayuda | fichero HTML copiado al directorio local | `FileSystemObject.CopyFile` | `:28-29` |

Solo lectura de datos. Sin transacciones.

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Ya se ve: literales de fecha Jet `#mm/dd/yyyy#` (`IR.bas:15, 27, 37-38, 56-57`), `Borrado = False` sobre campo Sí/No (`IR.bas:14`), subconsulta en `FROM` con alias obligatorio (`IS.cls:59`) y `Max` en `HAVING` (`IR.bas:56-57`); en PostgreSQL vinculado conviene parametrizar fechas y comprobar la traducción de booleanos.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El aviso «No hay datos.» nunca se muestra: se lanza como error 1000 (`:117-118`) y el manejador solo muestra los distintos de 1000 (`:138`); también se ocultan los errores del repositorio (`:116`).
- Un fallo de SQL en un recuento se muestra como `-1` sin aviso (`IS.cls:72-76`).
- Las asociaciones de riesgo se borran y reinsertan con la fecha del día en cada guardado (`RiesgoRepositorio.bas:96-106`): el indicador «con riesgo» cuenta la NC en el periodo de su último guardado, no en el de la decisión original.
- El comentario de `SQL_NC_ConRiesgo` dice usar `FechaRegistro` y la SQL usa `FechaDecison` (`IR.bas:47-57`).
- Los tips de `cmdAyuda` y `btnSalir` dicen «Grabar la decisión» y los de exportación «Aplicar filtro» (`.form.txt:541, 890, 1860`).
- Un año no numérico en `txtAnio` provoca error de tipo al pasar a `Integer` (`:82`; `IS.cls:26`), mostrado solo como «Error al calcular».

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: indicadores calculados en servidor con periodo libre; exportación CSV/XLSX descargable en lugar de Excel local; filtro por proyecto (hoy ignorado).

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, año en curso, semestre según el mes y los cuatro resultados a `-`.
2. «Calcular» con `S1` de un año cuenta solo NC no borradas con apertura entre el 1 de enero y el 30 de junio de ese año.
3. El stock activo cuenta NC abiertas antes del inicio del periodo y no cerradas antes de él.
4. Cada botón de exportación genera un Excel con las NC del indicador correspondiente y ofrece abrirlo.
5. Sin plantilla de ayuda, «Ayuda» avisa «No se encuentra la ayuda en el servidor».

## Pendientes

- Columnas y formato exactos del Excel por tipo de informe (`FG.bas:4367-4400`, `blnEsReplan`/`blnEsRiesgo`); no trazados en esta ficha.
- Si `FECHAAPERTURA` guarda hora, el límite `<= #fin#` excluiría el último día (`IR.bas:15`); depende del esquema (pendiente del ERD de origen).
