# Form_FormPlanificacionAlta

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormPlanificacionAlta.cls` (208 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta de una planificación semanal (mantenimiento preventivo) para un equipo ya elegido: el usuario indica año, mes y semana y el formulario crea un registro de `TbPlanificacion`. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); el equipo llega por `OpenArgs`. Toda la lógica de negocio está en `FG.bas` (= `modules/Funciones Generales.bas`): `AltaPlanificacion`, `FechaViernes`, `EquipoPlanificable`, `PlanificacionYaExistente`, `DameID`. Las citas son `Form_FormPlanificacionAlta.cls:línea` salvo indicación; capacidad de todos los controles: CAP-6 Planificación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormPlanificacionesBusqueda` (`ComandoAltaPlanificacion_Click`) | Exige una planificación seleccionada; compone `BUI\|SUBSISTEMA\|IDEquipo\|Equipo` con las columnas 7, 8, 9 y 1 de `ListaFiltrados`, cierra el formulario si ya estaba abierto y lo abre con ese `OpenArgs` | `Form_FormPlanificacionesBusqueda.cls:78-98` |

`OpenArgs` debe tener exactamente cuatro partes separadas por `|` (`:166-170`); la tercera (`IDEquipo`) debe ser numérica (`:175-178`). Salida: tras registrar, refresca `FormPlanificacionEquipos` (`Form_FormPlanificacionEquipos.ActualizarLista`, `:104-106`) y `ListaFiltrados` de `FormPlanificacionesBusqueda` (`:112-114`) si están abiertos, y se cierra sin guardar diseño (`:115`). No lanza eventos propios.

## Controles y eventos

Controles: 25 en la sección Detalle (34 `Begin` en el `.form.txt`: 1 `Form`, 7 plantillas de estilo por defecto en `Form_FormPlanificacionAlta.form.txt:141-244`, 1 `Section` y los 25).

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `EQUIPO`, `SUBSISTEMA`, `BUI` | texto no enlazado | — | Solo informativos; se rellenan en `Form_Load` si el argumento no está vacío. `EQUIPO` figura con `Enabled`/`Locked` distintos del valor por defecto | `:179-187`; `Form_FormPlanificacionAlta.form.txt:547-565` |
| `AÑO` | texto, `Tag` «OBLIGATORIO» | `OnExit` → `Año_Exit` | Recalcula la lista de semanas | `:48-57`; `Form_FormPlanificacionAlta.form.txt:690-692` |
| `ComboMes` | combo lista de valores 1-12, `LimitToList` | `OnExit` → `ComboMes_Exit` | Recalcula la lista de semanas | `:131-140`; `Form_FormPlanificacionAlta.form.txt:736, 754-759` |
| `ComboSemana` | combo lista de valores, `LimitToList` | `OnExit` → `ComboSemana_Exit` | Handler sin efecto (solo cursor de espera) | `:141-149`; `Form_FormPlanificacionAlta.form.txt:801, 819-824` |
| `ComandoAlta` (rótulo «Registrar») | botón | `OnClick` → `ComandoAlta_Click` | Valida, da de alta y cierra | `:58-126`; `Form_FormPlanificacionAlta.form.txt:357-359` |
| `cmdSalir` | botón | `OnClick` → `cmdSalir_Click` | Cierra sin guardar | `:127-130`; `Form_FormPlanificacionAlta.form.txt:400-402` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título; `Form_Load` copia `lblTitulo1` a `Me.Caption` | `:159`; `Form_FormPlanificacionAlta.form.txt:847, 867` |
| `Etiqueta191`, `192`, `197`, `242`, `245`, `249` | 6 etiquetas | — | Rótulos EQUIPO, SUBSIST., BUI, AÑO, MES, SEMANA | `Form_FormPlanificacionAlta.form.txt:489-785` |
| `Etiqueta0`, `Etiqueta1`, `Etiqueta6` | 3 etiquetas «*» | — | Marca de campo obligatorio | `Form_FormPlanificacionAlta.form.txt:888, 910, 971` |
| `ComandoAyudaFormulario` | botón con imagen | sin `OnClick` | Sin handler en el `.cls` | `Form_FormPlanificacionAlta.form.txt:931` |
| `Cuadro69`, `CuadroSuperior`, `Cuadro2`, `Marco175`, `Marco189` | 3 rectángulos y 2 grupos de opciones | — | Decorativos | `Form_FormPlanificacionAlta.form.txt:285, 308, 331, 443, 465` |

## Flujos

### Apertura (`Form_Load`, `:150-201`)

1. `AjustarTamaño Me` (`FG.bas:261`) y título (`:157-159`).
2. Valida `OpenArgs` y rellena `EQUIPO`, `BUI` y `SubSistema` (`:161-187`).
3. Error: `MsgBox` crítico y cierre del formulario (`:194-200`).

### Selección de periodo (`EstablecerListaSemana`, `:5-47`)

1. Si `AÑO` o `ComboMes` no son numéricos, el `RowSource` de `ComboSemana` queda en `#ERR` (`:13-14, 35`).
2. Si lo son, `FechaViernes(mes, año)` con semana 0 devuelve una línea `NN fecha` por semana con viernes; se sustituyen los saltos por `;` y se asigna como `RowSource` (`:26-36`; `FG.bas:659-703`).
3. Error de `FechaViernes` (cadena con `|`): `MsgBox` con el detalle (`:27-31, 41-46`).

### Registro (`ComandoAlta_Click`, `:58-126`)

1. Comprueba `strIDEquipo`, año numérico, mes numérico y semana con al menos 3 caracteres y los dos primeros numéricos (`:64-86`); toma `Left(semana, 2)` (`:87`).
2. `AltaPlanificacion(IDEquipo, semana, mes, año)` (`:93`; `FG.bas:2238-2396`). Si devuelve `#ERR|texto`, lo muestra (`:94-98`).
3. Refresca los formularios abiertos y se cierra (`:104-115`).
4. Cualquier error: `MsgBox` crítico y el formulario sigue abierto (`:120-125`).

## Reglas, validaciones y cálculos

- **Equipo planificable:** debe existir en `TbPlanificacionEquipos` y su `FechabajaParaPlanificacion` no debe ser una fecha (`IsDate` ⇒ «No»); si no, «Este equipo no está marcado como planificable» (`FG.bas:2278-2287, 2138-2165`).
- **Semana válida:** entre 1 y el número de semanas del mes (`NumeroSemanasTieneElMes`) y con viernes; la fecha prevista de cierre es ese viernes (`FG.bas:2297-2326`, `FG.bas:2662`).
- **Duplicado:** no puede existir otra planificación del mismo equipo, semana, mes y año (`FG.bas:2334-2343, 2210-2217`).
- **Identificador:** `DameID("TbPlanificacion", "IDPlanificacion")` = máximo + 1 (`FG.bas:2352-2362, 371-420`).
- **Campos escritos:** `IDPlanificacion`, `FechaRegistro = Now()`, `IDEquipo`, `Anio`, `Mes`, `Semana`, `FechaPrevistaCierre` (`FG.bas:2372-2381`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| ¿Equipo planificable? | `TbPlanificacionEquipos` | `Dame(...)`, SQL dinámico | no | `FG.bas:2138-2155, 290-300` |
| ¿Planificación existente? | `TbPlanificacion` | SQL dinámico concatenado con equipo, semana, mes y año | no | `FG.bas:2210-2217` |
| Siguiente ID | `TbPlanificacion` | recorrido completo del campo ID en VBA | no | `FG.bas:409-411` |
| Comprobación del ID | `TbPlanificacion` | SQL dinámico concatenado | no | `FG.bas:2363-2366` |
| Alta | `TbPlanificacion` | DAO `AddNew`/`Update` | no | `FG.bas:2372-2381` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: identificador `Max + 1` calculado en VBA sin transacción (`FG.bas:371-420`); fechas construidas con `CDate("01/" & mes & "/" & año)` (`FG.bas:636`, `FG.bas:2682`), sensibles a la configuración regional; `Weekday(..., vbMonday)` y `DateAdd`/`DateDiff` de VBA (`FG.bas:637-642`); tabla `TbPlanificacion` con columna `Anio` sin eñe.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La validación `intMes < 1 And intMes > 12` nunca puede ser cierta (`FG.bas:623`); el rango se comprueba después en `AltaPlanificacion` (`FG.bas:2288-2291`).
- Dos altas simultáneas podrían obtener el mismo ID y la comprobación posterior no lo impide de forma atómica (`FG.bas:2352-2371`).
- Al cambiar año o mes, la lista de semanas se regenera pero no se vacía el valor ya elegido en `ComboSemana` (`:35-36`); la semana antigua podría registrarse con el periodo nuevo.
- Textos de error incompletos: «No se ha introducido una válida» (`:80, 84`) y «una mes no válido» (`:76`).
- `Año_Exit` y `ComboMes_Exit` usan `On Error Resume Next` (`:49, 132`): un fallo al recalcular quedaría silenciado.
- Con `Option Compare Database` y fechas de texto, `FechaViernes` depende del formato regional (`FG.bas:636`).
- Comentario obsoleto en `FG.bas:604` (cita un método `CerrarPlanificacionSinTrabajos` que este formulario no contiene).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: identificador generado por la base de datos y restricción única (equipo, año, mes, semana); borrar la semana al cambiar el periodo.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin `OpenArgs` válidos muestra «No se ha abierto el formulario con los parámetros adecuados» y cierra el formulario (`ComandoAlta`, `Form_Load`).
2. Elegir año y mes numéricos lista solo las semanas con viernes, con su fecha (`AÑO`, `ComboMes`, `ComboSemana`).
3. Registrar para un equipo no planificable muestra «Este equipo no está marcado como planificable» y no escribe.
4. Registrar una semana ya planificada para el equipo muestra «ya está planificada para este equipo» y no escribe.
5. Un alta correcta crea un registro con `FechaPrevistaCierre` igual al viernes de la semana y refresca las listas abiertas.
6. `cmdSalir` cierra sin escribir.

## Pendientes

- Estructura de `TbPlanificacion` y `TbPlanificacionEquipos`, índices y tipos (ERD de origen).
- Tramo `FG.bas:2113-2137` de `EquipoPlanificable` (comprobaciones previas a la lectura de `TbPlanificacionEquipos`) no analizado en esta ficha.
- Acción prevista de `ComandoAyudaFormulario` (sin handler, `Form_FormPlanificacionAlta.form.txt:931`).
