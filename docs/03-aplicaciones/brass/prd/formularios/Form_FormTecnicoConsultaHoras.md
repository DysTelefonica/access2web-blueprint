# Form_FormTecnicoConsultaHoras

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormTecnicoConsultaHoras.cls` (450 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Informe de horas por técnico y día: el usuario indica un rango de fechas y, opcionalmente, un técnico; el formulario genera un libro Excel local con una fila por técnico y día (horas laborables, horas extra, actividades, tipo de día, libranza, festivo y observaciones de incoherencia). Es un formulario **no enlazado** (sin `RecordSource` en `Form_FormTecnicoConsultaHoras.form.txt`). Los datos salen de `TecnicoHoras.cls` (clase de `src/classes`), que consulta actividades, ausencias y festivos. Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`, `VG.bas` = `modules/Variables Globales.bas`. Las citas sin fichero son `Form_FormTecnicoConsultaHoras.cls:línea`. Capacidad: CAP-9 Técnicos, ausencias, fiestas y precios (el informe se apoya en actividades, CAP-4).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormTecnicoGestion` | Cierra el formulario si estaba abierto, fija `m_ObjTecnicoActivo` con el técnico seleccionado y lo abre | `Form_FormTecnicoGestion.cls:230-234` |
| `Form_FormActividadGestion` | Misma secuencia con su técnico seleccionado | `Form_FormActividadGestion.cls:567-571` |

El formulario no devuelve valores ni lanza eventos. Salida: un fichero `.xlsx` en el directorio local de la aplicación (`Entorno.URLDirectorioLocal`, bajo la carpeta de datos del usuario; `Entorno.cls:191`), que se abre con la aplicación asociada mediante `Ejecutar` (`:433`; API de shell declarada en `VG.bas:15`). `cmdSalir` cierra sin guardar (`:9-12`).

## Controles y eventos

Controles: 15 (24 `Begin` en el `.form.txt`: 1 `Form`, 7 plantillas de estilo por defecto en `:141-244`, 1 `Section` y 15 controles).

| Control | Tipo | Evento y handler | Qué hace / origen de datos / validación | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `FechaInicial` | cuadro de texto no enlazado, formato «Short Date» | — | Fecha inicial; obligatoria y debe ser fecha válida (`:108-111`) | CAP-9 | `Form_FormTecnicoConsultaHoras.form.txt:371-372`; `:28` |
| `FechaFinal` | cuadro de texto no enlazado, formato «Short Date» | — | Fecha final; obligatoria, fecha válida y no anterior a la inicial (`:112-119`) | CAP-9 | `Form_FormTecnicoConsultaHoras.form.txt:399-400`; `:29` |
| `ComboTecnico` | cuadro combinado, lista de valores, 3 columnas (alias; nombre; tipo), `LimitToList`, etiqueta `Tag` «OBLIGATORIO» | — (relleno en `Form_Load`) | Técnico; vacío significa todos los técnicos activos (`:120-125`) | CAP-9 | `Form_FormTecnicoConsultaHoras.form.txt:454-478`; `:30, 71`; `FG.bas:7224-7243` |
| `ComandoInforme` (rótulo «Informe ») | botón | `OnClick` → `ComandoInforme_Click` | Valida, construye la colección de técnico-día y genera el Excel | CAP-10 | `:17-57`; `Form_FormTecnicoConsultaHoras.form.txt:424-426` |
| `cmdSalir` (rótulo «Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra el formulario sin guardar | CAP-9 | `:9-12`; `Form_FormTecnicoConsultaHoras.form.txt:330-332` |
| `lblEstado` | etiqueta, oculta por defecto | — | Línea de progreso; `Avance` la hace visible y le escribe el texto del paso | CAP-10 | `Form_FormTecnicoConsultaHoras.form.txt:658-672`; `FG.bas:7262-7280`; `:75, 45` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título «CONSULTA DE LAS HORAS DE UN TÉCNICO» (dos etiquetas; `lblTitulo1` da el `Caption` del formulario) | sin asignar | `Form_FormTecnicoConsultaHoras.form.txt:582-603`; `:70` |
| `Etiqueta142`, `Etiqueta226`, `Etiqueta227` | etiquetas «F.INICIAL», «F.FINAL», «TÉCNICO» | — | Rótulos de los campos | sin asignar | `Form_FormTecnicoConsultaHoras.form.txt:503, 530, 557` |
| `Etiqueta4`, `Etiqueta1` | etiquetas «*» | — | Marca de obligatorio | sin asignar | `Form_FormTecnicoConsultaHoras.form.txt:622, 644` |
| `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | sin asignar | `Form_FormTecnicoConsultaHoras.form.txt:285, 308` |

## Flujos

### Apertura (`Form_Load`, `:58-90`)

1. Cursor de espera y `AjustarTamaño Me` (`:62, 66`; `FG.bas:261`).
2. Si hay técnico activo (`m_ObjTecnicoActivo`, `VG.bas:245`), asigna su alias a `ComboTecnico` (`:67-69`).
3. `Caption` del formulario = `lblTitulo1.Caption` (`:70`).
4. `RellenarComboTecnicos` con `p_EnTodos = No` y `p_SoloEnActivo = Sí`: solo técnicos sin fecha de baja (`:71`; `FG.bas:7224-7243`; `Constructor.bas:4266-4268`).
5. Oculta `lblEstado` (`:75`). Error: `MsgBox` crítico y el formulario queda abierto (`:84-89`).

### Generación del informe (`ComandoInforme_Click`, `:17-57`)

1. Lee fechas y técnico con `Nz` (`:28-30`).
2. `getTecnicosHoras` (`:92-167`): valida fechas (`:108-119`); sin alias toma `Entorno.ColTecnicosActivos` (`:121`; `Entorno.cls:693-707`); con alias lo carga con `Constructor.getTecnico` y falla si no existe (`:127-134`).
3. Para cada día `i` de `0` a `DateDiff(...) - 1` y cada técnico crea un `TecnicoHoras` con alias y fecha (`:142-157`).
4. `PasarAExcel` (`:172-449`): nombre `Informe de Horas_<h>_<m>_<s>.xlsx` en el directorio local (`:208-209`); si ya existe y está abierto, «Tiene una consulta abierta» (`:210-214`), si no lo borra (`:215`). Crea una instancia oculta de Excel (`:217-218`), escribe la cabecera de 10 columnas (`:225-234`) y una fila por `TecnicoHoras` (`:311-322`).
5. Salta los sábados y domingos (`:265-269`). Para el resto toma `TipoDia`, `LIbranza`, `Festivo`, `Hlab`, `HExt` y las cadenas de IDs de actividad (`:245-306`).
6. Da formato (negrita en cabecera, bordes, columnas autoajustadas; `:363-428`), guarda, cierra Excel (`:429-432`) y abre el fichero (`:433`).
7. Oculta `lblEstado` y restablece el cursor (`:45-48`).

### Errores

Cualquier error se acumula en `m_Error` (`:5`), se muestra en `MsgBox` crítico y restablece el cursor (`:51-56`). `PasarAExcel` cierra libro y Excel en su manejador (`:440-447`). No se envía correo.

## Reglas, validaciones y cálculos

- **Rango:** fechas obligatorias y `FechaFinal >= FechaInicial` (`:108-119`). Mensajes: «No se ha indicado la fecha inicial/Final» y «La fecha final no puede ser anterior a la inicial».
- **Técnicos:** vacío ⇒ técnicos activos del entorno (con caché en `Entorno`); informado ⇒ solo ese alias.
- **Tipo de día** (`TecnicoHoras.cls:58-108`): «No Laborable» si sábado/domingo (`:79`) o si hay festivo (`:86-91`); en otro caso «Laborable».
- **Libranza:** existe un registro del técnico y la fecha en `TbTecnicosAusencias` (`TecnicoHoras.cls:384-417`). **Festivo:** existe la fecha en `TbTecnicosFiestas` (`TecnicoHoras.cls:419-449`), sin distinguir técnico.
- **Horas:** `Hlab` y `HExt` suman `HorasLaborables` y `HorasExtras` de las actividades del técnico y día (`TecnicoHoras.cls:206-247, 248-289`). `m_Horas = HorasLab + HorasExt` (`:306`).
- **Observaciones** (`:323-337`, fila en rojo y negrita `:338-343`): laborable sin libranza y `m_Horas <> 7.5` ⇒ «Día Laborable y no cumple las 7.5 horas registradas»; libranza con horas ⇒ «Día de Libranza con horas registradas»; no laborable con horas ⇒ «Día no laborable con horas registradas».
- **Cadena de IDs:** primero las actividades con horas laborables > 0 y después las de horas extra, separadas por `;` (`:280-289`; `TecnicoHoras.cls:290-336`).
- **Fecha en Excel:** `Format(m_Fecha, "mm/dd/yyyy")` (`:312`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de técnicos activos (combo) | `TbTecnicos` | SQL dinámico (`WHERE FECHABAJA Is Null`) | — | `Constructor.bas:4266-4278` |
| Técnicos activos por tipo (caché del entorno) | `TbTecnicos` | SQL dinámico concatenado | — | `Entorno.cls:703`; `Constructor.bas:4334-4355` |
| Técnico por alias | `TbTecnicos` | SQL dinámico con el alias entre comillas | — | `Constructor.bas:1940-1953` |
| Actividades de técnico y día | `TbActividades` | SQL dinámico con alias y fecha `#mm/dd/yyyy#` | — | `Constructor.bas:4414-4419` |
| Libranza | `TbTecnicosAusencias` | SQL dinámico con alias y fecha `#...#` | — | `TecnicoHoras.cls:399-402` |
| Festivo | `TbTecnicosFiestas` | SQL dinámico con fecha `#...#` | — | `TecnicoHoras.cls:432-435` |
| Informe | fichero `.xlsx` local | Automatización de Excel y `FileSystemObject` | — | `:208-220, 429-433` |

Ninguna escritura en base de datos. Hay lecturas por cada técnico y día (varias por fila, con caché solo dentro de cada `TecnicoHoras`).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: fechas como literales `#mm/dd/yyyy#` (dialecto Jet/ACE) en tres consultas (`Constructor.bas:4417`; `TecnicoHoras.cls:401, 434`); `FECHABAJA Is Null` y `Not FECHABAJA Is Null` (`Constructor.bas:4268, 4272`); alias concatenado entre comillas simples; `WeekdayName(Weekday(...))` dependiente de la configuración regional (`TecnicoHoras.cls:45`); N consultas por técnico y día que contra PostgreSQL vinculado multiplicarían los viajes de red.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El último día del rango nunca se procesa: el bucle va de `0` a `DateDiff - 1` (`:142-143`); con fechas iguales no se genera ninguna fila y, al no haber colección, no se crea el Excel (`:204-206`).
- `m_CadenaActividades` no se reinicia al final de cada fila (`:347-355` reinicia las demás variables): las filas siguientes podrían arrastrar los IDs de las anteriores y repetirlos (`:280-289, 315-317`).
- `intFila` es `Integer` (`:186`); más de 32 767 filas (técnicos × días) provocaría desbordamiento.
- La asignación de `ComboTecnico` en `:67-69` ocurre antes de rellenar la lista (`:71`): el técnico activo podría no quedar seleccionado.
- `ComboTecnico` lleva `Tag` «OBLIGATORIO» (`Form_FormTecnicoConsultaHoras.form.txt:478`) pero el código acepta vacío como «todos» (`:120-125`), y el combo solo lista activos (`:71`): no se puede consultar a un técnico de baja desde el combo.
- Los días de la semana dependen de `WeekdayName(Weekday(fecha - 1))` (`TecnicoHoras.cls:45`) y de comparar con «Sábado»/«Domingo» en castellano (`TecnicoHoras.cls:79`; `:266`): podría fallar con otra configuración regional.
- Si `PasarAExcel` falla, `lblEstado` queda visible (`:45` solo se alcanza en éxito) y el fichero parcial queda en disco (`:220, 440-447`).
- Cada ejecución crea un fichero nuevo con hora/minuto/segundo (`:208`) y no se eliminan los anteriores.
- La clase `TecnicoHoras` asigna `ObservacionesFestivo = Festivo` (`TecnicoHoras.cls:90`) usando el valor enumerado, no un texto: dato no usado por el informe.
- El festivo no es por técnico (`TecnicoHoras.cls:432-435`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: un único cálculo por rango en el servidor en vez de consultas por técnico y día; incluir el último día del rango; descarga del informe desde el navegador en lugar de abrir un fichero local; calendario laboral y jornada de 7,5 h parametrizables.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin `FechaInicial` o `FechaFinal`, `ComandoInforme` muestra el mensaje correspondiente y no genera fichero (`:108-115`).
2. Con `FechaFinal` anterior a `FechaInicial` muestra «La fecha final no puede ser anterior a la inicial» (`:116-119`).
3. Con `ComboTecnico` vacío el informe incluye a todos los técnicos activos; con un técnico, solo ese (`:120-138`).
4. El Excel tiene las 10 columnas de cabecera indicadas y no contiene filas de sábado ni domingo (`:225-234, 265-269`).
5. Un día laborable sin libranza con horas distintas de 7,5 aparece en rojo con la observación correspondiente (`:323-343`).
6. Un día de libranza o no laborable con horas aparece con su observación (`:329-336`).
7. Si el fichero del mismo nombre está abierto, se informa «Tiene una consulta abierta» (`:210-214`).
8. Abierto desde la gestión de técnicos o de actividades, `ComboTecnico` muestra el técnico seleccionado (`:67-69`; `Form_FormTecnicoGestion.cls:233-234`).

## Pendientes

- Contenido de `AjustarTamaño` más allá de la cabecera (`FG.bas:261-275`) y su efecto con el formulario no enlazado.
- Alcance de `Format(m_Fecha, "mm/dd/yyyy")` en el Excel (`:312`): si la celda se almacena como texto o como fecha.
- Esquema y claves de `TbTecnicosAusencias`, `TbTecnicosFiestas` y `TbActividades` (ERD de origen).
- Contenido de `Actividad.HorasLaborables` y `Actividad.HorasExtras` (origen y tipo), citados en `TecnicoHoras.cls:231, 273`.
