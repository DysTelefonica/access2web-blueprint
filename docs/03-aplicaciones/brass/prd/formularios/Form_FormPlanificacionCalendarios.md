# Form_FormPlanificacionCalendarios

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormPlanificacionCalendarios.cls` (200 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Exportación a un libro Excel («Informe_Planificados_<año>.xls») de tres hojas —eventos, actividad realizada y material utilizado y reparado— a partir de tablas auxiliares ya rellenas. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) y no escribe en la base de datos. Abreviatura: `FG.bas` = `modules/Funciones Generales.bas`. Las citas son `Form_FormPlanificacionCalendarios.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| (ningún llamador en `src`) | No hay `OpenForm` hacia este formulario en `src/forms`, `modules` ni `classes`; la apertura (menú, macro o ventana de base de datos) queda pendiente | ver Pendientes |

Salida: un fichero `.xls` en una carpeta elegida por el usuario, que se abre al terminar con `ShellExecute` (`Ejecutar`, verbo `open`) (`:170`). Sale con `cmdSalir`, que cierra sin guardar (`:6-9`).

## Controles y eventos

Controles: 17 (17 `Begin` en el `.form.txt`): el propio `Form` (1), la sección `Detalle` (1), seis plantillas de estilo por defecto (`Label`, `Rectangle`, `Image`, `CommandButton`, `OptionGroup`, `TextBox`) (6) y nueve controles reales.

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `Año` | texto no enlazado | — | Año del informe; se lee como `Me.AÑO` y debe ser numérico | CAP-6 Planificación | `:128-132`; `.form.txt:393-420` |
| `ComandoGenerar` («Generar») | botón | `OnClick` → `ComandoGenerar_Click` | Pide carpeta, genera el libro y lo abre | CAP-6 Planificación | `:124-187`; `.form.txt:421-447` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra el formulario sin guardar | CAP-6 Planificación | `:6-9`; `.form.txt:301-326` |
| `ComandoAyudaFormulario` («Comando46») | botón | `OnClick` = `[Event Procedure]` en el `.form.txt`, **sin handler** en el `.cls` | Sin efecto observable | sin asignar | `.form.txt:327-338` |
| `lblTitulo1`, `lblTitulo2` | etiquetas «Calendario» | — | `lblTitulo1` se copia a `Caption` del formulario | CAP-6 Planificación | `:194`; `.form.txt:448-492` |
| `Etiqueta202` («AÑO») | etiqueta | — | Rótulo del campo año | CAP-6 Planificación | `.form.txt:366-392` |
| `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | sin asignar | `.form.txt:255-300` |

## Flujos

### Apertura (`Form_Load`, `:188-200`)

Con `On Error Resume Next` (`:189`): cursor de espera, `AjustarTamaño Me` (`FG.bas:261`), título de ventana desde `lblTitulo1` (`:193-194`) y repintado.

### Generar (`ComandoGenerar_Click`, `:124-187`)

1. Lee `Año` con `Nz`; si no es numérico, «Se ha de introducir primero el año» (`:128-132`).
2. `Seleccionar(False, ...)` abre el selector de carpeta (`:138`; `FG.bas:1117-1148`). Si devuelve «#ERR|…», lo propaga como error (`:139-143`). Si la carpeta devuelta no existe (p. ej. se cancela el selector), **sale sin mensaje** y sin quitar el cursor de espera (`:145-150`).
3. Cursor de espera y `Generar(carpeta, año)` (`:152-159`); un resultado con `|` es error (`:160-164`). `Generar` devuelve el literal «Edición correcta» en lugar de la ruta (`:101`), por lo que `FileExists("Edición correcta")` es falso y se lanza «El método Generar ha devuelto un resultado con formato desconocido» (`:165-169`): el flujo normal no llega a `Ejecutar` salvo que exista un fichero con ese nombre (a verificar).
4. Error: `MsgBox` crítico «Error» y se quita el cursor de espera (`:178-186`).

### `Generar` (`:10-123`)

1. Valida año numérico y carpeta existente (`:28-35`); añade `\` final (`:36-38`).
2. Nombre `Informe_Planificados_<año>.xls` (`:39-40`). Si existe: abierto ⇒ «El documento existe y está abierto» (`:47-50`); si no, lo borra (`:51`).
3. Crea la instancia global `appexcel` oculta y un libro (`:53-55`; `Variables Globales.bas:207`).
4. Hoja 1 «EVENTOS» con `ObtenerHojaEventosDeExportacion` (`:56-68`; `FG.bas:3654`); hoja 2 «ACTIVIDAD REALIZADA» (`:69-81`; `FG.bas:3376`); hoja 3 «MATERIAL UTILIZADO Y REPARADO» (`:82-94`; `FG.bas:3510`). Cada una devuelve «#ERR|texto» en fallo.
5. `SaveAs`, cierra Excel y devuelve «Edición correcta» (`:95-101`). En error cierra libro sin guardar y Excel (`:103-122`).

## Reglas, validaciones y cálculos

- **Año:** solo `IsNumeric` (`:129`, `:28`); no se acota ni se usa para filtrar datos, solo para el nombre del fichero (`:39`).
- **Fichero abierto:** `FicheroAbierto` intenta abrir el fichero en exclusiva; un error se interpreta como «abierto» (`FG.bas:195-217`).
- **Carpeta:** `FolderExists` (`:32`); la ruta de red configurada o local es indiferente para el código.
- **Contenido:** lo determinan las tablas auxiliares; el formulario no las rellena (ver Lecturas).
- **Cabecera de la hoja de actividad:** 13 columnas, fondo gris y negrita (`FG.bas:3402-3421`).

## Lecturas y escrituras

| Operación | Origen | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Actividad | `TbAuxActividad` con `TbCodActividad` | SQL estática en cadena, `OpenRecordset`, orden por `IDEVENTO` descendente | — | `FG.bas:3427-3430` |
| Materiales | `TbAuxMateriales` | SQL estática en cadena, orden por `IDEVENTO` | — | `FG.bas:3565-3568` |
| Eventos | `TbAuxEventos` | `OpenRecordset` por nombre de tabla | — | `FG.bas:3681, 3712-3713` |
| Fichero | carpeta elegida | `FileSystemObject` (borrado) y Excel `SaveAs` | — | `:51, 95` |

No hay escrituras en la base de datos. Quién rellena las tablas `TbAux*` no se localiza en `src` (ver Pendientes).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: dependencia de Excel por automatización COM y de `FileSystemObject`; las tablas auxiliares son locales de la sesión y deberían seguir existiendo en el entorno vinculado; la SQL no usa funciones Jet específicas.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Generar` devuelve «Edición correcta» (`:101`) pero el llamador lo trata como ruta del fichero (`:165-170`): el libro se guarda y luego se informa de «formato desconocido» y no se abre.
- Selección cancelada o carpeta inexistente: `Exit Sub` sin `Hourglass False` ni mensaje (`:145-150`); el cursor de espera aún no se había activado en ese punto (`:152`), por lo que el efecto real es solo la ausencia de mensaje.
- Los informes no filtran por año: el valor solo da nombre al fichero (`:39`; `FG.bas:3427-3429`).
- Si falla la creación de Excel, `appexcel` global puede quedar a medias; el mensaje de error nombra «El método Registrar» (`:105`).
- `ComandoAyudaFormulario` declara evento sin handler (`.form.txt:338`).
- `Form_Load` con `On Error Resume Next` oculta cualquier fallo de apertura (`:189`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: descarga del fichero en el navegador en lugar de carpeta local; filtrar por el año indicado.

## Criterios de aceptación

Propuestos, no aprobados:

1. Generar con `Año` vacío o no numérico muestra «Se ha de introducir primero el año» y no crea fichero (`:128-132`).
2. Con año válido y carpeta alcanzable se crea `Informe_Planificados_<año>.xls` con tres hojas «EVENTOS», «ACTIVIDAD REALIZADA» y «MATERIAL UTILIZADO Y REPARADO» (`:39, 57, 70, 83`).
3. Si el fichero destino existe y está abierto, se avisa y no se sobrescribe (`:47-50`).
4. Si el fichero existe y no está abierto, se sustituye (`:51`).
5. Un fallo en una hoja muestra el texto del error de esa hoja (`:64-68, 77-81, 90-94`).
6. `cmdSalir` cierra sin guardar (`:6-9`).

## Pendientes

- Desde dónde se abre el formulario (sin llamador en `src`; posible menú o macro del `.accdb`, no consultado).
- Quién rellena `TbAuxEventos`, `TbAuxActividad` y `TbAuxMateriales` (`FG.bas:3387, 3521, 3665`); solo se encuentra `TbAuxEventosParaInforme` en `Constructor.bas:1363`.
- Contenido de las hojas de eventos y materiales (columnas) y de `ObtenerHoja*` más allá de las líneas citadas.
