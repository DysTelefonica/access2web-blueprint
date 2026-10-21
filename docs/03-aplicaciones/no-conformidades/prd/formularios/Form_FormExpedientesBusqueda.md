# Form_FormExpedientesBusqueda

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormExpedientesBusqueda.cls` (379 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Selector de expedientes «EXPEDIENTES REGISTRADOS»: lista los expedientes de `TbExpedientes` filtrados por palabra clave y responsable de calidad y devuelve el elegido al formulario que lo abrió mediante el evento público `Seleccionado` (`:10, 88`). Es un formulario **no enlazado** (sin `RecordSource`, vista única `DefaultView =0`, `.form.txt:15`); la lista se rellena por código con `AddItem`. Las citas son `Form_FormExpedientesBusqueda.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_FormNCProyectoGestion | Cierra la instancia previa, `OpenForm`, se suscribe con `WithEvents`; si trae palabra clave la copia y llama a `Filtrar` | `Form_FormNCProyectoGestion.cls:10, 251-258, 1579-1587` |
| Form_FormNCProyectoGeneral | Igual | `Form_FormNCProyectoGeneral.cls:4, 225-232, 432-440` |
| Form_FormNCProyectoGeneralConVinculoNC | Igual | `Form_FormNCProyectoGeneralConVinculoNC.cls:5, 28-35, 82-89, 116-124` |
| Form_FormNCProyectoSeguimientoNC | Igual | `Form_FormNCProyectoSeguimientoNC.cls:8, 84-91, 117-125` |
| Form_FormNCProyectoSeguimientoTareas | Igual | `Form_FormNCProyectoSeguimientoTareas.cls:9, 85-92, 165-173` |

**Salida:** «Elegir» lanza `Seleccionado(m_Exp)` con el objeto `Expediente` y cierra el formulario (`:88-90`). Los receptores (`m_FormExpedientes_Seleccionado`) copian identificador, nemotécnico o código y responsable de calidad a su propio filtro (`Form_FormNCProyectoGestion.cls:1318-1335`, `Form_FormNCProyectoGeneral.cls:12-25`). «Salir» cierra sin devolver nada (`:126-129`).

## Controles y eventos

`.form.txt` declara 24 bloques `Begin`: el formulario (`.form.txt:5`), 8 de valores por defecto (`.form.txt:141-278`), 3 secciones (`EncabezadoDelFormulario` `.form.txt:299`, `Detalle` `.form.txt:397`, `PieDelFormulario` `.form.txt:1140`) y los 12 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `PalabraClave` | cuadro de texto | `OnExit` | Si el texto cambió respecto al último filtrado, llama a `Filtrar` y memoriza el valor | `:346-363`, `.form.txt:630-631` |
| `IDResponsableCalidad` | combo, lista de valores (id oculto; nombre), `LimitToList` | `BeforeUpdate` | Si el id cambió, llama a `Filtrar` y memoriza el valor; filas cargadas en `Form_Open` | `:134-151, 181, 202-228`, `.form.txt:595-600` |
| `ComandoLimpiarPalabraClave` | botón | `OnClick` | Vacía `PalabraClave`; refiltra solo si había un filtro previo | `:116-123`, `.form.txt:657-659` |
| `ComandoLimpiarIDResponsableCalidad` | botón | `OnClick` | Vacía el combo; refiltra solo si había un filtro previo | `:106-113`, `.form.txt:836-838` |
| `ComandoActualizar` | botón (texto «Comando47», tip «Actualizar») | `OnClick` | Olvida los últimos valores y vuelve a filtrar | `:14-27`, `.form.txt:337-340` |
| `ListaFiltrados` | lista de valores, 4 columnas con encabezado | `OnClick` | Carga el expediente seleccionado, muestra su título y habilita «Elegir» | `:284-326`, `.form.txt:423-424, 492` |
| `Titulo` | cuadro de texto deshabilitado | — | Título del expediente seleccionado | `:296, 325`, `.form.txt:1009, 1022` |
| `ComandoElegir` | botón | `OnClick` | Lanza `Seleccionado` y cierra | `:47-90`, `.form.txt:1111-1113` |
| `ComandoSalir` | botón | `OnClick` | `DoCmd.Close` sin guardar | `:126-129`, `.form.txt:1157-1159` |
| `lblTitulo1`, `Etiqueta29` («Palabra clave»), `Etiqueta30` («Resp. Calidad») | etiquetas | — | `lblTitulo1` da el título de la ventana | `:179`, `.form.txt:318, 1051, 1081` |

## Flujos

### Apertura (`Form_Open`, `:168-201`)

1. `Ajustar Me` (`:175`; `FG.bas:615`); la ventana toma el título de `lblTitulo1` (`:179`).
2. Deja en la lista solo la fila de encabezado `IDExp;Cod_Exp;Nemotécnico;Título` (`:180`).
3. Rellena el combo con los usuarios de calidad cacheados en `m_ObjEntorno.ColUsuariosCalidad` (`:181, 213-227`; `Entorno.cls:931-947`).
4. No filtra: la lista queda vacía hasta que el usuario o el formulario llamante invoca `Filtrar`.

### Filtrado (`Filtrar`, `:236-282`)

1. Vacía `Titulo` y reinicia la lista al encabezado (`:248-251`).
2. Obtiene la colección con `constructor.getExpedientesBusqueda(PalabraClave, id del combo)` (`:254-257`).
3. Añade una fila por expediente: id, código, nemotécnico y título con `;` sustituido por `:` (`:264-268`).
4. Selecciona la primera fila de datos y simula su clic (`:272-273`).

### Selección y devolución (`:47-90`)

Si no hay fila seleccionada, no hace nada (`:57-63`). Reutiliza el objeto cargado si coincide el id; si no, lo relee con `getExpediente` (`:65-80`). Lanza `Seleccionado` y cierra (`:88-90`).

### Errores

Error 1000 ⇒ `MsgBox` de advertencia con `m_Error`; otro ⇒ `CorreoAlAdministrador` y `MsgBox` crítico (patrón en `:35-43`). Si falla `Form_Open`, el formulario se cierra (`:200`). Los botones de limpiar usan `On Error Resume Next` y llaman a `Filtrar` sin recoger su error (`:107-110, 117-120`).

## Reglas, validaciones y cálculos

- **Filtro por responsable de calidad:** `WHERE IDResponsableCalidad = <id>`; sin id, toda la tabla (`constructor.bas:2410-2417`).
- **Palabra clave:** un expediente pasa si la palabra es **igual a su `IDExpediente`** o está contenida en `Nemotecnico` o en `Titulo` (`constructor.bas:2433-2438`); `InStr` sin argumento de comparación sigue `Option Compare Database` del módulo (`constructor.bas:2`). El filtro se aplica en memoria tras leer todas las filas.
- **Sin cambios no hay consulta:** ni la salida de `PalabraClave` ni el cambio de combo refiltran si el valor coincide con el último aplicado (`:141-146, 353-358`).
- **Usuarios de calidad:** los de `TbUsuariosAplicaciones` con `EsUsuarioCalidad = 'Sí'` en `TbUsuariosAplicacionesPermisos` para la aplicación, unidos por usuario (`constructor.bas:378-382`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Búsqueda de expedientes | `TbExpedientes` | SQL dinámico concatenado (`SELECT *`, id numérico del combo) + filtro en memoria | `constructor.bas:2410-2421, 2433-2438` |
| Expediente seleccionado | `TbExpedientes` | SQL dinámico concatenado (`WHERE IDExpediente=<id>`) | `constructor.bas:2259-2262` |
| Combo de responsables | `TbUsuariosAplicaciones` ⋈ `TbUsuariosAplicacionesPermisos` | SQL dinámico concatenado, cacheado en `Entorno` | `constructor.bas:378-383`; `Entorno.cls:939-947` |
| Registro de errores no controlados | `TbCorreosEnviados` | DAO `AddNew/Update` | `FG.bas:1291-1296` |

Solo lectura en el flujo funcional. Sin transacciones.

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `SELECT *` de toda `TbExpedientes` con filtro de texto en VBA (`constructor.bas:2415-2416, 2433-2436`), que en PostgreSQL vinculado conviene empujar al servidor; la semántica de mayúsculas de `InStr` bajo `Option Compare Database` no equivale a `LIKE` de PostgreSQL (sensible a mayúsculas) sin `ILIKE` o collation; identificadores concatenados en el SQL (`constructor.bas:2261, 2413`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El tip de `PalabraClave` promete buscar por «código expediente» (`.form.txt:633`), pero el filtro compara con `IDExpediente` por igualdad y nunca con `CodExp` (`constructor.bas:2434-2436`).
- Solo el título escapa `;` (`:267-268`); un código o nemotécnico con `;` desplazaría columnas en la lista de valores.
- `Filtrar` llamado desde otro formulario o desde «Actualizar» no actualiza `m_PalabraClaveUltima` (`:22-24, 236-282`): la siguiente salida de `PalabraClave` repite la consulta.
- Si `getExpediente` devuelve `Nothing` (expediente borrado entre filtrar y elegir), «Elegir» sale sin aviso (`:81-86`).
- `ColUsuariosCalidad` vacía deja el combo sin filas y sin aviso (`:219-221`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: selector modal con búsqueda en servidor (código, nemotécnico, título) y paginación; búsqueda por `CodExp` explícita; devolución del expediente por identificador en lugar de por objeto.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, la lista está vacía y el combo ofrece los usuarios de calidad de la aplicación.
2. Con palabra clave, la lista muestra solo los expedientes cuyo nemotécnico o título la contiene, o cuyo identificador es exactamente ese valor.
3. Con responsable de calidad, solo aparecen sus expedientes; ambos filtros se combinan.
4. Tras filtrar, la primera fila queda seleccionada, su título se muestra y «Elegir» se habilita.
5. «Elegir» devuelve el expediente al formulario llamante y cierra; «Salir» cierra sin devolver.

## Pendientes

- Efecto exacto de `Option Compare Database` sobre `InStr` con la configuración regional de la base (`constructor.bas:2`, `constructor.bas:2435`).
- `Ajustar` redimensiona por código (`FG.bas:615`); su efecto sobre este formulario no se ha trazado.
