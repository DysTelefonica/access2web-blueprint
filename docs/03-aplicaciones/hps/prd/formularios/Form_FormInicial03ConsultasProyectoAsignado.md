# Form_FormInicial03ConsultasProyectoAsignado

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial03ConsultasProyectoAsignado.cls` (193 líneas) y `.form.txt` (976 líneas) · **PRD:** [README](../README.md)

Formulario de consulta que filtra los usuarios por **proyecto (expediente) asignado** y estado HPS. No escribe ningún dato: escribe dos claves en el diccionario global `m_ObjColParaFiltros` y delega el filtrado en `EstablecerFiltroUsuarios`. Es **no enlazado** (el `.form.txt` no declara `RecordSource`). Es el único de la tanda que **no delega el relleno del combo**: trae su propia función `RellenarCombo` declarada en el módulo del propio formulario (`:103-193`), que recorre dos `SELECT` de expedientes y deduplica por `IDExpediente` en un `Scripting.Dictionary`. Las citas son `Form_FormInicial03ConsultasProyectoAsignado.cls:línea` salvo indicación; `Filtro.bas` es `modules/Filtro.bas` y `FG.bas`, `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial02ConsultasPrincipal` | Se monta como `SubFormParametrosConsulta`; el padre resuelve el nombre desde el ítem de `ListaConsultas` («03;Por Proyecto Asignado») | `Form_FormInicial02ConsultasPrincipal.cls:263, 272`, `.form.txt:589-591` |
| `Form_FormInicial00Principal` → `SubFormCentral` → `SubFormResultados` | Destino del filtro: `EstablecerFiltroUsuarios` le asigna el `RecordSource` | `Filtro.bas:21-26, 107` |

No abre ningún formulario.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ComboProyectoAsignado` | combo (`Value List`, dos columnas) | `AfterUpdate` → `ComboProyectoAsignado_AfterUpdate` | Publica el expediente como filtro `"IDExpediente"` y refiltra | `:39-54`, `.form.txt:317-320` |
| `ComboEstadoHPS` | combo (`Value List`) | `AfterUpdate` → `ComboEstadoHPS_AfterUpdate` | Publica el estado como filtro `"Estado"` y refiltra | `:55-70`, `.form.txt:388-393` |
| `ComandoLimpiarProyectoAsignado` | botón | `Click` → `ComandoLimpiarProyectoAsignado_Click` | Pone el combo a `Null` y reejecuta su `AfterUpdate` | `:31-34`, `.form.txt:601, 603` |
| `ComandoLimpiarEstadoHPS` | botón | `Click` → `ComandoLimpiarEstadoHPS_Click` | Pone el combo a `Null` y reejecuta su `AfterUpdate` | `:35-38`, `.form.txt:419, 421` |
| `lblPAsignado`, `Etiqueta0` | etiquetas | — | Rótulos del combo de expediente y de «Estado»; sin lógica | `.form.txt:281, 352` |
| `Detalle` | sección | — | Agrupa los seis controles; sin lógica | `.form.txt:257-258` |

El `RowSource` del combo de estado viene escrito en el binario (`.form.txt:390`) y lo sobrescribe `RellenarComboEstado` en cada apertura (`FG.bas:1242`); el combo de expediente no trae `RowSource` en diseño (`.form.txt:317-320`).

## Flujos

### Filtrar (`Filtrar`, `:8-29`)

1. Si `m_ObjColParaFiltros` es `Nothing`, crea un `Scripting.Dictionary` con `CompareMode = TextCompare` (`:11-14`).
2. Escribe `"Estado"` con `Nz(Me.ComboEstadoHPS, "")` (`:15`) y `"IDExpediente"` con `Nz(Me.ComboProyectoAsignado, "")` (`:16`).
3. Llama a `EstablecerFiltroUsuarios Me, m_Error` (`:18`) y, si `m_Error` no quedó vacío, relanza el error interno 1000 (`:19-21`).
4. `errores:` anexa la descripción y avisa al administrador solo si `Err.Number <> 1000` (`:24-27`); el `MsgBox` final muestra el `m_Error` pendiente, guardado en el global `pregunta` (`VariablesEntorno.bas:220`).

### Cambio de combo y botones «limpiar» (`:31-70`)

Los dos `AfterUpdate` (`:39-54`, `:55-70`) comparten el esqueleto `On Error GoTo errores`, `m_Error = ""`, `Filtrar`, `If m_Error <> "" Then Err.Raise 1000`, correo al administrador si `Err.Number <> 1000` y `MsgBox` crítico. Los dos botones «limpiar» (`:31-38`) ponen su combo a `Null` y llaman directamente al `AfterUpdate`.

### Relleno del combo de expedientes (`RellenarCombo`, `:103-193`)

Función **pública declarada en el propio módulo del formulario**, con firma `(p_Cbo As ComboBox, Optional ByRef p_Error)`:

1. `SELECT distinct TbExpedientes.* FROM TbExpedientes INNER JOIN TbUsuarios ON TbExpedientes.IDExpediente = TbUsuarios.IDExpediente ORDER BY TbExpedientes.CodExp` sobre `getdb()` (`:117-121`), leyendo `IDExpediente`, `Nemotecnico` y `CodExp` (`:126-128`).
2. Para cada fila elige el nemotécnico si no está vacío y si no el `CodExp` (`:129-133`), y añade al diccionario `CStr(IDExpediente) → <ID>;<texto>` solo si la clave no existe (`:134-140`).
3. Repite el recorrido con `INNER JOIN TbUsuariosHistoricos` (`:149-152`, `:157-171`), de modo que un mismo expediente no se duplica aunque esté en las dos tablas.
4. Si el diccionario quedó vacío, sale sin tocar el combo (`:179-181`); si no, recorre sus **claves** y hace `AddItem` del valor (`:182-186`).
5. `errores:` rellena `p_Error` con «El método RellenarCombo ha devuelto el error: …» solo si `Err.Number <> 1000` (`:190-192`).

### Apertura (`Form_Load`, `:72-102`)

1. `Set m_ObjColParaFiltros = Nothing` (`:77`), descartando el filtro heredado.
2. `EstablecerControlesDatos Me, m_Error` (`:78`; `FG.bas:1005-1055`) con `Err.Raise 1000` encadenado (`:79-81`).
3. `RellenarCombo Me.ComboProyectoAsignado, m_Error` (`:82`), la versión local de `:103`, con el raise encadenado (`:83-85`); la llamada a `RellenarComboExpediente` queda comentada y no se ejecuta (`:86-89`).
4. `RellenarComboEstado Me.ComboEstadoHPS, m_Error` (`:90`) con el raise encadenado (`:91-93`).
5. En error: mensaje, `CorreoAlAdministrador` si `Err.Number <> 1000`, `MsgBox` crítico y `DoCmd.Close acForm, Me.Name, acSaveNo` (`:96-101`).

## Reglas, validaciones y cálculos

- **Filtro combinado de dos claves, todo con `AND`.** El nombre del formulario no es `FormInicial03ConsultasEstado` ni `FormInicial03ConsultasGrado`, así que `EstablecerFiltroUsuarios` entra por la rama genérica (`Filtro.bas:69-107`): compone `getWhereParametro` (`:70`) y lo une al `WHERE` de estado en las cuatro variantes de `:98-106`.
- **Guard de contenedor abierto.** Sin `FormInicial00Principal` abierto, la función sale sin tocar nada (`Filtro.bas:21-23`).
- **Exigencia de la clave `Estado`.** Sin diccionario o sin `"Estado"`, `p_Error` se rellena con «No se ha rellenado la colección de parámetros» y se lanza el error interno 1000 (`Filtro.bas:31-38`). Aquí `Estado` se escribe siempre (`:15`), también vacío.
- **`IDExpediente` es la única clave sin comillas.** `getWhereParametro` le asigna el comparador `=` y sin comilla final (`Filtro.bas:138-141`), mientras el resto de claves usan `='<valor>'` (`Filtro.bas:142-146`); `"Estado"` se salta por comparación explícita (`:130-132`) y las restantes se encadenan con `AND` (`:151`).
- **La clave es la primera columna del combo.** El combo es de dos columnas (`ColumnWidths ="0;4847"`, `.form.txt:319`) y cada ítem es `<ID>;<Nemotecnico|CodExp>` (`:139, :170`); lo que va al diccionario es `Nz(Me.ComboProyectoAsignado, "")` (`:16`), la columna de valor, no el texto visible.
- **Precedencia del nemotécnico.** Un expediente con `Nemotecnico` vacío se muestra por su `CodExp` (`:129-133`), y la lista es la unión de los expedientes referenciados por `TbUsuarios` y por `TbUsuariosHistoricos`, ambas con `DISTINCT` y `ORDER BY CodExp` (`:117-119, 149-151`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Expedientes de usuarios activos | `TbExpedientes` `INNER JOIN` `TbUsuarios` (backend, vía `getdb()`) | `SELECT DISTINCT … ORDER BY CodExp` | `:117-121` |
| Expedientes de usuarios históricos | `TbExpedientes` `INNER JOIN` `TbUsuariosHistoricos` (backend, vía `getdb()`) | `SELECT DISTINCT … ORDER BY CodExp` | `:149-152` |
| Listado de resultados | `TbDatosLocal` | `RecordSource` reasignado al subform de resultados | `Filtro.bas:28-29, 107` |
| Estados visibles disponibles | `m_ObjEntorno.ColEstadosVisiblesHPS` | colección en memoria, sin consulta | `FG.bas:1245-1246` |

Sin escrituras, sin transacciones.

## Coexistencia

Pendiente de `../coexistencia.md`. El valor del expediente se concatena en el SQL **sin comillas y sin comillas dobles** (`Filtro.bas:140-141`), lo que solo es válido mientras la columna sea numérica y el combo aporte el identificador y no el texto visible. La lista se arma con dos consultas casi idénticas duplicadas en el propio `.cls` (`:122-146` y `:153-176`), y `RellenarCombo` colisiona en nombre con la función de propósito general de `FG.bas:1285`. Los errores llegan al usuario como `MsgBox` y el cierre es un `DoCmd.Close` (`:101`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Colisión de nombres: el formulario declara su propio `Public Function RellenarCombo(p_Cbo, p_Error)` (`:103-106`) mientras existe una global `RellenarCombo(p_NombreCampo, p_Cbo, p_NombreTabla, p_Error)` (`FG.bas:1285-1290`). La llamada sin cualificar de `:82` resuelve a la local; conviene verificar en el binario que la resolución es la esperada.
- El `ORDER BY TbExpedientes.CodExp` de ambas consultas (`:119, :151`) no se traslada al combo: los ítems se emiten recorriendo las **claves** de un `Scripting.Dictionary` (`:182-186`), cuyo orden de recorrido en Access no está garantizado por la API; el orden real del desplegable queda a verificar.
- El valor concatenado en `IDExpediente=<valor>` (`Filtro.bas:140-141`) no lleva comillas ni conversión a número. Si el combo devolviera el ítem completo `<ID>;<texto>` en lugar de su columna de valor, el `RecordSource` sería inválido; que devuelva la columna correcta depende de la resolución de `ColumnCount` a partir de `ColumnWidths` (.form.txt:319), a verificar.
- `RellenarCombo` no vacía `p_Cbo.RowSource` antes de los `AddItem` (`:184`), a diferencia de la versión de `FG.bas:1303`; y si ambos `Recordset` devuelven cero filas, el combo conserva su contenido de diseño sin avisar (`:179-181`).
- El bloque comentado de `RellenarComboExpediente` (`:86-89`) es código muerto; la función existe en `FG.bas:1254`.
- `m_Error` no se vacía al entrar en `Filtrar` (`:10-14`); la relyga la vacían los dos llamantes (`:41, :57`) pero `Form_Load` llama a `EstablecerControlesDatos` sin vaciarla (`:78`). Ni `EstablecerControlesDatos` (`FG.bas:1005-1055`) ni `RellenarComboEstado` (`FG.bas:1233-1253`) vacían su `p_Error` al entrar, al contrario que `EstablecerFiltroUsuarios` (`Filtro.bas:20`); este último traga el error sin relanzar (`Filtro.bas:110-114`), de modo que la propagación depende solo de la cadena `If m_Error <> "" Then Err.Raise 1000`.
- El filtro se propaga por un global compartido (`VariablesEntorno.bas:256`) que todos los formularios de consulta sobrescriben. `m_Parametro` se declara y no se usa (`:109`), igual que en `Filtro.bas:17`.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: catálogo de expedientes servido por el backend en vez de dos consultas duplicadas; y filtro por parámetro vinculado en lugar de `RecordSource` compuesto.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con ambos controles vacíos, la lista muestra todos los registros de `TbDatosLocal`.
2. Elegir un expediente restringe la lista a los usuarios de ese expediente; estado vacío.
3. Elegir un estado restringe la lista por las cuatro líneas HPS; expediente vacío.
4. Con ambos rellenos, las restricciones se combinan con `AND`.
5. La lista de expedientes no repite un `IDExpediente` aunque esté presente en `TbUsuarios` y en `TbUsuariosHistoricos`.
6. Un expediente sin nemotécnico se muestra por su `CodExp`.
7. Si no hay ningún expediente, el combo queda vacío y no se produce error visible.

## Pendientes

- Contenido de `m_ObjEntorno.ColTiposConsultasFormulario`: decide qué ítem de `ListaConsultas` monta este formulario y se construye en `Entorno.cls`, no en `src/`.
- Orden real de los ítems del desplegable: depende del recorrido del diccionario, a verificar contra el binario.
- Dónde quedó la versión previa que usaba `RellenarComboExpediente` (`FG.bas:1254`), ya que aquí la llamada está comentada (`:86-89`); y la relación exacta con `FormInicial03ConsultasGrado`, que comparte `EstablecerFiltroUsuarios` por su rama propia (`Filtro.bas:53-67`) y exige seis claves (`Filtro.bas:245-253`). Esta ficha **no** entra por esa rama.