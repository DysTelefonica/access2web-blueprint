# Form_FormInicial03ConsultasEstado

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial03ConsultasEstado.cls` (56 líneas) y `.form.txt` (607 líneas) · **PRD:** [README](../README.md)

Formulario de consulta que filtra los usuarios por su **estado HPS** y nothing más. No escribe ningún dato: compone una clave en el diccionario global `m_ObjColParaFiltros` y delega el filtrado real en `EstablecerFiltroUsuarios`, que reescribe el `RecordSource` del subformulario de resultados del formulario principal. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`). Las citas son `Form_FormInicial03ConsultasEstado.cls:línea` salvo indicación; `Filtro.bas` es `modules/Filtro.bas` y `FG.bas`, `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial00Principal` | Se instala como subformulario del panel de consultas; el contenedor decide cuál se muestra | no hay `DoCmd.OpenForm` en `src` |
| `Form_FormInicial00Principal` → `SubFormCentral` → `SubFormResultados` | Destino del filtro: `EstablecerFiltroUsuarios` alcanza el subform de resultados por nombre y le asigna el `RecordSource` | `Filtro.bas:21-26` |

No abre ningún formulario.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Prepara estilo, vacía el diccionario de filtros y rellena el combo de estados | `:34-55` |
| `ComboEstadoHPS` | combo (`RowSourceType = Value List`) | `AfterUpdate` → `ComboEstadoHPS_AfterUpdate` | Publica el estado elegido como filtro y refresca los resultados | `:15-31`, `.form.txt:339-341` |
| `ComandoLimpiarEstadoHPS` | botón | `Click` → `ComandoLimpiarEstadoHPS_Click` | Pone el combo a `Null` y reaplica el filtro sin estado | `:10-12`, `.form.txt:370` |
| `lbl1` | etiqueta («Estado») | — | Rótulo del combo | `.form.txt:304` |
| `Detalle` | marco | — | Agrupa los controles; sin lógica | `.form.txt:281` |

## Flujos

### Apertura (`Form_Load`, `:34-55`)

1. Pone `m_ObjColParaFiltros` a `Nothing` (`:39`), lo que fuerza a que cada filtro se reconstruya desde cero.
2. `EstablecerControlesDatos Me, m_Error` aplica la paleta de colores a etiquetas, campos y listas (`:40`; `FG.bas:1005-1053`). Si devuelve error, salta al manejador (`:41-43`).
3. `RellenarComboEstado Me.ComboEstadoHPS, m_Error` vacía el `RowSource` del combo y añade un ítem por clave de `m_ObjEntorno.ColEstadosVisiblesHPS` (`:44`; `FG.bas:1233-1253`). Si devuelve error, salta al manejador (`:45-47`).
4. En error: compone el mensaje, envía correo al administrador si el error no es interno 1000, muestra un `MsgBox` crítico y **cierra el formulario** sin guardar (`:48-55`).

### Cambio de estado (`ComboEstadoHPS_AfterUpdate`, `:15-31`)

1. Si `m_ObjColParaFiltros` es `Nothing`, crea un `Scripting.Dictionary` con `CompareMode = TextCompare` (`:16-19`).
2. Escribe la clave `"Estado"` con el valor del combo, normalizado con `Nz` para que vacío y `Null` se comporten igual (`:22`).
3. Llama a `EstablecerFiltroUsuarios Me, m_Error` (`:23`).
4. **No comprueba `m_Error`**: si el filtrado falla, el error solo se ve si el propio helper lo propaga (`:23-31`).

### Limpiar (`ComandoLimpiarEstadoHPS_Click`, `:10-12`)

Pone `ComboEstadoHPS` a `Null` y llama a `ComboEstadoHPS_AfterUpdate` directamente (`:11-12`), lo que equivale a elegir el valor vacío en el combo.

### Errores

Los dos manejadores repiten el mismo patrón: si `Err.Number <> 1000`, anexan la descripción al error, envían `CorreoAlAdministrador` y muestran `MsgBox` crítico (`:26-31`, `:48-54`). El valor de retorno del `MsgBox` se guarda en el global `pregunta`, declarado en `VariablesEntorno.bas:220`.

## Reglas, validaciones y cálculos

- **Filtro exclusivo por estado.** `EstablecerFiltroUsuarios` ramifica por el nombre del formulario; esta rama es la de `FormInicial03ConsultasEstado` (`Filtro.bas:44`) y **descarta cualquier otro filtro**: solo añade `getWhereEstado` y, si está vacío, deja `SELECT * FROM TbDatosLocal` sin `WHERE` (`Filtro.bas:45-51, 28-29`).
- **Guard de contenedor abierto.** Si `FormInicial00Principal` no está abierto, la función sale sin tocar nada (`Filtro.bas:21-23`). El filtro se aplica, por tanto, por efecto secundario sobre otro formulario.
- **Exigencia de la clave `Estado`.** Si el diccionario no existe o no contiene `"Estado"`, `p_Error` se rellena con «No se ha rellenado la colección de parámetros» y se relanza el error interno 1000 (`Filtro.bas:31-39`).
- **Traducción de estado a predicado.** `getWhereEstado` traduce el estado visible elegido a cuatro alternativas HPS (NAC, OTAN, UE, ESA) unidas por `OR` (`Filtro.bas:163-206`). Cada estado visible es a su vez una composición de estados HPS elementales (`VariablesEntorno.bas:110-122`).
- **Estado desconocido.** Si el valor no está en `m_ObjEntorno.ColVisiblesHPSEstados`, `getWhereEstado` falla con «No existe el estado indicado» (`Filtro.bas:181-184`).
- **Sin concatenación de valor del usuario.** El texto del estado no se interpola en el SQL: solo se elige una de diez cadenas literales ya escritas (`Filtro.bas:185-204`). Es el patrón más seguro de los tres formularios de esta tanda.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Listado de resultados | `TbDatosLocal` | `RecordSource` asignado al subform de resultados | `Filtro.bas:28-29, 48-53` |
| Estados visibles disponibles | `m_ObjEntorno.ColEstadosVisiblesHPS` | colección en memoria, sin consulta | `FG.bas:1245-1246` |

Sin escrituras, sin transacciones.

## Coexistencia

Pendiente de `../coexistencia.md`. El `WHERE` se compone concatenando literales, así que la traducción es tolerable, pero el identificador `HPS_*` está replicado en diez ramas `If` (`Filtro.bas:185-204`) y un estado nuevo obliga a tocar esa función, no una tabla. Los errores llegan al usuario como `MsgBox` y el cierre es un `DoCmd.Close` (`:55`), sin equivalente web directo.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComboEstadoHPS_AfterUpdate` ignora el valor de `m_Error` que devuelve `EstablecerFiltroUsuarios` (`:23`), al contrario que `Form_Load`, que sí encadena `Err.Raise 1000` tras cada llamada (`:41-47`). Un fallo de filtrado queda en silencio.
- El `RowSource` del combo viene escrito en el binario como lista de valores (`form.txt:340-342`) pero `RellenarComboEstado` lo vacía y lo reconstruye en cada apertura (`FG.bas:1242, 1245-1246`). La lista del binario es código muerto y una fuente de divergencia.
- `EstablecerFiltroUsuarios` silenciosamente no hace nada si `FormInicial00Principal` no está abierto (`Filtro.bas:21-23`), sin avisar: el usuario ve un combo que responde y una lista que no cambia.
- El filtro se propaga por un global compartido (`VariablesEntorno.bas:256`) que todos los formularios de consulta sobrescriben; dos consultas abiertas a la vez interfieren.
- `getWhereEstado` incluye un `OR` mal formado en el estado de renovación: `... OR HPS_UE_Renovacion='Sí'OR HPS_ESA_Renovacion='Sí'` (`Filtro.bas:202`), sin espacio entre la cadena y el `OR`. Access lo tolera; otros motores no.
- `pregunta` es un `Long` global compartido (`VariablesEntorno.bas:220`) usado como retorno de `MsgBox` y nunca leído en este formulario.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtros como parámetros vinculados en lugar de `RecordSource` compuesto; tabla de equivalencia estado→columnas en lugar de la cadena de `If`; y un estado por defecto declarado en vez de «sin filtro» como valor vacío.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin estado elegido, la lista muestra todos los registros de `TbDatosLocal`.
2. Elegir cada uno de los diez estados visibles restringe la lista a los usuarios que lo tienen activo en al menos una de las cuatro líneas (NAC, OTAN, UE, ESA).
3. «Limpiar» devuelve la lista al estado sin filtro.
4. Si el formulario principal no está abierto, el filtro no altera ninguna otra lista y no se produce error visible.
5. Un estado que no exista en la configuración se comunica al usuario en vez de dejar la lista sin tocar.

## Pendientes

- Contenido de `m_ObjEntorno.ColEstadosVisiblesHPS` y `ColVisiblesHPSEstados`: se construyen en `Entorno.cls`, no en el formulario.
- Si algún menú o macro del binario abre este formulario directamente: no hay `DoCmd.OpenForm` en `src`.
- Relación exacta con `Form_FormInicial03ConsultasGrado`, que comparte `EstablecerFiltroUsuarios` por otra rama (`Filtro.bas:53-67`).