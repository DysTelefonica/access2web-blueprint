# Form_FormInicial03ConsultasEmpresaUsuario

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial03ConsultasEmpresaUsuario.cls` (189 líneas) y `.form.txt` (1574 líneas) · **PRD:** [README](../README.md)

Formulario de consulta que filtra los usuarios por **empresa a la que pertenecen** y estado HPS. No escribe ningún dato: escribe dos claves en el diccionario global `m_ObjColParaFiltros` y delega el filtrado en `EstablecerFiltroUsuarios`. Es **no enlazado** (el `.form.txt` no declara `RecordSource`). Es el único de la tanda que **no guarda el identificador en el filtro sino el nombre real resuelto**: elige un `IDEmpresaUsuario`, consulta `Constructor.getSuministrador` y escribe el `Nombre` del suministrador (`:68-77`); el comentario del propio código lo declara como corrección del fallo de búsqueda (`:65-67`). Trae además su propia rutina de carga del combo, `CargarEmpresasUsadas` (`:120-189`), que cruza **dos bases de datos distintas**. Las citas son `Form_FormInicial03ConsultasEmpresaUsuario.cls:línea` salvo indicación; `Filtro.bas` es `modules/Filtro.bas` y `FG.bas`, `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial02ConsultasPrincipal` | Se monta como `SubFormParametrosConsulta`; el padre resuelve el nombre desde el ítem de `ListaConsultas` («06;Por Empresa que Pertenece el Usuario») | `Form_FormInicial02ConsultasPrincipal.cls:263, 272`, `.form.txt:589-591` |
| `Form_FormInicial00Principal` → `SubFormCentral` → `SubFormResultados` | Destino del filtro: `EstablecerFiltroUsuarios` le asigna el `RecordSource` | `Filtro.bas:21-26, 107` |

No abre ningún formulario.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `IDEmpresaUsuario` | combo (`Value List`, dos columnas) | `AfterUpdate` → `IDEmpresaUsuario_AfterUpdate` | Publica la empresa como filtro `"EmpresaUsuario"` y refiltra | `:22-34`, `.form.txt:921-924` |
| `ComboEstadoHPS` | combo (`Value List`) | `AfterUpdate` → `ComboEstadoHPS_AfterUpdate` | Publica el estado como filtro `"Estado"` y refiltra | `:36-48`, `.form.txt:991-995` |
| `ComandoLimpiarEmpresaUsuario` | botón | `Click` → `ComandoLimpiarEmpresaUsuario_Click` | Pone el combo a `Null` y reejecuta su `AfterUpdate` | `:10-13`, `.form.txt:1021, 1023` |
| `ComandoLimpiarEstadoHPS` | botón | `Click` → `ComandoLimpiarEstadoHPS_Click` | Pone el combo a `Null` y reejecuta su `AfterUpdate` | `:15-18`, `.form.txt:1203, 1205` |
| `lbl1`, `Etiqueta0` | etiquetas | — | Rótulos del combo de empresa y de «Estado»; sin lógica | `.form.txt:886, 956` |
| `Detalle` | sección | — | Agrupa los seis controles; sin lógica | `.form.txt:860-863` |

El `RowSource` del combo de estado viene escrito en el binario (`.form.txt:993`) y lo sobrescribe `RellenarComboEstado` en cada apertura (`FG.bas:1242`); el combo de empresa se puebla íntegramente en tiempo de ejecución y no trae `RowSource` en diseño (`.form.txt:921-923`).

## Flujos

### Filtrar (`Filtrar`, `:52-90`)

1. Si `m_ObjColParaFiltros` es `Nothing`, crea un `Scripting.Dictionary` con `CompareMode = TextCompare` (`:57-60`).
2. Escribe `"Estado"` con `Nz(Me.ComboEstadoHPS, "")` (`:63`).
3. Si `Nz(Me.IDEmpresaUsuario, "")` no está vacío, resuelve el suministro con `Constructor.getSuministrador(p_ID:=Me.IDEmpresaUsuario)` (`:69`) y escribe `"EmpresaUsuario"` con `m_Suministrador.Nombre` si el objeto existe (`:70-71`) o con `""` si no (`:72-74`); con el combo vacío escribe `""` directamente (`:75-77`).
4. Llama a `EstablecerFiltroUsuarios Me, m_Error` (`:80`) y, si `m_Error` no quedó vacío, relanza el error interno 1000 (`:81`).
5. `errores:` anexa la descripción y avisa al administrador solo si `Err.Number <> 1000` (`:85-88`); el `MsgBox` final muestra el `m_Error` pendiente, guardado en el global `pregunta` (`VariablesEntorno.bas:220`).

### Cambio de combo y botones «limpiar» (`:10-48`)

Los dos `AfterUpdate` (`:22-34`, `:36-48`) comparten el esqueleto `On Error GoTo errores`, `m_Error = ""`, `Filtrar`, `If m_Error <> "" Then Err.Raise 1000`, correo al administrador si `Err.Number <> 1000` y `MsgBox` crítico. Los dos botones «limpiar» (`:10-18`) ponen su combo a `Null` y llaman directamente al `AfterUpdate`.

### Carga del combo de empresas (`CargarEmpresasUsadas`, `:120-189`)

Rutina **privada del propio formulario**, con tres pasos y dos bases de datos:

1. `Set db = getdb()` (`:131`, backend de usuarios, `VariablesEntorno.bas:587-608`) y `SELECT DISTINCT IDEmpresaUsuario FROM TbUsuarios WHERE IDEmpresaUsuario IS NOT NULL` (`:135`), acumulando cada identificador en un `Scripting.Dictionary` local (`:137-141`).
2. Repite sobre `SELECT DISTINCT IDEmpresaUsuario FROM TbUsuariosHistoricos WHERE IDEmpresaUsuario IS NOT NULL` (`:144-150`) contra la misma base, de modo que una empresa usada solo por usuarios históricos también entra en el desplegable.
3. `Set db = getdbExpedientes()` (`:155`, `VariablesEntorno.bas:662-682`, otra clave de configuración) y `Me.IDEmpresaUsuario.RowSource = ""` (`:156`); después recorre `SELECT IDSuministrador, Nombre, CIF FROM TbSuministradores ORDER BY Nombre` (`:159-162`) y hace `AddItem` **solo** de las empresas presentes en el diccionario y con nombre no vacío (`:166-174`), con el formato `<ID>;<Nombre> (<CIF>)` (`:173`).
4. `errores:` fija `m_Error = "Error al CargarEmpresasUsadas: …"` y relanza 1000 (`:186-188`).

### Apertura (`Form_Load`, `:92-117`)

1. `m_Error = ""` y `Set m_ObjColParaFiltros = Nothing` (`:95-96`), descartando el filtro heredado.
2. `EstablecerControlesDatos Me, m_Error` (`:99`; `FG.bas:1005-1055`) con `Err.Raise 1000` encadenado (`:100`).
3. `CargarEmpresasUsadas` (`:103`), **sin** comprobar `m_Error`: la única protección es el `Err.Raise 1000` de `:188` y el salto al manejador de `Form_Load`.
4. `RellenarComboEstado Me.ComboEstadoHPS, m_Error` (`:106`) con el raise encadenado (`:107`).
5. En error: mensaje, `CorreoAlAdministrador` si `Err.Number <> 1000`, `MsgBox` crítico y `DoCmd.Close acForm, Me.Name, acSaveNo` (`:111-116`).

## Reglas, validaciones y cálculos

- **Filtro combinado de dos claves, todo con `AND`.** El nombre del formulario no es `FormInicial03ConsultasEstado` ni `FormInicial03ConsultasGrado`, así que `EstablecerFiltroUsuarios` entra por la rama genérica (`Filtro.bas:69-107`): compone `getWhereParametro` (`:70`) y lo une al `WHERE` de estado en las cuatro variantes de `:98-106`.
- **Guard de contenedor abierto.** Sin `FormInicial00Principal` abierto, la función sale sin tocar nada (`Filtro.bas:21-23`).
- **Exigencia de la clave `Estado`.** Sin diccionario o sin `"Estado"`, `p_Error` se rellena con «No se ha rellenado la colección de parámetros» y se lanza el error interno 1000 (`Filtro.bas:31-38`). Aquí `Estado` se escribe siempre (`:63`), también vacío.
- **El filtro viaja por nombre, no por identificador.** `getWhereParametro` no tiene rama propia para `"EmpresaUsuario"`, así que cae en la genérica y produce `EmpresaUsuario='<nombre>'` con comillas (`Filtro.bas:142-146`); `"Estado"` se salta por comparación explícita (`:130-132`) y el resto se encadena con `AND` (`:151`).
- **Resolución del nombre antes de filtrar.** `Constructor.getSuministrador` busca por `IDSuministrador=<valor>` en la base de expedientes y, si encuentra fila, construye un `Suministrador` rellenando todas sus propiedades desde el recordset (`Constructor.bas:39-61`); si el recordset está vacío devuelve `Nothing` (`Constructor.bas:49-53`), lo que aquí se traduce en un filtro por empresa vacío (`Filtro.bas:98-99`) en lugar de un aviso.
- **Desaparición de empresas sin usuarios.** El combo solo ofrece proveedores que aparecen en `TbUsuarios` o `TbUsuariosHistoricos` (`:135, 144, 166`) y descarta los de nombre vacío (`:170`).
- **Carga transversal de dos bases.** El formulario lee del backend de usuarios (`getdb()`, `VariablesEntorno.bas:596`) y del backend de expedientes (`getdbExpedientes()`, `VariablesEntorno.bas:671`); ambas rutas salen de `TbConfiguracionHPS` vía `m_ObjEntorno.GetConfigOrFail`, y `getSuministrador` concatena el `p_ID` sin comillas ni validación numérica (`Constructor.bas:40-41`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Empresas usadas por usuarios activos | `TbUsuarios` (backend de usuarios, vía `getdb()`) | `SELECT DISTINCT IDEmpresaUsuario … IS NOT NULL` | `:131, 135-136` |
| Empresas usadas por usuarios históricos | `TbUsuariosHistoricos` (backend de usuarios, vía `getdb()`) | `SELECT DISTINCT IDEmpresaUsuario … IS NOT NULL` | `:144-145` |
| Catálogo de proveedores | `TbSuministradores` (backend de expedientes, vía `getdbExpedientes()`) | `SELECT IDSuministrador, Nombre, CIF … ORDER BY Nombre` | `:155, 159-160` |
| Nombre del proveedor seleccionado | `TbSuministradores` (vía `Constructor.getSuministrador`) | `SELECT * … WHERE IDSuministrador=<valor>` | `:69`, `Constructor.bas:40-41` |
| Listado de resultados | `TbDatosLocal` | `RecordSource` reasignado al subform de resultados | `Filtro.bas:28-29, 107` |
| Estados visibles disponibles | `m_ObjEntorno.ColEstadosVisiblesHPS` | colección en memoria, sin consulta | `FG.bas:1245-1246` |

Sin escrituras, sin transacciones: el formulario solo lee y compone un `RecordSource`.

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Es el único de la tanda que **cruza dos bases de datos** en el mismo `Form_Load` (`:131` y `:155`), cada una abierta por su propia clave de configuración, y el que concatena el `p_ID` del proveedor en el SQL sin comillas (`Constructor.bas:40-41`). Convive con las otras formas que escriben el mismo global: `FormInicial03ConsultasContratistas` escribe `"Estado"` y `"CadenaContratistas"` (`Form_FormInicial03ConsultasContratistas.cls:56-57`). Los errores llegan al usuario como `MsgBox` y el cierre es un `DoCmd.Close` (`:116`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `CargarEmpresasUsadas` relanza 1000 (`:188`) y el manejador de `Form_Load` solo envía correo si `Err.Number <> 1000` (`:111`): un fallo de carga del combo muestra el `MsgBox` con el texto de `:187` pero **no avisa al administrador**, a diferencia del resto de handlers del formulario. Además `Form_Load` llama a la rutina (`:103`) sin comprobar `m_Error` ni encadenar un raise, al contrario que las dos llamadas vecinas (`:100, :107`).
- `getSuministrador` concatena el identificador sin comillas ni conversión (`Constructor.bas:40-41`); que el combo aporte el identificador y no el ítem completo `<ID>;<Nombre> (<CIF>)` (`:173`) depende de la resolución de `ColumnCount` a partir de `ColumnWidths` (.form.txt:923), a verificar.
- El nombre del proveedor se compara por igualdad exacta en el `WHERE` (`Filtro.bas:144-145`): un cambio de razón social en `TbSuministradores` deja de encontrar los usuarios ya filtrados, porque `TbDatosLocal.EmpresaUsuario` guarda una copia del nombre (`Mod_Sincronizacion_Historico.bas:52`).
- El `RowSource` del combo de estado escrito en el binario (`.form.txt:993`) es código muerto: `RellenarComboEstado` lo vacía en cada apertura (`FG.bas:1242`).
- `m_Error` no se vacía al entrar en `Filtrar` (`:55`); la relyga la vacían los dos llamantes (`:24, :38`). Ni `EstablecerControlesDatos` (`FG.bas:1005-1055`) ni `RellenarComboEstado` (`FG.bas:1233-1253`) vacían su `p_Error` al entrar, al contrario que `EstablecerFiltroUsuarios` (`Filtro.bas:20`); este último traga el error sin relanzar (`Filtro.bas:110-114`), de modo que la propagación depende solo de la cadena `If m_Error <> "" Then Err.Raise 1000`.
- El filtro se propaga por un global compartido (`VariablesEntorno.bas:256`) que todos los formularios de consulta sobrescriben; y `getdb()` y `getdbExpedientes()` se abren sin cerrar explícitamente (`:151, 181` solo hacen `Set db = Nothing`), dejando referencias vivas en la sesión de Access.
- El comentario «CORREGIDO» de `:65-67` documenta un arreglo ya aplicado en el `.cls`; el histórico del fallo no está en `src/`.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtro por identificador de empresa en lugar de por razón social, para que un cambio de nombre no rompa la búsqueda; y catálogo de empresas servido por el backend en vez del cruce de dos consultas.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con ambos controles vacíos, la lista muestra todos los registros de `TbDatosLocal`.
2. Elegir una empresa restringe la lista a los usuarios cuyo `EmpresaUsuario` coincide con su razón social exacta; estado vacío.
3. Elegir un estado restringe la lista por las cuatro líneas HPS; empresa vacía.
4. Con ambos rellenos, las restricciones se combinan con `AND`.
5. El desplegable solo ofrece proveedores con al menos un usuario activo o histórico y con nombre no vacío.
6. Cada ítem del desplegable muestra el identificador, el nombre y el CIF.
7. Si el identificador seleccionado no resuelve a ningún proveedor, la lista queda sin filtro de empresa y el error se comunica al usuario.
8. Si el formulario principal no está abierto, el filtro no altera ninguna otra lista y no se produce error visible.

## Pendientes

- Contenido de `m_ObjEntorno.ColTiposConsultasFormulario`: decide qué ítem de `ListaConsultas` monta este formulario y se construye en `Entorno.cls`, no en `src/`.
- Confirmar si el `MsgBox` de un fallo de `CargarEmpresasUsadas` debe llegar también al administrador, como ocurre en los demás handlers.
- Confirmar en el binario el valor real de `.Value` del combo (identificador frente a ítem completo), del que depende la validez del SQL de `Constructor.bas:40-41`; y la relación exacta con `FormInicial03ConsultasGrado`, que comparte `EstablecerFiltroUsuarios` por su rama propia (`Filtro.bas:53-67`) y exige seis claves (`Filtro.bas:245-253`). Esta ficha **no** entra por esa rama.