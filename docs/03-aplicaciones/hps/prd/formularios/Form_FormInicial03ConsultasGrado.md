# Form_FormInicial03ConsultasGrado

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial03ConsultasGrado.cls` (256 líneas) y `.form.txt` (2132 líneas) · **PRD:** [README](../README.md)

Formulario de consulta que filtra los usuarios de `TbDatosLocal` por **grado HPS** (NAC, OTAN, UE, ESA), por **especialidad** y por **estado HPS**, y combina además esos dos criterios según un conmutador Y/O. No escribe ningún dato: rellena siete claves del diccionario global `m_ObjColParaFiltros` y delega el SQL en `Filtro.EstablecerFiltroUsuarios`, que alcanza el subformulario de resultados del formulario principal por nombre. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`). Las citas son `Form_FormInicial03ConsultasGrado.cls:línea` salvo indicación; `Filtro.bas` es `modules/Filtro.bas` y `FG.bas`, `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial02ConsultasPrincipal` | Se instala como `SourceObject` de `SubFormParametrosConsulta`; el contenedor elige el formulario de consulta desde `m_ObjEntorno.ColTiposConsultasFormulario` | `Form_FormInicial02ConsultasPrincipal.cls:263, 272` |
| `Form_FormInicial00Principal` → `SubFormCentral` → `SubFormResultados` | Destino del filtro: `EstablecerFiltroUsuarios` alcanza el subform de resultados por nombre y le asigna el `RecordSource` | `Filtro.bas:21-26, 67` |

No abre ningún formulario.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Prepara estilo, vacía el diccionario, carga combos y fija `ComboYO = "O"` | `:178-252`, `.form.txt:81` |
| `ComboGradoNacional`, `ComboGradoOtan`, `ComboGradoUE`, `ComboGradoESA` | 4 combos (`RowSourceType = Value List`) | `AfterUpdate` → `ComboGrado*_AfterUpdate` | Publican el grado elegido y refiltran | `:36-100`, `.form.txt:320, 389, 459, 529` |
| `ComboEstadoHPS` | combo (`Value List`) | `AfterUpdate` → `ComboEstadoHPS_AfterUpdate` | Publica el estado y refiltra | `:144-160`, `.form.txt:671, 693` |
| `ComboYO` | combo (`Value List`: `Y;O`) | `AfterUpdate` → `ComboYO_AfterUpdate` | Elige si los grados se unen con `AND` u `OR` | `:128-143`, `.form.txt:599, 621` |
| `Especialidad` | combo (`Value List`) | `AfterUpdate` → `Especialidad_AfterUpdate` | Publica la especialidad y refiltra | `:161-177`, `.form.txt:1653, 1673` |
| `ComandoLimpiarGradoNacional`, `ComandoLimpiarGradoOTAN`, `ComandoLimpiarGradoUE`, `ComandoLimpiarGradoESA`, `ComandoLimpiarEstadoHPS`, `ComandoLimpiarEspecialidad` | 6 botones idénticos (`Caption = "Comando36"`) | `OnClick` → `ComandoLimpiar*_Click` | Ponen su combo a `Null` y llaman al `AfterUpdate` correspondiente | `:7-35`, `.form.txt:721, 903, 1085, 1267, 1449, 1699` |
| `lbl1` … `lbl5`, `Etiqueta0`, `Etiqueta141` | etiquetas («Nac.», «OTAN», «UE», «ESA», «Y/O», «Estado», «Especialidad») | — | Rótulos; sin lógica | `.form.txt:304, 373, 443, 513, 583, 655, 1637` |
| `Detalle` | sección | — | `Height = 1015`; sin lógica | `.form.txt:281` |

## Flujos

### Apertura (`Form_Load`, `:178-252`)

1. `Set m_ObjColParaFiltros = Nothing` (`:186`): fuerza a que el filtro se reconstruya desde cero.
2. `EstablecerControlesDatos Me, m_Error` aplica la paleta de colores; si devuelve error, eleva 1000 (`:187-190`; `FG.bas:1005-1053`).
3. `RellenarComboEstado Me.ComboEstadoHPS, m_Error` reconstruye el combo de estados; si falla, eleva 1000 (`:191-194`; `FG.bas:1233-1253`).
4. Exige que `m_ObjEntorno.colGradosHPS` exista y contenga las cuatro claves `Nacional`, `OTAN`, `UE` y `ESA`; cada ausencia produce un `m_Error` distinto y `Err.Raise 1000` (`:195-218`).
5. Vuelca cada diccionario en su combo con `RowSource = ""` + `AddItem` en un `For Each`, cuatro veces, una por línea (`:219-242`).
6. Fija `Me.ComboYO = "O"` (`:243`): el valor por defecto es **OR**, no `Y`.
7. En error: compone el mensaje, envía correo al administrador si el error no es interno 1000, muestra `MsgBox` crítico y **cierra el formulario sin guardar** (`:245-251`).

### Cualquier cambio de filtro (`*_AfterUpdate`, `:36-100, 128-177`)

Los siete handlers son idénticos salvo el rótulo del mensaje: `m_Error = ""`, llamada a `Filtrar`, `Err.Raise 1000` si vuelve relleno, y en el manejador, si el error no es 1000, mensaje + `CorreoAlAdministrador` + `MsgBox` (`:36-50` y sus gemelos). El rótulo es copiado: `ComboGradoOtan_AfterUpdate` y `ComboGradoUE_AfterUpdate` escriben literalmente «Al ComboGradoESA…» (`:61, 77`).

### Filtrado (`Filtrar`, `:101-127`)

1. `On Error GoTo errores` (`:103`).
2. Si el diccionario global es `Nothing`, crea un `Scripting.Dictionary` con `CompareMode = TextCompare` (`:104-107`).
3. Escribe **siete** claves, todas normalizadas con `Nz`: `GradoNAC`, `GradoOTAN`, `GradoESA`, `GradoUE`, `Especialidad` y `YO` (`:108-113`), más `Estado` (`:115`). Las seis primeras no son opcionales: su presencia es lo que `getWhereGrado` exige.
4. Llama a `EstablecerFiltroUsuarios Me, m_Error` (`:116`) — **esta llamada sí propaga** `m_Error` con `Err.Raise 1000` (`:117-119`), al contrario que la rama de `FormInicial03ConsultasEstado`.
5. En error: mensaje + correo + `MsgBox` (`:121-127`).

### Traducción a SQL (`Filtro.bas`, `:53-68, 214-342`)

1. `EstablecerFiltroUsuarios` exige que `FormInicial00Principal` esté abierto; si no, sale sin tocar nada (`Filtro.bas:21-23`). Exige además que el diccionario exista y contenga `"Estado"` (`Filtro.bas:31-39`).
2. Calcula siempre `m_WhereEstado = getWhereEstado(p_Error)` **antes** de ramificar (`Filtro.bas:41-43`).
3. En la rama `p_Formulario.Name = "FormInicial03ConsultasGrado"` (`Filtro.bas:53`) llama a `getWhereGrado` y compone el SQL con los cuatro casos: sin ninguno, solo grado, solo estado, o ambos entre paréntesis unidos por `AND` (`Filtro.bas:54-67`). El `AND` final es fijo, **no** usa el comparador Y/O.
4. Asigna `frm.RecordSource = m_SQL` (`Filtro.bas:67`).
5. `getWhereGrado` (`Filtro.bas:214`) falla con «Sin datos» si el diccionario es `Nothing` (`Filtro.bas:241-244`) y con **«Sin datos suficientes»** si falta cualquiera de las seis claves `GradoNAC`, `GradoOTAN`, `GradoESA`, `GradoUE`, `Especialidad` o `YO` (`Filtro.bas:245-253`).
6. Lee las seis claves (`Filtro.bas:254-259`) y elige el comparador: `" AND "` si `YO = "Y"`, `" OR "` en cualquier otro caso (`Filtro.bas:261-265`).
7. Traduce cada línea: si el valor es `"No Posee"` emite `HPS_<LÍNEA>_ESTADO='No Posee'`; si no está vacío emite `HPS_<LÍNEA>_Grado='<valor>'`; si está vacío no emite nada (`Filtro.bas:266-292`).
8. Traduce la especialidad a un `OR` de las cuatro columnas `HPS_*_Especialidad` (`Filtro.bas:294-301`).
9. Si alguna línea o la especialidad está informativa, acumula las líneas no vacías en una `Collection` y las une con el comparador (`Filtro.bas:304-330`); si solo hay especialidad, la usa ella sola (`Filtro.bas:332-333`); si hay ambas, las une con el comparador (`Filtro.bas:335-338`).

### Errores

Todos los handlers siguen el patrón: `If Err.Number <> 1000` compone el mensaje con `vbNewLine`, llama a `CorreoAlAdministrador` y muestra `MsgBox` crítico guardando el resultado en `pregunta` (`VariablesEntorno.bas:220`). Solo `Form_Load` cierra el formulario (`:251`).

## Reglas, validaciones y cálculos

- **Cadena de seis claves obligatorias.** `getWhereGrado` no tolera diccionarios parciales: exige las seis (`Filtro.bas:245-253`). Por eso `Filtrar` escribe las seis siempre, aunque estén vacías (`:108-113`): la clave vacía es un valor válido, la clave ausente es un error. Es la diferencia estructural con `FormInicial03ConsultasEstado`, que solo necesita `"Estado"`.
- **Y/O solo afecta a los grados.** El conmutador `ComboYO` se traduce en `m_Comparador` (`Filtro.bas:261-265`) y se usa para unir las cuatro líneas HPS **y** para añadir la especialidad (`Filtro.bas:336`). No afecta a la combinación con el estado, que es siempre `AND` (`Filtro.bas:64-66`).
- **`No Posee` es una palabra reservada del dominio.** Se compara en el código, no contra la configuración, y produce una columna distinta (`_ESTADO` en vez de `_Grado`) (`Filtro.bas:266-292`). También aparece como último elemento de las listas de los combos (`.form.txt:340, 410, 480, 550`).
- **El `RowSource` del binario es código muerto.** `Form_Load` vacía los cuatro combos de grado con `cmb.RowSource = ""` y los repuebla desde `m_ObjEntorno.colGradosHPS` (`:219-242`); las listas literales del `.form.txt` (`.form.txt:340, 410, 480, 550`) solo existen hasta la primera apertura.
- **Concatenación de valores en el SQL.** `getWhereGrado` interpola el grado y la especialidad directamente en el predicado (`Filtro.bas:269, 275, 281, 287, 295-301`) sin parametrizar. Los valores vienen de listas fijas de combos, lo que acota el riesgo, pero no lo elimina.
- **Guard de contenedor abierto.** Sin `FormInicial00Principal` abierto el filtro no hace nada y no avisa (`Filtro.bas:21-23`).
- **Estado desconocido.** `getWhereEstado` falla con «No existe el estado indicado» si el valor no está en `m_ObjEntorno.ColVisiblesHPSEstados` (`Filtro.bas:181-184`).
- **`Form_Load` no filtra.** Solo prepara los combos; la primera lista de resultados se obtiene desde `Form_Load` de `FormInicial02ConsultasPrincipal`, que llama a `EstablecerFiltroUsuarios` con el diccionario recién vaciado (`:186`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Listado de resultados | `TbDatosLocal` (`SELECT * FROM TbDatosLocal …`) | `RecordSource` asignado al subform de resultados | `Filtro.bas:28-29, 67` |
| Grados disponibles por línea | `m_ObjEntorno.colGradosHPS("Nacional"\|"OTAN"\|"UE"\|"ESA")` | diccionario en memoria, sin consulta | `:203-241`; `Entorno.cls:2399-2405` |
| Estados visibles disponibles | `m_ObjEntorno.ColEstadosVisiblesHPS` | colección en memoria, sin consulta | `:191`; `FG.bas:1245-1246` |

Sin escrituras, sin transacciones.

## Coexistencia

Pendiente de `../coexistencia.md`. Lo concreto: el filtro se propaga por un global compartido (`VariablesEntorno.bas:256`) que todos los formularios de consulta sobrescriben, así que abrir este filtro y otro a la vez interfiere. `getWhereGrado` es la única función de `Filtro.bas` que exige un conjunto fijo de claves, lo que ata este formulario a la forma del diccionario: cualquier filtro nuevo debe respetar el mismo contrato o ramificar antes de la llamada. El `DoCmd.Close` del manejador (`:251`), el `MsgBox` y el `DoCmd.Hourglass` no tienen equivalente directo en web; la comparación `Or`/`And` entre líneas HPS es una regla de negocio que en web conviene modelar como parámetro estructurado y no como SQL.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **Rótulos de error copiados.** `ComboGradoOtan_AfterUpdate` y `ComboGradoUE_AfterUpdate` escriben «Al ComboGradoESA ha devuelto el error» (`:61, 77`): si falla el filtro tras cambiar el grado OTAN o el grado UE, el mensaje nombra el combo equivocado y dificulta el diagnóstico.
- **La condición de entrada a `getWhereGrado` compara las variables equivocadas** (`Filtro.bas:304`): `Not (m_LineaNAC = "" And m_LineaOTAN = "" And m_ValorUE = "" And m_ValorESA = "" And …)` usa `m_ValorUE` y `m_ValorESA` donde debería usar `m_LineaUE` y `m_LineaESA`. El caso divergente es grado UE o ESA = `"No Posee"`: la variable de valor es no vacía, así que entra, pero la línea se añade solo si `m_LineaUE <> ""` (`Filtro.bas:312-317`), que sí se cumple. El resultado observable es el mismo salvo en el caso combinado «NAC vacío + UE=`No Posee` + especialidad vacía»: entra al bloque, `m_ObjColLineas` queda vacía, `m_WhereGrado` queda `""` y `getWhereGrado` devuelve `""` (`Filtro.bas:318-331, 341`) — el mismo resultado que no entrar. Defecto latente, no observable (a verificar con datos reales).
- **`getWhereGrado` deja un espacio sobrante.** `m_Resto` acumula con `" "` al final (`Filtro.bas:321, 323`) y el recorte de `Right(m_Resto, Len(m_Comparador))` (`Filtro.bas:326-327`) solo funciona cuando el último separador acaba la cadena; con el comparador de 5 caracteres (`" AND "`) puede no recortar como se espera. El `WHERE` resultante lleva espacios duplicados; Access los tolera.
- **`Err.Raise 100` en vez de `1000`.** No aplica a este formulario, pero el helper compartido de borrado sí; verificar el mismo patrón en los formularios hermanos.
- **`ComboYO = "O"` se fija por código** (`:243`) ignorando el valor guardado en el binario (`.form.txt:621`): el usuario no ve nunca `Y` por defecto y el conmutador no tiene etiqueta explicativa más allá de «Y/O» (`.form.txt:584`).
- **Los seis botones de limpiar son clones** con `Caption = "Comando36"` (`.form.txt:722, 904, 1086, 1268, 1450, 1700`): texto no descriptivo, sin imagen ni accesibilidad, y `TabStop = NotDefault` los saca del recorrido de tabulación.
- **El filtro se propaga por global.** Dos consultas abiertas a la vez se pisan los diccionarios; además `Filtrar` solo crea el diccionario si es `Nothing` (`:104`), de modo que conserva claves de otros formularios y el `Filtro.bas:245-253` puede dar «Sin datos suficientes» por claves prestadas.
- **`pregunta` es un `Long` global** (`VariablesEntorno.bas:220`) cuyo valor nunca se lee en este formulario.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: parámetros vinculados en lugar de `RecordSource` compuesto; el conmutador Y/O como parámetro explícito de la API en lugar de una cadena SQL; tabla de equivalencia grado→columna en lugar de las cuatro ramas `If`; y valor por defecto declarado («cualquiera de las líneas») en lugar del `OR` implícito.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, los cuatro combos de grado muestran los grados de `NAC`, `OTAN`, `UE` y `ESA` respectivamente, y el conmutador Y/O está en «O».
2. Elegir un grado en una línea restringe la lista a los usuarios que lo tienen en esa línea.
3. Elegir «No Posee» en una línea filtra por `HPS_<LÍNEA>_ESTADO='No Posee'`, no por la columna de grado.
4. Elegir especialidad filtra por la especialidad en cualquiera de las cuatro líneas (OR interno).
5. Con el conmutador en «O» y dos líneas elegidas, un usuario que cumple solo una de ellas aparece; con el conmutador en «Y», solo aparecen los que cumplen las dos.
6. Elegir grado y estado a la vez restringe por ambos (AND entre paréntesis).
7. Los seis botones de limpiar devuelven su criterio a vacío y refiltran.
8. Si falta cualquiera de las seis claves en el diccionario de filtros, el sistema comunica «Sin datos suficientes» en vez de devolver una lista sin filtrar.

## Pendientes

- Contenido de `m_ObjEntorno.colGradosHPS`: se construye en `Entorno.cls:2399-2405`, fuera de este formulario; confirmar de qué tabla o configuración lo carga y si incluye «No Posee».
- Quién llama a `EstablecerFiltroUsuarios` la primera vez: `Form_Load` no filtra; el arranque real está en `Form_FormInicial02ConsultasPrincipal`.
- Relación exacta con `Form_FormInicial03ConsultasEstado`, que comparte `EstablecerFiltroUsuarios` por otra rama (`Filtro.bas:44-52`) y comparte el control `ComboEstadoHPS`: los dos formularios escriben la misma clave `"Estado"` del mismo global.