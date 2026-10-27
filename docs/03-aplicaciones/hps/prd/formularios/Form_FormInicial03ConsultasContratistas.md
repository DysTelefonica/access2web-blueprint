# Form_FormInicial03ConsultasContratistas

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial03ConsultasContratistas.cls` (97 líneas) y `.form.txt` (879 líneas) · **PRD:** [README](../README.md)

Formulario de consulta que cruza **dos** criterios: la jurídica de contratación y el estado HPS. No escribe ningún dato: compone dos claves en el diccionario global `m_ObjColParaFiltros` y delega el filtrado real en `EstablecerFiltroUsuarios`, que reescribe el `RecordSource` del subformulario de resultados del formulario principal. Es el tipo de consulta «2» del catálogo de `m_ObjEntorno` (`Entorno.cls:1775`) y un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`). Comparte esqueleto con `Form_FormInicial03ConsultasEstado` pero, a diferencia de aquel, no tiene rama propia en `EstablecerFiltroUsuarios` y cae en la genérica. Las citas son `Form_FormInicial03ConsultasContratistas.cls:línea` salvo indicación; `Filtro.bas` es `modules/Filtro.bas` y `FG.bas`, `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial00Principal` | Se instala como subformulario del panel de consultas; el contenedor elige cuál según el tipo elegido en `m_ObjEntorno.ColTiposConsultasFormulario` | `Entorno.cls:1775` |
| `Form_FormInicial00Principal` → `SubFormCentral` → `SubFormResultados` | Destino del filtro: `EstablecerFiltroUsuarios` alcanza el subform de resultados por nombre y le asigna el `RecordSource` | `Filtro.bas:21-26, 107` |
| `Constructor.gelCadenaContratistas` → `TbExpedientesConEntidades` | Origen de los valores del combo de jurídica | `Constructor.bas:195-197`, invocado en `Entorno.cls:792` |

No abre ningún formulario.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Vacía el diccionario de filtros, aplica la paleta y rellena los dos combos | `:70-97`, `.form.txt:80` |
| `CadenaContratistas` | combo (`RowSourceType ="Value List"`) | `AfterUpdate` → `CadenaContratistas_AfterUpdate` | Publica la jurídica elegida y refresca los resultados | `:16-31`, `.form.txt:297-319` |
| `ComboEstadoHPS` | combo (`RowSourceType ="Value List"`) | `AfterUpdate` → `ComboEstadoHPS_AfterUpdate` | Publica el estado elegido y refresca los resultados | `:32-47`, `.form.txt:367-392` |
| `ComandoLimpiarJuridicaContrato` | botón («cepillo») | `OnClick` → `ComandoLimpiarJuridicaContrato_Click` | Pone `CadenaContratistas` a `Null` y reaplica el filtro | `:7-10`, `.form.txt:592-603` |
| `ComandoLimpiarEstadoHPS` | botón («cepillo») | `OnClick` → `ComandoLimpiarEstadoHPS_Click` | Pone `ComboEstadoHPS` a `Null` y reaplica el filtro | `:11-14`, `.form.txt:410-421` |
| `lblJuridicaContratacion` | etiqueta («Jurídica Contratación») | — | Rótulo del combo de jurídica | `.form.txt:281-282` |
| `Etiqueta0` | etiqueta («Estado») | — | Rótulo del combo de estado; nombre reutilizado de un control copiado | `.form.txt:351-352` |
| `Detalle` | sección | — | Aloja los seis controles; `Height =1015`, un solo registro | `.form.txt:255-258` |

## Flujos

### Apertura (`Form_Load`, `:70-97`)

1. `Set m_ObjColParaFiltros = Nothing` (`:75`), lo que fuerza a reconstruir el filtro desde cero.
2. `EstablecerControlesDatos Me, m_Error` aplica la paleta de colores a etiquetas, campos y listas (`:76`; `FG.bas:1005-1055`). Si devuelve error, salta al manejador (`:77-79`).
3. `RellenarCadenaContratistas Me.CadenaContratistas, m_Error` vacía el `RowSource` y añade un ítem por clave de `m_ObjEntorno.ColCadenaContratistas` (`:80`; `FG.bas:1147-1179`; `Entorno.cls:783-796`). Si devuelve error, salta al manejador (`:81-83`).
4. `RellenarComboEstado Me.ComboEstadoHPS, m_Error` hace lo propio con `m_ObjEntorno.ColEstadosVisiblesHPS` (`:85`; `FG.bas:1233-1253`; `Entorno.cls:1560-1597`). Si devuelve error, salta al manejador (`:86-88`).
5. En error: compone el mensaje, envía correo al administrador si no es el interno 1000, muestra un `MsgBox` crítico y **cierra el formulario** sin guardar (`:90-97`).

### Cambio de un criterio (`CadenaContratistas_AfterUpdate`, `:16-31`; `ComboEstadoHPS_AfterUpdate`, `:32-47`)

Los dos manejadores son idénticos salvo por el texto del mensaje: `On Error GoTo errores`, `m_Error = ""`, llamada a `Filtrar` y, si vuelve con error, `Err.Raise 1000` (`:17-22`, `:33-38`). No hay ninguna lógica propia: los dos delegan entero.

### Limpiar (`ComandoLimpiarJuridicaContrato_Click`, `:7-10`; `ComandoLimpiarEstadoHPS_Click`, `:11-14`)

Ponen su combo a `Null` y **llaman directamente al `AfterUpdate` correspondiente** (`:9`, `:13`), que equivale a elegir el valor vacío.

### Composición del filtro (`Filtrar`, `:48-69`)

1. Si `m_ObjColParaFiltros` es `Nothing`, crea un `Scripting.Dictionary` con `CompareMode = TextCompare` (`:51-54`).
2. Escribe la clave `"Estado"` con el valor del combo, normalizado con `Nz` para que vacío y `Null` se comporten igual (`:56`).
3. Escribe la clave `"CadenaContratistas"` con el valor del otro combo, también con `Nz` (`:57`).
4. `EstablecerFiltroUsuarios Me, m_Error` (`:58`) y, si vuelve con error, `Err.Raise 1000` (`:59-61`).

### Errores

Los cuatro manejadores repiten el patrón: si `Err.Number <> 1000`, anteponen su prefijo al error, envían `CorreoAlAdministrador` y muestran `MsgBox` crítico (`:26-31`, `:42-47`, `:64-69`, `:91-96`). El valor de retorno del `MsgBox` se guarda en el global `pregunta` (`VariablesEntorno.bas:220`).

## Reglas, validaciones y cálculos

- **Dos criterios acumulativos, ninguno obligatorio.** El `WHERE` final se arma con ambos: si solo hay estado, `WHERE <estado>`; si solo hay jurídica, `WHERE <jurídica>`; si hay los dos, `WHERE (<jurídica>) AND (<estado>)`; si no hay ninguno, `SELECT * FROM TbDatosLocal` sin `WHERE` (`Filtro.bas:98-106`). `EstablecerFiltroUsuarios` solo distingue `FormInicial03ConsultasEstado` y `FormInicial03ConsultasGrado`, así que este cae en el `Else` genérico (`Filtro.bas:44, 53, 69-107`).
- **Guard de contenedor abierto.** Si `FormInicial00Principal` no está abierto, la función sale sin tocar nada (`Filtro.bas:21-23`): el filtro se aplica por efecto secundario sobre otro formulario.
- **Exigencia de la clave `Estado`.** Si el diccionario no existe o no contiene `"Estado"`, `p_Error` se rellena con «No se ha rellenado la colección de parámetros» y se relanza el 1000 (`Filtro.bas:31-38`). `Filtrar` la escribe siempre (`:56`), pero cualquier otro consumidor del global puede romper la colección.
- **La clave del diccionario se usa como nombre de columna.** `getWhereParametro` recorre todas las claves salvo `"Estado"` y usa la clave como columna, con el comparativo `='` por defecto (`Filtro.bas:129-146, 149-151`).
- **Traducción de estado a predicado.** `getWhereEstado` traduce el estado visible elegido a cuatro alternativas HPS (NAC, OTAN, UE, ESA) unidas por `OR` (`Filtro.bas:163-205`). Un estado desconocido falla con «No existe el estado indicado» (`Filtro.bas:181-184`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Valores del combo de jurídica | `TbExpedientesConEntidades` (`gelCadenaContratistas`) | `SELECT DISTINCT CadenaContratistas … WHERE Not … Is Null` | `Constructor.bas:195-197, 209-215` |
| Listado de resultados | `TbDatosLocal` | `RecordSource` asignado al subform de resultados | `Filtro.bas:28-29, 107` |
| Estados visibles disponibles | `m_ObjEntorno.ColEstadosVisiblesHPS` | colección en memoria, nueve claves | `Entorno.cls:1570-1583` |
| Jurídicas disponibles | `m_ObjEntorno.ColCadenaContratistas` | colección en memoria, construida bajo demanda | `Entorno.cls:783-796` |

Sin escrituras, sin transacciones.

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Lo que ya se ve y distingue esta ficha de `Form_FormInicial03ConsultasEstado`: aquí el valor del usuario **sí** se concatena en el SQL, porque `getWhereParametro` usa la clave como columna y el valor con el comparativo `='` (`Filtro.bas:149-151`), sin escapado de comillas simples. La etiqueta de la jurídica es texto libre procedente de `TbExpedientesConEntidades` (`Constructor.bas:195-197`), de modo que el contenido de esa tabla decide si el `WHERE` resultante es válido.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **El nombre de la clave no coincide con el nombre de la columna.** `Filtrar` escribe `"CadenaContratistas"` (`:57`) y `getWhereParametro` la convierte en columna (`Filtro.bas:143, 149`), pero el `INSERT` que puebla `TbUsuariosHistoricosLocal` —y por tanto `TbDatosLocal`— guarda ese campo con el alias `JuridicaContrato` (`Mod_Sincronizacion_Historico.bas:51-52`). Con el esquema del snapshot, el `WHERE` saldría como `WHERE (CadenaContratistas='…')` sobre una tabla sin esa columna; a verificar contra el binario, porque el resultado depende del esquema real de `TbDatosLocal`.
- **Concatenación sin escapado.** `getWhereParametro` interpola el valor entre comillas simples sin duplicarlas ni parametrizar (`Filtro.bas:149-151`): una jurídica con apóstrofo rompe el SQL, y el `m_Error` resultante solo llega al `MsgBox` del manejador (`:64-68`).
- **`getWhereEstado` encadena un `OR` sin espacio** en el estado de renovación: `... HPS_UE_Renovacion='Sí'OR HPS_ESA_Renovacion='Sí'` (`Filtro.bas:202`). Access lo tolera; otros motores no.
- `CadenaContratistas_AfterUpdate` y `ComboEstadoHPS_AfterUpdate` solo miran `m_Error`, no `Err` de `Filtrar` (`:19-22`, `:35-38`): al contrario que `Form_Load`, no encadenan un `Err.Raise` por detrás de cada llamada.
- El `RowSource` del combo de estado viene escrito en el binario (`.form.txt:388-390`) pero `RellenarComboEstado` lo vacía y lo reconstruye en cada apertura (`FG.bas:1242, 1245-1246`): lista del binario como código muerto y fuente de divergencia.
- `EstablecerFiltroUsuarios` no hace nada si `FormInicial00Principal` no está abierto (`Filtro.bas:21-23`), sin avisar: el usuario ve combos que responden y una lista que no cambia.
- El filtro viaja por un global compartido (`VariablesEntorno.bas:256`) que todos los formularios de consulta sobrescriben; dos consultas abiertas a la vez interfieren.
- `pregunta` es un `Long` global compartido (`VariablesEntorno.bas:220`) usado como retorno de `MsgBox` y nunca leído en este formulario.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtros como parámetros vinculados en lugar de `RecordSource` compuesto; tabla de equivalencia clave→columna en lugar de usar la clave como nombre de columna; y que la jurídica sea un dato normalizado de la ficha en vez de texto libre concatenado.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin criterios elegidos, la lista muestra todos los registros de `TbDatosLocal`.
2. Elegir una jurídica restringe la lista a los usuarios con esa jurídica de contratación.
3. Elegir un estado visible restringe la lista a los usuarios que lo tienen activo en al menos una de las cuatro líneas (NAC, OTAN, UE, ESA).
4. Con ambos criterios, la lista muestra la intersección de los dos.
5. «Limpiar» en cualquiera de los dos combos devuelve esa dimensión a «sin filtro» y conserva la otra.
6. Si el formulario principal no está abierto, el filtro no altera ninguna otra lista y no se produce error visible.
7. Un estado o una jurídica que no exista en la configuración se comunica al usuario en vez de dejar la lista sin tocar.

## Pendientes

- Esquema real de `TbDatosLocal`: de él depende si la clave `CadenaContratistas` de `Filtrar` (`:57`) corresponde a una columna existente o al campo `JuridicaContrato` (`Mod_Sincronizacion_Historico.bas:51-52`).
- Contenido de `m_ObjEntorno.ColCadenaContratistas` y `ColEstadosVisiblesHPS`: se construyen en `Entorno.cls` y `Constructor.bas`, no en el formulario.
- Si algún menú o macro del binario abre este formulario directamente: en `src/` solo consta el tipo «2» del catálogo (`Entorno.cls:1775`).
- Relación con `Form_FormInicial03ConsultasJuridicaTramitacion`, que también usa jurídica y comparte el mismo `Else` genérico de `EstablecerFiltroUsuarios`.