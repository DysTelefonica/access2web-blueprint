# Form_FormInicial03ConsultasNombre

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial03ConsultasNombre.cls` (124 líneas) y `.form.txt` (1192 líneas) · **PRD:** [README](../README.md)

Formulario de consulta que filtra los usuarios por **DNI**, **nombre y apellidos** y estado HPS a la vez. No escribe ningún dato: escribe tres claves en el diccionario global `m_ObjColParaFiltros` y delega el filtrado real en `EstablecerFiltroUsuarios`, que reescribe el `RecordSource` del subformulario de resultados. Es **no enlazado** (el `.form.txt` no declara `RecordSource`). Es el `formBatch` más rico de la tanda: su clave `"Nombre"` sí alcanza la rama genérica de `getWhereParametro`, que `Form_FormInicial03ConsultasEstado` no llega a tocar. Las citas son `Form_FormInicial03ConsultasNombre.cls:línea` salvo indicación; `Filtro.bas` es `modules/Filtro.bas` y `FG.bas`, `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial02ConsultasPrincipal` | Se monta como `SubFormParametrosConsulta`; el padre resuelve el nombre desde el ítem de `ListaConsultas` | `Form_FormInicial02ConsultasPrincipal.cls:263, 272`, `.form.txt:589-591` |
| `Form_FormInicial00Principal` → `SubFormCentral` → `SubFormResultados` | Destino del filtro: `EstablecerFiltroUsuarios` le asigna el `RecordSource` | `Filtro.bas:21-26, 107` |

No abre ningún formulario: el padre elige el formulario de parámetros por colección `m_ObjEntorno.ColTiposConsultasFormulario`, no por `DoCmd.OpenForm` (`Form_FormInicial02ConsultasPrincipal.cls:255-263`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Nombre` | cuadro de texto | `OnExit` → `Nombre_Exit` | Publica el texto como filtro `"Nombre"` y refiltra | `:48-63`, `.form.txt:413, 424-425` |
| `ComboDNI` | combo (`Value List`) | `AfterUpdate` → `ComboDNI_AfterUpdate` | Publica el DNI como filtro `"DNI"` y refiltra | `:64-79`, `.form.txt:350, 370-372, 395` |
| `ComboEstadoHPS` | combo (`Value List`) | `AfterUpdate` → `ComboEstadoHPS_AfterUpdate` | Publica el estado como filtro `"Estado"` y refiltra | `:81-96`, `.form.txt:471, 491-493, 496` |
| `ComandoLimpiarNombre` | botón | `Click` → `ComandoLimpiarNombre_Click` | Pone `Nombre` a `Null` y reejecuta `Nombre_Exit` con `Cancel = False` | `:40-43`, `.form.txt:696, 704-706` |
| `ComandoLimpiarDNI` | botón | `Click` → `ComandoLimpiarDNI_Click` | Pone `ComboDNI` a `Null` y reejecuta su `AfterUpdate` | `:35-38`, `.form.txt:878, 886-888` |
| `ComandoLimpiarEstadoHPS` | botón | `Click` → `ComandoLimpiarEstadoHPS_Click` | Pone el combo a `Null` y reejecuta su `AfterUpdate` | `:44-47`, `.form.txt:514, 522-524` |
| `lblNombre`, `lblDNI`, `Etiqueta0` | etiquetas | — | Rótulos «Nombre», «DNI» y «Estado»; sin lógica | `.form.txt:304, 334, 455` |
| `Detalle` | sección | — | Agrupa los ocho controles; sin lógica | `.form.txt:278, 281` |

Los controles de cabecera del `.form.txt` (`.form.txt:140-266`: etiqueta, botón, casilla, cuadro, lista, combo y subformulario sin `Name`) los ignora `EstablecerControlesDatos`, que solo repinta por tipo (`FG.bas:1014-1048`).

## Flujos

### Filtrar (`Filtrar`, `:7-33`)

1. Si `m_ObjColParaFiltros` es `Nothing`, crea un `Scripting.Dictionary` con `CompareMode = TextCompare` (`:11-14`).
2. Si `Nombre` no está vacío, calcula `m_NombreCompleto` con `ParsearDatoParaConsulta` (`:16`) y **lo sobrescribe acto seguido con el texto crudo** (`:17`).
3. Escribe `"Estado"` con `Nz(Me.ComboEstadoHPS, "")` (`:19`), `"Nombre"` con `m_NombreCompleto` (`:20`) y `"DNI"` con `Nz(Me.ComboDNI, "")` (`:21`).
4. Llama a `EstablecerFiltroUsuarios Me, m_Error` (`:22`) y, si `m_Error` no quedó vacío, relanza el error interno 1000 (`:23-25`).
5. `errores:` anexa la descripción y avisa al administrador solo cuando `Err.Number <> 1000` (`:28-31`); el `MsgBox` final muestra el `m_Error` pendiente, guardado en el global `pregunta` (`VariablesEntorno.bas:220`).

### Cambio de control y botones «limpiar» (`:35-96`)

Los tres `AfterUpdate` (`ComboDNI_AfterUpdate` `:64-79`, `ComboEstadoHPS_AfterUpdate` `:81-96`) y `Nombre_Exit` (`:48-63`) comparten un único esqueleto: `On Error GoTo errores`, `m_Error = ""` (`:50, :66, :83`), `Filtrar`, `If m_Error <> "" Then Err.Raise 1000` y, en error, correo al administrador solo si `Err.Number <> 1000` más `MsgBox` crítico. Los tres botones «limpiar» (`:35-47`) ponen su control a `Null` y **llaman directamente al `AfterUpdate`/`Exit`** en vez de simular el evento.

### Apertura (`Form_Load`, `:99-124`)

1. `Set m_ObjColParaFiltros = Nothing` (`:102`), descartando cualquier filtro heredado de otro formulario de consulta.
2. `EstablecerControlesDatos Me, m_Error` (`:103`; `FG.bas:1005-1055`), encadenando `Err.Raise 1000` si falla (`:104-106`).
3. `RellenarCombo "DNI", Me.ComboDNI, , m_Error` (`:108`; `FG.bas:1285-1317` → `Constructor.getListaParaCombo`, `Constructor.bas:3061-3106`), encadenando el raise (`:109-111`).
4. `RellenarComboEstado Me.ComboEstadoHPS, m_Error` (`:112`; `FG.bas:1233-1253`), encadenando el raise (`:113-115`).
5. En error: mensaje, `CorreoAlAdministrador` si `Err.Number <> 1000`, `MsgBox` crítico y `DoCmd.Close acForm, Me.Name, acSaveNo` (`:118-123`).

## Reglas, validaciones y cálculos

- **Filtro combinado de tres claves, todo con `AND`.** El formulario no es `FormInicial03ConsultasEstado` ni `FormInicial03ConsultasGrado`, así que `EstablecerFiltroUsuarios` entra por la rama genérica (`:69-107`): compone `getWhereParametro` (`:70`) y lo une al `WHERE` de estado en las cuatro variantes de `:98-106`.
- **Guard de contenedor abierto.** Si `FormInicial00Principal` no está abierto, la función sale sin tocar nada (`Filtro.bas:21-23`): el filtro se aplica por efecto secundario sobre otro formulario.
- **Exigencia de la clave `Estado`.** Sin diccionario o sin `"Estado"`, `p_Error` se rellena con «No se ha rellenado la colección de parámetros» y se lanza el error interno 1000 (`Filtro.bas:31-38`). Aquí `Estado` se escribe siempre (`:19`), también vacío.
- **Traducción de `Nombre`.** `getWhereParametro` mapea `"Nombre"` a `like'*<valor>*'` sobre la concatenación `[Nombre] & ' ' & [Apellido_1] & ' ' & [Apellido_2] ` (`Filtro.bas:134-137`); `"DNI"` cae en la rama genérica `='<valor>'` (`Filtro.bas:142-146`); `"Estado"` se salta por comparación explícita (`Filtro.bas:130-132`) y el resto se encadena con `AND` (`Filtro.bas:151`).
- **Estado a predicado.** `getWhereEstado` traduce el estado visible elegido a cuatro alternativas HPS (NAC, OTAN, UE, ESA) unidas por `OR` (`Filtro.bas:185-204`); si está vacío, no hay `WHERE` de estado (`Filtro.bas:98-99`), y si no existe en la configuración falla con «No existe el estado indicado» (`Filtro.bas:181-184`).
- **Enmascarado de letras previsto y no aplicado.** `ParsearDatoParaConsulta` sustituye cada letra por `?` (`FG.bas:2124-2134`) usando `EsLetra`, que solo reconoce vocales acentuadas y sin acentuar (`FG.bas:2100-2104`); su resultado se descarta en `:17`.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Listado de resultados | `TbDatosLocal` | `RecordSource` reasignado al subform de resultados | `Filtro.bas:28-29, 107` |
| Lista de DNI disponibles | `TbUsuarios` (backend, vía `getdb()`) | `SELECT DISTINCT TbUsuarios.DNI … WHERE Not Is Null` | `Constructor.bas:3076-3082` |
| Estados visibles disponibles | `m_ObjEntorno.ColEstadosVisiblesHPS` | colección en memoria, sin consulta | `FG.bas:1245-1246` |

Sin escrituras, sin transacciones. Los dos `SELECT` de lectura son `DISTINCT` sobre el backend; el filtro se evalúa contra la tabla local `TbDatosLocal`.

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). A diferencia del resto de la tanda, este filtro **sí concatena texto del usuario dentro del SQL** (`Filtro.bas:136-137, 144`) sin comillas de escape, y solo el estado se traduce a literales fijos. Convive con las otras formas que escriben el mismo global: `FormInicial07UsuariosSICA` escribe `"Nombre"` (`Form_FormInicial07UsuariosSICA.cls:89`) y `FormInicial03ConsultasContratistas` escribe `"Estado"` y `"CadenaContratistas"` (`Form_FormInicial03ConsultasContratistas.cls:56-57`). Los errores llegan al usuario como `MsgBox` y el cierre es un `DoCmd.Close` (`:123`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `m_NombreCompleto = Me.Nombre` (`:17`) descarta el resultado de `ParsearDatoParaConsulta` (`:16`): línea muerta, y el texto crudo queda interpolado en el `LIKE` (`Filtro.bas:136-137`). En consecuencia, una comilla simple en el nombre o en el DNI cierra el literal y rompe el `RecordSource`, porque el camino no contiene ningún `Replace` (`Filtro.bas:136-137, 144`).
- Los `RowSource` del binario son código muerto: los vacían `RellenarCombo` (`FG.bas:1303`) y `RellenarComboEstado` (`FG.bas:1242`) en cada apertura; el de `ComboDNI` (.form.txt:372) además lleva DNIs reales embebidos.
- `m_Error` no se vacía al entrar en `Filtrar` (`:10-14`); la relyga la vacían los tres llamantes (`:50, :66, :83`) pero `Form_Load` llama a `EstablecerControlesDatos` sin vaciarla (`:103`). Ni `EstablecerControlesDatos` (`FG.bas:1005-1055`) ni `RellenarComboEstado` (`FG.bas:1233-1253`) vacían su `p_Error` al entrar, al contrario que `EstablecerFiltroUsuarios` (`Filtro.bas:20`); este último traga el error sin relanzar (`Filtro.bas:110-114`), de modo que la propagación depende solo de la cadena `If m_Error <> "" Then Err.Raise 1000`.
- El filtro viaja por un global compartido (`VariablesEntorno.bas:256`) que todos los formularios de consulta sobrescriben; dos consultas vivas a la vez interfieren. `pregunta` es un `Long` global compartido (`VariablesEntorno.bas:220`) usado solo como retorno de `MsgBox` y nunca leído aquí.
- La máscara de letras de `ParsearDatoParaConsulta` depende de `EsLetra`, que solo reconoce vocales y no distingue mayúsculas (`FG.bas:2100-2104`): si algún día se conectara, convertiría cada vocal en comodín y dejaría consonantes y espacios sin proteger.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtros como parámetros vinculados en lugar de `RecordSource` compuesto; y búsqueda por texto con escapado en servidor en lugar de un `LIKE` concatenado.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con los tres controles vacíos, la lista muestra todos los registros de `TbDatosLocal`.
2. Un DNI restringe la lista a los usuarios con ese DNI; nombre y estado vacíos.
3. Un texto en `Nombre` busca por coincidencia parcial sobre nombre y apellidos concatenados; estado y DNI vacíos.
4. Con dos o los tres controles rellenos, las restricciones se combinan con `AND`.
5. «Limpiar» en cualquiera de los tres devuelve la lista al estado sin ese filtro, sin tocar los otros dos.
6. Si el formulario principal no está abierto, el filtro no altera ninguna otra lista y no se produce error visible.

## Pendientes

- Contenido de `m_ObjEntorno.ColTiposConsultasFormulario`: decide qué ítem de `ListaConsultas` monta este formulario y se construye en `Entorno.cls`, no en `src/`.
- Confirmar en el binario si algún otro punto del flujo vuelve a disparar `Form_Load` sin pasar por `ListaConsultas_Click`; y la relación exacta con `FormInicial03ConsultasGrado`, que comparte `EstablecerFiltroUsuarios` por su rama propia (`Filtro.bas:53-67`) y exige seis claves (`Filtro.bas:245-253`). Esta ficha **no** entra por esa rama.