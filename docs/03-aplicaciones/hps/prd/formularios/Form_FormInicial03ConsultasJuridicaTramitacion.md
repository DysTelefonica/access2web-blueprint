# Form_FormInicial03ConsultasJuridicaTramitacion

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial03ConsultasJuridicaTramitacion.cls` (101 líneas) y `.form.txt` (902 líneas) · **PRD:** [README](../README.md)

Formulario de consulta que filtra los usuarios por la **empresa que tramita jurídicamente** y por el **estado HPS**. Es el único de la tanda que resuelve un identificador contra la base antes de filtrar: el combo guarda `IDSuministrador;Nombre` y `Filtrar` recupera el nombre de la empresa para usarlo como criterio. No escribe ningún dato; el efecto es un `RecordSource` nuevo en el subformulario de resultados del formulario principal. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`). Las citas son `Form_FormInicial03ConsultasJuridicaTramitacion.cls:línea` salvo indicación; `Filtro.bas` es `modules/Filtro.bas`, `FG.bas`, `modules/Funciones Generales.bas` y `Constructor.bas`, `modules/Constructor.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial00Principal` | Se instala como subformulario del panel de consultas | no hay `DoCmd.OpenForm` en `src` |
| `Form_FormInicial00Principal` → `SubFormCentral` → `SubFormResultados` | Destino del filtro: `EstablecerFiltroUsuarios` le asigna el `RecordSource` | `Filtro.bas:21-26` |

No abre ningún formulario.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Prepara estilo y rellena los dos combos | `:76-101` |
| `IDEmpresaHPS` | combo de dos columnas | `AfterUpdate` → `IDEmpresaHPS_AfterUpdate` | Publica la empresa elegida y refiltra | `:44-58`, `.form.txt:335` |
| `ComboEstadoHPS` | combo (`RowSourceType = Value List`) | `AfterUpdate` → `ComboEstadoHPS_AfterUpdate` | Publica el estado y refiltra | `:60-74`, `.form.txt:406` |
| `ComandoLimpiarJuridicaQueTramitaHPS` | botón | `Click` → `ComandoLimpiarJuridicaQueTramitaHPS_Click` | Vacía el combo de empresa y refiltra | `:36-38`, `.form.txt:619` |
| `ComandoLimpiarEstadoHPS` | botón | `Click` → `ComandoLimpiarEstadoHPS_Click` | Vacía el combo de estado y refiltra | `:40-42`, `.form.txt:437` |
| `Etiqueta0` | etiqueta | — | Rótulo | `.form.txt:370` |
| `lbl1` | etiqueta | — | Rótulo del combo de empresa | `.form.txt:299` |
| `Detalle` | marco | — | Agrupa los controles; sin lógica | `.form.txt:276` |

## Flujos

### Apertura (`Form_Load`, `:76-101`)

1. Vacía `m_Error` y pone `m_ObjColParaFiltros` a `Nothing` (`:79-71`).
2. `EstablecerControlesDatos Me, m_Error` aplica la paleta de colores (`:82`; `FG.bas:1005-1053`).
3. `RellenarComboEmpresas Me.IDEmpresaHPS, EnumSiNo.Sí, , m_Error` rellena el combo **solo con empresas tramitadoras** (`:86`; `FG.bas:1100-1130`).
4. `RellenarComboEstado Me.ComboEstadoHPS, m_Error` rellena el de estados (`:90`; `FG.bas:1233-1253`).
5. Cada paso encadena `Err.Raise 1000` si `m_Error` no está vacío (`:83-84, 87-88, 91-92`).
6. En error: correo al administrador, `MsgBox` crítico y cierre sin guardar (`:94-101`).

### Cambio de empresa (`IDEmpresaHPS_AfterUpdate`, `:44-58`) y de estado (`ComboEstadoHPS_AfterUpdate`, `:60-74`)

Idénticos salvo el rótulo del mensaje: vacían `m_Error`, llaman a `Filtrar` y relanzan 1000 si el filtro falló (`:46-50`, `:62-66`).

### Filtrado (`Filtrar`, `:9-33`)

1. Crea el `Scripting.Dictionary` con `CompareMode = TextCompare` si no existe (`:11-14`).
2. Escribe `"Estado"` con el valor del combo (`:17`).
3. Resuelve el empresa: `Constructor.getSuministrador(p_ID:=Nz(Me.IDEmpresaHPS.Column(0), ""))` (`:18`).
4. Si se encontró, escribe `"EmpresaTramitadora"` con **el nombre** del suministrador (`:19-21`). Si no, **no borra la clave anterior**: puede quedar el valor de la empresa filtrada en una ejecución previa.
5. Llama a `EstablecerFiltroUsuarios` y encadena error (`:23-26`).

### Errores

Los tres manejadores repiten el patrón `Err.Number <> 1000` ⇒ anexar descripción, `CorreoAlAdministrador`, `MsgBox` crítico (`:28-33`, `:53-58`, `:69-74`). El retorno del `MsgBox` va al global `pregunta` (`VariablesEntorno.bas:220`).

## Reglas, validaciones y cálculos

- **El filtro usa el nombre, no el identificador.** La clave `"EmpresaTramitadora"` recibe `m_Suministrador.Nombre` (`:20`), y `getWhereParametro` la compara por igualdad exacta (`Filtro.bas:143-145`). Un cambio de denominación en `TbSuministradores` invalida los filtros guardados y cualquier aviso calculado sobre ese texto.
- **Lectura por clave primaria sin comillas.** `getSuministrador` concatena el identificador directo: `"WHERE IDSuministrador=" & p_ID` (`Constructor.bas:40-41`). El identificador proviene de la columna 0 del combo, no de texto tecleado, lo que evita el riesgo de inyección aquí; sí lo habría si el control fuera editable.
- **Sin empresa ⇒ clave intacta.** Si `getSuministrador` devuelve `Nothing` —no existe el identificador o no se pasó ninguno— (`Constructor.bas:49-52`), el `If Not … Is Nothing` impide escribir, pero no elimina `"EmpresaTramitadora"` del diccionario (`:19-21`). Al filtrar, la clave previa sigue filtrando.
- **Combinación por `AND`.** Rama general de `EstablecerFiltroUsuarios`: `getWhereParametro` más `getWhereEstado`, unidos como `WHERE (…) AND (…)` (`Filtro.bas:69-107`).
- **Exigencia de la clave `Estado`.** Falla con «No se ha rellenado la colección de parámetros» si falta (`Filtro.bas:31-39`).
- **Guard de contenedor abierto.** Sin `FormInicial00Principal` abierto, no se filtra nada y no se avisa (`Filtro.bas:21-23`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Resolver empresa desde el combo | `TbSuministradores` | SQL concatenado por `OpenRecordset`, solo lectura | `Constructor.bas:40-41, 47` |
| Listado de resultados | `TbDatosLocal` | `RecordSource` asignado al subform de resultados | `Filtro.bas:28-29, 100-109` |
| Empresas tramitadoras del combo | `m_ObjEntorno.ColEmpresasTramitadoras` | colección en memoria | `FG.bas:1119` |
| Estados visibles disponibles | `m_ObjEntorno.ColEstadosVisiblesHPS` | colección en memoria | `:90`, `FG.bas:1245-1246` |

Sin escrituras, sin transacciones.

## Coexistencia

Pendiente de `../coexistencia.md`. Este es el único formulario de la tanda que **lee de la base durante el filtrado**: `Constructor.getSuministrador` abre un recordset nuevo por cada cambio de combo (`:18`). En web, resolver el nombre en el cliente o en un endpoint de referencia evita el N+1 por interacción. El identificador concatenado sin comillas (`Constructor.bas:40-41`) depende de que llegue siempre como número.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **La clave obsoleta sobrevive al limpiado.** «Limpiar» pone `IDEmpresaHPS` a `Null` (`:37`), `getSuministrador` devuelve `Nothing` y el `If` de `:19-21` no hace nada; `"EmpresaTramitadora"` conserva el nombre anterior y **el botón de limpiar no limpia**. Para limpiar de verdad habría que quitar la clave o escribirla vacía.
- `Nz(Me.IDEmpresaHPS.Column(0), "")` (`:18`) llama a `.Column` sobre un combo recién puesto a `Null`; Access devuelve `Null` y `Nz` lo neutraliza, pero el comportamiento de `.Column` sin fila seleccionada depende del motor y merece verificación en el binario.
- `RellenarComboEmpresas` se invoca con `EnumSiNo.Sí` en el primer parámetro y el cuarto argumento posicional (`m_Error`) vacío (`:86`). El salto de argumentos depende de que `p_Error` sea el cuarto formal; si se añadiera un parámetro intermedio, la llamada pasaría a escribir en él.
- El `RowSource` de `ComboEstadoHPS` viene escrito en el binario (`.form.txt:408-409`) y luego lo vacía y reconstruye `RellenarComboEstado` (`:90`; `FG.bas:1242`): lista del binario como código muerto.
- Filtrar por **nombre** en lugar de por identificador hace el filtro frágil ante cambios de denominación (`:20`; `Filtro.bas:143-145`).
- Un `OpenRecordset` por interacción sin `Close` explícito en la ruta de éxito de este formulario (lo cierra `Constructor.bas:41` solo en el caso `EOF`; el resto depende del final de la función).
- Estado compartido mutable: los dos combos escriben en el global `m_ObjColParaFiltros` (`VariablesEntorno.bas:256`).
- `pregunta` es un `Long` global (`VariablesEntorno.bas:220`) cuyo valor se descarta.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtrar por `IDSuministrador` con tabla de referencia en lugar de por nombre; catálogos precargados en el cliente para no consultar en cada cambio; y modelo de filtros declarativo con estado vacío explícito en lugar de ausencia de clave.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin empresa y sin estado, la lista muestra todos los registros de `TbDatosLocal`.
2. Elegir una empresa restringe la lista a los usuarios cuyo empresa tramitadora es la elegida.
3. Elegir solo estado restringe por estado, según las cuatro líneas HPS.
4. Elegir ambos devuelve la intersección de ambos criterios.
5. «Limpiar» en cualquiera de los dos combos deja su criterio fuera sin tocar el otro.
6. Si la empresa seleccionada deja de existir en `TbSuministradores`, el filtro deja de aplicarse en vez de conservar el valor anterior.

## Pendientes

- Cómo se puebla `m_ObjEntorno.ColEmpresasTramitadoras` y con qué criterio una empresa se considera tramitadora (`FG.bas:1119` la toma del entorno).
- Contenido de `m_ObjEntorno.ColEstadosVisiblesHPS` y `ColVisiblesHPSEstados`, construidos en `Entorno.cls`.
- Si algún menú o macro del binario abre este formulario directamente: no hay `DoCmd.OpenForm` en `src`.
- Si `IDEmpresaHPS` admite edición manual: el `.form.txt` marca `AllowValueListEdits`; conviene confirmarlo antes de fijar el criterio de aceptación 6.