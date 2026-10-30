# Form_FormInicial03ConsultasMotivoHPS

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial03ConsultasMotivoHPS.cls` (92 líneas) y `.form.txt` (900 líneas) · **PRD:** [README](../README.md)

Formulario de consulta que filtra los usuarios por **motivo HPS** y por **estado HPS**, combinando ambos en un `WHERE`. Comparte casi todo con `Form_FormInicial03ConsultasEstado`: la diferencia es que aquí los dos filtros pasan por un procedimiento `Filtrar` común. No escribe ningún dato; el efecto es un `RecordSource` nuevo en el subformulario de resultados del formulario principal. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`). Las citas son `Form_FormInicial03ConsultasMotivoHPS.cls:línea` salvo indicación; `Filtro.bas` es `modules/Filtro.bas` y `FG.bas`, `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial00Principal` | Se instala como subformulario del panel de consultas | no hay `DoCmd.OpenForm` en `src` |
| `Form_FormInicial00Principal` → `SubFormCentral` → `SubFormResultados` | Destino del filtro: `EstablecerFiltroUsuarios` le asigna el `RecordSource` | `Filtro.bas:21-26` |

No abre ningún formulario.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Prepara estilo y rellena los dos combos | `:68-92` |
| `ComboMotivoHPS` | combo (`RowSourceType = Value List`) | `AfterUpdate` → `ComboMotivoHPS_AfterUpdate` | Publica el motivo y refiltra | `:35-49`, `.form.txt:338` |
| `ComboEstadoHPS` | combo (`RowSourceType = Value List`) | `AfterUpdate` → `ComboEstadoHPS_AfterUpdate` | Publica el estado y refiltra | `:52-66`, `.form.txt:413` |
| `ComandoLimpiarMotivoHPS` | botón | `Click` → `ComandoLimpiarMotivoHPS_Click` | Vacía el combo de motivo y refiltra | `:27-29`, `.form.txt:626` |
| `ComandoLimpiarEstadoHPS` | botón | `Click` → `ComandoLimpiarEstadoHPS_Click` | Vacía el combo de estado y refiltra | `:31-33`, `.form.txt:444` |
| `Etiqueta0` | etiqueta | — | Rótulo | `.form.txt:377` |
| `lbl1` | etiqueta («Motivo») | — | Rótulo del combo de motivo | `.form.txt:303` |
| `Detalle` | marco | — | Agrupa los controles; sin lógica | `.form.txt:280` |

## Flujos

### Apertura (`Form_Load`, `:68-92`)

1. Pone `m_ObjColParaFiltros` a `Nothing` (`:71`).
2. `EstablecerControlesDatos Me, m_Error` aplica la paleta de colores (`:73`; `FG.bas:1005-1053`).
3. `RellenarCombo "Motivo_HPS", Me.ComboMotivoHPS, , m_Error` rellena el combo de motivos desde el campo `Motivo_HPS` (`:77`; `FG.bas:1285-1303`).
4. `RellenarComboEstado Me.ComboEstadoHPS, m_Error` rellena el de estados (`:81`; `FG.bas:1233-1253`).
5. Cada paso encadena `Err.Raise 1000` si `m_Error` no está vacío (`:74-75, 78-79, 82-83`).
6. En error: correo al administrador, `MsgBox` crítico y cierre sin guardar (`:85-92`).

### Cambio de motivo o de estado (`ComboMotivoHPS_AfterUpdate` `:35-49`, `ComboEstadoHPS_AfterUpdate` `:52-66`)

Ambos hacen lo mismo y solo cambian el rótulo del mensaje de error (`:38` y `:55`):

1. Vacían `m_Error` (`:37`, `:54`).
2. Llaman a `Filtrar` (`:39`, `:56`).
3. Si `Filtrar` devolvió error, relanzan 1000 para caer en el manejador (`:40-42`, `:57-59`).

### Filtrado (`Filtrar`, `:6-25`)

1. Crea el `Scripting.Dictionary` con `CompareMode = TextCompare` si no existe (`:8-11`).
2. Escribe `"Estado"` y `"Motivo_HPS"` con `Nz`, de modo que combo vacío equivale a filtro ausente (`:13-14`).
3. Llama a `EstablecerFiltroUsuarios` (`:15`) y encadena error (`:16-18`).

### Errores

Los tres manejadores repiten el patrón: `Err.Number <> 1000` ⇒ anexar descripción, `CorreoAlAdministrador` y `MsgBox` crítico (`:20-25`, `:44-49`, `:61-66`). El retorno del `MsgBox` va al global `pregunta` (`VariablesEntorno.bas:220`).

## Reglas, validaciones y cálculos

- **Combinación por `AND`.** Como el nombre del formulario no es ni `FormInicial03ConsultasEstado` ni `FormInicial03ConsultasGrado`, `EstablecerFiltroUsuarios` cae en la rama general: `getWhereParametro` para los campos del diccionario más `getWhereEstado` para el estado, unidos con `WHERE (…) AND (…)` (`Filtro.bas:69-107`).
- **Normalización de motivos.** `getWhereParametro` recorre todas las claves del diccionario salvo `"Estado"` (`Filtro.bas:129-132`). Para una clave no especial la compara con `='<valor>'` (`Filtro.bas:143-145`), así que `Motivo_HPS` se compara por igualdad exacta.
- **Contraste con `Nombre`.** La clave `"Nombre"` tiene un tratamiento aparte —concatena `[Nombre] & ' ' & [Apellido_1] & ' ' & [Apellido_2]` y usa `like'*valor*'` (`Filtro.bas:134-137`)—, que este formulario no usa pero que comparte el mismo constructor de `WHERE`.
- **Exigencia de la clave `Estado`.** Igual que en el formulario de estado, `EstablecerFiltroUsuarios` falla con «No se ha rellenado la colección de parámetros» si falta (`Filtro.bas:31-39`).
- **Traducción de estado.** `getWhereEstado` traduce el estado visible a la cadena de cuatro `OR` sobre NAC, OTAN, UE y ESA (`Filtro.bas:163-206`).
- **Guard de contenedor abierto.** Sin `FormInicial00Principal` abierto, no se filtra nada y no se avisa (`Filtro.bas:21-23`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Listado de resultados | `TbDatosLocal` | `RecordSource` asignado al subform de resultados | `Filtro.bas:28-29, 100-109` |
| Valores de motivo | campo `Motivo_HPS` | colección en memoria vía `getListaParaCombo` | `:77`, `FG.bas:1285-1303` |
| Estados visibles disponibles | `m_ObjEntorno.ColEstadosVisiblesHPS` | colección en memoria | `:81`, `FG.bas:1245-1246` |

Sin escrituras, sin transacciones.

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). El valor del motivo elegido **sí se concatena dentro del literal SQL** (`Filtro.bas:143-145`), a diferencia del estado. Un motivo con comilla simple rompería la sentencia; el patrón web equivalente es un parámetro vinculado. El `WHERE` se construye con `If/ElseIf` sobre cuatro combinaciones de vacío/no vacío (`Filtro.bas:98-107`), una forma que no se traslada tal cual.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El motivo se interpola en el `WHERE` sin escapar (`Filtro.bas:143-145`). Un valor con apóstrofo produce un error de SQL en tiempo de ejecución.
- Los dos `RowSource` vienen escritos en el binario como listas de valores (`.form.txt:340, 415`) pero `RellenarCombo` y `RellenarComboEstado` los vacían y reconstruyen al abrir (`:77, 81`; `FG.bas:1242, 1245-1246`). La lista de motivos incrustada en el binario incluye denominaciones reales de clientes y queda como código muerto: no debe propagarse al modelo web.
- Si el usuario elige motivo y estado, `getWhereParametro` y `getWhereEstado` se combinan entre paréntesis (`Filtro.bas:105`); si solo elige motivo, el `WHERE` no lleva paréntesis. El resultado es equivalente, pero la asimetría sugiere que la construcción se hizo por acumulación.
- Estado compartido mutable: los dos combos escriben en el global `m_ObjColParaFiltros` (`VariablesEntorno.bas:256`), que cualquier otro formulario de consulta puede sobrescribir entre la selección y el filtrado.
- Los dos `AfterUpdate` almost idénticos repiten el patrón de error; el de motivo (`:35-49`) y el de estado (`:52-66`) solo cambian la cadena del mensaje.
- `pregunta` es un `Long` global (`VariablesEntorno.bas:220`) cuyo valor se descarta.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: motivo como dimensión de consulta de catálogo en lugar de columna de texto; parámetros vinculados en lugar de `WHERE` compuesto; y motivo con identificador estable en vez de la denominación.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin motivo y sin estado, la lista muestra todos los registros de `TbDatosLocal`.
2. Elegir solo motivo restringe la lista a los usuarios con ese motivo exacto.
3. Elegir solo estado restringe por estado, según las cuatro líneas HPS.
4. Elegir ambos devuelve la intersección de ambos criterios.
5. «Limpiar» en cualquiera de los dos combos deja su criterio fuera sin tocar el otro.
6. Un motivo con apóstrofo se trata como valor literal y no produce error.

## Pendientes

- De dónde saca `getListaParaCombo` la lista de `Motivo_HPS` y si esa lista es configurable o está codificada en la base.
- Contenido de `m_ObjEntorno.ColEstadosVisiblesHPS` y `ColVisiblesHPSEstados`, construidos en `Entorno.cls`.
- Si algún menú o macro del binario abre este formulario directamente: no hay `DoCmd.OpenForm` en `src`.