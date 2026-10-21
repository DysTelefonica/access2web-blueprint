# Form_FormExpedientesParaCambioTipo

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormExpedientesParaCambioTipo.cls` (321 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Selector auxiliar del cambio de tipo: busca expedientes por palabra clave, los muestra en una lista y devuelve al formulario que lo abrió el expediente elegido, que ha de ser un Acuerdo Marco o un Lote, como futuro padre. Es un formulario **no enlazado** (el `.form.txt` no tiene `RecordSource`); no escribe en la base de datos. Las citas son `Form_FormExpedientesParaCambioTipo.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormExpedienteCambioTipo` | `ComandoBuscar_Click` cierra el selector si ya estaba abierto, lo abre sin `OpenArgs` y lo guarda en `m_Form` (`WithEvents`) | `Form_FormExpedienteCambioTipo.cls:5, 46-53` |

Salida: evento `Seleccionado(m_Expediente)` (`:7`, lanzado en `:181`) y cierre del formulario (`:185`). `Form_FormExpedienteCambioTipo.m_Form_Seleccionado` copia el `IDExpediente` recibido en su control `IDExpedientePadre` (`Form_FormExpedienteCambioTipo.cls:263-270`). `cmdSalir` cierra sin devolver nada (`:13-18`).

No se ha encontrado otro llamador: la búsqueda de `ParaCambioTipo` en `src` solo da `Form_FormExpedienteCambioTipo`.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `PalabraClave` | texto (`Tag = "BUSQUEDA"`) | `OnKeyDown` → `PalabraClave_KeyDown` | Con Intro, si `ComandoPalabraClave` está habilitado, le pasa el foco y lanza la búsqueda | `:304-320`, `.form.txt:623-627` |
| `ComandoPalabraClave` (rótulo «Ver Todos») | botón | `OnClick` | Ejecuta `Filtrar` con la palabra clave | `:114-139`, `.form.txt:529-531` |
| `ComandoIniciarPalabraClave` (rótulo «Ver Todos») | botón | `OnClick` | Vacía `PalabraClave` y llama a `Filtrar` sin parámetro de error | `:109-112`, `.form.txt:653-655` |
| `ListaFiltrados` | lista (`Value List`, 7 columnas, encabezados, `IDExpediente` con ancho 0) | `OnClick` → `ListaFiltrados_Click` | Carga el expediente de la fila y habilita `ComandoSeleccionar` | `:236-300`, `.form.txt:830-851` |
| `ListaFiltrados` | — | `OnDblClick` → `ListaFiltrados_DblClick` | Si `ComandoSeleccionar` está habilitado, lo ejecuta | `:229-234` |
| `ComandoSeleccionar` | botón (deshabilitado en diseño) | `OnClick` | Valida AM o Lote, lanza `Seleccionado` y cierra | `:141-195`, `.form.txt:871, 882-884` |
| `ComandoAyuda` (rótulo «Salir», tooltip «Ayuda») | botón | `OnClick` | `AbrirAyuda` (`FUNCIONES UTILES.bas:3076`) | `:81-106`, `.form.txt:729-733` |
| `cmdSalir` | botón | `OnClick` | Cierra el formulario sin guardar | `:13-18`, `.form.txt:940-942` |

## Flujos

### Apertura (`Form_Open`, `:197-227`)

1. `Ajustar Me` ajusta el tamaño de la ventana (`:206`; `FUNCIONES UTILES.bas:2703`).
2. Título de ventana = global `m_TituloFormulario` (`:208`); el rótulo fijo del diseño queda sustituido (`.form.txt:41`).
3. La lista se reinicia solo con la fila de encabezados `IDExpediente;Tipo;NºExp;Título;FInicio;FFin;FFinGarantía` (`:209`) y el foco va a `PalabraClave` (`:210`).
4. `AvanceCerrar` oculta el indicador de progreso (`:212`; `FUNCIONES UTILES.bas:2770`).

### Búsqueda (`Filtrar`, `:20-78`)

1. Deshabilita `ComandoSeleccionar`, olvida la selección previa y reinicia la lista a los encabezados (`:37-41`).
2. Con `PalabraClave` vacía sale sin mensaje y con la lista vacía (`:43-45`, `:74`).
3. `getColBusquedaParaCambio` (`FUNCIONES UTILES.bas:1214-1268`) devuelve un diccionario de `Expediente` por `IDExpediente` (`FUNCIONES UTILES.bas:1248-1254`); sin resultados, deja la lista vacía (`:51-54`).
4. Añade una fila por expediente: `IDExpediente`, `TipoParaLista`, `CodExp`, `Titulo`, `FechaInicioContrato`, `FechaFinContrato`, `FechaFinGarantia` (`:55-62`).
5. Selecciona el índice 2 de la lista y llama a `ListaFiltrados_Click` (`:66-67`).

### Selección (`ListaFiltrados_Click`, `:236-300`, y `ComandoSeleccionar_Click`, `:141-195`)

1. Al pulsar una fila se relee el expediente con `constructor.getExpediente` si cambia el `IDExpediente` (`:257-274`) y se habilita `ComandoSeleccionar` (`:283`). Fila vacía o expediente inexistente: se deshabilita (`:247-256`, `:275-282`).
2. Al seleccionar se vuelve a comprobar la fila: sin fila, «No ha seleccionado ningún elemento de la lista» (`:149-163`).
3. Si `TipoCalculado` no es `AM` ni `Lote`, «Se ha de seleccionar un AM o un Lote» (`:177-180`). `TipoCalculado` sale de `EsAM`, `EsLote`, `EsExpediente` y `EsBasado`, por ese orden (`FUNCIONES UTILES.bas:5417-5450`).
4. Lanza `Seleccionado` y cierra (`:181-185`).

### Errores

Todos los handlers con `On Error GoTo` acumulan el texto en `m_Error`, envían `CorreoAlAdministrador` salvo el error interno 1000 y muestran `MsgBox` crítico (p. ej. `:130-138`, `:187-194`). `Form_Open` además cierra el formulario (`:226`). `Filtrar` solo devuelve el texto en `p_Error` (`:72-76`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Búsqueda con `Like '*<palabra>*'` sobre `CodExp`, `Titulo` o `Nemotecnico`, sin filtrar por tipo ni estado | `getColBusquedaParaCambio` | `FUNCIONES UTILES.bas:1226-1230` |
| Sin palabra clave no se busca (no hay «ver todos» real) | `Filtrar` | `:43-45` |
| Columna «Tipo» = «AM» si `EsAM`, «C» si `EsExpediente`; vacía en el resto | `Expediente.TipoParaLista` | `classes/Expediente.cls:688-706` |
| Solo se devuelve un AM o un Lote | `ComandoSeleccionar_Click` | `:177-180` |
| Doble clic equivale a «Seleccionar» solo si el botón está habilitado | `ListaFiltrados_DblClick` | `:231-233` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Búsqueda | `TbExpedientes` (`SELECT *`) | SQL dinámico concatenado, DAO | no | `FUNCIONES UTILES.bas:1226-1232` |
| Carga de la fila elegida | `TbExpedientes` por `IDExpediente` | SQL dinámico concatenado, DAO | no | `:150, 259-261`; `constructor.bas:1151-1165` |

Sin escrituras: el cambio de tipo lo persiste `Form_FormExpedienteCambioTipo.ComandoRegistrar_Click` (`Form_FormExpedienteCambioTipo.cls:68`).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: comodín `*` de Jet en `Like` (en PostgreSQL, `%` o `ILIKE`) y palabra clave concatenada sin escapar (`FUNCIONES UTILES.bas:1228-1230`); `IDExpediente` concatenado sin comillas (`constructor.bas:1154`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Filtrar` marca `Selected(2)` (`:66`): con encabezados, el índice 2 sería la segunda fila de datos; con un único resultado podría fallar y mostrar un error en vez de la lista.
- `ComandoIniciarPalabraClave` («Ver Todos») vacía la palabra clave y `Filtrar` sale sin buscar (`:110-111`, `:43-45`): deja la lista vacía en lugar de mostrar todos.
- Ambos botones de búsqueda rotulan «Ver Todos» y `ComandoAyuda` rotula «Salir» en el diseño (`.form.txt:530, 654, 730`); no se ha comprobado si una imagen los sustituye.
- El tooltip de `PalabraClave` anuncia «CodExpCorto y Nemotécnico» (`.form.txt:627`), pero se busca en `CodExp`, `Titulo` y `Nemotecnico` (`FUNCIONES UTILES.bas:1228-1230`).
- Un apóstrofo en la palabra clave rompe el SQL concatenado (`FUNCIONES UTILES.bas:1228-1230`).
- Un `;` en `Titulo` o `CodExp` desplazaría columnas en la lista de valores (`:58-59`).
- La lista ofrece todos los tipos aunque solo se aceptan AM y Lote (`:177-180`); la columna «Tipo» queda vacía para lotes y basados (`classes/Expediente.cls:699-703`).
- Se acepta un Lote como padre, pero «Convertir a Lote» lo rechaza después (`Form_FormExpedienteCambioTipo.cls:102-105`); solo «Convertir a Basado» lo admite (`:125-128` del mismo fichero).
- El manejador de `ComandoSeleccionar_Click` llama a `CorreoAlAdministrador` sin `On Error Resume Next`, a diferencia del resto (`:188-191`).
- Sin `ORDER BY` en la búsqueda (`FUNCIONES UTILES.bas:1226-1230`): el orden de la lista no está garantizado.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtrar en servidor solo AM y Lote (o según la conversión elegida); búsqueda parametrizada con paginación; mostrar el tipo calculado completo.

## Criterios de aceptación

Propuestos, no aprobados:

1. Buscar sin palabra clave no consulta y deja la lista vacía (fila `PalabraClave` y regla «Sin palabra clave»).
2. Buscar devuelve los expedientes cuyo código, título o nemotécnico contienen el texto, también con apóstrofos.
3. Seleccionar un expediente que no es AM ni Lote muestra «Se ha de seleccionar un AM o un Lote» y no devuelve nada.
4. Seleccionar un AM o un Lote rellena el padre en el cambio de tipo y cierra el selector.
5. «Seleccionar» y el doble clic están deshabilitados hasta que hay una fila válida elegida.
6. Una búsqueda con un único resultado muestra ese resultado sin error.

## Pendientes

- Comportamiento real de `Selected(2)` con encabezados y uno o dos resultados (`:66`).
- Si los botones llevan imagen que oculte los rótulos del diseño (`.form.txt:530, 654, 730`).
- Valor de `m_TituloFormulario` al abrir desde `Form_FormExpedienteCambioTipo` (`:208`).
- Tratamiento de `;` y `,` en `AddItem` de una lista de valores (`:58-59`).
