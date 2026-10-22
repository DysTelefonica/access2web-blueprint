# Form_FormExpedientesBusqueda

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormExpedientesBusqueda.cls` (352 líneas) y `.form.txt` · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `fichero:línea` del commit fijado.

Buscador de expedientes: filtra por palabra clave, muestra el resultado en una lista de seis columnas y devuelve el expediente elegido al contenedor, que lo carga en la solicitud en edición. Es un formulario **no enlazado** (la lista se puebla con `AddItem`, `.form.txt:1068-1070`) y **no escribe nada**. Es además la vía por la que se abre la ficha de detalle de expediente.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormAltaDatosSolicitante` | `DoCmd.OpenForm` y suscripción con `WithEvents` al evento `Seleccionado` | `Form_FormAltaDatosSolicitante.cls:7`, `:25-33`, `:107-114` |
| `Form_FormSolicitudRenovacionDatos` | mismo patrón | `Form_FormSolicitudRenovacionDatos.cls:7`, `:24-32`, `:127-134` |
| `Form_FormSolicitudesAltaTraspasoDatos` | mismo patrón | `Form_FormSolicitudesAltaTraspasoDatos.cls:10`, `:28-36`, `:132-139` |

Los tres contenedores usan el mismo ritual: cierran el buscador si ya estaba abierto, lo abren, se suscriben y llaman a `Filtrar` con la palabra clave ya escrita en el control.

**Contrato de salida:** declara un único evento público, `Seleccionado(m_Exp As Expediente)` (`:9`), que se lanza en `ComandoElegir_Click` (`:94`) inmediatamente antes de cerrarse (`:95`). El parámetro es el objeto `Expediente` completo, ya cargado desde su base de datos.

Abre otros formularios: `FormExpedienteDetalle`, previa cierre de cualquier instancia previa (`:149-153`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `PalabraClave` | texto | `OnExit` | Filtra al salir del campo, si la palabra clave cambió | `:319-350`, `.form.txt:1221-1232` |
| `ListaFiltrados` | lista 6 columnas con cabecera | `OnClick` | Selecciona un expediente y muestra su título y códigos | `:252-315`, `.form.txt:1053-1070` |
| `Titulo` | texto deshabilitado | — | Título del expediente seleccionado | `:265`, `:294`, `.form.txt:1429-1442` |
| `Otros` | texto deshabilitado | — | Código corto y código largo del expediente seleccionado | `:266`, `:295-296`, `.form.txt:1528-1541` |
| `ComandoActualizar` | botón | `OnClick` | Vacía el filtro y vuelve a filtrar | `:13-44`, `.form.txt:919-921` |
| `ComandoLimpiarPalabraClave` | botón | `OnClick` | Vacía el campo de palabra clave | `:112-119`, `.form.txt:1258-1260` |
| `ComandoElegir` | botón | `OnClick` | Lanza `Seleccionado` y cierra | `:74-109`, `.form.txt:1501-1503` |
| `ComandoVerDetalleExpediente` | botón | `OnClick` | Abre la ficha de detalle del expediente seleccionado | `:130-169`, `.form.txt:1571-1573` |
| `ComandoAyuda` | botón | `OnClick` | Abre el PDF de ayuda del formulario | `:46-72`, `.form.txt:986-988` |
| `ComandoSalir` | botón | `OnClick` | Cierra el formulario | `:122-125`, `.form.txt:1617-1619` |
| `lblTitulo1` | etiqueta | — | Único rótulo; `Form_Open` copia su texto al título de la ventana | `:182`, `.form.txt:900-901` |

## Flujos

### Apertura (`Form_Open`, `:171-199`)

1. Dimensiona con `Ajustar Me` (`:178`; `FU.bas:735-758`) y copia el rótulo al título (`:182`).
2. Fija las cabeceras de la lista (`:183`). **No filtra**: la lista queda con la fila de cabecera hasta que el contenedor llama a `Filtrar` o el usuario escribe una palabra clave.

### Filtrado (`Filtrar`, `:202-250`, `Public`)

1. Vacía los rótulos del expediente seleccionado (`:215-216`) y vuelve a fijar las cabeceras (`:218`).
2. Llama a `constructor.getExpedientesBusqueda` con la palabra clave y **el segundo argumento vacío** (`:221`): el parámetro `p_IDResponsableCalidad` nunca se rellena desde aquí.
3. Si la colección es `Nothing` —es decir, sin coincidencias— sale sin poblar la lista (`:225-227`).
4. Añade una fila por expediente con seis campos separados por punto y coma, sustituyendo el `;` del título por `:` para no romper la separación (`:233-238`).
5. Selecciona la primera fila (`:240`) y dispara `ListaFiltrados_Click` (`:241`).

### Selección (`ListaFiltrados_Click`, `:252-315`)

1. Deshabilita elegir y ver detalle, y vacía los rótulos (`:263-266`).
2. Si no hay fila seleccionada, anula `m_ExpSeleccionado` y sale (`:267-273`).
3. Si ya había expediente seleccionado y el identificador de la fila es otro, lo recarga (`:279-286`).
4. Rellena `Titulo` y `Otros` (`:293-297`) y rehabilita los dos botones (`:299-300`).

### Alternativos

- **Elegir** (`:74-109`): si no hay expediente seleccionado lo carga por identificador (`:82-90`), lanza `Seleccionado` (`:94`) y se cierra (`:95`).
- **Ver detalle** (`:130-169`): carga el expediente si hace falta (`:139-146`), lo publica en la global `m_ObjExpedienteActivo` (`:149`), cierra la ficha previa (`:150-152`), abre `FormExpedienteDetalle` (`:153`) y cierra la barredora (`:155`). No cierra este formulario.
- **Limpiar palabra clave** (`:112-119`): vacía el campo y vuelve a filtrar **solo si la última palabra clave aplicada no estaba vacía** (`:115-117`); si nunca se filtró, el botón no hace nada visible.
- **Palabra clave** (`:319-350`): si el texto no ha cambiado desde la última aplicación, no vuelve a filtrar (`:327-333`); si ha cambiado, filtra y recuerda el valor (`:335-336`).

### Errores

Todos los handlers repiten el patrón de la casa: `Err.Raise 1000` como error interno controlado, `DoCmd.Hourglass False`, `CorreoAlAdministrador` solo si `Err.Number <> 1000` y `MsgBox` crítico o de advertencia según el caso (`:34-43`, `:63-71`, `:100-108`, `:160-168`, `:189-197`, `:306-314`, `:341-349`).

## Reglas, validaciones y cálculos

- **Alcance del filtro:** el formulario pasa la palabra clave pero deja vacío `p_IDResponsableCalidad` (`:221`), así que la consulta es `SELECT * FROM TbExpedientes` sin restricción (`constructor.bas:3872-3873`).
- **Comparación en memoria, sensible a mayúsculas:** la palabra clave se contrasta con `InStr` en cinco propiedades —identificador, nemotecnico, título, código corto y código largo— usando el comparador predeterminado (binario), no `vbTextCompare` (`constructor.bas:3888-3894`). El resultado es que «exp» y «EXP» no devuelven lo mismo, a diferencia del filtrado de `Form_FormSolicitudesGestion` (`:317` de ese fichero).
- **Sin paginación ni límite:** se recorren todos los registros y se materializan todos en un `Scripting.Dictionary` (`constructor.bas:3896-3905`); el coste crece con el tamaño de `TbExpedientes`.
- **Segundo backend:** expedientes viven en una conexión distinta de solicitudes —`getdbExpedientes()` frente a `getdb()`— (`constructor.bas:3876`, `constructor.bas:2310`).
- **Selección implícita:** al terminar de filtrar se selecciona siempre la primera fila y se dispara el manejador de clic (`:240-241`), de modo que hay una selección aunque el usuario no haya pulsado nada.
- **Sin control de rol:** este formulario no comprueba permisos. La comprobación vive en los contenedores que lo abren.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Búsqueda de expedientes | `TbExpedientes` (tabla completa, en el segundo backend) | SQL estática + filtro en memoria | `constructor.bas:3867-3876`, `constructor.bas:3888-3905` |
| Expediente por identificador | `TbExpedientes` | SQL por concatenación | `constructor.bas:2306-2310` |
| Global de expediente activo | `m_ObjExpedienteActivo` | escritura en memoria | `VG.bas:202`; `:149` |
| Ayuda | PDF homónimo en el directorio de documentación | E/S de fichero | `:54`; `FU.bas:475-500` |

Sin escrituras en base de datos ni en el sistema de ficheros.

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: la búsqueda es «cargar todo y filtrar en el cliente» (`constructor.bas:3872-3873`, `:3888-3894`), patrón que en PostgreSQL se sustituye por `ILIKE '%…%'` con índices; el paso de datos al contenedor es un evento VBA con un objeto de sesión, no una respuesta de petición (`:94`); y la selección se comunica por una global mutada antes de abrir otra ventana (`:149`, `Form_FormExpedienteDetalle.cls:49`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-10.** Las cabeceras de la lista no coinciden con el contenido de las columnas: se anuncia `IDExp;Cod_Exp;Nemotécnico;Eslote;Nº;Título` (`:183`, `:218`) y se escriben `IDExpediente;CodExp;Nemotecnico;TipoInforme;Ordinal;Título` (`:233-238`). Es decir, la columna rotulada «Eslote» muestra el tipo de informe y la rotulada «Nº» muestra el ordinal. A verificar el origen de ese desajuste: `TipoInforme` no lo aporta ningún control de este formulario.
- **H-HPS-S-11.** El filtro se resuelve cargando `TbExpedientes` entera desde el segundo backend y descartando en memoria las no coincidencias, con comparación sensible a mayúsculas (`constructor.bas:3872-3873`, `:3888-3894`), sin paginación ni límite; la palabra clave nunca llega al motor de datos. En esa misma línea, `m_IDResponsableCalidadUltimo` se declara (`:8`) y se reinicia en `ComandoActualizar_Click` (`:22`) pero **no se lee en ningún punto**, y el segundo parámetro de `getExpedientesBusqueda` queda siempre vacío (`:221`): código muerto que sugiere un filtro por responsable de calidad retirado a medias.
- **H-HPS-S-12.** La entrada y el refresco del formulario son desiguales: `Form_Open` no llama a `Filtrar` (`:171-199`), de modo que abierto directamente muestra solo la fila de cabecera y aparenta no traer datos; `ComandoLimpiarPalabraClave` solo vuelve a filtrar si la última palabra clave aplicada era distinta de vacía (`:115-117`), de modo que limpiar tras una apertura sin filtro no produce ningún efecto observable; y `ComandoActualizar_Click` vacía `m_PalabraClaveUltima` (`:21`) pero **no** la selección `m_ExpSeleccionado` (`:5`, `:268`), por lo que tras actualizar puede sobrevivir la selección anterior si el identificador coincide con el de la primera fila (`:279-286`).

**Defectos de interfaz heredados (a verificar):** el título de diseño es `Usuario conectado: Administrador ( Administrador )(64 bits) EN PRUEBAS` (`.form.txt:40`), que describe la sesión y el entorno, no el formulario; `Form_Open` lo sobrescribe con el texto de la etiqueta (`:182`), así que el defecto solo es visible en diseño. `ComandoActualizar` y `ComandoLimpiarPalabraClave` conservan rótulos internos `Comando47` y `Comando36` (`.form.txt:920`, `:1259`). El `RowSource` de diseño de la lista contiene filas de datos reales (`.form.txt:1070`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: búsqueda en servidor con paginación y orden explícito; el expediente elegido como identificador en la ruta del contenedor en lugar de evento con objeto de sesión; la ficha de detalle como panel dentro del propio buscador; y criterio de ordenación declarado en lugar del que devuelva el recorrido de la tabla.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al escribir una palabra clave y salir del campo, la lista contiene los expedientes cuyo identificador, nemotécnico, título, código corto o código largo la contienen.
2. Sin coincidencias, la lista queda vacía y no se produce ningún aviso.
3. Con resultados, la primera fila queda seleccionada y se muestran el título y los códigos en los campos de lectura.
4. Elegir devuelve el expediente al contenedor y cierra el buscador.
5. Ver detalle abre la ficha del expediente seleccionado y el buscador permanece abierto.
6. Limpiar la palabra clave repone la lista completa.
7. Actualizar repone el estado inicial del filtro.

## Pendientes

- Si el contenedor que llama a `Filtrar` con una palabra clave ya escrita llega a ver resultados: `Form_Open` no filtra y el filtro solo se dispara al salir del campo (`:319-350`).
- Volumen real de `TbExpedientes` en el segundo backend, para dimensionar H-HPS-S-11.
- Si el rótulo «Eslote» de H-HPS-S-10 corresponde a un tipo de informe concreto en el dominio HPS: `Expediente.TipoInforme` se escribe en la lista pero el formulario no lo expone en ninguna parte visible.
- Si `Expediente.Clase` expone un identificador de expediente estable y no nullable (`.form.txt:1070` sugiere que el identificador viaja como texto en la lista).