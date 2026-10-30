# Form_FormParteGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormParteGestion.cls` (657 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Gestión de los partes de trabajo (informes de avería) de una factura: el usuario elige una `IDFactura`, genera los partes por centro (BUI) y sus Word, abre el Word para firma, anexa el parte firmado, elimina los partes de la factura y consulta los eventos de cada parte. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); las dos listas son de tipo «Value List» y se rellenan por código (`Form_FormParteGestion.form.txt:361, 458, 493`). Estado en variables de módulo: factura, parte y evento elegidos (`Form_FormParteGestion.cls:6-10`). Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`; `VG.bas` = `modules/Variables Globales.bas`. Las citas son `Form_FormParteGestion.cls:línea` salvo indicación; las de la forma `:N` sin fichero son de este `.cls`.

**Controles: 19 (29 `Begin` en el `.form.txt`: 1 `Form`, 8 de plantilla por defecto, 1 `Section` y 19 controles).**

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form0BDInformes` | Rótulo `lblGestionDePartes`: abre el formulario y cierra el propio | `Form_Form0BDInformes.cls:44-48` |
| `FormEventoGestion` | `ComandoGestionPartes_Click`: abre el formulario | `Form_FormEventoGestion.cls:2027-2030` |
| `FormInformesPorFactura` | `ComandoGenerarInformesPorCentroParaFirma_Click`: abre el formulario | `Form_FormInformesPorFactura.cls:234-237` |
| `FormFacturacionGestion` | `ComandoAnexarFirmados_Click`: si no está abierto, lo abre, fija `IDFactura` y llama a `IDFactura_Exit False` | `Form_FormFacturacionGestion.cls:77, 97-101` |

Salida: abre `FormEventoGestion` desde `ComandoVerDetalleEventoEnParte` (`:232-270`). Declara el evento público `Adjuntado` (`:11`), que **no se lanza** en este fichero. `cmdSalir` cierra sin guardar (`:313-316`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `IDFactura` | combo, lista de valores | `OnChange` → `IDFactura_Change` | Carga la factura elegida y la lista de partes; sin valor, borra la factura y deshabilita Generar | CAP-8 Facturación, gastos y subcontratación | `:318-372`; `.form.txt:457-461` |
| `ListaFiltrados` | lista, 4 columnas (`0;2835;1418;1418`) | `OnClick` → `ListaFiltrados_Click` | Carga el parte elegido, habilita Adjuntar/Eliminar/Word y rellena los eventos | CAP-8 Facturación, gastos y subcontratación | `:496-557`; `.form.txt:352-363` |
| `ListaEventosEnIDParte` | lista, 2 columnas (`1701;6436`) | `OnClick` → `ListaEventosEnIDParte_Click` | Carga el evento elegido y habilita Detalle | CAP-3 Eventos | `:451-495`; `.form.txt:483-495` |
| `ComandoGenerarPartesDeTrabajo` | botón (deshabilitado de inicio) | `OnClick` | Pide carpeta y genera los partes de la factura y sus Word | CAP-8 Facturación, gastos y subcontratación | `:151-199`; `.form.txt:755, 766` |
| `ComandoGenerarWordParaFirma` | botón (deshabilitado de inicio), rótulo «Wor Para Firma» | `OnClick` | Genera el Word del parte elegido y lo abre | CAP-8 Facturación, gastos y subcontratación | `:201-230`; `.form.txt:1004, 1014` |
| `ComandoAdjuntarParteFirmado` | botón (deshabilitado de inicio) | `OnClick` | Anexa el documento firmado al parte elegido | CAP-11 Anexos | `:12-75`; `.form.txt:795, 805` |
| `ComandoEliminarParte` | botón (deshabilitado de inicio) | `OnClick` | Elimina **todos** los partes de la factura del parte elegido | CAP-8 Facturación, gastos y subcontratación | `:79-117`; `.form.txt:907, 918` |
| `ComandoVerDetalleEventoEnParte` | botón «Detalle» (deshabilitado de inicio) | `OnClick` | Abre `FormEventoGestion` posicionado en el evento | CAP-3 Eventos | `:232-270`; `.form.txt:835, 845` |
| `ComandoExportarExcelEventos` | botón (deshabilitado de inicio), rótulo «Comando227» | `OnClick` | Exporta a Excel `EventosEnInformeAverias` con la lista de eventos | CAP-10 Informes | `:119-148`; `.form.txt:673, 682` |
| `ImagenParteAnexado` | imagen | `OnClick` | Copia el anexo a la carpeta local y lo abre | CAP-11 Anexos | `:431-450`; `.form.txt:524` |
| `cmdSalir` | botón | `OnClick` | Cierra el formulario | sin asignar | `:313-316`; `.form.txt:390-392` |
| `lblNumeroPartes`, `lblNumeroEventos` | etiquetas | — | «Elementos en la lista (n)» | CAP-8 Facturación, gastos y subcontratación | `:287-288, 599, 645` |
| `lblEstado` | etiqueta (oculta) | — | Mensaje de avance; se oculta al abrir y al terminar | sin asignar | `:303, 187, 222`; `.form.txt:875, 888` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título; `lblTitulo1` se copia a `Caption` | sin asignar | `:284`; `.form.txt:722, 743` |
| `Etiqueta183`, `CuadroSuperior`, `Cuadro2` | etiqueta y 2 rectángulos | — | Decoración | sin asignar | `.form.txt:308, 331, 424` |

## Flujos

### Apertura (`Form_Open`, `:273-312`)

1. `AjustarTamaño Me` (`:282`); título tomado de `lblTitulo1` (`:284`); listas con solo la cabecera (`:285-286`).
2. `getValoresDistintos` rellena `IDFactura` con los `IDFactura` distintos y no nulos de `TbFacturaPrincipal` (`:294`; `FG.bas:6753-6775`).
3. Error: `MsgBox` y cierre del formulario (`:305-311`).

### Selección de factura (`IDFactura_Change` `:318`, `IDFactura_Exit` `:374`)

Ambos handlers tienen cuerpo idéntico (`:318-372` y `:374-429`). Vacían las listas y ocultan el anexo (`:327-331`). Sin factura: deshabilitan Generar (`:333-338`). Con factura: `Constructor.getFactura` si cambia (`:339-351`); si `ColPartes` es `Nothing` habilitan Generar (`:352-354`); si no, `EstablecerListaPartes` (`:355`).

### Listas (`:559-655`)

`EstablecerListaPartes` deshabilita los cuatro botones, añade una fila por parte (`IDParte;Centro;FechaInicial;FechaFinal`) y, si solo hay un parte, lo selecciona y llama a `ListaFiltrados_Click` (`:576-604`). `EstablecerListaEventos` hace lo mismo con los eventos del parte (`:626-649`).

### Generar partes (`:151-199`)

Exige factura (relee `IDFactura.Column(0)` si hace falta, `:162-173`); pide carpeta con `Seleccionar(False, …)` (`:174`); sale en silencio si no existe (`:175-180`). `Factura.GenerarPartesDeFactura` crea un parte por BUI y los Word en la carpeta (`Factura.cls:4013-4075`). Después anula `ColPartes` y recarga la lista (`:185-186`).

### Word, anexo, eliminar, detalle, Excel

- **Word** (`:201-230`): `Parte.GeneraWordParaFirma` abre la plantilla de averías configurada y rellena el Word (`Parte.cls:222-262`); lo abre con `Ejecutar … "open"` si existe (`:219-221`).
- **Adjuntar** (`:12-75`): pide confirmación si ya hay anexo (`:25-32`) y aborta si está abierto (`:33-36`); `Seleccionar(True, …)`; `Parte.AnexarParte` copia el fichero a la carpeta de partes firmados como `<IDParte>_PA.<ext>` y actualiza `NombreArchivoAdjunto` (`Parte.cls:167-215`).
- **Eliminar** (`:79-117`): confirmación con `vbDefaultButton2` (`:88-92`); `Factura.EliminarPartes` borra cada parte (`Factura.cls:5120-5150`).
- **Detalle** (`:232-270`): si `FormEventoGestion` está cerrado, lo abre y fija `ComboIDEventoBusqueda`; si está abierto con otro evento, llama antes a `ComandoActualizarTodo_Click` (`:248-258`).
- **Excel** (`:119-148`): `GenerarConsultas("EventosEnInformeAverias", Me.hWnd, , Me.ListaEventosEnIDParte, "No")` (`:131`).

### Errores

Patrón común: `Err.Raise 1000` con `m_Error` propio; cualquier otro error se compone con número y descripción; `Hourglass` se desactiva y se muestra `MsgBox` crítico (p. ej. `:69-75`). El formulario solo se cierra en `Form_Open` (`:311`).

## Reglas, validaciones y cálculos

- Adjuntar exige parte elegido: «Se ha de seleccionar un elemento de la lista superior» (`:21-24`). Eliminar: «Ha de seleccionar el parte» (`:84-87`) y factura conocida (`:96-99`). Word: «Seleccione un Parte de la lista » (`:208-211`). Detalle: «Seleccione un evento de la lista superior» (`:236-239`).
- Un parte tiene anexo si `URLParte <> ""` (`:536-540`); `URLParte` es la carpeta de partes firmados configurada más `NombreArchivoAdjunto` (`Parte.cls:46-50`).
- Los contadores muestran `ListCount - 1` (descuenta la cabecera) (`:599, 645`).
- Fechas del parte: mínima y máxima de `FechaFinal` de sus eventos, vía `getFechaInicialYFinalDeColeEventos` (`Parte.cls:327-336`; `FG.bas:6789-6832`).
- Identificador del parte: `IDParteCalculado` (`Parte.cls:343, 461`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Valores del combo | `TbFacturaPrincipal` | SQL dinámico (`SELECT DISTINCT` con tabla y campo concatenados) | — | `FG.bas:6764-6766` |
| Leer factura | `TbFacturaPrincipal` | SQL dinámico con `IDFactura` concatenado sin comillas | — | `Constructor.bas:2440-2441` |
| Leer parte | `TbPartesPpal` | SQL dinámico con `IDParte` entre comillas | — | `Constructor.bas:1781-1782` |
| Partes de la factura | `TbPartesPpal` | SQL dinámico concatenado | — | `Constructor.bas:919-921` |
| Regenerar partes | `TbPartesPpal`, `TbPartesDetalle` | SQL dinámico `DELETE` por factura y centro | no | `Parte.cls:319-326` |
| Alta de parte y detalle | `TbPartesPpal`, `TbPartesDetalle` | DAO `AddNew`/`Update` | no | `Parte.cls:344-354, 374-393` |
| Anexar parte | `TbPartesPpal` | SQL dinámico `UPDATE` de `NombreArchivoAdjunto` | no | `Parte.cls:208-212` |
| Eliminar parte | `TbPartesPpal`, `TbPartesDetalle`, `TbEventos` | SQL dinámico `DELETE` y `UPDATE … IDParte = Null` | no | `Parte.cls:431-449` |
| Exportar a Excel | consulta guardada `EventosEnInformeAverias` | indirecta, vía `GenerarConsultas` | — | `:131` |
| Ficheros | carpeta elegida, carpeta local y de partes firmados | `FileSystemObject` / ShellExecute | — | `Parte.cls:195-200`; `FG.bas:218-250` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL dinámico concatenado con identificadores sin parametrizar; `DELETE tabla.*` y `DELETE *` con sintaxis Jet/ACE (`Parte.cls:319, 431`); `Date` de VBA en `FechaObtencion` (`Parte.cls:346`); `IDFactura` numérico concatenado sin comillas y `IDParte` como texto; altas y borrados de varias tablas sin transacción; dependencia de Word, Excel, ShellExecute y rutas de red configuradas, que no existen en la web.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComandoEliminarParte` borra todos los partes de la factura aunque haya uno elegido (`:88-89, 100`; `Factura.cls:5120-5145`); el mensaje lo dice, pero el botón se habilita al elegir un parte (`:529`).
- `EliminarPartes` comprueba que el anexo no esté abierto parte a parte y borra a medida; si uno está abierto, los anteriores ya se han eliminado (`Factura.cls:5133-5146`).
- `Parte.Alta` borra los partes previos del mismo centro antes de insertar, sin transacción (`Parte.cls:319-326`); un fallo posterior deja la factura sin ese parte.
- `ImagenParteAnexado_Click` usa `m_ObjParteSeleccionado` sin comprobar `Nothing` (`:436`).
- `Adjuntar` no actualiza `m_ObjParteSeleccionado` de forma explícita: tras anexar, `ListaFiltrados_Click` reutiliza el objeto (`:63, 519`), y depende de que `AnexarParte` reinicie `URLParte` (`Parte.cls:204`).
- `ListaEventosEnIDParte` no se vacía en `ListaFiltrados_Click` si el parte no tiene eventos, y `m_ObjEventoSeleccionado` no se reinicia al cambiar de parte (`:541`; `:626`).
- `ComandoGenerarPartesDeTrabajo_Click` nombra el error `ComandoPartesParaCentros_Click` (`:194`) y `IDFactura_Change` nombra `IDFactura_AfterUpdate` (`:368`).
- `Form_Open` y `IDFactura_Change` fuerzan `Hourglass False` al entrar (`:324`), lo que oculta el estado de espera.
- `Err.Number` distinto de 1000 en `GeneraWordParaFirma` accede a `appWord` sin comprobar `Nothing` (`Parte.cls:280-287`).
- Rótulos de diseño: «Comando227» (`.form.txt:682`), «Wor Para Firma» (`.form.txt:1014`), título «Gestión de los Informes de Avería» (`.form.txt:39`) frente a «Gestión de Partes de Trabajo» (`.form.txt:722`).
- `Adjuntado` se declara y no se lanza (`:11`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: eliminar un solo parte o confirmar que se eliminan todos de forma explícita; operación de regenerado transaccional; carga y descarga de anexos por el navegador en lugar de Word/ShellExecute local.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, `IDFactura` ofrece los identificadores distintos de `TbFacturaPrincipal` (`:294`).
2. Elegir una factura sin partes habilita «Generar Partes de Trabajo»; con partes, los lista y deshabilita Generar (`:352-355, 576-592`).
3. Una factura con un único parte lo selecciona y muestra sus eventos (`:601-604`).
4. Elegir un parte habilita Adjuntar, Eliminar y Word, y muestra la imagen de anexo solo si hay fichero asociado (`:528-540`).
5. Adjuntar con un parte que ya tiene anexo pide confirmación y no continúa si está abierto (`:25-36`).
6. Eliminar pide confirmación y borra todos los partes de la factura (`:88-100`).
7. «Detalle» abre `FormEventoGestion` en el evento elegido (`:248-258`).
8. Cualquier fallo muestra un mensaje de error y deja el formulario abierto, salvo en la apertura (`:70-74, 305-311`).

## Pendientes

- Contenido de `AjustarTamaño` (`:282`) y de `Avance` usado por las clases durante la generación.
- Consulta guardada `EventosEnInformeAverias` (`:131`): no resuelta en `src`.
- Ubicación configurada de plantilla, carpetas local y de partes firmados (valores en `m_ObjEntorno`; `Parte.cls:46-50, 195`).
- Si `IDParteCalculado` y `Parte.Alta` colisionan en concurrencia (`Parte.cls:343, 461`).
