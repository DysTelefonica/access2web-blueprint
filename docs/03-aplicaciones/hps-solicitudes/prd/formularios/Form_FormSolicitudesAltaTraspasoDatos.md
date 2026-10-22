# Form_FormSolicitudesAltaTraspasoDatos

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormSolicitudesAltaTraspasoDatos.cls` (953 líneas) y `.form.txt` (66 controles `Begin`) · **PRD:** [README](../README.md)

Subformulario de datos de la solicitud de alta y de traspaso: reúne los campos del solicitante, del expediente asociado, del motivo, del responsable y de la justificación, y es el único punto donde se validan y persisten esos campos. Actúa como coordinador de cuatro diálogos (búsqueda de expedientes, detalle de expediente, cambio de tipo y alta de responsable) y decide, tras guardar, si hay que reprogramar o anular correos automáticos. Es un formulario **no enlazado**: pinta los controles desde una instancia de `solicitud` y escribe a través de `SolicitudOperaciones`. Las citas son `Form_FormSolicitudesAltaTraspasoDatos.cls:línea` salvo indicación; `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`, `classes/` = `src/classes/`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormSolicitudAltaTraspaso` | `FrmDetalle.SourceObject` apunta a este formulario; no se le pasan argumentos | `Form_FormSolicitudAltaTraspaso.cls:118` |
| `FormSolicitudRenovacionDatos` | Ruta homóloga; comprueba el nombre del padre al guardar | `:818-822` |
| `FormSolicitudesGestion` | Al releer datos, invoca el filtrado de esa pantalla | `:468-470` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `FormExpedientesBusqueda` | `DoCmd.OpenForm` modeless, con `WithEvents m_FormExpedientes` | `:36-41`, `:135-140`, `:10` |
| `FormExpedienteDetalle` | `DoCmd.OpenForm` modeless tras cargar el expediente activo | `:157-192` |
| `FormCambioTipo` | `DoCmd.OpenForm` modeless, con `WithEvents m_FormCambioTipo` | `:231-243`, `:9` |
| `FormResponsable` | `DoCmd.OpenForm` modeless, con `WithEvents m_FormResponsable` | `:292-301`, `:8` |
| `frmBuscarJustificacion` | `DoCmd.OpenForm` con el argumento de modo selección, con `WithEvents f_BusquedaJustif` | `:930-938`, `:11` |

Estado compartido: `m_ObjSolicitudActiva` y `m_ObjSolicitudAlInicio` (base de datos y fotografía inicial), `m_ObjEntorno` (catálogos en caché), `m_ObjExpedienteActivo`, `m_ObjResposableActivo`, `m_ObjUsuarioConectado` y `pregunta` (`VG.bas:182-204`). Variables propias: `m_Justificacion`, `m_JServicio` y `m_CorreosAutoAlInicio`, que memoriza el valor de «correos automáticos» con el que se abrió el formulario para detectar la transición de «No» a «Sí» (`:6-12`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` → `Form_Load` | Llama a `EstablecerDatos` y propaga el error | `:323-351`, `.form.txt:221` |
| `Gestor` | combo (tramitadores) | — | Tramitador de la solicitud; se carga con `RellenarComboTramitadores` | `:379`, `FU.bas:370-398` |
| `Nombre`, `Apellido1`, `Apellido2` | texto | — | Nombre del solicitante, capitalizado al pintar | `:404-416`, `:583-585`, `FU.bas:1032-1046` |
| `FNacimiento`, `LugarNacimiento` | texto | — | Datos de nacimiento | `:418-424` |
| `email`, `emailResponsable`, `Telefono` | texto / combo | — | `emailResponsable` se rellena con el catálogo de responsables (correo; nombre; descripción) | `:426-436`, `:753-768` |
| `IDEmpresaUsuario`, `IDEmpresaTramitadora` | combo | — | Suministrador del solicitante y tramitador; cada fila es «identificador;nombre», con el nemotécnico por delante | `:1318-1341`, `:1342-1358` |
| `DNI` | texto | — | Se valida con el algoritmo de número y letra | `:465`, `constructor.bas:2985` |
| `TIPO` | texto | — | Tipo de la solicitud; solo se habilita el cambio cuando no es «Traspaso» | `:394-402`, `:580` |
| `CorreosAutomaticos` | combo Sí/No | — | Dispara la lógica de reprogramación o de limpieza de cola al guardar | `:633-685`, `.form.txt:829` |
| `Motivo_HPS` | combo | — | Motivo; el botón de al lado permite crear uno nuevo | `:260-289`, `:771-776` |
| `Observaciones` | texto multilínea | — | Notas libres | `:459-461` |
| `IDSolicitud` | texto | — | Identificador, solo lectura en pantalla | `:462-464` |
| `PalabraClave` | texto | `OnKeyDown` → `PalabraClave_KeyDown` | Enter o tabulador lanzan la búsqueda de expediente | `:195-212`, `:2122` |
| `ComandoBuscarExpediente` | botón | `OnClick` → `ComandoBuscarExpediente_Click` | Resuelve la palabra clave a un expediente | `:57-155`, `:2209` |
| `ComandoIrABusquedaExpedientes` | botón | `OnClick` → `ComandoIrABusquedaExpedientes_Click` | Abre el buscador con la palabra clave ya escrita | `:20-55`, `:2287` |
| `Titulo`, `IDExpediente` | texto | — | Expediente resuelto; el título cae al texto del expediente si el propio está vacío | `:104-110, 440-450` |
| `ComandoVerDetalleExpediente` | botón | `OnClick` → `ComandoVerDetalleExpediente_Click` | Abre el detalle del expediente activo | `:157-192`, `:2477` |
| `ComandoActualizar` | botón (rótulo heredado «Comando46») | `OnClick` → `ComandoActualizar_Click` | Recalcula y repinta el estado calculado de la solicitud | `:213-229`, `.form.txt:1974` |
| `ComandoActualizarar` | botón (rótulo «Contador») | `OnClick` → `ComandoActualizarar_Click` | Invalida la caché de expedientes del entorno | `:927-929`, `:2347-2349` |
| `ComandoCambioTipo` | botón | `OnClick` → `ComandoCambioTipo_Click` | Abre el diálogo de cambio de tipo | `:231-243`, `:2031` |
| `ComandoMotivos` | botón | `OnClick` → `ComandoMotivos_Click` | Pide un motivo nuevo por diálogo y lo da de alta | `:260-289`, `:1573` |
| `ComandoNuevoResponsable` | botón | `OnClick` → `ComandoNuevoResponsable_Click` | Abre el alta de responsable | `:292-319`, `:1852` |
| `btnBuscarJustif`, `txtJustificacion`, `txtIdJustificacion` | botón + texto | `OnClick` → `btnBuscarJustif_Click` | Abre el buscador de justificación y recibe el identificador elegido | `:930-953`, `:2540-2693` |
| `cmdLimpiarJustificacion` | botón | `OnClick` → `cmdLimpiarJustificacion_Click` | Vacía la justificación seleccionada | `:15-18`, `:2693` |
| `lblEstadoSolicitud` | etiqueta | — | Título del estado calculado | `:466-467`, `:1701` |
| `ComandoRegistrar` | botón | `OnClick` → `ComandoRegistrar_Click` | Valida, guarda y resuelve la cola de correos | `:617-707`, `:1391` |

## Flujos

### Normal: guardar (`ComandoRegistrar_Click`, `:617-707`)

1. Bloquea la edición del formulario y llama a `Registrar EnumSiNo.Sí, m_Error` (`:624-627`), que: detecta cambios (`HaHabidoCambios`), aborta con «No hay cambios que guardar» si no hay ninguno (`:796-806`), vuelca los controles al objeto (`RellenarObjeto`, `:581-604`), delega en `SolicitudOperaciones.RegistrarDatos` (`SolicitudOperaciones.cls:42-277`) y, por último, refresca el encabezado del padre (`:818-822`).
2. Si la configuración tiene los correos automáticos activos (`:633`), evalúa la transición del campo «correos automáticos» (`:636, 672`):
   - **De «No» a «Sí»**: solo pregunta si el proceso ya ha empezado, es decir, si existe alguna de las dos fechas disparadoras del trámite (`:637-651`); si el usuario lo acepta, reprograma los avisos pendientes (`:653-665`).
   - **De «Sí» a «No»**: anula sin preguntar todas las programaciones de envío que sigan sin salir (`:672-686`; `SolicitudOperaciones.cls:2633-2649`).
3. Devuelve la edición a su estado anterior, cierra el lanzador si está abierto, avisa de que los datos quedaron grabados y memoriza el valor actual para la próxima vez (`:689-697`).

### Alternativos

- **Alta en curso** (sin identificador): `EstablecerDatos` no carga fotografía inicial, fija el gestor al usuario conectado, oculta el estado y el botón de refresco, deshabilita el cambio de tipo y salta (`:378-386`). Toda alta pasa por esa rama.
- **Búsqueda de expediente por texto** (`:57-155`): con palabra clave vacía no hace nada; pide el catálogo de expedientes al constructor (`:79`); si no hay coincidencias, limpia los campos y deriva al buscador (`:83-95`); si la clave es numérica, la trata como identificador directo (`:99-113`); si hay exactamente un resultado, lo asigna sin preguntar (`:115-128`); en cualquier otro caso abre el buscador.
- **Selección desde el buscador** (`m_FormExpedientes_Seleccionado`, `:890-926`): vuelca identificador y título, habilita o deshabilita los botones de refresco y detalle según haya expediente, y limpia la palabra clave (`:898-911`).
- **Justificación** (`:930-953`): abre el buscador en modo selección y, al recibir el evento, rellena el identificador y la etiqueta; si la búsqueda no devuelve un identificador numérico, no toca nada.
- **Motivo nuevo** (`:260-289`): diálogo de entrada, alta en la tabla de motivos, invalidación de la caché del entorno, recarga del combo y selección del valor tecleado.
- **Cambio de tipo** (`:834-855`): al recibir el evento, cierra el formulario padre; quien repinta la sesión es el propio diálogo (`Form_FormCambioTipo.cls:63`).

### Errores

Dos patrones distintos conviven: los handlers de botón capturan el error, notifican al administrador y muestran un `MsgBox` crítico o de advertencia según el número (`errores:` de `:44-54, 146-154, 184-192, 916-926`), mientras que las funciones públicas devuelven el fallo por `p_Error` y lanzan internamente el código 1000 (`:360-478, 480-572, 574-615, 709-786, 788-832`). `HaHabidoCambios` y `RellenarObjeto` no devuelven valor de retorno: quien las llama debe mirar el parámetro de error (`:480, 574`).

## Reglas, validaciones y cálculos

- **Detección de cambios:** diecisiete comparaciones manuales entre la fotografía inicial y los controles, con salida anticipada en la primera diferencia (`:480-566`). Devuelve `True` sin comparar cuando no hay fotografía inicial (`:485-488`).
- **Volcado al modelo:** cada control se copia al objeto con `Nz`, de modo que un campo vacío llega como cadena vacía y no como nulo (`:581-604`). No hay validación de formato en esta capa.
- **Validación de negocio:** la delega `SolicitudOperaciones.MotivoNoOK`, que se ejecuta al inicio del guardado (`SolicitudOperaciones.cls:54-60, 343-446`): exige gestor conocido, nombre, primer apellido, correo con formato válido, correo responsable con formato válido si viene informado, expediente resoluble y con contratista principal cuando se informa, elección explícita de correos automáticos, empresa tramitadora registrada si se informa, y DNI válido si se informa. Tras validar, el propio registro recalcula el estado y lo persiste (`SolicitudOperaciones.cls:64-65`).
- **Estado calculado:** no se almacena el estado en este formulario; se deriva de las fechas del trámite y se traduce con el catálogo de títulos del entorno (`Solicitud.cls:1039-1075, 1122-1141`; `Entorno.cls:1781-1790`).
- **Programación de correos:** la transición a «Sí» solo ofrece reprogramar si el trámite ya tiene alguna de sus dos fechas disparadoras; la transición a «No» borra siempre (`SolicitudOperaciones.cls:2633-2649`).
- **Capitalización:** nombres y apellidos se pintan con mayúscula inicial y resto en minúsculas (`FU.bas:1032-1046`), lo que altera el texto tal como se introdujo.
- **Catálogos:** los combos se rellenan desde diccionarios del entorno con Invalidación manual de caché (`Entorno.cls:268-295, 297-320, 543, 578`); el de responsables concatena correo, nombre y descripción en un solo elemento (`:753-768`).

## Lecturas y escrituras

| Operación | Tabla o recurso | Tipo | Evidencia |
|---|---|---|---|
| Relectura de la solicitud | `TbSolicitudes` | DAO por identificador | `:373`; `constructor.bas:483-495` |
| Persistencia de la solicitud | `TbSolicitudes` | inserción o actualización campo a campo | `SolicitudOperaciones.cls:78-256` |
| Búsqueda de expedientes | `TbExpedientes` (base de expedientes) | `SELECT *` de la tabla entera y filtrado en memoria | `:79`; `constructor.bas:3855-3900` |
| Detalle de expediente | expedientes y su relación de contratistas | lectura por identificador | `:174`; `constructor.bas:2294` |
| Catálogos de tramitadores, suministros y responsables | tablas de suministradores, usuarios y responsables | lectura en caché del entorno | `:379-768`; `Entorno.cls:268-320, 1601, 578` |
| Motivos | tabla de motivos HPS | lectura de combo e inserción de un motivo nuevo | `:271-275`; `FU.bas:1442-1475` |
| Justificación | catálogo de justificaciones | lectura por identificador | `:352-358`; `classes/JustificacionServicio.cls:26` |
| Limpieza de cola de correo | tabla de correos enviados | `DELETE` de lo programado y no enviado | `SolicitudOperaciones.cls:2641-2644` |
| Estado calculado | fechas del trámite | lectura derivada | `:218, 467`; `Solicitud.cls:1039-1075` |

## Coexistencia

Pendiente de `../coexistencia.md`. Puntos previsibles: el buscador de expedientes carga la tabla completa y filtra con comparación de subcadena en memoria (`constructor.bas:3872-3896`), lo que en PostgreSQL exigiría una consulta con `ILIKE` y un límite, y es hoy el mayor coste de esta pantalla; los combos se construyen concatenando identificador y nombre en un mismo elemento (`:731-768`) en lugar de usar dos columnas con valor y texto; la validación depende de leer catálogos completos en caché; el estado es un entero traducido por un diccionario del entorno (`Entorno.cls:1830-1866`) y no una clave foránea. La lógica está **duplicada** en el `.cls` y en el bloque `CodeBehindForm` del `.form.txt` (`.form.txt:2870-3830`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **Llamadas a la instancia por defecto del formulario:** en cinco puntos se invoca un método cualificando el formulario por su nombre de clase en lugar de la instancia abierta: el filtrado del buscador de expedientes (`:36, 139`), el filtrado de la gestión de solicitudes (`:469`) y el refresco del encabezado del padre (`:819`). Con `DoCmd.OpenForm` la instancia abierta no es la instancia por defecto, de modo que esos «refrescos» pueden no afectar a la ventana visible. Verificar en ejecución.
- **`PalabraClave_KeyDown` no anula el tabulador:** atiende el tabulador como si fuera Enter y pone el foco en el botón (`:201-210`), lo que provoca un salto extra de foco al tabular; además solo actúa si el botón está habilitado (`:197-199`).
- **Búsqueda de expediente con coste creciente:** cada búsqueda carga la tabla de expedientes completa y filtra en memoria (`constructor.bas:3872-3896`); con palabra clave numérica se hace además una segunda lectura por identificador (`:99-113`).
- **Título del expediente sin cualificar el receptor:** al recibir la selección, el título se asigna con el nombre del control sin el prefijo del formulario (`:899-903`), que en VBA resuelve contra una variable local del módulo en lugar de escribir en el control; el resto de puntos del formulario sí lo cualifica (`:106-110, 444-450`). En la misma rama, la comprobación de si hay expediente o no (`:904-910`) se hace justo después de haber asignado el identificador recibido, de modo que su rama vacía es inalcanzable.
- **Nombre del botón «Contador»:** `ComandoActualizarar_Click` no cuenta nada; solo invalida la caché de expedientes del entorno (`:927-929`), y su rótulo no describe lo que hace (`.form.txt:2347-2349`). Convive con `ComandoActualizar`, que sí recalcula el estado (`:213-229`) y cuyo rótulo sigue siendo «Comando46» (`.form.txt:1974`).
- **Integración de los diálogos por `WithEvents`:** los cuatro diálogos se abren sin comprobar que no estén ya abiertos, salvo el buscador de expedientes (`:36-38, 134-136`), y se capturan con `Forms(...)` (`:41, 242, 300, 938`). Si el formulario ya estuviera abierto, la referencia quedaría apuntando a otra instancia.
- **Diálogos que nunca se cierran:** tras usar el buscador de expedientes, el detalle o la justificación, este formulario no los cierra; la siguiente pulsación los cierra sin guardar (`:29-30, 178-180`) o abre una segunda instancia (`:932-937`).
- **`m_FormResponsable_Alta` fuerza el alta** (`:861-889`): ante cualquier alta de responsable, incluso de la de otro usuario en otra ventana, invalida los catálogos y sobrescribe el correo responsable.
- **Cambio de tipo que cierra el padre** (`:840`): al recibir el evento se cierra el contenedor, y la ejecución continúa después sobre un formulario cuyo padre ha desaparecido; además el cierre del diálogo de tipo lo hace el propio diálogo (`Form_FormCambioTipo.cls:74`).
- **Comprobación de cambios frágil:** `HaHabidoCambios` compara diecisiete campos a mano (`:490-562`); un campo nuevo que no se añada a la lista se guarda sin que nada lo detecte. Además compara el identificador de expediente (`:543`) pero no el título, que es informativo.
- **`Err.Numbr` en el manejador de refresco** (`:224`), que impide compilar el módulo tal como está exportado; hay seis apariciones en el proyecto.
- **Titulares sin filtro de rol ni de expediente:** este subformulario guarda los datos de cualquier solicitante y permite cambiar el tipo, abrir cualquier expediente y dar de alta motivos, sin comprobar permisos. La autorización se apoya exclusivamente en la navegación heredada. Sensible.
- **`InformeDelRegistro` y archivos locales:** la trazabilidad se genera como HTML en disco y se abre en un visor propio (`FU.bas:1721-1745`), modelo que no sobrevive a un despliegue web.
- **Sin concurrencia optimista:** la actualización escribe todos los campos a partir de la fotografía inicial (`:818-819`; `SolicitudOperaciones.cls:151-256`); dos ventanas sobre la misma solicitud se pisan sin aviso.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: formulario de datos como vista única con subformularios sustituidos por búsquedas asíncronas; conjunto de campos tipado en lugar de diecisiete comparaciones manuales; validación declarativa reutilizada entre alta, traspaso y renovación, dado que el formulario tiene un gemelo casi idéntico (`Form_FormSolicitudRenovacionDatos.cls`); búsqueda de expedientes en servidor con paginación y límite; lista de motivos y de responsables como referencias cargadas bajo demanda; idempotencia y control de versión en la actualización; traducción de errores centralizada en un servicio único en lugar de cadenas dispersas por cada manejador.

## Criterios de aceptación

Propuestos, no aprobados:

1. En un alta nueva el gestor es el usuario conectado, no se muestra el estado y el cambio de tipo está deshabilitado.
2. Guardar sin cambios avisa y no escribe.
3. Guardar con datos incompletos avisa el primer motivo y no escribe.
4. Guardar persiste los campos y refresca el encabezado del padre y la lista de la gestión.
5. Buscar por identificador de expediente rellena identificador y título; sin coincidencias se abre el buscador.
6. Seleccionar un expediente en el buscador habilita el botón de detalle.
7. Activar los correos automáticos con el trámite ya iniciado ofrece reprogramar; desactivarlos elimina lo pendiente.
8. Buscar justificación deja el identificador y su etiqueta; el botón de limpiar los vacía.

## Pendientes

- Si la instancia por defecto del formulario es o no la instancia abierta en los puntos citados (`:36, 139, 469, 819`): comportamiento decisivo para los refrescos.
- Si `Form_Load` llega a ejecutarse cuando el formulario actúa como subformulario, dado que `OnLoad` está declarado en su propio `.form.txt` (`.form.txt:221`) y el flujo de alta depende de ello (`:323-335`).
- Contenido de `constructor.getSolicitud` y su comportamiento con la caché del entorno (`constructor.bas:483-495`).
- Tamaño real de la tabla de expedientes y su efecto en el filtrado en memoria (`constructor.bas:3872-3896`).
- Diferencias reales entre este formulario y su gemelo `Form_FormSolicitudRenovacionDatos`, que repite la misma arquitectura (`Form_FormSolicitudRenovacionDatos.cls:6, 851-858, 914-941`).
- Sincronía entre el `.cls` y la copia del bloque `CodeBehindForm` del `.form.txt` (`.form.txt:2870`).