# Form_FormSolicitudesAltaTraspasoDatos

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormSolicitudesAltaTraspasoDatos.cls` (953 líneas) y `.form.txt` (65 bloques `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`, `constructor.bas` = `modules/constructor.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_FormSolicitudesAltaTraspasoDatos.cls:línea` salvo indicación.

Formulario de datos del alta y del traspaso. Comparte casi todo el código con `Form_FormSolicitudRenovacionDatos` —misma detección de cambios, mismo volcado a la entidad, misma lógica de correos automáticos— y añade la **justificación** de la solicitud, que no existe en el hermano. Se monta en dos contenedores distintos (`FormSolicitudAltaTraspaso` y `FormSolicitudRenovacion`) y lo resuelve por `Me.Parent.Name` al guardar (`:818-822`).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudAltaTraspaso` | Lo monta como subformulario con el contexto de sesión ya establecido | `:818-819` |
| `Form_FormSolicitudRenovacion` | El mismo formulario sirve al contenedor de renovación; ambos lo distinguen por nombre al guardar | `:820-821` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `frmBuscarJustificacion` | `btnBuscarJustif_Click` la abre en modo selección y escucha `JustificacionElegida` | `:930-939, 942-949` |
| `Form_FormExpedientesBusqueda` | Apertura desde el buscador de expedientes, con `WithEvents` | `:20-56, 890` |
| `Form_FormResponsable`, `Form_FormCambioTipo` | Apertura con `WithEvents` | `:292, 231, 834, 861` |
| `Form_FormSolicitudesGestion` | `EstablecerDatos` refresca la rejilla principal como efecto secundario de cargar | `:468-470` |
| Contenedor padre | `EstablecerDatos` del padre tras guardar | `:818-822` |
| `SolicitudOperaciones` | `RegistrarDatos`, `ReprogramarAvisosFuturos`, `AnularTodosLosProgramados` | `:661, 678, 813` |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` → `Form_Load` | `EstablecerDatos`; las dos líneas de `AllowEdits` están comentadas | `:323-351` |
| `txtIdJustificacion`, `txtJustificacion` | texto | — | Identificador y rótulo de la justificación elegida | `:356-357, 599`, `.form.txt:2668, 2540` |
| `btnBuscarJustif` | botón | `OnClick` → `btnBuscarJustif_Click` | Abre el buscador de justificaciones en modo selección | `:930-939`, `.form.txt:2598` |
| `cmdLimpiarJustificacion` | botón | `OnClick` → `cmdLimpiarJustificacion_Click` | Vacía los dos controles de justificación | `:15-18`, `.form.txt:2693` |
| `ComandoRegistrar` | botón | `OnClick` → `ComandoRegistrar_Click` | Guarda y ejecuta la lógica de correos automáticos | `:617-707` |
| `ComandoActualizar` | botón | `OnClick` → `ComandoActualizar_Click` | Repinta el rótulo de estado; oculto en alta | `:213-228, 383, 465` |
| `ComandoActualizarar` | botón | `OnClick` → `ComandoActualizarar_Click` | Invalida la caché de expedientes | `:927-929` |
| `ComandoBuscarExpediente`, `ComandoIrABusquedaExpedientes`, `ComandoVerDetalleExpediente`, `PalabraClave` | botón, botón, botón, texto | varios | Búsqueda y detalle de expediente | `:20, 57, 157, 195` |
| `ComandoCambioTipo`, `ComandoMotivos`, `ComandoNuevoResponsable` | botones | varios | Cambio de tipo, motivos y alta de responsable | `:231, 260, 292` |
| `TIPO`, `Gestor`, `Nombre`, `Apellido1`, `Apellido2`, `FNacimiento`, `LugarNacimiento`, `email`, `emailResponsable`, `Telefono`, `DNI`, `Observaciones`, `Motivo_HPS`, `IDEmpresaUsuario`, `IDEmpresaTramitadora`, `IDExpediente`, `IDSolicitud`, `CorreosAutomaticos` | controles enlazados por nombre | — | Los campos que la entidad expone | `:490-570, 582-600` |
| `lblEstadoSolicitud` | etiqueta | — | Estado calculado de la solicitud; sólo visible en edición | `:382, 466-467` |

## Flujos

### Normal — carga

1. `Form_Load` llama a `EstablecerDatos` y propaga el error (`:333-336`).
2. `EstablecerDatos` (`:360-479`) empieza por los combos (`:366`), carga la fotografía inicial si la solicitud ya existe (`:371-376`) y **en alta** se detiene pronto: fija gestor y tipo, da el foco a `CorreosAutomaticos`, oculta el rótulo de estado y el botón de actualizar, y deshabilita el cambio de tipo (`:378-386`). En esta rama los datos del solicitante **no** se copian, a diferencia del hermano de renovación.
3. **En edición** repinta los campos desde `m_ObjSolicitudAlInicio` (`:390-464`), recupera la justificación si la solicitud la tiene (`:391-393`), muestra el estado calculado (`:465-467`), refresca la rejilla de `FormSolicitudesGestion` (`:468-470`) y recuerda el valor inicial de `CorreosAutomaticos` (`:471`).

### Normal — justification

1. `btnBuscarJustif_Click` abre `frmBuscarJustificacion` con `OpenArgs:="MODO_SELECCION"` y se enlaza con `WithEvents` sin `acDialog`, con un comentario que explica por qué (`:930-938`).
2. `f_BusquedaJustif_JustificacionElegida` valida que el identificador sea numérico y llama a `RellenarJustificacion`, que resuelve la entidad y pinta id y rótulo (`:942-946`, `:352-358`), y suelta la referencia (`:949`).
3. `cmdLimpiarJustificacion_Click` pone los dos controles a `Null` (`:15-17`), lo que `RellenarObjeto` traduce a cadena vacía (`:599`) y `HaHabidoCambios` detecta como cambio (`:551`).

### Normal — guardado

1. `ComandoRegistrar_Click` llama a `Registrar EnumSiNo.Sí` (`:627`), que aborta con «No hay cambios que guardar» si `HaHabidoCambios` es falso (`:801-804`), vuelca los controles con `RellenarObjeto` (`:806`) y persiste con `RegistrarDatos` (`:813`).
2. Repinta el contenedor padre según su nombre (`:818-822`).
3. De vuelta al handler: si `CorreosAutomaticos` gira de «No» a «Sí» y el proceso ya tiene `FechaEnvioExcel` o `FechaTramitacionAltaMarga`, pregunta y reprograma (`:636-668`); si gira de «Sí» a «No», anula lo programado (`:672-683`).
4. Cierra `FormSolicitudAlta` si estuviera abierta (`:690-692`) y confirma el guardado (`:695`).

### Alternativos

- **Reprogramación fallida tras un guardado correcto:** aviso diferenciado y el guardado no se invalida (`:663-666`).
- **Fallo al limpiar la cola:** se avisa y `m_Error` se limpia (`:680-683`).
- **Búsqueda con Enter o Tab:** `PalabraClave_KeyDown` lanza la búsqueda de expediente (`:195-212`).
- **Baja de justificación:** la consulta abre el buscador, que aplica su propia regla de «en uso no se elimina».

### Errores

- Patrón uniforme de `On Error GoTo errores` con el centinela `1000` para los errores propios; los reales se formatean y se envían a `CorreoAlAdministrador` (`FU.bas:585`).
- `ComandoRegistrar_Click` muestra `MsgBox m_Error` incluso con el texto vacío (`:706`), igual que su gemelo de renovación.

## Reglas, validaciones y cálculos

- **Sin cambios no se escribe:** `Registrar` compara el conjunto de campos con `Nz` y aborta si todo coincide (`:490-570, 801-804`).
- **Justificación numérica:** `EstablecerDatos` sólo la recupera si `IsNumeric(.IDJustificacion)` (`:391`) y el evento del buscador sólo la pinta si el identificador es numérico (`:944`).
- **Cambio de tipo bloqueado en alta y traspaso** (`:384, 396-402`), como en el formulario hermano.
- **Reprogramación condicionada al proceso ya iniciado** (`:640-647`), igual que en renovación.
- **Justificación obligatoria:** no hay ninguna guarda que exija `txtIdJustificacion` antes de registrar; la validación, si existe, estará en `SolicitudOperaciones.RegistrarDatos`, fuera de este `.cls`.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Fotografía inicial | `constructor.getSolicitud` | DAO | `:372` |
| Persistencia | `SolicitudOperaciones.RegistrarDatos` | DAO | `:813` |
| Cola de correos automáticos | `ReprogramarAvisosFuturos`, `AnularTodosLosProgramados` | DAO sobre automatismos | `:661, 678` |
| Justificación elegida | `JustificacionServicio.ObtenerPorID` → `JustificacionRepositorio` | DAO | `:353-354` |
| Rótulo de la justificación | `getJustificacionParaCampo`, que trunca la descripción a 100 caracteres | Derivado | `:357`, `JustificacionServicio.cls:105-111` |
| Combos y rejilla de expedientes | diccionarios de `Entorno` | DAO, cacheados | `:709-786` |

## Coexistencia

Pendiente de `../coexistencia.md`. Este formulario y `Form_FormSolicitudRenovacionDatos` son casi el mismo código con dos diferencias (la justificación y la recarga de la fotografía inicial). En web, la traducción razonable es una única operación de alta/edición de solicitud con la justificación como campo opcional, no dos formularios casi idénticos; pero eso exige decidir si la justificación es obligatoria en algún tipo de trámite, y hoy esa regla no está en el formulario.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-54.** `Err.Numbr` en `ComandoActualizar_Click` (`:224`): la propiedad no existe y el aviso de error sale sin código, como ya ocurre en `Form_FormSolicitudRenovacion.cls:19` y en `Form_FormSolicitudRenovacionDatos.cls:289`. Es la tercera ocurrencia del mismo error en tres formularios distintos.
- **H-HPS-S-55.** `Public Event Alta()` y `Public Event Edicion()` se declaran al principio del formulario (`:4-5`) y **no hay ningún `RaiseEvent` en el `.cls`**: los dos eventos nunca se lanzan, de modo que ningún contenedor puede reaccionar a que la solicitud se haya dado de alta o se haya editado.
- **H-HPS-S-56.** `EstablecerDatos` lee `m_ObjSolicitudActiva.IDSolicitud` sin comprobar que el objeto exista (`:371`). El formulario hermano sí lo hace y aborta con «No se ha abierto el formulario con los parámetros adecuados» (`Form_FormSolicitudRenovacionDatos.cls:368-371`): aquí el mismo escenario termina en un error no controlado.
- **H-HPS-S-57.** `Registrar` no recarga la fotografía inicial después de `RegistrarDatos` (`:813-822`), a diferencia de su gemelo, que sí la vuelve a pedir con `constructor.getSolicitud` (`Form_FormSolicitudRenovacionDatos.cls:856`). `m_ObjSolicitudAlInicio` conserva los valores previos, así que un segundo guardado consecutivo compara contra datos obsoletos: puede no detectar cambios y responder «No hay cambios que guardar», o pasar a `RegistrarDatos` una fotografía anterior a la escritura real.
- `EstablecerDatos` refresca la rejilla de `Form_FormSolicitudesGestion` como efecto secundario de cargar los datos del subformulario (`:468-470`), lo que hace que abrir la ventana ya provoque un efecto sobre otra.
- `btnBuscarJustif_Click` no cierra una instancia previa del buscador antes de abrir otra (`:934`), a diferencia de los patrones que sí lo hacen en otros formularios (`Form_FormSolicitudRenovacionDatos.cls:24-27`).
- `RellenarJustificacion` sólo escribe los controles si la entidad se resuelve (`:355-358`): con un identificador que no devuelve fila, los controles conservan la justificación anterior y el usuario no ve cambio alguno.
- La línea `Me.CorreosAutomaticos.SetFocus` de la rama de alta (`:381`) se ejecuta desde `Form_Load`, antes de que el formulario tenga el foco.
- `Form_Load` tiene comentadas las dos líneas de `AllowEdits` (`:331, 341, 348`), de modo que este formulario no aplica el patrón de bloqueo que sí usa su gemelo en los mismos puntos.
- **Defectos de interfaz heredados (a verificar):** el `Caption` de diseño es «EXPEDIENTES Versión 2024-001 Administrador» (`.form.txt:39`), y seis combos llevan `RowSourceType ="Value List"` con datos reales de empresa (`.form.txt:745, 1330`), de personas (`.form.txt:1464`), de motivos (`.form.txt:1536`) y **correos electrónicos de empleados con una entrada `ddd` sin rellenar** (`.form.txt:1805`) versionados en el fichero.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: un único endpoint de alta/edición de solicitud compartido por alta, traspaso y renovación, con la justificación como campo opcional validado en el servidor; la justificación elegida como referencia a un recurso, no como dos controles de texto; y el refresco de la rejilla como efecto explícito de la respuesta, no como efecto secundario de abrir un formulario.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin solicitud activa, el formulario no se abre y el motivo se comunica.
2. En alta, los rótulos de estado y el botón de actualizar no están visibles; en edición, sí.
3. Elegir una justificación muestra su rótulo y guarda su identificador; limpiarla deja ambos campos vacíos y el guardado no escribe la justificación anterior.
4. Elegir una justificación inactiva no es posible desde este formulario.
5. Sin cambios, no se escribe nada y se informa de que no hay nada que guardar.
6. Tras guardar, un segundo guardado consecutivo detecta correctamente que no hay cambios.
7. Activar los correos automáticos con el proceso ya iniciado pide confirmación y reprograma; desactivarlos anula lo pendiente.
8. Un fallo al reprogramar o al limpiar la cola no invalida el guardado ya realizado.

## Pendientes

- Si la justificación es obligatoria en algún tipo de trámite: la única validación posible estaría en `SolicitudOperaciones.RegistrarDatos`, que no se analiza aquí.
- Quién deveria recibir `Alta` y `Edicion` si se llegaran a lanzar (H-HPS-S-55): no hay ningún `WithEvents` de este formulario en `src`.
- Qué devuelve `JustificacionServicio.ObtenerPorID` con un identificador inexistente y qué deja entonces en pantalla el formulario (`:355-358`).
- Recuento de controles: el inventario fija 66 y el `.form.txt` contiene 65 bloques `Begin <Tipo>`; la diferencia no se reproduce con una regla simple de recuento.