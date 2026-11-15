# Form_FormAdjuntarExcelSolicitante

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormAdjuntarExcelSolicitante.cls` (215 líneas) y `.form.txt` · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_X.cls:línea` del commit fijado.

Registro del Excel que devuelve la persona solicitante: elige el fichero ya cumplimentado en el equipo del tramitador, comprueba la fecha de recepción y copia el adjunto al directorio de la solicitud. Es un formulario **no enlazado** (sin `RecordSource`); avisa al formulario padre con el evento `Adjuntado` (`:4`, lanzado en `:61`).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudesAltaFechas` | `ComandoRegistrarExcelRecibido_Click` cierra la instancia previa y abre este formulario; se suscribe con `WithEvents` | `Form_FormSolicitudesAltaFechas.cls:474-489` |
| `Form_FormSolicitudesTraspasoFechas` | `DoCmd.OpenForm "FormAdjuntarExcelSolicitante"` | `Form_FormSolicitudesTraspasoFechas.cls:330` |
| `Form_FormSolicitudRenovacionFechas` | `DoCmd.OpenForm "FormAdjuntarExcelSolicitante"` | `Form_FormSolicitudRenovacionFechas.cls:486` |

No declara parámetros. No abre formularios. Al terminar el registro se cierra a sí mismo (`:74`). El estado se toma de las globales `m_ObjSolicitudActiva` (`VG.bas:197`) y `m_TituloFormulario` (`VG.bas:196`); el fichero elegido vive en la variable de módulo `m_URLLocal` (`:5`), no en un control.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `URL` | texto deshabilitado | — | Muestra `..\<nombre>` como ruta relativa ficticia | `:90`, `:127`, `.form.txt:511` |
| `FECHA` | combo con fechas | — (sin handler) | Fecha de recepción del Excel | `:57`, `.form.txt:546` |
| `ImagenURL` | imagen pulsable | `OnClick` | Abre el adjunto con el visor del sistema | `:182-215`, `.form.txt:241-242` |
| `Etiqueta109` | etiqueta | — | Rótulo «URL» | `.form.txt:390-391` |
| `Etiqueta250` | etiqueta | — | Rótulo «FECHA» | `.form.txt:577-578` |
| `ComandoSeleccionar` | botón | `OnClick` | Abre el diálogo de selección de fichero | `:115-147`, `.form.txt:411-414` |
| `ComandoRegistrar` | botón | `OnClick` | Registra el adjunto y avisa al padre | `:41-77`, `.form.txt:610-614` |
| `ComandoAyuda` | botón | `OnClick` | Abre el PDF de ayuda del formulario | `:13-38`, `.form.txt:289-293` |
| `cmdSalir` | botón «Salir» | `OnClick` | Cierra el formulario | `:8-11`, `.form.txt:701-705` |
| `lblTitulo` | etiqueta de encabezado | — | Rótulo «EXCEL DEL SOLICITANTE» | `.form.txt:226-227` |

## Flujos

### Apertura (`Form_Open`, `:149-180`)

1. `Ajustar Me` dimensiona el formulario (`:158`; `FU.bas:735-759`) y `AllowEdits` se desactiva (`:160`).
2. `EstablecerDatos` rellena el estado (`:162`; `:80-112`): sin `m_ObjSolicitudActiva` devuelve el error «No hay datos previos de la solicitud» (`:83-86`); con anexo previo muestra `ImagenURL` y fija `URL` con el nombre del fichero (`:88-94`); con `SolicitudFechas` presente, propone como fecha la de recepción ya registrada o, si no hay, la del sistema (`:95-101`). El título de la ventana se fija al final (`:103`).
3. `AvanceCerrar` y se reactiva `AllowEdits` (`:166-171`).

### Selección del fichero (`ComandoSeleccionar_Click`, `:115-147`)

1. Llama a `Seleccionar(True, "Seleccione el archivo", m_Error)` (`:120`; `FU.bas:678-716`), que muestra el diálogo de `Application.FileDialog` con el último directorio usado por la sesión.
2. Refleja el resultado en `URL` como ruta relativa; si el usuario cancela, `URL` queda a `Null` (`:124-128`).

### Registro (`ComandoRegistrar_Click`, `:41-77`)

1. Sin fichero elegido, aborta con «Debe seleccionar el archivo a registrar» (`:50-53`).
2. Crea `SolicitudOperaciones`, le asigna `m_ObjSolicitudActiva` y llama a `RegistroExcelDevuelto` (`:54-61`).
3. Con éxito, lanza `Adjuntado` (`:61`) para que el formulario padre repinte, llama a `PintarIndicadores` (`FU.bas:41-144`) y cierra la ventana (`:63`, `:67`).

### Apertura del adjunto (`ImagenURL_Click`, `:182-215`)

1. Si el fichero local aún existe, lo abre directamente con `Ejecutar … "open"` (`:189-196`).
2. Si no, lo abre desde el destino mediante `AbrirEnLocal` (`:197-200`; `FU.bas:433-472`).

### Errores

Todos los handlers acumulan el texto en `m_Error`, envían `CorreoAlAdministrador` (`FU.bas:585`) salvo el error interno 1000 y muestran `MsgBox` crítico (`:29-37`, `:69-76`, `:137-146`, `:171-179`, `:206-214`). `cmdSalir_Click` usa `On Error Resume Next` sin aviso (`:8-11`).

## Reglas, validaciones y cálculos

- **Fichero obligatorio:** el registro no procede sin `m_URLLocal` (`:50-53`).
- **Fecha por defecto:** la fecha del sistema solo se propone cuando existe `SolicitudFechas` y aún no hay recepción registrada (`:95-101`); si `SolicitudFechas` es `Nothing`, `FECHA` conserva el valor de diseño y ese valor es el que se envía.
- **Precondiciones en la capa de datos:** `RegistroExcelDevuelto` exige solicitud registrada, envío del Excel previo y recepción no anterior al envío (`SolicitudOperaciones.cls:1467-1491`), comprueba que el fichero es alcanzable (`:1475-1478`) y que ni el destino ni la copia actual están abiertos o en solo lectura (`:1496-1508`).
- **Sin filtro de extensión:** el diálogo ofrece «All Files» con el patrón `*.*` (`FU.bas:699-700`); la validación del contenido del Excel no está en este formulario.
- **Copia en el destino:** el fichero se copia al directorio de anexos de la solicitud (`SolicitudOperaciones.cls:1509-1514`) y se anulan los avisos automáticos futuros de esa fase (`:1535`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Nombre y carpeta del anexo | propiedades calculadas sobre `m_ObjEntorno` | Derivado, sin DAO | `Solicitud.cls:966-1010` |
| Actualizar el nombre del anexo | `TbSolicitudes` | `UPDATE` con SQL concatenado | `SolicitudOperaciones.cls:1516-1519` |
| Actualizar la fecha de recepción | `TbSolicitudesFechas` | `UPDATE` con SQL concatenado | `SolicitudOperaciones.cls:1522-1525` |
| Copiar el fichero | sistema de ficheros del servidor de ficheros | E/S binaria | `SolicitudOperaciones.cls:1509-1514` |
| Anular avisos programados | tabla de automatismos, vía `AutomatizacionRepositorio` | DAO | `SolicitudOperaciones.cls:1535` |
| Registro en el log y actualización de estado | log general y estado derivado de la solicitud | DAO | `SolicitudOperaciones.cls:1528-1530` |

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Lo que ya se ve: los `UPDATE` se construyen concatenando el nombre del fichero y la fecha formateada con `mm/dd/yyyy` (`SolicitudOperaciones.cls:1516-1525`); en PostgreSQL ambos pasarían a consultas parametrizadas, con el formato de fecha resuelto por el servidor. El almacenamiento del anexo es un sistema de ficheros, con la ruta calculada a partir de la sesión (`Solicitud.cls:993-1010`), y la apertura se delega en el visor del sistema operativo.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-18.** El nombre del fichero y la fecha de recepción se interpolan en el SQL (`SolicitudOperaciones.cls:1516-1525`); un nombre con apóstrofe rompe la sentencia y la fecha depende del formato regional para la máscara literal `#mm/dd/yyyy#`.
- **H-HPS-S-19.** Rótulos heredados de otra aplicación: la ventana arranca como «Edición de Suministrador» (`.form.txt:40`), el botón de ayuda muestra «Salir» con texto emergente «Ayuda» (`.form.txt:289-293`), el de selección dice «Extender» (`.form.txt:412`) y el de registro dice «Editar» frente a un texto emergente «Registrar» (`.form.txt:610-614`).
- `EstablecerDatos` muestra el nombre de un anexo ya registrado (`:88-90`) pero deja `m_URLLocal` a vacío, de modo que el registro se rechaza hasta que se vuelva a elegir el fichero (`:50-53`): el rótulo y el estado real no coinciden.
- `Seleccionar` escribe `m_ObjEntorno.URLArchivoUltimo` con la última selección aunque el usuario cancele el diálogo, con `varFile` sin valor asignado (`FU.bas:708-710`).
- El aviso de un error en `EstablecerDatos` no distingue el fallo real del texto que el propio método construye (`FU.bas:467-470`), lo que puede confundirse con un problema de permisos (`FU.bas:465-470`).
- `RaiseEvent Adjuntado` se lanza en el propio formulario que lo declara (`:4`, `:61`): solo lo escucha el padre que se suscribe con `WithEvents`; no hay ningún otro suscriptor.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: subida del fichero por HTTP en lugar de selección desde una ruta del equipo del tramitador; almacenamiento de objetos en vez de rutas en el sistema de ficheros;restrictor por extensión y validación de la plantilla antes de aceptar el anexo; fecha de recepción obligatoria con valor por defecto explícito en la interfaz.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin fichero seleccionado, no se registra el adjunto y se informa del motivo.
2. Con fichero seleccionado, se copia al expediente, se guardan nombre y fecha de recepción, se lanza el refresco del formulario padre y la ventana se cierra.
3. Si la fecha elegida es anterior a la del envío del Excel, no se registra nada.
4. Si la solicitud aún no se ha registrado o falta el envío del Excel, se informa y no se escribe.
5. Si el anexo ya está en el destino y está abierto o en solo lectura, se informa y no se sobrescribe.
6. Con un anexo registrado, el nombre mostrado coincide con el fichero registrado.

## Pendientes

- Si el valor de diseño de `FECHA` se usa alguna vez sin `SolicitudFechas` (`.form.txt:546`).
- Comportamiento de `ComandoRegistrarExcelRecibido_Click` al reabrir el anexo ya registrado (`Form_FormSolicitudesAltaFechas.cls:474-489`).
- Dónde se aplica el «anexo ya existente en el destino» al reutilizar un mismo nombre de fichero en dos solicitudes.