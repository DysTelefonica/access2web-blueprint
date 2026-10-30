# Form_FormAdjuntaTraspasoONS

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormAdjuntaTraspasoONS.cls` (234 líneas) y `.form.txt` (25 bloques `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_FormAdjuntaTraspasoONS.cls:línea` salvo indicación.

Adjunta a la solicitud el documento de traspaso que el tramitador remite a la ONS. Es un formulario **no enlazado**: elige el fichero ya cumplimentado en el equipo del tramitador, propone la fecha de envío y copia el fichero al directorio de anexos de la solicitud mediante `SolicitudOperaciones.RegistroEnvioTraspasoONS`. Avisa al formulario padre con el evento `AdjuntadoEnvio` (`:7`, lanzado en `:68`).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudesTraspasoFechas` | `ComandoRegistrarEnvioTraspasoONS_Click` cierra la instancia previa, abre este formulario y se suscribe con `WithEvents` | `Form_FormSolicitudesTraspasoFechas.cls:285-301` |

No declara parámetros: toma la solicitud de la global `m_ObjSolicitudActiva` (`VG.bas:197`). No abre formularios. Al terminar el registro se cierra a sí mismo (`:73`). El fichero elegido vive en la variable de módulo `m_URLLocal` (`:9`), no en un control.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnOpen` → `Form_Open` | Dimensiona, carga el estado y cierra la etapa de carga | `:166-197`, `.form.txt:40-41` |
| `lblTitulo` | etiqueta de encabezado | — | Rótulo de diseño «DOC DE LA ONS»; el código lo cambia a «ENVÍO DOC PARA LA ONS» | `:101`, `.form.txt:226-227` |
| `ImagenURL` | imagen pulsable | `OnClick` → `ImagenURL_Click` | Abre el adjunto ya registrado con el visor del sistema | `:199-233`, `.form.txt:241-242` |
| `ComandoAyuda` | botón | `OnClick` → `ComandoAyuda_Click` | Abre el PDF de ayuda | `:18-44`, `.form.txt:289-291` |
| `Etiqueta109` | etiqueta | — | Rótulo «URL» | `.form.txt:390-391` |
| `ComandoSeleccionar` | botón | `OnClick` → `ComandoSeleccionar_Click` | Abre el diálogo de selección de fichero | `:132-164`, `.form.txt:411-413` |
| `URL` | cuadro de texto deshabilitado | — | Muestra `..\<nombre>` como ruta relativa ficticia | `:105`, `:149`, `.form.txt:511` |
| `FECHA` | cuadro de texto con formato de fecha corta | — | Fecha de envío a la ONS | `:107-116`, `.form.txt:546-547` |
| `Etiqueta250` | etiqueta | — | Rótulo «FECHA» | `.form.txt:577-578` |
| `ComandoRegistrar` | button | `OnClick` → `ComandoRegistrar_Click` | Registra el adjunto y avisa al padre | `:46-90`, `.form.txt:610-612` |
| `cmdSalir` | botón «Salir» | `OnClick` → `cmdSalir_Click` | Cierra el formulario | `:13-16`, `.form.txt:701-703` |

## Flujos

### Apertura (`Form_Open`, `:166-197`)

1. `Ajustar Me` dimensiona el formulario (`:174`; `FU.bas:735-759`) y desactiva la edición (`:176`).
2. `EstablecerDatos` (`:92-129`) carga el estado; si no hay solicitud activa, aborta con «No hay datos previos de la solicitud» (`:96-99`).
3. Con adjunto ya registrado, muestra `ImagenURL` y pinta `URL` con `..\` y el nombre del fichero (`:102-105`); si no, oculta la imagen y vacía `URL` (`:106-109`).
4. Si la solicitud tiene `SolicitudFechas`, propone como fecha la de envío ONS ya registrada o, si no hay, la del sistema (`:111-118`).
5. Fija el título de la ventana con `m_TituloFormulario` (`:120`) y cierra la etapa de carga (`:182`).

### Selección del fichero (`ComandoSeleccionar_Click`, `:132-164`)

1. Llama a `Seleccionar(True, "Seleccione el archivo", m_Error)` (`:138`; `FU.bas:678-716`), que muestra el diálogo de `Application.FileDialog` con el último directorio de la sesión.
2. Refleja el resultado en `URL` como ruta relativa; si el usuario cancela, `URL` queda a `Null` (`:143-147`).

### Registro (`ComandoRegistrar_Click`, `:46-90`)

1. Sin fichero elegido, aborta con «Debe seleccionar el archivo a registrar» (`:55-58`).
2. Crea `SolicitudOperaciones`, le asigna `m_ObjSolicitudActiva` y llama a `RegistroEnvioTraspasoONS` con la ruta local y la fecha del control (`:59-66`).
3. Con éxito, lanza `AdjuntadoEnvio` (`:71`), llama a `PintarIndicadores` (`FU.bas:41-83`) y cierra la ventana (`:79`).

### Apertura del adjunto (`ImagenURL_Click`, `:199-233`)

1. Si el fichero local aún existe, lo abre directamente con `Ejecutar … "open"` (`:207-213`).
2. Si no, lo abre desde el destino mediante `AbrirEnLocal` (`:215-219`; `FU.bas:433-472`).

### Errores

- `cmdSalir_Click` usa `On Error Resume Next` sin aviso (`:13-16`).
- `ComandoAyuda_Click`, `ComandoRegistrar_Click`, `ComandoSeleccionar_Click`, `Form_Open` e `ImagenURL_Click` ignoran el centinela `Err.Number = 1000` («error propio ya descrito en `m_Error`»); los errores ajenos rellenan el texto, envían `CorreoAlAdministrador` (`FU.bas:585`) y muestran un `MsgBox` crítico (`:27-42`, `:82-87`, `:156-162`, `:189-195`, `:226-232`).
- `EstablecerDatos` tiene su propio manejador (`:126-129`) que sólo propaga `p_Error`.

## Reglas, validaciones y cálculos

- **Fichero obligatorio:** no se registra sin `m_URLLocal` (`:53-56`).
- **Solicitud activa obligatoria:** `EstablecerDatos` aborta si `m_ObjSolicitudActiva` es `Nothing` (`:95-98`).
- **Solicitud registrada:** `RegistroEnvioTraspasoONS` exige `IDSolicitud` informado (`SolicitudOperaciones.cls:1644-1647`).
- **Recepción del Excel previa:** la fase de envío a la ONS no se puede registrar sin la recepción del Excel del solicitante (`SolicitudOperaciones.cls:1649-1652`).
- **Orden de fechas:** el envío a la ONS no puede ser anterior a la recepción del Excel (`SolicitudOperaciones.cls:1653-1656`).
- **Fichero alcanzable:** se comprueba que exista en la ruta local (`SolicitudOperaciones.cls:1640-1643`).
- **Copia al destino:** el fichero se copia al directorio de anexos de la solicitud, creando la carpeta si falta (`SolicitudOperaciones.cls:1674-1679`), y no se sobrescribe si el destino está abierto o en solo lectura (`SolicitudOperaciones.cls:1661-1673`).
- **Fecha por defecto:** si `SolicitudFechas` no existe, `FECHA` conserva el valor de diseño y la capa de datos sustituye la fecha por la del sistema (`SolicitudOperaciones.cls:1636-1638`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Adjunto y fecha de envío ONS | propiedades calculadas sobre `m_ObjSolicitudActiva` | Derivado, sin DAO en este `.cls` | `:100-119` |
| Nombre del adjunto | `TbSolicitudes.URLAdjuntoEnvioONS` | `UPDATE` con SQL concatenado | `SolicitudOperaciones.cls:1681-1684` |
| Fecha de envío ONS | `TbSolicitudesFechas.FechaEnvioONS` | `UPDATE` con SQL concatenado | `SolicitudOperaciones.cls:1686-1689` |
| Copia del fichero | sistema de ficheros del servidor de ficheros | E/S binaria | `SolicitudOperaciones.cls:1674-1679` |
| Registro en el log y actualización de estado | log general y estado derivado de la solicitud | DAO | `SolicitudOperaciones.cls:1691-1694` |

## Coexistencia

Pendiente de `../coexistencia.md`. Lo que ya se ve: la fase de envío a la ONS es una escritura sobre dos tablas (`TbSolicitudes` y `TbSolicitudesFechas`) más una copia de fichero, y se hace sin transacción explícita: si la segunda sentencia falla, el nombre del adjunto queda guardado sin la fecha, o al revés. En PostgreSQL ambas pasarían a consultas parametrizadas y la fase cabría en una transacción con la copia del fichero.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-80.** Rótulos heredados de otra aplicación: la ventana arranca como «Edición de Suministrador» (`.form.txt:40`) y ningún handler la sobrescribe, el botón de ayuda muestra «Salir» (`.form.txt:290`), el de selección dice «Extender» (`.form.txt:412`) y el de registro dice «Editar» frente a un texto emergente «Registrar» (`.form.txt:611-612`). El rótulo interno de diseño es «DOC DE LA ONS» (`.form.txt:227`) y el código lo cambia por «ENVÍO DOC PARA LA ONS» (`:101`), de modo que el texto visible y el de diseño no coinciden.
- **H-HPS-S-81.** `EstablecerDatos` sólo propone fecha si la solicitud ya tiene `SolicitudFechas` (`:111-118`): cuando no lo tiene, `FECHA` conserva el valor de diseño y `RegistroEnvioTraspasoONS` sustituye la fecha por la del sistema sin avisar (`SolicitudOperaciones.cls:1636-1638`), de modo que la fecha registrada no es la que el usuario ve.
- **H-HPS-S-82.** `RegistroEnvioTraspasoONS` interpola el nombre del adjunto y la fecha en las sentencias SQL (`SolicitudOperaciones.cls:1681-1689`): un apóstrofe en el nombre rompe la sentencia y la fecha se enmascara con el formato literal `mm/dd/yyyy`, que depende de la configuración regional.
- **H-HPS-S-83.** `cmdSalir_Click` arranca con `On Error Resume Next` (`:13-16`): traga cualquier fallo sin aviso ni correo.
- El botón «Registrar» no comprueba que exista solicitud activa: `EstablecerDatos` lo aborta al abrir, pero el estado global puede cambiar entre la apertura y la pulsación, y `RegistroEnvioTraspasoONS` abortaría con «No hay datos» (`SolicitudOperaciones.cls:1632-1635`).
- `m_URLLocal` (`:9`) no se inicializa en `EstablecerDatos`: si ya hay un adjunto registrado, `ImagenURL_Click` sólo lo abre si el fichero sigue existiendo en la ruta local (`:207`), y en caso contrario recurre a `URLAdjuntoEnvioONSCalculado` (`:215`).
- El formulario no comprueba ningún permiso: no hay referencia a `EsTecnico` ni a los permisos por entidad, de modo que la fase de envío a la ONS queda disponible para cualquiera que llegue al formulario de traspaso.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: subida del documento por HTTP en lugar de selección desde una ruta del equipo del tramitador; almacenamiento de objetos en vez de rutas en el sistema de ficheros; y la fecha de envío como un campo obligatorio con valor por defecto explícito y visible.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin fichero seleccionado, no se registra el adjunto y se informa del motivo.
2. Con fichero seleccionado, se copia al expediente, se guardan nombre y fecha de envío y la ventana se cierra.
3. Sin solicitud registrada, o sin la recepción previa del Excel del solicitante, no se escribe y se informa del motivo.
4. Si la fecha elegida es anterior a la de recepción del Excel, no se registra nada.
5. Si el adjunto ya está en el destino y está abierto o en solo lectura, se informa y no se sobrescribe.
6. Con un adjunto ya registrado, el nombre mostrado coincide con el fichero registrado.
7. La fecha que se registra es la que el usuario ve.

## Pendientes

- Si el documento de traspaso a la ONS debe validarse por contenido antes de aceptarlo: en este `.cls` no hay ninguna validación del fichero más allá de que exista.
- Qué ocurre con el adjunto ya registrado cuando se vuelve a abrir el formulario y se elige otro fichero con el mismo nombre.
- Dónde se aplica el aviso de «existe en el destino y está abierto» cuando el directorio de anexos seCalcula a partir de la sesión (`Solicitud.cls`, propiedad `URLDirectorioAnexo`).
- Recuento de controles: el inventario fija 28 y el `.form.txt` contiene 25 bloques `Begin`; la diferencia no se reproduce con una regla simple de recuento.
