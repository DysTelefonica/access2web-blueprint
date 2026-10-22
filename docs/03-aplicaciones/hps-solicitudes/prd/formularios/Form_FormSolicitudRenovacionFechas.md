# Form_FormSolicitudRenovacionFechas

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormSolicitudRenovacionFechas.cls` (1652 líneas) y `.form.txt` (81 controles `Begin`) · **PRD:** [README](../README.md)

Subformulario de la pestaña «Fechas» de la renovación: es el **panel de hitos** del ciclo de vida de la solicitud. Cada hito tiene un botón de registro, un botón de envío, un botón de vista previa y —cuando ya existe— un botón de eliminación. Todas las escrituras pasan por `SolicitudOperaciones`; este `.cls` decide qué botones se habilitan según qué hitos hay ya registrados. Es **no enlazado** (sin `RecordSource`). Las citas son `Form_FormSolicitudRenovacionFechas.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormSolicitudRenovacion` | `Me.NavDatos.NavigationTargetName` del padre lo declara como destino de la pestaña «Fechas» | `Form_FormSolicitudRenovacion.cls:82`; `.form.txt:497` |

Abre `FormEnvioCorreo` (para los botones de envío y de registro), `FormAdjuntarExcelSolicitante` (para el Excel recibido) y `FormWeb` (para las vistas previas). Reacciona a dos hijos con `WithEvents`: el adjuntador (`:5`, evento en `:1452`). Al cambiar el estado de la solicitud, refresca el listado de gestión si está abierto (`:1136-1139`). El contexto llega por las globales `m_ObjSolicitudActiva` y `m_SolFechas` (`:7`).

## Controles y eventos

Solo los grupos de hitos y los botones; el resto son rótulos de fecha.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ComandoRegistrarExcelEnviado` | botón | `OnClick` | Abre `FormEnvioCorreo` con el hito de envío de Excel; es el control al que salta el foco al cargar | `:910-939`, `:1494`, `.form.txt:944-946` |
| `FechaEnvioExcel` | texto de fecha | — | Fecha del envío del Excel de solicitud | `.form.txt:915` |
| `ComandoRegistrarExcelRecibido` | botón | `OnClick` | Abre `FormAdjuntarExcelSolicitante`; al volver, el evento repinta | `:474-502`, `:1452-1478`, `.form.txt:1013-1015` |
| `FechaRecepcionExcel` | texto de fecha | — | Fecha de recepción del Excel; su presencia muestra el botón de abrir el adjunto | `:1532-1539`, `.form.txt:984` |
| `ComandoURLExcelRespuesta` | botón «E» | `OnClick` | Abre en local el adjunto del solicitante | `:1149-1178`, `.form.txt:2244-2246` |
| `ComandoRegistrarAltaMarga` | botón | `OnClick` | Registra el alta en Marga por el tramitador | `:505-547`, `.form.txt:2163-2165` |
| `FechaTramitacionAltaMarga` | texto de fecha | — | Puerta de entrada: sin ella no se puede registrar el envío a ONS | `:427`, `.form.txt:1255` |
| `ComandoRegistrarDPS` | botón | `OnClick` | Registra el envío del DPS | `:373-416`, `.form.txt:3200` |
| `FechaEnvioDPS` | texto de fecha | — | Fecha del envío del DPS | `.form.txt:3140` |
| `ComandoRegistrarEnvioONS` | botón | `OnClick` | Pide por diálogo la fecha y registra el alta en ONS | `:418-471`, `.form.txt:3688` |
| `FechaEnvioONS` | texto de fecha | — | Fecha del envío a ONS; su presencia habilita el registro en HPS | `:1583-1607`, `.form.txt:3563` |
| `ComandoRegistroHPS` | botón | `OnClick` | Cierra la tramitación creando o actualizando el usuario en el aplicativo de HPS | `:1068-1147`, `.form.txt:3523` |
| `FechaRegistroEnHPS` | texto de fecha | — | Cierre de la tramitación | `.form.txt:3494` |
| `ComandoRegistrarDesestimado` | botón | `OnClick` | Marca la solicitud como desestimada | `:332-371`, `.form.txt:2126-2128` |
| `ComandoRegistrarCancelacion` | botón | `OnClick` | Registra la cancelación | `:291-330` |
| `FechaDesestimado` / `FechaCancelado` | texto de fecha | — | Fechas de los dos desenlaces | `.form.txt:1489, 1613` |
| `ComandoEliminar*` (Excel enviado, Excel recibido, alta Marga, recordatorios 1 y 2, DPS, ONS, desestimado, cancelación) | botones | `OnClick` | Borrado del hito correspondiente | `.form.txt:2615-2617, 2680-2682, 2745-2747, 2810-2812, 2875-2877` |
| `ComandoEnvioRecordatorio*` / `ComandoEnvioCancelacion*` | botones | `OnClick` | Envían el correo del hito sin registrar la fecha | `:848-1066`, `.form.txt:1114-1116, 1888-1890` |
| `FechaPrevista*` (seis campos) | texto de fecha | — | Próximo correo previsto por cada recordatorio | `.form.txt:1707, 1771, 1833, 1895, 1958, 2059` |
| `DiasParaCancelacion` | texto de fecha | — | Días que faltan para la autocancelación, derivado de la fecha prevista | `:1632-1638`, `.form.txt:1553` |

## Flujos

### Carga (`Form_Load` → `EstablecerDatos`, `:1422-1450`, `:1480-1650`)

1. Sale sin hacer nada si no hay solicitud activa o si la solicitud no tiene registro de fechas asociado (`:1486-1493`).
2. Pone el foco en el botón de envío de Excel (`:1494`) y vuelca las dieciséis fechas del objeto, cada una con su botón de eliminación visible solo si la fecha existe (`:1495-1629`).
3. Calcula los días que faltan para la autocancelación: toma la prevista de post-Marga si existe y, si no, la de pre-Marga (`:1631-1639`).

### Habilitación en cascada (`:1583-1607`)

1. Si hay fecha de envío a ONS y además hay fecha de registro en HPS, deshabilita los botones de ONS y de HPS y oculta el de eliminación: el ciclo ya está cerrado (`:1586-1590`).
2. Si hay envío a ONS sin registro en HPS, habilita ambos y muestra el de eliminación (`:1591-1594`).
3. Si no hay envío a ONS, habilita el registro en ONS solo si ya hay fecha de envío del DPS (`:1597-1603`) y deja el registro en HPS deshabilitado (`:1607`).

### Registro de un hito (patrón repetido, ejemplo `:418-471`)

1. Comprueba la precondición del hito anterior: sin alta en Marga no se registra el envío a ONS (`:427-430`).
2. Pide la fecha con `InputBox`, con la fecha de hoy como valor por defecto, y sale si el usuario cancela (`:431-434`).
3. Delega en `SolicitudOperaciones`, invalida la caché de fechas de la solicitud y vuelve a pintar (`:436-465`).
4. Avisa al administrador ante errores no internos y restaura el estado de la interfaz (`:467-471`).

### Cierre de la tramitación (`ComandoRegistroHPS_Click`, `:1068-1147`)

1. Exige la fecha de envío a ONS (`:1080-1083`) y consulta `MotivoNoOKRegistrarUsuarioHPSEnHPS` para no dejar la solicitud a medias (`:1087-1097`).
2. Llama a `RegistrarUsuarioHPSEnHPS` y distingue entre usuario creado y usuario actualizado por el valor devuelto (`:1100-1117`).
3. Relee la solicitud, compara el estado antes y después y, si ha cambiado, refresca el listado de gestión (`:1120-1141`).
4. Repinta hitos e indicadores globales (`:1142-1144`).

### Borrado de un hito (`Eliminar`, `:10-46`)

1. Delega en `SolicitudOperaciones.EliminarFase` con el tipo de hito (`:16-23`).
2. Invalida la caché, repinta y llama a `PintarIndicadores`, que refresca contadores globales y menús (`FU.bas:41-72`).

### Errores

Todos los handlers capturan, envían el aviso al administrador cuando el error no es el 1000 interno, restauran `AllowEdits` y muestran `MsgBox` (`:59-69, 128-140, 199-211, 233-245, 268-280, 303-315, 344-356, 385-397, 430-442, 487-499, 518-530, 561-573, 604-616, 643-655, 679-691, 723-735, 792-804, 825-837, 860-871, 892-903, 973-985, 1005-1017, 1037-1049, 1071-1083, 1130-1142, 1162-1174, 1193-1205, 1222-1234, 1251-1263, 1280-1292, 1309-1321, 1338-1350, 1367-1379, 1396-1408, 1425-1437, 1465-1478, 1136-1147`). La rama de `m_FormAnexo_Adjuntado` muestra el `MsgBox` vacío: compone el aviso pero no lo enseña (`:1471-1477`).

## Reglas, validaciones y cálculos

- **Secuencia obligatoria de hitos:** alta en Marga → envío a ONS → cierre en HPS; y envío de DPS → envío a ONS (`:427, 1080`). El botón se habilita solo cuando la precondición está satisfecha (`:1583-1607`).
- **Cierre irreversible por interfaz:** con fecha de registro en HPS, los botones de ONS y de HPS quedan deshabilitados y el de eliminación de ONS se oculta (`:1586-1590`).
- **Fechas nulas explícitas:** cada campo se pone a `Null`, no a cadena vacía, cuando `IsDate` falla (`:1496-1499` y siguientes). La comprobación equivalente en la rama de ONS usa comparación con cadena vacía (`:1598`), criterio distinto del resto del método.
- **Días para la autocancelación:** `DateDiff("d", Date, fechaPrevista)`, con prioridad de la fecha de post-Marga sobre la de pre-Marga (`:1631-1639`). Si ninguna existe, el campo conserva el valor anterior.
- **Fecha por defecto en los diálogos:** los diálogos de registro ofrecen la fecha de hoy como valor inicial (`:432`).
- **Borrado en cascada:** eliminar un hito limpia también los posteriores que dependen de él (`SolicitudOperaciones.cls:1332-1351`).
- **Validación del cierre:** `MotivoNoOKRegistrarUsuarioHPSEnHPS` bloquea el registro en HPS si la solicitud no cumple las condiciones mínimas (`:1088-1097`; `SolicitudOperaciones.cls:1081`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Actualización de hitos | `TbSolicitudesFechas` (SQL concatenado con `IDSolicitud`) | DAO | `SolicitudOperaciones.cls:1339-1360` |
| Alta de fecha de ONS | `TbSolicitudesFechas` (alta o actualización según exista) | DAO | `SolicitudOperaciones.cls:2043-2057` |
| Cierre en HPS | `TbSolicitudesFechas` + escritura en el aplicativo de HPS | DAO + servicio | `SolicitudOperaciones.cls:1158` |
| Relectura de la solicitud | `TbSolicitudes` (vía `constructor.getSolicitud`) | DAO | `Form_FormSolicitudRenovacionFechas.cls:1120` |
| Relectura del adjunto | fichero del adjunto del solicitante | E/S de fichero | `Form_FormSolicitudRenovacionFechas.cls:1157` |
| Programación de recordatorios | tabla de envíos programados (vía `FormEnvioCorreo`) | DAO | `Form_FormSolicitudRenovacionFechas.cls:919-921` |

## Coexistencia

Pendiente del capítulo `coexistencia.md`, que todavía no existe en este PRD. Lo que ya se ve: la tabla de hitos tiene más de una quincena de columnas de fecha y se actualiza con sentencias `UPDATE` construidas por concatenación (`SolicitudOperaciones.cls:1339-1360`), un modelo que en web pide una tabla de hitos con una fila por evento en lugar de columnas fijas; la programación de recordatorios depende de un formulario aparte y de una tabla de cola, y la vista previa se genera como HTML en disco y se abre en `FormWeb`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `EstablecerDatos` sale sin composición cuando no hay solicitud activa o no hay registro de fechas (`:1486-1493`), dejando los controles con los valores del diseño en lugar de vacíos: la diferencia entre «sin hitos» y «no hay solicitud» no se ve en pantalla.
- El foco se pone en `ComandoRegistrarExcelEnviado` sin comprobar que esté habilitado (`:1494`); si el control está deshabilitado por diseño, la llamada falla o queda sin efecto.
- La rama de ONS en `EstablecerDatos` usa `If Me.FechaEnvioDPS <> "" Then` (`:1598`) mientras que el resto del método usa `IsDate`. Una fecha nula y una cadena vacía se comportan distinto: la condición puede quedar satisfecha con un valor no fechado.
- La cascada de habilitación se recalcula al leer, no al escribir: tras marcar un hito desde otro camino (por ejemplo desde `FormEnvioCorreo`), este subformulario no se refresca salvo que se llame a `EstablecerDatos` (`:1480`).
- `ComandoRegistrarEnvioONS_Click` y `ComandoRegistroHPS_Click` comparten literalmente el mismo texto de aviso de error (`"Al ComandoRegistrarAltaMarga_Click se ha producido el error n: "`, `:467` y `:1131`): el aviso miente sobre el origen del fallo.
- La rama `errores:` de `m_FormAnexo_Adjuntado` compone el mensaje, envía el correo y no muestra ningún `MsgBox` (`:1465-1478`): el usuario no ve el fallo tras adjuntar el Excel.
- `Me.DiasParaCancelacion` conserva el valor previo cuando no hay ninguna fecha prevista (`:1636-1638`): al abrir una solicitud ya caducada puede mostrarse un contador obsoleto.
- El borrado de un hito llama a `PintarIndicadores` sin pasar `p_Error` (`:41-45`); si el repintado falla, el error se traga (`FU.bas:69-72`).
- `EstablecerDatos` es el punto de relectura tras casi todos los handlers, pero dos de ellos llaman a `PintarIndicadores` **antes** de repintar (`:1120-1144` frente a `:456-465`): el orden no es homogéneo.
- Los seis campos `FechaPrevista*` no tienen ningún botón de edición: son de solo lectura y su origen no es evidente desde el formulario.
- `ComandoRegistroHPS_Click` distingue el resultado con literales `"0"` y `"1"` (`:1108-1116`); cualquier valor nuevo deja el trámite sin mensaje.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: los hitos como tabla de eventos en lugar de columnas; los botones de registro y de borrado unificados en una acción de máquina de estados con la transacción en servidor; la programación de recordatorios como tarea programada en segundo plano; la vista previa como plantilla en disco o como documento servido.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin alta en Marga registrada, el botón de envío a ONS rechaza la acción y explica que falta ese paso.
2. Sin envío del DPS, el botón de envío a ONS aparece deshabilitado.
3. Registrar un hito abre el diálogo con la fecha de hoy propuesta y no hace nada si se cancela.
4. Registrar un hito deja su campo de fecha informado y hace visible su botón de eliminación.
5. Eliminar un hito borra su fecha y su botón de eliminación y limpia los hitos posteriores dependientes.
6. El campo de días para la autocancelación muestra la diferencia con la fecha prevista de post-Marga si existe, y con la de pre-Marga si no.
7. Con fecha de registro en HPS, los botones de ONS y de HPS quedan inutilizables.
8. Cerrar la tramitación crea o actualiza el usuario en el aplicativo de HPS y avisa de cuál de las dos cosas ocurrió.
9. Si el estado de la solicitud cambia, el listado de gestión se refresca.
10. Un error no interno se comunica al usuario además de enviarse al administrador.

## Pendientes

- Valores del enumerado `EnumTipoEnvioCorreo` que cada botón de eliminación pasa a `Eliminar` (`Form_FormSolicitudRenovacionFechas.cls:16-23`).
- Dónde se calculan y escriben las seis fechas `FechaPrevista*`, que este formulario no modifica (`.form.txt:1707-2059`).
- Si `RegistrarUsuarioHPSEnHPS` escribe en otro sistema y con qué credenciales (`SolicitudOperaciones.cls:1158`).
- Qué hace `m_ObjSolicitudActiva.SolicitudFechas = Nothing`: invalida la caché o descarta el objeto vivo (`:454`).
- Origen y permisos del fichero que abre `AbrirEnLocal` (`:1157`).