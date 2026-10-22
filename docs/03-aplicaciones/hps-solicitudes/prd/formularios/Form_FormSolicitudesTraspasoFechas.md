# Form_FormSolicitudesTraspasoFechas

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormSolicitudesTraspasoFechas.cls` (1166 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Fechas» del alta/traspaso de una solicitud: concentra el ciclo de vida del envío y la cancelación que va desde la plantilla Excel hasta el traspaso a la ONS. Es un formulario **no enlazado** (sin `RecordSource`); pinta los campos desde `m_ObjSolicitudActiva.SolicitudFechas` y concentra ~28 botones «Comando…» que disparan `SolicitudOperaciones` o `SolicitudFechasOperaciones`. No cierra al padre: es la página de fechas de un `NavigationControl` cargado por `Form_FormSolicitudAltaTraspaso`. Las citas son `Form_FormSolicitudesTraspasoFechas.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudAltaTraspaso` (alta+traspaso) | `Me.NavFechas.NavigationTargetName = "FormSolicitudesTraspasoFechas"` cuando `m_Tipo <> "Alta"` | `Form_FormSolicitudAltaTraspaso.cls:122`, `.form.txt:1171` |

Sin salida propia hacia otros formularios; los eventos hacia atrás del flujo de fechas son llamadas a `EstablecerDatos` (re-pintar fechas y visibilidad de los botones) por parte de:

- `Form_FormEnvioCorreo` (programación de recordatorios), tras el envío de cada correo: `Form_FormEnvioCorreo.cls:102, 285`, declarado en `Form_FormEnvioCorreo.form.txt:1305, 1488`.
- `Form_FormAdjuntarExcelSolicitante`, vía evento `Adjuntado` consumido por `m_FormAnexo_Adjuntado` (`:993-1019`); la referencia de la instancia se captura en `ComandoRegistrarExcelRecibido_Click` con `Set m_FormAnexo = Forms(...)` (`:334-345`).
- `Form_FormAdjuntaTraspasoONS`, vía evento `AdjuntadoEnvio` consumido por `m_FormAnexoTraspaso_AdjuntadoEnvio` (`:1131-1166`); la referencia se captura en `ComandoRegistrarEnvioTraspasoONS_Click` con `Set m_FormAnexoTraspaso = Forms(...)` (`:298-303`).

Estado compartido: las globales `m_ObjSolicitudActiva`, `m_ObjEntorno`, `m_TituloFormulario`, y el `FileSystemObject` `fso` (declaradas fuera de este `cls`). El formulario no recibe parámetros.

## Controles y eventos

El formulario es declarativo: las fechas (`FechaEnvioExcel`, `FechaRecepcionExcel`, `FechaEnvioONS`, `FechaCancelado`, `FechaDesestimado`, `FechaRegistroEnHPS`, `FechaCorreoRecordatorioExcel1/2`, `FechaPrevistaCorreoRecordatorioExcel1/2`, `FechaPrevistaAutocancelacionPreMarga`, `DiasParaCancelacion`) son `TextBox` simples sin eventos; los botones se enumeran a continuación. `ComandoURL…` abren la URL local con `AbrirEnLocal`; los `ComandoVer…` abren HTML con `ObtenerHTMLEnLocal` desde `m_ObjEntorno`.

| Botón | Qué dispara | Evidencia |
|---|---|---|
| `ComandoRegistrarExcelEnviado` | genera correo de envío (abre `Form_FormEnvioCorreo` con la plantilla HTML de solicitud de información) | `:39` (focus inicial en `EstablecerDatos`), `.form.txt` |
| `ComandoRegistrarExcelRecibido` | abre `Form_FormAdjuntarExcelSolicitante` y captura `m_FormAnexo` con `WithEvents` | `:318-345` |
| `ComandoEliminarExcelEnviado` / `ComandoEliminarExcelRecibido` | `Eliminar(EnumTipoEnvioCorreo.EnvioExcel, …)` / `SolicitudOperaciones.EliminarExcelDevuelto` | `:385-420`, `:421-461` |
| `ComandoEnvioRecordatorioExcel1/2` | envía el recordatorio Excel 1/2 vía `Form_FormEnvioCorreo` | `:599-625`, `:630-657` |
| `ComandoEliminarRecordatorioExcel1/2` | `Eliminar(EnumTipoEnvioCorreo.RecordatorioExcel1/2, …)` | `:465-494`, `:500-528` |
| `ComandoActualizarRecordatoriosExcelPreMarga` | copia los `…Calculado` a los reales en `m_SolFechas` y los persiste con `SolicitudFechasOperaciones.Registrar` | `:49-113` |
| `ComandoEnvioCancelacionPreMarga` | envía el correo de cancelación pre-MARGA vía `Form_FormEnvioCorreo` | `:537-564` |
| `ComandoVerCancelacionPreMarga`, `ComandoVerHTMLEnvioExcel`, `ComandoVerRecordatorioExcel1/2` | previsualizar el HTML de cada correo mediante `ObtenerHTMLEnLocal` | `:849-899`, `:878-899`, `:907-931`, `:936-959` |
| `ComandoURLExcelSolicitud` | abre la plantilla Excel (alta/traspaso según `TipoCalculado`) en local | `:810-844` |
| `ComandoURLExcelRespuesta` | abre el Excel del solicitante desde `URLAdjuntoCalculado` | `:779-805` |
| `ComandoRegistrarEnvioTraspasoONS` | abre `Form_FormAdjuntaTraspasoONS` modal para anexar el documento de traspaso y captura `m_FormAnexoTraspaso` | `:285-313` |
| `ComandoEliminarEnvioTraspasoONS` | tras confirmar, llama a `SolicitudOperaciones.EliminarEnvioTraspasoONS` | `:154-194` |
| `ComandoURLDocumentoTraspasoEnviado` | abre el adjunto `URLAdjuntoEnvioONSCalculado` en local | `:746-771` |
| `ComandoRegistraCancelacion` | `InputBox("Fecha de Cancelación", …)` → `SolicitudOperaciones.RegistroCancelado` | `:200-237` |
| `ComandoEliminarCancelacion` | tras confirmar, `Eliminar(EnumTipoEnvioCorreo.cancelacion, …)` | `:118-146` |
| `ComandoRegistrarDesestimado` | `InputBox("Fecha de Desestimación", …)` → `SolicitudOperaciones.RegistroDesestimado` | `:241-279` |
| `ComandoEliminarDesestimado` | tras confirmar, `Eliminar(EnumTipoEnvioCorreo.Desestimado, …)` | `:350-380` |
| `ComandoRegistroHPS` | valida `FechaEnvioONS`, llama `SolicitudOperaciones.RegistrarUsuarioHPSEnHPS` y refresca `Form_FormSolicitudesGestion.Filtrar` si cambió el estado | `:664-743` |
| `cmdSalir` | cerrar (no está en este `.cls`; vive en el contenedor `Form_FormSolicitudAltaTraspaso`) | — |

## Flujos

### Carga (`Form_Load`, `:966-988`)

1. `On Error GoTo errores`; `DoCmd.Hourglass True`; `Me.AllowEdits = False` (`:970-973`).
2. `EstablecerDatos m_Error`; aborta con `Err.Raise 1000` si devuelve error (`:975-980`).
3. `DoCmd.Hourglass False`; `Me.AllowEdits = True` (`:982-984`).

### Pintado (`EstablecerDatos`, `:1023-1128`)

1. Si `m_ObjSolicitudActiva` es `Nothing`, sale sin tocar controles (`:1026-1031`).
2. `Set m_SolFechas = m_ObjSolicitudActiva.SolicitudFechas`; si es `Nothing`, también sale (`:1032-1036`).
3. `Me.ComandoRegistrarExcelEnviado.SetFocus` (`:1039`).
4. Por cada campo, asigna `Me.<campo> = m_SolFechas.<campo>` cuando `IsDate`, o `Null` cuando no, y hace `Visible = True/False` del botón «Eliminar» emparejado (`:1042-1120`). Para `FechaEnvioONS` además activa/desactiva `ComandoRegistrarEnvioTraspasoONS`, `ComandoEliminarEnvioTraspasoONS` y `ComandoRegistroHPS` en función de si la solicitud ya está registrada en HPS (`:1091-1097`).
5. `DiasParaCancelacion = DateDiff("d", Date, CDate(Me.FechaPrevistaAutocancelacionPreMarga))` solo si `IsDate` (`:1126-1128`).

### Recepción de eventos diferidos (`m_FormAnexo_Adjuntado`, `:993-1019` y `m_FormAnexoTraspaso_AdjuntadoEnvio`, `:1131-1166`)

Repiten el patrón de pintado: `Me.AllowEdits = False`, `EstablecerDatos`, `Me.AllowEdits = True`, sin lógica de retorno al contenedor. El gancho `WithEvents` se conecta cuando el botón «abrir el anexo» ejecuta `Set m_FormAnexo = Forms("FormAdjuntarExcelSolicitante")` (`:343`) o `Set m_FormAnexoTraspaso = Forms("FormAdjuntaTraspasoONS")` (`:302`), tras cerrar la instancia previa si ya estaba abierta (`:296-298`, `:331-333`).

### Registro de envío ONS (`ComandoRegistrarEnvioTraspasoONS_Click`, `:285-313`)

1. Cierra el `FormAdjuntaTraspasoONS` si estaba abierto (`:295-297`).
2. `DoCmd.OpenForm "FormAdjuntaTraspasoONS"` (`:299`).
3. Captura `Set m_FormAnexoTraspaso = Forms("FormAdjuntaTraspasoONS")` (`:301-303`) para recibir después `AdjuntadoEnvio` → `m_FormAnexoTraspaso_AdjuntadoEnvio` que repinta este formulario.

### Registro en HPS (`ComandoRegistroHPS_Click`, `:664-743`)

1. Verifica que `m_ObjSolicitudActiva.SolicitudFechas.FechaEnvioONS` sea fecha; aborta con «Primero hay que registrar la fecha de envío ONS» si no (`:678-681`).
2. `m_MotivoNoOKRegistrarUsuarioHPSEnHPS = m_SolicitudOp.MotivoNoOKRegistrarUsuarioHPSEnHPS(m_Error)` → si no vacío, aborta (`:686-693`).
3. `SolicitudOperaciones.RegistrarUsuarioHPSEnHPS`; con resultado `"0"` o `"1"`, `MsgBox` («Fin de tramitación») (`:698-707`).
4. Recarga `m_ObjSolicitudActiva` desde `constructor.getSolicitud`; si el estado final ≠ estado inicial, llama `Form_FormSolicitudesGestion.Filtrar` (`:715-726`). Esto significa que el refresco del listado padre depende de la detección local del cambio de estado.

### Apertura de adjuntos (`ComandoURLDocumentoTraspasoEnviado_Click`, `:746-771`)

Si `URLAdjuntoEnvioONSCalculado = ""`, aborta con «No hay ningún Adjunto de envío…» (`:758-760`); si no, llama `AbrirEnLocal`. Equivalente para `ComandoURLExcelRespuesta` contra `URLAdjuntoCalculado` (`:779-805`).

### Eliminaciones confirmadas (`ComandoEliminar*_Click`, parcial)

Patrón común: `pregunta = MsgBox("¿Desea borrar…?", vbExclamation + vbYesNo + vbDefaultButton2, …)`; si ≠ `vbYes`, sale sin tocar nada; si sí, llama a `Eliminar(EnumTipoEnvioCorreo.<fase>, m_Error)` (`SolicitudOperaciones.EliminarFase`) o equivalente directo (`:118-146, :350-380, :385-420, :421-461, :465-494, :500-528`).

## Reglas, validaciones y cálculos

- **Visibilidad por estado:** `EstablecerDatos` conmutada por `IsDate(...)` activa o oculta los botones de borrado y los enlaces a adjuntos. Existe estado «ya registrado en HPS» que deshabilita `ComandoRegistrarEnvioTraspasoONS` y oculta `ComandoEliminarEnvioTraspasoONS` (`:1091-1097`) — la ventana se queda en solo lectura sobre la fase ONS cuando `IsDate(m_SolFechas.FechaRegistroEnHPS)`.
- **Conteo de días para cancelación:** `DiasParaCancelacion = DateDiff("d", CDate(Date), CDate(Me.FechaPrevistaAutocancelacionPreMarga))` solo cuando la fecha prevista es `IsDate` (`:1126-1128`); valor neutro si no.
- **`m_ObjSolicitudActiva.SolicitudFechas = Nothing`:** el pintado sale sin error (`:1034-1036`), pero los botones no se actualizan y un usuario puede intentar registrar una fase sobre un objeto cuyo `SolicitudFechas` ya no existe: el método de operaciones lo trata y devuelve `p_Error` (a verificar cadena exacta para cada fase).
- **Saltarse la fase ONS:** `ComandoRegistroHPS_Click` exige `IsDate(FechaEnvioONS)` antes de invocar `MotivoNoOKRegistrarUsuarioHPSEnHPS` (`:678-681`). El motivo agrega al menos las mismas comprobaciones que el método subyacente (defectuoso hacerlo en dos sitios; ver §Defectos).
- **Concurrencia:** `SolicitudOperaciones.RegistroEnvioTraspasoONS` ejecuta dos `UPDATE` contra `TbSolicitudes` y `TbSolicitudesFechas` sin `BeginTrans`/`CommitTrans` (`SolicitudOperaciones.cls:1680-1690`): un fallo a mitad deja el adjunto copiado y solo parte del estado persistido.
- **`pregunta` y `m_Resultado` son globales implícitas:** los handlers usan `pregunta = MsgBox(...)` y `m_Resultado = .RegistrarUsuarioHPSEnHPS(...)` sin `Dim` (`:130, 711-714, 1219`); se asume que residen en `VG.bas`. A verificar nombres exactos.
- **`Set m_ObjSolicitudActiva = constructor.getSolicitud(...)`** recarga la solicitud tras el registro HPS (`:715-718`) — la frase «delegada al módulo constructor» confirma la indirección: el refresco no aprovecha `m_ObjSolicitudActiva.refresh` u homónimo.
- **Doble pintado:** prácticamente todos los handlers llaman a `EstablecerDatos` y luego `PintarIndicadores`; ambos pintan parcialmente el mismo formulario (defecto menor si `PintarIndicadores` no se usa ya en este `cls`, ver §Defectos).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Evidencia |
|---|---|---|---|
| Pintar fechas | `m_SolFechas.<*>` (cargado por `Solicitud.SolicitudFechas`) | memoria | `:1023-1128` |
| `RegistroEnvioTraspasoONS` | `TbSolicitudes.URLAdjuntoEnvioONS`, `TbSolicitudesFechas.FechaEnvioONS` | `UPDATE` por `Execute` | `SolicitudOperaciones.cls:1680-1690` |
| `EliminarEnvioTraspasoONS` | `TbSolicitudes.URLAdjuntoEnvioONS=Null`, `TbSolicitudesFechas.FechaEnvioONS=Null` | `UPDATE` por `Execute` | `SolicitudOperaciones.cls:1736-1742` |
| `EliminarFase` (cancelación, desestimado, recordatorios, excel enviado) | `TbSolicitudesFechas` | `UPDATE` por `Execute` | `SolicitudOperaciones.cls` (no leído) |
| `SolicitudFechasOperaciones.Registrar` | `TbSolicitudesFechas` | `UPDATE` por `Execute` (llamado desde `ComandoActualizarRecordatoriosExcelPreMarga_Click`) | `:99-105`, `SolicitudFechasOperaciones.cls:1-…` |
| `constructor.getSolicitud` | `TbSolicitudes` | SQL | `Form_FormSolicitudesTraspasoFechas.cls:716` |

`PintarIndicadores` se llama con frecuencia pero su definición está en `FU.bas:41-80` (`PintarIndicadores(p_Form, p_Solicitud, p_Error)`); no se ha verificado que opere sobre este formulario concreto — a confirmar leyendo `FU.bas:41-…` en una pasada posterior.

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `SolicitudOperaciones.RegistroEnvioTraspasoONS` hace dos `UPDATE` sin transacción (`SolicitudOperaciones.cls:1680-1690`); en PostgreSQL ambos `UPDATE` deben envolver en `BEGIN/COMMIT/ROLLBACK` (cambio obligatorio). `Format(Me.IDSolicitud, "000000")` para componer el nombre del archivo (`Solicitud.cls:1231`) es un detalle VBA sin coste web. `fso.CopyFile p_URLLocal, m_CarpetaAnexo & m_NombreArchivo` (`SolicitudOperaciones.cls:1679`) pide reemplazo de almacenamiento: blob storage + URL firmada, no copia a sistema de archivos.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `EstablecerDatos` filtra `m_SolFechas Is Nothing` saliendo sin error (`:1034-1036`); el resto del formulario queda con valores anteriores y la UI no avisa: combinación visible para un usuario tras pulsar uno de los `ComandoEliminar*` que llame a un `Set m_ObjSolicitudActiva.SolicitudFechas = Nothing` interno.
- `ComandoEliminarExcelRecibido_Click` y `ComandoEliminarExcelEnviado_Click` reusan el mismo literal «Al ComandoEliminarExcelEnviado_Click se ha producido el error» (`:457`); confirma patrón copy-paste y no aporta mensaje útil al administrador.
- `ComandoRegistroHPS_Click` duplica la comprobación `IsDate(FechaEnvioONS)` (`:678`) que ya hace `MotivoNoOKRegistrarUsuarioHPSEnHPS`: si la lógica cambia en un sitio, el otro puede divergir.
- `SolicitudOperaciones.RegistroEnvioTraspasoONS` (`cls:1680-1690`) ejecuta dos `UPDATE` sin transacción; un fallo entre ambos deja la copia del archivo en disco sin reflejar en una tabla. A verificar si la app ya no cumple ninguna ACID parcial por flujo.
- `ComandoEliminarCancelacion_Click`, `ComandoEliminarExcelEnviado_Click`, `ComandoEliminarRecordatorioExcel1/2_Click` no vuelven a llamar a `EstablecerDatos` tras `Eliminar`, solo `PintarIndicadores` (`:118-146, :385-420, :465-494, :500-528`) — `PintarIndicadores` puede no repintar los botones de borrado del propio formulario (contrastar con `FU.bas:41`).
- `ComandoEliminarExcelRecibido_Click` sí llama a `EstablecerDatos` (`:455`); la asimetría con los otros `ComandoEliminar*_Click` es un defecto menor de consistencia.
- `pregunta` y `m_Resultado` se asignan sin `Dim` (`:130, 1219`); esto obliga a que existan como globales en `VG.bas`. Cualquier cambio accidental de nombre las silencia.
- `Set m_FormAnexo = Forms("FormAdjuntarExcelSolicitante")` (`:343`) y `Set m_FormAnexoTraspaso = Forms("FormAdjuntaTraspasoONS")` (`:302`) se hacen después de `DoCmd.OpenForm`. Si Access no ha terminado la apertura, el `Forms(...)` puede ser `Nothing` y el `WithEvents` queda inerte.
- Texto heredado: el `.form.txt` deja `Caption = "EXPEDIENTES Versión 2024-001 Administrador"` (`.form.txt:38`) — la ventana de fechas aparece rotulada como un expediente. Es un defecto cosmetic claro.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: dividir el formulario en pestañas (Excel / Pre-MARGA / ONS / HPS); persistir `SolicitudFechas` con un `PUT /solicitudes/<id>/fechas` atómico que abarque los `UPDATE` de `TbSolicitudes` y `TbSolicitudesFechas`; mover la copia del archivo a blob storage firmado; asociar el «Fin de tramitación» a un evento del backend en lugar de comparar `EstadoAlInicio/EstadoAlFinal`.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir el alta+traspaso con `m_Tipo <> "Alta"` carga esta pestaña como `NavFechas`.
2. `EstablecerDatos` rellena cada fecha cuando existe en `SolicitudFechas` y la deja en `Null` en caso contrario, con la visibilidad de los botones coherente.
3. Sin `FechaEnvioONS`, `ComandoRegistroHPS` no actúa (mensaje «Primero hay que registrar la fecha de envío ONS»).
4. Tras registrar el envío ONS, el botón `ComandoURLDocumentoTraspasoEnviado` aparece y abre el adjunto correcto.
5. Tras eliminar el envío ONS, ambos controles (`ComandoURLDocumentoTraspasoEnviado` y `ComandoEliminarEnvioTraspasoONS`) se ocultan.
6. Tras `SolicitudOperaciones.RegistrarUsuarioHPSEnHPS`, `ComandoRegistrarEnvioTraspasoONS` queda deshabilitado y `ComandoEliminarEnvioTraspasoONS` oculto.
7. Las eliminaciones confirmadas piden confirmación con `vbExclamation + vbYesNo + vbDefaultButton2` antes de ejecutarse.

## Pendientes

- Forma exacta de `PintarIndicadores` (`FU.bas:41`): si pinta o no este formulario concreto.
- Definición exacta de `m_Resultado`, `pregunta` y resto de globales implícitas en `VG.bas`.
- Cadena exacta de los mensajes `p_Error` en cada uno de los `SolicitudOperaciones.EliminarFase` para fases `cancelacion`, `Desestimado`, `RecordatorioExcel1/2`, `EnvioExcel`.
- Comportamiento de `m_ObjSolicitudActiva.SolicitudFechas` tras `EliminarFase` (a quién se asigna `Nothing`, dónde se vuelve a cargar).
- `ComandoEliminar*_Click` adicionales que no se han lido explícitamente (más allá de los citados arriba).
- Si el envío del recordatorio 2 depende o no de que haya habido un recordatorio 1 (`ComandoEnvioRecordatorioExcel2_Click` no se ha leído).
- Texto heredado «EXPEDIENTES Versión 2024-001 Administrador» en el `Caption`.
