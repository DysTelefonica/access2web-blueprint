# Form_FormCorreoPruebas

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormCorreoPruebas.cls` (259 líneas) y `.form.txt` (14 controles `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `fichero:línea` del commit fijado.

Consola de envío de correos de prueba, pensada para uso interno y entornos de pre-producción. Carga dinámicamente un combo de solicitudes y un combo de tipos de correo (desde `m_ObjEntorno.ColTiposCorreos`), hidrata los campos de destinatario y asunto calculándolos con la misma lógica de negocio que el envío real, y permite previsualizar la plantilla o generar un correo en la cola de envío usando `Automiatizacion.EnviarCorreoGenerico` con `p_IDSolicitudVinculada:=CONST_IDSOLICITUD_PRUEBA`. La pantalla convive con un flag `m_ObjEntorno.ModoNuevoEnvioCorreo` que se activa al generar y se restaura en el bloque `errores:`. Las citas son `Form_FormCorreoPruebas.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormCorreosGestion` | `DoCmd.OpenForm "FormCorreoPruebas"` (sin `OpenArgs`) | `Form_FormCorreosGestion.cls:13-15` |
| Otras vías | (no en `src/` para esta pantalla) | (no en `src/`) |

Estado compartido: globales `m_ObjEntorno`, `m_ObjUsuarioConectado`, `m_ObjSolicitudActiva` (no usadas aquí; usa `m_ObjEntorno.ColTiposCorreos` y carga la solicitud directamente con `constructor.getSolicitud`).

| Hacia | Cómo | Evidencia |
|---|---|---|
| `TbCorreosEnviados` (inserción indirecta) | `Automiatizacion.EnviarCorreoGenerico` con `CONST_IDSOLICITUD_PRUEBA` (inserta fila con `IDSolicitud = -1`) | `:215-232` |
| Sistema de ficheros local | `ObtenerHTMLEnLocal` con `m_SolicitudOrigen` para previsualizar | `:175-182` |

No emite eventos ni abre otros formularios.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `cboSolicitudOrigen` | combo 2 columnas (ID oculto, texto visible) | `RowSourceType = "Value List"`; `AfterUpdate` → `CargarDatosSimulados` | `:38-44, 117-120` |
| `cboTipoCorreo` | combo 2 columnas (clave oculta, etiqueta visible) | `RowSourceType = "Value List"`; `AfterUpdate` → `CargarDatosSimulados` | `:71-78, 122-125` |
| `txtDestinatarios` | texto | `CargarDatosSimulados` lo rellena con `m_SolicitudOrigen.DestinatariosCalculados`; envío lo pasa a `p_DESTINATARIOS` | `:134, 221` |
| `txtCopia` | texto | `CargarDatosSimulados` lo rellena con `m_SolicitudOrigen.DestinatariosConCopiaCalculados(tipoCorreo)`; envío lo pasa a `p_DestinatariosCopia` | `:135, 224` |
| `txtAsunto` | texto | `CargarDatosSimulados` lo rellena con `m_SolicitudOrigen.AsuntoCalculado(tipoCorreo)` con prefijo «PRUEBA: »; envío lo pasa a `p_asunto` | `:136, 222` |
| `txtFechaOrden` | fecha | `IsDate(Me.txtFechaOrden)` decide si la generación es inmediata o programada; si es fecha, `p_FechaOrden:=CDate(Me.txtFechaOrden)`; si no, `p_FechaOrden:=Null` | `:204-209, 226` |
| `cmdCargarDatos` | botón | `cmdCargarDatos_Click` → `CargarDatosSimulados` | `:128-137` |
| `cmdVerHTML` | botón | `cmdVerHTML_Click` → `ObtenerHTMLEnLocal rutaPlantilla, , m_SolicitudOrigen, m_Error` | `:159-187` |
| `cmdEnviarPrueba` | botón | `cmdEnviarPrueba_Click` → confirmación + flag + `Automiatizacion.EnviarCorreoGenerico` | `:190-251` |
| `cmdSalir` | botón | `cmdSalir_Click` cierra sin guardar | `:81-83` |
| `lblInfo` | etiqueta | `CargarDatosSimulados` pinta un mensaje de feedback (`Datos cargados desde Solicitud {id} para fase {tipo}`) | `:138` |

## Flujos

### Carga (`Form_Load`, `:23-79`)

1. `Ajustar Me`; `Me.Caption = "CONSOLA DE ENVÍO DE CORREOS DE PRUEBA"` (`:24-26`).
2. Inicializa `cboSolicitudOrigen` con `RowSourceType = "Value List"`, `ColumnCount = 2`, `ColumnWidths = "0cm;10cm"`, `BoundColumn = 1` (`:30-36`).
3. Construye `m_SQL = "SELECT IDSolicitud, Nombre & ' ' & Apellido1 & ' ' & Apellido2 & ' (' & IDSolicitud & ')' AS NombreCompleto FROM TbSolicitudes ORDER BY IDSolicitud DESC;"` y la ejecuta con `getdb().OpenRecordset(m_SQL)`, recorriendo `rcd` y haciendo `Me.cboSolicitudOrigen.AddItem id & ";" & texto` (`:38-54`).
4. Inicializa `cboTipoCorreo` y lo puebla iterando `m_ObjEntorno.ColTiposCorreos.Keys` con `AddItem key & ";" & m_ObjEntorno.ColTiposCorreos(key)` (`:60-65`).
5. `Me.txtFechaOrden.Enabled = True` y `Me.txtFechaOrden = Null` (envío inmediato por defecto) (`:68-70`).

### Hidratación (`CargarDatosSimulados`, `:128-156`)

1. Sale silenciosamente si `cboSolicitudOrigen` o `cboTipoCorreo` son nulos (`:130-131`).
2. `idSol = Me.cboSolicitudOrigen.Column(0)`, `tipoCorreo = CLng(Me.cboTipoCorreo.Column(0))` (`:134-135`).
3. `Set m_SolicitudOrigen = constructor.getSolicitud(idSol, m_Error)`; aborta si falla (`:138-141`).
4. Pinta los campos: `txtDestinatarios = .DestinatariosCalculados`, `txtCopia = .DestinatariosConCopiaCalculados(tipoCorreo)`, `txtAsunto = "PRUEBA: " & .AsuntoCalculado(tipoCorreo)` (`:144-146`).
5. Feedback en `lblInfo.Caption = "Datos cargados desde Solicitud " & idSol & " para fase " & Me.cboTipoCorreo.Column(1)` (`:148`).

### Previsualización (`cmdVerHTML_Click`, `:159-187`)

1. Sale con aviso si no hay `m_SolicitudOrigen` (`:161-165`).
2. Calcula `tipoCorreo` y `nombrePlantilla = m_SolicitudOrigen.NombrePlantillaParaCorreoCalculada(tipoCorreo)`, monta `rutaPlantilla = m_ObjEntorno.URLDirectorioPlantillasHTMLVersionEstablecida & nombrePlantilla` (`:170-175`).
3. `ObtenerHTMLEnLocal rutaPlantilla, , m_SolicitudOrigen, m_Error` (`:177-179`).

### Envío (`cmdEnviarPrueba_Click`, `:190-251`)

1. Salidas tempranas: sin `txtDestinatarios`, sin `m_SolicitudOrigen`, o con `MsgBox vbNo` en la confirmación (`:192-209`).
2. Activa flag: `m_ObjEntorno.ModoNuevoEnvioCorreo = True` (`:211`).
3. Recalcula `sRecursos = m_SolicitudOrigen.ColParametrosParaHTMLTexto`, `sPlantilla = .NombrePlantillaParaCorreoCalculada(tipoCorreo)`, `sAccion = "PRUEBA: " & .AccionParaCorreoCalculada(tipoCorreo)` (`:213-220`).
4. Si `tipoCorreo = EnvioExcel`, `sAdjunto = m_SolicitudOrigen.URLPlantillaExcel` (`:223-225`).
5. `FechaEnvio = IIf(IsDate(Me.txtFechaOrden), CDate(Me.txtFechaOrden), Null)` (`:228-232`).
6. `Set m_CorreoGen = Automiatizacion.EnviarCorreoGenerico(..., p_IDSolicitudVinculada:=CONST_IDSOLICITUD_PRUEBA, p_FechaOrden:=FechaEnvio, p_Error:=m_Error)` (`:234-241`).
7. Restaura el flag `m_ObjEntorno.ModoNuevoEnvioCorreo = False` siempre (en éxito o en `errores:`) y muestra `MsgBox vbInformation` con `m_CorreoGen.IDCORREO` y `m_CorreoGen.EstadoTexto` (`:243-247`).

### Errores

`Form_Load` no tiene `On Error GoTo errores:`; el `On Error Resume Next` del `cmdSalir` cierra sin diagnóstico. `CargarDatosSimulados` y `cmdVerHTML_Click` saltan al bloque `errores:` y muestran `MsgBox vbCritical` con `m_Error` y `Err.Description`. `cmdEnviarPrueba_Click` asegura la restauración del flag en cualquier rama.

## Reglas, validaciones y cálculos

- **Acoplamiento al dominio:** los tres campos calculados (`Destinatarios`, `Copia`, `Asunto`) usan la misma lógica que el envío real. La previsualización y el envío comparten `m_SolicitudOrigen` y los métodos de cálculo de `Solicitud` y `m_ObjEntorno`. A verificar si la rama de pre-producción se mantiene en sincronía con producción.
- **Flag `m_ObjEntorno.ModoNuevoEnvioCorreo`:** sirve para que `CorreoOperaciones.Registrar` (u otros consumidores) distingan entre un correo real y uno de prueba y desactiven efectos colaterales. Está documentado en el nombre pero no en línea; en este formulario se asegura su restauración en `errores:` y en éxito (`:211, 243-251`).
- **Fecha de orden opcional:** vacío = envío inmediato (`p_FechaOrden:=Null`); fecha válida = programado (`p_FechaOrden:=CDate(Me.txtFechaOrden)`). No se valida que la fecha sea futura, así que un valor pasado también se acepta como «programado en el pasado».
- **Sin validación de tipo de correo:** si `m_ObjEntorno.ColTiposCorreos` no contiene la clave esperada, `cboTipoCorreo` queda con `Null` y `CargarDatosSimulados` aborta silenciosamente.
- **Concurrencia:** el envío se hace sin transacción explícita: un fallo entre el INSERT en `TbCorreosEnviados` y la respuesta de `Automiatizacion.EnviarCorreoGenerico` puede dejar la fila persistida con `Estado = "Pendiente"` para siempre (a verificar el manejo interno del helper).
- **`CONST_IDSOLICITUD_PRUEBA`:** constante simbólica para que la fila de prueba quede identificable. A verificar dónde se define y su valor (no aparece en el `.cls` leído; probablemente en `VG.bas` o en el módulo de la constante).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Carga del combo de solicitudes | `TbSolicitudes` (ordenado por ID DESC) | DAO `OpenRecordset` con SQL | `:38-54` |
| Carga del combo de tipos | `m_ObjEntorno.ColTiposCorreos` (en memoria) | `Scripting.Dictionary` con `key & ";" & m_ObjEntorno.ColTiposCorreos(key)` | `:60-65` |
| Hidratación de campos | `m_SolicitudOrigen` (en memoria) | `Solicitud.DestinatariosCalculados` / `AsuntoCalculado(tipoCorreo)` / `DestinatariosConCopiaCalculados(tipoCorreo)` | `:144-146` |
| Previsualización HTML | Fichero local en `m_ObjEntorno.URLDirectorioPlantillasHTMLVersionEstablecida` | E/S de fichero + render | `:170-179` |
| Generación de correo de prueba | `TbCorreosEnviados` (indirecto, vía `Automiatizacion.EnviarCorreoGenerico`) con `IDSolicitud = -1` | DAO a través del helper central | `:234-241` |

Sin escrituras en `TbSolicitudes` u otras tablas de negocio: la pantalla opera con entidades hidratadas en memoria.

## Coexistencia

`TbCorreosEnviados` migra a `lanzadera.hps_solicitudes.correos_enviados` con `IDSolicitud` como `BIGINT NOT NULL` (la fila de prueba queda con `IDSolicitud = -1`). El combo de solicitudes pasa a un `GET /solicitudes?limit=200` con paginación. `Automiatizacion.EnviarCorreoGenerico` se sustituye por un endpoint `POST /correos/preview` que devuelve el HTML firmado y un `POST /correos` que crea la fila. `m_ObjEntorno.ModoNuevoEnvioCorreo` deja de ser global y pasa a ser un parámetro `modo: "prueba"` en la petición.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-37.** La consulta SQL del combo concatena `Nombre & ' ' & Apellido1 & ' ' & Apellido2 & ' (' & IDSolicitud & ')'` sin alias `& " " &` alrededor de los apellidos: un solicitante sin `Apellido1` muestra el ID pegado al nombre con un espacio doble. `Form_FormCorreoPruebas.cls:41-42`
- **H-HPS-S-38.** El `.cls` declara `m_SolicitudOrigen` (`:5`) pero `Form_Load` no la inicializa; `cboSolicitudOrigen_AfterUpdate` y `cboTipoCorreo_AfterUpdate` ejecutan `CargarDatosSimulados` que sí lo hace, pero si el usuario pulsa «Ver HTML» o «Enviar» sin pasar primero por «Cargar Datos», `m_SolicitudOrigen` queda `Nothing` y el `cmdVerHTML_Click` aborta con `MsgBox "Seleccione primero…"`; el `cmdEnviarPrueba_Click` en cambio aborta con `MsgBox "Debe cargar…"` sin la misma protección homogénea.
- **H-HPS-S-39.** `cmdEnviarPrueba_Click` no valida que `p_FechaOrden` no esté en el pasado: un valor pasado se persiste como programado para una fecha anterior. `Form_FormCorreoPruebas.cls:228-232`
- **H-HPS-S-40.** `Form_Load` no usa `On Error GoTo errores:`: si el `OpenRecordset` de la línea 38 falla (p. ej. permisos), se sale por la rama `errores:` que solo muestra `MsgBox` con `Err.Description` y deja `rcd` sin cerrar. `Form_FormCorreoPruebas.cls:23-79`
- **H-HPS-S-41.** El `Caption` del formulario es literal `"CONSOLA DE ENVÍO DE CORREOS DE PRUEBA"`, sin el sufijo de versión de `m_TituloFormulario` que sí se usa en `Form_FormCorreo` y `Form_FormEnvioCorreo`.
- **H-HPS-S-42.** `CONST_IDSOLICITUD_PRUEBA` se referencia (`:241`) pero no se define en este `.cls`. A verificar en `VG.bas` o en el módulo donde resida; cualquier cambio de valor requiere un redeploy.

**Defectos de seguridad sensibles (no publicados con detalle, fuera del lote):** `m_ObjEntorno.URLDirectorioPlantillasHTMLVersionEstablecida` se evalúa sin allowlist y la ruta llega a `ObtenerHTMLEnLocal` que copia imágenes a un directorio local; una ruta maliciosa apuntaría a un directorio arbitrario. `m_ObjUsuarioConectado.UsuarioRed` se usa como `DesencadenadoPor` en el clon y queda registrado en `TbCorreosEnviados`, sin anonimización.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: validar formato de `IDSolicitud` (entero positivo) en el combo; renombrar `m_SolicitudOrigen` a `m_SolicitudSimulada` para que el dominio de pruebas quede explícito; helper único de mensajes de error y de feedback (eliminar los `MsgBox` dispersos); el flag `ModoNuevoEnvioCorreo` deja de ser global y se inyecta por composición.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir el formulario se cargan las solicitudes y los tipos de correo.
2. Seleccionar una solicitud y un tipo, pulsar «Cargar datos», rellena destinatario, copia y asunto según la misma lógica que el envío real.
3. Pulsar «Ver HTML» muestra la plantilla con los datos de la solicitud cargada.
4. Pulsar «Enviar prueba» genera una fila en `TbCorreosEnviados` con `IDSolicitud = CONST_IDSOLICITUD_PRUEBA` y deja el flag `m_ObjEntorno.ModoNuevoEnvioCorreo = False` aunque falle.
5. Si la fecha de orden está vacía, la generación es inmediata; si es fecha válida, es programada.

## Pendientes

- Valor y ubicación de `CONST_IDSOLICITUD_PRUEBA`.
- Forma del registro de `TbCorreosEnviados` cuando `IDSolicitud = CONST_IDSOLICITUD_PRUEBA`: ¿hay FK virtual? ¿se acepta el `-1`?
- Comportamiento de `Automiatizacion.EnviarCorreoGenerico` ante fallo: la fila queda persistida o se hace rollback.
- Si `m_ObjEntorno.ColTiposCorreos` se actualiza en tiempo de ejecución (recarga de la colección) o se cachea.
- Estructura del combo de tipos: ¿se rellena también `texto` (visible) o solo `clave` (oculta) en la columna 1?
- Si el `.form.txt` declara explícitamente el `Caption` o si este vive como literal en el `.cls` (parece esto último, sin `m_TituloFormulario`).
