# Form_FormCorreosGestion

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormCorreosGestion.cls` (540 líneas) y `.form.txt` (12 controles `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `fichero:línea` del commit fijado.

Listado de los correos generados por la aplicación, con filtros por estado y por solicitud. Se apoya en `WithEvents m_Form As Form_FormCorreo` para reaccionar a los eventos `Alta` y `Editado` y refrescar la lista. Ofrece tres acciones: generar un informe de control en HTML (`ComandoInformeControl_Click` → `InformesOperaciones.GenerarCuerpoInformeControl`), editar el correo seleccionado abriendo `Form_FormCorreo` en `MODO_GESTION`, y eliminar el correo siempre que no esté enviado. Las citas son `Form_FormCorreosGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpciones` | `ComandoGestionCorreos_Click` → `DoCmd.OpenForm "FormCorreosGestion"` (sin `OpenArgs`) | `Form_Form0BDOpciones.cls:135-138` |
| `Form_FormCorreo` | eventos `Alta` y `Editado` que re-enganchan `Filtrar` para refrescar la lista | `:397-403` |
| `Form_FormConfiguracion` | (no explícito en este `.cls`; el menú de opciones abre esta pantalla como una de las opciones) | (no en `src/`) |

Estado compartido: globales `m_ObjUsuarioConectado`, `m_ObjCorreoActivo` (que se fija en `ComandoEditar_Click` antes de abrir el subformulario) y `m_ObjEntorno`.

| Hacia | Cómo | Evidencia |
|---|---|---|
| `Form_FormCorreo` | `DoCmd.OpenForm "FormCorreo", OpenArgs:="MODO_GESTION"` con `m_ObjCorreoActivo` prefijado | `:280-305` |
| `Form_FormCorreoPruebas` | `cmdCorreoPruebas_Click` → `DoCmd.OpenForm "FormCorreoPruebas"` | `:11-15` |
| Sistema de ficheros local (HTML) | `ComandoInformeControl_Click` → `InformesOperaciones.GenerarCuerpoInformeControl` + `GenerarArchivoConHTML` + `Ejecutar(Me.hwnd, "open", m_RutaArchivo, "", "", 1)` | `:17-58` |
| `TbCorreosEnviados` (borrado) | `ComandoEliminar_Click` → `CorreoOperaciones.Eliminar` | `:323-345` |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Estado` | combo `Value List` | `AfterUpdate` y `BeforeUpdate` → `Filtrar` | `:75-78, 372-375` |
| `solicitud` | combo 2 columnas (ID oculto, texto visible) | `AfterUpdate` → `Filtrar` | `:104-141, 200-203` |
| `ListaFiltrados` | lista 6 columnas (cabecera + ID oculto) | `Click` carga `m_CorreoSeleccionado`; `DblClick` → `ComandoEditar_Click` | `:381-395, 432-481` |
| `lblListaDeResultados` | etiqueta | `Filtrar` actualiza su `Caption` con el conteo (`Registros filtrados (N)`) | `:251-254` |
| `lblTitulo` | etiqueta | `Form_Open` copia su texto al `Caption` si no está vacío | `:495-498` |
| `ComandoActualizar` | botón | `ComandoActualizar_Click` → `Filtrar` | `:209-225` |
| `ComandoEditar` | botón | `ComandoEditar_Click` abre `Form_FormCorreo` con `MODO_GESTION` y captura `WithEvents` | `:260-307` |
| `ComandoEliminar` | botón | `ComandoEliminar_Click` borra si no enviado; refresca y deshabilita | `:309-368` |
| `ComandoLimpiarEstado` | botón | `ComandoLimpiarEstado_Click` → `Estado = "Todos"` y `Filtrar` | `:67-72` |
| `ComandoLimpiarSolicitud` | botón | `ComandoLimpiarSolicitud_Click` → `solicitud = "0"` y `Filtrar` | `:80-83` |
| `ComandoInformeControl` | botón | `ComandoInformeControl_Click` | `:17-58` |
| `cmdCorreoPruebas` | botón | `cmdCorreoPruebas_Click` abre `FormCorreoPruebas` | `:11-15` |
| `cmdSalir` | botón | `cmdSalir_Click` cierra | `:6-10` |

## Flujos

### Apertura (`Form_Open`, `:482-509`)

1. `Ajustar Me`; `Me.Caption = m_TituloFormulario`; si `lblTitulo.Caption <> ""`, sobreescribe con él (`:486-498`).
2. `Set m_CorreoServicio = New CorreoServicio` (`:501`).
3. `EstablecerCombos` puebla `Estado` y `solicitud` (`:504`).
4. `Filtrar m_Error` carga la lista inicial con `Estado = "Todos"`, `solicitud = "0"` (`:506-509`).
5. `AvanceCerrar` y libera (`:512-516`).

### Construcción de combos (`EstablecerCombos`, `:85-145`)

- `Estado` como `Value List` con `Todos`, `Enviado`, `Pendiente`, `Programado`; valor por defecto `Todos` (`:88-94`).
- `solicitud` como `Value List` con `ColumnCount = 2`, `ColumnWidths = "0cm;10cm"`, `BoundColumn = 1`. Por defecto `0; Todas las solicitudes` (`:97-115`).
- `m_SQL = "SELECT IDSolicitud, Nombre & ' ' & Apellido1 & ' ' & Apellido2 & ' (' & IDSolicitud & ')' AS NombreCompleto FROM TbSolicitudes ORDER BY Nombre ASC, Apellido1 ASC, Apellido2 ASC, IDSolicitud DESC;"` ejecutado con `getdb().OpenRecordset(m_SQL)`, recorriendo `rcd` y haciendo `Me.solicitud.AddItem id & ";" & rcd!NombreCompleto` (`:117-141`).
- `Me.solicitud.value = "0"` (todas) (`:143`).

### Filtrado (`Filtrar`, `:144-258`)

1. Si `On Error GoTo errores`; `DoCmd.Hourglass True` (`:147-149`).
2. Limpia la lista y pinta la cabecera: `lst.AddItem "ID;Estado;Solicitud;Acción;F. Programada;F. Envío"` (`:152-156`).
3. Captura `m_EstadoSel = Nz(Me.Estado, "Todos")`, vacío si es "Todos" (`:158-161`).
4. `Set m_Col = m_CorreoServicio.ObtenerCorreosGestion(m_EstadoSel, m_SolicitudSel, p_Error)`. Si error, `MsgBox` y sale (`:163-170`).
5. Itera `For Each m_ID In m_Col`; por cada correo, resuelve `m_NombreSolicitud = m_Correo.solicitud.NombreCompletoCapitalizado & " (" & m_Correo.IDSolicitud & ")"` si hay `solicitud`, si no `"Sol. ID: " & m_Correo.IDSolicitud` (`:176-188`).
6. Pinta la línea con 6 columnas: `ID;Estado;Solicitud;Acción;F. Programada;F. Envío` (`:190-196`).
7. `Me.lblListaDeResultados.Caption = "Registros filtrados (" & total & ")"` donde `total = lst.ListCount - 1` (restando la cabecera) (`:249-254`).
8. Deshabilita `ComandoEditar` y `ComandoEliminar` (`:255-256`).

### Selección (`ListaFiltrados_Click`, `:432-481`)

1. Deshabilita `ComandoEditar` y `ComandoEliminar`; sale si la columna 0 está vacía (`:438-447`).
2. Si `m_CorreoSeleccionado Is Nothing` o cambió de ID, `Set m_CorreoSeleccionado = constructor.getCorreo(p_IDCorreo:=Me.ListaFiltrados.Column(0), p_Error)` (`:450-462`).
3. `Me.ComandoEditar.Enabled = True` (`:474`).
4. `Me.ComandoEliminar.Enabled = IIf(IsDate(m_CorreoSeleccionado.FechaEnvio), False, True)` (`:476-479`).

### Edición (`ComandoEditar_Click`, `:260-307`)

1. Sale con `MsgBox` si no hay selección (`:271-275`).
2. Cierra `Form_FormCorreo` si está abierto (`:277-279`).
3. Fija `Set m_ObjCorreoActivo = m_CorreoSeleccionado` y abre `DoCmd.OpenForm "FormCorreo", OpenArgs:="MODO_GESTION"` (`:282-285`).
4. Reengancha los eventos: `If FormularioAbierto("FormCorreo") Then Set m_Form = Forms("FormCorreo")` (`:287-290`).

### Eliminación (`ComandoEliminar_Click`, `:309-368`)

1. Sale con `MsgBox` si no hay selección (`:319-323`).
2. `If IsDate(m_CorreoSeleccionado.FechaEnvio)` → `MsgBox` de «Acción denegada» y sale (`:325-333`).
3. `pregunta = MsgBox("¿Desea realmente eliminar...?", vbYesNo)` (`:339-341`).
4. `Set m_SolOp = New CorreoOperaciones`, `.Eliminar p_Error:=m_Error` (`:344-350`).
5. `ComandoActualizar_Click` y `Filtrar p_Error:=m_Error`; deshabilita los botones (`:352-360`).

### Informe de control (`ComandoInformeControl_Click`, `:17-58`)

1. `m_Cuerpo = InformesOperaciones.GenerarCuerpoInformeControl(m_Error)` (`:24-27`).
2. `m_RutaArchivo = GenerarArchivoConHTML(m_Cuerpo, m_Error)` (`:32-35`).
3. `Ejecutar Me.hwnd, "open", m_RutaArchivo, "", "", 1` (`:40`).
4. `RegistrarLogGeneral "Consulta Informe Control", "Visualización en pantalla por " & m_ObjUsuarioConectado.UsuarioRed, Nothing, m_Error` (`:42`).

### Errores

`Form_Open`, `ComandoActualizar_Click`, `ComandoAyuda_Click`, `ComandoEditar_Click`, `ComandoEliminar_Click` y `ListaFiltrados_Click` comparten el patrón de la casa: `Err.Raise 1000` para errores internos, `CorreoAlAdministrador m_Error` solo si `Err.Number <> 1000`, `DoCmd.Hourglass False` y `MsgBox` crítico. `Filtrar` tiene su propio manejo: `MsgBox` con el `p_Error` y `DoCmd.Hourglass False`. `cmdSalir_Click` y `cmdCorreoPruebas_Click` usan `On Error Resume Next`.

## Reglas, validaciones y cálculos

- **Visibilidad de `ComandoEliminar`:** la única condición es `IsDate(m_CorreoSeleccionado.FechaEnvio)` — un correo «Pendiente» o «Programado» es eliminable, un correo «Enviado» no (`:328-331`).
- **Visibilidad de `ComandoEditar`:** siempre `True` mientras haya selección (`:474`); la lógica de bloqueos la decide `Form_FormCorreo` con `AplicarReglasNegocioVisuales` al abrirse.
- **Filtro compuesto por `m_CorreoServicio.ObtenerCorreosGestion(estado, solicitud)`:** la traducción a SQL no se ve en este `.cls` (es de la capa `CorreoServicio`); a verificar cómo se compone el `WHERE`.
- **Concurrencia:** no hay versión de fila ni `LastModified` en `Eliminar` ni en `Filtrar`. Dos ventanas simultáneas podrían mostrar correos recién borrados (lectura coherente) o hacer clic en `ComandoEditar` sobre una fila borrada por la otra ventana.
- **Reenganche `WithEvents`:** el módulo de clase declara `Private WithEvents m_Form As Form_FormCorreo` (`:5`); el handler `m_Form_Alta` y `m_Form_Editado` (`:397-403`) llama a `Filtrar`. Si `Form_FormCorreo` se cierra por `DoCmd.Close` antes del reenganche, el `m_Form` queda con referencia colgante; el reenganche posterior a `DoCmd.OpenForm` re-asigna `m_Form`, pero cualquier `Filtrar` entre la apertura del nuevo y el reenganche opera contra la referencia anterior.
- **SQL con concatenación de texto:** `m_SQL` concatena `Nombre & ' ' & Apellido1 & ' ' & Apellido2` en la cláusula `SELECT` y la inyecta en `OpenRecordset` por el helper de `getdb()`. El alias usa comillas dobles en el nombre compuesto: según el motor puede requerir comillas o corchetes (no aparecen). A verificar con la configuración regional.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Carga del combo de solicitudes | `TbSolicitudes` | DAO `OpenRecordset` con SQL dinámico (concatenado) | `:117-141` |
| Filtrado de correos | `TbCorreosEnviados` (vía `CorreoServicio.ObtenerCorreosGestion`) | DAO encapsulado en `CorreoServicio` | `:163-170` |
| Eliminación de correo | `TbCorreosEnviados` (DELETE) | `CorreoOperaciones.Eliminar` | `:344-350` |
| Carga del correo seleccionado | `TbCorreosEnviados` (SELECT por `IDCorreo`) | `constructor.getCorreo` | `:450-462` |
| Apertura de `Form_FormCorreo` | (subformulario) | `DoCmd.OpenForm` con `MODO_GESTION` | `:282-305` |
| Apertura de `Form_FormCorreoPruebas` | (subformulario) | `DoCmd.OpenForm` | `:11-15` |
| Informe de control (HTML) | Sistema de ficheros local | `InformesOperaciones.GenerarCuerpoInformeControl` + `GenerarArchivoConHTML` | `:24-40` |
| Log de auditoría | `TbLogsGeneral` (vía `RegistrarLogGeneral`) | DAO centralizado | `:42` |

Sin escrituras en `TbSolicitudes` ni en catálogos. No invoca APIs externas (correo, plantillas) directamente.

## Coexistencia

`TbCorreosEnviados` migra a `lanzadera.hps_solicitudes.correos_enviados` con `Estado` como `ENUM('Programado','Pendiente','Enviado')`. El listado pasa a `GET /correos?estado=...&solicitud=...&page=...` con paginación en servidor. `m_CorreoServicio.ObtenerCorreosGestion` se traduce a una vista materializada o un query con índices. `WithEvents` se sustituye por un store compartido o notificaciones por websocket. `Eliminar` se convierte en `DELETE /correos/{id}` con control de concurrencia (`If-Match` + `ExpectedVersion`). `GenerarArchivoConHTML` + `Ejecutar` se sustituyen por un endpoint que devuelve el HTML firmado por tiempo limitado.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-43.** `ComandoLimpiarSolicitud` y `ComandoLimpiarEstado` no actualizan `m_SolicitudSel`/`m_EstadoSel` si la condición previa es la misma (p. ej. `solicitud = "0"` → `ComandoLimpiarSolicitud` no hace nada por el `If` del `:81`); el comportamiento es correcto, pero el nombre del botón («Limpiar») sugiere que debería limpiarse, no solo reaccionar cuando difiere del valor. `Form_FormCorreosGestion.cls:80-83`
- **H-HPS-S-44.** `ComandoEliminar_Click` permite eliminar correos «Programados»: la columna `FechaEnvio` está vacía y el `IsDate` da `False`, pero el correo está en cola. Si llega a enviarse en la ventana entre `ComandoEliminar_Click` y la ejecución del `DELETE`, se puede borrar la fila de un correo ya entregado. A verificar la operación atómica de la capa `CorreoOperaciones.Eliminar`. `Form_FormCorreosGestion.cls:325-333, 344-350`
- **H-HPS-S-45.** `ComandoInformeControl_Click` registra un log con `m_ObjUsuarioConectado.UsuarioRed` literal; si el usuario de red está en blanco o coincide con el usuario de Windows, la traza no diferencia sesiones concurrentes del mismo usuario (a verificar cómo se compone la traza en `RegistrarLogGeneral`). `Form_FormCorreosGestion.cls:42`
- **H-HPS-S-46.** `m_Form_Alta` y `m_Form_Editado` invocan `Filtrar` sin parámetro: `p_Error` queda como `""` (variable local no declarada con `Optional`); VBA resuelve el parámetro a `Variant/Empty` y la firma `Filtrar(Optional ByRef p_Error As String)` lo acepta. A verificar que `Filtrar` no escriba en `p_Error` cuando la llamada es sin argumentos.
- **H-HPS-S-47.** El `.form.txt` tiene `Estado_BeforeUpdate` además de `Estado_AfterUpdate` (`:372-375` y `:75-78`): ambos llaman a `Filtrar`. Si `Filtrar` modifica el valor del control (no lo hace aquí, pero un futuro cambio podría introducir un bucle), el patrón `BeforeUpdate` + `AfterUpdate` lo ejecutaría dos veces.
- **H-HPS-S-48.** `ComandoInformeControl_Click` no comprueba permisos: cualquier usuario con acceso al formulario imprime el informe. La pantalla se muestra en el menú principal sin chequeo de rol en este `.cls`. `Form_FormCorreosGestion.cls:23-58`
- **H-HPS-S-49.** `ComandoLimpiarSolicitud` y `ComandoLimpiarEstado` no deshabilitan sus botones cuando el valor ya es el de «Todos»/«0», así que el usuario ve un botón inerte y un handler que no hace nada. La precondición `If ... = "0"` es un atajo silencioso.

**Defectos de seguridad sensibles (no publicados con detalle, fuera del lote):** `ComandoInformeControl_Click` ejecuta un informe que pasa por `InformesOperaciones.GenerarCuerpoInformeControl` y `GenerarArchivoConHTML`: el primero puede componer datos personales (correos, destinatarios, adjuntos, fechas) y el segundo escribe a un directorio local temporal que `Ejecutar(Me.hwnd, "open", ...)` abre con el navegador predeterminado. Sin validación de permisos en este formulario, cualquier usuario con acceso al menú puede generar el informe.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: traducir el SQL dinámico del combo a un endpoint `GET /solicitudes` paginado; revisar la firma de `Filtrar(Optional ByRef p_Error As String)` para que la llamada `Filtrar` sin argumento sea explícita y consistente; en la capa `Eliminar`, validar la concurrencia contra el `Estado` actual (rechazar si ha pasado a `Enviado` desde otra ventana); separar `GenerarCuerpoInformeControl` y la persistencia de `TbLogsGeneral` para que la consulta no quede registrada como log de auditoría si la generación falla.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir la pantalla se cargan los combos de estado y solicitud y la lista inicial muestra todos los correos con cabecera.
2. Cambiar el estado o la solicitud filtra la lista y actualiza el contador.
3. Seleccionar un correo sin enviar habilita «Editar» y «Eliminar»; uno enviado solo habilita «Editar».
4. «Editar» abre el subformulario en `MODO_GESTION` con los datos cargados.
5. «Eliminar» pide confirmación y refresca; un correo enviado no se elimina.
6. «Informe de control» genera un HTML en local, lo abre y registra log.

## Pendientes

- Firma de `Filtrar` cuando se invoca desde `m_Form_Alta` / `m_Form_Editado` (sin argumento): tipo de `p_Error` y comportamiento del `ByRef` con un argumento vacío.
- Sentido de la columna `Solicitud` del `RowSource` del combo: ¿es `m_ObjSolicitudActiva.ID` o el correo asociado? (a verificar en el `.form.txt`).
- Implementación real de `m_CorreoServicio.ObtenerCorreosGestion`: SQL concreto, índices, paginación.
- Comportamiento de `ComandoEliminar` cuando la operación falla tras `ComandoActualizar` y antes de `Filtrar`: la lista refleja el estado anterior a la actualización.
- Si `Estado` admite otros valores fuera de los cuatro hardcodeados (p. ej. «Cancelado»).
- Política de retención de los HTML generados en `GenerarArchivoConHTML` (purga o acumulación).
