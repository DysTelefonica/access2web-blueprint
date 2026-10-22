# Form_FormSolicitudesAltaRenovacionFechas

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormSolicitudesAltaRenovacionFechas.cls` (751 líneas) y `.form.txt` (62 controles `Begin`) · **PRD:** [README](../README.md)

Tercer panel de fechas del sistema: una vista **de solo lectura y de solo consulta** de los hitos, pensada para mostrar el avance a quien no tramita. Reagrupa en un mismo formulario los hitos de alta, traspaso y renovación, de ahí el nombre. No tiene ningún botón de registro ni de eliminación de hitos: solo abre `FormEnvioCorreo` para enviar un correo ya programado y `FormWeb` para ver el contenido. Es **no enlazado** (sin `RecordSource`). Las citas son `Form_FormSolicitudesAltaRenovacionFechas.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormEnvioCorreo` | Tras registrar un envío, invoca su `EstablecerDatos` bajo `On Error Resume Next` para repintar | `Form_FormEnvioCorreo.cls:285-288` |

**No hay ninguna otra vía de apertura.** No aparece como subformulario en ningún `.form.txt`, ni se invoca con `DoCmd.OpenForm`, ni lo nombra ningún control `SourceObject` o `NavigationTargetName`. En el rastreo sobre `src/` solo lo menciona `FormEnvioCorreo`, y esa llamada apunta al **módulo de clase**, no a una instancia de formulario (`Form_FormEnvioCorreo.cls:287`). Es, hasta donde llega la lectura, un formulario huérfano: su `.form.txt` no está embebido en ninguna parte. Queda **a verificar** si se abre por algún mecanismo no indexado.

Abre `FormEnvioCorreo` en ocho botones (`:23, 57, 91, 125, 159, 193, 227` y el de registro de ONS) y `FormWeb` a través del visor de plantillas. No declara `WithEvents` ni expone eventos. El contexto llega por la global `m_ObjSolicitudActiva` (`:622`) y la bandera `m_EnOficina` (`VG.bas:191`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ComandoEnvioExcel` | botón | `OnClick` | Abre `FormEnvioCorreo` con el hito de envío de Excel | `:74-105`, `.form.txt:641-643` |
| `FechaEnvioExcel` / `FechaRecepcionExcel` | texto de fecha | — | Fechas de salida y de recepción del Excel | `.form.txt:612, 681` |
| `ComandoAdjuntarExcelRecibido` | botón | **sin handler** | Botón visible en pantalla sin evento asociado | `.form.txt:710-711` |
| `ComandoEnvioRecordatorioExcel1` / `2` | botones | `OnClick` | Envían el correo de recordatorio correspondiente | `:107-172`, `.form.txt:810-812, 911-913` |
| `ComandoEnvioRecordatorioMarga1` / `2` | botones | `OnClick` | Envían el recordatorio de relleno de Marga | `:173-239`, `.form.txt:1076-1078, 1177-1179` |
| `ComandoEnvioCancelacionMarga` / `PreMarga` | botones | `OnClick` | Envían el correo de cancelación | `:8-73`, `.form.txt:1888-1890` |
| `ComandoRegistroEnvioONS` | botón | `OnClick` | Registra la fecha de envío a ONS, con diálogo previo | `:240-293`, `.form.txt:2356` |
| `FechaEnvioONS` | texto de fecha | — | Habilita el registro en ONS solo si el usuario está en oficina | `:674-678`, `.form.txt:2296` |
| `ComandoRegistroHPS` | botón | `OnClick` | Cierra la tramitación; cambia de rótulo según el estado | `:296-376`, `:696-737`, `.form.txt:2255` |
| `FechaRegistroEnHPS` | texto de fecha | — | Cierre de la tramitación | `.form.txt:983` |
| `FechaDesestimado` / `FechaAutocancelacion` | texto de fecha | — | Desenlaces de la solicitud | `.form.txt:1217, 1341` |
| `DiasParaCancelacion` | texto de fecha | — | Días que faltan para la autocancelación | `:686-693`, `.form.txt:1281` |
| `FechaPrevista*` (seis campos) | texto de fecha | — | Próximo correo previsto de cada recordatorio | `.form.txt:1435, 1499, 1561, 1623, 1958, 2069` |
| `ComandoVerHTMLEnvioExcel` / `ComandoVerRecordatorio*` / `ComandoVerCancelacionMarga` | botones «ojo» | `OnClick` | Abre la plantilla HTML del hito en `FormWeb` | `:378-580`, `.form.txt:1651-1653` |

## Flujos

### Carga (`Form_Load` → `EstablecerDatos`, `:581-609`, `:612-750`)

1. Sale sin hacer nada si no hay solicitud activa o si la solicitud no tiene registro de fechas (`:618-626`).
2. Vuelca las dieciséis fechas del objeto, **solo** las que son fecha válida (`:628-673`); a diferencia del panel de renovación, no limpia con `Null` las que faltan.
3. Habilita el registro en ONS solo si `m_EnOficina` es afirmativo (`:674-678`).
4. Calcula los días para la autocancelación con la misma prioridad de post-Marga sobre pre-Marga (`:686-693`).

### Cierre de la tramitación (`ComandoRegistroHPS_Click`, `:296-376`)

1. Exige la fecha de envío a ONS (`:309-311`), guarda el estado inicial y consulta `MotivoNoOKRegistrarUsuarioHPSEnHPS` (`:316-324`).
2. Llama a `RegistrarUsuarioHPSEnHPS` y avisa según haya creado o actualizado el usuario (`:328-344`).
3. Relee la solicitud, repinta y, si el estado ha cambiado, refresca el listado de gestión (`:347-361`).

### Envío de un correo (patrón repetido, ejemplo `:8-39`)

1. Cierra cualquier `FormEnvioCorreo` previa y lo reabre con el `EnumTipoEnvioCorreo` correspondiente (`:20-23`).
2. Desactiva la casilla de envío programado en el formulario de correo (`:25`).
3. Avisa al administrador ante errores no internos (`:32-38`).

### Errores

Los handlers capturan, envían el aviso al administrador cuando el error no es el 1000 interno y muestran `MsgBox` (`:31-38, 64-71, 97-104, 130-137, 163-170, 196-203, 230-237, 285-293, 363-374, 396-404, 425-433, 454-462, 483-491, 512-520, 541-549, 570-578, 592-608, 623-624`).

## Reglas, validaciones y cálculos

- **El registro en ONS depende del puesto:** solo se habilita si `m_EnOficina` es afirmativo (`:674-678`; `VG.bas:191, 341`). El panel de renovación no comprueba esta bandera.
- **Secuencia obligatoria:** sin fecha de envío a ONS no se cierra la tramitación (`:308-311`), y sin alta en Marga no se registra el envío a ONS (`:248-250`).
- **Botón de cierre con doble rótulo:** con fecha de registro en HPS el botón pasa a «Quitar registro en HPS (Fin tramitación solicitud)» y se deshabilitan el envío, el adjuntar y todos los botones de vista previa y envío de recordatorios y cancelaciones (`:696-711`).
- **Sin Excel recibido, no hay recordatorios:** si no hay fecha de recepción, el envío de Excel queda habilitado y todos los botones de recordatorio y cancelación quedan deshabilitados (`:713-737`).
- **Días para la autocancelación:** `DateDiff("d", Date, fechaPrevista)`, con prioridad de post-Marga (`:686-693`).
- **Sin escritura de hitos:** este formulario no registra ni elimina hitos: los botones de los demás paneles no existen aquí. Su única escritura es la que delega en el envío de correo o en el cierre de tramitación.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Lectura del registro de fechas | `m_ObjSolicitudActiva.SolicitudFechas`, ya cargado en memoria | Solo memoria | `Form_FormSolicitudesAltaRenovacionFechas.cls:622-626` |
| Registro de la fecha de ONS | `TbSolicitudesFechas` (vía `RegistroAltaONS`) | DAO | `Form_FormSolicitudesAltaRenovacionFechas.cls:259-261` |
| Cierre en HPS | `TbSolicitudesFechas` + aplicativo de HPS | DAO + servicio | `Form_FormSolicitudesAltaRenovacionFechas.cls:326-344` |
| Relectura de la solicitud | `TbSolicitudes` (vía `constructor.getSolicitud`) | DAO | `Form_FormSolicitudesAltaRenovacionFechas.cls:348` |
| Envío de correo programado | tabla de cola de envíos (vía `FormEnvioCorreo`) | DAO | `Form_FormSolicitudesAltaRenovacionFechas.cls:22-23` |

## Coexistencia

Pendiente del capítulo `coexistencia.md`, que todavía no existe en este PRD. Lo que ya se ve: la vista de hitos se alimenta de un objeto en memoria y no consulta, así que en web dependería por completo de la consulta al servidor; y la condición `m_EnOficina` es una bandera de sesión local, sin equivalente directo: en web sería el rol o el perfil del usuario autenticado.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El formulario no se abre desde ningún punto del código: no figura como subformulario en ningún `.form.txt` ni se invoca con `DoCmd.OpenForm`. La única referencia, `Form_FormEnvioCorreo.cls:287`, invoca `EstablecerDatos` sobre el **nombre del módulo de clase** y no sobre una instancia (`Forms("FormSolicitudesAltaRenovacionFechas")`), dentro de un `On Error Resume Next` (`:285`): si la resolución de nombre falla, el repintado se pierde en silencio. A verificar: si el formulario se abre por algún mecanismo no indexado y si esa llamada llega a ejecutarse.
- `ComandoAdjuntarExcelRecibido` aparece en pantalla con su rótulo completo (`.form.txt:710-711`) y el código lo deshabilita (`:705`), pero **no tiene handler**: en esta vista no hay forma de adjuntar el Excel recibido. El botón queda inerte.
- `EstablecerDatos` solo escribe las fechas que ya son válidas y no pone a `Null` las que faltan (`:628-673`), a diferencia del panel de renovación: al cambiar de solicitud sobre el mismo formulario, los campos pueden conservar valores de la solicitud anterior.
- `Me.DiasParaCancelacion` conserva el valor previo cuando no hay ninguna fecha prevista (`:691-692`): mismo problema que en el panel de renovación, y mismo origen.
- `ComandoRegistroEnvioONS_Click` y `ComandoRegistroHPS_Click` arrastran el mismo literal de aviso copiado de otro formulario, «Al ComandoRegistrarAltaMarga_Click…» (`:288` y `:365`).
- `ComandoRegistroHPS` cambia el rótulo a «Quitar registro en HPS» (`:697`) pero no tiene rama de des-hacer: el texto anuncia una acción que el handler no implementa.
- `Form_Load` es `Private` (`:581`) frente a `Public` en los otros tres formularios del lote: si el formulario se abriera desde el diálogo estándar, el evento no se ejecutaría.
- `Private m_Form As Form` (`:4`) es una referencia sin tipar al formulario de correo; cualquier otro uso sería ambiguo.
- `EstablecerDatos` devuelve tipo `String` y nunca le asigna valor (`:612, 748`), igual que en los otros paneles.
- El `Caption` de diseño (`.form.txt:39`) fija versión y nombre de una persona como metadato residual.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: vista de hitos como línea de tiempo de solo lectura con los mismos filtros de rol; los hitos como tabla de eventos en lugar de columnas fijas; el «Quitar registro en HPS» como una operación real y auditada o como rótulo honesto; la bandera `m_EnOficina` sustituida por el rol autenticado.

## Criterios de aceptación

Propuestos, no aprobados:

1. La vista muestra los hitos registrados de la solicitud y deja en blanco los que no existen.
2. Sin usuario en oficina, el botón de registro en ONS aparece deshabilitado.
3. Sin alta en Marga, el registro en ONS rechaza la acción y explica que falta ese paso.
4. Con fecha de registro en HPS, el botón de cierre muestra el rótulo de quitar registro y los botones de envío y vista previa quedan deshabilitados.
5. Cerrar la tramitación avisa de si el usuario de HPS se ha creado o se ha actualizado.
6. Si el estado de la solicitud cambia, el listado de gestión se refresca.
7. Los botones de vista previa abren la plantilla del hito sin registrar nada.
8. Enviar un correo abre el formulario de envío con el hito correcto y sin envío programado por defecto.

## Pendientes

- Por qué motivo el formulario existe separado si no hay ningún punto de apertura: si el heredero lo dejó como código muerto o si depende de una carga de formulario por nombre no rastreada.
- Si `m_EnOficina` distingue «en oficina» de «en remoto» de verdad o si hoy siempre vale `No` (`VG.bas:339-341`).
- Qué devuelve exactamente `RegistrarUsuarioHPSEnHPS` y por qué el resultado se compara con literales de texto (`:332-344`).
- Si los ocho `OpenForm` a `FormEnvioCorreo` de este `.cls` son los mismos ocho valores del enumerado que los del panel de renovación, o si difieren en algún hito.
- Si `AbrirEnLocal` deja ficheros temporales en la estación del usuario (`.form.txt:1651`).