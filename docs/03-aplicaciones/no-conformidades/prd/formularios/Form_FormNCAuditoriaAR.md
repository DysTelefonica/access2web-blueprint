# Form_FormNCAuditoriaAR

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCAuditoriaAR.cls` (654 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ventana de alta y detalle de una tarea (acción realizada, AR) de una acción correctiva (AC) de una NC de auditoría: pide el texto de la tarea, las fechas de inicio, fin previsto y fin real, notas y responsable, y permite adjuntar evidencias. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); trabaja con las globales `m_ObjACAuditoriaActiva`, `m_ObjARAuditoriaActiva` y `m_ObjNCAuditoriaActiva` y guarda con `ARAuditoriaOperaciones.Registrar`. Las citas son `Form_FormNCAuditoriaAR.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`, `ARAuditoriaOperaciones.cls` abrevia `classes/ARAuditoriaOperaciones.cls` y `VG.bas`, `modules/Variables Globales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormNCAuditoriaAcciones` | `DoCmd.OpenForm "FormNCAuditoriaAR", OpenArgs:=m_ObjNCAuditoriaActiva.Estado` | `Form_FormNCAuditoriaAcciones.cls:363, 499` |

Sin la AC activa muestra «El formulario se ha abierto con parámetros insuficientes» (`:45-48`). Abre `FormARAuditoriaDocumentos` desde `ComandoDocumentos`, desde `ImagenAnexo` y al cerrar una tarea sin evidencia (`:354, 447`), y escucha sus eventos `Alta` y `Eliminado` (`:5, 564-609`; eventos declarados en `Form_FormARAuditoriaDocumentos.cls:7-9`). Se cierra con `cmdSalir` (`:227-230`), tras registrar (`:504`) o ante un error de apertura (`:562`). Declara los eventos `Alta`, `Editado`, `AltaAnexo` y `EliminadoAnexo` (`:6-9`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | form | `OnOpen` | `Ajustar Me`, guarda `OpenArgs` y llama a `EstablecerDatos` | `.form.txt:41`, `:535-563` |
| `lblTitulo` | etiqueta | — | «Alta de tarea de AC de auditoría» o «Detalle de tarea de AC de auditoría» (y `Me.Caption`) | `:62-63, 93-94` |
| `AC` | texto, bloqueado (rótulo «A.Correc») | — | Texto de la AC padre | `:58`, `.form.txt:754-755` |
| `AccionRealizada` (rótulo «Tarea») | texto | `AfterUpdate` | Texto de la tarea; repinta bordes | `:221-225`, `.form.txt:534-535` |
| `FechaInicio`, `FechaFinPrevista`, `FechaFinReal` | texto (fecha) | `AfterUpdate` | Fechas de la tarea; repintan bordes | `:517-533`, `.form.txt:589-711` |
| `ESTADOAR` | texto, bloqueado | — | Estado de la tarea; oculto en el alta | `:49-50, 86-87, 116-118`, `.form.txt:809-810` |
| `Notas` | texto | — | Notas libres | `:102-106` |
| `Responsable` | combo (lista de valores) | — | Nombres de usuarios de Calidad (`FG.bas:3075-3105`) | `:22, 125-129` |
| `ImagenAnexo` | imagen | `OnClick` | Visible si hay documentos; con uno solo lo abre en local | `:120-124, 611-654`, `.form.txt:939-947` |
| `ComandoActualizar` (rótulo «Contador») | botón | `OnClick` | Recalcula el estado de la tarea y de su AC | `:232-298`, `.form.txt:865-878` |
| `ComandoDocumentos` («Documentos») | botón | `OnClick` | Abre la gestión de evidencias | `:343-369`, `.form.txt:1265-1278` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` | Guarda la tarea | `:371-515`, `.form.txt:1217-1230` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | `:227-230`, `.form.txt:1114-1125` |
| `lblEstadoAR`, `Etiqueta38/52/55/61/78/80/82`, `EncabezadoDelFormulario`, `Detalle`, `PieDelFormulario` | rótulos y secciones | — | Sin código (`lblEstadoAR` se oculta en el alta) | `:49`, `.form.txt:475-1104` |

## Flujos

### Apertura (`EstablecerDatos`, `:40-150`)

1. Oculta estado y `ComandoActualizar`, rellena el combo y muestra el texto de la AC (`:49-58`).
2. **Alta** (sin AR activa): título de alta, responsable de la AC, deshabilita Actualizar y Documentos y pinta bordes (`:61-75`).
3. **Detalle**: recarga la AR con `constructor.getARAuditoria` (`:77`); si no existe, «No parece estar registrada la tarea» (`:82-85`); muestra estado, activa Actualizar y Documentos y vuelca campos (`:86-129`).
4. Si la NC está borrada, bloquea la edición y Registrar (`:133-135`); pinta bordes (`:142`).

### Registrar (`ComandoRegistrar_Click`, `:371-515`)

1. Rechaza a técnicos («No tiene autorización para esa acción») y la ausencia de cambios («No hay cambios que guardar») (`:387-398`, `:301-341`).
2. Vuelca los controles en el objeto `ARAuditoria` (`:400-426`).
3. En edición, si se informa `FechaFinReal` por primera vez y no hay documentos, pregunta si adjuntar la evidencia; si responde «Sí», abre `FormARAuditoriaDocumentos` con `OpenArgs` = fecha real y **sale sin guardar**; si «No», sale (`:433-452`).
4. Si cambia una `FechaFinPrevista` ya informada, pide por `InputBox` el motivo de la replanificación; vacío ⇒ «No se ha grabado el cambio» (`:454-462`).
5. Llama a `ARAuditoriaOperaciones.Registrar`, avisa «Registrado correctamente», lanza `Alta`/`Editado`, refresca `FormNCAuditoriaGestion` si cambió estado o texto y cierra (`:465-504`).

### Actualizar (`ComandoActualizar_Click`, `:232-298`)

Recalcula la AR (`ActualizarDatosARAuditoria`); si cambia el estado, recalcula la AC y, si ésta cambia, recarga la NC de `FormNCAuditoria` y la lista de acciones (`:244-283`).

### Anexos

`m_FormAnexos_Alta` y `m_FormAnexos_Eliminado` invalidan los documentos de la AR y re-ejecutan `Form_Open` (`:573-574, 596-597`).

### Errores

Error 1000 ⇒ advertencia; otros ⇒ correo al administrador y `MsgBox` crítico (`:291-297, 508-514`).

## Reglas, validaciones y cálculos

- **Validación `MotivoNoOK` sin invocar** (`ARAuditoriaOperaciones.cls:16-111`): la función privada define estas reglas, pero ningún punto de la clase ni del formulario la llama, por lo que `Registrar` no las aplica (ver Posibles defectos). Reglas definidas: exige acción realizada y AC (`:27, 36, 40`); fin real sin fin previsto, fin previsto sin inicio, o inicio sin fin previsto, se rechazan (`:43-58`); fechas coherentes con `BuenaFecha` (`:63-67`, p. ej. «La fecha Final real es anterior a la inicial»); cerrar con fin real exige documentos (`:70-74, 82-86`); no admite otra tarea con el mismo nombre en la misma AC (`:91-99`).
- **Bordes de obligatorios** (`:151-220`): rojo si falta la tarea o las fechas son incoherentes (fin previsto anterior al inicio o a hoy sin fin real; fin real anterior al inicio); se omiten si la NC está borrada.
- **Saneado y responsable:** elimina comillas simples y sustituye «;» por «:» en la tarea; si el responsable está vacío, toma el de la AC (`ARAuditoriaOperaciones.cls:138-143`).
- **Estado:** `EstadoCalculadoTexto` en alta y edición (`ARAuditoriaOperaciones.cls:207, 250`).
- **Replanificación:** si cambia el fin previsto se inserta un registro de replanificación con las fechas inicial y nueva y las observaciones (`ARAuditoriaOperaciones.cls:259-277`).
- **Tras registrar:** recalcula la AC y la NC y reinicia indicadores si cambió el estado de la AC o hubo replanificación (`ARAuditoriaOperaciones.cls:288-316`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Evidencia |
|---|---|---|---|
| Leer AR | vía `constructor.getARAuditoria` | Pendiente: no leído | `:77` |
| Alta de AR | `TbNCAuditoriaAccionesRealizadas` | DAO `AddNew`/`Update` | `ARAuditoriaOperaciones.cls:145, 171, 212` |
| Edición de AR | `TbNCAuditoriaAccionesRealizadas` | SQL dinámico concatenado + DAO `Edit`/`Update` | `ARAuditoriaOperaciones.cls:159-160, 220, 253` |
| Replanificación | `ReplanificacionesAuditoriaOperaciones.Registrar` | Alta en tabla de replanificaciones (no leída) | `ARAuditoriaOperaciones.cls:269-272` |
| Log de cambios | `RegistrarLogAuditoria` | Alta/Edición de AR | `ARAuditoriaOperaciones.cls:213, 255` |

Sin transacción: AR, log, replanificación y recálculos son operaciones separadas, por lo que un fallo intermedio deja datos parciales.

## Coexistencia

Recordset DAO sobre `TbNCAuditoriaAccionesRealizadas` y `FechaAccionRealizada = Now()` (`ARAuditoriaOperaciones.cls:175-176`): con PostgreSQL vinculado, revisar tipos fecha/hora, el cálculo de identificadores en cliente y la ausencia de transacción entre AR, log y replanificación.

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual:** la tarea se cierra solo con evidencia adjunta; la replanificación exige motivo; los bordes avisan pero no impiden guardar.

**Posibles defectos (a verificar):**
- `ARAuditoriaOperaciones.MotivoNoOK` es privada y no se invoca (`ARAuditoriaOperaciones.cls:16`, `:134-146`): coherencia de fechas, evidencia obligatoria al cerrar y tarea duplicada no se comprueban al registrar; sí lo hacen el aviso de evidencia del formulario (`:433-452`) y, para las fechas, solo los bordes (`:151-220`).
- En el alta, `Me.Responsable = m_ObjACAuditoriaActiva.Responsable` (`:71`) sobrescribe el valor que el bloque anterior toma del responsable de implantación de la NC (`:64-70`).
- `HaHabidoCambios` compara fechas con texto vacío (`.FechaInicio <> Nz(Me.FechaInicio, "")`, `:318-329`).
- `ComandoActualizar` rotulado «Contador» (`.form.txt:877`).

**Propuestas web pendientes de decisión:**
- Adjuntar la evidencia en el mismo paso de cierre, sin salir sin guardar.
- Registrar tarea, log y replanificación en una única transacción.

## Criterios de aceptación

1. En el alta se ocultan estado y Actualizar y se deshabilitan Documentos y Actualizar (`:49-51, 72-73`).
2. En el detalle se muestran estado y botones y se cargan todos los campos (`:86-129`).
3. Técnico o NC borrada: no se puede registrar (`:133-135, 387-390`).
4. Sin cambios, no se guarda (`:395-398`).
5. Informar fin real sin evidencia ofrece adjuntarla y no guarda (`:433-452`).
6. Cambiar el fin previsto exige motivo y crea una replanificación (`:454-462`, `ARAuditoriaOperaciones.cls:259-277`).
7. Las reglas de fechas y duplicados de `MotivoNoOK` (`ARAuditoriaOperaciones.cls:43-99`) solo pueden exigirse si la web las aplica; el código actual no las invoca (ver Posibles defectos).
8. Tras registrar se recalculan AC y NC y se cierra (`ARAuditoriaOperaciones.cls:288-316`, `:504`).

## Pendientes

- `constructor.getARAuditoria`, `BuenaFecha`, `ActualizarDatosARAuditoria` y `ReplanificacionesAuditoriaOperaciones.Registrar` no leídos.
- Receptores de los eventos `Alta`, `Editado`, `AltaAnexo` y `EliminadoAnexo` (`:6-9, 487-489`).
