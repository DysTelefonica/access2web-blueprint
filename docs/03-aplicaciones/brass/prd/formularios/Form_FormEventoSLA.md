# Form_FormEventoSLA

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormEventoSLA.cls` (450 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Edición de los campos SLA de un evento de `TbEventos`: tiempo de respuesta (TRES), tiempo de reparación (TRCM), restablecimiento de servicio (TRSS) y tipo de reparación in situ (SLA-4). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): lee el registro con DAO al cargar y lo reescribe entero al guardar. El evento llega por `OpenArgs`. Las citas son `Form_FormEventoSLA.cls:línea` salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEventoEdicion` | `ComandoSLU_Click` cierra `FormEventoSLA` si está abierto y lo abre con `OpenArgs` = ID del evento | `Form_FormEventoEdicion.cls:276-298` |
| `Form_FormInformeSLA` | `cmdDetalleSLA_Click` lo abre con `acFormReadOnly` y `OpenArgs` = ID de la fila elegida (no la cabecera) | `Form_FormInformeSLA.cls:266-298` |

Sin `OpenArgs` el título es «Datos SLA - Sin evento» y no se carga nada (`:53-55, 71-73`). No lanza eventos ni devuelve valores; se cierra con `cmdCerrar` (`:265-268`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `FechaRecepcionNotificacion`, `FechaInicioContactoCliente` | texto con máscara de fecha y hora | — | TRES; se rellenan con formato «dd/mm/yyyy hh:nn» | CAP-3 | `:120-125`, `.form.txt:665, 712` |
| `cmbIncidenciaAveriaOReparacion` | combo «Sí;No», limitado a la lista | — | TRCM; `Sí`/`No`/vacío ⇒ `True`/`False`/`Null` | CAP-3 | `:128-142, 315-321`, `.form.txt:787, 807` |
| `FechaInicioTiempoAdquisicion`, `FechaFinTiempoAdquisicion` | texto, formato «General Date» | — | Inicio y fin de reparación | CAP-3 | `:144-150`, `.form.txt:852, 899` |
| `cmbTipoReparacion` | combo «Contratista;Fabricante», limitado a la lista | — | Tipo de reparación | CAP-3 | `:152-154`, `.form.txt:934, 954` |
| `chkUrgente` | casilla | — | Urgente | CAP-3 | `:156-158`, `.form.txt:991` |
| `cmbEventoConServicioAfectado` | combo «Sí;No», limitado a la lista | — | TRSS; igual que la incidencia | CAP-3 | `:161-173, 328-334`, `.form.txt:1060, 1080` |
| `FechaRestablecimientoServicio` | texto, formato «General Date» | — | Fecha de restablecimiento | CAP-3 | `:175-177`, `.form.txt:1125` |
| `chkTipoRepInsitu`, `chkTipoRepNoSMT`, `chkTipoRepValvulas` | casillas | — | SLA-4 | CAP-3 | `:180-188`, `.form.txt:1184, 1217, 1250` |
| `cmdGuardar` («Guardar») | botón | `OnClick` → `cmdGuardar_Click` | Valida y guarda | CAP-3 | `:213-259`, `.form.txt:1280-1282` |
| `cmdCerrar` («Cerrar») | botón | `OnClick` → `cmdCerrar_Click` | Cierra sin guardar | CAP-3 | `:265-268`, `.form.txt:1317-1319` |
| `lblTitulo1` | etiqueta | — | «Datos SLA - Evento <id>», con «(Solo lectura)» si está franqueado; pasa a `Caption` | CAP-3 | `:66-75`, `.form.txt:570` |
| 16 etiquetas, 3 rectángulos | etiquetas (`lblTRES`, `lblTRCM`, `lblTRSS`, `lblSLA4` y 12 de campo) y `Cuadro57-59` | — | Rótulos y recuadros de las cuatro zonas | sin asignar | `.form.txt:595-1260` |

Controles: 34 (46 Begin en el .form.txt: 1 `Form`, 8 plantillas de estilo por defecto, `FormHeader`, `Section` y `FormFooter`, y los 34 controles anteriores: 12 de entrada, 2 botones, 17 etiquetas incluida `lblTitulo1`, 3 rectángulos). El único evento del formulario es `OnLoad` (`.form.txt:219`).

## Flujos

### Carga (`Form_Load`, `:42-92`)

1. Lee `OpenArgs`; con ID, consulta `EstaFranqueado("EV", id)` y, si devuelve «Sí», pone `AllowEdits = False` y deshabilita `cmdGuardar` (`:53-61`; `FG.bas:6080-6090`).
2. Guarda el ID en `m_IDEvento`, rotula el título y llama a `CargarDatosEvento` (`:64-70`).
3. `CargarDatosEvento` limpia los controles (`:18-36, 110`), abre `TbEventos` por ID y vuelca cada campo si no es nulo/fecha válida (`:112-189`). `False` en incidencia o servicio afectado se muestra «No» solo si hay datos asociados; si no, queda vacío (`:132-140, 165-171`).
4. Error: mensaje crítico salvo código 1000 y fin del reloj de arena (`:84-91`).

### Guardado (`cmdGuardar_Click`, `:213-259`)

1. `SLAValidator.ValidarCamposSLA` con `p_EnFranqueo:=False`: solo reglas generales (`:224-236`; `SLAValidator.bas:67-109`). Si devuelve código, se muestra el texto acumulado (`:237-239, 250-258`).
2. `GuardarSLA` (`:274-449`): revalida franqueo y exige evento, convierte «Sí»/«No» a booleano, abre el registro con `Edit`, escribe los 12 campos (fechas inválidas o vacías ⇒ `Null`) y `FechaRegistroModificacion = Now()` (`:294-309, 315-339, 342-431`).
3. Éxito: «Datos SLA guardados correctamente.» (`:244`). El formulario sigue abierto.

## Reglas, validaciones y cálculos

- **SLA-001:** fecha de contacto sin recepción; fin de reparación sin inicio; restablecimiento sin «Sí» en servicio afectado (`SLAValidator.bas:69-85`).
- **SLA-007:** contacto anterior a recepción; fin anterior a inicio de reparación; restablecimiento anterior a recepción (`SLAValidator.bas:87-109`).
- **Franqueo (SLA-002 a 005)** —fechas y tipo obligatorios si hay incidencia, restablecimiento si afecta al servicio— solo con `p_EnFranqueo = True` (`SLAValidator.bas:113-141`); este formulario pasa `False`, así que no se aplican.
- **Evento franqueado:** solo lectura en la carga y rechazo al guardar (`:58-61, 296-301`); `EstaFranqueado` devuelve «Sí» si `Franqueado` es «Verdadero» o «True» (`FG.bas:6087-6089`).
- Un fallo inesperado del validador devuelve `SLA-999` (`SLAValidator.bas:153-159`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Cargar evento | `TbEventos` | SQL dinámico concatenado (`SELECT *` por `IDEvento` entre comillas) + `OpenRecordset` | — | `:112-115` |
| Comprobar franqueo | `TbEventos` | `Dame` con SQL concatenado, primero sin comillas y luego con ellas | — | `FG.bas:290-330` |
| Guardar | `TbEventos` | SQL dinámico concatenado + DAO `Edit`/`Update` | no | `:342-431` |

Campos escritos: `FechaRecepcionNotificacion`, `FechaInicioContactoCliente`, `IncidenciaAveriaOReparacion`, `FechaInicioTiempoAdquisicion`, `FechaFinTiempoAdquisicion`, `TipoReparacion`, `Urgente`, `EventoConServicioAfectado`, `FechaRestablecimientoServicio`, `TipoRepInsitu`, `TipoRepNoSMT`, `TipoRepValvulas` y `FechaRegistroModificacion` (`:355-430`).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL concatenado con el ID entre comillas simples; booleanos y `Null` en columnas `Yes/No` de Jet/ACE (`:369-427`) que en PostgreSQL serán `boolean` anulable; fechas convertidas con `Format`/`CDate`/`IsDate`, dependientes de la configuración regional (`:121, 124, 355-409`); `Franqueado` leído como texto «Verdadero»/«True» (`FG.bas:6087`); `Now()` evaluado en el cliente (`:430`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Los mensajes de error de `GuardarSLA` se pierden: relanza el error 1000 (`:446-448`) y `cmdGuardar_Click` solo muestra su `strTextoError` local, vacío en ese caso (`:250-258`). «No hay evento seleccionado», «No se encontró el evento» y «…evento franqueado» no se verían.
- Fallo abierto del control de franqueo: si `EstaFranqueado` devuelve error («#ERR|…»), la carga no bloquea (`:57-61`) y `GuardarSLA` tampoco (`:297-302`).
- Abierto desde `Form_FormInformeSLA` con `acFormReadOnly`, el formulario no deshabilita `cmdGuardar` si el evento no está franqueado y `GuardarSLA` escribe por DAO (`Form_FormInformeSLA.cls:298`; `:58-61, 345-431`); `AllowEdits = False` sobre controles no enlazados podría no bloquearlos.
- Una fecha mal escrita en los campos con «General Date» se guarda como `Null` y borra el valor previo sin avisar (`:374-384, 405-409`).
- La carga de fechas y su reescritura dependen del formato regional del equipo (`:120-125, 355-365`); además, recepción y contacto se cargan sin segundos y se guardan sin ellos (`:121, 124`).
- Un `False` sin datos asociados se muestra vacío y al guardar pasa a `Null` (`:132-140, 319-321, 368-372`); una casilla sin tocar se guarda como `Null`, no como `False` (`:326, 337-339`).
- `FechaRegistroModificacion` se actualiza en cada guardado aunque no cambie nada (`:430`).
- El comentario de `Evento.ActualizarSLAApartirDeFormulario` dice que lo usa este formulario (`Evento.cls:2829`), pero el formulario escribe directamente por DAO y la búsqueda no halla otros llamadores; lógica duplicada.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: mostrar los errores de guardado; bloquear la edición cuando se abre en solo lectura; decidir si el guardado normal debe aplicar el validador de franqueo; no guardar `Null` donde antes había `False`.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con un evento no franqueado se cargan los 12 campos; con uno franqueado el título indica «(Solo lectura)» y «Guardar» está deshabilitado (`:58-69`).
2. Fecha de contacto sin recepción, o anterior a ella, no se guarda y muestra el motivo (`SLAValidator.bas:69-73, 87-93`).
3. Fin de reparación sin inicio, o anterior a él, no se guarda (`SLAValidator.bas:75-79, 95-101`).
4. Fecha de restablecimiento sin «Sí» en servicio afectado, o anterior a la recepción, no se guarda (`SLAValidator.bas:81-85, 103-109`).
5. Guardar con datos válidos actualiza `TbEventos` y muestra «Datos SLA guardados correctamente.» (`:244, 355-430`).
6. Los campos vaciados se guardan como `Null` (`:358, 364, 371, 377, 383, 389`).
7. `cmdCerrar` cierra sin guardar (`:265-268`).

## Pendientes

- Tipo de columna y valores permitidos de `TbEventos.Franqueado` y de los campos SLA (ERD de origen).
- Si `AllowEdits = False` bloquea los controles no enlazados (`:59`).
- Contenido de `FormularioAbierto` (`Form_FormEventoEdicion.cls:285`).
