# Form_FormMaterialIntervencionEdicion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormMaterialIntervencionEdicion.cls` (735 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Edición (o consulta, si está bloqueada) de una intervención de seguimiento de material, es decir, un registro de `TbMaterialSeguimiento`: fechas de inicio y fin, motivo, destino, resultado y marca de última intervención. Es un formulario **no enlazado** (sin `RecordSource`/`ControlSource` en el `.form.txt`): carga los datos con recordsets DAO sobre SQL concatenado y graba con `IntervencionEdicion`, función pública del propio formulario (`Form_FormMaterialIntervencionEdicion.cls:461`). Citas `Form_FormMaterialIntervencionEdicion.cls:línea` salvo indicación; se abrevian `:N` en el propio `.cls`, `FG.bas` = `modules/Funciones Generales.bas` y `VG.bas` = `modules/Variables Globales.bas`. Controles: 32 (40 Begin en el `.form.txt`: 32 controles + `Form` + `Section` + 6 plantillas de estilo por defecto, `.form.txt:5, 145-253`).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormMaterialSeguimiento` | Cierra la instancia abierta si existe y abre el formulario pasando el `IDSeguimiento` en `OpenArgs` | `Form_FormMaterialSeguimiento.cls:152-155` |

Al guardar con éxito, si `FormMaterialSeguimiento` está abierto llama a su `EstablecerDatos` (`:162-164`). Al salir guardando, si `FormEventoGestion` está abierto refresca su lista `ListaActividades` (`:37-39`). `ComandoVerEvento` abre `FormEventoGestion` posicionado en el evento (`:230-241`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `FechaInicioIntervencion` | texto, formato fecha corta | — | Inicio de la intervención; obligatorio al guardar | CAP-7 | `.form.txt:343-344`; `:509-512` |
| `FechaFinIntervencion` | texto, formato fecha corta | — | Fin de la intervención; opcional | CAP-7 | `.form.txt:694-695`; `:524-535` |
| `MOTIVO` | texto | — | Motivo; obligatorio | CAP-7 | `.form.txt:750`; `:520-523` |
| `RESULTADO` | texto | — | Resultado; exige fecha fin y viceversa | CAP-7 | `.form.txt:869`; `:530-541` |
| `DESTINO` | combo `Table/Query` | — | Destinos distintos ya usados en `TbMaterialSeguimiento`; obligatorio | CAP-7 | `.form.txt:838-842`; `:516-519` |
| `ComboUltimaIntervencion` | combo lista de valores `Sí;No`, valor por defecto `'No'` | — | Marca de intervención final; «Sí» exige fecha fin | CAP-7 | `.form.txt:958-962`; `:542-547` |
| `IDEVENTO`, `EQUIPO`, `Material`, `Garantia` | texto, etiqueta `FIJOS` (`Locked`) | — | Datos de contexto de solo lectura; `Garantia` no se rellena en el código | CAP-7 | `.form.txt:527-528, 558-559, 616-617, 1017-1018`; `:396-402` |
| `ComandoRegistrar` | botón «Registrar» | `OnClick` → `ComandoRegistrar_Click` | Guarda y cierra | CAP-7 | `:176-213`; `.form.txt:368-370` |
| `cmdSalir` | botón «Salir» | `OnClick` → `cmdSalir_Click` | Si hay cambios y es editable, pregunta si guardar; cierra | CAP-7 | `:7-53`; `.form.txt:406-408` |
| `ComandoVerEvento` | botón «Ver Evento» (inicia deshabilitado) | `OnClick` → `ComandoVerEvento_Click` | Abre la gestión del evento origen | CAP-3 | `:215-252`; `.form.txt:1141, 1150-1152` |
| `lblFranqueado`, `lblFacturado` | etiquetas (inician ocultas) | — | Avisos «FRANQUEADO» / «Facturado» | CAP-7 | `.form.txt:1078, 1109`; `:407-417` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título; `lblTitulo1` cambia entre «Edición…» y «Detalle…» y fija `Me.Caption` | CAP-7 | `:425, 437, 448` |
| `Etiqueta83`, `84`, `91`, `94`, `96`, `98`, `102`, `107`, `108`, `110` | 10 etiquetas | — | Rótulos de campo | CAP-7 | `.form.txt:459-1000` |
| `CuadroPrincipal`, `CuadroSuperior`, `Cuadro2`, `Marco77`, `Marco93` | 3 rectángulos y 2 grupos de opciones | — | Decoración y agrupación visual | CAP-7 | `.form.txt:264-330, 437-457, 633-653` |

`Form_Open` (`OnOpen`, `.form.txt:42`) es el único evento de formulario.

## Flujos

### Apertura (`Form_Open`, `:254-291`)

1. `AjustarTamaño Me` (`:260`; definido en `FG.bas:261`) y lectura de `OpenArgs` en `strIDIntervencion`; vacío ⇒ «Se ha abierto con parámetros inadecuados» (`:261-265`).
2. `EstablecerDatos` (`:292-460`) lee la intervención con su material y la actividad (`:301-307`), guarda los valores iniciales para detectar cambios (`:313-319`) y los vuelca a los controles si no están vacíos (`:336-353`).
3. Si el material viene de una actividad (`TbMaterial.IDActividad` no vacío) lee equipo, evento y material por `TbActividades`/`TbEventos`/`TbEquipos` (`:361-378`); si no, por `TbMaterial.IDEquipo` (`:379-394`). Con evento habilita `ComandoVerEvento` (`:401-406`).
4. Calcula `strEscritura`: «No» si `strFranqueado = "Sí"` o `strEstaFacturado = "Sí"` (`:418-422`). En «No» deshabilita `ComandoRegistrar`, rotula «Detalle de Intervención de Material» y bloquea combos y textos no `FIJOS` con `lngColorCampoNoEditable` (`:423-434`; color en `VG.bas:351`); en «Sí» los deja editables con `lngColorCampoEditable` (`:435-447`; `VG.bas:350`).
5. Cualquier error: `MsgBox` crítico y cierre del formulario (`:284-290`).

### Guardado (`ComandoRegistrar_Click`, `:176-213`)

1. Sin cambios ⇒ «No se ha registrado ningún cambio» (`:187-190`).
2. `Registrar` (`:119-172`) exige `strIDIntervencion`, convierte vacíos en `"#"` y llama a `IntervencionEdicion` (`:135-151`); si `FormMaterialSeguimiento` está abierto, lo refresca (`:162-164`).
3. Éxito ⇒ cierra el formulario (`:205`). Error ⇒ `MsgBox` crítico y el formulario sigue abierto (`:207-212`).

### Salida (`cmdSalir_Click`, `:7-53`)

Con cambios y `strEscritura = "Sí"` pregunta «¿Desea guardar los cambios producidos?»; «Sí» ejecuta `Registrar` y refresca `ListaActividades` de `FormEventoGestion` (`:16-39`). Después cierra sin guardar el diseño (`:45`). Un error de `Registrar` muestra el mensaje y no cierra (`:28-32, 47-52`).

## Reglas, validaciones y cálculos

Todas en `IntervencionEdicion` (`:505-565`):

- Inicio obligatorio y válido como fecha (`:509-512`); `Destino` y `Motivo` obligatorios (`:516-523`).
- Fin no anterior al inicio (`:524-529`); con fecha fin, `Resultado` obligatorio (`:530-535`); con `Resultado`, fecha fin obligatoria (`:536-541`).
- `UltimaIntervencion = "Sí"` exige fecha fin (`:542-547`); cualquier valor distinto de «Sí»/«No» se trata como «No» (`:513-515`).
- Si no llega `strIDMaterial`, se obtiene con `Dame` sobre `IDSeguimiento` (`:548-561`; `FG.bas:290`).
- Vaciar fecha fin o resultado: el marcador `"#"` graba `Null` (`:669-682`).
- **Cierre del seguimiento:** tras grabar, `SeguimientoCerrado` (`FG.bas:3872-3919`) devuelve «Sí» si algún registro del material tiene «Sí»; entonces `TbMaterial.FechaEntrega` se actualiza con `UltimaFechaFinSeguimiento` (`:693-718`; `FG.bas:3930-3990`).
- **Solo lectura:** franqueado (valor de `FechaInicioIntervencion` comparado con `True`, `:324-328`) o facturado bloquea la edición (`:418-422`).
- **Detección de cambios:** compara con los valores iniciales y trata cualquier error como cambio (`:54-118`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura de intervención + material + factura | `TbMaterialSeguimiento`, `TbMaterial`, `TbFacturaMaterialesInvolucrados` | SQL dinámico concatenado (`IDSeguimiento` entre comillas), `OpenRecordset` | — | `:301-307` |
| Equipo/evento/material de una actividad | `TbMaterial`, `TbActividades`, `TbEventos`, `TbEquipos` | SQL dinámico concatenado | — | `:362-367` |
| Equipo/material sin actividad | `TbMaterial`, `TbEquipos` | SQL dinámico concatenado | — | `:380-383` |
| Lista de destinos | `TbMaterialSeguimiento` | `RowSource` guardado en el formulario (`SELECT DISTINCT`) | — | `.form.txt:838-842` |
| Material de la intervención | `TbMaterialSeguimiento` | `Dame`: SQL concatenado con reintento entre comillas | — | `FG.bas:290-310` |
| Edición | `TbMaterialSeguimiento` | DAO `Edit`/`Update` | no | `:656-685` |
| Estado de cierre / última fecha | `TbMaterialSeguimiento`, `TbMaterial` | SQL concatenado | — | `FG.bas:3890-3891, 3966-3967` |
| Fecha de entrega del material | `TbMaterial` | `UPDATE` concatenado con fecha `#mm/dd/yyyy#`, `DoCmd.RunSQL` con avisos desactivados | no | `:713-717` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: literales de fecha `#mm/dd/yyyy#` y `Format` (`:713`); identificadores entre comillas simples concatenadas; `DoCmd.RunSQL` y `SetWarnings`; la edición y la actualización de `FechaEntrega` no comparten transacción; comparación del campo con `True` (`:324`) y valores «Sí»/«No» como texto.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `strEstaFacturado` solo se lee (`:413, 418`) y nunca se asigna; `lblFacturado` no llega a mostrarse y la facturación no bloquea la edición. La consulta trae `IDFactura` (`:301`) sin usarlo.
- El franqueo se decide comprobando `.Fields("FechaInicioIntervencion") = True` (`:324`), comparando una fecha con `True`; parece una columna equivocada.
- `HaHabidoCambios` asigna la fecha fin a `strFechaInicioIntervencion` (`:81-83`): un cambio solo en la fecha de inicio puede no detectarse y se compara la fecha fin con la de inicio.
- No hay validación de solapamiento entre intervenciones ni de inicio frente a `TbMaterial`: el bloque que las hacía está comentado (`:566-655`), aunque la cabecera lo describe (`:476-491`).
- `UltimaFechaFinSeguimiento` lee `TbMaterial.FechaIncio` (`FG.bas:3966, 3975`) mientras los comentarios hablan de `FechaInicio` (`:478`): a verificar el nombre real de la columna.
- Las fechas se convierten con `Format(..., "mm/dd/yyyy")` para el `UPDATE` (`:713`) y se graban como texto sobre campos fecha (`:666, 670`); dependiente de la configuración regional.
- Con `IDSeguimiento` concatenado entre comillas (`:306, 658`), un valor con apóstrofo rompería la consulta.
- `Dame` sobre `On Error Resume Next` (`FG.bas:299-300`) puede dejar un recordset sin abrir si fallan las dos consultas.
- `Garantia` es un campo de solo lectura que ningún código rellena (`.form.txt:1017`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: validar el orden y solapamiento entre intervenciones en servidor; ejecutar la edición y la actualización de `FechaEntrega` en una única transacción; derivar «franqueado» y «facturado» de datos explícitos.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin `OpenArgs` o con un identificador inexistente, se muestra «Se ha abierto con parámetros inadecuados» y se cierra (`:261-265, 309-312`).
2. Una intervención de actividad franqueada abre en modo detalle, sin botón Registrar y con campos bloqueados (`:418-434`).
3. Guardar sin inicio, sin destino o sin motivo muestra el error correspondiente y no escribe (`:509-523`).
4. Guardar con fin anterior al inicio, con fin sin resultado o con resultado sin fin muestra el error y no escribe (`:524-541`).
5. Marcar «Sí» en última intervención sin fecha fin se rechaza (`:542-547`).
6. Vaciar fecha fin o resultado deja el campo a `Null` (`:669-682`).
7. Si tras guardar el seguimiento queda cerrado, `TbMaterial.FechaEntrega` pasa a la última fecha fin (`:693-718`).
8. Registrar sin cambios muestra «No se ha registrado ningún cambio» (`:187-190`).
9. «Ver Evento» abre la gestión del evento origen; deshabilitado si no hay evento (`:215-241, 401-406`).

## Pendientes

- Contenido de `AjustarTamaño` más allá de su firma (`FG.bas:261`) y de `FormEventoGestion.ComandoActualizarTodo_Click` y `ComboIDEventoBusqueda_AfterUpdate`, invocados en `:233, 237, 239`.
- Tipo real de las columnas fecha de `TbMaterialSeguimiento` y de `FechaEntrega` (ERD de origen).
- Dónde se rellena `TbFacturaMaterialesInvolucrados` y si debería informar `strEstaFacturado` (`:305`, `:413`).
- Cómo se determina el franqueo en el resto de la aplicación (`:324`).
