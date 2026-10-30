# Form_FormEventoFranqueo

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormEventoFranqueo.cls` (102 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Diálogo para **franquear** un evento: pide la causa de fin y delega en `Evento.Franquear`, que fija fecha y hora finales y marca el evento como franqueado. Es un formulario **no enlazado** (sin `RecordSource`). Capacidad: CAP-3 Eventos (franqueo). Las citas `:N` sin fichero son de `Form_FormEventoFranqueo.cls`. Lo abre `Form_FormEventoGestion` con `m_ObjEventoActivo` ya asignado.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEventoGestion` (botón con rótulo «Franquear») | Rechaza si ya está franqueado, cierra el diálogo si estaba abierto, asigna `m_ObjEventoActivo` y lo abre | `Form_FormEventoGestion.cls:1983-1990` |

Evento hacia el llamador: `Franqueado` (`:8`, lanzado en `:25`); el llamador reajusta `strSeleccionadoFranqueado` y repinta la lista (`Form_FormEventoGestion.cls:2869-2879`). Tras franquear se cierra solo (`:26`); `cmdSalir` cierra sin guardar (`:100-103`). El rótulo «Quitar Franqueo» del otro botón no abre este diálogo: lo gestiona el llamador (`Form_FormEventoGestion.cls:1991-2010`).

## Controles y eventos

Controles: 9 (18 `Begin` en el `.form.txt`: `Form`, `Section`, 7 plantillas de tipo por defecto y los 9 controles; sin cabecera ni pie).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `CausaFin` | cuadro de texto no enlazado, `Tag` «OBLIGATORIO» | — | Causa de fin; obligatoria al franquear; se precarga calculada o la ya guardada | CAP-3 | `:14, 67, 79`; `Form_FormEventoFranqueo.form.txt:276-303` |
| `ComandoFranquear` («Franquear») | botón | `OnClick` → `ComandoFranquear_Click` | Franquea el evento y cierra | CAP-3 | `:10-38`; `Form_FormEventoFranqueo.form.txt:328-365` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | CAP-3 | `:100-103`; `Form_FormEventoFranqueo.form.txt:366-404` |
| `Etiqueta214`, `Etiqueta3` | etiquetas | — | Rótulo «CAUSA FIN» y asterisco de obligatorio | CAP-3 | `Form_FormEventoFranqueo.form.txt:405-431, 499-510` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título «Franqueo de Evento»; `lblTitulo1` pasa a `Caption` del formulario | CAP-3 | `:84`; `Form_FormEventoFranqueo.form.txt:454-498` |
| `Cuadro2`, `Cuadro3` | rectángulos | — | Marcos decorativos | CAP-3 | `Form_FormEventoFranqueo.form.txt:304-327, 432-453` |

## Flujos

### Apertura (`Form_Load`, `:40-98`)

1. `AjustarTamaño Me` (`:48`). Sin evento activo: «No hay un evento activo» (`:51-54`).
2. **Relee** el evento con `Constructor.getEvento(m_ObjEventoActivo.IDEVENTO)` y reasigna la global (`:55-62`).
3. Si está franqueado: deshabilita el botón, `AllowEdits = False`, muestra la `CausaFin` guardada y rotula «Quitar Franqueo» (`:63-69`).
4. Si no: habilita el botón, `AllowEdits = True`, precarga `CausaFinCalculada` si no está vacía y rotula «Franquear» (`:70-82`; `Evento.cls:267-300`).
5. Cualquier error: `MsgBox` crítico y cierre del formulario (`:91-98`).

### Franquear (`ComandoFranquear_Click`, `:10-38`)

1. `CausaFin` vacía: «Ha de indicar alguna causa del franqueo» (`:14-17`).
2. `m_ObjEventoActivo.Franquear Me.CausaFin` (`:21`; `Evento.cls:1255-1310`); lanza `Franqueado` y cierra (`:25-26`).
3. Error: `MsgBox` crítico y el diálogo sigue abierto (`:32-38`).

## Reglas, validaciones y cálculos

- **Causa obligatoria:** en el formulario (`:14`) y de nuevo en la clase, «No se ha indicado la causa fin de franqueo» (`Evento.cls:1288-1291`).
- **Motivos de no franqueable:** `MotivosNoFranqueable` y la validación SLA con `ValidarSLA(p_EnFranqueo:=True)` bloquean con su mensaje (`Evento.cls:1269-1286, 1172`).
- **Fecha final:** `FechaFinalCalculada` (`Evento.cls:1292`); **hora final** fija 15:00 (`Evento.cls:1297`).
- **Efecto:** `UPDATE` de `FechaFinal`, `HORAFINALEVENTO`, `CAUSAFIN` y `Franqueado = True`, y copia el valor a la instancia (`Evento.cls:1298-1307`).
- **Causa precalculada:** `CausaFinCalculada` junta descripciones de actividades (`Evento.cls:267-300`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Relectura del evento | `TbEventos` | construcción por `Constructor.getEvento` | no | `:55`; `Constructor.bas:2242` |
| Franqueo | `TbEventos` | SQL dinámico `UPDATE` concatenado con `Execute` | no | `Evento.cls:1298-1302` |
| Cálculos previos (SLA, fecha final, actividades) | varias | lecturas de la clase | — | `Evento.cls:1269-1296` |

## Coexistencia

Adaptar: las fechas se incrustan con literales `#mm/dd/yyyy#` y la hora `15:0:0` (`Evento.cls:1298-1301`), sintaxis Jet/ACE que PostgreSQL no admite; `Franqueado = True` como literal booleano y `CAUSAFIN` entre comillas simples concatenadas; `IDEvento` tratado como texto (`Evento.cls:1301`). Se requiere SQL parametrizado y sentencias coherentes con tipos fecha/hora vinculados; el `UPDATE` único evita transacción, pero conviene conservarlo atómico.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `CAUSAFIN` se concatena entre comillas (`Evento.cls:1300`): una causa con apóstrofo rompería la sentencia.
- La hora final queda fija a las 15:00 y se escribe también en `HORAFINALEVENTO` con el formato de fecha (`Evento.cls:1297-1300`).
- No hay confirmación antes de franquear (`:10-38`); el cierre del diálogo ocurre antes de restablecer el cursor (`:26-28`).
- `m_ObjEventoAlInicio` está declarado y no se usa (`:6`).
- `Form_Load` reasigna la global `m_ObjEventoActivo` con la relectura (`:55`), lo que afecta al llamador.
- Con `AllowEdits = False` en eventos ya franqueados, `CausaFin` no enlazado podría no quedar bloqueado (`:65`; a verificar).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: pedir confirmación, parametrizar el `UPDATE` y mostrar de antemano los motivos que impiden franquear.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin evento activo, el diálogo muestra «No hay un evento activo» y se cierra (`:51-54`).
2. En un evento sin franquear, `CausaFin` aparece precargada con la causa calculada si existe (`:73-80`).
3. Franquear con `CausaFin` vacía muestra «Ha de indicar alguna causa del franqueo» y no escribe (`:14-17`).
4. Un evento con motivos que impiden franquear o con incumplimiento SLA muestra el motivo y no cambia (`Evento.cls:1269-1286`).
5. Franquear con éxito marca el evento, cierra el diálogo y el listado de eventos se repinta (`:25-26`; `Form_FormEventoGestion.cls:2869-2879`).
6. En un evento ya franqueado el botón queda deshabilitado, rotulado «Quitar Franqueo», con la causa guardada (`:63-69`).

## Pendientes

- Contenido de `MotivosNoFranqueable`, `ValidarSLA` y `FechaFinalCalculada` (`Evento.cls:1172, 1269-1293`).
- Si `pregunta` es global (`:37`; no localizada en `Variables Globales.bas`).
- Alcance de `AllowEdits = False` sobre un control no enlazado (`:65`).
