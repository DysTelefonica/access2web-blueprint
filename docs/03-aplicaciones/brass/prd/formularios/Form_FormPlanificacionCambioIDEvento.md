# Form_FormPlanificacionCambioIDEvento

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormPlanificacionCambioIDEvento.cls` (282 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Cambia el evento de mantenimiento programado que cierra una planificación ya cerrada (`TbPlanificacion`). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): recibe todo por `OpenArgs` y escribe con la función local `CerrarPlanificacionCambioTrabajo` (`Form_FormPlanificacionCambioIDEvento.cls:173-275`). Las citas son `Form_FormPlanificacionCambioIDEvento.cls:línea` (aquí `:línea`) salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`, `VG.bas` = `modules/Variables Globales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormPlanificacionesBusqueda` (`ComandoCambiarEventoDeCierre_Click`) | Cierra el formulario si estaba abierto y lo abre con `OpenArgs` = `IDPlanificacion\|IDEventoLoCierra\|IDEquipo\|FechaPrevistaCierre` (columnas 9 y 2 de `ListaFiltrados`) | `Form_FormPlanificacionesBusqueda.cls:255-271` |

Salida: tras registrar, refresca `ListaFiltrados` de `FormPlanificacionesBusqueda` si está abierto y se cierra sin guardar (`:32-35`). `cmdSalir` cierra sin hacer nada más (`:5-8`). Abre `FormEventoGestion` desde los dos botones «Ver Evento» (`:63-74`, `:101-112`). Comparte el formato de `OpenArgs` con la apertura desde `Form_FormPlanificacionesBusqueda.cls:271`.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `IDEvento` | texto bloqueado | — | Muestra el evento de cierre actual (`strIDEventoCierre`) | CAP-6 Planificación | `:151`, `.form.txt:515-529` |
| `FechaPrevistaCierre` | texto bloqueado | — | Muestra la fecha prevista en `dd/mm/yyyy` | CAP-6 Planificación | `:152`, `.form.txt:445-459` |
| `ComboIDEventoNuevo` | combo, 3 columnas, `LimitToList` | — | Eventos candidatos; su `RowSource` se reasigna en `Form_Load` | CAP-6 Planificación | `:153-158`, `.form.txt:476-497` |
| `ComandoRegistrar` | botón «Registrar» | `OnClick` → `ComandoRegistrar_Click` | Valida y registra el cambio; cierra | CAP-6 Planificación | `:9-47`, `.form.txt:545-555` |
| `ComandoVerEventoCierre` | botón «Ver Evento» | `OnClick` | Abre/posiciona `FormEventoGestion` en el evento actual | CAP-3 Eventos | `:48-85`, `.form.txt:689-700` |
| `ComandoVerNuevoEventoCierre` | botón «Ver Evento», **deshabilitado en diseño** | `OnClick` | Igual, con el evento elegido en el combo | CAP-3 Eventos | `:86-123`, `.form.txt:795-807` |
| `cmdSalir` | botón «Salir» | `OnClick` | Cierra con `acSaveNo` | CAP-6 Planificación | `:5-8`, `.form.txt:418-428` |
| `ComandoAyudaFormulario` | botón «Comando46», deshabilitado | sin handler | Sin comportamiento | sin asignar | `.form.txt:757-766` |
| `lblTitulo1`, `lblTitulo2`, `lbl1`, `Etiqueta10`, `Etiqueta11` | etiquetas (5) | — | Rótulos; `lblTitulo1` da el `Caption` de la ventana | CAP-6 Planificación | `:132`, `.form.txt:571-743` |
| `Marco120`, `Cuadro69`, `CuadroSuperior`, `Cuadro2` | marco y rectángulos (4) | — | Decoración | sin asignar | `.form.txt:349-404`, `638-648` |

Controles: 30 (30 Begin en el .form.txt): 1 `Form`, 11 plantillas de tipo por defecto (`.form.txt:153-328`), 1 `Section` y 17 controles con nombre, que son las 17 unidades de la tabla (8 individuales + 5 etiquetas + 4 marco/rectángulos).

## Flujos

### Apertura (`Form_Load`, `:124-172`)

1. `AjustarTamaño Me` (`FG.bas:261`) y `Caption` desde `lblTitulo1` (`:130-132`).
2. `OpenArgs` debe contener `\|` y exactamente 4 partes (`:134-142`); `IDPlanificacion` y `IDEquipo` numéricos, `IDEventoCierre` no vacío y la fecha válida (`:147`). Si no, «No se ha abierto el formulario con los parámetros adecuados», aviso y cierre (`:165-171`).
3. Rellena `IDEvento` y `FechaPrevistaCierre` (`:151-152`) y limita el combo a eventos `MANT. PROGRAMADO` del mismo equipo que no cierran ya otra planificación (`:153-158`).

### Registro (`ComandoRegistrar_Click`, `:9-47`)

1. Llama a `CerrarPlanificacionCambioTrabajo(strIDPlanificacion, ComboIDEventoNuevo)` (`:15-21`); un retorno con `\|` es un error y se muestra (`:22-26`, `:45`).
2. Refresca `FormPlanificacionesBusqueda` y cierra (`:32-35`). El formulario permanece abierto si hay error (`:40-46`).

## Reglas, validaciones y cálculos

- Planificación numérica y evento nuevo no vacío (`:195-202`).
- El evento nuevo no puede estar ya en `TbPlanificacion.IDEvento` (`:208-217`).
- Su `TIPOEVENTO` debe ser `MANT. PROGRAMADO` (`:223-232`).
- Debe estar cerrado: la «fecha fin» es la `FechaAlta` más reciente de sus actividades (`FG.bas:798-822`); sin actividades, «El evento ha de estar cerrado» (`:238-246`).
- La planificación debe existir (`:253-256`). Se graba `FechaCierre` = esa fecha en `dd/mm/yyyy` e `IDEvento` (`:257-260`).
- Mensaje de éxito «Cambio de evento que cierra la planificación a Fecha …» (`:264`); `ComandoRegistrar_Click` no lo muestra.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Candidatos del combo | `TbEventos` ⟕ `TbPlanificacion` | SQL dinámico concatenado (`IDEquipo`) | — | `:153-157` |
| Evento ya usado / tipo | `TbPlanificacion`, `TbEventos` | `Dame`: SQL dinámico concatenado, reintento con comillas | — | `:208`, `:223`, `FG.bas:290-325` |
| Fecha fin del evento | `TbActividades` | SQL dinámico concatenado | — | `FG.bas:810-822` |
| Lectura de la planificación | `TbPlanificacion` | SQL dinámico concatenado | — | `:248-251` |
| Cambio de cierre | `TbPlanificacion` | DAO `Edit`/`Update` (`FechaCierre`, `IDEvento`) | no | `:257-260` |

## Coexistencia

Adaptar a PostgreSQL vinculado: SQL concatenado con `IDEquipo` e `IDEvento` (`:155`, `FG.bas:812`); `Dame` que reintenta cambiando literal numérico por texto según el error (`FG.bas:304-315`); fecha formateada como cadena `dd/mm/yyyy` en un campo de fecha (`:258`); `LEFT JOIN … Is Null` (`:154-156`); `Edit`/`Update` DAO sin transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComandoVerNuevoEventoCierre` tiene `Enabled` no por defecto (`.form.txt:796`) y ningún código lo activa en este formulario; podría quedar inutilizable.
- No se comprueba que el evento nuevo sea del mismo equipo en la función, solo en el filtro del combo (`:155` frente a `:195-260`); tampoco el estado de la planificación (abierta/cerrada), solo que exista (`:253`).
- `Format(strFechaCierre, "dd/mm/yyyy")` guardado en un campo de fecha depende de la configuración regional (`:152`, `:258`).
- `ComandoVerEventoCierre` y `ComandoVerNuevoEventoCierre` duplican el mismo bloque (`:63-74`, `:101-112`).
- `ComandoAyudaFormulario` rotulado «Comando46», deshabilitado y sin handler (`.form.txt:765-766`).
- Condición `intMes < 1 And intMes > 12` no es nunca cierta; no afecta a este formulario (`FG.bas:623`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: validar equipo y estado de la planificación en el servidor; transacción única; eliminar el botón de ayuda.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin 4 parámetros válidos muestra «No se ha abierto el formulario con los parámetros adecuados» y se cierra.
2. El combo solo ofrece eventos `MANT. PROGRAMADO` del equipo sin planificación asociada.
3. Registrar sin evento muestra «Ha de introducir el evento que va a ser permutado».
4. Registrar un evento ya usado muestra «Ese evento ya ha sido usado para cerrar otra planificación anteriormente.»
5. Registrar un evento sin actividades muestra «El evento ha de estar cerrado» y no escribe.
6. Un registro correcto actualiza `FechaCierre` e `IDEvento`, refresca la búsqueda y cierra.
7. «Ver Evento» abre `FormEventoGestion` posicionado en el evento mostrado.

## Pendientes

- Contenido de `FormEventoGestion.ComboIDEventoBusqueda_AfterUpdate` y `ComandoActualizarTodo_Click` (`:66`, `:70`, `:72`).
- Si `TbPlanificacion.IDEvento` tiene índice único (ERD de origen).
- Nombre con distinta caja `formPlanificacionCambioIDEvento` al abrir (`Form_FormPlanificacionesBusqueda.cls:268-271`).
