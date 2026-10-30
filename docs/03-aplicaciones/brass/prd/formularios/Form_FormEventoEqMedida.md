# Form_FormEventoEqMedida

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormEventoEqMedida.cls` (282 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ventana auxiliar del **alta** de evento para elegir los equipos de medida que se vincularán al evento aún no guardado. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) y **no escribe en la base de datos**: mantiene la selección en la global `m_ColEquiposMedidaPorEventoActivos` (`Variables Globales.bas:255`) y es el formulario que consume `Form_FormEventoAlta` al guardar. Capacidad: CAP-3 Eventos (equipos de medida del evento). Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`; `VG.bas` = `modules/Variables Globales.bas`. Las citas `:N` sin fichero son de `Form_FormEventoEqMedida.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEventoAlta` | `ComandoEquiposDeMedida_Click` pone `m_ObjEventoActivo = Nothing`, abre el formulario si no estaba abierto y se engancha a sus eventos | `Form_FormEventoAlta.cls:156-167` |

Eventos hacia el llamador: `Alta` (`:9`, lanzado en `:42`) y `Eliminado` (`:10`, lanzado en `:87`); el llamador solo refresca el rótulo del botón con el recuento de la global (`Form_FormEventoAlta.cls:831-840, 856-866`). El alta de evento recorre la global al guardar (`Form_FormEventoAlta.cls:777-780`). `VG.bas:291` la pone a `Nothing` en una rutina de reinicio. Sale con `cmdSalir` (`:12-15`).

## Controles y eventos

Controles: 8 (20 `Begin` en el `.form.txt`: 8 plantillas de tipo por defecto, `Form`, `FormHeader`, `Section` y `FormFooter` más los 8 controles).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComboEquipoDeMedida` | combo, lista de valores de 3 columnas (ID oculto) | `AfterUpdate` → `ComboEquipoDeMedida_AfterUpdate` | Resuelve el equipo, avisa si está fuera de calibración y, si `ComandoAlta` queda habilitado, lo ejecuta | CAP-3 | `:110-163`; `Form_FormEventoEqMedida.form.txt:416-468` |
| `ListaEqMedida` | lista de valores de 3 columnas (ID oculto) | `OnClick` → `ListaEqMedida_Click` | Selecciona un equipo ya añadido y habilita `ComandoEliminar` | CAP-3 | `:236-283`; `Form_FormEventoEqMedida.form.txt:351-388` |
| `ComandoAlta` («+») | botón | `OnClick` | Añade el equipo del combo a la global | CAP-3 | `:17-61`; `Form_FormEventoEqMedida.form.txt:469-506` |
| `ComandoEliminar` («Eliminar»; deshabilitado de diseño) | botón | `OnClick` | Quita de la global el equipo seleccionado en la lista | CAP-3 | `:63-108`; `Form_FormEventoEqMedida.form.txt:507-540` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | CAP-3 | `:12-15`; `Form_FormEventoEqMedida.form.txt:552-561` |
| `lblTitulo1`, `lblTitulo2`, `Etiqueta13` | etiquetas | — | Título (`lblTitulo1` pasa a `Caption` del formulario) y rótulo «EQ. Medida» | CAP-3 | `:190`; `Form_FormEventoEqMedida.form.txt:307, 327, 400` |

## Flujos

### Apertura (`Form_Load`, `:167-203`)

1. `AjustarTamaño Me` (`:177`; `FG.bas:261`) y **vacía** la global (`:178`): cada apertura empieza sin selección.
2. `RellenarListaEquiposMedida` carga el combo con todos los equipos de `m_ObjEntorno.ColEquiposMedida`, con texto «ID;nombre;estado de calibración» (`:181`; `FG.bas:7293-7322`).
3. `EstablecerDatos` deja la lista con su cabecera, ambos botones deshabilitados y la rellena con la global (`:186, 206-228`).
4. Cualquier error: `MsgBox` crítico y cierre del formulario (`:196-203`).

### Alta (`ComboEquipoDeMedida_AfterUpdate` → `ComandoAlta_Click`)

1. Combo vacío: descarta el equipo y deshabilita `ComandoAlta` (`:117-124`). Si no, relee el equipo con `Constructor.getEquipoMedida` (`:125-131`).
2. `ComandoAlta` se habilita salvo que haya evento activo facturado o franqueado (`:132-144`); en este flujo el llamador lo deja a `Nothing` (`Form_FormEventoAlta.cls:160`), luego se habilita siempre.
3. Si `Calibrado = No`: `MsgBox` «Equipo fuera de calibración» (`:148-150`; estado en `EquipoMedida.cls:68-95, 109-124`). Después llama a `ComandoAlta_Click` (`:151-154`).
4. `ComandoAlta_Click` relee el equipo (`:28`), crea el diccionario si hace falta (`:36-39`) y, si el ID no estaba, lo añade y lanza `Alta` (`:40-43`). Si ya estaba, no hace nada y sin aviso. Vacía el combo y repinta la lista (`:44-46`).

### Eliminación (`ComandoEliminar_Click`, `:63-108`)

Relee el equipo de la fila (`:73`), lo quita de la global y, si queda vacía, la deja en `Nothing` (`:81-86`); lanza `Eliminado` (`:87`), repinta la lista y deshabilita el botón (`:90-95`). No pide confirmación.

### Errores

Patrón común: código 1000 = mensaje ya cargado en `m_Error`; otro error se compone con número y descripción; `MsgBox` crítico y el formulario sigue abierto (`:55-61, 102-108, 157-163, 277-283`).

## Reglas, validaciones y cálculos

- **Equipo obligatorio:** sin equipo, «Primero seleccione un equipo de medida» (`:32-35, 77-80`).
- **Sin duplicados:** la clave del diccionario es `IDEquipoMedida` con `CompareMode = TextCompare` (`:38, 40`).
- **Equipo fuera de calibración:** solo avisa, no bloquea (`:148-150`).
- **Bloqueo por evento facturado o franqueado:** código presente (`:138, 264`), pero inalcanzable en el flujo actual por el `Nothing` del llamador (`Form_FormEventoAlta.cls:160`).
- **Estado mostrado:** `EstadoCalibracion` toma «Fuera de servicio», «Sin Calibraciones», «Calibrado» o «Fuera de Calibración» (`EquipoMedida.cls:68-95`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura de un equipo por ID | `TbEquiposMedida` | SQL dinámico concatenado (`WHERE IDEquipoMedida=<id>`) | no | `Constructor.bas:4878-4896` |
| Catálogo para el combo | caché `m_ObjEntorno.ColEquiposMedida` | no se ve SQL en esta ficha | — | `FG.bas:7293-7322` |
| Escrituras | ninguna | solo memoria (global) | — | `:40-43, 83` |

## Coexistencia

El `ID` se concatena sin comillas en el SQL (`Constructor.bas:4894`): con PostgreSQL vinculado es válido mientras la clave sea numérica; si pasara a texto o `uuid`, habría que entrecomillar. El estado de calibración se calcula con `CDate` y `Date` en VBA (`EquipoMedida.cls:90`), sin dialecto de base de datos.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComandoAlta_Click` declara un `m_EquipoEnCombo` local (`:19`) que sombrea al de módulo (`:6`); el de módulo solo lo actualiza el `AfterUpdate` del combo.
- Si el equipo no se resuelve, `:132-133` deshabilita el botón pero `:148` desreferencia igualmente `m_EquipoEnCombo`: error 91 mostrado en `MsgBox` (`:157-163`).
- Con `ComandoAlta` y duplicado, no se avisa de que el equipo ya estaba (`:40-43`).
- `ListaEqMedida_Click` usa `Me.ListaEqMedida.Column(0)` sin comprobar fila (`:252-255`).
- Variable `pregunta` sin declarar localmente (`:60, 107, 149`); depende de una global (a verificar).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: guardar la selección en el estado del asistente de alta en lugar de una global, y avisar del duplicado.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, la lista de equipos añadidos está vacía y `ComandoEliminar` deshabilitado (`ListaEqMedida`, `:178, 216`).
2. Elegir un equipo lo añade a la lista sin pulsar `+` y vacía el combo (`ComboEquipoDeMedida`, `:44, 151-154`).
3. Un equipo fuera de calibración muestra el aviso y se añade igualmente (`:148-154`).
4. Elegir dos veces el mismo equipo deja una sola fila (`:40`).
5. Seleccionar una fila habilita «Eliminar» y eliminarla la quita sin confirmación (`:81-95`).
6. El botón del alta de evento refleja el número de equipos tras cada alta y baja (`Form_FormEventoAlta.cls:837-840, 863-866`).

## Pendientes

- Contenido de `NombreParaCombo` y de `UltimaCalibracion` (`EquipoMedida.cls:25, 68-95`).
- Quién consume la global tras guardar el alta (`Form_FormEventoAlta.cls:777-780`) y qué escribe en base de datos.
- Si `pregunta` es global (`:60`; no localizada en `VG.bas`).
