# Form_FormEventoCopia

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormEventoCopia.cls` (454 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Diálogo de copia de un evento ya existente: crea un evento nuevo con sus actividades, los materiales de cada actividad, los anexos de cada material y los seguimientos de cada material. Es un formulario **no enlazado** (sin `RecordSource` en `Form_FormEventoCopia.form.txt`): toma el evento de la global `m_ObjEventoActivo` (`Form_FormEventoCopia.cls:57, 85`) y escribe directamente por DAO desde funciones públicas definidas en el propio formulario (`:109`, `:345`, `:409`). Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`; `VG.bas` = `modules/Variables Globales.bas`. Las citas sin fichero son del `.cls` del formulario.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEventoGestion` (`ComandoCopiarEvento_Click`) | Toma el evento seleccionado (o lo carga de `ListaEventos`; si no hay, «Seleccione un evento de la lista»), lo asigna a `m_ObjEventoActivo`, cierra el formulario si ya estaba abierto y lo abre; guarda la referencia `WithEvents` | `Form_FormEventoGestion.cls:1511-1530` |

Salida: evento público `AltaEvento(m_Evento As Evento)` (`:5`), lanzado tras copiar con el evento nuevo (`:35`); el formulario padre refresca la lista, selecciona el evento nuevo y muestra «Se ha generado el evento con ID» (`Form_FormEventoGestion.cls:2844-2858`). Tras copiar se cierra solo (`:36`); `cmdSalir` cierra sin copiar (`:8-11`).

## Controles y eventos

Controles: 19 (19 `Begin` en el `.form.txt`): 1 `Begin Form`, 1 `Begin Section`, 7 plantillas de control por defecto de la cabecera y 10 controles reales.

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `Form` | formulario | `OnOpen` → `Form_Open` | Ajusta tamaño, exige evento activo, rellena controles | CAP-3 | `:49-78`; `Form_FormEventoCopia.form.txt:5, 63` |
| `Detalle` | sección | — | Contenedor de los controles | CAP-3 | `Form_FormEventoCopia.form.txt:277-279` |
| Plantillas por defecto (Label, Rectangle, Image, CommandButton, OptionGroup, TextBox, ComboBox) | 7 plantillas de estilo | — | Valores por defecto del diseño, sin nombre ni lógica | sin asignar | `Form_FormEventoCopia.form.txt:140-256` |
| `CuadroSuperior`, `Cuadro2` | 2 rectángulos | — | Decoración | sin asignar | `Form_FormEventoCopia.form.txt:288-333` |
| `lblTitulo1`, `lblTitulo2` | 2 etiquetas | — | Rótulo «COPIA DE EVENTO <id>» (el segundo copia el texto del primero); también el título de ventana | CAP-3 | `:85-87` |
| `Etiqueta210`, `Etiqueta2` | 2 etiquetas | — | Rótulos «BUI» y «FECHA ALTA» | sin asignar | `Form_FormEventoCopia.form.txt:359-371, 514-526` |
| `BUI` | cuadro combinado, lista de valores | — | Se vacía y se llena con `m_ObjEntorno.Centros`; valor inicial el BUI del evento origen; destino de la copia | CAP-3 | `:88-96`; `Form_FormEventoCopia.form.txt:386-405` |
| `FECHAALTA` (en código `Me.FechaAlta`) | cuadro de texto no enlazado | — | Fecha propuesta: la del evento origen en `dd/mm/yyyy`; editable | CAP-3 | `:97`; `Form_FormEventoCopia.form.txt:487-498` |
| `ComandoCopiar` | botón | `OnClick` → `ComandoCopiar_Click` | Confirma, copia, avisa al padre y cierra | CAP-3 | `:13-47`; `Form_FormEventoCopia.form.txt:419-429` |
| `cmdSalir` | botón | `OnClick` → `cmdSalir_Click` | Cierra sin guardar, ignorando errores | CAP-3 | `:8-11`; `Form_FormEventoCopia.form.txt:334-343` |

## Flujos

### Apertura (`Form_Open`, `:49-78`)

1. Cursor de espera y `AjustarTamaño Me` (`:53-55`; `FG.bas:261`).
2. Sin `m_ObjEventoActivo`: «No se conoce el evento a copiar» (`:57-60`).
3. `EstablecerDatos` (`:81-104`): rótulos y título con el id del evento (`:85-87`); llena `BUI` con `m_ObjEntorno.Centros` (`:93-95`, que si es `Nothing` sale sin rellenar BUI ni fecha, `:90-92`); `BUI` y `FechaAlta` toman los del evento origen (`:96-97`).
4. Cualquier error: `MsgBox` crítico y cierre del formulario (`:71-77`).

### Copia (`ComandoCopiar_Click`, `:13-47`)

1. Confirmación Sí/No, con «Sí» por defecto; «No» sale sin hacer nada (`:19-23`).
2. `CopiarEventoCompletoConActividadesYMateriales(IDEVENTO, FechaAlta, BUI, m_Error)` (`:27`).
3. Con error: `Err.Raise 1000` (`:28-30`). Si no, relee el evento nuevo con `Constructor.getEvento` (`:31-34`), lanza `AltaEvento` (`:35`) y cierra (`:36`).
4. Error: `MsgBox` crítico con el texto acumulado; el formulario queda abierto (`:41-46`).

### Función de copia (`:109-344`)

1. Fecha: si el texto no es fecha, usa la del día (`:122-124`). Lee el evento origen; si no existe, «El Evento seleccionado no está registrado» (`:131-139`). BUI vacío ⇒ el del origen; sigue vacío ⇒ error (`:140-147`).
2. Id nuevo con `DameIDEvento(BUI, fecha)` (`:153-159`). Inserta el evento copiando todos los campos salvo `HORAFINALEVENTO`, `CAUSAFIN`, `Franqueado`, `IDExportacion`, `CodImportacion` e `IDParte`; `FechaRegistroAlta` toma la fecha indicada (`:160-173`).
3. Por cada actividad origen exige `ALIASTECNICO` y `FechaAlta` (`:194-204`), genera id con `DameIDActividad` (`:212-218`) y copia todos los campos salvo `IDExportacion`, `CodImportacion` e `IDFacturacion` (`:219-230`).
4. Por cada material de la actividad: `MaterialDeActividadAlta` con los datos del origen y `EsReparacion = "Sí"` fijo (`:234-258`), después `CopiarAnexoMaterial` (`:267`, `:345-408`) y `CopiarSeguimientoMaterial` (`:271`, `:409-455`).
5. Error: cualquier fallo salta a `errores` (`:298-344`), que intenta deshacer a mano (ver defectos). Devuelve el id del evento nuevo (`:296`).

## Reglas, validaciones y cálculos

- **Id de evento:** iniciales de `TbBuiIDEvento` (o tres primeras letras del BUI, mínimo 3) + `yy` + `mm` + ordinal de tres cifras siguiente al máximo existente con ese prefijo (`FG.bas:6171-6246`).
- **Id de actividad:** `DameValorSiguienteAlMaximoDeUnCampo("TbActividades", "IDActividad", ..., 2)` con el alias y la fecha (`FG.bas:6535-6575`).
- **Id de anexo y de seguimiento:** `DameID("TbAnexos", "IDAnexo")` = máximo numérico + 1 (`:369`; `FG.bas:371-430`); `DameIDIntervencion` para seguimientos (`:435`; `FG.bas:4289`).
- **Anexos:** solo se copian los que tienen `NombreArchivo`; título y descripción solo si no están vacíos (`:368-381`). El fichero no se duplica: el anexo nuevo apunta al mismo `NombreArchivo` (`:380`).
- **Fecha de la copia:** solo afecta al id y a `FechaRegistroAlta` (`:153, 164-165`); el resto de campos de fecha del evento se copian tal cual (`:169`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura del evento origen | `TbEventos` | SQL dinámico concatenado (`IDEvento='...'`) | no | `:131-135` |
| Alta del evento | `TbEventos` | DAO `AddNew`/`Update` sobre recordset de tabla | no | `:129-130, 160-173` |
| Lectura de actividades | `TbEventos` + `TbActividades` | SQL dinámico concatenado | no | `:187-190` |
| Alta de actividades | `TbActividades` | DAO `AddNew`/`Update` | no | `:206-230` |
| Lectura de materiales | `TbMaterial` | SQL dinámico concatenado | no | `:234-237` |
| Alta de material | `TbMaterial` | DAO vía `MaterialDeActividadAlta` | no | `FG.bas:5111, 5233` |
| Lectura/alta de anexos | `TbAnexos` | SQL dinámico concatenado / DAO | no | `:357-381` |
| Lectura/alta de seguimientos | `TbMaterialSeguimiento` | SQL dinámico concatenado / DAO | no | `:425-445` |
| Máximos para ids | `TbEventos`, `TbActividades`, `TbAnexos` | SELECT dinámico | no | `FG.bas:6225-6228, 6573`; `FG.bas:409-410` (DameID) |
| Deshacer tras error | `TbEventos`, `TbActividades`, `TbMaterial` | `DoCmd.RunSQL` DELETE concatenado | no | `:311-339` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `OpenRecordset` sobre nombres de tabla y recorrido de `Fields` por nombre; SQL concatenado con comillas simples; ids calculados en VBA con máximo + 1 y sin transacción; `Format(..., "dd/mm/yyyy")` y `Like '...*'` con comodín Jet (`FG.bas:6227`); `DoCmd.RunSQL` con `SetWarnings`; el reverso del copiado depende de borrados manuales, no de transacción de la base.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba. No hay control de permisos en el formulario: ningún control se habilita o deshabilita por rol (`:49-104`).

**Posibles defectos (sin corregir, a verificar):**

- El deshacer es incompleto: la consulta sobre `TbAnexos` es un `SELECT` ejecutado con `RunSQL` (`:326-331`) que no borra, y no se borran los seguimientos copiados (`TbMaterialSeguimiento`) ni los anexos copiados; además un `On Error Resume Next` (`:310`) oculta los fallos. Un error a medias podría dejar anexos y seguimientos huérfanos.
- El deshacer borra actividades por `IDEvento` (`:318-320`) y materiales solo por los ids acumulados en `ColMateriales` (`:324-338`); un material insertado cuyo id no llegó a la colección no se borraría.
- `p_FechaAlta` se sobrescribe con la fecha de cada actividad origen (`:195`) y, al ser parámetro de la propia función, el valor del evento ya no se usa después; no afecta a la copia del evento (ya insertado) pero sí es una reutilización de variable a verificar.
- El campo `FECHAALTAEVENTO` del evento copiado conserva la fecha del origen (`:169`), de modo que el id lleva la fecha nueva (`:153`) pero el campo no.
- `EsReparacion` se fija a «Sí» (`:257`) para todos los materiales, y `MaterialDeActividadAlta` solo escribe `FechaEntrega` cuando es «No» (`FG.bas:5257-5261`); `FechaEntrega` no se copia en la práctica.
- Ids por máximo + 1 sin bloqueo (`FG.bas:371-430`, `:369`): dos copias simultáneas podrían colisionar.
- Variables compartidas no declaradas localmente (`pregunta`, `flag`, `dato`, `m_SQL`; `VG.bas:203-206`), reutilizadas entre funciones anidadas.
- Si `BUI` queda vacío, `Nz(Me.BUI, "")` cae al BUI del evento origen (`:27, 140-142`). Errata «tieme» en un mensaje (`:145`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: copia en una única transacción; ids generados por la base; copiar `FECHAALTAEVENTO` según criterio a decidir; decidir si el anexo debe duplicar el fichero o compartirlo.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir el formulario sin evento activo muestra «No se conoce el evento a copiar» y lo cierra (`:57-60`).
2. Al abrir se muestran el id del evento en los rótulos, el BUI y la fecha de alta del origen (`:85-97`).
3. Responder «No» a la confirmación no crea nada (`:21-23`).
4. Copiar crea un evento con id nuevo según BUI y fecha, sin `HORAFINALEVENTO`, `CAUSAFIN`, `Franqueado`, `IDExportacion`, `CodImportacion` ni `IDParte` (`:167-170`).
5. Cada actividad, material, anexo con archivo y seguimiento del origen existe en el evento nuevo (`:206-281`).
6. Una actividad sin alias técnico o sin fecha impide la copia con su mensaje (`:197-204`).
7. Tras copiar, el padre muestra el evento nuevo y avisa de su id (`Form_FormEventoGestion.cls:2844-2858`).
8. Ante un fallo, no quedan eventos, actividades, materiales, anexos ni seguimientos de la copia (hoy no se cumple del todo; ver defectos).

## Pendientes

- Contenido de `Constructor.getEvento` y de `Dame` en `DameIDEvento` (`:31`; `FG.bas:6205`), no analizado aquí.
- Cuerpo completo de `DameValorSiguienteAlMaximoDeUnCampo` y `DameIDIntervencion` (`FG.bas:6573`, `FG.bas:4289`).
- Origen y contenido de `m_ObjEntorno.Centros` frente a la tabla `TbBuiIDEvento` (`:90-95`).
- Si `TbEventos`, `TbActividades`, `TbMaterial`, `TbAnexos` y `TbMaterialSeguimiento` tienen índices únicos sobre sus ids (ERD de origen).
