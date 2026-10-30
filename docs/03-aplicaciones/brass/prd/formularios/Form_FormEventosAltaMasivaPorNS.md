# Form_FormEventosAltaMasivaPorNS

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormEventosAltaMasivaPorNS.cls` (629 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta por lotes de eventos de reparación: el usuario rellena los datos comunes y una lista de números de serie (N/S) separados por «;», y el formulario crea, por cada N/S distinto, un evento (`TbEventos`), una actividad de entrada (`TbActividades`) y un material en reparación (`TbMaterial`). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`). La escritura la hacen tres funciones públicas del propio formulario (`AltaEvento`, `AltaActividad`, `AltaMaterial`) con DAO. Las citas son `Form_FormEventosAltaMasivaPorNS.cls:línea` salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`, `VG.bas` = `modules/Variables Globales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEventoGestion` | `ComandoAltaMasivaPorNS_Click` cierra el formulario si estaba abierto y lo abre sin `OpenArgs` | `Form_FormEventoGestion.cls:176-198` |

Al terminar con éxito abre `FormEventoGestion` si no estaba abierto, ejecuta `ComandoActualizarTodo_Click`, fija en sus filtros la fecha de alta (inicial y final) y el BUI, y ejecuta `ComboBUIBusqueda_AfterUpdate` (`:79-87`). Después muestra un aviso y se cierra (`:91-93`). No lanza eventos propios. `cmdSalir` cierra sin guardar (`:254-257`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `BUI` | combo limitado a la lista | `OnChange` → `BUI_Change` | Vacía `SubSistema`, `IDEquipo` y `Material`, recalcula `SubSistema` y la lista de `Contacto` | CAP-3 | `:11-57`, `.form.txt:527-531` |
| `Subsistema` | combo limitado a la lista | `OnEnter` → `Subsistema_Enter`; `OnChange` → `SubSistema_Change` | Recalcula `IDEquipo` | CAP-3 | `:190-249`, `.form.txt:387-394` |
| `IDEquipo` | combo de 3 columnas (ID y datos ocultos) | `OnEnter` → `IDEquipo_Enter`; `OnChange` → `IDEquipo_Change` | Recalcula `Material` | CAP-3 | `:110-160`, `.form.txt:742-747` |
| `Material` | combo | `OnExit` → `Material_Exit` | Si `Descripcion` está vacía, la rellena con material y nombre del equipo | CAP-3 | `:166-174`, `.form.txt:1602-1606` |
| `TipoEvento`, `Originador` | combos limitados a la lista | `OnEnter` → `Requery` | Catálogos `TbTipoEvento` y `TbOriginador` | CAP-3 | `:185-188, 250-253`, `.form.txt:595-600, 673-678` |
| `Criticidad` | combo de 2 columnas | — | Catálogo `TbCriticidad` | CAP-3 | `.form.txt:636-640` |
| `ALIASTECNICO` | combo de 3 columnas (lista de valores) | — | Se rellena en la carga con `RellenarComboTecnicos` | CAP-3 | `:272`, `.form.txt:1452-1454` |
| `Contacto` | combo | `OnEnter` → `Contacto_Enter` | Ejecuta `Me.Requery` del formulario | CAP-3 | `:259-261`, `.form.txt:1543-1549` |
| `Descripcion`, `PMPR`, `FECHAALTAEVENTO` | textos (fecha en formato «Short Date») | — | Datos comunes del evento | CAP-3 | `.form.txt:316-317, 559-560, 705-707` |
| `CadenaNS` | texto | — | Lista de N/S separados por «;»; obligatorio | CAP-3 | `:327, 349`, `.form.txt:1387-1388` |
| `Numero_Actividad` | texto deshabilitado y bloqueado | — | Con `ControlSource` y valor por defecto heredados; el código no lo usa | sin asignar | `.form.txt:346-348` |
| `ComandoAlta` («Alta») | botón | `OnClick` → `ComandoAlta_Click` | Confirma y lanza `AltaLote` | CAP-3 | `:59-101`, `.form.txt:440-442` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra | CAP-3 | `:254-257`, `.form.txt:478-480` |
| `ComandoAyudaFormulario` | botón | `OnClick` declarado, sin handler en el `.cls` | Sin efecto | sin asignar | `.form.txt:1133-1135` |
| 26 etiquetas, 2 rectángulos | rótulos (11 asteriscos, 13 de campo, `lblTitulo1`, `lblTitulo2`), `Cuadro2`, `Cuadro3` | — | `lblTitulo1` pasa a `Caption` | sin asignar | `:281-282`, `.form.txt:1058-1112` |

Controles: 45 (54 Begin en el .form.txt: 1 `Form`, 7 plantillas de estilo por defecto, 1 `Section` y los 45 controles anteriores: 14 de entrada, 3 botones, 26 etiquetas, 2 rectángulos). El campo `Tag` marca los obligatorios («OBLIGATORIO») y los que se copian al evento («DATOS»): `BUI` y `Material` y `CadenaNS` solo obligatorios; `PMPR` y `Contacto` solo «DATOS» (`.form.txt:532, 560, 1388, 1549, 1606`).

## Flujos

### Carga (`Form_Load`, `:263-296`)

Rellena `ALIASTECNICO` con los técnicos activos (`:272`; `FG.bas:7224`) y deja valores literales por defecto en `Criticidad`, `ALIASTECNICO`, `Originador` y `TipoEvento` (`:277-280`). Cualquier error: mensaje y cierre del formulario (`:289-295`).

### Cascada de combos (`:11-57, 190-244, 110-155, 602-628`)

1. `BUI_Change`: vacía los tres combos dependientes, carga `SubSistema` por BUI (`FG.bas:1209-1224`) y la lista de `Contacto` con los contactos de eventos de ese BUI; sin BUI, limita `Contacto` a la lista, y con BUI no (`:36-42, 609-621`).
2. `SubSistema_Change`: sin BUI o subsistema, vacía `IDEquipo` y `Material`; si no cambia el subsistema, no hace nada; si cambia, carga `IDEquipo` con los equipos del BUI y subsistema (`:205-232`; `FG.bas:1192-1207`).
3. `IDEquipo_Change`: si cambia el equipo, vacía `Material` y lo recarga con los materiales de eventos anteriores del equipo (`:123-143`; `FG.bas:1226-1245`).

### Alta del lote (`ComandoAlta_Click`, `AltaLote`, `:59-101, 299-409`)

1. Confirmación Sí/No con «No» por defecto (`:65-69`).
2. `AltaLote` recorre los controles: falla en el primero con «OBLIGATORIO» vacío (`:318-326`). `Descripcion` no puede contener «S/N» (`:331-335`).
3. Copia al `Evento` los controles con «DATOS» mediante `SetPropiedad` (`:339-347`; `Evento.cls:2601`).
4. Trocea `CadenaNS` por «;» y elimina duplicados sin distinguir mayúsculas (`:349-358`).
5. Por cada N/S: descripción = descripción del usuario + «S/N» + N/S; `BUI`; `NODO` desde `TbNodoBUI`; `AltaEvento`; `AltaActividad` (actividad, ubicación, horas y descripción literales en el código); `AltaMaterial` con `EsReparacion` = «Sí» y la fecha de inicio igual a la del evento (`:361-396`).
6. Un error en cualquier paso detiene el bucle y se muestra; no hay vuelta atrás de lo ya grabado (`:404-408, 95-100`). Éxito: abre y filtra `FormEventoGestion` (`:79-87`).

## Reglas, validaciones y cálculos

- **IDEvento:** inicial del BUI (de `TbBuiIDEvento` o sus tres primeras letras) + año y mes + ordinal de 3 cifras, máximo existente más uno (`FG.bas:6171-6242`).
- **IDActividad:** alias del técnico + `aammdd` de la fecha de alta + hora y minuto actuales + contador de 2 cifras, máximo más uno; si pasa de 99 ⇒ error (`:480-491`; `FG.bas:440-507`).
- **IDMaterial:** `IDActividad` + «_» + contador de 2 cifras (`:430`).
- **Valores fijos del evento:** hora inicial y tiempo de respuesta literales; fechas de registro con `Now()` (`:578-579, 585-586`).
- **Unicidad:** el evento y la actividad comprueban que el ID calculado no exista (`:556-564, 493-501`); el material no se crea si ya existe el mismo N/S en la actividad (`:422-429`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Listas de la cascada | `TbNodoBUI`, `TbSubsistemaBui`, `TbEquipos`, `TbMaterial`, `TbActividades`, `TbEventos` | `RowSource` guardado o SQL concatenado asignado en ejecución | — | `:37, 142, 231, 615-620`; `FG.bas:1192-1245` |
| Nodo del BUI | `TbNodoBUI` | `Dame` con SQL concatenado | — | `:363`; `FG.bas:290-330` |
| Siguiente ID | `TbEventos`, `TbActividades`, `TbMaterial` | `LIKE`/recorrido completo de la tabla en VBA | no | `FG.bas:6225-6232, 456-480` |
| Alta de evento | `TbEventos` | DAO `AddNew`/`Update` | no | `:556-588` |
| Alta de actividad | `TbActividades` | DAO `AddNew`/`Update` | no | `:493-524` |
| Alta de material | `TbMaterial` | DAO `AddNew`/`Update` | no | `:422-441` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL concatenado con textos entre comillas simples y `Like '…*'` con comodín de Jet (`FG.bas:6227-6229`); `Format(…, "yy")`, `"mm"` y `"000"` para construir claves (`:481-484`; `FG.bas:6242`); identificadores calculados en VBA con «máximo más uno» sin transacción; campos de fecha escritos desde cadenas (`:577`) y horas como texto (`:578-579`); varias inserciones encadenadas que en PostgreSQL convendría agrupar en una transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Sin transacción: si falla el N/S k, los k-1 anteriores (eventos, actividades y materiales) quedan grabados, el formulario sigue abierto y reintentar los duplicaría (`:359-401, 404-408`).
- Los errores de los generadores de ID no se comprueban: `DameIDEvento`, `DameValorSiguienteAlMaximoDeUnCampo` y `Dame` devuelven «#ERR|…» y el código solo valida cadena vacía o ignora el resultado, de modo que ese texto podría usarse como ID o nodo (`:363, 430, 491, 551-555`; `FG.bas:6249, 507`).
- `SubSistema_Change` comprueba la global `flag` sin asignarla desde `EstablecerComboEquipo` (`:225-230`; `VG.bas:205`): puede dar un error espurio por un valor residual o no detectar el error real, y «ERR|…» se usaría como `RowSource`.
- Con más de 99 N/S en el mismo minuto, el contador de 2 cifras de `IDActividad` se agota (`:491`; `FG.bas:498-501`); el lote se corta a medias.
- Entradas vacías en `CadenaNS` («A;;B» o «;» final) no se descartan y generarían un evento con «S/N» sin número (`:349-358, 361`).
- `Contacto_Enter` ejecuta `Me.Requery` del formulario completo (`:259-261`); en un formulario no enlazado podría ser inocuo o recargar los controles, a verificar.
- `BUI_Change` no vacía `Contacto` ni `Descripcion` al cambiar de BUI (`:21-24`), y `Material_Exit` solo rellena `Descripcion` si está vacía, por lo que una descripción previa queda desfasada (`:168-174`).
- `AltaEvento` escribe `FECHAALTAEVENTO` desde una cadena sin validarla como fecha (solo `DameIDEvento` lo hace) y el formato de `FECHAALTAEVENTO` depende de la configuración regional (`:577`; `FG.bas:6192-6195`).
- El aviso final habla de «tantos eventos como números de serie» aunque se hayan eliminado duplicados (`:91, 349-358`).
- `ComandoAyudaFormulario` tiene `OnClick` sin handler (`.form.txt:1135`); `Numero_Actividad` conserva un `ControlSource` y un `DCount` heredados (`.form.txt:347-348`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: lote en una única transacción con informe de resultado por N/S; identificadores generados por la base; validar y normalizar la lista de N/S (vacíos, espacios, duplicados) antes de grabar; mover los valores literales a configuración.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir se cargan los técnicos activos y los valores por defecto (`:272-280`).
2. Al elegir BUI se vacían subsistema, equipo y material y el subsistema solo ofrece los de ese BUI (`:21-24, 36-38`).
3. Elegir subsistema ofrece solo los equipos de ese BUI y subsistema (`:225-232`).
4. Sin cumplimentar un campo obligatorio, «Alta» no crea nada y nombra el campo (`:318-326`).
5. Una descripción con «S/N» se rechaza (`:331-335`).
6. Con N números de serie distintos se crean N eventos, N actividades y N materiales, cada uno con su descripción «… S/N <número>» (`:359-396`).
7. Un número de serie repetido en la lista se crea una sola vez (`:355-357`).
8. Responder «No» a la confirmación no crea nada (`:67-69`).
9. Tras el éxito se abre la gestión de eventos filtrada por la fecha y el BUI del lote (`:79-87`).

## Pendientes

- Si `TbEventos`, `TbActividades` y `TbMaterial` tienen índices únicos sobre los IDs y sobre N/S (ERD de origen).
- Contenido de `Form_FormEventoGestion.ComandoActualizarTodo_Click` y `ComboBUIBusqueda_AfterUpdate`, invocados desde `:82, 87`.
- Efecto real de `Me.Requery` en un formulario sin origen de datos (`:260`).
- Qué ocurre con el parámetro `p_EnTodos = EnumSino.No` en `Constructor.getTecnicos` (`:272`; `FG.bas:7224`).
