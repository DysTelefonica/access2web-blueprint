# Form_FormMaterialRepuestoAlta

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormMaterialRepuestoAlta.cls` (246 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta de un material como repuesto de una actividad («ALTA DE REPARACIÓN DE MATERIAL», `.form.txt:65`). Es un formulario **no enlazado** (sin `RecordSource`): recibe en `OpenArgs` el contexto (BUI, subsistema, equipo, evento, actividad), rellena unos controles de solo lectura y delega el alta en `FG.bas:MaterialDeActividadAlta` (`FG.bas` = `modules/Funciones Generales.bas`), que inserta en `TbMaterial` con `EsReparacion = "No"`. Las citas son `Form_FormMaterialRepuestoAlta.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEventoGestion` | Abre `FormMaterialEleccion` pasando `IDEvento\|IDActividad` | `Form_FormEventoGestion.cls:1206-1209` |
| `Form_FormMaterialEleccion` (elección «No» a «¿es reparación?») | Cierra este formulario si estaba abierto, lee BUI/subsistema/equipo del evento y lo abre con seis campos separados por `\|` | `Form_FormMaterialEleccion.cls:81-90` |

Argumento esperado: `BUI\|Subsistema\|Equipo\|IDEquipo\|IDEvento\|IDActividad` (`:188-199`); si no hay `\|` o no hay exactamente seis campos, error «abierto con parámetros erróneos» y cierre (`:184-193`, `:231-237`). Salida: tras guardar refresca `FormEventoGestion` si está abierto (`:47-49`) y se cierra (`:53`); `cmdSalir` cierra sin guardar (`:64-67`).

## Controles y eventos

Controles: 57 `Begin` en el `.form.txt` = 1 `Form` + 1 `Section` (`Detalle`) + 7 plantillas de controles por defecto (`.form.txt:168-292`) + 48 controles reales (tabla).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `Form` | formulario | `OnLoad` → `Form_Load` | Valida `OpenArgs`, vuelca contexto, fija garantía «No», carga `RowSource` de `ComboMaterial` y pone la fecha de hoy | CAP-7 | `:174-238`, `.form.txt:109` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` | Recoge los controles y llama a `MaterialDeActividadAlta`; cierra | CAP-7 | `:6-62`, `.form.txt:357-359` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | CAP-7 | `:64-67`, `.form.txt:395-397` |
| `ComandoVerActividad` | botón (rótulo «Ver Evento») | `OnClick` | Abre o enfoca `FormActividadGestion` y selecciona la actividad | CAP-4 | `:68-104`, `.form.txt:1622-1624` |
| `ComandoVerEvento` | botón | `OnClick` | Abre o enfoca `FormEventoGestion` y selecciona el evento | CAP-3 | `:105-142`, `.form.txt:1582-1584` |
| `ComboMaterial` | combo (lista de materiales ya dados de alta en el equipo) | `OnChange` → `ComboMaterial_Change` | Si hay valor: vacía `Material` y copia PN y NS de las columnas 1 y 2; si no, y `Material` está vacío, vacía PN y NS | CAP-7 | `:143-172`, `.form.txt:445-454` |
| `Material` | texto (material nuevo, no existente) | `OnExit` → `Material_Exit` | Si el texto cambia y no es vacío, anula `ComboMaterial` | CAP-7 | `:239-247`, `.form.txt:1555-1556` |
| `PN`, `NS`, `COSTE`, `OBSERVACIONES` | 4 textos | — | Datos libres del material; PN/NS se autorrellenan desde el combo | CAP-7 | `:26-28, 152-155`, `.form.txt:502, 530, 559, 1074` |
| `FechaIncio` | texto | — | Fecha de inicio; se inicializa con `Format(Now(), "dd/mm/yyyy")` | CAP-7 | `:223` |
| `ComboReparadoPor` | combo (distintos `TbMaterial.ReparadoPor`) | — | Suministrador; `LimitToList` | CAP-7 | `.form.txt:577, 595-600` |
| `ComboTipoAccion` | combo (`TbTipoAccion.ESTADOREPARABLE`) | — | Tipo de acción; `LimitToList` | CAP-7 | `.form.txt:614, 632-637` |
| `ComboGarantia` | combo, lista «Sí;No» | — | Garantía; valor inicial «No»; `LimitToList` | CAP-7 | `:215`, `.form.txt:1027, 1045-1049` |
| `IDEvento`, `IDActividad`, `BUI`, `SubSistema`, `IDEquipo` | 5 textos bloqueados (`Locked`) | — | Contexto recibido en `OpenArgs` | CAP-7 | `:208-212`, `.form.txt:1118, 1147, 1177, 1207, 1237` |
| `CuadroSuperior`, `Cuadro2` | 2 rectángulos | — | Decoración | sin asignar | `.form.txt:312, 335` |
| `MarcoEvento`, `Marco240`, `Marco17` | 3 grupos de opciones | — | Marcos visuales | sin asignar | `.form.txt:477, 741, 1480` |
| 24 etiquetas (`Etiqueta236-238, 241-244, 247, 248, 252, 264, 265, 98, 99, 8, 9, 10, 13, 14, 16, 21, 22`, `lblTitulo1`, `lblTitulo2`) | etiquetas | — | Rótulos; 7 son «\*». `lblTitulo1` alimenta `Me.Caption` | sin asignar | `:182`, `.form.txt:1298, 1319-1320` |

Desglose de los 48: 4 botones, 4 combos, 11 textos (`Material`, `PN`, `NS`, `COSTE`, `OBSERVACIONES`, `FechaIncio` y los 5 bloqueados), 2 rectángulos, 3 marcos y 24 etiquetas. Cuadre: 1 + 1 + 7 + 48 = 57.

## Flujos

### Apertura (`Form_Load`, `:174-238`)
1. `AjustarTamaño Me` y título desde `lblTitulo1` (`:181-182`). Valida los seis campos de `OpenArgs`: BUI, subsistema, `IDEquipo`, evento y actividad no vacíos (`:200-207`).
2. Vuelca contexto en los controles bloqueados y habilita los dos botones «Ver» (`:208-214`); garantía «No» (`:215`).
3. Compone en `m_SQL` los materiales distintos (`Material`, `PN`, `NS`) de actividades de eventos del mismo `IDEquipo`, con el identificador **concatenado** (`:216-222`).
4. Fecha de hoy y foco en `ComboMaterial` (`:223-224`).

### Registro (`ComandoRegistrar_Click`, `:6-62`)
1. Sin `strIDActividad`: «No se sabe para qué actividad es el repuesto» (`:13-16`).
2. Material = `ComboMaterial` o, si está vacío, `Material` (`:17-20`); resto de campos con `Nz` (`:21-28`).
3. `MaterialDeActividadAlta(..., "No", strFecha)`: el penúltimo argumento es `EsReparacion`, el último `FechaEntrega`; **la misma fecha se pasa como fecha inicial y de entrega** (`:34-36`).
4. Un resultado con `\|` es error y se muestra (`:37-41`). Éxito: refresca `FormEventoGestion` con `ListaActividades_Click` si está abierto (`:47-49`) y cierra (`:53`).
5. Error: `MsgBox` crítico; el formulario sigue abierto (`:55-61`). Un error no 1000 usa `On Error Resume Next` en el manejador (`:56-59`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Actividad obligatoria | formulario y función | `:13-16`; `FG.bas:5152-5155` |
| Material obligatorio («Se ha de indicar el material») | función | `FG.bas:5156-5159` |
| Coste, si no vacío, numérico | función | `FG.bas:5160-5165` |
| Fecha inicial no válida ⇒ `Now()`; garantía distinta de Sí/No ⇒ «No» | función | `FG.bas:5166-5171` |
| Con `EsReparacion = "No"`, `FechaEntrega` no vacía debe ser fecha | función | `FG.bas:5175-5182` |
| Duplicado: misma actividad, material, `EsReparacion`, PN y NS (nulos tratados con `Is Null`) ⇒ «Ya se ha dado de alta este material para esta actividad» | función | `FG.bas:5183-5209` |
| `IDMaterial` = `IDActividad_` + siguiente máximo de dos dígitos; si ya existe, error | función | `FG.bas:5217-5232`; `FG.bas:492-496` |
| Campos opcionales solo se escriben si no vacíos; `FechaEntrega` solo si es fecha | función | `FG.bas:5235-5260` |

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de materiales del equipo | `TbMaterial`, `TbActividades`, `TbEventos` | SQL dinámico concatenado (`IDEquipo`) como `RowSource` | — | `:216-222` |
| Suministradores y tipos de acción | `TbMaterial`, `TbTipoAccion` | SQL guardada en el `RowSource` del control | — | `.form.txt:597, 634` |
| Búsqueda de duplicado | `TbMaterial`, `TbActividades` | DAO `OpenRecordset` sobre SQL concatenado con comillas | no | `FG.bas:5183-5204` |
| Siguiente `IDMaterial` | `TbMaterial` | DAO, recorre y calcula el máximo en VBA | no | `FG.bas:470-496` |
| Comprobación de existencia del ID | `TbMaterial` | SQL concatenado | no | `FG.bas:5224-5232` |
| Alta | `TbMaterial` | DAO `AddNew`/`Update` | no | `FG.bas:5233-5263` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `IDMaterial` calculado en VBA (máximo + 1 sobre cadena con prefijo) fuera de transacción; SQL concatenado con texto entre comillas simples; fecha pasada como cadena `dd/mm/yyyy` a un campo fecha (`:223`; `FG.bas:5166-5168, 5240`); `DISTINCT` y `Is Null` en el dialecto Jet/ACE.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Orden de argumentos invertido: el origen envía `IDEquipo` y luego `Equipo` (`Form_FormMaterialEleccion.cls:89`), pero el formulario guarda `dato(2)` en `strEquipo` y `dato(3)` en `strIDEquipo` (`:196-197`). El `WHERE` de `ComboMaterial` (`:219`) y `Me.IDEquipo` (`:212`) usarían el nombre del equipo.
- La comprobación de parámetros repite `strIDEquipo` y no valida `strEquipo` (`:200`).
- Se pasa la misma fecha como inicial y de entrega (`:36`); la entrega queda informada en el alta (`FG.bas:5258-5260`).
- El comentario de la función exige que el evento no esté franqueado (`FG.bas:5136`), pero no hay comprobación de franqueo en el código revisado (`FG.bas:5111-5268`).
- Máximo + 1 sin transacción ni bloqueo (`FG.bas:5217`, `5233-5263`): dos altas simultáneas podrían chocar.
- SQL concatenado con material, PN o NS entre comillas simples (`FG.bas:5184-5195`): un apóstrofo rompería la consulta.
- Rótulo «Ver Evento» en `ComandoVerActividad` (`.form.txt:1623`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: identificador generado por la base; unicidad (actividad, material, PN, NS) como restricción; comprobar franqueo del evento en servidor.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin los seis campos de contexto muestra «parámetros erróneos» y cierra (`:184-207`).
2. Registrar con material vacío muestra «Se ha de indicar el material» y no escribe (`FG.bas:5156-5159`).
3. Un coste no numérico muestra el error de coste (`FG.bas:5160-5165`).
4. Registrar un material con la misma actividad, PN y NS lo rechaza (`FG.bas:5206-5209`).
5. Elegir un material del combo copia PN y NS y vacía `Material` (`:149-155`).
6. Un alta correcta crea un `IDMaterial` con prefijo de la actividad, refresca la gestión del evento y cierra (`:47-53`; `FG.bas:5217`).

## Pendientes

- Contenido de `AjustarTamaño` (`:181`) y comportamiento exacto de `FormularioAbierto` (`FG.bas:171-186`).
- `ListaActividades_Click` de `Form_FormEventoGestion` (`:48`) y `ComboIDActividadBusqueda_Change` (`:85, 91`), no leídos.
- Campos de `TbMaterial` y su tipo (ERD de origen).
