# Form_FormMaterialAltaReparacion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormMaterialAltaReparacion.cls` (594 líneas) y `.form.txt` (2580 líneas) · **PRD:** [README](../README.md)

Alta de un material en reparación asociado a un equipo o a una actividad de un evento. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): recibe el contexto por `OpenArgs`, rellena controles y persiste en `TbMaterial` mediante DAO. Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`; las citas sin fichero son `Form_FormMaterialAltaReparacion.cls:línea`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormMaterialEleccion` (reparación = «Sí») | Cierra el formulario si ya estaba abierto y lo abre con `OpenArgs` de 4 elementos (`BUI\|SubSistema\|IDEquipo\|IDEquipo`) o de 6 (añade `IDEvento\|IDActividad`) | `Form_FormMaterialEleccion.cls:31-33, 53-58` |

Tras guardar se cierra solo (`:238`); `cmdSalir` cierra sin guardar (`:249-252`). Si el alta es de actividad y `FormEventoGestion` está abierto, invoca `Form_FormEventoGestion.ListaActividades_Click` (`:218-220`). No lanza eventos propios. Un comentario de `FG.bas:5142` menciona también como llamante a `Form_FormMaterialRepuestoAlta`, pero solo como comentario.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComandoAlta` («Registrar») | botón | `OnClick` → `ComandoAlta_Click` | Lee los controles y da de alta el material | CAP-7 | `:172-247`; `Form_FormMaterialAltaReparacion.form.txt:348-358` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | CAP-7 | `:249-252`; `Form_FormMaterialAltaReparacion.form.txt:386-397` |
| `ComandoVerActividad` (rótulo «Ver Evento») | botón, deshabilitado de inicio | `OnClick` | Abre/enfoca `FormActividadGestion` y selecciona la actividad | CAP-4 | `:253-289`; `Form_FormMaterialAltaReparacion.form.txt:1939-1950` |
| `ComandoVerEvento` | botón, deshabilitado de inicio | `OnClick` | Abre/enfoca `FormEventoGestion` y selecciona el evento | CAP-3 | `:290-327`; `Form_FormMaterialAltaReparacion.form.txt:1886-1898` |
| `ComboMaterial` | combo | `OnChange` → `ComboMaterial_Change` | Material ya existente del equipo; copia datos a `PN` y `NS` | CAP-7 | `:329-355, 409-415`; `Form_FormMaterialAltaReparacion.form.txt:426-453` |
| `Material` | texto no enlazado | `OnExit` → `Material_Exit` | Material nuevo (no existente); si cambia, vacía `ComboMaterial` | CAP-7 | `:434-442`; `Form_FormMaterialAltaReparacion.form.txt:1772` |
| `PN`, `NS` | texto no enlazado | — | Referencia y nº de serie | CAP-7 | `:185-186`; `Form_FormMaterialAltaReparacion.form.txt:496, 524` |
| `COSTE` | texto, formato moneda | — | Coste opcional | CAP-7 | `:183`; `Form_FormMaterialAltaReparacion.form.txt:552-566` |
| `ComboReparadoPor` | combo (`SELECT DISTINCT ReparadoPor` de `TbMaterial`) | — | Quién repara | CAP-7 | `:187`; `Form_FormMaterialAltaReparacion.form.txt:582-602` |
| `ComboTipoAccion` | combo (`TbTipoAccion.ESTADOREPARABLE`) | — | Tipo de acción | CAP-7 | `:184`; `Form_FormMaterialAltaReparacion.form.txt:618-639` |
| `FechaIncio` | texto, fecha corta | — | Fecha inicial; por defecto hoy | CAP-7 | `:189, 419`; `Form_FormMaterialAltaReparacion.form.txt:948-961` |
| `ComboGarantia` | combo «Sí;No» | — | Garantía; por defecto «No» | CAP-7 | `:190, 408`; `Form_FormMaterialAltaReparacion.form.txt:1031-1052` |
| `FechaFinIntervencion` | texto, fecha corta | — | Se lee y se pasa como fecha de entrega | CAP-7 | `:191, 205-207`; `Form_FormMaterialAltaReparacion.form.txt:1186-1199` |
| `ComboUltimaIntervencion`, `OBSERVACIONES`, `ComboOrigen`, `ComboDestino` | combos y texto | — | Se rellenan o muestran pero `ComandoAlta_Click` no los lee | CAP-7 | `:6-43, 417-418`; `Form_FormMaterialAltaReparacion.form.txt:1094-1117, 1131, 1665, 1828` |
| `IDEvento`, `IDActividad`, `BUI`, `SubSistema`, `IDEquipo` | texto bloqueado (`Locked`) | — | Contexto recibido en `OpenArgs` | CAP-7 | `:400-407`; `Form_FormMaterialAltaReparacion.form.txt:1242-1375` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título; `lblTitulo1` da el `Caption` de la ventana | CAP-7 | `:365`; `Form_FormMaterialAltaReparacion.form.txt:1391-1422` |
| 10 etiquetas «*» y 21 etiquetas de rótulo (31 etiquetas en total con los títulos) | etiquetas | — | Marcas de obligatorio y rótulos | CAP-7 | `Form_FormMaterialAltaReparacion.form.txt:655-1590` |
| `MarcoEvento`, `Marco240`, `Marco17`, `CuadroSuperior`, `Cuadro2` | marcos y rectángulos | — | Decoración | sin asignar | `Form_FormMaterialAltaReparacion.form.txt:302-334, 473, 736, 1615` |

Controles: 67 (67 Begin en el .form.txt, sin contar `Begin Form`): 59 nombrados (31 etiquetas, 12 cuadros de texto, 7 combos, 4 botones, 3 marcos, 2 rectángulos), 7 estilos por defecto (`Form_FormMaterialAltaReparacion.form.txt:167-270`) y la sección `Detalle` (`Form_FormMaterialAltaReparacion.form.txt:291-293`).

## Flujos

### Apertura (`Form_Load`, `:357-433`)

1. `AjustarTamaño Me` (`FG.bas:261`) y `Caption` desde `lblTitulo1` (`:363-365`).
2. `OpenArgs` debe contener `|` y 4 o 6 partes; si no, «El formulario ha sido abierto con parámetros erróneos» y cierre (`:366-377`). Con 6 partes exige evento y actividad, los vuelca a controles y habilita `ComandoVerActividad` (`:395-403`).
3. Vuelca `BUI`, `SubSistema`, `IDEquipo`, fija garantía «No» (`:405-408`), construye el `RowSource` de `ComboMaterial` con los materiales de actividades de eventos del equipo (`:409-415`), llena origen/destino (`:416-418`) y pone la fecha de hoy (`:419`).
4. Error: mensaje crítico y cierre del formulario (`:426-433`).

### Alta (`ComandoAlta_Click`, `:172-247`)

1. Material = columna 1 de `ComboMaterial` si hay; si no, el texto de `Material` (`:181-197`). Motivo fijo «Reparación» (`:188`).
2. Con actividad: `MaterialDeActividadAlta` (`:199-212`), refresca la gestión del evento (`:218-220`). Sin actividad: `MaterialNODeActividadAlta` (`:221-233`). En ambos pasa «Sí» como `EsReparacion` (`:207, 228`).
3. Un retorno con `|` es error: se muestra y el formulario queda abierto (`:208-211, 240-246`). Si todo va bien, cierra (`:238`).

### Ver actividad / ver evento

`ComandoVerActividad` sale si no hay actividad (`:256-258`); abre `FormActividadGestion`, asigna `ComboIDActividadBusqueda` y llama a `ComboIDActividadBusqueda_Change` (`:267-278`). `ComandoVerEvento` hace lo propio con el `IDEVENTO` del control (`:293-316`).

## Reglas, validaciones y cálculos

- **Selección de material:** elegir en el combo anula `Material` y copia `Column(2)` a `NS` y `Column(3)` a `PN` (`:332-338`).
- **Alta con actividad** (`FG.bas:5152-5278`): actividad y material obligatorios (`FG.bas:5152-5159`); coste numérico si existe (`FG.bas:5160-5165`); no admite el mismo material, `EsReparacion`, PN y NS en la misma actividad («Ya se ha dado de alta…», `FG.bas:5183-5209`); `IDMaterial` = actividad + «_» + 2 cifras (`FG.bas:5217`).
- **Alta sin actividad** (`:443-594`): `IDEquipo` numérico (`:482-485`); coste numérico (`:486-491`); fecha inicial = `Now()` si no es fecha (`:492-494`); rechaza si `TbEquipos.FechaObsoleto` es posterior a hoy (`:506-518`); `IDMaterial` = `MNA_` + 4 cifras (`:533`); no comprueba material vacío.
- **Garantía y reparación** distintas de «Sí»/«No» se normalizan a «No» (`:495-500`; `FG.bas:5169-5174`).
- `FechaFinIntervencion` llega como fecha de entrega, pero solo se escribe si `EsReparacion` = «No» (`:574-578`; `FG.bas:5258-5262`), y aquí es «Sí».
- **`Registrar` (`:44-171`)**: función privada con las reglas de intervención (destino, motivo, última intervención, fechas) y alta de `MaterialIntervencionAlta` (`:157`). Ningún código del formulario la invoca.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Combo de materiales del equipo | `TbMaterial`, `TbActividades`, `TbEventos` | SQL dinámico (`IDEquipo` concatenado) | — | `:409-415` |
| Combos origen/destino | `TbSubsistemaBui` | SQL fija, recorrido DAO | — | `:15-29` |
| Equipo obsoleto | `TbEquipos` | SQL dinámico concatenado (`Dame`) | — | `:506`; `FG.bas:300-310` |
| Siguiente `IDMaterial` | `TbMaterial` | Recorrido DAO de toda la tabla y `Max` en VBA | no | `FG.bas:469-496` |
| Duplicado de actividad | `TbMaterial`, `TbActividades` | SQL dinámico concatenado | — | `FG.bas:5183-5210` |
| Comprobar ID libre | `TbMaterial` | SQL dinámico concatenado | — | `:540-548`; `FG.bas:5224-5232` |
| Alta de material | `TbMaterial` | DAO `AddNew`/`Update` | no | `:549-579`; `FG.bas:5233-5263` |

## Coexistencia

Adaptar a PostgreSQL vinculado: `IDMaterial` textual generado con `Max` en VBA, sin transacción ni secuencia; SQL concatenado con comillas simples y `IDEquipo` sin comillas; `Garantia` y `EsReparacion` como texto «Sí»/«No»; fechas asignadas desde cadenas locales (`Format(Now(), "dd/mm/yyyy")`, `:419`) y `Coste` desde cadena con formato regional; `SELECT DISTINCT … ORDER BY` de los combos.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La consulta de `ComboMaterial` devuelve `IDMaterial, Material, PN, NS` (`:409`) pero `ComboMaterial_Change` asigna `Column(2)` a `NS` y `Column(3)` a `PN` (`:335-338`): PN y NS podrían quedar intercambiados.
- `ComandoAlta_Click` no llama a `Registrar` ni a `MaterialIntervencionAlta`: no valida destino, última intervención ni resultado, y no registra la intervención (`:172-247` frente a `:44-171`).
- Origen, destino, última intervención y observaciones se muestran pero no se leen (`:172-247`).
- La fecha final se pasa como fecha de entrega pero no se escribe (`:191, 574-578`).
- `strIDEquipo` se asigna dos veces con `dato(2)` y `dato(3)`; prevalece la cuarta parte (`:381-382, 386-387`); la comprobación repite la misma condición (`:391`).
- El rechazo por obsolescencia se activa con fecha de obsolescencia futura (`:514`), lo que parece inverso al mensaje (`:515`).
- Sin actividad no se comprueba material vacío (`:443-594`) y no se refresca el evento.
- `ComandoVerEvento` queda siempre deshabilitado (`Form_FormMaterialAltaReparacion.form.txt:1887`; `:402-403`).
- Rótulo «Ver Evento» en `ComandoVerActividad` (`Form_FormMaterialAltaReparacion.form.txt:1950`).
- `Max` + 1 sin bloqueo (`FG.bas:469-496`): dos altas simultáneas podrían colisionar.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: identificador generado por la base de datos y registro de la intervención en la misma operación.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin `OpenArgs` válidos muestra el error y cierra (`:366-377`).
2. Elegir un material existente rellena PN y NS con los valores correctos del material (`:332-338`).
3. Alta con actividad: un material repetido en la actividad se rechaza (`FG.bas:5207`).
4. Alta sin actividad con coste no numérico se rechaza (`:486-491`).
5. Un alta correcta cierra el formulario y refresca la gestión del evento si estaba abierta (`:218-220, 238`).
6. Un error muestra el mensaje y mantiene el formulario abierto (`:240-246`).

## Pendientes

- Estructura de `TbMaterial` (nulabilidad de `Material`, tipos de fecha y coste) para el ERD.
- Declaración de las globales `m_SQL`, `flag`, `dato` y `pregunta`, usadas sin declarar en el formulario (`:15, 129, 133, 246`).
- Contenido de `ListaActividades_Click` de `Form_FormEventoGestion` (`:219`) y de `ComboIDActividadBusqueda_Change` (`:270`).
- Si `Form_FormMaterialRepuestoAlta` llama a `MaterialNODeActividadAlta` (`:474`).
