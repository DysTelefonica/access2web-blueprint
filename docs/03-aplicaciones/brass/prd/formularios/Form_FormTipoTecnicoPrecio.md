# Form_FormTipoTecnicoPrecio

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormTipoTecnicoPrecio.cls` (170 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta y edición de un tramo de precios de un perfil técnico (`TbTipoTecnicoPrecios`): tipo de técnico, fecha inicial, fecha final, precio de hora laboral y de hora extra. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): lee el tramo de la global `m_ObjTipoTecnicoPrecioActivo` (`Variables Globales.bas:248`) y persiste con `TipoTecnicoPrecio.Registrar`. Las citas son `Form_FormTipoTecnicoPrecio.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormTipoTecnicoPreciosGestion` (alta) | Fija `m_ObjTipoTecnicoPrecioActivo` al tramo seleccionado con `ParaAlta = Sí`, cierra el formulario si estaba abierto y lo abre | `Form_FormTipoTecnicoPreciosGestion.cls:48-54` |
| `Form_FormTipoTecnicoPreciosGestion` (edición) | Igual con `ParaAlta = No`, solo si el tramo es el último de su tipo | `Form_FormTipoTecnicoPreciosGestion.cls:84-95` |

Eventos hacia quien lo abrió: `Alta` (`:5`, lanzado en `:121`) y `Edicion` (`:6`, lanzado en `:123`), según exista o no el tramo inicial. Tras guardar se cierra solo (`:127`); `cmdSalir` cierra sin guardar (`:140-143`).

## Controles y eventos

Controles: 25 (25 `Begin` en el `.form.txt`): el propio `Form` (1), la sección `Detalle` (1), siete plantillas de estilo por defecto (`Label`, `Rectangle`, `Image`, `CommandButton`, `OptionGroup`, `TextBox`, `ComboBox`) (7) y dieciséis controles reales.

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `TIPOTECNICO` | texto, deshabilitado | — | Perfil técnico del tramo; se rellena desde el código, no se edita | CAP-9 Técnicos, ausencias, fiestas y precios | `:56, 71`; `.form.txt:426-453` |
| `FechaInicial` | texto, formato fecha corta | — | Inicio de vigencia | CAP-9 | `:57-59`; `.form.txt:497-525` |
| `FechaFinal` | texto, formato fecha corta | — | Fin de vigencia; vacío = tramo abierto | CAP-9 | `:60-62`; `.form.txt:553-582` |
| `PrecioHoraLab` | texto, formato moneda | — | Precio de hora laboral | CAP-9 | `:63-65`; `.form.txt:664-694` |
| `PrecioHoraExtra` | texto, formato moneda | — | Precio de hora extra | CAP-9 | `:66-68`; `.form.txt:695-726` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` → `ComandoRegistrar_Click` | Guarda y cierra | CAP-9 | `:88-138`; `.form.txt:322-359` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra sin guardar | CAP-9 | `:140-143`; `.form.txt:360-398` |
| `lblTitulo1`, `lblTitulo2` | etiquetas «PERFIL TÉCNICO PRECIO» | — | `lblTitulo1` pasa a `Caption` del formulario | CAP-9 | `:153`; `.form.txt:454-496` |
| `Etiqueta5`, `Etiqueta10`, `Etiqueta12`, `Etiqueta13`, `Etiqueta14` | etiquetas | — | Rótulos «Tipo Técnico», «Fecha Inicial», «Fecha Final», «€ Hora Lab», «€ Hora Extra» | CAP-9 | `.form.txt:399-425, 526-552, 583-663` |
| `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | sin asignar | `.form.txt:276-321` |

## Flujos

### Apertura (`Form_Load`, `:144-170`)

1. Cursor de espera, `AjustarTamaño Me`, título desde `lblTitulo1` (`:147-153`) y `EstablecerDatos` (`:154`).
2. `EstablecerDatos` (`:41-86`) distingue tres casos: (a) global activa con `ParaAlta = No`: **relee** el tramo con `Constructor.getTipoTecnicoPrecio(IDTipoTecnicoFecha)` en `m_ObjTipoTecnicoPrecioAlInicio` y copia al formulario los valores que sean fecha o número (`:45-69`); sin registro, «No se puede obtener los datos del precio del perfil» (`:51-54`); (b) global activa con `ParaAlta = Sí`: solo copia el tipo y crea un objeto nuevo para alta (`:70-75`); (c) sin global: objeto nuevo en alta, sin tipo (`:76-79`).
3. Cualquier error: `MsgBox` crítico y cierre del formulario (`:163-169`).

### Guardado (`ComandoRegistrar_Click`, `:88-138`)

1. `HaHabidoCambios`: en alta (sin tramo inicial) siempre cierto (`:13-16`); en edición compara las cuatro cajas con el tramo inicial (`:17-32`). Sin cambios: «Ningún cambio que registrar» (`:103-106`).
2. Copia las cinco cajas al objeto activo con `Nz(..., "")` (`:108-113`) y llama a `Registrar` (`:114`).
3. `Registrar` (`TipoTecnicoPrecio.cls:403-519`): exige alta/edición (`TipoTecnicoPrecio.cls:418-421`) y pasa `MotivoNoOK` (`TipoTecnicoPrecio.cls:424-430`); si existe un tramo activo del mismo tipo, le pone `FechaFinal` = hoy (`TipoTecnicoPrecio.cls:431-439`); en alta asigna ID calculado (`TipoTecnicoPrecio.cls:442-456`); escribe por DAO (`TipoTecnicoPrecio.cls:462-507`). Si falla, repone `FechaFinal = Null` del tramo cerrado (`TipoTecnicoPrecio.cls:511-515`).
4. Lanza `Alta` o `Edicion` y cierra (`:120-127`). Error: `MsgBox` crítico y el formulario sigue abierto (`:132-138`).

## Reglas, validaciones y cálculos

Validaciones de `MotivoNoOK` (`TipoTecnicoPrecio.cls:287-345`):

- Tipo técnico obligatorio y existente (`TipoTecnicoPrecio.cls:300-303`).
- Edición de un tramo no activo prohibida: «No se puede cambiar los datos de un tipo no activo» (`TipoTecnicoPrecio.cls:304-309`).
- Ambos precios deben ser numéricos (`TipoTecnicoPrecio.cls:310-317`).
- En alta no se admite fecha final (`TipoTecnicoPrecio.cls:318-323`).
- En edición, cerrar un tramo que estaba abierto (poner fecha final) se rechaza si algún técnico de ese tipo tiene actividades o eventos sin facturar (`TipoTecnicoPrecio.cls:324-335`; consultas `TipoTecnicoPrecio.cls:361-380`).
- Cálculo: ID = `DameID("TbTipoTecnicoPrecios", "IDTipoTecnicoFecha")` (`TipoTecnicoPrecio.cls:109`).
- No se valida que `FechaInicial` sea una fecha, que sea anterior a `FechaFinal` ni que no solape con otros tramos: el código del formulario y de `MotivoNoOK` no lo comprueba.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura por ID | `TbTipoTecnicoPrecios` | SQL dinámico concatenado (`WHERE IDTipoTecnicoFecha=` + texto) | — | `Constructor.bas:1994-1996` |
| Tramo activo del tipo | `TbTipoTecnicoPrecios` | SQL dinámico con el tipo entre comillas | — | `Constructor.bas:2211-2218` |
| Comprobación de pendientes de facturar | `TbTecnicos`, `TbActividades`, `TbEventos` | SQL dinámico con el tipo entre comillas | — | `TipoTecnicoPrecio.cls:361-380` |
| Cierre del tramo previo | `TbTipoTecnicoPrecios` | SQL dinámica con fecha `#mm/dd/yyyy#` | no | `TipoTecnicoPrecio.cls:433-436` |
| Alta | `TbTipoTecnicoPrecios` | DAO `AddNew`/`Update` sobre la tabla | no | `TipoTecnicoPrecio.cls:470-485` |
| Edición | `TbTipoTecnicoPrecios` | DAO `Edit`/`Update` sobre `WHERE IDTipoTecnicoFecha=` | no | `TipoTecnicoPrecio.cls:458-460, 486-501` |
| Compensación en error | `TbTipoTecnicoPrecios` | SQL dinámica `FechaFinal = Null` | no | `TipoTecnicoPrecio.cls:511-515` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: literal de fecha de Jet `#mm/dd/yyyy#` (`TipoTecnicoPrecio.cls:433`); identificador calculado en VBA con `DameID` fuera de transacción; varias escrituras (cerrar previo, insertar) sin transacción, con compensación manual; SQL concatenada con el tipo entre comillas simples.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- No hay transacción: cerrar el tramo previo y grabar el nuevo son operaciones separadas; si el proceso se interrumpe tras el cierre, la compensación (`TipoTecnicoPrecio.cls:511-515`) no se ejecuta (`TipoTecnicoPrecio.cls:433-440`).
- En edición del tramo activo, `Registrar` le pone antes `FechaFinal` = hoy (`TipoTecnicoPrecio.cls:431-437`, si `getTipoTecnicoActivo` lo devuelve) y luego la sobrescribe con el valor de la caja (`TipoTecnicoPrecio.cls:486-495`); si el segundo paso falla, la compensación lo deja abierto aunque el usuario hubiera indicado una fecha final.
- `HaHabidoCambios` compara cadenas con `<>` (`:17-31`) tras `Nz(..., "")`: `100` y `100,00` o dos formatos de la misma fecha se tomarían como cambio; con `Option Compare Database` (`:2`) las comparaciones de texto no distinguen mayúsculas.
- `TipoEnAlgunTecnicoPorFacturar` cierra y anula `rcdDatos` al detectar pendientes y vuelve a ejecutar `rcdDatos.Close` sobre `Nothing` (`TipoTecnicoPrecio.cls:366-373`): la rama «hay técnicos por facturar» podría acabar en error de objeto en lugar de devolver `Sí`.
- En `Registrar`, `m_Error` privado se rellena pero se informa `p_Error` (`TipoTecnicoPrecio.cls:446-454`), por lo que el mensaje «No se ha podido obtener el ID» podría no llegar al formulario.
- No hay validación de coherencia entre `FechaInicial` y `FechaFinal` ni de solapes (ver Reglas).
- `ComandoRegistrar_Click` escribe en la global `m_ObjTipoTecnicoPrecioActivo` antes de guardar (`:108-113`); si el guardado falla, queda modificada.
- El `Format` de moneda «€» está fijo en el diseño (`.form.txt:678, 709`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: cierre del tramo previo y alta en una sola transacción; validar fechas y no solape; identificador generado por la base de datos.

## Criterios de aceptación

Propuestos, no aprobados:

1. Alta desde la gestión: el tipo aparece relleno y bloqueado (`:71`; `.form.txt:427`).
2. Guardar un alta sin precios numéricos muestra «Se ha de indicar un precio de hora laboral/Extra» (`TipoTecnicoPrecio.cls:310-317`).
3. Guardar un alta con fecha final muestra «En el alta no se pude poner fecha final» (`TipoTecnicoPrecio.cls:318-323`).
4. Un alta cierra con fecha final de hoy el tramo activo del mismo tipo (`TipoTecnicoPrecio.cls:431-437`).
5. Editar sin cambiar nada muestra «Ningún cambio que registrar» (`:103-106`).
6. Cerrar un tramo con técnicos sin facturar del tipo es rechazado (`TipoTecnicoPrecio.cls:324-335`).
7. Tras guardar se lanza `Alta` o `Edicion` y el formulario se cierra (`:120-127`).

## Pendientes

- Si `DameID` (`TipoTecnicoPrecio.cls:109`) calcula `Max + 1`; no leído.
- Contenido de `getTipoTecnicoActivo` más allá de la consulta (`Constructor.bas:2200-2218`) y qué devuelve en edición del tramo activo.
- Índices o restricciones de `TbTipoTecnicoPrecios` (ERD de origen).
- Contenido de `AjustarTamaño` (`:151`) en lo que afecta a este formulario.
