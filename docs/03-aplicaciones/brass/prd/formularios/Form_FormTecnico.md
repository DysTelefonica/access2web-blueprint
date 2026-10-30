# Form_FormTecnico

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormTecnico.cls` (270 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta y edición de un técnico (`TbTecnicos`: alias, tipo, nombre, teléfono, correo, fechas de alta y baja). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): lee el técnico de la global `m_ObjTecnicoActivo` (`VG.bas:245`) y persiste con `Tecnico.Registrar` (`classes/Tecnico.cls:218-327`). Abreviatura: `VG.bas` = `modules/Variables Globales.bas`. Las citas sin fichero son `Form_FormTecnico.cls:línea`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormTecnicoGestion` (alta) | `m_ObjTecnicoActivo = Nothing`, cierra el formulario si estaba abierto y lo abre; escucha sus eventos con `WithEvents` | `Form_FormTecnicoGestion.cls:15-24, 9` |
| `Form_FormTecnicoGestion` (edición) | `m_ObjTecnicoActivo = m_ObjTecnicoSeleccionado` y lo abre | `Form_FormTecnicoGestion.cls:94-100` |

Eventos hacia quien lo abrió: `Alta` (`:8`, lanzado en `:43`), `Editado` (`:9`, lanzado en `:46`) y `Borrado` (`:10`, **nunca lanzado** en este formulario). En alta y edición, la gestión refresca sus combos (`Form_FormTecnicoGestion.cls:580-596`). Tras guardar se cierra solo (`:48`); `cmdSalir` cierra sin guardar (`:62-65`).

## Controles y eventos

Controles: 24 (33 `Begin` en el `.form.txt`: 1 formulario, 7 plantillas de estilo por defecto, 1 sección y 24 controles).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ALIAS` | texto no enlazado, `Tag="DATO"` | — | Alias; obligatorio y único | CAP-9 Técnicos | `.form.txt:1008-1020`; `Tecnico.cls:148-151, 165-173, 182-189` |
| `Tipo` | combo lista de valores, `Tag="DATO"` | `OnExit` y `OnChange` declarados en el `.form.txt` sin procedimiento en el `.cls` | Tipo; se llena con los tipos con precio vigente | CAP-9 Técnicos | `.form.txt:1063-1087`; `:99-131` |
| `Nombre` | texto, `Tag="DATO"` | — | Nombre; obligatorio | CAP-9 Técnicos | `.form.txt:1185-1198`; `Tecnico.cls:156-159` |
| `Telefono` | texto, `Tag="DATO"` | — | Teléfono opcional | CAP-9 Técnicos | `.form.txt:1214-1227` |
| `Email` | texto, `Tag="DATO"` | — | Correo opcional | CAP-9 Técnicos | `.form.txt:1297-1310` |
| `FechaAlta` | texto, formato fecha corta, `Tag="DATO"` | — | Obligatoria | CAP-9 Técnicos | `.form.txt:1101-1115`; `Tecnico.cls:160-163` |
| `FechaBaja` | texto, formato fecha corta, `Tag="DATO"` | — | Opcional; no anterior a la de alta | CAP-9 Técnicos | `.form.txt:1353-1367`; `Tecnico.cls:175-180, 190-195` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` | Valida, guarda y cierra | CAP-9 Técnicos | `.form.txt:904-914`; `:12-60` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | CAP-9 Técnicos | `.form.txt:942-952`; `:62-65` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | «Alta de Técnico» o «Edición de <nombre>» | CAP-9 Técnicos | `.form.txt:1405-1415, 1383-1394`; `:147-150, 178-181` |
| 7 etiquetas de campo (`Etiqueta118`, `Etiqueta0`, `Etiqueta3`, `Etiqueta5`, `Etiqueta8`, `Etiqueta9`, `Etiqueta11`) | etiquetas | — | Rótulos de los campos | CAP-9 Técnicos | `.form.txt:981-1338` |
| 4 etiquetas «*» (`Etiqueta1`, `Etiqueta2`, `Etiqueta4`, `Etiqueta6`) | etiquetas | — | Marcas de campo obligatorio (a qué campo corresponde cada una, a verificar con el diseño) | CAP-9 Técnicos | `.form.txt:1426-1500` |
| `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decorativos | sin asignar | `.form.txt:858-890` |

## Flujos

### Apertura (`Form_Load`, `:68-97`)

`AjustarTamaño` (`:77`) y `EstablecerDatos` (`:133-187`). Llena el combo `Tipo` (`:142`). Sin técnico activo: rótulos de alta y crea un `Tecnico` con `ParaAlta = Sí` (`:146-153`). Con técnico activo: `ParaAlta = No` y relee el registro con `Constructor.getTecnico(Alias)` (`:155-156`); si no existe, «Técnico no reconocido en el Sistema» (`:160-163`). Copia a cada control con `Tag` «DATO» el valor de la propiedad homónima, solo si no está vacío (`:164-177`). Error: `MsgBox` crítico, **el formulario queda abierto** (`:91-96`).

### Guardado (`ComandoRegistrar_Click`, `:12-60`)

1. Sin técnico activo: «No se conoce el técnico activo» (`:21-24`).
2. Vuelca los controles `DATO` en `m_ObjTecnicoActivo` con `Nz` (`:25, 189-206`).
3. En edición, `HaHabidoCambios` compara siete campos con la copia inicial; sin cambios: «Ningún cambio que guardar» (`:29-37, 215-261`).
4. `Tecnico.Registrar` valida (`MotivoNoOK`) y escribe (`:38-41`).
5. Lanza `Alta` o `Editado` y cierra (`:42-48`).
6. Error: `MsgBox` crítico con `m_Error`, formulario abierto (`:54-59`).

## Reglas, validaciones y cálculos

- **Obligatorios:** `Alias`, `Tipo`, `Nombre` y `FechaAlta` (`Tecnico.cls:148-163`).
- **Alias único:** en alta, o en edición si cambia el alias, `ValorExistente("TbTecnicos","Alias",...)`; si existe, «El alias ya ha sido usado para otro técnico» (`Tecnico.cls:164-173, 182-189`).
- **Fecha de baja:** en alta, si es fecha, no puede ser anterior a la de alta; en edición solo se comprueba si el original no tenía baja (`Tecnico.cls:175-180, 190-195`). Mensaje: «La fecha de baja ha de ser posterior a la de alta».
- **Tipo registrado:** debe existir en `TbTipoTecnico`, «El Tipo de técnico no ha sido registrado previamente» (`Tecnico.cls:197-204`).
- **Opcionales:** teléfono, correo, fechas y DNI se escriben solo si no están vacíos (`Tecnico.cls:273-288, 294-309`).
- **Combo `Tipo`:** `SELECT TIPOTECNICO FROM TbTipoTecnicoPrecios WHERE FechaFinal Is Null`, ordenado (`:108-111`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Tipos del combo | `TbTipoTecnicoPrecios` | SQL estática DAO | — | `:108-112` |
| Releer técnico | `TbTecnicos` | SQL dinámico con el alias entre comillas | — | `Constructor.bas:1951-1953` |
| Alias y tipo existentes | `TbTecnicos`, `TbTipoTecnico` | SQL dinámico concatenado (`ValorExistente`) | — | `FG.bas:7197-7203`; `Tecnico.cls:166, 182, 197` |
| Alta | `TbTecnicos` | DAO `AddNew`/`Update` | no | `Tecnico.cls:253, 268-289` |
| Edición | `TbTecnicos` | DAO `Edit`/`Update` sobre `WHERE Alias=` concatenado | no | `Tecnico.cls:256-257, 290-310` |

La conexión sale de `getdb` (`VG.bas:548-575`), con contraseña literal en el código (`VG.bas:573`).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL concatenado con comillas simples sobre el alias; escritura DAO por tabla (portable si se vincula a PostgreSQL, con fechas como texto asignado a campos de fecha, `Tecnico.cls:281-284`); sin transacción; la clave natural es el alias (no hay autonumérico).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La rama de edición no escribe `ALIAS` (`Tecnico.cls:291-309`) y localiza el registro por el alias inicial (`Tecnico.cls:256-257`): un alias modificado pasa la validación de unicidad (`Tecnico.cls:165-173`) pero no se persistiría.
- En edición, vaciar `FechaBaja`, `Telefono` o `Email` no los borra: solo se asignan si no están vacíos (`Tecnico.cls:294-306`).
- El combo se llena de `TbTipoTecnicoPrecios` (`:108-111`) y la validación consulta `TbTipoTecnico` (`Tecnico.cls:197`): dos tablas distintas para el mismo concepto.
- `HaHabidoCambios` no compara `DNI` (`:223-254`), que `Registrar` sí escribe (`Tecnico.cls:286-288, 307-309`); el formulario no tiene control de DNI.
- `FechaAlta` y `FechaBaja` se pasan a `CDate` sin comprobar `IsDate` en la rama de edición (`Tecnico.cls:176`): una fecha inválida acabaría en el mensaje genérico de error.
- `Tipo` declara `OnExit` y `OnChange` sin procedimiento en el `.cls` (`.form.txt:1085-1086`).
- `Borrado` se declara y nunca se lanza (`:10`), aunque la gestión lo escucha (`Form_FormTecnicoGestion.cls:603`).
- Se vuelca el formulario en la global antes de guardar (`:25`); si el guardado falla, la global queda modificada.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: unificar el origen de los tipos de técnico y permitir blanquear campos opcionales.

## Criterios de aceptación

Propuestos, no aprobados:

1. Guardar con alias, tipo, nombre o fecha de alta vacíos muestra «Obligatorio …» y no escribe (`Tecnico.cls:148-163`).
2. Dar de alta un alias ya existente muestra «El alias ya ha sido usado para otro técnico».
3. Una fecha de baja anterior a la de alta se rechaza con el mensaje de la regla.
4. Guardar en edición sin cambios muestra «Ningún cambio que guardar».
5. El combo `Tipo` ofrece solo los tipos con precio vigente (`:108-111`).
6. Tras guardar, la gestión de técnicos refresca su combo (`Form_FormTecnicoGestion.cls:580-596`).

## Pendientes

- Efecto real de `NotDefault` en `LimitToList` del combo `Tipo` (`.form.txt:1064`).
- Si `TbTecnicos` tiene índice único sobre `Alias` y longitudes de campo (ERD de origen).
- Contenido de `AjustarTamaño` (`:77`) y de `ValorExistente` completo (`FG.bas:7190`).
