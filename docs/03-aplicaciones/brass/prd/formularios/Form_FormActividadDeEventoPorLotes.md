# Form_FormActividadDeEventoPorLotes

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormActividadDeEventoPorLotes.cls` (383 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta masiva de actividades para el evento activo: el usuario indica tipo de actividad, ubicación, descripción, horas totales, fecha del primer día y una lista de técnicos, y el formulario pide al evento (`Evento.GenerarActividadesPorLotes`) que reparta las horas entre los técnicos en los días laborables disponibles. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); no escribe él mismo en la base: delega en `Evento.cls`. Las citas son `Form_FormActividadDeEventoPorLotes.cls:línea` salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`. Capacidad de todo el formulario: CAP-3 Eventos (alta masiva) y CAP-4 Actividades.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEventoGestion` (`ComandoAltaActividadesPorLotes_Click`) | Exige evento seleccionado, asigna `m_ObjEventoActivo = m_ObjEventoSeleccionado`, cierra el formulario si estaba abierto y lo abre sin `OpenArgs` | `Form_FormEventoGestion.cls:152-166` |

Entrada: solo la global `m_ObjEventoActivo` (`VG.bas:244`). Salida: al terminar refresca `ListaEventos` de `FormEventoGestion` y `ListaFiltrados` de `FormActividadGestion` si están abiertos y se cierra solo (`:280-293`). `cmdSalir` cierra sin guardar (`:303-306`).

## Controles y eventos

Controles: 21 (30 `Begin` en el `.form.txt`: 21 controles + `Form` + `Section` + 7 plantillas de estilo por defecto).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComboActividad` | combo 2 columnas, lista de valores | — | Código de actividad; se rellena en `Form_Load` desde `TbCodActividad`; obligatorio al registrar | CAP-4 | `:362-364`, `:225-228`, `.form.txt:1152-1186` |
| `ComboUbicacion` | combo lista de valores | — | Ubicación; se rellena desde `TbUbicacion`; obligatoria | CAP-12 | `:365-367`, `:229-232`, `.form.txt:1115-1151` |
| `Descripcion` | texto | — | Descripción común; obligatoria | CAP-4 | `:220-224`, `.form.txt:867-895` |
| `HorasTotales` | texto, formato número general | — | Horas a repartir; numérico y distinto de 0 | CAP-6 | `:233-240`, `.form.txt:1054-1084` |
| `FechaPrimeraActividad` | texto, fecha corta | — | Primer día a considerar; debe ser fecha | CAP-6 | `:245-248`, `.form.txt:1085-1114` |
| `ComboTecnico` | combo 3 columnas (alias, nombre, tipo), con encabezados | `AfterUpdate` → `ComboTecnico_AfterUpdate` | Habilita `ComandoAltaTecnico` solo si hay valor | CAP-9 | `:308-314`, `:358`, `.form.txt:1365-1403` |
| `ComandoAltaTecnico` (rótulo «Registrar») | botón, inicia deshabilitado | `OnClick` | Añade el técnico elegido a `ListaTecnicos` y a `m_ColTecnicos` | CAP-9 | `:159-177`, `:51-104`, `.form.txt:1404-1455` |
| `ComandoEliminarTecnico` (rótulo «Registrar») | botón | `OnClick` | Quita de la lista y del diccionario los técnicos seleccionados | CAP-9 | `:179-197`, `:106-156`, `.form.txt:1456-1510` |
| `ListaTecnicos` | lista 2 columnas con encabezados | — | Técnicos del lote; la fila 0 es el encabezado (`RowSource = "ALIAS;NOMBRE"`) | CAP-9 | `:368`, `.form.txt:1511-1560` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` | Valida, confirma, genera las actividades y cierra | CAP-3 | `:199-301`, `.form.txt:977-1014` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | CAP-3 | `:303-306`, `.form.txt:1015-1053` |
| `Detalle` (sección) | sección | `OnDblClick` → `Detalle_DblClick` | Rutina de pruebas: rellena los campos con valores fijos y da de alta un técnico | sin asignar | `:316-327`, `.form.txt:857-858` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título «ALTA DE ACTIVIDADES POR LOTES DE <evento>»; también `Caption` del formulario | CAP-3 | `:354-357`, `.form.txt:1295-1337` |
| `Etiqueta209`, `Etiqueta210`, `lbl1`, `Etiqueta214`, `Etiqueta241` | etiquetas | — | Rótulos de campo | sin asignar | `.form.txt:1187-1294, 1338-1364` |
| `Numero_Actividad` | texto bloqueado y deshabilitado | — | Residuo de diseño: `ControlSource` `IDActividad` y valor por defecto con `DCount` sobre `Tabla_Parte_Actividad`; el código no lo usa | sin asignar | `.form.txt:896-930` |
| `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decorativos | sin asignar | `.form.txt:931-976` |

## Flujos

### Apertura (`Form_Load`, `:329-382`)

1. Cursor de espera y `AjustarTamaño Me` (`:337-339`).
2. Sin `m_ObjEventoActivo`: «No se sabe el evento activo» (`:342-345`). Si `FechaFacturado` no está vacía o `Franqueado` es verdadero, error y cierre (`:346-353`).
3. Rotula con `UCase(IDEVENTO)` (`:354-357`), carga `ComboTecnico` con `RellenarComboTecnicos(…, Sí, No)` (`:358`; `FG.bas:7224-7260`) y los otros dos combos con `RellenaCombo` (`:362-367`; `FG.bas:1149-1191`).
4. Cualquier error: `MsgBox` crítico y cierre (`:375-382`).

### Alta de técnico y baja

`ComandoAltaTecnico_Click` sale sin hacer nada si el combo está vacío (`:163-165`); `AddTecnico` ignora alias ya presentes (`:68-75`), carga el técnico con `Constructor.getTecnico` y, si no existe, «No existe un técnico con ese alias» (`:76-83`); añade la fila `Alias;Nombre;Tipo` y el objeto al diccionario (`:85-94`). `DeleteTecnico` no hace nada sin filas o sin selección (`:121-128`) y borra las filas seleccionadas (`:144-146`).

### Registro (`ComandoRegistrar_Click`, `:199-301`)

1. Validaciones en este orden: descripción, actividad, ubicación, horas totales numéricas y distintas de 0, al menos un técnico, fecha válida (`:221-248`).
2. Crea `Actividad` con actividad, ubicación, descripción y `Observaciones = "De Lote"` (`:249-255`).
3. Pregunta Sí/No: texto distinto con un técnico o con varios (`:256-266`); «No» sale sin cambios.
4. Llama a `m_ObjEventoActivo.GenerarActividadesPorLotes` (`:271`; `Evento.cls:2100-2132`) y, si devuelve error, lo muestra (`:272-274`, `:295-300`).
5. Refresca los formularios abiertos y cierra (`:280-293`).

### Generación en `Evento.cls`

`RellenarTbActividadesLotes` vacía `TbActividadesLotes`, calcula las actividades y las graba una a una con DAO (`Evento.cls:2061-2087`); luego `GenerarActividadesPorLotes` las copia a `TbActividades` con `INSERT … SELECT` (`Evento.cls:2118-2124`). Las horas por técnico son `horas totales / nº de técnicos` (`Evento.cls:2203`); `getColFechasHoras` recorre hasta 10 001 días desde la fecha inicial, toma las horas libres de cada día (tope 7,5 menos las horas laborables ya registradas) y asigna el resto al último día (`Evento.cls:2340-2364`, `2295`). El identificador de cada actividad es `IDActividadCalculada` (`Evento.cls:2074`; `Actividad.cls:375-383`; `FG.bas:6535-6580`).

## Reglas, validaciones y cálculos

- **Evento editable:** no facturado y no franqueado, comprobado al abrir (`:346-353`) y de nuevo en `Evento.cls:2179-2186`.
- **Obligatorios:** actividad, ubicación, descripción, horas totales ≠ 0, fecha y al menos un técnico (`:221-248`); el evento repite las comprobaciones (`Evento.cls:2158-2199`).
- **Reparto:** igualitario entre técnicos; cada técnico ocupa días consecutivos hasta agotar sus horas; los días no laborables se saltan (`Evento.cls:2203-2206`, `2282-2288`).
- **Tope diario:** 7,5 horas literales en el código (`Evento.cls:2295`).
- **Nomenclatura del identificador:** alias, fecha de la actividad, hora y minuto actuales y un correlativo de dos cifras (`FG.bas:6560-6573`).

## Lecturas y escrituras

| Operación | Objeto | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Cargar actividades | `TbCodActividad` | SQL fija vía `RellenaCombo` | — | `:362-364` |
| Cargar ubicaciones | `TbUbicacion` | SQL fija vía `RellenaCombo` | — | `:365-367` |
| Técnico por alias | técnicos | `Constructor.getTecnico` | — | `:76`, `:28` |
| Horas del técnico en un día | `TbActividades` | SQL dinámico concatenado con fecha `#mm/dd/yyyy#` y alias | no | `FG.bas:549-552` |
| Vaciar el área de trabajo | `TbActividadesLotes` | `DELETE *` ejecutado con `CurrentDb().Execute` | no | `Evento.cls:2061-2062` |
| Escribir el lote | `TbActividadesLotes` | DAO `AddNew`/`Update` | no | `Evento.cls:2064, 2076-2087` |
| Pasar a definitivas | `TbActividades` | SQL fija `INSERT … SELECT` | no | `Evento.cls:2118-2124` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: fecha como literal `#mm/dd/yyyy#` (`FG.bas:551`) y `DELETE *` (`Evento.cls:2061`), propios del dialecto Jet/ACE; `INSERT … SELECT` entre tablas vinculadas; identificador textual calculado en VBA, no autonumérico; tres escrituras sin transacción explícita sobre la tabla de lote compartida.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La tabla `TbActividadesLotes` es única y se vacía entera en cada ejecución (`Evento.cls:2061-2062`): dos usuarios generando lotes a la vez podrían mezclar o perder filas; un fallo entre el vaciado y el `INSERT … SELECT` deja la tabla de lote con residuos.
- `Observaciones = "De Lote"` se asigna en el formulario (`:254`), pero `getActividadesPorLotes` no la copia a las actividades generadas (`Evento.cls:2233-2239`) ni `RellenarTbActividadesLotes` la escribe (`Evento.cls:2078-2085`), aunque el `INSERT` sí la selecciona (`Evento.cls:2118-2123`): podría llegar vacía.
- `IDActividadCalculada` devuelve el texto de error `ERR|…` de `DameIDActividad` sin comprobarlo (`Actividad.cls:381-383`; `FG.bas:6586`), y `RellenarTbActividadesLotes` lo graba como identificador (`Evento.cls:2074-2078`).
- El reparto divide las horas entre técnicos sin redondeo (`Evento.cls:2203`); si un técnico no tiene días libres, `getColFechasHoras` devuelve `Nothing` y sus horas se descartan sin aviso (`Evento.cls:2210-2218`), por lo que el total generado podría ser inferior al pedido.
- Código inerte: `RellenarCole` no tiene llamadores y su bloque de diccionario nunca añade nada (`:11-49`, `:32-39`); `m_ObjTecnicoSeleccionado` no se usa (`:6`).
- El diálogo de confirmación usa el texto «las <n> horas» sin unidades y la rutina `Detalle_DblClick` queda accesible a cualquier usuario (`:257-262`, `:316-327`).
- Rótulos de diseño: `ComandoAltaTecnico` y `ComandoEliminarTecnico` se rotulan ambos «Registrar» (`.form.txt:1414, 1466`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: generar el lote en memoria o en una tabla por sesión y confirmar todo en una única transacción; mostrar una previsualización del reparto antes de grabar; parametrizar el tope diario de 7,5 horas.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir el formulario sin evento activo, o con evento facturado o franqueado, muestra el error correspondiente y lo cierra.
2. Registrar con cualquiera de los seis datos obligatorios vacíos muestra su mensaje y no escribe.
3. Con un técnico, la confirmación habla de «el técnico seleccionado»; con varios, indica cuántos.
4. Añadir dos veces el mismo técnico deja una sola fila.
5. Con N técnicos y H horas, cada técnico recibe H/N horas repartidas en días laborables con tope de 7,5 h por día y técnico.
6. Tras registrar, el evento y la lista de actividades abiertos se refrescan.
7. Contestar «No» a la confirmación no crea actividades.

## Pendientes

- Contenido de `Constructor.getTecnico` y de `DiaLaborableParaTecnico` (`Evento.cls:2282`) para el criterio de día laborable.
- Estructura y clave de `TbActividadesLotes` y de `TbActividades` (ERD de origen).
- Si `Observaciones` llega a grabarse (`:254`; `Evento.cls:2078-2085`).
- Contenido de `AjustarTamaño` (`:339`).
