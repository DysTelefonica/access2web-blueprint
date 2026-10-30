# Form_FormSubSistemaEdicion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormSubSistemaEdicion.cls` (288 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Renombra un subsistema de un BUI (`TbSubsistemaBui`: pareja `SUBSISTEMA`/`BUI`). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): recibe BUI y subsistema por `OpenArgs`, muestra el BUI de solo lectura y deja editar el nombre del subsistema. Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`. Las citas son `Form_FormSubSistemaEdicion.cls:línea` salvo indicación; capacidad de todos los controles: CAP-12 Estructura y catálogos.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSubSistemaGestion` | Cierra el formulario si ya estaba abierto y lo abre con `OpenArgs = BUI & "\|" & subsistema` seleccionados | `Form_FormSubSistemaGestion.cls:59-62` |

Salida: tras guardar, si `FormSubSistemaGestion` está abierto, hace `Requery` de su control `ListaSubSistema` (`:60-62`, `:107-109`) y cierra el formulario sin guardar diseño (`:68`, `:110`). No lanza eventos propios.

## Controles y eventos

Controles: 22 (22 `Begin` en el `.form.txt`): `Form` y `Section` (2), 7 definiciones de plantilla por defecto sin nombre (`.form.txt:141-244`) y 13 en la sección `Detalle`.

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `BUI` | texto, bloqueado | — | Muestra el BUI recibido; no editable | CAP-12 | `.form.txt:485-499, 486`; `:185` |
| `SubSistema` | texto no enlazado | — | Nombre nuevo del subsistema; se inicializa con el actual y es obligatorio (asterisco `Etiqueta0`) | CAP-12 | `.form.txt:601-614`; `:186` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` → `ComandoRegistrar_Click` | Guarda y cierra | CAP-12 | `:77-121`, `.form.txt:372-383` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Pregunta si guardar cuando hay cambios y cierra | CAP-12 | `:34-76`, `.form.txt:414-426` |
| `lblTitulo1`, `lblTitulo2` | etiquetas «Edición de SubSistema» | — | `lblTitulo1` da el `Caption` de la ventana | CAP-12 | `:166`; `.form.txt:515-545` |
| `Etiqueta92`, `Etiqueta118`, `Etiqueta0` | etiquetas (SUBSISTEMA, BUI, «*») | — | Rótulos | CAP-12 | `.form.txt:299-311, 458-470, 579-588` |
| `CuadroPrincipal`, `CuadroSuperior`, `Cuadro2`, `Marco122` | rectángulos y grupo de opciones | — | Decoración | CAP-12 | `.form.txt:276-285, 326-335, 348-358, 556-566` |
| Plantilla por defecto (7) y `Form`/`Section` (2) | — | — | Sin comportamiento | CAP-12 | `.form.txt:5, 141-265` |

## Flujos

### Apertura (`Form_Open`, `:159-200`)

1. Cursor de espera, `AjustarTamaño Me` (`FG.bas:261`) y `Caption` desde `lblTitulo1` (`:163-166`).
2. `OpenArgs` vacío o sin `|` ⇒ «No se han cargado todos los datos» (`:168-176`). Se parte por `|`: `dato(0)` = BUI, `dato(1)` = subsistema inicial; si alguno queda vacío, mismo error (`:177-184`).
3. Rellena `BUI` y `SubSistema` (`:185-186`).
4. Cualquier error: `MsgBox` crítico y cierre del formulario (`:193-199`).

### Guardado (`ComandoRegistrar_Click`, `:77-121`)

1. `HaHabidoCambios` compara `SubSistema` (con `Nz`) con el valor inicial (`:22-29`). Sin cambios: «No ha habido cambios que guardar» y sale sin cerrar (`:84-90`).
2. `Registrar` (`:122-158`) lee `SubSistema` y llama a `Edicion` (`:138-144`); un retorno con `|` es error (`:145-149`).
3. Refresca la lista de `FormSubSistemaGestion` si está abierta y cierra (`:107-110`).
4. Error: `MsgBox` crítico y el formulario queda abierto (`:115-120`).

### Salir (`cmdSalir_Click`, `:34-76`)

Con cambios pregunta «¿Desea guardar los cambios producidos?»; con «Sí» ejecuta `Registrar` y refresca la lista (`:38-62`); con «No» descarta. Siempre cierra al final (`:68`). Si `Registrar` falla, muestra el error y no cierra (`:70-75`).

## Reglas, validaciones y cálculos

- **Datos de entrada:** BUI, subsistema inicial y subsistema nuevo no vacíos (`:223-234`).
- **Existencia:** debe existir la pareja subsistema inicial/BUI en `TbSubsistemaBui`; si no, «No está registrado el Subsistema…» (`:235-245`).
- **Uso en eventos:** si el subsistema inicial aparece en `TbEventos.SUBSISTEMA` (`Dame`, `:253`) se rechaza: «…ha formado parte de algún evento» (`:259-262`). La búsqueda es solo por nombre, sin filtrar por BUI.
- **Escritura:** `Edit` sobre el campo `SUBSISTEMA` del registro encontrado (`:263-267`); resultado «Se ha editado correctamente el SubSistema» (`:270`).
- No hay comprobación de que el nombre nuevo no exista ya para ese BUI, aunque el comentario de cabecera lo enumera (`:212`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura de la pareja a editar | `TbSubsistemaBui` | SQL dinámico concatenado (subsistema y BUI entre comillas) con `OpenRecordset` | no | `:235-240` |
| Comprobación de uso | `TbEventos` | SQL dinámico concatenado vía `Dame`: primero sin comillas y, si falla, con comillas | no | `:253`; `FG.bas:290-310` |
| Renombrado | `TbSubsistemaBui` | DAO `Edit`/`Update` | no | `:263-267` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL concatenado con valores entre comillas simples (`:235-238`); `Dame` genera primero una comparación numérica sin comillas y recurre al segundo intento por captura de error (`FG.bas:296-309`), lo que dependerá del tipo de columna y de los mensajes de error del motor vinculado; `Update` DAO sobre un recordset abierto con SQL.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Un nombre con apóstrofo rompe las consultas concatenadas (`:237`, `FG.bas:309`).
- La comprobación de eventos no filtra por BUI: un subsistema homónimo usado en otro BUI bloquearía el renombrado (`:253`).
- No se valida duplicidad del nombre nuevo (`:212`, `:263-267`); depende de un posible índice único en la tabla (a verificar con el ERD).
- `Dame` con un valor que contenga apóstrofo o texto numérico puede dar resultados distintos según el intento (`FG.bas:296-309`).
- El mensaje de error de `ComandoRegistrar_Click` habla de «registrar el gasto» (`:117`), texto heredado de otro formulario.
- Los renombrados no se propagan a otras tablas que guarden el nombre del subsistema (solo se consulta `TbEventos`, `:253`); a verificar con el ERD.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: renombrado por identificador numérico del subsistema, con restricción única (BUI, nombre).

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin BUI o sin subsistema muestra «No se han cargado todos los datos» y cierra (`:168-184`).
2. El campo BUI no es editable (`.form.txt:486`).
3. Guardar sin cambiar el nombre muestra «No ha habido cambios que guardar» (`:84-90`).
4. Vaciar `SubSistema` y guardar muestra «No se ha indicado el nombre del SubSistema que queremos que tenga» (`:231-234`).
5. Renombrar un subsistema ya usado en algún evento se rechaza con el mensaje de uso (`:259-262`).
6. Renombrar uno no usado cambia el nombre y refresca `ListaSubSistema` de la gestión (`:263-267, 107-109`).
7. Salir con cambios pregunta y respeta Sí/No (`:38-66`).

## Pendientes

- Índices y claves de `TbSubsistemaBui` y `TbEventos` (ERD de origen).
- Contenido de `AjustarTamaño` más allá de la cabecera (`FG.bas:261`) y de `FormularioAbierto` (`FG.bas:171`).
- `pregunta`, `flag`, `dato` y `m_SQL` no están declaradas en el formulario (`:39, 49, 51, 235`); se asume que son globales de módulo, sin resolver.
