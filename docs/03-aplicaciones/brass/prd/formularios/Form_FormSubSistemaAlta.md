# Form_FormSubSistemaAlta

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormSubSistemaAlta.cls` (146 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Asocia un subsistema existente a un BUI en `TbSubsistemaBui`. Es un formulario no enlazado: recibe el BUI en `OpenArgs`, lo guarda en la variable de módulo `strBUI` (`Form_FormSubSistemaAlta.cls:4, 57`) y persiste con la función pública `Alta` (`:84-139`). Abreviatura: `FG.bas` = `modules/Funciones Generales.bas`. Las citas son `Form_FormSubSistemaAlta.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSubSistemaGestion` | Exige `ComboBui` informado, cierra el formulario si estaba abierto y lo abre pasando el BUI como `OpenArgs` | `Form_FormSubSistemaGestion.cls:10-25` |

Salida: tras guardar, si `FormSubSistemaGestion` está abierto llama a su `ActualizarLista` (`:28-30`) y se cierra (`:31`). `cmdSalir` cierra sin guardar (`:44-47`). No lanza eventos propios.

## Controles y eventos

Controles: 22 (22 Begin en el .form.txt: 1 formulario, 7 plantillas de control por defecto, 1 sección `Detalle` y 13 controles).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `BUI` | texto bloqueado | — | Muestra el BUI recibido (`Me.BUI = strBUI`) | CAP-12 | `:62`, `.form.txt:492-520` |
| `SubSistema` | combo | — | Subsistema a asociar; su `RowSource` se reasigna en la carga con los subsistemas que no pertenecen al BUI | CAP-12 | `:63-69`, `.form.txt:522-553` |
| `ComandoAlta` («Registrar») | botón | `OnClick` → `ComandoAlta_Click` | Da de alta | CAP-12 | `:5-42`, `.form.txt:378-419` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra sin guardar | CAP-12 | `:44-47`, `.form.txt:421-463` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título «Alta de SubSistema»; `lblTitulo1` da el título de la ventana | CAP-12 | `:56`, `.form.txt:555-594` |
| `Etiqueta92`, `Etiqueta118` | etiquetas | — | Rótulos «SUBSISTEMA» y «BUI» | CAP-12 | `.form.txt:305-330, 465-490` |
| `Etiqueta0` | etiqueta «*» | — | Marca de obligatorio | CAP-12 | `.form.txt:619-640` |
| `Marco122` | grupo de opciones | — | Marco decorativo | CAP-12 | `.form.txt:596-617` |
| `CuadroPrincipal`, `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | CAP-12 | `.form.txt:282-303, 332-376` |

## Flujos

### Apertura (`Form_Load`, `:48-83`)

1. Ajusta el tamaño (`AjustarTamaño`, `FG.bas:261`) y copia el título a `Me.Caption` (`:55-56`).
2. Sin `OpenArgs`: «No se ha abierto el formulario con los parámetros adecuadosW», `MsgBox` y cierre (`:57-61, 76-82`).
3. Rellena `BUI`, da foco a `SubSistema` y reasigna su `RowSource` con los subsistemas distintos de las filas cuyo BUI no es el actual (`:62-69`).

### Alta (`ComandoAlta_Click`, `:5-42`)

1. Lee `SubSistema` y llama a `Alta` (`:11, 17`). Si el resultado contiene «|» muestra el error (`:18-22`).
2. Actualiza la lista de gestión si está abierta y cierra (`:28-34`).
3. Error: `MsgBox` crítico y el formulario queda abierto (`:36-42`).

## Reglas, validaciones y cálculos

- BUI vacío ⇒ «No se ha indicado el nombre del BUI» (`:103-106`); subsistema vacío ⇒ «No se ha indicado el nombre del SubSistema» (`:107-110`).
- Unicidad: si ya existe la pareja subsistema + BUI ⇒ «Ya existe un subsistema llamado …» (`:111-120`).
- No hay identificador: la clave es la pareja de texto (`:122-123`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de subsistemas | `TbSubsistemaBui` | SQL dinámico con el BUI concatenado, asignado al `RowSource` | — | `:64-69` |
| Comprobación de duplicado | `TbSubsistemaBui` | SQL dinámico con subsistema y BUI entre comillas | no | `:111-115` |
| Alta | `TbSubsistemaBui` | DAO `AddNew`/`Update` | no | `:121-125` |

## Coexistencia

Adaptar a PostgreSQL vinculado: `SELECT DISTINCT` y `Not (…='…')` en dialecto Access (`:64-67`); literales entre comillas simples concatenados; clave compuesta de texto sin autonumérico.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `WHERE Not (BUI='x')` (`:66`) excluye filas del BUI actual, pero un subsistema presente en este BUI y en otro seguiría apareciendo en la lista por la otra fila; la regla de duplicado de `Alta` (`:111-120`) lo rechazaría.
- Un BUI o subsistema con apóstrofo rompería el SQL concatenado (`:66, 113-114`).
- Mensaje de error de carga con «W» sobrante (`:59`) y mensaje de `ComandoAlta_Click` que habla de «Eliminar» (`:20`).
- Comprobación y alta no son atómicas (`:115-125`).
- El BUI no se valida contra su tabla; depende del llamador.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: restricción única sobre la pareja subsistema + BUI y selector de BUI cerrado.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin BUI muestra el error y cierra el formulario (`:57-61, 76-82`).
2. Abierto con un BUI, el campo `BUI` lo muestra bloqueado y el combo no ofrece subsistemas de ese BUI (`:62-69`).
3. Registrar sin subsistema muestra «No se ha indicado el nombre del SubSistema» (`:107-110`).
4. Registrar una pareja ya existente muestra «Ya existe un subsistema llamado …» (`:116-120`).
5. Un alta válida crea el registro, refresca la gestión si estaba abierta y cierra (`:121-128`, `:28-31`).

## Pendientes

- Si el combo admite texto fuera de la lista (`LimitToList` no aparece en `.form.txt:522-553`).
- Estructura y claves de `TbSubsistemaBui` (ERD de origen).
- Contenido de `ActualizarLista` de `FormSubSistemaGestion` (`:29`).
