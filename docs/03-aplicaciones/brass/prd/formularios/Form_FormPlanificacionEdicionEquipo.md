# Form_FormPlanificacionEdicionEquipo

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormPlanificacionEdicionEquipo.cls` (287 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Edición del alias y de la periodicidad de mantenimiento recomendada (en meses) de un equipo ya marcado como planificable (`TbPlanificacionEquipos`, con el nombre del equipo leído de `TbEquipos`). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): recibe el `IDEquipo` por `OpenArgs`, carga los datos por DAO y los guarda con `EditarEquipo`. Las citas son `Form_FormPlanificacionEdicionEquipo.cls:línea` salvo indicación; capacidad de todos los controles: CAP-6 Planificación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormPlanificacionEquipos` | Cierra el formulario si ya estaba abierto y lo abre con `OpenArgs = strIDEquipoSeleccionado` | `Form_FormPlanificacionEquipos.cls:207-210` |

No lanza eventos ni refresca el formulario que lo abrió: tras guardar solo se cierra (`:27`, `:72`).

## Controles y eventos

Controles: 21 (21 `Begin` en el `.form.txt`): `Form` y `Section` (2), 6 definiciones de plantilla por defecto sin nombre (`.form.txt:141-223`) y 13 en la sección `Detalle`.

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComboEquipoPlanificable` | texto bloqueado | — | Muestra el nombre del equipo; no editable | CAP-6 | `.form.txt:531-545, 532`; `:271` |
| `Alias` | texto no enlazado | — | Alias opcional del equipo | CAP-6 | `.form.txt:420-431`; `:272` |
| `PERIODICIDADMESES` | texto no enlazado | — | Periodicidad en meses; debe ser ≥ 1 | CAP-6 | `.form.txt:501-514`; `:273`, `:210-213` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` → `ComandoRegistrar_Click` | Guarda y cierra | CAP-6 | `:5-38`, `.form.txt:301-311` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Pregunta si guardar cuando hay cambios y cierra | CAP-6 | `:85-119`, `.form.txt:327-337` |
| `ComandoAyudaFormulario` (rótulo «Comando46») | botón | `OnClick = [Event Procedure]` | Sin procedimiento `ComandoAyudaFormulario_Click` en el `.cls` | CAP-6 | `.form.txt:354-363` |
| `lblTitulo1`, `lblTitulo2` | etiquetas «EDICIÓN DE EQUIPO PARA MTO PREVENTIVO» | — | `lblTitulo1` da el `Caption` de la ventana | CAP-6 | `:127`; `.form.txt:561-591` |
| `Etiqueta235`, `Etiqueta199`, `Etiqueta200` | etiquetas (EQUIPO, ALIAS, PERIODICIDAD…) | — | Rótulos | CAP-6 | `.form.txt:393-405, 447-459, 474-486` |
| `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | CAP-6 | `.form.txt:255-264, 277-287` |
| Plantilla por defecto (6) y `Form`/`Section` (2) | — | — | Sin comportamiento | CAP-6 | `.form.txt:5, 141-244` |

## Flujos

### Apertura (`Form_Load`, `:120-157`)

1. Cursor de espera, `AjustarTamaño Me` y `Caption` desde `lblTitulo1` (`:124-127`).
2. `OpenArgs` no numérico ⇒ «El formulario se ha abierto con datos insuficientes» (`:128-132`).
3. `EstablecerDatos` (`:248-285`): `SELECT` con `INNER JOIN` a `TbEquipos` filtrado por `IDEquipo`; si no hay fila, «Equipo planificable no registrado» (`:256-264`). Guarda alias y periodicidad iniciales y rellena los tres controles (`:265-273`).
4. Cualquier error: `MsgBox` crítico y cierre (`:150-156`).

### Guardado (`ComandoRegistrar_Click`, `:5-38`)

1. `HaHabidoCambios` compara alias y periodicidad con los iniciales como texto (`:172-180`). Sin cambios: «Ningún cambio que grabar» como error (`:12-15`).
2. `Registrar` (`:39-83`) exige `IDEquipo` numérico (`:55-58`), toma `Alias` (`"#"` si es nulo) y `PERIODICIDADMESES` y llama a `EditarEquipo` con `CInt` (`:59-66`). Tras el éxito cierra el formulario (`:72`).
3. Cierra de nuevo desde el manejador (`:27`).
4. Error en `Registrar`: se muestra un `MsgBox` desde `Registrar` (`:77-82`), que devuelve cadena vacía (no `#ERR\|`) y el manejador cierra igualmente (`:21-27`).

### Salir (`cmdSalir_Click`, `:85-119`)

Con cambios pregunta «¿Desea guardar los cambios producidos?»; con «Sí» ejecuta `Registrar` (`:89-105`); con «No» descarta. Cierra siempre al final (`:111`).

## Reglas, validaciones y cálculos

- **Equipo numérico:** `IDEquipo` ha de ser numérico (`:55-58`, `:206-209`).
- **Periodicidad:** `intNumeroMesesMto < 1` ⇒ «La peridicidad es en meses» (`:210-213`). Un valor no entero hace fallar `CInt` (`:66`).
- **Alias:** vacío ⇒ no se modifica el campo; `"#"` (alias nulo en el formulario) ⇒ se guarda `Null` (`:224-230`). Como `Registrar` convierte `Null` en `"#"` (`:59`), vaciar el alias lo borra.
- **Existencia:** el registro de planificación del equipo debe existir; `EditarEquipo` no lo crea (`:214-222`, cabecera `:195`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Carga de datos | `TbPlanificacionEquipos` ⨝ `TbEquipos` | SQL dinámico concatenado (`IDEquipo` numérico sin comillas) | no | `:256-259` |
| Lectura para edición | `TbPlanificacionEquipos` | SQL dinámico concatenado (`SELECT .*`) | no | `:214-217` |
| Actualización | `TbPlanificacionEquipos` (`Alias`, `PeriodicidadEnMesesRecomendada`) | DAO `Edit`/`Update` | no | `:223-232` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL concatenado con `INNER JOIN` y `SELECT .*` (portables a PostgreSQL con adaptación de nombres); `Update` DAO sobre recordset abierto por SQL; `Null` frente a cadena vacía en `Alias` (`:228`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Con `PERIODICIDADMESES` vacío o no numérico, `CInt` falla (`:66`); el error lo captura `Registrar` (`:77-82`), que muestra el mensaje y devuelve cadena vacía, tras lo cual `ComandoRegistrar_Click` cierra el formulario sin haber guardado (`:21-27`).
- Doble cierre del formulario: `Registrar` lo cierra (`:72`) y `ComandoRegistrar_Click` lo vuelve a intentar (`:27`).
- Con alias previo y el campo vaciado, `Nz(Me.Alias, "#")` borra el alias (`:59`, `:227-229`); si el alias no cambia pero sí la periodicidad, un alias previo vacío se envía como `"#"` y se guarda `Null` (`:224-230`).
- «Ningún cambio que grabar» se trata como error del manejador (`:12-15`), a diferencia del resto de formularios análogos.
- Errata en el mensaje «peridicidad» (`:211`).
- `ComandoAyudaFormulario` declara `OnClick` sin procedimiento en el `.cls` (`.form.txt:363`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: validar la periodicidad como entero positivo antes de guardar y refrescar la lista de equipos planificables al cerrar.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir con un identificador no numérico muestra «El formulario se ha abierto con datos insuficientes» y cierra (`:128-132`).
2. Abrir con un equipo sin fila de planificación muestra «Equipo planificable no registrado» y cierra (`:261-264`).
3. El nombre del equipo se muestra y no es editable (`.form.txt:532`; `:271`).
4. Guardar sin cambios muestra «Ningún cambio que grabar» (`:12-15`).
5. Periodicidad menor que 1 se rechaza con el mensaje de periodicidad (`:210-213`).
6. Guardar alias y periodicidad válidos los actualiza en `TbPlanificacionEquipos` y cierra (`:223-232`).
7. Vaciar el alias lo deja nulo (`:59`, `:227-229`).
8. Salir con cambios pregunta y respeta Sí/No (`:89-111`).

## Pendientes

- Tipo, longitud y restricciones de `Alias` y `PeriodicidadEnMesesRecomendada` (ERD de origen).
- Contenido de `AjustarTamaño` (`:126`).
- Si existe un manejador de ayuda en otro módulo para `ComandoAyudaFormulario` (`.form.txt:363`).
- `pregunta`, `flag`, `dato` y `m_SQL` no se declaran en el formulario (`:37, 21, 23, 214`); se asumen globales, sin resolver.
