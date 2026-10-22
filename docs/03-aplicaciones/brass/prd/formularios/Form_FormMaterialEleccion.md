# Form_FormMaterialEleccion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormMaterialEleccion.cls` (163 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Selector previo al alta de material: el usuario elige si es una reparación («Sí») o un repuesto («No») y el formulario abre `FormMaterialAltaReparacion` o `FormMaterialRepuestoAlta` con un argumento compuesto que describe el contexto (evento y actividad, o BUI, subsistema y equipo). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) y no escribe en la base de datos. Las citas son `Form_FormMaterialEleccion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEventoGestion` | Cierra el formulario si estaba abierto y lo abre con `OpenArgs` = `IDEvento\|IDActividad` | `Form_FormEventoGestion.cls:1206-1209` |

El resto de llamadores no aparecen en `src` (ver Pendientes). `Form_Open` (`:104-157`) admite tres formas de `OpenArgs` (comentario en `:112-114`): vacío (`:116-120`), dos valores `IDEvento|IDActividad` (`:126-132`) o cuatro valores `BUI|Subsistema|IDEquipo|IDEquipo` (`:133-141`); cualquier otro número de partes da «No se ha abierto el formulario con los datos necesarios» (`:122-125`).

Salida: abre `FormMaterialAltaReparacion` (`:58`) o `FormMaterialRepuestoAlta` (`:90`) con `OpenArgs` = `BUI|SUBSISTEMA|IDEquipo|<IDEquipo o Equipo>|IDEvento|IDActividad` (`:53, 89`) y cierra este formulario (`:92`). `cmdSalir` cierra sin guardar (`:6-9`).

## Controles y eventos

Controles: 17 (17 `Begin` en el `.form.txt`): el propio `Form` (1), la sección `Detalle` (1), siete plantillas de estilo por defecto (`Label`, `Rectangle`, `Image`, `CommandButton`, `OptionGroup`, `TextBox`, `ComboBox`) (7) y ocho controles reales.

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComboEsReparacion` | desplegable, lista de valores «Sí»; «No», limitado a la lista | — | Elección reparación/repuesto; se lee con `Nz` al pulsar «Alta» | CAP-7 Materiales, reparaciones y repuestos | `:18`; `.form.txt:386-419` |
| `ComandoAlta` («Alta») | botón | `OnClick` → `ComandoAlta_Click` | Valida la elección, resuelve el contexto y abre el formulario de alta | CAP-7 | `:11-103`; `.form.txt:420-444` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra sin guardar | CAP-7 | `:6-9`; `.form.txt:334-358` |
| `lblTitulo1`, `lblTitulo2` | etiquetas «TIPO MATERIAL» | — | `lblTitulo1` pasa a `Caption` del formulario | CAP-7 | `:143`; `.form.txt:445-490` |
| `Etiqueta210` («REPARACIÓN») | etiqueta | — | Rótulo del desplegable | CAP-7 | `.form.txt:359-385` |
| `CuadroSuperior`, `Cuadro2` | rectángulos | — | Decoración | sin asignar | `.form.txt:288-333` |

## Flujos

### Apertura (`Form_Open`, `:104-157`)

1. Cursor de espera y `AjustarTamaño Me` (`:107-110`).
2. Sin argumentos, quita el cursor y sale **sin** poner el título de ventana (`:116-120`).
3. Con dos valores guarda `strIDEvento` y `strIDActividad` y exige ambos no vacíos (`:126-132`); con cuatro guarda `strBUI`, `strSUBSISTEMA` e `strIDEquipo` (`:134-137`) y exige que no estén vacíos (`:138-141`).
4. Error: `MsgBox` crítico y cierre del formulario (`:150-156`).

### Alta (`ComandoAlta_Click`, `:11-103`)

1. Si `ComboEsReparacion` no es «Sí» ni «No»: «No se ha elegido sí o no» (`:18-22`). Parte `strArgumento` por `|` en la global `dato` (`:23`).
2. **Reparación («Sí»)** (`:24-58`): cierra `FormMaterialAltaReparacion` si está abierto (`:31-33`). Con dos partes lee BUI, subsistema, equipo y nombre del evento (`:34-51`), error «Ne está registrado el evento» si no existe (`:41-44`), y recompone el argumento (`:53`). Con cuatro partes recompone el argumento sin consultar (`:54-56`). Abre `FormMaterialAltaReparacion` (`:58`).
3. **Repuesto («No»)** (`:59-91`): lee el contexto del evento con la misma consulta (`:60-75`), cierra `FormMaterialRepuestoAlta` si está abierto (`:81-83`), exige dos partes: si no, «No venimos de una actividad» (`:84-87`), y abre `FormMaterialRepuestoAlta` con el nombre del equipo en la cuarta posición (`:89-90`).
4. Cierra este formulario (`:92`). Error: quita el cursor y `MsgBox` crítico; el formulario queda abierto (`:97-102`).

## Reglas, validaciones y cálculos

- Elección obligatoria entre «Sí» y «No» (`:19-22`).
- Reparación: admite contexto de actividad (evento + actividad) o de equipo (BUI + subsistema + equipo) (`:34-57`).
- Repuesto: solo se admite desde una actividad (`:84-87`).
- El evento debe existir en `TbEventos` (`:41-44, 65-68`).
- El argumento de salida tiene seis partes separadas por `|` (`:53, 89`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Contexto del evento (reparación, desde actividad) | `TbEventos` con `TbEquipos` (LEFT JOIN) | SQL dinámico concatenado con `IDEvento` entre comillas, `OpenRecordset` | — | `:35-39` |
| Contexto del evento (repuesto) | `TbEventos` con `TbEquipos` (LEFT JOIN) | Misma consulta | — | `:60-63` |

No hay escrituras. El formulario usa las globales `m_SQL` y `dato` (`Variables Globales.bas:203, 206`).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL concatenado con el identificador del evento entre comillas simples; `LEFT JOIN` estándar sin funciones Jet. La navegación entre formularios por argumento de texto separado por `|` no tiene equivalente directo en PostgreSQL.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Con cuatro valores, `strIDEquipo` se asigna dos veces y el segundo valor (cuarto) sobrescribe al tercero (`:136-137`); `strEquipo` no se asigna nunca en ese caso (`:134-137`), de modo que la cuarta posición del argumento se rellena con el identificador del equipo (`:56`).
- En reparación desde actividad, la cuarta posición del argumento lleva `strIDEquipo`, mientras que en repuesto lleva `strEquipo` (`:53` frente a `:89`); la intención no queda clara.
- Repuesto con cuatro valores o sin argumentos: `strIDEvento` está vacío, la consulta no encuentra registro y sale «Ne está registrado el evento: » (`:60-68`) antes de llegar al mensaje «No venimos de una actividad» (`:84-87`), que parece inalcanzable.
- Sin argumentos y «Sí»: `Split` de cadena vacía no cumple ninguna rama (`:34, 54`) y se abre `FormMaterialAltaReparacion` con argumento vacío (`:58`).
- Errata en el mensaje «Ne está registrado el evento» (`:42, 66`).
- Sin argumentos, el título de ventana no se asigna (`:116-120` frente a `:143`).
- En error, el recordset no se cierra y no se informa de cuál era el formulario de destino (`:97-102`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: sustituir el argumento textual por parámetros explícitos de navegación (evento, actividad, equipo).

## Criterios de aceptación

Propuestos, no aprobados:

1. Pulsar «Alta» sin elegir «Sí» o «No» muestra «No se ha elegido sí o no» y no abre nada (`:19-22`).
2. Desde una actividad con «Sí» se abre el alta de reparación con evento, actividad, BUI, subsistema y equipo del evento (`:34-58`).
3. Desde una actividad con «No» se abre el alta de repuesto con los mismos datos (`:60-90`).
4. Con un evento inexistente se muestra el mensaje de evento no registrado y no se abre nada (`:41-44, 65-68`).
5. Un argumento con un número de partes distinto de dos o cuatro cierra el formulario con aviso (`:122-125, 150-156`).
6. Tras abrir el destino, este formulario se cierra (`:92`).
7. `cmdSalir` cierra sin abrir nada (`:6-9`).

## Pendientes

- Otros llamadores del formulario (solo se localiza `Form_FormEventoGestion.cls:1209`; el caso de cuatro valores y el de sin argumentos no tienen origen visible en `src`).
- Cómo interpretan `FormMaterialAltaReparacion` (`Form_FormMaterialAltaReparacion.cls:366`) y `FormMaterialRepuestoAlta` (`Form_FormMaterialRepuestoAlta.cls:183`) las seis partes del argumento.
- Contenido de `AjustarTamaño` (`:110`) y de `FormularioAbierto` (`Funciones Generales.bas:171`).
