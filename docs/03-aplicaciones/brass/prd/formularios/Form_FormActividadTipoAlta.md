# Form_FormActividadTipoAlta

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormActividadTipoAlta.cls` (148 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta de un tipo de actividad no programada: pide un código (`CodTipo`) y una descripción (`TIPOACTIVIDAD`) y los inserta en la tabla `TbCodActividad` (`:121-128`). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`). Capacidad provisional: CAP-4 Actividades y tipos de actividad. `FG.bas` = `modules/Funciones Generales.bas`. Las citas son `Form_FormActividadTipoAlta.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormActividadTipoGestion` | `ComandoAlta_Click` ejecuta `DoCmd.OpenForm "FormActividadTipoAlta"` sin argumentos y con `On Error Resume Next` | `Form_FormActividadTipoGestion.cls:6-9` |

Al guardar con éxito, si `FormActividadTipoGestion` está abierto (`FormularioAbierto`, `FG.bas:171-193`) refresca su lista `ListaTiposDeActividad` con `Requery` (`:28-30`) y se cierra sin guardar el diseño (`:31`). No lanza eventos ni escribe globales. `cmdSalir` cierra sin guardar (`:44-47`).

## Controles y eventos

Controles: 14 (22 Begin en el .form.txt: 1 `Form`, 6 bloques de valores por defecto de control, 1 `Section` y 14 controles con `Name`).

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `CodTipo` | texto no enlazado | — | Código del tipo; se lee con `Nz` al registrar | `:10`, `.form.txt:496-507` |
| `TIPOACTIVIDAD` | texto no enlazado | — | Descripción del tipo; se lee con `Nz` | `:11`, `.form.txt:523-535` |
| `ComandoAlta` («Registrar») | botón | `OnClick` → `ComandoAlta_Click` | Valida, inserta y cierra | `:4-42`, `.form.txt:327-339` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra sin guardar | `:44-47`, `.form.txt:370-382` |
| `ComandoAyudaFormulario` («Comando46») | botón | `OnClick` = `[Event Procedure]` | **Sin handler** en el `.cls` | `.form.txt:414-425`; ausente en `:1-148` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | «Alta de Tipo de Actividad»; `lblTitulo1` pasa a `Caption` del formulario | `:54`, `.form.txt:551-581` |
| `Etiqueta92`, `Etiqueta118`, `Etiqueta1`, `Etiqueta2` | etiquetas | — | Rótulos «Tipo», «Cod» y dos «*» (obligatoriedad visual) | `.form.txt:254-266, 469-481, 615-646` |
| `CuadroSuperior`, `Cuadro2`, `Marco122` | rectángulos y grupo de opciones | — | Decoración | `.form.txt:281-313, 592-602` |

## Flujos

### Apertura (`Form_Load`, `:48-60`)

Con `On Error Resume Next`: reloj de arena, `AjustarTamaño Me` (`FG.bas:261-287`), título de ventana igual a `lblTitulo1` y repintado (`:49-59`). No hay `Form_Open` ni carga de datos.

### Registrar (`ComandoAlta_Click`, `:4-42`)

1. Lee `CodTipo` y `TIPOACTIVIDAD` con `Nz(…, "")` (`:10-11`) y llama a `Alta` (`:17`).
2. Si el resultado contiene `|`, lo trata como error y lo muestra (`:18-22`).
3. Refresca la gestión si está abierta y cierra (`:28-31`).
4. Error: oculta el reloj de arena y muestra `MsgBox` crítico; el formulario **queda abierto** (`:36-41`).

### `Alta` (función pública, `:61-142`)

Rechaza código vacío («No se ha indicado el código de la actividad») y descripción vacía (`:83-90`); comprueba unicidad de ambos con `Dame` (`:96-120`); inserta con DAO (`:121-128`) y devuelve «Se ha dado de Alta el Tipo de Actividad correctamente » (`:131`). Cualquier error devuelve `"#ERR|…"` (`:133-141`).

## Reglas, validaciones y cálculos

- **Código obligatorio** y **descripción obligatoria** (`:83-90`); no se recortan espacios.
- **Código único:** `Dame("TbCodActividad","CODIGO","CODIGO",código)` distinto de vacío ⇒ «Ya existe un código de actividad como ese» (`:96-105`).
- **Descripción única:** misma comprobación sobre `ACTIVIDADNOPROGRAMADA` ⇒ «Ya existe un tipo de actividad como esa» (`:111-120`).
- La descripción del formulario se guarda en el campo `ACTIVIDADNOPROGRAMADA` (`:126`).
- Sin comprobación de permisos de usuario en este formulario.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Existe código | `TbCodActividad` | SQL dinámico concatenado, vía `Dame` (primero sin comillas; si falla, con comillas) | — | `:96`, `FG.bas:298-312` |
| Existe descripción | `TbCodActividad` | ídem sobre `ACTIVIDADNOPROGRAMADA` | — | `:111`, `FG.bas:298-312` |
| Alta | `TbCodActividad` | DAO `OpenRecordset` de tabla, `AddNew`/`Update` | no | `:121-128` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: el SQL de `Dame` se compone por concatenación y se reintenta con comillas ante error (`FG.bas:298-312`), dependiente de que el motor rechace el valor textual sin comillas; la apertura de la tabla completa con DAO exigirá tabla vinculada con clave en PostgreSQL; la unicidad se comprueba en VBA, no en el motor.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComandoAyudaFormulario` tiene `OnClick = [Event Procedure]` (`.form.txt:414-425`) pero no existe `ComandoAyudaFormulario_Click` en el `.cls`: el clic no hace nada.
- Comprobar unicidad y luego insertar sin transacción ni bloqueo (`:96-128`): dos altas simultáneas con el mismo código podrían ambas pasar la validación.
- `Dame` concatena el valor sin escapar (`FG.bas:298-312`): un apóstrofo en el código o la descripción haría fallar ambas variantes de la consulta y el alta se rechazaría con el error del motor.
- Con `Option Compare Database` (`:2`) y comparaciones del motor, la unicidad podría ignorar mayúsculas; no se verifica el índice de la tabla.
- Mensaje de éxito de `Alta` con espacio final (`:131`) que el formulario descarta; el usuario no ve confirmación, solo el cierre (`:31`).
- Un valor numérico en `CodTipo` se interpreta en `Dame` como número sin comillas (`FG.bas:298-300`); si el campo es texto el motor podría compararlo de forma distinta (a verificar).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: restricción única en base de datos sobre código y descripción; mensaje de confirmación; validación de longitud.

## Criterios de aceptación

Propuestos, no aprobados:

1. Registrar con `CodTipo` vacío muestra «No se ha indicado el código de la actividad» y no escribe (`:83-86`).
2. Registrar con `TIPOACTIVIDAD` vacío muestra «No se ha indicado el tipo de actividad» y no escribe (`:87-90`).
3. Un código ya existente en `TbCodActividad` muestra «Ya existe un código de actividad como ese» (`:102-105`).
4. Una descripción ya existente muestra «Ya existe un tipo de actividad como esa» (`:117-120`).
5. Un alta válida añade una fila con ambos valores y cierra el formulario (`:125-127, 31`).
6. Tras un alta válida, la lista `ListaTiposDeActividad` de la gestión abierta incluye el nuevo tipo (`:28-30`).
7. Tras un error el formulario sigue abierto con los valores escritos (`:36-41`).

## Pendientes

- Estructura de `TbCodActividad` (tipos, índices únicos, longitudes): ERD de origen.
- Contenido de `Form_FormActividadTipoGestion.ListaTiposDeActividad` y de su evento `Requery` (`:29`).
- Propiedad `Name` del grupo `Marco122` y su uso (sin referencias en el `.cls`).
