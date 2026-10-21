# Form_FormOficinaPrograma

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormOficinaPrograma.cls` (198 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta y edición de un registro del catálogo de oficinas de programa (`TbOficinasPrograma`): nombre y descripción. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); el modo se decide por la global `m_ObjOficinaProgramaActiva` (`Variables Globales.bas:260`): vacía, alta; informada, edición (`:130-135`). Sigue el mismo patrón que [Form_FormEjercito](Form_FormEjercito.md); esta ficha cita sus propias líneas y destaca las diferencias. Las citas son `Form_FormOficinaPrograma.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormOficinasProgramaGestion` (Alta) | Vacía la global, cierra el formulario si estaba abierto y lo abre; escucha `Alta`/`Editado` | `Form_FormOficinasProgramaGestion.cls:74-82, 381-425` |
| `Form_FormOficinasProgramaGestion` (Editar) | Asigna a la global la oficina seleccionada y lo abre | `Form_FormOficinasProgramaGestion.cls:223-231` |
| `Form_FormExpedienteGeneral` (`ComandoNuevaOficinaPrograma`) | Vacía la global y lo abre en alta; escucha solo `Alta` | `Form_FormExpedienteGeneral.cls:81-88, 693-709` |

Eventos hacia quien lo abrió: `Alta` (`:5`, lanzado en `:79`) y `Editado` (`:6`, lanzado en `:81`). Tras guardar, cierra el formulario (`:86`); `cmdSalir` cierra sin guardar (`:10-13`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `OficinaPrograma` | texto | — | Nombre de la oficina (obligatorio al guardar) | `.form.txt:930`; `:66, 148` |
| `DESCRIPCION` | texto | — | Descripción opcional | `.form.txt:965`; `:67, 151` |
| `ComandoRegistrar` (rótulo «Editar») | botón, deshabilitado en diseño | `OnClick` | Guarda el registro | `.form.txt:1040-1054`; `:42-96` |
| `ComandoAyuda` (rótulo «Salir», tip «Ayuda») | botón | `OnClick` | `AbrirAyuda` | `.form.txt:824-828`; `:15-40` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | `.form.txt:1154-1156`; `:10-13` |
| `lblTitulo` | etiqueta | — | «ALTA DE OFICINA DE PROGRAMA» o «EDICIÓN DE OFICINA DE PROGRAMA» | `:131, 135` |

## Flujos

### Apertura (`Form_Open`, `:165-198`)

1. `Ajustar Me` (`:173`; `FUNCIONES UTILES.bas:2703`). Deshabilita `ComandoRegistrar` y solo lo habilita con `EsAdministrador = Sí` (`:176-180`). A diferencia de `Form_FormEjercito.cls:176`, **no** fija `AllowEdits = False`.
2. `EstablecerDatos`: sin oficina activa, título de alta y fin (`:130-133`). Con ella, relee el registro con `constructor.getOficinaPrograma` por `IDOficinaPrograma` y lo guarda en `m_OficinaProgramaAlInicio` (`:137`); si no existe, «No se ha podido encontrar la OficinaPrograma registrada» (`:141-144`). Copia nombre y descripción no vacíos (`:146-153`).
3. Cualquier error: correo al administrador, `MsgBox` crítico y cierre (`:190-197`).

### Guardado (`ComandoRegistrar_Click`, `:42-96`)

1. `HaHabidoCambios`: en alta siempre; en edición compara nombre y descripción con el estado inicial (`:98-116`). Sin cambios, «No hay cambios que guardar» (`:58-61`).
2. Crea el objeto si la global está vacía y le copia los controles con `Nz` (`:62-68`).
3. `OficinaProgramaOperaciones.Registrar` (`:70-77`; detalle en Reglas).
4. Lanza `Alta` o `Editado` y cierra (`:78-86`). En alta, la global conserva el `IDOficinaPrograma` asignado (`OficinaProgramaOperaciones.cls:51`), que `Form_FormExpedienteGeneral` usa para seleccionar el nuevo valor (`Form_FormExpedienteGeneral.cls:706-707`).

**Errores:** distintos del interno 1000 se envían con `CorreoAlAdministrador` y se muestran con `MsgBox` crítico (`:88-95`); el formulario sigue abierto.

## Reglas, validaciones y cálculos (`classes/OficinaProgramaOperaciones.cls`)

- **Permiso:** solo administradores pueden guardar (`:178-180`).
- **Nombre obligatorio:** vacío ⇒ «Falta el campo OficinaPrograma» (`:28-31`).
- **Nombre único:** en alta o si el nombre cambia, busca por nombre y rechaza con «Ese nombre de OficinaPrograma ya está registrado» (`:32-48`; `constructor.bas:691-693`).
- **Identificador:** en alta, `IDOficinaProgramaCalculado` = `Max(IDOficinaPrograma) + 1` (`:51`; `OficinaPrograma.cls:30`; `FUNCIONES UTILES.bas:2851-2864`).
- **Descripción:** en alta solo se escribe si no está vacía; en edición, vacía pasa a `Null` (`:64-75`).
- **Caché y resincronización:** invalida `m_ObjEntorno.OficinasDePrograma` (`:81`); si una edición cambió el nombre, `ResyncCatalogo "oficina"` resincroniza los expedientes con ese `IDOficinaPrograma` (`:85-92`; `ExpedienteOperaciones.cls:695-697, 712-742`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Relectura / unicidad | `TbOficinasPrograma` por ID o por nombre | SQL dinámico concatenado | no | `constructor.bas:686-695` |
| Nuevo identificador | `Max(IDOficinaPrograma)` | SQL dinámico | no | `FUNCIONES UTILES.bas:2851-2853` |
| Alta / edición | `TbOficinasPrograma` (`AddNew`/`Edit` + `Update`) | DAO | no | `OficinaProgramaOperaciones.cls:59-79` |
| Resincronización | `TbExpedientes` por `IDOficinaPrograma` y derivados | SQL dinámico + DAO | no | `ExpedienteOperaciones.cls:712-742` |

## Coexistencia

Pendiente de `../coexistencia.md`. Igual que el resto de catálogos: identificador `Max + 1` en cliente, unicidad con `=` sobre texto (sensibilidad a mayúsculas distinta en Jet y PostgreSQL) y escritura más resincronización sin transacción común.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Sin `AllowEdits = False`, un no administrador puede teclear en los campos aunque no pueda guardar (`:176-180`).
- Un error de `DameID` queda en `OficinaPrograma.Error` sin consultar: el alta seguiría con ID vacío (`OficinaPrograma.cls:30-41`; `OficinaProgramaOperaciones.cls:51`).
- Patrón de familia (descrito en [Form_FormEjercito](Form_FormEjercito.md#comportamiento-actual--posibles-defectos--propuestas-web-pendientes-de-decisión)): `Max + 1` sin bloqueo, apóstrofo en el nombre rompe la consulta (`constructor.bas:693`), edición del objeto compartido con la gestión antes de validar (`:65-68`; `Form_FormOficinasProgramaGestion.cls:223`).
- Restos de diseño: `Caption` «Edición de Suministrador» (`.form.txt:38`, sustituido en `:177`), `lblTitulo` «EDICIÓN JURIDICA» (`.form.txt:806`), `ComandoAyuda` rotulado «Salir» (`.form.txt:825`).
- `Form_Open` llama a `CorreoAlAdministrador` sin `On Error Resume Next` (`:193`).

**Propuestas web pendientes de decisión:** ninguna aprobada; las mismas candidatas de la familia de catálogos.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un no administrador no puede guardar ni modificar los campos.
2. Guardar sin nombre muestra «Falta el campo OficinaPrograma» y no escribe.
3. Un nombre ya existente se rechaza; guardar el mismo registro sin renombrar, no.
4. Vaciar la descripción en edición la deja nula.
5. Renombrar una oficina resincroniza los expedientes que la usan.
6. Crear una oficina desde la ficha de expediente la deja seleccionada.

## Pendientes

- Si `AllowEdits` por defecto del `.form.txt` permite editar a no administradores en ejecución.
- Contenido de `SincronizarDerivados`, `AbrirAyuda` y `CorreoAlAdministrador`.
