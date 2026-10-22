# Form_FormLugarEjecucion

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormLugarEjecucion.cls` (197 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta y edición de un lugar de ejecución del catálogo (`TbLugaresEjecucion`: nombre y descripción). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): lee el registro de la global `m_ObjLugarEjecucionActiva` (`Variables Globales.bas:259`) y persiste con `LugarEjecucionOperaciones.Registrar`. Mismo patrón que [Form_FormCPV](Form_FormCPV.md). Las citas son `Form_FormLugarEjecucion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormLugarEjecucionGestion` (alta) | `m_ObjLugarEjecucionActiva = Nothing`, cierra el formulario si estaba abierto y lo abre; escucha `Alta`/`Editado` | `Form_FormLugarEjecucionGestion.cls:73-81, 373-417` |
| `Form_FormLugarEjecucionGestion` (edición) | `m_ObjLugarEjecucionActiva = m_LugarEjecucionSeleccionada` y lo abre | `Form_FormLugarEjecucionGestion.cls:222-230` |
| `Form_FormExpedienteEntidades` | `ComandoNuevoLugar_Click`: alta desde la pestaña Entidades; en `Alta` recarga combos, selecciona el nuevo lugar y lo añade al expediente | `Form_FormExpedienteEntidades.cls:1204-1218, 1853-1869` |

Eventos hacia quien lo abrió: `Alta` (`:5`, lanzado en `:79`) y `Editado` (`:6`, lanzado en `:81`). Tras guardar se cierra solo (`:86`); `cmdSalir` cierra sin guardar (`:10-13`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `LugarEjecucion` | texto no enlazado | — | Nombre; obligatorio y único al guardar | `.form.txt:916-930`; `LugarEjecucionOperaciones.cls:29-48` |
| `DESCRIPCION` | texto no enlazado | — | Descripción opcional | `.form.txt:950-965` |
| `lblTitulo` | etiqueta | — | «ALTA DE LUGAR EJECUCIÓN» o «EDICIÓN DE LUGAR EJECUCIÓN» | `:131, 135` |
| `ComandoRegistrar` (rótulo «Editar») | botón | `OnClick` → `ComandoRegistrar_Click` | Guarda y cierra | `:42-96`, `.form.txt:1039-1052` |
| `ComandoAyuda` (rótulo «Salir» en `.form.txt`) | botón | `OnClick` | `AbrirAyuda` (`FUNCIONES UTILES.bas:3076`) | `:15-40`, `.form.txt:814-826` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | `:10-13`, `.form.txt:1144-1156` |

El `Caption` de diseño del formulario («Edición de Suministrador», `.form.txt:38`) se sustituye en ejecución por `m_TituloFormulario` (`:176`).

## Flujos

### Apertura (`Form_Open`, `:165-198`)

1. `Ajustar Me` (`FUNCIONES UTILES.bas:2703`); `ComandoRegistrar` deshabilitado salvo con `EsAdministrador = Sí` (`:173-179`). No se toca `AllowEdits`: un usuario no administrador puede escribir en los campos, pero no guardar.
2. `EstablecerDatos` (`:126-162`): sin lugar activo, rótulo de alta y nada más (`:130-134`). Con lugar activo, **relee** el registro con `constructor.getLugarEjecucion(p_IDLugarEjecucion)` en `m_LugarEjecucionAlInicio` (`:137`); si no existe, «No se ha podido encontrar la LugarEjecucion registrada» (`:141-144`). Copia los dos campos no vacíos a los controles (`:146-153`).
3. Cualquier error: correo al administrador, `MsgBox` crítico y cierre (`:190-197`).

### Guardado (`ComandoRegistrar_Click`, `:42-96`)

1. `HaHabidoCambios`: en alta siempre cierto (`:101-104`); en edición compara nombre y descripción con `m_LugarEjecucionAlInicio` (`:105-116`). Sin cambios: «No hay cambios que guardar» (`:58-61`).
2. Crea `m_ObjLugarEjecucionActiva` si no existía y le copia los controles con `Nz` (`:62-68`).
3. `LugarEjecucionOperaciones.Registrar m_LugarEjecucionAlInicio` (`:70-77`; reglas abajo).
4. Lanza `Alta` o `Editado` (`:78-82`) y cierra (`:86`). En alta, la global queda con el `IDLugarEjecucion` asignado, que usa `Form_FormExpedienteEntidades` (`Form_FormExpedienteEntidades.cls:1866-1868`).
5. Error: correo salvo error interno 1000 y `MsgBox` crítico; el formulario queda abierto (`:88-95`).

## Reglas, validaciones y cálculos

- **Obligatorio:** nombre vacío ⇒ «Falta el campo LugarEjecucion» (`LugarEjecucionOperaciones.cls:29-32`).
- **Unicidad del nombre:** en alta, o en edición si cambia el nombre, busca otro registro con el mismo texto (`constructor.bas:641-643`); si existe, «Ese nombre de LugarEjecucion ya está registrado» (`LugarEjecucionOperaciones.cls:33-48`).
- **Identificador:** en alta `IDLugarEjecucion = IDLugarEjecucionCalculado` (`LugarEjecucionOperaciones.cls:52`), que es `Max(IDLugarEjecucion) + 1` (`classes/LugarEjecucion.cls:54`; `FUNCIONES UTILES.bas:2851-2853`).
- **Descripción vacía:** en alta no se asigna; en edición se guarda `Null` (`LugarEjecucionOperaciones.cls:65-76`).
- **Caché y derivados:** tras escribir invalida `m_ObjEntorno.LugaresEjecucion` (`LugarEjecucionOperaciones.cls:82`). Si en una edición cambia el nombre, `ResyncCatalogo "lugar"` recalcula los derivados de cada expediente vinculado en `TbExpedientesLugaresEjecucion` (`LugarEjecucionOperaciones.cls:84-93`; `ExpedienteOperaciones.cls:683-685, 712-742`). Lo cubre la prueba de `Test_Issue237CacheAltasSueltas.bas:2130`.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura por ID / por nombre | `TbLugaresEjecucion` | SQL dinámico concatenado (nombre entre comillas) | — | `constructor.bas:636-643` |
| Siguiente ID | `TbLugaresEjecucion` | `SELECT Max(...)` | no | `FUNCIONES UTILES.bas:2851-2853` |
| Alta / edición | `TbLugaresEjecucion` | DAO `AddNew` / `Edit` + `Update` | no | `LugarEjecucionOperaciones.cls:56-78` |
| Resincronización por renombrado | `TbExpedientesLugaresEjecucion` y derivados | DAO, `SincronizarDerivados` por expediente | no, posterior a la escritura | `ExpedienteOperaciones.cls:712-742` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: identificador `Max + 1` calculado en VBA; búsqueda por nombre con `=` sobre texto, cuya sensibilidad a mayúsculas difiere entre Jet/ACE y PostgreSQL; escritura y resincronización sin transacción común.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Renombrado sin transacción: si `ResyncCatalogo` falla, el nombre ya está escrito y el formulario sigue abierto (`LugarEjecucionOperaciones.cls:70-93`). Un reintento vuelve a comparar con el nombre antiguo y la unicidad encuentra el propio registro: «ya está registrado» (`LugarEjecucionOperaciones.cls:34-48`).
- `IDLugarEjecucionCalculado` deja el error en `.Error` sin propagarlo (`classes/LugarEjecucion.cls:54-65`); `Registrar` no lo comprueba (`LugarEjecucionOperaciones.cls:52`) y podría insertar un ID vacío.
- `Max + 1` y la comprobación de unicidad no son atómicos: dos altas simultáneas pueden chocar (`LugarEjecucionOperaciones.cls:41-68`).
- Un apóstrofo en el nombre rompe la búsqueda por nombre (`constructor.bas:641-643`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: identificador generado por la base; unicidad garantizada por índice; renombrado y resincronización en una transacción.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario no administrador ve el lugar pero no puede guardarlo.
2. Guardar sin nombre muestra «Falta el campo LugarEjecucion» y no escribe.
3. Guardar un nombre ya existente lo rechaza; cambiar solo la descripción no lo comprueba.
4. Renombrar un lugar actualiza los datos derivados de los expedientes que lo usan.
5. Un alta desde la pestaña Entidades deja el nuevo lugar añadido al expediente.

## Pendientes

- Si `TbLugaresEjecucion` tiene índice único sobre el nombre (`erd-origen.md`).
- Contenido de `Ajustar` y de `SincronizarDerivados` (fuera de esta ficha).
- `Form_FormExpedienteAltaParaHPS` declara `WithEvents m_FormLugar` y su handler `Alta` (`Form_FormExpedienteAltaParaHPS.cls:12, 452-454`), pero no se ha encontrado ninguna asignación de `m_FormLugar` en ese formulario.
