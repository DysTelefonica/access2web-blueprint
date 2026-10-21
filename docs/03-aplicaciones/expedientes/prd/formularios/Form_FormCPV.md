# Form_FormCPV

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormCPV.cls` (199 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta y edición de un código CPV del catálogo (`TbCPV`: código y descripción). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): lee el CPV de la global `m_ObjCPVActivo` (`Variables Globales.bas:252`) y persiste con `CPVOperaciones.Registrar`. Las citas son `Form_FormCPV.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormCPVsGestion` (alta) | `m_ObjCPVActivo = Nothing`, cierra el formulario si estaba abierto y lo abre; escucha `Alta`/`Editado` | `Form_FormCPVsGestion.cls:73-81` |
| `Form_FormCPVsGestion` (edición) | `m_ObjCPVActivo = m_CPVSeleccionado` y lo abre | `Form_FormCPVsGestion.cls:222-230` |
| `Form_FormExpedienteEntidades` | `ComandoNuevoCPV_Click`: alta desde la pestaña Entidades; escucha `Alta` | `Form_FormExpedienteEntidades.cls:1175-1189, 1291-1307` |

Eventos hacia quien lo abrió: `Alta` (`:5`, lanzado en `:79`) y `Editado` (`:6`, lanzado en `:81`). Tras guardar se cierra solo (`:86`); `cmdSalir` cierra sin guardar (`:10-13`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `CPV` | texto no enlazado | — | Código CPV; obligatorio al guardar | `.form.txt:917-931`; `CPVOperaciones.cls:27-30` |
| `DESCRIPCION` | texto no enlazado | — | Descripción opcional | `.form.txt:951-966` |
| `lblTitulo` | etiqueta | — | «ALTA DE CPV» o «EDICIÓN DE CPV» según haya CPV activo | `:131, 135, 145` |
| `ComandoRegistrar` (rótulo «Editar») | botón | `OnClick` → `ComandoRegistrar_Click` | Guarda y cierra | `:42-96`, `.form.txt:1144-1158` |
| `ComandoAyuda` (rótulo «Salir» en `.form.txt`) | botón | `OnClick` | `AbrirAyuda` (`FUNCIONES UTILES.bas:3076`) | `:15-40`, `.form.txt:825-827` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | `:10-13`, `.form.txt:1049-1051` |

## Flujos

### Apertura (`Form_Open`, `:166-200`)

1. `Ajustar Me`; `ComandoRegistrar` deshabilitado y `AllowEdits = False`; ambos se activan solo con `EsAdministrador = Sí` (`:174-181`).
2. `EstablecerDatos` (`:126-163`): sin CPV activo, rótulo de alta y nada más (`:130-134`). Con CPV activo, **relee** el registro con `constructor.getCPV(p_IDCPV)` en `m_CPVAlInicio` (`:137`); si no existe, «No se ha podido encontrar el CPV registrado» (`:141-144`). Copia `CPV` y `DESCRIPCION` no vacíos a los controles (`:147-154`).
3. Cualquier error: correo al administrador, `MsgBox` crítico y cierre del formulario (`:192-199`).

### Guardado (`ComandoRegistrar_Click`, `:42-96`)

1. `HaHabidoCambios`: en alta siempre cierto (`:101-104`); en edición compara `CPV` y `DESCRIPCION` con `m_CPVAlInicio` (`:107-114`). Sin cambios: «No hay cambios que guardar» (`:58-61`).
2. Crea `m_ObjCPVActivo` si no existía y le copia los dos controles con `Nz` (`:62-68`).
3. `CPVOperaciones.Registrar m_CPVAlInicio` (`:70-77`; reglas abajo).
4. Lanza `Alta` o `Editado` según hubiera registro inicial (`:78-82`) y cierra (`:86`).
5. Error: acumula en `m_Error`, correo salvo error interno 1000 y `MsgBox` crítico; el formulario queda abierto (`:88-95`).

## Reglas, validaciones y cálculos

- **Obligatorio:** `CPV` vacío ⇒ «Falta el campo CPV» (`CPVOperaciones.cls:27-30`).
- **Unicidad del código:** en alta, o en edición si cambia el código, busca otro CPV con el mismo texto (`constructor.bas:396-398`); si existe, «Ese nombre de CPV ya está registrado» (`CPVOperaciones.cls:31-47`).
- **Identificador:** en alta `IDCPV = IDCPVCalculado` (`CPVOperaciones.cls:49`), que es `Max(IDCPV) + 1` (`classes/CPV.cls:30`; `FUNCIONES UTILES.bas:2851-2864`).
- **Descripción vacía:** en alta no se asigna; en edición se guarda `Null` (`CPVOperaciones.cls:61-72`).
- **Caché:** tras guardar invalida `m_ObjEntorno.CPVs` (`CPVOperaciones.cls:78`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura por ID | `TbCPV` | SQL dinámico concatenado | — | `constructor.bas:391-394` |
| Búsqueda por código | `TbCPV` | SQL dinámico con el texto entre comillas | — | `constructor.bas:396-398` |
| Siguiente ID | `TbCPV` | `SELECT Max(...)` | no | `FUNCIONES UTILES.bas:2851-2853` |
| Alta | `TbCPV` | DAO `AddNew`/`Update` | no | `CPVOperaciones.cls:56-64` |
| Edición | `TbCPV` | DAO `Edit`/`Update` sobre `WHERE IDCPV=` | no | `CPVOperaciones.cls:52-54, 66-73` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: identificador calculado en VBA con `Max + 1` fuera de transacción; SQL concatenado con el código entre comillas simples.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Max(IDCPV) + 1` sin transacción ni bloqueo (`FUNCIONES UTILES.bas:2851-2864`; `CPVOperaciones.cls:49-64`): dos altas simultáneas podrían calcular el mismo ID.
- La búsqueda de duplicado concatena el código entre comillas (`constructor.bas:398`): un código con apóstrofo rompería la consulta.
- Con `Option Compare Database` (`:2`) las comparaciones de `HaHabidoCambios` (`:107, 111`) y de `Registrar` (`CPVOperaciones.cls:32`) podrían no distinguir mayúsculas: cambiar solo la caja daría «No hay cambios que guardar».
- `ComandoRegistrar_Click` escribe el formulario en la global `m_ObjCPVActivo` antes de guardar (`:65-68`); si el guardado falla, la global queda modificada.
- Rótulos de diseño incoherentes: título de ventana «Edición de Suministrador» (`.form.txt:38`, sobrescrito en `:177`), `ComandoAyuda` rotulado «Salir» (`.form.txt:826`) y `ComandoRegistrar` rotulado «Editar» también en altas (`.form.txt:1157`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: identificador generado por la base de datos y restricción única sobre el código.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario no administrador ve el CPV sin poder editarlo ni guardar.
2. Guardar con el código vacío muestra «Falta el campo CPV» y no escribe.
3. Guardar un código ya existente muestra «Ese nombre de CPV ya está registrado».
4. Guardar en edición sin cambios muestra «No hay cambios que guardar».
5. Vaciar la descripción en edición la deja en blanco en el registro.
6. Crear un CPV desde la pestaña Entidades lo deja seleccionado en su combo `CPVs` (`Form_FormExpedienteEntidades.cls:1304-1306`).

## Pendientes

- Si `TbCPV` tiene índice único sobre `CPV` y longitudes de columna (ERD de origen).
- Si `AllowEdits = False` bloquea los controles no enlazados de este formulario (`:176`).
- Contenido de `Ajustar` y `AvanceCerrar` (`:174, 187`).
