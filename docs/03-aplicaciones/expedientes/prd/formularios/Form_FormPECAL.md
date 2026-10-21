# Form_FormPECAL

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormPECAL.cls` (197 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta y edición de una PECAL del catálogo (`TbPECAL`: código y descripción). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): lee la PECAL de la global `m_ObjPECALActiva` (`Variables Globales.bas:262`) y persiste con `PECALOperaciones.Registrar`. Sigue el patrón de alta de catálogo descrito en [Form_FormCPV](Form_FormCPV.md); esta ficha solo detalla lo propio de la PECAL. Las citas son `Form_FormPECAL.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| [Form_FormPECALESGestion](Form_FormPECALESGestion.md) (alta) | `m_ObjPECALActiva = Nothing`, cierra el formulario si estaba abierto y lo abre; escucha `Alta`/`Editado` | `Form_FormPECALESGestion.cls:73-81` |
| [Form_FormPECALESGestion](Form_FormPECALESGestion.md) (edición) | `m_ObjPECALActiva = m_PECALSeleccionada` y lo abre | `Form_FormPECALESGestion.cls:222-230` |
| [Form_FormExpedienteEntidades](Form_FormExpedienteEntidades.md) | `ComandoNuevaPECAL_Click`: alta desde la pestaña Entidades; escucha `Alta` | `Form_FormExpedienteEntidades.cls:1088-1102, 1883-1899` |

Eventos hacia quien lo abrió: `Alta` (`:5`, lanzado en `:79`) y `Editado` (`:6`, lanzado en `:81`). Tras guardar se cierra solo (`:86`); `cmdSalir` cierra sin guardar (`:10-13`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `PECAL` | texto no enlazado | — | Código de la PECAL; obligatorio al guardar | `.form.txt:916-930`; `PECALOperaciones.cls:29-32` |
| `DESCRIPCION` | texto no enlazado | — | Descripción opcional | `.form.txt:950-965` |
| `lblTitulo` | etiqueta | — | «ALTA DE PECAL» o «EDICIÓN DE PECAL» según haya PECAL activa | `:131, 135` |
| `ComandoRegistrar` (rótulo «Editar») | botón | `OnClick` → `ComandoRegistrar_Click` | Guarda y cierra | `:42-96`, `.form.txt:1143-1157` |
| `ComandoAyuda` (rótulo «Salir» en `.form.txt`) | botón | `OnClick` | `AbrirAyuda` | `:15-40`, `.form.txt:824-826` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | `:10-13`, `.form.txt:1048-1050` |

## Flujos

### Apertura (`Form_Open`, `:165-198`)

1. `Ajustar Me`; `ComandoRegistrar` deshabilitado y habilitado solo con `EsAdministrador = Sí` (`:173-180`). A diferencia de `Form_FormCPV` (`Form_FormCPV.cls:176, 180`), **no** fija `AllowEdits`.
2. `EstablecerDatos` (`:126-162`): sin PECAL activa, rótulo de alta y nada más (`:130-134`). Con PECAL activa, relee el registro con `constructor.getPecal(p_IDPECAL)` en `m_PECALAlInicio` (`:137`); si no existe, «No se ha podido encontrar la PECAL registrada» (`:141-144`). Copia `PECAL` y `DESCRIPCION` no vacíos (`:146-153`).
3. Cualquier error: correo al administrador, `MsgBox` crítico y cierre (`:190-197`).

### Guardado (`ComandoRegistrar_Click`, `:42-96`)

1. `HaHabidoCambios`: en alta siempre cierto (`:101-104`); en edición compara `PECAL` y `DESCRIPCION` (`:107-114`). Sin cambios: «No hay cambios que guardar» (`:58-61`).
2. Crea `m_ObjPECALActiva` si no existía y le copia los dos controles con `Nz` (`:62-68`).
3. `PECALOperaciones.Registrar m_PECALAlInicio` (`:70-77`; reglas abajo).
4. Lanza `Alta` o `Editado` (`:78-82`) y cierra (`:86`). Error: correo salvo error 1000, `MsgBox` crítico y el formulario queda abierto (`:88-95`).

## Reglas, validaciones y cálculos

- **Obligatorio:** `PECAL` vacío ⇒ «Falta el campo PECAL» (`PECALOperaciones.cls:29-32`).
- **Unicidad del código:** en alta, o en edición si cambia el código, busca otra PECAL con el mismo texto (`constructor.bas:791-793`); si existe, «Ese nombre de PECAL ya está registrado» (`PECALOperaciones.cls:33-49`).
- **Identificador:** en alta `IDPECAL = IDPECALCalculado` (`PECALOperaciones.cls:52`), es decir `DameID("TbPECAL", "IDPECAL")` sobre `Max` (`classes/PECAL.cls:31`; `FUNCIONES UTILES.bas:2851-2853`).
- **Descripción vacía:** en alta no se asigna; en edición se guarda `Null` (`PECALOperaciones.cls:65-76`).
- **Renombrado (issue #237):** si en edición cambia el código, `ResyncCatalogo "pecal"` recalcula los derivados (`SincronizarDerivados`, ámbito `Todo`) de cada expediente que la usa vía `TbExpedientesPECAL` (`PECALOperaciones.cls:83-92`; `ExpedienteOperaciones.cls:674-676, 712-742`).
- **Caché:** `Registrar` no invalida `m_ObjEntorno.PECALES`; lo hacen quienes escuchan los eventos (`Form_FormPECALESGestion.cls:388, 411`; `Form_FormExpedienteEntidades.cls:1891`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura por ID / por código | `TbPECAL` | SQL dinámico concatenado | — | `constructor.bas:786-793` |
| Siguiente ID | `TbPECAL` | `SELECT Max(...)` | no | `FUNCIONES UTILES.bas:2851-2853` |
| Alta / edición | `TbPECAL` | DAO `AddNew`/`Edit` + `Update` | no | `PECALOperaciones.cls:53-80` |
| Resync tras renombrar | `TbExpedientesPECAL` y derivados | SQL dinámico + `SincronizarDerivados` | no, tras la escritura | `ExpedienteOperaciones.cls:712-742` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: ID con `Max + 1` en VBA; SQL concatenado con el código entre comillas; escritura del catálogo y resync de expedientes en pasos separados sin transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Si `ResyncCatalogo` falla, la PECAL ya está renombrada (`PECALOperaciones.cls:70-79, 85-91`), el formulario sigue abierto y no se lanza `Editado` (`:74-76`); reintentar volvería a comprobar el nombre nuevo y encontraría la propia PECAL: «Ese nombre de PECAL ya está registrado» (`PECALOperaciones.cls:34-47`).
- La PECAL «N/A», protegida contra el borrado (`PECALOperaciones.cls:112-115`), puede renombrarse sin restricción (`PECALOperaciones.cls:28-49`).
- Sin `AllowEdits = False` (`:173-180`), un no administrador puede teclear en los controles aunque no pueda guardar.
- Comunes a la familia de catálogos (ver [Form_FormCPV](Form_FormCPV.md)): `Max + 1` sin bloqueo; apóstrofo en el código rompe la búsqueda (`constructor.bas:793`); cambios solo de mayúsculas no cuentan con `Option Compare Database` (`:2, 107, 111`); la global se modifica antes de guardar (`:65-68`); rótulos de diseño «Edición de Suministrador» y «EDICIÓN PECAL» (`.form.txt:38, 806`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: renombrado y resync en una sola transacción; identificador generado por la base y restricción única sobre el código.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un no administrador no puede guardar una PECAL.
2. Guardar con el código vacío muestra «Falta el campo PECAL» y no escribe.
3. Guardar un código ya existente muestra «Ese nombre de PECAL ya está registrado».
4. Renombrar una PECAL en uso actualiza la cadena PECAL de todos los expedientes que la referencian.
5. Crear una PECAL desde la pestaña Entidades la deja seleccionada y asociada al expediente (`Form_FormExpedienteEntidades.cls:1896-1898`).

## Pendientes

- Índice único sobre `TbPECAL.PECAL` y longitudes de columna (ERD de origen).
- Qué escribe exactamente `SincronizarDerivados` (`ExpedienteOperaciones.cls:735`); se trata en la ficha de expedientes.
- Si el cierre sin `acSaveNo` desde Entidades (`Form_FormExpedienteEntidades.cls:1096`) puede pedir confirmación de guardado de diseño.
