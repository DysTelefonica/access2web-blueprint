# Form_FormRAC

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormRAC.cls` (205 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta y edición de un RAC del catálogo (`TbRACS`: nombre, correo y descripción). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): lee el RAC de la global `m_ObjRACActivo` (`Variables Globales.bas:263`) y persiste con `RACOperaciones.Registrar`. Mismo esqueleto que [Form_FormCPV](Form_FormCPV.md); esta ficha recoge lo propio del RAC. Las citas son `Form_FormRAC.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| [Form_FormRACSGestion](Form_FormRACSGestion.md) (alta) | `m_ObjRACActivo = Nothing`, cierra el formulario si estaba abierto y lo abre; escucha `Alta`/`Editado` | `Form_FormRACSGestion.cls:72-80` |
| [Form_FormRACSGestion](Form_FormRACSGestion.md) (edición) | `m_ObjRACActivo = m_RACSeleccionado` y lo abre | `Form_FormRACSGestion.cls:221-229` |
| [Form_FormExpedienteEntidades](Form_FormExpedienteEntidades.md) | `ComandoNuevoRAC_Click`: alta desde la pestaña Entidades; escucha `Alta` | `Form_FormExpedienteEntidades.cls:11, 1117-1131, 1913-1929` |

Eventos hacia quien lo abrió: `Alta` (`:5`, lanzado en `:80`) y `Editado` (`:6`, lanzado en `:82`). Tras guardar se cierra solo (`:87`); `cmdSalir` cierra sin guardar (`:10-13`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `RAC` | texto no enlazado | — | Nombre del RAC; obligatorio y único al guardar | `.form.txt:930`; `RACOperaciones.cls:29-49` |
| `CORREO` (etiqueta «EMAIL») | texto no enlazado | — | Correo opcional; sin validación de formato | `.form.txt:1017, 1063` |
| `DESCRIPCION` | texto no enlazado | — | Descripción opcional | `.form.txt:965` |
| `lblTitulo` | etiqueta | — | «ALTA DE RAC» o «EDICIÓN DE RAC» según haya RAC activo | `:136, 140` |
| `ComandoRegistrar` (rótulo «Editar») | botón | `OnClick` → `ComandoRegistrar_Click` | Guarda y cierra | `:42-97`, `.form.txt:1211-1213` |
| `ComandoAyuda` (rótulo «Salir» en `.form.txt`) | botón | `OnClick` | `AbrirAyuda` (`FUNCIONES UTILES.bas:3076`) | `:15-40`, `.form.txt:824-826` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | `:10-13`, `.form.txt:1104-1106` |

## Flujos

### Apertura (`Form_Open`, `:173-206`)

1. `Ajustar Me`; `ComandoRegistrar` deshabilitado y habilitado solo con `EsAdministrador = Sí` (`:181-188`). A diferencia de `Form_FormCPV`, **no** fija `AllowEdits = False`: un no administrador puede escribir en los controles, pero no guardar.
2. `EstablecerDatos` (`:131-170`): sin RAC activo, rótulo de alta y nada más (`:135-139`). Con RAC activo, **relee** el registro con `constructor.getRAC(p_IDRAC)` en `m_RACAlInicio` (`:142`); si no existe, «No se ha podido encontrar la RAC registrada» (`:146-149`). Copia `RAC`, `CORREO` y `DESCRIPCION` no vacíos a los controles (`:151-161`).
3. Cualquier error: correo al administrador, `MsgBox` crítico y cierre del formulario (`:198-205`).

### Guardado (`ComandoRegistrar_Click`, `:42-97`)

1. `HaHabidoCambios`: en alta siempre cierto (`:102-105`); en edición compara los tres campos con `m_RACAlInicio` (`:106-121`). Sin cambios: «No hay cambios que guardar» (`:58-61`).
2. Crea `m_ObjRACActivo` si no existía y le copia los tres controles con `Nz` (`:62-69`).
3. `RACOperaciones.Registrar m_RACAlInicio` (`:71-78`; reglas abajo).
4. Lanza `Alta` o `Editado` según hubiera registro inicial (`:79-83`) y cierra (`:87`).
5. Error: acumula en `m_Error`, correo salvo error interno 1000 y `MsgBox` crítico; el formulario queda abierto (`:89-96`).

## Reglas, validaciones y cálculos

- **Obligatorio:** `RAC` vacío ⇒ «Falta el campo RAC» (`RACOperaciones.cls:29-32`).
- **Unicidad del nombre:** en alta, o en edición si cambia el nombre, busca otro RAC con ese texto (`constructor.bas:840-842`); si existe, «Ese nombre de RAC ya está registrado» (`RACOperaciones.cls:33-49`).
- **Identificador:** en alta `IDRAC = IDRACCalculado` (`RACOperaciones.cls:52`), que es `DameID("TbRACS", "IDRAC")` = `Max + 1` (`classes/RAC.cls:31`; `FUNCIONES UTILES.bas:2851-2864`).
- **Opcionales vacíos:** en alta no se asignan; en edición se guardan como `Null` (`RACOperaciones.cls:65-70, 75-84`).
- **Renombrado:** si una edición cambia el nombre, `ResyncCatalogo "rac"` recalcula los derivados de cada expediente que lo usa en `TbExpedientesRACS` (`RACOperaciones.cls:92-101`; `ExpedienteOperaciones.cls:677-679, 723-736`).
- **Caché:** tras guardar invalida `m_ObjEntorno.RACs` (`RACOperaciones.cls:90`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura por ID / por nombre | `TbRACS` | SQL dinámico concatenado | — | `constructor.bas:835-842` |
| Siguiente ID | `TbRACS` | `SELECT Max(...)` | no | `FUNCIONES UTILES.bas:2851-2853` |
| Alta | `TbRACS` | DAO `AddNew`/`Update` | no | `RACOperaciones.cls:53, 60-71` |
| Edición | `TbRACS` | DAO `Edit`/`Update` sobre `WHERE IDRAC=` | no | `RACOperaciones.cls:56-60, 73-85` |
| Resincronización | expedientes de `TbExpedientesRACS` | DAO por expediente | no | `ExpedienteOperaciones.cls:713-736` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: identificador `Max + 1` en VBA fuera de transacción; SQL concatenado con el nombre entre comillas simples; edición del RAC y resincronización de expedientes en pasos separados sin transacción común.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Edición y `ResyncCatalogo` no comparten transacción (`RACOperaciones.cls:73-101`): si la resincronización falla, el nombre nuevo ya está guardado, el formulario sigue abierto con `m_RACAlInicio` antiguo (`:89-96`) y un segundo «Guardar» encontraría su propio registro y daría «Ese nombre de RAC ya está registrado» (`RACOperaciones.cls:34-47`).
- `IDRACCalculado` deja su error en `RAC.Error` y devuelve vacío (`classes/RAC.cls:31-42`); `Registrar` no lo comprueba (`RACOperaciones.cls:52`) y haría `AddNew` con `IDRAC` vacío.
- Sin `AllowEdits = False` (`:184-188`), un no administrador puede modificar los campos en pantalla sin poder guardarlos.
- Patrones compartidos con `Form_FormCPV` y ya registrados en su ficha: `Max + 1` sin bloqueo, comillas sin escapar (`constructor.bas:842`), comparación sin distinguir mayúsculas por `Option Compare Database` (`:2`, `RACOperaciones.cls:10`) y la global `m_ObjRACActivo` modificada antes de guardar (`:65-69`).
- Rótulos de diseño incoherentes: ventana «Edición de Suministrador» (`.form.txt:38`, sobrescrito en `:185`), `lblTitulo` «EDICIÓN PECAL» (`.form.txt:806`, sobrescrito en `:136, 140`), `ComandoAyuda` «Salir» y `ComandoRegistrar` «Editar» también en altas.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: validar el formato del correo; identificador generado por la base de datos con restricción única sobre el nombre; renombrado y resincronización en una sola transacción.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un no administrador ve el RAC sin poder guardar.
2. Guardar con el nombre vacío muestra «Falta el campo RAC» y no escribe.
3. Guardar un nombre ya existente muestra «Ese nombre de RAC ya está registrado».
4. Guardar en edición sin cambios muestra «No hay cambios que guardar».
5. Renombrar un RAC en uso actualiza los datos derivados de sus expedientes.
6. Crear un RAC desde la pestaña Entidades lo deja seleccionado y añadido (`Form_FormExpedienteEntidades.cls:1926-1929`).

## Pendientes

- Índice único y longitudes de `TbRACS` (ERD de origen).
- Contenido de `SincronizarDerivados` con ámbito `Todo` para RAC (`ExpedienteOperaciones.cls:735`) y de `ComandoAltaRACs_Click` en Entidades.
- Contenido de `Ajustar` y `AvanceCerrar` (`:181, 193`).
