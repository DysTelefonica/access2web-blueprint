# Form_FormComercial

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormComercial.cls` (200 líneas), `.form.txt` y `src/classes/ComercialOperaciones.cls` · **PRD:** [README](../README.md)

Alta y edición de un comercial del catálogo `TbComerciales` (nombre y descripción). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): lee el comercial con `constructor.getComercial` y lo guarda con `ComercialOperaciones.Registrar`. Las citas son `Form_FormComercial.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormComercialesGestion` | `ComandoAlta` (global a `Nothing`) y `ComandoEditar` (global = seleccionado); escucha `Alta` y `Editado` | `Form_FormComercialesGestion.cls:72-80, 221-229, 372-416` |
| `Form_FormExpedienteEntidades` | `ComandoNuevoComercial`: global a `Nothing` y apertura; escucha solo `Alta` | `Form_FormExpedienteEntidades.cls:1146-1160, 1261-1277` |

El comercial a editar llega por la global `m_ObjComercialActivo` (`Variables Globales.bas:251`), no por `OpenArgs`. Eventos: `Alta` (`:5`, lanzado en `:79`) y `Editado` (`:6`, lanzado en `:81`). `cmdSalir` cierra sin guardar (`:10-13`); tras guardar se cierra solo (`:86`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Comercial` (etiqueta «NOMBRE») | texto | — | Nombre del comercial | `.form.txt:931, 996` |
| `DESCRIPCION` (etiqueta «DESCRIP.») | texto | — | Descripción opcional | `.form.txt:966, 1016` |
| `ComandoRegistrar` (rótulo «Editar», deshabilitado en diseño) | botón | `OnClick` → `ComandoRegistrar_Click` | Valida cambios y guarda | `:42-96`, `.form.txt:1039-1051` |
| `ComandoAyuda` (rótulo «Salir» en `.form.txt`) | botón | `OnClick` | `AbrirAyuda` (`FUNCIONES UTILES.bas:3076`) | `:15-40`, `.form.txt:824-826` |
| `cmdSalir` | botón | `OnClick` | `DoCmd.Close` con `On Error Resume Next` | `:10-13`, `.form.txt:1152-1154` |
| `lblTitulo` | etiqueta | — | «ALTA DE COMERCIAL» o «EDICIÓN DE COMERCIAL» | `:131, 135, 145` |

## Flujos

### Apertura (`Form_Open`, `:166-200`)

1. `ComandoRegistrar` deshabilitado y `AllowEdits = False`; ambos pasan a `True` solo si `EsAdministrador = Sí` (`:175-181`). El título de ventana se toma de `m_TituloFormulario` (`:177`).
2. `EstablecerDatos`: sin global activa es un alta y solo rotula el título (`:130-134`). Con global, **relee** el comercial por `IDComercial` en `m_ComercialAlInicio` (`:137`); si no existe, «No se ha podido encontrar el Comercial registrado» (`:141-144`). Copia nombre y descripción no vacíos a los controles (`:147-154`).

### Guardado (`ComandoRegistrar_Click`, `:42-96`)

1. `HaHabidoCambios`: un alta siempre cuenta como cambio (`:101-104`); en edición compara nombre y descripción con `m_ComercialAlInicio` (`:105-114`). Sin cambios, «No hay cambios que guardar» (`:58-61`).
2. Copia los controles a la global (la crea si es alta) (`:62-68`) y llama a `Registrar` con el comercial inicial (`:70-77`).
3. Éxito: `Alta` si no había comercial inicial, `Editado` si lo había (`:78-82`), y cierra (`:86`).

### Errores

Patrón común: error en `m_Error`, `CorreoAlAdministrador` salvo el error interno 1000 y `MsgBox` crítico (`:88-95`). Un fallo en `Form_Open` cierra el formulario (`:192-199`).

## Reglas, validaciones y cálculos (`classes/ComercialOperaciones.cls`)

- Nombre obligatorio: «Falta el campo Comercial» (`:29-32`).
- Nombre único: en alta, o en edición si el nombre cambia, busca por nombre; si existe, «Ese nombre de comercial ya está registrado» (`:33-49`). No se recorta el texto.
- Identificador: en alta, `IDComercialCalculado` = máximo de `IDComercial` + 1 (`classes/Comercial.cls:23-34`; `DameID`, `FUNCIONES UTILES.bas:2836-2864`).
- Descripción: en alta solo se escribe si no está vacía; en edición vacía pasa a `Null` (`:65-76`).
- Tras guardar invalida la caché `m_ObjEntorno.Comerciales` (`:82`). Si fue una edición con cambio de nombre, `ResyncCatalogo "comercial"` resincroniza los derivados de cada expediente que usa el comercial (`:86-93`; `classes/ExpedienteOperaciones.cls:680-682, 712-742`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura por ID | `TbComerciales` | SQL dinámico concatenado | — | `constructor.bas:342-345` |
| Duplicado por nombre | `TbComerciales` | SQL dinámico con comillas simples | — | `constructor.bas:347-349` |
| Nuevo ID | `Max(IDComercial)` de `TbComerciales` | SQL dinámico | — | `FUNCIONES UTILES.bas:2851-2853` |
| Alta / edición | `TbComerciales` (`AddNew` / `Edit`) | DAO Recordset | No | `ComercialOperaciones.cls:60-80` |
| Resincronización | `TbExpedientesComerciales` → derivados del expediente | DAO | No (tras el `Update`) | `ExpedienteOperaciones.cls:712-742` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: ID calculado en VBA con `Max + 1` en lugar de autonumérico; SQL concatenado; escritura y resincronización sin transacción común.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Un nombre con apóstrofo rompe la consulta de duplicados (`constructor.bas:349`): el alta o el cambio de nombre fallarían con error de SQL.
- `Max + 1` sin bloqueo ni transacción (`FUNCIONES UTILES.bas:2851`; `ComercialOperaciones.cls:52, 60-68`): dos altas simultáneas podrían obtener el mismo ID.
- Con `Option Compare Database` (`:2`; `ComercialOperaciones.cls:10`), un cambio solo de mayúsculas no cuenta como cambio (`:107`) y se rechaza con «No hay cambios que guardar».
- Si `ResyncCatalogo` falla, el nombre ya está guardado pero se muestra error y no se lanza `Editado` (`ComercialOperaciones.cls:77, 88-92`; `:74-82`): la lista de gestión no se refresca.
- La global se modifica con los valores de pantalla antes de guardar (`:65-68`); si `Registrar` falla, el objeto compartido con la gestión (`Form_FormComercialesGestion.cls:221`) queda con valores no guardados.
- Títulos de diseño heredados: ventana «Edición de Suministrador» (`.form.txt:38`, sobrescrita en `:177`); `ComandoAyuda` rotula «Salir» (`.form.txt:825`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: recortar espacios y comparar sin distinguir mayúsculas solo para la unicidad; identificador generado por la base; restricción `UNIQUE` sobre el nombre.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario no administrador ve el comercial sin poder editar ni guardar.
2. Guardar sin nombre muestra «Falta el campo Comercial» y no escribe.
3. Guardar un nombre ya existente muestra «Ese nombre de comercial ya está registrado».
4. Guardar sin cambios muestra «No hay cambios que guardar».
5. Renombrar un comercial actualiza el nombre en los expedientes que lo usan.
6. Vaciar la descripción en edición la deja nula.

## Pendientes

- Si Access aplica `AllowEdits = False` a controles no enlazados en este formulario (`:176`).
- Contenido de `SincronizarDerivados` y qué derivados actualiza para comerciales (`ExpedienteOperaciones.cls:735`).
- Si `ComandoAyuda` muestra en ejecución un icono que oculte el rótulo «Salir».
