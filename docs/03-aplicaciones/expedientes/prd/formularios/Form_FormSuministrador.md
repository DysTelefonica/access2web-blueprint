# Form_FormSuministrador

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormSuministrador.cls` (237 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta y edición de una empresa del catálogo de suministradores (`TbSuministradores`): nombre, nemotécnico, CIF, dirección, ciudad, CP y si es tramitadora HPS. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); el modo se decide por la global `m_ObjSuministradorActivo`: vacía, alta; informada, edición (`:154-159`). Sigue el patrón de alta/edición de catálogo descrito en [Form_FormEjercito](Form_FormEjercito.md); aquí solo se detalla lo propio. Las citas son `Form_FormSuministrador.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSuministradoresGestion` (Alta) | Vacía la global, cierra el formulario si estaba abierto y lo abre; escucha `Alta`/`Editado` | `Form_FormSuministradoresGestion.cls:76-84` |
| `Form_FormSuministradoresGestion` (Editar) | Asigna a la global el seleccionado y lo abre | `Form_FormSuministradoresGestion.cls:227-235` |
| `Form_FormExpedienteSuministradores` («Ver Detalle», «Ver Ficha») | Carga la global con `getSuministrador` y lo abre en diálogo | `Form_FormExpedienteSuministradores.cls:589-592, 665-666` |
| `Form_FormExpedienteSuministradores` («Nuevo Suministrador») | Vacía la global y lo abre en alta; escucha `Alta` | `Form_FormExpedienteSuministradores.cls:713-724, 730` |

Eventos: `Alta(p_ID)` (`:5`, lanzado en `:84`) y `Editado` (`:6`, lanzado en `:86`). Tras guardar cierra (`:91`); `cmdSalir` cierra sin guardar (`:10-13`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Nombre`, `CIF` | textos | — | Obligatorios al guardar | `.form.txt:624, 701`; `:66, 68` |
| `Nemotecnico` | texto | — | Opcional; único si se informa | `.form.txt:736`; `:67` |
| `Direccion`, `Ciudad`, `CP` | textos | — | Opcionales | `.form.txt:792, 869, 925`; `:69-71` |
| `TramitadoraHPS` | combo `Sí;No` (`LimitToList`) | — | Si tramita HPS | `.form.txt:966-993`; `:72` |
| `ComandoRegistrar` (rótulo «Editar») | botón, deshabilitado en diseño | `OnClick` | Guarda | `.form.txt:1126-1141`; `:42-101` |
| `ComandoAyuda` (rótulo «Salir», tip «Ayuda») | botón | `OnClick` | `AbrirAyuda` | `.form.txt:517-521`; `:15-40` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | `.form.txt:1030-1032`; `:10-13` |
| `lblTitulo` | etiqueta | — | «ALTA DE SUMINISTRADOR» o «EDICIÓN DE SUMINISTRADOR» | `:155, 159` |

## Flujos

### Apertura (`Form_Open`, `:204-237`)

1. `Ajustar Me`; `ComandoRegistrar` se habilita solo con `EsAdministrador = Sí` (`:212-219`). **No** fija `AllowEdits = False`, a diferencia de `Form_FormEjercito.cls:176`.
2. `EstablecerDatos`: sin suministrador activo, título de alta y fin (`:154-158`). Con él, relee el registro con `constructor.getSuministrador` por ID en `m_SuministradorAlInicio`; si no existe, «No se ha podido encontrar el suministrador registrado» (`:161-168`). Copia los campos no vacíos; `TramitadoraHPS` solo si vale «Sí» o «No» (`:170-191`).
3. Error: correo al administrador, `MsgBox` crítico y cierre (`:229-236`).

### Guardado (`ComandoRegistrar_Click`, `:42-101`)

1. `HaHabidoCambios` compara los siete campos con `m_SuministradorAlInicio`; en alta siempre hay cambios (`:103-140`). Sin cambios, «No hay cambios que guardar» (`:58-61`).
2. Crea el objeto si la global está vacía y copia los controles con `Nz` (`:62-73`).
3. `SuministradorOperaciones.Registrar` (`:75-82`), lanza `Alta` o `Editado` y cierra (`:83-91`). Un error deja el formulario abierto (`:93-100`).

## Reglas, validaciones y cálculos

Todas en `classes/SuministradorOperaciones.cls`:

- **Obligatorios:** «Falta el campo Nombre» y «Falta el campo CIF» (`:29-36`).
- **Tramitadora HPS:** cualquier valor distinto de «Sí»/«No» se guarda como «No» (`:37-39`).
- **CIF único:** en alta o si cambia el CIF, «Ese CIF de Suministrador ya está registrado» (`:40-56`; `constructor.bas:5577-5580`).
- **Nemotécnico único:** en alta si se informa, o si cambia; «Ese Nemotécnico de Suministrador ya está registrado» (`:57-76`; `constructor.bas:5585-5588`). El nombre **no** se valida como único.
- **Identificador:** `IDSuministradorCalculado` = `DameID` (máximo + 1) y, si es menor que 1000, se le suma 1000 (`:79`; `classes/Suministrador.cls:100-106`; `FUNCIONES UTILES.bas:2851-2864`).
- **Vacíos:** en alta, los opcionales vacíos no se escriben; en edición se guardan como `Null` (`:99-116, 122-151`). `DESCRIPCION` se escribe aunque no hay control: en edición conserva el valor leído.
- **Caché y resincronización:** invalida `m_ObjEntorno.Suministradores` (`:157`); si en edición cambia nombre o nemotécnico, `ResyncCatalogo "suministrador"` resincroniza los expedientes de `TbExpedientesSuministradores` (`:83-88, 161-168`; `ExpedienteOperaciones.cls:686-688`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Relectura del registro | `TbSuministradores` por `IDSuministrador` | SQL dinámico concatenado | no | `constructor.bas:5573-5576` |
| Unicidad de CIF y nemotécnico | `TbSuministradores` | SQL dinámico, valor entre comillas | no | `constructor.bas:5577-5588` |
| Nuevo identificador | `Max(IDSuministrador)` | SQL dinámico | no | `FUNCIONES UTILES.bas:2851-2853` |
| Alta / edición | `TbSuministradores` (`AddNew`/`Edit` + `Update`) | DAO | no | `SuministradorOperaciones.cls:93-155` |
| Resincronización | `TbExpedientesSuministradores` y derivados | SQL dinámico + DAO | no | `SuministradorOperaciones.cls:161-168` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: identificador `Max + 1` con suelo 1000 calculado en cliente; unicidad con `=` sobre texto (sensibilidad a mayúsculas distinta entre Jet y PostgreSQL); alta, edición y resincronización sin transacción común.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Un no administrador puede teclear en todos los campos aunque no pueda guardar: falta `AllowEdits = False` (`:215-219`).
- Un error de `DameID` queda en `Suministrador.Error` y no se consulta; `CLng("")` provocaría además un error de tipo (`classes/Suministrador.cls:100-105`; `SuministradorOperaciones.cls:79`).
- Un apóstrofo en el CIF o el nemotécnico rompe la consulta de unicidad (`constructor.bas:5580, 5588`).
- Si `getSuministrador` no encuentra el ID, «Ver Detalle» abre el formulario en modo alta (`Form_FormExpedienteSuministradores.cls:591-592`; `:154-155`).
- Patrón de familia ya descrito en [Form_FormEjercito](Form_FormEjercito.md): `Max + 1` sin bloqueo, resincronización sin transacción, objeto seleccionado alterado antes de validar (`:65-73`; `Form_FormSuministradoresGestion.cls:227`) y `CorreoAlAdministrador` sin protección en `Form_Open` (`:232`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: CIF y nemotécnico únicos como restricción de base; validar el formato del CIF; formulario de solo lectura para no administradores.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un no administrador no puede guardar.
2. Guardar sin nombre o sin CIF muestra el mensaje correspondiente y no escribe.
3. Un CIF o nemotécnico ya registrado se rechaza; el mismo registro sin cambiarlos no.
4. Tramitadora HPS vacía se guarda como «No».
5. Vaciar un campo opcional en edición lo deja nulo.
6. Renombrar una empresa actualiza los expedientes que la referencian.

## Pendientes

- Uso de `ConsorcioPropio` (`classes/Suministrador.cls:23`): no se edita ni se escribe aquí.
- Si existe validación de formato del CIF fuera de este flujo.
- Detalle de `SincronizarDerivados` para la categoría `suministrador` (`ExpedienteOperaciones.cls:686-688`).
