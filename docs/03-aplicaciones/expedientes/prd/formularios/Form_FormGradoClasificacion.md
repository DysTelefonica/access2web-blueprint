# Form_FormGradoClasificacion

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormGradoClasificacion.cls` (197 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta y edición de un grado de clasificación del catálogo `TbGradosClasificacion` (campos grado y descripción). Formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): lee el registro con `constructor.getGradoClasificacion` y escribe con `GradoClasificacionOperaciones.Registrar`. Las citas son `Form_FormGradoClasificacion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| [Form_FormGradosClasificacionGestion](Form_FormGradosClasificacionGestion.md) | Alta: vacía `m_ObjGradoClasificacionActivo` y abre; edición: asigna el seleccionado a la global y abre. Escucha `Alta` y `Editado` | `Form_FormGradosClasificacionGestion.cls:75-83, 224-232, 382-426` |
| [Form_FormExpedienteGeneral](Form_FormExpedienteGeneral.md) | `ComandoNuevaClasificacion`: vacía la global y abre; escucha solo `Alta` | `Form_FormExpedienteGeneral.cls:45-60, 601` |

`Form_FormExpedienteAltaParaHPS` declara un `WithEvents` del formulario (`Form_FormExpedienteAltaParaHPS.cls:9`) pero no se ha encontrado que lo abra.

El registro llega por la global `m_ObjGradoClasificacionActivo` (`Variables Globales.bas:254`), no por `OpenArgs`. Eventos hacia quien lo abre: `Alta` (`:5`, lanzado en `:79`) y `Editado` (`:6`, lanzado en `:81`). `CmdSalir_Click` cierra sin guardar (`:10-13`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `GradoClasificacion` (etiqueta «GRADO») | texto | — | Nombre del grado; obligatorio y único | `.form.txt:930, 995` |
| `DESCRIPCION` (etiqueta «DESCRIP.») | texto | — | Descripción opcional | `.form.txt:965, 1015` |
| `ComandoEditar` (rótulo «Editar») | botón | `OnClick` → `[Event Procedure]` | **No hay `ComandoEditar_Click` en el `.cls`**; el guardado vive en `ComandoRegistrar_Click` (`:42-96`), sin control con ese nombre | `.form.txt:1153-1155` |
| `ComandoAyuda` (rótulo «Salir» en `.form.txt`) | botón | `OnClick` | `AbrirAyuda` | `:15-40`, `.form.txt:824-826` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | `:10-13`, `.form.txt:1046-1048` |
| `lblTitulo` | etiqueta | — | «ALTA…» o «EDICIÓN DE GRADO CLASIFICACIÓN» | `:131, 135` |

## Flujos

### Apertura (`Form_Open`, `:166-198`)

1. `Ajustar` dimensiona la ventana; `ComandoEditar` se deshabilita y solo se habilita si `EsAdministrador = Sí` (`:174-180`). El `Caption` de diseño se sustituye por `m_TituloFormulario` (`:177`).
2. `EstablecerDatos`: sin global activa, título de alta y fin (`:130-134`). Con global, relee el registro por `IdGradoClasificacion` en `m_GradoClasificacionAlInicio` (`:137`); si no existe, error «No se ha podido encontrar el Grado Clasificación registrado» (`:141-144`); copia grado y descripción no vacíos a los controles (`:147-154`).
3. Cualquier error envía `CorreoAlAdministrador`, muestra `MsgBox` y cierra (`:190-197`).

### Guardado (`ComandoRegistrar_Click`, `:42-96`)

1. `HaHabidoCambios`: un alta siempre cuenta como cambio (`:101-104`); en edición compara grado y descripción (`:107-114`). Sin cambios: «No hay cambios que guardar» (`:58-61`).
2. Crea la global si no existe y le copia los dos controles con `Nz` (`:62-68`).
3. `GradoClasificacionOperaciones.Registrar` con el estado inicial (`:70-77`; reglas abajo).
4. Lanza `Alta` o `Editado` según hubiera estado inicial y cierra (`:78-86`). Los errores se muestran con `MsgBox`; los no controlados envían correo (`:88-95`).

## Reglas, validaciones y cálculos

- Grado vacío: «Falta el campo GradoClasificacion» (`GradoClasificacionOperaciones.cls:29-32`).
- Unicidad por nombre, solo en alta o si el nombre cambia: busca por `GradoClasificacion` y rechaza con «Ese nombre de GradoClasificacion ya está registrado» (`GradoClasificacionOperaciones.cls:33-49`).
- Identificador en alta: `IdGradoClasificacionCalculado` = `Max(IdGradoClasificacion) + 1` vía `DameID` (`GradoClasificacion.cls:31`; `FUNCIONES UTILES.bas:2851-2864`).
- Descripción vacía: en alta no se asigna; en edición se guarda `Null` (`GradoClasificacionOperaciones.cls:65-67, 72-76`).
- Tras escribir invalida la caché `m_ObjEntorno.GradosClasificaciones` (`GradoClasificacionOperaciones.cls:82`). Si fue un renombrado, `ExpedienteOperaciones.ResyncCatalogo "grado"` resincroniza los expedientes que usan ese `IDGradoClasificacion` (`GradoClasificacionOperaciones.cls:86-93`; `ExpedienteOperaciones.cls:689-691, 712-714`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura por ID o por nombre | `TbGradosClasificacion` | SQL dinámico concatenado | no | `constructor.bas:538-547` |
| Siguiente ID | `Max` sobre `TbGradosClasificacion` | SQL dinámico | no | `FUNCIONES UTILES.bas:2851-2853` |
| Alta / edición | `TbGradosClasificacion` (`AddNew` / `Edit` + `Update`) | DAO | no | `GradoClasificacionOperaciones.cls:53-80` |
| Resincronización de expedientes | `TbExpedientes` por `IDGradoClasificacion` | SQL dinámico + `SincronizarDerivados` | fuera de la escritura del catálogo | `ExpedienteOperaciones.cls:640-646, 712-715` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: identificador `Max + 1` calculado en VBA, sin bloqueo (`FUNCIONES UTILES.bas:2851-2864`); nombre entre comillas simples sin escapar (`constructor.bas:545`); escritura del catálogo y resincronización sin transacción común.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El botón de guardar se llama `ComandoEditar` (`.form.txt:1153-1155`) y el handler es `ComandoRegistrar_Click` (`:42`): si el binario coincide con el fuente, pulsar «Editar» no guarda. Los formularios hermanos sí nombran el botón `ComandoRegistrar` (p. ej. `Form_FormEjercito.form.txt:842`).
- Si `DameID` falla, `IdGradoClasificacionCalculado` devuelve vacío y solo deja el error en `.Error` (`GradoClasificacion.cls:38-42`); `Registrar` no lo consulta (`GradoClasificacionOperaciones.cls:52`).
- Si `ResyncCatalogo` falla, el renombrado ya está escrito y el usuario ve error sin reversión (`GradoClasificacionOperaciones.cls:70-92`).
- Un nombre con apóstrofo rompe la consulta de unicidad (`constructor.bas:545`).
- `Me.lblTitulo.Caption` de edición se asigna dos veces (`:135, 145`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: identificador generado por la base; unicidad con índice único; renombrado y resincronización en una sola transacción.

## Criterios de aceptación

Propuestos, no aprobados:

1. Guardar un grado vacío devuelve «Falta el campo GradoClasificacion» y no escribe.
2. Dar de alta o renombrar a un nombre existente devuelve «Ese nombre de GradoClasificacion ya está registrado».
3. Guardar sin cambios devuelve «No hay cambios que guardar».
4. Renombrar un grado actualiza los expedientes que lo referencian.
5. Un usuario no administrador ve el registro pero no puede guardar.

## Pendientes

- Comprobar en el binario el nombre real del botón de guardado y si `ComandoRegistrar_Click` está enlazado.
- Contenido de `SincronizarDerivados` (fichas de expediente).
- Si `OpenRecordset("TbGradosClasificacion")` abre tipo tabla sobre una tabla vinculada (`GradoClasificacionOperaciones.cls:53, 60`).
