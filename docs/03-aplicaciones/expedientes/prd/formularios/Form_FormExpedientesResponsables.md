# Form_FormExpedientesResponsables

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormExpedientesResponsables.cls` (187 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Editor de un responsable del expediente: muestra el nombre del usuario y permite fijar tres indicadores `Sí`/`No` (jefe de proyecto, notificaciones siempre y preventa). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): recibe los datos por `InitializeEditor` y delega en `modResponsablesEdit.ProcesarEdit`, que actúa sobre el DTO global `m_ObjExpedienteDTOActivo`. Las citas son `Form_FormExpedientesResponsables.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormExpedienteEntidades.Responsables_AfterUpdate` (elegir usuario en el combo `Responsables`) | `OpenForm` oculto, `InitializeEditor` con modo `create` o `edit` según exista ya la fila del usuario, y `Visible = True`; después vacía el combo | `Form_FormExpedienteEntidades.cls:1384-1448` |
| `Form_FormExpedienteEntidades.ComandoAltaReponsable_Click` | Si el combo tiene valor, llama a `Responsables_AfterUpdate`; si no, despliega el combo | `Form_FormExpedienteEntidades.cls:367-382` |
| `Form_FormExpedienteEntidades.ComandoEditarResponsable_Click` | Lee la fila de `TbExpedientesResponsables` del usuario seleccionado en `ListaResponsables` y abre en modo `edit` | `Form_FormExpedienteEntidades.cls:384-453` |

El `.form.txt` no declara `Modal` ni `PopUp`, y el padre no usa `acDialog` (`Form_FormExpedienteEntidades.cls:450, 1442`): pese al comentario de cabecera («Modal form», `:15`), la apertura no es modal en el sentido de Access.

Salidas: `cmdSalir_Click` cierra sin cambios (`:175-177`); `ComandoRegistrar_Click` cierra tras un resultado correcto (`:160`), antes de lo cual refresca la lista del padre llamando a `Forms("FormExpediente").FrmDetalle.Form.RellenarListaResponsables` (`:153-159`; `Form_FormExpedienteEntidades.cls:1585-1615`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | `Ajustar Me` ajusta el tamaño a las secciones; errores ignorados | `:183-186`, `.form.txt:40`; `FUNCIONES UTILES.bas:2703-2727` |
| (formulario) | — | `InitializeEditor` (público) | Guarda el estado, fija el título, rellena combos y valores, cablea los botones y habilita el guardado | `:44-98` |
| `lblTitulo` | etiqueta | — | Rótulo fijo de diseño; el código no lo cambia | `.form.txt:498-499`, `:64-65` |
| `Nombre` | texto deshabilitado y bloqueado | — | Muestra el nombre recibido | `:66`, `.form.txt:693-695, 711` |
| `EsJefeProyecto` (rótulo «Es JP») | combo | — | `RowSource = "Sí;No"`, `LimitToList = True`; valor inicial solo si el recibido no está vacío | `:70-71, 79`, `.form.txt:585, 615-617` |
| `CorreoSiempre` (rótulo «Notificaciones») | combo | — | Igual que el anterior | `:72-73, 80`, `.form.txt:644, 674-676` |
| `EsPreventa` (rótulo «Es Preventa») | combo | — | Igual que el anterior | `:74-75, 81`, `.form.txt:527, 556-558` |
| `ComandoRegistrar` (rótulo «Editar») | botón | `OnClick` cableado en ejecución → `ComandoRegistrar_Click` | Valida invariantes y llama a `ProcesarEdit` | `:84, 103-171`, `.form.txt:846-847, 858-859` |
| `cmdSalir` (rótulo «Salir») | botón | `OnClick` cableado en ejecución → `cmdSalir_Click` | `DoCmd.Close … acSaveNo` | `:85, 175-177`, `.form.txt:752-753` |

El `RowSource` de diseño de los tres combos es `Todos;Sí;No` (`.form.txt:558, 617, 676`); `InitializeEditor` lo sustituye por `Sí;No` (`:70-75`). El título de diseño «Edición de Suministrador» (`.form.txt:39`) se sustituye por «Editar responsable» en modo `edit` y «Guardar responsable» en otro caso (`:63`).

## Flujos

### Apertura (`InitializeEditor`, `:44-98`)

1. Copia a variables privadas modo, `IDExpediente`, `IdUsuario`, `IDExpedienteResponsable`, nombre y los tres indicadores en bruto (`:54-61`).
2. Fija título, nombre y combos (`:63-75`). Un indicador vacío deja el combo vacío (`:79-81`).
3. `ComandoRegistrar` arranca deshabilitado en diseño (`.form.txt:847`) y se habilita solo con `EsAdministrador = Sí` (`:92`).
4. Un error se muestra con `MsgBox` crítico; no cierra el formulario ni avisa al administrador (`:95-97`).

### Guardado (`ComandoRegistrar_Click`, `:103-171`)

1. Construye un `clsResponsablesEditState` con el estado de apertura y los valores actuales de los combos (`:111-118`).
2. `ExpedienteOperaciones.ValidarInvariantesResponsables` sobre el DTO **tal como está antes del cambio** (`:125-131`); si devuelve motivo, `MsgBox` y sale.
3. `ProcesarEdit(m_ObjExpedienteDTOActivo, m_Estado, …, Nothing)` (`:133`): con `p_db = Nothing` no hay acceso a la base (`modResponsablesEdit.bas:241, 321`).
4. Resultado `dto-only`, `created` o `updated`: mensaje «Responsable guardado correctamente.» con el título del formulario, refresco de la lista del padre y cierre (`:135-160`).
5. Resultado fallido: `MsgBox` con el mensaje del helper y el formulario queda abierto (`:161-165`).

### Qué hace `ProcesarEdit` con este llamador (`modResponsablesEdit.bas:110-372`)

- Normaliza vacío a «No» en los tres indicadores (`:172-187`) y exige `Sí` o `No` (`:189-207`; resultado `validation-failure`).
- Busca el usuario por valor en `ColResponsables` (`:209-216`).
- `create`: si ya está, `duplicate` (`:221-226`); si no, añade un `ExpedienteResponsable` con clave `CStr(IdUsuario)` (`:230-236`) y devuelve `dto-only` (`:274-278`).
- `edit`: clave `IDExpedienteResponsable` o, si viene vacío, `IdUsuario` (`:282-286`); actualiza la entrada o la **añade** si no existe (`:305-316`) y devuelve `updated` (`:349`).
- Cualquier error en tiempo de ejecución devuelve `error` con el detalle (`:359-371`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Guardar solo si `EsAdministrador = Sí` | `InitializeEditor` | `:92` |
| Indicadores restringidos a `Sí`/`No` en el control | `InitializeEditor` | `:70-75` |
| Vacío se guarda como «No» | `ProcesarEdit` | `modResponsablesEdit.bas:173-187` |
| Al menos un responsable con `EsJefeProyecto = Sí` y otro (o el mismo) con `CorreoSiempre = Sí`; colección `Nothing` cuenta como válida | `ValidarInvariantesResponsables` | `ExpedienteOperaciones.cls:1103-1124` |
| Un usuario no se añade dos veces al DTO | `ProcesarEdit` | `modResponsablesEdit.bas:221-226` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Mutación de `ColResponsables` del DTO global | memoria | — | `modResponsablesEdit.bas:236, 306-315` |
| Ramas `INSERT`/`UPDATE` de `TbExpedientesResponsables` | no se ejecutan desde este formulario (`p_db = Nothing`) | DAO, SQL dinámico | `:133`; `modResponsablesEdit.bas:241-273, 321-348` |

La persistencia de un alta la hace `ExpedienteOperaciones.Registrar` al guardar el expediente, con `RegistraResponsable` por cada entrada (`ExpedienteOperaciones.cls:468-476`), que solo escribe `EsJefeProyecto` y `CorreoSiempre` (`ExpedienteOperaciones.cls:1511-1517`). Los detalles se guardan solo en alta (`ExpedienteOperaciones.cls:400-402`).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: las ramas no usadas de `ProcesarEdit` concatenan identificadores en SQL (`modResponsablesEdit.bas:243, 323`) y calculan la clave con `DameID` en lugar de autonumérico (`:260`); los indicadores son texto con «í» (`ChrW$(237)`, `:190`), lo que exige codificación coherente en PostgreSQL.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Con expediente ya guardado, el cambio solo llega al DTO (`:133`) y `Registrar` no guarda detalles en edición (`ExpedienteOperaciones.cls:400-402`): no se ha encontrado otra escritura de `TbExpedientesResponsables` desde esta pantalla, pero el usuario ve «Responsable guardado correctamente.» (`:142`).
- La invariante se valida antes de aplicar el cambio (`:127` frente a `:133`), al contrario de lo que exige su contrato (`ExpedienteOperaciones.cls:1072-1075`): permite dejar el expediente sin jefe de proyecto y bloquea corregir un estado ya inválido (p. ej. añadir el segundo responsable en un alta cuyo primero tiene ambos indicadores en «No»).
- En un alta, `EsPreventa` se guarda en el DTO (`modResponsablesEdit.bas:234`) pero `RegistraResponsable` no lo escribe (`ExpedienteOperaciones.cls:1511-1517`).
- En modo `edit` sin entrada con la clave `IDExpedienteResponsable` en el DTO, se añade una entrada nueva (`modResponsablesEdit.bas:309-315`) en lugar de fallar; si el DTO usa otra clave para ese usuario, quedaría duplicado.
- El error de `RellenarListaResponsables` se recoge en `m_Error` y se descarta (`:158`); si la pestaña visible no es Entidades, la llamada falla tras guardar y muestra error sin cerrar (`:155-158, 168-170`).
- `g_NecesitaRefrescarResponsables` ya no se asigna en ningún sitio (`Variables Globales.bas:5-11`; lectura en `Form_FormExpedienteEntidades.cls:1372`): el refresco por `Form_Activate` está muerto.
- La habilitación del guardado no consulta `Form_FormExpediente.m_AbiertoParaEditar` (`:92`), a diferencia de las pestañas (`Form_FormExpedienteEntidades.cls:1339-1342`).
- El botón de guardado rotula «Editar» también en alta (`.form.txt:859`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: persistir el responsable en el momento o declarar explícitamente que se guarda con el expediente; validar la invariante sobre el estado resultante; editar los indicadores en línea en la lista de responsables.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario no administrador ve el editor con el guardado deshabilitado.
2. Los tres indicadores solo admiten «Sí» o «No»; dejarlos vacíos guarda «No».
3. Añadir un usuario que ya es responsable del expediente muestra un error y no lo duplica.
4. Un cambio que deje el expediente sin jefe de proyecto o sin notificación siempre se rechaza con el motivo correspondiente.
5. Tras guardar, la lista de responsables de la pestaña Entidades muestra los valores nuevos.
6. Cancelar no modifica ni la lista ni los datos.

## Pendientes

- Si alguna otra ruta persiste en edición los cambios de `ColResponsables` (ficha de `Form_FormExpedienteEntidades` y resto de `ExpedienteOperaciones`).
- Si `ColResponsables` llega `Nothing` o vacía en un alta (`ExpedienteDTO.cls:23`), lo que decide si la invariante bloquea el primer responsable.
- Comportamiento de Access al asignar por código un valor fuera de lista con `LimitToList = True` (`:79-81`), p. ej. «Todos» guardado en datos antiguos.
- Si los botones de alta y edición del padre se deshabilitan en solo lectura.
