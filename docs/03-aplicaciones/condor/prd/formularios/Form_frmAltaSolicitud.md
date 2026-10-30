# Form_frmAltaSolicitud

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frmAltaSolicitud.cls` (646 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pre-registro de una solicitud: el usuario elige un expediente PECAL, un tipo de solicitud y, opcionalmente, una no conformidad (NC); el formulario propone el código `DC-<clave>-<ordinal>` y crea o edita el registro en `tbSolicitudes`. Es **no enlazado** (sin `RecordSource` en el `.form.txt`). Las citas son `:línea` de `Form_frmAltaSolicitud.cls` salvo indicación. Abreviaturas: `SS.cls` = `classes/SolicitudServicio.cls`, `SR.bas` = `modules/SolicitudRepositorio.bas`, `ES.cls` = `classes/ExpedienteServicio.cls`, `NCS.cls` = `classes/NoConformidadServicio.cls`, `NCR.bas` = `modules/NoConformidadRepositorio.bas`, `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`.

## Entrada y salida

| Desde | Modo | Evidencia |
|---|---|---|
| `Form_frm0Ppal` («Nueva solicitud») | Alta: sin `OpenArgs`, ventana normal; se suscribe con `WithEvents` | `Form_frm0Ppal.cls:5, 103-112` |
| `Form_frmGestionSolicitud` («Editar pre-registro») | Edición: `OpenArgs` = id de solicitud, `acDialog` | `Form_frmGestionSolicitud.cls:1787` |

`OpenArgs > 0` activa la edición; si no, alta (`:28, 32-40`). Abre `frmBuscarExpediente` (`:275, 312`) y `frmDetalleExpediente` (`:342`). Tras el alta, `frm0Ppal` abre `frmGestionSolicitud` con el nuevo id y llama a `AbrirFormularioDeDatos` (`Form_frm0Ppal.cls:289-296`). Sin control de rol en el formulario.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `txtPalabraClave` | texto | `KeyDown` | Intro lanza la búsqueda rápida | `:633-639`, `.form.txt:721` |
| `cmdBuscarExpediente` («Buscar») | botón | `OnClick` | Busca expedientes por código o nemotécnico | `:247-303` |
| `cmdIrABusquedaExpedientes` («Búsqueda Avanzada») | botón | `OnClick` | Abre el buscador en diálogo sin filtro | `:305-337` |
| `cmdVerDetalleExpediente` | botón (desactivado al inicio) | `OnClick` | Abre `frmDetalleExpediente` en diálogo | `:339-349`, `.form.txt:1381` |
| `txtIDExpediente` | texto desactivado y bloqueado | — | Id del expediente elegido | `:152`, `.form.txt:769-770` |
| `txtContratistaPrincipal`, `txtSubContratista` | texto desactivado | — | Nombres de suministradores separados por « \| » o «N/A» | `:160-162`, `ES.cls:189-211` |
| `cboTipoSolicitud` | combo, lista de valores | — | Tipos según el expediente | `:164-180` |
| `cboNCAsociada` | combo, lista de valores | `AfterUpdate` | NC del expediente; confirma si ya estaba vinculada | `:357-401, 420-443` |
| `txtCodigoSolicitudPropuesto` | texto desactivado | — | Código propuesto; fondo rojo si está duplicado | `:156, 612`, `.form.txt:866` |
| `txtOrdinalManual` («Forzar Ordinal») | texto | `AfterUpdate` | Sustituye el ordinal del código | `:583-631` |
| `cmdAltaSolicitud` | botón (desactivado al inicio) | `OnClick` | «Crear Solicitud» o «Guardar Cambios» | `:34, 38, 445-581` |
| `cmdSalir` («Cancelar») | botón | `OnClick` | Cierra sin guardar ni preguntar | `:641-643` |
| `lblBloqueo` | etiqueta oculta | — | Aviso de espera durante búsqueda y guardado | `:261, 470` |

`ControlBox` y `CloseButton` tienen valor no predeterminado (`.form.txt:10, 13`): el cierre previsto es «Cancelar» (a verificar en el binario).

## Flujos

### Apertura (`Form_Load`, `:24-52`)

1. `Ajustar Me`, lee `OpenArgs` y deja el combo NC con `0;'(Ninguna)'` (`:27-30, 351-356`).
2. Alta: título «Alta de Nueva Solicitud», limpia y desactiva el paso 2 (`:37-39, 215-237`).
3. Edición: carga la solicitud (`:73-121`): expediente, código guardado, NC guardada, tipo; si la NC existe en la base de NC, guarda su código para la confirmación (`:98-106`); si el estado no es pre-registro (`estadoPreregistro = 1`, `modules/modEnumeradores.bas:37`), desactiva el tipo (`:114-116`).
4. Guarda los valores originales para detectar cambios (`:44, 54-61`).

### Selección de expediente

1. **Búsqueda rápida** (`:247-303`): si ya hay expediente, pide confirmación (`:254-256`); exige texto (`:259`); `getExpedientes` (`ES.cls:18-87`). 0 resultados ⇒ «No se encontraron resultados.» y reinicia (`:268-271`); 1 ⇒ lo carga (`:272-273`); varios ⇒ abre el buscador en diálogo con el texto y lee su `Tag` (`:274-294`).
2. **Búsqueda avanzada** (`:305-337`): igual, sin texto inicial.
3. **Carga** (`CargarDatosExpediente`, `:123-214`): vacía tipo, código, ordinal y NC (`:132, 239-245`); rellena NC del expediente (`:136`; `NCR.bas:40-43`); lee el expediente y el nemotécnico (`:148-152`); propone código en alta o si cambia el expediente en edición (`:155-157`); muestra suministradores (`:160-162`); fija los tipos (`:168-180`) y habilita el paso 2 (`:194-201`).

### Código y ordinal

- `ProponerCodigoSolicitud` (`SS.cls:397-422`): clave = `CodigoActividad` si existe (`ES.cls:110-113`); si no, `AA` + `XXXX` de un `CodExp` con forma `XXXX/AA…` (`ES.cls:115-146`); si no se puede, error 513 (`ES.cls:152`). Ordinal = máximo numérico tras el último guion entre los códigos `DC-<clave>-*` + 1 (`SR.bas:107-116`).
- Ordinal manual numérico ⇒ código `<base>-<n>`; si existe y no es el propio, fondo rojo y «¡ATENCIÓN! El código '…' ya existe.» (`:592-616`). Vacío o no numérico ⇒ vuelve a proponer (`:620-624`).

### Guardado (`cmdAltaSolicitud_Click`, `:445-581`)

1. Edición sin cambios ⇒ «No se han detectado cambios en la solicitud.» y cierra (`:461-467`).
2. Abre otra conexión al backend Condor en el workspace compartido e inicia transacción (`:479-482, 511-514`; `FU.bas:56-62`).
3. **Alta:** `CrearNuevaSolicitud` asigna id = máximo + 1 (`SS.cls:323`; `FU.bas:486-495`), estado pre-registro, fechas, usuario y `revisionCalidadEstado = "PENDIENTE"` (`SS.cls:327-335`); `Validar` (`SS.cls:348-383`); `INSERT` por recordset (`SR.bas:75-80`); escribe el código en la NC elegida (`SS.cls:707-731`; `NCR.bas:132-139`); `CommitTrans` (`:495`).
4. Refresca contadores (`VG.bas:165-177`), muestra el código creado (`:501-504`), lanza `SolicitudCreadaConExito` (`:505`) y cierra (`:563`).
5. **Edición:** relee la solicitud, actualiza expediente, tipo, código, NC, fecha y usuario (`:516-527`), valida contra duplicados de otra solicitud (`SS.cls:369-377`), `UPDATE` (`SR.bas:62-72`), escribe el código en la NC nueva y vacía `CodConcesionAsociada` de la anterior si cambió (`SS.cls:715-725`; `NCR.bas:156-158`). Si cambió el expediente, recalcula los campos dependientes de `DatosPC`, `DatosCDCA` o `DatosCDCASUB` (`:539-552`; `classes/DatosPCServicio.cls:15-43`, `classes/DatosCDCAServicio.cls:16-36`, `classes/DatosCDCASUBServicio.cls:560-580`). `CommitTrans`, aviso y `SolicitudModificada` (`:554-560`).

### Errores

Cualquier error revierte la transacción si estaba abierta (`:574-576`) y pasa por `ManejadorErroresFormulario` (`:579`): el 513 se muestra como «Validación: Datos Requeridos», el 514 como error operativo y el resto como técnico (`FU.bas:621-636`); se registra en `tbLogErrores` si existe (`FU.bas:789-798`). Un error en `Form_Load` cierra el formulario (`:48-51`); en la carga del expediente, reinicia los campos (`:210-213`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Tipos ofrecidos: `PC` siempre; `CD_CA` si `ContratistaPrincipal = "Sí"`, si no `CD_CA_SUB`. `PC_SUB` no existe en `src` (resuelve D-CON-1 para este formulario) | `CargarDatosExpediente` | `:168-180` |
| El tipo solo se cambia en estado pre-registro | carga y edición | `:114-116, 185-198` |
| Obligatorios: expediente, tipo y código | `Validar` | `SS.cls:354-356` |
| Código único en alta; en edición, único salvo la propia solicitud | `Validar` | `SS.cls:369-382` |
| Cambiar una NC vinculada exige confirmación; «No» restaura la original | `ValidarCambioNC` | `:424-443` |
| Detección de cambios: expediente, tipo, código y NC (el ordinal se guarda pero no se compara) | `HayCambiosSinGuardar` | `:60, 63-71` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Expedientes PECAL con responsables y RAC | `TbExpedientes`, `TbExpedientesResponsables`, `TbUsuariosAplicaciones`, `TbExpedientesRACS`, `TbRACS` | SQL dinámico parametrizado | No | `ES.cls:35-79` |
| NC del expediente y por id | `TbNoConformidades` (backend NC) | SQL parametrizado | No | `NCR.bas:40-49, 104, 182-189` |
| Siguiente ordinal y existencia de código | `tbSolicitudes` | QueryDef temporal | No | `SR.bas:107-112, 141-145` |
| Id nuevo | `tbSolicitudes` | `Max + 1` | Sí (lectura) | `FU.bas:486-495` |
| Alta o edición de la solicitud | `tbSolicitudes` | DAO `AddNew`/`Edit` | Sí | `SR.bas:62-80` |
| Vincular y desvincular NC | `TbNoConformidades.CodConcesionAsociada` | `UPDATE` parametrizado | Sí (mismo workspace) | `NCR.bas:132-139, 156-158`; `FU.bas:252` |
| Campos dependientes del expediente | `DatosPC`, `DatosCDCA`, `DatosCDCASUB` (repositorios) | DAO | Sí | `:539-552` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: la transacción abarca dos ficheros Jet (Condor y NC) abiertos en el mismo workspace (`FU.bas:252`; `:481`), y no consta que su confirmación sea atómica entre ficheros (a verificar); en PostgreSQL vinculado debe decidirse el ámbito transaccional. Comodín `*` en `LIKE` (`ES.cls:57`; `SR.bas:110`), `FIRST()` (`ES.cls:37, 40`), `Val/Mid/InStrRev` en SQL (`SR.bas:107`), booleanos como texto `'Sí'` (`ES.cls:53`) e id por `Max + 1` en lugar de autonumérico (`FU.bas:486-495`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `DetectarNCVinculadaEnBD` nunca se ejecuta: `CargarDatosExpediente` vacía el código (`:132, 242`) antes de leerlo (`:141-143`).
- Id y ordinal por `Max + 1` (`FU.bas:486-495`; `SR.bas:107-116`): dos altas simultáneas pueden chocar; el ordinal además se calcula fuera de la transacción.
- Si solo cambia el ordinal en edición, los campos `refSuministrador` de los datos técnicos no se actualizan: dependen de `haCambiadoExpediente` (`:539`; `classes/DatosPCServicio.cls:34`).
- La conexión abierta en `:480, 512` solo se cierra si hay error (`:567`, `usaTransaccion` ya es `False` tras el commit).
- Doble mensaje de éxito en el alta: el del formulario (`:501`) y el de `frm0Ppal` (`Form_frm0Ppal.cls:283`).
- `SolicitudModificada` (`:7, 560`) no tiene suscriptor: `frmGestionSolicitud` lanza el suyo al volver del diálogo, haya o no guardado (`Form_frmGestionSolicitud.cls:1787-1789`).
- `cmdNuevaSolicitud_Click` enfoca el formulario abierto y aun así vuelve a llamar a `OpenForm` (`Form_frm0Ppal.cls:106-112`).
- `GetPasswordDB` usa una contraseña fija de respaldo si falla la lectura (`FU.bas:131-133`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: código y id asignados por la base dentro de la transacción; buscador de expedientes integrado en el mismo paso.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin expediente no se habilitan tipo, NC, ordinal ni «Crear Solicitud» (`:227-234`).
2. Un expediente con `ContratistaPrincipal = "Sí"` ofrece `PC` y `CD_CA`; otro, `PC` y `CD_CA_SUB`.
3. El código propuesto sigue `DC-<clave>-<máximo + 1>`; p. ej. `CodExp` ficticio `0123/24` da `DC-240123-1` si no hay códigos previos.
4. Un ordinal que produce un código existente avisa y el guardado lo rechaza con el mensaje de duplicado.
5. Alta correcta: solicitud en pre-registro, NC con el código y apertura de la gestión de la solicitud.
6. Cualquier fallo al vincular la NC deja sin cambios la solicitud y ambas NC.
7. Edición sin cambios cierra con «No se han detectado cambios en la solicitud.».

## Pendientes

- Contenido de `ExpedienteRepositorio.getExpedientePorID` y de los repositorios `Datos*` (no leídos para esta ficha).
- Propiedades `Modal`/`PopUp` y botón de cierre en el binario.
- Si `TbUsuariosAplicaciones` del backend de expedientes es la misma tabla que la de Lanzadera.
