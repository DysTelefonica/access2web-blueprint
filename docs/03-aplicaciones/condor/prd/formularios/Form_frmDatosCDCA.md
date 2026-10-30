# Form_frmDatosCDCA

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frmDatosCDCA.cls` (735 líneas) y `.form.txt` · **PRD:** [README](../README.md) · **Variante subcontratista:** [Form_frmDatosCDCASUB](Form_frmDatosCDCASUB.md)

Formulario contenedor de los datos técnicos de una solicitud de concesión o desviación (tipo `CD_CA`). Aloja seis subformularios en un control de navegación, decide qué bloque es editable según rol y estado, orquesta el guardado por bloque y propone las transiciones de estado. Es **no enlazado**: el `.form.txt` no declara `RecordSource`; los datos llegan por un ViewModel (`:549-607`).

Abreviaturas de las citas: sin prefijo, `Form_frmDatosCDCA.cls`; `WF`, `classes/WorkflowServicio.cls`; `Srv`, `classes/DatosCDCAServicio.cls`; `Repo`, `modules/DatosCDCARepositorio.bas`; `FPA`, `modules/FormulariosPadreAuxiliares.bas`; `FU`, `modules/FUNCIONES UTILES.bas`; `GS`, `forms/Form_frmGestionSolicitud.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Gestión de solicitud (`frmGestionSolicitud`), botón «Ver datos» | Si `tipoSolicitud = "CD_CA"`, `DoCmd.OpenForm "frmDatosCDCA"` con `OpenArgs` = id de solicitud; si ya está abierto, solo lo muestra | `GS:304, 326-335` |
| Gestión de solicitud | Se suscribe al evento `SolicitudHaCambiado` (`WithEvents`) | `GS:33, 345, 1918-1940` |
| Subformularios `subfrmDatosCDCA_*` | Llaman a `Me.Parent.GuardarDesdeSubform` y, en Decisión final aprobada, a `Me.Parent.ProcesarCierreFormalizacion` | `subfrmDatosCDCA_DecisionFinal.cls:85, 90`; `subfrmDatosCDCA_Generales.cls:124` |

Abre o devuelve el foco a `frmGestionSolicitud` (`:68-85`) y muestra el selector de ficheros del sistema en el cierre (`:686-696`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `Load` | Valida `OpenArgs`, carga el ViewModel y navega a la pestaña del estado | `:20-54`, `.form.txt:241` |
| Formulario | — | `Unload` | Libera ViewModel y servicios | `:56-62`, `.form.txt:30` |
| `NavigationButton` | navegación | `Click` | Reconfigura permisos y repuebla el subformulario activo | `:90-100`, `.form.txt:680-681` |
| `tabGeneral` («GENERAL») … `tabDecisionFinal` («AUTORIDAD COMPETENTE») | botones de navegación (6) | `Click` | Delegan en `NavigationButton_Click` | `:116-121`, `.form.txt:706-1006` |
| `BotónDeNavegación1` («[Agregar nuevo]») | botón de navegación | — | Sin destino ni handler | `.form.txt:1043-1044` |
| `NavigationSubform` | subformulario | — | Hospeda `subfrmDatosCDCA_*` | `.form.txt:1093` |
| `cmdCerrar` («Salir») | botón | `Click` | Cierra el formulario | `:64-66`, `.form.txt:1131-1133` |
| `cmdVolverResumen` («Ver detalle») | botón | `Click` | Oculta este formulario y muestra o abre `frmGestionSolicitud` | `:68-85`, `.form.txt:1253-1255` |
| `txtEstado` | texto | — | Nombre del estado actual | `:624`, `.form.txt:1231` |
| `lblTitulo`, `Caption` | etiqueta, título | — | Código de solicitud y nemotécnico del expediente | `:622-623` |
| `lblBannerRechazo` | etiqueta oculta | — | Visible con rechazo de Calidad en Modificación | `:650-655`, `.form.txt:1294-1302` |
| `lblBloqueo` | etiqueta oculta | — | «Grabando, por favor espere…» durante el guardado | `:174`, `.form.txt:652-663` |

Pestañas y bloques: General → Generales; «DETALLE Y DESCRIPCIÓN» → Propuesta; «MOTIVOS» → Impacto; «APROBACIÓN SUMINSTRADOR» → Aprobación suministrador; «RAC» → Dictamen RAC; «AUTORIDAD COMPETENTE» → Decisión final (`.form.txt:707-1006`; `FPA:103-160, 179-189`).

## Flujos

### Apertura (`Form_Load`, `:20-54`)

1. Sin id válido en `OpenArgs` lanza 513 «No se recibió un idSolicitud válido.», muestra el error y cierra (`:31-33, 50-53`).
2. `RecargarViewModel` lee solicitud, expediente y datos CD/CA; si no hay fila, usa valores iniciales en memoria (`:563-570`; `Srv:116`).
3. Calcula los permisos de los seis bloques con `PuedeEditarBloque` (`:592-597`; `WF:1419-1445`).
4. Navega a la pestaña del estado: 1-2 General, 3-4 Propuesta, 5 Aprobación, 6 RAC, 7-8 Decisión final; en otro caso, General (`:39-43`; `WF:1844-1897`).
5. `NavegarAPestana` usa `DoCmd.BrowseTo` y simula Intro con `SendKeys` sobre la pestaña (`:127-157`).

### Guardado de un bloque (`GuardarDesdeSubform`, `:160-406`)

1. Identifica el bloque por el nombre del subformulario; si no lo reconoce, 513 (`:188-189`).
2. Lee la fila actual, la clona, recoge los valores del subformulario y copia solo los campos del bloque (`:192-196, 492-547`).
3. Si nada cambió, «No se han detectado cambios.» y no guarda, salvo Dictamen RAC con decisión `APROBADO` (`:199-210`; comparación por texto en `FPA:25-47`).
4. Dictamen RAC: abre una segunda conexión, `BeginTrans`, guarda y confirma (`:215-230`; `Srv:547-589`).
5. Resto de bloques: llama al guardado del servicio (`:236-247`). Aprobación suministrador comprueba además `hayCambios()` del subformulario (`:240-245`).
6. Aprobación suministrador en Modificación: recarga la lista de comprobación de `frmGestionSolicitud` (`:250-255`).
7. Generales en Registro (o recién pasado a Registro): pregunta «¿Asignar tarea al Técnico ahora?» y, si se acepta, transita a Desarrollo técnico (`:258-274`).
8. En Desarrollo técnico (salvo si se viene de Registro): sin Motivos completos y sin rechazo previo, avisa y lleva a «MOTIVOS»; desde Motivos sin Detalle, lleva a «DETALLE»; con ambos, pregunta si enviar a Calidad y transita a Modificación (`:280-325`).
9. Mensaje genérico «Datos guardados correctamente.», `RaiseEvent SolicitudHaCambiado`, recarga y repoblado (`:375-383`).

### Cierre con aprobación (`ProcesarCierreFormalizacion`, `:669-728`)

1. Aviso obligatorio de adjuntar el PDF de cierre; Cancelar sale (`:678-682`).
2. Selector de un único `*.pdf`; sin selección, «Operacion cancelada…» (`:685-696`).
3. Copia decisión, observaciones y firmante sobre la fila leída (`:703-708`).
4. `EjecutarCierreFormalizacion` en una transacción: guarda Decisión final, borra adjuntos de cierre previos, guarda el PDF como adjunto y transita a Aprobada (`:711`; `WF:3334-3397`).
5. Mensaje de éxito, evento, recarga (`:714-721`).

### Errores

Las validaciones del servicio se lanzan como 513 y se muestran con `ManejadorErroresFormulario` (`:403-405`). En el bloque RAC, el error revierte la transacción (`:399-401`). `RecargarViewModel` propaga con `CondorError` (`:603-606`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Generales obligatorios: nº de contrato, referencia y nombre del suministrador, «requiere modificación» | Servicio | `Srv:300-303` |
| Guardar Generales en Preregistro transita a Registro en la misma transacción | Servicio | `Srv:305-323` |
| Propuesta obligatoria: identificación del material y causa de la NC | Servicio | `Srv:359-360` |
| Aprobación: Ingeniería y Calidad obligatorias; Producción, Diseño y Representante obligatorios si «aplica», y se vacían si no | Servicio | `Srv:512-532` |
| «Aplica» se deduce de que el nombre no esté vacío; no se persiste | Formulario | `:463-488` |
| RAC: código y nombre obligatorios salvo decisión `RECHAZADO` | Servicio | `Srv:555-558` |
| Decisión final obligatoria | Servicio y subformulario | `Srv:594-596`; `subfrmDatosCDCA_DecisionFinal.cls:75-78` |
| Detalle completo: material, causa y descripción del impacto | Servicio | `Srv:664-682` |
| Motivos completos: coste en {AUMENTARÁ, DISMINUIRÁ, NO VARIARÁ}, clasificación {MAYOR, MENOR}, autoridad de diseño informada | Servicio | `Srv:684-707` |
| Edición: Rechazada nadie; Aprobada solo Decisión final (Calidad/Administrador); Formalización solo Aprobación y Decisión final; Propuesta e Impacto el Técnico en Desarrollo técnico y Calidad después; Aprobación y RAC Calidad desde Modificación | Workflow | `WF:1478-1553` |
| Sin permiso: botones de guardado desactivados y controles bloqueados salvo etiqueta `INMUTABLE` | Formulario | `:640-641`; `FPA:268-273`; `FU:946-971` |
| Transición: rol permitido y precondiciones del tipo `CD_CA` | Workflow | `WF:1244-1249, 895-925` |
| Tiempo de ciclo: hasta `fechaModificacion` en estado final; si no, hasta ahora «(en curso)» | Formulario | `:579-584` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura de datos CD/CA | `tbDatosCDCA` | SQL parametrizado `[p_IDSol]` | No | `Repo:264-292` |
| Generales (alta o edición completa) | `tbDatosCDCA` | DAO `AddNew/Edit`, id = `Max+1` | Sí (servicio) | `Repo:13-24, 316-343`; `Srv:306-323` |
| Propuesta, Impacto, Aprobación, RAC, Decisión | `tbDatosCDCA` | DAO `Edit` con SQL concatenado | Según bloque | `Repo:86-262` |
| Cambio de estado e historial | solicitud y log de estados | Repositorios | Sí | `WF:1282, 1291, 1298-1301` |
| Adjunto de cierre | adjuntos | Borrado y alta | Sí | `WF:3372-3380` |
| Notificación por correo | — | Fuera de la transacción, errores ignorados | No | `WF:1308-1312` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: transacciones DAO sobre `CreateWorkspace` y una segunda conexión abierta con contraseña de base (`:216-218`; `FU:56-62`); identificador por `Max+1` (`Repo:18-24`); SQL concatenado con el id (`Repo:47`); comparaciones de texto con tildes (`Srv:651`). Debe preservarse la atomicidad de guardar Generales con el paso a Registro y del cierre con su adjunto y su transición.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El bloque RAC se anuncia como «GUARDAR + AVANZAR» (`:213`), pero solo guarda (`:222-225`); la excepción de `APROBADO` al control de cambios (`:202-204`) vuelve a guardar sin transitar.
- El bloque «avance guiado técnico» (`:330-369`) solo se alcanza si Detalle o Motivos están incompletos, y entonces `EsParteTecnicaCompleta` (`Srv:633-662`), que exige ambos, es falso: el mensaje específico de corrección tras rechazo (`:340-342`) no se muestra nunca.
- Tras enviar a Calidad o navegar a otra pestaña se sale por `LimpiarYSalir` (`:296, 303, 315, 366`) sin `RaiseEvent` ni recarga (`:380-383`): la gestión de solicitud y la cabecera pueden quedar con el estado anterior.
- La transición a Desarrollo técnico (`:270`) no actualiza el estado en memoria antes de evaluar los escenarios siguientes.
- `EjecutarCierreFormalizacion` cierra la conexión global de `getdb()` (`WF:3350, 3389`; `FU:84-108`); falta comprobar en qué workspace está abierta para saber si `BeginTrans` la cubre (`WF:3353-3355`).
- `esRolCalidad` se calcula y no se usa (`:647`); `silenciarErrorValidacion` no lo pasa ningún llamador.
- `SendKeys "{Enter}"` (`:150`) depende del foco de la ventana.
- `Repo:317-319` indica que `Guardar` no debe usarse desde edición normal, pero `GuardarDatosGenerales` lo usa (`Srv:312`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: una ruta por solicitud con pestañas por bloque; guardado y transición como acciones separadas y explícitas; subida del PDF de cierre en la misma operación transaccional.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin id de solicitud muestra el error y no presenta datos (fila `Load`).
2. La pestaña inicial corresponde al estado (paso 4 de Apertura).
3. Cada bloque es editable solo con el rol y el estado de la regla de edición.
4. Guardar sin cambios muestra «No se han detectado cambios.» y no escribe.
5. Guardar Generales en Preregistro deja la solicitud en Registro y la fila guardada, o ninguna de las dos cosas.
6. Las validaciones obligatorias de cada bloque impiden el guardado con su mensaje.
7. Con Detalle y Motivos completos en Desarrollo técnico, se ofrece enviar a Calidad y, al aceptar, la solicitud pasa a Modificación.
8. Aprobar exige un PDF; sin él no cambia nada; con él, la decisión, el adjunto y el estado Aprobada se guardan juntos.
9. Con rechazo de Calidad en Modificación se muestra el aviso con los comentarios.

## Pendientes

- Precondiciones concretas de cada transición `CD_CA` (`WF:895-925` y `CumplePaso*`, `WF:981-1210`).
- Contenido de `GetDatosGeneralesIniciales` y de su suministrador por defecto (`Srv:116-262`).
- Deriva del código: el `CodeBehind` del `.form.txt` (`.form.txt:1321`) usa helpers locales y el `.cls`, los de `FPA`; falta saber cuál está en el binario.
- Comportamiento de `ManejadorErroresFormulario` y de `CondorError.Raise` ante el 513.
