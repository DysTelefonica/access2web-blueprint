# Form_frmDatosPC

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frmDatosPC.cls` (811 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Contenedor de navegación de una solicitud de tipo Propuesta de Cambio (`PC`): seis pestañas con un subformulario cada una, guardado por bloque, permisos por rol y estado, y avances de estado guiados por mensajes. Es **no enlazado**: trabaja sobre un `DatosPCViewModel` que recarga desde los servicios (`:12-16, 498-589`). Las citas son `Form_frmDatosPC.cls:línea` salvo indicación; `FPA.bas` abrevia `modules/FormulariosPadreAuxiliares.bas` y `FU.bas`, `modules/FUNCIONES UTILES.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_frmGestionSolicitud, «Ver datos» | Para tipo `PC` abre `frmDatosPC` y se suscribe a su evento `SolicitudHaCambiado` | `Form_frmGestionSolicitud.cls:295-303, 344, 1758-1774` |
| Subformularios `subfrmDatosPC_*` | Llaman a `GuardarDesdeSubform` o, con decisión «APROBADO», a `ProcesarCierreFormalizacion` | `Form_subfrmDatosPC_Generales.cls:114`, `Form_subfrmDatosPC_DecisionFinal.cls:107-115` |

«Ver detalle» oculta el formulario y devuelve el foco a Form_frmGestionSolicitud, o lo abre si no está abierto (`:26-38`). «Salir» lo cierra (`:19-21`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad`, `OnUnload` | Carga el modelo y navega a la pestaña del estado; al descargar libera referencias | `:50-92`, `.form.txt:30, 241` |
| `tabGeneral` … `tabDecisionFinal` | botones de navegación | `OnClick` | Seis pestañas: General, Detalle y descripción, Motivos, Aprobación suministrador, RAC, Autoridad competente | `:701-723`, `.form.txt:706-1006` |
| `NavigationButton` | control de navegación | `OnClick` | Reconfigura permisos y repuebla el subformulario | `:134-142`, `.form.txt:680-681` |
| `NavigationSubform` | subformulario | — | Destino de `BrowseTo` | `:113`, `.form.txt:1093` |
| `txtEstado` | texto bloqueado | — | Nombre del estado | `:435`, `.form.txt:1270-1282` |
| `lblBannerRechazo` | etiqueta | — | Visible si Calidad rechazó y el estado es modificación | `:479-485`, `.form.txt:1302` |
| `lblBloqueo` | etiqueta oculta | — | Aviso de proceso en curso | `:171, 773-774`, `.form.txt:652-663` |
| `cmdVolverResumen` («Ver detalle»), `cmdCerrar` («Salir») | botones | `OnClick` | Ver «Entrada y salida» | `.form.txt:1131-1133, 1229-1231` |

## Flujos

### Apertura (`Form_Load`, `:50-82`)

1. Exige `OpenArgs` mayor que 0 (`:62-63`) y llama a `RecargarViewModel` (`:66`): lee solicitud, expediente y `tbDatosPC`; si no hay fila, precarga las dos referencias con el código largo del expediente y el código de la solicitud (`:513-522`).
2. Calcula el nombre del estado, el tiempo de ciclo desde el historial de estados (`:529-565`) y un permiso de guardado por bloque con `PuedeEditarBloque` (`:573-578`).
3. Pestaña inicial por estado: 1-2 General; 3-4 Detalle; 5 Aprobación; 6 RAC; 7-8 Autoridad competente (`WorkflowServicio.cls:1844-1883`). `NavegarAPestana` usa `BrowseTo` y simula Intro con `SendKeys` (`:97-132`).

### Guardado por bloque (`GuardarDesdeSubform`, `:157-419`)

1. Clona la fila de `tbDatosPC` y vuelca solo los campos del bloque desde el modelo (`:191-201, 648-695`).
2. Sin cambios ⇒ «No se han detectado cambios.» (`:204-216`), salvo RAC con decisión «APROBADO».
3. RAC: guarda en una transacción propia (`:221-236`). Resto: `GuardarDatosGenerales`, `GuardarPropuesta`, `GuardarImpacto`, `GuardarAprobacionSuministrador` (con su propia comprobación de cambios) o `GuardarDecisionFinal` (`:242-255`).
4. General: si el estado era preregistro, el servicio pasa a registro en la misma transacción (`DatosPCServicio.cls:259-276`); en registro se pregunta «¿Asignar tarea al Técnico ahora?» y, si se acepta, transición a desarrollo técnico (`:265-281`).
5. En desarrollo técnico (si no se vino de registro): sin motivos ⇒ lleva a «Motivos»; motivos sin detalle ⇒ lleva a «Detalle»; ambos completos ⇒ propone enviar a Calidad (transición a modificación) (`:287-332`). Una segunda comprobación con `EsParteTecnicaCompleta` repite la propuesta con texto distinto si hubo rechazo (`:337-378`).
6. Aprobación suministrador en modificación: recarga la lista de comprobación de Form_frmGestionSolicitud (`:258-262`).
7. Mensaje genérico, evento `SolicitudHaCambiado`, recarga y repoblado (`:384-396`).

### Cierre (`ProcesarCierreFormalizacion`, `:744-804`)

1. Aviso obligatorio y selector de un PDF; cancelar en cualquiera de los dos sale (`:753-771`).
2. Copia la fila, aplica decisión, observaciones, autoridad y firmante del modelo (`:778-784`) y llama a `EjecutarCierreFormalizacion` (`:787`): en una transacción guarda la decisión, borra filas de adjuntos de cierre previos, adjunta el PDF con tipo «Generado» y transiciona a aprobada (`WorkflowServicio.cls:3352-3385`).

- **Errores:** las validaciones de servicio lanzan 513 (p. ej. campos obligatorios de General, `DatosPCServicio.cls:253-257`; código RAC, `DatosPCServicio.cls:526-530`; decisión final, `DatosPCServicio.cls:565`) y se muestran con `ManejadorErroresFormulario`; con `silenciarErrorValidacion` el 513 se calla (`:417`). En RAC se hace `Rollback` (`:412-414`). Un error en la carga cierra el formulario (`:79-80`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Rechazada: nadie edita; aprobada: solo Calidad/Administrador en Autoridad competente | `PermiteEdicion_PC` | `WorkflowServicio.cls:1610-1617` |
| Formalización: solo Aprobación suministrador y Autoridad competente, Calidad/Administrador | ídem | `WorkflowServicio.cls:1620-1628` |
| General: Calidad/Administrador; Detalle y Motivos: técnico en desarrollo técnico, Calidad/Administrador después | ídem | `WorkflowServicio.cls:1634-1651` |
| Aprobación y RAC: Calidad/Administrador desde modificación | ídem | `WorkflowServicio.cls:1654-1663` |
| Cinco campos obligatorios en General | servicio | `DatosPCServicio.cls:253-257` |
| Parte técnica completa: dos descripciones, incidencia de coste y clasificación válidas y al menos un motivo | servicio | `DatosPCServicio.cls:711-731` |
| Controles marcados `INMUTABLE` no se bloquean | `HabilitarControlesSubformulario` | `FU.bas:946-956` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Datos PC por solicitud | `tbDatosPC` | SQL parametrizado | no | `DatosPCRepositorio.bas:53-64` |
| Alta o actualización de General | `tbDatosPC` | SQL concatenado + DAO | sí | `DatosPCRepositorio.bas:10-40`; `DatosPCServicio.cls:260-276` |
| Dictamen RAC | `tbDatosPC` | SQL concatenado + DAO | sí | `DatosPCRepositorio.bas:338-371` |
| Decisión final | `tbDatosPC` | `QueryDef` parametrizado + DAO | según llamador | `DatosPCRepositorio.bas:384-412` |
| Transición de estado | solicitud, registro de estados y de cambios | servicio | propia o compartida | `WorkflowServicio.cls:1282-1296` |
| PDF de cierre | `tbAdjuntos` y carpeta | servicio | compartida | `WorkflowServicio.cls:3372-3380` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL concatenado con `idSolicitud` (`DatosPCRepositorio.bas:20, 353`); transacciones DAO en un `Workspace` propio con contraseña de Jet/ACE (`FU.bas:56-61, 252`); textos de negocio comparados en mayúsculas con tildes («AUMENTARÁ», `DatosPCServicio.cls:720`). Debe preservarse la atomicidad de guardar y transicionar en General y en el cierre.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ImportarEntidad_aVM` carga `NombreAutoridadDiseno` en `vm_decisionFinalNombreFirmante`, que se sobrescribe después (`:642, 645`); `vm_decisionAutoridadDisenioNombre` nunca se carga. El subformulario lo mostraría vacío (`Form_subfrmDatosPC_DecisionFinal.cls:22`) y al guardar se escribiría vacío (`:690`; `DatosPCRepositorio.bas:408`).
- El cierre no copia `obsAprobacionAutoridadDiseno` del modelo (`:781-784`): esa observación tecleada no se guarda.
- El RAC «APROBADO» sin cambios salta la comprobación «para intentar el avance de fase» (`:209`), pero la rama RAC solo guarda (`:221-236`).
- `mostrarBtnTecnica` se calcula y no se usa (`:466-470`).
- Si `RecargarViewModel` falla tras el cierre, el usuario ya vio «Solicitud APROBADA» (`:791-795`).
- Con `getDatosPC` vacío, `ClonarEntidad(Nothing)` en el cierre (`:778-779`).
- Coexisten dos `.form.txt` con código distinto (`frmDatosPC.form.txt:1319` en adelante, frente al `.cls`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: pestañas como rutas con permisos calculados en servidor; avances de estado como acciones explícitas en lugar de preguntas tras guardar.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, la pestaña activa corresponde al estado según la tabla de `getPaginaActivaGenerica`.
2. Cada bloque solo es editable para el rol y estado de `PermiteEdicion_PC`.
3. Guardar sin cambios avisa y no escribe.
4. Guardar General en preregistro pasa a registro en la misma transacción.
5. En desarrollo técnico, con detalle y motivos completos, se ofrece enviar a Calidad; si se acepta, el estado pasa a modificación.
6. Aprobar en Autoridad competente exige un PDF y deja decisión, adjunto y estado aprobado, o nada.

## Pendientes

- Cuál de los dos `.form.txt` corresponde al binario en uso.
- Contenido de `PrecondicionesCumplidas_PC` y de `EsTransicionPermitida` (`WorkflowServicio.cls:954, 1701`).
- Si el `m_ViewModel` del subformulario es el mismo objeto que el del padre.
