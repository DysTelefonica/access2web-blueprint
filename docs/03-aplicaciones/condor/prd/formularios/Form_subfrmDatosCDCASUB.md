# Form_subfrmDatosCDCASUB (grupo de seis subformularios)

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_subfrmDatosCDCASUB_<Bloque>.cls` y `.form.txt` · **PRD:** [README](../README.md)

Seis subformularios que el contenedor `Form_frmDatosCDCASUB` carga en su `NavigationSubform`, uno por pestaña, para editar por bloques los datos técnicos de una concesión o desviación de subsuministrador (tipo `CD_CA_SUB`). Ninguno tiene `RecordSource` (0 apariciones en los seis `.form.txt`): son **no enlazados** y trabajan sobre un `DatosCDCASUBViewModel` que les pasa el contenedor. Guardar, validar, transitar y recargar es responsabilidad del contenedor y del servicio; el subformulario solo vuelca y recoge valores.

Abreviaturas: `Gen`, `Pro`, `Imp`, `Apr`, `RAC` y `Dec` son `Form_subfrmDatosCDCASUB_Generales|Propuesta|Impacto|AprobacionSuministrador|DictamenRAC|DecisionFinal.cls`; `P`, `forms/Form_frmDatosCDCASUB.cls`; `SSrv`, `classes/DatosCDCASUBServicio.cls`; `SRepo`, `modules/DatosCDCASUBRepositorio.bas`; `WF`, `classes/WorkflowServicio.cls`; `FPA`, `modules/FormulariosPadreAuxiliares.bas`; `FU`, `modules/FUNCIONES UTILES.bas`; `VM`, `classes/DatosCDCASUBViewModel.cls`.

## Contrato común con el contenedor

| Paso | Qué ocurre | Evidencia |
|---|---|---|
| Pestaña pulsada | El contenedor aplica permisos y luego llama a `PoblarControles(vm)` del subformulario activo | `P:172-175, 182-187` |
| Permisos | `HabilitarGuardarEnSubform` habilita `cmdGuardar`, `cmdCargarDefaults` y `cmdCargarDatosPredeterminados`; `HabilitarControlesSubformulario` bloquea textos (`Locked`) y casillas y deshabilita grupos de opción, salvo los de `Tag = "INMUTABLE"` | `P:93-104`; `FPA:268-273`; `FU:946-980, 931-945` |
| Permiso por bloque | `PuedeEditarBloque` ⇒ `PermiteEdicion_CDCASUB` según rol y estado | `P:684-689`; `WF:1419-1437, 1557-1598` |
| Grabar | El subformulario llama a `Me.Parent.GuardarDesdeSubform Me.Name`; el contenedor identifica el bloque por el nombre, invoca `RecogerValores`, copia **solo los campos del bloque** a la entidad y guarda | `P:268-269, 283-284, 544-615`; `FPA:192-204` |
| Sin cambios | Aprobación usa `hayCambios()` del subformulario; los demás comparan entidad leída y entidad editada con `SonEntidadesIguales`; si no hay cambios: «No se han detectado cambios.» | `P:286-305`; `FPA:25-47` |
| Tras grabar | `RaiseEvent SolicitudHaCambiado`, recarga del ViewModel (instancia nueva) y repoblado del subformulario | `P:459-466, 628-691` |
| Errores | Errores del subformulario: `ManejadorErroresFormulario` (`FU:588`); errores de guardado: `ErrorPresenter.MostrarError` del contenedor | `P:478-481` |

Permisos por bloque (`WF:1557-1598`): con la solicitud Rechazada no se edita nada (`:1564`); en Aprobada solo Decisión final, para Calidad o Administrador (`:1565-1567`); en Formalización solo Aprobación y Decisión final, para los mismos roles (`:1569-1571`). En el resto de estados: Generales para Calidad o Administrador (`:1575-1576`); Propuesta e Impacto para el Técnico en Desarrollo técnico y para Calidad o Administrador en Modificación o posterior (`:1578-1587`); Aprobación y RAC para Calidad o Administrador a partir de Modificación (`:1589-1593`).

## Controles y eventos

| Subformulario | Control | Tipo | Evento y handler | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| Generales | `refSuministrador` | texto, `Tag INMUTABLE` | — | Se vuelca y recoge; siempre deshabilitado | `Gen:56-57, 95, 28-33`; `.form.txt:630-633` |
| Generales | `refSubSuministrador` | texto, `Tag INMUTABLE` | — | Solo lectura; no se recoge | `Gen:58-59, 96`; `.form.txt:758-761` |
| Generales | `suministradorPrincipalNombreDir`, `subSuministradorNombreDir`, `refDesviacionesPrevias` | texto | — | Editables si `PuedeGuardar_Generales` | `Gen:61-66, 98-100` |
| Generales | `grpModContrato` | grupo Sí=1 / No=2 | — | Triestado: 1 ⇒ verdadero, 2 ⇒ falso, otro ⇒ nulo | `Gen:68-74, 102-106`; `.form.txt:1117-1150` |
| Generales | `cmdGuardar` | botón | `Click` `Gen:113-129` | Confirma edición (`CommitEditsOnForm`, `FU:847-865`), guarda y oculta «Restaurar» | `Gen:115-122` |
| Generales | `cmdCargarDefaults` | botón | `Click` `Gen:131-176` | Guarda instantánea y aplica valores calculados (ver Flujos) | `Gen:143-164` |
| Generales | `cmdRestaurarDatos` | botón, oculto | `Click` `Gen:178-205` | Devuelve los cuatro valores de la instantánea | `Gen:182-198` |
| Propuesta | `identificacionMaterial`, `numPlanoEspecificacion`, `cantidadPeriodo`, `numSerieLote`, `causaNC`, `descripcionImpactoNC` | texto | — | Vuelco y recogida directos | `Pro:6-23` |
| Propuesta | `cmdGuardar` | botón | `Click` `Pro:25-27` | Delega en el contenedor; sin manejador de errores | `Pro:26` |
| Impacto | `afectaPrestaciones` … `afectaOtros` (9) | casillas | — | Booleanos | `Imp:10-18, 50-58` |
| Impacto | `grpIncidenciaCoste` | grupo 1-3 | — | «AUMENTARÁ», «DISMINUIRÁ», «NO VARIARÁ»; otro ⇒ «» | `Imp:21-26, 60-65` |
| Impacto | `grpNCConsiderada` | grupo 1-2 | — | «MAYOR», «MENOR»; otro ⇒ «» | `Imp:28-32, 67-71` |
| Impacto | `grpAutoridadDisenio` | grupo 1-2 | — | `esSubSuministradorAD` triestado | `Imp:34-38, 73-77`; `.form.txt:1159` |
| Impacto | `identificacionAutoridadDiseno`, `efectoFechaEntrega` | texto | — | Opcionales | `Imp:40-41, 79-80` |
| Impacto | `cmdGuardar` | botón | `Click` `Imp:87-99` | Recoge y delega | `Imp:90-92` |
| Aprobación | `firmaAprobacionRespIngenieriaNombre`, `firmaAprobacionRespCalidadNombre` | texto | — | Obligatorias | `Apr:31-32`; `SSrv:319-320` |
| Aprobación | `firmaAprobacionRespProduccionNombre`, `…RespDisenioNombre`, `…RepresentanteSumNombre` | texto | — | Habilitadas solo si su grupo «Aplica» vale Sí; con No se vacían | `Apr:35-37, 141-163` |
| Aprobación | `grpRespProduccion`, `grpRespDisenio`, `grpRepresentanteSum` | grupos Sí=1 / No=2 | `AfterUpdate` `Apr:137-139` | Ajustan el campo de firma correspondiente | `Apr:40-47, 75-77` |
| Aprobación | `cmdGuardar` | botón | `Click` `Apr:85-94` | Delega; oculta «Restaurar» | `Apr:87-90` |
| Aprobación | `cmdCargarDatosPredeterminados` | botón | `Click` `Apr:96-122` | Propone las dos firmas obligatorias | `Apr:104-114` |
| Aprobación | `cmdRestaurar` | botón, oculto | `Click` `Apr:124-135` | Devuelve las dos firmas obligatorias | `Apr:126-131` |
| RAC | `racCodigo`, `racNombre`, `racNombreDelegador` | texto | — | Editables si `PuedeGuardar_DictamenRAC` | `RAC:15-31` |
| RAC | `observacionesRAC`, `observacionesRACDelegador` | texto | — | Siempre deshabilitadas; se recogen igualmente | `RAC:29-30, 49, 53` |
| RAC | `cmdGuardar` | botón | `Click` `RAC:61-70` | Recoge y delega | `RAC:64-65` |
| RAC | `cmdCargarDatosPredeterminados` | botón | `Click` `RAC:74-105` | Propone el RAC asignado al expediente | `RAC:87-93` |
| RAC | `cmdRestaurarDatos` | botón, oculto | `Click` `RAC:107-122` | Devuelve `racNombre` y `racNombreDelegador` | `RAC:110-117` |
| Decisión | `grpDecisionFinal` | grupo 1-2 | — | «APROBADO», «NO APROBADO»; vacío ⇒ 0 | `Dec:12-16, 45-49`; `.form.txt:814-849` |
| Decisión | `observacionesFinales`, `NombreFirmanteFinal` | texto | — | Deshabilitados en estado 8 (Aprobada) | `Dec:18-19, 29-34` |
| Decisión | `cmdGuardar` | botón | `Click` `Dec:59-76` | Valida y cierra o guarda (ver Flujos) | `Dec:61-72` |

Los `OnChange` declarados en los textos de los `.form.txt` (por ejemplo `Gen .form.txt:632`, `RAC .form.txt:667`) no tienen handler en los `.cls`.

## Flujos

### Normal (cualquier bloque salvo Decisión final con «APROBADO»)

1. El contenedor vuelca el ViewModel con `PoblarControles` (`P:187`).
2. El usuario edita y pulsa «Grabar»; el subformulario llama a `GuardarDesdeSubform` (`Gen:118`, `Pro:26`, `Imp:92`, `Apr:87`, `RAC:65`, `Dec:71`).
3. El contenedor recoge, comprueba cambios, valida en el servicio y escribe en `tbDatosCDCASUB` (`P:283-324`).
4. Para Propuesta, Impacto y Generales sin transición registra una traza de edición (`P:333-346`).
5. Generales en Preregistro transita a Registro y pregunta si se asigna al Técnico (`SSrv:153-157`; `P:352-363`); Propuesta e Impacto en Desarrollo técnico pueden ofrecer el envío a Calidad (`P:370-452`).
6. Muestra «Datos guardados correctamente.», recarga y repuebla (`P:455-466`).

### Alternativos

- **Cargar datos predeterminados, Generales.** Guarda instantánea de cuatro campos (`Gen:143-146`); `GetDatosGeneralesIniciales` calcula `refSubSuministrador` = código de la solicitud y los nombres con dirección de suministrador principal y subsuministrador a partir de las relaciones del expediente, con «UNION TEMPORAL DE EMPRESAS: A y B» si hay varios (`SSrv:43-102`); recoge, repuebla y muestra «Restaurar» (`Gen:152-164`).
- **Sin registro previo.** Al recargar sin fila en `tbDatosCDCASUB`, el ViewModel se rellena en memoria con esos mismos valores iniciales, sin persistir (`P:647-651`); el primer guardado de Generales da de alta la fila con id `Max+1` (`SRepo:28-40, 6-16`).
- **Cargar datos predeterminados, Aprobación.** Propone el responsable técnico del expediente con sufijo «(Jefe de Proyecto)» y el jefe de calidad del entorno con «(Responsable de Calidad)» (`SSrv:114-124, 714-726`; `Apr:107-110`).
- **Cargar datos predeterminados, RAC.** Busca el RAC asignado al expediente (`RAC:87`; `classes/ExpedienteServicio.cls:293-296`; `modules/ExpedienteRepositorio.bas:86-96`); si no hay, avisa «Este expediente no tiene un RAC asignado. Hable con el gestor de expedientes» (`RAC:92`).
- **Decisión final «APROBADO».** Llama a `ProcesarCierreFormalizacion` (`Dec:68-69`): pide un PDF firmado (`P:491-508`), guarda la decisión, sustituye el adjunto de cierre y transita a Aprobada en una transacción (`P:514-521`; `WF:3346-3385`).
- **Decisión final «NO APROBADO».** Solo guarda los tres campos; no hay transición asociada en el contenedor (`Dec:71`; `P:323`).
- **Aprobación en Modificación.** Tras guardar recarga la gestión de solicitud si está abierta (`P:327-331`).

### Errores

| Situación | Mensaje o efecto | Evidencia |
|---|---|---|
| Decisión final vacía | «El campo 'Decisión Final' es obligatorio.»; no llama al contenedor | `Dec:63-66` |
| Generales incompleto | «Ref. Sub-Suministrador obligatoria.», «Suministrador Principal obligatorio.», «Sub-Suministrador obligatorio.», «Indique si requiere modificación.»; sin escritura | `SSrv:142-145` |
| Propuesta incompleta | «Identificacion obligatoria.», «Causa NC obligatoria.», «Descripcion Impacto obligatoria.» | `SSrv:180-182` |
| Impacto incompleto | «Complete campos obligatorios de Impacto.» | `SSrv:248-250` |
| Aprobación incompleta | «Firma Ingeniería obligatoria.», «Firma Calidad obligatoria.» | `SSrv:319-320` |
| RAC sin código | «Código RAC obligatorio.» salvo `racDecision = "RECHAZADO"` | `SSrv:334-336` |
| Bloque sin fila base | «No se encontro el registro base CDCASUB.» | `SRepo:143, 223, 299` |
| Cierre cancelado | «Operacion cancelada. No se ha seleccionado el documento firmado.» | `P:506` |
| Fallo en transacción | `Rollback` y error propagado | `SSrv:167-171`; `WF:3392-3396` |

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| La referencia del suministrador ya no es obligatoria | Servicio | `SSrv:141` |
| «Requiere modificación» se normaliza a booleano o nulo | Contenedor | `P:558-568` |
| Impacto: coste, clasificación y autoridad de diseño obligatorios; identificación de la autoridad y efecto en la entrega, no | Servicio | `SSrv:248-250` |
| Propuesta e Impacto con rechazo activo registran el antes y el después de cada campo en el JSON de rechazo | Servicio | `SSrv:186-220, 254-299` |
| Propuesta, Impacto y RAC abren transacción propia solo si la solicitud está en Validación | Servicio | `SSrv:194-200, 262-268, 337-343` |
| Aprobación: solo nombres de firma; los indicadores «Aplica» no se copian a la entidad | Contenedor | `P:594-599` |
| RAC con `racDecision = "APROBADO"` se guarda aunque no haya cambios | Contenedor | `P:296-298` |
| Decisión final no editable en estado 8 aunque el permiso lo conceda | Subformulario | `Dec:24-34`; `WF:1565-1567` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura del bloque | `tbDatosCDCASUB` | vía `getPorIdSolicitud` | No | `SSrv:14-16`; `SRepo:62` |
| Generales | `tbDatosCDCASUB` | DAO `AddNew/Edit`, id `Max+1` | Sí, con transición | `SSrv:147-159`; `SRepo:28-40` |
| Propuesta, Impacto, Aprobación, RAC, Decisión | `tbDatosCDCASUB` | DAO dynaset `Edit/Update`, SQL concatenado con `idSolicitud` | Según regla | `SRepo:140-153, 178-200, 220-232, 257-271, 296-306` |
| Valores iniciales de Generales | solicitud, expediente, relaciones y suministradores | servicios | No | `SSrv:43-57` |
| RAC asignado | `TbExpedientesRACS`, `TbRACS` (base de expedientes) | SQL dinámico con `FIRST()` | No | `modules/ExpedienteRepositorio.bas:86-91` |
| Traza de edición | log de cambios | `LogServicio.RegistrarCambio` | No | `P:344` |
| Cierre | `tbDatosCDCASUB`, adjuntos, estado | servicio de workflow | Sí | `WF:3360-3385` |

## Coexistencia

Pendiente de `../coexistencia.md`. Puntos visibles:

- Triestados (`requiereModificacionContrato`, `esSubSuministradorAD`) y nueve booleanos `afecta*`: el destino debe distinguir nulo de falso (`P:558-568`; `Imp:34-38`; `SRepo:184-196`).
- Literales con tilde guardados como dato («AUMENTARÁ», «NO VARIARÁ») y comparados con `UCase` (`Imp:21-24, 61-63`): requieren la misma codificación y comparación en PostgreSQL.
- Dynaset `Edit/Update` sobre tabla vinculada exige clave primaria en el destino; el id por `Max+1` (`SRepo:11`) y las transacciones DAO de espacio de trabajo propio (`SSrv:147-149`) deben preservar su atomicidad.
- `FIRST()` de Jet en la consulta del RAC (`modules/ExpedienteRepositorio.bas:87`) no existe en PostgreSQL si la consulta pasa a ejecutarse en el servidor.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Los indicadores «Aplica» de Aprobación no se leen ni se guardan (`P:594-599, 743-748`; `VM:71-73`): tras cada recarga valen No, el subformulario vacía las tres firmas opcionales al mostrarlas (`Apr:40-47, 144-146`), `hayCambios` detecta ese vaciado (`Apr:176-178`) y el siguiente guardado escribe «» sobre ellas (`SRepo:227-230`).
- `refSuministrador` lleva `Tag INMUTABLE` (`Gen .form.txt:633`) y queda siempre deshabilitado (`Gen:30-33`), aunque el código lo trata como manual (`Gen:56`) y la traza lo usa como campo (`P:340`).
- Las observaciones del RAC y del delegador se deshabilitan siempre (`RAC:29-30`) pero se recogen y guardan (`RAC:49, 53`; `SRepo:264, 268`).
- `racDecision` solo se fija a «RECHAZADO» (`SSrv:657`); no hay control ni código que la ponga a «APROBADO», por lo que la excepción de `P:296-298` no actuaría.
- En Generales, «Cargar» escribe `refSubSuministrador` (`Gen:152`) pero `RecogerValores` no lo recoge (`Gen:96`) y el repoblado posterior (`Gen:159`; `Gen:58`) muestra de nuevo el valor del ViewModel; «Restaurar» no incluye ese campo (`Gen:182-185`).
- En Aprobación, «Cargar» reemplaza la instantánea de las dos firmas obligatorias (`Apr:104-105`); tras «Restaurar», `hayCambios` compara contra esos valores y no contra los leídos (`Apr:174-175`).
- `GuardarAprobacionSuministrador` del subformulario (`Apr:186-197`) y el evento `SolicitudHaCambiado` de Decisión (`Dec:4`) no tienen uso en el código fuente.
- Propuesta no tiene manejo de errores ni comprobación de ViewModel nulo (`Pro:6-27`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: una sola pantalla por bloques compartida con la variante `CD_CA`; persistir o retirar los indicadores «Aplica».

## Criterios de aceptación

Propuestos, no aprobados:

1. Cada bloque solo es editable con el rol y el estado de `PermiteEdicion_CDCASUB` (`WF:1557-1598`).
2. Grabar sin cambios muestra «No se han detectado cambios.» y no escribe (`P:286-305`).
3. Cada bloque rechaza los campos obligatorios de su fila en la tabla de errores con su mensaje y sin escribir.
4. «Cargar datos predeterminados» propone los valores de Generales, Aprobación y RAC descritos, y «Restaurar» devuelve los anteriores.
5. Una firma opcional guardada se muestra de nuevo al volver a la pestaña.
6. «APROBADO» exige un PDF y, en una sola transacción, guarda la decisión, archiva el adjunto y pasa a Aprobada (`WF:3346-3385`).
7. Los grupos triestado conservan nulo, verdadero y falso tras guardar y recargar.

## Pendientes

- Si las observaciones del RAC se rellenan desde otro formulario o proceso (`RAC:29-30`).
- Deriva del código: el `CodeBehind` de `subfrmDatosCDCASUB_DictamenRAC.form.txt` incluye una comprobación de cambios con `racDecision` (`forms/subfrmDatosCDCASUB_DictamenRAC.form.txt:1271-1313`) que el `.cls` no tiene.
- Valor de `m_ObjEntorno.DirUTE` y `NombreJefeCalidad` en los valores iniciales (`SSrv:83, 123`).
- Contenido de `EjecutarTransicion` y del adjunto (`WF:3380-3383`).
