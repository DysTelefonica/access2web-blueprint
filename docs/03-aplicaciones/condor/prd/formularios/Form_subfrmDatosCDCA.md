# Form_subfrmDatosCDCA_* (grupo de seis subformularios)

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_subfrmDatosCDCA_<Bloque>.cls` y `.form.txt` · **PRD:** [README](../README.md)

Seis subformularios que el contenedor `Form_frmDatosCDCA` aloja en su `NavigationSubform`, uno por bloque de la solicitud de concesión o desviación (`CD_CA`). Todos son **no enlazados** (ningún `.form.txt` declara `RecordSource` ni `ControlSource`) y siguen el mismo contrato: `PoblarControles(vm)` vuelca el ViewModel en los controles, `RecogerValores(vm)` hace lo inverso y `cmdGuardar` delega el guardado en el padre (`Me.Parent.GuardarDesdeSubform`).

Abreviaturas: `AS`, `DF`, `DR`, `GE`, `IM` y `PR` = `forms/Form_subfrmDatosCDCA_` + `AprobacionSuministrador`, `DecisionFinal`, `DictamenRAC`, `Generales`, `Impacto` y `Propuesta` (`.cls` salvo que se indique `.form.txt`); `P`, `forms/Form_frmDatosCDCA.cls`; `Srv`, `classes/DatosCDCAServicio.cls`; `Repo`, `modules/DatosCDCARepositorio.bas`; `WF`, `classes/WorkflowServicio.cls`; `FPA`, `modules/FormulariosPadreAuxiliares.bas`; `FU`, `modules/FUNCIONES UTILES.bas`; `ExpRepo`, `modules/ExpedienteRepositorio.bas`; `Enum`, `modules/modEnumeradores.bas`; `VM`, `classes/DatosCDCAViewModel.cls`.

## Entrada, salida y permisos

| Subformulario | Pestaña del padre | Bloque | Quién edita (`PermiteEdicion_CDCA`) | Evidencia |
|---|---|---|---|---|
| `_Generales` | «GENERAL» | `Bloque_Generales` | Calidad o Administrador en estados anteriores a Formalización | `WF:1515-1516` |
| `_Propuesta` | «DETALLE Y DESCRIPCIÓN» | `Bloque_Propuesta` | Técnico en Desarrollo técnico (3); Calidad o Administrador en Modificación (4) y posteriores | `WF:1519-1532` |
| `_Impacto` | «MOTIVOS» | `Bloque_Impacto` | Igual que Propuesta | `WF:1519-1532` |
| `_AprobacionSuministrador` | «APROBACIÓN SUMINSTRADOR» | `Bloque_AprobacionSuministrador` | Calidad o Administrador con estado > 3, también en Formalización (7) | `WF:1503, 1535-1538` |
| `_DictamenRAC` | «RAC» | `Bloque_DictamenRAC` | Calidad o Administrador con estado > 3 y < 7 | `WF:1499-1508, 1541-1544` |
| `_DecisionFinal` | «AUTORIDAD COMPETENTE» | `Bloque_DecisionFinal` | Calidad o Administrador solo en Formalización (7) y Aprobada (8) | `WF:1491-1508, 1546` |

- **Rótulos de pestaña:** `forms/Form_frmDatosCDCA.form.txt:707, 762, 818, 874, 931, 988`. Estados: `Enum:37-45`.
- **Carga:** el padre navega con `DoCmd.BrowseTo` al nombre que devuelve `GetSubformNameFromTabName` (`P:131, 143`; `FPA:103-145`) y llama por `CallByName` a `PoblarControles` (`P:107`).
- **Permisos:** el padre calcula `PuedeGuardar_<Bloque>` con `WF.PuedeEditarBloque` (`P:592-597`; `WF:1419-1431`). En Rechazada (9) nadie edita (`WF:1488`). Después `ConfigurarYPoblarInterfaz` habilita `cmdGuardar`, `cmdCargarDefaults` y `cmdCargarDatosPredeterminados` (`FPA:268-273`) y bloquea o desbloquea los controles que no llevan `Tag` `INMUTABLE` (`P:640-641`; `FU:946-980`). `PoblarControles` se ejecuta **después** y algunos subformularios vuelven a fijar `Enabled` (`P:155-156`).
- **Abre:** ningún formulario. Decisión final «APROBADO» dispara en el padre el selector de PDF del cierre (`DF:85`; `P:669-696`).

## Controles y eventos

| Subformulario | Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|---|
| Generales | `numContrato`, `refSuministrador` | texto, `Tag` `INMUTABLE` | — | Siempre deshabilitados; solo se rellenan con «Cargar datos predeterminados» | `GE.form.txt:632-635, 759-762`; `GE:27-32` |
| Generales | `SuministradorNombreDir`, `refDesviacionesPrevias` | texto | — | Editables si `PuedeGuardar_Generales` | `GE:58-61` |
| Generales | `grpModContrato` | grupo Sí (1) / No (2) | — | Tri-estado: `Null` si nunca se informó | `GE:64-70, 106-110`; `GE.form.txt:1148-1180` |
| Generales | `cmdCargarDefaults` | botón | `Click` | Copia expediente, código de solicitud y suministrador | `GE:137-180` |
| Generales | `cmdRestaurarDatos` | botón | `Click` | Restaura los cuatro textos previos a la carga | `GE:182-209` |
| Generales | `cmdGuardar` | botón | `Click` | `CommitEditsOnForm` y guardado en el padre | `GE:121-135`; `FU:847-862` |
| Propuesta | `identificacionMaterial`, `numPlanoEspecificacion`, `cantidadPeriodo`, `numSerieLote`, `causaNC`, `descripcionImpactoNC` | texto (6) | — | Volcado directo del ViewModel | `PR:16-21, 34-39` |
| Propuesta | `cmdGuardar` | botón | `Click` | Guardado en el padre, sin `RecogerValores` propio | `PR:48-56` |
| Impacto | `afectaPrestaciones` … `afectaOtros` | casillas (9) | — | Características afectadas | `IM:18-26, 76-84` |
| Impacto | `grpIncidenciaCoste` | grupo 1/2/3 | — | «AUMENTARÁ», «DISMINUIRÁ», «NO VARIARÁ» | `IM:35-40, 87-92`; `IM.form.txt:1174-1224` |
| Impacto | `grpNCConsiderada` | grupo 1/2 | — | «MAYOR», «MENOR» | `IM:42-46, 94-98` |
| Impacto | `grpAutoridadDisenio` | grupo Sí/No | — | Tri-estado `esSuministradorAD` | `IM:48-52, 100-104`; `VM:73` |
| Impacto | `identificacionAutoridadDiseno`, `efectoFechaEntrega` | texto | — | Textos libres | `IM:61-62, 107-108` |
| Impacto | `cmdGuardar` | botón | `Click` | `RecogerValores` y guardado en el padre | `IM:117-129` |
| Aprobación | `firmaAprobacionRespIngenieriaNombre`, `firmaAprobacionRespCalidadNombre` | texto | — | Firmas obligatorias | `AS:25-26`; `AS.form.txt:652, 1039` |
| Aprobación | `grpRespProduccion`, `grpRespDisenio`, `grpRepresentanteSum` | grupos Sí (1) / No (2) | `AfterUpdate` | «Aplica»: habilitan el nombre asociado; con «No» lo vacían | `AS:167-199`; `AS.form.txt:1419-1420, 1565-1566, 1710-1711` |
| Aprobación | `firmaAprobacionRespProduccionNombre`, `…RespDisenioNombre`, `…RepresentanteSumNombre` | texto | — | Firmas opcionales | `AS:29-31` |
| Aprobación | `cmdCargarDatosPredeterminados` | botón | `Click` | Rellena ingeniería y calidad | `AS:107-143` |
| Aprobación | `cmdRestaurar` | botón | `Click` | Restaura ingeniería y calidad | `AS:145-162` |
| Aprobación | `cmdGuardar` | botón | `Click` | Guardado en el padre | `AS:93-105` |
| RAC | `racCodigo`, `racNombre` | texto | — | Editables si `PuedeGuardar_DictamenRAC` | `DR:14-23` |
| RAC | `observacionesRAC` | texto | — | Siempre deshabilitado | `DR:24` |
| RAC | `cmdCargarDatosPredeterminados` | botón | `Click` | Rellena `racNombre` con el RAC del expediente | `DR:64-94` |
| RAC | `cmdRestaurarDatos` | botón | `Click` | Restaura `racNombre` | `DR:96-114` |
| RAC | `cmdGuardar` | botón | `Click` | `RecogerValores` y guardado en el padre | `DR:51-60` |
| Decisión | `grpDecisionFinal` | grupo 1/2 | — | «APROBADO» / «NO APROBADO»; 0 si vacío | `DF:17-21, 52-56`; `DF.form.txt:815-850` |
| Decisión | `observacionesFinales`, `NombreFirmanteFinal` | texto | — | Textos libres | `DF:23-24` |
| Decisión | `cmdGuardar` | botón | `Click` | Valida y bifurca: cierre o guardado | `DF:68-96` |

## Flujos

### Normal (común a los seis)

1. Al entrar en la pestaña, el padre configura permisos y llama a `PoblarControles` (`P:155-156`). Cada subformulario guarda la referencia al ViewModel (`GE:47`, `PR:13`, `IM:14`, `AS:21`, `DR:10`, `DF:13`).
2. `cmdGuardar` → `GuardarDesdeSubform(Me.Name)` (`P:160`). El padre lee `tbDatosCDCA`, clona la entidad, vuelve a llamar a `RecogerValores` y copia **solo los campos del bloque** (`P:192-196, 492-547`).
3. Si la entidad no cambia, avisa «No se han detectado cambios.» y sale (`P:199-210`).
4. Guarda según el bloque (`P:215-248`), muestra «Datos guardados correctamente.» si no hubo avance de estado (`P:375-377`), lanza `SolicitudHaCambiado`, recarga el ViewModel y repuebla el subformulario (`P:380-383`).

### Por subformulario

- **Generales:** la validación exige contrato, referencia, suministrador y Sí/No de modificación del contrato (`Srv:300-303`). Guarda la entidad completa en transacción y, si la solicitud estaba en Prerregistro (1), pasa a Registro (2) en la misma transacción (`Srv:305-323`). El padre pregunta «¿Asignar tarea al Técnico ahora?» y, con Sí, transiciona a Desarrollo técnico (`P:258-273`).
- **Cargar datos predeterminados (Generales):** `numContrato` ← código largo del expediente; `refSuministrador` ← código de la solicitud; `SuministradorNombreDir` ← subcontratistas o, si no hay, contratistas principales, en mayúsculas; con más de uno, «UNION TEMPORAL DE EMPRESAS: A y B»; añade la dirección entre paréntesis (`Srv:134-188`). Pasa los valores al ViewModel **sin guardar** y repuebla (`GE:160-163`).
- **Propuesta e Impacto:** validan los campos obligatorios (`Srv:359-361, 436-438`). Si hay un rechazo activo, registran el valor anterior y el nuevo de cada campo en el JSON de rechazo (`Srv:363-400, 440-448`). Solo abren transacción propia si la solicitud está en Validación (5) (`Srv:373-382, 451-460, 863-884`). En Desarrollo técnico el padre guía la navegación entre «MOTIVOS» y «DETALLE Y DESCRIPCIÓN» y ofrece enviar a Calidad (Modificación, 4) cuando la parte técnica está completa (`P:280-370`).
- **Impacto, sección sin guardar:** si `impactoCoste` está vacío, los tres grupos se muestran a `Null` aunque haya valores en los otros dos (`IM:31-58`).
- **Aprobación del suministrador:** valida ingeniería y calidad; cada firma opcional es obligatoria si «Aplica» = Sí y se vacía si = No (`Srv:511-532`). El padre aplica además `hayCambios` (`P:241-245`; `AS:201-214`). En Modificación (4), recarga `frmGestionSolicitud` (`P:251-255`). «Aplica» no se persiste: al cargar se deduce de que el nombre tenga texto (`P:464-488`).
- **Cargar datos predeterminados (Aprobación):** ingeniería ← responsable técnico del expediente + « (Jefe de Proyecto)»; calidad ← jefe de calidad del entorno + « (Responsable de Calidad)» (`Srv:263-275, 893-905`).
- **Dictamen RAC:** código y nombre obligatorios salvo `racDecision = "RECHAZADO"` (`Srv:555-558`). El padre abre una transacción DAO y guarda dentro de ella (`P:215-230`). Si no hay cambios pero `racDecision = "APROBADO"`, continúa igualmente (`P:201-204`). `racDecision` no tiene control: llega del ViewModel (`P:456, 539`).
- **Cargar datos predeterminados (RAC):** primer RAC del expediente (`ExpRepo:86-96`). Si no hay ninguno, avisa «Este expediente no tiene un RAC asignado…» (`DR:78-82`).
- **Decisión final «APROBADO»:** `ProcesarCierreFormalizacion` pide un PDF de cierre (`P:677-696`). Después, en una sola transacción, guarda la decisión, sustituye el adjunto de cierre y transiciona a Aprobada (8) (`P:703-711`; `WF:3347-3385`).
- **Decisión final «NO APROBADO»:** guardado estándar; solo escribe los tres campos (`DF:87-91`; `P:247`; `Srv:591-601`).

### Alternativos

- **Restaurar:** «Restaurar» devuelve los valores capturados antes de la carga de datos predeterminados y vuelve a habilitar la carga (`GE:186-202`; `AS:149-155`; `DR:99-106`).
- **Decisión final en Aprobada (8):** el subformulario deshabilita grupo, textos y `cmdGuardar` (`DF:26-40`).

### Errores

- Los errores del servicio (513) llegan al manejador del padre, que deshace la transacción del RAC si estaba abierta y muestra el mensaje con `ManejadorErroresFormulario` (`P:397-405`; `FU:588-620`).
- Decisión final vacía: «El campo 'Decisión Final' es obligatorio.» (`DF:75-77`); el servicio repite la regla (`Srv:594-596`).
- Cancelar el aviso o el selector de PDF: «Operacion cancelada…», sin escritura (`P:678-695`).
- Sin registro base en `tbDatosCDCA`: «No se encontró el registro base CDCA.» (`Repo:99, 133, 175, 214, 245`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Campos con `Tag` `INMUTABLE`: siempre deshabilitados y en gris | Generales y helper del padre | `GE:27-32`; `FU:954-956` |
| Firmas opcionales obligatorias si «Aplica» = Sí; vaciadas si = No | Servicio y subformulario | `Srv:516-532`; `AS:178-180` |
| Coste, clasificación NC y autoridad de diseño obligatorios | Servicio | `Srv:436-438` |
| Material, causa y descripción de la NC obligatorios | Servicio | `Srv:359-361` |
| Decisión final bloqueada con estado literal `8` (Aprobada) | Subformulario | `DF:27`; `Enum:44` |
| Mapeo de opciones a texto en mayúsculas acentuadas | Subformularios | `IM:87-98`; `DF:52-56` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Datos del bloque (lectura) | `tbDatosCDCA` | vía `Srv.GetDatosCDCA` | No | `P:192`; `Srv:95-98` |
| Generales | `tbDatosCDCA` (entidad completa) y transición | DAO, `Repo.Guardar` | Sí | `Srv:305-323` |
| Propuesta, Impacto | `tbDatosCDCA` | DAO `SELECT *` + `Edit/Update`, SQL concatenado | Solo en Validación | `Repo:96-108, 130-155` |
| Aprobación | `tbDatosCDCA` | DAO `Edit/Update` con `getdb()` | No | `Repo:167-187` |
| Dictamen RAC | `tbDatosCDCA` (`racCodigo`, `observacionesRAC`, `racNombre`, `racDecision`) | DAO `Edit/Update` | Sí, abierta por el padre | `P:216-225`; `Repo:211-222` |
| Decisión final | `tbDatosCDCA` (`decisionFinal`, `observacionesFinales`, `NombreFirmanteFinal`) | DAO `Edit/Update` | Sí en el cierre | `Repo:242-251`; `WF:3352-3385` |
| RAC por defecto | `TbExpedientesRACS` ⋈ `TbRACS` (base de expedientes) | SQL dinámico con `FIRST()` | No | `ExpRepo:86-91` |
| Estado en Validación | `tbSolicitudes` | SQL dinámico | No | `Srv:876-884` |

## Coexistencia

Detalle pendiente de `../coexistencia.md`. En este grupo:

- Las transacciones abren la base con `ws.OpenDatabase(getdb().Name, …, "MS Access;PWD=…")` (`P:217`; `Srv:307`). Con tablas de PostgreSQL vinculadas, debe comprobarse que la transacción del espacio de trabajo DAO cubre esas tablas.
- `FIRST()` es un agregado de Jet sin equivalente directo en PostgreSQL (`ExpRepo:87`).
- La actualización edita la **primera fila** por `idSolicitud` ordenada por `idDatosCDCA` (`Repo:96, 130, 172, 211, 242`). Debe preservarse la unicidad efectiva.
- Hay valores tri-estado (`requiereModificacionContrato`, `esSuministradorAD`) que exigen columnas booleanas con nulo (`VM:49, 73`; `Repo:151`) y literales de texto acentuados como dominio («AUMENTARÁ»…, `IM:88-90`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- En Aprobada (8) `PermiteEdicion_CDCA` permite editar la decisión «para borrarla/reabrir» (`WF:1490-1494`), pero el subformulario lo deshabilita todo (`DF:27-40`).
- «NO APROBADO» no cambia el estado; `RegistrarRechazoDesdeFormalizacion` (`Srv:812`) no tiene llamadas en `src`.
- `cmdCargarDefaults` y el botón de carga del RAC se deshabilitan mientras tienen el foco, sin `QuitarFocoControl` (`GE:82-84, 163, 166`; `DR:84`). Aprobación sí lo quita (`AS:129-130`).
- `observacionesRAC` nunca es editable, pero se guarda desde el control (`DR:24, 43`; `Repo:218`).
- La excepción «RAC aprobado sin cambios sigue adelante» no lleva a ningún avance: la rama RAC solo guarda (`P:201-204, 215-230`).
- La carga de datos predeterminados del RAC solo rellena el nombre, no el código (`DR:79`).
- `GuardarAprobacionSuministrador` del subformulario no tiene llamadas (`AS:216-223`). En Propuesta, `RecogerValores` lo invoca solo el padre (`PR:48-50`; `P:195`).
- El comentario de cabecera dice «no inferir por contenido de texto» (`AS:7`), pero el padre infiere «Aplica» del texto (`P:464-488`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: persistir «Aplica» como dato; unificar las transacciones de guardado por bloque; enumerar las opciones como códigos en lugar de textos acentuados.

## Criterios de aceptación

Propuestos, no aprobados:

1. Cada bloque es editable solo para el rol y el estado de la tabla de permisos; en Rechazada nada es editable.
2. Generales rechaza el guardado sin contrato, referencia, suministrador o Sí/No de modificación; en Prerregistro el guardado pasa la solicitud a Registro.
3. «Cargar datos predeterminados» rellena los campos según las reglas descritas y «Restaurar» devuelve los valores previos sin guardar.
4. Impacto exige coste, clasificación NC y autoridad de diseño; una sección sin guardar muestra los tres grupos sin selección.
5. Con «Aplica» = Sí la firma opcional es obligatoria; con No se vacía.
6. El RAC exige código y nombre salvo dictamen rechazado.
7. Decisión final «APROBADO» exige un PDF y, en una sola operación, guarda, archiva el adjunto y deja la solicitud en Aprobada; si algo falla, no queda nada guardado.
8. Con un rechazo activo, los cambios en Propuesta e Impacto quedan trazados (anterior y nuevo).

## Pendientes

- Qué debe ocurrir con «NO APROBADO» (rechazo, reapertura o solo registro).
- Contenido de `Repo.Guardar`, `EjecutarTransicion` y `JsonHelper` (fuera de este grupo).
- Si `tbDatosCDCA` admite más de una fila por solicitud.
- Origen del jefe de calidad y de la dirección UTE del entorno (`Srv:171, 274`).
