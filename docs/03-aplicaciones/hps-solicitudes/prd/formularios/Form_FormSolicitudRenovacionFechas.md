# Form_FormSolicitudRenovacionFechas

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormSolicitudRenovacionFechas.cls` (1652 líneas) y `.form.txt` (81 bloques `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_FormSolicitudRenovacionFechas.cls:línea` salvo indicación.

Cronograma de una solicitud de renovación o elevación de grado: es el segundo destino del control de navegación de `Form_FormSolicitudRenovacion` (`Form_FormSolicitudRenovacion.form.txt:497`). Es un formulario **no enlazado** que pinta 40 campos de fecha y decide, para cada hito, si se puede registrar, se puede eliminar o se puede abrir el correo o el documento associated. La persistencia la hacen `SolicitudOperaciones` (fases y estados) y `SolicitudFechasOperaciones` (fechas); este `.cls` sólo decide, calcula y coordina.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudRenovacion` | Es el destino «Fechas» del control de navegación; se abre por diseño al cambiar de pestaña | `Form_FormSolicitudRenovacion.form.txt:497`, `Form_FormSolicitudRenovacion.cls:110, 121, 135` |
| `Form_FormEnvioCorreo` | Tras enviar un correo, el padre invoca `EstablecerDatos` por código | `Form_FormEnvioCorreo.cls:105` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `Form_FormAdjuntarExcelSolicitante` | `DoCmd.OpenForm "FormAdjuntarExcelSolicitante"`, con `WithEvents` | `:486-488` |
| `Form_FormEnvioCorreo` | Siete aperturas, cada una con un `EnumTipoEnvioCorreo` distinto | `:862`, `:893`, `:924`, `:955`, `:986`, `:1017`, `:1048` |
| `FormWeb` (vía `ObtenerHTMLEnLocal`) | Siete botones de «ver documento» | `:1219-1420` |
| `Form_FormSolicitudesGestion` | `Form_FormSolicitudesGestion.Filtrar` cuando el estado de la solicitud ha cambiado | `:1124-1128` |

No declara parámetros: trabaja sobre `m_ObjSolicitudActiva` (`VG.bas:197`). La carga de datos la hace `EstablecerDatos` (`:1480-1652`), que es `Public` para que el contenedor y `Form_FormEnvioCorreo` puedan repintarla.

## Controles y eventos

El formulario tiene 40 procedimientos y unos 80 controles; la mayor parte son parejas «fecha + botón» que forman el cronograma. Se agrupan aquí por familia.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` → `Form_Load` | Carga el cronograma y activa la edición | `:1422-1450`, `.form.txt:38, 373` |
| `FechaEnvioExcel` | cuadro de fecha | — | Envío del Excel de solicitud de información | `:915`, `.form.txt:915` |
| `ComandoRegistrarExcelEnviado` | botón | `OnClick` | Abre `FormEnvioCorreo` con `EnvioExcel` | `:910-939`, `.form.txt:944-945` |
| `FechaRecepcionExcel` | cuadro de fecha | — | Recepción del Excel del solicitante | `.form.txt:984` |
| `ComandoRegistrarExcelRecibido` | botón | `OnClick` | Abre `FormAdjuntarExcelSolicitante` | `:474-503`, `.form.txt:1013-1014` |
| `FechaCorreoRecordatorioExcel1` / `2` | cuadros de fecha | — | Envío de los recordatorios del Excel | `.form.txt:1053, 1154` |
| `ComandoEnvioRecordatorioExcel1` / `2` | botones | `OnClick` | Abren `FormEnvioCorreo` con `RecordatorioExcel1` / `2` | `:941-1001`, `.form.txt:1114, 1215` |
| `FechaTramitacionAltaMarga` | cuadro de fecha | — | Alta de la ficha en Marga | `.form.txt:1255` |
| `FechaCorreoRecordatorioRellenoMarga1` / `2` | cuadros de fecha | — | Recordatorios de Marga | `.form.txt:1287, 1388` |
| `ComandoEnvioRecordatorioMarga1` / `2` | botones | `OnClick` | Abren `FormEnvioCorreo` con `RecordatorioMarga1` / `2` | `:1003-1065`, `.form.txt:1348, 1449` |
| `FechaDesestimado` | cuadro de fecha | — | Fecha de desestimación | `.form.txt:1489` |
| `ComandoRegistrarDesestimado` | botón | `OnClick` | Pide la fecha por `InputBox` y registra el desestimado | `:332-371`, `.form.txt:2126-2127` |
| `FechaCancelado` | cuadro de fecha | — | Fecha de cancelación | `.form.txt:1613` |
| `ComandoRegistraCancelacion` | botón | `OnClick` | Pide la fecha por `InputBox` y registra la cancelación | `:291-330`, `.form.txt:3453-3454` |
| `DiasParaCancelacion` | cuadro numérico | — | Días que faltan para la autocancelación | `:1618-1624`, `.form.txt:1553` |
| `FechaPrevista*` (6 controles) | cuadros de fecha | — | Próximos correos previstos | `.form.txt:1707, 1771, 1833, 1895, 1995, 2059` |
| `ComandoEnvioCancelacionMarga` / `PreMarga` | botones | `OnClick` | Abren `FormEnvioCorreo` con `CancelacionMarga` / `CancelacionPreMarga` | `:848-908`, `.form.txt:1925, 2089` |
| `ComandoRegistrarAltaMarga` | botón | `OnClick` | Pide la fecha y registra el alta en Marga | `:505-547`, `.form.txt:2163-2164` |
| `FechaEnvioDPS` | cuadro de fecha | — | Envío a la DGSFP | `.form.txt:3140` |
| `ComandoRegistrarDPS` | botón | `OnClick` | Pide la fecha y registra el envío DPS | `:373-416`, `.form.txt:3200-3201` |
| `FechaEnvioONS` | cuadro de fecha | — | Envío a la ONS | `.form.txt:3563` |
| `ComandoRegistrarEnvioONS` | botón | `OnClick` | Pide la fecha y registra el envío a la ONS | `:418-472`, `.form.txt:3688-3689` |
| `FechaRegistroEnHPS` | cuadro de fecha | — | Fecha de registro en el aplicativo de HPS | `.form.txt:3494` |
| `ComandoRegistroHPS` | botón | `OnClick` | Crea o actualiza el `UsuarioHPS` y cierra la tramitación | `:1068-1147`, `.form.txt:3523-3524` |
| `ComandoURLExcelSolicitud` / `ComandoURLExcelRespuesta` | botones | `OnClick` | Abren la plantilla del Excel o el Excel recibido | `:1149-1217`, `.form.txt:2200, 2244` |
| `ComandoVer*` (7 botones) | botones | `OnClick` | Proyectan en `FormWeb` el HTML de recordatorios y cancelaciones | `:1219-1420`, `.form.txt:2286-2568` |
| `ComandoEliminar*` (11 botones) | botones | `OnClick` | Borran la fase o la fecha correspondiente, con confirmación | `:187-846`, `.form.txt:2615-3387` |
| `ComandoActualizarRecordatoriosExcelPreMarga` / `Marga` | botones «Contador» | `OnClick` | Recalculan las fechas previstas de recordatorio y las graba | `:48-186`, `.form.txt:3300, 3343` |
| `lblTitulo`, etiquetas de fase | etiquetas | — | Rótulos del cronograma | `.form.txt:1083-2026` |

## Flujos

### Carga del cronograma (`Form_Load` → `EstablecerDatos`, `:1422-1450`, `:1480-1652`)

1. `Form_Load` desactiva la edición, llama a `EstablecerDatos` y la reactiva (`:1431-1447`).
2. `EstablecerDatos` sale sin pintar si no hay solicitud activa (`:1486-1489`) o si la solicitud no tiene `SolicitudFechas` (`:1490-1493`).
3. Fija el foco en `Me.ComandoRegistrarExcelEnviado` (`:1494`) y, para cada fecha, la pinta o la vacía según sea una fecha válida, y deja visible el botón de eliminar correspondiente sólo cuando la fecha existe (`:1496-1609`).
4. Calcula `DiasParaCancelacion` como la diferencia en días entre hoy y la fecha prevista de autocancelación, con prioridad para la de Marga (`:1632-1638`).

### Registro de hitos

- **Alta en Marga** (`:505-547`), **DPS** (`:373-416`), **ONS** (`:418-472`), **desestimado** (`:332-371`) y **cancelación** (`:291-330`) piden la fecha con un `InputBox` cuyo valor por defecto es la fecha del sistema; si el usuario cancela, el diálogo devuelve cadena vacía y el handler sale sin escribir (`:296-298`, `:338-340`, `:381-383`, `:429-431`).
- El DPS exige que antes se haya registrado el alta en Marga (`:378-382`); el envío a la ONS exige lo mismo (`:426-429`).
- Tras cada registro se invalida la caché de fechas de la solicitud (`Set m_ObjSolicitudActiva.SolicitudFechas = Nothing`), se repinta el cronograma y se llama a `PintarIndicadores` (`:528-535` y equivalentes).

### Recálculo de fechas previstas

- `ComandoActualizarRecordatoriosExcelPreMarga_Click` (`:48-115`) copia a las fechas previstas las fechas calculadas (`…Calculado`) de los dos recordatorios de Excel y de la autocancelación pre-Marga, y las graba con `SolicitudFechasOperaciones.Registrar` (`:88-98`).
- `ComandoActualizarRecordatoriosExcelMarga_Click` (`:116-186`) hace lo mismo con los dos recordatorios de Marga y la autocancelación de Marga (`:152-166`).

### Envío de correos

- Los siete botones de envío abren `FormEnvioCorreo` con un `EnumTipoEnvioCorreo` distinto: `CancelacionMarga` (`:862`), `CancelacionPreMarga` (`:893`), `EnvioExcel` (`:924`), `RecordatorioExcel1` (`:955`), `RecordatorioExcel2` (`:986`), `RecordatorioMarga1` (`:1017`) y `RecordatorioMarga2` (`:1048`). En todos los casos se cierra antes la instancia previa de `FormEnvioCorreo`.

### Borrado de fases

- `Eliminar` (`:10-46`) delega en `SolicitudOperaciones.EliminarFase`, invalida la caché, repinta y llama a `PintarIndicadores` (`:31-41`).
- Los handlers de borrado piden confirmación con un `MsgBox` cuyo botón por defecto es «No» y, si se confirma, llaman a `Eliminar` con la fase correspondiente (`:187-219` y equivalentes).
- Dos borrados no pasan por `Eliminar`: el Excel recibido (`:667-709`) y el alta en Marga (`:549-594`), que llaman directamente a `EliminarExcelDevuelto` y `EliminarMarga`.

### Cierre de la tramitación

- `ComandoRegistroHPS_Click` (`:1068-1147`) exige la fecha de envío a la ONS (`:1082-1085`), pide a la capa de datos el motivo que impide registrar el usuario HPS (`:1088-1095`), llama a `RegistrarUsuarioHPSEnHPS` (`:1100`) y avisa del resultado (`:1104-1112`). Si el estado de la solicitud ha cambiado, refresca `FormSolicitudesGestion` (`:1123-1128`).

### Errores

Todos los handlers usan el patrón `On Error GoTo errores` con el centinela `Err.Number = 1000` para los errores ya descritos en `m_Error`; los errores ajenos rellenan el mensaje, envían `CorreoAlAdministrador` (`FU.bas:585`) y muestran un `MsgBox` crítico. `Form_Load` no cierra el formulario tras el aviso (`:1440-1449`), a diferencia de otros formularios del legacy.

## Reglas, validaciones y cálculos

- **Sin solicitud o sin registro de fechas:** el cronograma se queda vacío, sin aviso (`:1487-1497`).
- **Orden de fases:** el DPS exige el alta en Marga (`:378-382`) y el envío a la ONS también (`:426-429`); en la capa de datos, el envío a la ONS exige el DPS previo y no puede ser anterior a él (`SolicitudOperaciones.cls:2045-2051`).
- **Fechas no futuras:** `SolicitudOperaciones.FechaValidaNoFutura` (`SolicitudOperaciones.cls:15`) se aplica en las fases de correo; en los `InputBox` de este formulario no hay validación de la fecha tecleada: se pasa como cadena a la capa de datos.
- **Desestimado y cancelación:** se registran con la fecha que el usuario teclee; ambas anulan los automatismos programados (`SolicitudOperaciones.cls:2158-2161`, `:2251-2254`).
- **Días para la autocancelación:** `DateDiff("d", Date, FechaPrevista…)` (`:1632-1638`), sin acarreo de signo; un valor negativo indicaría una fecha pasada.
- **Visibilidad de los botones de eliminar:** depende de que la fecha exista (`:1500-1610`), de modo que no se puede borrar lo que no está registrado.
- **Escritura:** todas las fases pasan por `SolicitudOperaciones`, que encadena `UPDATE` sobre `TbSolicitudesFechas` y registra el log (`SolicitudOperaciones.cls:1311-1450`); las fechas previstas se graban con `SolicitudFechasOperaciones.Registrar`, que compara campo a campo con la fotografía inicial (`SolicitudFechasOperaciones.cls:32-190`).
- **Sin transacciones:** ninguna de estas rutas abre transacción; la única del legacy está en `SolicitudOperaciones.ReprogramarAvisosFuturos` (`SolicitudOperaciones.cls:2456`, `:2613`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Fechas de la solicitud | `SolicitudFechas` de la solicitud de sesión | Derivado, sin DAO en este `.cls` | `:1490` |
| Alta, DPS, ONS, desestimado, cancelación | `TbSolicitudesFechas` y estado de `TbSolicitudes` | `UPDATE` con SQL concatenado | `SolicitudOperaciones.cls:1311-1450`, `:2017-2076`, `:2118-2210` |
| Fechas previstas recalculadas | `TbSolicitudesFechas` | DAO (`Recordset` en `SolicitudFechasOperaciones`) | `SolicitudFechasOperaciones.cls:190-414` |
| Excel recibido | `TbSolicitudes` y `TbSolicitudesFechas` | `UPDATE` con SQL concatenado | `SolicitudOperaciones.cls:1544-1618` |
| Alta en Marga | `TbSolicitudesFechas` | `UPDATE` con SQL concatenado | `SolicitudOperaciones.cls:1773-1859` |
| Registro en HPS | usuario de HPS en la base de HPS | DAO remoto | `SolicitudOperaciones.cls:1158-1242` |
| Log de fases | log general | DAO | `FU.bas:895-937` |
| HTML de recordatorios | ficheros locales generados y `FormWeb` | E/S de fichero | `FU.bas:1336-1392` |

## Coexistencia

Pendiente de `../coexistencia.md`. Lo que ya se ve: el cronograma es un mapa de fases con dependencias (alta en Marga → DPS → ONS → registro en HPS) más un conjunto de automatismos de correo programados. Las escrituras son `UPDATE` con SQL concatenado sobre `TbSolicitudesFechas` (`SolicitudOperaciones.cls:1311-1450`) y no hay transacción que agrupe la fecha, el estado y el log. En la web el equivalente es una vista de hitos con su regla de precedencia declarada, y un servicio de dominio que la aplique en el servidor.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-84.** `ComandoActualizarRecordatoriosExcelMarga_Click` decide con una fecha que no es la que copia: la guarda es `IsDate(.FechaPrevistaCorreoRecordatorioExcel2Calculado)` (`:139`) pero el valor que asigna es `.FechaPrevistaCorreoRecordatorioRellenoMarga2Calculado` (`:140`). Si el segundo recordatorio de Excel no tiene fecha prevista calculada, el segundo recordatorio de Marga no se recalcula aunque su propio valor calculado esté disponible.
- **H-HPS-S-85.** `EstablecerDatos` decide si habilita «Registrar envío ONS» y «Registrar en HPS» leyendo `Me.FechaRegistroEnHPS` (`:1586`), que se pinta más abajo en el mismo método (`:1610-1615`): en el primer repintado tras registrar en HPS, el control todavía está vacío, de modo que los dos botones quedan habilitados y es posible registrar dos veces.
- **H-HPS-S-86.** En la rama «no hay fecha de envío ONS», la habilitación de «Registrar envío ONS» depende de `Me.FechaEnvioDPS <> ""` (`:1598`): como `FechaEnvioDPS` se pinta antes y a `Null` cuando no hay fecha de DPS (`:1581`), la comparación vale `Null` y el botón queda deshabilitado (`:1602`) aunque el flujo admitiese el registro.
- **H-HPS-S-87.** Los mensajes de error identifican mal el handler: el de `ComandoRegistrarEnvioONS_Click` dice «Al ComandoRegistrarAltaMarga_Click» (`:466`), el de `ComandoEliminarExcelRecibido_Click` dice «Al ComandoEliminarExcelEnviado_Click» (`:703`) y el de `ComandoEliminarAltaMarga_Click` vuelve a decir «Al ComandoRegistrarAltaMarga_Click» (`:588`).
- **H-HPS-S-88.** La ventana conserva el `Caption` de diseño «EXPEDIENTES Versión 2024-001 Administrador» (`.form.txt:38`): ningún handler del `.cls` lo sobrescribe, de modo que la pestaña de fechas se presenta como si fuera otra aplicación.
- **H-HPS-S-89.** `ComandoRegistroHPS_Click` accede a `m_ObjSolicitudActiva.SolicitudFechas.FechaEnvioONS` (`:1082`) sin comprobar que `SolicitudFechas` exista: con una solicitud sin registro de fechas se produce el error de objeto no establecido, que el manejador convierte en correo al administrador.
- **H-HPS-S-90.** `EstablecerDatos` fija el foco en `Me.ComandoRegistrarExcelEnviado` (`:1494`) en cada repintado, incluidos los que dispara el propio usuario al terminar un registro: el foco salta de forma inesperada y el repintado posterior puede dejar el foco sobre un control que ya no corresponde al estado.
- Las fechas tecleadas en los `InputBox` no se validan antes de escribirse: se pasan como cadena a la capa de datos, que las reinterpreta con `IsDate` (`SolicitudOperaciones.cls:2032-2034`).
- `DiasParaCancelacion` se calcula con `DateDiff` sin acarreo de signo (`:1632-1638`) y sólo se recalcula si hay fecha prevista de Marga o, en su defecto, pre-Marga: si no hay ninguna, el valor anterior se queda en pantalla.
- El formulario no comprueba ningún permiso: no hay referencia a `EsTecnico`, a `EsAdministrador` ni a los permisos por entidad, de modo que todas las fases (alta en Marga, DPS, ONS, registro en HPS) y todos sus borrados quedan disponibles para cualquiera que llegue a la pestaña.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: cronograma como vista de hitos con dependencias declaradas en el dominio y no en el botón; los `InputBox` sustituidos por campos de fecha con validación y valor por defecto visible; el registro en HPS como una acción única e idempotente; y las fechas previstas recalculadas por el servidor, no por un botón «Contador».

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin solicitud activa o sin registro de fechas, el cronograma se muestra vacío y se explica por qué.
2. Sólo se puede registrar el DPS si el alta en Marga está registrada, y sólo el envío a la ONS si el DPS lo está.
3. Registrar un hito deja la fecha visible, habilita su botón de borrar y deshabilita los hitos que dependen de él.
4. Borrar un hito pide confirmación y deja el cronograma y los botones en el estado que corresponde.
5. El registro en HPS exige el envío a la ONS, avisa de si crea o actualiza el usuario y no puede repetirse sobre una solicitud ya registrada.
6. Los botones de «ver documento» abren el documento de la fase correspondiente.
7. Un perfil sin permiso no ve ni puede ejecutar las fases de registro y de borrado.
8. El número de días para la autocancelación se corresponde con la fecha prevista vigente.

## Pendientes

- Por qué el formulario de renovación tiene fase de Marga (alta, DPS, recordatorios de Marga) y el de traspaso no: la diferencia no está documentada en el código.
- Qué ocurre con los automatismos de correo ya programados cuando se borra una fase: `EliminarFase` vacía la fecha pero no anula los correos programados de esa fase, salvo en el caso de desestimado y cancelación (`SolicitudOperaciones.cls:2158-2161`, `:2251-2254`).
- Si `FechaPrevistaAutocancelacionPreMarga` y `FechaPrevistaAutocancelacionMarga` se recalculan con la lógica de `SolicitudFechasOperaciones` o con la del entorno: en este `.cls` sólo se copian valores ya calculados.
- Recuento de controles: el inventario fija 82 y el `.form.txt` contiene 81 bloques `Begin`; la diferencia no se reproduce con una regla simple de recuento.
