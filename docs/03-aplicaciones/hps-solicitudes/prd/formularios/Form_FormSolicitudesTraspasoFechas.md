# Form_FormSolicitudesTraspasoFechas

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormSolicitudesTraspasoFechas.cls` (1166 líneas) y `.form.txt` (55 bloques `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_FormSolicitudesTraspasoFechas.cls:línea` salvo indicación.

Cronograma de una solicitud de traspaso. Es el segundo destino del control de navegación de `Form_FormSolicitudAltaTraspaso` cuando el tipo de trámite es «Traspaso» (`Form_FormSolicitudAltaTraspaso.cls:122`; `Form_FormSolicitudAltaTraspaso.form.txt:774`). Comparte la estructura de `Form_FormSolicitudRenovacionFechas`, de la que se diferencia en que no tiene fase de Marga ni de DPS y en que añade la fase propia del envío del formulario de traspaso a la ONS.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudAltaTraspaso` | Destino «Fechas» del control de navegación cuando el tipo es «Traspaso» | `Form_FormSolicitudAltaTraspaso.cls:122`, `.form.txt:774` |
| `Form_FormEnvioCorreo` | Tras enviar un correo, el padre invoca `EstablecerDatos` por código | `Form_FormEnvioCorreo.cls:102, 285` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `Form_FormAdjuntarExcelSolicitante` | `DoCmd.OpenForm "FormAdjuntarExcelSolicitante"`, con `WithEvents` | `:330-332` |
| `Form_FormAdjuntaTraspasoONS` | `DoCmd.OpenForm "FormAdjuntaTraspasoONS"`, con `WithEvents` | `:298-301` |
| `Form_FormEnvioCorreo` | Cuatro aperturas, cada una con un `EnumTipoEnvioCorreo` distinto | `:551`, `:582`, `:613`, `:644` |
| `FormWeb` (vía `ObtenerHTMLEnLocal`) | Cuatro botones de «ver documento» | `:849-935` |
| `Form_FormSolicitudesGestion` | `Form_FormSolicitudesGestion.Filtrar` cuando el estado de la solicitud ha cambiado | `:721-725` |

No declara parámetros: trabaja sobre `m_ObjSolicitudActiva` (`VG.bas:197`). La carga de datos la hace `EstablecerDatos` (`:1025-1138`), que es `Public` para que el contenedor y `Form_FormEnvioCorreo` puedan repintarla.

## Controles y eventos

El formulario tiene 29 procedimientos y unos 55 controles, casi todos parejas «fecha + botón» del cronograma. Se agrupan por familia.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` → `Form_Load` | Carga el cronograma y activa la edición | `:966-994`, `.form.txt:39, 374` |
| `FechaEnvioExcel` | cuadro de fecha | — | Envío del Excel de solicitud de información | `.form.txt:917` |
| `ComandoRegistrarExcelEnviado` | botón | `OnClick` | Abre `FormEnvioCorreo` con `EnvioExcel` | `:568-597`, `.form.txt:946-947` |
| `FechaRecepcionExcel` | cuadro de fecha | — | Recepción del Excel del solicitante | `.form.txt:986` |
| `ComandoRegistrarExcelRecibido` | botón | `OnClick` | Abre `FormAdjuntarExcelSolicitante` | `:318-348`, `.form.txt:1015-1016` |
| `FechaCorreoRecordatorioExcel1` / `2` | cuadros de fecha | — | Envío de los recordatorios del Excel | `.form.txt:1057, 1158` |
| `ComandoEnvioRecordatorioExcel1` / `2` | botones | `OnClick` | Abren `FormEnvioCorreo` con `RecordatorioExcel1` / `2` | `:599-662`, `.form.txt:1118, 1219` |
| `FechaEnvioONS` | cuadro de fecha | — | Envío a la ONS | `.form.txt:1259` |
| `ComandoRegistrarEnvioTraspasoONS` | botón | `OnClick` | Abre `FormAdjuntaTraspasoONS` para adjuntar el formulario de traspaso | `:285-316`, `.form.txt:1741-1742` |
| `ComandoEliminarEnvioTraspasoONS` | botón | `OnClick` | Borra la fase de envío a la ONS, con confirmación | `:154-198`, `.form.txt:2247-2248` |
| `ComandoURLDocumentoTraspasoEnviado` | botón | `OnClick` | Abre el documento de traspaso ya enviado | `:746-777`, `.form.txt:2553-2554` |
| `FechaDesestimado` | cuadro de fecha | — | Fecha de desestimación | `.form.txt:1291` |
| `ComandoRegistrarDesestimado` | botón | `OnClick` | Pide la fecha por `InputBox` y registra el desestimado | `:241-283`, `.form.txt:1704-1705` |
| `FechaCancelado` | cuadro de fecha | — | Fecha de cancelación | `.form.txt:1415` |
| `ComandoRegistraCancelacion` | botón | `OnClick` | Pide la fecha por `InputBox` y registra la cancelación | `:200-239`, `.form.txt:2596-2597` |
| `DiasParaCancelacion` | cuadro numérico | — | Días que faltan para la autocancelación | `:1128-1130`, `.form.txt:1355` |
| `FechaPrevista*` (3 controles) | cuadros de fecha | — | Próximos correos previstos | `.form.txt:1509, 1573, 1637` |
| `ComandoEnvioCancelacionPreMarga` | botón | `OnClick` | Abre `FormEnvioCorreo` con `CancelacionPreMarga` | `:537-566`, `.form.txt:1667-1668` |
| `FechaRegistroEnHPS` | cuadro de fecha | — | Fecha de registro en el aplicativo de HPS | `.form.txt:2637` |
| `ComandoRegistroHPS` | botón | `OnClick` | Crea o actualiza el `UsuarioHPS` y cierra la tramitación | `:664-744`, `.form.txt:2666-2667` |
| `ComandoURLExcelSolicitud` / `ComandoURLExcelRespuesta` | botones | `OnClick` | Abren la plantilla del Excel o el Excel recibido | `:779-847`, `.form.txt:1778, 1822` |
| `ComandoVer*` (4 botones) | botones | `OnClick` | Proyectan en `FormWeb` el HTML de recordatorios y cancelaciones | `:849-935`, `.form.txt:1864-2005` |
| `ComandoEliminar*` (7 botones) | botones | `OnClick` | Borran la fase o la fecha correspondiente, con confirmación | `:118-536`, `.form.txt:2052-2486` |
| `ComandoActualizarRecordatoriosExcelPreMarga` | botón «Contador» | `OnClick` | Recalcula las fechas previstas de recordatorio y las graba | `:49-117`, `.form.txt:2442-2443` |

## Flujos

### Carga del cronograma (`Form_Load` → `EstablecerDatos`, `:966-994`, `:1025-1138`)

1. `Form_Load` desactiva la edición, llama a `EstablecerDatos` y la reactiva (`:974-984`).
2. `EstablecerDatos` sale sin pintar si no hay solicitud activa (`:1031-1034`) o si la solicitud no tiene `SolicitudFechas` (`:1035-1038`).
3. Fija el foco en `Me.ComandoRegistrarExcelEnviado` (`:1039`) y, para cada fecha, la pinta o la vacía según sea una fecha válida, dejando visible el botón de eliminar correspondiente sólo cuando la fecha existe (`:1041-1085`).
4. En el bloque de la ONS habilita el acceso al documento y el registro en HPS si hay fecha de envío, y habilita el adjunto del formulario de traspaso si todavía no se ha registrado en HPS (`:1086-1098`).
5. Calcula `DiasParaCancelacion` con la fecha prevista de autocancelación pre-Marga (`:1128-1130`).

### Registro de hitos

- **Desestimado** (`:241-283`) y **cancelación** (`:200-239`) piden la fecha con un `InputBox` cuyo valor por defecto es la fecha del sistema; si el usuario cancela, el diálogo devuelve cadena vacía y el handler sale sin escribir (`:206-209`, `:247-250`).
- El **envío del formulario de traspaso a la ONS** no se registra aquí: este botón abre `FormAdjuntaTraspasoONS` (`:285-316`) y es ese formulario el que escribe la fecha y el adjunto (`SolicitudOperaciones.RegistroEnvioTraspasoONS`).
- Tras cada registro se invalida la caché de fechas de la solicitud, se repinta el cronograma y se llama a `PintarIndicadores` (`:222-229` y equivalentes).

### Recálculo de fechas previstas

- `ComandoActualizarRecordatoriosExcelPreMarga_Click` (`:49-117`) copia a las fechas previstas las fechas calculadas (`…Calculado`) de los dos recordatorios de Excel y de la autocancelación pre-Marga, y las graba con `SolicitudFechasOperaciones.Registrar` (`:87-95`).

### Envío de correos

- Los cuatro botones de envío abren `FormEnvioCorreo` con un `EnumTipoEnvioCorreo` distinto: `CancelacionPreMarga` (`:551`), `EnvioExcel` (`:582`), `RecordatorioExcel1` (`:613`) y `RecordatorioExcel2` (`:644`). En todos los casos se cierra antes la instancia previa de `FormEnvioCorreo`.

### Borrado de fases

- `Eliminar` (`:11-47`) delega en `SolicitudOperaciones.EliminarFase` (`:23`), invalida la caché, repinta y llama a `PintarIndicadores` (`:32-39`).
- Los handlers de borrado piden confirmación con un `MsgBox` cuyo botón por defecto es «No» y, si se confirma, llaman a `Eliminar` con la fase correspondiente (`:118-152` y equivalentes).
- Dos borrados no pasan por `Eliminar`: el Excel recibido (`:421-463`) y el envío a la ONS (`:154-198`), que llaman directamente a `EliminarExcelDevuelto` y `EliminarEnvioTraspasoONS`.

### Cierre de la tramitación

- `ComandoRegistroHPS_Click` (`:664-744`) exige la fecha de envío a la ONS (`:678-681`), pide a la capa de datos el motivo que impide registrar el usuario HPS (`:684-691`), llama a `RegistrarUsuarioHPSEnHPS` (`:696`) y avisa del resultado (`:700-709`). Si el estado de la solicitud ha cambiado, refresca `FormSolicitudesGestion` (`:720-725`).

### Errores

Todos los handlers usan el patrón `On Error GoTo errores` con el centinela `Err.Number = 1000` para los errores ya descritos en `m_Error`; los errores ajenos rellenan el mensaje, envían `CorreoAlAdministrador` (`FU.bas:585`) y muestran un `MsgBox` crítico. `Form_Load` no cierra el formulario tras el aviso (`:986-993`).

## Reglas, validaciones y cálculos

- **Sin solicitud o sin registro de fechas:** el cronograma se queda vacío, sin aviso (`:1031-1038`).
- **Recepción del Excel previa al envío a la ONS:** lo exige la capa de datos al registrar el adjunto (`SolicitudOperaciones.cls:1649-1656`).
- **Desestimado y cancelación:** se registran con la fecha tecleada; ambas anulan los automatismos programados (`SolicitudOperaciones.cls:2158-2161`, `:2251-2254`).
- **Días para la autocancelación:** `DateDiff("d", Date, FechaPrevista…)` (`:1128-1130`), sin acarreo de signo.
- **Visibilidad de los botones de eliminar:** depende de que la fecha exista (`:1041-1085`).
- **Escritura:** todas las fases pasan por `SolicitudOperaciones`, que encadena `UPDATE` sobre `TbSolicitudesFechas` y registra el log (`SolicitudOperaciones.cls:1311-1450`); las fechas previstas se graban con `SolicitudFechasOperaciones.Registrar` (`SolicitudFechasOperaciones.cls:190-414`).
- **Sin transacciones:** ninguna de estas rutas abre transacción.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Fechas de la solicitud | `SolicitudFechas` de la solicitud de sesión | Derivado, sin DAO en este `.cls` | `:1035` |
| Alta, desestimado, cancelación | `TbSolicitudesFechas` y estado de `TbSolicitudes` | `UPDATE` con SQL concatenado | `SolicitudOperaciones.cls:1311-1450`, `:2118-2261` |
| Fechas previstas recalculadas | `TbSolicitudesFechas` | DAO (`Recordset` en `SolicitudFechasOperaciones`) | `SolicitudFechasOperaciones.cls:190-414` |
| Excel recibido | `TbSolicitudes` y `TbSolicitudesFechas` | `UPDATE` con SQL concatenado | `SolicitudOperaciones.cls:1544-1618` |
| Envío del formulario de traspaso a la ONS | `TbSolicitudes`, `TbSolicitudesFechas` y copia del fichero | `UPDATE` con SQL concatenado y E/S binaria | `SolicitudOperaciones.cls:1619-1706` |
| Registro en HPS | usuario de HPS en la base de HPS | DAO remoto | `SolicitudOperaciones.cls:1158-1242` |
| Log de fases | log general | DAO | `FU.bas:895-937` |
| HTML de recordatorios | ficheros locales generados y `FormWeb` | E/S de fichero | `FU.bas:1336-1392` |

## Coexistencia

Pendiente de `../coexistencia.md`. Lo que ya se ve: el cronograma de traspaso es una versión recortada del de renovación, con menos fases y una fase propia (el formulario de traspaso a la ONS). Las escrituras son `UPDATE` con SQL concatenado sobre `TbSolicitudesFechas` y no hay transacción que agrupe la fecha, el estado y el log. En la web el equivalente es una vista de hitos con su regla de precedencia declarada en el dominio y aplicada en el servidor.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-91.** `EstablecerDatos` deja `Me.ComandoRegistroHPS` habilitado siempre que haya fecha de envío a la ONS (`:1089`), incluso cuando la solicitud ya está registrada en HPS: el bloque interno deshabilita el adjunto del formulario de traspaso cuando `FechaRegistroEnHPS` existe (`:1090-1092`), pero el botón de cierre de tramitación queda activo, de modo que la solicitud puede registrarse dos veces en HPS.
- **H-HPS-S-92.** `EstablecerDatos` decide si habilita el adjunto del formulario de traspaso leyendo `m_SolFechas.FechaRegistroEnHPS` (`:1090`) —el dato del modelo— y, en cambio, pinta `Me.FechaRegistroEnHPS` más abajo (`:1106-1111`): son dos fuentes distintas para la misma decisión, y la que se pinta es la que llega a la interfaz, de modo que el rótulo de la fecha puede no coincidir con el estado que ha decidido la habilitación de los botones.
- **H-HPS-S-93.** El mensaje de error de `ComandoEliminarExcelRecibido_Click` identifica a otro handler: dice «Al ComandoEliminarExcelEnviado_Click se ha producido el error n» (`:457`).
- **H-HPS-S-94.** Rótulos heredados: la ventana conserva el `Caption` de diseño «EXPEDIENTES Versión 2024-001 Administrador» (`.form.txt:39`), que ningún handler sobrescribe; el botón de adjuntar el Excel recibido dice «ADJUNTAR EXCEL RECIBIDO DE SOLCITUD DE INFORMACIÓN», con el typo «SOLCITUD» (`.form.txt:1016`); los cuatro botones de ver documento llevan el rótulo de diseño «Comando36» (`.form.txt:1865, 1912, 1959, 2006`) y los siete de borrar el rótulo «Comando47» (`.form.txt:2053, 2118, 2183, 2248, 2313, 2379, 2487`), de modo que ninguno tiene texto visible; y la etiqueta «PRÓX CORREO» se repite en tres rótulos distintos (`.form.txt:1477, 1540, 1604`).
- **H-HPS-S-95.** `ComandoRegistroHPS_Click` accede a `m_ObjSolicitudActiva.SolicitudFechas.FechaEnvioONS` (`:678`) sin comprobar que `SolicitudFechas` exista: con una solicitud sin registro de fechas se produce el error de objeto no establecido, que el manejador convierte en correo al administrador.
- `EstablecerDatos` fija el foco en `Me.ComandoRegistrarExcelEnviado` (`:1039`) en cada repintado, incluidos los que dispara el propio usuario al terminar un registro.
- `ComandoRegistrarEnvioTraspasoONS_Click` no comprueba que exista solicitud activa antes de abrir `FormAdjuntaTraspasoONS` (`:298`), y sólo asigna `m_FormAnexoTraspaso` si el formulario ha quedado abierto (`:299-301`): si la apertura falla, el evento `AdjuntadoEnvio` no llega y el padre no se repinta.
- Las fechas tecleadas en los `InputBox` no se validan antes de escribirse: se pasan como cadena a la capa de datos, que las reinterpreta con `IsDate`.
- `DiasParaCancelacion` se calcula con `DateDiff` sin acarreo de signo (`:1128-1130`) y sólo se recalcula si hay fecha prevista pre-Marga; si no la hay, el valor anterior se queda en pantalla.
- El formulario no comprueba ningún permiso: no hay referencia a `EsTecnico`, a `EsAdministrador` ni a los permisos por entidad, de modo que todas las fases del traspaso y sus borrados quedan disponibles para cualquiera que llegue a la pestaña.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: cronograma como vista de hitos con dependencias declaradas en el dominio; el cierre de tramitación en HPS como una acción idempotente; el adjunto del formulario de traspaso como una subida con su fecha en el mismo formulario de la fase; y las fechas previstas recalculadas por el servidor en lugar de por un botón «Contador».

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin solicitud activa o sin registro de fechas, el cronograma se muestra vacío y se explica por qué.
2. El envío del formulario de traspaso a la ONS exige la recepción previa del Excel del solicitante y no puede ser anterior a ella.
3. Registrar un hito deja la fecha visible, habilita su botón de borrar y deshabilita los hitos que dependen de él.
4. Borrar un hito pide confirmación y deja el cronograma y los botones en el estado que corresponde.
5. El registro en HPS exige el envío a la ONS, avisa de si crea o actualiza el usuario y no puede repetirse sobre una solicitud ya registrada.
6. Los botones de «ver documento» abren el documento de la fase correspondiente.
7. Un perfil sin permiso no ve ni puede ejecutar las fases de registro y de borrado.
8. El número de días para la autocancelación se corresponde con la fecha prevista vigente.

## Pendientes

- Por qué el traspaso no tiene fase de DPS ni de Marga y la renovación sí: la diferencia no está documentada en el código.
- Qué ocurre con los automatismos de correo ya programados cuando se borra una fase: `EliminarFase` vacía la fecha pero no anula los correos programados de esa fase, salvo en el caso de desestimado y cancelación (`SolicitudOperaciones.cls:2158-2161`, `:2251-2254`).
- Si `FechaPrevistaAutocancelacionPreMarga` se recalcula con la lógica de `SolicitudFechasOperaciones` o con la del entorno: en este `.cls` sólo se copian valores ya calculados.
- Recuento de controles: el inventario fija 56 y el `.form.txt` contiene 55 bloques `Begin`; la diferencia no se reproduce con una regla simple de recuento.
