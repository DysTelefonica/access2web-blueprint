# Form_FormSolicitudesAltaFechas

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormSolicitudesAltaFechas.cls` (1661 líneas) y `.form.txt` (5080 líneas) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_X.cls:línea` del commit fijado.

Panel de fechas de una solicitud de alta: es la única pantalla donde el tramitador recorre y corrige el ciclo completo de la tramitación (envío del Excel, recordatorios, recepción, alta en Marga, envío a las unidades, registro en HPS, desestimación y cancelación). Es un formulario **no enlazado** (sin `RecordSource`): los 82 controles son presentacionales y el estado vive en la solicitud de sesión. Concentra 38 botones de acción y una función pública `EstablecerDatos` que los repinta a todos.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudAltaTraspaso` | El contenedor lo carga como subformulario del panel de navegación, solo cuando el tipo es alta: `Me.NavFechas.NavigationTargetName = "FormSolicitudesAltaFechas"`; para traspaso carga su propio hermano | `Form_FormSolicitudAltaTraspaso.cls:119-123` |
| `Form_FormEnvioCorreo` | Al abrirse, comprueba cuál es el subformulario activo del contenedor y, si es este, lo repinta llamando a su `EstablecerDatos` | `Form_FormEnvioCorreo.cls:98-101` |

No abre formularios salvo los de apoyo: `FormEnvioCorreo`, desde los siete botones de envío y de registro de fase (`:313`, `:862`, `:895`, `:928`, `:961`, `:993`, `:1026`, `:1059`), y `FormAdjuntarExcelSolicitante` (`:486`), al que se suscribe con `WithEvents` para recibir `Adjuntado` (`:5`, `:1465`).

No declara parámetros. Lee `m_ObjSolicitudActiva` (`VG.bas:197`), `m_ObjEntorno` (`VG.bas:190`) y el tipo calculado de la solicitud. Avisa al exterior con una llamada a la función pública `Filtrar` de la gestión de solicitudes (`:1142-1143`).

## Controles y eventos

Solo los controles con decisión de negocio; el resto son campos de fecha y etiquetas de estado puramente presentacionales.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `FechaEnvioExcel` | texto | — | Fecha de salida del Excel hacia la empresa | `.form.txt:609`; pintado en `:1511-1516` |
| `ComandoRegistrarExcelEnviado` | botón | `OnClick` | Abre `FormEnvioCorreo` con la fase de envío | `:914-946`, `.form.txt:638-640` |
| `FechaRecepcionExcel` | texto | — | Fecha de vuelta del Excel cumplimentado | `.form.txt:678`; pintado en `:1547-1553` |
| `ComandoRegistrarExcelRecibido` | botón | `OnClick` | Abre el formulario de adjunto y engancha su evento `Adjuntado` | `:474-503`, `.form.txt:707-709` |
| `ComandoActualizarRecordatoriosExcelPreMarga` / `…Marga` | botones | `OnClick` | Recalcula las fechas previstas de recordatorios y cancelación automática y las persiste | `:48-114`, `:116-186`, `.form.txt:2993`, `.form.txt:3036` |
| `FechaPrevista*` (6 textos) | texto | — | Próximo correo previsto por fase | `.form.txt:1401, 1465, 1527, 1589, 1689, 1753` |
| `ComandoRegistrarAltaMarga` | botón | `OnClick` | Pide por `InputBox` la fecha de alta en Marga y la registra | `:505-544`, `.form.txt:1857-1859` |
| `ComandoRegistrarDesestimado` | botón | `OnClick` | Registra la desestimación de la solicitud | `:332-372`, `.form.txt:1820-1822` |
| `FechaEnvioDPS`, `FechaEnvioONS`, `FechaRegistroEnHPS` | texto | — | Hito de la derivación a las unidades y cierre en HPS | `.form.txt:2833, 3257, 3187`; pintado en `:1592-1596`, `:1598-1623`, `:1624-1629` |
| `ComandoRegistrarEnvioONS`, `ComandoRegistroHPS`, `ComandoEliminarONS` | botones | `OnClick` | Encadenan el envío ONS y el registro en la aplicación de HPS | `:418-473`, `:1081-1161`, `:258-292`, `.form.txt:3383, 3217, 3316` |
| `ComandoURLExcelSolicitud`, `ComandoURLExcelRespuesta` | botones | `OnClick` | Abre la plantilla que sale hacia la empresa y el anexo devuelto | `:1193-1231`, `:1162-1192`, `.form.txt:1894-1897, 1938-1941` |
| `ComandoVer*` (7 botones) | botones | `OnClick` | Abre en `FormWeb` el HTML de cada fase, enviado o previsto | `:1232-1434`, `.form.txt:1980-1983, 2027-2030, 2074-2077, 2121-2124, 2168-2171, 2215-2218, 2262-2265` |
| `ComandoEliminar*` (11 botones) | botones | `OnClick` | Borran la fecha de una fase, previa confirmación en `MsgBox` | `:187-847`, `.form.txt:2309-2312, 2374-2377, 2439-2442, 2504-2507, 2569-2572, 2634-2637, 2699-2702, 2764` |
| `ComandoEnvio*` (6 botones) | botones | `OnClick` | Envían a mano el correo de una fase y fijan su fecha | `:848-1080`, `.form.txt:1619-1621, 1783-1785` |

## Flujos

### Apertura (`Form_Load`, `:1435-1463`)

1. Desactiva `AllowEdits` (`:1443`) y llama a `EstablecerDatos` (`:1445`), que lee `m_ObjSolicitudActiva.SolicitudFechas` (`:1504`) y, para cada fecha, escribe el valor en su cuadro y muestra u oculta el botón de borrado correspondiente; si no hay solicitud o no hay registro de fechas, sale sin error (`:1499-1507`). El foco inicial va al botón de envío del Excel (`:1508`).
2. Reactiva `AllowEdits` (`:1453`) y, en error, avisa al administrador y muestra `MsgBox` sin cerrar (`:1455-1462`).

### Alta de una fase (patrón común de los 8 `ComandoRegistrar*`)

1. Delega en `SolicitudOperaciones` o en `SolicitudFechasOperaciones` con la solicitud de sesión (`:516-522`, `:1099-1101`).
2. Descarta la caché con `Set m_ObjSolicitudActiva.SolicitudFechas = Nothing` (`:528`), lo que obliga a releer `TbSolicitudesFechas`.
3. Vuelve a pintar con `EstablecerDatos` (`:530`) y refresca los contadores globales con `PintarIndicadores` (`:534`; `FU.bas:41-144`).

### Baja de una fase (`Eliminar`, `:10-46`, compartido por los 11 botones)

1. Pide confirmación con `MsgBox` de sí/no (`:192`, `:672`).
2. Llama a `SolicitudOperaciones.EliminarFase` con el valor de `EnumTipoEnvioCorreo` (`VG.bas:150-165`; `SolicitudOperaciones.cls:1311-1452`).
3. Descarta la caché, repinta y refresca (`:30-38`).

### Cerrar la tramitación (`ComandoRegistroHPS_Click`, `:1081-1161`)

1. Exige la fecha de envío ONS (`:1096`) y pasa el control previo `MotivoNoOKRegistrarUsuarioHPSEnHPS` (`:1101`).
2. Llama a `RegistrarUsuarioHPSEnHPS` (`:1113`) y comunica si se creó o actualizó el usuario en la aplicación de HPS (`:1117-1130`).
3. Relee la solicitud con `constructor.getSolicitud` (`constructor.bas:483`; `:1132`) y, si el estado ha cambiado, refresca la lista de la gestión (`:1140-1146`).

### Alternativos

- **Adjunto ya disponible:** en lugar de registrar a mano, `ComandoRegistrarExcelRecibido_Click` abre `FormAdjuntarExcelSolicitante` y espera su evento `Adjuntado`, que vuelve a pintar todo el panel (`:1465-1491`).
- **Alta en Marga:** la fecha se pide con `InputBox` (`:512`) y se abandona en silencio si se cancela (`:513-515`).
- **Plantillas:** `ComandoURLExcelSolicitud` elige la plantilla según el tipo calculado —alta o renovación, traspaso— y avisa «Tipo de desconocido» si no reconoce ninguno (`:1201-1211`).

### Errores

Los 38 handlers repiten el mismo esquema: `On Error GoTo errores`, acumulación en `m_Error`, `CorreoAlAdministrador` (`FU.bas:585`) solo si el número no es el 1000 interno y `MsgBox` crítico. `m_FormAnexo_Adjuntado` y `EstablecerDatos` son los dos únicos que no muestran el `MsgBox` (`:1484-1490`, `:1655-1659`).

## Reglas, validaciones y cálculos

- **Cronología interna:** `EstablecerDatos` no valida el orden de las fechas; las reglas viven en la capa de datos (`SolicitudOperaciones.cls:1474-1491`, `SolicitudFechasOperaciones.cls:444`).
- **Borrado en cascada:** borrar el envío del Excel anula también los dos recordatorios y la autocancelación prevista (`SolicitudOperaciones.cls:1327-1345`).
- **Cierre en HPS bloquea la fase:** con fecha de registro en HPS, `ComandoRegistrarEnvioONS` y `ComandoRegistroHPS` se deshabilitan y el botón de baja de ONS se oculta (`:1600-1607`).
- **Reprogramación:** los dos botones de recálculo toman las fechas previstas calculadas por la clase y las escriben con `SolicitudFechasOperaciones.Registrar EnumSiNo.no` (`:62-92`, `:130-160`).
- **Refresco de contadores:** todo cambio pasa por `PintarIndicadores`, que recalcula los pendientes del entorno y repinta tareas y menú (`FU.bas:41-144`).
- **Visibilidad por estado:** cada fecha gobierna la visibilidad de su propio botón de borrado (`:1510-1649`); no hay un único modelo de permisos en esta pantalla.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Leer el registro de fases | `TbSolicitudesFechas` | DAO por `IDSolicitud` | `constructor.bas:440`; `Solicitud.cls:966` |
| Reler la solicitud tras el cierre | `TbSolicitudes` | DAO por `IDSolicitud` | `constructor.bas:483` (`:1132`) |
| Dar de alta una fase | `TbSolicitudesFechas` y tabla de automatismos | `INSERT` / `UPDATE` | `SolicitudFechasOperaciones.cls:190-413`; `SolicitudOperaciones.cls:452` |
| Borrar una fase | `TbSolicitudesFechas` | `UPDATE` con SQL concatenado | `SolicitudOperaciones.cls:1337-1345` |
| Registrar el cierre en HPS | solicitud, usuario HPS y fecha de cierre | DAO multi-tabla | `SolicitudOperaciones.cls:1158-1266` |
| Anular o reprogramar avisos | tabla de automatismos | DAO | `AutomatizacionRepositorio.bas`; `SolicitudOperaciones.cls:2318, 2428` |
| Abrir el HTML de una fase | sistema de ficheros y `FormWeb` | E/S binaria | `FU.bas:433-472`; `FU.bas:1738-1741` |
| Registrar en el log | log general | DAO | `SolicitudOperaciones.cls:1528-1530` |

## Coexistencia

Pendiente de `../coexistencia.md`. Lo que ya se ve: la pantalla es un mapa de fases con botones por fase, no un formulario de datos; el `EnumTipoEnvioCorreo` de 14 valores (`VG.bas:150-165`) y el `EnumEstado` de 16 (`VG.bas:122-139`) son el modelo de dominio que debe sobrevivir a la traducción, no los rótulos de los botones. El borrado y el registro se expresan como `UPDATE` con SQL concatenado sobre `TbSolicitudesFechas` (`SolicitudOperaciones.cls:1337-1345`); la versión web necesita transacciones explícitas, porque aquí cada fase se persiste por separado. La reprogramación de avisos depende de un planificador del sistema de ficheros, sin equivalente directo en PostgreSQL.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-20.** `EstablecerDatos` decide el estado de los botones leyendo controles ya pintados en lugar del modelo: comprueba `Me.FechaRegistroEnHPS` en `:1600` (que solo se pinta en `:1624-1629`) y `Me.FechaEnvioDPS` en `:1612` (que solo se pinta en `:1592-1596`). En la primera pasada ambos valen lo que tuviera el formulario, de modo que con ONS enviado y HPS ya registrado se pueden dejar habilitados `ComandoRegistrarEnvioONS` y `ComandoEliminarONS`.
- **H-HPS-S-21.** Rótulos residuales de otra aplicación y de diseño: la ventana arranca como «EXPEDIENTES Versión 2024-001 Administrador» (`.form.txt:39`), el botón de recepción dice «ADJUNTAR EXCEL RECIBIDO DE SOLCITUD DE INFORMACIÓN» con falta de «u» (`.form.txt:708`), siete botones «Ver» comparten el rótulo `Comando36` (`.form.txt:1981, 2028, 2075, 2122, 2169, 2216, 2263`) y once botones de borrado comparten `Comando47` (`.form.txt:2310, 2375, 2440, 2505, 2570, 2635, 2700, 2764`), dos botones de apertura de Excel se llaman solo «E» (`.form.txt:1895, 1939`) y los rótulos de recordatorio y cancelación están duplicados entre la fase Excel y la de Marga (`.form.txt:809` y `.form.txt:1043`; `.form.txt:1620` y `.form.txt:1784`).
- **H-HPS-S-22.** Los mensajes de error no identifican el handler: los 38 repiten «se ha producido el error n» y `ComandoEliminarExcelRecibido_Click` anuncia «Al ComandoEliminarExcelEnviado_Click» (`:703`), lo que desvía la localización del fallo.
- `ComandoRegistrarExcelRecibido_Click` cierra `FormAdjuntarExcelSolicitante` con `DoCmd.Close` sin especificar `acSaveNo` (`:482-484`) y reencola `WithEvents` sobre la instancia recién abierta sin anular antes la referencia anterior (`:487-489`).
- `EstablecerDatos` devuelve sin error cuando no hay solicitud o no hay registro de fases (`:1499-1507`): la pantalla queda con los valores del diseño sin avisar a nadie.
- La fase de cierre reescribe la solicitud de sesión con `constructor.getSolicitud` (`:1132`) mientras el resto de handlers solo descartan `SolicitudFechas`; dos caminos distintos para el mismo problema de caché.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: línea de tiempo por fases en lugar de cuadrícula de botones; una operación de «deshacer fase» con registro de auditoría en vez de once confirmaciones independientes; el conjunto de fases habilitadas como dato del servidor en lugar de calcularse pintando controles; recálculo de fechas previstas en segundo plano, sin botón manual.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir la pantalla, cada fecha registrada aparece en su cuadro y su botón de borrado queda visible; las fases no registradas aparecen vacías y sin botón.
2. Registrar una fase la persiste, repinta la pantalla y actualiza los contadores de pendientes de la aplicación.
3. Borrar una fase pide confirmación y, tras aceptarla, deja la fecha y el botón en el estado inicial.
4. Borrar el envío del Excel deja también vacíos sus dos recordatorios y su autocancelación prevista.
5. Registrar el Excel recibido se puede hacer desde el botón o desde el formulario de adjunto, y ambos caminos dejan la pantalla igual.
6. Con el registro en HPS hecho, el envío ONS y el registro en HPS quedan bloqueados.
7. El cierre en HPS informa de si creó o actualizó el usuario y, si cambia el estado, la lista de la gestión se refresca.
8. Abrir el HTML o el Excel de una fase deja el fichero en el equipo del tramitador y lo abre con la aplicación asociada.

## Pendientes

- Los handlers de los formularios hermanos `FormSolicitudRenovacionFechas` (1652 líneas), `FormSolicitudesAltaRenovacionFechas` (751) y `FormSolicitudesTraspasoFechas` (1166) replican este patrón con diferencias; el alcance de la ficha es solo el de alta.
- Si `DiasParaCancelacion` (`.form.txt:1247`) es de lectura o lo fija algún cálculo externo.
- Qué ocurre con los botones de borrado cuando la fase se ha enviado pero su aviso programado sigue vivo (`SolicitudOperaciones.cls:2318-2427`).
- La lista completa de controles presentacionales, no cubierta aquí por decisión de alcance.