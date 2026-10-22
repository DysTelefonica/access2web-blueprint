# Form_FormSolicitudesTraspasoFechas

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormSolicitudesTraspasoFechas.cls` (1 166 líneas) y `.form.txt` (55 bloques `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_FormSolicitudesTraspasoFechas.cls:línea` salvo indicación.

Panel de hitos del traspaso, con la misma arquitectura que `Form_FormSolicitudesAltaRenovacionFechas` pero con dos diferencias que la hacen distinta: admite **deshacer** cada hito (diez botones `ComandoEliminar*`) y trabaja con dos formularios de adjuntos. Cada hito se puede registrar, deshacer, previsualizar o enviar por correo. **Enlazado por control** a `TbSolicitudesFechas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudAltaTraspaso` | Se declara como segundo destino del control de navegación del contenedor | `Form_FormSolicitudAltaTraspaso.cls:82-83` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `Form_FormAdjuntarExcelSolicitante` | `WithEvents`; recibe `Adjuntado` y repinta | `:5, 996-1023` |
| `Form_FormAdjuntaTraspasoONS` | `WithEvents`; recibe `AdjuntadoEnvio` | `:6, 1140` |
| `FormEnvioCorreo` | Los `ComandoEnvio*_Click` la abren por tipo y deshabilitan `Programado` | `:537-663` |
| `FormWeb` | Los `ComandoVer*_Click` proyectan el HTML de la plantilla | `:849-965` |
| `SolicitudOperaciones` | `EliminarFase`, `EliminarExcelDevuelto` y el resto de hitos | `:23, 438` y hermanos |
| `Form_FormSolicitudesGestion` y menú | `PintarIndicadores` recalcula los contadores | `:39, 135` y hermanos |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` → `Form_Load` | `EstablecerDatos` y reactiva la edición | `:966-994` |
| `FechaEnvioExcel`, `FechaRecepcionExcel`, `FechaCorreoRecordatorioExcel1/2`, `FechaEnvioONS`, `FechaDesestimado`, `FechaCancelado`, `FechaRegistroEnHPS`, `FechaPrevistaCorreoRecordatorioExcel1/2`, `FechaPrevistaAutocancelacionPreMarga` | controles enlazados a `SolicitudFechas` | — | Los hitos que `EstablecerDatos` pinta | `:1030-1127`, `.form.txt:917-2637` |
| `DiasParaCancelacion` | etiqueta o texto | — | Días restantes hasta la autocancelación pre-Marga | `:1128-1130`, `.form.txt:1355` |
| `ComandoRegistrarExcelEnviado`, `ComandoRegistrarExcelRecibido`, `ComandoRegistrarDesestimado`, `ComandoRegistrarEnvioTraspasoONS`, `ComandoRegistraCancelacion`, `ComandoRegistroHPS` | botones | `OnClick` | Registran cada hito | `:200-317, 568, 664` |
| `ComandoEliminarExcelEnviado`, `ComandoEliminarExcelRecibido`, `ComandoEliminarRecordatorioExcel1/2`, `ComandoEliminarDesestimado`, `ComandoEliminarCancelacion`, `ComandoEliminarEnvioTraspasoONS` | botones | `OnClick` | Deshacen cada hito, con confirmación | `:118-537` |
| `ComandoEnvioCancelacionPreMarga`, `ComandoEnvioRecordatorioExcel1/2` | botones | `OnClick` | Abre `FormEnvioCorreo` por tipo | `:537-663` |
| `ComandoVerCancelacionPreMarga`, `ComandoVerHTMLEnvioExcel`, `ComandoVerRecordatorioExcel1/2` | botones | `OnClick` | Proyecta el HTML de la plantilla | `:849-965` |
| `ComandoURLDocumentoTraspasoEnviado`, `ComandoURLExcelRespuesta`, `ComandoURLExcelSolicitud`, `ComandoActualizarRecordatoriosExcelPreMarga` | botones | `OnClick` | Documentos y recarga de recordatorios | `:746-848, 49` |

## Flujos

### Normal — carga

1. `Form_Load` llama a `EstablecerDatos` y propaga el error (`:976-979`).
2. `EstablecerDatos` (`:1025-1138`) sale sin error si no hay solicitud activa o sin `SolicitudFechas` (`:1031-1038`), da el foco a «Registrar Excel enviado» (`:1039`) y a continuación recorre una cadena de `If … Else` que decide, para cada hito ya registrado, si su botón de deshacer y sus accesos a documento son visibles (`:1041-1125`).
3. Calcula los días hasta la autocancelación, esta vez sí protegido con `IsDate` (`:1128-1130`).

### Normal — registrar un hito

1. Los handlers `ComandoRegistrar*` preguntan la fecha o lanzan `FormEnvioCorreo` / los formularios de adjuntos y, al terminar, invalidan `SolicitudFechas`, repintan y recalculan los indicadores (`:200-317, 568, 664` y hermanos).
2. `m_FormAnexo_Adjuntado` (`:996-1023`) y `m_FormAnexoTraspaso_AdjuntadoEnvio` (`:1140`) repintan cuando el formulario de adjuntos avisa de que ha terminado.

### Normal — deshacer un hito

1. `Public Function Eliminar` (`:11-46`) es la implementación común de cinco botones: llama a `SolicitudOperaciones.EliminarFase` con el `EnumTipoEnvioCorreo` correspondiente, invalida `SolicitudFechas`, repinta y recalcula los indicadores (`:20-39`).
2. Los cinco llamadores piden confirmación con `MsgBox` y `vbDefaultButton2` antes de invocarla (`:123-130, 359, 395, 475, 509`).
3. Los otros dos botones de deshacer no la usan y llaman a `SolicitudOperaciones` directamente, repitiendo lainvalidación y el repintado (`:154, 421-453`).

### Alternativos

- **Botón de deshacer sobre hito no registrado:** no se pinta (visible `False`), y el botón de registro queda como única vía.
- **Registro de traspaso ONS:** dos rutas distintas según haya o no documento de traspaso ya enviado (`:1088-1096`).
- **Cancelación del registro ONS:** `Exit Sub` sin escribir, como en el formulario hermano.

### Errores

- Patrón uniforme de `On Error GoTo errores` con el centinela `1000`, correo al administrador (`FU.bas:585`) y `MsgBox` crítico, repetido en los treinta handlers.
- `m_FormAnexo_Adjuntado` es la excepción: su rama de error **no muestra mensaje** (`:1015-1023`).

## Reglas, validaciones y cálculos

- **Los hitos se muestran o se ocultan, no se habilitan:** `EstablecerDatos` juega con `Visible` para los botones de deshacer y de documento (`:1043-1125`), mientras el formulario hermano usa `Enabled` para los de envío. La diferencia impide que el usuario sepa por qué una acción no está disponible.
- **Deshacer un hito no pregunta qué va a pasar con el correo ya enviado:** la confirmación sólo menciona el fichero o la fecha (`:123, 427`).
- **`PintarIndicadores` se llama en cada registro y en cada deshacer**, lo que recalcula los contadores del menú y repinta el panel de tareas (`FU.bas:41-72`).
- **Los correos no se programan desde aquí:** los `ComandoEnvio*_Click` deshabilitan `Programado` en `FormEnvioCorreo`.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Fechas de la solicitud | `m_ObjSolicitudActiva.SolicitudFechas`, derivado de `TbSolicitudesFechas` | Derivado, sin DAO en este `.cls` | `:1035-1127` |
| Deshacer fase | `SolicitudOperaciones.EliminarFase` | DAO | `:23` |
| Deshacer Excel recibido | `SolicitudOperaciones.EliminarExcelDevuelto` | DAO | `:438` |
| Registro de hitos | handlers `ComandoRegistrar*` | DAO, no acotados en este `.cls` | `:200-317, 568, 664` |
| Previsualización de plantillas | `ObtenerHTMLEnLocal` | DAO y E/S de fichero | `:849-965`, `FU.bas:1336-1392` |
| Contadores del menú | `PintarIndicadores` → `Entorno.CalcularTotalesPendientes` | DAO | `:39`, `FU.bas:52` |

## Coexistencia

Pendiente de `../coexistencia.md`. El patrón es «registro + deshacer» por hito, con el estado de la interfaz calculado en cliente. En web, deshacer un hito es una operación destructiva que necesita su propia confirmación explícita y su auditoría, porque el correo ya puede haberse enviado; hoy la confirmación no menciona ese efecto. El cálculo de disponibilidad de cada acción debe pasar al servidor junto con las fechas.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-66.** `EstablecerDatos` sale **sin error** cuando no hay solicitud activa o no hay registro de fechas (`:1031-1038`), de modo que `Form_Load` da la carga por buena (`:976-979`) y el panel muestra los valores de diseño como si fueran los hitos reales de una solicitud.
- **H-HPS-S-67.** El mensaje de error de `ComandoEliminarExcelRecibido_Click` nombra otro procedimiento: «Al ComandoEliminarExcelEnviado_Click se ha producido el error n» (`:457`). Es la misma confusión que H-HPS-S-65 en el camino de registro, aquí en el camino de borrado, y en ambos casos el correo al administrador atribuye el fallo al handler equivocado.
- **H-HPS-S-68.** La rama de error de `m_FormAnexo_Adjuntado` **no informa al usuario**: sólo envía el correo al administrador (`:1015-1019`). Si falla el repintado tras adjuntar el Excel del solicitante, la pantalla queda con las fechas anteriores y ninguna explicación visible.
- **H-HPS-S-69.** El borrado de fases tiene **dos implementaciones**: cinco botones delegan en `Public Function Eliminar`, que va a `SolicitudOperaciones.EliminarFase` (`:11-46`, llamado en `:127, 359, 395, 475, 509`), y los otros dos (`ComandoEliminarEnvioTraspasoONS`, `ComandoEliminarExcelRecibido`) llaman a `SolicitudOperaciones` directamente (`:154, 438`). La invalidación de `SolicitudFechas`, el repintado y `PintarIndicadores` quedan duplicados en ocho sitios, y una corrección aplicada a `Eliminar` no alcanza a los otros dos botones.
- `ComandoEliminarCancelacion_Click` activa el reloj de arena (`:131`) y lo desactiva dos líneas después (`:138`) sin ninguna operación entre medias: código muerto.
- `Public Function Eliminar` **apaga el reloj de arena** (`:30`) aunque ningún llamador lo haya encendido antes de invocarla (`:127`).
- `EstablecerDatos` da el foco a `ComandoRegistrarExcelEnviado` (`:1039`) antes de resolver la visibilidad del resto de botones, de modo que el foco puede caer en un control que la cadena posterior deja inaccesible.
- Los treinta handlers repiten literalmente la estructura `On Error`/`Hourglass`/`AllowEdits`/`MsgBox`, cada uno con su propio texto; H-HPS-S-67 es una de las consecuencias de esa repetición.
- **Defectos de interfaz heredados (a verificar):** el `Caption` de diseño es «EXPEDIENTES Versión 2024-001 Administrador» (`.form.txt:39`), de otra aplicación; y el formulario no declara rótulo de estado propio, a diferencia de los restantes de la aplicación.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: historial de hitos con acciones de deshacer explícitas y auditadas; confirmación que advierta del efecto sobre el correo ya enviado; y el conjunto de acciones disponibles devuelto por la API en lugar de calcularse con `Visible` y `Enabled` en cliente.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sólo se ofrecen las acciones cuyo hito previo está registrado, con independencia del orden de pintado.
2. Sin solicitud activa o sin registro de fechas, el panel informa de que no hay nada que mostrar.
3. Registrar un hito con su precondición cumplida lo deja visible junto con su botón de deshacer.
4. Deshacer un hito pide confirmación, ejecuta una única ruta de código y repinta el panel y los contadores del menú.
5. Un fallo al repintar tras adjuntar informa al usuario y no deja las fechas con el valor anterior sin explicación.
6. Deshacer el Excel recibido no deja habilitado ningún botón cuyo hito ya no existe.
7. Un error en un camino de borrado se atribuye al procedimiento correcto.

## Pendientes

- Los diez botones `ComandoRegistrar*` delegan en `SolicitudOperaciones`; el detalle de cada registro queda fuera de esta ficha.
- Qué operación usa exactamente `ComandoEliminarEnvioTraspasoONS_Click` (`:154`), distinto de `EliminarFase` y de `EliminarExcelDevuelto`.
- `DiasParaCancelacion` (`:1355`): si es etiqueta o cuadro de texto no se deduce del `.form.txt`, y su valor no se usa en ningún handler del `.cls`.
- Recuento de controles: el inventario fija 56 y el `.form.txt` contiene 55 bloques `Begin <Tipo>`; la diferencia no se reproduce con una regla simple de recuento.