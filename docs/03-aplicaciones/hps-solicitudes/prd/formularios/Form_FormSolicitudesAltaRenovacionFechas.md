# Form_FormSolicitudesAltaRenovacionFechas

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormSolicitudesAltaRenovacionFechas.cls` (751 líneas) y `.form.txt` (62 bloques `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`, `constructor.bas` = `modules/constructor.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_FormSolicitudesAltaRenovacionFechas.cls:línea` salvo indicación.

Panel de fechas de la solicitud: la segunda pestaña del contenedor `FormSolicitudRenovacion`. Es un **cajón de acciones** más que un formulario de datos: no pinta un registro, sino que decide qué de los dieciocho hitos del trámite están disponibles según las fechas ya registradas, y lanza la acción correspondiente (enviar correo, abrir el HTML de la plantilla, registrar fecha). **Enlazado por control**, con los nombres de campo de `TbSolicitudesFechas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudRenovacion` | Se declara como segundo destino del control de navegación del contenedor | `Form_FormSolicitudRenovacion.form.txt:497` |
| `Form_FormSolicitudesAltaFechas` | Su gemelo de alta, con el mismo esqueleto de handlers | `Form_FormSolicitudesAltaFechas.cls` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `FormEnvioCorreo` | Los siete `ComandoEnvio*_Click` la abren con el `EnumTipoEnvioCorreo` que corresponda y deshabilitan su casilla `Programado` | `:21-27` y hermanos |
| `FormWeb` | Los `ComandoVer*_Click` generan el HTML de la plantilla y lo proyectan | `:445` y hermanos, `FU.bas:1336-1392` |
| `SolicitudOperaciones` | `RegistroAltaONS` y `RegistrarUsuarioHPSEnHPS` registran los dos hitos finales | `:263, 328` |
| `Form_FormSolicitudesGestion` | Se refresca si el estado de la solicitud ha cambiado | `:357-359` |
| Menú principal | `PintarIndicadores` recalcula los contadores del menú | `:280, 362` |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` → `Form_Load` | `EstablecerDatos` y reactiva la edición | `:581-609` |
| `FechaEnvioExcel`, `FechaRecepcionExcel`, `FechaTramitacionAltaMarga`, `FechaRegistroEnHPS`, `FechaDesestimado`, `FechaAutocancelacion`, `FechaEnvioONS`, `FechaCorreoRecordatorioExcel1/2`, `FechaCorreoRecordatorioRellenoMarga1/2`, `FechaPrevistaCorreoRecordatorioExcel1/2`, `FechaPrevistaCorreoRecordatorioRellenoMarga1/2`, `FechaPrevistaAutocancelacionMarga`, `FechaPrevistaAutocancelacionPreMarga` | controles enlazados a `SolicitudFechas` | — | Los dieciocho hitos que `EstablecerDatos` pinta si son fecha válida | `:628-670`, `.form.txt:612-2296` |
| `ComandoEnvioExcel`, `ComandoEnvioRecordatorioExcel1/2`, `ComandoEnvioRecordatorioMarga1/2`, `ComandoEnvioCancelacionMarga`, `ComandoEnvioCancelacionPreMarga` | botones | `OnClick` | Abre `FormEnvioCorreo` con el tipo correspondiente | `:8-239`, `.form.txt:641-2099` |
| `ComandoVerHTMLEnvioExcel`, `ComandoVerRecordatorioExcel1/2`, `ComandoVerRecordatorioMarga1/2`, `ComandoVerCancelacionMarga`, `ComandoVerCancelacionPreMarga` | botones | `OnClick` | Proyecta el HTML de la plantilla | `:378-580`, `.form.txt:1651-2135` |
| `ComandoRegistroEnvioONS` | botón | `OnClick` → `ComandoRegistroEnvioONS_Click` | Pide la fecha por `InputBox` y registra el alta en ONS | `:240-294`, `.form.txt:2356` |
| `ComandoRegistroHPS` | botón | `OnClick` → `ComandoRegistroHPS_Click` | Crea o actualiza el usuario HPS y termina la tramitación | `:296-376`, `.form.txt:2255` |
| `ComandoAdjuntarExcelRecibido`, `ComandoDesestimar`, `ComandoAltaMarga` | botones | declarados en diseño | Hitos cuyo handler no está en este `.cls` | `.form.txt:710, 2183, 2219` |

## Flujos

### Normal — carga y habilitación

1. `Form_Load` llama a `EstablecerDatos` y propaga el error (`:591-594`).
2. `EstablecerDatos` (`:612-751`) sale sin error si no hay solicitud activa o si `SolicitudFechas` no existe (`:618-626`), y a continuación copia al formulario cada hito que sea fecha válida (`:628-670`).
3. La segunda mitad del método decide qué botones quedan operativos según la rama en la que esté la solicitud (`:680-738`): si la tramitación ha terminado, desactiva los once botones de correo y aviso (`:700-711`); si no, habilita «Registrar en HPS» y decide los botones de recordatorio a partir de `FechaRecepcionExcel` (`:715-721`) y de `FechaEnvioExcel` (`:724-737`).

### Normal — envío y previsualización

1. Cada `ComandoEnvio*_Click` cierra `FormEnvioCorreo` si ya estuviera abierta, la abre con su `EnumTipoEnvioCorreo` y desactiva su casilla `Programado`, de modo que el envío no se programa desde la pantalla de hitos (`:21-27`).
2. Cada `ComandoVer*_Click` llama a `ObtenerHTMLEnLocal` con la plantilla de `m_ObjEntorno` correspondiente y la solicitud activa (`:445` y hermanos), que genera el fichero y abre `FormWeb` (`FU.bas:1381-1385`).

### Normal — cierre del trámite

1. `ComandoRegistroEnvioONS_Click` exige `FechaTramitacionAltaMarga` (`:248-251`), pide la fecha con `InputBox` (`:252`), llama a `RegistroAltaONS` (`:263`), invalida `SolicitudFechas` (`:273`) y repinta (`:276-280`).
2. `ComandoRegistroHPS_Click` exige `FechaEnvioONS` (`:309-312`), consulta `MotivoNoOKRegistrarUsuarioHPSEnHPS` (`:316`), registra en HPS y avisa según devolviera `"0"` (usuario nuevo) o `"1"` (usuario existente) (`:328-341`), recarga la solicitud (`:347`), repinta (`:351`) y refresca la rejilla principal si el estado cambió (`:356-360`).

### Alternativos

- **Sin solicitud activa o sin `SolicitudFechas`:** `EstablecerDatos` sale sin error (`:618-626`) y el formulario muestra los valores de diseño.
- **Cancelación del `InputBox`:** cadena vacía y `Exit Sub` sin escribir (`:252-255`).
- **Resultado inesperado de `RegistrarUsuarioHPSEnHPS`:** si no es `"0"` ni `"1"`, no se avisa nada (`:332-341`).

### Errores

- Patrón uniforme de `On Error GoTo errores` con el centinela `1000`, correo al administrador (`FU.bas:585`) y `MsgBox` crítico, repetido en los dieciocho handlers.
- Los handlers de previsualización y de envío **no cierran el formulario** en su rama de error, a diferencia de los dos de registro, que tampoco lo hacen pero sí dejan la pantalla repintable.

## Reglas, validaciones y cálculos

- **Precondiciones de los hitos finales:** «Registrar en HPS» exige `FechaEnvioONS` (`:309`) y «Registrar envío ONS» exige `FechaTramitacionAltaMarga` (`:248`); ambas se validan contra el control ya pintado, no contra la entidad.
- **La habilitación de botones depende del camino ejecutado**, no sólo del estado: ver H-HPS-S-64.
- **Los correos de esta pantalla no se programan:** los siete `ComandoEnvio*_Click` desactivan `Programado` en `FormEnvioCorreo` (`:27`), de modo que la automatización se gestiona desde la pantalla de envío, no desde la de hitos.
- **`EstablecerDatos` no propaga la ausencia de datos:** es la única diferencia entre «no hay nada que pintar» y «error».

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Fechas de la solicitud | `m_ObjSolicitudActiva.SolicitudFechas`, derivado de `TbSolicitudesFechas` | Derivado, sin DAO en este `.cls` | `:623-670` |
| Recarga tras el registro ONS | `constructor.getSolicitud` | DAO | `:347` |
| Registro del alta en ONS | `SolicitudOperaciones.RegistrarAltaONS` | DAO | `:263` |
| Registro en HPS | `MotivoNoOKRegistrarUsuarioHPSEnHPS` y `RegistrarUsuarioHPSEnHPS` | DAO | `:316, 328` |
| Previsualización de plantillas | `ObtenerHTMLEnLocal` → `GenerarHTML` | DAO y E/S de fichero | `:445` y hermanos, `FU.bas:1179, 1381` |
| Envío de correo | `FormEnvioCorreo`, fuera de esta ficha | Fuera de alcance | `:21-27` |

## Coexistencia

Pendiente de `../coexistencia.md`. La pestaña de fechas es, en la práctica, una **máquina de estados por hito**: cada fecha habilita un conjunto de acciones y el registro de cada hito se hace con un parámetro libre. En web el equivalente es un historial de hitos consultable, con el conjunto de acciones disponibles derivado de las fechas ya registradas y cada registro de hito devolviendo su propio resultado. La fecha de registro ONS, que hoy se teclea en un `InputBox`, pasa a ser un campo del formulario con validación de servidor.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-62.** `ComandoRegistroEnvioONS_Click` y `ComandoRegistroHPS_Click` leen `m_ObjSolicitudActiva.SolicitudFechas` sin comprobar que el objeto exista (`:248, 309`), mientras que `EstablecerDatos` sí lo comprueba y sale con calma (`:624-626`). Con una solicitud sin registro de fechas, pulsar cualquiera de los dos botones termina en error no controlado en lugar del aviso de precondición.
- **H-HPS-S-63.** La fecha del alta en ONS se pide con `InputBox` y se pasa tal cual a `RegistroAltaONS` (`:252, 263`), sin comprobar que sea una fecha ni un formato aceptado: un texto cualquiera llega a la capa de datos, y el formato válido sólo se define allí.
- **H-HPS-S-64.** La rama `If IsDate(Me.FechaEnvioExcel) Then` de `EstablecerDatos` tiene el cuerpo **vacío** (`:726-728`) y su `Else` es el que desactiva los cuatro botones de recordatorio (`:729-733`). El estado de esos botones queda por tanto determinado por qué rama se ejecutó antes y no por las fechas: con `FechaEnvioExcel` ya registrada, los botones conservan el valor que les dejó la rama exterior, que puede ser habilitado aunque el resto de hitos no lo estén. Es el mismo patrón ya descrito en el formulario gemelo (`H-HPS-S-20`, `Form_FormSolicitudesAltaFechas.cls:1592-1596`).
- **H-HPS-S-65.** El mensaje de error de `ComandoRegistroEnvioONS_Click` nombra otro procedimiento: «Al ComandoRegistrarAltaMarga_Click se ha producido el error n» (`:288`), aunque el handler es `ComandoRegistroEnvioONS_Click`. El error llega al correo del administrador atribuído al handler equivocado.
- `EstablecerDatos` sale sin comunicar nada cuando no hay solicitud activa ni registro de fechas (`:618-626`), de modo que `Form_Load` da la carga por buena (`:591-594`) y el formulario muestra los valores de diseño como si fueran los reales.
- Si `RegistrarUsuarioHPSEnHPS` devuelve un valor distinto de `"0"` o `"1"`, no se muestra ningún mensaje (`:332-341`): el registro ha ocurrido y el usuario no se entera.
- `ComandoEnvioExcel_Click` cierra `FormEnvioCorreo` **sin** `acSaveNo` (`:21-23`), a diferencia de los patrones que sí lo usan en otros formularios del legacy.
- Los dieciocho handlers son copias con dos o tres cambios: la misma estructura `On Error`/`Hourglass`/`AllowEdits`/`MsgBox` se repite sin abreviar, con la consecuencia de que cada uno lleva su propio literal de mensaje y cada literal puede quedar desincronizado (H-HPS-S-65 es un ejemplo).
- **Defectos de interfaz heredados (a verificar):** el `Caption` de diseño es «EXPEDIENTES Versión 2024-001 Administrador» (`.form.txt:39`); los rótulos «RECORDATORIO MARGA 1/2» y «RECORDATORIO EXCEL 1/2» identifican la etapa, mientras los botones correspondientes dicen ambos «Envío Recordatorio 1» (`.form.txt:811, 912, 1077, 1178`); y `ComandoAdjuntarExcelRecibido` lleva el typo «EXCEL DE SOLCITUD» (`.form.txt:711`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: historial de hitos consultable en lugar de una rejilla de botones habilitados por código; campo de fecha validado en el servidor en lugar de `InputBox`; y el conjunto de acciones disponibles devuelto por la API junto con las fechas, en vez de calcularse en cliente.

## Criterios de aceptación

Propuestos, no aplicados:

1. Sólo se ofrecen las acciones cuyo hito previo está registrado, y esa condición no depende del orden en que se pintó la pantalla.
2. Sin solicitud activa o sin registro de fechas, la pantalla informa de que no hay nada que mostrar.
3. «Registrar envío ONS» exige el alta en Marga y una fecha válida; si falta la precondición o la fecha, no se escribe y se explica el motivo.
4. «Registrar en HPS» exige el envío ONS registrado y muestra siempre el resultado, incluidos los valores distintos de usuario nuevo o existente.
5. Abrir la pantalla de hitos no programa correos por su cuenta.
6. Un error en el registro de un hito no se atribuye a otro procedimiento y no invalida lo ya registrado.

## Pendientes

- `ComandoAdjuntarExcelRecibido`, `ComandoDesestimar` y `ComandoAltaMarga` existen en el diseño (`.form.txt:710, 2183, 2219`) pero no tienen handler en este `.cls`: su comportamiento está en otro formulario o no está implementado.
- Formato de fecha que acepta `RegistroAltaONS`: el `InputBox` no lo restringe y el tipo de columna en `TbSolicitudesFechas` (`SHORT_DATE_TIME`) no lo declara.
- Valor de retorno de `RegistrarUsuarioHPSEnHPS` fuera de `"0"` y `"1"`: no está documentado en este `.cls`.
- Recuento de controles: el inventario fija 63 y el `.form.txt` contiene 62 bloques `Begin <Tipo>`; la diferencia no se reproduce con una regla simple de recuento.