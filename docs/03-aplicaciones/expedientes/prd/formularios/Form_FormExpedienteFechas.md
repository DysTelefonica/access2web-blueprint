# Form_FormExpedienteFechas

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormExpedienteFechas.cls` (327 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Fechas» de [`Form_FormExpediente`](Form_FormExpediente.md): muestra y edita las doce fechas del ciclo de vida del expediente (licitación, oferta, adjudicación, contrato, recepción y garantía) y recalcula en pantalla la fecha de fin de garantía. Las citas son `Form_FormExpedienteFechas.cls:línea` salvo indicación.

## Entrada y salida

| Aspecto | Comportamiento | Evidencia |
|---|---|---|
| Origen de datos | Formulario independiente: el `.form.txt` no declara `RecordSource` ni `ControlSource`; los valores salen del DTO global `m_ObjExpedienteDTOActivo.Expediente` | `.cls:187-225` |
| Se abre desde | Botón de navegación `NavFechas` del contenedor, que vuelca antes la pestaña visible | `Form_FormExpediente.cls:418-440` |
| Abre | Nada; no lanza formularios, informes ni consultas | `.cls:1-327` |
| Salida | `Form_Unload` vuelca los controles al DTO con `RellenarDTO` | `.cls:238-257, 298-318` |
| Permisos | Edición solo si el contenedor se abrió para editar y `EsAdministrador = Sí` | `.cls:137-144, 174-186` |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `FECHAINICIOLICITACION`, `FECHAPREOFERTA`, `FECHAOFERTA`, `FECHADESESTIMADA`, `FECHAPERDIDA`, `FECHAADJUDICACION`, `FechaInicioContrato`, `FECHAFIRMACONTRATO` | TextBox `Short Date` | — | Se cargan desde el DTO y se vuelcan al DTO | `.cls:188-220, 305-315`; `.form.txt:401, 587, 617, 709, 865, 968, 1263, 1366` |
| `FechaFinContrato` | TextBox `Short Date` | `BeforeUpdate` | Recalcula `FechaFinGarantia` | `.cls:90-127`; `.form.txt:771-773` |
| `GARANTIAMESES` (rótulo «MESES GARANTÍA») | TextBox sin formato | `BeforeUpdate` | Recalcula `FechaFinGarantia` | `.cls:259-296`; `.form.txt:833-834` |
| `FECHACERTIFICACION` (rótulo «F.RECEPCIÓN») | TextBox `Short Date` | `BeforeUpdate` | Recalcula `FechaFinGarantia` | `.cls:51-88`; `.form.txt:997, 1028-1030` |
| `FechaFinGarantia` (rótulo «F.F.GARANTÍA») | TextBox bloqueado | — | Solo lectura; lo escribe el código | `.form.txt:1161, 1173` |
| `ComandoFECHA<X>` (siete botones: inicio licitación, preoferta, oferta, desestimada, perdida, adjudicación, firma contrato) | CommandButton, `Tag = PARAFECHA`, deshabilitado por defecto | `OnClick` | Escribe `Date` (hoy) en su campo | `.cls:10-45`; `.form.txt:420-446` y equivalentes |
| `ComandoFECHACERTIFICACION` | CommandButton `PARAFECHA` | `OnClick` | Escribe hoy e intenta recalcular la garantía | `.cls:14-22`; `.form.txt:1098-1114` |
| `lblEstadoExpediente` | Label | — | «ESTADO: » + `ESTADOCalculadoTitulo` en mayúsculas | `.cls:224`; `Expediente.cls:644-658` |

`FechaInicioContrato` y `FechaFinContrato` no tienen botón «Poner fecha actual» (`.form.txt` sin `ComandoFechaInicioContrato` ni `ComandoFechaFinContrato`).

## Flujos

### Apertura (`Form_Load`, `.cls:129-162`)

1. Parte con `AllowEdits = False`; lo activa solo si `Form_FormExpediente.m_AbiertoParaEditar` y `EsAdministrador = Sí` (`:137-144`).
2. `EstablecerDatos` exige `m_ObjExpedienteDTOActivo.Expediente`; si falta, «El formulario se ha abierto con parámetros insuficientes» (`:170-173`).
3. Habilita los botones `PARAFECHA` con la misma condición de permiso (`:174-186`).
4. Copia cada fecha del DTO al control solo si no está vacía (`:188-223`) y pinta el estado calculado (`:224`).

### Edición y cálculo de fin de garantía

- Al cambiar `FECHACERTIFICACION`, `FechaFinContrato` o `GARANTIAMESES`, el `BeforeUpdate` aplica la misma regla (`:58-76, 97-115, 266-284`); no cancela nunca la actualización.
- Los botones de fecha escriben hoy sin confirmación, aunque el campo ya tuviera valor (`:10-45`).

### Salida y guardado

- `Form_Unload` llama a `RellenarDTO` (`:245`); el contenedor también vuelca esta pestaña con su propio `RellenarDTO` (`Form_FormExpediente.cls:548-561`) al guardar (`Form_FormExpediente.cls:65`).
- El guardado lo hace `ExpedienteOperaciones.Registrar`; esta pestaña no escribe en base de datos.

### Errores

Los `BeforeUpdate` y `Form_Load` muestran `MsgBox` crítico y envían `CorreoAlAdministrador` salvo el error controlado 1000 (`:81-87, 155-161`). `EstablecerDatos` y `RellenarDTO` devuelven el error por `p_Error` (`:230-235, 324-327`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Fin de garantía = `GARANTIAMESES` meses sumados a `FECHACERTIFICACION` si es fecha; si no, a `FechaFinContrato`. Vacía si los meses no son numéricos o `FechaFinContrato` no es fecha | Pantalla y guardado | `.cls:58-76`; `Expediente.cls:474-494` |
| Al guardar, `FechaFinGarantia` se sustituye por `FechaFinGarantiaCalculada`: el valor de pantalla no se persiste tal cual | `Registrar` | `ExpedienteOperaciones.cls:69` |
| Coherencia de fechas (solo si `APLICAESTADO = "Sí"`): inicio de contrato sin fin, fin sin inicio, inicio sin meses de garantía, firma sin inicio/fin/meses, recepción sin inicio/fin/meses/firma → rechazo con el motivo de `SonFechasIrregulares`; con fecha de desestimación o pérdida no se valida nada | `MotivoNoOK` | `ExpedienteOperaciones.cls:845-850`; `FUNCIONES UTILES.bas:5615-5659` |
| Estado derivado de las fechas: Desestimado > Perdido > Oferta/Preoferta (sin adjudicación) > Adjudicada (sin fin de contrato) > En ejecución (hoy ≤ recepción o fin de contrato) > En garantía / Cerrado según `FechaFinGarantia` | `GetESTADOCalculado` | `FUNCIONES UTILES.bas:5486-5534` |
| No se valida el orden cronológico entre fechas ni el rango de `GARANTIAMESES` | — | `FUNCIONES UTILES.bas:5607-5659`; `.cls:58-76` |

## Lecturas y escrituras

Ninguna directa. Lee el DTO global (`.cls:187-224`) y lo escribe (`:304-317`). La persistencia en `TbExpedientes` (doce columnas, cadena vacía → `Null`) ocurre en `ExpedienteOperaciones.cls:294-363`, dentro de la transacción descrita en la ficha del contenedor.

## Coexistencia

Las fechas viajan como `String` en el DTO (`Expediente.cls:26, 60`) y se asignan a los campos del `Recordset` sin conversión explícita (`ExpedienteOperaciones.cls:294-363`): depende de la configuración regional de Access. En PostgreSQL vinculado hay que fijar tipo `date`, `GARANTIAMESES` entero y un cálculo de garantía idéntico en ambos lados. Véase `../coexistencia.md` (pendiente).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComandoFECHACERTIFICACION_Click` comprueba `IsNumeric(Me.FECHACERTIFICACION)` en vez de `IsDate` (`.cls:17`): con una fecha es previsible que no recalcule la garantía; además, asignar por código no dispara `BeforeUpdate`.
- `Registrar` calcula `ESTADO` (`ExpedienteOperaciones.cls:68`) antes de recalcular `FechaFinGarantia` (`:69`); `GetESTADOCalculado` usa el valor previo y hace `CDate(.FechaFinGarantia)` (`FUNCIONES UTILES.bas:5530`), que fallaría si está vacío.
- `GetESTADOCalculado` declara la etiqueta `errores` sin `On Error GoTo` (`FUNCIONES UTILES.bas:5459-5548`): un error se propaga sin pasar por `p_Error`.
- `DiasParaElFin` toma la mayor entre recepción y fin de contrato (`Expediente.cls:439-462`); garantía y estado toman la recepción sin comparar. Reglas divergentes.
- Los tres `BeforeUpdate` rotulan el error como «Al Meses_Exit» (`.cls:83, 122, 291`).
- `Form_Unload` ignora el error que devuelve `RellenarDTO` (`.cls:245`); `EstablecerDatos` no comprueba `Expediente.Error` tras leer `ESTADOCalculadoTitulo` (`:224`).
- `RellenarDTO` está duplicado en esta pestaña y en el contenedor (`.cls:298-318`; `Form_FormExpediente.cls:548-561`).
- `GARANTIAMESES` admite decimales o negativos (`IsNumeric`, `.cls:266`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: calcular fin de garantía y estado solo en servidor; validar orden cronológico; botón «hoy» también para inicio y fin de contrato.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin permiso de edición (solo lectura o no administrador) los campos y botones «hoy» quedan deshabilitados.
2. Con 24 meses y recepción informada, el fin de garantía es recepción + 24 meses; sin recepción, fin de contrato + 24 meses; sin fin de contrato o con meses no numéricos, vacío.
3. El fin de garantía guardado coincide siempre con el calculado, aunque se envíe otro valor.
4. Con `APLICAESTADO = "Sí"`, guardar firma de contrato sin fin de contrato devuelve el motivo de `SonFechasIrregulares` y no escribe.
5. El rótulo de estado refleja la tabla de precedencia de `GetESTADOCalculado`.

## Pendientes

- Confirmar en el binario si `Form_Unload` se dispara al cambiar de pestaña en el control de navegación (afecta a la conclusión de la ficha del contenedor sobre `NavHitos`).
- Contenido de `ColEstadosTituloMejorado` (`Entorno.cls`) y valores de `EnumEstados`.
- Comportamiento real de `IsNumeric` sobre el valor del control `FECHACERTIFICACION` y del redondeo de `DateAdd` con meses decimales.
