# Form_frmDetalleLogCambio

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frmDetalleLogCambio.cls` (58 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ficha de solo lectura de un registro de `tbLogCambios`. No tiene `RecordSource`; los cuadros de texto declaran `ControlSource` (`.form.txt:531-919`), pero el código los rellena a mano. Las citas son `Form_frmDetalleLogCambio.cls:línea` salvo indicación; `FU.bas` abrevia `modules/FUNCIONES UTILES.bas` y `RC.bas`, `modules/RepositorioComun.bas`.

## Entrada y salida

- **Se abre desde:** Form_frmVerLogCambios, con el identificador en `OpenArgs` (`Form_frmVerLogCambios.cls:118-128`).
- **Abre:** nada. **Permisos:** sin comprobación de rol propia.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` | Carga el registro por `OpenArgs` y rellena los campos | `:9-47`, `.form.txt:214` |
| `txtFechaHora`, `txtusuario`, `txttabla`, `txtRegistroID`, `txtcampo`, `txtvalorAnterior`, `txtvalorNuevo`, `txtOperacion` | texto | — | Reciben `fechaHora`, `usuario`, `tabla`, `registro`, `campo`, `valorAnterior`, `valorNuevo`, `tipoOperacion` | `:29-36` |
| `txtsuplantadoPor` | texto | — | Etiqueta «Suplantado Por»; el código no lo rellena | `.form.txt:568-569, 684` |
| `lblBloqueo` | etiqueta oculta | — | Sin uso en este formulario | `.form.txt:959-971` |
| `cmdCerrar` («Salir») | botón | `OnClick` | Cierra el formulario | `:50-58`, `.form.txt:999-1001` |

## Flujos

1. **Normal:** `Ajustar Me` y título «Detalle del Registro de Cambio» (`:16-17`); `idLog = CLng(Nz(OpenArgs, 0))` (`:18`); `LogServicio.getLogCambioPorID` (`:24`; `LogServicio.cls:135-148`) delega en `LogCambioRepositorio.getLogCambioPorID` (`LogCambioRepositorio.bas:55-65`); vuelca las propiedades (`:29-36`).
2. **Alternativo:** si el formulario ya está abierto con el mismo identificador, el visor solo le da el foco; con otro, lo cierra y lo reabre (`Form_frmVerLogCambios.cls:118-125`).
3. **Errores:** `OpenArgs` vacío o ≤ 0 ⇒ error 513 «No se ha especificado un registro de log para mostrar.» (`:20-22`); registro inexistente (`RC.bas:128-129`) ⇒ «No se pudo cargar el registro del log con ID …» (`:25-27`). En ambos casos `ManejadorErroresFormulario` muestra el mensaje y registra el error (`FU.bas:588-676`) y el formulario se cierra (`:43-46`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Evidencia |
|---|---|---|---|
| Registro por identificador | `tbLogCambios` | DAO `QueryDef` con parámetro `[p_ID]`, snapshot | `LogCambioRepositorio.bas:61-65`, `RC.bas:52-60` |
| Registro del error de interfaz | `tbLogErrores` | vía `LogErrorUI` | `FU.bas:672` |

## Coexistencia

Consulta parametrizada sin sintaxis específica de Jet (`LogCambioRepositorio.bas:61`); debe preservarse con `tbLogCambios` vinculada a PostgreSQL. Detalle común en `../coexistencia.md` (pendiente).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** el descrito en Flujos.

**Posibles defectos (a verificar):**

- Asignar valor a controles con `ControlSource` en un formulario sin `RecordSource` (`:29-36`; `.form.txt:531`) podría fallar o mostrar `#Nombre?`.
- `txtsuplantadoPor` nunca se rellena aunque `LogCambio` guarda `suplantadoPor` (`LogServicio.cls:27`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: vista de detalle dentro del visor de cambios, mostrando también el usuario suplantador.

## Criterios de aceptación (propuestos, no aprobados)

1. Con un identificador válido se muestran los ocho campos del registro sin permitir edición.
2. Sin identificador o con uno inexistente se muestra el mensaje correspondiente y no se abre la vista.

## Pendientes

- Confirmar en Access el efecto de `ControlSource` sin `RecordSource`.
