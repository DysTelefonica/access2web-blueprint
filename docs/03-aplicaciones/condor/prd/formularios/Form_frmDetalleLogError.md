# Form_frmDetalleLogError

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frmDetalleLogError.cls` (58 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ficha de solo lectura de un registro de `tbLogErrores`. Misma estructura que Form_frmDetalleLogCambio: sin `RecordSource`, controles con `ControlSource` rellenados por código. Las citas son `Form_frmDetalleLogError.cls:línea` salvo indicación; `FU.bas` abrevia `modules/FUNCIONES UTILES.bas` y `RC.bas`, `modules/RepositorioComun.bas`.

## Entrada y salida

- **Se abre desde:** Form_frmVerLogErrores, con el identificador en `OpenArgs` (`Form_frmVerLogErrores.cls:117-126`).
- **Abre:** nada. **Permisos:** sin comprobación de rol propia.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` | Carga el registro por `OpenArgs` y rellena los campos | `:9-47`, `.form.txt:210` |
| `txtFechaHora`, `txtusuario`, `txtsuplantadoPor`, `txtprocedimiento`, `txtnumeroError`, `txtdescripcionError`, `txtcontexto` | texto | — | Reciben las propiedades homónimas de `LogError` | `:29-36` |
| `txtModulo` | texto | — | El código asigna `modulo`; el `ControlSource` declarado es `procedimiento` | `:32`, `.form.txt:863-864` |
| `cmdCerrar` («Salir») | botón | ninguno declarado | El handler existe, pero el `.form.txt` no enlaza `OnClick` | `:50-58`, `.form.txt:938-947` |

## Flujos

1. **Normal:** `Ajustar Me` y título «Detalle del Registro de Error» (`:16-17`); `idLog = CLng(Nz(OpenArgs, 0))` (`:18`); `LogServicio.getLogErrorPorID` (`:24`; `LogServicio.cls:150-163`) delega en `LogErrorRepositorio.getLogErrorPorID` (`LogErrorRepositorio.bas:65-78`); vuelca las propiedades (`:29-36`).
2. **Alternativo:** reapertura con foco o cierre y reapertura según el identificador (`Form_frmVerLogErrores.cls:117-124`).
3. **Errores:** sin identificador ⇒ «No se ha especificado un registro de error para mostrar.» (`:20-22`); inexistente (`RC.bas:128-129`) ⇒ «No se pudo cargar el registro del log de error con ID …» (`:25-27`). `ManejadorErroresFormulario` avisa y registra (`FU.bas:588-676`); el formulario se cierra (`:43-46`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Evidencia |
|---|---|---|---|
| Registro por identificador | `tbLogErrores` | DAO `QueryDef` con `[p_ID]`, snapshot | `LogErrorRepositorio.bas:73-77`, `RC.bas:52-60` |
| Registro del error de interfaz | `tbLogErrores` | vía `LogErrorUI` | `FU.bas:672` |

## Coexistencia

Consulta parametrizada sin sintaxis específica de Jet (`LogErrorRepositorio.bas:73`). Detalle común en `../coexistencia.md` (pendiente).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** el descrito en Flujos.

**Posibles defectos (a verificar):**

- El botón «Salir» no ejecuta su handler por falta de `OnClick` (`.form.txt:938-947`); el cierre dependería del botón de ventana.
- `txtModulo` enlazado a `procedimiento` (`.form.txt:864`) mientras el código le asigna `modulo` (`:32`).
- Mismo riesgo que en el detalle de cambios al asignar controles con `ControlSource` sin `RecordSource`.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: detalle con descripción y contexto completos en texto multilínea.

## Criterios de aceptación (propuestos, no aprobados)

1. Con un identificador válido se muestran fecha, usuario, suplantador, módulo, procedimiento, número, descripción y contexto.
2. El módulo y el procedimiento se muestran en campos distintos con su propio valor.
3. «Salir» cierra la vista.

## Pendientes

- Confirmar en Access si «Salir» responde al clic.
