# Form_frmDetalleExpediente

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frmDetalleExpediente.cls` (88 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ficha de solo consulta de un expediente: título, códigos, responsables, RAC, contratistas, subcontratistas y fechas de contrato y garantía. Es un formulario **no enlazado** que no permite edición, altas ni bajas (`.form.txt:13-17`); los valores se vuelcan desde un `ExpedienteViewModel` (`:10, 56-78`; `classes/ExpedienteViewModel.cls:20-25`). Las citas son `Form_frmDetalleExpediente.cls:línea` salvo indicación; `ExpSrv.cls` abrevia `classes/ExpedienteServicio.cls`, `ExpRepo.bas`, `modules/ExpedienteRepositorio.bas`, y `FU.bas`, `modules/FUNCIONES UTILES.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_frmAltaSolicitud | Botón de ver expediente, si hay uno seleccionado; diálogo con el identificador en `OpenArgs` | `Form_frmAltaSolicitud.cls:339-343` |
| Form_frmBuscarExpediente | Expediente seleccionado en la lista; diálogo con el identificador en `OpenArgs` | `Form_frmBuscarExpediente.cls:218-228` |

No abre otros formularios. Al salir deja `Tag = "CLOSED"` y se cierra (`:80-84`).

## Controles y eventos

| Control | Tipo | Evento y handler | Origen del valor | Evidencia |
|---|---|---|---|---|
| `txtTitulo` | texto bloqueado | — | `Expediente.titulo` | `:59`, `.form.txt:683-701` |
| `txtNemotec` | texto | — | `Expediente.Nemotecnico` | `:60`, `.form.txt:598-614` |
| `txtCodigoActividad`, `txtCodExp`, `txtCodExpLargo` | texto | — | `CodigoActividad`, `CodExp`, `CodExpLargo` | `:61-63` |
| `txtContratistasPrinciaples`, `txtSubContratistas` | texto | — | Nombres concatenados con « \| », o «N/A» | `:64-65`; `ExpSrv.cls:189-211` |
| `txtRCalidad`, `txtRTecnico`, `txtRacAsignado` | texto | — | Responsable de calidad, jefe de proyecto y RAC | `:66-68` |
| `txtFInicioContrato` | texto bloqueado | `AfterUpdate` declarado, sin handler en el `.cls` | Fecha si es mayor que cero | `:73`, `.form.txt:629-648` |
| `txtFFinContrato`, `txtFFinGarantia` | texto bloqueado | — | Fecha si es mayor que cero | `:74-75`, `.form.txt:757-775, 810-828` |
| `cmdSalir` | botón | `OnClick` | Cierra el formulario | `:80-88`, `.form.txt:1308-1310` |

Nueve cuadros de texto llevan `Tag = "OBLIGATORIO"` (`.form.txt:616, 882, 935, 988, 1041, 1094, 1147, 1200, 1253`); el `.cls` no lo lee.

## Flujos

**Normal:**

1. `Form_Load` ajusta el formulario y lee el identificador de `OpenArgs` (`:22-23`).
2. `ExpedienteServicio.getExpedienteViewModelPorID` valida que el identificador sea numérico y carga el expediente (`:29`; `ExpSrv.cls:252-264, 89-96`).
3. El servicio añade los nombres de contratistas principales y subcontratistas (`ExpSrv.cls:272-274`): por cada relación del expediente obtiene el suministrador y lo clasifica según `ContratistaPrincipal = "Sí"` y `SubContratista = "Sí"` (`ExpSrv.cls:180-206`).
4. `PoblarFormularioDesdeViewModel` vuelca los valores en los controles (`:56-78`).

**Alternativos:** una fecha nula o cero deja el control vacío (`:70-75`); el expediente se modela con fechas de tipo `Date` (`classes/Expediente.cls:22-24`).

**Errores:**

- Sin `OpenArgs`: «No se ha proporcionado un expediente para mostrar.» (`:24-26`).
- Expediente no encontrado: «El expediente con ID … no ha sido encontrado.» (`:31-32`).
- Identificador no numérico: «Se requiere un ID de Expediente válido.» desde el servicio (`ExpSrv.cls:92-94`).
- En todos los casos muestra el mensaje con `ManejadorErroresFormulario` y cierra el formulario (`:45-48`; `FU.bas:588-677`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Jefe de proyecto mostrado: `FIRST` de los responsables con `EsJefeProyecto = 'Sí'` | subconsulta | `ExpRepo.bas:20-22` |
| RAC mostrado: `FIRST` de los RAC vinculados al expediente | subconsulta | `ExpRepo.bas:23-25, 33` |
| Varios contratistas o subcontratistas se separan con « \| »; sin ninguno, «N/A» | servicio | `ExpSrv.cls:193, 202, 210-211` |
| Un suministrador puede figurar a la vez como principal y como subcontratista | servicio | `ExpSrv.cls:189, 198` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Expediente con responsables y RAC | `TbExpedientes`, `TbUsuariosAplicaciones`, `TbExpedientesResponsables`, `TbExpedientesRACS`, `TbRACS` | SQL dinámico con parámetro `[p_ID]` | no | `ExpRepo.bas:18-41` |
| Relaciones con suministradores | `TbExpedientesSuministradores` | SQL dinámico, identificador concatenado | no | `ExpRepo.bas:58-69` |
| Suministrador por relación | `TbSuministradores` | una consulta por relación | no | `ExpSrv.cls:185` |

Solo lectura, sobre la conexión de expedientes (`ExpRepo.bas:40, 68`; `FU.bas:156-178`).

## Coexistencia

Pendiente del resumen de coexistencia (`../coexistencia.md`). Ya se ve: `FIRST()` sin orden determinista (`ExpRepo.bas:20, 23`); comparación de `'Sí'` en SQL y en VBA, sensible a la intercalación de destino (`ExpRepo.bas:22`; `ExpSrv.cls:189, 198`); fechas nulas tratadas como cero (`:73-75`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Con varios jefes de proyecto o varios RAC, `FIRST()` muestra uno arbitrario sin indicar que hay más (`ExpRepo.bas:20, 23`).
- El manejador de errores salta con `GoTo` a `LimpiarYSalir` sin `Resume` (`:45-48`), con el error aún activo; el `DoCmd.Close` dentro de `Form_Load` (`:47`) puede no cerrar según el evento en curso.
- El nombre `txtContratistasPrinciaples` contiene una errata (`:64`); afecta a cualquier referencia futura por nombre.
- `AfterUpdate` de `txtFInicioContrato` está declarado sin procedimiento (`.form.txt:648`); el control está bloqueado, por lo que no se dispara.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: listar todos los jefes de proyecto y RAC del expediente en lugar de uno.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con un identificador válido, la ficha muestra los trece campos de la tabla de controles sin permitir edición.
2. Sin identificador, con identificador no numérico o inexistente, se muestra el mensaje correspondiente y no se presenta la ficha.
3. Fechas vacías se muestran vacías, no como fecha cero.
4. Contratistas y subcontratistas múltiples aparecen todos; sin ninguno, «N/A».

## Pendientes

- Criterio de elección de jefe de proyecto y RAC cuando hay varios.
- Contenido de `getSuministradorPorID` y de la clase `ExpedienteSuministrador` (`classes/SuministradorServicio.cls:44`).
