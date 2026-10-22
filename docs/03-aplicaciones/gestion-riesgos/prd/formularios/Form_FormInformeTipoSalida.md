# Form_FormInformeTipoSalida

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormInformeTipoSalida.cls` (102 líneas) y `src/forms/FormInformeTipoSalida.form.txt` · **PRD:** [README](../README.md)

Selector del formato de salida del informe de una edición: Excel u HTML. Genera el informe y lo abre con el visor del sistema. No guarda datos de negocio, pero **escribe un fichero en disco** a través de `GenerarInformePublicacion`. Es uno de los dos formularios **añadidos** respecto al snapshot (`main` no lo tiene en staging). Las citas son `Form_FormInformeTipoSalida.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| **Nadie en `src`** | Ningún `DoCmd.OpenForm` de `src` abre `FormInformeTipoSalida` | búsqueda por nombre en `src`, 0 resultados fuera de sus propios ficheros |

El formulario espera un identificador de edición por `OpenArgs` (`:30`, `:70-71`), así que suене haber un llamador **en el binario que no está exportado** (macro o código de otro formulario). Queda pendiente de confirmar contra el binario; no se afirma aquí que esté huérfano.

| Hacia | Cómo | Evidencia |
|---|---|---|
| Navegador del sistema | `Ejecutar Me.hWnd, "open", url` | `:85` |
| Fichero en disco | `GenerarInformePublicacion` genera el artefacto y devuelve su URL | `:81`; `Funciones Generales.bas:1409` |
| Barra de progreso del splash | `AvanceCerrar` | `:87, 94`; `Funciones Generales.bas:4920` |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Ajusta tamaño y título, bloquea edición, carga el combo y lee la edición de `OpenArgs` | `:17-44` |
| (formulario) | — | `ComandoGenerarInforme_Click` | Genera el informe y abre el resultado | `:46-97` |
| `cmbTipoSalida` | combo | — | Opciones `Excel` y `HTML`; por defecto `Excel` | `:24-28` |
| `lblEstado` | etiqueta | — | «Generando informe...», oculta hasta pulsar | `:29, 55-56` |
| `ComandoGenerarInforme` | botón | `OnClick` | `:46` |
| `cmdCancelar` | botón | `OnClick` | Cierra sin generar | `:99-102` |

## Flujos

### Apertura (`Form_Open`, `:17-44`)

1. `AjustarTamaño Me` y copia `lblTitulo1.Caption` al título del formulario (`:21-22`).
2. `AllowEdits = False` mientras monta el combo (`:23`).
3. Vacía `RowSource` y añade `Excel` y `HTML`; selecciona `Excel` (`:24-28`).
4. Oculta `lblEstado` (`:29`).
5. Si `OpenArgs` trae identificador, carga la edición con `Constructor.getEdicion` (`:30-36`).
6. `AllowEdits = True` (`:37`).

### Generación (`ComandoGenerarInforme_Click`, `:46-97`)

1. Reloj de arena y texto «Generando informe...» (`:52-56`).
2. Lee el combo en mayúsculas; si está vacío, **asume HTML** (`:57-60`).
3. Traduce a `EnumTipoInformePublicacion`: `Excel` o `html`; cualquier otro valor cae a `html` (`:61-68`).
4. Si no tiene edición, reintenta leerla de `OpenArgs` (`:69-76`); si tampoco, error «No se ha indicado la edición para el informe» (`:77-80`).
5. `GenerarInformePublicacion(m_Edicion, TipoEnum, , "", m_Error)` (`:81`).
6. Abre la URL con `Ejecutar` (`:85`) y se cierra (`:91`).

### Errores (`:39-44`, `:93-96`)

Ambos manejadores llaman a `AvanceCerrar`, quitan el reloj de arena y muestran `m_Error` en un `MsgBox`. El de apertura **cierra además el formulario** (`:43`); el de generación **no**, y deja el formulario abierto para reintentar.

## Reglas, validaciones y cálculos

- **Dos valores, dos defaults distintos.** Si el combo está vacío, el valor por defecto es HTML (`:59`), pero el valor inicial del combo es Excel (`:28`). El fallback y el default no coinciden.
- **`AllowEdits` se manipulate dos veces** (`:23, 37`) para que el combo no parpadee al abrirse; el efecto final es editable.
- **`Err.Raise 1000` sin descripción** en `:34, 73, 79, 83`. Funciona solo porque `m_Error` ya está informado en esos cuatro puntos; si alguien añade una llamada nueva con el mismo patrón y sin informar `m_Error`, el `MsgBox` saldría vacío.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Edición por identificador | `Constructor.getEdicion` | `:32, 71` |
| Generación del informe | fichero en disco, ruta resuelta en ejecución | `:81`; `Funciones Generales.bas:1409` |

Sin escrituras en tablas.

## Coexistencia

Pendiente de `../coexistencia.md`. Aquí ya se ve: el resultado es un **fichero en disco** cuya URL se abre con `ShellExecuteA` (`Variables Globales.bas:15, 42`), no una respuesta HTTP. La web tendría que decidir si el informe es una descarga, una vista previa en el navegador o un fichero generado bajo demanda, y el formato Excel no tiene equivalente directo en ese modelo.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-1.** La línea `Me.cmbTipoSalida.AddItem "Excel"` está **comentada** con la nota «Deshabilitado por SPEC-052» (`:25`) e inmediatamente después hay otra activa idéntica (`:26`). El comentario sugiere que alguien quiso quitar la opción y no lo hizo: el comentario dice una cosa y el código hace otra.
- **H-GR-2.** El valor por defecto del combo es `Excel` (`:28`) pero el fallback ante combo vacío es `html` (`:59`), y el tercer fallback también es `html` (`:67`). Dos defaults distintos para la misma decisión.
- **H-GR-3.** `pregunta` se usa como variable desechable en los dos manejadores de error (`:42, 96`) sin declararla, con `Option Explicit` activo (`:12`). Compila porque `pregunta` es una **global pública** (`Variables Globales.bas:69`): dos formularios que la usen a la vez se pisan el valor. No afecta al resultado, pero es estado global compartido usado como variable local.
- **H-GR-4.** `Ejecutar` es `ShellExecuteA` declarado en `Variables Globales.bas:15, 42`, y la URL la genera la aplicación. Si la ruta se construye con datos del usuario, hay superficie de inyección; conviene verificar el origen de `url` antes de dar por cerrado el punto. **A verificar.**
- **H-GR-5.** El formulario queda abierto tras un error de generación (`:93-96` no cierra) pero **el reloj de arena sí se quita**; correcto. En cambio, si `AvanceCerrar` fallara, el usuario se queda con el cursor de espera. No está protegido por manejador.
- **H-GR-6.** `AvanceCerrar` (`Funciones Generales.bas:4920`) manipula el progreso del splash, que en este formulario no existe: si `frmSplash` ya se cerró (`:92` del propio splash), la llamada apunta a un formulario inexistente. Conviene comprobar que no lanza.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: elegir el formato en la propia pantalla de informe en lugar de en un modal; devolver el informe como descarga autenticada en lugar de abrir un fichero local con el visor del sistema; eliminar el concepto de «tipo de salida» si la web estandariza en un formato.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir el formulario sin edición devuelve un aviso claro y no genera informe.
2. Con edición válida, generar en Excel produce un fichero y en HTML una página, y ambos se abren.
3. Si el combo llega vacío, se usa el mismo valor por defecto que el inicial.
4. Un fallo al generar deja el formulario utilizable para reintentar.

## Pendientes

- Quién abre el formulario: no hay llamador en `src` (macro o código no exportado).
- Dónde escribe `GenerarInformePublicacion` y cómo compone la ruta (`:81`; `Funciones Generales.bas:1409`).
- Qué significa SPEC-052 y por qué la opción sigue activa (`:25-26`).
- Si `AvanceCerrar` es seguro cuando el splash ya está cerrado.