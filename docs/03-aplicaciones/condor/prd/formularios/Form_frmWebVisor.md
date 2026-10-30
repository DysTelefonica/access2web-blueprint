# Form_frmWebVisor

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frmWebVisor.cls` (426 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Contenedor emergente y modal (`.form.txt:7-8`) con un control Edge (`webVisor`) que muestra HTML generado por `WorkflowServicio` y devuelve a Form_frmGestionSolicitud lo que el usuario captura mediante enlaces `vba:`. No enlazado. Las citas son `Form_frmWebVisor.cls:línea` salvo indicación; `FGS.cls` abrevia `forms/Form_frmGestionSolicitud.cls`, `WS.cls`, `classes/WorkflowServicio.cls`, y `WVC.cls`, `classes/WebVisorCacheServicio.cls`.

## Entrada y salida

| Desde FGS.cls | Título | Caché | Evidencia |
|---|---|---|---|
| Motivo de rechazo de calidad | «Motivo de rechazo» | — | `FGS.cls:566, 598-599` |
| Deltas de subsanación | «Cambios detectados (Subsanación)» | — | `FGS.cls:601-606` |
| Historial RAC | «Historial de Validación y Revisión» | — | `FGS.cls:608-613` |
| Captura de fecha de envío RAC | «Notificar envio RAC» | `CAPTURA_FECHA_RAC` | `FGS.cls:663-672` |
| Captura de comentarios RAC | «Comentarios RAC» | `CAPTURA_COMENTARIOS_RAC` | `FGS.cls:674-679` |
| Rechazo técnico de calidad | «Rechazo calidad» | `CAPTURA_RECHAZO_CALIDAD` | `FGS.cls:785-791` |

Form_frmGestionSolicitud lo cierra si sigue abierto (`FGS.cls:919, 974, 1207, 1441, 1467`). Devuelve datos a `RecibirDatosRAC`, `RecibirDatosRechazoCalidad` y `RecibirImportacionWordRAC` (`FGS.cls:883-899`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnOpen` | Título, estado a cero e inicialización del control | `:40-50`, `.form.txt:42` |
| Formulario | — | `OnLoad` / `OnClose` | `Ajustar`; borra el fichero temporal | `:35-38, 52-58`, `.form.txt:43, 224` |
| `webVisor` | Edge | `OnDocumentComplete` | Marca el control como listo y pinta el HTML en cola | `:290-316`, `.form.txt:661-669` |
| `webVisor` | Edge | `OnBeforeNavigation` | Intercepta `vba:<COMANDO>:<dato>` | `:322-382`, `.form.txt:668` |
| `lblTitulo` | etiqueta | — | Recibe el título | `:136, 305`, `.form.txt:630` |
| `cmdCerrar` («Salir») | botón | `OnClick` | Cierra el visor | `:60-62`, `.form.txt:691-693` |

## Flujos

1. **Apertura:** `InicializarControl` crea, si falta, `%TEMP%\CONDOR\_blank.html` (`:79-107`) y navega a `https://msaccess/<ruta>`; si no puede crearlo, a `about:blank` (`:64-73`).
2. **Carga (`CargarContenido`, `:125-162`):** con tipo cacheable consulta `WebVisorCacheServicio` (`:131-134`); fija el título (`:136`); borra el temporal anterior (`:140-144`); si el control no está listo deja el HTML en cola (`:148-152`); si lo está, `RenderizarHTML` (`:155`).
3. **Caché (`WVC.cls:36-60`):** solo actúa con `UsarRenderizadoFichero = "Sí"` (`WVC.cls:66-72`), fijado así en el arranque (`modules/Variables Globales.bas:82`). Busca `captura_*.html` en `recursos\html\` del directorio de la aplicación (`WVC.cls:146-179`; `Entorno.cls:496-506`): si existe devuelve ese HTML; si no, guarda el generado (`WVC.cls:48-55`).
4. **Pintado:** con el indicador a «Sí» escribe `vis_<fecha>_<aleatorio>.html` en UTF-8 en `%TEMP%\CONDOR\WebVisor` y navega a él (`:191-209, 245-275`); si no, `document.write` por JavaScript tras quitar saltos de línea y escapar `'` (`:217-234`). El HTML en cola se pinta siempre por JavaScript (`:300-307`).
5. **Respuesta (`:322-376`):** cancela la navegación `vba:` y separa por `:` (`:327-332`).
   - `DATOS_RAC` y `DATOS_RECHAZO` ⇒ pasan el dato a Form_frmGestionSolicitud si está cargado (`:337-347`); los emiten los botones de captura (`WS.cls:3853-3854, 3884-3885, 3943-3947`).
   - `ABRIR_ADJUNTO` ⇒ `AdjuntosServicio.getAdjuntoPorID` (`classes/AdjuntosServicio.cls:267-269`), ruta `ANEXOS\` + nombre (`Entorno.cls:318-327`) y `AbrirEnLocal` (`modules/FUNCIONES UTILES.bas:404-425`) (`:349-366`); el enlace lo genera el historial RAC (`WS.cls:3790`).
   - `IMPORTAR_WORD_RAC` ⇒ `RecibirImportacionWordRAC` (`:368-371`).
   - Salvo `IMPORTAR_WORD_RAC`, el visor se cierra tras el comando (`:374-376`).
6. **Errores:** `CargarContenido` relanza como `CondorError` (`:158-161`); el pintado y los eventos del control registran con `LogServicio.RegistrarError` (`:236-238, 285-287, 313-315`); `RenderizarHTML`, `RenderizarModoArchivo` y `BeforeNavigate` los descartan sin aviso (`:183, 208, 380-381`).

## Lecturas y escrituras

| Operación | Destino | Tipo | Evidencia |
|---|---|---|---|
| Adjunto por identificador | `tbAdjuntos` | DAO parametrizado | `modules/AdjuntoRepositorio.bas:156-162` |
| Registro de errores | `tbLogErrores` | DAO `AddNew` | `LogServicio.cls:50-76`, `modules/LogErrorRepositorio.bas:23-24` |
| HTML temporal y cebador | `%TEMP%\CONDOR\` | fichero | `:85-98, 251-272` |
| Caché de capturas | `recursos\html\` compartido | fichero | `WVC.cls:103-124` |

## Coexistencia

Sin SQL propio. La lectura de `tbAdjuntos` y la escritura en `tbLogErrores` son parametrizadas. El contrato `vba:` y la ruta `ANEXOS\` dependen del cliente Windows; no dependen del backend.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** el descrito en Flujos.

**Posibles defectos (a verificar):**

- El control solo se marca listo con `about:blank` o URL vacía (`:296-298`), pero la inicialización navega a `https://msaccess/…` (`:69`): el HTML podría quedar en cola sin pintarse.
- Con caché, el HTML guardado sustituye al generado (`WVC.cls:51-52`); valores por defecto como la fecha del día (`FGS.cls:670`) quedarían fijados por la primera generación.
- El modo JavaScript no escapa `\` (`:223-229`).
- `IMPORTAR_WORD_RAC` no lo emite ningún HTML encontrado; el botón usa `window.colaComandos` con otro nombre (`WS.cls:250`).
- `EncodeBase64` no tiene llamadas (`:405-424`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: sustituir el visor por diálogos nativos de la web con formularios de captura y respuesta tipada.

## Criterios de aceptación (propuestos, no aprobados)

1. Cada uno de los seis orígenes muestra su contenido con su título.
2. Confirmar o cancelar una captura devuelve el dato a la solicitud y cierra el diálogo.
3. «[Word]» en el historial abre el adjunto indicado.
4. La captura de fecha propone la fecha del día en cada apertura.

## Pendientes

- Orden real de `DocumentComplete` del control Edge; origen del comando `vba:IMPORTAR_WORD_RAC`.
