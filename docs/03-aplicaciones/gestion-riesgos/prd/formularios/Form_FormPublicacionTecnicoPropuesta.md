# Form_FormPublicacionTecnicoPropuesta

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormPublicacionTecnicoPropuesta.cls` (425 líneas) y `src/forms/FormPublicacionTecnicoPropuesta.form.txt` (1387 líneas) · **PRD:** [README](../README.md)

Pantalla de la propuesta técnica de publicación de una edición. Es la tercera pestaña del bloque de publicación técnica: decide si la edición se propone o se retira de la propuesta, deja generar los dos informes asociados (el de publicación y el de publicabilidad), deja volcar el historial de publicaciones a fichero y da acceso a la pantalla de rechazo de calidad. Las citas son `Form_FormPublicacionTecnicoPropuesta.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormPublicacionTecnico` | Pestaña de navegación lateral: el contenedor fija `NavigationTargetName = "FormPublicacionTecnicoPropuesta"` y lo carga como subformulario | `Form_FormPublicacionTecnico.cls:124-125` |

No hay ninguna llamada a `DoCmd.OpenForm` con este formulario en el clon: se llega siempre como subformulario anidado, y la edición sobre la que opera no llega como parámetro, sino de la global `m_ObjEdicionActiva` (`Variables Globales.bas:79`). Abre `FormPublicacionCalidadRechazoObservaciones` (`:130`) y genera un fichero HTML en disco que lanza con el shell (`Ejecutar Me.hWnd, "open", url, "", "", 1`, `:42`). Cierra, según el camino, `FormPublicacionTecnico` (`:361`, `:404`), `FormRiesgosGestion` (`:364`, `:407`) y el propio formulario (`:367`, `:410`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Valida que haya edición activa y refresca rótulos y botones | `:147-183`, `.form.txt:41` |
| `lbl1` | etiqueta | — | Rótulo del bloque de evidencia UTE; el texto lo pone el código | `.form.txt:615` |
| `lbl2` | etiqueta | — | Rótulo del bloque de evidencias de suministradores | `.form.txt:639` |
| `lbl1Requerido` | etiqueta | — | «Requerido» / «No Requerido» según el proyecto esté en UTE | `:196`, `:204` |
| `lbl1Complementado` | etiqueta | — | «Complementado» / «No complementado» / «N/A» con color | `:197-202`, `:204-207` |
| `lbl2Requerido` | etiqueta | — | «Requerido» / «No Requerido» según `EvidenciasSuministradoresRequeridas` | `:210`, `:219` |
| `lbl2Complementado` | etiqueta | — | «Complementado» / «No complementado» / «N/A» con color | `:211-217`, `:219-222` |
| `ComandoProponerPublicacion` | botón | `OnClick` → `ComandoProponerPublicacion_Click` | Propone o quita la propuesta según su propio rótulo | `:291-425`, `.form.txt:759-770` |
| `ComandoHistorialPublicaciones` | botón | `OnClick` → `ComandoHistorialPublicaciones_Click` | Vuelca el historial de publicaciones a fichero HTML | `:91-115`, `.form.txt:796-807` |
| `ComandoGenerarInforme` | botón | `OnClick` → `ComandoGenerarInforme_Click` | Genera el informe de publicación y lo abre | `:16-59`, `.form.txt:833-844` |
| `ComandoVerRechazo` | botón | `OnClick` → `ComandoVerRechazo_Click` | Abre la pantalla de rechazo de calidad por observación | `:119-145`, `.form.txt:871-882` |
| `lblPublicacion` | etiqueta | — | Estado de la propuesta en texto y color | `:237-253`, `.form.txt:907` |
| `ComandoGenerarInformePublicabilidad` | botón | `OnClick` → `ComandoGenerarInformePublicabilidad_Click` | Genera el informe de publicabilidad y guarda su URL en la global | `:62-89`, `.form.txt:931-942` |

Las siete etiquetas del diseño se escriben desde `EstablecerDatos` y no tienen evento. `ComandoVerRechazo` llega al diseño con `Visible = NotDefault` (`.form.txt:871`) y su visibilidad la decide el código.

## Flujos

### Apertura (`Form_Open`, `:147-183`)

1. Reloj de arena con cedimientos y `m_Error = ""` (`:151-154`).
2. Si `m_ObjEdicionActiva Is Nothing`, escribe «No hay ninguna edición activa» y lanza el error controlado (`:158-160`).
3. Llama a `EstablecerDatos m_Error` y, si vuelve con error, lo relanza (`:162-164`).
4. `AvanceCerrar` (`:166`) y quita el reloj (`:167-169`).

### Normal: proponer (`ComandoProponerPublicacion_Click`, `:291-425`, ruta A)

1. Reloj de arena y `m_Error = ""` (`:297-301`).
2. **Lee su propio `Caption`** para decidir la operación: si no es exactamente «Proponer para publicación» ni «Quitar propuesta para publicación», aborta con «Texto del botón no reconocido» (`:303-306`).
3. Si la edición no es publicable, genera y abre el informe de publicabilidad (`:317`) y acto seguido aborta con un texto fijo (`:324-326`).
4. Pide confirmación: «No olvide lo siguiente: Riesgos del suministrador / Riesgos de los pedidos / ¿Desea seguir con la propuesta?», con `vbYesNo` y botón por defecto el sí (`:331-334`). Si no es `vbYes`, `AvanceCerrar` y sale sin proponer (`:335-337`).
5. `ProponerParaPublicacion` sobre la edición activa (`:342`; `Edicion.cls:3159`) y, si falla, error controlado (`:343-344`).
6. Informa de que la edición quedó propuesta y avisa de que se ha enviado correo a Calidad (`:349-350`).
7. Recarga el proyecto con `Constructor.getProyecto` (`:355`; `Constructor.bas:675`) y regenera las tareas de calidad con `EstablecerTareasCalidad p_Reseteando:=EnumSiNo.Sí` (`:356`; `Funciones Generales.bas:2054`).
8. Cierra `FormPublicacionTecnico` y `FormRiesgosGestion` si estaban abiertos (`:360-364`) y luego se cierra a sí mismo sin guardar (`:367`).

### Alternativo: quitar la propuesta (misma rutina, ruta B, `:371-411`)

1. Reloj de arena (`:371-373`).
2. Pide confirmación: «¿Desea quitar la propuesta de publicación», con `vbYesNo` y **sin** botón por defecto (`:374`). Si no es `vbYes`, sale sin borrar (`:375-377`).
3. `QuitarPropuestaParaPublicacion` (`:383`; `Edicion.cls:3237`).
4. Informa de que se ha enviado correo a Calidad (`:393-394`), recarga proyecto (`:398`) y tareas (`:399`), cierra los dos contenedores (`:403-407`) y se cierra (`:410`).

Los dos caminos comparten preámbulo, manejador de error y cierre de contenedores; solo cambia el cuerpo entre `:369` y `:371`.

### Alternativo: informes, historial y rechazo

- `ComandoGenerarInforme_Click` (`:16-59`) exige `TodosLosRiesgosPriorizados <> No` (`:22-24`), llama a `GenerarInformePublicacion` (`Funciones Generales.bas:1409`) y abre el HTML generado (`:34`, `:42`).
- `ComandoGenerarInformePublicabilidad_Click` (`:62-89`) llama a `GenerarInformePublicabilidadEdicionInteractivoHTML` (`PublicabilidadEdicion.bas:781`) y guarda la URL en la global `m_URLHTMLActivo` (`:75`), **sin abrirla**.
- `ComandoHistorialPublicaciones_Click` (`:91-115`) vuelca `m_ObjEdicionActiva.HTMLHistorialPublicaciones` (`Edicion.cls:529`) con `HTMLENTXT` (`Funciones Generales.bas:3438`), y `ComandoVerRechazo_Click` (`:119-145`) cierra la instancia previa de `FormPublicacionCalidadRechazoObservaciones` si la hay y vuelve a abrirla (`:127-130`).

### De error

Todos los manejadores siguen el mismo patrón: `On Error GoTo errores`, quitar el reloj y, si `Err.Number <> 1000`, componer el mensaje, escribirlo al administrador con `CorreoAlAdministrador` (`Funciones Generales.bas:964`) y avisar con diálogo crítico; si es 1000, solo aviso de advertencia (`:49-57`, `:80-87`, `:106-113`, `:136-143`, `:412-419`). Ninguno de ellos cierra el formulario.

## Reglas, validaciones y cálculos

- **No se propone sin riesgos priorizados.** El botón de informe aborta si `TodosLosRiesgosPriorizados = EnumSiNo.No` (`:22-24`), pero el botón de propuesta no repite la comprobación: delega en la edición con `ProponerParaPublicacion` (`:342`).
- **No se propone sin superar la publicabilidad.** `Publicable = EnumSiNo.No` dispara el informe de publicabilidad y el aborto (`:317-326`).
- **El permiso es global, no por rol.** `EstablecerDatos` habilita el botón de propuesta según `blnPermitidoEscribir` (`Variables Globales.bas:121`) y solo si la edición está activa (`:225-229`, `:260`).
- **El estado visible se deriva de dos fechas.** «SIN PROPUESTA TÉCNICA» si `FechaPreparadaParaPublicar` no es fecha (`:237-239`); «ESPERANDO A CALIDAD» si no hay fecha de rechazo (`:241-244`) o si la propuesta es posterior al rechazo (`:246-249`); «RECHAZADA» si la propuesta es anterior (`:250-253`).
- **El requisito UTE y el de suministradores se calculan por separado**, y ninguno de los dos toca la base de datos: no hay `getdb()`, `Recordset`, `Open`, `Execute`, SQL, `DLookup` ni transacciones en el fichero, todo pasa por clases. Bloque UTE con `Proyecto.EnUTE` (`Proyecto.cls:36`), `TieneAnexoEvidenciaUTE` (`Edicion.cls:1133`) y `AnexoEvidenciaUTE` (`Edicion.cls:1103`); bloque de suministradores con `EvidenciasSuministradoresRequeridas` (`Edicion.cls:4221`) y `…Completadas` (`Edicion.cls:4252`). Los colores salen de `Constantes.bas:96-97`.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Edición sobre la que se opera | `m_ObjEdicionActiva` (`Variables Globales.bas:79`) | `:22`, `:99`, `:158`, `:194` |
| Proyecto en UTE | `m_ObjEdicionActiva.Proyecto.EnUTE` (`Proyecto.cls:36`) | `:196` |
| Evidencias UTE y de suministradores | `Edicion.TieneAnexoEvidenciaUTE`, `EvidenciasSuministradoresRequeridas` / `…Completadas` | `:197`, `:210-211` |
| Fechas de propuesta, rechazo y publicación | `FechaPreparadaParaPublicar`, `PropuestaRechazadaPorCalidadFecha`, `FechaPublicacion` (`Edicion.cls:33, 35, 25`) | `:237-273` |
| Permiso de escritura | `blnPermitidoEscribir` (`Variables Globales.bas:121`) | `:225` |
| Informe de publicación | `GenerarInformePublicacion` (`Funciones Generales.bas:1409`) | `:34` |
| Informe de publicabilidad | `GenerarInformePublicabilidadEdicionInteractivoHTML` (`PublicabilidadEdicion.bas:781`) | `:71`, `:317` |
| Propuesta de publicación | `Edicion.ProponerParaPublicacion` (`Edicion.cls:3159`) | `:342` |
| Retirada de propuesta | `Edicion.QuitarPropuestaParaPublicacion` (`Edicion.cls:3237`) | `:383` |
| Recarga de proyecto y tareas de calidad | `Constructor.getProyecto` (`Constructor.bas:675`), `EstablecerTareasCalidad` (`Funciones Generales.bas:2054`) | `:355-356`, `:398-399` |
| Historial de publicaciones | `Edicion.HTMLHistorialPublicaciones` (`Edicion.cls:529`) | `:99` |

Escritura efectiva: una sola llamada a la clase `Edicion` en cada ruta; quien decide qué tabla toca es la clase, no la vista.

## Coexistencia

Pendiente de `../coexistencia.md`. Lo que ya se ve: la pantalla no habla con la base de datos, pero depende por completo de la sesión global (`m_ObjEdicionActiva`) y de que el contenedor haya cargado la edición; en web el contexto sería un recurso explícito y la propuesta una operación idempotente sobre esa edición, no un estado repartido entre variables globales y rótulos.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-101.** El botón decide entre proponer y quitar **leyendo su propio `Caption`** (`:303-306`), y `EstablecerDatos` es quien fija ese texto (`:230-234`, `:263`). El comportamiento queda acoplado a una cadena literal del rótulo: cualquier traducción, cambio de mayúsculas o rediseño deja la operación inservible con el mensaje «Texto del botón no reconocido».
- **H-GR-102.** En error, `Form_Open` avisa **dos veces**: el manejador muestra el diálogo (`:178`) y después vuelve a llamar `AvanceCerrar` (`:180`) y muestra un segundo diálogo crítico con el mismo `m_Error` (`:181`). El usuario ve el mismo fallo duplicado.
- **H-GR-103.** En la ruta A, cuando la edición no es publicable se genera y **se abre** el informe de publicabilidad (`:317`) y acto seguido `m_Error` se sobrescribe con un texto fijo y se aborta (`:324-326`). El motivo real que devolvió la clase se pierde, y el usuario sale con el formulario cerrado y sin saber qué le faltaba.
- **H-GR-104.** La comparación `CDate(.FechaPreparadaParaPublicar) > CDate(.PropuestaRechazadaPorCalidadFecha)` (`:246`) no comprueba `IsDate`, al contrario que las otras dos ramas (`:237`, `:266`). Si una de las dos fechas llega como texto no fecha, `CDate` revienta y el manejador lo traduce a «error inesperado» con correo al administrador.
- **H-GR-105.** `EstablecerDatos` es una `Public Function` sin tipo de retorno cuyo manejador **no avisa ni escribe a nadie**: solo rellena `p_Error` (`:283-285`). Como el error viaja únicamente por ese argumento, cualquier llamador que lo ignore deja la pantalla a medias sin rastro.
- **H-GR-106.** `EstablecerDatos` decide la visibilidad de `ComandoVerRechazo` (`:241-262`), pero el diseño lo entrega con `Visible = NotDefault` (`.form.txt:871`). Si el refresco falla antes de llegar a esa rama, manda el diseño y el botón queda visible sin una propuesta rechazada que mirar.
- **H-GR-107.** El camino de error de `ComandoProponerPublicacion_Click` no ejecuta el cierre de `FormPublicacionTecnico`, `FormRiesgosGestion` ni del propio formulario (`:412-419` frente a `:360-367` y `:403-410`). Cuando falla el guardado, las pantallas quedan abiertas con la edición ya propuesta o deshecha en la base de datos y la interfaz sin reflejarlo.
- **H-GR-108.** `ComandoHistorialPublicaciones_Click` llama a `HTMLENTXT` (`:99`) y **no comprueba `m_Error`**, al contrario que las demás rutinas de la pantalla. Un fallo al escribir el fichero termina en `Exit Sub` sin avisar a nadie.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: una acción explícita con identificador estable en vez de un rótulo que actúe de interruptor; verificación de publicabilidad y priorización en el servidor y no en el botón; historial de propuestas y retiradas como recurso auditable de la edición; permisos de escritura explícitos por rol en lugar de la global `blnPermitidoEscribir`; mensajes de motivo sin sobrescribir el error de la clase.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir la pantalla con una edición activa muestra el estado de la propuesta y los requisitos UTE y de suministradores.
2. Sin edición activa, la pantalla no abre y avisa de que no hay ninguna.
3. Proponer exige confirmación, deja la edición en estado propuesto y confirma la operación.
4. Quitar la propuesta pide confirmación y devuelve la edición al estado previo.
5. El botón de informe no hace nada si faltan riesgos por priorizar.
6. Un fallo al guardar avisa una sola vez y no deja pantallas abiertas con un estado que no corresponde al guardado.

## Pendientes

- Qué tablas y qué correo envía cada una de las dos rutas dentro de `Edicion.cls:3159` y `Edicion.cls:3237`, si queda traza de la propuesta retirada y qué comprueba exactamente `GenerarInformePublicabilidadEdicionInteractivoHTML` (`PublicabilidadEdicion.bas:781`), incluido si su HTML depende de un navegador embebido (**H-GR-103**).