# Form_FormRiesgo

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormRiesgo.cls` (576 líneas) y `src/forms/FormRiesgo.form.txt` (1866 líneas) · **PRD:** [README](../README.md)

Contenedor del riesgo. Es el primero de los seis formularios de detalle de un riesgo (definición, plazo/coste/calidad, vulnerabilidad, mitigación, retirado y materializado): no tiene datos propios, decide qué subformulario montar, vuelca los controles de ese subformulario al objeto `Riesgo` y decide si se puede grabar. Las citas son `Form_FormRiesgo.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormRiesgosGestionRiesgo` | `DoCmd.OpenForm "FormRiesgo"` | `Form_FormRiesgosGestionRiesgo.cls:87` |
| `Form_FormRiesgosGestionEdicion` | `DoCmd.OpenForm "FormRiesgo"` | `Form_FormRiesgosGestionEdicion.cls:79` |
| `Form_FormCalidadTareaRiesgosMaterializadosPorDecidir` | `DoCmd.OpenForm "FormRiesgo"` | `Form_FormCalidadTareaRiesgosMaterializadosPorDecidir.cls:77` |
| `Form_FormCalidadTareaRiesgosAceptadosRetirados` | `DoCmd.OpenForm "FormRiesgo"` | `Form_FormCalidadTareaRiesgosAceptadosRetirados.cls:99` |
| `Form_FormTecnicoTareaRiesgosAceptadosRetirados` | `DoCmd.OpenForm "FormRiesgo"` | `Form_FormTecnicoTareaRiesgosAceptadosRetirados.cls:31` |
| `Form_FormCalidadTareaRiesgosRetipificacion` | `DoCmd.OpenForm "FormRiesgo"` | `Form_FormCalidadTareaRiesgosRetipificacion.cls:34` |

Se abre como formulario normal (`DoCmd.OpenForm`), no como subformulario. El riesgo no se pasa por parámetro: se toma de la global `m_ObjRiesgoActivo` (`Variables Globales.bas:80`) y el alta se marca con `m_EsAlta` (`:82`). Los formularios de detalle hermanos también lo referencian: leen `Form_FormRiesgo.m_PermitidoEditar` y llaman a su `EstablecerDatos` (`Form_FormRiesgoDefinicion.cls:98-120`, `Form_FormRiesgoMitigacion.cls:103, 150`, `Form_FormRiesgoRetirado.cls:60, 115`). Abre `FormRiesgosGestion` (`:122`), cierra la instancia previa de ese formulario antes de abrirla de nuevo (`:118-120`) y cierra el propio formulario (`:107`, `:123`, `:377`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Fija el tamaño de la ventana y refresca | `:139-170`, `.form.txt` (evento `OnLoad`) |
| `lblTitulo` | etiqueta | — | Título con tipo de operación y código de riesgo | `:285`, `.form.txt:569` |
| `lblEstadoRiesgo` | etiqueta | — | Estado del riesgo; se muestra y se recalcula | `:245`, `:251`, `.form.txt:1263` |
| `ComandoAyuda` | botón | `OnClick` → `ComandoAyuda_Click` | Abre la ayuda de la pantalla | `:78-103`, `.form.txt:598-606` |
| `ComandoSalir` | botón | `OnClick` → `ComandoSalir_Click` | Cierra sin guardar | `:106-108`, `.form.txt:1080-1089` |
| `ComandoGrabar` | botón | `OnClick` → `ComandoGrabar_Click` | Vuelca, compara y guarda el riesgo | `:318-392`, `.form.txt:1148-1158` |
| `ComandoVerEdicion` | botón | `OnClick` → `ComandoVerEdicion_Click` | Reabre la gestión de riesgos | `:110-137`, `.form.txt:1183-1194` |
| `ComandoActualizar` | botón | `OnClick` → `ComandoActualizar_Click` | Recarga el riesgo y refresca el subformulario | `:29-76`, `.form.txt:1219-1230` |
| `ControlDeNavegación0` | control de navegación | — | Aloja las pestañas; el destino se reasigna en código | `.form.txt:652`, `:286-289` |
| `NavGeneral` | botón de navegación | — | Primera pestaña; su destino se reasigna en `EstablecerDatos` | `.form.txt:682`, `:244`, `:286-289` |
| `NavPlazoCosteCalidad` | botón de navegación | `OnEnter` → `NavPlazoCosteCalidad_Enter` | Pestaña de plazo, coste y calidad | `:557-562`, `.form.txt:729-755` |
| `NavVulnerabilidad` | botón de navegación | `OnEnter` → `NavVulnerabilidad_Enter` | Pestaña de vulnerabilidad | `:571-576`, `.form.txt:781-807` |
| `NavMitigacion` | botón de navegación | `OnEnter` → `NavMitigacion_Enter` | Pestaña de mitigación | `:489-495`, `.form.txt:833-859` |
| `NavRetirado` | botón de navegación | `OnEnter` → `NavRetirado_Enter` | Pestaña de retirado; deshabilitada según el estado | `:564-569`, `.form.txt:885-911` |
| `NavMaterializado` | botón de navegación | `OnEnter` → `NavMaterializado_Enter` | Pestaña de materializado; deshabilitada según el estado | `:482-487`, `.form.txt:938-964` |
| `BotónDeNavegación1` | botón de navegación | — | «[Agregar nuevo]» del control de navegación, sin evento ni destino | `.form.txt:991` |
| `FrmDetalle` | subformulario | — | Contenedor del formulario de detalle; su `SourceObject` se fija en código | `.form.txt:1042`, `:307`, `:371` |

## Flujos

### Apertura (`Form_Load` → `EstablecerDatos`, `:139-315`)

1. `Form_Load` pone `InsideHeight = 9060` e `InsideWidth = 13980` (`:148-149`), llama a `EstablecerDatos m_Error` (`:152`) y propaga el error (`:153-154`).
2. `EstablecerDatos` fija el título de la ventana con el usuario conectado (`:182`) y resuelve la edición desde `m_ObjRiesgoActivo.Edicion` (`:184-185`), con dos rutas de error (`:187-196`) y una **comprobación repetida** (`:200-210`).
3. Calcula `m_PermitidoEditar` si está vacío (`:211-223`): no si la edición no está activa (`:213-214`), sí si el usuario está autorizado en el proyecto (`:216-217`), no en cualquier otro caso (`:219-220`).
4. Con permiso habilita `AllowEdits` y el botón de grabar; sin permiso, los deshabilita (`:226-231`).
5. En alta (`m_EsAlta`), sin permiso avisa y aborta (`:237-239`), crea `New riesgo` (`:241`), asigna `IDEdicion` (`:242`), deja la copia inicial a `Nothing` (`:243`) y oculta la etiqueta de estado (`:245`).
6. En edición, sin datos, avisa y aborta (`:247-249`); con datos, muestra la etiqueta de estado (`:251`) y recarga la copia inicial con `Constructor.getRiesgo` (`:252-259`).
7. Calcula `m_EstadoRiesgo = m_ObjRiesgoActivo.EstadoEnum` (`Riesgo.cls:2351`) y deshabilita las pestañas de retirado y materializado si el estado es incompleto o vacío (`:264-269`).
8. Compone el título: «ALTA DE RIESGO PARA LA EDICIÓN …» (`:279-280`) o «DETALLE DEL RIESGO: <código> EDICIÓN …» (`:281-282`).
9. Elige el formulario de definición según `Proyecto.RequiereRiesgoDeBibliotecaCalculado` (`Proyecto.cls:667`): `FormRiesgoDefinicion` (`:286-287`) o `FormRiesgoDefinicionNoBiblioteca` (`:288-289`).
10. Refresca etiqueta de estado y color de pestañas (`:292-299`), captura el nodo seleccionado y el filtro de descripción si la gestión está abierta (`:301-302`) y monta el subformulario (`:307`).

### Normal: grabar (`ComandoGrabar_Click`, `:318-392`)

1. Reloj de arena y `m_Error = ""` (`:327-330`).
2. `RellenarObjeto` (`:331`) y comparación con la copia inicial: `HaHabidoCambios(m_Error)` (`:332`). Si no hay cambios, aborta con «Ningún cambio que guardar» (`:333-335`).
3. En alta asigna `IDEdicion` (`:339-341`) y guarda con `Registrar EnumSiNo.Sí, m_ObjRiesgoAlInicio, m_Error` (`:344`; `Riesgo.cls:4799`).
4. Recalcula el estado y la etiqueta (`:348-349`) y vuelve a habilitar o deshabilitar las pestañas de retirado y materializado (`:351-356`).
5. Si era alta, emite `RaiseEvent RiesgoNuevo` (`:359-360`). Si no, actualiza el icono del árbol (`:364`; `Funciones Generales.bas:6857`) y refresca `Form_FormRiesgosGestionRiesgo` si es la vista montada (`:372`).
6. **Cierra el formulario** (`:377`), llama `AvanceCerrar` (`:378`) y quita el reloj (`:379-381`).

### Alternativo: cambiar de pestaña

Los cinco `Nav*_Enter` hacen lo mismo: si `Me.AllowEdits = True`, llaman a `RellenarObjeto` para volcar los controles al objeto antes de abandonar la pestaña (`:483-484`, `:490-491`, `:558-559`, `:565-566`, `:572-573`). Ninguno mira el error que devuelven.

### Alternativo: actualizar y recargar (`ComandoActualizar_Click`, `:29-76`)

Recarga el riesgo activo y la copia inicial con `Constructor.getRiesgo` (`Constructor.bas:1827`) (`:38-44`), toma el formulario embebido (`:46`) y delega el refresco según cuál esté montado: `FormRiesgoDefinicion`, `FormRiesgoPlazoCosteCalidad`, `FormRiesgoVulnerabilidad`, `FormRiesgoMitigacion`, `FormRiesgoRetirado` o `FormRiesgoMaterializado` (`:47-61`).

### De error

Los manejadores siguen el patrón común: quitar el reloj y, si `Err.Number <> 1000`, componer el mensaje, escribirlo al administrador (`Funciones Generales.bas:964`) y avisar con diálogo crítico; si es 1000, solo advertencia (`:67-74`, `:94-101`, `:128-135`, `:161-168`, `:383-390`). En `ComandoGrabar_Click` el cierre (`:377`) y `AvanceCerrar` (`:378`) quedan **sin ejecutar** si hay error. `EstablecerDatos`, `HaHabidoCambios` y `RellenarObjeto` no avisan a nadie: solo rellenan su argumento de error (`:311-313`, `:476-478`, `:549-551`). `ComandoSalir_Click` es el único manejador sin `On Error` (`:106-108`).

## Reglas, validaciones y cálculos

- **El permiso se calcula una vez y se guarda en el formulario.** `m_PermitidoEditar` es un campo de la clase (`:22`): si ya está informado, `EstablecerDatos` no lo recalcula (`:211`), y los formularios hermanos lo leen directamente.
- **Alta sin permiso, no hay alta.** Con `m_EsAlta = Sí` y sin permiso, la pantalla se niega a abrir (`:237-239`).
- **El estado gobierna las pestañas.** `Incompleto` o `Empty` dejan sin usar las pestañas de retirado y materializado (`:264-269`), y el guardado vuelve a aplicar la misma regla (`:351-356`).
- **Solo se guarda si algo cambió.** `HaHabidoCambidos` compara 15 campos uno a uno y devuelve en la primera diferencia (`:407-472`); en alta devuelve `True` sin comparar (`:400-403`).
- **La comparación depende del proyecto:** `CausaRaiz` solo se compara dentro de `If .Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí` (`:415-419`). **Los valores ausentes se normalizan a cadena vacía:** `RellenarObjeto` envuelve todas las asignaciones en `Nz(..., "")` (`:503-539`). **Sin acceso directo a datos:** sin `getdb()`, `Recordset`, `Open`, `Execute`, SQL, `DLookup` ni transacciones, todo pasa por `Constructor` y por la clase `Riesgo`.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Riesgo sobre el que se opera | `m_ObjRiesgoActivo` (`Variables Globales.bas:80`) | `:38`, `:42`, `:184-262`, `:501-539` |
| Copia para detectar cambios | `m_ObjRiesgoAlInicio` (`Variables Globales.bas:83`) | `:42`, `:243`, `:252`, `:406-472` |
| Edición y proyecto | `m_ObjEdicionActiva` (`Variables Globales.bas:79`), `Constructor.getEdicion` (`Constructor.bas:24`) | `:185`, `:192` |
| Marcador de alta | `m_EsAlta` (`Variables Globales.bas:82`) | `:235`, `:279`, `:339` |
| Estado y pestaña | `Riesgo.EstadoEnum` (`Riesgo.cls:2351`) | `:262`, `:348` |
| Datos del riesgo | `Riesgo.cls` (`IDRiesgo:13`, `IDEdicion:14`, `CodigoRiesgo:16`, `CodRiesgoBiblioteca:47`) | `:241-242`, `:282`, `:406-472` |
| Permiso de usuario | `Proyecto.UsuarioAutorizado` (`Proyecto.cls:3780`) | `:216-217` |
| Elección del formulario de definición | `Proyecto.RequiereRiesgoDeBibliotecaCalculado` (`Proyecto.cls:667`) | `:286-289`, `:415` |
| Guardado del riesgo | `Riesgo.Registrar` (`Riesgo.cls:4799`) | `:344` |
| Refresco del árbol de riesgos | `ActualizarIconoRiesgo` (`Funciones Generales.bas:6857`) | `:364` |
| Filtro de vista heredado | `Forms("FormRiesgosGestion").Controls("comboVerDescripcion")` | `:302` |

Escritura efectiva: una llamada a `Riesgo.Registrar` (`:344`). Todo lo demás relee o delega en clases.

## Coexistencia

Pendiente de `../coexistencia.md`. Lo que ya aparece: el formulario no toca la base de datos, pero depende de variables globales compartidas entre seis subformularios y de un árbol de riesgos abierto en otra ventana, de modo que la concurrencia entre dos riskedores sobre el mismo riesgo no tiene ninguna barrera visible en esta pantalla.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-116.** `EstablecerDatos` repite la misma comprobación de edición activa con los mismos dos mensajes: primero en `:184-198` y otra vez en `:200-210`. Es un bloque duplicado dentro del mismo procedimiento.
- **H-GR-117.** `ComandoGrabar_Click` llama a `RellenarObjeto` (`:331`) y a `HaHabidoCambios` (`:332`) y **no comprueba el error que devuelven** antes de guardar: el `Err.Raise` conditioned solo llega tras `Registrar` (`:345-346`). Si el volcado o la comparación fallan, el guardado persiste el objeto a medias.
- **H-GR-118.** `RellenarObjeto` escribe **cadena vacía** en los campos del subformulario de mitigación cuando el control no está visible (`:525-528`, `:530-533`). Al venir el valor de `Nz`, un campo que en la base de datos es nulo o con contenido pasa a vaciarse sin que el usuario haya tocado nada.
- **H-GR-119.** Los cinco `Nav*_Enter` llaman a `RellenarObjeto` **sin manejador de error y sin mirar `m_Error`** (`:483-484`, `:490-491`, `:558-559`, `:565-566`, `:572-573`). Un fallo al cambiar de pestaña deja el objeto a medias y el guardado posterior escribe esos datos sin avisar.
- **H-GR-120.** La comparación de `Coste` y `Plazo` (`:441-448`) no normaliza valores, mientras que el volcado escribe `Nz(m_Form.ListaCoste.Column(0), "")` (`:520`). Si el objeto conserva un valor no nulo y el control devuelve cadena vacía, la comparación detecta un cambio fantasma.
- **H-GR-121.** `CausaRaiz` solo se compara cuando `RequiereRiesgoDeBibliotecaCalculado = Sí` (`:415-419`). Si el valor cambia sin que cambie ese indicador, el cambio **no se detecta** y el guardado se rechaza con «Ningún cambio que guardar».
- **H-GR-122.** La clase declara y expone dos eventos que nunca se emiten en el clon: `RiesgoEditado` (`:19`) y `cambioEnPriorizacion` (`:21`); el único `RaiseEvent` es `RiesgoNuevo` (`:360`). Contrato público sin implementador.
- **H-GR-123.** `ComandoSalir_Click` es el único manejador sin `On Error`, sin reloj y sin mensaje (`:106-108`): si `DoCmd.Close` falla, el fallo desaparece.
- **H-GR-124.** `Form_Load` fija el tamaño de la ventana en twips literales (`InsideHeight = 9060`, `InsideWidth = 13980`, `:148-149`) en vez de usar el mecanismo de ajuste que sí existe en el resto de la aplicación (`AjustarTamaño`, `Funciones Generales.bas:1043`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: detección de cambios en el servidor o por campo con marca de versión, en vez de comparar 15 campos uno a uno; detección de conflicto explícita cuando dos usuarios editan el mismo riesgo; pestañas como recursos del riesgo en una sola vista; subida del control al perder el foco con error visible; Elimination del contrato de eventos sin implementador; cierre del formulario solo después de confirmar el guardado.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir la pantalla muestra el riesgo con su estado y la pestaña de definición correcta según el tipo de proyecto.
2. Sin permiso sobre la edición, el formulario no permite grabar.
3. En alta, el riesgo se crea y la pantalla avisa de que el alta se ha realizado.
4. Sin cambios, el guardado avisa y no escribe nada.
5. Cambiar de pestaña conserva lo escrito en la pestaña de origen.
6. Un error al guardar avisa y no cierra la pantalla con un estado que no se corresponde con lo guardado.

## Pendientes

- Qué tablas y qué control de cambios aplica `Riesgo.Registrar` (`Riesgo.cls:4799`), y si admite edición concurrente del mismo riesgo.
- Qué hacen `ActualizarEtiquetaEstadoRiesgo` (`Funciones Generales.bas:7480`) y `EstablecerColorPestañasRiesgo` (`Funciones Generales.bas:7372`) exactamente (**H-GR-120**), y si `Proyecto.UsuarioAutorizado` (`Proyecto.cls:3780`) comprueba un rol o un usuario concreto, porque de eso depende `m_PermitidoEditar`.