# Form_FormGestionRiesgos

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormGestionRiesgos.cls` (423 líneas) y `src/forms/FormGestionRiesgos.form.txt` · **PRD:** [README](../README.md)

Formulario contenedor central de la gestión de riesgos. Aloja el encabezado, el pie y un `CtrNav` con cuatro botones de navegación que instancia[n] cuatro subformularios hijos distintos en un único control `FormDetalle`. Carga el contexto de la edición en curso, decide si la sesión es de alta o de edición y expone el flag público `blnPermitidoEditar` que leen sus hijos. No dibuja datos propios: los pinta el subformulario activo. Las citas son `Form_FormGestionRiesgos.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Menú principal | Cierra la instancia previa y abre esta | `Form_Form0BDOpciones.cls:291-294` |
| Menú principal | Conserva una referencia `Forms("FormGestionRiesgos")` | `Form_Form0BDOpciones.cls:295-296` |
| Listado de riesgos, gestión de proyectos y tareas de calidad | Cierra la instancia previa y reabre el contenedor | `Form_FormRiesgosGestion.cls:801-804`, `Form_FormProyectosGestion.cls:229-232`, `Form_FormCalidadTareasDetalleEdicion.cls:222-225` |

No acepta argumentos de apertura ni los necesita: el contexto llega por variables globales de sesión. **El patrón de entrada es «cerrar y reabrir»**, repetido en cuatro formularios distintos, para garantizar una única instancia y forzar la recarga de `EstablecerDatos`.

Salidas: abre el listado de riesgos y se cierra a sí mismo (`:221-225`). La subpestaña «Riesgos de oferta» abre a su vez el listado de la biblioteca pasando el nombre de este formulario como argumento de retorno (`Form_FormGestionRiesgosRiesgosOferta.cls:286`), que el destino lee con `Nz(Me.openArgs(), "")` (`Form_FormRiesgosBibliotecaGestion.cls:472`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Fija tamaño, llama a `EstablecerDatos`, apunta la navegación y `FormDetalle` a datos generales | `:265-298` |
| (formulario) | — | `Alta` / `Edicion` | Eventos públicos que este formulario **emite** al guardar | `:13-14`, emitidos en `:176` y `:179` |
| `blnPermitidoEditar` | — | — | Bandera pública de permiso de edición; la consume el hijo de datos generales | `:15`, `:350-356`; `Form_FormGestionRiesgosDatosGenerales.cls:495-499` |
| `CtrNav` | Marco de navegación | — | Aloja los cuatro botones que comparten `FormDetalle` | `.form.txt:405` |
| `NavDatosGenerales` | Navegación | — | «Datos generales» → `FormGestionRiesgosDatosGenerales` | `:280`, `.form.txt:434-449` |
| `NavResponsables` | Navegación | — | «Responsables» → `FormGestionRiesgosAutorizados` | `.form.txt:486-503` |
| `NavRiesgosOferta` | Navegación | — | «Riesgos de oferta» → `FormGestionRiesgosRiesgosOferta` | `.form.txt:540-557` |
| `NavSuministradores` | Navegación | — | «Suministradores» → `FormGestionRiesgosSuministradores` | `.form.txt:594-611` |
| `FormDetalle` | Subformulario | — | Hospeda el subformulario activo; su `SourceObject` se fija **en código**, no en diseño | `:281`, `.form.txt:697` |
| `lblTitulo` | Etiqueta | — | «ALTA DE GESTIÓN DE RIESGOS» o el nombre del proyecto en edición | `:346-348`, `.form.txt:365-366` |
| `lblPendienteGrabacion` | Etiqueta | — | «Pte. Grabación»; se oculta tras guardar | `:199`, `.form.txt:833-834` |
| `ComandoGrabar` | Botón | `ComandoGrabar_Click` | Refresca el objeto y llama a `grabar` | `.form.txt:855-857`, `:50-90` |
| `ComandoRegistrarRiesgos` | Botón | `ComandoRegistrarRiesgos_Click` | Abre el listado de riesgos | `.form.txt:890-892`, `:210-239` |
| `ComandoSalir` | Botón | `ComandoSalir_Click` | Cierra el formulario | `.form.txt:733-735`, `:241-263` |
| `ComandoAyuda` | Botón | `ComandoAyuda_Click` | Delega en `AbrirAyuda` | `.form.txt:300-302`, `:22-48` |

## Flujos

### Apertura (`Form_Load`, `:265-298`)

1. Ajusta `InsideHeight` e `InsideWidth` a valores fijos (`:275-276`).
2. Llama a `EstablecerDatos` y aborta si devuelve error (`:277-280`).
3. Fija `NavDatosGenerales.NavigationTargetName` y `FormDetalle.SourceObject` al subformulario de datos generales (`:280-281`).
4. Titula el formulario con el usuario conectado (`:282`) y cierra la barra de progreso (`:283`).

El manejador de error **cierra el formulario él mismo** (`:297`), a diferencia del resto de procedimientos, que solo avisan.

### Carga del contexto (`EstablecerDatos`, `:300-368`)

1. Si `m_EsAlta` sigue en `Empty`, aborta: no se sabe si es alta o edición (`:312-315`).
2. Sin proyecto activo crea uno en memoria y **oculta** el botón de riesgos; con proyecto activo lo muestra (`:316-321`).
3. **Rama de alta**: si el usuario es técnico, se rechaza con el mensaje «Es un usuario no autorizado» (`:322-326`); deja deshabilitados los tres botones de responsables, riesgos y suministradores (`:327-329`).
4. **Rama de edición**: recupera el proyecto de inicio si falta y aborta si no lo encuentra (`:331-341`); habilita los tres botones (`:342-344`).
5. Fija `blnPermitidoEditar` según el rol y, con él, el estado de `ComandoGrabar` (`:350-356`).

### Normal — guardar (`ComandoGrabar_Click` → `grabar`, `:50-208`)

1. Si el subformulario visible es el de datos generales, llama a su `RellenarDatosAlObjeto` para volcar los controles al objeto de proyecto (`:59-64`).
2. `grabar` crea el objeto de proyecto si no existe, copiando seis controles del subformulario (`:104-112`).
3. `HaHabidoCambios` compara siete campos con la instantánea de inicio; si no hay ninguno, aborta con «Ningún cambio que grabar» (`:113-121`, `:370-422`).
4. Si el indicador de informes pasa a «No», pide confirmación explícita y, si el usuario no acepta, **aborta en silencio** (`AvanceCerrar` y salida, `:123-143`).
5. Registra el proyecto (`:149`); en alta genera la primera edición, lanza las tareas del perfil y emite `Alta` (`:155-176`); en edición emite `Edicion` (`:178-179`).
6. Recarga ambos objetos de proyecto, fija la fecha máxima de publicación, cierra la barra y habilita botones y subpestañas (`:182-201`).

### Alternativo — ir a los riesgos (`:210-239`)

Cierra `FormRiesgosGestion` si estuviera abierto, abre ese listado y se cierra a sí mismo con `acSaveNo` (`:220-225`).

### Error

Los seis manejadores comparten una forma: distinguen su propio `Err.Raise 1000` de un error ajeno con `Err.Number <> 1000`, notifican al administrador y muestran un `MsgBox`; solo `Form_Load` añade además el cierre del formulario (`:297`). `grabar` y `EstablecerDatos` no muestran nada: devuelven el texto en el parámetro de salida (`:205-207`, `:365-367`).

## Reglas, validaciones y cálculos

- **El alta se decide por una variable global, no por un argumento.** `m_EsAlta` se inicializa en `Variables Globales.bas:167` y cada llamador lo fija antes de abrir (por ejemplo `Form_FormRiesgosGestion.cls:800`).
- **El permiso es derivado del rol, en un solo punto.** `blnPermitidoEditar` es `False` para técnico y `True` en el resto (`:350-353`), y gobierna tanto el botón Guardar como la editabilidad del hijo, que lo lee por referencia cruzada al módulo (`:495-499`).
- **`HaHabidoCambios` es un conjunto cerrado de siete campos** (`:377-408`): sin proyecto de inicio devuelve siempre `True`. Los formularios hermanos replican esta función con campos distintos, lo que sugiere copias paralelas de una utilidad no extraída.
- **Guardar no es idempotente desde el punto de vista del usuario**: el segundo intento sin cambios se trata como error, no como no-op.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Datos de la edición | Objetos `Proyecto` de sesión (`m_ObjProyectoActivo`, `m_ObjProyectoAlInicio`) | `:104-112`, `:183-184` |
| Proyecto persistido | `Proyecto.Registrar` | `:149` |
| Primera edición | `Proyecto.EdicionActiva` | `:157-165` |
| Tareas del perfil | `EstablecerTareasTecnico` / `EstablecerTareasCalidad` | `:166-170` |
| Identidad y título | `m_ObjEntorno.TituloUsuarioConectado` | `:282`, `:306` |
| Snapshot de cambios | Los siete campos de `HaHabidoCambios` | `:377-408` |

## Coexistencia

Coordinación con los hijos en tres niveles, sin un patrón único. **Dentro de `FormDetalle`**, el subformulario se elige asignando `SourceObject` y el código lee sus controles directamente (`:104-112`). **Entre pestañas**, el marco de navegación instancia cada hijo con su `NavigationTargetName` declarado en diseño (`.form.txt:449, 503, 557, 611`) y el contenedor solo habilita o deshabilita. **Entre hermanos**, la comunicación es global o por argumentos de apertura, en ambos sentidos. En una web esto se resuelve con un contenedor de rutas y estado compartido en el servidor; la colección de formularios abiertos y las variables globales de sesión no tienen equivalente.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-252.** `grabar` lee los controles del subformulario solo si `SourceObject` es exactamente el de datos generales (`:104`). Si el usuario está en «Responsables» o «Suministradores» y pulsa Guardar, el objeto de proyecto no se actualiza desde esos hijos.
- **H-GR-253.** `Err.Raise 1000` sin `vbObjectError` aparece en todo el archivo (`:32`, `:62`, `:117`…): depende de que el manejador siga el flujo correcto y no se comporta como un error de usuario de Access.
- **H-GR-254.** El aviso de «no generar correos» se abandona con `AvanceCerrar` y `Exit Function` (`:139-141`): ni mensaje de cancelación ni restauración de la barra; el usuario no distingue cancelar de un fallo de guardado.
- **H-GR-255.** `EstablecerDatos` rechaza el alta de un técnico con un `Err.Raise` (`:323-326`), pero `Form_Load` solo distingue ese error del resto: el mensaje mostrado es el acumulado en `p_Error`, no el específico de «usuario no autorizado».
- **H-GR-256.** El estado de las subpestañas depende de `blnPermitidoEditar`, pero los tres botones se habilitan sin comprobar ese flag (`:342-344`, frente a `:350-353`): un técnico puede navegar a subformularios que no puede editar.
- **H-GR-257.** `HaHabidoCambios` compara siete campos y no cubre el resto del objeto; si un hijo modifica un campo no listado, el guardado se rechaza como «sin cambios» (`:377-408`).
- **H-GR-258.** `EstablecerDatos` recalcula el título cada vez que se llama y `grabar` lo vuelve a invocar (`:72`), de modo que el rótulo depende del momento de la última llamada.
- **H-GR-259.** El «cerrar y reabrir» está replicado en cuatro formularios (`:221-225` y las cuatro entradas de la tabla); cualquier punto de entrada nuevo puede dejar dos instancias vivas.
- **H-GR-260.** `lblPendienteGrabacion` se oculta tras guardar (`:199`) pero este archivo no la muestra nunca: su estado inicial solo depende del diseño (`.form.txt:825`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: un único endpoint de guardado que devuelva el objeto completo en lugar de volcar controles del subformulario; derivar el permiso de edición del rol en el servidor y no de una bandera pública entre módulos; sustituir el conjunto cerrado de campos de «¿ha cambiado?» por comparación de versión o marca de tiempo; un solo punto de entrada con rutas en el servidor en vez del patrón «cerrar y reabrir»; pestañas como navegación de cliente con estado compartido en el servidor, sin controles INSTANCEADOS dinámicamente.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir el contenedor muestra la pestaña de datos generales con el contexto de la edición en curso.
2. Un alta de un usuario no autorizado se rechaza con un mensaje propio y el formulario se cierra.
3. Guardar sin cambios no escribe nada y avisa de que no hay nada que grabar; un error de guardado avisa y no deja la interfaz en estado intermedio.
4. Guardar deja habilitados la navegación secundaria y el acceso al listado de riesgos.
5. Cambiar de pestaña no pierde lo introducido en la anterior.

## Pendientes

- Qué escribe realmente cada subformulario hijo en el objeto de proyecto y por qué `grabar` solo recoge los datos generales; quién fija `lblEstado` en este formulario, ya que ningún procedimiento de este archivo lo hace.
- Si algún emisor escucha los eventos `Alta` y `Edicion`; en el código visible solo se emiten.
- Si `FormularioAbierto` resuelve una instancia o varias cuando el patrón «cerrar y reabrir» falla.