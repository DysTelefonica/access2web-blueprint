# Form_FormProyectosGestion

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormProyectosGestion.cls` (956 líneas) y `src/forms/FormProyectosGestion.form.txt` · **PRD:** [README](../README.md)

Listado de gestiones de riesgos del proyecto. Es la pantalla de entrada al trabajo diario: desde aquí se da de alta un proyecto, se editan sus datos, se entra a sus riesgos y anexos, se ve la evolución y se borra. Filtra por cuatro criterios y adapta la botonera al rol del usuario conectado, técnico o Calidad. Es el formulario con más superficie de eventos de la aplicación. Las citas son `Form_FormProyectosGestion.cls:línea` salvo indicación.

## Entrada y salida

No acepta parámetros ni `OpenArgs`: el contexto llega íntegro por globales de sesión (`m_ObjProyectoActivo`, `m_ObjEdicionActiva`, `m_ObjProyectoAlInicio`, `m_EsAlta`).

| Desde | Evidencia |
|---|---|
| Menú general de opciones | `Form_Form0BDOpciones.cls:177` |
| Menú de opciones de técnico | `Form_Form0BDOpcionesTecnico.cls:69` |

Salidas, todas con el patrón «cerrar si está abierta y abrir de nuevo», que las hace de instancia única: `FormGestionRiesgos` desde el alta y la edición (`:233-237`, `:405-409`), `FormRiesgosGestion` desde «Riesgos» y desde el doble clic (`:452-456`, `:780-787`), el Excel de evolución por el visor del sistema (`:504-506`) y el último informe publicado en local (`:649-651`).

Expone además `Public Function ActualizarProyecto` (`:909`), llamado por la propia clase de dominio (`Proyecto.cls:1563`) y por `Form_FormRiesgosGestionEdicion.cls:277` para refrescar una fila cuando el proyecto cambia.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Tamaño, combos, rol y primer filtrado | `:569-641` |
| `ListaFiltrados` | cuadro de lista | `ListaFiltrados_Click` | Resuelve el proyecto, enciende la botonera y busca el último informe | `:671-778`, `.form.txt:695` |
| `ListaFiltrados` | — | `ListaFiltrados_DblClick` | Doble clic equivale a «Riesgos» | `:780-787` |
| `comboActivos` | desplegable | `…_BeforeUpdate` | Refiltra al confirmar | `:557-559`, `.form.txt:1259-1263` |
| `ComboJefeProyecto` | desplegable de dos columnas | `…_BeforeUpdate` | Refiltra al confirmar | `:561-563`, `.form.txt:792-802` |
| `ComboResponsableCalidad` | desplegable | `…_BeforeUpdate` | Refiltra al confirmar | `:565-567`, `.form.txt:1546-1551` |
| `PalabraClave` | cuadro de texto | `PalabraClave_KeyDown` | Con Intro filtra, tras un salto de foco | `:314-321`, `.form.txt:751-753` |
| `ComandoFiltrar` / `ComandoActualizar` | botones | `…_Click` | Filtra; el segundo invalida la caché y refiltra | `:322-346`, `:145-180`, `.form.txt:2147`, `:1851` |
| `ComandoLimpiarActivos`, `ComandoLimpiarAutorizado`, `ComandoLimpiarPalabraClave`, `ComandoLimpiarResponsableCalidad` | botones | `…_Click` | Vacían un criterio y refiltran | `:347-378` |
| `ComandoNuevoProyecto` | botón | `…_Click` | Alta de proyecto; bloqueada a técnicos | `:379-421`, `.form.txt:1996-2000` |
| `ComandoEditarProyecto` | botón | `…_Click` | Datos principales y usuarios autorizados | `:209-248`, `.form.txt:2034-2038` |
| `ComandoEliminar` | botón | `…_Click` | Borra el proyecto con confirmación | `:249-308`, `.form.txt:2073-2077` |
| `ComandoRiesgos` | botón | `…_Click` | Entra a los riesgos | `:422-468`, `.form.txt:2111-2115` |
| `ComandoVerEvolucionRiesgos` | botón | `…_Click` | Abre el Excel de evolución | `:474-521`, `.form.txt:1927-1931` |
| `ComandoAnexar` | botón | — | **Sin manejador de clic**; se muestra y se habilita | `:36`, `:811`, `.form.txt:1472-1474` |
| `ImagenDocumentoUltimPublicacion` | imagen | `…_Click` | Abre en local el informe de la última edición | `:643-670`, `.form.txt:1192-1194` |
| `ComandoAyuda` / `ComandoSalir` | botones | `…_Click` | Ayuda y cierre | `:181-208`, `:469-473`, `.form.txt:596`, `:2245` |
| — | evento del hijo | `m_FormGestionRiesgos_Alta` | Refiltra cuando el hijo confirma un alta | `:309-312` |
| — | evento del hijo | `m_FormAnexos_AnexoAñadido` / `…_AnexoEliminado` | Recargan el proyecto, **pero la referencia nunca se asigna** | `:832-908`, `:23` |

## Flujos

### Apertura (`Form_Load`, `:569-641`)

1. `AjustarTamaño Me` (`:575`), rótulo de ventana y cabecera de la lista (`:582-584`), y `Avance "Estableciendo datos de usuarios...."` (`:586`), que es lo que ve el usuario durante la carga.
2. `RellenaComboJps` (`:592`; `:522-556`) puebla el desplegable de autorizados con **una o dos columnas según las que tenga el combo** (`:545-549`). El mismo método sirve para otros combos y se adapta en lugar de recibir un parámetro.
3. `EstablecerComboUsuariosCalidad` (`:599`) para el responsable de Calidad y `comboActivos = "Sí"` (`:600`): por defecto solo proyectos activos.
4. **Ramificación por rol** (`:602-609`): el técnico ve sus propios autorizados, el botón dice «Ver Datos Principales» y el de eliminar se oculta; para Calidad el botón dice «Editar Datos Principales» y el de eliminar se ve. Si el usuario es de Calidad, se preselecciona como responsable (`:610-612`).
5. `Filtrar` (`:615`) y `Me.Caption = m_ObjEntorno.TituloUsuarioConectado` (`:621`), que vuelve a pisar el rótulo del paso 1.

### Normal: filtrado (`Filtrar`, `:71-143`)

Descarta el proyecto seleccionado y el activo, deja la botonera en el estado inicial (`:87-92`) y pone la cabecera de ocho columnas (`:96`). Consulta `Constructor.getProyectosBusqueda1` con los cuatro criterios (`:98-101`), recorre el diccionario y compone una fila de ocho campos por proyecto (`:117-127`), sustituyendo por `-----` las fechas vacías o que no son fecha (`:108-116`). Termina seleccionando la primera fila y disparando a mano el manejador del clic (`:132-133`). No escribe nada: es una lectura repetible.

### Normal: selección (`ListaFiltrados_Click`, `:671-778`)

1. Apaga la etiqueta y la imagen del último informe y vacía la ruta (`:679-681`); si la lista tiene exactamente dos filas, selecciona la primera (`:682-685`).
2. Sin identificador, **anula seis globales de contexto** y devuelve la botonera al estado inicial (`:687-708`).
3. Carga el proyecto si no lo tiene o si ha cambiado (`:712-729`); si es el mismo, sale pronto sin rehacer nada (`:723-728`).
4. `EstablecerBotonera` (`:789-831`) habilita anexos, riesgos, edición y evolución, y habilita eliminar **solo si no es técnico** (`:817-821`).
5. Busca la última edición publicada, guarda la ruta de su informe (`:742-751`) y con ella enciende la imagen de «ver último informe» (`:755-761`).

### Normal: entrar a los riesgos (`:422-468`)

Fija el proyecto activo y elige edición: la activa si existe, y si no la última (`:439-444`); cierra `FormRiesgosGestion` si estaba abierta y la vuelve a abrir (`:446-452`).

### Alternativo: alta, edición y borrado

El alta (`:379-421`) bloquea a técnicos con aviso explícito (`:386-391`), marca `m_EsAlta = Sí`, cierra y abre `FormGestionRiesgos`, y **guarda referencia a ese hijo** (`:406-412`), que dispara `m_FormGestionRiesgos_Alta` para refiltrar (`:309-312`). La edición (`:209-248`) hace lo mismo con `m_EsAlta = No` y sin guardar referencia. El borrado (`:249-308`) pide confirmación con el botón por defecto en **No** (`:263-268`), borra, recalcula las tareas de Calidad (`:277-280`) y refiltra (`:282-285`): es **la única operación destructiva accesible desde esta pantalla**.

### Flujo de error

Todos los manejadores de acción repiten el mismo patrón (`:155-179`, `:210-247`, `:252-307`, `:324-345`, `:383-420`, `:425-467`, `:477-520`, `:646-669`, `:674-777`): error distinto de 1000, correo al administrador y crítico; 1000 exacto, advertencia con el texto ya preparado. `Form_Load` cierra el formulario incondicionalmente (`:640`); los cuatro manejadores de limpiar criterio no capturan error (`:347-378`).

## Reglas, validaciones y cálculos

- El rol se comprueba con dos banderas distintas: `EsTecnico` para el alta y el borrado, `EsCalidad` solo para preseleccionar el responsable (`:610`). Para el alta, la comprobación real es la negativa.
- La botonera se construye en dos pasos: `EstablecerBotoneraAlInicio` (`:29-69`) la apaga entera y `EstablecerBotonera` (`:789-831`) enciende lo que corresponde. Toda entrada a la pantalla pasa por el primero.
- Las fechas vacías o no válida se sustituyen por `-----` (`:110-116`) en vez de dejar la celda en blanco: es convención de presentación, no de dato.
- El filtro se dispara al confirmar, no al teclear: los tres combos filtran en `BeforeUpdate` y el texto solo con Intro (`:314-321`). El filtro es conjunción de cuatro criterios, con «Sí» como valor inicial de activos.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Búsqueda de proyectos | `Constructor.getProyectosBusqueda1` | `:98-101` |
| Proyecto concreto | `Constructor.getProyecto` | `:216`, `:256`, `:431`, `:482`, `:717`, `:800` |
| Datos del proyecto | `Proyecto`: cierres, publicación, edición activa y última | `:108-120`, `:439-444`, `:742-748` |
| Borrado | `Proyecto.Borrar` | `:274` |
| Tareas de Calidad | `EstablecerTareasCalidad` | `:277` |
| Gráficas de evolución | `Proyecto.ObtenerGraficasRiesgos`, fichero Excel | `:494`, `:504` |
| Informe de la última edición | `EdicionUltimaPublicada.URLArchivoInforme` | `:747` |
| Autorizados y usuarios de Calidad | `m_ObjEntorno.ColJPs`, `EstablecerComboUsuariosCalidad` | `:536`, `:599` |

## Coexistencia

Pendiente de la capa de coexistencia, sin enlace desde esta ficha. Aquí se concentran dos de las ataduras más fuertes del legacy con el sistema de ficheros del cliente: la evolución de riesgos se abre con el visor de Excel del sistema (`:504-506`) y el último informe publicado se abre en local (`:649-651`). Mientras convivan, ambos dependen de que el fichero exista en la máquina del usuario.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-307.** `m_FormAnexos` se declara `WithEvents` (`:23`) pero **nunca se le asigna un formulario**: no hay ningún `Set m_FormAnexos` en el archivo, así que `m_FormAnexos_AnexoAñadido` y `m_FormAnexos_AnexoEliminado` (`:832`, `:873`) **no pueden ejecutarse**. Son código muerto: el formulario de anexos se sigue pudiendo abrir, pero ningún cambio en él refresca esta lista.
- **H-GR-308.** `ComandoAnexar` se muestra, se habilita (`:36`, `:811`) y tiene globo explicativo (`.form.txt:1474`), pero **no existe `ComandoAnexar_Click`**: el botón no hace nada. Conviene confirmar con el propietario si la entrada a anexos se perdió al migrar.
- **H-GR-309.** `ActualizarProyecto` (`:909-956`) construye la línea actual con **seis** columnas (`:929-931`) y la nueva con **seis** campos en otro orden (`:934-936`), pero la fila real tiene **ocho** columnas tal como las compuso `Filtrar` (`:96`, `:124-127`). El `Replace` sobre el `RowSource` (`:941`) **no puede encontrar la línea**: el refresco de una fila editada no ocurre y, si ocurriera, dejaría la fila incompleta.
- **H-GR-310.** `Filtrar` hace `lst.Selected(1) = True` sin comprobar que haya resultados (`:132`). Con cero proyectos, solo queda la cabecera, la selección falla y el error sale como mensaje de método interno (`:137-139`) en lugar de «no hay resultados».
- **H-GR-311.** `ListaFiltrados_Click` autosecciona **solo** con exactamente dos filas (`:682`), contando la cabecera. Con tres o más, tras filtrar nada queda seleccionado, aunque `Filtrar` acabara de seleccionar la primera: la botonera depende de una cuenta exacta de filas.
- **H-GR-312.** El diseño guarda `RowSource` de muestra con **nombres de personas concretas** en la lista y en los tres desplegables (`.form.txt:697`, `:794`, `:1548`), y el rótulo de la ventana incluye el nombre del usuario que guardó el diseño (`.form.txt:39`). Dato personal incrustado en el artefacto; no se reproduce aquí. Conviene señalarlo al propietario.
- **H-GR-313.** Quedan llamadas de cronómetro y `Debug.Print` en la carga (`:588-598`): ejecutan trabajo y escriben en la ventana inmediato aunque en una versión compilada no se vean.
- **H-GR-314.** Al deseleccionar se anulan seis globales de contexto (`:690-695`) pero **no** `m_ObjProyectoActivo`: queda vivo el proyecto que se estaba editando aunque no haya nada seleccionado.
- **H-GR-315.** El Excel de evolución se abre solo si el fichero existe (`:503-506`) y, si no existe, no se avisa: la barra de progreso se cierra como si hubiera ido bien. El informe de la última edición solo se abre si hay ruta (`:648`), aunque la imagen solo se muestre cuando la hay (`:755-761`).
- **H-GR-316.** `EstablecerBotonera` (`:789-831`) inicializa `p_Error` (`:795`) pero internamente usa `m_Error` (`:800-806`): mezcla de dos variables de error en el mismo método.
- **H-GR-317.** Los cuatro manejadores de limpiar criterio no capturan error (`:347-378`) y no siguen la convención del control que limpian: el de autorizados es `ComandoLimpiarAutorizado` (`:355`) y controla `ComboJefeProyecto`.
- **H-GR-318.** `ComandoAyuda` tiene `Caption = "Salir"` (`.form.txt:597`), igual que `ComandoSalir` (`.form.txt:2246`): dos botones con el mismo texto visible y distinta función.
- **H-GR-319.** El borrado elimina el proyecto completo sin resumen de lo que se lleva (`:263-280`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: sustituir el listado por una tabla con búsqueda y filtros combinables en el servidor; dejar de anular globals al deseleccionar y pasar el contexto por la ruta; confirmar el borrado con un resumen de lo afectado; publicar la evolución y el último informe desde rutas autenticadas en lugar de abrirlos en el equipo del usuario; resolver por rol en el servidor los botones de alta y borrado en lugar de ocultarlos o deshabilitarlos en cliente.

## Criterios de aceptación

Propuestos, no aprobados:

1. El listado muestra por defecto solo los proyectos activos y filtra por los cuatro criterios a la vez.
2. Según el rol se ven unos botones u otros, y el alta y el borrado no están disponibles para el técnico.
3. Con un único resultado, ese resultado queda seleccionado y accionable.
4. Un proyecto sin resultados muestra un estado vacío explícito, no un error interno.
5. Tras editar un proyecto, el listado muestra la fila nueva.
6. Borrar pide confirmación y el listado queda coherente.

## Pendientes

- Si `ComandoAnexar` debe abrir el formulario de anexos y qué manejador se perdió, si es que se perdió alguno.
- Qué garantiza que `m_ObjProyectoActivo` no se solape entre instancias, dado que es una global y la ventana se abre desde dos menús.
- Qué criterio aplica `Constructor.getProyectosBusqueda1` a la palabra clave (`.form.txt:755` dice «cualquier parte del nemotécnico»).
- Si el Excel de evolución es un artefacto estático o se regenera en cada petición.
- Cómo se comporta `EstablecerTareasCalidad` (`:277`) tras un borrado con tareas de otros usuarios abiertas.