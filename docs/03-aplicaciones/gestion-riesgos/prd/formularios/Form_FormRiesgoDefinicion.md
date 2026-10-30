# Form_FormRiesgoDefinicion

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormRiesgoDefinicion.cls` (703 líneas) y `src/forms/FormRiesgoDefinicion.form.txt` (2362 líneas) · **PRD:** [README](../README.md)

Pestaña de definición del riesgo cuando el proyecto usa riesgo de biblioteca. Es el formulario más grande de la cadena: elige un riesgo tipo de una lista filtrada, arrastra su código y su descripción al riesgo, y valida la priorización contra el resto de riesgos de la edición. Las citas son `Form_FormRiesgoDefinicion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormRiesgo` | Subformulario: el contenedor fija `Me.NavGeneral.NavigationTargetName = "FormRiesgoDefinicion"` y lo monta en `Me.FrmDetalle` | `Form_FormRiesgo.cls:287`, `:307` |

No hay ningún `DoCmd.OpenForm` hacia él en el clon. El contenedor también lo refresca explícitamente en su botón de actualizar (`Form_FormRiesgo.cls:48-49`). No abre ningún formulario y no se cierra a sí mismo en ningún punto: su vida depende del contenedor.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Refresca la pantalla | `:125-151`, `.form.txt:221` |
| `ListaFiltrados` | lista | `OnClick` → `ListaFiltrados_Click` | Selección del riesgo tipo de biblioteca | `:154-235`, `.form.txt:778-818` |
| `ComboFamilia` | lista desplegable | `BeforeUpdate` → `ComboFamilia_BeforeUpdate` | Filtra la biblioteca por familia | `:47-78`, `.form.txt:1107-1111` |
| `ComandoLimpiarFamilia` | botón | `OnClick` → `ComandoLimpiarFamilia_Click` | Vacía la familia y recarga la lista | `:30-36`, `.form.txt:593-595` |
| `PalabraClave` | cuadro de texto | `OnExit` → `PalabraClave_Exit` | Filtra la biblioteca por palabra clave | `:598-629`, `.form.txt:1144-1145` |
| `ComandoLimpiarPalabraClave` | botón | `OnClick` → `ComandoLimpiarPalabraClave_Click` | Vacía la palabra clave y recarga | `:38-45`, `.form.txt:1172-1174` |
| `Descripcion` | cuadro de texto | `OnExit` → `Descripcion_Exit` | Descripción heredada del riesgo tipo | `:80-87`, `.form.txt:917-918` |
| `CausaRaiz` | cuadro de texto | `BeforeUpdate` → `CausaRaiz_BeforeUpdate` | Causa raíz; solo en proyectos con riesgo de biblioteca | `:20-28`, `.form.txt:949-950` |
| `EntidadDetecta` | lista desplegable | `BeforeUpdate` → `EntidadDetecta_BeforeUpdate` | Entidad que detecta el riesgo y puebla «detectado por» | `:100-108`, `.form.txt:988-992` |
| `DetectadoPor` | lista desplegable | `BeforeUpdate` → `DetectadoPor_BeforeUpdate` | Persona o empresa que detecta | `:89-96`, `.form.txt:1031-1035` |
| `FechaDetectado` | cuadro de texto | `BeforeUpdate` → `FechaDetectado_BeforeUpdate` | Fecha de detección | `:111-119`, `.form.txt:1067-1069` |
| `Origen` | lista desplegable | `BeforeUpdate` → `Origen_BeforeUpdate` | Origen del riesgo | `:588-595`, `.form.txt:1362-1371` |
| `Priorizacion` | cuadro de texto | `OnChange` → `Priorizacion_Change` | Valida la priorización contra el resto de riesgos | `:632-702`, `.form.txt:1403-1405` |
| `TextoPendienteRetipificacion` | cuadro de texto | — | Rótulo «Pendiente de Calidad» en retipificación | `:203-204`, `.form.txt:886` |
| `lblCausaRaiz` | etiqueta | — | Rótulo de la causa raíz | `:433`, `.form.txt:1582` |

El diseño conserva siete etiquetas sin código (`Etiqueta232`, `Etiqueta42` a `Etiqueta48`; `.form.txt:1432-1642`). Los bloques `Begin` sin nombre de `.form.txt:415-552` son valores por defecto de propiedades.

## Flujos

### Apertura (`Form_Load` → `EstablecerDatos`, `:125-522`)

1. `Form_Load` pone el reloj y llama a `EstablecerDatos` (`:132`), propagando el error (`:134`).
2. `EstablecerDatos` deshabilita ocho controles de golpe (`:332-339`) y exige edición activa (`:342-344`).
3. Si `Form_FormRiesgo.m_PermitidoEditar` está `Empty`, lo **calcula y escribe en el contenedor** (`:346-355`).
4. En alta sin permiso aborta con «Es un alta y el usuario no tiene permitida esa operación» (`:363-365`); en alta crea `New riesgo` y su copia inicial a `Nothing` (`:368-370`); en edición recarga el riesgo (`:380`) y su copia (`:399`).
5. Con permiso rehabilita `CausaRaiz`, `Origen`, `ListaFiltrados`, `ComboFamilia`, las dos palabras clave y sus botones (`:412-419`); deja `Descripcion` deshabilitado (`:435`).
6. Llena origen y familia (`:421`, `:429`), muestra el rótulo de causa raíz (`:433-434`) y, si el riesgo ya tiene código de biblioteca, recupera el riesgo tipo y lo selecciona en la lista (`:437-454`).
7. Con `RiesgoParaRetipificar = Sí` muestra el aviso y habilita la descripción (`:455-458`); si no, lo oculta (`:462`).
8. Vuelca los ocho campos del riesgo (`:467-489`), comprueba la priorización con `Constructor.EsPriorizacionCorrecta` (`Constructor.bas:6984`) (`:493`, `:496`) y, con edición activa, marca bordes y colorea (`:501-506`).

### Normal: elegir riesgo tipo (`ListaFiltrados_Click`, `:154-235`)

1. Si `CausaRaiz` no está visible, sale (`:164-165`): la lista solo trabaja en modo biblioteca.
2. Con la lista vacía, suelta la selección y sale (`:172-176`).
3. Recupera el riesgo tipo con `Constructor.getRiesgoBiblioteca` (`:182`) y, si cambió de riesgo, lo vuelve a pedir (`:186-188`).
4. En retipificación muestra el rótulo, habilita la descripción y fija «Pendiente de Calidad» (`:203-205`); en modo normal oculta el rótulo y la deja deshabilitada (`:207-208`).
5. Vuelca la descripción (`:210`), marca bordes si la edición está activa (`:211-212`), y escribe **directamente en el objeto riesgo**: `.CodRiesgoBiblioteca = .codigo` (`:214`) y `.Descripcion = Nz(Me.Descripcion, "")` (`:218`).

### Alternativo: filtrar la biblioteca

- `ComboFamilia_BeforeUpdate` (`:47-78`), `PalabraClave_Exit` (`:598-629`) y los dos botones de limpiar (`:30-36`, `:38-45`) recomponen la lista. Los tres comparan con un valor memorizado (`m_FamiliaUltima` `:15`, `m_PalabraClaveUltima` `:16`) y solo recargan cuando el valor ha cambiado de verdad (`:54`, `:605`).
- `EstablecerListaRiesgosBiblioteca` (`:237-284`) pide la colección con `Constructor.getRiesgosBiblioteca` (`Constructor.bas:1461`) (`:251`) y monta cada fila sustituyendo `;` por `:` y **recortando la descripción a 45 caracteres** (`:259-266`). Si la lista está habilitada y queda un único elemento, lo selecciona y llama a su propio manejador (`:271-274`).

### Alternativo: priorización no válida (`Priorizacion_Change`, `:632-702`)

1. Valida con `EsPriorizacionCorrecta`, por edición si es alta y por riesgo si es edición (`:647`, `:651`).
2. Colorea las pestañas (`:657`) y sale si la priorización está vacía (`:659-662`).
3. Si es incorrecta y hay copia inicial con valor numérico, **restaura la priorización anterior** en el control y en el objeto (`:668-673`); si no, la vacía (`:677`).
4. Avisa «Esa priorización no es correcta…» y lanza el error controlado (`:681-682`).

### De error

`ComboFamilia_BeforeUpdate` (`:69-76`), `Form_Load` (`:142-149`), `ListaFiltrados_Click` (`:226-233`), `PalabraClave_Exit` (`:620-627`) y `Priorizacion_Change` (`:693-700`) llevan el patrón común: quitar el reloj y, si `Err.Number <> 1000`, componer el mensaje, escribirlo al administrador (`Funciones Generales.bas:964`) y avisar con diálogo crítico; si es 1000, solo advertencia. Ninguno cierra. Los seis manejadores de `BeforeUpdate` y `Exit` que solo pintan bordes y colores usan `On Error Resume Next` **sin manejador** (`:21`, `:81`, `:90`, `:101`, `:112`, `:589`).

## Reglas, validaciones y cálculos

- **Solo en proyectos con riesgo de biblioteca.** El contenedor elige esta pestaña según `Proyecto.RequiereRiesgoDeBibliotecaCalculado` (`Proyecto.cls:667`); aquí la prueba indirecta es que `CausaRaiz` esté visible (`:164`).
- **La descripción no se escribe a mano.** Se deja deshabilitada (`:435`) salvo en retipificación (`:458`): su valor viene del riesgo tipo elegido en la lista.
- **La priorización se valida contra el resto de riesgos** de la edición con `Constructor.EsPriorizacionCorrecta` (`Constructor.bas:6984`), al cambiar el campo (`:647`, `:651`) y también al cargar (`:493`, `:496`).
- **«Detectado por» depende de la entidad.** `RellenarComboDetectadoPor` (`:524-585`) solo puebla la lista si la entidad es «Persona» (`:549-560`) o «Empresa» (`:565-573`): para «Persona» usa `Proyecto.CadenaNombreAutorizados` (`Proyecto.cls`), para «Empresa» los suministradores de la edición (`:566`).
- **El filtro de biblioteca es doble:** familia y palabra clave, con los valores anteriores memorizados para no recargar en cada pulsación (`:54`, `:605`).
- **Sin acceso directo a datos:** sin `getdb()`, `Recordset`, `Open`, `Execute`, SQL, `DLookup` ni transacciones; todo pasa por `Constructor` y por las clases.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Riesgo sobre el que se opera | `m_ObjRiesgoActivo` (`Variables Globales.bas:80`) y su copia `m_ObjRiesgoAlInicio` (`:83`) | `:370`, `:380`, `:399`, `:436-498`, `:668-671` |
| Edición activa | `m_ObjEdicionActiva` (`Variables Globales.bas:79`) | `:24`, `:83`, `:211`, `:342`, `:369` |
| Permiso de edición | `Form_FormRiesgo.m_PermitidoEditar` (`Form_FormRiesgo.cls:22`), **escrito aquí** | `:346`, `:349-355`, `:363`, `:410` |
| Riesgo tipo seleccionado | `Constructor.getRiesgoBiblioteca` (`Constructor.bas:2885`), `Constructor.getRiesgosBiblioteca` (`Constructor.bas:1461`) | `:182`, `:188`, `:251`, `:440`, `:446` |
| Riesgo completo | `Constructor.getRiesgo` (`Constructor.bas:1827`), `New riesgo` | `:380`, `:368-369` |
| Validación de priorización | `Constructor.EsPriorizacionCorrecta` (`Constructor.bas:6984`) | `:493`, `:496`, `:647`, `:651` |
| Personas autorizadas | `m_ObjEdicionActiva.Proyecto.CadenaNombreAutorizados` | `:551` |
| Suministradores de la edición | `m_ObjRiesgoActivo.Edicion.ColSuministradores` | `:566` |
| Escrituras directas en el objeto | `.CodRiesgoBiblioteca` (`:214`), `.Descripcion` (`:218`), `.Priorizacion` (`:671`) | `:214`, `:218`, `:671` |

Este formulario **no escribe en la base de datos**: solo prepara el objeto riesgo y lo deja en memoria para que el contenedor lo registre. El volcado definitivo lo hace `RellenarObjeto` del contenedor (`Form_FormRiesgo.cls:502-509`).

## Coexistencia

Pendiente de `../coexistencia.md`. Dos cosas que la versión web deberá decidir: la lista de riesgos tipo pasa de una colección en memoria a una consulta paginada, y la validación de priorización, hoy recalculada en el cliente, pasa al servidor para que no dependa del estado de la sesión.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-156.** `EstablecerDatos` **escribe el permiso del contenedor**: `Form_FormRiesgo.m_PermitidoEditar` se calcula y se guarda aquí (`:349-355`) además de leerse (`:346`). Un subformulario modifica el estado de su propio padre, de modo que el permiso final depende del orden en que se carguen las pestañas.
- **H-GR-157.** `EstablecerDatos` tiene un bloque redundante con `m_EsAlta = No` (`:393`) que repite la validación y la recarga de la copia inicial (`:399`) de lo ya hecho en la rama de edición.
- **H-GR-158.** Los seis manejadores que solo pintan bordes y colores usan `On Error Resume Next` sin manejador (`:21`, `:81`, `:90`, `:101`, `:112`, `:589`). Cualquier fallo al recalcular la obligatoriedad o el color se traga en silencio y la pantalla queda con indicadores que no corresponden a los datos.
- **H-GR-159.** `EstablecerListaRiesgosBiblioteca` **recorta la descripción a 45 caracteres y cambia los punto y coma por dos puntos** para montar la fila (`:259-266`). Dos riesgos tipo que compartan inicio son indistinguibles para el usuario, y la selección posterior se resuelve por identificador (`:297`), no por lo que se ve.
- **H-GR-160.** `ListaFiltrados_Click` escribe directo en el objeto riesgo (`.CodRiesgoBiblioteca` `:214`, `.Descripcion` `:218`) sin pasar por el volcado del contenedor. Hay dos rutas de escritura al mismo campo, y la de aquí actúa en cuanto se elige una fila, antes de guardar.
- **H-GR-161.** `Priorizacion_Change` corrige el campo **antes** de avisar: restaura el valor anterior (`:670-671`) o lo vacía (`:677`) y después lanza el error controlado con el mensaje (`:681-682`). El usuario ve una advertencia sobre un campo que ya ha cambiado a sus espaldas.
- **H-GR-162.** `RellenarComboDetectadoPor` propaga a su propio error el de la carga de nombres autorizados (`m_ObjEdicionActiva.Proyecto.Error`, `:552`). Un fallo ajeno a esta pantalla se muestra como si fuera suyo, y el mensaje resultante no encaja con lo que el usuario estaba haciendo.
- **H-GR-163.** Los dos botones de limpiar (`:30-36`, `:38-45`) y `EstablecerListaRiesgosBiblioteca` recargan la lista **sin comprobación de error**: si `Constructor.getRiesgosBiblioteca` falla, el `For Each` no se ejecuta y la lista queda vacía sin aviso al usuario.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: descripción completa de la fila con tooltip en lugar de recorte; escritura única del objeto desde el contenedor, sin doble vía; permiso calculado en un único sitio y pasado por parámetro; validación de priorización en el servidor; lista de entidades y personas como recursos consultables, no como `RowSource` en memoria.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir la pestaña en modo biblioteca muestra el riesgo con sus ocho campos y la lista de riesgos tipo.
2. Elegir un riesgo tipo arrastra su código y su descripción al riesgo en curso.
3. En un proyecto sin riesgo de biblioteca, esta pestaña no se monta.
4. Los filtros de familia y palabra clave recargan la lista solo cuando cambian de verdad.
5. Una priorización incorrecta no se guarda y el usuario recibe el aviso con el campo corregido.
6. Si el usuario no tiene permiso, los campos y el botón de guardar no están disponibles.

## Pendientes

- Qué devuelve exactamente `Constructor.getRiesgosBiblioteca` (`Constructor.bas:1461`) y si el truncado a 45 caracteres (`:259-266`) es por diseño de la interfaz o por el ancho del control.
- Cómo calcula `Constructor.EsPriorizacionCorrecta` (`Constructor.bas:6984`) la prioridad relativa y qué reglas de desempate aplica.
- Quién más escribe `Form_FormRiesgo.m_PermitidoEditar` y en qué orden se cargan las pestañas (**H-GR-156**).
- Si `RellenarBordeCamposObligatorios` (`Funciones Generales.bas:7095`) y `EstablecerColorPestañasRiesgo` (`Funciones Generales.bas:7372`) actúan sobre este subformulario o sobre el formulario activo, porque se les pasa `Me` desde un subformulario (**H-GR-158**).