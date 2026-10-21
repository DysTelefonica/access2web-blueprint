# Form_FormVideosGestion

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormVideosGestion.cls` (203 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Contenedor de la gestión de vídeos de ayuda: un árbol ActiveX (`MSComctlLib.TreeCtrl`) con una rama por aplicación y una hoja por vídeo, y un subformulario `FormDetalle` que muestra el alta (nodo de aplicación) o la edición (nodo de vídeo). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); no escribe en la base de datos. Las citas son `Form_FormVideosGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormMenuPrincipalAdmin` | `SUBMENU3_Click` anula `m_ObjUsuarioActivo` y fija `FormDetalle.SourceObject = "FormVideosGestion"` | `Form_FormMenuPrincipalAdmin.cls:268-275` |
| `Form_FormMenuPrincipalCalidad` | `SUBMENU1_Click`, mismo patrón | `Form_FormMenuPrincipalCalidad.cls:271-278` |
| `EliminarNodo`, `CrearNodo` | llaman a `Form_FormVideosGestion.Arbol_NodeClick` tras modificar el árbol | `Funciones Genreales.bas:2687, 2732` |

Abre en `FormDetalle` los subformularios [Form_FormVideosGestionArbolAplicacion](Form_FormVideosGestionArbolAplicacion.md) (`:56-57`) y `Form_FormVideosGestionArbolVideo` (`:62-63`).

Estado compartido (globales de `Variables Globales.bas`): `m_ObjAplicacionActiva` (`:101`), `m_ObjVideoActivo` (`:102`) y `NodoActivo` (`:103`). Los subformularios leen las dos primeras.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Arbol` | TreeView ActiveX | `NodeClick` → `Arbol_NodeClick` (público) | Resuelve vídeo o aplicación del nodo y carga el subformulario | `:23-81`, `.form.txt:836, 984-985` |
| `ListaImagenes` | ImageList ActiveX, invisible | — | Iconos del árbol | `:141`, `.form.txt:997, 1010` |
| `FormDetalle` | subformulario | — | Sin `SourceObject` de diseño; se fija por código | `:49-62`, `.form.txt:2117-2133` |
| `Sep` | etiqueta usada como separador | `MouseDown`, `MouseMove`, `MouseUp` | Redimensiona árbol y subformulario arrastrando | `:154-200`, `.form.txt:2105-2108` |
| `ComandoActualizarContador` (rótulo «Contador», ayuda «Actualizar Tareas») | botón | `OnClick` (público) | Recarga el árbol con `EstablecerDatos` | `:83-107`, `.form.txt:2144-2147` |

## Flujos

### Apertura (`Form_Load`, `:109-133`)

1. `EstablecerDatos` llama a `CargarArbol` con el árbol y la lista de imágenes (`:141`) y coloca `Sep` junto a `FormDetalle` (`:145`).
2. `CargarArbol` rellena la lista de imágenes con el icono de cada aplicación y tres iconos fijos de carpeta y vídeo; si falta uno fijo, error (`Funciones Genreales.bas:2522, 2464-2490`).
3. Vacía el árbol y añade un nodo raíz `A<IDAplicacion>` por aplicación de `m_ObjEntorno.ColAplicaciones`, salvo la de ID 1, las que están en pruebas y las que no tienen icono de árbol (`Funciones Genreales.bas:2550, 2559-2585`).
4. Bajo cada aplicación, un nodo hijo `V<IDVideo>` por vídeo de `ColVideosDeAyuda` (`Funciones Genreales.bas:2590-2600`).

### Selección de nodo (`Arbol_NodeClick`, `:23-81`)

1. Guarda el nodo en `NodoActivo` (`:30`) y busca el vídeo con `getVideoByNode`, que solo resuelve claves que empiezan por «V» (`:32`; `Constructor.bas:1302-1307`).
2. Con vídeo, la aplicación activa es la del vídeo (`:36-37`); sin vídeo, `getAplicacionByNode` resuelve claves «A» (`:43`; `Constructor.bas:1335-1338`).
3. Si no resuelve ninguno, vacía `FormDetalle` y sale (`:48-51`).
4. Vacía `FormDetalle` y carga `FormVideosGestionArbolAplicacion` + `CargarDatos` (alta) o `FormVideosGestionArbolVideo` + `MENU1_Click` (edición) (`:52-67`).

### Errores

Los handlers acumulan el error en `m_Error`, envían `CorreoAlAdministrador` salvo el error interno 1000 y muestran `MsgBox` crítico (`:72-80, 100-106, 126-132`). `EstablecerDatos` solo devuelve el texto en `p_Error` (`:148-151`).

## Reglas, validaciones y cálculos

- **Vídeos visibles en el árbol según el rol** (`Constructor.getVideosDeAplicacion` con `p_DeAyuda = Sí`): solo vídeos con categoría «Ayuda»; administrador, todos; usuario de Calidad, solo `ParaCalidad = 'Sí'`; resto, solo `ParaTecnicos = 'Sí'` (`Constructor.bas:1385-1438`).
- **Aplicaciones excluidas:** ID 1, `EnPruebasCalculado = Sí` y sin `NombreIconoArbol` (`Funciones Genreales.bas:2560-2583`).
- **Permisos:** el formulario no comprueba el rol. Se alcanza desde los menús de administración y de Calidad, que `LoginCorrecto` elige según `EsAdministradorCalculado` y `EsUsuarioCalidadCalculado` (`Variables Globales.bas:724-729`).
- **Separador:** ancho mínimo del árbol 60 twips y máximo la mitad del formulario (`:171-173, 192-193`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Aplicaciones | `TbAplicaciones` (caché de `Entorno`) | DAO, tabla entera | `Constructor.bas:110-111`; `Entorno.cls:425-429` |
| Vídeos de ayuda por aplicación | `TbVideos` ⋈ `TbVideosCategorias` ⋈ `TbCategorias` | SQL dinámico concatenado | `Constructor.bas:1390-1400, 1446` |
| Vídeo por nodo | `TbVideos` | SQL dinámico concatenado | `Constructor.bas:1259-1261` |
| Iconos | ficheros de icono en el servidor | E/S de ficheros | `Funciones Genreales.bas:2448-2490` |

Sin escrituras.

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: filtros por categoría con literal de texto (`NombreCategoria='Ayuda'`) e IDs concatenados (`Constructor.bas:1396`); el árbol depende de iconos en una carpeta de red, sin equivalente en PostgreSQL.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El árbol solo carga vídeos de categoría «Ayuda» (`Funciones Genreales.bas:2568`; `Constructor.bas:1396`): un vídeo solo de «Encuesta», o sin categoría, no aparece y no puede editarse ni borrarse desde aquí.
- `ComandoActualizarContador` recarga el árbol desde cachés que nadie invalida (`Entorno.cls:425-427`; `Aplicacion.cls:1238-1240`): tras borrar un vídeo, el nodo podría reaparecer y un alta reciente desaparecer.
- `Arbol_NodeClick` comprueba `m_Error` tras `MENU1_Click`, que no lo modifica (`:63-66`): los errores de carga de la edición no llegan aquí.
- Rótulo «Contador» y ayuda «Actualizar Tareas» heredados de otro formulario (`.form.txt:2145`).
- `SepW` se declara `Integer` con valor `60.1` (`:19`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: listado de vídeos por aplicación con filtro de categoría (ayuda y encuesta); la aplicación y el vídeo activos como parámetros de ruta, no globales.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un administrador ve todas las aplicaciones activas, no en pruebas, con sus vídeos de ayuda.
2. Un usuario de Calidad solo ve los vídeos de ayuda dirigidos a Calidad.
3. Seleccionar una aplicación abre el alta de vídeo; seleccionar un vídeo abre su edición.
4. Refrescar la lista refleja altas y bajas recientes.

## Pendientes

- Si Access enlaza el evento `NodeClick` del ActiveX con el `Public Sub Arbol_NodeClick` (`:23`) igual que con uno privado.
- Si `Form_FormVideosGestionArbolAplicacion` y `Form_FormVideosGestionArbolVideo` referencian la instancia del subformulario o crean otra (`:57, 63`).
- Contenido de `URLIconoParaArbol` y `AddImagen` (`Funciones Genreales.bas:2448, 2455`).
- Uso de `m_ObjUsuarioActivo`, que anulan los menús al abrir (`Form_FormMenuPrincipalAdmin.cls:270`).
