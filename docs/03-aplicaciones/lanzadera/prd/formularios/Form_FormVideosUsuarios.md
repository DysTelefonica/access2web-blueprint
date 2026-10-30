# Form_FormVideosUsuarios

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormVideosUsuarios.cls` (171 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Consulta de vídeos de ayuda para el usuario: un árbol (`TreeView`) con las aplicaciones y sus vídeos de la categoría «Ayuda» a la izquierda, un separador redimensionable y, a la derecha, el detalle del vídeo seleccionado. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); lee vídeos por DAO y no escribe. Las citas son `Form_FormVideosUsuarios.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Ninguno activo | Los tres menús principales tienen comentada la carga de `FormVideosUsuarios`; la opción «Videos-Guías» muestra «En construcción» | `Form_FormMenuPrincipalAdmin.cls:187-192`, `Form_FormMenuPrincipalCalidad.cls:234-239`, `Form_FormMenuPrincipalUsuario.cls:221-226`, `Form_FormMenuPrincipalUsuario.form.txt:590-591` |

Aloja `FormVideosArbolVideo` en `FormDetalle` (`.form.txt:2822-2823`; `:55-57`), que a su vez aloja [Form_FormVideoDetalle](Form_FormVideoDetalle.md). En el estado actual del código el formulario es **inalcanzable** desde la interfaz.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` → `Form_Load` | Llama a `EstablecerDatos` con reloj de arena | `:76-103`, `.form.txt:83` |
| `Arbol` | ActiveX `MSComctlLib.TreeCtrl.2` | `NodeClick` → `Arbol_NodeClick` | Carga el vídeo del nodo en `m_ObjVideoActivo` y refresca el detalle | `:24-72`, `.form.txt:294, 538-539` |
| `ListaImagenes` | ActiveX `MSComctlLib.ImageListCtrl.2` (oculto) | — | Iconos del árbol | `.form.txt:550, 561, 2789-2790` |
| `Sep` | etiqueta separadora | `OnMouseDown`, `OnMouseMove`, `OnMouseUp` | Redimensiona árbol y detalle arrastrando | `:123-170`, `.form.txt:2804-2807` |
| `FormDetalle` | subformulario | — | Aloja `FormVideosArbolVideo` | `.form.txt:2822-2823` |

El evento `NodeClick` del ActiveX se enlaza por nombre del procedimiento, sin propiedad en el `.form.txt`.

## Flujos

### Apertura (`Form_Load` y `EstablecerDatos`, `:76-122`)

1. `CargarArbol` rellena la lista de imágenes y el árbol (`:111`; `Funciones Genreales.bas:2498-2622`).
2. `RellenarListaImagenes` añade el icono de cada aplicación salvo la de ID 1, la de `NombreBase = "Registro"` y las que están en pruebas, más tres iconos fijos (carpeta abierta, carpeta cerrada, vídeo); si falta uno fijo, error (`Funciones Genreales.bas:2436-2490`).
3. El árbol se vacía (`Funciones Genreales.bas:2550`) y, por cada aplicación de `m_ObjEntorno.ColAplicaciones` (lee `TbAplicaciones`, `Constructor.bas:110-111`), omite la de ID 1, las que están en pruebas y las que no tienen icono de árbol (`Funciones Genreales.bas:2560-2583`).
4. Añade el nodo de aplicación con clave `A<ID>` y, bajo él, un nodo `V<IDVideo>` por cada vídeo de `ColVideosDeAyuda` (`Funciones Genreales.bas:2574-2599`).
5. Coloca el separador junto al subformulario (`:115`).

### Selección de nodo (`Arbol_NodeClick`, `:24-72`)

1. Nodo de vídeo (clave con «V»): extrae el ID con `Split` y `CInt` (`:30-32`).
2. Sin vídeo activo, lo carga con `Constructor.getVideo` (`:33-37`); con otro vídeo activo distinto, lo carga y vacía `FormDetalle` (`:39-45`).
3. Nodo de aplicación: `m_ObjVideoActivo = Nothing` (`:49-51`) y vacía `FormDetalle` (`:52-53`).
4. Con vídeo activo, recarga `FormVideosArbolVideo` si el subformulario está vacío (`:55-57`) y llama a `Form_FormVideoDetalle.EstablecerDatos` (`:59`).

### Errores

Los errores distintos de 1000 se envían con `CorreoAlAdministrador`; siempre se muestra `MsgBox` crítico (`:63-71`, `:94-102`). `EstablecerDatos` solo devuelve el texto en `p_Error` (`:118-121`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Solo vídeos de la categoría «Ayuda» | `Aplicacion.ColVideosDeAyuda` | `classes/Aplicacion.cls:1245`, `Constructor.bas:1395-1396` |
| Administrador: todos los vídeos de ayuda de la aplicación | `getVideosDeAplicacion` | `Constructor.bas:1385-1396` |
| Calidad: solo `ParaCalidad = 'Sí'` | `getVideosDeAplicacion` | `Constructor.bas:1403, 1413-1415` |
| Resto: solo `ParaTecnicos = 'Sí'` | `getVideosDeAplicacion` | `Constructor.bas:1433-1435` |
| Se excluyen la aplicación de ID 1 y las que están en pruebas | `CargarArbol` | `Funciones Genreales.bas:2560-2567` |
| Ancho mínimo del árbol 60 twips; el separador no pasa de la mitad del ancho | `Sep_MouseMove`, `Sep_MouseUp` | `:141-143, 162-163` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Aplicaciones | `TbAplicaciones` | DAO, tabla completa | No | `Constructor.bas:110-111` |
| Vídeos de ayuda por aplicación y rol | `TbVideos`, `TbVideosCategorias`, `TbCategorias` | SQL dinámico concatenado | No | `Constructor.bas:1390-1446` |
| Vídeo seleccionado | `TbVideos` por `IDVideo` | SQL dinámico concatenado | No | `Constructor.bas:1259-1261` |

## Coexistencia

`IDAplicacion` e `IDVideo` se concatenan sin parámetros (`Constructor.bas:1260, 1396`); los filtros comparan textos `'Sí'` (`Constructor.bas:1406, 1426`), que en PostgreSQL convendría mapear a booleanos. Se trata en `coexistencia.md` (tanda posterior).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba; formulario sin punto de entrada activo.

**Posibles defectos (sin corregir, a verificar):**

- `CInt(dato(1))` (`:32`) desborda con un `IDVideo` mayor que 32 767 (error 6, correo al administrador).
- La llamada a `EstablecerDatos` no recoge `p_Error` (`:59`): un fallo del detalle no se muestra.
- Un nodo de aplicación se añade aunque no tenga vídeos (`Funciones Genreales.bas:2585-2590`).
- `CargarArbol` no excluye «Registro» como `RellenarListaImagenes` (`Funciones Genreales.bas:2439` frente a `:2560-2567`); los fallos de `Nodes.Add` se ignoran con `On Error Resume Next` (`Funciones Genreales.bas:2584-2588`).
- `Const SepW As Integer = 60.1` (`:20`) se redondea a entero.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: decidir si la consulta de vídeos se publica o se retira, dado que está desactivada; filtro por rol en el servidor con consultas parametrizadas.

## Criterios de aceptación

Propuestos, no aprobados:

1. El árbol muestra solo aplicaciones que no están en pruebas y, bajo cada una, sus vídeos de ayuda visibles para el rol del usuario.
2. Seleccionar un vídeo muestra su detalle; seleccionar una aplicación lo vacía.
3. Un usuario de Calidad no ve vídeos sin `ParaCalidad`; un técnico, vídeos sin `ParaTecnicos`.

## Pendientes

- Si la clave de imagen `NombreIconoArbol` coincide con la `NombreBase` usada en `AddImagen` (`Funciones Genreales.bas:2455, 2576`).
- Si `Form_FormVideoDetalle` (`:59`) alcanza la instancia anidada en `FormVideosArbolVideo`.
- Motivo de la desactivación del menú «Videos-Guías» y si se prevé reactivarlo.
