# Form_FormExpedienteSuministradores

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormExpedienteSuministradores.cls` (961 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña Empresas de la ficha de expediente: árbol de suministradores (`tvCadena`, TreeView ActiveX) con el órgano de contratación como raíz, adjudicatarios en el primer nivel y subcontratistas debajo, y un catálogo de empresas a la derecha (`lvFuente`, ListView ActiveX) desde el que se arrastran. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`). Trabaja en modo híbrido: en un alta modifica la colección en memoria del DTO; en una edición escribe directamente en la base de datos, cada operación en su propia transacción (`:8-12, 248-273, 812-883`). Las citas son `Form_FormExpedienteSuministradores.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormExpediente` | botón `NavSuministradores` (rótulo «Empresas») con `NavigationTargetName = "FormExpedienteSuministradores"`; antes llama a `RellenarDTO` | `Form_FormExpediente.form.txt:903`, `Form_FormExpediente.cls:576-597` |

No recibe parámetros: lee el DTO global `m_ObjExpedienteDTOActivo` (`Variables Globales.bas:265`); sin DTO, error «No hay contexto de expediente activo.» y cancela la apertura (`:46-49, 103-105`). Abre `FormSuministradoresGestion` en diálogo (`:561`), `FormSuministrador` en diálogo para ver una ficha (`:592, 666`) y `FormSuministrador` en modo normal para dar de alta una empresa, escuchando su evento `Alta` con `WithEvents` (`:34, 709-729`; evento en `Form_FormSuministrador.cls:5, 84`). El menú contextual llama a `EjecutarAccionMenu`, que resuelve el formulario con `Screen.ActiveControl.Parent` y ejecuta el método con `CallByName` (`modules/MenuContextual.bas:8-29`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `tvCadena` | TreeView (`MSComctlLib.TreeCtrl.2`) | `NodeClick` | Habilita `cmdEliminar`/`cmdVerDetalle` según el nodo | `:597-609`, `.form.txt:658, 903` |
| `tvCadena` | TreeView | `MouseUp` (botón derecho) | Selecciona el nodo y abre el menú «ARBOL» | `:143-153` |
| `tvCadena` | TreeView | `OLEStartDrag`, `OLEDragOver`, `OLEDragDrop` | Mover un nodo dentro del árbol o soltar una empresa del catálogo | `:410-535` |
| `lvFuente` | ListView (`MSComctlLib.ListViewCtrl.2`) | `MouseUp` (botón derecho) | Menú «LISTA» | `:132-141`, `.form.txt:1368, 1549` |
| `lvFuente` | ListView | `OLEStartDrag` | Arrastra la clave `SRC_<IDSuministrador>` como copia | `:402-408` |
| `txtFiltro` | texto | `OnChange` → `txtFiltro_Change` | Recarga el catálogo con el texto tecleado | `:155-158`, `.form.txt:1603-1606` |
| `chkNuestras` («Nuestras Empresas») | casilla | `OnClick` | Recarga el catálogo solo con empresas propias | `:160-165`, `.form.txt:1567-1583` |
| `cmdEliminar` («Quitar de Expediente») | botón | `OnClick` | Pide confirmación y quita el nodo seleccionado | `:569-582`, `.form.txt:923-925` |
| `cmdVerDetalle` («Ver Detalle») | botón | `OnClick` | Carga `m_ObjSuministradorActivo` y abre `FormSuministrador` en diálogo | `:584-594`, `.form.txt:969-971` |
| `cmdNuevoSuministrador` («Nuevo Suministrador») | botón | `OnClick` | Cierra `FormSuministrador` si estaba abierto y lo abre en alta | `:709-729`, `.form.txt:1811-1813` |
| `cmdLimpiarFiltro` | botón | sin `OnClick` ni handler | Sin comportamiento («Quita el contenido del campo» solo como `ControlTipText`) | `.form.txt:1628-1630` |
| `ListaImagenes` | ImageList oculta | — | Iconos de órgano, empresa propia, ajena y lista, cargados desde URL del entorno | `:57-74`, `.form.txt:1019, 1328` |

Menú contextual (`MostrarMenuContextual`, `:611-643`): en el árbol, «Ver Detalle» (no raíz), «Añadir Aquí...» y «Eliminar» (no raíz) solo con `AllowEdits`; en la lista, «Ver Ficha» y, con `AllowEdits`, «Añadir al Expediente...». Cada opción llama a un método `Public_*` (`:649-679`).

## Flujos

### Apertura (`Form_Open`, `:40-106`)

1. `m_EsAlta = (IDExpediente = "")` decide el modo memoria o base de datos (`:52-53`).
2. Recarga los iconos y vincula la `ImageList` a los dos controles (`:57-74`); prepara el árbol y la lista (`:77-78, 116-130, 202-213`).
3. Pinta el árbol (`RefrescarAArbol`) y carga el catálogo sin filtro (`:81-82`).
4. Permisos: en alta siempre editable; en edición, solo si `EsAdministrador = Sí` y `Form_FormExpediente.m_AbiertoParaEditar` (`:85-90`). Aplica `AllowEdits` y habilita `cmdEliminar`, `txtFiltro`, `chkNuestras` y `cmdNuevoSuministrador` (`:92-98`).

### Pintado del árbol (`RefrescarAArbol`, `:215-372`)

1. Datos: en alta, de `ColArbolSuministradores` del DTO (`:248-259`); en edición, de `ExpedienteSuministradorRepositorio.getDatosArbol` (`:261-272`; SQL en `modules/ExpedienteSuministradorRepositorio.bas:20-24`).
2. Raíz única «Órgano de Contratación (<nombre>)», o «DESCONOCIDO» si el expediente no tiene órgano (`:278-292`).
3. Los nodos de primer nivel (`IdPadre` vacío) cuelgan de la raíz; si hay dos o más, cada uno añade « (UTE con: <otros>)» (`:297-302, 332-334, 374-395`).
4. Nombre mostrado: `Nemotecnico` si existe, si no `Nombre` (`:685-693`); empresa propia en negrita con icono de grupo (`:341-347`).
5. Inserta nodos en pasadas sucesivas hasta colocar todos o dejar de avanzar (`:307-362`). Etiqueta `Tag = "RELID=<id relación>;IDS=<id suministrador>"` (`:336`).
6. `ActualizarEtiquetaSituacion` fija `Expediente.ContratistaPrincipal` del DTO a «Sí» o «No» (`:367, 761-803`).

### Añadir empresa

- **Arrastrar del catálogo:** soltar en vacío o en la raíz crea un adjudicatario (`ROOT_CONTR`); soltar sobre un nodo crea un subcontratista suyo (`CHILD`) (`:480-505`).
- **Menú «Añadir al Expediente...»:** siempre como adjudicatario (`:673-679`).
- **Menú «Añadir Aquí...»:** abre `FormSuministradoresGestion` en diálogo y lee el elegido de `TempVars("SelectedIDSuministrador")` (`:541-567`).
- `EjecutarAlta` añade en memoria (`AgregarSuministradorMemoria`, `:887-913`) o llama a `ExpedienteSuministradorServicio.AgregarSuministrador`, y repinta (`:812-833`).

### Mover y quitar

- **Mover:** arrastrar un nodo sobre la raíz lo convierte en adjudicatario; sobre otro nodo, en subcontratista de este. No se mueve la raíz, ni sobre sí mismo, ni sobre un hijo directo (`:415, 508-526`). `EjecutarMovimiento` cambia memoria o base y repinta (`:835-856`).
- **Quitar:** rechaza la raíz («No se puede eliminar la raíz.»), pide «¿Eliminar seleccionado?» (`:569-581`). Con hijos falla: en base con «Tiene dependencias.» (`:866-868`); en memoria sin mensaje propio (`:943-951`). Si funciona, borra el nodo visual sin repintar (`:874-878`).

### Alta de una empresa nueva

`cmdNuevoSuministrador` abre `FormSuministrador` (`:713-724`). Cuando este lanza `Alta`, se desmarca `chkNuestras`, se pone el nombre en `txtFiltro` y se recarga el catálogo (`:730-756`). **No añade la empresa al árbol.**

### Errores

`MsgBox` crítico con «Error Alta: », «Error Mover: » o «Error Eliminar: » seguido de `m_Error` (`:830, 853, 880`); «Error DragDrop Nuevo: » y repintado (`:531-534`); «Error RefrescarAArbol_Nuevo: » (`:370-371`). Sin correo al administrador.

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Adjudicatario: `ContratistaPrincipal = "Sí"`, `SubContratista = "No"`, sin padre; subcontratista: lo contrario, con padre | memoria y servicio | `:902-903, 923-924`; `classes/ExpedienteSuministradorServicio.cls:57-61, 115-119` |
| Misma empresa no repetida bajo el mismo padre (solo en edición) | servicio | `ExpedienteSuministradorServicio.cls:37-40`; `ExpedienteSuministradorRepositorio.bas:318-365` |
| No se quita un nodo con hijos | formulario y servicio | `:866-868, 943-951`; `ExpedienteSuministradorServicio.cls:251-254` |
| Cabecera `TbExpedientes.ContratistaPrincipal = "Sí"` si hay una empresa propia (`ConsorcioPropio = 'Sí'`) en primer nivel como adjudicataria | servicio, tras cada operación | `ExpedienteSuministradorServicio.cls:285-309`; `ExpedienteSuministradorRepositorio.bas:256-276` |
| Mismo cálculo en memoria para el DTO; en edición vía `HayNuestrasEmpresasComoPrincipal`, que no filtra por `IdPadre` | formulario | `:766-802`; `ExpedienteSuministradorRepositorio.bas:228-233` |
| Rama `ROOT_SUB` («Participación» solo con empresas propias): existe en el servicio pero el formulario nunca la usa | servicio | `ExpedienteSuministradorServicio.cls:43-48, 110-113`; el formulario solo envía `ROOT_CONTR`/`CHILD` (`:494, 501, 516, 520`) |
| Catálogo: filtro `LIKE` por `Nombre`, `CIF` o `Nemotecnico`; con «Nuestras Empresas», solo `ConsorcioPropio = 'Sí'`; orden por `Nombre` | constructor | `modules/constructor.bas:5792-5804` |
| Id de relación: en memoria `TMP_<hhmmss>_<aleatorio>`; en base `DameID` sobre `TbExpedientesSuministradores` | formulario y clase | `:893`; `classes/ExpedienteSuministrador.cls:59-70` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Árbol | `TbExpedientesSuministradores` | SQL dinámico, snapshot | no | `ExpedienteSuministradorRepositorio.bas:20-27` |
| Catálogo | `TbSuministradores` | SQL dinámico concatenado | no | `constructor.bas:5792-5806` |
| Nombre y «propia» por nodo | `TbSuministradores` (una consulta por nodo y llamada) | SQL dinámico | no | `:329, 341`; `constructor.bas:5573-5590`; `ExpedienteSuministradorServicio.cls:327-336` |
| Añadir | `INSERT` en `TbExpedientesSuministradores` | SQL dinámico | `ws.BeginTrans`/`CommitTrans` | `ExpedienteSuministradorServicio.cls:65-79`; `ExpedienteSuministradorRepositorio.bas:60-72` |
| Mover | `UPDATE` de `IdPadre` y roles | SQL dinámico | sí | `ExpedienteSuministradorServicio.cls:121-137`; `ExpedienteSuministradorRepositorio.bas:200-205` |
| Quitar | `DELETE` por `IDExpedienteSuministrador` | SQL dinámico | sí | `ExpedienteSuministradorServicio.cls:256-270`; `ExpedienteSuministradorRepositorio.bas:95-96` |
| Cabecera | `UPDATE TbExpedientes SET ContratistaPrincipal` | SQL dinámico | dentro de la anterior | `ExpedienteSuministradorRepositorio.bas:268-270` |
| Caché | `RegistrarExpEntidades` (siempre ámbito `Todo`) | DAO | dentro; `RegistrarUltimoCambio` después del commit | `ExpedienteSuministradorServicio.cls:74-81`; `ExpedienteOperaciones.cls:607` |

En alta no escribe nada: el árbol queda en `ColArbolSuministradores` y lo debería persistir `ExpedienteOperaciones.Registrar` con `PersistirArbolDesdeMemoria` (`ExpedienteOperaciones.cls:488-495`).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: todo el SQL es concatenado (identificadores sin comillas, `LIKE '*…*'` con comodín Jet, `Nz` en `ORDER BY`, `ExpedienteSuministradorRepositorio.bas:24`; `constructor.bas:5796`); `DameID` genera el identificador en VBA; los valores «Sí»/«No» son texto; la transacción DAO cubre relación, cabecera y caché.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba. En edición, cada cambio del árbol se guarda en el acto, sin pasar por el botón de guardado del contenedor.

**Posibles defectos (sin corregir, a verificar):**

- `PersistirArbolDesdeMemoria` está vacío (`ExpedienteSuministradorServicio.cls:312-316`): el árbol montado en un alta no se guardaría (llamada en `ExpedienteOperaciones.cls:492`).
- El servicio llama al repositorio sin `p_Error` (`ExpedienteSuministradorServicio.cls:70, 126, 260, 308`) y el repositorio captura sus errores (`ExpedienteSuministradorRepositorio.bas:74-77`): un `INSERT`/`UPDATE`/`DELETE` fallido no revertiría y se haría `CommitTrans`.
- `EliminarRelacion` no comprueba `p_Error` tras `RegistrarExpEntidades` antes del commit (`ExpedienteSuministradorServicio.cls:267-270`), a diferencia de añadir y mover (`:77, 135`).
- El aviso de empresa repetida se pierde: `p_Error` se sobrescribe con «Error en AgregarSuministrador: » y la descripción genérica del `Err.Raise 1000` (`ExpedienteSuministradorServicio.cls:38-39, 87`). En memoria no hay control de repetidos (`:887-909`).
- «Añadir Aquí...» lee `TempVars("SelectedIDSuministrador")`, que ningún fichero de `src` asigna (`:560-564`; `Form_FormSuministradoresGestion.cls` no usa `TempVars`): la opción no añadiría nada.
- Mover solo impide soltar sobre un hijo directo (`:521-523`; servicio `:106`): soltar sobre un nieto crearía un ciclo, y `RefrescarAArbol` omite sin aviso los nodos que no puede colgar (`:353-361`).
- Tras quitar no se repinta ni se llama a `ActualizarEtiquetaSituacion` (`:874-878`): el texto «UTE con» y `ContratistaPrincipal` del DTO quedan desfasados; el guardado de cabecera lo escribe tal cual (`ExpedienteOperaciones.cls:387`) y podría deshacer el valor recalculado en base.
- Quitar con hijos en memoria muestra «Error Eliminar: » sin motivo (`:950-951`); `MoverRelacion` con `Err.Raise 1000` sin texto también (`ExpedienteSuministradorServicio.cls:105-106`).
- En edición, las operaciones no actualizan `ColArbolSuministradores` del DTO (`:817-823`), que el contenedor usa para «Ningún cambio que guardar» (`Form_FormExpediente.cls:71-76`).
- El filtro concatena el texto sin escapar comillas (`constructor.bas:5796`); el error se traga y la lista queda vacía (`:174-178`). El `ControlTipText` habla de «título, nemotécnico, código expediente» (`.form.txt:1606`).
- `cmdVerDetalle` y «Ver Ficha» abren la ficha maestra del suministrador también en solo lectura (`:584-594, 658-668`).
- Restos: `cmdLimpiarFiltro` sin handler; `Public_AnadirComoAdjudicatario` sin uso (`:669-672`); `MoverConsorcioCompleto` y `ContarPropiasEnRaiz` sin llamadas desde `src` (`ExpedienteSuministradorServicio.cls:155, 387`); `Rnd` sin `Randomize` en `src` (`:893`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: un único modelo de guardado (todo en memoria hasta guardar, o todo inmediato) para alta y edición; validar ciclos y repetidos en el servidor; sustituir el arrastre por acciones explícitas «añadir como adjudicatario/subcontratista».

## Criterios de aceptación

Propuestos, no aprobados:

1. Soltar una empresa en la raíz la añade como adjudicataria; sobre un nodo, como subcontratista de ese nodo.
2. Con dos o más adjudicatarios, cada uno muestra «UTE con» y los demás.
3. Añadir la misma empresa dos veces bajo el mismo padre se rechaza con un mensaje que lo explica.
4. Quitar un nodo con subcontratistas se rechaza; la raíz no se puede quitar.
5. `ContratistaPrincipal` del expediente vale «Sí» si y solo si hay una empresa propia como adjudicataria, tras añadir, mover o quitar.
6. Sin permiso de edición no se puede añadir, mover ni quitar (`:85-98, 403, 411, 425, 570`).
7. El árbol de un alta queda guardado al registrar el expediente.

## Procedimientos auxiliares

Citas `Form_FormExpedienteSuministradores.cls:línea`.

| Procedimiento | Qué hace | Evidencia |
|---|---|---|
| `ExtraerTag(tag, clave)` | Lee el valor de una clave en el `Tag` de un nodo con formato `CLAVE=valor;CLAVE2=valor2`; no distingue mayúsculas, recorta espacios y devuelve vacío si la clave falta. El formulario solo pide `RELID`, el identificador de la relación del nodo | `:695-707`; usos `:500, 512, 519, 555, 579` |
| `MoverRelacionMemoria` | Solo en alta (modo memoria): reasigna el padre del elemento en `ColArbolSuministradores`; `ROOT_CONTR` lo deja como adjudicatario (`ContratistaPrincipal = "Sí"`, `SubContratista = "No"`) y `CHILD` como subcontratista (lo contrario). Si la relación no existe devuelve falso con «Elemento no encontrado en memoria.» | `:915-936` |
| `ObtenerTextoSocios` | Con dos o más adjudicatarios, devuelve « (UTE con: <los demás>)» separando por coma y excluyendo al propio nodo | `:374-395` |
| `ConfigurarListView`, `PrepararTreeView` | Configuración visual de los dos controles ActiveX: la lista, en modo informe, con una columna «Empresa / Suministrador» de ancho el del control menos 300 (2000 si queda por debajo de 1000) y arrastre OLE automático; el árbol, con sangría 300 y arrastre y soltado OLE. No tienen reglas de negocio | `:116-130`, `:202-213` |
| `Form_Close` | Libera `m_Servicio` y nada más | `:108-115` |

## Pendientes

- Si `Enabled = NotDefault` de `tvCadena` y `lvFuente` los deja deshabilitados en diseño (`.form.txt:649, 1358`) y si el código los habilita.
- Si `FormSuministradoresGestion` devuelve el elegido por otro canal (macro o `Eval`) para «Añadir Aquí...».
- Persistencia real del árbol en alta: se confirma con un alta en el binario o con la ficha de `Form_FormExpedienteAltaParaHPS`, que monta un árbol análogo (`Form_FormExpedienteAltaParaHPS.cls:274-332, 399`).
- Contenido de `EstablecerDTODeExpediente` al copiar desde un padre: carga el árbol del padre con sus id reales (`FUNCIONES UTILES.bas:2105, 2142-2163`).
- Orden en que Access dispara `MouseUp` y `NodeClick` frente al menú contextual; resolución de `Screen.ActiveControl.Parent` en un subformulario (`MenuContextual.bas:18`).
