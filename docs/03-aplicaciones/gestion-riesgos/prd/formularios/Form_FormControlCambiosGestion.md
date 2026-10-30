# Form_FormControlCambiosGestion

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormControlCambiosGestion.cls` (455 líneas) y `src/forms/FormControlCambiosGestion.form.txt` · **PRD:** [README](../README.md)

Buscador del control de versiones del gestor de riesgos. Combina tres criterios (versión principal, versión y texto libre dentro de los cambios de esa versión), muestra el resultado en una lista y permite generar la ficha web de control de cambios de la versión elegida. No escribe nada en la base de datos: solo filtra, selecciona y delega la generación. Las citas son `Form_FormControlCambiosGestion.cls:línea` salvo indicación.

## Entrada y salida

No acepta parámetros ni `OpenArgs`. Los dos únicos puntos que lo abren son los menús de opciones, técnico y general, y ambos lo hacen sin argumentos:

| Desde | Evidencia |
|---|---|
| `Form0BDOpcionesTecnico` | `Form_Form0BDOpcionesTecnico.cls:170` |
| `Form0BDOpciones` | `Form_Form0BDOpciones.cls:385` |

El contexto —la colección de versiones y sus cambios— llega por la global `m_ObjEntorno` (`ColVersiones`, `:345`), y el objeto `CCVersion` de la fila elegida se carga con `Constructor.getCCVersion` (`:165`, `:419`, `:425`). Las clases que modelan el dominio son `CCVersion`, `CCCambio` y `CCDocumentoCambio` (`src/classes/`).

Salidas: no abre formularios. Genera HTML y lo entrega a `MostrarControlCambiosVersion` (`:174`), que es la función genérica de `Funciones Generales.bas:1537`, y de ahí al visor de informes.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Tamaño, título, combo y primer filtrado | `:285-321` |
| `ListaFiltrados` | cuadro de lista | `ListaFiltrados_Click` | Resuelve el objeto de la fila elegida | `:403-455`, `.form.txt:1016-1023` |
| `ListaFiltrados` | — | `ListaFiltrados_DblClick` | Doble clic equivale a «Informe» | `:249-256`, `.form.txt:1022` |
| `VersionPrincipal` | desplegable | `VersionPrincipal_BeforeUpdate` | Refiltra al confirmar el valor | `:199-223`, `.form.txt:1089-1093` |
| `VersionPrincipal` | — | `AfterUpdate` | **Declarado en el diseño sin manejador en el código** | `.form.txt:1094` |
| `Version` | desplegable | `Version_BeforeUpdate` | Refiltra al confirmar el valor | `:224-248`, `.form.txt:1134-1137` |
| `Cambio` | cuadro de texto | `Cambio_Exit` | Refiltra al salir del campo | `:258-284`, `.form.txt:1052-1053` |
| `ComandoInforme` | botón | `ComandoInforme_Click` | Genera la ficha web de la versión | `:156-193`, `.form.txt:1257-1261` |
| `ComandoLimpiarVersionPrincipal` | botón | `…_Click` | Vacía el criterio y refiltra | `:69-96`, `.form.txt:1292-1296` |
| `ComandoLimpiaVersion` | botón | `…_Click` | Vacía el criterio y refiltra | `:97-124`, `.form.txt:1471-1475` |
| `ComandoLimpiarCambio` | botón | `…_Click` | Vacía el criterio y refiltra | `:125-155`, `.form.txt:1650-1654` |
| `ComandoAyuda` | botón | `ComandoAyuda_Click` | Abre la ayuda | `:41-68`, `.form.txt:934-937` |
| `ComandoSalir` | botón | `ComandoSalir_Click` | Cierra sin guardar | `:194-197`, `.form.txt:1840-1843` |

Los tres botones de limpiar comparten rótulo `Comando36` (`.form.txt:1293`, `:1472`, `:1651`), es decir, marcador de Access en pantalla; solo el globo los distingue.

## Flujos

### Apertura (`Form_Open`, `:285-321`)

1. `AjustarTamaño Me` (`:292`) y `Me.Caption = m_ObjEntorno.TituloUsuarioConectado` (`:294`), que **pisa** el rótulo de diseño, «BUSCADOR DE GESTIÓN DE RIESGOS» (`.form.txt:40`).
2. `RellenarComboVersiones` (`:295`): vacía el desplegable y lo recorre con `m_ObjEntorno.VersionesAplicacion` (`:22-32`). Si no hay colección, sale sin poblarlo.
3. `Filtrar` (`:299`), que carga la lista con el estado actual de los tres criterios (vienen vacíos).
4. Cualquier fallo termina en `DoCmd.Close` incondicional (`:319`).

### Normal: filtrado (`Filtrar`, `:324-402`)

Es el corazón del formulario y **vuelve a ejecutarse en cada interacción** con cualquiera de los tres criterios.

1. Descarta el objeto seleccionado (`:341`) y pone el `RowSource` a la fila de cabecera `IDVersion;Version;Descripcion` (`:344`); los datos se añaden después con `AddItem` (`:384`).
2. Toma `m_ObjEntorno.ColVersiones` y, si la propiedad devuelve error, lo propaga (`:345-349`). Si la colección es `Nothing`, sale (`:350-352`).
3. Recorre todas las versiones y **descarta** por salto a la etiqueta `siguiente` (`:386`) la que no cumple:
   - `VersionPrincipal` distinto al criterio (`:357-361`);
   - `Version` distinta al criterio (`:362-366`);
   - si hay texto de cambio, exige que **algún** cambio de esa versión contenga ese texto (`:367-382`).
4. La versión que pasa los tres filtros se añade a la lista (`:383-384`).
5. **Si el resultado tiene más de una fila**, selecciona la primera y guarda el objeto (`:391-395`).

### Normal: selección (`ListaFiltrados_Click`, `:403-455`)

Deshabilita el botón de informe (`:410`), resuelve el `CCVersion` de la columna 0 si no lo tiene o si ha cambiado (`:415-431`), lo suelta si ya no existe (`:433-436`) y vuelve a habilitar el botón (`:437`). El doble clic (`:249-256`) delega en el mismo manejador, protegido por `On Error Resume Next`.

### Normal: generar la ficha (`ComandoInforme_Click`, `:156-193`)

1. Si no hay objeto, lo carga desde la columna 0; si no aparece, avisa de que hay que elegir un elemento (`:163-173`).
2. `MostrarControlCambiosVersion m_VersionSeleccionada` (`:174`), que devuelve el error por parámetro si falla.

### Alternativo: limpiar un criterio

Los tres manejadores hacen lo mismo (`:69-96`, `:97-124`, `:125-155`): si el campo no está vacío, lo ponen a `Null` y refiltran. **No hay un botón que limpie los tres a la vez**; el usuario los limpia de uno en uno.

### Error

Ocho manejadores repiten el patrón (`:60-67`, `:85-95`, `:113-123`, `:141-154`, `:180-192`, `:210-222`, `:235-247`, `:272-283`, `:309-321`): error con número distinto de 1000, correo al administrador y crítico; 1000 exacto, advertencia con el texto ya preparado. `Form_Open` es el único que además **cierra** el formulario (`:319`).

## Reglas, validaciones y cálculos

- **La búsqueda de texto es por subcadena y sin distinguir mayúsculas**: `InStr(1, m_Cambio.NombreCambio, Me.Cambio) <> 0` (`:374`). El criterio `Me.Cambio` no se normaliza ni se escapa.
- **El filtro es conjunction de tres criterios, todos opcionales**: vacíos, no filtran.
- **Un resultado único deja el formulario sin acción rápida**: con `ListCount = 1` no se selecciona nada ni se habilita el botón de informe (`:391`). El usuario tiene que hacer clic en la fila, aunque sea la única.
- **El botón de informe solo se habilita con selección válida**, nunca por defecto: el estado depende del último `ListaFiltrados_Click` o de `Filtrar` (`:394`, `:437`).
- Los tres criterios llevan `Tag = "BUSQUEDA"` (`.form.txt:1056`, `:1095`, `:1139`). Ningún código de este formulario recorre controles por etiqueta, así que el mecanismo que la usa está fuera y conviene localizarlo antes de replicarlo en la web.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Versiones de la aplicación | `m_ObjEntorno.VersionesAplicacion` | `:20-30` |
| Versiones y sus cambios | `m_ObjEntorno.ColVersiones`, `CCVersion.Cambios` | `:345`, `:368` |
| Versión concreta | `Constructor.getCCVersion` | `:165`, `:419`, `:425` |
| Nombre del cambio | `CCCambio.NombreCambio` | `:374` |
| Ficha web | `MostrarControlCambiosVersion`, `Funciones Generales.bas:1537` | `:174` |

**Sin escrituras**: ni en base de datos ni en disco. El formulario solo prepara la selección y pide la generación.

## Coexistencia

Pendiente de la capa de coexistencia, sin enlace desde esta ficha. Todo el peso de esta pantalla está en la salida: `MostrarControlCambiosVersion` genera HTML y termina en el visor local, igual que el resto de la aplicación. Mientras convivan, la ficha generada tiene que existir en un fichero que el cliente alcance, y el buscador funciona sin cambios pero su resultado no se ve.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-297.** `VersionPrincipal` declara `AfterUpdate = "[Event Procedure]"` en el diseño (`.form.txt:1094`) pero **no existe `VersionPrincipal_AfterUpdate` en el código**. Access puede lanzar un error en tiempo de ejecución al salir del desplegable, o dejar el evento sin efecto según cómo se haya compilado el binario. Conviene comprobarlo en ejecución.
- **H-GR-298.** `If lst.ListCount > 1 Then` (`:391`) excluye el caso de un único resultado: no selecciona la fila ni habilita el botón de informe. Con una búsqueda muy específica el formulario parece no responder.
- **H-GR-299.** `Set m_VersionSeleccionada = m_Version` (`:393`) guarda **la última versión iterada**, no la de la primera fila que se acaba de seleccionar. Tras filtrar, el botón de informe puede quedar habilitado apuntando a una versión distinta de la que el usuario ve marcada.
- **H-GR-300.** El diseño guarda un `RowSource` de muestra con la descripción de una versión real y el nombre de una persona (`.form.txt:1018`). Dato de negocio y personal incrustado en el artefacto; no se reproduce aquí. Conviene señalarlo al propietario.
- **H-GR-301.** El rótulo del botón `ComandoAyuda` es «Salir» (`.form.txt:935`) y solo el globo dice «Ayuda» (`.form.txt:937`); junto al botón `ComandoSalir`, también rotulado «Salir» (`.form.txt:1841`), hay dos botones indistinguibles a la vista.
- **H-GR-302.** `Form_Open` cierra el formulario incondicionalmente en su manejador de error (`:319`), sin asignar `Cancel`: un fallo de red al leer las versiones deja al usuario sin buscador y sin salida más allá del mensaje.
- **H-GR-303.** `Application.Echo False` se enciende en `ListaFiltrados_Click` (`:409`) y en `ListaFiltrados_DblClick` (`:251`); el doble clic llama a `ComandoInforme_Click` con el eco apagado y lo vuelve a encender al salir (`:255`). Si el manejador de informe falla con `Resume`, el eco puede quedar apagado y Access «se cuelga» aparentemente sin hacer nada.
- **H-GR-304.** La lógica de filtro usa `GoTo escribe` desde dentro de un `For` anidado (`:376`) hacia una etiqueta situada dentro del `With` (`:383`), y `GoTo siguiente` como salida del bucle (`:359`, `:364`, `:370`, `:381`). Funciona, pero cualquier refactor que reordene esas líneas cambia el comportamiento sin aviso.
- **H-GR-305.** El criterio de texto se compara con `InStr` sobre `NombreCambio` sin normalizar espacios ni acentos (`:374`): lo que el usuario escribe no tiene por qué coincidir con lo que escribió quien registró el cambio.
- **H-GR-306.** `Me.Caption` se sobrescribe con el usuario conectado (`:294`) y pierde el nombre de la pantalla que tenía en diseño (`.form.txt:40`); en una sesión con varios formularios abiertos no se distingue uno de otro.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: devolver el resultado de la búsqueda como listado paginado con filtros combinables, en lugar de un `RowSource` de texto; habilitar el botón de informe con un único resultado; resolver la búsqueda en el servidor con una consulta parametrizada en lugar de filtrar en memoria; añadir un botón que limpie los tres criterios a la vez; mantener el detalle de versión como ruta propia en lugar de generar HTML en disco.

## Criterios de aceptación

Propuestos, no aprobados:

1. Combinar los tres criterios devuelve solo las versiones que los cumplen todos.
2. Un criterio vacío no restringe el resultado.
3. Con un solo resultado, ese resultado queda seleccionado y se puede generar su ficha sin un clic adicional.
4. La ficha generada corresponde a la fila marcada, no a otra.
5. Vaciar un criterio refiltra y deja la lista coherente.
6. Un fallo al leer las versiones avisa y deja el buscador utilizable.

## Pendientes

- Si el `AfterUpdate` declarado en `:1094` del diseño llega a ejecutarse y qué efecto tendría.
- Qué componente usa la etiqueta `BUSQUEDA` de los tres criterios y si hay más pantallas con el mismo patrón.
- Qué hace `MostrarControlCambiosVersion` exactamente (`Funciones Generales.bas:1537`) y en qué formato produce la ficha.
- De qué depende que `CCVersion.Cambios` esté ya cargado en memoria, y quién lo carga.