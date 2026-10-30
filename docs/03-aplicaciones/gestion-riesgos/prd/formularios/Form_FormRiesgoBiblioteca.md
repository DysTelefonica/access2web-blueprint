# Form_FormRiesgoBiblioteca

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormRiesgoBiblioteca.cls` (301 líneas) y `src/forms/FormRiesgoBiblioteca.form.txt` (1588 líneas) · **PRD:** [README](../README.md)

Alta y edición de un riesgo tipo de la biblioteca: tres campos (familia, si está activo y descripción). No pertenece al contenedor de riesgo: es una pantalla independiente, abierta por la gestión de la biblioteca y dirigida por eventos. Las citas son `Form_FormRiesgoBiblioteca.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormRiesgosBibliotecaGestion` | `DoCmd.OpenForm "FormRiesgoBiblioteca"` desde el botón de alta, tras cerrar la instancia previa | `Form_FormRiesgosBibliotecaGestion.cls:192-195` |
| `Form_FormRiesgosBibliotecaGestion` | `DoCmd.OpenForm "FormRiesgoBiblioteca"` desde el botón de editar | `Form_FormRiesgosBibliotecaGestion.cls:236-240` |

El contenedor conserva el puntero con `Set m_FormRiesgoBiblioteca = Forms("FormRiesgoBiblioteca")` (`:197`, `:242`) y se declara `WithEvents` para sus dos eventos (`:19`). La comunicación es de vuelta: esta pantalla emite `AltaRiesgoBiblioteca` (`:76`) y `EdicionRiesgoBiblioteca` (`:78`), que el contenedor atiende en `:723` y `:758`.

No abre ningún formulario. Se cierra a sí mismo en cuatro puntos (`:86`, `:97`, `:117`, `:150`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Ajusta el tamaño y refresca | `:121-151` |
| `ComandoGrabar` | botón | `OnClick` → `ComandoGrabar_Click` | Valida, vuelca, guarda y emite el evento correspondiente | `:49-111`, `.form.txt:1264-1266` |
| `ComandoAyuda` | botón | `OnClick` → `ComandoAyuda_Click` | Abre la ayuda de la pantalla | `:21-47`, `.form.txt:904-906` |
| `ComandoSalir` | botón | `OnClick` → `ComandoSalir_Click` | Cierra sin guardar | `:115-118`, `.form.txt:1189-1191` |
| `Familia` | lista desplegable | — | Familia del riesgo tipo; se lee y se escribe | `:988`, `:90`, `:200`, `:226`, `:260` |
| `Activo` | lista desplegable | — | Si el riesgo tipo está activo | `:1056`, `:92`, `:203`, `:230`, `:261` |
| `Descripcion` | cuadro de texto | — | Descripción del riesgo tipo | `:1021`, `:91`, `:206`, `:234`, `:262` |
| `lblTitulo1` | etiqueta | — | Título de alta, detalle o edición | `:883`, `:190` |

El diseño conserva además tres etiquetas sin código asociado (`Etiqueta232`, `Etiqueta20`, `Etiqueta21`; `.form.txt:1089`, `:1119`, `:1149`). Los bloques `Begin` sin nombre de `.form.txt:722-843` son valores por defecto de propiedades, no controles instanciados.

## Flujos

### Apertura (`Form_Load` → `EstablecerDatos`, `:121-215`)

1. `Form_Load` pone el reloj, ajusta el tamaño con `AjustarTamaño` (`:130`) y llama a `EstablecerDatos` (`:132`), propagando el error (`:134`).
2. `EstablecerDatos` fija el permiso con la global `EsTecnico` (`Variables Globales.bas:109`): si es `Sí`, deshabilita la edición y el botón de grabar (`:161-163`); si no, los habilita (`:166-167`).
3. Llena la lista de familias con `EstablecerLista` (`:169`) y elige el título: «Alta de riesgo de biblioteca» sin objeto activo (`:175`), «Detalle» si no se puede editar (`:177`) o «Edición» (`:179`).
4. Si hay riesgo activo, lo recarga con `Constructor.getRiesgoBiblioteca` (`Constructor.bas:2885`) (`:181`).
5. Con `RiesgoParaRetipificar = Sí` deshabilita edición y botón (`:185-187`), ajusta el rótulo (`:190`) y el título de la ventana (`:191`).
6. Sin copia inicial sale (`:194-195`); con ella vuelca los tres campos (`:200`, `:203`, `:206`).

### Normal: guardar (`ComandoGrabar_Click`, `:49-111`)

1. Reloj de arena y `m_Error = ""` (`:56-58`).
2. `HaHabidoCambios` (`:59`); si devuelve error, aborta (`:61`); si no hay cambios, avisa «No ha habido cambios que guardar» y aborta (`:63-65`).
3. `RellenarObjetoConDatosFormulario` (`:67`), con comprobación de error (`:69`).
4. `Registrar m_ObjRiesgoBibliotecaAlIncio, EnumSiNo.Sí` (`:71`; `RiesgoBiblioteca.cls:322`), con comprobación (`:73`).
5. Si era alta, emite `AltaRiesgoBiblioteca` (`:76`) y, si el evento ha dejado el objeto a `Nothing`, pregunta si se quiere otro alta (`:80-84`); si se dice que sí, vacia los tres controles (`:90-92`), pone el objeto a `Nothing` (`:93`) y **no cierra** (`:94`).
6. Si era edición, emite `EdicionRiesgoBiblioteca` (`:78`) y cierra (`:97`).

### Alternativo: ayuda y salida

`ComandoAyuda_Click` abre la ayuda de la pantalla (`:29`; `Funciones Generales.bas:5327`). `ComandoSalir_Click` cierra sin guardar y sin manejador de error (`:115-118`).

### De error

`Form_Load` (`:141-150`), `ComandoGrabar_Click` (`:102-109`), `ComandoAyuda_Click` (`:38-45`) y `EstablecerDatos` (`:211-212`) llevan el patrón común: quitar el reloj y, si `Err.Number <> 1000`, componer el mensaje, escribirlo al administrador (`Funciones Generales.bas:964`) y avisar con diálogo crítico; si es 1000, solo advertencia. `Form_Load` es el único que **cierra siempre**, también en error (`:150`). `EstablecerDatos` no avisa a nadie: solo rellena su argumento de error.

## Reglas, validaciones y cálculos

- **El permiso viene de una global de rol.** Toda la pantalla se bloquea si `EsTecnico = EnumSiNo.Sí` (`:161-163`), una variable global que se calcula en `Variables Globales.bas:348-368` y que no tiene relación con la biblioteca de riesgos ni con el proyecto.
- **La retipificación bloquea también el guardado.** `RiesgoParaRetipificar = Sí` deshabilita edición y botón (`:185-187`), y ese estado no se vuelve a mirar al guardar.
- **Solo se guarda si algo de los tres campos cambió.** `HaHabidoCambios` compara `Familia`, `Activo` y `Descripcion` con `Nz` contra la copia inicial (`:226`, `:230`, `:234`) y devuelve en la primera diferencia; sin copia inicial devuelve `True` sin comparar (`:220-223`).
- **La lista de familias viene de una colección global.** `EstablecerLista` vacía el `RowSource` (`:283`) y lo repuebla con `For Each` sobre `m_ObjEntorno.ColRiesgosBibliotecasFamilias` (`:284-288`).
- **Sin acceso directo a datos.** No hay `getdb()`, `Recordset`, `Open`, `Execute`, SQL, `DLookup` ni transacciones: la persistencia la hace `RiesgoBiblioteca.Registrar` (`RiesgoBiblioteca.cls:322`).

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Riesgo tipo sobre el que se opera | `m_ObjRiesgoBibliotecaActivo` (`Variables Globales.bas:94`), con copia local `m_ObjRiesgoBibliotecaAlIncio` (`:15`) | `:59`, `:71`, `:174`, `:181`, `:220` |
| Familias disponibles | `m_ObjEntorno.ColRiesgosBibliotecasFamilias` (`Variables Globales.bas:73`) | `:284`, `:286` |
| Rol que decide el permiso | `EsTecnico` (`Variables Globales.bas:109`), calculado en `Variables Globales.bas:348-368` | `:161` |
| Título de la ventana | `m_ObjEntorno.TituloUsuarioConectado` | `:191` |
| Recarga del riesgo tipo | `Constructor.getRiesgoBiblioteca` (`Constructor.bas:2885`) | `:181` |
| Guardado | `RiesgoBiblioteca.Registrar` (`RiesgoBiblioteca.cls:322`) | `:71` |
| Avisos al usuario | `AltaRiesgoBiblioteca` (`:76`), `EdicionRiesgoBiblioteca` (`:78`), atendidos en `Form_FormRiesgosBibliotecaGestion.cls:723` y `:758` | `:76`, `:78` |

Escritura efectiva: una llamada a `RiesgoBiblioteca.Registrar`; el resto es estado en memoria y dos eventos.

## Coexistencia

Pendiente de `../coexistencia.md`. Esta pantalla es el caso más sencillo de todos los formularios: tres campos y una escritura. Lo que sí conviene decidir en la web es si el alta encadena un segundo formulario o devuelve el control al usuario con la lista ya actualizada.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-151.** El permiso de escritura de la biblioteca de riesgos no depende de la biblioteca ni del proyecto, sino de la global `EsTecnico` (`:161-163`), que se calcula en otro punto de la aplicación (`Variables Globales.bas:348-368`). Un usuario técnico queda sin poder mantener la biblioteca desde esta pantalla, sin que el motivo sea visible en ella.
- **H-GR-152.** `Form_Load` cierra el formulario **siempre**, también en error (`:150`), al contrario que el resto de la cadena, que deja al usuario en la pantalla con el aviso. Un fallo puntual al cargar deja al usuario sin contexto y con el mensaje en un diálogo que ya no pertenece a ninguna ventana.
- **H-GR-153.** En la ruta de alta, si el evento `AltaRiesgoBiblioteca` deja el objeto a `Nothing`, se pregunta por otro alta y, al decir que sí, se vacían los tres controles y el objeto (`:90-93`) **sin cerrar el formulario** (`:94`). La pantalla queda en pie, vacía, con el botón de guardar habilitado y sin objeto detrás.
- **H-GR-154.** `EstablecerDatos` deshabilita edición y botón cuando `RiesgoParaRetipificar = Sí` (`:185-187`), pero `ComandoGrabar_Click` no vuelve a comprobar ese estado: si el indicador cambia con la pantalla ya abierta, el botón manda y guarda un riesgo tipo en retipificación.
- **H-GR-155.** `EstablecerLista` borra el `RowSource` antes de repoblarlo (`:283`) y, si la colección global viene vacía, sale sin reponer nada (`:284`): la lista de familias queda en blanco sin aviso y el usuario no puede distinguir «no hay familias» de «falló la carga».

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: permiso derivado del recurso y del proyecto en vez de una global de rol; encadenar el alta en una sola operación en lugar de preguntar por un segundo formulario; lista de familias con estado vacío explícito y una distinción clara entre «sin familias» y «falló la carga».

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir la pantalla en alta muestra los tres campos vacíos y la lista de familias cargada.
2. Abrirla en edición muestra los valores del riesgo tipo.
3. Sin permiso de escritura, los campos y el botón de guardar no están disponibles.
4. Guardar sin cambios avisa y no escribe nada.
5. Guardar con cambios registra el riesgo tipo y avisa.
6. Cerrar la pantalla en alta deja la gestión de biblioteca actualizada.

## Pendientes

- Qué tablas toca `RiesgoBiblioteca.Registrar` (`RiesgoBiblioteca.cls:322`) y si deja traza de quién da de alta o modifica un riesgo tipo.
- Qué calcula exactamente `EsTecnico` en `Variables Globales.bas:348-368` y si es una decisión de rol o de contexto de sesión (**H-GR-151**).
- Qué hace el contenedor con `AltaRiesgoBiblioteca` y `EdicionRiesgoBiblioteca` (`Form_FormRiesgosBibliotecaGestion.cls:723` y `:758`) y si puede reposicionar el objeto (`H-GR-153`).
- Si `FormRiesgosBibliotecaGestion.form.txt` mantiene su volcado de código duplicado sincronizado con el `.cls` (p. ej. `.form.txt:2363` frente a `.cls:195`).