# Form_Form0BDGestorEntidades

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_Form0BDGestorEntidades.cls` (187 líneas), `.form.txt` y `src/modules/Form0BDGestorEntidadesLogic.bas` · **PRD:** [README](../README.md)

Menú «ENTIDADES» (`.form.txt:835-836`): una lista de opciones y un botón que abre el formulario de gestión de la entidad elegida (comerciales, CPV, ejércitos, empresas, lugares de ejecución, oficinas de programa, órganos de contratación, PECAL, RAC, grados de clasificación, gestor E2E y panel de cola E2E). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) y no lee ni escribe en la base de datos. Las citas son `Form_Form0BDGestorEntidades.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpciones` | `ComandoGestorEntidades_Click`: cierra el gestor si estaba abierto y lo abre sin `OpenArgs`; también por doble clic en su lista | `Form_Form0BDOpciones.cls:196-206, 353-354` |

El botón de origen solo se habilita si `EsAdministrador = Sí` (`Form_Form0BDOpciones.cls:254-256`). Salidas: `cmdSalir` cierra sin guardar (`:8-12`); tras abrir un destino, el gestor se cierra solo (`:76`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaOtrasOpciones` | lista, 2 columnas (clave oculta, `ColumnWidths ="0;4081"`) | `OnClick` → `ListaOtrasOpciones_Click` | Habilita `ComandoAbrirFormulario` solo si hay clave seleccionada | `:153-179`, `.form.txt:940-954` |
| `ListaOtrasOpciones` | — | `OnDblClick` → `ListaOtrasOpciones_DblClick` | Si el botón está habilitado, ejecuta `ComandoAbrirFormulario_Click` | `:181-187` |
| `ComandoAbrirFormulario` (rótulo «Editar») | botón | `OnClick` | Resuelve el destino y lo abre; cierra el gestor | `:41-89`, `.form.txt:994-996` |
| `ComandoAyuda` (rótulo «Salir» en `.form.txt`) | botón | `OnClick` | `AbrirAyuda` (`FUNCIONES UTILES.bas:3076`) | `:91-113`, `.form.txt:855-857` |
| `cmdSalir` | botón | `OnClick` | `DoCmd.Close` con `On Error Resume Next` | `:8-12`, `.form.txt:1081-1083` |
| `lblTitulo1` | etiqueta | — | Rótulo fijo «ENTIDADES» | `.form.txt:835-836` |

La lista tiene en diseño un `RowSource` de tipo «Value List» con claves y textos antiguos (`.form.txt:948-950`); en ejecución se vacía y se rellena desde `Entorno.OtrasOpciones` (`:23-32`).

## Flujos

### Apertura (`Form_Open`, `:115-151`)

1. `Ajustar Me` (`FUNCIONES UTILES.bas:2703`) (`:123`).
2. Oculta la cinta salvo con `TempVars("EnDesarrollo") = "Sí"` (`:125-129`).
3. Título del formulario desde la global `m_TituloFormulario` (`:130`); el `Caption` de diseño (`.form.txt:39`) queda sustituido.
4. Habilita `ComandoAbrirFormulario` (`:131`) y llama a `EstablecerLista` (`:132`), que recorre el diccionario `m_ObjEntorno.OtrasOpciones` y añade `clave;texto` por cada entrada (`:24-32`). El diccionario se construye una vez y queda en caché (`classes/Entorno.cls:682-701`); el orden de la lista es el de inserción.

### Abrir una entidad (`ComandoAbrirFormulario_Click`, `:41-89`)

1. Deshabilita el botón (`:53`). Sin selección, error «No se ha seleccionado ninguna opción» (`:54-57`).
2. `ResolverFormDestino(clave)` traduce la clave al nombre del formulario (`Form0BDGestorEntidadesLogic.bas:12-37`); clave desconocida ⇒ cadena vacía y error «No se ha seleccionado ninguna entidad reconocida» (`:58-63`).
3. Si `TipoAccion(clave) = "html"` (solo la clave 12, `Form0BDGestorEntidadesLogic.bas:55-58`) llama a `E2EColaHtmlSnapshot_Open` en lugar de abrir `FormE2ECola` (`:67-69`). En otro caso, `DoCmd.OpenForm` con o sin `OpenArgs` (`:70-74`).
4. Llama a `ListaOtrasOpciones_Click` para recalcular el botón (`:75`) y cierra el gestor (`:76`).

### Panel de cola E2E (clave 12)

`E2EColaHtmlSnapshot_Open` inicializa el entorno si falta (`E2EColaHtmlSnapshot.bas:65-71`), exige administrador o gestor con `ControlPanelGateAllowed` (`:75-79`; `E2EEnlacesService.bas:374-377`), borra los HTML previos, genera el informe con `E2EColaHtmlSnapshot_BuildReport` y lo abre con `Application.FollowHyperlink` (`E2EColaHtmlSnapshot.bas:82-108`). Sus errores se muestran con `MsgBox` propio (`:111-121`) y no llegan al gestor, que se cierra igualmente (`:76`).

### Errores

Todos los handlers salvo `cmdSalir` y el doble clic acumulan el error en `m_Error`, envían `CorreoAlAdministrador` salvo el error interno 1000 y lo muestran con `ShowErrorIfNotSuppressed`, que no muestra nada si `g_SuppressMsgBox` está activo (`:81-88`; `modUiSuppress.bas:19-24`). Un error en `Form_Open` cierra el formulario (`:150`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Opciones del menú: claves 1-12 de `EnumOtrasOpciones` con su texto | `Entorno.OtrasOpciones` | `classes/Entorno.cls:690-701`; `Variables Globales.bas:146-158` |
| Clave → formulario (1 comerciales … 12 cola E2E; 4 «Empresas» → `FormSuministradoresGestion`) | `ResolverFormDestino` | `Form0BDGestorEntidadesLogic.bas:13-36` |
| Clave 12 se resuelve como HTML, no como formulario | `TipoAccion` | `Form0BDGestorEntidadesLogic.bas:50-63` |
| Botón «Editar» solo habilitado con selección | `ListaOtrasOpciones_Click` | `:162-165` |
| Acceso al gestor: administrador (en el formulario de origen) | `Form_Form0BDOpciones` | `Form_Form0BDOpciones.cls:254-256` |
| Gestor E2E y cola E2E repiten su propia comprobación de rol | formularios destino | `Form_FormE2EGestionEnlaces.cls:241`, `Form_FormE2ECola.cls:36` |

## Lecturas y escrituras

Ninguna propia. Solo lee la caché en memoria de `Entorno` (`:24`). La opción 12 lee la base con `getdb` dentro de `E2EColaHtmlSnapshot_Open` (`E2EColaHtmlSnapshot.bas:90`); su contenido corresponde a la ficha de ese módulo. La persistencia de cada entidad es de los formularios destino.

## Coexistencia

Sin SQL ni transacciones: no hay adaptación a PostgreSQL en este formulario. Debe preservarse el catálogo de opciones y su correspondencia con los formularios destino. Ver `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La correspondencia clave → formulario está duplicada: `ResolverFormDestino` la codifica a mano (`Form0BDGestorEntidadesLogic.bas:12-37`) y `Entorno.OtrasOpcionesFormularios` (`classes/Entorno.cls:715`) la repite, aunque fuera de las pruebas nadie la usa; pueden divergir.
- `ResolverFormDestino` nunca asigna `p_OpenArgs` distinto de vacío (`Form0BDGestorEntidadesLogic.bas:10`): la rama con `OpenArgs` es código muerto (`:72-73`).
- Si falla la resolución o el `OpenForm`, el botón queda deshabilitado (`:53`) hasta volver a pulsar la lista.
- `ListaOtrasOpciones_Click` se invoca justo antes de cerrar el formulario (`:75-76`): no tiene efecto visible.
- El gestor se cierra aunque el panel HTML falle o se deniegue por rol (`:69, 76`; `E2EColaHtmlSnapshot.bas:75-79`).
- Comentarios obsoletos: atribuyen la bifurcación HTML a `ListaOtrasOpciones_Click` (`Form0BDGestorEntidadesLogic.bas:48`, `E2EColaHtmlSnapshot.bas:44`) cuando está en `ComandoAbrirFormulario_Click` (`:67-69`), y citan `Form_Form0BDGestorEntidades.cls:222`, línea inexistente (`E2EColaHtmlSnapshot.bas:64`).
- El `RowSource` de diseño usa otras claves y textos (p. ej. clave 11 con otro significado, `.form.txt:949`); solo es inocuo mientras `EstablecerLista` lo sobrescriba.
- `ComandoAyuda` rotula «Salir» y `ComandoAbrirFormulario` «Editar» en el `.form.txt` (`.form.txt:856, 995`); no se ha visto código que los cambie.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: sustituir el menú por navegación con rutas por entidad; una única tabla de configuración clave → destino; aplicar el control de rol en el propio menú y ocultar las opciones no permitidas.

## Criterios de aceptación

Propuestos, no aprobados:

1. El menú lista las doce opciones en el orden de `Entorno.OtrasOpciones`.
2. Sin selección no se puede abrir ninguna entidad.
3. Cada opción abre su pantalla de gestión según la tabla de `ResolverFormDestino`.
4. «Panel de cola E2E» genera y abre el informe solo para administradores o gestores; al resto muestra «acceso restringido».
5. Solo los administradores llegan al menú desde la pantalla principal.

## Pendientes

- Estado inicial (`Enabled`) de `ComandoGestorEntidades` en el `.form.txt` de `Form0BDOpciones`: no localizado, y con él si un no administrador puede llegar al gestor.
- Contenido de `Ajustar` y `AbrirAyuda` (`FUNCIONES UTILES.bas:2703, 3076`).
- Origen de `m_TituloFormulario` y de `TempVars("EnDesarrollo")`.
- Comportamiento de `E2EColaHtmlSnapshot_BuildReport` (fuera del alcance de esta ficha).
