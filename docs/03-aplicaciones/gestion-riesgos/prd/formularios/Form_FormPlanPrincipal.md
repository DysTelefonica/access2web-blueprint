# Form_FormPlanPrincipal

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormPlanPrincipal.cls` (199 líneas) y `src/forms/FormPlanPrincipal.form.txt` (1504 líneas; el código empieza en `.form.txt:1314`) · **PRD:** [README](../README.md)

Ficha de **alta** de un plan del riesgo activo. Es una ventana modal sencilla: un solo campo editable (el disparador o la denominación del plan), los datos del riesgo como contexto de solo lectura, y dos botones: grabar y salir. Al grabar crea el objeto de plan en memoria, lo registra y avisa al contenedor mediante un evento. No edita planes existentes ni gestiona acciones: eso lo hace su contenedor. Las citas son `Form_FormPlanPrincipal.cls:línea` salvo indicación.

La distinción **PM / PC** (plan de mitigación frente a plan de contingencia) es el eje de todo el formulario. No hay dos formularios ni dos modos de diseño: hay un parámetro, y cada rama cambia rótulo, título y clase de negocio.

## Entrada y salida

Se abre como ventana modal desde la ficha del riesgo, que cierra antes cualquier instancia previa:

| Quién abre | Parámetro | Evidencia |
|---|---|---|
| `Form_FormRiesgosGestionRiesgo`, alta de contingencia | `"No"` | `Form_FormRiesgosGestionRiesgo.cls:47-49` |
| `Form_FormRiesgosGestionRiesgo`, alta de mitigación | `"Sí"` | `Form_FormRiesgosGestionRiesgo.cls:70-72` |

No acepta nada más. `Me.openArgs()` (`:134`) tiene que valer exactamente `"Sí"` o `"No"`; cualquier otro valor aborta la apertura (`:135-138`). Además exige que exista un riesgo activo global (`:139-142`).

Salidas: **se cierra solo** tras grabar (`:77`), **se cierra solo** si la apertura falla (`:195`) y **se cierra** con `ComandoSalir` (`:100-104`). No hay ningún camino que devuelva un valor: la cesión de responsabilidades se hace por el evento `PlanNuevo` (`:64`, `:74`), no por el cierre.

### El evento `PlanNuevo`

Declarado en `:17` y recogido por `Form_FormRiesgosGestionRiesgo.cls:351-371`, que recarga el nodo del árbol con `CargarArbolPM` (`:356`) o `CargarArbolPC` (`:360`), **eligiendo por el tipo real del objeto recibido** (`TypeOf p_Plan Is PM`, `:355`), y termina seleccionando el nodo nuevo (`:365`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Valida los parámetros, monta los rótulos y decide si se puede grabar | `:118-196` |
| `lblTitulo1` | etiqueta | — | Título del alta; distingue mitigación de contingencia | `:157`, `.form.txt:900-902` |
| `ComandoAyuda` | botón | `ComandoAyuda_Click` | Abre la ayuda del formulario activo | `:19-45`, `.form.txt:920-923` |
| `DisparadorDelPlan` | cuadro de texto | `BeforeUpdate`, `Exit` | Recolorea los bordes de obligatorios | `:107-116`, `.form.txt:1000` |
| `Etiqueta12` | etiqueta | — | Asterisco de campo obligatorio | `.form.txt:1030-1031` |
| `Descripcion` | cuadro de texto | — | Descripción del riesgo, solo lectura | `:145`, `.form.txt:1057` |
| `CausaRaiz` | cuadro de texto | — | Causa raíz del riesgo, solo lectura | `:146`, `.form.txt:1087` |
| `Etiqueta9`, `Etiqueta29` | etiquetas | — | Rótulos de causa raíz y del campo de plan | `.form.txt:1113-1142` |
| `lblDisparador` | etiqueta | — | Su texto pasa a ser «Denominación» (PM) o «Disparador» (PC) | `:150`, `:153`, `.form.txt:1169-1170` |
| `ComandoSalir` | botón | `ComandoSalir_Click` | Cierra sin grabar | `:100-104`, `.form.txt:1207-1208` |
| `ComandoGrabar` | botón | `ComandoGrabar_Click` | Crea y registra el plan | `:47-99`, `.form.txt:1283-1284` |

## Flujos

### Apertura (`Form_Load`, `:118-196`)

1. Reloj de arena y `AjustarTamaño Me` (`:130-133`).
2. Valida el parámetro (`:134-138`) y la existencia del riesgo activo (`:139-142`).
3. Copia descripción y causa raíz del riesgo al formulario (`:145-146`) y compone el título según la rama (`:148-154`).
4. **Sobrescribe el título de la ventana** con el usuario conectado (`:159`), de modo que el rótulo del formulario deja de decir qué está haciendo.
5. Calcula `blnPermitidoEditar` (`:160-169`) y enciende o apaga `ComandoGrabar` y la edición (`:171-177`).
6. `RellenarBordeCamposObligatorios Me` (`:180`).

### Normal — alta (`ComandoGrabar_Click`, `:47-99`)

1. **Si es mitigación** (`:55`): instancia `PM`, le asigna el riesgo activo y el texto del campo (`:56-58`) y la registra (`:60`). **Si es contingencia** (`:66`): el mismo camino con `PC` (`:67-70`).
2. Si `Registrar` devuelve error, `Err.Raise 1000` y salta al manejador (`:61-63`, `:71-73`).
3. **`RaiseEvent PlanNuevo`** con el objeto recién creado (`:64`, `:74`), que es quien inserta el nodo en el árbol.
4. Cierra el formulario (`:77`).

### Alternativo — ayuda

`ComandoAyuda_Click` llama a `AbrirAyuda` (`Funciones Generales.bas:5327`), que resuelve la ayuda **a partir del formulario activo**; por eso este botón depende de tener el foco y no de un identificador propio.

### De error (`:81-98`)

1. Quita el reloj de arena.
2. Si el número no es 1000, compone el mensaje, **notifica por correo al administrador** y muestra un `MsgBox` crítico (`:85-88`); si es 1000, muestra solo la advertencia (`:89-91`).
3. **Vuelve a leer el plan** con `getPM` o `getPC` (`:92-97`) para dejar la sesión coherente.

La validación de fondo no está aquí: `Registrar` llama a `MotivoNoOK` (`PM.cls:788-838`), que exige riesgo conocido, ficha de riesgo completa, edición activa y disparador no vacío.

## Reglas, validaciones y cálculos

- **El parámetro es un texto, no una enumeración.** `m_EsMitigacion` (`:15`) se contrasta con `"Sí"` / `"No"` (`:135`, `:148`, `:152`) mientras el resto de la aplicación usa `EnumSiNo`. La conversión ocurre en el contenedor, no aquí.
- **La comprobación de autorización no cambia el resultado.** En `:163-168` las dos ramas del `If` asignan `blnPermitidoEditar = True`: el permiso depende solo de que la edición esté activa (`:160-161`).
- **El alta no compara con ningún estado anterior.** `Registrar` se invoca sin argumento de instancia previa (`:60`, `:70`), de modo que el objeto se valida como si fuera nuevo aunque el alta se lance sobre un riesgo que ya tenga planes.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Riesgo activo | global de sesión | `:144`, `:147` |
| Edición y permiso | `m_ObjRiesgoActivo.Edicion` | `:143`, `:160-163` |
| Estado del plan | propiedades calculadas de `PM` / `PC` | `PM.cls:321-335` |
| Alta del plan | `Registrar` de la clase de negocio | `:60`, `:70` |

Escritura: una fila de plan, una por alta. Más la inserción del nodo correspondiente en el árbol, que hace el contenedor.

## Coexistencia

Pendiente de `../../forms.md`. Lo relevante aquí es que el formulario **no guarda nada por su cuenta**: el estado vive en objetos de sesión y en las clases de negocio, y la cesión se hace por evento. Es la forma más fácil de portar, porque un formulario web puede recibir el plan ya creado desde el servidor sin imitar la ventana ni el evento.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-42.** La etiqueta del botón de ayuda dice «Salir» (`.form.txt:921`) y su imagen es la de ayuda (`.form.txt:924`), pero el manejador abre la ayuda (`:19`). Hay además un segundo botón, `ComandoSalir`, con el mismo rótulo (`.form.txt:1208`): dos botones con el mismo texto y efectos distintos.
- **H-GR-43.** La comprobación de autorización es inocua: las dos ramas de `:163-168` conceden el permiso. Quien lo lea concluirá que hay un control de usuario autorizado que no existe; el control real vive en `Form_FormRiesgosGestion.blnPermitidoEditar` (`Form_FormRiesgosGestion.cls:202-209`), que este formulario **no consulta**.
- **H-GR-44.** El título de la ventana se sobrescribe con el usuario conectado (`:159`), de modo que se pierde la información de qué se está dando de alta. `Me.Caption` era el rótulo del formulario en diseño (`.form.txt:40`).
- **H-GR-45.** Los dos títulos se componen con convenciones distintas: «ALTA DE PLAN DE MITIGACIÓN PARA» (`:149`) en mayúsculas y «ALTA DE plan de contingencia PARA» (`:152`) en minúsculas. Es el mismo rótulo con dos estilos.
- **H-GR-46.** El manejador de error recarga el plan por identificador (`:92-97`) **después** de un alta fallida. Si la escritura no llegó a producir identificador, `m_ObjPMActivo.IDMitigacion` está vacío y esa relectura falla a su vez, sin manejador propio; conviene verificar ese camino en ejecución.
- **H-GR-47.** Todos los `Err.Raise 1000` (`:29`, `:37`, `:62`, `:72`, `:137`, `:141`) se lanzan sin `vbObjectError` ni `Source`, al contrario que los del resto de la aplicación. Conviene comprobar si Access ejecuta el manejador o si el error sube al formulario que abrió la ventana.
- **H-GR-48.** La respuesta de los diálogos se guarda en `pregunta`, una **global pública compartida por toda la aplicación** (`Variables Globales.bas:69`), no en una variable local. Dos formularios abiertos a la vez se pisan el valor.
- **H-GR-49.** `RellenarBordeCamposObligatorios` (`:108`, `:114`, `:180`) y `EstablecerColorPestañasRiesgo` (`:110`, `:115`) **no están definidas en `src`**: su comportamiento no es verificable desde el repositorio, aunque de ellas depende la señal visual de campo obligatorio.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: devolver el plan creado desde el servidor y cerrar, sin evento ni instancia global; comprobar el permiso en el servidor en vez de duplicar una comprobación que no cambia nada; mantener el rótulo de ventana constante y mostrar el usuario conectado en la barra de la aplicación.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin parámetro o sin riesgo activo avisa y no deja ventana abierta.
2. Abrir con mitigación muestra «Denominación»; con contingencia, «Disparador».
3. Grabar con el campo vacío avisa y no crea nada.
4. Grabar con éxito cierra la ventana y el árbol muestra el plan nuevo ya seleccionado.
5. Sin permiso de edición el botón de grabar aparece deshabilitado.

## Pendientes

- Qué hace exactamente `AbrirAyuda` con un formulario de ventana modal (`Funciones Generales.bas:5327`).
- Dónde están definidas `RellenarBordeCamposObligatorios` y `EstablecerColorPestañasRiesgo`.
- Si la relectura del manejador de error (`:92-97`) es alcanzable con un identificador vacío.
- Qué reglas de dominio devuelve `MotivoNoOK` en la variante de contingencia: el cuerpo leído es el de `PM`, y no se ha confirmado que `PC.cls:793-846` sea idéntico.