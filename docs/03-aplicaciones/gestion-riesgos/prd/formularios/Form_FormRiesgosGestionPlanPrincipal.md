# Form_FormRiesgosGestionPlanPrincipal

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormRiesgosGestionPlanPrincipal.cls` (400 líneas) y `src/forms/FormRiesgosGestionPlanPrincipal.form.txt` (1619 líneas; el código empieza en `.form.txt:1225`) · **PRD:** [README](../README.md)

Contenedor de la edición del plan del riesgo seleccionado. A diferencia de su hermano `Form_FormPlanPrincipal`, que da de alta, este **no se abre**: se carga como subformulario en el área de detalle del árbol de riesgos y cambia de contenido cada vez que el usuario selecciona un nodo distinto. Edita un único campo del plan y da de alta las acciones asociadas. Las citas son `Form_FormRiesgosGestionPlanPrincipal.cls:línea` salvo indicación.

La distinción **PM / PC** se resuelve aquí: no por parámetro, sino leyendo el flag que el contenedor dejó al seleccionar el nodo. Toda la ficha es, en la práctica, un `If m_EsMitigacion = "Sí"` repetido.

## Entrada y salida

No tiene `OpenArgs`. Se instala como `SourceObject` del control de detalle del formulario contenedor:

| Contexto | Rama | Evidencia |
|---|---|---|
| Se selecciona un nodo de mitigación | `Me.FormDetalle.SourceObject = "FormRiesgosGestionPlanPrincipal"` | `Form_FormRiesgosGestion.cls:340` |
| Se selecciona un nodo de contingencia | el mismo `SourceObject` | `Form_FormRiesgosGestion.cls:384` |

El subformulario se recarga en cada cambio de `SourceObject`, de modo que su `Form_Load` (`:210-238`) es de hecho **la routine de cambio de contexto**. No existe forma de abrirlo aislado.

La información de contexto no llega por parámetros sino por estado compartido: el flag `Form_FormRiesgosGestion.m_EsMitigacion`, los objetos globales `m_ObjPMActivo` / `m_ObjPCActivo`, el permiso `blnPermitidoEditar` y el nodo seleccionado del árbol.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Carga los datos mediante `EstablecerDatos` y cierra si falla | `:210-238` |
| `Cuadro38` | marco | — | Agrupa el bloque de estado del plan | `.form.txt:905` |
| `DisparadorDelPlan` | cuadro de texto | `BeforeUpdate`, `Exit` | Recolorea los bordes de obligatorios | `:202-208`, `.form.txt:926` |
| `EstadoCompletoPlanEstado` | etiqueta | — | Estado del plan en texto | `:297`, `:305`, `.form.txt:962` |
| `EstadoCompletoFechaActivacion` | etiqueta | — | Primera fecha de inicio, en texto | `:298`, `:306`, `.form.txt:995` |
| `EstadoCompletoFechaDesactivacion` | etiqueta | — | Fecha de desactivación, en texto | `:299`, `:307`, `.form.txt:1027` |
| `lblTitulo` | etiqueta | — | Identifica el plan y su riesgo | `:295`, `:303`, `.form.txt:1059` |
| `lblDisparador` | etiqueta | — | «Denominación» en PM, «Disparador» en PC | `:267`, `:277`, `.form.txt:1086` |
| `ComandoAltaAccion` | botón | `ComandoAltaAccion_Click` | Abre la ficha de acción en ventana aparte | `:26-56`, `.form.txt:1117-1118` |
| `ComandoEliminar` | botón | `ComandoEliminar_Click` | Borra el plan completo | `:58-131`, `.form.txt:1156-1157` |
| `ComandoGrabar` | botón | `ComandoGrabar_Click` | Guarda el campo editado y refresca el árbol | `:133-201`, `.form.txt:1194-1195` |

## Flujos

### Apertura (`Form_Load` → `EstablecerDatos`, `:210-238`, `:240-323`)

1. `EstablecerDatos` se ejecuta y su error se comprueba (`:222-224`); si falla, el manejador cierra el formulario (`:238`).
2. Dentro de `EstablecerDatos`: exige que el contenedor sepa si es mitigación o contingencia (`:248-250`) y traduce la enumeración a texto (`:252-256`).
3. Si es PM, exige un plan activo, carga una **copia del estado de apertura** con `getPM` (`:258-265`) y rotula el campo «Denominación» (`:267`); si es PC, el mismo camino con `getPC` (`:269-278`).
4. Enciende o apaga los tres comandos y la edición según el permiso (`:279-290`).
5. Rellena título, campo y las tres etiquetas de estado (`:293-309`).
6. Si el contenedor está abierto, guarda el nodo seleccionado y la preferencia de vista (`:311-314`).

### Normal — grabación (`ComandoGrabar_Click`, `:133-201`)

1. Si no hay permiso, **sale sin decir nada** (`:141-143`).
2. Vuelca el campo del formulario sobre el objeto activo, no sobre la copia de apertura (`:145-148`).
3. `HaHabidoCambios` compara (`:149-152`); si no hay cambios, avisa «No hay nada que grabar» (`:153-156`).
4. `Registrar` pasando la copia de apertura como estado anterior (`:158`, `:165`).
5. Refresca el icono del nodo (`:171-174`).

### Normal — alta de acción (`ComandoAltaAccion_Click`, `:26-56`)

Cierra `FormPlanAcciones` si estuviera abierta (`:34-36`), **anula los dos objetos de acción activos** (`:37-38`), la abre pasando el flag (`:39`) y se suscribe a su evento `AccionNueva` mediante `WithEvents` (`:17`).

### Alternativo — recepción de una acción nueva (`m_FormAccion_AccionNueva`, `:356-400`)

1. Decide mitigación o contingencia **por los dos primeros caracteres de la clave del nodo** (`:365`, `:371`).
2. Refresca las tres etiquetas de estado (`:366-374`).
3. Inserta el nodo de la acción en el árbol del contenedor (`:381`).

### De error (`:189-198`, `:123-131`, `:392-400`)

Los tres manejadores siguen la convención de la aplicación: número distinto de 1000 compone mensaje, **notifica por correo al administrador** y muestra `MsgBox` crítico; 1000 exacto muestra solo advertencia. Tras un fallo de grabación, el manejador **recarga el plan desde el constructor** (`:191-197`).

## Reglas, validaciones y cálculos

- **El tipo de plan se decide en tres sitios distintos**: por el flag del contenedor (`:252`), por el tipo del objeto global (`:258`, `:269`) y por el prefijo de la clave del nodo (`:365`, `:371`). Los tres tienen que coincidir y nada lo comprueba.
- **El permiso se lee del contenedor, no se calcula.** Se consulta `Form_FormRiesgosGestion.blnPermitidoEditar` (`:279`), que a su vez lo calcula `Form_FormRiesgosGestion.cls:202-209`. El subformulario hereda la decisión.
- **La copia de apertura se usa solo para detectar cambios.** `m_ObjPMAlInicio` / `m_ObjPCAlInicio` (`:18-19`) alimentan `HaHabidoCambios` (`:330-352`) y se pasan a `Registrar`; la asignación del campo, en cambio, va al objeto activo (`:145-148`).
- **Borrar el plan es una operación de una sola pieza**, con confirmación y aviso de éxito (`:87-89`, `:110`). No distingue si el plan tiene acciones: la integridad la resuelve `Borrar` (`PM.cls:993`).

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Tipo de plan y permiso | estado del contenedor | `Form_FormRiesgosGestion.cls:340`, `:384`, `:202-209` |
| Copia de apertura | `Constructor.getPM` / `getPC` | `Constructor.bas:2012`, `:2218` |
| Estado legible | propiedades calculadas de `PM` / `PC` | `PM.cls:321-335` |
| Nodo del árbol | `Form_FormRiesgosGestion.m_NodoSeleccionado` | `:312` |
| Grabación | `Registrar` con estado anterior | `:158`, `:165` |
| Baja del plan | `Borrar` de la clase de negocio | `:96`, `:99` |
| Inserción de la acción | `CargarArbolAccion` del contenedor | `:381`, `Form_FormRiesgosGestion.cls:1598` |

## Coexistencia

Pendiente de `../../forms.md`. Aquí se ve el patrón que más cuesta portar: **la ventana de acción vive fuera del subformulario** y se comunica con él por evento (`WithEvents`, `:17`), mientras que el subformulario, a su vez, depende por nombre de otro formulario abierto. Es una cadena de tres formularios vivos a la vez, gobernada por variables globales de sesión.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-50.** `m_NodoSeleccionado` solo se asigna si el contenedor está abierto (`:311-314`). Si no lo está, la variable queda vacía y `m_FormAccion_AccionNueva` la desreferencia sin comprobación (`:365`), con lo que el alta de acción depende de un orden de ventanas que nada garantiza.
- **H-GR-51.** La rama PM / PC del evento se decide por `Left(m_NodoSeleccionado.Key, 2)` (`:365`, `:371`). Es una convención sobre la nomenclatura de las claves del árbol; si una clave deja de empezar por `PM` o `PC`, la acción se registra contra el objeto equivocado y no hay aviso.
- **H-GR-52.** `ComandoGrabar` sale **en silencio** cuando no hay permiso (`:141-143`). El botón ya está deshabilitado por `:287-288`, así que el camino solo se alcanza si el permiso cambia entre la carga y el clic; en ese caso el usuario pulsa y no ocurre nada.
- **H-GR-53.** Tras `Borrar` se vacía la colección del riesgo (`:97`, `:100`) pero **no se pone a `Nothing` el objeto del plan activo**. El subformulario sigue cargado con una referencia a un plan que ya no existe, y un `ComandoGrabar` posterior escribiría sobre una instancia borrada.
- **H-GR-54.** El aviso «Plan borrado con éxito» (`:110`) se muestra **antes** de tocar el árbol, y el orden de la actualización es contradictorio: se selecciona el nodo padre (`:111`), se dispara su manejador de clic (`:112`) y solo después se elimina el nodo borrado (`:116`). El `MsgBox` afirma un resultado que todavía no se ha aplicado en pantalla.
- **H-GR-55.** `m_NodoActual` (`:83`) no se comprueba antes de usarlo (`:111-116`). `SelectedItem` devuelve `Nothing` si el árbol no tiene selección, y entonces el borrado ya se ha ejecutado pero la actualización del árbol falla.
- **H-GR-56.** `Registrar` recibe la copia de apertura (`PM.cls:842-846`) y se la pasa a `MotivoNoOK` (`PM.cls:851`), pero **`MotivoNoOK` no la lee en ningún punto de su cuerpo** (`PM.cls:788-838`): el parámetro se acepta y se descarta. El estado anterior no participa en la validación, que es idéntica para alta y edición.
- **H-GR-57.** `CambiaIconoElementoArbolDeRiesgos` (`:171`) es un método del **control de árbol embebido** (`Arbol.Object`) y no está definido en `src`. De la actualización del icono tras grabar depende la señal visual del plan en el árbol, y no es verificable desde el repositorio.
- **H-GR-58.** Los manejadores de error guardan la respuesta del diálogo en `pregunta`, una global pública compartida (`Variables Globales.bas:69`). Con la ventana de acción abierta sobre el subformulario, ambos formularios pueden pisarse el valor.
- **H-GR-59.** `EstablecerDatos` (`:240-323`) es un `Function` cuyo valor de retorno no se usa: la firma declara `As String` (`:242`) y nunca asigna a `EstablecerDatos`; el resultado viaja solo por el parámetro de error. Conviene confirmar que no haya un segundo llamante que sí espere el valor.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: un único formulario de plan con ruta de alta y de edición en la misma pantalla; sustituir la comunicación por evento entre la ventana de acción y el detalle por una acción de servidor que devuelva el plan y la acción afectada; decidir el tipo de plan en un solo punto y propagarlo como dato, no por prefijo de clave de árbol; devolver el estado del plan como datos y no como etiquetas preformateadas.

## Criterios de aceptación

Propuestos, no aprobados:

1. Seleccionar un nodo de mitigación muestra la edición de su plan; seleccionar uno de contingencia, la del suyo.
2. Abrir el detalle sin un plan detrás avisa y no deja el formulario en blanco.
3. Cambiar el campo y grabar actualiza el plan y el icono del árbol.
4. Grabar sin cambios avisa y no escribe.
5. Dar de alta una acción desde el detalle añade su nodo al árbol y refresca el estado del plan.
6. Borrar un plan pide confirmación y desaparece del árbol junto con sus acciones.

## Pendientes

- Si `MotivoNoOK` en `PC.cls:793-846` replica el cuerpo de `PM.cls:788-838` y confirma el mismo descarte del parámetro de estado anterior.
- Qué comprobaciones hace `Borrar` sobre las acciones y las acciones inversas (`PM.cls:993`, `PC.cls:1001`).
- Qué hace exactamente `CambiaIconoElementoArbolDeRiesgos` y dónde está definido.
- Si el evento `AccionNueva` de `Form_FormPlanAcciones` distingue mitigación y contingencia, dado que recibe un único parámetro sin tipo.