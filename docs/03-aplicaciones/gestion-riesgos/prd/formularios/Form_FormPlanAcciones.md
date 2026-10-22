# Form_FormPlanAcciones

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormPlanAcciones.cls` (378 líneas) y `src/forms/FormPlanAcciones.form.txt` (1.993 líneas) · **PRD:** [README](../README.md)

Alta de una acción dentro de un plan de mitigación o de contingencia. Es una ventana modal de un solo uso: recibe por `OpenArgs` si el plan activo es de mitigación o de contingencia, escribe una fila de acción nueva y devuelve el evento `AccionNueva` al contenedor. **No edita ni borra**: eso vive en los dos formularios contenedores. Las citas son `Form_FormPlanAcciones.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormRiesgosGestionPlanPrincipal` | `DoCmd.OpenForm "FormPlanAcciones", , , , , , m_EsMitigacion`, previa cierre de la instancia anterior | `Form_FormRiesgosGestionPlanPrincipal.cls:39`, cierre en `:34-35` |

Es la **única** apertura en todo el código fuente. Recibe un único parámetro, `m_EsMitigacion`, con el literal `«Sí»` o `«No»`, que lee en `Form_Load` (`:234`). Los dos contenedores (`FormRiesgosGestionPlanPrincipal` y `FormRiesgosGestionPlanAcciones`) comparten el mismo modelo de datos: los planes en `m_ObjPMActivo` y `m_ObjPCActivo`, el riesgo en `m_ObjRiesgoActivo` y la acción en curso en `m_ObjPMAccionActiva` o `m_ObjPCAccionActiva`.

Salidas:

- **Al contenedor**, mediante el evento `AccionNueva(m_PMAccion As Object)` (`:17`, lanzado en `:85` y `:127`), que el principal atiende con `WithEvents` (`Form_FormRiesgosGestionPlanPrincipal.cls:17, 356`) para refrescar la etiqueta de estado del plan y recargar el árbol (`Form_FormRiesgosGestion.cls`, vía `CargarArbolAccion`).
- **Cierre propio**, sin guardar: `ComandoSalir` (`:182`) y el final del ciclo de alta (`:100` y `:145`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Ajusta el tamaño, lee `OpenArgs` y delega todo el montaje en `EstablecerDatos` | `:224-254` |
| (formulario) | — | — | `EstablecerDatos` construye título, contexto del riesgo, permisos y origen del responsable | `:257-366` |
| `lblTitulo1` | etiqueta | — | Título «ALTA ACCIÓN PARA EL PLAN» + código del plan | `:282, 292`, `.form.txt:901-903` |
| `DisparadorDelPlan` | cuadro de texto | — | Muestra el disparador o la denominación del plan, según el tipo | `:284, 293`, `.form.txt:1240` |
| `Descripcion` / `CausaRaiz` | cuadro de texto | — | Contexto del riesgo: la causa raíz solo si el proyecto exige riesgo de biblioteca calculado, y entonces `Descripcion` reduce su alto y su cuerpo tipográfico | `:333-349`, `.form.txt:1210, 1180` |
| `Accion` / `ResponsableAccion` | texto / combo | `BeforeUpdate`, `Exit` | Delimitan los campos obligatorios y repintan las pestañas; el combo recibe la lista de autorizados | `:24-31, 370-377`, `:358`, `.form.txt:1102, 1141` |
| `FechaInicio`, `FechaFinPrevista`, `FechaFinReal` | cuadro de fecha | `BeforeUpdate`, `Change`, `Exit` | Repintado del color de cada fecha según su cumplimento | `:185-222`, `.form.txt:999, 1033, 1067` |
| `lblDisparador` | etiqueta | — | Alterna su texto entre «Denominación» y «Disparador» | `:283, 292`, `.form.txt:1327` |
| `lblCRaiz` | etiqueta | — | Se oculta cuando el riesgo no procede de biblioteca | `:340, 345`, `.form.txt:1297` |
| `ComandoGrabar` | botón | `Click` | Da de alta la acción y ofrece encadenar otra | `:63-177`, `.form.txt:1590` |
| `ComandoSalir` | botón | `Click` | Cierra sin guardar | `:179-183`, `.form.txt:1516` |
| `ComandoAyuda` | botón | `Click` | Abre la ayuda de la aplicación | `:34-61`, `.form.txt:921` |

## Flujos

### Apertura (`Form_Load`, `:224-254`)

1. Reloj de arena, `AjustarTamaño Me` y `m_EsMitigacion = Nz(Me.openArgs(), "")` (`:231-234`).
2. `EstablecerDatos m_Error`; si devuelve texto, `Err.Raise 1000` (`:236-238`).
3. El manejador de error distingue el error interno 1000 (advertencia, sin correo) del resto (mensaje crítico **y** correo al administrador, `:249`), y **cierra el formulario pase lo que pase** (`:254`).

### Montaje del contexto (`EstablecerDatos`, `:257-366`)

1. Activa `ComandoGrabar` y pone la barra de título con el usuario conectado (`:273-274`).
2. **Ramifica por `m_EsMitigacion`**: si es `«Sí»` exige `m_ObjPMActivo`, toma su riesgo y su edición, rotula «Denominación» y copia el disparador; si no, hace lo propio con `m_ObjPCActivo` y etiqueta «Disparador» (`:276-294`). Si el plan activo no existe, avisa y aborta (`:277-278`, `:287-288`).
3. Copia la descripción del riesgo y, si el proyecto exige riesgo de biblioteca calculado, también la causa raíz (`:296-299`).
4. Calcula si se permite editar según la edición y el usuario (`:303-311`) y lo aplica con `AllowEdits` (`:328-332`).
5. **Si `FormRiesgosGestion` está abierto**, toma el nodo seleccionado del árbol y deduce el nodo del plan: si el nodo es de acción, el nodo del plan es su padre (`:312-325`).
6. Ajusta el alto y el cuerpo de `Descripcion` según el modo de cálculo, y muestra u oculta `CausaRaiz` (`:334-349`).
7. Carga en `ResponsableAccion` la lista de usuarios autorizados del proyecto y comprueba el error acumulado del proyecto (`:352-357`).
8. Delimita los campos obligatorios y repinta las pestañas (`:359-360`).

### Alta (`ComandoGrabar_Click`, `:63-177`)

Flujo espejo en los dos tipos de plan, con dos familias de objetos distintas:

1. **Mitigación** (`:72-114`): crea `PMAccion`, copia el identificador del plan y los cuatro campos del formulario, y llama a `Registrar` (`:73-81`). **Contingencia** (`:116-159`): idéntico con `PCAccion` y el identificador de contingencia (`:117-125`).
2. Si `Registrar` devuelve error, se aborta con el error interno (`:82-84`, `:126-128`).
3. Lanza el evento `AccionNueva` con la acción recién creada (`:85`, `:127`), que es quien refresca el árbol del contenedor.
4. Invalida la caché de acciones del riesgo y repinta el icono del riesgo y el del plan (`:89-96`, `:134-141`).
5. Pregunta «¿Desea dar de alta otra acción» con **Sí por defecto** (`:98`, `:143`).

### Encadenar otra acción (rama alternativa, `:99-113` y `:144-157`)

Si el usuario contesta **Sí**, libera la acción activa, vacía los cinco controles (`Null`, no cadena vacía) y devuelve el foco a `Accion` (`:104-110`, `:149-155`). Si contesta **No**, cierra el formulario sin guardar (`:100`, `:145`).

### Errores (`:168-177`)

Quita el reloj de arena y, según `Err.Number`, compone el mensaje con número y detalle, **envía correo al administrador** y muestra un `MsgBox` crítico (error distinto de 1000), o bien un `MsgBox` de advertencia con `m_Error` (error interno). En ninguno de los dos casos revierte lo ya registrado: la acción queda grabada aunque la pantalla de actualización del árbol falle.

## Reglas, validaciones y cálculos

- **El tipo de plan viaja como cadena, no como enumeración.** `m_EsMitigacion` se compara con el literal `«Sí»` (`:72`, `:116`, `:275`); cualquier otro valor se interpreta como contingencia.
- **La edición se decide antes de pintar, pero el permiso no distingue.** Las dos ramas de `UsuarioAutorizado` dan `True` (`:307-310`): hay que comprobar si el resto de la aplicación resuelve el permiso en otro sitio, porque aquí la condición es un no-op.
- **El origen del responsable se recalcula en cada apertura** a partir de la cadena de autorizados del proyecto (`:352`), no se guarda con la acción.
- **La forma de la pantalla depende de la edición, no de los datos**: sin riesgo de biblioteca calculado se oculta la causa raíz y `Descripcion` crece y sube de cuerpo tipográfico (`:334-341`); con ella, la causa raíz se rellena solo si no está vacía (`:342-349`).
- **El alta no valida las fechas entre sí.** Los tres campos de fecha solo pintan color según su cumplimento (`:185-222`); no hay ninguna comprobación de que la fecha de fin prevista sea posterior a la de inicio, ni de que la real lo sea.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Contexto de sesión | `m_ObjPMActivo`, `m_ObjPCActivo`, `m_ObjRiesgoActivo`, `m_ObjEntorno` | `:274, 280, 290, 357` |
| Nodo seleccionado del árbol | `Form_FormRiesgosGestion.m_NodoSeleccionado` y el combo de vista | `:313-315` |
| Lista de responsables | `Riesgo.Edicion.Proyecto.CadenaNombreAutorizados` | `:352` |
| **Escritura** | `PMAccion.Registrar` / `PCAccion.Registrar`; invalidación de caché con `ColPMs` o `ColPCs` a `Nothing` | `:81`, `:89`, `:125`, `:134` |

## Coexistencia

Pendiente de `../coexistencia.md`. Los puntos que la condicionan: este formulario no tiene identidad propia —depende de cuatro objetos globales de sesión que otro formulario debe haber fijado—, depende del árbol de `FormRiesgosGestion` para localizar el nodo del plan (`:312-325`), y devuelve el resultado por un evento que solo el contenedor principal escucha. Al trasplantarlo, la ventana pierde el modo de instancia única, el reloj de arena y el `WithEvents`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** alta de una acción de plan, con encadenado opcional de varias acciones en la misma sesión, y refresco del árbol y de los iconos por evento.

**Posibles defectos (sin corregir, a verificar):**

- **Hallazgo.** Las dos ramas de la condición de permiso sobre `UsuarioAutorizado` asignan `True` (`:307-310`), de modo que la comprobación no restringe nada. La edición no depende de `Form_FormRiesgosGestion.blnPermitidoEditar`, que sí se usa al borrar en el contenedor (`Form_FormRiesgosGestionPlanAcciones.cls:51`): conviene confirmar si el alta está protegida en otro punto.
- **Hallazgo.** Si falla el repintado del árbol o de los iconos **después** de `Registrar` (`:92`, `:96`, `:137`, `:141`), el error se muestra pero la acción ya está grabada y el árbol no la refleja: el usuario ve un aviso que parece un fallo de alta cuando en realidad es un fallo de refresco.
- **Hallazgo.** El manejador de error de `ComandoGrabar_Click` no distingue el estado del reloj de arena ni libera `m_ObjPMAccionActiva` o `m_ObjPCAccionActiva` (`:168-177`): tras un fallo, la sesión queda con una acción activa que el contenedor puede leer como válida.
- **Hallazgo.** `EstablecerDatos` sobrescribe su parámetro de error en `:353` con el error acumulado del proyecto; si `EstablecerDatos` falla antes, `p_Error` conserva el texto puesto. Es un doble uso del mismo canal para dos orígenes de error distintos.
- **Hallazgo.** Si `FormRiesgosGestion` no está abierto, ni `m_NodoSeleccionado` ni `m_VerDescripcion` ni `m_NodoPlan` se asignan y el alta continúa (`:312-325`): `ActualizarIconoRiesgo` y `ActualizarIconoPlan` reciben `Nothing` y su comportamiento depende del código de esos procedimientos.
- **Hallazgo.** `Err.Raise 1000` se usa sin `vbObjectError` ni tercer argumento en todo el formulario (`:44, 83, 92, 96, 128, 137, 141, 238, 278, 288, 355`); el mensaje de advertencia depende de que `m_Error` esté informado, y en `:238` lo está por el retorno de `EstablecerDatos`, lo que funciona, pero el patrón es frágil ante cualquier nueva rama.
- **Hallazgo.** La pregunta de encadenado usa `vbDefaultButton1` (`:98`, `:143`), de modo que el botón por defecto es «Sí»: un `Intro` accidental crea una segunda acción en blanco.
- **Hallazgo.** El formulario se cierra siempre con `acSaveNo` (`:100, 145, 182, 254`), incluso tras un alta correcta: cualquier dato tecleado y no confirmado por `ComandoGrabar` se pierde sin aviso, y no hay aviso de cambios pendientes como el que sí implementa el contenedor (`HaHabidoCambios`, `Form_FormRiesgosGestionPlanAcciones.cls:398`).
- **Hallazgo.** `ComandoGrabar` se activa siempre al inicio (`:273`) y nunca se desactiva: el control de habilitación no depende del permiso ni del estado de los campos obligatorios, que se marcan solo por color de borde (`:359`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: sustituir los objetos de sesión por un identificador de plan en la ruta y devolver la acción creada al servidor; modelar el encadenado como varias filas en un formulario en lugar de un bucle modal; mover la validación de fechas al servidor; devolver un identificador de la acción creada en la respuesta en vez de depender de un evento en memoria.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir el alta desde un plan de mitigación titula el plan como mitigación y rotula el campo como «Denominación»; desde contingencia, como «Disparador».
2. Grabar con la acción y el responsable informados crea la acción, la muestra en el árbol del contenedor y ofrece encadenar otra.
3. Encadenar vacía los campos y mantiene la sesión abierta; no encadenar cierra la ventana.
4. Sin plan activo, el alta no se abre y se explica el motivo.
5. Un proyecto que no exige riesgo de biblioteca calculado no muestra la causa raíz.
6. Si el refresco del árbol falla, la acción sigue creada y el aviso distingue alta correcta de refresco fallido.
7. Cerrar sin grabar no deja datos ni crea ninguna acción.

## Pendientes

- Qué garantiza exactamente que `m_EsMitigacion` vale `«Sí»` o `«No»` y no llega vacío (`Nz(Me.openArgs(), "")`, `:234`): un valor vacío se interpreta como contingencia.
- Si el permiso de alta se comprueba en algún punto anterior a este formulario, dado que aquí no discrimina (`:307-310`).
- Qué hacen `ActualizarIconoRiesgo` y `ActualizarIconoPlan` cuando reciben un nodo `Nothing` (`:90`, `:94`).
- Qué valida `Registrar` en `PMAccion` y `PCAccion`: si el propio objeto cubre fechas y campos obligatorios, parte de los hallazgos sobre validación quedan sin objeto.
- Si existe en el modelo alguna noción de acción inversa o de reversión de un alta: en este commit fijo no aparece en el código fuente de `src/`, y el árbol solo distingue plan de mitigación, plan de contingencia y sus acciones ([../../forms.md](../../forms.md)).