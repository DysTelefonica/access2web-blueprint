# Form_FormPublicacionTecnico

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormPublicacionTecnico.cls` (142 líneas) y `src/forms/FormPublicacionTecnico.form.txt` · **PRD:** [README](../README.md)

Contenedor de la parte técnica de la publicación. No edita nada por sí mismo: decide qué pestañas de navegación están disponibles según los datos de la edición activa y deja cargado el subformulario de propuesta. Es un contenedor de tres pestañas alrededor de un mismo contexto de sesión. Las citas son `Form_FormPublicacionTecnico.cls:línea` salvo indicación.

## Entrada y salida

No acepta parámetros ni `OpenArgs`: todo el contexto viene de la global `m_ObjEdicionActiva`. Si esa global está vacía, el formulario se cierra solo (`:91-94`, `:141`).

| Qué decide | Evidencia |
|---|---|
| Título de la ventana: referencia del usuario conectado | `:95`, `m_ObjEntorno.TituloUsuarioConectado` |
| Título interno: proyecto y número de edición | `:98` |
| Qué pestañas se pueden pulsar | `:111-121` |

Salidas: `ComandoSalir` cierra el formulario (`:72`). Desde aquí no se abre ningún otro formulario: los destinos de las pestañas son formularios hijos.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Todo el armado, el permiso y las pestañas | `:75-142` |
| `NavCtrl` | control de navegación | — | Aloja las tres pestañas | `.form.txt:957` |
| `NavPropuesta` | pestaña | — | Propuesta de publicación; su destino se fija en código, no en el diseño | `:124`, `.form.txt:985` |
| `NavUTE` | pestaña | — | Acuerdo UTE; destino `FormPublicacionUTE` fijo en el diseño | `:111-116`, `.form.txt:1035, 1052` |
| `NavAcuerdoSuministradores` | pestaña | — | Solución de riesgos de suministradores; destino `FormPublicacionSuministradores` | `:117-121`, `.form.txt:1089, 1106` |
| `FormDetalle` | subformulario | — | Contiene siempre la propuesta de publicación | `:125`, `.form.txt:1193` |
| `lblTitulo` | etiqueta | — | Proyecto y edición | `:98`, `.form.txt:879-880` |
| `ComandoAyuda` | botón | `ComandoAyuda_Click` | Abre la ayuda; su globo dice «Actualizar» | `:44-69`, `.form.txt:904, 907` |
| `ComandoSalir` | botón | `ComandoSalir_Click` | Cierra | `:71-73`, `.form.txt:1259-1262` |
| `lblEstado` | etiqueta | — | Visible = No en el diseño; el pie es decorativo | `.form.txt:1224, 1233` |
| — | — | `ComandoActualizar_Click` | No existe ningún control con ese nombre en el diseño | `:15-42` |

## Flujos

### Apertura (`Form_Load`, `:75-142`)

1. Reloj de arena y `m_Error` a vacío (`:82-85`).
2. Tamaño fijo en twips: `InsideHeight = 8490`, `InsideWidth = 13220` (`:86-87`). No usa `AjustarTamaño` como el resto de la aplicación.
3. Deja a `Nothing` la global de anexo UTE, pone `blnPermitidoEscribir = False` y vacía `m_URLInforme` (`:88-90`).
4. Si no hay edición activa, escribe el aviso y lanza el error interno 1000 (`:91-94`).
5. Pone el título de la ventana con el usuario conectado (`:95`).
6. Calcula el permiso de escritura (`:100-108`): solo es `True` si la edición está activa y el usuario conectado está autorizado sobre ella. Son dos condiciones, no una.
7. Habilita pestañas según los datos de la edición (`:109-121`):
   - `NavUTE` se habilita si el proyecto declara UTE; al habilitarla, guarda en la global el anexo de evidencia UTE.
   - `NavAcuerdoSuministradores` se habilita si la edición requiere evidencias de suministradores.
8. Fija en código el destino de `NavPropuesta` y el `SourceObject` del subformulario (`:124-125`).
9. `AvanceCerrar` (`:127`) y quita el reloj de arena.

### Alternativo: ayuda

`ComandoAyuda_Click` (`:44-69`) y `ComandoActualizar_Click` (`:15-42`) son idénticos salvo el nombre del manejador: los dos limpian `m_Error`, llaman a `AbrirAyuda` y convierten un resultado no vacío en `Err.Raise 1000`. El segundo no tiene control asociado.

### Error (`:132-141`)

Quita el reloj de arena y, según `Err.Number`:

- Distinto de 1000: compone el mensaje con número y detalle, envía correo al administrador y muestra `MsgBox` crítico (`:135-137`).
- 1000 exacto: muestra `MsgBox` de advertencia con el texto de `m_Error`, sin correo (`:138-139`).

En ambos casos cierra el formulario con `acSaveNo` (`:141`). El cierre va fuera del `If`, así que es incondicional: hasta un fallo ajeno a la lógica de este formulario deja al técnico sin la pantalla de publicación abierta.

## Reglas, validaciones y cálculos

- El permiso es global, no por pestaña. `blnPermitidoEscribir` es una variable global que este formulario escribe (`:89`, `:101`, `:104`, `:106`) y que los formularios hijos leen. El hijo no lo calcula: lo hereda.
- Habilitar no es mostrar. Las pestañas no aplicables se deshabilitan, no se ocultan: el técnico ve las tres y sabe cuáles no le tocan.
- `EnUTE` se compara como texto. `m_EnUTE = "Sí"` (`:111`) depende de que el valor exacto sea ese literal, mientras que el segundo flag se compara contra la enumeración `EnumSiNo.Sí` (`:117`). Misma idea, dos convenciones.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Contexto de edición | global `m_ObjEdicionActiva` | `:91-123` |
| Usuario conectado | `m_ObjEntorno.TituloUsuarioConectado` | `:95` |
| Anexo de evidencia UTE | global `m_ObjAnexoEvicenciaUTE` | `:88`, `:113` |
| Permiso de escritura | global `blnPermitidoEscribir` | `:89`, `:101`, `:104`, `:106` |
| Ruta de informe pendiente | global `m_URLInforme`, puesta a vacío | `:90` |

Sin escrituras en base de datos. Solo escribe en globales de sesión, que es el mecanismo por el que los hijos saben qué pueden hacer.

## Coexistencia

Pendiente de la capa de coexistencia, sin enlace desde esta ficha. Lo relevante aquí es que el permiso y el anexo viajan por globales de proceso, no por la petición del usuario: en una web con varios usuarios concurrentes ese modelo no tiene equivalente directo y obliga a decidir si el permiso se resuelve en el servidor por sesión o se propaga en cada petición.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-330.** `ComandoActualizar_Click` (`:15-42`) es un handler huérfano: no hay ningún control llamado `ComandoActualizar` en el diseño, que solo declara `ComandoAyuda` y `ComandoSalir`. O es código muerto de una versión anterior, o el botón se renombró sin actualizar el evento. Conviene comprobarlo con el propietario.
- **H-GR-331.** El único botón de acción real, `ComandoAyuda`, tiene `ControlTipText = "Actualizar"` (`.form.txt:907`). Al usuario le aparece «Actualizar» sobre un botón que abre la ayuda; su rótulo es además el marcador sin texto `Comando47` (`.form.txt:905`), que se ve literalmente en la interfaz.
- **H-GR-332.** El rótulo de la ventana guardado en el diseño contiene el nombre y apellidos de una persona concreta junto con su perfil y la arquitectura del cliente (`.form.txt:45`). Es dato personal incrustado en el artefacto exportado; no se reproduce aquí. Conviene señalarlo al propietario para decidir si se limpia del binario.
- **H-GR-333.** El cierre incondicional del final del manejador de error (`:141`) hace que cualquier fallo, incluso ajeno a la lógica de este formulario, deje al técnico sin la pantalla de publicación, sin reintento y sin estado. La distinción entre error interno (aviso) y error real (correo al administrador) se pierde justo en el momento en que se cierra.
- **H-GR-334.** Los tres `Err.Raise 1000` (`:93`, `:26`, `:53`) no llevan `vbObjectError`, a diferencia de los errores reales: Access podría tratarlos como errores de usuario y no entrar en el manejador, con lo que el formulario se quedaría abierto a medias o no cerraría.
- **H-GR-335.** `Set m_ObjAnexoEvicenciaUTE = Nothing` (`:88`) se ejecuta antes de la comprobación de edición activa: si el formulario se abre sin contexto, deja esa global a `Nothing` para cualquier otro consumidor.
- **H-GR-336.** `Me.Caption = m_ObjEntorno.TituloUsuarioConectado` (`:95`) accede al entorno sin una comprobación previa de que esté inicializado. Si falla, cae en el manejador y cierra el formulario.
- **H-GR-337.** `AvanceCerrar` (`:127`) se invoca sin argumentos y siempre, incluso si la carga no terminó bien. Conviene comprobar que la barra de progreso no queda desfasada para el resto de la sesión.
- **H-GR-338.** El tamaño fijo en twips (`:86-87`) ignora `AjustarTamaño`, que sí usa el resto de la aplicación: con escalado de Windows o en pantallas de alta resolución este formulario queda más pequeño que los demás.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: resolver el permiso de escritura en el servidor por sesión en lugar de propagarlo por globales; ocultar las pestañas no aplicables en lugar de deshabilitarlas; devolver un error de página en lugar de cerrar el formulario entero ante un fallo de carga; sustituir el subformulario incrustado por navegación por ruta.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin edición activa avisa y no deja la pantalla a medias.
2. Las pestañas no aplicables a la edición no se pueden pulsar.
3. Un usuario sin autorización sobre la edición no puede escribir en ninguna de las tres pestañas.
4. Un fallo de carga deja al técnico con un reintento, no con la ventana cerrada.

## Pendientes

- Si `ComandoActualizar_Click` corresponde a un control eliminado o a código muerto, y cuál de las dos hipótesis es la buena.
- Qué hace `AbrirAyuda` y por qué el botón de ayuda hereda el globo de «Actualizar».
- Quién más consume `blnPermitidoEscribir` y `m_ObjAnexoEvicenciaUTE`; su alcance real no es visible desde aquí.
- Qué hace `AvanceCerrar` con el contador global de pasos.