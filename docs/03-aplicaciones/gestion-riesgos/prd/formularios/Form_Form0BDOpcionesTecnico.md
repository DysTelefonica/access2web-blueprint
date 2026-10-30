# Form_Form0BDOpcionesTecnico

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_Form0BDOpcionesTecnico.cls` (314 líneas, 10 procedimientos) y `src/forms/Form0BDOpcionesTecnico.form.txt` · **PRD:** [README](../README.md)

Menú principal del perfil técnico. Es una reducción de `Form_Form0BDOpciones`: mismo esqueleto (patrón `DoEvents` → reloj de arena → `m_Error` → acción, mismo manejador de error), menos botonera y con dos diferencias que sí importan: la puerta de retorno al perfil inicial y el salto de `EVE`. No comprueba `EsTecnico` en ningún handler. Las citas son `Form_Form0BDOpcionesTecnico.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `frmSplash` | `DoCmd.OpenForm` cuando el usuario es técnico | `Form_frmSplash.cls:72` |
| `Form0BDOpciones` | `ComandoEntrarComo`, con el argumento «Para Otro Usuario» | `Form_Form0BDOpciones.cls:86` |

**Sí acepta un parámetro**, y le importa: `Form_Load` solo llama a `EVE` si `OpenArgs` viene vacío (`:235-240`). Entrado desde el menú de calidad, el entorno **no** se reinicializa.

Salidas: `FormProyectosGestion` (`:69`), `FormTecnicoTareas` (`:96`), `FormRiesgosGestion` (`:138`), `FormControlCambiosGestion` (`:170`), el menú de calidad (`:205`), la ayuda (`:35`), el vídeo externo (`:300`) y la aplicación (`:23`).

## Controles y eventos

| Control | Evento y handler | Qué abre | Evidencia |
|---|---|---|---|
| (formulario) | `Form_Load` | Rótulos, puerta de retorno y botonera | `:223-287`, `.form.txt:221` |
| `ComandoTareasPendientes` | `ComandoTareasPendientes_Click` | `FormTecnicoTareas`, sin comprobación de rol | `:88-111`, `.form.txt:1626` |
| `ComandoGestionesDeRiesgo` | `ComandoGestionesDeRiesgo_Click` | `FormProyectosGestion` tras refrescar objetos activos | `:54-85`, `.form.txt:941` |
| `ComandoUltimoProyecto` | `ComandoUltimoProyecto_Click` | `FormRiesgosGestion`, o la lista si no hay último proyecto | `:113-153`, `.form.txt:1254` |
| `ComandoVerControlCambios` | `ComandoVerControlCambios_Click` | `FormControlCambiosGestion` | `:163-186`, `.form.txt:1878` |
| `ComandoVolverAlPerfilInicial` | `ComandoVolverAlPerfilInicial_Click` | `Form0BDOpciones` tras cerrar sesión y llamar a `EVE` | `:188-221`, `.form.txt:2693` |
| `ComandoAyuda` | `ComandoAyuda_Click` | Ayuda de la aplicación | `:26-52`, `.form.txt:849` |
| `cmdSalir` | `CmdSalir_Click` | Cierra la aplicación | `:15-24`, `.form.txt:2605` |
| `lblVideo` | `lblVideo_Click` | Navegador externo al vídeo configurado | `:289-314`, `.form.txt:2540` |

Los rótulos (`lblTareasPendientes`, `lblMisProyectos`, `lblUltimoProyecto`, `lblControlCambios`) son etiquetas fijas sin evento (`.form.txt:2224-2322`). **No hay `Form_Activate`**: el diseño solo declara `OnLoad` (`.form.txt:221`), a diferencia del menú de calidad, que declara `OnActivate` (`.form.txt:221` del otro diseño).

## Flujos

### Apertura (`Form_Load`, `:223-287`)

1. Reloj de arena y `AjustarTamaño Me` (`:234`).
2. `EVE` solo si `OpenArgs` está vacío (`:235-240`).
3. `Me.Caption`, `lblVersion` y `lblUltimoProyecto` (`:246-255`).
4. **Puerta de retorno**: si `m_ObjUsuarioConectadoInicialmente` existe y su `EsTecnicoCalculado` **no** es `Sí`, muestra y habilita `ComandoVolverAlPerfilInicial`; en cualquier otro caso lo oculta (`:257-267`).
5. `ActivarBotonera` y `AvanceCerrar` (`:269-270`).

Su manejador de error, como en el menú de calidad, **termina la aplicación** (`:286`).

### Normal

`DoEvents` → `Hourglass True` → `m_Error = ""` → acción → `Hourglass False`. Errores previsibles con `Err.Raise 1000` y aviso; los demás, correo al administrador y error crítico (`:45-51`).

### Alternativo

- **Sin último proyecto:** `ComandoUltimoProyecto_Click` delega en `ComandoGestionesDeRiesgo_Click` (`:126-131`) y, a diferencia del gemelo, toma el proyecto directamente de `m_ObjUltimoProyecto.Proyecto` (`:132-133`) sin volver a consultar `Constructor`.
- **Volver al perfil inicial** (`:195-205`): cierra los formularios abiertos, pone `m_ObjUsuarioConectadoInicialmente = Nothing`, llama a `EVE` sin argumentos y abre el menú de calidad.

### De error

`CmdSalir_Click` y `ActivarBotonera` usan `On Error Resume Next` sin manejador (`:17`, `:155`). `Form_Load` y `ComandoVolverAlPerfilInicial` son los dos únicos que pueden acabar la sesión o cambiar de usuario.

## Reglas, validaciones y cálculos

- **La condición es negativa y está en la apertura:** `m_ObjUsuarioConectadoInicialmente.EsTecnicoCalculado <> EnumSiNo.Sí` (`:259`) habilita el retorno. El atributo `EsTecnicoCalculado` es del **usuario inicial**, no del actual.
- **Ningún handler mira el perfil.** Los cuatro comandos se ejecutan sin comprobación, incluido `ComandoTareasPendientes`, que abre `FormTecnicoTareas` sin validar nada (`:96`).
- **`ActivarBotonera` habilita cuatro controles** (`:154-160`), uno menos por duplicado que su gemelo, y **no incluye** `ComandoVolverAlPerfilInicial`: el estado de ese botón queda fijado en la apertura y no se revisa al recuperar el foco.
- No expone el cambio de modelo de árbol: `EstablecerModeloCadenaJerarquica` solo existe en el menú de calidad (`Form_Form0BDOpciones.cls:23`).

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Identidad, tareas, cachés | Globales de sesión `m_ObjEntorno`, `m_ObjUltimoProyecto` | `Variables Globales.bas:73, 130` |
| Usuario inicial de la sesión | `m_ObjUsuarioConectadoInicialmente` | `Variables Globales.bas:104`, `:199` |
| Reinicialización del entorno | `EVE` | `Variables Globales.bas:204`, `:201, 236` |
| Rótulos de avance | `AvanceCerrar` | `Funciones Generales.bas:4920` |

Sin escrituras en base de datos.

## Comportamiento actual

Panel de navegación con cuatro destinos, un botón de retorno al perfil inicial que solo aparece cuando la sesión empezó siendo de calidad o administración, y salida. El destino de «tareas pendientes» es el formulario del propio perfil, lo que resuelve la divergencia con el menú de calidad sin necesidad de comprobación en ejecución.

## Posibles defectos (sin corregir, a verificar)

- **H-GR-30.** El control de acceso **no está en los handlers sino en la visibilidad** de un botón (`:259-267`). Si el diseño no declara la propiedad `Visible` de `ComandoVolverAlPerfilInicial` (`.form.txt:2693-2695`), el control aparece con la visibilidad por defecto hasta que `Form_Load` la ajusta: conviene verificar en ejecución si hay un destello del botón en sesiones iniciadas como técnico.
- **H-GR-31.** `ActivarBotonera` no incluye `ComandoVolverAlPerfilInicial` (`:154-160`): el botón queda en el estado que le dio la rama de `Form_Load` y no se revalida nunca. Al no haber `Form_Activate`, un menú reutilizado tras cambiar de usuario en otra ruta conservaría un estado obsoleto.
- **H-GR-32.** La ausencia de `Form_Activate` (`.form.txt:221` declara solo `OnLoad`) rompe la simetría con el menú de calidad (`.form.txt:221-222`). El menú que se abre desde el splash no recibe foco después de que su destino se cierre, salvo que lo reponga la navegación.
- **H-GR-33.** El salto de `EVE` depende de un argumento de texto mágico (`:235`): abierto desde «Entrar como» (`Form_Form0BDOpciones.cls:86`) el entorno **no** se reinicializa, y `m_ObjUsuarioConectadoInicialmente` conserva el usuario anterior (`:257`). La condición no distingue «abierto por el splash» de «abierto imitando a otro usuario».
- **H-GR-34.** `ComandoVolverAlPerfilInicial_Click` llama a `EVE` sin argumentos (`:201`) para reponer la identidad del usuario real de la sesión. Es una dependencia fuerte del orden de llamadas entre los dos menús, no verificable leyendo uno solo.
- **H-GR-35.** `ComandoTareasPendientes_Click` abre `FormTecnicoTareas` sin ninguna validación (`:96`), mientras su gemelo bloquea al técnico antes de abrir el suyo (`Form_Form0BDOpciones.cls:203-207`): dos handlers con el mismo nombre y propósito opuesto en el mismo menú. La corrección depende de que el destino sea correcto, no de una comprobación.
- **H-GR-36.** La divergencia entre menús es de **superficie, no de permiso**: aquí no existen `ComandoInformeIndicador`, `ComandoNuevaGestionDeRiesgo`, `ComandoBibliotecaRiesgos` ni `ComandoEntrarComo`, y el diseño no los declara. Un técnico no puede crear una gestión, consultar el indicador, ir a la biblioteca ni cambiar de usuario desde su menú.
- **H-GR-37.** El manejador de error de `Form_Load` termina la aplicación sin preguntar (`:286`), igual que en el menú de calidad (`Form_Form0BDOpciones.cls:468`): una inicialización fallida en el perfil técnico cierra la sesión completa.
- **H-GR-38.** Los dos menús comparten el patrón de error con `Err.Raise 1000` sin descripción (`:197, 203, 238, 298`) y con el global `pregunta` (`Variables Globales.bas:69`) usado como retorno de `MsgBox` sin declarar; también aquí un error real con número 1000 se mostraría como advertencia y no se enviaría al administrador.
- **H-GR-39.** Los rótulos de versión del diseño están desincronizados entre los dos menús (`2024-001` en `.form.txt:2194` frente a `2024-025` en `Form0BDOpciones.form.txt:3542`) y el rótulo del último riesgo usa distinta capitalización («Últimos riesgos», `.form.txt:2505`). `Form_Load` los sobrescribe, así que no afecta a ejecución, pero delata copias de diseño sin sincronizar.

## Propuestas web pendientes de decisión

Ninguna aprobada. Candidatas: resolver el perfil en el servidor y no pintar ni esconder opciones en cliente; sustituir el argumento de texto por un modo de sesión explícito (arranque normal frente a identidad supuesta), para que el reinicio del entorno no dependa de un argumento mágico; hacer reversible y con caducidad el cambio de identidad; unificar los dos menús en una sola pantalla construida a partir de las capacidades del perfil, en lugar de mantener dos formularios paralelos.

## Criterios de aceptación

Propuestos, no aprobados:

1. El menú técnico ofrece solo sus cuatro destinos y el retorno al perfil inicial solo cuando la sesión empezó como perfil no técnico.
2. Volver al perfil inicial repone la identidad real de la sesión y cierra los formularios abiertos del usuario impersonado.
3. Abrir el menú técnico como otro usuario no reinicializa el entorno más de lo necesario, y esa regla no depende de un texto libre.
4. Un fallo al cargar el menú técnico avisa y no termina la aplicación.
5. Ningún comando del menú técnico depende de una comprobación ausente en su propio handler.

## Pendientes

- Si el diseño declara o no la visibilidad por defecto de `ComandoVolverAlPerfilInicial`.
- Qué distingue en `EVE` un arranque normal de una identidad supuesta, y si la comparación de `OpenArgs` cubre todos los caminos de entrada.
- Si `m_ObjUsuarioConectadoInicialmente` sobrevive a `Application.Quit` o se reinicializa con la sesión.