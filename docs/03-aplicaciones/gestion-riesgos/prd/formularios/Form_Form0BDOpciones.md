# Form_Form0BDOpciones

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_Form0BDOpciones.cls` (499 líneas, 16 procedimientos) y `src/forms/Form0BDOpciones.form.txt` · **PRD:** [README](../README.md)

Menú principal del perfil de calidad y administración. Es un panel de botones: cada `ComandoXxx_Click` abre un formulario concreto y ninguno consulta datos de negocio por su cuenta. La única decisión de permisos que toma es `EsTecnico`, y solo para rechazar tres de sus botones. No escribe nada en base de datos, salvo el estado de sesión que delega en `EVE`. Las citas son `Form_Form0BDOpciones.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `frmSplash` | `DoCmd.OpenForm` cuando el usuario no es técnico | `Form_frmSplash.cls:83` |
| `Form0BDOpcionesTecnico` | «Volver al perfil inicial» | `Form_Form0BDOpcionesTecnico.cls:205` |

No acepta parámetros: no lee `OpenArgs` y no usa `Filter` ni `Where`. Lo que sí consume es la TempVar de modelo de árbol, que puede cambiar desde aquí (`:23-44`).

Salidas: `Form0BDOpcionesTecnico` (`:86`), `FormIndicador` (`:113`), `FormProyectosGestion` (`:177`), `FormCalidadTareas` (`:207`), `FormRiesgosBibliotecaGestion` (`:236`), `FormGestionRiesgos` (`:294`), `FormRiesgosGestion` (`:345`), `FormControlCambiosGestion` (`:385`), la ayuda (`:141`) y la aplicación completa (`:56`).

## Controles y eventos

| Control | Evento y handler | Qué abre | Evidencia |
|---|---|---|---|
| (formulario) | `Form_Load` | Prepara rótulos, lista de usuarios y botonera | `:425-470`, `.form.txt:222` |
| (formulario) | `Form_Activate` | `AvanceCerrar` | `:421-423`, `.form.txt:221` |
| `ComandoTareasPendientes` | `ComandoTareasPendientes_Click` | `FormCalidadTareas`, bloqueado si es técnico | `:196-222`, `.form.txt:1742` |
| `ComandoGestionesDeRiesgo` | `ComandoGestionesDeRiesgo_Click` | `FormProyectosGestion` tras refrescar objetos activos | `:160-192`, `.form.txt:1994` |
| `ComandoNuevaGestionDeRiesgo` | `ComandoNuevaGestionDeRiesgo_Click` | `FormGestionRiesgos` en modo alta, bloqueado si es técnico | `:258-315`, `.form.txt:2307` |
| `ComandoBibliotecaRiesgos` | `ComandoBibliotecaRiesgos_Click` | `FormRiesgosBibliotecaGestion` y oculta un botón suyo | `:223-256`, `.form.txt:943` |
| `ComandoInformeIndicador` | `ComandoInformeIndicador_Click` | `FormIndicador`, bloqueado si es técnico | `:102-127`, `.form.txt:2648` |
| `ComandoUltimoProyecto` | `ComandoUltimoProyecto_Click` | `FormRiesgosGestion`, o la lista si no hay último proyecto | `:318-359`, `.form.txt:2819` |
| `ComandoVerControlCambios` | `ComandoVerControlCambios_Click` | `FormControlCambiosGestion` | `:377-401`, `.form.txt:3226` |
| `ComandoAyuda` | `ComandoAyuda_Click` | Ayuda de la aplicación | `:129-158`, `.form.txt:852` |
| `ComandoEntrarComo` | `ComandoEntrarComo_Click` | `Form0BDOpcionesTecnico` como otro usuario | `:60-100`, `.form.txt:4157` |
| `cmdSalir` | `CmdSalir_Click` | Cierra la aplicación | `:48-58`, `.form.txt:4231` |
| `lblVideo` | `lblVideo_Click` | Navegador externo al vídeo configurado | `:473-499`, `.form.txt:3950` |
| `EntrarComo` | `EntrarComo_KeyDown` | Enter actúa como clic del botón | `:404-419`, `.form.txt:4026` |

Los rótulos de texto son etiquetas fijas (`lblMisProyectos`, `lblNuevoProyecto`, `lblBibliotecaRiesgos`, `lblInformeIndicador`, `lblControlCambios`), sin evento asociado: en el diseño, cada botón lleva encima su etiqueta descriptiva (`.form.txt:3572-3766`). `lblTitulo1` (`:439`) y `lblEstado` (`:452`) se ocultan o vacían en ejecución.

## Flujos

### Apertura (`Form_Load`, `:425-470`)

1. Reloj de arena y `AjustarTamaño Me` (`:436`; `Funciones Generales.bas:1043`).
2. `Me.Caption` pasa a título del usuario conectado y `lblVersion` a la versión de la aplicación (`:438, 448`).
3. `Avance "Obteniendo las tareas"` y `EntrarComo.RowSource = m_ObjEntorno.ListaUsuarios` (`:441-442`).
4. `lblUltimoProyecto` muestra el nombre del último proyecto o el texto genérico (`:443-447`).
5. `ActivarBotonera` habilita siete controles y `AvanceCerrar` cierra el rótulo de estado (`:450-452`).

Su manejador de error **termina la aplicación** (`:469`).

### Normal (patrón común de los once handlers)

`DoEvents` → `Hourglass True` → `m_Error = ""` → validación → acción → `Hourglass False`. Los errores previsibles se señalan con `Err.Raise 1000` y el manejador los muestra como advertencia sin correo; los inesperados se describen, se envían al administrador por `CorreoAlAdministrador` y se muestran como error crítico (`:92-99`).

### Alternativo

- **Sin último proyecto:** `ComandoUltimoProyecto_Click` delega en `ComandoGestionesDeRiesgo_Click` (`:330-334`).
- **Biblioteca:** si el formulario ya estaba abierto, lo cierra sin guardar y lo reabre; después oculta un botón suyo desde fuera (`Form_FormRiesgosBibliotecaGestion.ComandoNoExisteRiesgo.Visible = False`, `:237-240`).
- **Alta de gestión:** cierra `FormGestionRiesgos` si estaba abierto y reabre en modo alta (`:291-296`).

### De error

`Err.Raise 1000` no lleva descripción (`:72` y diez ocurrencias más): el texto mostrado es siempre el `m_Error` asignado justo antes. `CmdSalir_Click` y `ActivarBotonera` usan `On Error Resume Next` y no tienen manejador (`:49`, `:361`).

## Reglas, validaciones y cálculos

- **El perfil decide dentro del handler, no en la botonera.** Tres comandos rechazan al técnico con `EsTecnico = EnumSiNo.Sí`: tareas pendientes (`:203`), nueva gestión (`:267`) e informe indicador (`:109`). Los demás no miran el rol.
- **`ActivarBotonera` (`:360-372`) no deshabilita nada:** solo pone `Enabled = True` a siete controles. El estado por defecto del diseño es habilitado, de modo que un técnico ve tres botones que le van a rechazar.
- **Cambio de modelo de árbol:** `EstablecerModeloCadenaJerarquica` (`:23-44`) valida `antiguo` o `nuevo` y escribe la TempVar que el resto de la aplicación lee. Es la única escritura de este formulario.
- **Cambio de usuario:** `Constructor.getUsuario` con el valor del combo (`:70`), cierre de formularios abiertos, `EVE` con el correo del usuario elegido y apertura del menú técnico con el argumento «Para Otro Usuario» (`:76-86`).

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Identidad, tareas, cachés | Globales de sesión `m_ObjEntorno`, `m_ObjUltimoProyecto` | `Variables Globales.bas:73, 130` |
| Lista de usuarios | `m_ObjEntorno.ListaUsuarios` | `:442` |
| Modelo de árbol | TempVar `CadenaJerarquicaModelo` | `:33-38` |
| Rótulos de avance | `Avance` y `AvanceCerrar` | `Funciones Generales.bas:4879, 4920` |
| Apertura externa | `ShellExecuteA` vía `Ejecutar` | `Variables Globales.bas:15`, `:485` |

Sin escrituras en base de datos.

## Comportamiento actual

Panel de navegación con once destinos, un selector de usuario, un enlace al vídeo de ayuda y salida de la aplicación. La comprobación de perfil ocurre en los handlers, la carga de datos en `EVE` y el refresco en `ActualizarObjetosActivos`. La versión técnica del menú es `Form_Form0BDOpcionesTecnico`.

## Posibles defectos (sin corregir, a verificar)

- **H-GR-192.** `ActivarBotonera` (`:360-372`) habilita indiscriminadamente bajo `On Error Resume Next` y nunca deshabilita: el único filtro real de rol son tres `If` dentro de handlers (`:109, 203, 267`), de modo que la botonera **muestra** opciones que el propio formulario **rechaza**. Conviene verificar si el diseño deja algún control deshabilitado por defecto.
- **H-GR-193.** `Me.ComandoVerControlCambios.Enabled = True` aparece dos veces seguidas (`:369-370`): la segunda es inerte y apunta a copia de código.
- **H-GR-194.** La misma routine no toca `ComandoEntrarComo`, `cmdSalir` ni `ComandoAyuda` (`:360-372`): la «botonera activada» no es la botonera completa, y esos tres quedan en el estado que traiga el diseño.
- **H-GR-195.** `ComandoNuevaGestionDeRiesgo_Click` repite el bloque `On Error GoTo errores` / `DoEvents` / `Hourglass True` / `m_Error = ""` ya ejecutado (`:276-282`). El segundo bloque **descarta** el `m_Error` anterior y reinicia el reloj de arena sin cerrar el anterior.
- **H-GR-196.** `Form_Load` asigna `Me.lblTitulo1.Caption = m_Titulo` (`:439`) con una variable declarada y **nunca asignada** (`:430`): el rótulo queda siempre vacío.
- **H-GR-197.** El manejador de error de `Form_Load` acaba en `DoCmd.Quit acQuitSaveNone` (`:469`): un fallo de carga —por ejemplo, `m_ObjEntorno` sin inicializar— cierra la aplicación sin preguntar y descarta el trabajo de sesión.
- **H-GR-198.** `CmdSalir_Click` (`:48-58`) muestra el aviso bajo `On Error Resume Next` (`:50`) y con un texto truncado («aplicaciónº», `:51`), y la llamada a `CerrarAplicacion` está comentada (`:55`); `Application.Quit acQuitSaveNone` (`:56`) descarta cambios sin guardar.
- **H-GR-199.** Los `Err.Raise 1000` (`:72, 76, 80, 84, 111, 142, 170, 205, 274, 338, 483`) no llevan `vbObjectError`, y el manejador decide entre aviso y crítico comparando `Err.Number <> 1000`. Un error real de Access con número 1000 se mostraría como advertencia y **no se enviaría al administrador**.
- **H-GR-200.** El selector de usuario carga la lista completa sin filtrar por rol (`:442`) y `ComandoEntrarComo_Click` no comprueba `EsTecnico` en ningún punto: el único control del salto de perfil está en `Form_Load` del menú técnico (`Form_Form0BDOpcionesTecnico.cls:259`). El acceso está en el destino, no en el origen.
- **H-GR-201.** El diseño congela rótulos capturados en una ejecución real: `lblUltimoProyecto` guarda un nombre de proyecto concreto (`.form.txt:3733`) y `lblVersion` una versión concreta (`.form.txt:3542`). No es un fallo en ejecución (`Form_Load` los sobrescribe) pero documenta que el diseño se ha extraído con datos cargados.
- **H-GR-202.** `lblVideo_Click` comprueba existencia con `fso.FileExists` sobre la ruta configurada (`:481`) y luego delega en `ShellExecuteA` (`:485`, `Variables Globales.bas:15`): la comprobación y la apertura son dos actos separados y la disponibilidad real depende del momento del clic.

## Propuestas web pendientes de decisión

Ninguna aprobada. Candidatas: navegación por rutas en vez de formularios modales; decidir el perfil en el servidor y devolver solo las opciones permitidas, sin rechazos en cliente; sustituir el combo «Entrar como» por un mecanismo de impersonación con caducidad propia; incrustar el vídeo explicativo en lugar de abrir un visor externo; eliminar el concepto de modelo de árbol conmutable y dejar un único modelo.

## Criterios de aceptación

Propuestos, no aprobados:

1. El menú muestra solo las acciones permitidas al perfil del usuario, sin botones que rechacen al pulsarlos.
2. Cada acción abre su destino y el anterior se cierra sin dejar sesiones abiertas.
3. Cambiar de usuario deja el entorno y los datos del usuario anterior sin rastro en la sesión.
4. Un fallo al cargar el menú avisa y no termina la aplicación.
5. Salir avisa y no descarta trabajo sin guardar.

## Pendientes

- Cómo se abre realmente el formulario en el arranque (propiedad de inicio o macro, fuera del código fuente) y qué papel cumple `EstablecerModeloCadenaJerarquica` desde el menú.
- Si el diseño deja algún control deshabilitado por defecto, para saber si `ActivarBotonera` es realmente la única fuente de habilitación.
- Qué decide `ActualizarObjetosActivos` y cuánto puede tardar, porque bloquea la navegación con reloj de arena.