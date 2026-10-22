# Form_FormInicial

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial.cls` (338 líneas) y `.form.txt` (1352 líneas) · **PRD:** [README](../README.md)

Pantalla de arranque de HPS. No muestra datos: encadena la carga pesada con un temporizador de 500 ms para que Access llegue a pintar la ventana (`Form_Load` fija `Me.TimerInterval = 500`, `:336`; `Form_Timer` lo anula y lanza el trabajo, `:12-21`), decide si la sesión será de solo lectura o de lectura/escritura según si el usuario está en la oficina, y solo entonces abre el contenedor raíz. También aloja el botón «Modo local…», que escribe en la tabla de configuración. De `Form_FormInicial00Principal` heredó la convención de `AvanceNuevo`/`AvanceCerrar`, que escriben en `lblEstado` (`FG.bas:4402-4404, 4421`). Las citas son `Form_FormInicial.cls:línea` salvo indicación; `FG.bas` abrevía `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Arranque de la aplicación | Es el primer formulario: `Form_Timer` arranca la carga sin ninguna entrada del usuario | `:12-21`, `.form.txt:29` |
| `FG.bas` (`AvanceNuevo`, `AvanceCerrar`) | Escriben el texto de progreso en `Forms("FormInicial").lblEstado`, así que el nombre del formulario y el de la etiqueta son un contrato | `FG.bas:4402-4404`, `FG.bas:4421` |
| `ComandoAccesoLectura_Click` | `DoCmd.OpenForm "FormInicial00Principal"` tras fijar el modo solo lectura | `:113-142`, apertura en `:132` |
| `ComandoAccesoEscritura_Click` | `DoCmd.OpenForm "FormInicial00Principal"` tras fijar el modo lectura/escritura | `:90-111`, apertura en `:101` |
| `ComandoSalir_Click` | `DoCmd.Quit acQuitSaveNone` | `:152` |

## Controles y eventos

`DefaultView =0` (`.form.txt:15`), `TimerInterval =500` en diseño (`.form.txt:29`), `Caption ="Usuario conectado: Administrador"` persistido en el diseño (`.form.txt:41`), `OnTimer` y `OnLoad` enganchados (`.form.txt:223-224`).

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Oculta la cinta de opciones, deshabilita los dos accesos y programa el temporizador | `:317-337`, `.form.txt:224` |
| (formulario) | — | `OnTimer` → `Form_Timer` | Anula el temporizador y ejecuta la carga pesada | `:12-21`, `.form.txt:223` |
| `lblEstado` | etiqueta | — | Barra de progreso de texto; la escriben `AvanceNuevo` y el propio `.cls` | `.form.txt:1139-1140`, `FG.bas:4404` |
| `ComandoAccesoLectura` | botón | `Click` → `ComandoAccesoLectura_Click` | Valida el permiso de entrada desde casa, fija solo lectura y abre el contenedor raíz | `:113-142`, `.form.txt:860-862` |
| `ComandoAccesoEscritura` | botón | `Click` → `ComandoAccesoEscritura_Click` | Fija lectura/escritura y abre el contenedor raíz | `:90-111`, `.form.txt:988-990` |
| `ComandoSalir` | botón | `Click` → `ComandoSalir_Click` | Pregunta y cierra la sesión | `:144-153`, `.form.txt:1063-1065` |
| `ComandoConfigurarModoLocal` | botón | `Click` → `ComandoConfigurarModoLocal_Click` | Alterna el `TempVar` de modo local o aplica un perfil de configuración | `:171-277`, `.form.txt:1166-1168` |
| `SituacionDatos`, `AccesoRedTelefonica` | cajas de texto | — | Etiquetas informativas que el `.cls` rellena con literales | `.form.txt:899, 960`, valores en `:64-65, 70-74` |
| `lblTitulo1`, `Etiqueta10`, `Etiqueta17` | etiquetas | — | Rótulos de la pantalla | `.form.txt:599-600, 927-928, 1024-1025` |

## Flujos

### Carga inicial (`Form_Timer` → `EjecutarCargaInicial`, `:12-21`, `:23-89`)

1. `Form_Timer` pone `Me.TimerInterval = 0`, llama a `DoEvents` para forzar el repintado y arranca `EjecutarCargaInicial` (`:14-20`).
2. Muestra `lblEstado` y avisa «Cargando configuración del entorno...» con `AvanceNuevo` (`:27, 30`).
3. `EVE` (`:34`; `VariablesEntorno.bas:301`): resuelve la identidad del usuario, deja `Application.TempVars("DatosEnLocal") = "No"` (`VariablesEntorno.bas:330`) y carga `m_ObjEntorno`.
4. Avisa «Sincronizando tablas locales...» y llama a `EjecutarSincronizacionCargaInicialFrontend p_Error:=sError` (`:42-43`; `Mod_StartupCacheInitialization.bas:18`), que delega en `SincronizarCachesLocalesFrontend` (`Mod_StartupCacheInitialization.bas:25`, `:62`).
5. Si hay error, muestra «Error en la carga» y deja `lblEstado.Caption = "Error."` (`:45-47`); si no, avisa «Carga completada.» (`:50`).
6. Exige `m_EnOficina` resuelto (`:53-56`) y, si no se está en la oficina, que `m_AccesoADatosTE` también lo esté (`:57-62`); los dos cortes lanzan `Err.Raise 1000` con el mensaje en `sError`.
7. En oficina deja `SituacionDatos = "En Oficina"`, `AccesoRedTelefonica = "Sí"` y habilita escritura (`:63-68`); en local deja `"En Local"`, decide `AccesoRedTelefonica` según `m_AccesoADatosTE`, deshabilita escritura y deja el foco en el acceso de solo lectura (`:69-82`).
8. En cualquier salida esconde `lblEstado` (`:84`); el manejador de errores solo muestra «Error crítico» sin correo ni limpieza (`:87-88`).

### Acceso con escritura (`ComandoAccesoEscritura_Click`, `:90-111`)

1. Hourglass y `m_Error = ""` (`:92-93`).
2. `CerrarFormulariosAbiertos` (`:94`; `FG.bas:2411`): cierra todos los formularios cargados salvo el activo, sin guardar.
3. `EstablecerSoloLectura p_SoloLectura:=EnumSiNo.No, p_Error:=m_Error` (`:95`; `FG.bas:29`), que rehace por su cuenta la comprobación de oficina (`FG.bas:37-40`), escribe `Application.TempVars("SoloLectura") = "No"` y `m_EstadoConexion = "Lectura/Escritura"` (`FG.bas:43-44`).
4. `AvanceCerrar` (`:100`; `FG.bas:4415`), `DoCmd.OpenForm "FormInicial00Principal"` (`:101`) y fin del hourglass (`:102`).
5. En error: correo al administrador solo si `Err.Number <> 1000` y `MsgBox` crítico (`:105-110`).

### Acceso en solo lectura (`ComandoAccesoLectura_Click`, `:113-142`)

Repite el mismo camino, pero antes comprueba que, si `m_EnOficina = EnumSiNo.No`, el usuario tenga `PermisoEntradaLocalEnCasa = EnumSiNo.Sí` (`:118-123`; `Usuario.cls:51`); si no, avisa y no abre nada. La llamada es `EstablecerSoloLectura p_SoloLectura:=EnumSiNo.Sí` (`:125`), que repite esa misma comprobación (`:47-52` en `FG.bas`) y escribe `Application.TempVars("SoloLectura") = "Sí"` y `m_EstadoConexion = "Solo lectura"` (`FG.bas:55-56`). Abre el contenedor raíz en `:132`.

### Salir (`ComandoSalir_Click`, `:144-153`)

`On Error Resume Next` (`:146`), `MsgBox` de confirmación con «Sí» por omisión en el segundo botón (`:147`); cualquier respuesta distinta de `vbYes` cancela (`:148-150`) y solo entonces `DoCmd.Quit acQuitSaveNone` (`:152`).

### Modo local (`ComandoConfigurarModoLocal_Click`, `:171-277`)

1. Lee `Application.TempVars(k_LocalModeTempVar)`, con `k_LocalModeTempVar = "DatosEnLocal"` declarado en `:6-8`, y normaliza el valor a «(sin inicializar)» si no es «Sí» ni «No» (`:183-186`).
2. Un `MsgBox` de tres botones ofrece dos acciones (`:189-197`): **Sí**, alternar el `TempVar`; **No**, aplicar un perfil. Cancelar sale sin tocar nada (`:195-197`).
3. Rama de alternancia: calcula el nuevo estado (`:200-205`), pide opcionalmente la ruta del backend con `InputBox` solo al activar (`:209-214`), escribe el `TempVar` (`:217`) y confirma que el cambio se aplicará al reiniciar (`:220-224`).
4. Rama de perfil: pregunta Local o Producción (`:229-237`) y llama a `ConfigurarTablaConfiguracion_PerfilLocal` o `…PerfilProduccion` (`:240-244`; `modConfiguracionHPS.bas:325` y `modConfiguracionHPS.bas:364`), que crean `TbConfiguracionHPS` y `TbVinculosTablas` si faltan y escriben las claves de entorno (`modConfiguracionHPS.bas:30-31, 89-100`).
5. Refresca `m_ObjEntorno.CargarConfiguracion` (`:247`; `Entorno.cls:3199`), lee las dos rutas de backend con `GetConfig` (`:252-253`; `Entorno.cls:3269`), calcula el directorio común con `CommonParentDir` (`:254`; `:281-291`) y, si no lo encuentra, con `ParentDirOfPath` (`:256`; `:295-313`).
6. `m_ObjEntorno.RelinkAllTables(sParentDir)` (`:261`; `Entorno.cls:3431`) y un `MsgBox` final informa del resultado (`:264-269`).

## Reglas, validaciones y cálculos

- **El temporizador es la única barrera de repintado.** `Form_Load` programa 500 ms (`:336`) y `Form_Timer` los anula antes de trabajar (`:14`); sin ese margen, `EVE` bloquearía la interfaz antes del primer repintado.
- **La ventana se prepara antes que los datos.** `Form_Load` solo deshabilita botones y pone el rótulo inicial (`:328-332`); ninguna llamada a objetos globales pesados ocurre en el `Load`.
- **Los dos accesos convergen en la misma transición.** `EstablecerSoloLectura` es la única función que fija el modo de sesión, y escribe en dos estados compartidos: `Application.TempVars("SoloLectura")` y `m_EstadoConexion` (`FG.bas:43-44, 55-56`).
- **La habilitación de escritura depende de la topología, no de un permiso.** `m_EnOficina` y `m_AccesoADatosTE` los fija `EVE` (`VariablesEntorno.bas:333`); si quedan sin resolver, la carga se corta con `Err.Raise 1000` (`:53-62`).
- **`CerrarFormulariosAbiertos` respeta el formulario activo** (`FG.bas:2415-2429`): en estos dos caminos el activo es esta pantalla, que sobrevive al `DoCmd.OpenForm` del contenedor raíz (`:101, 132`).
- **Las constantes del modo local van antes de cualquier miembro** (`k_LocalMode*`, `:6-8`), con el motivo anotado en el propio código (`:5`).
- **Los valores de `SituacionDatos` y `AccesoRedTelefonica` son literales** (`"En Oficina"`, `"En Local"`, `"Sí"`, `"No"`; `:64-65, 70-74`); no hay tabla de configuración que los gobierne.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Entorno e identidad del usuario conectado | `TbConfiguracionHPS` y tablas de usuario (vía `EVE`) | `SELECT` + caché en memoria | `:34`, `VariablesEntorno.bas:301-333` |
| Sincronización de arranque | múltiples tablas `Tb…Local` | `DELETE` + `INSERT` masivos | `:43`, `Mod_StartupCacheInitialization.bas:62` |
| Modo de sesión | `Application.TempVars("SoloLectura")`, `m_EstadoConexion` | estado en memoria, sin tabla | `FG.bas:43-44, 55-56` |
| Modo local | `Application.TempVars("DatosEnLocal")` | estado en memoria | `:183, 217`; lectura en `VariablesEntorno.bas:365` |
| Perfil de configuración | `TbConfiguracionHPS`, `TbVinculosTablas` | `INSERT`/`UPDATE` por clave | `:241, 243`, `modConfiguracionHPS.bas:89-100, 120-130` |
| Rutas de backend | `TbConfiguracionHPS` (vía `GetConfig`, `GetConfigTablePath`) | `SELECT` | `:252-253, 266`, `Entorno.cls:3269, 3318` |
| Relink de `TableDefs` | enlaces de la base actual | `TableDef.Connect` reasignado | `:261`, `Entorno.cls:3431` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `AvanceNuevo` y `AvanceCerrar` localizan esta pantalla por nombre (`FG.bas:4402-4404, 4421`), así que el nombre `FormInicial` y el control `lblEstado` son un contrato entre tres módulos; la sesión se transporta en `Application.TempVars("SoloLectura")` y `Application.TempVars("DatosEnLocal")` (`FG.bas:43, 55`; `VariablesEntorno.bas:330, 365`), estado compartido sin equivalente directo en web; `m_EstadoConexion`, `m_EnOficina`, `m_AccesoADatosTE` y `m_ObjUsuarioConectado` son variables de módulo públicas (`VariablesEntorno.bas:269-270, 267`); `CerrarFormulariosAbiertos` recorre `AllForms` y cierra por nombre (`FG.bas:2422-2431`), lo que en web es una navegación y no un barrido; y `DoCmd.ShowToolbar "Ribbon", acToolbarNo` (`:324`) y `DoCmd.Quit` (`:152`) no tienen representación.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El mensaje de error de `ComandoAccesoEscritura_Click` dice «Al ComandoAccesoLectura_Click…» (`:106`), igual que el de su gemelo (`:137`): es una copia sin corregir.
- El toggle del modo local se confirma con «El cambio se aplicará al reiniciar la aplicación» (`:223`), pero `EVE` —que es quien lee ese `TempVar` para resolver las rutas (`VariablesEntorno.bas:365`)— lo sobrescribe con «No» en su primera línea (`VariablesEntorno.bas:330`). Tal como está, el valor alternado no llega a ningún punto de decisión posterior; hay que comprobar si algún punto posterior a `EVE` lo vuelve a leer antes de concluir que es inerte.
- La comprobación de `PermisoEntradaLocalEnCasa` está duplicada: la hace este `.cls` (`:118-123`) y también `EstablecerSoloLectura` (`FG.bas:47-52`), con mensajes de error distintos para la misma condición.
- **`pregunta` es un global compartido, no una variable local.** Se usa en cuatro puntos (`:110, 141, 147-148, 276`) sin estar declarada en este `.cls`, pero resuelve al `Public pregunta As Long` de `VariablesEntorno.bas:220`: compila sin problema y el valor del `MsgBox` queda en un único hueco global que todos los formularios de la aplicación se pisan. `lbl` en `Form_FormInicial00Principal.cls:21` es un caso distinto: allí la variable sí es local al formulario y no está declarada en su `.cls`.
- `sBackendPath` se pide y se muestra (`:210-214, 222`) pero nunca se persiste ni se usa: solo se escribe el `TempVar` (`:217`). El backend sandbox se resuelve por la configuración, no por la ruta escrita.
- Los dos perfiles escriben las claves de entorno incluyendo una contraseña en claro (`modConfiguracionHPS.bas:352` y la línea equivalente del perfil de producción); el valor no se reproduce aquí. Está fuera de la superficie de esta ficha, pero el botón «Modo local…» es quien lo invoca.
- Si `EjecutarSincronizacionCargaInicialFrontend` falla, `lblEstado.Caption` queda en «Error.» pero **los dos accesos siguen deshabilitados** (`:328-329`, sin volver a habilitarlos en la rama de error `:45-47`): la pantalla queda inservible sin más salida que «Salir».
- El manejador `errores:` de `EjecutarCargaInicial` (`:87-88`) muestra «Error crítico» sin correo al administrador, sin `AvanceCerrar` y sin restaurar el hourglass, a diferencia del resto de handlers del archivo (`:105-110, 136-141, 272-276`).
- `Me.lblEstado.Visible = True` se repite en `:27` y `:39` sin motivo aparente, y `AvanceCerrar` (`:100, 131`) ya lo oculta antes de abrir el contenedor raíz; `:84` lo oculta otra vez.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: sustituir el splash con un endpoint de arranque que devuelva `{entorno, modoSesion, errores}` y una pantalla de espera con progreso real, en lugar de un temporizador de 500 ms; convertir el modo local en un perfil de configuración versionado en servidor, sin escritura de rutas desde el cliente; y mover la persistencia de la contraseña a un gestor de secretos en lugar de una clave en la tabla de configuración.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al arrancar, la pantalla se pinta con un mensaje de progreso antes de empezar la carga pesada y el temporizador no vuelve a dispararse.
2. Si el entorno no se puede determinar, la carga se detiene con un mensaje que dice si falta la detección de oficina o la de acceso a datos.
3. En oficina se habilita el acceso de lectura/escritura y en local solo el de solo lectura, con el estado de sesión publicado para el resto de formularios.
4. El acceso en solo lectura fuera de la oficina se rechaza, sin abrir el contenedor raíz, cuando el usuario no tiene permiso de entrada desde casa.
5. Ambos accesos abren `FormInicial00Principal` tras cerrar los formularios abiertos salvo el activo.
6. «Salir» pide confirmación y solo cierra la sesión si la respuesta es afirmativa.
7. El botón «Modo local…» ofrece las acciones de alternar el modo y de aplicar un perfil, y cancela sin tocar nada; el perfil aplicado informa del directorio de backends y del resultado del relink.
8. Ante un error de carga, el usuario ve un mensaje con el detalle del fallo y puede salir de la aplicación.

## Pendientes

- Confirmar contra el binario qué componente abre esta pantalla como primer formulario: no hay `DoCmd.OpenForm "FormInicial"` en `src/`.
- Confirmar si el `TempVar("DatosEnLocal")` alternado desde `:217` lo lee algún punto posterior a `EVE`, dado que `VariablesEntorno.bas:330` lo sobrescribe en cada arranque.
- Confirmar el destino de `sBackendPath` (`:210-214`), que hoy se pide y se muestra pero no se persiste.
- Verificar si `Form_FormInicial` es realmente el padre declarado de `Form_FormInicial00Principal` en el binario, o si la relación es solo de código heredado.