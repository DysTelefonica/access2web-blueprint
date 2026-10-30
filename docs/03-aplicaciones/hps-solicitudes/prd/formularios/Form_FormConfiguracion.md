# Form_FormConfiguracion

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormConfiguracion.cls` (418 líneas) y `.form.txt` (50 controles `Begin`) · **PRD:** [README](../README.md)

Pantalla de configuración de la aplicación: edita los parámetros de `Configuracion` (correos automáticos, buzón de seguridad, recordatorios, autogestión MARGA/pre-MARGA y versiones de plantillas HTML/Excel) y abre los directorios de plantillas. Es un formulario **no enlazado** (sin `RecordSource`): pinta los controles desde una instancia de la clase `Configuracion` que se carga por `constructor.getConfiguracion`. Emite el evento público `Modificado` para que `Form_Form0BDOpciones` refresque `lblParametros` y `lblConfCorreosAutomaticos`. Las citas son `Form_FormConfiguracion.cls:línea` salvo indicación; `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; `classes/` = `src/classes/`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form0BDOpciones` | `DoCmd.OpenForm "FormConfiguracion"` desde `ComandoConfiguracion_Click`; el formulario es `WithEvents m_FormConf` para reaccionar al `Modificado` | `Form_Form0BDOpciones.cls:135-137`, este cls `:5` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `Form0BDOpciones` | `RaiseEvent Modificado` tras guardar; `DoCmd.Close acForm, Me.Name, acSaveNo` | `:78, 81` |
| Explorador de Windows | `Ejecutar(Me.hwnd, "open", m_ObjEntorno.URLDirectorioPlantillasHTMLVersionEstablecida, ...)` | `:202, 240` |

Estado compartido: globales `m_TituloFormulario`, `m_ObjEntorno`, `TempVars("ActivadoCorreoAutomatico")`, `m_FaltaAlgunParametroEnConfiguracion` (declaradas fuera de este `cls`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` → `Form_Load` | `Ajustar Me`, `EstablecerDatos m_Error`, oculta reloj | `:266-291` |
| `cmdSalir` | botón | `OnClick` → `cmdSalir_Click` | Cierra sin guardar (`acSaveNo`) | `:13-15` |
| `ComandoAyuda` | botón | `OnClick` → `ComandoAyuda_Click` | Llama a `AbrirAyuda` (sin marcadores de error) | `:17-49` |
| `ComandoRegistrar` | botón | `OnClick` → `ComandoRegistrar_Click` | Llama a `HaHabidoCambios`, `RellenarObjeto`, `ConfiguracionOperaciones.Registrar`; emite `Modificado` y cierra | `:51-90` |
| `ComandoURLPlantillasHTML` | botón | `OnClick` → `ComandoURLPlantillasHTML_Click` | Llama a `Ejecutar Me.hwnd, "open", m_ObjEntorno.URLDirectorioPlantillasHTMLVersionEstablecida` | `:228-249` |
| `ComandoURLPlantillasExcel` | botón | `OnClick` → `ComandoURLPlantillasExcel_Click` | Igual sobre `URLDirectorioPlantillasEXCELConVersion` | `:198-225` |
| `CorreosAutomaticos` | combo Sí/No | `AfterUpdate` → `CorreosAutomaticos_AfterUpdate` | Activa o desactiva los campos con `Tag = "CORREO"` mediante `ActivarCamposCorreo` | `:254-281` |
| `DiasParaRecordatorioExcel1`, `DiasParaRecordatorioExcel2` | texto | — | Días entre recordatorios del Excel (cada uno con `Tag = "CORREO"`) | `.form.txt:751, 812` |
| `DiasCancelacionPreMARGA`, `DiasCancelacionMARGA` | texto | — | Días para cancelación automática (sí/no en combo) | `.form.txt:873, 994-1057` |
| `DiasParaRecordatorioRellenoMarga1`, `…RellenoMarga2` | texto | — | Recordatorios del proceso MARGA | `.form.txt:933, 1055` |
| `BuzonSeguridad` | texto | — | Buzón del director de seguridad (Tag=`CORREO`) | `.form.txt:994` |
| `EmailDirectorSeguridad` | texto | — | Correo del director de seguridad (Tag=`CORREO`) | `.form.txt:1055` |
| `CorreodeEnvio` | texto | — | Correo de envío (Tag=`CORREO`, no asignado en `EstablecerDatos`) | `.form.txt` (en `src/modules/FU.bas` figura como clave INI; ver §Pendientes) |
| `VersionPlantillasHTML`, `VersionPlantillasEXCEL` | texto | — | Versión activa de las plantillas (no se validan) | `:428, 446` |

## Flujos

### Carga (`Form_Load`, `:266-291`)

1. `On Error GoTo errores`; `DoCmd.Hourglass True`; `Ajustar Me` (`:276-279`).
2. `EstablecerDatos m_Error`; si devuelve error, `Err.Raise 1000` que dispara la rama `errores:` (`:282-285`).
3. La rama `errores:` envía el error al administrador con `CorreoAlAdministrador`, muestra `MsgBox` con `Err.Description` y cierra el formulario (`:328-336`).

### Pintado (`EstablecerDatos`, `:294-359`)

1. `Me.Caption = m_TituloFormulario` (`:301`).
2. `constructor.getConfiguracion p_Error` devuelve `m_ConfAlInicio` (`classes/ConfiguracionOperaciones.cls`); si es `Nothing` (registro único de `TbConfiguracion` aún no creado), asume `AutocancelacionMARGA = "Sí"` y `AutocancelacionPreMARGA = "Sí"` y desactiva `ActivarCamposCorreo False` (`:307-318`).
3. Con configuración existente, asigna uno a uno los campos control (`:319-352`) si el valor original no está vacío. Los campos `Tag = "CORREO"` se activan o desactivan según `m_ConfAlInicio.CorreosAutomaticos` (`:323-329`).
4. `TempVars("ActivadoCorreoAutomatico") = "No"` → muestra `lblConfCorreosAutomaticos` con texto «Aún no implementado correos automáticos» (`:354-358`).

### Activación de campos de correo (`ActivarCamposCorreo`, `:94-114`)

Itera `Me.Controls`, y para cada `ctl.Tag = "CORREO"` fija `ctl.Enabled = p_Activar` (`:106-110`). Es un patrón declarativo: añadir un campo nuevo de correo al formulario basta con ponerle `Tag = "CORREO"` en el `.form.txt`.

### Detección de cambios (`HaHabidoCambios`, `:382-461`)

Compara cada campo cargado contra su valor inicial en `m_ConfAlInicio`. Si encuentra una diferencia, devuelve `True` y sale inmediatamente (`:388-457`); si no hay diferencias, el flujo cae al `Exit Function` final y `HaHabidoCambios` queda en su valor por defecto `False`. Si `m_ConfAlInicio Is Nothing`, devuelve `True` directamente (`:386-389`) — un registro inicial forzado.

### Guardado (`ComandoRegistrar_Click`, `:51-90`)

1. `HaHabidoCambios m_Error`; si no hay cambios, aborta con `m_Error = "No hay cambios que guardar"` (`:62-65`).
2. `RellenarObjeto m_Error` (`:67-69`) copia cada `Nz(Me.<campo>, "")` a `m_Conf.<campo>`; mantiene `m_Conf.id = m_ConfAlInicio.id` cuando existe (`:474-476`); si no existe, el nuevo ID se asignará al guardar.
3. `Set m_ObjEntorno.Configuracion = Nothing` (`:72`) invalida la caché del entorno antes de guardar — sin esto, `m_ObjEntorno.getConfiguracion` devolvería la versión previa.
4. `ConfiguracionOperaciones.Registrar m_Conf, m_Error` (`classes/ConfiguracionOperaciones.cls:1-402`): escribe la fila en `TbConfiguracion` (`schema.sql:tb-configuracion`).
5. `RaiseEvent Modificado`; `DoCmd.Close acForm, Me.Name, acSaveNo` (`:78-81`).

### Rellenar objeto (`RellenarObjeto`, `:467-501`)

Set `m_Conf = New Configuracion` y copia los 14 campos desde los controles. Sin validación: cada `Nz(Me.<x>, "")` se asigna tal cual; la validación de tipos se hará en `ConfiguracionOperaciones.Registrar` (a verificar).

### Apertura de plantillas (`ComandoURLPlantillas…_Click`)

Valida `fso.FolderExists(m_ObjEntorno.URLDirectorioPlantillas<X>ConVersion)` (`:213, 234`); si el directorio no existe, aborta con error; en caso contrario invoca `Ejecutar(Me.hwnd, "open", <URL>, "", "", 1)` (`modules/FU.bas`) para abrirlo en el explorador.

## Reglas, validaciones y cálculos

- **Snapshot vs draft:** `m_ConfAlInicio` se carga una sola vez en `EstablecerDatos`; `m_Conf` es el objeto borrador que se construye en `RellenarObjeto`. Si el usuario edita y cancela con `cmdSalir`, `m_Conf` se descarta al cerrar el formulario.
- **Detección de cambios:** 14 comparaciones manuales con `Nz` (`:386-457`); el patrón se rompe si se añade un campo nuevo sin agregar el `If … <> Nz(…) Then` correspondiente — defecto silencioso.
- **Cache del entorno:** `Set m_ObjEntorno.Configuracion = Nothing` se ejecuta al guardar pero NO al descartar cambios con `cmdSalir`: un usuario que abre el formulario, no guarda y vuelve al menú deja `m_ObjEntorno.Configuracion` con la versión anterior en memoria (a verificar que `m_ObjEntorno.getConfiguracion` la recarga; ver `classes/Entorno.cls`).
- **Comprobación de Tag `CORREO`:** `InStr(1, Nz(ctl.Tag, ""), "CORREO") <> 0` también activa controles cuyo nombre de clase sea `CORREO*` o contenga la cadena en otra propiedad del `.form.txt` distinta del `Tag` (no aplica a VBA, pero la coincidencia podría propagarse si se reusa `Tag`).
- **Sin validación de versiones.** `VersionPlantillasHTML`/`VersionPlantillasEXCEL` se leen/escriben tal cual; la comparación contra la versión activa no se hace en este formulario.
- **`fso` (FileSystemObject) global:** `ComandoURLPlantillas*_Click` lo usa sin instanciarlo localmente (`fso.FolderExists`); se asume `fso` declarado en `VG.bas` o `FU.bas` (a verificar).
- **Errores de redacción en handlers de error.** Cada `errores:` tiene un mensaje con la cadena «Al Form_Load se ha producido» pese a estar en manejadores como `ComandoAyuda_Click`, `ComandoRegistrar_Click`, `ComandoURLPlantillasExcel_Click` (`:42-48, 84-89, 218-223`): copy-paste que pierde utilidad de log (defecto claro, a verificar si se conserva o se corrigió en runtime).
- **`TempVars("ActivadoCorreoAutomatico") = "No"`** no desactiva `ComandoGestionCorreos` en el menú (contrastar con `Form_Form0BDOpciones.cls:251-262` donde la rampa está comentada y el botón se queda siempre activo): aquí sí pinta la etiqueta «Aún no implementado correos automáticos» (`:355-358`).
- **Concurrencia:** no hay `BeginTrans`/`CommitTrans`/`Rollback` en `ConfiguracionOperaciones.Registrar`; una escritura concurrente desde otra ventana puede sobrescribir la fila (a verificar).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Configuración actual (al pintar) | `TbConfiguracion` | `QueryDef` / DAO (`SELECT TOP 1 *`) | `classes/ConfiguracionOperaciones.cls:1-100`, `:301-307` |
| Persistencia de cambios | `TbConfiguracion` (la única fila, autoincremento manual con `Max+1` si es primer insert) | `INSERT` / `UPDATE` DAO | `classes/ConfiguracionOperaciones.cls:Registrar` |
| Apertura de carpeta de plantillas | API Shell (`Ejecutar`) | `ShellExecute` con `"open"` y `fso.FolderExists` | `:213, 234` |
| Invalida caché de entorno | `m_ObjEntorno.Configuracion = Nothing` | setter del composition root | `:72` |

## Coexistencia

`TbConfiguracion` debe migrar a `lanzadera.hps_solicitudes.configuracion` (candidata por nombre). El `Max+1` para `ID` colisiona con una clave primaria `SERIAL` de PostgreSQL; idoneidad de mantener el patrón o migrar a `INSERT … RETURNING id`. `m_ObjEntorno.Configuracion = Nothing` se invoca al guardar pero NO al cerrar — punto de rediseño.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **Mensajes de error que mienten.** Los `errores:` de los cuatro manejadores principales repiten el literal «Al Form_Load se ha producido» (`:42, 84, 218, 280`): un fallo en `ComandoRegistrar_Click` se reporta como «Al Form_Load», ocultando el origen real al depurar.
- **`HaHabidoCambios` se rompe en silencio si se añade un campo.** Cada nueva propiedad exige una nueva comparación manual (`:386-457`); sin test automatizado, los campos nuevos quedan «modificables» pero sin marcar como cambio → `ComandoRegistrar` los ignora hasta que otro campo dispare `m_HaHabidoCambios = True`. La alta inicial está cubierta por la rama `m_ConfAlInicio Is Nothing` (`:386-389`) pero un campo nuevo en alto se guarda sin validación de cambio.
- **`cmdSalir` descarta cambios sin pedir confirmación** (`:13-15`): si el usuario editó cinco campos y pulsa «Salir» sin guardar, los cambios se pierden sin aviso.
- **`Set m_ObjEntorno.Configuracion = Nothing` solo al guardar** (`:72`): cerrar con «Salir» deja la caché con datos potencialmente obsoletos.
- **`TempVars("ActivadoCorreoAutomatico") = "No"` desactiva campos del formulario pero no así el botón del menú** (`Form_Form0BDOpciones.cls:251-262`): desde `Form_FormConfiguracion` la rampa se nota; desde el menú no.
- **`fso` global sin inicialización local** (`:213, 234`): si el usuario llega con la sesión de EVE sin `fso`, los dos botones de URL fallan sin diagnóstico claro.
- **`RellenarObjeto` no valida formato** (`:467-501`): si el usuario teclea «abc» en `DiasParaRecordatorioExcel1`, se guarda tal cual y la conversión a entero queda en manos del backend (o se rechaza silenciosamente).
- **`ComandoRegistrar_Click` renueva `m_ObjEntorno.Configuracion = Nothing` antes de `Registrar`** (`:72, 77`): si `Registrar` falla, la caché queda vacía y `Form0BDOpciones` lee el siguiente `getConfiguracion` desde BD; sin embargo el evento `Modificado` se lanza solo en éxito (`:78`) — el menú solo se entera del éxito, no del fallo.
- **`Event Modificado` se lanza también si no había cambios reales** porque `HaHabidoCambios` puede devolver `True` por una comparación spuria (p. ej. `""` vs `Null`).
- **`VersionPlantillasHTML`/`VersionPlantillasEXCEL` no se validan contra la versión fija del backend**.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: detección de cambios derivada (un único `Iterator` sobre los pares control-propiedad); `fso` pasado por composición al composition root; `cmdSalir` con confirmación por diff; eventos `Modificado` siempre diferidos al commit; validación de campos numéricos en cliente antes de `RellenarObjeto`.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir la pantalla muestra los valores actuales de `TbConfiguracion` y la etiqueta «Parámetros COMPLETOS» en el menú padre.
2. Cambiar un campo numérico y pulsar «Registrar» persiste el cambio; «Salir» descarta sin guardar.
3. Cambiar «Correos automáticos» desactiva/activa los campos `Tag = "CORREO"` y notifica al menú.
4. La validación de días rechaza valores no numéricos antes de «Registrar».
5. La rampa «ActivadoCorreoAutomatico = "No"» se aplica al menú y a la pantalla.

## Pendientes

- Origen exacto de las globales `m_TituloFormulario`, `m_ObjEntorno`, `m_FaltaAlgunParametroEnConfiguracion`, `TempVars("ActivadoCorreoAutomatico")` (`VG.bas`, `FU.bas`).
- Definición de `constructor.getConfiguracion` en `modules/constructor.bas` y el `ID` que asigna en `ConfiguracionOperaciones.Registrar` cuando la tabla está vacía.
- Comportamiento del `Max+1` para nuevos `ID` de `TbConfiguracion` si la tabla está vacía o se purga (riesgo de colisión con autoincrementos en coexistencia con PostgreSQL).
- Si el campo `CorreodeEnvio` (Tag=`CORREO`) debería pintarse desde `TbConfiguracion` (no aparece en `EstablecerDatos`) — candidato a omisión.
- Relación con los `TempVars("EnDesarrollo")` que el splash fija y este formulario no consulta.
- Tipo de dato real de cada columna de `TbConfiguracion` y precisión de los días (entero corto vs entero largo).
