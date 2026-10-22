# Form_FormUsuario07SICAAlta

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormUsuario07SICAAlta.cls` (516 líneas) y `.form.txt` (2711 líneas) · **PRD:** [README](../README.md)

Formulario modal de **alta de usuario en el sistema SICA**. No está enlazado a datos (`RecordSource` no aparece en el `.form.txt`): construye un `UsuarioSICA` nuevo en memoria (`VariablesEntorno.bas` lo declara como global), copia desde los controles etiquetados con `DATOS` y lo graba en `TbUsuariosSICA`. La ficha de datos se divide según `TramitacionExterna`: con valor `Sí` se teclean los identificadores de SICA y las cuatro fechas; con valor `No` solo se indica la empresa tramitadora. Cerrar sin grabar es `CmdSalir` (`:41-43`). Las citas son `Form_FormUsuario07SICAAlta.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas` y `Constructor.bas` `modules/Constructor.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial07UsuariosSICA.ComandoNuevoUsuarioSICA_Click` | Descarta el global, cierra el formulario si ya está abierto y abre este con `DoCmd.OpenForm` | `Form_FormInicial07UsuariosSICA.cls:166, 167-170` |
| salida de `ComandoGrabar_Click` | Tras el alta avisa al administrador y refresca la caché y el formulario central del inicial | `:122-126`; `FG.bas:3402-3417` |
| `CmdSalir_Click` / error de `Form_Load` | Cierra con `acSaveYes` y con `acSaveNo` respectivamente | `:42, 201` |

Abre el selector de ficheros del sistema (`:52`; `FG.bas:880-882`) y el visor de la aplicación registrada para el anexo (`:218`; `VariablesEntorno.bas:5`). No abre otros formularios de Access.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Ajusta tamaño, decide solo lectura, crea el `UsuarioSICA`, rellena los tres combos y los botones | `:144-202`; `.form.txt:373` |
| (formulario) | — | — | `Caption = "Alta de Usuario"`, sustituido en `:187` por el de `lblTitulo` («ALTA DE USUARIO SICA») | `.form.txt:38, 919` |
| Trece controles de datos: `Nombre`, `Apellido_1`, `Apellido_2`, `Correo_e`, `Observaciones`, `UsuarioSICATDE`, `UsuarioSICATSOL`, `IDEmpresaTramitadora`, `TramitacionExterna`, `FAltaTdE`, `FBajaTdE`, `FAltaTSOL`, `FbajaTSOL` | textbox y combobox, todos con `Tag ="DATOS"` | — | Los única que `RellenarObjetoConDatos` vuelca al objeto | grupo `.form.txt:1023-1504`, fechas en `.form.txt:1700-1887`; `FG.bas:4131-4137` |
| `UsuarioHPS`, `UsuarioHistorico` | combobox `Value List` | `OnExit` → `UsuarioHPS_Exit`, `UsuarioHistorico_Exit` | Copian la identidad de un usuario HPS o histórico hacia los campos de SICA | `:354-516`; `.form.txt:1380, 2069` |
| `TramitacionExterna` | combobox | `BeforeUpdate` → `TramitacionExterna_BeforeUpdate` | Activa, limpia o bloquea el resto de la ficha según el valor elegido | `:233-352`; `.form.txt:1339` |
| `ComandoAnexar`, `ImagenAnexo` | botón «Anexo», imagen adjunta | `OnClick` → `ComandoAnexar_Click`, `ImagenAnexo_Click` | Elige un único fichero (no se copia hasta grabar) y lo abre | `:45-75`, `:204-230`; `.form.txt:1913, 2005` |
| `ComandoEliminarAnexoActual` | botón «Eliminar» | `OnClick` → `ComandoEliminarAnexoActual_Click` | Descarta el fichero elegido sin llegar a tocar disco | `:77-99`; `.form.txt:1960` |
| `ComandoGrabar`, `ComandoEmpresas`, `cmdSalir` | botones «Grabar», «Editar», «Salir» | `OnClick` de los tres | Grabar es el que persiste; «Salir» cierra; «Editar» no tiene handler en este `.cls` | `:103-141`; `.form.txt:1529, 2109, 2152` |
| Cuatro etiquetas con `Tag ="NOCAMBIABLE"` | etiqueta «*» | — | Marcan los campos de SICA como opcionales | grupo `.form.txt:1591-1647` |

## Flujos

### Carga (`Form_Load`, `:144-202`)

1. Ajusta el tamaño y traduce `Application.TempVars("SoloLectura")` a `AllowEdits` (`:153-158`).
2. Crea el objeto de negocio en memoria y marca `ParaAlta = EnumSiNo.Sí` (`:161-162`).
3. Aplica la paleta de controles (`:163`; `FG.bas:1005`) y habilita `ComandoGrabar` y `ComandoEmpresas` solo si `AllowEdits` (`:167-173`).
4. Rellena los tres combos: empresas tramitadoras (`EnumSiNo.No`, `:174`), usuarios HPS (`:179`) e históricos (`:183`).
5. Copia el rótulo del formulario (`:187`) y el estilo de los botones (`:188-189`).
6. En error compone el mensaje, avisa al administrador, muestra `MsgBox` crítico y **cierra el formulario** (`:194-201`).

### Volcar el formulario al objeto (`RellenarObjetoConDatosLocal`, `:10-39`)

1. `RellenarObjetoConDatos` recorre `frm.Controls` y solo copia los controles cuyo `Tag` contiene `DATOS` (`:19`; `FG.bas:4124-4139`). Los combos `UsuarioHPS` y `UsuarioHistorico` no llevan esa etiqueta, y por eso quedan fuera.
2. Si el combo `UsuarioHPS` no está vacío, extrae el identificador del texto con `getIDUsuario` (`:24-25`; `FG.bas:2565`) y lo asigna a `IDHPS` (`:30`). El lector toma el valor entre paréntesis del elemento elegido (`FG.bas:2577-2588`).

### Elegir usuario HPS o histórico (`UsuarioHPS_Exit`, `:436-516`; `UsuarioHistorico_Exit`, `:354-434`)

Ambos handlers son simétricos salvo en la fuente del identificador: `getIDUsuario` para el combo de HPS (`:454`) y `getIDUsuarioHistorico` para el de histórico (`:372`), seguidos de `getUsuarioHPS` (`:464`) o `getUsuarioHistorico` (`:382`).

1. Si el formulario está en solo lectura o el identificador está vacío, sale sin tocar nada (`:448-456`, `:366-374`).
2. Si el objeto no existe, sale sin error (`:468-473`, `:385-390`).
3. Rellena el combo de empresas con `EnumSiNo.Sí` y fija la empresa del usuario, deshabilitando el combo y el botón «Editar» (`:475-481`, `:393-399`).
4. Copia nombre, apellidos y correo, los deja deshabilitados y habilita los cuatro campos de SICA más observaciones (`:482-500`, `:400-419`).
5. `Apellido_2` se anula cuando el usuario de origen no lo tiene (`:484-489`, `:404-408`), y el foco queda en `UsuarioSICATDE` (`:501, 419`).

### Cambiar `TramitacionExterna` (`:233-352`)

1. En solo lectura, sale sin hacer nada (`:241-246`).
2. Si el valor queda vacío, anula y bloquea empresa, nombre, apellidos, correo, los dos identificadores de SICA y las tres fechas, recuerda el estado anterior como cadena vacía y sale (`:249-278`). `FBajaTdE` no aparece en esta lista.
3. Si el valor es `Sí` y el anterior ya era `Sí`, sale (`:281-289`). Si no, anula el combo de `UsuarioHPS`, deshabilita la ficha de identidad y habilita empresa, identificadores de SICA, cuatro fechas y observaciones (`:290-306`).
4. Si el valor es `No` y el anterior ya era `No`, sale (`:307-313`). Si no, anula empresa, nombre, apellidos, correo, identificadores de SICA, tres fechas y observaciones (`:314-337`). `FBajaTdE` vuelve a quedar fuera de la lista.
5. El botón «Anexo» se habilita en cuanto el combo tiene valor (`:279`).

### Grabar (`ComandoGrabar_Click`, `:103-141`)

1. Vuelca los controles al objeto (`:113-116`).
2. `m_ObjUsuarioSICAActivo.Alta m_URLLocal, m_Error` valida, asigna `ID`, inserta la fila en `TbUsuariosSICA` y, si hay anexo, lo copia y lo anexa (`:118`; `UsuarioSICA.cls:773-856`).
3. Recoge el `ID` nuevo, muestra «Alta correcta» y llama a `ActualizarUsuarioSICA`, que invalida y reconstruye las cachés locales del SICA y refresca el formulario central si el inicial está en esa pantalla (`:122-126`; `FG.bas:3402-3417`; `CacheConsistencyAudit.bas:139-143`).
4. En error compone el mensaje, avisa al administrador y muestra `MsgBox` crítico (`:135-140`).

### Elegir o abrir el anexo (`:45-75`, `:204-230`)

1. `Seleccionar1` devuelve **un solo** fichero y recuerda la última carpeta usada (`:52`; `FG.bas:880-891`).
2. Si el fichero no existe, vacía la ruta y oculta la imagen; si existe, la muestra (`:56-63`). En ambos casos activa el botón «Eliminar» (`:59, 62`).
3. Al pulsar la imagen, comprueba la existencia y abre el fichero con `Ejecutar` (`:213-218`); si no existe, muestra «No es alcanzable en estos momentos el anexo seleccionado» con la ruta (`:214`).
4. El botón «Eliminar» solo vacía la ruta y oculta la imagen (`:85-86`): la copia al destino se produce al grabar.

## Reglas, validaciones y cálculos

- **Solo los controles con `Tag ="DATOS"` viajan al objeto** (`.form.txt:1024, 1304, 1504, 1887`; `FG.bas:4132`). Es la convención que sostiene todo el alta y la edición: un control nuevo sin esa etiqueta se guarda vacío sin aviso.
- **La regla de negocio vive en `UsuarioSICA`, no en el formulario.** `Alta` llama a `MotivoNoOK` antes de escribir (`UsuarioSICA.cls:787-794`); el formulario solo recoge el mensaje y lo propaga.
- **Los combos llevan el identificador entre paréntesis** y `getIDUsuario` / `getIDUsuarioHistorico` lo extraen por troceado de cadena (`FG.bas:2581-2588, 2612`). Es acoplamiento al formato de la etiqueta, no al `ListIndex`.
- **El anexo no se copia al elegirlo.** `ComandoAnexar` solo guarda la ruta en `m_URLLocal`; la copia la hace `UsuarioSICA.Anexar` dentro de `Alta` (`UsuarioSICA.cls:843-848`).
- **El alta no cierra el formulario.** La línea de cierre está comentada (`:131`), de modo que la pantalla permanece con los mismos datos tras «Alta correcta».
- **Identidad por `TramitacionExterna`.** Con `Sí`, `IDHPS` no procede; con `No`, se teclean nombre, apellidos, correo y las cuatro fechas (`:249-338`).
- **Solo lectura llega por `Application.TempVars("SoloLectura")`** y se traduce a `AllowEdits` y a los dos handlers de `Exit` (`:154-158, 241, 366, 448`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Inserción del usuario SICA | `AddNew` sobre `TbUsuariosSICA` con `FechaCreacion` y `FechaModificacion` a `Now()` | escritura | `:118`; `UsuarioSICA.cls:795, 803-839` |
| Copia del anexo al alta | `AnexoUsuarioSICA` / fichero local | escritura de fichero | `:118`; `UsuarioSICA.cls:843-848` |
| Identidad del usuario HPS elegido | `getUsuarioHPS` → usuario HPS | lectura | `:464`; `Constructor.bas:743` |
| Identidad del histórico elegido | `getUsuarioHistorico` | lectura | `:382`; `Constructor.bas:1198` |
| Listas de los combos | `getListaUsuariosHPSParaCombo`, `getListaUsuariosHistoricosParaCombo`, empresas | lectura en memoria | `:174, 179, 183`; `FG.bas:2175, 2207` |
| Refresco de cachés tras el alta | `DeleteSicaCacheRows` + `UpsertSicaCacheRow` sobre `TbUsuariosSICALocal` y `TbUsuariosSICALocalParaIndicadores` | escritura en la base local | `:124`; `CacheConsistencyAudit.bas:139-143` |
| Apertura del anexo | `ShellExecuteA(hwnd, "open", url, "", "", 1)` | API declarada en `VariablesEntorno.bas` | `:218`; `VariablesEntorno.bas:5` |

## Coexistencia

Pendiente de `../coexistencia.md`. Lo concreto: `Application.FileDialog`, `DoCmd.Hourglass`, `DoCmd.Close` y `Application.TempVars` son APIs de Access sin equivalente directo; los ocho `DoCmd.Hourglass True` / `DoCmd.Hourglass False` que envuelven cada operación (`:48-50, 64-66`, entre otros) son ruido imperceptible en web. `fso` (`VariablesEntorno.bas:218`) y `Ejecutar` (`VariablesEntorno.bas:5`) no se traducen: el anexo debe ser un recurso del servidor. El trasvase de datos por `Tag` (`FG.bas:4132`) es un patrón de reflexión que en web debe reemplazarse por un formulario con binding declarado. La edición de empresa («Editar», `.form.txt:1529`) no tiene handler propio en este `.cls`: se resuelve en otro formulario, a verificar.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **`FBajaTdE` se queda fuera de la limpieza en las dos ramas restrictivas.** La rama de valor vacío (`:249-278`) y la rama `No` (`:307-337`) anulan y deshabilitan `FAltaTSOL`, `FAltaTdE` y `FbajaTSOL`, pero no `FBajaTdE`, que la rama `Sí` sí habilita (`:303`). Consecuencia posible: pasar de `Sí` a `No` deja visible y editable una fecha de baja de TDE con el valor que tuviera, y `Alta` la escribiría en el registro (`UsuarioSICA.cls:833`). A verificar contra el binario.
- **El histórico elegido nunca llega al registro.** `UsuarioHistorico_Exit` (`:354-434`) copia nombre, apellidos, correo y empresa, pero no guarda el `IDHPSHistorico` en el objeto, y el control `UsuarioHistorico` no lleva `Tag ="DATOS"` (`.form.txt:2069`), así que `RellenarObjetoConDatos` lo ignora. Solo el combo de HPS recibe su clave (`:30`). A verificar: quizá esa relación la rellene el alta a partir de la empresa o del correo.
- **El evento `AltaRegistro` está declarado y nunca se lanza** (`:7`): no hay `RaiseEvent` en el módulo. Código muerto o contrato previsto para un consumidor que ya no existe.
- **Grabar dos veces inserta dos filas.** El cierre está comentado (`:131`) y ni el objeto ni el formulario se invalidan tras el alta, así que un segundo `ComandoGrabar` vuelve a pasar por `Alta` con la misma `URLLocal` y vuelve a anexar el fichero (`:118`; `UsuarioSICA.cls:844`). A verificar si algo externo cierra la pantalla.
- **`CmdSalir` guarda y el error de carga no** (`:42` frente a `:201`): cerrar con cambios en la pantalla de solo lectura intentaría guardar; el formulario no está enlazado, así que el efecto real depende del binario.
- **La validación real no está en el formulario.** `MotivoNoOK` (`UsuarioSICA.cls:787`) decide, y el formulario solo distingue `Err.Number = 1000` de cualquier otro error para no pisar el mensaje (`:136-139`). Un `Err` no previsto por la clase llega al usuario como descripción en crudo.
- **`pregunta` es un global compartido** (`VariablesEntorno.bas:220`) usado aquí solo para el «Alta correcta» (`:123`) y los `MsgBox` de error; su valor no se vuelve a leer.
- **Restos:** `RellenarObjetoConDatosLocal` declara `m_NombreCampo`, `m_ValorCampo`, `m_objColCampos` y `ctl` sin usarlas (`:12-16`), y devuelve `String` sin asignar valor (`:10`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: paso de wizard o una sola pantalla con secciones condicionadas en lugar de habilitar y deshabilitar controles sobre la marcha; validación declarativa en cliente además de la de servidor; selector de fichero con subida inmediata en lugar de diferirla al grabar; previsualización del anexo con botón de descarga; mensaje de éxito como notificación y no como `MsgBox` bloqueante; cierre y retorno automático a la lista tras un alta correcta.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, el combo de usuarios HPS e históricos está relleno y el resto de campos, vacíos.
2. Con `TramitacionExterna` vacío, los campos de identidad y de SICA están vacíos y deshabilitados.
3. Con `TramitacionExterna = Sí`, se habilitan los dos identificadores de SICA y las cuatro fechas, y el combo de usuario HPS queda vacío y deshabilitado.
4. Con `TramitacionExterna = No`, se habilita la empresa tramitadora y se vacían y bloquean el resto de campos de identidad y de SICA, **incluidas las cuatro fechas**.
5. Elegir un usuario HPS o un histórico rellena empresa, nombre, apellidos y correo, los deja no editables y habilita la ficha de SICA.
6. Elegir fichero muestra su icono; un segundo clic sobre el icono lo abre con la aplicación del sistema; «Eliminar» lo descarta sin tocar la carpeta de origen.
7. «Grabar» con los datos válidos inserta una fila en `TbUsuariosSICA`, copia el anexo si se eligió uno, confirma con «Alta correcta» y refresca la caché local.
8. Con datos que la clase rechaza, se muestra el mensaje de `MotivoNoOK` y no se inserta nada.
9. En sesión de solo lectura, «Grabar» y «Editar» están deshabilitados y los `Exit` de los combos no modifican la pantalla.

## Pendientes

- Si el alta de un usuario SICA a partir de un histórico debe rellenar `IDHPSHistorico`: en el snapshot no hay ninguna asignación a esa propiedad.
- Qué reglas concretas impone `UsuarioSICA.MotivoNoOK` (`UsuarioSICA.cls:354`) y cuáles son las que el usuario ve como mensaje, para poder derivar los criterios de aceptación desde la clase y no desde el formulario.
- Qué hace el botón «Editar» (`.form.txt:1529`), que no tiene handler en este `.cls`.
- Si `UsuarioSICA.Anexar` (`UsuarioSICA.cls:692`) copia el fichero o solo registra la ruta, y qué ocurre si el destino ya existe.
- Si el cierre de la pantalla tras el alta depende de algún evento externo (`AltaRegistro`, `:7`, nunca lanzado).