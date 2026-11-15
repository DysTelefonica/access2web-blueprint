# Form_FormInicial05Configuracion

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial05Configuracion.cls` (79 líneas) y `.form.txt` (1241 líneas) · **PRD:** [README](../README.md)

Formulario de **diagnóstico y transferencia**: muestra en tres cajas de solo lectura en qué régimen se está ejecutando la aplicación (estado de conexión, si se está en la oficina y si se han descargado datos de oficina) y ofrece un único botón que replica, con `ROBOCOPY /MIR`, las carpetas de anexos y recursos de la oficina en la carpeta local del aplicativo. No escribe en ninguna tabla: su único efecto es sobre el sistema de archivos. Se monta como `SubFormCentral` del principal al pulsar la etiqueta `lblConfiguracion` del lateral (`Form_FormInicial01Lateral.cls:424-432`). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`). Las citas son `Form_FormInicial05Configuracion.cls:línea` salvo indicación; `FG.bas` abrevía `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial01Lateral` | Al pulsar `lblConfiguracion` se asigna `SubFormCentral.SourceObject = "FormInicial05Configuracion"` y se recolorea el lateral | `Form_FormInicial01Lateral.cls:424-428` |
| `Application.CurrentProject.path` | El destino de la copia es la carpeta donde está el archivo `.accdb` abierto, no una ruta elegida por el usuario | `:21` |
| `m_ObjEntorno.URLCarpetaAnexos`, `URLCarpetaAnexosSICA`, `URLCarpetaAnexosHistoricos`, `URLCarpetaRecursos` | Carpetas de origen que la transferencia debe alcanzar | `FG.bas:2333, 2341, 2349, 2357` |
| `m_EnOficina` | Régimen de ejecución; gobierna la habilitación del botón y es la primera validación de la copia | `:47-57`, `FG.bas:2321-2324` |

No abre ningún formulario.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Rellena los tres campos de diagnóstico y decide si el botón de importación queda habilitado | `:40-78`, `.form.txt:372` |
| `ComandoImportarDatosOficinaALocal` | botón | `OnClick` → `ComandoImportarDatosOficinaALocal_Click` | Copia las carpetas de la oficina en la carpeta local y confirma con un `MsgBox` informativo | `:7-38`, `.form.txt:899-910` |
| `EstadoConexion` | caja de texto, `Enabled` y `TabStop` a `NotDefault` | — | Réplica de `m_EstadoConexion`: «Lectura/Escritura» o «Solo lectura» | `:45`, `.form.txt:1010-1024` |
| `EstoyEnLaOficina` | caja de texto, no editable | — | «Sí» o «No» según `m_EnOficina` | `:53-57`, `.form.txt:1068-1082` |
| `DatosDeLaOficina` | caja de texto, no editable | — | «Sí» o «No»; hoy refleja el mismo `m_EnOficina` que la anterior | `:47-51`, `.form.txt:1126-1140` |
| `Marco23` | marco de opciones | — | Agrupa el rótulo del botón; no contiene ninguna opción marcada, el `.form.txt` solo declara una etiqueta dentro | `.form.txt:945-979` |
| `Etiqueta24` | etiqueta («Obtener Datos de Oficina») | — | Rótulo de la zona de importación | `.form.txt:962-970` |
| `Etiqueta10`, `Etiqueta14`, `Etiqueta18` | etiquetas | — | Rótulos «ESTADO DE CONEXIÓN», «ESTOY EN LA OFICINA» y «DATOS DE LA OFICINA» | `.form.txt:995, 1053, 1111` |

## Flujos

### Apertura (`Form_Load`, `:40-78`)

1. `Me.EstadoConexion = m_EstadoConexion` (`:45`), cuyo valor decide `EstablecerSoloLectura` (`FG.bas:44, 56`).
2. Traduce `m_EnOficina` a «Sí»/«No» y lo escribe en `DatosDeLaOficina` (`:47-51`) y en `EstoyEnLaOficina` (`:53-57`). **Los dos leen la misma variable global.**
3. Deshabilita el botón de importación (`:59`).
4. Solo lo rehabilita si se cumplen las dos condiciones de `:62`: `EstoyEnLaOficina = "Sí"` **y** `DatosDeLaOficina = "Sí"`; y, dentro de ellas, solo si la ruta del proyecto **no** contiene `AppData\Roaming` (`:63-66`).
5. Si la ruta es la de un aplicativo instalado por usuario en Windows, el botón queda deshabilitado aunque se esté en la oficina.

`errores:` (`:71-77`) envía los errores distintos de 1000 al administrador y muestra un `MsgBox` crítico; el retorno se guarda en el global `pregunta` (`VariablesEntorno.bas:220`).

### Importación de datos (`ComandoImportarDatosOficinaALocal_Click`, `:7-38`)

1. El bloque que pedía la carpeta al usuario — `Seleccionar False, "Seleccione la carpeta para exportar los datos", …` y su guarda de vacío — está **comentado** (`:12-18`): la selección de destino ya no se produce.
2. `DoCmd.Hourglass True` (`:19`).
3. `URLCarpetaLocal = Application.CurrentProject.path` (`:21`).
4. `CopiarDatosDeOficinaACarpetaLocal URLCarpetaLocal, m_Error` (`:22`): valida régimen y existencia de carpeta (`FG.bas:2321-2328`), valida que las cuatro carpetas de origen existan (`FG.bas:2333-2354`), crea `ANEXOS`, `ANEXOS\HPS`, `ANEXOS\SICA`, `ANEXOS\HISTORICO` y `recursos` (`FG.bas:2361-2365`) y lanza cuatro `ROBOCOPY … /MIR /R:0 /W:0 /NP` (`FG.bas:2371-2392`).
5. Si la función devolvió error, relanza el 1000 y cae al manejador (`:23-25`).
6. `DoCmd.Hourglass False` y `MsgBox` informativo «Datos transferidos al aplicativo local correctamente» (`:26-27`).

`errores:` (`:31-37`) antepone su propio prefijo al error, envía correo al administrador, baja el cursor y muestra el `MsgBox` crítico.

## Reglas, validaciones y cálculos

- **Doble condición para habilitar la importación.** El botón exige estar en la oficina (`m_EnOficina = Sí`, evaluado dos veces, `:47-68`) y que la ruta no sea `AppData\Roaming` (`:63`), es decir, que el aplicativo no viva en el perfil del usuario.
- **El régimen «en la oficina» es una comprobación de IP.** `m_EnOficina` se fija en el arranque con `EnOficina`, que busca la subred de oficina entre las direcciones IP de la máquina (`VariablesEntorno.bas:333`; `FG.bas:2811-2816`; `SubRedOficina` en `VariablesEntorno.bas:219`).
- **La copia es espejo, no sincronización.** Las cuatro llamadas usan `/MIR` (`FG.bas:2371-2392`): lo que exista en el destino y no en el origen se borra. No hay confirmación previa ni registro de lo eliminado.
- **Origen y destino se calculan, no se eligen.** El destino es la ruta del proyecto abierto (`:21`) y los orígenes son cuatro propiedades de `m_ObjEntorno` (`FG.bas:2333-2358`); el usuario solo puede dispararlo, no dirigirlo.
- **Ninguna escritura en base de datos.** Todo el efecto es `DELETE`/`INSERT` de ficheros vía `ROBOCOPY`; no hay `CurrentDb.Execute` en la cadena (`FG.bas:2371-2392`).
- **`m_EstadoConexion` es una etiqueta, no un estado vivo.** La asigna `EstablecerSoloLectura` (`FG.bas:44, 56`), que además fija `Application.TempVars("SoloLectura")` (`FG.bas:43, 55`); este formulario solo lo refleja (`:45`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Copia de anexos HPS | carpeta de anexos de oficina → `<proyecto>\ANEXOS\HPS` | `ROBOCOPY /MIR` | `FG.bas:2333-2338, 2377-2380` |
| Copia de anexos SICA | carpeta de anexos SICA → `<proyecto>\ANEXOS\SICA` | `ROBOCOPY /MIR` | `FG.bas:2341-2346, 2383-2386` |
| Copia de anexos históricos | carpeta de anexos históricos → `<proyecto>\ANEXOS\HISTORICO` | `ROBOCOPY /MIR` | `FG.bas:2349-2354, 2389-2392` |
| Copia de recursos | carpeta de recursos → `<proyecto>\recursos` | `ROBOCOPY /MIR` | `FG.bas:2357-2358, 2371-2374` |
| Creación de carpetas destino | `ANEXOS`, `ANEXOS\HPS`, `ANEXOS\SICA`, `ANEXOS\HISTORICO`, `recursos` | `CrearCarpetaSiNoExiste` | `FG.bas:2361-2365` |
| Estado de la aplicación | `Application.TempVars("SoloLectura")`, `m_EstadoConexion`, `m_EnOficina` | variables en memoria | `FG.bas:43-44, 55-56`, `VariablesEntorno.bas:333` |

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Lo que ya se ve: la frontera de la copia es el sistema de archivos con una herramienta externa del sistema operativo y el modificador `/MIR`, cuyo borrado no es transaccional ni auditable desde la aplicación (`FG.bas:2371-2392`); el régimen se decide por inspección de IP (`FG.bas:2811-2816`); y la condición de habilitación comprueba una cadena de ruta del sistema operativo, `AppData\Roaming` (`:63`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `DatosDeLaOficina` y `EstoyEnLaOficina` muestran **el mismo valor**: ambos leen `m_EnOficina` (`:47-51` y `:53-57`). La condición de habilitación (`:62`) exige los dos a «Sí», de modo que una de las dos cajas es redundante y la etiqueta «DATOS DE LA OFICINA» sugiere un dato que no existe. El proyecto asigna `Application.TempVars("DatosEnLocal")` (`VariablesEntorno.bas:330`) y este formulario no lo usa, a verificar.
- La condición de `:63` comprueba `InStr(1, Application.CurrentProject.path, "AppData\Roaming") = 0`: es una heurística sobre la ruta, sensible a cómo se haya iniciado la aplicación y no a si el destino es escribible.
- El bloque de selección de carpeta está comentado (`:12-18`) y la copia se dirige al proyecto abierto (`:21`): el mensaje al usuario sigue diciendo «Datos transferidos…» sin indicar qué carpeta se ha tocado, y no hay confirmación antes de un `/MIR` que puede borrar destino.
- `Marco23` es un marco de opciones sin ninguna opción dentro; solo aloja una etiqueta (`.form.txt:945-979`): andamiaje de un control que ya no existe.
- `Form_Load` no lleva `DoCmd.Hourglass False` en su camino feliz ni en el manejador de error (`:71-77`), a diferencia de `ComandoImportarDatosOficinaALocal_Click` (`:36`): si otro procedimiento dejó el cursor de espera activo, este formulario no lo restaura.
- Los tres campos son de solo lectura por diseño (`Enabled = NotDefault`, `TabStop = NotDefault`, `.form.txt:1011-1012, 1069-1070, 1127-1128`): no hay forma de consultar el estado por otro medio que el rótulo, y no hay ningún indicador del error concreto si `EnOficina` falló al arrancar.
- El `MsgBox` de éxito (`:27`) se muestra aunque alguna de las cuatro copias no se haya realizado: la función devuelve `"OK"` siempre que `EjecutarShell` no devuelva error (`FG.bas:2394`), sin comparar archivos ni verificar destino.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: que el destino se elija explícitamente y se muestre antes de copiar; sustituir `/MIR` por una copia sin borrado, o exigir confirmación explícita para el borrado; y presentar el régimen (oficina, solo lectura, datos descargados) como estado declarado del cliente en lugar de tres textos derivados de variables globales.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrirse, `EstadoConexion` refleja el valor vigente de `m_EstadoConexion` (`:45`).
2. `EstoyEnLaOficina` y `DatosDeLaOficina` muestran «Sí» o «No» y solo habilitan la importación cuando ambos son «Sí» (`:47-68`).
3. El botón queda deshabilitado cuando la ruta del proyecto contiene `AppData\Roaming` (`:63`).
4. Al pulsar el botón, los anexos HPS, SICA e históricos y los recursos terminan reflejados en las subcarpetas `ANEXOS\HPS`, `ANEXOS\SICA`, `ANEXOS\HISTORICO` y `recursos` de la carpeta del proyecto (`FG.bas:2338-2358, 2371-2392`).
5. Cualquier carpeta de origen inalcanzable detiene la operación y el mensaje explica cuál es (`:22-25`; `FG.bas:2334-2337, 2342-2345, 2350-2353`).
6. Los errores distintos de 1000 se envían por correo al administrador y se muestran en un `MsgBox` crítico (`:32-37`).

## Pendientes

- Si `Application.TempVars("DatosEnLocal")` (`VariablesEntorno.bas:330`) debía alimentar `DatosDeLaOficina`: en `src/` solo se asigna «No».
- Contenido real de las cuatro carpetas de `m_ObjEntorno` y de la subred `SubRedOficina` (`VariablesEntorno.bas:219`), que no están en el snapshot.
- Si existe en el binario otro punto que asigne `SubFormCentral.SourceObject` a este formulario además del lateral.
- Por qué está comentado el bloque de selección de carpeta (`:12-18`): el comentario no lo explica.