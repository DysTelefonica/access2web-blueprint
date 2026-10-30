# Form_FormAnexos

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormAnexos.cls` (493 líneas) y `src/forms/FormAnexos.form.txt` · **PRD:** [README](../README.md)

Gestor de anexos y evidencias de una edición o de un riesgo. Sube un fichero desde el disco del usuario, lo copia al directorio de documentación de la aplicación, da de alta la fila y permite renombrar el título y borrar el anexo. Es el formulario vivo del par: `Form_FormAnexos1` es su predecesor y no lo abre nadie. Las citas son `Form_FormAnexos.cls:línea` salvo indicación.

## Entrada y salida

No tiene `OpenArgs` con clave y valor: recibe **una sola palabra** por `openArgs` (`:420`) que decide el ámbito, y el objeto concreto lo toma de una variable global de sesión.

| Ámbito | `openArgs` | Objeto de contexto | Evidencia |
|---|---|---|---|
| Proyecto | `Proyecto` | `m_ObjProyectoActivo` | `:423-428` |
| Edición | `Edición` | `m_ObjEdicionActiva` | `:429-439` |
| Riesgo | `Riesgo` | `m_ObjRiesgoActivo` | `:440-447` |

| Quién lo abre | Evidencia |
|---|---|
| `FormRiesgosGestion`, botón de anexos | `Form_FormRiesgosGestion.cls:686` (con cierre previo si ya estaba abierto, `:681-683`) |

Ese único punto **solo genera `Edición` o `Riesgo`** (`Form_FormRiesgosGestion.cls:669-675`); la rama `Proyecto` está comentada (`:676-678`). La rama de proyecto, por tanto, solo se alcanza desde el binario. `Form_FormProyectosGestion.cls:23` declara `WithEvents m_FormAnexos As Form_FormAnexos` e implementa los dos manejadores (`:832`, `:873`), pero **nunca abre el formulario**: los eventos `AnexoAñadido` y `AnexoEliminado` no tienen oyente efectivo.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Hourglass, tamaño, contexto por `openArgs`, permiso de escritura y primer listado | `:408-466` |
| `btnExaminar` | botón | `Click` | Diálogo de selección de fichero y sugerencia de título desde el nombre del fichero | `:62-97`, `.form.txt:972` |
| `txtRuta` | texto | — | Ruta local del fichero elegido | `.form.txt:1382` |
| `btnAceptar` | botón | `Click` | Valida, instancia `Anexo` y persiste | `:103-197`, `.form.txt:1340` |
| `TituloDoc` | texto | — | Título del anexo; se autocompleta con `fso.GetBaseName` | `:89`, `.form.txt:819` |
| `EvidenciaUTE` | opción | — | Solo habilitada si el proyecto está en UTE | `:433-435`, `.form.txt:916` |
| `btnActualizarNombreArchivo` | botón | `Click` | Renombra el título del anexo seleccionado | `:199-231`, `.form.txt:1439` |
| `ListaDocumentos` | lista | `Click` / `DblClick` | Selecciona, carga el anexo y habilita acciones; doble clic abre el fichero | `:331-373`, `:490-492`, `.form.txt:781` |
| `ImagenAnexo` | imagen | `Click` | Abre el anexo con `AbrirEnLocal` | `:479-484`, `.form.txt:627` |
| `ComandoEliminarAnexo` | botón | `Click` | Confirma y borra fila y fichero | `:375-406`, `.form.txt:1305` |
| `cmdSalir` | botón | `Click` | Cierra sin guardar | `:486-488`, `.form.txt:1485` |
| `lblTitulo1` | etiqueta | — | Encabezado según el ámbito | `:428, 438, 447`, `.form.txt:608` |
| `ComandoAyuda` | botón | — | **Sin manejador en la clase**; su rótulo en el diseño es «Salir» | `.form.txt:699-700` |

## Flujos

### Apertura (`Form_Open`, `:408-466`)

1. Hourglass, `AjustarTamaño Me` y deshabilita `EvidenciaUTE`, `ComandoEliminarAnexo` y `btnAceptar` (`:413-418`).
2. Lee `openArgs` (`:420`) y carga el objeto global correspondiente (`:422-448`).
3. **Permiso de escritura por `EsActivo` del contenedor** (`:450-456`): si no está activo, `btnAceptar` sigue deshabilitado.
4. `AdaptarTamañoFormulario` muestra u oculta el bloque de evidencia UTE (`:457`, `:468-477`) y `EstablecerLista` puebla la lista (`:458`).

El manejador de error muestra el mensaje y **cierra el formulario** (`:463-466`), sin asignar `Cancel`.

### Flujo normal: subir un anexo (`btnAceptar_Click`, `:103-197`)

1. Ruta vacía → aviso y foco en `btnExaminar` (`:115-118`); título vacío y no-UTE → aviso y foco en `TituloDoc` (`:121-125`).
2. **Duplicado de título** con `ExisteNombreEnAmbito(m_Titulo, "0")` → aviso con el texto del título (`:130-134`).
3. Instancia `Anexo`, asigna el ámbito y, si es evidencia UTE sobre una edición, **fuerza el título** a un texto fijo con el número de edición (`:149-152`).
4. Si ya hay evidencia UTE, pregunta y **borra la anterior antes de registrar la nueva** (`:162-166`).
5. `Registrar m_ObjDELAnexo, m_URL, m_Error` (`:173`): valida, inserta la fila y copia el fichero.
6. `RaiseEvent AnexoAñadido` (`:178`), recarga la lista (`:179`), limpia los campos (`:182-186`) y confirma (`:187`).

### Flujo alternativo: renombrar (`btnActualizarNombreArchivo_Click`, `:199-231`)

Excluye el propio `IDAnexo` de la comprobación de duplicados (`:217-221`), llama a `CambiarNombre` (`:225`), recarga y confirma. Sin cambios no hace nada (`:214`).

### Flujo de error

`Err.Raise 1000` significa «mensaje de negocio ya preparado»: el manejador lo muestra tal cual; cualquier otro error se muestra con `Err.Description` (`:190-197`). En `EstablecerLista` el criterio es el mismo (`:328`). El renombrado **no distingue** los dos casos (`:234-236`).

## Reglas, validaciones y cálculos

- **El permiso no es un rol, es el estado del contenedor**: `EsActivo` de proyecto, edición o riesgo (`:450`). El perfil del usuario solo se comprueba dentro de `Anexo` (`Anexo.cls:301-308` y `:459-465`).
- **La unicidad del título es de interfaz, no de base de datos.** `ExisteNombreEnAmbito` recorre el diccionario en memoria (`:243-267`), sin distinguir mayúsculas (`StrComp`, `:255`).
- **Una sola evidencia UTE por edición.** La comprobación se hace con `Edicion.AnexoEvidenciaUTE` (`Edicion.cls:1103-1122`), que resuelve con `Constructor.getAnexoEvidenciaUTE` (`Constructor.bas:2659-2685`, `ORDER BY IDAnexo DESC`).
- **La lista cambia de alcance según el ámbito:** con proyecto usa `ColAnexosTotales`, que recorre el proyecto, sus ediciones y sus riesgos (`Proyecto.cls:2003-2050`); con edición o riesgo usa `ColAnexos` directo (`:287-292`).

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Fichero de origen | Ruta local elegida por el usuario | `:77-78`, `:173` |
| Fichero de destino | `m_ObjEntorno.URLDirectorioDocumentacion` + nombre generado | `Anexo.cls:330-332` |
| Alta de fila | `TbAnexos`, `AddNew`/`Update` | `Anexo.cls:376-402` |
| Copia del fichero | `fso.CopyFile` origen → destino | `Anexo.cls:407` |
| Renombrado | `UPDATE TbAnexos SET Titulo = …` concatenado | `Anexo.cls:433-436` |
| Baja de fila y fichero | `DELETE FROM TbAnexos` y `fso.DeleteFile` | `Anexo.cls:497-502` |
| Apertura | Copia a directorio local y `Shell open` | `Funciones Generales.bas:874-887` |

## Coexistencia

Pendiente de `../coexistencia.md`. Aquí el punto crítico es el fichero: el destino es una ruta de red compartida (`URLDirectorioDocumentacion`) y la apertura copia a un directorio local del usuario antes de lanzar el visor. Mientras convivan, la web y el legacy tienen que escribir en el mismo repositorio de ficheros, y `fso.FileExists` es la comprobación que decide si un anexo existe.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-60.** La sustitución de la evidencia UTE **borra la anterior antes de registrar la nueva** (`:162-166` frente a `:173`). Si `Registrar` falla, la edición se queda sin evidencia UTE y sin el fichero anterior: pérdida doble. El orden inverso es el que usa el predecesor (`Form_FormAnexos1.cls:217-226`).
- **H-GR-61.** No hay rama para un `openArgs` ausente o distinto de los tres valores (`:422-447`): `m_ObjDELAnexo` queda sin asignar y `:450` falla con «objeto no definido». El usuario ve un error genérico y el formulario se cierra. El predecesor sí tiene la rama explícita («parámetros insuficientes», `Form_FormAnexos1.cls:450`).
- **H-GR-62.** La comprobación de duplicados se hace solo contra el diccionario en memoria (`:243-267`): dos usuarios que anexan a la vez pasan los dos, y no hay restricción de unicidad en la tabla. Conviene verificar si existe un índice único sobre `TbAnexos`.
- **H-GR-63.** `CambiarNombre` concatena el título en el `UPDATE` (`Anexo.cls:433-435`) y el título es texto libre del usuario: un apóstrofe rompe la sentencia. Verificar si algún día se probó.
- **H-GR-64.** `EstablecerLista` lee `.riesgo.CodigoRiesgo` para `Tipo = "R"` (`:315`) sin comprobar que `riesgo` exista ni su `.Error`; si la fila quedó huérfana, el error se traga en `:328` y la lista sale incompleta sin aviso.
- **H-GR-65.** `ImagenAnexo_Click` usa `On Error Resume Next` (`:480`) y pasa `m_Error` a `AbrirEnLocal` (`:482`), pero **nadie lee `m_Error` después**: los mensajes «no es accesible la ruta» y «tiene el archivo abierto» (`Funciones Generales.bas:863-872`) no llegan nunca al usuario. Abrir un anexo inaccesible no produce ninguna señal.
- **H-GR-66.** `AbrirEnLocal` copia el fichero al directorio local del usuario y **no lo borra después** (`Funciones Generales.bas:874-887`): cada apertura deja una copia del documento en el perfil local, sin limpieza ni caducidad.
- **H-GR-67.** `Anexo.Registrar` inserta la fila **antes** de `fso.CopyFile` (`Anexo.cls:407`) y `EliminarAnexo` borra la fila **antes** de `fso.DeleteFile` (`Anexo.cls:501`): un fallo de disco deja la base de datos y el repositorio de ficheros desincronizados, en los dos sentidos.
- **H-GR-68.** El diseño guarda datos personales en el rótulo del formulario: `Caption =`«Usuario conectado: …» (nombre y apellidos y perfil) en `FormAnexos.form.txt:40`. No se reproduce aquí. Es texto de diseño, no un valor resuelto en ejecución, y conviene decidir si se limpia.
- **H-GR-69.** El control `ComandoAyuda` existe (`FormAnexos.form.txt:699`) pero no hay `ComandoAyuda_Click` en la clase, y su rótulo en el diseño es «Salir» (`:700`), duplicando `cmdSalir` (`:1485-1486`). Control muerto.
- **H-GR-70.** Los eventos `AnexoAñadido` y `AnexoEliminado` (`:178`, `:388`) no tienen oyente: `Form_FormProyectosGestion.cls:23` los implementa sin abrir nunca el formulario, y `Form_FormPublicacionCalidadPublicar.cls:16` declara `WithEvents` sin un solo manejador. El refresco del padre no depende de ellos, pero la suscripción sugiere una intención que no se cumplió.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: subida por HTTP con almacenamiento en servidor en lugar de un diálogo de fichero local y una ruta de red; título único garantizado por la base de datos en lugar de por el formulario; baja y alta del anexo en una sola operación transaccional con el fichero; apertura en streaming autenticado en vez de copiar a disco local.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con el contenedor activo, se puede examinar, aceptar, renombrar y borrar anexos.
2. Un título repetido en el mismo ámbito se rechaza antes de escribir.
3. Sustituir la evidencia UTE deja exactamente una evidencia UTE en la edición, o ninguna si el alta falla.
4. Con el contenedor no activo, los comandos de escritura aparecen deshabilitados.
5. Un anexo no alcanzable se comunica al usuario con un mensaje explicativo.

## Pendientes

- Si la rama `Proyecto` de `Form_Open` sigue siendo alcanzable desde el binario, dado que su único emisor en `src` la tiene comentada.
- Si existe índice único de título sobre `TbAnexos`.
- Dónde queda el directorio de documentación y si es alcanzable desde la red y desde el futuro servidor web.
- Qué hace `AjustarTamaño` con este formulario (`Funciones Generales.bas:1043`).