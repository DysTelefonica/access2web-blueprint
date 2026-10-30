# Form_FormPlanificacionAnexos

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormPlanificacionAnexos.cls` (351 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Gestión de los documentos anexos de una planificación (`TbPlanificacionAnexos`): lista los anexos de la planificación recibida, permite adjuntar un fichero con un título, abrir el anexo seleccionado y eliminarlo (fichero y registro) mientras la planificación esté abierta. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); la lista se llena en `Form_Open` con SQL concatenada (`Form_FormPlanificacionAnexos.cls:135-140`). Las citas son `Form_FormPlanificacionAnexos.cls:línea` salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`. Capacidad: CAP-11 Anexos (con CAP-6 Planificación en la lectura de estado).

Controles: 15 (24 `Begin` en el `.form.txt`: 15 controles, `Begin Form`, `Begin Section` y 7 plantillas de estilo de las líneas 139-262).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormPlanificacionesBusqueda` (botón de anexos) | cierra la instancia abierta y hace `DoCmd.OpenForm "FormPlanificacionAnexos"` con el ID de planificación en `OpenArgs` | `Form_FormPlanificacionesBusqueda.cls:125-128` |
| `Form_FormPlanificacionesBusqueda` (segundo punto de apertura) | igual | `Form_FormPlanificacionesBusqueda.cls:886-889` |

`OpenArgs` es el ID de planificación y debe ser numérico (`:130-134`). No devuelve valores ni lanza eventos. Escribe en disco (copia y borrado de ficheros en el directorio de planificaciones configurado) y en `TbPlanificacionAnexos`. Sale con `cmdSalir` (`:6-9`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ListaDocumentos` | lista 5 columnas con encabezados, tipo «Value List» en diseño | handler `ListaDocumentos_Click` en el `.cls`; **sin `OnClick` en el `.form.txt`** | Toma el ID del anexo, resuelve la ruta, muestra el icono y habilita Eliminar según el estado | CAP-11 | `:183-254`, `.form.txt:343-358` |
| `TituloDoc` | texto no enlazado | — | Título del documento; obligatorio al adjuntar | CAP-11 | `:17-21`, `.form.txt:434` |
| `ComandoAnexar` («Adjuntar») | botón | handler `ComandoAnexar_Click`; sin `OnClick` en el `.form.txt` | Pide el título, abre el selector de fichero y llama a `AnexarAPlanificacion` | CAP-11 | `:10-69`, `.form.txt:626-627` |
| `ComandoEliminarDocumento` («Eliminar») | botón, deshabilitado en diseño | handler `ComandoEliminarDocumento_Click`; sin `OnClick` en el `.form.txt` | Confirma y llama a `EliminarAnexo` | CAP-11 | `:71-119`, `.form.txt:453, 463-464` |
| `ImagenDocumentoAnexado` | imagen, oculta en diseño | `OnClick` → `ImagenDocumentoAnexado_Click` | Abre el anexo seleccionado con `AbrirEnLocal` | CAP-11 | `:155-182`, `.form.txt:657, 664-665` |
| `cmdSalir` («Salir») | botón | handler `cmdSalir_Click`; sin `OnClick` en el `.form.txt` | Cierra sin guardar | CAP-11 | `:6-9`, `.form.txt:491-492` |
| `ComandoAyudaFormulario` | botón con icono | sin handler en el `.cls` ni `OnClick` | Sin comportamiento | CAP-11 | `.form.txt:518-519` |
| `lblTitulo1`, `lblTitulo2`, `Etiqueta142` | etiquetas | — | `lblTitulo1` pasa a `Caption` del formulario (`:128`); `Etiqueta142` rotula «TÍTULO» | CAP-11 | `.form.txt:560-561, 585-586, 605-606` |
| `CuadroPrincipal`, `CuadroSuperior`, `Cuadro2`, `Marco103`, `Marco104` | 3 rectángulos y 2 grupos de opciones | — | Decoración | sin asignar | `.form.txt:283, 305, 328, 386, 409` |

El único evento de formulario enlazado en el `.form.txt` es `OnOpen` (`.form.txt:37`).

## Flujos

### Apertura (`Form_Open`, `:120-154`)

1. `AjustarTamaño Me`, título de ventana desde `lblTitulo1` y `strIDPlanificacion = Nz(Me.OpenArgs, "")` (`:126-130`).
2. Si no es numérico: «No se ha abierto el formulario con los datos necesarios» (`:131-134`).
3. Asigna a `ListaDocumentos` la SQL de anexos de esa planificación (ID, planificación, título, fecha) y hace `Requery` (`:135-140`).
4. Cualquier error: `MsgBox` crítico y cierre del formulario (`:147-153`).

### Selección de anexo (`ListaDocumentos_Click`, `:183-254`)

1. Guarda el ID de la columna 0 (`:189`) y lee `URLAnexo` con `Dame` (`:199`; `FG.bas:290-318`).
2. Sin valor oculta el icono; con valor lo muestra y compone la ruta como directorio de planificaciones configurado más el nombre de fichero (`:206-217`).
3. Lee `EstadoPlanificacion` (`:227`) y habilita Eliminar solo si el resultado es «abierta» (`:233-237`).

### Adjuntar (`ComandoAnexar_Click`, `:10-69`)

1. Exige planificación y título (`:13-21`) y abre el selector de fichero con `Seleccionar` (`:27`; `FG.bas:1117-1145`).
2. Si el fichero no existe, sale sin avisar (`:34-36`).
3. `AnexarAPlanificacion`: valida ID y fichero, calcula el nombre como `<título>_<n.º de anexos + 1>.<extensión>` (`FG.bas:2445-2450`), rechaza un nombre ya registrado (`FG.bas:2456-2463`), copia el fichero al directorio de planificaciones sobrescribiendo (`FG.bas:2478`), obtiene el ID con `DameID` (`FG.bas:2484`) e inserta el registro con usuario de red y `Now()` (`FG.bas:2498-2507`).
4. Hace `Requery` de la lista (`:52`).

### Abrir anexo (`ImagenDocumentoAnexado_Click`, `:155-182`)

`AbrirEnLocal` comprueba que el fichero exista y no esté abierto, lo copia al directorio local configurado (borrando la copia previa) y lo abre con `Ejecutar` (`FG.bas:218-251`).

### Eliminar (`ComandoEliminarDocumento_Click`, `:71-119`; `EliminarAnexo`, `:255-348`)

1. Valida planificación y anexo seleccionado (`:77-84`), confirma con «No» por defecto (`:88-92`).
2. `EliminarAnexo` rechaza si la planificación está «Cerrada por reprogramación» (`:294-303`), lee el registro (`:306-309`), borra el fichero si existe (`:315-325`) y después el registro con `.Delete` (`:327`).
3. Refresca la lista, oculta el icono y deshabilita Eliminar (`:105-108`).

Errores: `MsgBox` crítico con el texto acumulado, sin correo; el error interno 1000 conserva el texto propio (`:60-68, 113-118, 176-181, 245-253`).

## Reglas, validaciones y cálculos

- **Título obligatorio:** «Se ha de poner un título al documento» (`:17-21`).
- **Nombre del anexo:** `GetBaseName(título)` + `_` + número + extensión del fichero original (`FG.bas:2450`); no se admite un nombre ya registrado para la planificación (`FG.bas:2456-2463`).
- **Estado de planificación:** «Abierta» si `FechaCierre` es fecha; «Trabajo Realizado» y «Reprogramación» dan los otros dos estados; cualquier otro motivo es error (`FG.bas:2073-2084`).
- **Permiso de baja:** el botón solo se habilita con planificación abierta (`:233-237`); `EliminarAnexo` solo bloquea la reprogramada (`:300-303`).
- **Persistencia de la ruta:** `URLAnexo` guarda solo el nombre de fichero (`FG.bas:2506`), no la ruta.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de anexos | `TbPlanificacionAnexos` | `RowSource` con SQL dinámica concatenada (ID numérico) | — | `:135-139` |
| `URLAnexo` del anexo | `TbPlanificacionAnexos` | `Dame` con SQL concatenada | no | `:199`; `FG.bas:290-318` |
| Estado de planificación | `TbPlanificacion` | DAO `OpenRecordset` con SQL concatenada | no | `FG.bas:2059-2062` |
| Nº de anexos | `TbPlanificacionAnexos` | `DameNumeroRegistrosPorSQL` sobre SQL concatenada | no | `FG.bas:2764-2766` |
| Nombre duplicado | `TbPlanificacionAnexos` | SQL con el nombre entre comillas | no | `FG.bas:2551-2555` |
| Siguiente ID | `TbPlanificacionAnexos` | `DameID`: recorre el campo y calcula máximo + 1 | no | `FG.bas:371-430` |
| Alta | `TbPlanificacionAnexos` | DAO `AddNew`/`Update` | no | `FG.bas:2498-2507` |
| Baja | `TbPlanificacionAnexos` | DAO `Delete` tras borrar el fichero | no | `:306-327` |
| Ficheros | directorios de planificaciones y local | `CopyFile`, `DeleteFile`, `Ejecutar` | no | `FG.bas:2478`; `:324`; `FG.bas:246` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: identificadores calculados en VBA (máximo + 1) fuera de transacción, SQL concatenada con comillas simples, `Now()` de cliente en `FechaAnexo`, y almacenamiento de anexos en un directorio de red configurado cuyo nombre se reconstruye al leer; en la web habrá que decidir dónde viven los ficheros.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Las propiedades `OnClick` de `ListaDocumentos`, `ComandoAnexar`, `ComandoEliminarDocumento` y `cmdSalir` no figuran en el `.form.txt` (solo `OnOpen` y `ImagenDocumentoAnexado`, `.form.txt:37, 665`): sus handlers del `.cls` (`:6, 10, 71, 183`) podrían no dispararse. Puede ser una pérdida de la exportación y no del binario.
- `Format("00", CStr(n))` tiene los argumentos invertidos (`FG.bas:2450`); el número del nombre podría no ser el esperado.
- El número de anexo sale del recuento de filas (`FG.bas:2445`): tras borrar un anexo intermedio, el siguiente nombre podría coincidir con uno existente y el alta fallaría con «ya está registrado» (`FG.bas:2463`).
- Al cancelar el selector, `Seleccionar` asigna `URLArchivoUltimo` con una cadena vacía (`FG.bas:1139-1141`).
- `ListaDocumentos_Click` valida `strIDPlanificacion` y no el anexo elegido (`:190`); el mensaje habla del anexo.
- `EstadoPlanificacion` devuelve «Abierta» y el formulario compara con «abierta» (`FG.bas:2074`; `:233`): depende de que `Option Compare Database` (`:2`) no distinga mayúsculas.
- La baja borra primero el fichero y luego el registro sin transacción (`:323-327`): un fallo en medio deja un registro sin fichero.
- `ListaDocumentos` está definida como «Value List» con 5 columnas y la SQL tiene 4 (`.form.txt:348, 357`; `:135-139`).
- `Eliminar` valida el estado con dos criterios distintos entre formulario y función (`:233`, `:300`).
- El mensaje de error de `AnexarAPlanificacion` dice «Eliminar» (`:48`), y `AbrirEnLocal` dice «El archivo debe estar abierto» cuando está abierto (`FG.bas:231-232`).
- La búsqueda de duplicado concatena el título entre comillas (`FG.bas:2551-2555`): un apóstrofo rompería la consulta.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: identificador generado por la base de datos, borrado atómico de fichero y registro, y nombre de anexo independiente del recuento.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir el formulario con un ID de planificación no numérico muestra «No se ha abierto el formulario con los datos necesarios» y lo cierra (`:131-134`).
2. La lista muestra solo los anexos de la planificación recibida, con título y fecha (`:135-139`).
3. Adjuntar sin título muestra «Se ha de poner un título al documento» y no copia nada (`:17-21`).
4. Adjuntar crea un registro con usuario, fecha y nombre, copia el fichero al directorio de planificaciones y refresca la lista (`FG.bas:2478, 2498-2507`; `:52`).
5. Con una planificación cerrada, el botón Eliminar queda deshabilitado (`:233-237`).
6. Eliminar con confirmación borra fichero y registro; «No» no borra (`:88-92`, `:323-327`).
7. Una planificación reprogramada no permite eliminar anexos (`:300-303`).
8. Pulsar el icono abre una copia local del anexo seleccionado (`FG.bas:245-246`).

## Pendientes

- Referencia a `establecerbotonera` en comentarios (`FG.bas:2044`) sin procedimiento correspondiente en este formulario.
- Si los `OnClick` ausentes del `.form.txt` existen en el binario (`.form.txt:37`).
- Contenido de `FicheroAbierto` (`FG.bas:195`), no leído.
- Si `IDAnexoPlanificacion` es autonumérico en la tabla de origen (ERD de origen).
