# Form_FormVideoDetalle

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormVideoDetalle.cls` (81 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ficha de consulta de un vídeo de ayuda: muestra título y descripción del vídeo activo y permite reproducirlo. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) que lee la global `m_ObjVideoActivo` (`Variables Globales.bas:102`); no escribe en la base de datos. Las citas son `Form_FormVideoDetalle.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormVideosArbolVideo` | Subformulario de diseño `FormDetalle` (`SourceObject = "Form.FormVideoDetalle"`) y reasignado en `MENU1_Click` | `Form_FormVideosArbolVideo.form.txt:416`, `Form_FormVideosArbolVideo.cls:25-27` |
| `Form_FormVideosUsuarios` | Llama a `Form_FormVideoDetalle.EstablecerDatos` tras seleccionar un nodo de vídeo | `Form_FormVideosUsuarios.cls:59` |

Abre el fichero del vídeo con la aplicación asociada de Windows mediante `Ejecutar` (alias de `ShellExecuteA`, `Funciones Genreales.bas:6`; uso en `:22`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | Muestra `Titulo` del vídeo activo; texto de diseño genérico | `:61, 65`, `.form.txt:326-327` |
| `lblExplicacion` | etiqueta | — | Muestra `Descripcion` o «Sin descripción registrada» | `:62, 66-70`, `.form.txt:352-353` |
| `ComandoVerVideo` | botón «Ver Video» | `OnClick` → `ComandoVerVideo_Click` | Comprueba vídeo y fichero y lo abre | `:6-36`, `.form.txt:376-378` |
| Formulario | — | `OnLoad` → `Form_Load` | Vacío salvo el manejador de errores | `:38-52`, `.form.txt:82` |

## Flujos

### Carga de datos (`EstablecerDatos`, `:54-81`)

1. Sin vídeo activo, restablece los textos genéricos de título y explicación (`:60-62`).
2. Con vídeo activo, copia `Titulo` a `lblTitulo` (`:65`) y `Descripcion` a `lblExplicacion`, o el texto «Sin descripción registrada» si está vacía (`:66-70`).

`Form_Load` no llama a `EstablecerDatos` (`:38-43`): al cargar, el formulario muestra los textos de diseño hasta que el contenedor lo invoca (`Form_FormVideosUsuarios.cls:59`).

### Ver vídeo (`ComandoVerVideo_Click`, `:6-36`)

1. Activa el reloj de arena (`:11`).
2. Sin vídeo activo, error interno 1000 con «Ningún video seleccionado» (`:14-17`).
3. Si el fichero de `m_ObjVideoActivo.url` no existe para `fso.FileExists`, error 1000 con «En estos momentos no es accesible el video» (`:18-21`).
4. Abre la ruta con el verbo `open` (`:22`) y desactiva el reloj de arena (`:24`).

La ruta se compone como directorio de vídeos de la aplicación del vídeo más `NombreArchivo` (`classes/Video.cls:237-238`), con el directorio derivado de `m_ObjEntorno.URLDirVideos` (`classes/Aplicacion.cls:1079-1095`).

### Errores

Los errores distintos de 1000 se envían con `CorreoAlAdministrador` (`:28-31`, `:45-48`); en todos los casos se muestra `MsgBox` crítico con `m_Error` (`:34`, `:51`). `EstablecerDatos` solo devuelve el texto en `p_Error` (`:77-80`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Sin descripción se muestra un texto fijo | `EstablecerDatos` | `:66-70` |
| Reproducir exige vídeo activo y fichero accesible | `ComandoVerVideo_Click` | `:14-21` |
| La ruta del vídeo es directorio de la aplicación + nombre de archivo | `Video.url` | `classes/Video.cls:216-240` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Ninguna propia | — | — | — | El vídeo llega ya cargado en `m_ObjVideoActivo` |
| Existencia del fichero | sistema de ficheros | `FileSystemObject` | — | `:18` |

## Coexistencia

Sin SQL propio. La reproducción depende de que el fichero sea accesible desde el puesto (`:18`); en la web, el vídeo deberá servirse por HTTP. Véase `coexistencia.md` (tanda posterior).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Me.lblExplicacion = "…"` asigna a una etiqueta sin propiedad `Caption` explícita (`:62`); una etiqueta no tiene valor por defecto y podría fallar. El error quedaría en `p_Error`, que el llamador no recoge (`Form_FormVideosUsuarios.cls:59`).
- `fso.FileExists` (`:18`) solo admite rutas de fichero: un vídeo publicado en una URL web se daría siempre por inaccesible.
- El resultado de `Ejecutar` no se comprueba (`:22`); un fallo al abrir no avisa.
- `Form_Load` captura errores con `m_Error` vacío (`:38-51`): mostraría un `MsgBox` sin texto.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: reproductor integrado con URL servida por la aplicación en lugar de abrir el fichero localmente.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al seleccionar un vídeo se muestran su título y su descripción, o «Sin descripción registrada».
2. «Ver vídeo» sin vídeo seleccionado muestra «Ningún video seleccionado».
3. «Ver vídeo» con el fichero no accesible muestra «En estos momentos no es accesible el video».

## Pendientes

- Si `Form_FormVideoDetalle` (`Form_FormVideosUsuarios.cls:59`) resuelve la instancia anidada como subformulario o crea una instancia oculta nueva; de ello depende que los textos lleguen a pantalla.
- Contenido real de `m_ObjEntorno.URLDirVideos` y formato de `NombreArchivo` (`classes/Aplicacion.cls:1090`).
