# Form_frmJustificacion

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_frmJustificacion.cls` (139 líneas) y `.form.txt` (26 bloques `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`, `Factoria.bas` = `modules/Factoria.bas`; las clases de `src/classes/` y los repositorios de `src/modules/` se citan por el nombre del fichero. Las citas son `Form_frmJustificacion.cls:línea` salvo indicación.

Alta y edición de una justificación de HPS. Es un formulario **no enlazado** que usa `OpenArgs` como único contrato: sin argumento significa alta, con un `Long` significa edición (`:19-23`). Escribe a través de una arquitectura de tres capas —`JustificacionServicio` (validación y caché), `JustificacionRepositorio` (SQL) y la entidad `Justificacion`— y avisa al buscador padre con el evento `JustificacionGuardada` (`:7, 126`). No implementa borrado: la baja es lógica y vive en el buscador.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_frmBuscarJustificacion` (`btnAlta_Click`) | `DoCmd.OpenForm "frmJustificacion"` sin argumentos; el padre se enlaza con `WithEvents` para recibir `JustificacionGuardada` | `Form_frmBuscarJustificacion.cls:148-151` |
| `Form_frmBuscarJustificacion` (`btnEditar_Click`) | `DoCmd.OpenForm "frmJustificacion", , , , , , idSel` con el identificador en `OpenArgs` | `Form_frmBuscarJustificacion.cls:165-168` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `JustificacionServicio` | `ObtenerPorID` en carga, `VerificarUso` para el bloqueo, `GuardarJustificacion` y `ObtenerUltimoIDCreado` al guardar | `:37, 51, 104, 115` |
| `Entorno` (implícito) | `GuardarJustificacion` invalida la caché global de justificaciones | `JustificacionServicio.cls:57-59` |

El formulario se cierra siempre a sí mismo al terminar: `cmdGuardar_Click` (`:129`), `cmdCancelar_Click` (`:138`) y `cmdSalir_Click` (`:14`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` → `Form_Load` | Decide alta o edición por `OpenArgs`, carga y aplica las reglas | `:17-27` |
| `lblTitulo` | etiqueta de encabezado | — | «Nueva Justificación» / «Editar Justificación» | `:34, 38`, `.form.txt:505-506` |
| `txtTitulo` | texto | — | Título; se bloquea si la justificación está en uso | `:42, 53-54`, `.form.txt:538` |
| `txtDescripcion` | texto multilínea | — | Descripción; mismo bloqueo que el título | `:43, 56-57`, `.form.txt:594` |
| `chkActiva` | casilla | — | Estado activo; **nunca** se bloquea | `:44, 98`, `.form.txt:671` |
| `lblAviso` | etiqueta | — | Aviso de «en uso»; visible sólo en edición bloqueada | `:60, 65, 68`, `.form.txt:696-697` |
| `cmdGuardar` | botón | `OnClick` → `cmdGuardar_Click` | Valida cambios, guarda, anuncia el id y cierra | `:72-135`, `.form.txt:736-738` |
| `cmdCancelar` | botón | `OnClick` → `cmdCancelar_Click` | Cierra sin guardar | `:137-139`, `.form.txt:775-777` |
| `Etiqueta232`, `Etiqueta256`, `Etiqueta45` | etiquetas | — | Rótulos «Título», «DESC.» y `¿Activa?`; nombres internos de plantilla | `.form.txt:568, 624, 644` |
| Sin nombre | botón, `CheckBox`, `OptionGroup`, `TextBox`, `ListBox`, `ComboBox`, `Subform`, `CustomControl`, `Image`, `Rectangle` | — | Restos de plantilla en el cuerpo del formulario; el `.cls` no los menciona | `.form.txt:411-486` |

## Flujos

### Normal — alta

1. `Form_Load` ve `OpenArgs` nulo y fija `m_IdActual = 0` (`:19-23`).
2. `CargarDatos` construye la entidad con `Factoria.CreateEntity("Justificacion")` (`:32`; `Factoria.bas:9-10`), la marca activa (`:33`) y rotula «Nueva Justificación» (`:34`).
3. `AplicarReglasNegocio` no hace nada en alta: oculta el aviso (`:67-69`).
4. `cmdGuardar_Click` compara los tres controles contra la entidad (`:81-83`); si no hay diferencias, avisa «No se han detectado cambios para guardar.» y **no cierra** (`:86-89`, con el cierre comentado en `:88`).
5. Vuelca los datos al objeto (`:96-98`) y llama a `m_Servicio.GuardarJustificacion` (`:104`), que valida objeto no nulo, título no vacío y longitud ≤ 255 (`JustificacionServicio.cls:37-49`) y luego delega en el repositorio (`:53`).
6. Recupera el id: en alta llama a `ObtenerUltimoIDCreado` (`:115`), en edición conserva `m_IdActual` (`:119`).
7. Lanza `JustificacionGuardada(idFinal, m_Entidad.Titulo)` (`:126`) y se cierra (`:129`).

### Normal — edición

1. `Form_Load` convierte `OpenArgs` con `CLng` (`:20`); `CargarDatos` pide la entidad al servicio y rotula «Editar Justificación» (`:37-38`).
2. `AplicarReglasNegocio` consulta `VerificarUso` (`:51`, que cuenta filas en `TbSolicitudes`, `JustificacionRepositorio.bas:69-82`). Si está en uso, bloquea título y descripción, los pinta en gris, muestra el aviso y lleva el foco a `chkActiva` (`:53-63`).
3. El guardado sigue el mismo camino; el id es el ya cargado (`:119`), así que en teoría no depende de la consulta posterior.

### Alternativos

- **Sin cambios que guardar:** el mensaje informativo y el `Exit Sub` de `:86-89` dejan el formulario abierto con los datos en pantalla.
- **Justificación en uso:** la interfaz ofrece deliberadamente sólo el cambio de `chkActiva` (`:62-63`); la regla de negocio es «no se puede cambiar ni borrar si ya se ha usado» (comentario en `:48`).
- **Baja de la justificación:** no está en este formulario. La ejecuta el buscador, que llama a `EliminarJustificacion` y por tanto a un `UPDATE … SET activa = False` (`JustificacionRepositorio.bas:62-67`).

### Errores

- `Manejador` (`:133-134`) muestra «Error al guardar: » con `Err.Description` y no cierra niRepite: el formulario queda abierto con los datos ya volcados en el objeto `m_Entidad` pero no necesariamente en la base de datos.
- `Form_Load` (`:17-27`), `CargarDatos` (`:29-45`) y `AplicarReglasNegocio` (`:47-70`) **no tienen manejador de error**: cualquier excepción en la apertura sale por la vía estándar de Access.

## Reglas, validaciones y cálculos

- **Título obligatorio y ≤ 255 caracteres**, validado en la capa de servicio, no en el formulario (`JustificacionServicio.cls:42-49`; el límite coincide con el `TEXT(255)` del ERD de origen).
- **Detección de cambios por comparación explícita** de los tres campos (`:81-83`), con `Nz` para neutralizar nulos; el formulario no marca «sucio» de ninguna otra forma.
- **Bloqueo por uso:** `VerificarUso` → `EstaSiendoUsada` cuenta `TbSolicitudes` con `idjustificacion = id` (`JustificacionRepositorio.bas:75`). El comentario del repositorio admite que el nombre de tabla y clave se supusieron (`:73-74`); contra el ERD de origen ambos nombres existen (`schema.sql:72-76` para `TbJustificaciones` y la columna `idjustificacion` de `TbSolicitudes`), pero `src` no declara ninguna clave foránea ni primaria, así que la integridad real no es verificable en el código.
- **El estado activo no es un filtro de lectura:** `ObtenerTodas` trae activas e inactivas sin `WHERE` (`JustificacionRepositorio.bas:12-14`) y es el buscador quien discrimina.
- **Alta con id calculado, no autonumérico:** el repositorio asigna `ObtenerUltimoID() + 1` (`:36`); el formulario está escrito sobre esa premisa (comentario en `:113-114`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Cargar la entidad | `SELECT * FROM TbJustificaciones WHERE idjustificacion = ` & `id` | DAO, SQL concatenado con un `Long` | `JustificacionRepositorio.bas:21-23` |
| Comprobar uso | `SELECT Count(*) … FROM TbSolicitudes WHERE idjustificacion = ` & `id` | DAO, SQL concatenado | `JustificacionRepositorio.bas:75-77` |
| Reservar id de alta | `SELECT Max(idjustificacion) FROM TbJustificaciones` | DAO, sobre el back-end con `getdb()` | `JustificacionRepositorio.bas:90-93` |
| Guardar | `INSERT INTO TbJustificaciones (IdJustificacion, titulo, descripcion, activa) VALUES ([pId], …)` | `QueryDef` temporal **parametrizada** (`dbFailOnError`) | `JustificacionRepositorio.bas:41-51` |
| Invalidar caché | (ninguna) | Derivado: `Entorno.RecargarJustificaciones` pone la colección a `Nothing` | `JustificacionServicio.cls:57-59`, `Entorno.cls:132-134` |
| Releer tras guardar | `SELECT Max(idjustificacion)` (segunda consulta) | DAO | `:115`, `JustificacionRepositorio.bas:86-104` |

La escritura es la única parte parametrizada de todo el camino; las lecturas y el `UPDATE` de baja concatenan el identificador, que al ser un `Long` controls la inyección.

## Coexistencia

Pendiente de `../coexistencia.md`. Lo que este formulario aporta al plan: la arquitectura de tres capas (servicio → repositorio → entidad con propiedades) es el patrón más limpio del legacy y el candidato natural a equivalente web, donde `JustificacionRepositorio.Guardar` se traduce en un `INSERT` parametrizado y `JustificarServicio.GuardarJustificacion` en una operación de API. Los tres puntos que no sobreviven tal cual son el `Max(id)+1` como generador de claves (debe ser `IDENTITY`/`GENERATED`), la segunda lectura de `Max` para recuperar el id (debería devolverlo el propio `INSERT`) y la validación de duplicados por código.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-34.** El repositorio sólo sabe insertar: `Guardar` ejecuta siempre `INSERT INTO TbJustificaciones` (`JustificacionRepositorio.bas:41-51`) y sólo calcula el identificador cuando viene a cero (`:35-38`). Editar una justificación existente, por tanto, reenvía un `INSERT` con la clave primaria ya cargada. Si `idjustificacion` es clave primaria o tiene índice único, el guardado falla; si no lo es, se crea una fila duplicada en lugar de actualizar. `src` no declara ninguna de las dos cosas, así que el desenlace exacto queda como hipótesis (`schema.sql:72-76` no incluye restricciones).
- **H-HPS-S-35.** El error de persistencia se pierde: `JustificacionServicio.GuardarJustificacion` es un `Sub` que invoca `JustificacionRepositorio.Guardar obj` sin pasar ni leer `p_Error` (`JustificacionServicio.cls:53`), mientras el repositorio comunica el fallo sólo por su valor de retorno y por el `p_Error` opcional (`JustificacionRepositorio.bas:28, 54-59`). Un `INSERT` fallido devuelve `False` al vacío y el formulario continúa: lanza `JustificacionGuardada`, anuncia un identificador y se cierra como si se hubiera guardado.
- **H-HPS-S-36.** El identificador del alta se recupera con una **segunda** consulta de `Max` posterior al `INSERT` (`:111-120` → `JustificacionRepositorio.bas:86-104`), y el propio alta lo reserva con `Max(id) + 1` (`:36`). Entre ambas lecturas otra inserción puede desplazar el `Max`, de modo que el evento puede anunciar un identificador distinto del creado; y dos altas simultáneas sobre el mismo `Max` colisionan.
- **H-HPS-S-37.** `Form_Load` no tiene manejador de error: `CLng(Me.OpenArgs)` (`:20`) falla si el argumento no es numérico y `m_Entidad.Titulo` (`:42-44`) falla si `ObtenerPorID` devuelve `Nothing` —lo hace con `id <= 0` (`JustificacionServicio.cls:28`) y también cuando la consulta no devuelve filas (`RepositorioComun.bas:44-48`)—, con lo que abrir el formulario con un id inexistente produce un error no controlado en lugar de un aviso.
- Rótulos y plantilla heredados: el `Caption` de `Etiqueta45` es `Caption ="\"¿Activa?"` (`.form.txt:645`), con una comilla doble literal dentro del texto; las etiquetas `Etiqueta232` y `Etiqueta256` (`.form.txt:568, 624`) conservan nombres internos; y el cuerpo del formulario conserva `Button`, `CheckBox`, `OptionGroup`, `TextBox`, `ListBox`, `ComboBox`, `Subform` y `CustomControl` sin nombre ni handler (`.form.txt:411-486`).
- `chkActiva.SetFocus` (`:63`) se invoca desde `Form_Load`, antes de que el formulario tenga el foco, y sin comprobar `TabStop` ni `Visible`.
- El cierre de `:88` está comentado: pulsar «Guardar» sin cambios deja la ventana abierta, a diferencia de «Cancelar» y «Salir».

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: alta con clave delegada en la base y el identificador devuelto por la propia escritura; edición y baja en la misma ruta que la alta con bloqueo explícito cuando la justificación esté referenciada; y el evento de guardado sustituido por la actualización de la colección en la respuesta, sin ventana modal intermedia.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un alta sin título no escribe nada y señala el campo obligatorio.
2. Un título de más de 255 caracteres no se guarda y el aviso indica el límite.
3. Un alta correcta crea una sola fila, con la activa por defecto, y el buscador la ve en la lista sin recargar la aplicación.
4. Editar una justificación existente modifica su fila y no crea una nueva.
5. Una justificación con solicitudes asociadas no permite cambiar título ni descripción, pero sí activar o desactivar.
6. Una justificación en uso no se elimina; una que no está en uso queda desactivada, no borrada.
7. Guardar sin cambios no escribe en la base de datos.
8. Si la escritura falla, se informa del motivo y el formulario permanece abierto con los datos introducidos.

## Pendientes

- Si `idjustificacion` es clave primaria o tiene índice único en `TbJustificaciones`: no es verificable en `src`; el ERD de origen exporta las cuatro columnas sin restricciones (`schema.sql:72-76`). De ello depende el desenlace de H-HPS-S-34.
- Si existe algún lanzador de `frmJustificacion` fuera de `Form_frmBuscarJustificacion`: en `src` sólo aparece en `:148` y `:165` de ese formulario.
- Tratamiento de la clave foránea `TbSolicitudes.idjustificacion` (`schema.sql:139`): no hay ninguna relación declarada en `src` ni en el ERD, sólo la cuenta de `EstaSiendoUsada` (`JustificacionRepositorio.bas:75`).
- Longitud real de `Descripcion`: el tipo es `MEMO` en el ERD, sin límite declarado; el truncado a 100 caracteres de `getJustificacionParaCampo` (`JustificacionServicio.cls:106-110`) sugiere un límite de campo que no aparece en el código.
- `JustificacionServicio.cls:88` deja un comentario de cabecera de clase en mitad del fichero, justo antes de `ObtenerUltimoIDCreado` (`:90`): sugiere un archivo ensamblado a mano y no exportado desde un único origen.