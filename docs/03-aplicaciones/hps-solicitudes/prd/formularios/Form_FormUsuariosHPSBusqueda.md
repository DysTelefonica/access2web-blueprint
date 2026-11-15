# Form_FormUsuariosHPSBusqueda

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormUsuariosHPSBusqueda.cls` (295 líneas) y `.form.txt` (23 bloques `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`, `constructor.bas` = `modules/constructor.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_FormUsuariosHPSBusqueda.cls:línea` salvo indicación.

Buscador de usuarios HPS (solicitantes) que, en vez de limitarse a devolver un nombre, **dirige el alta de una solicitud**: al seleccionar, establece el usuario HPS activo, descarta la solicitud en curso y abre `FormSolicitudRenovacion`. Es **no enlazado**: la lista se puebla desde un diccionario en memoria y el filtro se evalúa en VBA.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Menú de entidades de `Form0BDOpciones` | `Entorno.EntidadesFormularios` lo asocia a un tipo de entidad; se abre con `OpenArgs` = tipo de trámite | `classes/Entorno.cls:240-252` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `FormSolicitudRenovacion` | `ComandoSeleccionar_Click` cierra la instancia previa con `acSaveNo`, borra `m_ObjSolicitudActiva` y `m_ObjSolicitudAlInicio`, y abre el formulario con `m_Tipo` | `:150-156` |
| Suscriptores de `Seleccionado` | `Public Event Seleccionado(p_Usuario As UsuarioHPS)` | `:7, 147` |
| `Entorno` | `m_ObjEntorno.colUsuariosHPS` como origen de la lista; `Set … = Nothing` para forzar recarga | `:46, 93` |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnOpen` → `Form_Open(Cancel)` | Valida `OpenArgs`, filtra y ajusta el rótulo | `:180-221` |
| `PalabraClave` | texto | `OnBeforeUpdate` → `PalabraClave_BeforeUpdate` | Refiltra al salir del campo | `:293-295` |
| `ListaFiltrados` | lista de 3 columnas | `OnClick` → `ListaFiltrados_Click` | Resuelve el usuario seleccionado desde la base y habilita el botón | `:223-290`, `.form.txt:1005-1011` |
| `ListaFiltrados` | lista | `OnDblClick` → `ListaFiltrados_DblClick` | Equivalent a pulsar «Seleccionar» | `:170-178` |
| `ComandoSeleccionar` | botón | `OnClick` → `ComandoSeleccionar_Click` | Fija el usuario activo y lanza el alta de la solicitud | `:136-168`, `.form.txt:1063-1066` |
| `ComandoActualizar` | botón | `OnClick` → `ComandoActualizar_Click` | Limpia la palabra clave e invalida la caché de usuarios | `:85-106`, `.form.txt:903-906` |
| `ComandoAyuda` | botón | `OnClick` → `ComandoAyuda_Click` | Abre el PDF de ayuda | `:109-134`, `.form.txt:810-813` |
| `cmdSalir` | botón | `OnClick` → `cmdSalir_Click` | Cierra el formulario | `:13-16` |
| `lblTitulo` | etiqueta de encabezado | — | «USUARIOS HPS»; `Form_Open` copia su rótulo al del formulario | `:193`, `.form.txt:791-792` |
| `lblListaDeResultados` | etiqueta | — | «Registros filtrados (N)», recounted en cada filtro | `:44, 69`, `.form.txt:1041-1042` |

## Flujos

### Normal — carga

1. `Form_Open` valida que `OpenArgs` no esté vacío (`:194-198`) y guarda el valor en `m_Tipo`.
2. `Filtrar m_Error` (`:199`) recorre `m_ObjEntorno.colUsuariosHPS` (`:46`), que es la caché que `Entorno` rellena con `constructor.getUsuariosHPS` la primera vez (`Entorno.cls:245-252`).
3. `Filtrar` limpia la selección anterior (`:34-40`), fija la cabecera `"ID;DNI;Nombre"` (`:43`), añade una fila por usuario con `id & ";" & .DNI & ";" & nombre completo` (`:60-68`) y actualiza el contador (`:69`).
4. Si hay exactamente un usuario, selecciona la primera fila y dispara `ListaFiltrados_Click` (`:70-74`).

### Normal — búsqueda y selección

1. `PalabraClave_BeforeUpdate` llama a `Filtrar` sin argumentos (`:293-295`), que al no recibir palabra clave usa la caché completa (`:45-52`); con palabra clave delega en `constructor.getUsuariosHPSPorPalabraClave` (`constructor.bas:3633`).
2. `ListaFiltrados_Click` deshabilita el botón, lee el identificador de la fila (`:239`) y resuelve el `UsuarioHPS` completo con `constructor.getUsuarioHPS` (`constructor.bas:3581`), reutilizando el objeto ya resuelto si el identificador no ha cambiado (`:248-261`). Con usuario resuelto, habilita «Seleccionar» (`:271`).
3. `ComandoSeleccionar_Click` exige selección (`:143-146`), lanza `Seleccionado` (`:147`), asigna `m_ObjUsuarioHPSActivo` (`:149`), **cierra `FormSolicitudRenovacion` con `acSaveNo`** si estaba abierta (`:150-152`), pone a `Nothing` la solicitud activa y su fotografía (`:153-154`), abre `FormSolicitudRenovacion` con `m_Tipo` (`:155`) y se cierra (`:156`).

### Alternativos

- **Búsqueda con un único resultado:** `Filtrar` pre-selecciona la fila y dispara el click, de modo que basta «Seleccionar» (`:70-74`).
- **Recarga de la caché:** `ComandoActualizar_Click` borra la palabra clave y pone `m_ObjEntorno.colUsuariosHPS` a `Nothing` (`:92-93`), con lo que el siguiente `Filtrar` vuelve a pedir los usuarios a `constructor`.
- **Fila sin identificador:** `ListaFiltrados_Click` limpia `m_UsuarioSeleccionado` y sale sin habilitar el botón (`:239-247`).

### Errores

- Cuatro manejadores (`ComandoActualizar_Click`, `ComandoAyuda_Click`, `ComandoSeleccionar_Click`, `ListaFiltrados_Click`) distinguen el centinela `Err.Number = 1000` y, para los errores reales, envían correo al administrador (`FU.bas:585`) y muestran un `MsgBox` crítico (`:105, 133, 167, 288`).
- `Form_Open` añade el cierre del formulario a su manejador (`:220`), de modo que un error de apertura no deja una ventana en blanco.

## Reglas, validaciones y cálculos

- **Origen de los datos:** sin palabra clave, la lista sale de la caché global (`Entorno.colUsuariosHPS`), que se cargó una sola vez; con palabra clave, la consulta va directa a `constructor.getUsuariosHPSPorPalabraClave` (`:45-52`). Los dos caminos pueden devolver conjuntos distintos.
- **`Filtrar` es `Public`** (`:18`) y reinicia `m_UsuarioSeleccionado` a `Nothing` en cada llamada (`:32`), lo que obliga a volver a resolver el usuario desde la base.
- **Selección obligatoria:** sin `m_UsuarioSeleccionado` el botón no se habilita (`:271`) y `ComandoSeleccionar_Click` lo comprueba otra vez (`:143-146`).
- **La fila lleva el DNI** como segunda columna (`:64`), que es un dato personal visible en la lista y que no se filtra ni se oculta.
- **Contador de resultados:** `lst.ListCount - 1` descuenta la cabecera (`:69`); el valor de diseño de la etiqueta dice «(1)» (`.form.txt:1042`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Listar usuarios HPS | `constructor.getUsuariosHPS`, cacheada en `Entorno` | DAO, no acotado en este fichero | `:46`, `Entorno.cls:245-252` |
| Buscar por palabra clave | `constructor.getUsuariosHPSPorPalabraClave` | DAO | `:48`, `constructor.bas:3633` |
| Resolver un usuario | `constructor.getUsuarioHPS` | DAO | `:249, 255`, `constructor.bas:3581` |
| Estado de la solicitud en curso | variables de sesión, sin DAO | Derivado | `:153-154` |

Este formulario no escribe en ninguna tabla: sólo reposiciona el estado de la sesión y lanza el alta en otro formulario.

## Coexistencia

Pendiente de `../coexistencia.md`. Lo que aporta: la búsqueda de solicitante y el alta de solicitud están acopladas en el mismo botón, lo que en web se traduce en un flujo de dos pasos (buscar solicitante → iniciar solicitud) con el identificador del usuario en la ruta. La lista completa cacheada en memoria y el filtro en VBA (`Entorno.colUsuariosHPS`) son incompatibles con un volumen grande de usuarios: ahí la búsqueda debe ir a la base con paginación.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-46.** `ComandoSeleccionar_Click` cierra `FormSolicitudRenovacion` con `acSaveNo` y la vuelve a abrir (`:150-155`), además de poner a `Nothing` la solicitud activa y su fotografía (`:153-154`): si el usuario tenía una renovación a medias, esa ventana se cierra sin aviso y su trabajo se pierde. El doble clic de la lista (`:170-178`) dispara el mismo camino, de modo que un doble clic accidental también lo provoca.
- **H-HPS-S-47.** Las filas se arman concatenando `id & ";" & .DNI & ";" & nombre` (`:43, 64`): un punto y coma en cualquiera de los tres campos desplaza las columnas siguientes, y la búsqueda y la selección leen la fila por posición (`ListaFiltrados.Column(0)`, `:239`). El nombre completo se compone de tres campos (`Apellido_1`, `Apellido_2`), lo que amplía la superficie.
- **H-HPS-S-48.** `cmdSalir_Click` (`:14`) y `ListaFiltrados_DblClick` (`:171`) arrancan con `On Error Resume Next` y no avisan: cualquier fallo —incluida la propia llamada a `ComandoSeleccionar_Click` del doble clic— se traga sin correo al administrador ni mensaje, a diferencia de los otros cinco handlers del formulario.
- **H-HPS-S-49.** `ComandoAyuda` y `ComandoActualizar` comparten el rótulo de diseño `Caption ="Comando46"` (`.form.txt:811, 904`): ninguno de los dos botones tiene texto visible, sólo el nombre interno.
- `Me.Caption` se escribe dos veces seguidas en `Form_Open`, primero con `m_TituloFormulario` (`:190`) y acto seguido con el rótulo de la etiqueta (`:193`); la primera asignación no sirve para nada.
- `m_UsuarioSeleccionado` es `Public` (`:5`) pero ningún otro fichero de `src` lo lee o escribe: la superficie pública es innecesaria y deja el estado de selección fuera del control del formulario.
- El `Caption` de diseño es «GESTIÓN DE LAS SOLICITUDES» (`.form.txt:38`), que describe otro formulario; `Form_Open` lo sobrescribe con el texto de `lblTitulo` (`:193`).
- `ListaFiltrados_Click` deshabilita `ComandoSeleccionar` al entrar (`:235`) y sólo lo rehabilita al finalizARS con éxito (`:271`): cada error deja el botón deshabilitado hasta el siguiente click sobre la lista.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: búsqueda de solicitante en servidor con paginación y referencia a la DNI explícita en el listado; separar la búsqueda del alta de solicitud, con el usuario elegido como paso intermedio; y no exponer el DNI en la lista salvo que el perfil lo permita.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin tipo de trámite en `OpenArgs` el formulario no se abre y el motivo se comunica.
2. Sin palabra clave se listan todos los usuarios HPS; con palabra clave, sólo los que coinciden, y el contador refleja el número real.
3. Sin fila elegida, «Seleccionar» no hace nada y no hay mensaje de error.
4. Elegir un usuario fija el solicitante de la solicitud que se abre a continuación.
5. Elegir un usuario con una solicitud en curso avisa antes de descartarla, y no se pierde trabajo sin confirmación.
6. Un doble clic sobre una fila no descarta una solicitud en curso sin confirmación.
7. «Actualizar» vacía la palabra clave y vuelve a pedir los usuarios a la base de datos.
8. Ante un fallo en la resolución del usuario, el aviso contiene el motivo y el botón «Seleccionar» vuelve a quedar disponible.

## Pendientes

- Si algún otro formulario de `src` consume `Public Event Seleccionado`: la búsqueda de suscriptores no devuelve ninguno fuera de este fichero, de modo que el evento se lanza sin destinatario conocido.
- Volumen real de `TbUsuariosHPS`: no consta en el ERD de origen consultado y condiciona si el filtrado en memoria es viable.
- Qué plantilla usa `AbrirAyuda` (`FU.bas`) para este formulario: el rótulo del botón no lo identifica (H-HPS-S-49).
- Recuento de controles: el inventario fija 26 y el `.form.txt` contiene 23 bloques `Begin <Tipo>`; la diferencia no se reproduce con una regla simple de recuento.