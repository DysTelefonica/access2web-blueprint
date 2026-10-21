# Form_FormUsuariosGestion

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormUsuariosGestion.cls` (194 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Búsqueda» de la gestión de usuarios: busca usuarios por nombre, fija el **usuario activo** de la sesión (global `m_ObjUsuarioActivo`, `Variables Globales.bas:100`) y permite restablecer su contraseña. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); la lista se rellena por código. Las citas son `Form_FormUsuariosGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormUsuarioGeneral` | `MENU1_Click` carga el formulario en el subformulario `FormDetalle` y escucha sus eventos con `WithEvents` | `Form_FormUsuarioGeneral.cls:3, 66-72` |
| `Form_FormMenuPrincipalAdmin` | `SUBMENU1_Click` («Usuario») carga `FormUsuarioGeneral` y llama a `Form_FormUsuarioGeneral.MENU1_Click` | `Form_FormMenuPrincipalAdmin.cls:221-237` |

Eventos hacia el contenedor: `UsuarioSeleccionado(usuario)` (`:4`, lanzado en `:156`) y `QuitarUsuarioSeleccionado` (`:5`, lanzado en `:98`). No abre otros formularios.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Nombre` | texto | `AfterUpdate` | Lanza `Filtrar` con el texto escrito | `:169-193`, `.form.txt:554-555` |
| `ComandoLimpiarNombre` | botón | `OnClick` | Vacía `Nombre` y vuelve a filtrar (lista completa) | `:10-33`, `.form.txt:371-374` |
| `ListaUsuarios` | lista, 3 columnas (ID oculto, correo, nombre), con cabecera | `OnClick` | Carga el usuario por ID y lo fija como activo | `:139-167`, `.form.txt:297-316` |
| `UsuarioSeleccionado` | texto bloqueado y deshabilitado | — | Muestra el nombre del usuario activo | `:83, 134, 155`, `.form.txt:600-614` |
| `ComandoQuitarUsuarioActivo` | botón | `OnClick` | Vacía el usuario activo y lanza `QuitarUsuarioSeleccionado` | `:95-100` |
| `ComandoResetearPass` (rótulo «Resetear Pass») | botón, deshabilitado en diseño | `OnClick` | Restablece la contraseña del usuario seleccionado | `:102-129`, `.form.txt:858-871` |

El `RowSource` de diseño de `ListaUsuarios` conserva filas residuales con datos personales (`.form.txt:312`); el código lo sustituye al cargar (`:132`).

## Flujos

### Apertura (`Form_Load`, `:131-137`)

1. Deja la lista solo con la fila de cabecera (`:132`).
2. Si ya hay usuario activo, muestra su nombre (`:133-135`). Deshabilita `ComandoResetearPass` (`:136`).

### Búsqueda (`Filtrar`, `:35-92`)

1. Deshabilita el restablecimiento, vacía el usuario activo y la caja `UsuarioSeleccionado` (`:51-53`).
2. `Constructor.getListaUsuariosPorNombre` lee `TbUsuariosAplicaciones` completa si el nombre está vacío, o filtra con `Like '*<nombre>*'` (`modules/Constructor.bas:841-847`).
3. Añade una fila por usuario; con `p_SoloActivos = Sí` omitiría los desactivados (`:64-77`). Ningún llamador pasa `Sí` (`:18, 177`): **se listan también los usuarios desactivados**.
4. Si queda exactamente un resultado (cabecera + 1, `:78`), lo selecciona y ejecuta `ListaUsuarios_Click` (`:79-83`).

### Selección (`ListaUsuarios_Click`, `:139-167`)

1. Con ID no numérico (fila de cabecera) vacía el seleccionado y el activo (`:144-149`).
2. `Constructor.getUsuario(ID)` relee el usuario (`:150`; `modules/Constructor.bas:25-92`), lo fija como activo, muestra su nombre, lanza `UsuarioSeleccionado` y habilita `ComandoResetearPass` (`:154-157`).

### Restablecer contraseña (`ComandoResetearPass_Click`, `:102-129`)

1. Sin seleccionado: «Seleccione un usuario de la lista» (`:105-108`). No pide confirmación.
2. `Usuario.ResetearPass` (`classes/Usuario.cls:729-782`): exige correo (`classes/Usuario.cls:736-739`), genera una clave aleatoria (`classes/Usuario.cls:740`), localiza la fila por `CorreoUsuario` (`classes/Usuario.cls:748-751`) y guarda su resumen criptográfico, marca `TieneQueCambiarLaContrasenia` y guarda la clave en claro en `PassIncialPlana` (`classes/Usuario.cls:757-765`). Después envía el correo de restablecimiento (`classes/Usuario.cls:770`; cuerpo en `classes/Usuario.cls:786-835`).
3. Éxito: «Se le ha mandado una clave de un solo uso al usuario por correo» (`:119`).

### Errores

Cada handler acumula el error en `m_Error`, envía `CorreoAlAdministrador` (salvo error interno 1000) y muestra `MsgBox` crítico (`:122-128, 159-166, 185-192`). `ComandoLimpiarNombre_Click` envía el correo incluso con el error 1000 (`:30`).

## Reglas, validaciones y cálculos

- La lista se indexa por `UsuarioRed`: si dos filas comparten usuario de red, solo aparece la primera (`modules/Constructor.bas:867-869`).
- El restablecimiento solo se habilita tras seleccionar un usuario en la lista (`:51, 99, 136, 157`).
- Permisos: el formulario no comprueba rol; solo es alcanzable desde el menú de administración, que se abre si el usuario conectado es administrador (`Variables Globales.bas:724-725`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Búsqueda por nombre | `TbUsuariosAplicaciones` | SQL dinámico concatenado (`Like`) | no | `modules/Constructor.bas:841-847` |
| Carga del usuario | `TbUsuariosAplicaciones` por `ID` | SQL dinámico | no | `modules/Constructor.bas:46-73` |
| Restablecer contraseña | `TbUsuariosAplicaciones` (`Password`, `TieneQueCambiarLaContrasenia`, `PassIncialPlana`) | DAO `Edit`/`Update` | no | `classes/Usuario.cls:748-765` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: comodín `*` de Jet en `Like` (en PostgreSQL, `%` o `ILIKE`); nombre concatenado sin escapar comillas (`modules/Constructor.bas:846`); el correo se envía tras el `Update` sin transacción, así que un fallo de envío deja la clave ya cambiada.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La clave de un solo uso se guarda en claro en `PassIncialPlana` y viaja en claro por correo (`classes/Usuario.cls:763, 810`).
- Un nombre con apóstrofo rompe la consulta de búsqueda (`modules/Constructor.bas:846`).
- Usuarios con `UsuarioRed` repetido o vacío se ocultan en la lista (`modules/Constructor.bas:867`).
- Si falla el correo, la contraseña ya se cambió y el usuario no la recibe (`classes/Usuario.cls:757-773`).
- `ComandoQuitarUsuarioActivo` no limpia `m_ObjUsuarioSeleccionado` ni la selección de la lista (`:95-100`); tras `Filtrar` con varios resultados, `m_ObjUsuarioSeleccionado` queda con el último iterado (`:75`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: pedir confirmación antes de restablecer; enlace de restablecimiento con caducidad en lugar de clave en claro; filtro «solo activos» visible.

## Criterios de aceptación

Propuestos, no aprobados:

1. Buscar por parte del nombre lista los usuarios coincidentes, incluidos los desactivados (regla actual).
2. Con un único resultado, queda seleccionado y se notifica al contenedor.
3. Restablecer sin usuario seleccionado no es posible; con usuario, obliga al cambio de contraseña en el siguiente acceso.
4. Quitar el usuario activo vacía la etiqueta del contenedor.

## Pendientes

- Contenido de `CorreoAlAdministrador` (`Funciones Genreales.bas:1447`) y de `GenerarPassValida` (`Funciones Genreales.bas:1070`).
- Si `Form_FormUsuarioGeneral.MENU1_Click` llamado desde el menú actúa sobre la instancia alojada en el subformulario (`Form_FormMenuPrincipalAdmin.cls:236`).
