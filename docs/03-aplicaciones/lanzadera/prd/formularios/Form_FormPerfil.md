# Form_FormPerfil

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormPerfil.cls` (220 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ficha de solo lectura con los datos del usuario conectado y sus permisos por aplicación; permite ir al cambio de contraseña. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) que lee la global `m_ObjUsuarioConectadoLogin`. Las citas son `Form_FormPerfil.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormMenuPrincipalUsuario` | Subformulario: `FormDetalle.SourceObject = "FormPerfil"` (`MENU3_Click`) | `Form_FormMenuPrincipalUsuario.cls:246-252` |

Abre `FormUsuarioCambioPass`, cerrándolo antes si ya estaba abierto (`:201-204`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Nombre`, `Matricula`, `Telefono`, `Movil`, `CorreoUsuario`, `UsuarioRed` | textos bloqueados (`Tag = DATO`) | — | Copia del usuario si no está vacío | `:21-41`, `.form.txt:343-344, 400-401, 457-458, 514-515, 571-572, 798-799` |
| `FechaAlta`, `FechaBaja` | textos bloqueados | — | Fecha con formato `dd/mm/yyyy` si es fecha | `:30-32, 50-52`, `.form.txt:627, 684` |
| `TextoTiempoParaExpirarPass` | texto bloqueado | — | Texto de días hasta la caducidad | `:42-44`; `Usuario.cls:841-879` |
| `MantenerLanzaderaAbierta` | texto bloqueado | — | «Sí» o «No» | `:45-49` |
| `Responsable` | texto bloqueado | — | Nombre del jefe del usuario | `:53-59` |
| `ID` | texto oculto | — | Identificador del usuario | `:61`, `.form.txt:816, 830` |
| `ListaUsuarios` | lista de 8 columnas (1.ª oculta) | — | Permisos por aplicación | `:87-165`, `.form.txt:848-870` |
| `ComandoCambiarPass` («Cambiar Pass») | botón | `OnClick` | Abre `FormUsuarioCambioPass` | `:195-219`, `.form.txt:1109-1111` |

## Flujos

### Apertura (`Form_Open` → `EstablecerDatos`, `:166-194`, `:8-86`)

1. Sin usuario conectado, error «No se sabe quién ha hecho login» (`:15-18`).
2. Copia los datos del usuario a los controles (`:19-63`). El responsable se busca por `JefeDelUsuario` como usuario de red (`Usuario.cls:95-98`).
3. Activa la edición del formulario, pero habilita y bloquea los textos con `Tag = DATO` (`:64-72`).
4. `EstablecerLista` vacía la lista dejando la cabecera (`:107`), lee `ColAplicacionesPermisos` (`:113`; consulta en `Constructor.bas:1019-1022`) y añade una fila por aplicación con «Sí»/«No» en Admin, Calidad, Economía, Técnico, Secretaría y Calidad avisos (`:121-159`).

### Errores

`EstablecerDatos` y `EstablecerLista` devuelven el texto en `p_Error` (`:81-85, 161-164`). `Form_Open` envía los errores no internos con `CorreoAlAdministrador` y muestra `MsgBox` crítico, sin cerrar el formulario (`:185-193`).

## Reglas, validaciones y cálculos

- **Texto de caducidad** (`Usuario.cls:841-879`): vacío si la cuenta no está activada; «Su Pass nunca caduca» si `PasswordNuncaCaduca`; si no, días hasta `FechaProximaPassCalculado` con mensajes para 1 día, hoy, futuro y pasado.
- **Permisos:** cada columna sale de una propiedad calculada de `UsuarioAplicacionPermisos` (`:124-153`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Permisos del usuario | `TbUsuariosAplicacionesPermisos` | SQL dinámico concatenado | `Constructor.bas:1019-1022` |
| Responsable | `TbUsuariosAplicaciones` por `UsuarioRed` | SQL dinámico | `Usuario.cls:98`; `Constructor.bas:53-73` |

Sin escrituras.

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: correo concatenado en el `WHERE` de permisos (`Constructor.bas:1021`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La lista trae en diseño filas de ejemplo con aplicaciones y permisos fijos (`.form.txt:865-869`); si `EstablecerDatos` falla antes de `EstablecerLista` (`:15-18, 53-57`), se verían esos datos y el formulario sigue abierto (`:185-193`).
- Si el usuario no tiene permisos (`ColAplicacionesPermisos` vacía), la lista queda solo con la cabecera, sin mensaje (`:118-120`).
- Solo `FormMenuPrincipalUsuario` lo muestra (`Form_FormMenuPrincipalUsuario.cls:249-250`); los menús de administrador y calidad no lo cargan en `src`.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: vista «Mi perfil» común a todos los roles; tabla de permisos generada desde el catálogo de roles en vez de seis columnas fijas.

## Criterios de aceptación

Propuestos, no aprobados:

1. La ficha muestra los datos del usuario conectado y no permite editarlos.
2. La tabla de permisos tiene una fila por aplicación con los seis roles en «Sí»/«No».
3. El texto de caducidad coincide con las reglas de días restantes.
4. «Cambiar Pass» abre el cambio de contraseña.

## Pendientes

- Contenido de `getAplicacionesPermisos` más allá de la consulta y de las propiedades calculadas de `UsuarioAplicacionPermisos`.
- Si otros formularios cargan `FormPerfil` por nombre en una variable.
