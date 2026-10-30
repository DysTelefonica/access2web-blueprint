# Form_FormUsuarioPerfilAplicaciones

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormUsuarioPerfilAplicaciones.cls` (385 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Perfil Aplicaciones» del usuario activo: lista las aplicaciones a las que tiene permiso con sus seis perfiles (administrador, calidad, economía, técnico, secretaría, avisos de calidad), y permite dar de alta, sobrescribir o eliminar los permisos de una aplicación en `TbUsuariosAplicacionesPermisos`. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`). Las citas son `Form_FormUsuarioPerfilAplicaciones.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormUsuarioGeneral` | `MENU3_Click` lo carga en `FormDetalle` si hay usuario activo | `Form_FormUsuarioGeneral.cls:86-94` |
| `Funciones Genreales.bas` | `UsuarioAltaEdicionEstado` y `UsuarioSeleccionadoEstado` lo asignan a `CuadroInferior` de `Form00Inicial`; no tienen llamadores y `Form00Inicial` no existe en `src` | `Funciones Genreales.bas:1680, 1702, 1745, 1781` |

No abre otros formularios.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ComboAplicacion` | combo 2 columnas (ID oculto) | — | Aplicaciones con icono en la lanzadera más una fija | `:297-331`, `.form.txt:349-370` |
| `EsUsuarioAdministrador`, `EsUsuarioCalidad`, `EsUsuarioEconomia`, `EsUsuarioSecretaria`, `EsUsuarioTecnico`, `EsUsuarioCalidadAvisos` | casillas | — | Perfiles que se asignarán con «Alta» | `:36-65`, `.form.txt:562, 605, 648, 691, 734, 777` |
| `ComandoAltaPerfil` («Alta») | botón | `OnClick` | Crea o sobrescribe los permisos de la aplicación del combo | `:8-113`, `.form.txt:396-398` |
| `ListaUsuarios` | lista 8 columnas (ID oculto) con cabecera | `OnClick` | Selecciona la aplicación para eliminar | `:333-386`, `.form.txt:436-460` |
| `ComandoEliminar` («Eliminar») | botón, deshabilitado en diseño | `OnClick` | Borra los permisos de la aplicación seleccionada, con confirmación | `:115-171`, `.form.txt:480-491` |

Los `RowSource` de diseño de `ComboAplicacion` y `ListaUsuarios` conservan filas residuales de datos (`.form.txt:367, 452`); el código los sustituye al abrir (`:239, 306`).

## Flujos

### Apertura (`Form_Open`, `:173-217`)

1. Sin usuario activo: «No hay un usuario activo» (`:181-184`).
2. `AllowEdits` según si el usuario **conectado** es administrador (`:185-189`).
3. `EstablecerCombo` y `EstablecerLista` (`:190-197`). `ComandoEliminar` queda deshabilitado; `ComandoAltaPerfil` sigue a `AllowEdits` (`:198-203`).

### Lista (`EstablecerLista`, `:219-293`)

Lee `m_ObjUsuarioActivo.ColAplicacionesPermisos` (`:241`; `classes/Usuario.cls:107-125`), que consulta `TbUsuariosAplicacionesPermisos` por correo (`modules/Constructor.bas:1019-1022`), y añade una fila por aplicación con «Sí»/«No» en cada perfil (`:249-287`).

### Alta (`ComandoAltaPerfil_Click`, `:8-113`)

1. Sin aplicación en el combo: «Seleccione un elemento de la lista» (`:24-35`).
2. Sin ninguna casilla marcada: «Para dar de alta una aplicación, al menos ha de haber uno de los permisos» (`:66-69`).
3. Busca el permiso existente por correo y aplicación; si no hay, crea uno nuevo (`:70-78`).
4. `RegistrarPermisosCompletos`: `AddNew` o `Edit` con los seis perfiles y `EsUsuarioSinAcceso = "No"` (`classes/UsuarioAplicacionPermisos.cls:335-415`). Si ya existía, **sobrescribe** todos los perfiles.
5. Relee el usuario activo y repinta la lista (`:93-100`).

### Selección y eliminación (`:333-386`, `:115-171`)

1. Clic en la lista: toma el ID de la aplicación, la carga con `getAplicacion` si cambia y, con `AllowEdits`, habilita «Eliminar» y «Alta» (`:342-371`). Las casillas no se rellenan con los permisos de la fila.
2. «Eliminar»: sin selección, «Seleccione un elemento de la lista»; sin registro, «No registrado» (`:122-134`). Pide confirmación con «No» por defecto (`:138-141`), borra la fila (`classes/UsuarioAplicacionPermisos.cls:417-447`), relee el usuario y repinta (`:149-156`).

### Errores

`CorreoAlAdministrador` salvo error 1000 y `MsgBox` crítico (`:105-112, 162-170, 208-216`). `ListaUsuarios_Click` no muestra mensaje (`:378-385`).

## Reglas, validaciones y cálculos

- Solo un administrador conectado puede asignar o retirar permisos (`:185-203, 365-371`).
- El combo incluye las aplicaciones con `ConIconoEnLanzadera = "Sí"` (`:316-322`; `classes/Aplicacion.cls:1166-1176`) y siempre una aplicación fija por ID (`:323-325`).
- Un permiso exige al menos un perfil (`:66-69`); al guardar, `EsUsuarioSinAcceso` pasa a «No» (`classes/UsuarioAplicacionPermisos.cls:369, 404`).
- Clave lógica del permiso: `CorreoUsuario` + `IDAplicacion` (`classes/UsuarioAplicacionPermisos.cls:355`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Permisos del usuario | `TbUsuariosAplicacionesPermisos` | SQL dinámico | no | `modules/Constructor.bas:1019-1022` |
| Permiso puntual | `TbUsuariosAplicacionesPermisos` | SQL dinámico | no | `modules/Constructor.bas:977-981` |
| Aplicación | `TbAplicaciones` por `IDAplicacion` | SQL dinámico | no | `modules/Constructor.bas:482-484` |
| Alta o sobrescritura | `TbUsuariosAplicacionesPermisos` | DAO `AddNew`/`Edit` | no | `classes/UsuarioAplicacionPermisos.cls:353-406` |
| Eliminación | `TbUsuariosAplicacionesPermisos` | DAO `Delete` | no | `classes/UsuarioAplicacionPermisos.cls:435-442` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: perfiles como texto «Sí»/«No»; clave por correo concatenado sin escapar comillas (`classes/UsuarioAplicacionPermisos.cls:355`); la web debe ver al instante los permisos escritos por Access y viceversa.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- «Alta» reasigna `m_ObjAplicacionSeleccionada` a la aplicación del combo (`:28`) y no deshabilita «Eliminar» (`:93-100`): si se había seleccionado otra fila, «Eliminar» borraría los permisos de la aplicación del combo.
- La aplicación fija se añade sin comprobar que no esté ya en el combo (`:323-324`): puede salir duplicada, y si falta en la caché `ColAplicaciones("…")` da error.
- La selección de la lista no carga las casillas (`:333-372`): para cambiar un perfil hay que volver a marcarlos todos y pulsar «Alta», que sobrescribe.
- `ListaUsuarios_Click` traga el error sin avisar al usuario (`:378-385`).
- Desde una ficha en modo alta sin guardar, el usuario activo no tiene correo y el alta falla con «No se sabe el usuario» (`classes/UsuarioAplicacionPermisos.cls:343-345`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: edición en línea de perfiles por fila; clave por ID de usuario en lugar de correo; catálogo de aplicaciones asignables explícito en lugar de ID fijo.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin usuario activo no se abre la pestaña.
2. Un no administrador ve los permisos pero no puede dar de alta ni eliminar.
3. «Alta» sin aplicación o sin ningún perfil se rechaza con su mensaje.
4. «Alta» sobre una aplicación ya asignada sustituye sus seis perfiles.
5. «Eliminar» pide confirmación y, con «Sí», retira la fila de la lista.

## Pendientes

- Por qué una aplicación concreta se añade siempre al combo (`:323`).
- Uso de `EsUsuarioSinAcceso` en el acceso a aplicaciones (no trazado).
- Si `UsuarioAltaEdicionEstado`/`UsuarioSeleccionadoEstado` (`Funciones Genreales.bas:1680, 1745`) se invocan desde el binario.
