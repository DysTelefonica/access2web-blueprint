# Form_FormUsuario07SICAEdicion

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormUsuario07SICAEdicion.cls` (910 líneas) y `.form.txt` (3153 líneas) · **PRD:** [README](../README.md)

Edición de un usuario del sistema **SICA**: mantiene su correspondencia con un usuario de HPS (actual o histórico) y, si la tramitación es externa, también los datos de la empresa tramitadora. Es la ficha hermana de `Form_FormUsuario07SICAAlta`, con la que comparte los cuatro campos de fecha y el bloque de permisos. Formulario **no enlazado**: los valores viajan por reflexión con `RellenarObjetoConDatosLocal` (`:14`) y `HaHabidoCambiosLocal` (`:52`), definidos en este mismo fichero. Las citas son `Form_FormUsuario07SICAEdicion.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormUsuario02Edicion` | `ComandoDetalleUsuarioSICA_Click` cierra la edición SICA abierta y abre esta con `m_ObjUsuarioSICAActivo.ParaAlta = No` | `Form_FormUsuario02Edicion.cls:221-225` |
| `Form_FormUsuario07SICAAlta` | Mismo objeto global, con `ParaAlta = Sí` | `Form_FormUsuario07SICAAlta.cls` |
| `Form_FormUsuario07SICADatos` | Se lo nombra como origen del usuario activo, pero **no tiene ficha todavía** | — |
| `Form_FormInicial00Principal` | `ComandoDetalleUsuarioHPS_Click` lo abre si no está abierto | `:180-182` |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Decide el modo lectura/escritura, carga la copia de inicio y rellena los datos | `:324-376` |
| (formulario) | — | declara `Public Event EdicionRegistro()` | Se emite solo al grabar | `:12`, `:308` |
| `TramitacionExterna` | combo | `BeforeUpdate` → `TramitacionExterna_BeforeUpdate` | Al cambiar, habilita o bloquea el bloque de empresa tramitadora | `:629-748` |
| `UsuarioHPS` | combo | `Exit` → `UsuarioHPS_Exit` | Resuelve el usuario HPS elegido y vuelca sus datos a la pantalla | `:831-908` |
| `UsuarioHistorico` | combo | `Exit` → `UsuarioHistorico_Exit` | Idem para el usuario histórico | `:750-829` |
| `IDEmpresaTramitadora` | combo | — | Empresa tramitadora; la habilita `TramitacionExterna_BeforeUpdate` | `.form.txt:1347`, `:645-647` |
| `ComandoGrabar` | botón | `Click` → `ComandoGrabar_Click` | Valida que haya cambios, registra y refresca | `:274-322` |
| `ComandoDetalleUsuarioHPS` | botón | `Click` | Abre la edición del usuario HPS relacionado | `:171-225` |
| `ComandoAnexar`, `ComandoEliminarAnexoActual`, `ImagenAnexo` | botón, botón, imagen | `Click` | Gestión de anexos de la correspondencia | `:114-170`, `:226-273`, `:590-628` |
| `ComandoEmpresas` | botón | **sin handler** | Se habilita y deshabilita (`:351`, `:648`, `:701`) pero no tiene procedimiento en el `.cls` | `.form.txt:1575` |
| `cmdSalir` | botón | `Click` | Cierra sin guardar | `:110-112` |
| `Nombre`, `Apellido_1`, `Apellido_2`, `Correo_e`, `Observaciones`, `UsuarioSICATDE`, `UsuarioSICATSOL`, `lblTitulo`, `lblExpedienteActividad`, `Etiqueta*`, `EncabezadoDelFormulario`, `Detalle` | textos y etiquetas | — | Datos de la correspondencia y rótulos | `.form.txt:900-1716` |

## Flujos

### Apertura (`Form_Load`, `:324-376`)

1. Pone el reloj de arena (`:329-331`).
2. Traduce `Application.TempVars("SoloLectura")` a `AllowEdits` (`:332-336`). **Único punto del formulario donde se lee el modo.**
3. Exige `m_ObjUsuarioSICAActivo` (`:338-341`).
4. Carga la **copia de inicio** con `Constructor.getUsuarioSICA` (`:342-345`; `Constructor.bas:1254`).
5. Aplica la paleta de colores (`:346-348`).
6. Deja apagados `ComandoGrabar` y `ComandoEmpresas`, y reactiva `ComandoGrabar` si hay edición (`:349-353`).
7. `EstablecerDatos` vuelca el objeto a la pantalla (`:355-358`; `:378-589`).
8. Copia el estilo de tres botones, apaga el reloj (`:359-365`).
9. En error: correo al administrador, `MsgBox` crítico y cierre (`:367-376`).

### Cambiar el modo de tramitación (`TramitacionExterna_BeforeUpdate`, `:629-748`)

1. Sale sin hacer nada si `AllowEdits` es falso (`:637-642`).
2. Si el valor queda vacío, **anula y deshabilita** ocho controles: empresa tramitadora, nombre, dos apellidos, correo, las dos fechas SICA y el botón de empresas (`:645-662`).
3. En la rama contraria los habilita y rellena (`:663-748`). Es el bloque más extenso del fichero.

### Elegir usuario HPS (`UsuarioHPS_Exit`, `:831-908`) o histórico (`UsuarioHistorico_Exit`, `:750-829`)

Ambos siguen el mismo molde:
1. Salen si el formulario es de solo lectura (`:843-848`, `:761-766`).
2. Resuelven el identificador con `getIDUsuario` (`:849`; `FG.bas:2565`) o `getIDUsuarioHistorico` (`:767`), y salen si queda vacío.
3. Cargan el objeto y vuelcan sus propiedades a la pantalla (`:859-908`, `:779-829`).

### Grabar (`ComandoGrabar_Click`, `:274-322`)

1. Reloj de arena (`:280-282`).
2. `RellenarObjetoConDatosLocal` vuelca los controles al objeto (`:284-287`; `:14`).
3. `HaHabidoCambiosLocal` compara contra la copia de inicio; si no hay cambios, aborta con «No hay cambios que guardar» (`:288-295`; `:52`).
4. Marca `ParaAlta = No` y llama a `Edicion` pasando la copia de inicio (`:296-300`).
5. **Muestra el aviso «Edición correcta» antes de refrescar** (`:303`).
6. `ActualizarUsuarioSICA` propaga el cambio a las tablas replicadas (`:304-307`; `FG.bas:3389`).
7. Emite `EdicionRegistro`, apaga el reloj y se cierra (`:308-313`).
8. En error: apaga el reloj, correo al administrador y `MsgBox` (`:315-322`).

### Errores

Todos los manejadores repiten: `Err.Number <> 1000` ⇒ anexar descripción, `CorreoAlAdministrador` y `MsgBox` crítico. El retorno va al global `pregunta` (`VariablesEntorno.bas:220`).

## Reglas, validaciones y cálculos

- **La copia de inicio es la referencia de cambios.** `getUsuarioSICA` la carga al abrir (`:342-345`) y tanto `Edicion` como los handlers de exit la consultan. Sin ella no se puede ni detectar ni aplicar el cambio.
- **Comparación por reflexión, pero local.** Aquí los dos helpers están definidos en este mismo fichero (`:14`, `:52`) y no en `FG.bas` como en `Form_FormUsuario02Edicion`. Dosimplementaciones del mismo patrón.
- **`TramitacionExterna` gobierna ocho controles** (`:645-662`): el bloque de empresa tramitadora solo existe si la tramitación es externa.
- **El modo lectura se decide una sola vez**, con `AllowEdits`, y los tres handlers de exit lo consultan antes de actuar (`:637`, `:763`, `:844`). A diferencia de la edición de usuario, aquí no hay doble aplicación.
- **El aviso de éxito precede al refresco** (`:303` antes de `:304`): si `ActualizarUsuarioSICA` falla, el usuario ya ha visto «Edición correcta».

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Copia de inicio | vía `Constructor.getUsuarioSICA` | lectura por `OpenRecordset` | `:342-345`, `Constructor.bas:1254` |
| Registro de la edición | `UsuarioSICA.Edicion` | escritura con la copia de inicio | `:298-300` |
| Propagación a réplicas | `ActualizarUsuarioSICA` | escritura sobre las tablas replicadas | `:304-307`, `FG.bas:3389` |
| Anexos de la correspondencia | handlers del formulario | ver `:114-273`, `:590-628` | — |
| Identificador del usuario elegido | `getIDUsuario`, `getIDUsuarioHistorico` | lectura auxiliar | `:849`, `:767`, `FG.bas:2565` |

## Coexistencia

Pendiente de `../coexistencia.md`. El patrón dominante es el **estado global**: el usuario en curso, la copia de inicio y hasta el retorno de los `MsgBox` son variables públicas de módulo (`VariablesEntorno.bas:226`, `:246`, `:220`), no parámetros. En web serían un identificador en la ruta y un retorno de la llamada. `DoCmd.Hourglass` aparece en cinco procedimientos (`:281`, `:333`, `:634`, `:764`, `:839`) y no tiene equivalente directo.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **El aviso de éxito se muestra antes de propagar.** `pregunta = MsgBox("Edición correcta"…)` está en `:303` y `ActualizarUsuarioSICA` en `:304-307`: si la propagación a las tablas replicadas falla, el usuario ha leído que la edición fue correcta y después recibe un error.
- `ComandoEmpresas` se habilita y deshabilita en tres puntos (`:351`, `:648`, `:701`) pero **no tiene procedimiento en el `.cls`**: el control existe en el diseño (`.form.txt:1575`) y no hace nada en este código. A verificar contra el binario.
- Los cuatro handlers (`ComandoGrabar`, `Form_Load`, `TramitacionExterna_BeforeUpdate`, `UsuarioHPS_Exit`, `UsuarioHistorico_Exit`) repiten el bloque de reloj de arena sin `On Error Resume Next`: un error dentro del propio manejador deja el cursor de espera encendido, porque la línea que lo apaga está tras la que falla.
- `RellenarObjetoConDatosLocal` y `HaHabidoCambiosLocal` (`:14`, `:52`) duplican lo que `FG.bas:4124` y `FG.bas:4147` hacen para el usuario de HPS. Dos implementaciones del mismo patrón que solo difieren en el destino del volcado.
- `CmdSalir_Click` cierra sin `On Error` ni comprobación (`:110-112`): a diferencia del resto de handlers, no tiene manejador.
- El evento `EdicionRegistro` se declara público (`:12`) pero no hay ningún `WithEvents` ni manejador en `src`: quién lo escucha es una pregunta abierta.
- `m_Error` es variable de módulo y los helpers la sobrescriben `ByRef` desde fuera (`:284`, `:288`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: confirmar la escritura antes de avisar al usuario; un único servicio de alta/edición compartido por usuario HPS y SICA en vez de dos implementaciones de volcado por reflexión; y devolver un resultado de la validación en vez de escribir en un global.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin usuario SICA activo, el formulario se abre explicando que faltan los datos y se cierra.
2. En modo solo lectura los campos no admiten escritura y la pantalla sigue siendo consultable.
3. Poner `TramitacionExterna` vacío anula y deshabilita los ocho controles de empresa tramitadora; ponerlo con valor los habilita.
4. Elegir un usuario HPS o histórico vuelca sus datos a la pantalla; sin identificador no ocurre nada.
5. Grabar sin cambios no escribe y explica por qué.
6. Tras una edición correcta, el aviso aparece **después** de que la propagación termine, y si esta falla no se muestra el aviso de éxito.
7. Cerrar con `cmdSalir` no guarda nada.

## Pendientes

- Quién escucha `EdicionRegistro` y si recarga la lista de origen: no hay manejador en `src`.
- Si `ComandoEmpresas` tiene handler en el binario o si es un control muerto.
- Contenido de `ActualizarUsuarioSICA` (`FG.bas:3389`): qué réplicas escribe y con qué orden respecto al registro.
- `Form_FormUsuario07SICADatos`, que fija `m_ObjUsuarioSICAActivo` y decide la selección, **no tiene ficha**: es el hueco de esta familia.
- Qué componente abre este formulario además de la edición de usuario.