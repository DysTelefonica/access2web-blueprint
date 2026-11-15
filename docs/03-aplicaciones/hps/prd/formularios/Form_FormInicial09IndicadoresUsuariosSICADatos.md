# Form_FormInicial09IndicadoresUsuariosSICADatos

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial09IndicadoresUsuariosSICADatos.cls` (47 líneas) y `.form.txt` (986 líneas) · **PRD:** [README](../README.md)

Hoja-datos alternativa del panel de indicadores para los usuarios SICA. Es un **subformulario hoja** que `Form_FormInicial09Indicadores` puede cargar como `SubFormIndicadores` (`SourceObject ="Form.FormInicial09IndicadoresDatosLocalDatos"` por defecto, `.form.txt:794`); este formulario es el valor alternativo cuyo nombre compara el padre en `If Me.SubFormIndicadores.SourceObject = "FormInicial09IndicadoresUsuariosSICADatos" Then` (`Form_FormInicial09Indicadores.cls:617`, `.form.txt:1893`). Su `.form.txt` se muestra como datasheet (`DefaultView =2`, `.form.txt:14`; `AllowFormView =0`, `.form.txt:122`) con 17 columnas enlazadas a `TbUsuariosSICALocalParaIndicadores`. Llama a `CopiarDatosAIndicadores` y deshabilita el botón `ComandoEditarUsuario` del padre. Las citas son `Form_FormInicial09IndicadoresUsuariosSICADatos.cls:línea` salvo indicación; `FG.bas` abrevía `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial09Indicadores` | `Me.SubFormIndicadores.SourceObject` puede tomar este nombre (la asignación concreta no aparece en `src/`; solo la comprobación) | `Form_FormInicial09Indicadores.cls:617`, `.form.txt:1893` |

Abre nada. El padre reutiliza `Me.SubFormIndicadores.Form` para `ConsultaAExcel` (`Form_FormInicial09Indicadores.form.txt:1493`) y para `SetFocus` (`.form.txt:1927`).

## Controles y eventos

`RecordSource` en diseño: `SELECT TbUsuariosSICALocalParaIndicadores.* FROM TbUsuariosSICALocalParaIndicado...` (recortado por longitud, `.form.txt:76`). `OnCurrent ="[Event Procedure]"` (`.form.txt:78`) y `OnLoad ="[Event Procedure]"` (`.form.txt:121`).

| Grupo | Controles | Inicio en `.form.txt` |
|---|---|---|
| Cajas de texto de identidad | `ID`, `IDHPS`, `Nombre`, `Apellido_1`, `Apellido_2`, `EmpresaTramitadora`, `TramitacionExterna`, `UsuarioSICATSOL`, `UsuarioSICATDE` | `.form.txt:336, 369, 404, 439, 474, 509, 544, 580, 615` |
| Cajas de texto de fechas y datos personales | `FechaCreacion`, `FechaModificacion`, `Correo_e`, `Observaciones` | `.form.txt:651, 686, 721, 760` |
| Cajas de texto de fechas SICA | `FAltaTSOL`, `FAltaTdE`, `FbajaTSOL`, `FBajaTdE` | `.form.txt:795, 831, 867, 903` |
| Etiquetas emparejadas | `Etiqueta603` … `Etiqueta619` con `Caption` igual al nombre del campo | `.form.txt:347, 380, 415, 450, 485, 520, 556, 591, 626, 661, 696, 731, 770, 805, 841, 877, 913` |

Cada caja lleva su `ControlSource` apuntando al campo homónimo (p. ej. `Nombre`, `.form.txt:405`); `EmpresaTramitadora` está mapeado al campo con espacio `Empresa Tramitadora` (`.form.txt:510`).

## Flujos

### Apertura (`Form_Load`, `:32-46`)

1. `CopiarDatosAIndicadores m_Error` (`:33`): sincroniza las tablas físicas `TbDatosLocalParaIndicadores` y `TbUsuariosSICALocalParaIndicadores` desde `TbDatosLocal` y `TbUsuariosSICALocal` (`FG.bas:3006-3054`). La función guarda un `Static s_YaCopiadoEnEstaCarga` (`FG.bas:3015`) para evitar repetir el `DELETE + INSERT` durante la cascada del `Form_Load` (comentario en `FG.bas:3017-3024`).
2. `Me.RecordSource = "SELECT TbUsuariosSICALocalParaIndicadores.* FROM TbUsuariosSICALocalParaIndicadores WHERE id='-1';"` (`:38-40`): `WHERE id='-1'` deja la hoja vacía hasta que el padre la rellene.
3. `Me.Requery` (`:41`).
4. `Me.Parent.ComandoEditarUsuario.Enabled = False` (`:42`).

`errores:` (`:44-46`) envía los errores no 1000 al administrador con `CorreoAlAdministrador m_Error` (`FG.bas:67`) y muestra `MsgBox` crítico (`:46`).

### Cambio de registro (`Form_Current`, `:9-21`)

1. `On Error Resume Next`; `m_ID = Nz(Me.ID, "")` (`:11-12`).
2. Si `Err.Number <> 0` (p. ej. `Me.ID` no accesible con el `RecordSource` `id='-1'`), `Me.Parent.ComandoEditarUsuario.Enabled = False` (`:13-14`); si no, `.Enabled = True` (`:15-16`).
3. `EstablecerSemaforo m_ID` (`:18`): misma rutina que en los subformularios 03 y 07 (`FG.bas:2865-2904`); la búsqueda usa `Me.ID` y el resultado se pinta en `ImagenSemaforo`.
4. `Err.Clear` (`:19`).

## Reglas, validaciones y cálculos

- El `RecordSource` de la hoja lleva `WHERE id='-1'` (`:39`): el datasheet arranca vacío y se rellena cuando el padre cambia `SubFormIndicadores.SourceObject` al nombre de un formulario con `RecordSource` real y asigna un origen concreto (la mecánica vive en `FG.bas:3083-3140`, donde cada rama del `EnumTipoIndicador` fija el `SourceObject` al `Form.FormInicial09IndicadoresDatosLocalDatos`; el caso SICA queda solo mencionado por nombre en `Form_FormInicial09Indicadores.cls:617` y `.form.txt:1893`, sin asignación real en `src/`).
- `CopiarDatosAIndicadores` borra e inserta las dos tablas-fuente cada vez que se la llama con `s_YaCopiadoEnEstaCarga = False` (`FG.bas:3032-3046`); la cascada del bloque 09 la invoca varias veces y solo la primera ejecuta el `DELETE + INSERT` (`FG.bas:3015, 3027-3031`).
- `Form_Current` no recarga la entidad activa: depende de la rama del padre que compara `SourceObject` y llama a `getUsuarioSICA` (`Form_FormInicial09Indicadores.cls:617-627`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Sincronización inicial de las tablas-fuente | `TbDatosLocal`, `TbDatosLocalParaIndicadores`, `TbUsuariosSICALocal`, `TbUsuariosSICALocalParaIndicadores` | `DELETE *` + `INSERT INTO … SELECT` (DAO `Execute`) | `FG.bas:3032-3046` |
| Origen vacío del subformulario | `TbUsuariosSICALocalParaIndicadores` | `SELECT … WHERE id='-1';` | `:38-40` |
| Semáforo del usuario | `TbDatosLocal` (consulta de `URLImagenSemaforo`) | `SELECT` por `ID` | `FG.bas:2891-2897` |
| Entidad activa (rama SICA del padre) | `TbUsuariosSICA` (`getUsuarioSICA`) | `SELECT * FROM TbUsuariosSICA WHERE IDHPS=<ID>;` | `Constructor.bas:1273-1284`, llamada en `Form_FormInicial09Indicadores.cls:617-623` |

Sin `INSERT`/`UPDATE`/`DELETE` directos en este `.cls` (`:1-47`).

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Ya se ve: literal Jet `WHERE id='-1'` (`:39`); SQL concatenado en `getUsuarioSICA` (`Constructor.bas:1273-1284`); dos llamadas de `DoCmd.Hourglass False` en ramas distintas (`:46, 21`) que se solapan si el error llega después de un `On Error Resume Next`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Me.Parent.ComandoEditarUsuario` se deshabilita dos veces: en `Form_Load` (`:42`) y en `Form_Current` cuando `Me.ID` no resuelve (`:14`). Si `Me.ID` quedara vacío con `Err.Number = 0`, el botón se habilita (`:16`) sobre una hoja que solo tiene `WHERE id='-1'`: la acción del padre abre `FormUsuario07SICAEdicion` sobre un usuario sin datos (`Form_FormInicial09Indicadores.cls:154-159`).
- `On Error Resume Next` en `Form_Current` (`:11`) silencia el fallo de `Nz(Me.ID, "")` cuando el `RecordSource` aún no tiene columna `ID` resuelta; la rama `:13-14` queda entonces activa incluso con `m_ID = ""`.
- `EstablecerSemaforo` (`:18`) recibe `m_ID = Nz(Me.ID, "")`: con `id='-1'` siempre será vacío y la rutina termina por ocultar `ImagenSemaforo` (`FG.bas:2882-2888`), lo que no avisa al usuario.
- `Me.RecordSource` se asigna después de `CopiarDatosAIndicadores` (`:38`); si esa rutina falló con 1000, el `.RecordSource` no se ejecuta y la hoja conserva el del `.form.txt` (`.form.txt:76`).
- `CopiarDatosAIndicadores` puede borrar las filas sincronizadas hace un instante por otra cascada: `s_YaCopiadoEnEstaCarga` (estático) sobrevive al ciclo del `.form.txt` (`FG.bas:3015`) pero la rama con `p_CacheAcabDeRegenerar` (`:20-22`) lo reinicia — esa rama llega solo desde `Mod_StartupCacheInitialization.bas:376`.
- El padre solo comprueba `If Me.SubFormIndicadores.SourceObject = "FormInicial09IndicadoresUsuariosSICADatos" Then` (`Form_FormInicial09Indicadores.cls:617`) pero `FG.bas:3083-3140` fija `SourceObject` a `FormInicial09IndicadoresDatosLocalDatos` en todas las ramas del `EnumTipoIndicador`: **no hay asignación real a este nombre en `src/`** (búsqueda sin resultados), así que el formulario queda definido pero nunca montado en la cascada de indicadores.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: o se elimina este subformulario (la rama SICA del padre queda vacía) o se documenta un `WHERE` parametrizable para que el padre pueda conmutar a la vista SICA sin reintroducir el `id='-1'` literal.

## Criterios de aceptación

Propuestos, no aprobados:

1. La hoja se monta vacía al `Form_Load` con `RecordSource WHERE id='-1'`.
2. `CopiarDatosAIndicadores` se invoca una vez por cascada (la segunda llamada, `skip` por `Static s_YaCopiadoEnEstaCarga`).
3. `Me.Parent.ComandoEditarUsuario.Enabled` se desactiva mientras `Me.ID` no resuelva, y se reactiva cuando resuelve.
4. `EstablecerSemaforo` se llama con `m_ID = Nz(Me.ID, "")` y oculta el semáforo si está vacío.
5. Los errores que no sean 1000 se notifican al administrador y muestran `MsgBox` crítico.

## Pendientes

- Confirmar en el binario si existe un punto del flujo que asigne `SourceObject = "Form.FormInicial09IndicadoresUsuariosSICADatos"` desde fuera de `src/` (macro, `AutoExec`, o etiqueta de opción).
- Decidir si este subformulario se conserva, se elimina, o se le quita el `WHERE id='-1'` para que sirva como vista alternativa real.
