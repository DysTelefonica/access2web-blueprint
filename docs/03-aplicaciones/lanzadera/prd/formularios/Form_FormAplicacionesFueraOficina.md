# Form_FormAplicacionesFueraOficina

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormAplicacionesFueraOficina.cls` (493 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado por defecto del catálogo (aplicaciones utilizables fuera de la oficina): una botonera fija de nueve aplicaciones, cada una con icono, rótulo y dos etiquetas de versión. Cada botón lleva en `Tag` el `IDAplicacion` que lanza. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); lee el catálogo y los permisos desde cachés del `Entorno` y del usuario conectado y no escribe en la base de datos. Las citas son `Form_FormAplicacionesFueraOficina.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| [Form_FormAplicaciones](Form_FormAplicaciones.md) | pestaña `MENU1`: `FormDetalle.SourceObject = "FormAplicacionesFueraOficina"` | `Form_FormAplicaciones.cls:39-46` |
| `Form_FormMenuPrincipalAdmin`, `…Calidad`, `…Usuario` | lo fijan directamente en el subformulario de `FormAplicaciones` al cargar | `Form_FormMenuPrincipalAdmin.cls:68-70`, `Form_FormMenuPrincipalCalidad.cls:62-64`, `Form_FormMenuPrincipalUsuario.cls:57-59` |

Abre: la aplicación de escritorio del botón pulsado mediante `Lanzar` (`Funciones Genreales.bas:1792`), que copia ficheros del servidor al equipo del usuario y la ejecuta con `Shell` (`classes/Aplicacion.cls:1305-1499`).

## Controles y eventos

| Control | `Tag` | Evento y handler | Evidencia |
|---|---|---|---|
| `ComandoAgedysEco` | 2 | `OnClick` → `Lanzar` | `:348-375`, `.form.txt:603-607` |
| `ComandoAgedysTec` | 3 | `OnClick` → `Lanzar` | `:318-345`, `.form.txt:1017-1021` |
| `ComandoGestionRiesgos` | 5 | `OnClick` → `Lanzar` | `:406-433`, `.form.txt:1944-1948` |
| `ComandoBrass` | 6 | `OnClick` → `Lanzar` | `:464-490`, `.form.txt:1659-1663` |
| `ComandoNC` | 8 | `OnClick` → `Lanzar` | `:435-462`, `.form.txt:2459-2463` |
| `ComandoRegistro` | 10 | `OnClick` → `Lanzar` | `:377-404`, `.form.txt:960-964` |
| `ComandoExpediente` | 19 | `OnClick` → `Lanzar` | `:192-217`, `.form.txt:3415-3419` |
| `ComandoSHPS` | 22 | `OnClick` → `Lanzar` | `:248-273`, `.form.txt:3751-3755` |
| `ComandoCondor` | 23 | `OnClick` → `Lanzar` | `:221-246`, `.form.txt:4124-4128` |
| `ComandoActualizar` («Recargar Versiones») | — | `OnClick`: botonera, borrado de `.ini` locales y versiones | `:116-157`, `.form.txt:3262-3263` |
| `ComandoActualizarVersiones` | — | `OnClick`: solo `setVersiones` | `:161-190`, `.form.txt:3627` |
| `lbl<App>VerServidor` / `lbl<App>VerUsuario` | `<ID>;S` / `<ID>;U` | — | `.form.txt:2873-3236, 3356, 3389, 4000, 4033, 4474, 4507` |
| `UsuariosConectadosAgedysTec` | 3 | sin handler; cuadro de texto oculto | `.form.txt:4049, 4063-4065` |

Ocho de los nueve botones nacen deshabilitados en diseño (`Enabled = NotDefault`, `.form.txt:594, 949, 1007, 1649, 1934, 2449, 3405, 3741`); `ComandoCondor` no.

## Flujos

### Apertura (`Form_Load`, `:275-316`)

1. Si no hay `m_ObjEntorno`, llama a `EVE` sin recoger su error (`:285-287`).
2. Sin usuario conectado, error «No se puede determinar el usuario que ha abierto la aplicación» (`:288-291`).
3. `EstablecerBotonera` (`:294`) y `setVersiones` (`:298`; `Funciones Genreales.bas:2285-2349`).

### Botonera (`EstablecerBotonera`, `:7-114`)

1. Conjunto permitido: administrador ⇒ `ColAplicacionesActivas`; resto ⇒ `ColAplicacionesPermisos` (`:25-31`).
2. Recorre `ColAplicaciones` (`:34-40`); salta las que no tienen `ConIconoEnLanzadera = "Sí"` (`:41-43`) o no tienen botón con ese `Tag` (`:46-52`; `getBoton`, `Funciones Genreales.bas:2350-2361`).
3. Aplicación de oficina con usuario fuera ⇒ deshabilitada con icono de no acceso (`:55-61`).
4. Sin colección ⇒ deshabilitada (`:83-88`); con el ID en la colección ⇒ habilitada; sin él ⇒ deshabilitada (`:89-97`). Al administrador le basta con que la aplicación exista en `TbAplicaciones`.
5. La aplicación de nombre «EXPEDIENTES» se habilita siempre, sin permiso (`:98-101`).
6. A diferencia de los otros listados, no llama a `ActualizarAperturas` ni consulta `m_ObjEntorno.Error`.

### Lanzamiento (`Comando<App>_Click` → `Lanzar`)

1. Reloj de arena y `Lanzar Me.<botón>` (p. ej. `:199`).
2. `Lanzar` exige `Tag` numérico, relee la aplicación (`Constructor.getAplicacion`) y bloquea las de oficina fuera de ella con «Esta aplicación sólo se puede ejecutar desde la Oficina» (`Funciones Genreales.bas:1803-1821`). No comprueba permisos.
3. `Aplicacion.Lanzar` sincroniza ficheros, ejecuta la aplicación pasando el correo del usuario y cierra la Lanzadera si el usuario no tiene `MantenerLanzaderaAbierta` (`classes/Aplicacion.cls:1426-1498`).

### Errores

Cada handler acumula el error, envía `CorreoAlAdministrador` salvo el error 1000 y muestra `MsgBox` crítico (p. ej. `:208-216`). `ComandoActualizar` y `ComandoActualizarVersiones` ponen `AllowEdits = False` al empezar y lo restauran a `True` al acabar o fallar (`:125, 144, 154`; `:169, 175, 186`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| «En oficina»: alguna IP del equipo contiene la subred de oficina configurada | `EnOficina` | `Funciones Genreales.bas:1003-1012` |
| `Sí`/`No` en texto para `EjecucionEnOficina` y `ConIconoEnLanzadera` | `Aplicacion` | `classes/Aplicacion.cls:158-161, 1171-1174` |
| Versión de servidor y de usuario por `Tag` `<ID>;S` o `<ID>;U`; «##.##.####» si la aplicación no existe | `setVersiones` | `Funciones Genreales.bas:2314-2335` |
| «Recargar» borra el `.ini` local de cada aplicación con icono y uno adicional de ruta fija | `LimpiarCache` | `:135`; `Funciones Genreales.bas:2998-3017` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Catálogo | `TbAplicaciones` | DAO, tabla abierta entera; caché en `Entorno` | no | `Constructor.bas:110-111`; `classes/Entorno.cls:420-433` |
| Conjunto del administrador | `TbAplicaciones` (todas) | DAO, SQL fijo | no | `Constructor.bas:386-388` |
| Permisos del usuario | `TbUsuariosAplicacionesPermisos` por `CorreoUsuario` | DAO, SQL concatenado | no | `Constructor.bas:1019-1022` |
| Aplicación a lanzar | `TbAplicaciones` por `IDAplicacion` | DAO, SQL concatenado | no | `Constructor.bas:482-484` |

Sin escrituras en base de datos; sí copia y borrado de ficheros locales (`classes/Aplicacion.cls:1428-1431`; `Funciones Genreales.bas:3007, 3016`).

## Coexistencia

Filtros concatenados (`Constructor.bas:483, 1021`) a parametrizar; `Sí`/`No` en texto a mapear. El lanzamiento por ficheros locales y `Shell` no tiene equivalente web. Ver `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- «EXPEDIENTES» se habilita para cualquier usuario (`:98-101`) y `Lanzar` no valida permisos (`Funciones Genreales.bas:1803-1823`).
- `ComandoCondor` nace habilitado (`.form.txt:4124`, sin `Enabled = NotDefault`): si la aplicación 23 falta o no tiene icono, el botón queda activo.
- Doble `MsgBox` ante un fallo de lanzamiento (`Funciones Genreales.bas:1838` y p. ej. `:216`).
- Textos de error con el handler equivocado: `ComandoSHPS` informa «ComandoNC_Click» (`:266`).
- Comprobación `Err.Number <> 0` sin efecto tras un bloque comentado (`:62-69`).
- `Aplicacion.Lanzar` comprueba el icono en lugar de la imagen BMP (`classes/Aplicacion.cls:1407`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: botonera generada desde el catálogo en lugar de botones fijos por `Tag`; permiso validado en el servidor al lanzar.

## Criterios de aceptación

Propuestos, no aprobados:

1. Es el listado que se muestra al entrar.
2. Cada aplicación aparece habilitada solo si el usuario tiene permiso o es administrador.
3. Fuera de la oficina, las aplicaciones de oficina aparecen deshabilitadas.
4. «Recargar Versiones» actualiza las versiones mostradas.

## Pendientes

- Si la excepción de «EXPEDIENTES» es una regla de negocio o un parche temporal.
- Uso previsto de `UsuariosConectadosAgedysTec` (código comentado en `:62-65`).
- Si la aplicación lanzada registra la apertura.
