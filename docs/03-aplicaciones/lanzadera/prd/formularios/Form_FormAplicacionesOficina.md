# Form_FormAplicacionesOficina

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormAplicacionesOficina.cls` (371 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado «Solo Oficina» del catálogo: una botonera fija de cinco aplicaciones, cada una con su icono, rótulo y dos etiquetas de versión (servidor y usuario). Cada botón lleva en `Tag` el `IDAplicacion` que lanza. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); lee el catálogo y los permisos desde cachés del `Entorno` y del usuario conectado y no escribe en la base de datos. Las citas son `Form_FormAplicacionesOficina.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| [Form_FormAplicaciones](Form_FormAplicaciones.md) | pestaña `MENU2`: `FormDetalle.SourceObject = "FormAplicacionesOficina"` | `Form_FormAplicaciones.cls:51-58` |

Abre: la aplicación de escritorio del botón pulsado mediante `Lanzar` (`Funciones Genreales.bas:1792`), que copia ficheros del servidor al equipo del usuario y la ejecuta con `Shell` (`classes/Aplicacion.cls:1305-1499`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ComandoSeguridad` (`Tag` 16) | botón | `OnClick` | `Lanzar Me.ComandoSeguridad` | `:277-304`, `.form.txt:794-798` |
| `ComandoHPS` (`Tag` 17) | botón | `OnClick` | `Lanzar Me.ComandoHPS` | `:247-274`, `.form.txt:1153-1157` |
| `ComandoGestionContratos` (`Tag` 18) | botón | `OnClick` | `Lanzar Me.ComandoGestionContratos` | `:220-245`, `.form.txt:1663-1667` |
| `ComandoSICA` (`Tag` 20), `lblSICA`, `lblSICAVerServidor`, `lblSICAVerUsuario` | botón y etiquetas | `OnClick` | `Lanzar Me.ComandoSICA`; botón oculto en diseño y visible solo para administradores | `:306-331, 66-83`, `.form.txt:2007, 2019-2023` |
| `ComandoAGEDO` (`Tag` 4) | botón | `OnClick` | `Lanzar Me.ComandoAGEDO` | `:193-218`, `.form.txt:2556-2560` |
| `ComandoActualizar` («Recargar Versiones») | botón | `OnClick` | Recalcula la botonera, borra los `.ini` locales y repinta versiones | `:122-162`, `.form.txt:2377-2379` |
| `ComandoActualizarVersiones` | botón | `OnClick` | Solo repinta versiones (`setVersiones`) | `:164-190`, `.form.txt:2431-2433` |
| `lbl<App>VerServidor`, `lbl<App>VerUsuario` | etiquetas | — | `Tag` `<ID>;S` / `<ID>;U`: `setVersiones` escribe la versión del servidor o del usuario | `.form.txt:1441-1540, 1926-1959, 2318-2351, 2892-2925` |

Los cinco botones de aplicación nacen deshabilitados en diseño (`Enabled = NotDefault`, `.form.txt:784, 1142, 1652, 2008, 2545`).

## Flujos

### Apertura (`Form_Load`, `:333-370`)

1. Si no hay `m_ObjEntorno`, llama a `EVE` sin recoger su error (`:342-344`).
2. Sin usuario conectado, error «No se puede determinar el usuario que ha abierto la aplicación» (`:345-348`).
3. `EstablecerBotonera` (`:349`) y `setVersiones` (`:353`; `Funciones Genreales.bas:2285-2349`).

### Botonera (`EstablecerBotonera`, `:6-119`)

1. Conjunto permitido: administrador ⇒ `ColAplicacionesActivas`; resto ⇒ `ColAplicacionesPermisos` del usuario (`:30-36`).
2. Recorre `ColAplicacionesActivas` (`:38-39`); salta las que tienen `ConIconoEnLanzadera` distinto de «Sí» (`:45-47`) y las que no tienen botón con ese `Tag` en el formulario (`getBoton`, `:49-55`; `Funciones Genreales.bas:2350-2361`).
3. Aplicación de oficina y usuario fuera de la oficina ⇒ botón deshabilitado con icono de no acceso, **también para administradores** (`:58-64`).
4. Administrador ⇒ habilitado con icono de acceso; para el ID 20 muestra el botón y las etiquetas de SICA (`:66-76`). No administrador ⇒ oculta SICA (`:78-83`).
5. Sin colección de permisos ⇒ deshabilitado (`:85-89`); con permiso ⇒ habilitado; sin él ⇒ deshabilitado (`:90-96`).
6. La aplicación de nombre «EXPEDIENTES» se habilita siempre (`:97-100`).
7. Administrador ⇒ `ActualizarAperturas` (`:104-109`).

### Lanzamiento (`Comando<App>_Click` → `Lanzar`)

1. Reloj de arena y `Lanzar Me.<botón>` (p. ej. `:196-200`).
2. `Lanzar` exige `Tag` numérico, relee la aplicación con `Constructor.getAplicacion` y, si es de oficina y el usuario no está en ella, error «Esta aplicación sólo se puede ejecutar desde la Oficina» (`Funciones Genreales.bas:1803-1821`). No comprueba permisos.
3. `Aplicacion.Lanzar` copia ejecutable, icono, imagen e `.ini` si hay cambio de versión o faltan en local, ejecuta la aplicación pasando el correo del usuario y cierra la Lanzadera si el usuario no tiene `MantenerLanzaderaAbierta` (`classes/Aplicacion.cls:1426-1498`).

### Errores

Cada handler acumula el error en `m_Error`, envía `CorreoAlAdministrador` salvo el error interno 1000 y muestra `MsgBox` crítico (p. ej. `:209-217`). `Lanzar` ya muestra su propio `MsgBox` antes de devolver el error (`Funciones Genreales.bas:1837-1838`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| «En oficina»: alguna IP del equipo contiene la subred de oficina configurada | `EnOficina` | `Funciones Genreales.bas:1003-1012`; `Variables Globales.bas:598` |
| Campos `Sí`/`No` en texto: `EjecucionEnOficina`, `ConIconoEnLanzadera`, `EsAdministrador` | propiedades calculadas | `classes/Aplicacion.cls:158-161, 1171-1174`; `classes/Usuario.cls:72-75` |
| Icono: `<recursos del servidor>\iconos\<NombreIconoLanzadera>_C.png` para acceso | `URLIconoAcceso` | `classes/Aplicacion.cls:215-219` |
| Versión mostrada: `VersionServidor` / `VersionUsuario` de la aplicación del `Tag`; «##.##.####» si no existe | `setVersiones` | `Funciones Genreales.bas:2314-2335` |
| «Recargar» borra el `.ini` local de cada aplicación con icono | `LimpiarCache` | `:141`; `Funciones Genreales.bas:2998-3008` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Catálogo | `TbAplicaciones` (todas las filas) | DAO, SQL fijo, caché en `Entorno` | no | `Constructor.bas:386-388`; `classes/Entorno.cls:442-455` |
| Permisos del usuario | `TbUsuariosAplicacionesPermisos` por `CorreoUsuario` | DAO, SQL concatenado | no | `Constructor.bas:1019-1022`; `classes/Usuario.cls:107-120` |
| Aplicación a lanzar | `TbAplicaciones` por `IDAplicacion` | DAO, SQL concatenado | no | `Constructor.bas:482-484` |

Escrituras: ninguna en base de datos. Sí en disco local: copia y borrado de ficheros de la aplicación (`classes/Aplicacion.cls:1428-1431`; `Funciones Genreales.bas:3007`).

## Coexistencia

`SELECT * FROM TbAplicaciones` y los filtros concatenados (`Constructor.bas:483, 1021`) deben pasar a consultas parametrizadas; los `Sí`/`No` en texto deben mapearse o conservarse. El lanzamiento depende de rutas de red, del sistema de ficheros local y de `Shell`: no tiene equivalente directo en web. Ver `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Las etiquetas de versión de SICA y de AGEDO llevan `Tag` 18 (el de contratos) en lugar de 20 y 4 (`.form.txt:2318, 2351, 2892, 2925`): mostrarían las versiones de otra aplicación.
- `ColAplicacionesActivas` no filtra por estado: lee toda `TbAplicaciones` (`Constructor.bas:386-387`).
- `Lanzar` no valida permisos (`Funciones Genreales.bas:1803-1823`); el único control es el `Enabled` del botón, y «EXPEDIENTES» se habilita sin permiso (`:97-100`).
- Doble `MsgBox` ante un fallo de lanzamiento: el de `Lanzar` (`Funciones Genreales.bas:1838`) y el del handler (p. ej. `:217`).
- Con SICA de oficina y usuario fuera de ella, el salto de `:62` deja visibles las etiquetas de SICA con el botón oculto en diseño.
- `LimpiarCache` ignora la colección que recibe y recorre todas las aplicaciones (`Funciones Genreales.bas:2998`).
- `ActualizarAperturas` (`:104-109`) solo actúa sobre etiquetas `lblAperturas*` (`Funciones Genreales.bas:3461`) y este `.form.txt` no tiene ninguna: la llamada no tiene efecto visible.
- `Aplicacion.Lanzar` comprueba la existencia del icono donde debería comprobar la imagen BMP (`classes/Aplicacion.cls:1407`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: catálogo filtrado en servidor por permisos y ubicación; validar el permiso en el endpoint de lanzamiento, no solo en la vista; sustituir la copia local por enlaces o un agente de escritorio.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario sin permiso sobre una aplicación la ve deshabilitada con el icono de no acceso.
2. Fuera de la oficina, toda aplicación de oficina aparece deshabilitada, incluso para administradores.
3. Solo los administradores ven la aplicación SICA.
4. Lanzar una aplicación de oficina desde fuera de la oficina muestra «Esta aplicación sólo se puede ejecutar desde la Oficina».
5. «Recargar Versiones» actualiza las versiones de servidor y usuario de cada botón.

## Pendientes

- Si la aplicación lanzada registra la apertura (no hay escritura en el camino trazado).
- Comportamiento de `Application.Screen.ActiveForm.AllowEdits` cuando el formulario activo es el menú principal (`Funciones Genreales.bas:1801`).
