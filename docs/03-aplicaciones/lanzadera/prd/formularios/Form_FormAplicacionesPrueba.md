# Form_FormAplicacionesPrueba

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormAplicacionesPrueba.cls` (293 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado «En Pruebas» del catálogo: dos botones de aplicaciones en versión de prueba, con sus etiquetas de versión, y un botón de recarga. Las versiones de este listado no salen del catálogo sino de ficheros `.ini` en rutas fijas del código. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) y no escribe en la base de datos. Las citas son `Form_FormAplicacionesPrueba.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| [Form_FormAplicaciones](Form_FormAplicaciones.md) | pestaña `MENU3`: `FormDetalle.SourceObject = "FormAplicacionesPrueba"` | `Form_FormAplicaciones.cls:63-70` |

Ningún menú principal lo selecciona por código; solo se llega pulsando la pestaña. Es además el `SourceObject` de diseño de `FormDetalle`, que `Form_Load` del contenedor vacía (`Form_FormAplicaciones.form.txt:604`, `Form_FormAplicaciones.cls:35`).

Abre: la aplicación del botón mediante `Lanzar` (`Funciones Genreales.bas:1792`; `classes/Aplicacion.cls:1305-1499`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ComandoCondorPrueba` (`Tag` 23) | botón | `OnClick` → handler con el cuerpo comentado | No hace nada; además se deshabilita siempre al final de la botonera | `:143-169, 86`, `.form.txt:563, 573-577` |
| `ComandoNoConformidadesPrueba` (`Tag` 81) | botón | `OnClick` | `Lanzar Me.ComandoNoConformidadesPrueba` | `:172-197`, `.form.txt:1013-1017` |
| `ComandoActualizar` («Recargar Versiones») | botón | `OnClick` | Botonera, borrado de `.ini` locales, `setVersiones` y versiones de prueba | `:98-141`, `.form.txt:486-488` |
| `lblCondorPruebaVerServidor`, `lblCodorPruebaVerUsuario` | etiquetas | — | Versión del `.ini` remoto y local de la primera aplicación de prueba | `:270-271`, `.form.txt:920-956` |
| `lblNoConformidadesPruebaVerServidor`, `…VerUsuario` | etiquetas | — | Ídem para la segunda | `:284-285`, `.form.txt:1423-1459` |

## Flujos

### Apertura (`Form_Load`, `:199-236`)

1. Si no hay `m_ObjEntorno`, llama a `EVE` sin recoger su error (`:208-210`).
2. Sin usuario conectado, error «No se puede determinar el usuario que ha abierto la aplicación» (`:211-214`).
3. `EstablecerBotonera` (`:215`) y `setVersionAplicacionesEnPrueba` (`:219`). No llama a `setVersiones`.

### Botonera (`EstablecerBotonera`, `:6-95`)

1. Conjunto permitido: administrador ⇒ `ColAplicacionesActivas`; resto ⇒ `ColAplicacionesPermisos` (`:30-36`).
2. Recorre `ColAplicacionesActivas`; salta las que no tienen icono en la Lanzadera o no tienen botón con ese `Tag` (`:38-54`).
3. Administrador ⇒ habilitada (`:57-61`). Resto: sin colección o sin permiso ⇒ deshabilitada; con permiso ⇒ habilitada (`:64-75`).
4. **No** comprueba si la aplicación es de oficina; esa regla solo la aplica `Lanzar` al pulsar (`Funciones Genreales.bas:1817-1821`).
5. Administrador ⇒ `ActualizarAperturas` (`:80-85`); al final, `ComandoCondorPrueba.Enabled = False` (`:86`).

### Versiones de prueba (`setVersionAplicacionesEnPrueba`, `:238-293`)

1. Construye la ruta local (bajo el perfil del usuario) y la remota (bajo `m_URLRutaAplicacionesRemotas`) del `.ini` de cada aplicación con literales fijos (`:261-264, 275-278`).
2. Lee la clave `VersionAplicacion` con `getVersion`; si el fichero no existe devuelve vacío (`:267-268, 281-282`; `Funciones Genreales.bas:259-270`).
3. Escribe «Ver. Servidor: …» y «Ver. Usuario: …» en las cuatro etiquetas (`:270-271, 284-285`).

### Errores

Cada handler acumula el error, envía `CorreoAlAdministrador` salvo el error 1000 y muestra `MsgBox` crítico (p. ej. `:188-196`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Permiso por botón: administrador o fila en `TbUsuariosAplicacionesPermisos` | `EstablecerBotonera` | `:57-75` |
| La primera aplicación de prueba está siempre deshabilitada | `EstablecerBotonera` | `:86` |
| Lanzamiento con línea de comandos propia para la aplicación de prueba de no conformidades | `Aplicacion.Lanzar` | `classes/Aplicacion.cls:1481-1489` |
| «En oficina» solo se comprueba al lanzar | `Lanzar` | `Funciones Genreales.bas:1817-1821` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Catálogo | `TbAplicaciones` (todas) | DAO, SQL fijo, caché | no | `Constructor.bas:386-388` |
| Permisos del usuario | `TbUsuariosAplicacionesPermisos` por `CorreoUsuario` | DAO, SQL concatenado | no | `Constructor.bas:1019-1022` |
| Versiones de prueba | ficheros `.ini` local y remoto | API `GetPrivateProfileString` | — | `Funciones Genreales.bas:263-268` |

Sin escrituras en base de datos; «Recargar» borra `.ini` locales (`:117`; `Funciones Genreales.bas:2998-3017`).

## Coexistencia

Además de lo común a los listados (SQL concatenado en `Constructor.bas:1021`), las versiones dependen de rutas de ficheros codificadas (`:261-262, 275-276`) sin equivalente en el catálogo de base de datos. Ver `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComandoCondorPrueba` usa el `Tag` 23, el mismo que `ComandoCondor` de producción (`.form.txt:577`; `Form_FormAplicacionesFueraOficina.form.txt:4128`); queda inoperante por `:86` y por el handler comentado (`:143-169`).
- Las cuatro etiquetas de versión llevan `Tag` 231 (`.form.txt:923, 956, 1426, 1459`), también las de no conformidades; `ComandoActualizar` llama a `setVersiones` (`:121`) y después las sobrescribe (`:125`).
- `getVersion` solo rellena el error si ya traía texto (`Funciones Genreales.bas:274-275`): los fallos de lectura se pierden y las etiquetas quedan con versión vacía.
- `ComandoNoConformidadesPrueba` nace habilitado (`.form.txt:1013`, sin `Enabled = NotDefault`): si su aplicación falta o no tiene icono, queda activo.
- Versiones de prueba con rutas fijas en el código (`:261-264, 275-278`) en lugar del catálogo.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: tratar «en pruebas» como atributo de la aplicación (ya existe `EnPruebas`, `classes/Aplicacion.cls:172-182`) y generar este listado desde el catálogo; decidir si la pestaña es visible para no administradores.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario con permiso sobre la aplicación de prueba de no conformidades la ve habilitada; sin permiso, deshabilitada.
2. La aplicación de prueba deshabilitada por código no se puede lanzar.
3. Cada aplicación de prueba muestra su versión de servidor y de usuario, o vacío si no hay fichero.

## Pendientes

- Si `EnPruebas` se usa en algún filtro del catálogo (no se usa en este listado).
- Por qué la primera aplicación de prueba está deshabilitada (sin comentario en `:86`).
- Si la pestaña de pruebas debe ocultarse a usuarios no administradores.
