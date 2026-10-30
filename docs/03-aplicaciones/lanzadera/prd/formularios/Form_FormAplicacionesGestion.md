# Form_FormAplicacionesGestion

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormAplicacionesGestion.cls` (149 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Búsqueda» de la gestión de aplicaciones: busca aplicaciones por nombre, fija la aplicación activa de la sesión y muestra quién la está usando. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); no escribe en la base de datos. Las citas son `Form_FormAplicacionesGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormAplicacionGeneral` | `MENU1_Click` fija `FormDetalle.SourceObject = "FormAplicacionesGestion"` y se suscribe a sus eventos con `WithEvents` | `Form_FormAplicacionGeneral.cls:3, 62-68` |

No abre otros formularios. Eventos hacia el contenedor: `AplicacionSeleccionada` (`:4`, lanzado en `:112`) y `QuitarAplicacionSeleccionada` (`:5`, lanzado en `:83`).

Estado compartido: la global `m_ObjAplicacionActiva` (`Variables Globales.bas:101`), que lee `Form_FormAplicacionDatosGenerales` (`Form_FormAplicacionDatosGenerales.cls:90-94`). El formulario no recibe parámetros.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `NombreAplicacion` | texto | `AfterUpdate` | Filtra la lista por nombre | `:124-148`, `.form.txt:832-833` |
| `ComandoLimpiarNombre` | botón | `OnClick` | Vacía `NombreAplicacion` y vuelve a filtrar (lista completa) | `:7-31`, `.form.txt:647-651` |
| `ListaAplicaciones` | lista 2 columnas (ID oculto), con cabecera | `OnClick` | Carga la aplicación, la fija como activa y muestra los usuarios conectados | `:93-122`, `.form.txt:583-589` |
| `AplicacionSeleccionada` | texto deshabilitado y bloqueado | — | Nombre de la aplicación activa | `:45, 69, 89, 109`, `.form.txt:881-896` |
| `ComandoQuitarAplicacionActiva` | botón | `OnClick` | Vacía la aplicación activa y avisa al contenedor | `:80-84`, `.form.txt:923-925` |
| `UsuariosConectados` | texto | — | Correos de los usuarios que usan la aplicación, uno por línea | `:97, 111`, `.form.txt:1185` |

## Flujos

### Apertura (`Form_Load`, `:86-91`)

1. Inicializa la lista solo con la fila de cabecera `ID;Nombre` (`:87`).
2. Si ya hay `m_ObjAplicacionActiva`, muestra su nombre en `AplicacionSeleccionada` (`:88-90`).

### Búsqueda (`Filtrar`, `:33-78`)

1. Anula la aplicación activa y vacía `AplicacionSeleccionada` (`:44-45`); reinicia la lista (`:47`).
2. `Constructor.getListaAplicacionesPorNombre` devuelve las aplicaciones cuyo nombre contiene el texto; con texto vacío, toda `TbAplicaciones` (`:49`; `Constructor.bas:917-922`).
3. Añade una fila `ID;Nombre` por aplicación (`:56-63`).
4. Si queda un único resultado (`ListCount = 2`, cabecera incluida), lo selecciona y ejecuta `ListaAplicaciones_Click` (`:64-70`).

### Selección (`ListaAplicaciones_Click`, `:93-122`)

1. Vacía `UsuariosConectados` y lee el ID de la fila (`:97-98`). Si no es numérico (fila de cabecera), anula la aplicación activa y sale sin avisar al contenedor (`:99-103`).
2. `Constructor.getAplicacion` lee la aplicación (`:104`; `Constructor.bas:482-494`) y la fija como activa (`:108-109`).
3. `CadenaCorreoUsuariosConectados` rellena `UsuariosConectados` (`:111`; `Aplicacion.cls:1662-1698`).
4. Lanza `AplicacionSeleccionada`; el contenedor actualiza su rótulo (`:112`; `Form_FormAplicacionGeneral.cls:58-60`).

### Errores

Los handlers acumulan el error en `m_Error`, envían `CorreoAlAdministrador` salvo el error interno 1000 y muestran `MsgBox` crítico (`:23-30, 114-121, 140-147`). `Filtrar` solo devuelve el texto en `p_Error` (`:73-76`).

## Reglas, validaciones y cálculos

- **Filtro por nombre:** `LIKE '*<texto>*'` sobre `NombreAplicacion` (`Constructor.bas:920-922`); el resultado se indexa por `IDAplicacion` sin duplicados (`Constructor.bas:939-944`). Sin `ORDER BY`: el orden de la lista es el que devuelva el motor.
- **Usuarios conectados:** `ColUsuariosUsandola` lee el fichero de bloqueo `.laccdb` del backend de la aplicación, extrae cadenas ASCII de tres o más caracteres y las busca como usuario de red con `Constructor.getUsuario` (`Aplicacion.cls:1537-1611`). Sin bloqueo, muestra «Nadie está usando la aplicación» (`Aplicacion.cls:1677-1679`).
- **Permisos:** este formulario no comprueba el rol. Solo es accesible desde el menú de administración, que abre `LoginCorrecto` cuando `EsAdministradorCalculado = Sí` (`Variables Globales.bas:724-725`; `Form_FormMenuPrincipalAdmin.cls:241-257`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Lista por nombre | `TbAplicaciones` | SQL dinámico concatenado (o tabla entera) | `Constructor.bas:917-925` |
| Aplicación por ID | `TbAplicaciones` | SQL dinámico concatenado | `Constructor.bas:482-484` |
| Usuarios conectados | fichero `.laccdb` del backend + usuario por usuario de red | E/S binaria + DAO | `Aplicacion.cls:1549-1558, 1595` |

Sin escrituras.

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `LIKE` con comodín `*` de Jet/ACE (en PostgreSQL, `ILIKE '%…%'`); el texto de búsqueda se concatena sin escapar comillas (`Constructor.bas:922`). La detección de usuarios conectados depende del fichero `.laccdb` de Access y no tiene equivalente directo en PostgreSQL.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Filtrar` anula la aplicación activa (`:44-45`) sin lanzar `QuitarAplicacionSeleccionada`: el rótulo del contenedor sigue mostrando la aplicación anterior (`Form_FormAplicacionGeneral.cls:59`). Lo mismo al pulsar la cabecera de la lista (`:99-103`).
- Un nombre con apóstrofo rompe la consulta del filtro (`Constructor.bas:922`).
- `ListaAplicaciones_Click` llama a `CadenaCorreoUsuariosConectados` sin `p_Error` (`:111`): si no se encuentra el backend, el campo queda vacío sin aviso (`Aplicacion.cls:1538-1540`).
- La primera captura de error en `ColUsuariosUsandola` desactiva el manejador con `On Error GoTo 0` (`Aplicacion.cls:1570-1572`); el mensaje de error dice «Lanzar» (`Aplicacion.cls:1620`).
- Rótulos heredados de otro formulario: el formulario se titula «Gestión de Usuarios» (`.form.txt:39`) y ambos botones dicen «Vacía el campo de usuario de búsqueda» (`.form.txt:651, 926`).
- `m_ObjAplicacionSeleccionada` acumula la última aplicación del bucle aunque haya varias (`:61`); solo se usa si hay una.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: búsqueda paramétrica ordenada por nombre; usuarios conectados a partir de sesiones registradas en lugar del fichero de bloqueo; la aplicación activa como parámetro de ruta, no como global.

## Criterios de aceptación

Propuestos, no aprobados:

1. Escribir parte del nombre lista las aplicaciones que lo contienen.
2. Con un único resultado, la aplicación queda seleccionada y el contenedor muestra «Aplicación Activa: <nombre>».
3. Limpiar el nombre lista todas las aplicaciones.
4. Quitar la aplicación activa vacía el campo y el rótulo del contenedor.
5. Seleccionar una aplicación sin usuarios conectados muestra «Nadie está usando la aplicación».
6. Solo un administrador llega al formulario.

## Pendientes

- Si el contenedor recibe los eventos cuando se abre desde el menú de administración por la instancia por defecto (`Form_FormMenuPrincipalAdmin.cls:256`).
- Formato real del fichero `.laccdb` y fiabilidad de la extracción de usuarios (`Aplicacion.cls:1565-1583`).
- Contenido de `URLDirectorioAplicacion` (`Aplicacion.cls:538`).
- Si el filtro distingue mayúsculas: depende de la intercalación del backend, no del código (`Constructor.bas:922`).
