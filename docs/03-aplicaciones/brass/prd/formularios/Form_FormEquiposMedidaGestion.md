# Form_FormEquiposMedidaGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormEquiposMedidaGestion.cls` (470 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado y gestión del catálogo de equipos de medida (`TbEquiposMedida`): filtro por nombre, alta, edición, borrado, acceso a calibraciones, exportación a Excel y modo selector que devuelve el equipo elegido. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): el listado es un cuadro de lista de lista de valores que rellena `Filtrar` (`Form_FormEquiposMedidaGestion.cls:9-59`). Las citas son `Form_FormEquiposMedidaGestion.cls:línea` salvo indicación; las abreviadas `:N` son de ese fichero. `FG.bas` = `modules/Funciones Generales.bas`. Capacidad de todos los controles: CAP-5 Equipos, equipos de medida y calibraciones.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpcionesOtras` | Etiqueta `lblEquiposMedida`: abre el formulario con `On Error Resume Next` | `Form_Form0BDOpcionesOtras.cls:25-28` |
| `FormEquipoMedida` (alta y edición) | Lo abre este formulario y escucha sus eventos `EquipoAlta` y `EquipoModificado` | `:6`, `:79-80`, `:157-158`, `:398-445`; `Form_FormEquipoMedida.cls:7-8` |
| `FormEquipoMedidaCalibracionesGestion` | Lo abre `ComandoCalibraciones` con el equipo en la global `m_ObjEquipoMedidaActivo` | `:114-118`; `Variables Globales.bas:253` |

Salida: evento `EquipoSeleccionado(m_Equipo As EquipoMedida)` (`:7`), lanzado en `:300` por `ComandoSeleccionarEquipo`, que después cierra el formulario (`:304`). Ningún otro formulario del clon declara un `WithEvents` sobre éste (búsqueda de `EquiposMedidaGestion` solo en `Form_Form0BDOpcionesOtras`), véase Pendientes.

## Controles y eventos

Controles: 26 (26 `Begin` en el `.form.txt`): 14 controles con nombre, más 12 `Begin` estructurales (el propio formulario, 8 plantillas de control por defecto en `.form.txt:142-271` y las secciones de cabecera, detalle y pie en `.form.txt:283, 339, 773`).

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo1`, `lblTitulo2` (2) | etiquetas | — | Título «GESTIÓN DE LOS EQUIPOS DE MEDIDA»; `lblTitulo1` da el `Caption` del formulario | `.form.txt:304, 325`; `:323` |
| `Etiqueta162` | etiqueta | — | Rótulo «EQUIPOS» junto al filtro | `.form.txt:726` |
| `Nombre` | texto no enlazado | `OnExit` → `Nombre_Exit` | Texto de filtro; al salir del control relanza `Filtrar` | `.form.txt:754-755`; `:447-471` |
| `ComandoLimpiarNombre` | botón (icono) | `OnClick` | Vacía `Nombre` y relanza `Filtrar` | `.form.txt:557-559`; `:239-264` |
| `ListaFiltrados` | cuadro de lista, 5 columnas, con cabeceras | `OnClick` → `ListaFiltrados_Click` | Columnas ID (oculta, ancho 0), Equipo, Marca, Modelo, NS; al pulsar carga el equipo y habilita edición | `.form.txt:354-378`; `:342-396` |
| `ComandoAltaEquipo` | botón «Alta Equipo» | `OnClick` | Abre `FormEquipoMedida` vacío | `.form.txt:433-435`; `:66-91` |
| `ComandoEditarEquipo` | botón «Editar Equipo» | `OnClick` | Abre `FormEquipoMedida` con el equipo seleccionado | `.form.txt:406-408`; `:132-169` |
| `ComandoEliminarEquipo` | botón «Eliminar Equipo» | `OnClick` | Borra el equipo seleccionado | `.form.txt:461-463`; `:171-211` |
| `ComandoCalibraciones` | botón «Calibraciones» | `OnClick` | Abre `FormEquipoMedidaCalibracionesGestion` | `.form.txt:491-493`; `:93-130` |
| `ComandoExportarExcel` | botón (icono) | `OnClick` | Exporta el contenido de la lista con `GenerarConsultas` | `.form.txt:518-520`; `:213-237` |
| `cmdSalir` | botón «Salir» (pie) | `OnClick` | `DoCmd.Close` del formulario | `.form.txt:791-793`; `:62-64` |
| `ComandoSeleccionarEquipo` | botón «Seleccionar Equipo» (pie) | `OnClick` | Valida el equipo y lanza `EquipoSeleccionado` | `.form.txt:810, 819-821`; `:266-312` |
| `lblEstadoCalibracion` | etiqueta (pie) | — | Muestra `EstadoCalibracion` del equipo seleccionado; «Calibrado» en diseño | `.form.txt:849-850`; `:366-367, 374-375` |

## Flujos

### Apertura (`Form_Load`, `:314-340`)
`AjustarTamaño Me` (`FG.bas:261`), `Caption` desde `lblTitulo1` (`:322-323`) y `Filtrar` (`:324`). Cualquier error: `MsgBox` crítico y cierre del formulario (`:333-339`).

### Filtrado (`Filtrar`, `:9-59`)
1. Deja el estado inicial: sin selección, solo `ComandoAltaEquipo` habilitado, `ComandoExportarExcel` deshabilitado (`:21-27`) y `RowSource` reducido a la fila de cabecera (`:29`).
2. `Constructor.getEquiposMedida(Nz(Me.Nombre, ""))` (`:30`); sin resultados (`Nothing`) sale con la lista vacía y Excel deshabilitado (`:34-36`).
3. Añade una fila por equipo `IDEquipoMedida;Nombre;Marca;Modelo;NS` (`:37-43`) y habilita Excel (`:44`).
4. Si `ListCount = 2` (cabecera más un único equipo) lo selecciona y ejecuta `ListaFiltrados_Click` (`:45-49`).
5. Error: `p_Error` con número y descripción, salvo el interno 1000 (`:53-58`). Es `Public`: lo invocan los handlers de este formulario (`:196, 248, 324, 405, 430, 455`).

### Selección (`ListaFiltrados_Click`, `:342-396`)
Deshabilita Editar, Eliminar y Calibraciones (`:351-354`). Con `Column(0)` vacío oculta `lblEstadoCalibracion` y termina (`:356-360`). Si no hay equipo cargado, o el ID de la fila difiere del cargado, lo relee con `Constructor.getEquipoMedida` y muestra su estado (`:361-377`); con equipo habilita los tres botones (`:383-385`).

### Alta y edición (`:66-91`, `:132-169`, `:398-445`)
`ComandoAltaEquipo` pone `m_ObjEquipoMedidaActivo = Nothing` y abre `FormEquipoMedida` (cerrándolo antes si estaba abierto). `ComandoEditarEquipo` hace lo mismo con el equipo seleccionado. Si `m_EquipoSeleccionado` es `Nothing` recurre a `Constructor.getEquipo` (`:140-146`). Al llegar `EquipoAlta` o `EquipoModificado` se relanza `Filtrar`.

### Borrado (`:171-211`)
`m_EquipoSeleccionado.Eliminar` (`:190`), sin confirmación previa; limpia globales y relanza `Filtrar` (`:194-199`). `EquipoMedida.Eliminar` rechaza el borrado si el equipo figura en algún evento: «El aparato de medida está registrado en al menos un evento» (`EquipoMedida.cls:469-486`).

### Selección como selector (`:266-312`)
`Calibrado = No` ⇒ error «El equipo seleccionado no está activo. Ya sea porque tiene fecha de fin de uso o por estar fuera de calibración» (`:287-293`). Si es válido lanza `EquipoSeleccionado` y cierra (`:300-304`).

### Exportación (`:213-237`)
`GenerarConsultas("Lista de equipos de medida", Me.hWnd, , Me.ListaFiltrados, "No")` (`:221`); un retorno con `|` se trata como error (`:222-226`). `GenerarConsultas` descuenta la fila de cabecera (`FG.bas:1422-1423`).

### Errores
Todos los handlers: `Hourglass` activo, `m_Error` con el número y el detalle (salvo 1000, ya informado) y `MsgBox` crítico (`:85-90`, `:124-129`, `:163-168`, `:205-210`, `:231-236`, `:258-263`, `:306-311`, `:390-395`, `:414-419`, `:439-444`, `:465-470`).

## Reglas, validaciones y cálculos

- **Estado de calibración:** «Fuera de servicio» si hay `FechaFinServicio`; «Sin Calibraciones» si no hay última calibración; «Calibrado» si `FechaFinCalibracion >= Date`; si no, «Fuera de Calibración» (`EquipoMedida.cls:68-100`). `Calibrado` es Sí solo con «Calibrado» (`EquipoMedida.cls:109-132`). El resultado se cachea en el objeto (`EquipoMedida.cls:75-78`).
- **Borrado:** no se permite si el equipo tiene eventos (`EquipoMedida.cls:469-486`).
- **Filtro:** contiene el texto en `Nombre` (`Constructor.bas:4832-4833`).
- **Permisos:** este formulario no comprueba rol alguno (no hay referencia a `EsAdministrador` en el `.cls`).

## Lecturas y escrituras

| Operación | Origen | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Listado filtrado | `TbEquiposMedida` | SQL dinámico concatenado con `Like '*…*'`; sin filtro abre la tabla entera | — | `Constructor.bas:4830-4837` |
| Lectura por ID | `TbEquiposMedida` | SQL dinámico concatenado (ID de la lista) | — | `Constructor.bas:4892-4894` |
| Lectura por ID (rama alternativa) | `TbEquipos` | SQL dinámico concatenado | — | `Constructor.bas:2339-2340` |
| Eventos del equipo | `getEventosConEquipoMedida` | lectura (no analizada aquí) | — | `Constructor.bas:3137` |
| Borrado | `TbEquiposMedida` | SQL dinámico `Delete *` con `Execute` | no | `EquipoMedida.cls:491-494` |
| Alta y edición | `TbEquiposMedida` | delegado en `FormEquipoMedida` (`Registrar`) | — | `Form_FormEquipoMedida.cls:145` |

## Coexistencia

Adaptar a PostgreSQL vinculado: comodín `*` de `Like` y `Delete *` (dialecto Jet/ACE) en `Constructor.bas:4833` y `EquipoMedida.cls:491`; ID numérico concatenado como texto en `Constructor.bas:4894`; nombre con apóstrofo sin escapar en el filtro; `Nz(.Fields(...), "")` convierte nulos en cadena vacía al hidratar (`Constructor.bas:4849`). Pendiente de `../coexistencia.md` para el resto.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Los botones Editar, Eliminar, Calibraciones y Seleccionar, si `m_EquipoSeleccionado` es `Nothing`, llaman a `Constructor.getEquipo` (`:103, 142, 180, 277`), que lee `TbEquipos` y devuelve un `Equipo` (`Constructor.bas:2325-2340`), no un `EquipoMedida`: la asignación a la variable de tipo `EquipoMedida` (`:5`) fallaría. Con la lista usada con normalidad la rama no se alcanza, porque `ListaFiltrados_Click` ya carga el equipo.
- `ComandoSeleccionarEquipo` no se habilita ni deshabilita en `Filtrar`/`ListaFiltrados_Click`; en el `.form.txt` tiene `Visible = NotDefault` (`.form.txt:810`) y ningún código del `.cls` cambia su visibilidad: a verificar si es un botón oculto siempre.
- `EquipoSeleccionado` no tiene oyente en el clon; el modo selector parece no usarse.
- El handler de borrado informa de errores internos como `ComandoEditarEquipo_Click` (`:207`).
- El borrado se ejecuta sin confirmación (`:190`).
- Tras un error en `ComandoSeleccionarEquipo`, `m_EquipoSeleccionado` queda cargado en estado no válido hasta el siguiente `Filtrar` (`:275-293`).
- Si el usuario pulsa la fila de cabecera, `Column(0)` valdría «ID» y se concatenaría en el SQL (`:356`; `Constructor.bas:4894`): a verificar.
- `Nombre_Exit` filtra solo al salir del control; un cambio de texto y clic directo en la lista depende del orden de eventos (`:447-458`).
- Un nombre con apóstrofo rompe el `Like` concatenado (`Constructor.bas:4833`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: confirmación antes de borrar; filtro con búsqueda parametrizada; mostrar el motivo de rechazo del borrado.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, la lista muestra todos los equipos con Equipo, Marca, Modelo y NS, y solo «Alta Equipo» está habilitado (`:21-27`, `:37-43`).
2. Filtrar por texto muestra solo los equipos cuyo nombre lo contiene (`Constructor.bas:4833`).
3. Con un único resultado, el equipo queda seleccionado y se habilitan Editar, Eliminar y Calibraciones (`:45-49`, `:383-385`).
4. Seleccionar un equipo muestra su estado de calibración (`:366-367`; `EquipoMedida.cls:68-100`).
5. Alta o edición desde `FormEquipoMedida` refresca la lista (`:398-445`).
6. Eliminar un equipo con eventos muestra «El aparato de medida está registrado en al menos un evento» y no lo borra (`EquipoMedida.cls:469-486`).
7. Seleccionar un equipo no calibrado muestra el mensaje de «no está activo» y no lanza el evento (`:287-293`).
8. Exportar genera el Excel con las filas de la lista (`:221`).

## Pendientes

- Quién escucha `EquipoSeleccionado` (referencia dinámica no resuelta; solo se ve el `RaiseEvent` en `:300`) y cuándo se muestra `ComandoSeleccionarEquipo` (`.form.txt:810`).
- Contenido de `getEventosConEquipoMedida` (`Constructor.bas:3137`) y de `FormEquipoMedidaCalibracionesGestion`, que se abre por nombre en `:118`.
- Índices y tipos de `TbEquiposMedida` (ERD de origen).
