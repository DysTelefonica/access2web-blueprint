# Form_frmFiltrosAvanzadosSolicitudes

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frmFiltrosAvanzadosSolicitudes.cls` (133 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Diálogo de filtros avanzados del buscador de solicitudes: responsable de calidad, responsable técnico, suministrador, tipo de solicitud y estado. Es un formulario **no enlazado**; los cinco desplegables son `Value List` rellenados por código (`:78-102`). No escribe en la base: modifica un objeto `FiltrosSolicitud` en memoria (`classes/FiltrosSolicitud.cls:18-22`). Las citas son `Form_frmFiltrosAvanzadosSolicitudes.cls:línea` salvo indicación; `FU.bas` abrevia `modules/FUNCIONES UTILES.bas` y `Buscar.cls`, `Form_frmBuscarSolicitudes.cls`.

## Entrada y salida

Solo lo abre Form_frmBuscarSolicitudes, en modo diálogo (`Buscar.cls:252`). Al cerrar con «Aplicar» deja `Tag = "OK"` y se oculta; con «Cancelar», `Tag = "Cancel"` (`:126-127, 132-133`). El buscador lee `Tag` y la propiedad `filtros` y después lo cierra (`Buscar.cls:255-265`). No abre otros formularios.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `cboResponsableCalidad` («Resp. Calidad») | desplegable, `LimitToList` | — | Usuarios de calidad, «(Todos)» = 0 | `:83`, `.form.txt:625-644, 855` |
| `cboResponsableTecnico` («Resp. Técnico») | desplegable, `LimitToList` | — | Jefes de proyecto, «(Todos)» = 0 | `:82`, `.form.txt:875-895, 1106` |
| `cboSuministrador` («Contratista Ppal.») | desplegable, `LimitToList` | — | Suministradores, «(Todos)» = 0 | `:84`, `.form.txt:1126-1146, 1357` |
| `cboTipoSolicitud` («Tipo Solicitud») | desplegable, `LimitToList` | — | Lista fija de tres tipos más «(Todos)» | `:87-98`, `.form.txt:1377-1397` |
| `cboEstado` («Estado») | desplegable, `LimitToList` | — | Estados, «(Todos)» = 0 | `:101`, `.form.txt:1628-1647` |
| `ComandoLimpiarResponsableCalidad`, `ComandoLimpiarResponsableTecnico`, `cmdLimpiarSuministrador`, `cmdLimpiarEstado` | botones | `OnClick` | Ponen su desplegable a 0 | `:14-27, 34-37` |
| `cmdLimpiarTipoSolicitud` | botón | `OnClick` | Pone el tipo a cadena vacía | `:29-32` |
| `cmdAplicar` («Aplicar Filtros») | botón, deshabilitado en diseño | `OnClick` | Copia los valores al filtro, `Tag = "OK"`, se oculta | `:115-128, 65`, `.form.txt:1889-1901` |
| `cmdCancelar` | botón, deshabilitado en diseño | `OnClick` | `Tag = "Cancel"`, se oculta | `:130-134, 66`, `.form.txt:1940-1953` |

## Flujos

**Normal:**

1. `Form_Load` ajusta el formulario y, si el buscador está abierto, toma **la misma instancia** de su filtro (`FiltrosActivos`, `Buscar.cls:322-324`); si no, crea uno vacío (`:48-59`).
2. `PoblarCombos` rellena los desplegables con `LlenarComboDesdeDiccionario` sobre las cachés de `Entorno` y la lista fija de tipos (`:78-102`; `FU.bas:538-571`).
3. `PoblarControlesDesdeFiltros` selecciona en cada desplegable el valor actual del filtro (`:104-113`) y habilita «Aplicar» y «Cancelar» (`:65-66`).
4. «Aplicar» convierte los tres identificadores con `CLng(Nz(…, 0))`, el tipo con `Nz(…, "")` y el estado con `CStr(Nz(…, "0"))` (`:117-124`).

**Alternativos:** «Cancelar» no toca el filtro (`:130-134`); los botones de limpiar solo cambian el desplegable hasta «Aplicar» (`:14-37`).

**Errores:** `Form_Load` delega en `ManejadorErroresFormulario` (`:74-76`; `FU.bas:588-677`). `PoblarCombos`, `PoblarControlesDesdeFiltros`, `cmdAplicar_Click` y `cmdCancelar_Click` no tienen control de errores propio. Si falla la carga de suministradores, `Entorno` devuelve una colección vacía sin avisar (`classes/Entorno.cls:167-169`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Opción «(Todos)» con valor 0 en los desplegables cargados desde diccionario | `LlenarComboDesdeDiccionario` | `FU.bas:554-556, 569` |
| Responsables técnicos: jefe de proyecto de algún expediente con `Pecal = 'Sí'`, usuario sin `FechaBaja` | `UsuarioServicio.getResponsablesTecnicos` | `classes/UsuarioServicio.cls:112-121` |
| Responsables de calidad: permiso `EsUsuarioCalidad = 'Sí'` en la aplicación, sin `FechaBaja` | `UsuarioServicio.getResponsablesCalidad` | `classes/UsuarioServicio.cls:130-141` |
| Suministradores: todos los de `TbSuministradores`, por nombre | `SuministradorServicio.getSuministradores` | `classes/SuministradorServicio.cls:18-34` |
| Estados: todos los de `tbEstados`, por `orden` | `EstadoRepositorio.getTodosLosEstados` | `modules/EstadoRepositorio.bas:10-20`, `classes/Entorno.cls:101-109` |
| Tipos ofrecidos: `PC`, `CD_CA`, `CD_CA_SUB`; la opción «(Todos)» vale una comilla y «Aplicar» la convierte en cadena vacía | lista fija | `:95, 122` |
| Las cuatro listas se leen una vez por sesión y quedan en caché de `Entorno` | propiedades | `classes/Entorno.cls:104-106, 121-124, 140-143, 158-161` |
| Filtro activo: algún identificador > 0, tipo no vacío o estado distinto de `"0"` | `HayFiltrosActivos` | `classes/FiltrosSolicitud.cls:24-34` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Responsables técnicos | `TbExpedientes`, `TbExpedientesResponsables`, `TbUsuariosAplicaciones` (conexión de expedientes) | SQL dinámico | no | `classes/UsuarioServicio.cls:116-121` |
| Responsables de calidad | `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` (conexión de Lanzadera) | SQL dinámico, `IDAplicacion` concatenado | no | `classes/UsuarioServicio.cls:134-141` |
| Suministradores | `TbSuministradores` (conexión de expedientes) | SQL dinámico | no | `classes/SuministradorServicio.cls:25-34` |
| Estados | `tbEstados` (conexión de Condor) | SQL dinámico | no | `modules/EstadoRepositorio.bas:16-20` |

Sin escrituras en base de datos.

## Coexistencia

Pendiente del resumen de coexistencia (`../coexistencia.md`). Las cuatro listas salen de tres conexiones distintas (Condor, expedientes y Lanzadera; `FU.bas:64-70, 84-108, 156-178`); la versión web debe resolver los mismos conjuntos de usuarios y suministradores, con los literales `'Sí'` comparados según la intercalación de destino (`classes/UsuarioServicio.cls:119, 139`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La etiqueta dice «Contratista Ppal.» (`.form.txt:1357`), pero el filtro acepta cualquier suministrador vinculado al expediente (`classes/SolicitudServicio.cls:149-151`).
- El desplegable técnico lista solo jefes de proyecto de expedientes PECAL (`classes/UsuarioServicio.cls:119`); la búsqueda filtra por cualquier responsable (`classes/SolicitudServicio.cls:140`).
- El valor «(Todos)» del tipo es una comilla (`:95`), mientras `PoblarCombos` y `PoblarControlesDesdeFiltros` asignan cadena vacía (`:98, 111`): el desplegable podría quedar sin selección visible.
- `cmdAplicar_Click` modifica la instancia compartida con el buscador (`:56, 117-124`) sin control de errores; un fallo a mitad dejaría el filtro parcialmente cambiado.
- `idEstado` admite varios valores según su comentario (`classes/FiltrosSolicitud.cls:22`), pero el desplegable solo permite uno.
- El tipo `PC_SUB` del estudio previo no figura en la lista fija (`:95`); véase D-CON-1 en el [README](../README.md).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtros integrados en el propio buscador en lugar de un diálogo; selección múltiple de estados.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, cada desplegable muestra el valor del filtro activo o «(Todos)».
2. «Aplicar» devuelve los cinco valores al buscador y relanza la búsqueda; «Cancelar» lo deja igual.
3. Cada botón de limpiar devuelve su desplegable a «(Todos)» sin aplicar.
4. Las listas de responsables y suministradores contienen los mismos conjuntos que las reglas de la tabla.

## Pendientes

- Contenido de `FormatearDescripcionValidacion` y `LogErrorUI` (`FU.bas:625, 672`).
- Valor de `IDAplicacion` usado en `getResponsablesCalidad` (`classes/UsuarioServicio.cls:138`).
