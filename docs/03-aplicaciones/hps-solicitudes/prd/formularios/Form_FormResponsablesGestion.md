# Form_FormResponsablesGestion

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormResponsablesGestion.cls` (392 líneas) y `.form.txt` (25 bloques `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_FormResponsablesGestion.cls:línea` salvo indicación.

Catálogo de responsables: lista los responsables registrados, filtra por nombre y da el alta, la edición y el borrado. Es un formulario **no enlazado**; carga el listado con `constructor.getResponsables`, abre `FormResponsable` como ventana modal de facto (no hay `Dialog` ni `acDialog`, sino `DoCmd.OpenForm`) y se suscribe a sus eventos `Alta` y `Editado` para repintar la lista. Pertenece al menú de entidades: `Entorno.EntidadesFormularios` lo asocia a `EnumEntidades.Responsables` (`Entorno.cls:225`).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Menú de entidades | Se abre por el association `EnumEntidades.Responsables` → `FormResponsablesGestion` | `Entorno.cls:210-231` |
| `Form_FormSolicitudRenovacionDatos` / `Form_FormSolicitudesAltaTraspasoDatos` | No lo abren: sólo abren `FormResponsable` | `Form_FormSolicitudRenovacionDatos.cls:309`, `Form_FormSolicitudesAltaTraspasoDatos.cls:303` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `FormResponsable` | `DoCmd.OpenForm "FormResponsable"` en alta y en edición | `:146`, `:184` |
| `FormResponsable` (eventos) | `WithEvents m_Form`; `m_Form_Alta` y `m_Form_Editado` repintan la lista | `:5`, `:379-387` |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnOpen` → `Form_Open` | Dimensiona, fija el título y carga la lista | `:271-308`, `.form.txt:38-39` |
| `lblTitulo` | etiqueta de encabezado | — | Rótulo «RESPONSABLES» | `.form.txt:208-209` |
| `ComandoAyuda` | botón | `OnClick` → `ComandoAyuda_Click` | Abre el PDF de ayuda | `:105-131`, `.form.txt:227-229` |
| `ComandoActualizar` | botón | `OnClick` → `ComandoActualizar_Click` | Relee la lista | `:82-103`, `.form.txt:320-322` |
| `ListaFiltrados` | lista de valores | `OnClick` → `ListaFiltrados_Click`; doble clic → `ListaFiltrados_DblClick` | Lista de responsables; la selección fija `m_ResponsableSeleccionado` | `:310-377`, `:258-269`, `.form.txt:422-428` |
| `lblListaDeResultados` | etiqueta | — | «Registros filtrados (n)» | `:43`, `:66`, `.form.txt:458-459` |
| `ComandoAlta` | botón | `OnClick` → `ComandoAlta_Click` | Abre `FormResponsable` en modo alta | `:133-163`, `.form.txt:480-482` |
| `ComandoEditar` | botón «Editar/Ver» | `OnClick` → `ComandoEditar_Click` | Abre `FormResponsable` en modo edición con el responsable elegido | `:165-203`, `.form.txt:538-540` |
| `Etiqueta80` | etiqueta | — | Rótulo «Nombre» | `.form.txt:585-586` |
| `Nombre` | cuadro de texto | `BeforeUpdate` → `Nombre_BeforeUpdate` | Filtro por nombre | `:389-391`, `.form.txt:610` |
| `ComandoEliminar` | botón | `OnClick` → `ComandoEliminar_Click` | Pide confirmación y borra el responsable seleccionado | `:205-256`, `.form.txt:643-645` |
| `lblEstado` | etiqueta | — | Rótulo de diseño «URLDirectorioAplicaciones»; no lo usa el código | `.form.txt:703-704` |
| `cmdSalir` | botón «Salir» | `OnClick` → `cmdSalir_Click` | Cierra el formulario | `:12-15`, `.form.txt:724-726` |

## Flujos

### Apertura y filtrado (`Form_Open` → `Filtrar`, `:271-308`, `:17-79`)

1. `Form_Open` dimensiona el formulario (`:280`; `FU.bas:735-759`), deja la edición deshabilitada (`:282`) y carga la lista (`:286`).
2. `Filtrar` deselecciona las filas seleccionadas salvo la primera (`:33-40`), fija el `RowSource` de la lista con las cuatro columnas del responsable (`:42`), deshabilita «Editar/Ver» y «Eliminar» (`:45-46`) y pide los responsables a `constructor.getResponsables` con el texto del filtro (`:48`).
3. Si no hay resultados, sale (`:52-54`); si no, añade una fila por responsable (`:57-64`), actualiza el contador (`:66`), selecciona la primera fila (`:70`) y dispara `ListaFiltrados_Click` (`:71`).

### Selección (`ListaFiltrados_Click`, `:310-377`)

1. Si la columna 0 está vacía, limpia `m_ResponsableSeleccionado` y sale (`:326-334`).
2. Si no hay responsable seleccionado, lo carga con `constructor.getResponsable` por identificador (`:336-338`); si ya había uno distinto, lo recarga (`:341-345`).
3. Habilita «Editar/Ver» y «Eliminar» (`:358-359`).

### Alta (`:133-163`), edición (`:165-203`) y borrado (`:205-256`)

- El alta y la edición cierran la instancia previa de `FormResponsable` (`:142-144`, `:180-182`), preparan la global `m_ObjResposableActivo` (`Nothing` en alta, el responsable elegido en edición) y abren el formulario (`:146`, `:184`).
- El borrado pide confirmación (`:228`), llama a `ResponsableOperaciones.Eliminar` (`:235-238`) y repinta dos veces: con `ComandoActualizar_Click` (`:234`) y otra vez con `Filtrar` (`:235`).

### Errores

- `cmdSalir_Click` usa `On Error Resume Next` sin aviso (`:12-15`).
- `ListaFiltrados_DblClick` arranca con `On Error Resume Next` (`:259`).
- Las ramas `errores:` de los botones rellenan `m_Error`, envían `CorreoAlAdministrador` (`FU.bas:585`) y muestran un `MsgBox` (`:96-101`, `:118-129`, `:152-161`, `:193-200`, `:247-254`, `:298-306`, `:365-374`).
- `Filtrar` tiene su propio manejador (`:75-78`) que sólo rellena `p_Error` y no avisa por sí mismo.

## Reglas, validaciones y cálculos

- **Selección obligatoria:** «Editar/Ver» y «Eliminar» exigen un responsable seleccionado (`:170-173`, `:215-218`).
- **Confirmación de borrado:** el borrado pide confirmación explícita con el botón por defecto «No» (`:228`).
- **No se borra un responsable en uso:** `ResponsableOperaciones.Eliminar` llama a `ValidacionEliminar`, que aborta si el correo del responsable está en alguna solicitud (`ResponsableOperaciones.cls:137-165`, `:98-135`).
- **Filtro por nombre:** la consulta es un `LIKE '%nombre%'` sobre `TbResponsables` (`constructor.bas:2585-2589`).
- **Campos del responsable:** nombre obligatorio, correo con formato válido y correo único (`ResponsableOperaciones.cls:27-53`).
- **Escritura:** el borrado es un `DELETE` por identificador (`ResponsableOperaciones.cls:148-156`); el alta y la edición los hace `FormResponsable` (`ResponsableOperaciones.cls:56-87`).
- **Sin transacciones:** ninguna de las rutas de este formulario abre transacción.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Listado de responsables | vía `constructor.getResponsables` | DAO, `SELECT` con `LIKE` opcional | `:48`, `constructor.bas:2572-2617` |
| Responsable seleccionado | vía `constructor.getResponsable` | DAO, consulta por `IDResponsable` | `:336`, `:342`, `constructor.bas:3431` |
| Comprobación de uso | vía `getSolicitudesConResponsable` | DAO, `INNER JOIN` con `TbSolicitudes` | `ResponsableOperaciones.cls:110`, `constructor.bas:2640-2642` |
| Borrado | `TbResponsables` | `DELETE` con SQL concatenado | `ResponsableOperaciones.cls:148-156` |

## Coexistencia

Pendiente de `../coexistencia.md`. Lo que ya se ve: la pantalla es un CRUD clásico sobre `TbResponsables` con filtro por texto, sin paginación — la consulta trae todos los responsables y el filtrado se resuelve con un `LIKE` en la base (`constructor.bas:2585-2589`). En la web el equivalente es una tabla con búsqueda en servidor y paginación, y el borrado necesita decidir qué ocurre con las solicitudes que referencian al responsable, hoy bloqueado en cliente (`ResponsableOperaciones.cls:98-135`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-73.** `Form_Open` fija dos veces el título de la ventana: `Me.Caption = m_TituloFormulario` (`:281`) y, justo después, `Me.Caption = Me.lblTitulo.Caption` (`:284`), que es el rótulo estático «RESPONSABLES» (`.form.txt:209`): la segunda asignación anula la primera y la ventana se queda sin versión ni usuario.
- **H-HPS-S-74.** Las filas de la lista se arman concatenando con `;` los cuatro campos (`:61`): un `;` en el nombre, el correo o la descripción desplaza las columnas de esa fila, que es el mismo patrón ya señalado en otras pantallas del legacy.
- **H-HPS-S-75.** `ComandoActualizar_Click` (`:90`), `m_Form_Alta` (`:381`) y `m_Form_Editado` (`:386`) llaman a `Filtrar` sin argumento: `Filtrar` devuelve el error en `p_Error` (`:18`) y, al no pasárselo, el fallo se pierde en silencio y la lista queda sin pintar.
- **H-HPS-S-76.** `ListaFiltrados_DblClick` arranca con `On Error Resume Next` (`:259`): traga cualquier fallo sin aviso ni correo.
- **H-HPS-S-77.** El filtro de nombre se interpola en la sentencia SQL (`constructor.bas:2585-2589`): un apóstrofe en el nombre rompe la consulta, y el `Like` no distingue mayúsculas según la collation con la que esté configurada la base.
- **H-HPS-S-78.** El formulario no comprueba ningún permiso: no hay referencia a `EsTecnico`, a `EsAdministrador` ni a los permisos por entidad, de modo que quien llegue al menú de entidades puede dar de alta, editar y borrar responsables (`Entorno.cls:225`; `:133-256`).
- **H-HPS-S-79.** La lista conserva en diseño un `RowSource` de otra entidad, «IDSolicitud;DNI;Nombre;Estado» (`.form.txt:424`), con anchos de columna calculados para DNI, nombre y estado (`.form.txt:426`); el código lo sustituye por «IDResponsable;Nombre;Correo;Descripción» (`:42`), de modo que los anchos no encajan con las columnas reales.
- `ComandoEliminar_Click` repinta la lista dos veces (`:234-235`) y `ComandoActualizar_Click` repinta por tercera vez (`:90`).
- `lblEstado` conserva el rótulo de diseño «URLDirectorioAplicaciones» (`.form.txt:704`) y no lo usa ningún handler.
- `pregunta`, la variable que recoge el resultado de los `MsgBox`, es una global compartida (`VG.bas:182`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: tabla de responsables con búsqueda y paginación en servidor; alta y edición como formulario de recurso con identificador en la URL; y el borrado como operación que declare su efecto sobre las solicitudes que lo referencian en lugar de bloquearlo en cliente.

## Criterios de aceptación

Propuestos, no aprobados:

1. La lista muestra todos los responsables y el filtro por nombre acota el resultado.
2. Sin selección, «Editar/Ver» y «Eliminar» no están disponibles y se informa del motivo.
3. El alta abre el formulario de responsable en modo alta; la edición lo abre con los datos del responsable elegido.
4. Tras un alta o una edición, la lista se repinta sin pulsar «Actualizar».
5. El borrado pide confirmación y, si el responsable está en uso en alguna solicitud, no se borra y se informa de qué solicitudes lo bloquean.
6. Con permisos insuficientes, el borrado y la edición no están disponibles.

## Pendientes

- Quién puede llegar a esta pantalla: `Entorno.EntidadesFormularios` la declara (`:225`), pero no consta en este `.cls` ni en el `Entorno` qué condición de menú la hace visible.
- Si el filtro debería buscar también por correo o por descripción: el único campo de filtro es `Nombre`.
- Comportamiento esperado al abrir `FormResponsable` sin `Dialog`: la ventana queda detrás o delante según el estado de la aplicación, no del formulario.
- Recuento de controles: el inventario fija 28 y el `.form.txt` contiene 25 bloques `Begin`; la diferencia no se reproduce con una regla simple de recuento.
