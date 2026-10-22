# Form_FormNCAuditoriaSeguimientoNC

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCAuditoriaSeguimientoNC.cls` (358 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Rev. Estados» del seguimiento de auditoría: lista las NC de auditoría pendientes de revisión, filtrables por responsable de Calidad, estado y auditoría, y abre la NC elegida en `FormNCAuditoria`. Se aloja en `FrmDetalle` de [Form_FormNCAuditoriaSeguimiento](Form_FormNCAuditoriaSeguimiento.md). Es un formulario **no enlazado** (sin `RecordSource`); la lista no sale de una consulta sino de las colecciones `ColSegsNCAuditoria…` del entorno. Las citas son `Form_FormNCAuditoriaSeguimientoNC.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_FormNCAuditoriaSeguimiento | Botón de navegación `NavNC` carga este formulario en `FrmDetalle` | `Form_FormNCAuditoriaSeguimiento.form.txt:1061` |
| `PintarIndicadores` | Llama a `Filtrar` si `FrmDetalle.SourceObject` es este formulario | `FG.bas:332-334` |

Sale hacia `FormNCAuditoria`: cierra la instancia abierta sin guardar y la abre con `DoCmd.OpenForm`, tras fijar `m_ObjNCAuditoriaActiva` (`:81-86`). No se cierra a sí mismo salvo error en `Form_Load` (`:160`).

## Controles y eventos

`.form.txt` declara 24 bloques `Begin`: el formulario (`:4`), 9 de valores por defecto (`:713-772`), la sección `Detalle` (`:778-780`) y los 13 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaFiltrados` | lista (valores) | `OnClick`, `OnDblClick` | Muestra IDNoconformidad, Auditoria, Nº, Descripción, Responsable, Estado; el clic fija la NC seleccionada y habilita «Detalle»; el doble clic abre el detalle si está habilitado | `:267-327`, `.form.txt:802-808` |
| `lblListaDeResultados` («Registros filtrados») | etiqueta | — | Contador «Registros filtrados (n)» | `:230, 254`, `.form.txt:846-847` |
| `ComandoVerDetalle` («Detalle») | botón | `OnClick` | Abre `FormNCAuditoria` con la NC seleccionada | `:64-98`, `.form.txt:867-869` |
| `Responsable` + `Etiqueta57` («Responsable implant.») | combinado (valores) + etiqueta | `AfterUpdate` | Filtra por responsable de Calidad | `:330-356`, `.form.txt:915, 945` |
| `Estado` + `Etiqueta261` («Estado tarea») | combinado (valores) + etiqueta | `AfterUpdate` | Filtra por estado (5 valores) | `:100-126, 193-211`, `.form.txt:975, 1004` |
| `Auditoria` + `Etiqueta86` («Auditoría») | combinado (valores) + etiqueta | `AfterUpdate` | Filtra por auditoría | `:14-40`, `.form.txt:1413, 1628` |
| `ComandoLimpiarResposable` | botón | `OnClick` | Pone `Responsable` a Null y filtra | `:56-60`, `.form.txt:1034-1036` |
| `ComandoLimpiarEstado` | botón | `OnClick` | Pone `Estado` a Null y filtra | `:48-52`, `.form.txt:1217-1219` |
| `ComandoLimpiarAuditoria` | botón | `OnClick` | Pone `Responsable` (no `Auditoria`) a Null y filtra | `:42-46`, `.form.txt:1444-1446` |
| `Etiqueta237` | etiqueta | — | Decorativa | `.form.txt:829` |

## Flujos

### Carga (`Form_Load`, `:128-161`)

1. Cursor de espera, `Ajustar Me` y `RowSource` de la lista con la cabecera de columnas (`:134-136`).
2. `EstablecerCombos` rellena responsables de Calidad (`FG.bas:1921`), los 5 estados (`:193-211`) y las auditorías (`FG.bas:3137`) (`:163-191`).
3. Si el usuario conectado está en `ColUsuariosCalidad`, preselecciona su nombre en `Responsable` (`:142-144`).
4. Preselecciona el estado «Cerrada pendiente de CE» y filtra (`:145-146`).

### Filtrado (`Filtrar`, `:215-265`)

Deshabilita «Detalle», pone la cabecera, llama a `constructor.getNCsDeAuditoriaBusqueda` con responsable (texto), `Auditoria.Column(0)` y `Estado.Column(0)` (`:237-239`), añade una fila por NC (`:251`), actualiza el contador, selecciona la fila 1 y dispara `ListaFiltrados_Click` (`:254-256`).

### Selección y detalle

`ListaFiltrados_Click` carga con `constructor.getNCAuditoria` la NC de `Column(0)` si cambia (`:286-298`) y habilita «Detalle» (`:305`). «Detalle» reutiliza la NC seleccionada o la carga; sin selección avisa «Seleccione un elemento de la lista» (`:71-80`).

### Errores

Error 1000 ⇒ advertencia con `m_Error`; otro ⇒ `CorreoAlAdministrador` y mensaje crítico (`:33-39`, `:91-97`, `:153-160`).

## Reglas, validaciones y cálculos

- Los cinco estados de la combo: CERRADAPTECE, CERRADAPTECECADUCADA, CERRADACENOCONFORME, ACSSINTAREAS y REGISTRADA (`:201-205`).
- `getNCsDeAuditoriaBusqueda` elige la colección por título de estado y, sin estado reconocido, usa `ColSegsNCAuditoriaTotales`; filtra por nombre de responsable e id de auditoría (`constructor.bas:2672-2701`).
- Las descripciones sustituyen `;` por `:` porque la lista es de valores (`:251`); el número se muestra con dos dígitos (`:251`).
- No hay comprobación de rol.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Listado y combos | Colecciones del entorno (`ColSegsNCAuditoria…`, `ColUsuariosCalidad`, `ColAuditorias`, `ColEstadosNCTitulo`) | Memoria; carga en `Entorno`, fuera de esta ficha | `constructor.bas:2672-2686`, `FG.bas:1931, 3149` |
| Cargar la NC seleccionada | `TbNoConformidadesAuditoria` | SQL dinámico concatenado con `p_IDNC` | `constructor.bas:5479-5500` |
| Escrituras | — | Ninguna | `:1-358` |

## Coexistencia

La concatenación de `p_IDNC` en `getNCAuditoria` debe revisarse al pasar a PostgreSQL vinculado (tipo del id, comillas). Las comparaciones de responsable y estado se hacen por texto en memoria, sin dialecto. `Format(…, "00")` y `Nz` son de VBA, no del motor.

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual**

- Filtros combinados responsable + auditoría sobre una única colección elegida por estado (`constructor.bas:2672-2701`).
- Tras el filtrado se selecciona la fila 1 y se dispara el clic (`:255-256`).

**Posibles defectos (a verificar)**

- `ComandoLimpiarAuditoria_Click` limpia `Responsable`, no `Auditoria` (`:44`).
- `Auditoria_AfterUpdate` informa de error con el texto «Responsable_AfterUpdate» (`:34`).
- `Filtrar` hace `lst.Selected(1) = True` aunque no haya filas y el contador resta 1 a `ListCount` por la cabecera (`:254-255`).
- `ListaFiltrados_Click` vuelve a poner el cursor de espera al salir con `m_NCAuditoriaSeleccionada Is Nothing` (`:299-303`).
- El `RowSource` de diseño de la lista tiene columnas de tareas (IDAccionRealizada, Tarea…), distintas de las del código (`.form.txt:804`, `:136`).
- `Filtrar` sale sin limpiar el filtro cuando `getNCsDeAuditoriaBusqueda` devuelve `Nothing`, y deja `m_NCAuditoriaSeleccionada` previa (`:243-245`).

**Propuestas web pendientes de decisión**

- Sustituir las listas de valores por una tabla paginada con filtros en servidor.
- Fijar el estado inicial según rol y no solo para Calidad.

## Criterios de aceptación

1. Al abrir, el responsable es el usuario si es de Calidad y el estado «Cerrada pendiente de CE» (`:142-145`).
2. Cambiar responsable, estado o auditoría refiltra y actualiza el contador (`:14-40, 100-126, 330-356`).
3. Las tres «X» vacían su filtro (`:42-60`); el comportamiento de la de auditoría está pendiente de verificar.
4. «Detalle» (o doble clic) abre `FormNCAuditoria` con la NC elegida; sin selección avisa (`:71-86, 322-327`).
5. Cada error muestra el mensaje y avisa al administrador salvo error 1000 (`:33-39`).

## Pendientes

- Carga de `ColSegsNCAuditoria…` en `Entorno` (referencia dinámica no resuelta en esta ficha; `constructor.bas:2672-2686`).
- Valor inicial de `ColumnCount` de `Auditoria` y `Responsable` en el binario (afecta a `FG.bas:1938-1942`).
