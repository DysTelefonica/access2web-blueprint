# Form_FormE2ECola

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormE2ECola.cls` (336 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Panel de consulta, de solo lectura, de la cola de exportación E2E (`TbE2ECola`): lista filtrable y paginada de 50 filas y panel de detalle de la fila elegida. No tiene acciones sobre la cola: la llenan los productores (hook tras guardar y exportación manual) y la consume un proceso servidor (`Form_FormE2ECola.cls:9-12`). Es un formulario **no enlazado**: sin `RecordSource`, y el `RowSource` de la lista se asigna en ejecución (`.form.txt:289`, `.cls:137`). Las citas son `Form_FormE2ECola.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDGestorEntidades` (opción 12, «Panel de cola E2E») | `ResolverFormDestino("12") = "FormE2ECola"`, pero `TipoAccion("12") = "html"` desvía a `E2EColaHtmlSnapshot_Open` y **no** llama a `DoCmd.OpenForm` | `Form0BDGestorEntidadesLogic.bas:35-36, 55-57`; `Form_Form0BDGestorEntidades.cls:67-73` |
| `Entorno.OtrasOpcionesFormularios` | Mapea la opción `PanelColaE2E` (12) a `"FormE2ECola"`; no se ha encontrado consumidor fuera de `Entorno` | `classes/Entorno.cls:701, 740`; `Variables Globales.bas:158` |

En el commit fijado no hay ningún `DoCmd.OpenForm "FormE2ECola"`: el formulario se conserva como artefacto histórico y su sustituto es un HTML generado (`E2EColaHtmlSnapshot.bas:9-13`; `Form_Form0BDGestorEntidades.cls:64-66`). Salida: `cmdSalir` cierra sin guardar (`:313-316`). No lanza eventos ni abre otros objetos.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnOpen` → `Form_Open` | Control de rol, valores por defecto y primera carga | `:55-100`, `.form.txt:32` |
| `cmbFiltroEstado` | combo, lista de valores (`Todos;Pendiente;Enviado;Error`) | — | Filtro por `Estado`; «Todos» por defecto | `:76-78`, `.form.txt:201-203` |
| `cmbFiltroTipoOrigen` | combo, lista de valores (`Todos;Auto;Manual`) | — | Filtro por `TipoOrigen`; «Todos» por defecto | `:79-81`, `.form.txt:220-222` |
| `txtFechaDesde`, `txtFechaHasta` | textos, formato `dd/mm/yyyy` | — | Filtro sobre `FechaDeteccion` | `:127-128`, `.form.txt:239-258` |
| `cmdAplicarFiltros` | botón | `OnClick` | Vuelve a la página 1 y recarga | `:291-296` |
| `cmdLimpiarFiltros` | botón | `OnClick` | Combos a «Todos», fechas a `Null`, página 1 y recarga | `:298-307` |
| `ListaCola` | lista, 7 columnas (la 7.ª, `IDCola`, oculta y dependiente) | `OnClick` → `ListaCola_Click` | Carga el detalle de la fila | `:171-218`, `.form.txt:287-294` |
| `lblPagination` | etiqueta | — | «Página X de Y (N filas)» | `:156-163` |
| `cmdAnterior`, `cmdSiguiente` | botones | `OnClick` | Retroceden o avanzan una página dentro de los límites | `:269-285` |
| `cmdSalir` (rótulo «Cerrar») | botón | `OnClick` | Cierra el formulario | `:313-316`, `.form.txt:330-331` |
| `txtIDCola` … `txtJsonGenerado` (10 textos) | textos bloqueados y deshabilitados | — | Panel de detalle; `txtJsonGenerado` multilínea con barra vertical | `:226-263`, `.form.txt:359-528` |

Miembros públicos para pruebas sin abrir el formulario: `CheckControlPanelGate`, `ListaCola_ColumnCount` (devuelve 7) y las propiedades `CurrentPage`, `TotalRows` y `SelectedIDCola` (`:35-48, 326-336`).

## Flujos

### Apertura (`Form_Open`, `:55-100`)

1. `Ajustar Me` redimensiona la ventana y el título pasa a `m_TituloFormulario` (`:60-61`; `FUNCIONES UTILES.bas:2703`; `Variables Globales.bas:273`).
2. Sin permiso, «Solo Gestores o Administradores pueden acceder al panel de control E2E» y `Cancel = True` (`:64-69`).
3. Página 1, tamaño `PAGE_SIZE` = 50 y filtros vacíos a «Todos» (`:72-81`; `modE2EColaPanelHelper.bas:27`).
4. `RefreshList` bajo `On Error Resume Next` (`:85-87`).

### Recarga de la lista (`RefreshList`, `:107-149`)

1. Obtiene la base con `getdb`; error si falla o devuelve `Nothing` (`:118-123`).
2. `CountColaFiltered` cuenta las filas filtradas (`:130-132`; `modE2EColaPanelHelper.bas:160-198`).
3. `BuildColaSQL` construye la consulta de la página y se asigna al `RowSource` con `Requery` (`:134-138`).
4. Actualiza la etiqueta de paginación (`:140`).

### Detalle (`ListaCola_Click`, `:171-218`)

1. Lee `IDCola` de la columna 6 (base 0); si es ≤ 0, vacía el panel (`:180-184`).
2. `GetColaDetail` abre la fila por `IDCola` (`:190`; `modE2EColaPanelHelper.bas:224-228`) y `PopulateDetailPanel` copia los 10 campos con `Nz` (`:226-243`). `IDUsuarioOrigen` nulo se muestra como «(N/A)» (`:236-237`; `modE2EColaPanelHelper.bas:30`).

### Errores

`Form_Open` acumula el error, envía `CorreoAlAdministrador` salvo el interno 1000, muestra `MsgBox` y cierra (`:91-99`). `RefreshList` y `ListaCola_Click` solo muestran `MsgBox`, sin correo (`:144-148, 209-217`). Paginación, filtros y salida usan `On Error Resume Next` (`:270, 279, 292, 299, 314`).

## Reglas, validaciones y cálculos

- **Rol:** acceso si `EsAdministrador = Sí` **o** `EsGestor = Sí` (`:36`; `E2EEnlacesService.bas:374-376`).
- **Filtros:** «Todos» o vacío omite la condición; `Estado` y `TipoOrigen` se comparan como texto con comillas escapadas; las fechas solo se aplican si `IsDate` y se formatean `yyyy-mm-dd` (`modE2EColaPanelHelper.bas:113-150`).
- **Orden:** `FechaDeteccion DESC, IDCola DESC` (`modE2EColaPanelHelper.bas:75`).
- **Paginación:** la página 1 es `TOP 50`; la página N excluye con `NOT IN (SELECT TOP (N-1)·50 …)` (`modE2EColaPanelHelper.bas:64-90`). Total de páginas = techo de filas/50, mínimo 1 (`modE2EColaPanelHelper.bas:261-267`). «Siguiente» solo avanza si `página × 50 < total` (`:280`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Recuento filtrado | `TbE2ECola` | SQL dinámico, `dbOpenSnapshot` | `modE2EColaPanelHelper.bas:183-185` |
| Página de la lista | `TbE2ECola` | SQL dinámico como `RowSource` | `:134-138` |
| Detalle | `TbE2ECola` por `IDCola` | SQL dinámico, `dbOpenSnapshot` | `modE2EColaPanelHelper.bas:224-228` |

Sin escrituras. Las filas las insertan `ManualE2EExport` (`TipoOrigen` «Manual» o «Auto», `Estado = 'Pendiente'`, `FechaDeteccion = Now()`, `ManualE2EExport.bas:132-138`) y el hook `modNotifyChangeE2E` (`modNotifyChangeE2E.bas:119`).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: paginación con `TOP` + `NOT IN`, propia de ACE (`modE2EColaPanelHelper.bas:38-40, 84-88`), equivalente a `LIMIT/OFFSET`; literales de fecha `#yyyy-mm-dd#` (`modE2EColaPanelHelper.bas:134-144`); el formulario no muestra las columnas `EstadoImportacion`, `FechaVerificacion` y `MotivoFallo` que añade `ProvisionarTbE2EColaVerificacion` (`M_E2EMigration.bas:406-420`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba; en producción el menú abre el sustituto HTML, no este formulario.

**Posibles defectos (sin corregir, a verificar):**

- `EsGestor` solo se declara (`Variables Globales.bas:247`) y no se asigna fuera de los módulos `Test_`: en ejecución real solo pasarían los administradores (`E2EEnlacesService.bas:376`).
- «Fecha hasta» compara `FechaDeteccion <= #día#` (`modE2EColaPanelHelper.bas:143-144`) y `FechaDeteccion` lleva hora (`Now()`, `ManualE2EExport.bas:135`): las filas de ese mismo día quedarían fuera.
- El comentario dice que la primera carga falla «en silencio» (`:83-84`), pero `RefreshList` tiene su propio manejador con `MsgBox` (`:144-148`).
- Una fila no encontrada devuelve error desde `GetColaDetail` (`modE2EColaPanelHelper.bas:229-233`), así que la rama `m_RS Is Nothing` → `ClearDetailPanel` (`:195-198`) no se alcanza en ese caso: se muestra un error.
- Filtrar o paginar no vacía el panel de detalle ni `m_SelectedIDCola` (`:291-307`): el detalle puede no corresponder a la lista visible.
- `txtJsonGenerado` está deshabilitado (`.form.txt:527-528`); en Access un control deshabilitado no recibe el foco, por lo que no se podría desplazar ni copiar un JSON largo.
- El combo de estado ofrece «Error» (`.form.txt:203`), pero ningún módulo VBA escribe ese valor; solo «Pendiente» y «Enviado» constan (`M_E2EMigration.bas:20-21`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: decidir si la web sustituye este formulario o el panel HTML; filtro «hasta» inclusivo del día completo; mostrar las columnas de verificación.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario sin rol de administrador ni de gestor no accede al panel y ve el mensaje de acceso restringido.
2. Sin filtros, la lista muestra 50 filas ordenadas por fecha de detección descendente y «Página 1 de Y (N filas)».
3. Cada filtro de estado, origen y fechas reduce el recuento y la lista de forma coherente; «Limpiar filtros» vuelve a la página 1 sin filtros.
4. «Anterior» en la página 1 y «Siguiente» en la última no cambian de página.
5. Elegir una fila muestra sus 10 campos; un usuario de origen nulo se muestra como «(N/A)».
6. El panel no ofrece ninguna acción que modifique la cola.

## Pendientes

- Si alguna vía (macro, cinta, ventana Inmediato) abre todavía `FormE2ECola`; en `src` no se ha encontrado.
- Quién debería asignar `EsGestor` y si existe esa asignación fuera de `src`.
- Valores reales de `Estado` que escribe el proceso servidor (fuera de `src`), en particular «Error».
- Si Access permite desplazarse en un cuadro de texto deshabilitado y bloqueado (`.form.txt:524-528`).
