# Form_FormSubContratacionGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormSubContratacionGestion.cls` (762 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado y gestión de subcontrataciones (`TbSubcontrataciones`) con filtros por empresa, rango de fecha de alta e ID. Es un formulario **no enlazado** (sin `RecordSource`): rellena el `RowSource` de `ListaFiltrados` con SQL construido en `Filtrar` (`Form_FormSubContratacionGestion.cls:163-230`) y desde él abre alta, edición, anexos, exportación a Excel y borrado. Las citas son `Form_FormSubContratacionGestion.cls:línea` salvo indicación. Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`; `Constructor.bas` = `modules/Constructor.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpciones` | `DoCmd.OpenForm "FormSubContratacionGestion"` | `Form_Form0BDOpciones.cls:154` |
| `Form_FormSubContratacionAlta` | Si este formulario está abierto, invoca `ActualizarLista` (pública) tras el alta | `Form_FormSubContratacionAlta.cls:32-33` |

Hacia otros formularios: `FormSubContratacionAlta` (cierra la instancia previa y abre; `:105-108`), `FormSubContratacionEdicion` con `OpenArgs = ID|EstaFacturado|Escritura` (`:255-259`) y `FormAnexos` con la subcontratación en la global `m_ObjEntidadParaAnexoActiva` (`:145-151`; `Variables Globales.bas:252`). No lanza eventos ni devuelve valores.

## Controles y eventos

Controles: 24 (24 `Begin` de control en la sección `Detalle` del `.form.txt:317-1639`; los 8 `Begin` del bloque de valores por defecto, `.form.txt:161-285`, no son controles del formulario).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ListaFiltrados` | cuadro de lista, 6 columnas, con encabezados | `OnClick` → `ListaFiltrados_Click` | Muestra el resultado; col. 0 = ID, col. 5 = IDFactura (ancho 0). Al seleccionar habilita Editar/Eliminar/Adjuntar, carga la subcontratación y muestra `ImagenAnexo` si tiene anexos; Eliminar queda deshabilitado si hay factura | CAP-8 | `.form.txt:363-388`; `:691-757` |
| `ComboEmpresaBusqueda` | combo (`SELECT DISTINCT Empresa`) | `OnEnter`, `OnChange`, `OnExit` | `Requery` al entrar (`:541-544`); en `Change` y `Exit` llama `ActualizarLista` solo si el valor difiere del anterior | CAP-8 | `.form.txt:616-641`; `:519-566` |
| `FechaInicialBusqueda` | texto, formato fecha corta | — | Límite inferior de `FechaAlta`; sin evento propio, se aplica en la siguiente actualización | CAP-8 | `.form.txt:762-775`; `:52-53` |
| `FechaFinalBusqueda` | texto, formato fecha corta | `OnExit` → `FechaFinalBusqueda_Exit` | Límite superior; refresca si cambia | CAP-8 | `.form.txt:792-806`; `:615-636` |
| `ComboIDSubContratacion` | combo 4 columnas (ID, fechas, empresa) | `OnEnter`, `OnChange`, `OnExit` | Igual que el de empresa, con `strIDBusquedaAnterior` | CAP-8 | `.form.txt:850-879`; `:567-614` |
| `ComandoLimpiarEmpresa` / `ComandoLimpiarFechas` / `ComandoLimpiarID` | 3 botones | `OnClick` | Llaman `LimpiarUltimos`; si el filtro tiene valor lo anulan y refrescan | CAP-8 | `.form.txt:1099-1110, 1279-1290, 1459-1470`; `:473-518` |
| `ComandoAltaSubContratacion` | botón «Alta» | `OnClick` | Abre `FormSubContratacionAlta` | CAP-8 | `.form.txt:406-416`; `:94-119` |
| `ComandoEditarSubContratacion` | botón «Editar», inicialmente deshabilitado | `OnClick` | Exige ID numérico seleccionado; abre `FormSubContratacionEdicion` | CAP-8 | `.form.txt:506-517`; `:231-270` |
| `ComandoEliminarSubContratacion` | botón «Eliminar», inicialmente deshabilitado | `OnClick` | Confirma y llama `Eliminar` | CAP-8 | `.form.txt:545-558`; `:271-333` |
| `ComandoAnexar` | botón «Adjuntar», inicialmente deshabilitado | `OnClick` | Abre `FormAnexos` para la subcontratación | CAP-11 | `.form.txt:1639-1651`; `:120-162` |
| `ComandoExportarExcelSubContrataciones` | botón (rótulo de diseño «Comando227»), inicialmente deshabilitado | `OnClick` | Exporta el SQL vigente del listado con `GenerarConsultas` | CAP-8 | `.form.txt:656-667`; `:443-472` |
| `ImagenAnexo` | imagen, invisible por defecto | `OnClick` | Un anexo: `AbrirEnLocal`; varios (`;`): `ComandoAnexar_Click` | CAP-11 | `.form.txt:894-906`; `:655-690` |
| `cmdSalir` | botón «Salir» | `OnClick` | Cierra sin guardar | CAP-8 | `.form.txt:444-454`; `:651-654` |
| `lblTitulo1`, `lblTitulo2` | 2 etiquetas | — | Título; `Form_Open` copia `lblTitulo1` a `Caption` | CAP-8 | `.form.txt:1054-1090`; `:644` |
| `Etiqueta210`, `Etiqueta230`, `Etiqueta231`, `Etiqueta235` | 4 etiquetas («EMPRESA», «ENTRE», «Y», «ID») | — | Rótulos de filtros | CAP-8 | `.form.txt:589-601, 708-747, 823-835` |
| `CuadroSuperior`, `Cuadro2`, `Marco122` | 2 rectángulos y 1 grupo de opciones | — | Decoración / agrupación sin código | sin asignar | `.form.txt:317-350, 483-494` |

Apertura (`Form_Open`, `:637-650`): `AjustarTamaño Me` (`FG.bas:261`) y título; todo bajo `On Error Resume Next`, sin cargar datos: el listado nace con el SQL del `.form.txt` (`IDSubcontratacion = -1`, vacío) (`.form.txt:380-386`).

## Flujos

**Normal.** (1) El usuario fija empresa, fechas y/o ID; cada cambio llama `ActualizarLista` (`:33-93`). (2) Deshabilita Editar, Eliminar, Exportar y Adjuntar y oculta `ImagenAnexo` (`:40-45`). (3) `Filtrar` devuelve el SQL (`:52-53`), que se asigna a `ListaFiltrados.RowSource` y se hace `Requery` (`:59-62`). (4) Sin filas: Exportar queda deshabilitado (`:66-67`); con una fila: la selecciona y ejecuta `ListaFiltrados_Click` (`:68-72`); con varias: habilita Exportar y reaplica el clic sobre la fila ya seleccionada (`:73-81`). (5) Al seleccionar una fila se habilitan los botones y se muestra el icono de anexo (`:716-729`).

**Eliminar** (`:271-333`, `Eliminar` `:334-442`): ID obligatorio (`:278-281`); si la columna 5 tiene valor, «No se puede eliminar un elemento que está facturado» (`:282-291`); confirmación `vbYesNo` con «No» por defecto (`:302-307`). `Eliminar` valida ID numérico (`:355-358`), consulta `EstaFacturado("SUB", id)` (`:365`; `FG.bas:5758-5830`), carga la subcontratación (`:377`), comprueba que ningún anexo esté abierto con `FicheroAbierto` (`:382-411`; `FG.bas:195-208`), borra el registro (`:415-423`) y después los anexos con `Subcontratacion.EliminarAnexos` (`:425`; `Subcontratacion.cls:215-250`), que borra cada fila de `TbAnexos` y su fichero (`Anexo.cls:912-932`). Termina con `ActualizarLista` (`:323`).

**Exportar** (`:443-472`): pasa `ListaFiltrados.RowSource` a `GenerarConsultas("FiltroDeSubContrataciones", …, "No")` (`:449-455`; `FG.bas:1384`).

**Errores.** Cada handler distingue el error propio (nº 1000, texto ya preparado) del inesperado y muestra `MsgBox` crítico. `ComandoActualizarTodo_Click` (`:6-27`) cierra el formulario tras mostrar el error (`:26`); `ComandoActualizarTodo` no figura entre los controles del `.form.txt` (a verificar si queda código huérfano).

## Reglas, validaciones y cálculos

- **Listado vacío por defecto:** sin ningún filtro, `Filtrar` devuelve `IDSubcontratacion = -1` (`:189-195`).
- **Fechas:** `FechaInicial` no válida ⇒ `01/01/1900`; `FechaFinal` no válida ⇒ `01/01/2100` (`:196-201`); se formatean `mm/dd/yyyy` para el literal `#…#` (`:202-203, 215`). El filtro es sobre `FechaAlta` (`:215`).
- **Empresa e ID:** igualdad exacta con la empresa entre comillas (`:213`) e ID numérico concatenado (`:208`); vacíos ⇒ `Like '*' Or Is Null` (`:206, 211`). Orden `FechaAlta DESC` (`:220`).
- **Factura:** columna `IDFactura` por `LEFT JOIN` con `TbFacturaSubcontratacionesInvolucradas` (`:185-188`); no vacía ⇒ «facturada», sin Editar con escritura (`Escritura = "No"`, `:243-249`) y sin Eliminar.
- **Protocolo de error:** las funciones devuelven `"#ERR|texto"` y el llamante detecta cualquier `|` (`:54, 317, 366`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Listado filtrado | `TbSubcontrataciones` ⟕ `TbFacturaSubcontratacionesInvolucradas` | SQL dinámico concatenado como `RowSource` | — | `:185-221, 61` |
| Combos de empresa e ID | `TbSubcontrataciones` | `RowSource` estático (`DISTINCT Empresa`; ID, fechas, empresa) | — | `.form.txt:636, 873` |
| Carga de la subcontratación | `TbSubcontrataciones` | SQL concatenado `WHERE IDSubcontratacion=` (`getdb()`) | — | `Constructor.bas:3859-3861` |
| Anexos de la subcontratación | `TbAnexos` | `Constructor.getAnexosDeSubContratacion` | — | `Subcontratacion.cls:34-52` |
| ¿Facturada? | `TbFacturaSubcontratacionesInvolucradas` | SQL concatenado con `CurrentDb()` | — | `FG.bas:5806-5815` |
| Borrado de la subcontratación | `TbSubcontrataciones` | DAO `OpenRecordset` + `.Delete` | no | `:412-423` |
| Borrado de anexos | `TbAnexos` + ficheros | SQL `DELETE` por `IDAnexo` y `fso.DeleteFile` | no | `Anexo.cls:927-932` |
| Exportación | resultado del SQL del listado | `GenerarConsultas` | — | `:455` |

## Coexistencia

Adaptar a PostgreSQL vinculado: literal de fecha `#mm/dd/yyyy#` y `Between` (`:215`); comodín `Like '*'` de Jet (`:206, 211`); `DELETE *` (`Anexo.cls:927`); `Nz`, `Format` e `IsDate` dependientes de configuración regional (`:196-203`); ID numérico y empresa concatenados sin parámetros; borrado de registro y de anexos en dos pasos sin transacción (`:421-425`); alias `AS ID` y nombres `Tb…` en el `RowSource` fijo (`.form.txt:380-386`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Se borra primero el registro (`:421`) y luego los anexos (`:425`); si `EliminarAnexos` falla, la subcontratación ya no existe y quedan anexos huérfanos, sin transacción.
- Empresa concatenada entre comillas (`:213`): un nombre con apóstrofo rompería el SQL; con `|` en el nombre, el `InStr(flag,"|")` (`:54`) lo trataría como error.
- `LEFT JOIN` con la tabla de facturas (`:185-188`): si una subcontratación figura en varias facturas, podría listarse más de una vez.
- Los handlers `Change` leen `Nz(Me.ComboEmpresaBusqueda,"")` (`:525`, `:573`), valor que durante `Change` aún podría no reflejar el texto tecleado; el refresco se apoyaría entonces en `Exit`.
- En `ComandoEliminarSubContratacion_Click` los errores no restablecen el cursor de espera (`:275`, `:328-332`) y en `ImagenAnexo_Click` se sale con `Hourglass True` hacia `ComandoAnexar_Click` (`:663-667`).
- En `Eliminar`, con varios anexos `FicheroAbierto` se llama sin `FileExists` previo (`:386-396`), mientras que con uno sí (`:399`); `FicheroAbierto` abre el fichero en modo binario (`FG.bas:195-208`), que podría crear uno vacío si falta.
- `FechaInicialBusqueda` no dispara refresco (sin evento, `.form.txt:762-775`).
- `Form_Open` con `On Error Resume Next` (`:638`) oculta fallos de apertura.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtros parametrizados y búsqueda con refresco único; borrado transaccional de registro y anexos; rango de fechas con validación explícita; confirmación con resumen de anexos afectados.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir sin filtros el listado está vacío (`:189-195`).
2. Filtrar por empresa, por fechas o por ID devuelve solo coincidencias, ordenadas por fecha de alta descendente (`:205-220`).
3. Con una única fila, queda seleccionada y habilita Editar, Eliminar y Adjuntar (`:68-72, 716-718`).
4. Una subcontratación con factura no se puede eliminar y se abre en edición sin escritura (`:243-249, 288-291, 373-376`).
5. Eliminar pide confirmación y, si se acepta, borra el registro y sus anexos; el listado se refresca (`:302-323`).
6. Eliminar falla con el mensaje correspondiente si algún anexo está abierto (`:393-408`).
7. El icono de anexo solo aparece si la subcontratación tiene anexos; con varios abre `FormAnexos` (`:724-729, 665-668`).
8. Exportar genera el Excel del filtro vigente y está deshabilitado si no hay filas (`:66-67, 449-455`).
9. Tras un alta, el listado se actualiza (`Form_FormSubContratacionAlta.cls:32-33`).

## Pendientes

- Contenido de `GenerarConsultas` (destino y formato del Excel) (`:455`).
- Quién refresca el listado tras editar: solo se localizó la llamada desde el alta; `Form_FormSubContratacionEdicion` no se ha revisado (`:255-259`).
- Si `ComandoActualizarTodo_Click` (`:6-27`) tiene control asociado en el binario.
- Contenido de `Ejecutar` (invocado desde `AbrirEnLocal`, `FG.bas:218-250`).
