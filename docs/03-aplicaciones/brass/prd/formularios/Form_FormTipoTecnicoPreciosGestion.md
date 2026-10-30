# Form_FormTipoTecnicoPreciosGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormTipoTecnicoPreciosGestion.cls` (447 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pantalla de gestión del histórico de precios por hora de cada perfil (tipo de técnico). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): construye a mano una lista de valores (`ListaPreciosPerfil`) con `TbTipoTecnicoPrecios` filtrada por vigencia (`Form_FormTipoTecnicoPreciosGestion.cls:302-375`) y delega alta, edición y borrado en `FormTipoTecnicoPrecio`, `FormTecnicoTipo` y en la clase `TipoTecnicoPrecio.cls`. Las citas son `Form_FormTipoTecnicoPreciosGestion.cls:línea` salvo indicación; `VG.bas` = `modules/Variables Globales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormTecnicoGestion` | `DoCmd.OpenForm "FormTipoTecnicoPreciosGestion"` desde `ComandoTipoTecnico_Click` | `Form_FormTecnicoGestion.cls:335-339` |
| `Form_FormFacturaAlta` | misma apertura desde `ComandoPerfiles_Click` | `Form_FormFacturaAlta.cls:125-129` |

Hacia fuera: abre `FormTecnicoTipo` (alta de perfil, `:19`) y `FormTipoTecnicoPrecio` (alta/edición de precio, `:53, 94`), cerrándolos antes si estaban abiertos (`:16-18, 50-52, 91-93`) mediante `FormularioAbierto` (`Funciones Generales.bas:171`). Pasa el precio elegido en la global `m_ObjTipoTecnicoPrecioActivo` (`VG.bas:248`). Escucha `Alta`/`Edicion` de `FormTipoTecnicoPrecio` (`:8, 387-428`; eventos en `Form_FormTipoTecnicoPrecio.cls:5-6, 121-123`) y `Alta` de `FormTecnicoTipo` (`:9, 430-447`; `Form_FormTecnicoTipo.cls:5, 62`); en los tres casos recarga la lista con el filtro vigente. `cmdSalir` cierra sin guardar (`:248-251`).

## Controles y eventos

Controles: 13 (13 `Begin` de control en la sección `Detalle`, `.form.txt:297-896`; además hay 8 plantillas de control por defecto en `.form.txt:141-265` que no son controles del formulario).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ListaPreciosPerfil` | lista, 6 columnas, lista de valores | `OnClick` → `ListaPreciosPerfil_Click` | Muestra ID (columna 0 oculta), perfil, fechas y precios; al seleccionar carga el precio y habilita Alta/Editar/Eliminar precio | CAP-9 | `.form.txt:494-553`; `:252-300` |
| `ComboEnActivo` | combo, valores `Sí;No` | `AfterUpdate` → `ComboEnActivo_AfterUpdate` | Filtra por vigencia y recarga la lista | CAP-9 | `.form.txt:638-658`; `:193-216` |
| `ComandoLimpiarEnActivo` (rótulo «ComandoLimpiarEquipo») | botón | `OnClick` | Vacía el combo y recarga sin filtro | CAP-9 | `.form.txt:716-727`; `:167-191` |
| `ComandoAltaPerfil` | botón «Alta Perfil» | `OnClick` | Abre `FormTecnicoTipo` para dar de alta un perfil | CAP-9 | `.form.txt:896-907`; `:11-26` |
| `ComandoAltaPerfilHora` | botón «Alta Nuevo Precio» (deshabilitado al diseñar) | `OnClick` | Abre `FormTipoTecnicoPrecio` en modo alta | CAP-9 | `.form.txt:343-355`; `:28-60` |
| `ComandoEdicionPerfilHora` | botón «Editar Precio» (deshabilitado al diseñar) | `OnClick` | Abre `FormTipoTecnicoPrecio` en modo edición solo si es el último precio | CAP-9 | `.form.txt:410-422`; `:62-106` |
| `ComandoEliminarPerfilHora` | botón «Eliminar Precio» (deshabilitado al diseñar) | `OnClick` | Confirma y borra el precio seleccionado | CAP-9 | `.form.txt:450-463`; `:117-163` |
| `cmdSalir` | botón «Salir» | `OnClick` | Cierra el formulario | CAP-9 | `.form.txt:371-381`; `:248-251` |
| `lblTitulo1`, `lblTitulo2`, `Etiqueta5` | 3 etiquetas («Gestión de Precios Por Perfil» ×2, «Activos») | — | Títulos (`lblTitulo1` da el `Caption` de la ventana) y rótulo del combo | CAP-9 | `.form.txt:611-623, 671-704`; `:230` |
| `CuadroSuperior`, `Cuadro2` | 2 rectángulos | — | Decoración | sin asignar | `.form.txt:297-329` |

Sin control `ComandoEditarPerfil`: el handler vacío `ComandoEditarPerfil_Click` (`:112-114`) no tiene botón asociado en el `.form.txt`.

## Flujos

### Apertura (`Form_Open`, `:220-247`)
1. Cursor de espera, `AjustarTamaño Me` (`:227`; `Funciones Generales.bas:261`), `Caption` = rótulo de `lblTitulo1` (`:230`).
2. `Filtrar` sin argumento (`:231`): lista todos los precios, vigentes o cerrados.
3. Error: `MsgBox` crítico y cierre del formulario (`:240-247`).

### Carga de la lista (`Filtrar`, `:302-383`)
1. Deshabilita Alta/Editar/Eliminar precio y habilita `ComandoAltaPerfil` (`:320-324`); fija como cabecera de la lista una fila fija de textos (`:327`).
2. Vigencia: sin filtro ⇒ todos; «Sí» ⇒ `FechaFinal Is Null`; otro valor ⇒ `Not FechaFinal Is Null` (`:331-340`); orden por `TIPOTECNICO` (`:342-344`).
3. Por cada fila añade `IDTipoTecnicoFecha;TIPOTECNICO;FechaInicial;FechaFinal;PrecioHoraLab;PrecioHoraExtra` con los precios formateados `#,##0.00 €` (`:352-355`).
4. Si queda una sola fila de datos (`ListCount = 2`) la selecciona y simula el clic (`:363-366`); si no, y hay selección previa, repite el clic sobre ella (`:367-374`).

### Selección (`ListaPreciosPerfil_Click`, `:252-300`)
Sin ID en la columna 0 ⇒ descarta la selección (`:265-268`). Si no hay objeto seleccionado o el ID difiere, lo relee con `Constructor.getTipoTecnicoPrecio` (`:269-281`; `Constructor.bas:1980-2033`). Si lo obtiene habilita Alta y Editar (`:286-287`) y Eliminar precio (`:289`).

### Alta de precio (`:28-60`)
Exige selección (`:33-46`), marca `ParaAlta = Sí` en el precio activo (`:48-49`) y abre `FormTipoTecnicoPrecio` (`:53`). Al guardar allí, `TipoTecnicoPrecio.Registrar` cierra el precio vigente del tipo con la fecha del día y inserta el nuevo (`TipoTecnicoPrecio.cls:435-439, 470-486`); si falla, reabre el cerrado (`TipoTecnicoPrecio.cls:511-515`).

### Edición (`:62-106`)
Solo se permite si `EsUltima = Sí` (`:85-88`; `TipoTecnicoPrecio.cls:225-251`); marca `ParaAlta = No` (`:89`) y abre el formulario hijo (`:94`). El hijo llama a `Registrar` (`Form_FormTipoTecnicoPrecio.cls:108-115`).

### Borrado (`:117-163`)
Confirmación con «No» por defecto (`:137-140`); `Borrar` (`:142`) y recarga con el filtro actual (`:150`). Errores: `MsgBox` crítico (`:157-162`).

## Reglas, validaciones y cálculos

- **Solo el último precio de cada tipo** se edita (`:85-88`) o elimina (`TipoTecnicoPrecio.cls:597-600`).
- **No se elimina** si algún técnico tiene el perfil (`TipoTecnicoPrecio.cls:602-609`) ni, cuando no queda precio anterior, si hay eventos o actividades sin facturar (`TipoTecnicoPrecio.cls:617-646`).
- **Efecto del borrado:** `DELETE` del precio y reapertura (`FechaFinal = Null`) del anterior; si no hay anterior, borra el tipo de técnico (`TipoTecnicoPrecio.cls:765-782`; `Abrir` en `TipoTecnicoPrecio.cls:565-567`).
- **Validación de registro** (en `MotivoNoOK`): tipo obligatorio, precios numéricos, sin fecha final en el alta, sin cambiar un tipo no activo, y no cerrar fecha si hay técnicos por facturar (`TipoTecnicoPrecio.cls:287-340`).
- **Formato de importe** en lista: `#,##0.00 €` con `Nz(...,"0")` (`:354-355`).
- No hay comprobación de rol en este formulario (ninguna referencia a administrador en el `.cls`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Carga de la lista | `TbTipoTecnicoPrecios` | SQL dinámico concatenado (`CurrentDb().OpenRecordset`), `Like '*'` | no | `:328-347` |
| Lectura por ID | `TbTipoTecnicoPrecios` | SQL dinámico concatenado | no | `Constructor.bas:1993-1996` |
| Cierre del precio vigente | `TbTipoTecnicoPrecios` | `UPDATE` con fecha literal `#mm/dd/yyyy#` | no | `TipoTecnicoPrecio.cls:436-438` |
| Alta / edición | `TbTipoTecnicoPrecios` | DAO `AddNew`/`Edit` + `Update` | no | `TipoTecnicoPrecio.cls:470-500` |
| Borrado | `TbTipoTecnicoPrecios` | `DELETE *` concatenado | no | `TipoTecnicoPrecio.cls:765-768` |
| Reapertura del anterior | `TbTipoTecnicoPrecios` | `UPDATE ... FechaFinal = Null` | no | `TipoTecnicoPrecio.cls:565-567` |

`getdb` abre el `.accdb` de datos por ruta de red configurada o local, con contraseña literal en el código (`VG.bas:548-576`).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: sintaxis Jet/ACE (`Like '*'`, `DELETE *`, `Not FechaFinal Is Null`, `Format` con `€`); fecha literal `#mm/dd/yyyy#` (`TipoTecnicoPrecio.cls:436`); identificador `IDTipoTecnicoFecha` calculado en VBA (`TipoTecnicoPrecio.cls:101-118`); cierre + alta + borrado + reapertura como operaciones sueltas sin transacción, a envolver en una sola al migrar; mezcla de `CurrentDb` (tablas vinculadas, `:347`) y `getdb` (apertura directa).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComandoAltaPerfilHora` y `ComandoEdicionPerfilHora` usan `Constructor.getTecnico` (`:38, 74`), que busca en `TbTecnicos` por `ALIAS` (`Constructor.bas:1949-1951`) con el ID de la columna 0 y lo asigna a un `TipoTecnicoPrecio`; la rama solo corre si no hay objeto seleccionado, pero sería un error de tipo o una búsqueda equivocada.
- `ListaPreciosPerfil_Click` sale con `Exit Sub` sin quitar el cursor de espera (`:258, 265-268`).
- `Err.Raise 100` (`:284`) no se trata como error propio: el manejador lo sustituye por el mensaje genérico (`:295-297`).
- `ComandoAltaPerfil_Click` y `ComandoAltaPerfilHora_Click` guardan `m_Error` pero no muestran `MsgBox` (`:22-25, 56-59`): el error quedaría silencioso.
- `Form_Open` no reinicia `m_Error` antes de `Filtrar` (`:231-234`).
- Cierre previo + alta + borrado/reapertura sin transacción (`TipoTecnicoPrecio.cls:436-438, 765-782`): un fallo intermedio podría dejar el tipo sin precio vigente.
- `CurrentDb` frente a `getdb` (comentario `TODO`, `:346-347`): la lista y las escrituras podrían apuntar a orígenes distintos.
- Restos de diseño: la lista conserva en el `.form.txt` una lista de valores de ejemplo con datos de técnicos, distinta de la cabecera de ejecución (`.form.txt:512` frente a `:327`); `ComandoLimpiarEnActivo` lleva el rótulo «ComandoLimpiarEquipo» (`.form.txt:726`); handler vacío `ComandoEditarPerfil_Click` (`:112-114`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: tabla con filtro de vigencia en servidor; cierre y alta de vigencia en una sola transacción; identificador generado por la base de datos; confirmación y errores homogéneos.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, la lista muestra todos los precios ordenados por perfil con importes en formato `#,##0.00 €` (`ListaPreciosPerfil`, `:327-355`).
2. «Sí» muestra solo precios sin fecha final; «No», solo los cerrados; el botón de limpiar vuelve a mostrar todos (`ComboEnActivo`, `ComandoLimpiarEnActivo`).
3. Sin selección, Alta/Editar/Eliminar precio están deshabilitados y `ComandoAltaPerfil` habilitado (`:320-324`).
4. Editar un precio que no es el último de su perfil muestra «Solo se puede editar el último perfil del seleccionado» (`:85-88`).
5. Eliminar pide confirmación; con «No» no cambia nada (`:137-140`).
6. Eliminar un precio de un perfil con técnicos asignados muestra «Al menos hay un técnico que tiene este perfil» y no borra (`TipoTecnicoPrecio.cls:602-609`).
7. Tras un alta o edición correctas en el formulario hijo, la lista se recarga con el filtro vigente (`:387-428`).
8. Al dar de alta un precio nuevo, el anterior del mismo perfil queda cerrado con la fecha del día (`TipoTecnicoPrecio.cls:435-439`).

## Pendientes

- `Constructor.getTipoTecnicoActivo` y `TipoTecnicoPrecioEnEventos` (`TipoTecnicoPrecio.cls:430, 620`): consultas sin leer en detalle.
- Efecto de `Abrir` y `ColTecnicos` en el flujo de borrado (`TipoTecnicoPrecio.cls:602, 770`).
- Contenido de `FormTecnicoTipo` más allá del evento `Alta` (`:430`).
- Datos de la columna de vigencia (`TbTipoTecnicoPrecios`): índices y tipos (ERD de origen).
