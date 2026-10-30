# Form_FormSubSistemaGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormSubSistemaGestion.cls` (384 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Gestión del catálogo de subsistemas por BUI (`TbSubsistemaBui`): el usuario elige un BUI en un combo y ve en una lista los subsistemas de ese BUI, con altas y ediciones delegadas en otros formularios, baja directa y exportación a Excel. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); el combo se alimenta de `TbBUI` (`Form_FormSubSistemaGestion.form.txt:1223`) y la lista de `TbSubsistemaBui` (`:359-360`). Las citas son `Form_FormSubSistemaGestion.cls:línea` salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`. Capacidad: CAP-12 Estructura y catálogos.

Controles: 16 (25 `Begin` en el `.form.txt`: 16 controles, `Begin Form`, `Begin Section` y 7 plantillas de estilo de las líneas 731-854).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpcionesOtras` | `DoCmd.OpenForm "FormSubSistemaGestion"` desde `lblGestionSubSistemas_Click` | `Form_Form0BDOpcionesOtras.cls:101-104` |
| `Form_FormSubSistemaAlta` | tras guardar, si este formulario está abierto, llama a `Form_FormSubSistemaGestion.ActualizarLista` | `Form_FormSubSistemaAlta.cls:28-29` |
| `Form_FormSubSistemaEdicion` | tras guardar, hace `Requery` de `ListaSubSistema` | `Form_FormSubSistemaEdicion.cls:60-61, 107-108` |

Hacia otros formularios: abre `FormSubSistemaAlta` con el BUI como `OpenArgs` (`:25`) y `FormSubSistemaEdicion` con `BUI|SUBSISTEMA` (`:62`); cierra antes la instancia previa con `FormularioAbierto` (`:22-24, 59-61`; `FG.bas:171-190`). Sale con `cmdSalir` (`:205-208`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComboBUI` | combo | `BeforeUpdate` → `ComboBui_BeforeUpdate` | Habilita `ComandoAlta` según haya BUI, rehace la lista y guarda `strBUIAnterior` | CAP-12 | `:165-189`, `.form.txt:1221-1223` |
| `ListaSubSistema` | lista (una columna) | `OnClick` → `ListaSubSistema_Click` | Habilita Editar y Eliminar si hay fila seleccionada | CAP-12 | `:209-219`, `.form.txt:948-953` |
| `ComandoAlta` («Alta») | botón, deshabilitado en diseño | `OnClick` | Exige BUI y abre `FormSubSistemaAlta` | CAP-12 | `:6-36`, `.form.txt:972-984` |
| `ComandoEditar` («Editar») | botón, deshabilitado en diseño | `OnClick` | Exige subsistema y BUI y abre `FormSubSistemaEdicion` | CAP-12 | `:38-73`, `.form.txt:1059-1070` |
| `ComandoEliminar` («Eliminar») | botón, deshabilitado en diseño | `OnClick` | Confirma y llama a `Eliminar` | CAP-12 | `:74-124`, `.form.txt:1103-1116` |
| `ComandoExportarExcel` | botón, deshabilitado en diseño | `OnClick` | Exporta el `RowSource` de la lista a `.xlsx` con `GenerarConsultasSQL` | CAP-12 | `:126-148`, `.form.txt:1500-1511` |
| `ComandoLimpiarBUI` | botón | `OnClick` | Vacía el combo y rehace la lista | CAP-12 | `:149-163`, `.form.txt:1340-1342` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | CAP-12 | `:205-208`, `.form.txt:1025-1027` |
| `ComandoAyudaFormulario` | botón con icono | `OnClick = [Event Procedure]` | El `.cls` no define el handler: no hace nada | CAP-12 | `.form.txt:1157-1159` |
| `lblTitulo1`, `lblTitulo2`, `Etiqueta118` | etiquetas | — | `lblTitulo1` pasa a `Caption` del formulario (`:198`); `lblTitulo2` es el título duplicado; `Etiqueta118` rotula «BUI» | CAP-12 | `.form.txt:1252, 1300, 1320` |
| `Cuadro2`, `CuadroFormulario`, `CuadroSuperior`, `Marco231` | 3 rectángulos y 1 grupo de opciones | — | Decoración | sin asignar | `.form.txt:875, 898, 921, 1277` |

## Flujos

### Apertura (`Form_Open`, `:192-204`)

`On Error Resume Next`; `AjustarTamaño Me` (`FG.bas:261-285`), título de ventana desde `lblTitulo1` y repintado (`:193-203`). No rellena la lista: queda con el `RowSource` de diseño (`.form.txt:950`), y `ComandoAlta`, `ComandoEditar`, `ComandoEliminar` y `ComandoExportarExcel` arrancan deshabilitados (`.form.txt:972, 1059, 1103, 1500`).

### Filtrado por BUI (`ComboBui_BeforeUpdate` → `ActualizarLista`, `:165-189, 293-342`)

1. `ComandoAlta` se habilita si el valor del combo no está vacío (`:171-176`) y se llama a `ActualizarLista` (`:177`).
2. `ActualizarLista` deshabilita Editar, Eliminar y Exportar (`:303-305`), obtiene la SQL de `Filtrar` (`:311-318`) y la asigna a `lst.RowSource`; con la lista vacía deja Exportar deshabilitado y sale (`:321-327`); si no, habilita Exportar, selecciona la primera fila y llama a `ListaSubSistema_Click` (`:328-330`).
3. `Filtrar` sin BUI devuelve una SQL con `BUI='-1'` (lista vacía, `:361-366`); con BUI, filtra por igualdad (`:370-375`).

### Alta y edición (`:6-73`)

Alta: sin BUI, «No se ha indicado para que BUI es el Subsistema» (`:12-16`). Edición: exige subsistema (`:46-49`) y BUI (`:50-53`). El subsistema sale de `ListaSubSistema.Column(0)` (`:44`). El formulario destino refresca la lista al guardar (véase Entrada).

### Baja (`ComandoEliminar_Click`, `:74-124`; `Eliminar`, `:220-292`)

1. Valida selección y BUI (`:79-88`) y pide confirmación con «No» por defecto (`:93-97`).
2. `Eliminar` valida los dos argumentos (`:244-251`), consulta `TbEventos` por `SUBSISTEMA` con `Dame` (`:257`) y, si existe algún evento, rechaza con «El Subsistema está incluido en al menos un evento» (`:263-266`).
3. Abre un `Recordset` sobre `TbSubsistemaBui` con SQL concatenada (`:267-271`); si no hay fila, error (`:273-276`); si la hay, `.Delete` (`:277`).
4. Un resultado con `|` se trata como error (`:107-111`); si no, devuelve el foco a `ComandoAlta` y rehace la lista (`:112-113`).

### Exportación (`ComandoExportarExcel_Click`, `:126-148`)

Pasa `ListaSubSistema.RowSource` a `GenerarConsultasSQL` (`:132-133`), que escribe un `.xlsx` temporal en el directorio local configurado, rechaza una consulta sin registros y lo abre con `Ejecutar` (`FG.bas:1584-1672`).

### Errores

Los handlers con `On Error GoTo errores` muestran un `MsgBox` crítico con el texto acumulado; el error interno 1000 conserva el texto propio (`:30-35, 67-72, 142-147, 183-188`). No envían correo.

## Reglas, validaciones y cálculos

- **Integridad en la baja:** un subsistema con eventos no se elimina (`:257-266`). La comprobación es solo por nombre de subsistema, sin BUI (`:257`).
- **Selección obligatoria:** subsistema y BUI para editar y eliminar (`:46-53, 81-88`); BUI para dar de alta (`:13-16`).
- **Lista vacía sin BUI:** el filtro `BUI='-1'` equivale a no mostrar nada (`:362`).
- **Confirmación:** «¿Desea realmente eliminar el SubSistema: …?» con «No» por defecto (`:93-94`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Combo de BUI | `TbBUI` | `RowSource` `SELECT DISTINCT` | — | `.form.txt:1223` |
| Lista de subsistemas | `TbSubsistemaBui` | `RowSource` con SQL dinámico concatenada con el BUI entre comillas | — | `:359-375` |
| Comprobación de uso | `TbEventos` | `Dame`: SQL concatenada, primero numérica y, si falla, con comillas | no | `:257`; `FG.bas:290-318` |
| Baja | `TbSubsistemaBui` | DAO `OpenRecordset` + `Delete` sobre SQL concatenada | no | `:267-277` |
| Exportación | resultado de la lista | DAO `OpenRecordset` sobre el `RowSource` | — | `FG.bas:1619` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL concatenada con comillas simples, comodín `Like '*'` (código inactivo, `:368`), `Dame` que reintenta con comillas ante error, y exportación mediante automatización de Excel y apertura del fichero; en PostgreSQL vinculado habrá que revisar el reintento por error de `Dame` y los tipos de `SUBSISTEMA`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComandoEliminar_Click` no declara `On Error GoTo errores` (`:74-124`) aunque tiene la etiqueta (`:118`); un `Err.Raise 1000` en `:83, 87, 110` se propagaría sin pasar por el `MsgBox`.
- `ComboBui_BeforeUpdate` consulta `Me.ComboBui` (`:171, 311`) en un evento previo a la actualización: podría filtrar por el valor anterior y no por el elegido. La variable `strBUIAnterior` (`:5, 178`) no se lee en ningún otro punto.
- La baja comprueba eventos solo por `SUBSISTEMA` (`:257`): el mismo nombre usado en otro BUI bloquearía la baja.
- Los mensajes de baja dicen «editar» (`:82, 86`).
- Si `Dame` devuelve el texto de error `#ERR|…`, la detección por `|` también atraparía un nombre de subsistema con barra vertical (`:258, 107`).
- `Filtrar` repite la condición `strBUI = ""` (`:361, 367`), por lo que la rama `Like '*'` es inalcanzable.
- Vaciar el combo con `ComandoLimpiarBUI` no deshabilita `ComandoAlta` (`:149-163`).
- `GenerarConsultasSQL` vuelca las fechas como `mm/dd/yyyy` (`FG.bas:1652`) y en la primera vuelta del bucle solo escribe las cabeceras y avanza con `MoveNext` (`FG.bas:1639-1646, 1665`), por lo que el primer registro podría no exportarse.
- `ComandoAyudaFormulario` tiene evento enlazado sin código (`.form.txt:1159`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: restricción de clave foránea entre subsistema y eventos, y borrado atómico con la comprobación.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin BUI elegido, la lista está vacía y «Alta» deshabilitado (`ComboBUI`, `:165-176`).
2. Al elegir un BUI, la lista muestra solo los subsistemas de ese BUI y selecciona el primero (`:311-330`).
3. «Editar» y «Eliminar» solo se habilitan con un subsistema seleccionado (`:209-219`).
4. Eliminar un subsistema con eventos muestra «El Subsistema está incluido en al menos un evento» y no borra (`:263-266`).
5. Responder «No» a la confirmación no borra (`:95-97`).
6. «Exportar» genera un fichero con los subsistemas listados y avisa si no hay registros (`FG.bas:1620-1622`).
7. «Limpiar» vacía el combo y la lista (`:158-159`).

## Pendientes

- `Ejecutar` es la declaración de `ShellExecuteA` (`Variables Globales.bas:15, 39`); `FicheroAbierto` en `FG.bas:195`; su cuerpo no se ha leído.
- Si `SUBSISTEMA` y `BUI` son clave de `TbSubsistemaBui` y su tipo (ERD de origen).
- Contenido de `FormSubSistemaAlta` y `FormSubSistemaEdicion` (fichas propias).
