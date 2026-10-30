# Form_FormSubContratacionEdicion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormSubContratacionEdicion.cls` (495 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Edición (o consulta, si ya está facturada) de una subcontratación de `TbSubcontrataciones`: empresa, fecha de alta, descripción, importe y fecha de fin. Es un formulario **no enlazado** (sin `RecordSource` ni `ControlSource` en el `.form.txt`): recibe el identificador por `OpenArgs`, lee el registro por DAO y lo escribe con la función pública `editar` del propio formulario. Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`, `VG.bas` = `modules/Variables Globales.bas`. Las citas son `Form_FormSubContratacionEdicion.cls:línea` (escrito `:línea`) salvo indicación; `.form.txt` = `Form_FormSubContratacionEdicion.form.txt`.

Controles: 22 (31 `Begin` en el `.form.txt`: 22 controles, 7 plantillas de tipo por defecto, `Form` y `Section`).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSubContratacionGestion` (botón de edición) | Toma el ID de la columna 0 de `ListaFiltrados` y la marca de facturada de la columna 5 (no vacía ⇒ `strEstaFacturado="Sí"`, `strEscritura="No"`; vacía ⇒ `"No"`/`"Sí"`); cierra el formulario si estaba abierto y lo abre con `OpenArgs = ID|facturada|escritura` | `Form_FormSubContratacionGestion.cls:231-259` |

Salidas: no lanza eventos ni deja globales de resultado al volver. Hacia `FormAnexos` abre con `OpenArgs` = nombre de este formulario y deja la subcontratación en la global `m_ObjEntidadParaAnexoActiva` (`:121-124`; `VG.bas:252`). Tras registrar con éxito muestra «Edición correcta» y se cierra (`:205-206`); `cmdSalir` cierra siempre (`:76`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `FechaAlta` | texto no enlazado, formato fecha corta | — | Fecha de alta; obligatoria y válida al guardar | CAP-8 | `.form.txt:365-366`; `:427-430` |
| `ComboEmpresa` | combo, `Table/Query` | — | Empresas distintas ya existentes en `TbSubcontrataciones` (`SELECT DISTINCT … Empresa … ORDER BY`); obligatoria al guardar | CAP-8 | `.form.txt:400-403`; `:423-426` |
| `Descripcion` | texto no enlazado | — | Descripción opcional | CAP-8 | `.form.txt:484`; `:157, 461-465` |
| `Importe` | texto no enlazado, formato `#,##0.00 €;-#,##0.00 €` | — | Importe; obligatorio y numérico al guardar | CAP-8 | `.form.txt:540-541`; `:431-434` |
| `FechaFin` | texto no enlazado, formato fecha corta | — | Fecha de fin opcional; si se informa, ≥ `FechaAlta` | CAP-8 | `.form.txt:596-597`; `:435-440` |
| `ImagenAnexo` | imagen | `OnClick` → `ImagenAnexo_Click` | Abre en local el anexo con `AbrirEnLocal(strURLAnexo, Me.hWnd)` (`FG.bas:218`); visible solo si hay URL de anexos | CAP-11 | `:373-396`; `:356-360`; `.form.txt:623-624` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` → `ComandoRegistrar_Click` | Guarda y cierra; sin cambios, «No hay cambios que registrar» | CAP-8 | `:180-214`; `.form.txt:780-782` |
| `ComandoAdjuntar` («Adjuntar») | botón | `OnClick` → `ComandoAdjuntar_Click` | Guarda si hay cambios y abre `FormAnexos` | CAP-11 | `:88-135`; `.form.txt:1000-1002` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra; si hay cambios y hay escritura, pregunta si guardar | CAP-8 | `:49-87`; `.form.txt:818-820` |
| `lblTitulo1` | etiqueta | — | Título: «EDICIÓN DE SUBCONTRATACIÓN: » o «DETALLE DE SUBCONTRATACIÓN: » + ID con formato `000`; copiado a `Caption` | CAP-8 | `:223, 321, 332`; `.form.txt:881-882` |
| `lblFacturada` | etiqueta («Facturado») | — | Visible solo si `strEstaFacturado="Sí"`; en diseño `Visible` no es por defecto | CAP-8 | `:344-348`; `.form.txt:960-973` |
| Agrupados sin evento: `lblTitulo2`, `Etiqueta142`, `Etiqueta198`, `Etiqueta201`, `Etiqueta219`, `Etiqueta222` | 6 etiquetas | — | Título secundario y rótulos de los campos | CAP-8 | `.form.txt:326-338, 418-430, 445-457, 500-512, 557-569, 849-860` |
| Agrupados sin evento: `Etiqueta4`, `Etiqueta1`, `Etiqueta2` | 3 etiquetas «*» | — | Marca visual de campos obligatorios | CAP-8 | `.form.txt:894-903, 916-925, 938-947` |
| Agrupados sin evento: `CuadroSuperior`, `Cuadro2` | 2 rectángulos | — | Decoración | CAP-8 | `.form.txt:280-289, 302-312` |

Recuento por tipo: 11 etiquetas, 4 cuadros de texto, 1 combo, 1 imagen, 3 botones, 2 rectángulos = 22.

## Flujos

### Apertura (`Form_Open`, `:215-272`)

1. `AjustarTamaño Me` (`FG.bas:261`) y `Caption` = título (`:222-223`).
2. Exige `OpenArgs` con exactamente tres partes separadas por `|`; si no, «No se han cargado todos los datos» y cierra el formulario (`:225-234, 265-271`). El ID vacío también aborta (`:238-241`).
3. Normaliza: `strEstaFacturado` distinto de «Sí»/«No» ⇒ «Sí»; `strEscritura` distinto de «Sí»/«No» ⇒ «No» (`:242-247`).
4. `EstablecerDatos` (`:274-372`): lee `TbSubcontrataciones` por ID (SQL concatenado, `:282-285`); sin fila, error (`:287-290`). Vuelca empresa, fechas (a `dd/mm/yyyy`), descripción e importe (`#,##0.00 €`) en los controles solo si hay valor (`:299-319`) y guarda copia «al inicio» para detectar cambios (`:291-295`).
5. Si no está facturada: título de edición, combos y textos con `lngColorCampoEditable`, `Locked = False` y botones Registrar/Adjuntar activos (`:320-330`). Si está facturada: título de detalle, foco en `cmdSalir`, campos con `lngColorCampoNoEditable` y `Locked = True`, botones desactivados (`:331-343`; colores en `VG.bas:350-351`).
6. Obtiene la subcontratación con `Constructor.getSubContratacion` (`:350`; `Constructor.bas:3846-3885`) y muestra `ImagenAnexo` según `URLAnexos` (`:355-360`; `Subcontratacion.cls:112-146`).
7. Cualquier error: `MsgBox` crítico y cierre (`:265-271`).

### Guardado (`ComandoRegistrar_Click`, `:180-214`; `Registrar`, `:137-179`; `editar`, `:397-491`)

1. Sin cambios (`HaHabidoCambios`, `:7-48`) ⇒ «No hay cambios que registrar» (`:187-190`).
2. `Registrar` lee los controles con `Nz`; descripción y fecha de fin vacías se envían como `"#"` (`:155-159`) para que `editar` las deje en `Null` (`:461-472`).
3. `editar` valida (ver reglas), abre el registro por DAO, `Edit`/`Update` (`:445-476`) y devuelve «Editada correctamente» (`:480`).
4. Éxito: «Edición correcta» y cierre (`:205-206`). Error: `MsgBox` crítico y el formulario sigue abierto (`:208-213`).

### Salir (`cmdSalir_Click`, `:49-87`)

Con cambios **y** `strEscritura="Sí"` pregunta «¿Desea guardar los cambios producidos?»; «Sí» guarda y luego cierra; «No» cierra sin guardar (`:53-76`). Un fallo al guardar muestra el error y no cierra (`:68-69, 81-86`).

### Adjuntar (`ComandoAdjuntar_Click`, `:88-135`)

Si hay cambios los guarda sin preguntar (`:99-112`); relee la subcontratación (`:113`), cierra `FormAnexos` si está abierto (`:118-120`), la deja en `m_ObjEntidadParaAnexoActiva` y lo abre (`:121-124`).

## Reglas, validaciones y cálculos

- **Empresa obligatoria**: vacía o `"#"` ⇒ «Se ha de introducir el nombre de la empresa subcontratada» (`:423-426`).
- **Fecha de alta** válida (`IsDate`) (`:427-430`); **importe** numérico (`IsNumeric`) (`:431-434`).
- **Fecha de fin** opcional; si existe debe ser ≥ fecha de alta (`:435-440`).
- **Facturada**: `editar` rechaza el guardado con «Ya está facturada, por lo que no se permiten cambios» (`:441-444`); la facturación la decide el formulario llamante (columna 5 de `ListaFiltrados`) y no se vuelve a consultar aquí (`Form_FormSubContratacionGestion.cls:238-244`).
- **Detección de cambios**: compara texto de los cinco campos con la copia inicial; fechas formateadas `dd/mm/yyyy`; ante cualquier error devuelve «hay cambios» (`:25-47`).
- **Sin transacción** ni comprobación de duplicados de empresa/fecha.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Carga del registro | `TbSubcontrataciones` | SQL dinámico concatenado (ID sin comillas), `CurrentDb().OpenRecordset` | no | `:282-285` |
| Relectura para anexos | `TbSubcontrataciones` | SQL dinámico concatenado | no | `Constructor.bas:3859-3861` |
| Lista de empresas | `TbSubcontrataciones` | `RowSource` guardado en el control | no | `.form.txt:402-403` |
| Edición | `TbSubcontrataciones` | DAO `Edit`/`Update` sobre SELECT concatenado | no | `:445-476` |
| URL de anexos | tablas de anexos (vía `ColAnexos`) | no resuelto aquí | — | `Subcontratacion.cls:112-146` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: ID numérico concatenado en SQL (`:284, 447`); `FechaAlta`, `FechaFin` e `Importe` se asignan desde cadena a campos fecha/moneda de Access (`:456, 467, 474`), lo que depende de la configuración regional y de la conversión implícita; `Nz` y `Format` con formato monetario `€` (`:318`); `Max`/autonumérico no intervienen (solo edición).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El bloqueo por facturación depende solo del argumento `OpenArgs`; `editar` lee la variable de módulo `strEstaFacturado` (`:441`), no la base. Con apertura sin partes válidas se aborta, pero un ID facturado tras abrir no se detecta.
- `strEscritura` solo se usa en `cmdSalir` (`:53`); `ComandoAdjuntar` y `ComandoRegistrar` guardan sin consultarlo (`:99, 187`), aunque con facturada quedan deshabilitados (`:341-342`).
- `HaHabidoCambios` trata `Importe` con `IsDate` en lugar de `IsNumeric` (`:36-38`): un importe como «1,5» puede interpretarse como fecha y convertirse con `CDbl`, y el valor inicial no se normaliza igual (`:295`), con lo que podría marcar cambios inexistentes.
- `Importe` se inicializa como `Format(…, "#,##0.00 €")` (`:318`) y `HaHabidoCambios` compara con `strImporteAlInicio` sin formatear (`:39-40, 295`): podría informar cambios espurios al salir.
- `Descripcion` vacía al leer se omite del control (`:308`) y al guardar se envía `"#"` (`:157`), pero un texto literal «#» también se guardaría como `Null` (`:461-464`).
- `ComandoAdjuntar` guarda sin confirmar si hay cambios (`:99-112`), mientras `cmdSalir` sí pregunta.
- `Importe`: un texto con separadores regionales podría fallar en `IsNumeric` o convertirse distinto de lo esperado (`:431, 474`).
- Rótulo de diseño: `lblTitulo1`/`lblTitulo2` con «Edición de Subcontratación» incluso en modo detalle hasta que `Form_Open` lo cambia (`.form.txt:860, 882`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: validar en servidor la facturación en vez de confiar en el argumento; unificar la confirmación de guardado al adjuntar; tipar importe y fechas.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con subcontratación facturada, el formulario muestra «DETALLE…», «Facturado», campos bloqueados y Registrar/Adjuntar desactivados (`:331-348`).
2. Con subcontratación no facturada, los campos son editables (`:320-330`).
3. Guardar sin empresa, sin fecha válida o con importe no numérico muestra el error correspondiente y no escribe (`:423-434`).
4. Una fecha de fin anterior a la de alta se rechaza (`:435-440`).
5. Vaciar descripción o fecha de fin las deja en blanco en el registro (`:461-472`).
6. Guardar sin cambios muestra «No hay cambios que registrar» (`:187-190`).
7. Salir con cambios pregunta si guardar; «No» cierra sin escribir (`:53-76`).
8. `ImagenAnexo` solo aparece si la subcontratación tiene anexos y los abre en local (`:355-360, 382`).
9. Adjuntar con cambios los guarda y abre `FormAnexos` con la subcontratación (`:99-124`).

## Pendientes

- Contenido de `AjustarTamaño` (`FG.bas:261`), `AbrirEnLocal` (`FG.bas:218`) y de `ColAnexos` en `Subcontratacion.cls` (`:112-146`).
- Origen de la marca de facturada en `Form_FormSubContratacionGestion` (columna 5 de `ListaFiltrados`): consulta de origen no revisada (`Form_FormSubContratacionGestion.cls:238`).
- `pregunta`, `flag`, `dato` y `m_SQL` son globales (`VG.bas:203-206`); el formulario las comparte con el resto de la aplicación.
- Tipos y restricciones de columnas de `TbSubcontrataciones` (ERD de origen).
