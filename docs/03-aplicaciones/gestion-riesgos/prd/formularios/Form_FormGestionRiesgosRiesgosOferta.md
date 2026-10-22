# Form_FormGestionRiesgosRiesgosOferta

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormGestionRiesgosRiesgosOferta.cls` (740 líneas) y `src/forms/FormGestionRiesgosRiesgosOferta.form.txt` · **PRD:** [README](../README.md)

Subformulario de **riesgos detectados en la oferta**: lista los riesgos externos asociados a la edición, permite dar de alta uno nuevo —manualmente o eligiéndolo del catálogo de biblioteca— y abrir su ficha, eliminarlo o saltar al riesgo trasladado. Es el subformulario más complejo de los cuatro. Las citas son `Form_FormGestionRiesgosRiesgosOferta.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormGestionRiesgos` | Pestaña «Riesgos de oferta»; el contenedor activa y desactiva la pestaña según el estado del alta | `FormGestionRiesgos.cls:197`, `:328`, `:343`, `FormGestionRiesgos.form.txt:557` |
| `FormRiesgosBibliotecaGestion` | Se suscribe al evento `RiesgoDeBibliotecaElegido` | `:16`, `:681` |
| `FormRiesgoExternoDetalle` | Se suscribe al evento `Registrado` para refrescar la lista | `:18`, `:716` |

La salida es doble: abre `FormRiesgoExternoDetalle` o `FormRiesgoExternoPedido` (`:168-183`) y, si el riesgo tiene asociado, abre `FormRiesgosGestion` (`:342-344`). `ActualizarRiesgoExt` (`:629`) es pública y se ofrece al formulario de detalle para sincronizar la fila ya editada sin recargar todo.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Carga la lista; ante error avisa y **cierra el subformulario** | `:514-540` |
| `ListaRiesgosOferta` | `ListBox` de 5 columnas | `Click`, `DblClick` | Selecciona el riesgo y habilita los comandos; doble clic abre el detalle | `:542-627`, `.form.txt:874`, `:880` (`ColumnCount =5`) |
| `ComboTrasladar` | `ComboBox` | `Change`, `Exit` | Activa o desactiva el motivo de no integración | `:362-380`, `.form.txt:912` |
| `Descripcion`, `FechaDetectado` | `TextBox` | `Exit` | Repiten la misma lógica de habilitación del motivo | `:382-400`, `.form.txt:956`, `:986` |
| `MotivoNoIntegrado`, `CodRiesgoBiblioteca`, `CausaRaiz` | `TextBox` | — | Datos del alta; los habilita o bloquea según el modo | `.form.txt:1017`, `:1047`, `:1078` |
| `ComandoActualizar` | botón | `Click` | Recarga la lista conservando la selección | `:27-78`, `.form.txt:1107` |
| `ComandoAltaRiesgoOferta` | botón | `Click` | Registra un riesgo externo nuevo con origen «Oferta» | `:80-154`, `.form.txt:1332` |
| `ComandoSeleccionarRiesgoDeBiblioteca` | botón | `Click` | Abre el catálogo en modo selección y lo reconfigura | `:276-316`, `.form.txt:1301` |
| `ComandoDetalle`, `ComandoEliminarRiesgoOferta`, `ComandoVerRiesgoTrasladado` | botones | `Click` | Ver, borrar y saltar al riesgo trasladado | `:156-197`, `:199-273`, `:318-360` |

## Flujos

### Apertura (`EstablecerDatos`, `:401-513`)

1. **Deshabilita todos los comandos** y pone `AllowEdits = True` de forma incondicional (`:415-423`).
2. Decide `blnPermitidoEscribir`: solo si el proyecto está activo, el usuario **no** es técnico, hay edición activa y **es la primera edición** (`:426-441`).
3. Si puede escribir, el modo de alta depende de `RequiereRiesgoDeBibliotecaCalculado`: con catálogo obligatorio se habilita elegir de biblioteca y se **bloquea la descripción**; sin catálogo se hace al revés (`:443-461`).
4. Obtiene la edición activa y vuelca la lista de riesgos externos (`:472-498`).
5. Selecciona la primera fila y llama al manejador de selección (`:503-505`).

### Alta (`ComandoAltaRiesgoOferta_Click`, `:80-154`)

1. Crea un `RiesgoExterno` nuevo con la edición activa, la descripción, la fecha de detección, si se traslada y el motivo (`:104-111`).
2. Si hay motivo de no integración, fija la fecha del motivo al día de hoy (`:109-111`).
3. Fija el **origen «Oferta»** y el usuario que registra (`:112-113`).
4. Copia el indicador de riesgo de catálogo calculado por el proyecto y, solo si procede, la causa raíz (`:115-118`).
5. Lo registra (`:120-123`) y añade la fila a la lista (`:129`).
6. Limpia los campos del alta y devuelve el foco al desplegable de traslado (`:133-135`, `:137-139`).

### Detalle (`ComandoDetalle_Click`, `:156-197`)

Abre `FormRiesgoExternoDetalle` salvo que el origen sea «Pedido», en cuyo caso abre `FormRiesgoExternoPedido` (`:168-183`). El código de origen se compara con una cadena literal.

### Eliminación (`ComandoEliminarRiesgoOferta_Click`, `:199-273`)

Pide confirmación (`:219`), valida que el riesgo asociado se pueda borrar (`:231-238`), borra el riesgo externo (`:239-242`), quita la fila (`:245-246`) y recarga el riesgo asociado (`:247-252`). La última parte comprueba si `FormRiesgo` está abierto y lo cierra, pero **no vuelve a abrirlo**: el efecto final es cerrar una ventana que no ha vuelto a abrir.

### Catálogo (`ComandoSeleccionarRiesgoDeBiblioteca_Click`, `:276-316`)

Abre la biblioteca filtrada por esta gestión de riesgos y **reconfigura el formulario ajeno**: muestra «Elegir», oculta «Alta», «Eliminar» y «Ver si usado», fija el filtro de activos y bloquea su limpiado (`:286-300`). Al elegir, copia el código al alta y, según `RiesgoParaRetipificar`, deja la descripción editable o la bloquea (`:688-700`).

## Reglas, validaciones y cálculos

- **Solo la primera edición permite escribir** (`:432-433`). En ediciones posteriores los campos quedan deshabilitados, aunque el alta original siga viva.
- **El catálogo obligatorio bloquea la descripción**: si el riesgo debe venir de biblioteca, el usuario no puede escribirla (`:446-449`).
- **`MotivoNoIntegrado` depende de `ComboTrasladar`**, y esa dependencia está replicada en cuatro manejadores (`:362`, `:372`, `:382`, `:392`) con cuerpo idéntico. Si se elige «No», el motivo se habilita; en cualquier otro caso se deshabilita **y se borra**.
- **`Origen` se fija en «Oferta»** al dar de alta (`:112`), y es lo que decide qué formulario de detalle se abre.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Riesgos de la edición | `m_ObjEdicionActiva.ColRiesgosExternos` | `:483-498` |
| Riesgo externo por identificador | `getRiesgoExterno` | `:564`, `:570` |
| Catálogo de biblioteca | `FormRiesgosBibliotecaGestion` (subformulario abierto) | `:286-300` |
| Alta | `RiesgoExterno.Registrar` | `:123` |
| Baja | `RiesgoExterno.Borrar` y `riesgo.ValidacionEliminar` | `:232-241` |

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** listado de los riesgos externos de la edición con cinco columnas —identificador, código, descripción, si se traslada y código del riesgo trasladado—, con alta manual o asistida por catálogo, consulta de la ficha, borrado con validación en cascada y acceso al riesgo ya trasladado. Solo es editable en la primera edición y por usuarios no técnicos.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-339.** `ListaRiesgosOferta_Click` lanza `Err.Raise 100` (`:585`) mientras que el resto del formulario usa `1000`. El manejador distingue `Err.Number <> 1000`, de modo que **este error no entra en el aviso al administrador** y cae en el `MsgBox` de advertencia con el mensaje ya compuesto, normalmente vacío: el usuario ve un diálogo en blanco.
- **H-GR-340.** `Me.AllowEdits = True` se fija **incondicionalmente** al principio de `EstablecerDatos` (`:423`), antes de calcular `blnPermitidoEscribir`. La protección real depende de que cada control se deshabilite después; si faltara uno, quedaría editable.
- **H-GR-341.** `ComandoActualizar_Click` comprueba `If lst.ListCount > 2` (`:55`) para recuperar la selección: con exactamente dos filas no se restaura la selección, aunque la condición de entrada (`:54`) sí se cumple.
- **H-GR-342.** La etiqueta `SiguienteRiesgoExt:` (`:501`) no está referenciada por ningún `GoTo`: código muerto.
- **H-GR-343.** `ComandoVerRiesgoTrasladado_Click` carga el riesgo en `m_ObjRiesgoDelExt` (`:340`) y declara cinco variables más que nunca usa (`:321-327`); abre `FormRiesgosGestion` (`:343`) sin pasarle el riesgo seleccionado, de modo que no es evident que el formulario muestre el riesgo pedido.
- **H-GR-344.** `ActualizarRiesgoExt` compara líneas completas con `Replace` sobre el `RowSource` (`:670`). Si dos filas son idénticas salvo en un carácter, `Replace` las cambia las dos; y si `m_LineaActual` queda vacía porque no se encontró la fila, se intentaría sustituir una cadena vacía.
- **H-GR-345.** La misma técnica de `Replace` sobre el `RowSource` se usa también en el formulario de suministros (`:262` de ese módulo): dos implements distintas del mismo problema.
- **H-GR-346.** Al abrir la ficha se cierra primero la que ya estuviera abierta con `acSaveNo` (`:170`, `:180`): si el usuario tenía cambios sin guardar en esa ficha, se pierden sin preguntar.
- **H-GR-347.** `Err.Raise 1000` no lleva `vbObjectError` en ninguno de los ~20 puntos de este formulario.
- **H-GR-348.** La comparación del origen con la cadena «Pedido» (`:168`) y la asignación de `Origen = "Oferta"` (`:112`) son literales mágicos repartidos por el código, sin enumeración compartida.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: representar los riesgos como una tabla con columna de acciones en lugar de una lista de cinco columnas; sustituir el catálogo como formulario aparte por un selector dentro de la propia página; modelar «trasladado» y su motivo como una decisión única con motivo obligatorio condicional, en lugar de dos campos acoplados por cuatro manejadores; ocultar por diseño, no por código, el botón de elegir de biblioteca cuando el proyecto no lo requiere; decidir si la edición no editable se comunica con un aviso visible en lugar de con controles deshabilitados.

## Criterios de aceptación

Propuestos, no aprobados:

1. La lista muestra los riesgos de la edición con su código, descripción, indicador de traslado y riesgo asociado.
2. En la primera edición y por usuario no técnico se puede dar de alta un riesgo, manualmente o desde el catálogo.
3. Un riesgo de catálogo obligatorio no admite descripción libre.
4. Elegir «No trasladar» habilita el motivo; cualquier otro valor lo deshabilita y lo limpia.
5. No se puede eliminar un riesgo cuyo riesgo asociado no admita el borrado, y el motivo se explica.
6. En ediciones posteriores o para usuarios técnicos, la vista es de solo lectura.

## Pendientes

- Qué decide exactamente `RequiereRiesgoDeBibliotecaCalculado` y quién lo calcula.
- Si `FormRiesgoExternoPedido` existe en el mismo contenedor o pertenece a otro flujo.
- Si el catálogo debe seguir siendo un formulario aparte o puede reducirse a un selector.
- Cómo se comporta el botón de ver riesgo trasladado cuando el formulario destino no acepta el riesgo como criterio (ver [forms](../../forms.md)).