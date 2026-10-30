# Form_FormEventoEqMedidaEdicion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormEventoEqMedidaEdicion.cls` (306 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ventana auxiliar de la **edición** de un evento ya existente: lista los equipos de medida vinculados al evento activo (`m_ObjEventoActivo`) y permite añadir y desvincular equipos **escribiendo directamente en la base de datos** a través de `Evento.RegistrarEquipoMedida` y `Evento.EliminarEquipoMedida`. Es un formulario **no enlazado** (sin `RecordSource`). Capacidad: CAP-3 Eventos (equipos de medida del evento). Abreviatura: `FG.bas` = `modules/Funciones Generales.bas`. Las citas `:N` sin fichero son de `Form_FormEventoEqMedidaEdicion.cls`. Es la variante de edición de `Form_FormEventoEqMedida`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEventoEdicion` | `ComandoEquiposDeMedida_Click`: cierra el formulario si estaba abierto, lo abre y se engancha a sus eventos | `Form_FormEventoEdicion.cls:256-262` |

Requiere `m_ObjEventoActivo` (si es `Nothing`: «No hay evento activo» y se cierra, `:186-189, 207-214`). Eventos hacia el llamador: `Alta` (`:9`, lanzado en `:40`) y `Eliminado` (`:10`, lanzado en `:98`); el llamador solo reescribe el rótulo del botón con el recuento de `ColEquiposMedida` (`Form_FormEventoEdicion.cls:1254-1280, 1282-1304`). Sale con `cmdSalir` (`:12-15`).

## Controles y eventos

Controles: 8 (20 `Begin` en el `.form.txt`: 8 plantillas de tipo por defecto, `Form`, `FormHeader`, `Section` y `FormFooter` más los 8 controles).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComboEquipoDeMedida` | combo, lista de valores de 3 columnas (ID oculto) | `AfterUpdate` → `ComboEquipoDeMedida_AfterUpdate` | Resuelve el equipo, avisa si está fuera de calibración y llama a `ComandoAlta_Click` | CAP-3 | `:119-170`; `Form_FormEventoEqMedidaEdicion.form.txt:418-470` |
| `ListaEventoEqMedida` | lista de valores de 3 columnas (ID oculto) | `OnClick` → `ListaEventoEqMedida_Click` | Selecciona la vinculación y habilita `ComandoEliminar` si el evento es modificable | CAP-3 | `:262-307`; `Form_FormEventoEqMedidaEdicion.form.txt:353-390` |
| `ComandoAlta` («+») | botón | `OnClick` | Vincula el equipo del combo al evento | CAP-3 | `:17-58`; `Form_FormEventoEqMedidaEdicion.form.txt:471-508` |
| `ComandoEliminar` («Eliminar»; deshabilitado de diseño) | botón | `OnClick` | Pide confirmación y desvincula | CAP-3 | `:60-117`; `Form_FormEventoEqMedidaEdicion.form.txt:509-542` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | CAP-3 | `:12-15`; `Form_FormEventoEqMedidaEdicion.form.txt:554-563` |
| `lblTitulo1`, `lblTitulo2`, `Etiqueta13` | etiquetas | — | Título (`lblTitulo1` pasa a `Caption` del formulario) y rótulo «EQ. Medida» | CAP-3 | `:201`; `Form_FormEventoEqMedidaEdicion.form.txt:309, 329, 402` |

## Flujos

### Apertura (`Form_Load`, `:174-214`)

1. `AjustarTamaño Me` (`:184`), comprueba el evento activo y vacía el `RowSource` del combo (`:186-191`).
2. `RellenarListaEquiposMedida` carga todos los equipos del catálogo (`:192`; `FG.bas:7293-7322`).
3. `EstablecerDatos` rellena la lista con `m_ObjEventoActivo.ColEquiposMedida`; cada fila es «IDEventoEquipoMedida;nombre;Calibrado|Sin Calibrar», según `IDCalibracion` sea numérico (`:216-252`).

### Alta (`ComboEquipoDeMedida_AfterUpdate` → `ComandoAlta_Click`)

1. Combo vacío: deshabilita `ComandoAlta` (`:126-133`). Evento no modificable: sale sin mensaje (`:134-139`).
2. Relee el equipo (`:140-146`), habilita `ComandoAlta` (`:153`), avisa «Equipo fuera de calibración» si `Calibrado = No` (`:158-160`) y llama a `ComandoAlta_Click` (`:161`).
3. `ComandoAlta_Click` llama a `m_ObjEventoActivo.RegistrarEquipoMedida` (`:36`), lanza `Alta`, vacía el combo y repinta (`:40-43`).
4. `RegistrarEquipoMedida` rechaza si el evento está facturado o franqueado, busca la calibración vigente a la fecha con `getCalibracionPorFecha` y, si el vínculo no existía, inserta el registro con DAO (`Evento.cls:1879-1918`). Si ya existía, sale sin error (`Evento.cls:1901-1905`).

### Eliminación (`ComandoEliminar_Click`, `:60-117`)

1. Exige `Modificable = Sí` («El evento no es modificable», `:70-73`) y una fila seleccionada (`:74-83`).
2. `MsgBox` de confirmación con «No» por defecto (`:87-90`). Con «Sí», `EliminarEquipoMedida` borra el vínculo y lanza `Eliminado` (`:94-98`; `Evento.cls:1941-1956`).

## Reglas, validaciones y cálculos

- **Evento modificable:** `Modificable = No` si hay fecha de facturado o fecha final (`Evento.cls:428-462`); gobierna `ComandoEliminar` (`:70, 279`) y la entrada del combo (`:134`).
- **No registrar si facturado o franqueado:** `Evento.cls:1879-1886`.
- **No eliminar si no es editable:** repite la comprobación en la clase (`Evento.cls:1942-1945`).
- **Calibración asociada al registrar:** la vigente a la fecha, si existe (`Evento.cls:1888-1891`); el aviso del formulario usa `Calibrado` (`:158`).
- **Identificador del vínculo:** `DameID("TbEventosEquipoMedida", ...)` (`Evento.cls:1906`), según su cabecera el máximo más uno (`FG.bas:371-390`).
- **Sin duplicados:** consulta previa por evento y equipo (`Evento.cls:1895-1904`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Equipos del evento | `TbEventosEquipoMedida` | SQL dinámico concatenado (`WHERE IDEvento='<id>'`) | no | `Constructor.bas:4965-4981` |
| Equipo por ID | `TbEquiposMedida` | SQL dinámico concatenado | no | `Constructor.bas:4878-4896` |
| Vínculo por ID | `TbEventosEquipoMedida` | SQL dinámico concatenado | no | `Constructor.bas:4922-4940` |
| Existencia del vínculo | `TbEventosEquipoMedida` | SQL dinámico concatenado | no | `Evento.cls:1895-1898` |
| Alta del vínculo | `TbEventosEquipoMedida` | DAO `AddNew`/`Update` | no | `Evento.cls:1907-1914` |
| Baja del vínculo | `TbEventosEquipoMedida` | SQL dinámico `DELETE` con `Execute` | no | `Evento.cls:1951-1954` |

## Coexistencia

Adaptar: el `DELETE  *` (`Evento.cls:1951`) usa sintaxis Jet/ACE que PostgreSQL rechaza (`DELETE FROM`); el ID del vínculo se calcula en VBA con máximo más uno, sin transacción, y chocaría con una secuencia o identidad de PostgreSQL; la restricción por `IDEvento` entrecomillado frente a `IDEquipoMedida` numérico exige respetar los tipos de las columnas vinculadas.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `:147-153`: si el equipo no se resuelve, se deshabilita `ComandoAlta` pero a continuación se vuelve a habilitar (sin `Else`) y se desreferencia `m_EquipoEnCombo` en `:158`.
- `ComandoAlta_Click` declara un `m_EquipoEnCombo` local (`:19`) que sombrea al de módulo (`:6`).
- Si el vínculo ya existe, `RegistrarEquipoMedida` sale sin error (`Evento.cls:1901-1905`) y el formulario lanza `Alta` igualmente (`:40`).
- El alta no comprueba `Modificable` (solo facturado y franqueado en la clase, `Evento.cls:1879-1886`): un evento con fecha final sin franquear podría admitir altas, pero no bajas (a verificar).
- Con evento no modificable, el combo no avisa de nada (`:134-139`).
- La baja usa `DELETE  *` y `Execute` sin transacción (`Evento.cls:1951-1954`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: mostrar el motivo cuando el evento no admite cambios y unificar la regla de modificabilidad entre alta y baja.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin evento activo, el formulario muestra «No hay evento activo» y se cierra (`:186-189`).
2. La lista muestra cada equipo vinculado con «Calibrado» o «Sin Calibrar» (`ListaEventoEqMedida`, `:243-249`).
3. Elegir un equipo en un evento facturado o franqueado no lo vincula y muestra el motivo (`:36`; `Evento.cls:1879-1886`).
4. Elegir un equipo fuera de calibración muestra el aviso y lo vincula (`:158-161`).
5. «Eliminar» exige confirmación y, con «No», no cambia nada (`ComandoEliminar`, `:87-90`).
6. Con evento no modificable, «Eliminar» permanece deshabilitado (`:279-284`).

## Pendientes

- Contenido de `getCalibracionPorFecha` y de `DameID` (`Evento.cls:1888, 1906`; `FG.bas:371`).
- Cuerpo de `FechaFacturado` y `FechaFinal` (`Evento.cls:651`, `Evento.cls:440-455`).
- Quién escribe la columna `IDCalibracion` de `ColEquiposMedida` (`:243`).
