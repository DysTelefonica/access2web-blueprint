# Form_FormActividadTipoGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormActividadTipoGestion.cls` (180 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pantalla de gestión del catálogo de tipos de actividad (`TbCodActividad`): lista, alta, edición y borrado. Es un formulario **no enlazado**; los datos vienen del `RowSource` de `ListaTiposDeActividad`. `FG.bas` = `modules/Funciones Generales.bas`. Las citas son `Form_FormActividadTipoGestion.cls:línea` salvo indicación. Capacidad de todos los controles: CAP-4 Actividades y tipos de actividad.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form0BDOpcionesOtras` | `DoCmd.OpenForm "FormActividadTipoGestion"` | `Form_Form0BDOpcionesOtras.cls:112` |

Abre `FormActividadTipoAlta` (`:7`) y `FormActividadTipoEdicion` con el código seleccionado en `OpenArgs` (`:29`). Ambos refrescan la lista al guardar (`Form_FormActividadTipoAlta.cls:28-29`; `Form_FormActividadTipoEdicion.cls:111-113`).

Controles: 11 (18 `Begin` en el `.form.txt`: 1 `Form`, 5 estilos por defecto, 1 `Section` y 11 controles).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaTiposDeActividad` | lista, 2 columnas | `OnClick` → `ListaTiposDeActividad_click` | Muestra `Código` y `Tipo` de `TbCodActividad`; al pulsar guarda el código (columna 0) y habilita Editar/Eliminar | `:104-114`; `.form.txt:543-563` |
| `ComandoAlta` («Alta») | botón | `OnClick` → `ComandoAlta_Click` | Abre `FormActividadTipoAlta` | `:5-8`; `.form.txt:299-301` |
| `ComandoEditar` («Editar») | botón, deshabilitado al abrir | `OnClick` → `ComandoEditar_Click` | Abre la edición del código seleccionado | `:10-40`; `.form.txt:398-411` |
| `ComandoEliminar` («Eliminar») | botón, deshabilitado al abrir | `OnClick` → `ComandoEliminar_Click` | Pide confirmación y borra | `:41-86`; `.form.txt:442-455` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra con `acSaveNo` | `:100-103`; `.form.txt:342-344` |
| `ComandoAyudaFormulario` | botón | `OnClick = [Event Procedure]` sin handler en el `.cls` | Ninguno visible | `.form.txt:497-499` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | «TIPOS DE ACTIVIDAD»; `lblTitulo1` a `Me.Caption` | `:93`; `.form.txt:591, 611` |
| `CuadroSuperior`, `Cuadro2`, `Marco120` | rectángulos y marco | — | Decoración | `.form.txt:253, 276, 385` |

## Flujos

1. **Apertura** (`Form_Open`, `:87-99`): `On Error Resume Next`; ajusta tamaño y título. No hay más inicialización: los botones Editar/Eliminar parten deshabilitados por diseño.
2. **Selección** (`:104-114`): lee `Column(0)`; vacío ⇒ deshabilita Editar y Eliminar, si no los habilita.
3. **Alta:** `DoCmd.OpenForm "FormActividadTipoAlta"` con `On Error Resume Next` (`:5-8`).
4. **Edición:** exige selección («Se ha de seleccionar el código del tipo de actividad que desea editar», `:16-20`), cierra el editor si estaba abierto (`FormularioAbierto`, `:26-28`) y lo abre con el código (`:29`).
5. **Eliminación:** exige selección (`:47-50`); confirma «¿Desea realmente eliminar el tipo de Actividad…?» con «No» por defecto (`:54-58`); llama a `Eliminar` (`:67`); éxito: `Requery`, foco en `ComandoAlta`, deshabilita Eliminar (`:73-75`).
6. **Error:** `MsgBox` crítico y el formulario sigue abierto (`:34-39, 80-85`).

## Reglas, validaciones y cálculos

`Eliminar` (`:115-178`):

- Código no vacío (`:135-138`).
- **Sin uso:** si `Dame("TbActividades", "Actividad", "Actividad", código)` devuelve algo, «El Tipo de actividad … forma parte de alguna actividad» (`:144-153`).
- El código existe en `TbCodActividad`; si no, «no está registrado en la tabla TbCodActividad» (`:154-162`).
- Borra el registro con DAO (`:163`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de tipos | `TbCodActividad` | `RowSource` SQL guardada en el control (`CODIGO`, `ACTIVIDADNOPROGRAMADA`) | — | `.form.txt:559-562` |
| Comprobación de uso | `TbActividades` | `Dame`: SQL dinámico concatenado | no | `:144`; `FG.bas:300-310` |
| Borrado | `TbCodActividad` | SQL dinámico concatenado + DAO `Delete` | no | `:154-163` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL concatenado con el código entre comillas (`:154-156`); `Dame` depende de un fallo de Jet para cambiar de numérico a texto (`FG.bas:303-310`); borrado físico sin restricción referencial declarada visible desde este formulario (la regla de «sin uso» está solo en VBA, `:144-153`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Tras borrar, `strCodTipoActividadSeleccionada` no se vacía (`:67-75`): el código borrado sigue como «seleccionado» hasta la siguiente pulsación en la lista (`:106`).
- El código pasa a SQL sin escapar (`:156`; `FG.bas:300-302`): un código con apóstrofo rompería la consulta.
- El handler se llama `ListaTiposDeActividad_click` (minúscula) y `Form_Open(Canel …)` tiene el parámetro mal escrito (`:87, 104`); en VBA no afecta, pero indica código copiado.
- `On Error Resume Next` en `Form_Open`, `ComandoAlta_Click`, `cmdSalir_Click` y la selección (`:6, 88, 101, 105`) oculta cualquier fallo, incluido un `FormActividadTipoAlta` inexistente.
- Mensaje de error de borrado habla de «eliminar la ubicación» (`:82`), texto heredado.
- Comprobación y borrado sin transacción: una actividad creada entre ambos pasos quedaría con un tipo inexistente.
- `ComandoAyudaFormulario` con `OnClick` sin handler (`.form.txt:499`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: clave foránea entre actividades y tipos para que la base impida el borrado en uso.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, la lista muestra todos los registros de `TbCodActividad` con código y tipo, y Editar/Eliminar están deshabilitados (`ListaTiposDeActividad`, `.form.txt:399, 443`).
2. Seleccionar una fila habilita Editar y Eliminar.
3. Editar abre el editor con el código de la fila seleccionada y, al guardar allí, la lista se refresca.
4. Eliminar pide confirmación con «No» por defecto; cancelar no borra.
5. Eliminar un tipo usado en alguna actividad muestra «forma parte de alguna actividad» y no borra.
6. Eliminar un tipo libre lo quita de la lista y deja Eliminar deshabilitado.
7. Alta abre el formulario de alta y su guardado refresca la lista.

## Pendientes

- Qué contiene realmente `TbActividades.Actividad` (código frente a texto), que determina la regla de «sin uso» (ERD de origen).
- `ComandoAyudaFormulario` sin handler (`.form.txt:499`).
- Acción de `FormActividadTipoAlta` (ficha de otro lote, no analizada aquí).
