# Form_FormNCAuditoriaControlEficacia

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCAuditoriaControlEficacia.cls` (415 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Control de eficacia» de la ficha de una NC de auditoría: planifica el control (fecha prevista y descripción) y registra su resultado (fecha, resultado y conformidad). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): lee y escribe sobre el objeto global `m_ObjNCAuditoriaActiva` (`Variables Globales.bas:127`) y delega la persistencia en `NCaUDITORIAOperaciones`. Las citas son `Form_FormNCAuditoriaControlEficacia.cls:línea` salvo indicación; `NCaUDITORIAOperaciones.cls` abrevia `classes/NCaUDITORIAOperaciones.cls`, `NCAuditoria.cls`, `classes/NCAuditoria.cls` y `VG.bas`, `modules/Variables Globales.bas`.

## Entrada y salida

Se carga como subformulario de navegación `FrmDetalle` de `FormNCAuditoria` mediante el botón `NavControlEficacia` (`Form_FormNCAuditoria.form.txt:1089`, `:951`). En el alta de una NC (sin `id`) el botón queda deshabilitado (`Form_FormNCAuditoria.cls:155, 172`). Sin objeto activo o sin `id`, `EstablecerDatos` avisa «El formulario se ha abierto con parámetros insuficentes» (`:306-313`).

Desde «Grabar» puede conmutar la pestaña a `FormNCAuditoriaAcciones` (`:132-139`) y refrescar la línea de `FormNCAuditoriaGestion` si el estado cambia y está abierto (`:157-161`).

## Controles y eventos

`.form.txt` declara 23 bloques `Begin`: el formulario (`.form.txt:5`), 8 bloques de valores por defecto (`:141-290`), la sección `Detalle` (`:302`) y los 13 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `RequiereControlEficacia` | cuadro de texto bloqueado, `Tag=DATO` | `AfterUpdate` (nunca disparable: bloqueado) | Se rellena siempre con «Sí» al cargar | `:319, 350-354`, `.form.txt:706-723` |
| `FechaPrevistaControlEficacia` | cuadro de texto, fecha corta, `Tag=DATO` | `AfterUpdate` | Fecha prevista del control; recalcula bordes | `:199-203`, `.form.txt:322-326` |
| `ControlEficacia` | cuadro de texto, `Tag=DATO` | `AfterUpdate` | Descripción del control | `:187-191`, `.form.txt:417-420` |
| `FechaControlEficacia` | cuadro de texto, fecha corta, `Tag=DATO` | `AfterUpdate` | Fecha del resultado | `:193-197`, `.form.txt:480-484` |
| `ResultadoControlEficacia` | cuadro de texto, `Tag=DATO` | `AfterUpdate` | Resultado | `:412-416`, `.form.txt:581-584` |
| `ConformeControlEficacia` | lista desplegable, lista de valores «Sí;No», `Tag=DATO` | `AfterUpdate` | Conformidad | `:181-185`, `.form.txt:651-657` |
| `ComandoGrabar` | botón «Grabar» | `OnClick` | Valida, registra y recalcula estado | `:35-179`, `.form.txt:541-543` |
| `Etiqueta254`, `255`, `256`, `264`, `63`, `65` | 6 etiquetas | — | Rótulos de los campos | `.form.txt:353, 384, 447, 511, 611, 686` |

## Flujos

### Apertura (`Form_Load`, `:274-300`)

`EstablecerDatos` copia a los controles los valores de `m_ObjNCAuditoriaActiva` solo si no están vacíos (`:318-335`), habilita «Grabar» (`:315`), guarda el estado calculado (`:314`) y pinta bordes (`:337`).

### Grabar (`:35-179`)

1. Un técnico recibe «No tiene autorización para esa acción» (`:48-51`).
2. Detecta cambios frente a `Form_FormNCAuditoria.m_NCAlInicio`: plan (fecha prevista, descripción; `:355-382`) y resultado (resultado, conforme; `:383-409`). Nota: la fecha del resultado no cuenta como cambio (ver defectos). Sin cambios ⇒ «No hay cambios que guardar» (`:60-63`).
3. Vuelca los controles `DATO` al objeto (`:65`, `:9-34`).
4. Con fecha prevista: si cambió el plan, `RegistrarControlEficacia` (`:74-81`). Sin fecha prevista, y si al inicio la tenía, `EliminarResultadoControlEficacia` y salta a `fin` (`:82-91`).
5. Con fecha de resultado: si cambió el resultado, `RegistrarResultadoControlEficacia` (`:94-102`). Sin ella, y si al inicio la tenía, `EliminarResultadoControlEficacia` (`:103-111`).
6. En `fin`: recarga `m_NCAlInicio`, le copia los cinco valores del formulario, regraba el estado con `EstadoGrabar` y vuelve a pintar (`:116-130`).
7. Si `ConformeControlEficacia = "No"` y requiere control, pregunta si desea planificar una nueva acción y, con «Sí», cambia a la pestaña de acciones (`:132-142`).
8. Avisa del cierre («La NC de Auditoria pasa a estar cerrada») o reapertura de la NC según el cambio de estado (`:143-156`).

### Errores

Error 1000 ⇒ aviso con `m_Error`; otro error ⇒ mensaje con número y descripción, `CorreoAlAdministrador` y `MsgBox` crítico (`:169-178`, `:291-299`). En los errores de operaciones recarga la NC desde base de datos (`:78, 86, 98, 107`).

## Reglas, validaciones y cálculos

- **Rol:** los técnicos no pueden grabar (`:48-51`; `EsTecnico` en `VG.bas:98`). El resto de la comprobación de roles no está en el formulario.
- **Plan** (`NCaUDITORIAOperaciones.cls:498-509`): exige NC registrada, fecha prevista y descripción. Si es «Conforme», exige todas las tareas de acción finalizadas (`:510-515`, `:517-525`).
- **Resultado** (`NCaUDITORIAOperaciones.cls:575-625`): exige que la NC requiera acciones correctivas (`:579-582`), que `RequiereControlEficacia` sea «Sí» (`:583-590`), plan completo, fecha de resultado, resultado y conforme «Sí»/«No»; si «Sí» con acción correctiva, exige acciones con tarea y todas las tareas finalizadas (`:612-625`).
- **Eliminar resultado:** vacía fecha, resultado y conforme y recalcula el estado (`NCaUDITORIAOperaciones.cls:676-695`).
- **Indicador visual:** el borde se pinta como relleno o no relleno según los campos (`:241-269`) y el color de la pestaña `NavControlEficacia` según `EficaciaOK` (`:235-239`; `NCAuditoria.cls:815-870`). Con estado inicial «Cerrada» no se recalcula (`:221-225`). Constantes de borde en `VG.bas:83-86`.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Alta/modificación del plan | `TbNoConformidadesAuditoria` (`RequiereControlEficacia`, `FechaPrevistaControlEficacia`, `ControlEficacia`) | SQL dinámico concatenado por `ID` + DAO `Edit/Update` | `NCaUDITORIAOperaciones.cls:526-540` |
| Registro del resultado | misma tabla (`FechaControlEficacia`, `ResultadoControlEficacia`, `ConformeControlEficacia`, `Estado`) | ídem | `NCaUDITORIAOperaciones.cls:626-642` |
| Borrado del resultado | misma tabla (campos a `Null`, `Estado`) | ídem | `NCaUDITORIAOperaciones.cls:680-697` |
| Recarga de la NC | `constructor.getNCAuditoria` | DAO (contenido no revisado) | `:78, 117` |
| Persistencia del estado | `NCA.EstadoGrabar` | contenido pendiente | `:128`, `NCAuditoria.cls:1058` |
| Traza de cambios | `RegistrarLogAuditoria` | pendiente | `NCaUDITORIAOperaciones.cls:543, 645, 700` |
| Indicadores | `PintarIndicadores` | pendiente | `NCaUDITORIAOperaciones.cls:549, 651, 707` |

Sin transacciones: cada operación hace su `Edit/Update` y registra el log después, sin envolverlos.

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). A adaptar: el `WHERE ID=` concatenado (`NCaUDITORIAOperaciones.cls:528`), el uso de `Edit/Update` de DAO sobre recordsets abiertos con `SELECT *`, y los valores `Null` escritos por DAO (`NCaUDITORIAOperaciones.cls:691-693`) frente a PostgreSQL vinculado. Las fechas viajan como texto vía `Nz(..., "")` (`:122-126`); la comparación por cadena no equivale a comparar fechas.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `HaHabidoCambiosResultadoControlEficacia` no compara `FechaControlEficacia` (`:394-401`): un cambio solo de esa fecha puede acabar en «No hay cambios que guardar».
- Tras `RellenarDatosObjeto` lanza `Err.Raise 100` y no 1000 (`:65-68`): el error se mostraría como técnico y enviaría correo al administrador.
- Con la fecha de resultado vacía y sin fecha al inicio no se llama a nada, pero resultado/conforme se quedan en el objeto (`:103-112`).
- El flujo sin fecha prevista llama a `EliminarResultadoControlEficacia` (no al plan) y salta a `fin` omitiendo el resto (`:82-91`); el plan anterior no se borra en `TbNoConformidadesAuditoria`.
- `RegistrarResultadoControlEficacia` exige «Sí» en `RequiereControlEficacia` (`NCaUDITORIAOperaciones.cls:583-590`), pero el formulario lo fija siempre a «Sí» (`:319`) y la rama «No» es inalcanzable desde aquí.
- `EstableceColorBordes` usa `On Error Resume Next` (`:206`) y puede ocultar errores.
- El nombre `NCaUDITORIAOperaciones` y el texto «Auditoria» (`:148`) carecen de coherencia tipográfica.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: marcar los campos obligatorios por regla en lugar de por borde; separar «Plan» y «Resultado» en dos acciones; sustituir el aviso de «nueva acción» por un enlace directo.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un técnico no puede grabar (`:48-51`).
2. Sin cambios, «Grabar» informa de que no hay nada que guardar (`:60-63`).
3. No se registra un plan sin fecha prevista ni descripción (`NCaUDITORIAOperaciones.cls:502-509`).
4. No se registra un resultado sin fecha, resultado y conforme (`NCaUDITORIAOperaciones.cls:600-611`).
5. «Conforme: Sí» solo se acepta con todas las tareas cerradas (`NCaUDITORIAOperaciones.cls:510-515`, `:621-624`).
6. Con «Conforme: No», la aplicación ofrece planificar una nueva acción (`:132-139`).
7. El cambio de estado a/desde «Cerrada» se comunica al usuario (`:145-156`).

## Pendientes

- Cuerpo de `NCA.EstadoGrabar`, `RegistrarLogAuditoria`, `PintarIndicadores` y `constructor.getNCAuditoria`.
- `SetPropiedad` de `NCAuditoria` (invocado dinámicamente desde `:20` por el nombre de control).
- `FormNCAuditoriaAcciones` se abre por cadena (`:138`); su ficha no se enlaza aquí.
