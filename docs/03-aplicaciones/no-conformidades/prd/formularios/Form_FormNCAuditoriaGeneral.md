# Form_FormNCAuditoriaGeneral

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCAuditoriaGeneral.cls` (648 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «General» de la ficha de una NC de auditoría: alta y edición de los datos únicos de la NC (tipo, descripción, causa raíz, punto de la norma, responsable, fecha de apertura, si requiere acción correctiva y control de eficacia). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): copia el estado desde `Form_FormNCAuditoria.m_NCAlInicio` y vuelca los controles al objeto global `m_ObjNCAuditoriaActiva` antes de grabar con `NCaUDITORIAOperaciones`. Es también el formulario de alta: sin `id` muestra una NC nueva. Las citas son `Form_FormNCAuditoriaGeneral.cls:línea` salvo indicación; `NCaUDITORIAOperaciones.cls` abrevia `classes/NCaUDITORIAOperaciones.cls`.

## Entrada y salida

Se muestra como detalle de `FormNCAuditoria` (primera pestaña, `NavGeneral`; `Form_FormNCAuditoria.cls:134-139`). En el alta el resto de pestañas están deshabilitadas (`Form_FormNCAuditoria.cls:151-157`). Abre dos formularios auxiliares:

| Destino | Cómo | Evidencia |
|---|---|---|
| [Form_FormNCAuditoriaControlEficaciaAlta](Form_FormNCAuditoriaControlEficaciaAlta.md) | Botón `ComandoControlEficaciaDatos`: vuelca datos, exige datos generales completos y lo abre | `:50-64` |
| `FormMotivosNoRequiereControlEficacia` | Botón `ComandoVerNoRequiere`: lo abre con el motivo actual como argumento y se suscribe a su evento `MotivoRegistrado` | `:553-561, 410-419` |

También lo invoca desde fuera `Form_FormNCAuditoriaControlEficaciaAlta` (`ComandoGrabar_Click` público, `:78`).

## Controles y eventos

`.form.txt` declara 43 bloques `Begin`: el formulario, 8 bloques de valores por defecto (`.form.txt:447-554` y anteriores), la sección `Detalle` (`:578`) y 33 controles: 14 etiquetas y los 19 de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Tipo` | lista de valores `NC;OB;OP`, obligatorio | `BeforeUpdate` | Recalcula bordes de obligatorios | `:421-425`, `.form.txt:997-1003` |
| `RequiereAccionCorrectiva` | lista `Sí;No`, obligatorio | `BeforeUpdate` | Habilita o borra el motivo; con «No» anula `RequiereControlEficacia` | `:427-448`, `.form.txt:1559-1565` |
| `MotivoNoAccionCorrectiva` | cuadro de texto, deshabilitado por defecto | `BeforeUpdate` | Obligatorio solo si «No» (cambia su `Tag`) | `:450-454, 328-332`, `.form.txt:1597-1600` |
| `Descripcion`, `CAUSARAIZ` | cuadros de texto, obligatorios | `BeforeUpdate` | Bordes de obligatorios | `:456-465`, `.form.txt:600-603, 663-666` |
| `PuntoNorma` | lista de valores fija, obligatorio | `BeforeUpdate` | Bordes; alimentada además por `EstablecerComboPuntosNormaNCAuditorias` | `:466-470, 243`, `.form.txt:796-806` |
| `RESPONSABLEIMPLANTACION` | lista, obligatorio | `BeforeUpdate` | Se rellena con `EstablecerComboResponsablesImplantacion` | `:471-475, 239`, `.form.txt:1040-1047` |
| `FechaApertura` | cuadro de texto, obligatorio | — | Fecha de apertura; por defecto `Date` en el alta | `:284`, `.form.txt:725-728` |
| `RequiereControlEficacia` | lista `Sí;No`, obligatorio | `AfterUpdate` | Habilita el botón de datos o el de motivos | `:476-511`, `.form.txt:1116-1122` |
| `FPREVCIERRE`, `FECHACIERRE`, `ESTADO`, `Numero`, `NombreAuditoria` | 5 cuadros de texto bloqueados | — | Informativos; los calcula la NC | `:280-387`, `.form.txt:758, 1222, 1254, 1316, 1461` |
| `ComandoGrabar` | botón | `OnClick` | Valida y registra | `:78-198`, `.form.txt:1181-1183` |
| `ComandoActualizar` | botón | `OnClick` | Recalcula datos calculados y refresca | `:7-47`, `.form.txt:1373-1375` |
| `ComandoControlEficaciaDatos` | botón, deshabilitado por defecto | `OnClick` | Abre el alta del plan de control de eficacia | `:50-76`, `.form.txt:1420-1422` |
| `ComandoVerNoRequiere` | botón, deshabilitado por defecto | `OnClick` | Abre el registro del motivo de «no requiere» | `:542-578`, `.form.txt:1658-1660` |
| `ImagenMotivosNoCE` | imagen, oculta | — | Indica que existe motivo registrado | `:362-365`, `.form.txt:1693` |

## Flujos

### Apertura (`Form_Load`, `:203-229`; `EstablecerDatos`, `:259-408`)

Rellena los combos (`:239-246`). Sin objeto activo: «El formulario se ha abierto con parámetros insuficientes» (`:276-279`). **Alta** (`id` vacío): deshabilita «Actualizar», fecha de apertura = hoy, nombre de la auditoría y número calculado con `NumeroNCCalculado` (`:282-292`). **Edición**: recarga `m_NCAlInicio` si falta, regraba estado y fechas calculadas (`.EstadoGrabar`, `.FPREVCIERREGrabar`, `.FECHACIERREGrabar`), copia los campos no vacíos y ajusta botones de control de eficacia (`:295-387`); refresca la fila en `FormNCAuditoriaGestion` si está abierto (`:396-401`).

### Grabar (`:78-198`)

1. Técnico ⇒ «No tiene autorización para esa acción» (`:93-96`).
2. Sin cambios frente a `m_NCAlInicio` ⇒ «No hay cambios que guardar» (`:97-104`, `:580-648`).
3. Vuelca los controles `DATO` al objeto y asigna la auditoría (`:105-108`, `:515-540`).
4. Si `RequiereControlEficacia = "No"` sin motivo ⇒ pide usar «Meter Motivos No CE» (`:111-114`).
5. Si pasa a «No» habiendo control descrito, confirma «¿Desea guardar y perder ese dato?» con «No» por defecto (`:116-134`).
6. `RegistrarDatosUnicos` (`:136-146`): alta o edición según exista `m_NCAlInicio`; recarga la NC, repinta pestañas y llama a `ComandoActualizar_Click` (`:148-154`), refresca `FormNCAuditoriaGestion` (`:155-157`).
7. Avisos: «Registro correcto de … de Auditoria NN» y cierre o reapertura según el estado (`:165-182`).

### Errores

Error 1000 ⇒ aviso con `m_Error`; otro ⇒ mensaje con número y descripción más `CorreoAlAdministrador` (`:189-197`).

## Reglas, validaciones y cálculos

- **Rol:** los técnicos no graban (`:93-96`).
- **Obligatorios** (`NCaUDITORIAOperaciones.cls:28-66`, mensajes en `:30-64`): auditoría conocida, tipo, descripción, causa raíz, fecha de apertura válida, punto de la norma, responsable, `RequiereAccionCorrectiva` y `RequiereControlEficacia` en «Sí/No». «No» a acciones correctivas exige motivo (`NCaUDITORIAOperaciones.cls:68-73`); «Sí» a control de eficacia con «No» a acciones correctivas es rechazado (`NCaUDITORIAOperaciones.cls:74-78`); con control de eficacia «Sí» exige fecha prevista y descripción (`NCaUDITORIAOperaciones.cls:79-89`); la fecha prevista de cierre no puede ser anterior a la apertura (`NCaUDITORIAOperaciones.cls:90-95`).
- **Sin acción correctiva:** la NC se cierra en el acto, con fecha prevista y de cierre iguales a la de apertura (`NCaUDITORIAOperaciones.cls:176-183, 212-219`).
- **Sin control de eficacia:** si deja de ser «Sí», se vacían los cinco campos del control (`NCaUDITORIAOperaciones.cls:224-234`); si vuelve a «Sí» se limpia el motivo previo con aviso (`NCaUDITORIAOperaciones.cls:238-243`).
- **Descripción:** se eliminan apóstrofos y `;` se sustituye por `:` (`NCaUDITORIAOperaciones.cls:145-146`).
- **Numeración:** el `id` y el número se calculan en el alta (`NCaUDITORIAOperaciones.cls:134-139`).
- **Indicador visual:** `RellenarBordeCamposObligatorios` marca los obligatorios vacíos (`:291, 394`); el objeto se limpia de caché de puntos de la norma tras grabar (`:25`; `NCaUDITORIAOperaciones.cls:254-260`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Alta de NC | `TbNoConformidadesAuditoria` | DAO `AddNew/Update` | `NCaUDITORIAOperaciones.cls:133, 155-187` |
| Edición de NC | ídem, por `ID` | SQL dinámico concatenado + DAO `Edit/Update` | `NCaUDITORIAOperaciones.cls:142-143, 196-247` |
| Plan de control de eficacia (si «Sí») | ídem | ver `RegistrarControlEficacia` | `NCaUDITORIAOperaciones.cls:262-267, 484-559` |
| Traza | `RegistrarLogAuditoria` («Alta de NC» / «Edición de NC de Auditoría») | pendiente | `NCaUDITORIAOperaciones.cls:188, 248` |
| Punto de la norma nuevo | `PuntoNormaAuditoriaRegistrar` | pendiente | `NCaUDITORIAOperaciones.cls:254-260` |
| Recarga y recálculo | `constructor.getNCAuditoria`, `ActualizarDatosNCAuditoria` → `ActualizarDatosCalculados` | pendiente | `:21, 148`, `Funciones Generales.bas:2762-2781` |

Sin transacciones: alta, plan, log e indicadores son pasos sucesivos (`NCaUDITORIAOperaciones.cls:187-273`).

## Coexistencia

Pendiente de `../coexistencia.md`. A adaptar: `ID` y `Numero` calculados en código y no por secuencia (`NCaUDITORIAOperaciones.cls:134-139`), `WHERE ID=` concatenado (`NCaUDITORIAOperaciones.cls:143`), `Edit/Update` de DAO sobre `SELECT *`, escritura de `Null` por DAO (`NCaUDITORIAOperaciones.cls:230-234`) y listas de valores fijas en el formulario (`Tipo`, `PuntoNorma`, `RESPONSABLEIMPLANTACION`; esta última sustituida en ejecución por el combo, `:239`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `HaHabidoCambios` repite la comparación de `RequiereControlEficacia` (`:626-634`) y no compara `FPREVCIERRE`, `FECHACIERRE` ni `Numero` (`:593-641`).
- `ComandoGrabar_Click` es público y lo invoca `Form_FormNCAuditoriaControlEficaciaAlta` (`Form_FormNCAuditoriaControlEficaciaAlta.cls:84-92`); los avisos se muestran igualmente, aunque el formulario llamante lo cierre después.
- `Err.Raise 100` tras `RellenarDatosObjeto` (`:53-56, 105-108`) no se trata como aviso controlado: iría al administrador por correo.
- `RequiereAccionCorrectiva_BeforeUpdate` modifica controles antes de confirmar el cambio y con `On Error Resume Next` (`:427-448`).
- `RequiereControlEficacia_AfterUpdate` solo actúa si `m_PermitidoEditar = Sí` (`:479`); en caso contrario el valor puede cambiar sin ajustar botones.
- `NCaUDITORIAOperaciones.cls:74-78` rechaza «control de eficacia Sí con acción correctiva No», pero al poner «No» el formulario vacía `RequiereControlEficacia` (`:442-445`), forzando al usuario a rellenarlo de nuevo.
- Si `RequiereControlEficacia = "Sí"` la ficha de plan se pide aparte y `MotivoDatosUnicosNoOK` exige ya la fecha prevista (`NCaUDITORIAOperaciones.cls:79-89`): el alta no puede grabarse sin abrir antes `FormNCAuditoriaControlEficaciaAlta`.
- Un apóstrofo se borra de la descripción en silencio (`NCaUDITORIAOperaciones.cls:145`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: integrar plan y motivo de control de eficacia en la misma página; sustituir las listas fijas por catálogos; mostrar los obligatorios por validación de formulario y no por color.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un técnico no puede grabar (`:93-96`).
2. No se graba sin los obligatorios indicados en `NCaUDITORIAOperaciones.cls:28-66`.
3. «Requiere acción correctiva: No» exige motivo y cierra la NC en el momento (`NCaUDITORIAOperaciones.cls:68-73, 176-183`).
4. «Requiere control de eficacia: No» exige motivo registrado (`:111-114`).
5. Pasar de «Sí» a «No» con control descrito pide confirmación y, al aceptar, borra los datos del control (`:116-134`; `NCaUDITORIAOperaciones.cls:224-234`).
6. El alta calcula número e identificador y no permite abrir las demás pestañas hasta grabar (`NCaUDITORIAOperaciones.cls:134-139`; `Form_FormNCAuditoria.cls:151-157`).
7. Sin cambios, «Grabar» informa de que no hay nada que guardar (`:97-104`).

## Pendientes

- `SetPropiedad` de `NCAuditoria`, invocado por nombre de control (`:525`).
- `EstablecerComboResponsablesImplantacion`, `EstablecerComboPuntosNormaNCAuditorias`, `RellenarBordeCamposObligatorios`, `EstablecerColorPestañasNCAuditoria`, `ActualizarDatosCalculados`, `PuntoNormaAuditoriaRegistrar`, `RegistrarLogAuditoria`: cuerpos no revisados.
- Ficha de `FormMotivosNoRequiereControlEficacia` y su evento `MotivoRegistrado`.
