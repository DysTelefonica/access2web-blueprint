# Form_FormNCProyectoGeneral

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCProyectoGeneral.cls` (1228 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «General» de la NC de proyecto: alta y edición de los datos únicos (expediente, descripción, causa raíz, tipología, entidad responsable, detectado por, responsables, control de eficacia, riesgos, concesión) y presentación de estado y fechas calculadas. Es un formulario **no enlazado** (sin `RecordSource`): trabaja sobre `m_ObjNCProyectoActiva` y sobre `Form_FormNCProyecto` (`m_NCAlInicio`, `m_PermitidoEditar`, `m_Expediente`, `m_ColRiesgos`). Las citas son `Form_FormNCProyectoGeneral.cls:línea` (abreviado `:línea`) salvo indicación; `Ops` abrevia `classes/NCProyectoOperaciones.cls` y `FG.bas`, `modules/Funciones Generales.bas`. Existe una variante con vínculo a otra NC en la ficha de `FormNCProyectoGeneralConVinculoNC`.

## Entrada y salida

Se carga como detalle (`FrmDetalle`) de `FormNCProyecto` (`Form_FormNCProyectoControlEficaciaAlta.cls:117`). Abre `FormExpedientesBusqueda` (`:435`, `:222-224`), `FormNCProyectoControlEficaciaAlta` (`:263`), `formRiesgosSeleccion` en diálogo (`:1150`) y `FormMotivosNoRequiereControlEficacia` (`:1174`). Escucha los eventos de la búsqueda de expedientes (`:4`, `:12`) y de motivos (`:792`). Actualiza `FormNCProyectoGestion` si está abierto (`:776-777`).

## Controles y eventos

`.form.txt` declara 58 bloques `Begin`: el formulario (`:5`), 8 de valores por defecto, la sección `Detalle` (`:575`) y 48 controles.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `PalabraClave` | texto | `OnKeyDown` (`:803-820`), `OnChange` sin procedimiento | Intro/Tab lanza la búsqueda; sin permiso de edición no actúa | `.form.txt:587-602` |
| `ComandoBuscarExpediente` | botón | `OnClick` | Busca por palabra clave: sin resultado abre la búsqueda; con 1 lo asigna; con varios abre la búsqueda filtrada | `:160-247`, `.form.txt:690` |
| `ComandoIrABusquedaExpedientes` | botón | `OnClick` | Abre `FormExpedientesBusqueda` con la palabra clave | `:423-458`, `.form.txt:768` |
| `ComandoActualizararExpedientes` | botón | sin `OnClick` | Habilitado/deshabilitado según permisos | `.form.txt:817`, `:734, 743` |
| `IDExpediente`, `Juridica`, `Titulo` | texto bloqueado (`IDExpediente` con `Tag="DATO"`) | — | Datos del expediente | `.form.txt:861-923`, `:1006-1008` |
| `Descripcion`, `CausaYAnalisRaiz` | texto, `DATO;OBLIGATORIO` | `BeforeUpdate` | Repinta bordes obligatorios | `:822-837`, `.form.txt:984-1063` |
| `FechaApertura` | texto, `DATO;OBLIGATORIO` | — | Se fija a la fecha del día en el alta | `:550`, `.form.txt:1144` |
| `FPREVCIERRE`, `FECHACIERRE` | texto bloqueado | `BeforeUpdate` (`FPREVCIERRE`) | Fechas calculadas | `:921-928`, `.form.txt:1161, 1823` |
| `EntidadResponsable`, `IDTipo`, `DetectadoPor` | combos de lista de valores, `DATO;OBLIGATORIO` | `BeforeUpdate` | Selección cerrada; `IDTipo` se rellena con `EstablecerComboTipo` | `.form.txt:1216, 1611, 1658`, `:497`, `:877-902` |
| `ResponsableTelefonica`, `RESPONSABLECALIDAD` | combos, `DATO;OBLIGATORIO` | `BeforeUpdate` | Se rellenan por código con los responsables | `:501-506`, `:904-919`, `.form.txt:1322, 2330` |
| `RequiereControlEficacia` | combo Sí/No, `DATO;OBLIGATORIO` | `AfterUpdate` | Activa «Datos alta control eficacia» o «Motivos No CE» | `:839-875`, `.form.txt:1733` |
| `VEHICULO`, `CodConcesionAsociada` | texto (`VEHICULO` con `DATO`) | — | Datos libres | `.form.txt:2067`, `:958-961` |
| `CodigoNoConformidad`, `ESTADO` | texto bloqueado | — | Código y estado calculados | `.form.txt:1854, 1917`, `:588-589, 665-666` |
| `CodigoRiesgo` | texto deshabilitado | — | Códigos de riesgos separados por `|` | `.form.txt:2153`, `:1201-1224` |
| `ComandoGrabar` | botón | `OnClick` | Valida y registra | `:277-421`, `.form.txt:1797` |
| `ComandoActualizar` | botón | `OnClick` | Recalcula datos calculados y recarga | `:122-158`, `.form.txt:1989` |
| `ComandoControlEficaciaDatos` | botón | `OnClick` | Abre el alta de control de eficacia | `:249-275`, `.form.txt:2096` |
| `btnSeleccionarRiesgos` | botón | `OnClick` | Elige riesgos en diálogo | `:1125-1162`, `.form.txt:2422` |
| `ComandoVerNoRequiere`, `ImagenMotivosNoCE` | botón / imagen oculta | `OnClick` | Meter o ver motivos de no control de eficacia | `:1163-1199`, `.form.txt:2486, 2512` |
| `Cuadro53` y 18 etiquetas (`Etiqueta43` … `Etiqueta254`) | rectángulo y rótulos | — | Decorativos | `.form.txt:619-1755` |

## Flujos

### Apertura (`Form_Load` → `EstablecerDatos`, `:461-487`, `:519-789`)

1. `EstablecerCombos` (`:489-516`) rellena tipología, responsables de Calidad y de Telefónica (`FG.bas:1773, 1806, 1921`).
2. Sin NC activa: «El formulario se ha abierto con parámetros insuficientes» (`:539-541`). Alta (sin `IDNoConformidad`): fija `FechaApertura` a hoy, pinta obligatorios y da foco a `PalabraClave` (`:547-553`). Edición: carga `m_NCAlInicio` con `constructor.getNCProyecto` (`:557-565`) y vuelca los campos (`:576-724`).
3. Si hay `NCProyectoAsociada`, toma de ella causa, tipología, entidad, control de eficacia, fechas y estado, y deshabilita los campos heredados (`:608-629`; `DeshabilitarControlesPorVinculoNC` `:43-60`); si no, usa los propios y calcula `FPREVCIERRE`, `FECHACIERRE` y estado (`:630-656`).
4. Permisos: con datos generales incompletos permite editar salvo técnicos; con completos, solo si `m_PermitidoEditar` no es «No» (`:732-771`).

### Grabar (`:277-421`)

1. Técnico ⇒ «No tiene autorización para esa acción»; sin cambios ⇒ «No hay cambios que guardar» (`:283-296`).
2. `RellenarDatosObjeto` (`:930-970`): `SetPropiedad` por cada control con `Tag` «DATO», `ResponsableTelefonica.Column(1)` y `CodConcesionAsociada`.
3. «No» a control de eficacia exige motivo (`:301-304`); si se pierde un control de eficacia ya descrito, pide confirmación con «No» por defecto (`:305-324`).
4. Alta/edición: sin NC asociada, `RegistrarDatosUnicos`; con asociada, `RegistrarCambiosDatosUnicosConVinculoNC` (`:326-344`). Refresca `m_NCAlInicio`, colores de pestañas, código, `ComandoActualizar_Click` y la lista de gestión (`:349-380`).
5. Mensajes de «Registro correcto», «Cierre de NC» o «Reapertura de NC» según cambie el estado (`:381-410`).

### Otros

- Expediente seleccionado (`m_FormExpedientes_Seleccionado`, `:12-40`) o único resultado ⇒ `EstablecerDatosExpediente` (`:971-1027`): rellena identificador, título, jurídicas y, si existe, el jefe de proyecto como responsable de Telefónica (`:1004-1012`).
- `RequiereControlEficacia` = «Sí» con motivos previos pide confirmar su pérdida (`:849-856`).
- Errores: patrón 1000 ⇒ advertencia; otro ⇒ `CorreoAlAdministrador` y aviso crítico (`:412-420`); `btnSeleccionarRiesgos` solo muestra aviso (`:1160-1161`).

## Reglas, validaciones y cálculos

- Obligatorios (`Tag` «OBLIGATORIO», `.form.txt:1000-2332`) y cálculo de «datos generales OK» en `NCProyecto.DatosGeneralesOK` (`NCProyecto.cls:993-1050`); con «Sí» a control de eficacia exige además control y fecha prevista (`NCProyecto.cls:1018-1031`).
- Cambios: `HaHabidoCambios` compara 15 campos y el motivo con `m_NCAlInicio` (`:1028-1122`).
- Rol: técnico no graba ni edita (`:283`, `:732-748`).
- Estado, fechas y color del estado salen de `EstadoGrabar`, `FPREVCIERREGrabar`, `FECHACIERREGrabar` (`:617, 637-641`).

## Lecturas y escrituras

| Operación | Tabla o fuente | Tipo | Evidencia |
|---|---|---|---|
| Carga de la NC | por `constructor.getNCProyecto` | ver ficha de la clase | `:558` |
| Alta de NC | `TbNoConformidades` | DAO `AddNew/Update` | `Ops:351-436` |
| Edición de NC | `TbNoConformidades` | SQL dinámico concatenado + DAO `Edit/Update` | `Ops:363-366, 453-543` |
| Alta/edición con vínculo | `TbNoConformidades` | DAO | `Ops:632-695, 750-813` |
| Registro de cambios | `RegistrarLogProyecto` | tras cada alta/edición | `Ops:446, 547, 701, 816` |
| Datos calculados | `ActualizarDatosNCProyecto` + `CacheNCProyecto.InvalidarCache` | ver clase | `FG.bas:2728-2755` |
| Listas de expedientes y responsables | consultas en `constructor`/`FG.bas` | no rastreadas aquí | `:160-247`, `FG.bas:1773-1921` |

No se ha visto transacción en el flujo de grabar.

## Coexistencia

Identificador y código de NC se calculan en cliente; `Edit/Update` DAO sobre `TbNoConformidades` y SQL concatenado deben pasar a sentencias parametrizadas y secuencia/transacción en PostgreSQL. Las fechas viajan como texto. Las listas de valores de tipología, entidad, detectado por y Sí/No están en el diseño del formulario (`.form.txt:1216, 1611, 1658, 1733`) y deberían salir de catálogo.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `If m_VEHICULO = "" Then Me.VEHICULO = m_VEHICULO` (`:716-717`): la condición parece invertida; el valor no vacío no se muestra.
- `ImagenMotivosNoCE` queda oculta con motivos al cambiar a «No» (`:865-867`), pero visible en la carga (`:703-705`).
- `ComandoActualizararExpedientes` no tiene `OnClick` (`.form.txt:817`) y `PalabraClave` declara `OnChange` sin procedimiento (`.form.txt:602`).
- `Err.Raise 100` tras `RellenarDatosObjeto` (`:299`, `:254`): sigue la rama de error crítico.
- Los valores de responsables figuran en el diseño del formulario (`.form.txt:1260, 1325, 2268`); contienen nombres reales.
- Condición `If m_ObjNCProyectoActiva.IDNoConformidad = ""` con `EsTecnico` sin `Else` para técnicos (`:547-553`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: separar alta y edición; calcular estado y fechas en servidor.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un técnico ve la NC en solo lectura y no puede grabar.
2. Sin cambios, grabar avisa y no escribe (`:290-296`).
3. «No» a control de eficacia sin motivo no se graba (`:301-304`).
4. Con vínculo, los campos heredados están deshabilitados y muestran los de la NC asociada (`:608-629`).
5. La búsqueda de expediente con un único resultado lo selecciona y completa jurídicas y responsable (`:176-190`).
6. Tras grabar se muestra el código de la NC y se actualizan los cálculos.
7. «Datos alta control eficacia» solo se habilita con «Sí», sin fecha prevista y sin vínculo (`:680-696`).

## Pendientes

- Líneas `:164-190` y `:1164-1190` leídas parcialmente; detalle de `FormMotivosNoRequiereControlEficacia` y `formRiesgosSeleccion` en sus propias fichas.
- Funciones `RellenarBordeCamposObligatorios`, `EstablecerColorPestañasNCProyecto` (`FG.bas:2287`) y propiedades calculadas de `NCProyecto`: no leídas en detalle.
- `BellChar` (`:1226`) sin uso rastreado.
