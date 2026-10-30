# Form_FormMotivosHPS

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormMotivosHPS.cls` (323 líneas) y `.form.txt` (861 líneas) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_FormMotivosHPS.cls:línea` del commit fijado.

Gestión del catálogo de motivos (`TbMotivoHPS`): lista, alta, edición y eliminación. Diez procedimientos. El alta y la edición se canalizan abriendo el diálogo `Form_FormMotivoHPS` y suscribiéndose a sus eventos `Alta` y `Editado`; la eliminación usa directamente `MotivoHPS.EliminarMotivo`, que avisa si hay usuarios o solicitudes con ese motivo.

> **Hallazgo previo.** Una búsqueda global en `src/` no encuentra ningún `DoCmd.OpenForm "FormMotivosHPS"` ni un `SourceObject = "FormMotivosHPS"` en los `.form.txt`. En `cfa1937a` el catálogo queda compilado pero **sin camino de apertura** desde el resto de la aplicación; hace falta confirmación con el área responsable para descartar que sea residuo de un refactor previo.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| (nadie) | `DoCmd.OpenForm "FormMotivosHPS"` no aparece en `src/`; ver hallazgo previo. | — |
| Hacia `Form_FormMotivoHPS` (alta) | `ComandoAlta_Click` cierra instancia previa, fija `m_ObjMotivoActivo = Nothing` y abre el diálogo; queda suscrito por `WithEvents` | `:55-72` |
| Hacia `Form_FormMotivoHPS` (edición) | `ComandoEditar_Click` carga `m_ObjMotivoActivo = m_MotivoSeleccionado` y abre el diálogo | `:93-104` |

Eventos escuchados del diálogo: `Alta` (`:5`, manejado en `m_Form_Alta`, `:273-296`) y `Editado` (`:5`, manejado en `m_Form_Editado`, `:299-323`). Estado compartido: la global `m_ObjMotivoActivo` (`VG.bas:199`), `m_ObjEntorno` (`VG.bas:190`, cuya `ColMotivos` se anula tras altas y ediciones, `:284-286, 308-310`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo1` | etiqueta «MOTIVOS HPS» | — | `Form_Open` copia el texto a `Me.Caption` (`:179`) | `:253-255`, `:179` |
| `ListaFiltrados` | ListBox (`RowSourceType = "Value List"`) | `Click` y `DblClick` | `Click` fija `m_MotivoSeleccionado` y habilita los botones (`:204-260`); doble clic dispara `ComandoEditar_Click` | `:289-291`, `:204-270` |
| `ComandoAlta` | botón «Alta» | `OnClick` | Abre `FormMotivoHPS` con `m_ObjMotivoActivo = Nothing` (`:49-78`) | `:461-463`, `:49-78` |
| `ComandoEditar` | botón «Editar» | `OnClick` | Si hay selección, abre `FormMotivoHPS` con la instancia (`:81-115`) | `:422-424`, `:81-115` |
| `ComandoEliminar` | botón «Eliminar» | `OnClick` | Pide confirmación, exige oficina, llama a `m_MotivoSeleccionado.EliminarMotivo` y recarga la lista (`:118-166`) | `:501-503`, `:118-166` |
| `cmdSalir` | botón «Salir» | `OnClick` | Cierra con `acSaveNo` sin capturador de error (`:8-12`) | `:332-334`, `:8-12` |

## Flujos

### Apertura (`Form_Open`, `:169-202`)

1. `Ajustar Me` (`:177`; `FU.bas:735-759`).
2. `Me.Caption = Me.lblTitulo1.Caption` (`:179`).
3. `EstablecerLista m_Error` (`:183`) carga `m_ObjEntorno.ColMotivos` al ListBox y, si la propiedad trae error, lo propaga (`p_Error = m_ObjEntorno.Error`, `:28`).

### Carga del catálogo (`EstablecerLista`, `:14-45`)

1. Activa `ComandoEditar` y `ComandoEliminar` (`:22-23`); se desactivarán desde `ListaFiltrados_Click` si no hay selección válida.
2. Vacía el `RowSource` del ListBox (`:26`).
3. Lee `m_ObjEntorno.ColMotivos` (`:27`); si el `Entorno` trae error, propaga a `p_Error` y aborta (`:27-31`).
4. Añade una fila por cada `MotivoHPS` con `lst.AddItem m_Motivo.MotivoHPS` (`:32-37`). No muestra identificador.

### Selección (`ListaFiltrados_Click`, `:204-261`)

1. Desactiva los dos botones al entrar (`:213-214`).
2. Si no hay fila seleccionada, anula `m_MotivoSeleccionado` y sale (`:215-221`).
3. Si cambia la fila, recarga `m_MotivoSeleccionado` con `constructor.getMotivo(p_Motivo:=ListaFiltrados.Column(0), …)` (`:222-234`; `constructor.bas:2147-…`).
4. Si encuentra motivo, vuelve a habilitar los botones (`:245-246`).

### Alta y edición (delegadas al diálogo)

Se describen en [Form_FormMotivoHPS](Form_FormMotivoHPS.md). Los manejadores `m_Form_Alta` y `m_Form_Editado` (`:273-296, 299-323`) anulan `m_ObjEntorno.ColMotivos` y repintan la lista.

### Eliminación (`ComandoEliminar_Click`, `:118-167`)

1. Exige selección y oficina (`:123-130`).
2. `MsgBox` de confirmación (`vbYesNo + vbDefaultButton2`, `:131-133`).
3. `m_MotivoSeleccionado.EliminarMotivo m_Error` (`:138`; `MotivoHPS.cls:97-154`); si hay usuarios o solicitudes con el motivo, devuelve error y aborta (`:97-140`).
4. Tras éxito, invalida el caché (`Set m_ObjEntorno.ColMotivos = Nothing`, `:149`) y recarga la lista con `EstablecerLista m_Error` (`:150`).

### Errores

Los capturadores diferencian el error interno 1000 (aviso previsto, sin correo) del resto, donde envían `CorreoAlAdministrador m_Error` (`FU.bas:585`) y muestran `MsgBox` crítico (`:71-78, 108-115, 159-166, 193-200, 253-260, 289-296, 315-322`). `cmdSalir_Click` (`:8-12`) y `ListaFiltrados_DblClick` (`:263-270`) no capturan errores: el primero cierra con `On Error Resume Next` implícito y el segundo envuelve `ComandoEditar_Click` en `On Error Resume Next` + `Echo`.

## Reglas, validaciones y cálculos

- **Reapertura segura:** antes de abrir `FormMotivoHPS`, los dos botones que lo invocan cierran cualquier instancia previa con `acSaveNo` (`:56-59, 93-96`).
- **Oficina obligatoria para borrar:** el chequeo `m_EnOficina = EnumSiNo.no` aborta (`:127-130`); el alta y la edición no lo aplican aquí (lo aplica `Form_FormMotivoHPS`).
- **Borrado con dependencias:** `MotivoHPS.EliminarMotivo` recorre `constructor.getUsuariosConMotivo` y `constructor.getSolicitudesConMotivo` y aborta si cualquiera trae elementos (`:103-141`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Lectura del catálogo | `TbMotivoHPS` (cacheada en `m_ObjEntorno.ColMotivos`) | DAO Recordset vía `constructor.getMotivos` | `constructor.bas:2099-2145` |
| Búsqueda por motivo | `TbMotivoHPS` (`WHERE MotivoHPS='…'`) | DAO Recordset | `constructor.bas:2147-…` |
| Borrado | `TbMotivoHPS` (`DELETE FROM … WHERE MotivoHPS='…'`) | `getdb().Execute` | `MotivoHPS.cls:144-151` |

El formulario solo lanza el borrado y observa el resultado; las altas y ediciones las ejecuta `Form_FormMotivoHPS`.

## Coexistencia

Pendiente de `../coexistencia.md`. Lo que ya se ve: la lista es un `ListBox` plano con `AddItem` (`:33-36`); un esquema web lo sustituiría por una tabla paginada con `id` y `MotivoHPS`. La selección actual trabaja contra el `MotivoHPS` textual como clave, lo que obligará a introducir un `ID` surrogate en la migración.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-29.** Catálogo **no cableado**: `DoCmd.OpenForm "FormMotivosHPS"` no aparece en `src/`; en la versión analizada `cfa1937a` el formulario queda compilado sin camino de apertura, igual que `Form_FormAltaDatosSolicitante` (H-HPS-S-23). Confirmar con el área responsable si es residuo o si existe un lanzador externo.
- **H-HPS-S-30.** Botones habilitados por defecto: `EstablecerLista` activa `ComandoEditar` y `ComandoEliminar` al pintar la lista (`:22-23`) y solo los desactiva `ListaFiltrados_Click` cuando la fila activa es la cabecera o no hay selección. Antes del primer clic, los dos botones están habilitados; un doble clic accidental dispara `ComandoEditar_Click`, que aborta con «Seleccione una elemento de la lista» (typo «un elemento» también, `:91`). Hay que verificar si esto produce confusión en el operador.
- `ListaFiltrados_Click` deshabilita los dos botones antes de reevaluar (`:213-214`); si el clic recae en una fila vacía, `m_MotivoSeleccionado` queda `Nothing` (`:216`) y los botones permanecen deshabilitados, mientras que `EstablecerLista` los volverá a habilitar la próxima vez que se pinte.
- `ListaFiltrados_DblClick` (`:263-270`) depende de que `Echo False/True` se ejecute dentro de `On Error Resume Next`: si la pantalla no responde, el doble clic puede dejar la UI sin repintar.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: pantalla de catálogo como tabla con filtros por servidor; botones deshabilitados hasta selección válida; sustituto de `ListBox` por HTML estándar con paginación.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir el catálogo lista todos los motivos existentes en `TbMotivoHPS`.
2. Seleccionar una fila habilita «Editar» y «Eliminar»; seleccionar una fila vacía los deshabilita.
3. Pulsar «Alta» abre el diálogo, da de alta y refresca la lista al volver.
4. Pulsar «Editar» abre el diálogo con el motivo seleccionado y refresca la lista al volver.
5. Pulsar «Eliminar» pide confirmación y, si hay usuarios o solicitudes con ese motivo, aborta con el mensaje correspondiente.
6. Doble clic sobre una fila equivale a «Editar».

## Pendientes

- Confirmar el camino de apertura del catálogo.
- Política sobre la activación inicial de los botones (UX vs. coherencia con el resto del proyecto).
- Homogeneización del aviso «Seleccione un elemento» (typo en `:91` y `:124`).
