# Form_FormNCAuditoriaAC

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCAuditoriaAC.cls` (328 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ventana de alta y edición de una acción correctiva (AC) de una NC de auditoría: pide la descripción de la acción, notas y responsable, y muestra de solo lectura la NC, el estado y las fechas calculadas. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); los datos se vuelcan a los controles desde objetos (`ACAuditoria`) y se guardan con `ACAuditoriaOperaciones.Registrar`. Las citas son `Form_FormNCAuditoriaAC.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`, `ACAuditoriaOperaciones.cls` abrevia `classes/ACAuditoriaOperaciones.cls` y `VG.bas`, `modules/Variables Globales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormNCAuditoriaAcciones` | `DoCmd.OpenForm "FormNCAuditoriaAC", OpenArgs:=m_ObjNCAuditoriaActiva.Estado` (alta) y apertura sin `OpenArgs` | `Form_FormNCAuditoriaAcciones.cls:37, 200` |

El contexto viaja en las globales `m_ObjNCAuditoriaActiva` y `m_ObjACAuditoriaActiva`: sin la NC activa muestra «Se ha abierto el formulario con parámetros insuficientes» (`:41-44`). `OpenArgs` se guarda en `m_EstadoAlInicio` (`:305`). Se cierra con `cmdSalir` (`:155-163`), tras registrar (`:280`) o ante un error de apertura (`:325`). Declara los eventos `Alta` y `Editado` (`:8-9`), lanzados al registrar (`:265-270`); los receptores no se han localizado (ver Pendientes).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | form | `OnOpen` | `Ajustar Me`, lee `OpenArgs` y llama a `EstablecerDatos` | `.form.txt:41`, `:296-326` |
| `lblTitulo` | etiqueta | — | «Alta de acción correctiva» o «Edición de acción correctiva» (también `Me.Caption`) | `:57-58, 97-99` |
| `NoConformidad` | texto, bloqueado | — | Descripción de la NC o «Sin definir» | `:45-50`, `.form.txt:339-340` |
| `AccionCorrectiva` | texto | `AfterUpdate` | Descripción de la acción; repinta el borde de obligatorios | `:149-153`, `.form.txt:262-263` |
| `Notas` | texto | — | Notas libres | `:103-107, 254` |
| `Responsable` | combo (lista de valores) | — | Se rellena con los nombres de los usuarios de Calidad; en el alta precarga el responsable de implantación de la NC | `:22, 59-61`, `FG.bas:3075-3105` |
| `FechaInicialMinima` | texto, bloqueado | `AfterUpdate` declarado | Solo lectura; sin handler en el `.cls` | `.form.txt:281-282, 298`, `:111-115` |
| `FechaFinPrevistaUltima` | texto, bloqueado | — | Solo lectura | `:116-120`, `.form.txt:717-718` |
| `FechaFinalUltima` | texto, bloqueado | — | Solo lectura | `:121-125`, `.form.txt:415-416` |
| `Estado` | texto, bloqueado | — | Estado de la AC; se vacía antes de actualizar | `:108-110, 176`, `.form.txt:472-473` |
| `ComandoActualizar` (rótulo «Contador», sugerencia «Actualizar») | botón | `OnClick` | Recalcula los datos de la AC y refresca otras pantallas | `:165-218`, `.form.txt:530-543` |
| `ComandoRegistrar` («Registrar») | botón | `OnClick` | Guarda la AC | `:225-292`, `.form.txt:890-903` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra sin guardar | `:155-163`, `.form.txt:786-797` |
| `lblfecha1`, `lblfecha2`, `lblfecha3`, `lblEstado1` | etiquetas | — | Rótulos de fechas y estado; solo visibles al editar | `:63-66, 76-79` |
| `Etiqueta31`, `Etiqueta38`, `Etiqueta52`, `Etiqueta61`, `EncabezadoDelFormulario`, `Detalle`, `PieDelFormulario` | rótulos y secciones | — | Sin código | `.form.txt:201-776` |

## Flujos

### Apertura (`Form_Open`, `:296-326`; `EstablecerDatos`, `:35-146`)

1. Rellena `NoConformidad`, el combo `Responsable` y decide el modo (`:45-55`).
2. **Alta** (sin AC activa): título de alta, precarga el responsable de la NC si lo tiene y oculta `ComandoActualizar`, las tres fechas, el estado y sus rótulos (`:56-71`); pinta bordes de obligatorios y termina (`:72-73`).
3. **Edición**: muestra esos controles, recarga la AC con `constructor.getACAuditoria` y vuelca acción, notas, estado, fechas y responsable (`:75-130`). Si no la encuentra, «No se han encontrado datos de la acción» (`:89-92`).
4. Si la NC está `Borrado`, bloquea la edición y deshabilita Registrar; si no, la habilita y pinta obligatorios (`:133-140`).
5. Ante cualquier error el formulario se cierra (`:316-325`).

### Registrar (`ComandoRegistrar_Click`, `:225-292`)

1. Si `EsTecnico = Sí`: «No tiene autorización para esa acción» (`:243-246`).
2. Crea la AC activa si no existe y la rellena con acción, notas y responsable (`:247-256`).
3. Llama a `ACAuditoriaOperaciones.Registrar` con el estado anterior (`:257-264`); lanza `Alta` o `Editado` (`:265-270`), refresca la línea de `FormNCAuditoriaGestion` si está abierto (`:273-275`) y cierra (`:280`).

### Actualizar (`ComandoActualizar_Click`, `:165-218`)

Vacía fechas y estado, llama a `ActualizarDatosACAuditoria` (`FG.bas:2503-2534`), recarga el formulario y, si el estado cambió, refresca `FormNCAuditoriaGestion`, recarga la NC de `FormNCAuditoria` y, si el detalle es `FormNCAuditoriaAcciones`, actualiza su lista (`:174-203`).

### Errores

Error 1000 ⇒ advertencia; otros ⇒ correo al administrador y `MsgBox` crítico (`:211-217, 284-290, 318-324`).

## Reglas, validaciones y cálculos

- **Validación** (`ACAuditoriaOperaciones.cls:56-107`): exige NC, acción y responsable (`:66-74, 83-86`); si la acción es nueva o cambió, rechaza duplicados dentro de la misma NC: «Existe otra acción correctiva con el mismo nombre…» (`:87-95`, `:17-46`).
- **Saneado:** elimina comillas simples y sustituye «;» por «:» en la acción (`ACAuditoriaOperaciones.cls:128-129`).
- **Estado:** se recalcula con `EstadoCalculadoTexto` en el alta y en la edición (`ACAuditoriaOperaciones.cls:186, 209`).
- **Alta:** `IdAccionCorrectiva` y `NAccion` se calculan en el objeto (`ACAuditoriaOperaciones.cls:133, 141`).
- **Si la NC estaba registrada**, el alta reinicia los indicadores (`ACAuditoriaOperaciones.cls:192-194`).
- **Obligatorios:** `RellenarBordeCamposObligatorios` (`:72, 139, 152`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Evidencia |
|---|---|---|---|
| Leer AC | `TbNCAuditoriaAccionCorrectivas` | SQL dinámico concatenado por `IDAccionCorrectiva` (DAO) | `constructor.bas:5081-5085` |
| Alta de AC | `TbNCAuditoriaAccionCorrectivas` | DAO `AddNew`/`Update` sobre la tabla completa | `ACAuditoriaOperaciones.cls:132, 152, 163-190` |
| Edición de AC | `TbNCAuditoriaAccionCorrectivas` | SQL dinámico concatenado + DAO `Edit`/`Update` | `ACAuditoriaOperaciones.cls:149-150, 196-215` |
| Responsables | usuarios de Calidad (`ColUsuariosCalidad`) | Objeto del entorno | `FG.bas:3086` |
| Registro de auditoría de cambios | `RegistrarLogAuditoria` | Alta/Edición de AC | `ACAuditoriaOperaciones.cls:191, 216` |

Sin transacción explícita: la AC y el log son operaciones separadas.

## Coexistencia

`SELECT *` y recordset DAO sobre `TbNCAuditoriaAccionCorrectivas`: con PostgreSQL vinculado, revisar el `AddNew` y el cálculo del identificador en cliente (`IDAccionCorrectivaCalculada`), el tipo de las fechas y el orden de escritura sin transacción. Concatenación de `IDAccionCorrectiva` numérico (`ACAuditoriaOperaciones.cls:150`).

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual:** `Estado` y las tres fechas son derivadas y de solo lectura (`.form.txt:282, 340, 416, 473, 718`).

**Posibles defectos (a verificar):**
- Mensaje invertido: con responsable vacío dice «Se ha introducido un responsable» (`ACAuditoriaOperaciones.cls:83-85`).
- `FechaInicialMinima` declara `AfterUpdate` sin handler (`.form.txt:298`).
- `ComandoActualizar` rotulado «Contador» (`.form.txt:542`).
- El alta lanza `Alta(m_EstadoAlInicio)`, pero `m_EstadoAlInicio` local de `ComandoActualizar` sombrea al del módulo solo en ese procedimiento; en `Registrar` se usa el de módulo, vacío si se abrió sin `OpenArgs` (`:6, 166, 267`; apertura sin argumentos en `Form_FormNCAuditoriaAcciones.cls:200`).
- Un `Err.Raise 100` en `MotivoNoOK` no se trata como error controlado (`ACAuditoriaOperaciones.cls:90`).

**Propuestas web pendientes de decisión:**
- Sustituir `ComandoActualizar` por recálculo automático al guardar.
- Validar la unicidad de la acción en servidor.

## Criterios de aceptación

1. Sin AC activa, el título es de alta y se ocultan el botón Actualizar, fechas y estado (`:56-71`).
2. Con AC activa se cargan acción, notas, estado, fechas y responsable (`:95-130`).
3. Un técnico no puede registrar (`:243-246`); con NC borrada no se puede editar (`:133-135`).
4. Registrar sin acción o sin responsable no guarda y avisa (`ACAuditoriaOperaciones.cls:71-74, 83-86`).
5. No se guarda una acción duplicada en la misma NC (`ACAuditoriaOperaciones.cls:87-95`).
6. Tras registrar se refresca la lista y se cierra (`:273-280`).
7. Actualizar recalcula estado y fechas y propaga el cambio (`:174-203`).

## Pendientes

- `ActualizarDatosCalculados`, `EstadoCalculadoTexto` e `IDAccionCorrectivaCalculada` (`ACAuditoriaOperaciones.cls:133, 186, 224`) no leídos.
- Receptores de los eventos `Alta` y `Editado` (`:267, 269`).
- `RegistrarLogAuditoria` (`ACAuditoriaOperaciones.cls:191`).
