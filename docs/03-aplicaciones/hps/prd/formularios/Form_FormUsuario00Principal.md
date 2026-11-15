# Form_FormUsuario00Principal

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormUsuario00Principal.cls` (63 líneas) y `.form.txt` (1038 líneas) · **PRD:** [README](../README.md)

Contenedor de detalle/edición del usuario **actual** (`m_ObjUsuarioActivo`, declarado en `modules/VariablesEntorno.bas:225`). Aloja dos subformularios: `SubFormLateral` (panel de navegación vertical) y `SubFormCentral` (panel de datos). No está enlazado: en `.form.txt` no aparece `RecordSource`, por lo que toda la persistencia y la carga útil se delegan en el `UsuarioHPS` global y en los formularios hijos. Las citas son `Form_FormUsuario00Principal.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Apertura desde `Form_FormInicial08UsuariosHistoricos` o equivalente de selección de actual | `DoCmd.OpenForm "FormUsuario00Principal", , , …` con el `m_ObjUsuarioActivo` ya rellenado por el llamador | `modules/VariablesEntorno.bas:225` para el global; no hay `DoCmd.OpenForm` en `src` que apunte a este nombre |
| Subformulario `SubFormLateral` | Cargado como `FormUsuario01Lateral` desde código en la apertura | `:50-52`, `.form.txt:835, 841` |
| Subformulario `SubFormCentral` | Alojado en `Detalle` (a la derecha del lateral); el `.form.txt` no le asigna `SourceObject`, lo hace cada `lbl*_Click` del lateral | `.form.txt:873, 880` |

Sale con `DoCmd.Close acForm, Me.Name, acSaveNo` desde el botón `cmdSalir` (`:9-12`) o desde el manejador de errores tras un `1000` interno (`:62`). No abre otros formularios.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta (encabezado) | — | Título centrado; reescrito en `Form_Load` con «EDICIÓN DE …» o «DETALLE DE …» en mayúsculas | `:42-43`, `.form.txt:794, 804` |
| `SubFormLateral` | subformulario | — | Aloja `FormUsuario01Lateral` (panel de navegación) | `:50-52`, `.form.txt:835, 841` |
| `SubFormCentral` | subformulario | — | Aloja la vista activa (`FormUsuario02Edicion`, `FormUsuario02HPS`, `FormUsuario02Anexos`, …), seleccionada por el lateral | `.form.txt:873, 880` |
| `cmdSalir` («Salir») | botón | `OnClick` → `CmdSalir_Click` | Cierra el formulario sin guardar | `:8-12`, `.form.txt:924, 931` |
| (formulario) | — | `OnLoad` → `Form_Load` | Configura modo lectura/edición, títulos, estilo y carga del subform lateral | `:15-54`, `.form.txt:370` |
| `EncabezadoDelFormulario`, `PieDelFormulario` | secciones | — | Cabecera y pie de página; el estilo se aplica tras `EstablecerControlesDatos` | `.form.txt:787, 917` |

Sin ribbon de Access ni cinta personalizada. `NavigationButtons = NotDefault` y `RecordSelectors = NotDefault` (`.form.txt:6-8`).

## Flujos

### Apertura (`Form_Load`, `:15-54`)

1. `AjustarTamaño Me` reescala el formulario a las dimensiones del monitor (`:23`; helper en `FG.bas:901`).
2. Lee la variable temporal `SoloLectura` y pone `Me.AllowEdits` en consecuencia (`:27-31`); si no está definida se queda editable.
3. Si `m_ObjUsuarioActivo` es `Nothing`, rellena `m_Error` con «No se puede determinar el usuario activo» y relanza error interno `1000` (`:32-35`).
4. Marca `m_ObjUsuarioActivo.ParaAlta = EnumSiNo.No` para que cualquier `Registrar` posterior del objeto entre en modo edición (`:36`).
5. Elige título según `AllowEdits`: «EDICIÓN DE …» editable o «DETALLE DE …» solo lectura, en ambos casos con `UCase(NombreCompleto)` del usuario activo (`:37-41`). Lo aplica a `lblTitulo` y `Caption` (`:42-43`).
6. `EstablecerControlesDatos Me, m_Error` aplica la paleta de colores del entorno (helper en `FG.bas:1005`); un error se eleva como `1000` (`:45-48`).
7. Si `SubFormLateral.SourceObject` no es ya `FormUsuario01Lateral`, lo carga (`:50-52`).
8. `CopiarEstiloBoton Me.cmdSalir` aplica el aspecto de botón del entorno al botón «Salir» (helper en `FG.bas:2435`, firma de un solo argumento; `:53`).
9. En error: compone el mensaje, envía `CorreoAlAdministrador` si `Err.Number <> 1000`, muestra un `MsgBox` crítico y **cierra el formulario** sin guardar (`:55-62`). El valor de retorno del `MsgBox` se guarda en el global `pregunta`, declarado en `modules/VariablesEntorno.bas:220`.

### Cierre (`CmdSalir_Click`, `:8-12`)

`DoCmd.Close acForm, Me.Name, acSaveNo` con `On Error Resume Next`, sin comprobaciones (`:9-12`). No pregunta nada: la cancelación corresponde al subformulario `SubFormCentral`, que es quien maneja cambios pendientes.

### Errores

Manejador único en `:55-61`, patrón repetido en el resto del PRD: si `Err.Number <> 1000`, el helper compone «Al abrir el formulario se ha producido el error: …», llama a `CorreoAlAdministrador` y muestra `MsgBox` con título «Error» (`:57-60`). El formulario siempre se cierra con `acSaveNo` (`:61`), así que ningún error de apertura deja estado parcial.

## Reglas, validaciones y cálculos

- **Exigencia de `m_ObjUsuarioActivo`.** Único guard al abrir; sin usuario activo, el formulario no llega a pintarse (`:32-35`).
- **Modo lectura global.** `SoloLectura="Sí"` se evalúa una sola vez y bloquea el formulario completo; el lateral y el subform central heredan esa restricción (`:27-31`). El subform `FormUsuario02DatosObservacion` (mismo PRD) re-evalúa esta variable por su cuenta.
- **Forzado de `ParaAlta = No`.** El contenedor garantiza que `m_ObjUsuarioActivo` no llegará al registrar como alta, redirigiendo la responsabilidad a `FormUsuario02Edicion`, `FormUsuario02HPS` y los formularios `02*` (`:36`).
- **Carga perezosa del subform lateral.** Se asigna `SourceObject = "FormUsuario01Lateral"` solo si difiere del actual; volver a entrar no recarga el subform si ya está, evitando reinicializar su estado (`:50-52`).
- **No concurrencia.** El contenedor no impide abrir dos instancias en paralelo; el global `m_ObjUsuarioActivo` se sobrescribe en la segunda, dejando la primera con datos obsoletos. Sin `WindowMode = acDialog` ni `Modal`.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Recalcular el aspecto del contenedor | `m_ObjEntorno` (paleta) | en memoria | `:45`, `FG.bas:1005` |
| Plantilla de botones | `m_ObjEntorno` (estilo) | en memoria | `:53`, `FG.bas:2435` |
| Selección de subform lateral | `SubFormLateral.SourceObject` | propiedad del subform | `:50-52` |

Sin `RecordSource`, sin `INSERT`/`UPDATE`/`DELETE`. Las escrituras reales las hacen los subformularios `02*` desde el `m_ObjUsuarioActivo` global.

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Cosas que ya se ven: el formulario no declara `PopUp = True` ni `Modal = True`, así que en web se traducirá a una ruta con navegación explícita; la distinción «EDICIÓN» / «DETALLE» se basa en `AllowEdits`, no en una propiedad separada; `SourceObject = "FormUsuario01Lateral"` está escrito como literal y duplica una cadena que también aparece en `Form_FormUsuarioHistorico00Principal.cls:78` (mismo patrón, otro prefijo).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La asignación `m_ObjUsuarioActivo.ParaAlta = EnumSiNo.No` es incondicional (`:36`): si el llamador dejó el usuario en `Sí` esperando un alta, el contenedor lo revierte sin avisar.
- `CopiarEstiloBoton Me.cmdSalir` se llama con un solo argumento (`:53`); la versión de `FG.bas:2435` declara además un `ByRef p_Error`. La llamada con un argumento es legal en VB, pero deja el control sobre errores en el helper; conviene verificar la firma exacta.
- `CmdSalir_Click` ignora si hay cambios pendientes en el `SubFormCentral` (`:8-12`); el contenedor asume que el subform ya preguntó. Si el subform central es uno sin `CmdSalir` propio (p. ej. un visor), el cierre es silencioso.
- El manejador de errores trata todo `1000` como ya controlado (`Err.Number <> 1000`), lo que silencia cualquier `Err.Raise 1000` legítimo re-lanzado por los helpers (`:56-58`).
- `Form_Load` no cancela la apertura inicial con `DoCmd.CancelEvent` ni con `Me.Undo`: si la validación falla, el formulario igual se ha abierto y pintado antes de cerrarse.
- El título se construye con `UCase(NombreCompleto)` (`:38-40`) sin decoración («» … «»), así que una `NombreCompleto` con espacios duplicados o caracteres de control sale tal cual a la barra de título.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: pasar el usuario por la ruta (`/usuarios/:id/editar`) en vez de por variable global; traducir `AllowEdits` a un `disabled` global ligado al rol; eliminar la duplicación de `SourceObject` entre `FormUsuario00Principal` y `FormUsuarioHistorico00Principal` con un solo componente configurable.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin `m_ObjUsuarioActivo` rellenado muestra el aviso y cierra el formulario sin lanzar el subform lateral.
2. Con `SoloLectura="Sí"`, los campos del subform central quedan no editables aunque el subformulario concreto declare `AllowEdits=True`.
3. El título refleja el estado: «EDICIÓN DE …» editable, «DETALLE DE …» solo lectura, con el nombre del usuario en mayúsculas.
4. Al pulsar «Salir», el formulario se cierra y los globales `m_ObjUsuarioActivo` y `m_ObjObservacionActiva` mantienen su valor para el siguiente consumidor.
5. El subform lateral siempre está cargado como `FormUsuario01Lateral` antes de que el usuario pueda hacer clic en sus etiquetas.

## Pendientes

- Nombre del formulario que abre `Form_FormUsuario00Principal`: en `src` no hay `DoCmd.OpenForm` con ese nombre; la búsqueda en menús y macros del binario no se ha hecho.
- Firma exacta de `CopiarEstiloBoton` en `FG.bas:2435` (paso de `m_Error` por referencia): la llamada actual en `:53` no lo usa.
- Comportamiento de `AjustarTamaño Me` (`FG.bas:901`) en pantallas con DPI no estándar.
- `pregunta` (retorno de `MsgBox`) nunca se evalúa en este formulario (`VariablesEntorno.bas:220`).
