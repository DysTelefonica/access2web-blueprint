# Form_FormInicial02ConsultasPrincipal

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial02ConsultasPrincipal.cls` (317 líneas) y `.form.txt` (1310 líneas) · **PRD:** [README](../README.md)

Panel central de consultas del usuario: aloja la lista de tipos de consulta a la izquierda (`ListaConsultas`), el formulario de parámetros a la derecha (cambiado en caliente por `SourceObject`) y debajo la rejilla de resultados (`SubFormResultados`). Es **no enlazado** (el `.form.txt` solo trae el `RecSrcDt` interno del binario, sin `RecordSource` textual a `TbDatosLocal`): todos los datos viven en el global `m_ObjUsuarioActivo` (`modules/VariablesEntorno.bas:225`) y en `m_ObjEntorno.ColConsultas` / `m_ObjEntorno.ColTiposConsultasFormulario`. La pieza estructural que lo define es el patrón «cambiar el `SourceObject` de un panel según un valor»: está decidido en `ListaConsultas_Click` (`:223-286`), reutilizado en `Form_FormUsuario01Lateral.lbl*_Click` (`Form_FormUsuario01Lateral.cls:32-110`) y replicado, con la misma mecánica, por `EstablecerFuenteDatosIndicadores` (`modules/Funciones Generales.bas:3056-3146`) y `EstablecerFiltroUsuarios` (`modules/Filtro.bas:5-111`). Las citas son `Form_FormInicial02ConsultasPrincipal.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial00Principal` (`SubFormCentral`) | Alojado como `FormInicial02ConsultasPrincipal` con `SourceObject` fijo del `.form.txt` | `.form.txt:645-646` (cabecera del subform) |
| `ListaConsultas` (clics de la lista izquierda) | Cambia la consulta activa y dispara `ListaConsultas_Click` | `.form.txt:587-619`, `:223-286` |
| `m_ObjUsuarioActivo` (global) | Lectura y reinicio: los `Click` lo consultan (`If Is Nothing Then …`) y lo ponen a `Nothing` cuando cambia la consulta; lo rellena `FormInicial03ConsultasDatos` al elegir fila | `:72-77`, `:237-239`, `:246-248`; `VariablesEntorno.bas:225` |
| `ComandoNuevoUsuario_Click` | `DoCmd.OpenForm "FormUsuario02Alta"` (sin `OpenArgs`) | `:141-160` |
| `ComandoEditarUsuario_Click` | `DoCmd.OpenForm "FormUsuario00Principal"` después de cerrar el existente | `:44-68` |
| `ComandoAnexos_Click` | `DoCmd.OpenForm "FormUsuario00Principal", OpenArgs:="No"` y luego `Form_FormUsuario01Lateral.lblAnexos_Click` | `:14-43` |
| `ComandoEliminarUsuario_Click` | `m_ObjUsuarioActivo.Borrar m_Error` y reconstruye el entorno con `EVE` | `:69-113` |
| `ComandoListaEnExcel_Click` | `ConsultaAExcel(Me.SubFormResultados.Form, …)` y abre el `.xlsx` con `Ejecutar` (Shell) | `:114-140` |
| `m_FormInicial00Principal.ActualizadoUsuarios` (contenedor) | Requery del `SubFormResultados` y reselecciona el ítem 0 de `ListaConsultas` | `:314-318`, `:12` |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaConsultas` | `ListBox` de dos columnas (`Value List`) | `OnClick` → `ListaConsultas_Click` | Cambia la consulta activa: reescribe `SubFormParametrosConsulta.SourceObject` desde `ColTiposConsultasFormulario` y de `SubFormResultados` a `FormInicial03ConsultasDatos` | `:223-286`, `.form.txt:587-619` (lista) |
| `SubFormParametrosConsulta` | subformulario | — | Aloja el formulario de parámetros activo (`FormInicial03ConsultasEstado` por defecto en diseño) | `:166, :265`, `.form.txt:620-645` |
| `SubFormResultados` | subformulario | — | Aloja `FormInicial03ConsultasDatos`, que pinta la rejilla con el `RecordSource` reescrito por `EstablecerFiltroUsuarios` | `:251, :312`, `.form.txt:645-671`, `Filtro.bas:107` |
| `ComandoListaEnExcel` | botón (icono) | `OnClick` → `ComandoListaEnExcel_Click` | Exporta el `RecordSource` actual a `.xlsx` y lo abre | `:114-140`, `.form.txt:671-714` |
| `ComandoNuevoUsuario` («Nuevo») | botón (`Enabled = NotDefault` en diseño) | `OnClick` → `ComandoNuevoUsuario_Click` | Cierra `FormUsuario02Alta` si está abierto y lo abre para alta | `:141-160`, `.form.txt:714-759` |
| `ComandoEditarUsuario` («Detalle Usuario») | botón (`OnClick`) | `OnClick` → `ComandoEditarUsuario_Click` | Cierra `FormUsuario00Principal` si está abierto y lo reabre | `:44-68`, `.form.txt:759-805` |
| `ComandoAnexos` («Anexos») | botón (`OnClick`) | `OnClick` → `ComandoAnexos_Click` | Abre `FormUsuario00Principal` con `OpenArgs:="No"` y dispara la pestaña «Anexos» del lateral | `:14-43`, `.form.txt:805-852` |
| `ComandoEliminarUsuario` («Eliminar») | botón (`Enabled = NotDefault` en diseño) | `OnClick` → `ComandoEliminarUsuario_Click` | Pide confirmación y borra al usuario activo y sus registros asociados | `:69-113`, `.form.txt:852-899` |
| `Comando5`, `Comando6` | botones «Anexos» / «Eliminar» duplicados | — | Botones sin handler declarado, marcados `OverlapFlags =93`; el `.form.txt` los entrega como código muerto (`.form.txt:899, 945`) | `.form.txt:899-988` |
| (formulario) | — | `OnLoad` → `Form_Load` | Manda el `SourceObject = ""`, rellena `ListaConsultas` con `EstablecerTablasConsultas`, lee `TempVars("SoloLectura")` y dispara `ListaConsultas_Click` | `:161-222`, `.form.txt` |

Sin ribbon de Access; `NavigationButtons = NotDefault`, `RecordSelectors = NotDefault`, `ControlBox = NotDefault`, `CloseButton = NotDefault` (`.form.txt:5-9`). `DefaultView = 0`, `ScrollBars = 0`, `ViewsAllowed = 1` (`.form.txt:11-13`): vista única continua.

## Flujos

### Apertura (`Form_Load`, `:161-222`)

1. `m_Error = ""`, `Me.AllowEdits = False` y `Me.SubFormParametrosConsulta.SourceObject = ""` (`:165-167`): deja el panel de parámetros vacío hasta que `EstablecerTablasConsultas` rellene la lista.
2. `EstablecerTablasConsultas m_Error` (`:168`; helper privado en `:287-313`) vacía `ListaConsultas.RowSource` y lo rellena con `Format(m_ID, "00") & ";" & m_ObjConsulta.Nombre` por cada clave de `m_ObjEntorno.ColConsultas` (`:298-305`). Si la colección es `Nothing`, la lista queda vacía pero **no** eleva error (`:298`).
3. `EstablecerControlesDatos Me, m_Error` (`:174`; `FG.bas:1005-1055`) con `Err.Raise 1000` encadenado (`:175-177`).
4. Lee `Application.TempVars("SoloLectura")` (`:181`): con valor `"Sí"` deshabilita `ComandoNuevoUsuario` y `ComandoEliminarUsuario`, y cambia el `Caption` de `ComandoEditarUsuario` a «Detalle Usuario» (`:182-185`); en otro caso los tres quedan habilitados y el caption pasa a «Editar Usuario» (`:186-190`).
5. `CopiarEstiloBoton` aplicado a los tres botones (`ComandoNuevoUsuario`, `ComandoEditarUsuario`, `ComandoAnexos`) (`:193-204`; helper en `FG.bas:2435`, firma con `ByRef p_Error`).
6. `On Error Resume Next` + `If FormularioAbierto("FormInicial00Principal") Then Set m_FormInicial00Principal = Forms("FormInicial00Principal")` (`:207-211`): asociación `WithEvents` no fatal si el contenedor no está abierto. `Err.Clear` y vuelve a `On Error GoTo errores` (`:212-214`).
7. `Me.ListaConsultas.Selected(0) = True` y dispara `ListaConsultas_Click` (`:217-218`). Termina con `Me.AllowEdits = True` (`:220`).
8. En error: «El método abrir formulario ha dado el error: …», `CorreoAlAdministrador` si `Err.Number <> 1000`, `MsgBox` y `DoCmd.Close acForm, Me.Name, acSaveNo` (`:213-221`); el `AllowEdits = True` se restituye antes del `DoCmd.Close` (`:218`).

### Clic en la lista (`ListaConsultas_Click`, `:223-286`)

1. Lee `Me.ListaConsultas.Column(0)` y lo guarda en `m_IDListaSeleccionada` (`:230`).
2. **Rama cabecera no numérica** (`:231-241`): si la columna no es numérica — pasa en la cabecera del listbox al pulsar fuera de cualquier ítem — registra `m_IDListaSeleccionadaUltima`, vacía los dos `SourceObject`, pone `m_ObjUsuarioActivo = Nothing` y deshabilita `ComandoListaEnExcel`, `ComandoAnexos` y `ComandoEditarUsuario` (`:232-240`). Es la única rama que no pinta datos.
3. **Rama idempotente** (`:242-243`): si `m_IDListaSeleccionada = m_IDListaSeleccionadaUltima`, sale sin tocar nada. La primera apertura siempre entra, porque `:217-218` fuerza `Selected(0) = True` después del reset a «» de `:167`.
4. Marca `m_IDListaSeleccionadaUltima = m_IDListaSeleccionada`, vacía `m_ObjUsuarioActivo`, habilita `ComandoEditarUsuario` y `ComandoAnexos`, y reescribe `Me.SubFormResultados.SourceObject = "FormInicial03ConsultasDatos"` (`:244-251`). `Set m_FormInicial03ConsultasDatos = Me.SubFormResultados.Form` (`:252`) para poder reaccionar a eventos del subform.
5. Resuelve `m_SubFormParametrosConsultaNombreForm = m_ObjEntorno.ColTiposConsultasFormulario(CStr(CLng(m_IDListaSeleccionada)))` (`:263-265`); si la colección es `Nothing` o no contiene la clave, eleva `1000` con «No se han encontrado datos de la consulta nº …» (`:254-262`).
6. Después de la lectura, evalúa `m_Error = m_ObjEntorno.Error` (`:266`); si quedó algo del getter, vuelve a elevar `1000` (`:267-269`). Si la colección devolvió cadena vacía, eleva el mismo error (`:270-273`).
7. `Me.SubFormParametrosConsulta.SourceObject = m_SubFormParametrosConsultaNombreForm` (`:275`) y `Me.ComandoListaEnExcel.Enabled = True` (`:276`).
8. En error: «El método click ha dado el error: …», `CorreoAlAdministrador` si `Err.Number <> 1000`, `MsgBox`; **repone** `m_IDListaSeleccionadaUltima = 0` (`:278-283`), pero **no** cierra el formulario.

### Salir a edición (`ComandoEditarUsuario_Click`, `:44-68`)

Pide `m_ObjUsuarioActivo` no nulo (`:48-53`); si falta, eleva «Se ha de seleccionar el usuario primero». Cierra `FormUsuario00Principal` si está abierto y abre `FormUsuario00Principal` sin `OpenArgs` (`:58-62`). El lateral cargará `lblDatosGenerales` por defecto (ver `Form_FormUsuario01Lateral.cls:20`). En error: mensaje, `CorreoAlAdministrador`, `MsgBox` (no cierra).

### Salir a anexos (`ComandoAnexos_Click`, `:14-43`)

Igual que `ComandoEditarUsuario_Click` pero con `OpenArgs:="No"` (`:24`) y luego `Form_FormUsuario01Lateral.lblAnexos_Click` (`:26`), que fuerza la pestaña «Anexos». Como `OpenArgs:="No"` se pasa siempre, queda como invariante del anexo (no editable). En error: mensaje, `CorreoAlAdministrador`, `MsgBox` (no cierra).

### Nuevo usuario (`ComandoNuevoUsuario_Click`, `:141-160`)

`Set m_ObjUsuarioActivo = Nothing` (`:148`), cierra `FormUsuario02Alta` si está abierto y `DoCmd.OpenForm "FormUsuario02Alta"` (`:148-158`). El formulario de alta (`Form_FormUsuario02Alta`) rellena `m_ObjUsuarioActivo` con un `UsuarioHPS` nuevo (su `Form_Load` lo crea) y al guardar reentra por su cola llamando a `Form_FormInicial02ConsultasPrincipal.ComandoEditarUsuario_Click` (`Form_FormUsuario02Alta.cls:256`).

### Eliminar usuario (`ComandoEliminarUsuario_Click`, `:69-113`)

Confirma con `MsgBox(…, vbYesNo + vbDefaultButton2)` (`:80-81`); si no es «Sí», sale (`:83-84`). Llama a `m_ObjUsuarioActivo.Borrar m_Error` con `DoCmd.Hourglass True/False` alrededor (`:91-95`). Si dejó error, eleva 100 (no 1000: `:96-99`). Si tras borrar `m_ObjEntorno Is Nothing`, reconstruye el entorno vía `EVE m_Error` y eleva 1000 si queda error (`:100-105`). En error: mensaje, `CorreoAlAdministrador`, `MsgBox` (no cierra).

### Listado en Excel (`ComandoListaEnExcel_Click`, `:114-140`)

`m_URLExcelGenerado = ConsultaAExcel(Me.SubFormResultados.Form, , , EnumSiNo.Sí, m_Error)` (`:123`; helper en `FG.bas:382-437`); eleva 1000 si quedó `m_Error`. `Ejecutar Me.hWnd, "open", m_URLExcelGenerado, …` lanza el `.xlsx` con `Shell` (`FG.bas:2251`). En error: mensaje, `CorreoAlAdministrador`, `MsgBox` (no cierra).

### Recibir la notificación de usuarios actualizados (`m_FormInicial00Principal_ActualizadoUsuarios`, `:314-318`)

`Me.SubFormResultados.Form.Requery`, `Me.ListaConsultas.Selected(0) = True`, `ListaConsultas_Click`: repinta la rejilla y vuelve al primer ítem de la lista. Se dispara cuando el contenedor envía `ActualizadoUsuarios` (gestionado fuera de este `.cls`).

## Reglas, validaciones y cálculos

- **Inyección de subform por `SourceObject`.** Cada cambio de consulta reescribe `SubFormParametrosConsulta.SourceObject` (`:275`) y `SubFormResultados.SourceObject` (`:251`). La decisión de qué `Form_*` montar sale de la colección `m_ObjEntorno.ColTiposConsultasFormulario`, no de un `DoCmd.OpenForm` con nombre en duro.
- **Reset del usuario al cambiar de consulta.** `m_ObjUsuarioActivo` se vacía en `ListaConsultas_Click` (`:246-248`), en la rama cabecera (`:237-239`) y al abrir alta (`:148`). La lista así nunca arrastra el usuario seleccionado de la consulta anterior. El nuevo `m_ObjUsuarioActivo` lo fija `FormInicial03ConsultasDatos` al elegir un ítem.
- **El id leído del listbox como literal.** `m_IDListaSeleccionada = Nz(Me.ListaConsultas.Column(0), "")` (`:230`): se interpreta como número por `CLng(m_IDListaSeleccionada)` en `:264`. Si la cabecera del listbox entrega la cadena vacía o un texto, la rama `:231-241` sale sin pintar.
- **`SoloLectura` por `Application.TempVars`.** El estado de solo lectura es un global de sesión (`FG.bas:43, 55, 2151-2156, 2447-2448`); este formulario lo lee una sola vez (`:181`) y deshabilita comandos, pero **no** los oculta: el menú derecho sigue mostrándose. Los parámetros del menú lateral («indicadores») viajan por el mismo mecanismo (`FG.bas:1065-1098` lee `Application.TempVars("Command")` y reparte `Campo=Valor` en un diccionario): es el mismo `TempVars` y, por tanto, tiene el mismo ciclo de vida (sesión de Access). Sin equivalente directo en web — se traduce a un parámetro de ruta o a un contexto.
- **Idempotencia al cambiar de consulta.** `:242-243` evita recargar la consulta si el usuario hace clic otra vez sobre el mismo ítem; `m_IDListaSeleccionadaUltima` se reinicia a `0` solo cuando hay error (`:282`) o por la rama cabecera (`:232`).
- **Alta y vuelta a edición de un solo paso.** El alta «NUEVO» abre `FormUsuario02Alta`; al guardar, ese formulario cierra y llama a `Form_FormInicial02ConsultasPrincipal.ComandoEditarUsuario_Click` (`Form_FormUsuario02Alta.cls:256`), entrando de nuevo en `:44-68`. El global `m_ObjUsuarioActivo` ya lleva el nuevo ID, así que la edición resultante abre el recién creado.
- **`EVE` para reconstruir el entorno tras borrar.** El borrado puede dejar `m_ObjEntorno Is Nothing`; `:100-105` reconstruye vía `EVE` y relanza si la reconstrucción falla. No se documenta qué hace `EVE` en este PRD.
- **`m_FormInicial03ConsultasDatos` no se libera.** Se reasigna en cada `ListaConsultas_Click` (`:252`); no se cierra explícitamente al cerrar el formulario.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Catálogo de consultas (lista izquierda) | `m_ObjEntorno.ColConsultas` | `Dictionary` en memoria; `Format(m_ID, "00") & ";" & m_ObjConsulta.Nombre` por clave | `:298-305` |
| Asignación de `Form_*` por id | `m_ObjEntorno.ColTiposConsultasFormulario` | `Dictionary` en memoria | `:254-275` |
| Borrado del usuario activo | `m_ObjUsuarioActivo.Borrar` | método de `UsuarioHPS` (`classes/UsuarioHPS.cls`, ubicación exacta por verificar) | `:91` |
| Exportación a Excel | `Me.SubFormResultados.Form` (`RecordSource` vivo) | `ConsultaAExcel` (`FG.bas:382-437`) | `:123` |
| Apertura del `.xlsx` | sistema operativo | `Shell` vía `Ejecutar` (`FG.bas:2251`) | `:127` |

Sin `RecordSource` ni `INSERT`/`UPDATE`/`DELETE` directos: este formulario no toca `TbUsuarios`, `TbDatosLocal` ni `TbConfiguracionHPS` por su cuenta; las escrituras reales van por `m_ObjUsuarioActivo` y por los subformularios inyectados.

## Coexistencia

Pendiente de `../coexistencia.md`. Cosas que ya se ven: el formulario reutiliza el patrón «cambiar `SourceObject` por valor» que también preside `EstablecerFiltroUsuarios` (`Filtro.bas:5-111`) y `EstablecerFuenteDatosIndicadores` (`FG.bas:3056-3146`): en web se traducirá a enrutado dinámico (`/consultas/:id` o `/consultas/:id/parametros`) con un componente único para la rejilla; el botón `Comando5`/`Comando6` del `.form.txt` (`:899, 945`) es código muerto que el runtime nunca invoca — pendiente de borrado en la migración; la pareja `m_ObjUsuarioActivo` ↔ `m_ObjUsuarioActivo = Nothing` se reescribe tres veces en `:234, 247, 148` y debe consolidarse en web con una sola fuente (parámetro de ruta o store).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `m_SQLSinWhere` se declara en `:10` pero **no** se usa en ningún sitio del `.cls`: código muerto.
- `Set m_FormInicial03ConsultasDatos = Me.SubFormResultados.Form` (`:252`) se ejecuta en cada cambio de consulta, pero el módulo no declara el evento `m_FormInicial03ConsultasDatos_*`: si el subform dispara eventos que el contenedor captura a través del `WithEvents`, este formulario no los procesa — verificar que la recepción real de selección de fila viva en `FormInicial03ConsultasDatos` y no en este.
- `ComandoEliminarUsuario_Click` eleva `Err.Raise 100` (no 1000) tras un fallo de `Borrar` (`:97`): el manejador no distingue 100 de cualquier otro número, así que el guard `If Err.Number <> 1000 Then CorreoAlAdministrador` (`:108`) **no** se ejecuta y el envío de correo depende solo de que `:91` no haya dejado ya `m_Error`. La elevación con 100 es no estándar para este PRD.
- `ComandoEditarUsuario_Click`, `ComandoAnexos_Click`, `ComandoNuevoUsuario_Click`, `ComandoEliminarUsuario_Click`, `ComandoListaEnExcel_Click` **no** cierran el formulario en error: solo muestran `MsgBox`. En modo lectura el botón «Eliminar» queda deshabilitado (`:184`), pero el «Editar» sigue activo con `Caption = "Detalle Usuario"` (`:184`); sin embargo, `ComandoEditarUsuario_Click` no trata `OpenArgs:="Sí"` y abre siempre editable por defecto. Verificar intención.
- `Form_Load` deshabilita `ComandoEliminarUsuario` por `SoloLectura` (`:184`) pero **no** deshabilita `ComandoAnexos`, que sigue activo: el alta está cerrada (ComandoNuevo), pero los anexos siguen editándose. A verificar.
- `Aplicación.TempVars("SoloLectura")` se evalúa en `Form_Load` (`:181`) sin restaurar el flag por defecto si está vacío: la rama `:186-190` (Else) asume editable incluso si la variable no existe — riesgo de comportamiento por defecto no documentado.
- `m_IDListaSeleccionadaUltima = 0` se repone a «0» en la rama de error (`:282`), pero no se reescribe `m_IDListaSeleccionada`: el siguiente clic compara `Column(0)` con `0`, falla la igualdad y vuelve a cargar el panel de parámetros aunque el ítem activo sea el mismo.
- `m_FormInicial00Principal.ActualizadoUsuarios` (`:314-318`) siempre selecciona el ítem 0 de la lista — borra el ítem que el usuario tenía activo.
- `pregunta = MsgBox(…, vbCritical, "Error")` en cada `errores:` (`:38, :64, :107, :131, :153`) ignora la respuesta: el retorno solo se guarda, no se usa para decidir nada dentro del flujo.
- El cierre `DoCmd.Close acForm, Me.Name, acSaveNo` solo se llama desde `Form_Load` en error (`:221`) — este formulario no expone un botón «Salir» propio; el cierre llega por el contenedor padre (`Form_FormInicial00Principal`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: consolidar el triple «cambiar `SourceObject` por valor» de este formulario, `EstablecerFuenteDatosIndicadores` y `EstablecerFiltroUsuarios` en un componente declarativo (`/consultas/:tipo`) con un único `<router-outlet>` parametrizado; eliminar la duplicación de `m_ObjUsuarioActivo = Nothing` consolidando el global con un contexto de ruta; eliminar `m_SQLSinWhere` y `Comando5`/`Comando6` en la migración; reemplazar el uso de `Application.TempVars` por un mecanismo de sesión o por un parámetro explícito, de modo que `SoloLectura` y los parámetros del menú derecho vivan en una sola fuente; sustituir `MsgBox` modal por toasts no bloqueantes que no requieran un `Dim pregunta As Long` global.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir la app, la lista de la izquierda muestra los identificadores `01;Todos`, `02;Por Jurídica que Contrata`, etc., y el primer ítem queda seleccionado.
2. Elegir un ítem monta el formulario de parámetros de la derecha con el `SourceObject` resuelto desde `m_ObjEntorno.ColTiposConsultasFormulario(id)`.
3. La rejilla de abajo se inyecta siempre como `FormInicial03ConsultasDatos`, lista para que `EstablecerFiltroUsuarios` la reescriba.
4. Con `SoloLectura="Sí"`, los botones «Nuevo» y «Eliminar» quedan deshabilitados y el caption del «Editar» pasa a «Detalle Usuario».
5. Si el ítem activo no existe en `ColTiposConsultasFormulario`, el `MsgBox` correspondiente avisa y la lista no cambia de `m_IDListaSeleccionadaUltima` (queda a `0`).
6. Tras eliminar un usuario, `m_ObjEntorno` se reconstruye vía `EVE`; si la reconstrucción falla, el formulario se cierra por el flujo de error.
7. Pulsar dos veces el mismo ítem no recarga el formulario de parámetros.
8. La cabecera de la lista (no numérica) deja los tres botones de acción deshabilitados, `SubFormParametrosConsulta` y `SubFormResultados` con `SourceObject = ""`, y `m_ObjUsuarioActivo = Nothing`.

## Pendientes

- Contenido real de `m_ObjEntorno.ColConsultas` y `m_ObjEntorno.ColTiposConsultasFormulario`: se construyen en `classes/Entorno.cls` (no en `src/`), por lo que la correspondencia entre `01;Todos` y `FormInicial03ConsultasTodos` es opaca en el snapshot.
- Eventos que dispara `FormInicial03ConsultasDatos` y a quién notifica: el formulario asigna el `WithEvents` (`:12`, `:252`) pero no expone ningún handler `m_FormInicial03ConsultasDatos_*` en este `.cls`. Verificar si la selección de fila la gestiona el propio subform.
- Lista de `m_ObjConsulta.Nombre` y `m_ObjConsulta.Id` que devuelve `ColConsultas`: nombres exactos por verificar en `classes/Consulta.cls`.
- Definición exacta de `EVE` (reconstructor del entorno): vive fuera de `src/` y no se ha reproducido aquí.
- Firma exacta de `CopiarEstiloBoton` (`FG.bas:2435`): la llamada en `:193-204` usa la versión con `ByRef p_Error`; en este PRD hay también llamadas con un solo argumento (a verificar, por ejemplo en `Form_FormUsuario00Principal.cls:53`).
- `Borrar` y `cache_usuario_regenerar` en `classes/UsuarioHPS.cls`: el segundo se llama desde `Form_FormUsuario02Alta.cls:248` tras alta; no se llama desde este formulario, así que la consistencia de caché tras eliminar queda pendiente de un barrido equivalente.
