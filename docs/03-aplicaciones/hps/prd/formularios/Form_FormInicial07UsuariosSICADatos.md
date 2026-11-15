# Form_FormInicial07UsuariosSICADatos

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial07UsuariosSICADatos.cls` (39 líneas) y `.form.txt` (826 líneas) · **PRD:** [README](../README.md)

Hoja-datos de los usuarios SICA del bloque 07. Es un **subformulario hoja** montado por `Form_FormInicial07UsuariosSICA` mediante `Me.SubFormResultados.SourceObject = "Form.FormInicial07UsuariosSICADatos"` (`Form_FormInicial07UsuariosSICA.form.txt:896-897`). El formulario se muestra como datasheet (`DefaultView =2`, `.form.txt:14`; `AllowFormView =0`, `.form.txt:113`) con 14 columnas enlazadas a `TbUsuariosSICALocal`. Se suscribe al bus de indicadores `modIndicadores.m_Indicadores` (instancia única de `clsIndicadoresBus`) para reaccionar a `RaiseLifecycleMoveCompleted` (`:36-40`). Las citas son `Form_FormInicial07UsuariosSICADatos.cls:línea` salvo indicación; `FG.bas` abrevía `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial07UsuariosSICA` | `Me.SubFormResultados.SourceObject = "Form.FormInicial07UsuariosSICADatos"` | `Form_FormInicial07UsuariosSICA.form.txt:896-897` |

Abre nada por sí mismo. El padre reutiliza `Me.SubFormResultados.Form` para `ConsultaAExcel` (`Form_FormInicial07UsuariosSICA.form.txt:1617`).

## Controles y eventos

`RecordSource` en diseño: `SELECT TbUsuariosSICALocal.* FROM TbUsuariosSICALocal;` (`.form.txt:68`). `OnCurrent` y `OnLoad` enganchados a procedimientos del evento (`.form.txt:69, 112`).

| Grupo | Controles | Inicio en `.form.txt` |
|---|---|---|
| Cajas de texto de identidad | `ID`, `Nombre`, `Apellido_1`, `Apellido_2`, `TramitacionExterna`, `Correo_e`, `EmpresaTramit` | `.form.txt:328, 364, 399, 434, 470, 506, 542` |
| Cajas de texto de fechas y observaciones | `FAltaTSOL`, `FAltaTdE`, `FbajaTSOL`, `FBajaTdE`, `Observaciones` | `.form.txt:578, 613, 648, 683, 721` |
| Cajas de relación con HPS (sin etiqueta) | `IDHPS`, `IDHPSHistorico` | `.form.txt:748, 767` |
| Etiquetas emparejadas | `Etiqueta633` … `Etiqueta644` para los doce campos con `Caption` propio | `.form.txt:340, 375, 410, 445, 482, 518, 554, 589, 624, 659, 694, 733` |

Cada caja con etiqueta tiene `ControlSource = <nombre del campo>` y la etiqueta, `Caption` igual al nombre (p. ej. `.form.txt:329, 345`). `IDHPS` y `IDHPSHistorico` quedan sin etiqueta visible.

## Flujos

### Apertura (`Form_Load`, `:29-34`)

1. Sustituye `Me.RecordSource = "SELECT * " & "FROM TbUsuariosSICALocal;"` (`:30-31`) — coincide con el del `.form.txt` salvo el espacio final.
2. `Me.Requery` (`:32`).
3. `Set m_Indicadores = modIndicadores.m_Indicadores` (`:33`): se suscribe al singleton del bus (`modIndicadores.bas:8`); la declaración está en el `.cls` (`:4`).

### Cambio de registro (`Form_Current`, `:8-27`)

1. `m_ID = Nz(Me.IDHPS, "")` (`:13`).
2. `EstablecerSemaforo m_ID` (`:14`): misma rutina que en el subformulario 03 (`FG.bas:2865-2904`); usa `IDHPS` para resolver la imagen en `ImagenSemaforo` del panel lateral (`Form_FormInicial01Lateral.form.txt:1805`).
3. `Set m_ObjUsuarioSICAActivo = Constructor.getUsuarioSICA(Nz(Me.ID, ""), , , m_Error)` (`:15`): con `p_IDUsuario = Me.ID` la SQL filtra por `IDHPS=<Me.ID>` (`Constructor.bas:1273-1284`); resultado en la global `m_ObjUsuarioSICAActivo` (`VariablesEntorno.bas:226`).
4. Si `m_Error <> ""`, eleva 1000 (`:16-18`); si no es 1000, `CorreoAlAdministrador m_Error` y `MsgBox` crítico (`:21, 23-26`).

### Notificación de movimiento del bus (`m_Indicadores_LifecycleMoveCompleted`, `:36-40`)

1. `On Error Resume Next`; `Me.Requery` (`:37-38`).
2. Limpia el error con `On Error GoTo 0` (`:39`).

El bus se dispara desde `modIndicadores.Indicadores_NotificarMovimiento` (`modIndicadores.bas:11-15`), consumido por otros formularios; este subformulario se limita a refrescar la rejilla.

## Reglas, validaciones y cálculos

- Solo se carga el usuario SICA cuyo `IDHPS` de `TbUsuariosSICA` coincide con `Me.ID` (`Constructor.bas:1273-1284`); no se contempla la consulta por `IDUsuarioSICA` ni por `IDUsuarioHistorico` desde aquí (`:15` usa el segundo parámetro).
- Sin `WHERE` en `RecordSource`: se sirven todas las filas de `TbUsuariosSICALocal` (`:30-31`, `.form.txt:68`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Origen del subformulario | `TbUsuariosSICALocal` | `SELECT *` sin `WHERE` | `.form.txt:68`, `:30-31` |
| Semáforo del usuario | `TbDatosLocal` (consulta de `URLImagenSemaforo`) | `SELECT` por `ID` | `FG.bas:2891-2897` |
| Entidad activa | `TbUsuariosSICA` (`getUsuarioSICA`) | `SELECT * FROM TbUsuariosSICA WHERE IDHPS=<Me.ID>;` | `:15`, `Constructor.bas:1273-1284` |
| Suscripción al bus de indicadores | `modIndicadores.m_Indicadores` (`clsIndicadoresBus`) | singleton `WithEvents` | `:33`, `modIndicadores.bas:8` |

`SubFormResultados` se exporta a Excel desde el padre con `ConsultaAExcel(Me.SubFormResultados.Form, ...)` (`Form_FormInicial07UsuariosSICA.form.txt:1617`).

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Ya se ve: SQL concatenado en `getUsuarioSICA` (`Constructor.bas:1274, 1279, 1284`); `On Error Resume Next` en el handler del bus (`:37`) que silencia errores de `Requery`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `m_Indicadores` (declarado `WithEvents` en `:4`) queda suscrito al singleton del bus; `Me.Requery` (`:38`) ocurre tras cada notificación sin distinguir si el cambio afecta a la fila actual.
- `On Error Resume Next` en el handler del bus (`:37-40`) enmascara cualquier fallo de `Requery`; el error 1000 de `Form_Current` (`:17`) tampoco se loguea cuando es 1000.
- `IDHPS` e `IDHPSHistorico` (`.form.txt:757, 776`) son las únicas columnas sin etiqueta visible — el usuario SICA puede tener `IDHPS` vacío mientras sí lo tenga `IDHPSHistorico`, y la hoja no lo distingue visualmente.
- `DoCmd.Hourglass False` se invoca dos veces con la misma intención: una en el handler del bus (no, en realidad solo en `errores:` `:25`) y otra en el flujo principal que no existe aquí. Lo problemático es que `Form_Current` no la llama cuando termina con éxito, dejando el reloj de arena activo si Access lo encendió antes.
- La consulta `getUsuarioSICA` se llama con el segundo parámetro `p_IDUsuario` (`:15`); los consumidores que esperan pasar por `p_IDUsuarioSICA` o `p_IDUsuarioHistorico` (`Constructor.bas:1258-1260`) deben usar `Me.ID` de la fila actual, no de la clave de HPS.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: el bus singleton se sustituye por un bus de eventos tipado en backend (web socket / SSE); las dos columnas `IDHPS*` se ocultan o se distinguen con rótulo claro.

## Criterios de aceptación

Propuestos, no aprobados:

1. El subformulario muestra una fila por cada `TbUsuariosSICALocal` y cada `ControlSource` apunta al campo homónimo.
2. Al cambiar el registro actual, `ImagenSemaforo` refleja `URLImagenSemaforo` del usuario cuyo `IDHPS` coincide con `Me.IDHPS`.
3. La entidad global `m_ObjUsuarioSICAActivo` queda instanciada con un `UsuarioSICA` cuyo `ID` es `Me.ID` cuando la consulta devuelve fila; en caso contrario, queda en `Nothing`.
4. Una notificación `LifecycleMoveCompleted` del bus refresca la rejilla aunque el cambio afecte a otra fila.
5. Los errores no 1000 se notifican al administrador y muestran `MsgBox` crítico.

## Pendientes

- Confirmar en el binario que el `.form.txt` no declara columnas adicionales que el `.cls` no conoce (los `ControlSource` identificados llegan hasta 14; cualquier etiqueta «huérfana» debe verificarse).
- Decidir si la suscripción al bus se mantiene en la migración o se reemplaza por un canal reactivo explícito.
