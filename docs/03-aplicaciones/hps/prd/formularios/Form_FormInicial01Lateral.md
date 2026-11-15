# Form_FormInicial01Lateral

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial01Lateral.cls` (494 líneas) y `.form.txt` (2447 líneas) · **PRD:** [README](../README.md)

Lateral izquierdo de la pantalla principal (`Form_FormInicial00Principal`). Aloja el icono de usuario, tres etiquetas informativas de versión, ubicación y modo de escritura, la imagen del semáforo que escribe `EstablecerSemaforo` (`FG.bas:2865-2904`), un botón «Pasar a Lectura / Escritura» y catorce etiquetas-enlace de navegación agrupadas en seis bloques temáticos (Usuarios HPS, Indicadores con contadores, Configuración, Usuarios SICA, Usuarios Históricos, Irregulares / Caducados / Pendientes). No tiene `RecordSource` propio: es un **conmutador** que cambia `Me.Parent.SubFormCentral.SourceObject` del padre. Las citas son `Form_FormInicial01Lateral.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial00Principal` | Se carga como `Me.SubFormLateral.SourceObject = "FormInicial01Lateral"` | `Form_FormInicial00Principal.cls:120` |
| `FormInicial00Principal.SubFormCentral` | El lateral reescribe `SourceObject` a uno de los seis formularios central (`FormInicial02ConsultasPrincipal`, `FormInicial05Configuracion`, `FormInicial07UsuariosSICA`, `FormInicial08UsuariosHistoricos`, `FormInicial09Indicadores`); los seis puntos de asignación están en `:331, 378, 426, 436, 446, 455`. | `:331, 378, 426, 436, 446, 455`, `Form_FormInicial09Indicadores.cls:61-103` |
| `modIndicadores.m_Indicadores` (singleton) | `m_Indicadores_LifecycleMoveCompleted` repinta los contadores cuando el bus notifica un movimiento | `:4, :42, :83-88`, `modIndicadores.bas:8` |

No abre formularios por sí solo, salvo `ComandoCambiarEstado_Click` que reabre `FormInicial` y se cierra (`:7-21`).

## Controles y eventos

`OnLoad ="[Event Procedure]"` (`.form.txt:117`); sin `RecordSource` en el `.form.txt`.

| Grupo | Controles | Inicio en `.form.txt` |
|---|---|---|
| Icono y semáforo | `Imagen7` (`Policia_256.png`), `ImagenSemaforo` (URL `\\datoste\…\Semaforo_Rojo.png`) | `.form.txt:286, 1805` |
| Etiquetas informativas | `lblVersion`, `lblDatosEn`, `lblModoEscritura` | `.form.txt:1491, 1630, 1656` |
| Navegación de alto nivel | `lblUsuariosHPS` (seleccionada), `lblIndicadores`, `lblConfiguracion`, `lblUsuariosSICA`, `lblUsuariosHistoricos` | `.form.txt:264, 1857, 1573, 1754, 1782` |
| Contadores irregulares / caducidad | `lblHPSEstadoIrregular`, `lblHPSAPuntoDeCaducar`, `lblHPSCaducados` | `.form.txt:1545, 1517, 1601` |
| Contadores requiere curso | `lblRequiereCurso` y los tres `lblRequiereCursoPte*` | `.form.txt:1726, 1885, 1914, 1943` |
| Solicitados | `lblSolicitados` | `.form.txt:1829` |
| Botón de modo | `ComandoCambiarEstado` | `.form.txt:1679` |

Todas las etiquetas navegables llevan `Tag = "Elegibles;NoSeleccionado"` salvo `lblUsuariosHPS` que arranca como `Elegibles;Seleccionado` (`.form.txt:267, 2680, 1576, 1757, 1785, 1548, 1520, 1729, 1832, 1860, 1888, 1917, 1946`); `HyperlinkAddress ="#"` en todas las navegables (mismo rango). `Detalle` define el fondo `BackColor =3969911` (`.form.txt:248`).

## Flujos

### Apertura (`Form_Load`, `:27-74`)

1. Entra con `On Error GoTo errores` y vacía `m_Error` (`:29-30`).
2. Pinta `lblVersion`, `lblDatosEn`, `lblModoEscritura` con `m_ObjEntorno.{VersionAplicacion, DatosEn, ModoEscritura}` (`:34-36`).
3. `EstablecerContadores m_Error` recalcula las cinco etiquetas de contadores desde `m_ObjEntorno.Col*` (`:38`; `FG.bas:1440-1525`); eleva 1000 si falla (`:39-41`).
4. `Set m_Indicadores = modIndicadores.m_Indicadores` suscribe el bus singleton (`:42`; `modIndicadores.bas:8`).
5. `CopiarEstiloBoton Me.ComandoCambiarEstado, m_Error` aplica la paleta del entorno al botón de modo (`:43`; `FG.bas:2435-2502`); eleva 1000 si falla (`:44-46`).
6. Lee `Application.TempVars("SoloLectura")`. Si vale `"Sí"`, `Caption = "Pasar a Escritura"` y `Enabled = False` (`:47-48, :53-54`); si no, `Caption = "Pasar a Lectura"` y `Enabled = True` (`:50-51, :55-56`).
7. `SeleccionarEnlace , m_Error` resetea todas las etiquetas `Elegibles;*` a `NoSeleccionado` (`:60`; ver §Helper); eleva 1000 si falla (`:61-63`).
8. `lblUsuariosHPS_Click` se invoca explícitamente para fijar la pestaña de inicio (`:65`; ver §Clic en «Usuarios HPS»).
9. En error: compone mensaje, `CorreoAlAdministrador` si no es 1000, `MsgBox` crítico y **cierra el formulario sin guardar** (`:67-73`).

### Activación (`Form_Activate`, `:76-81`)

`On Error Resume Next`; `EstablecerContadores m_Error`; `On Error GoTo 0`. Repinta contadores cada vez que el lateral recupera el foco (sin propagar errores).

### Notificación del bus (`m_Indicadores_LifecycleMoveCompleted`, `:83-88`)

Mismo esquema silencioso: `On Error Resume Next`, `EstablecerContadores m_Error`, `On Error GoTo 0`. Se dispara desde `modIndicadores.Indicadores_NotificarMovimiento` (`modIndicadores.bas:11-15`).

### Clic en «Usuarios HPS» (`lblUsuariosHPS_Click` → `SeleccionarEnlace`, `:272-287, :332-358`)

1. Vacía `m_Error` y llama a `SeleccionarEnlace Me.lblUsuariosHPS, m_Error` (`:274-276`).
2. Si `m_Error <> ""`, eleva 1000 (`:277-279`).
3. Dentro de `SeleccionarEnlace`, cuando `p_lbl.Name = "lblUsuariosHPS"` (`:329`), dos ramas:
   - **Otro central** (`:330-340`): si `m_Frm.SubFormCentral.SourceObject <> "FormInicial02ConsultasPrincipal"`, asigna `SourceObject = "FormInicial02ConsultasPrincipal"`, reaplica `EstablecerControlesformlateral`, fuerza `ListaConsultas.Selected(0) = True` y llama a `Form_FormInicial02ConsultasPrincipal.ListaConsultas_Click` (`:339`) y `Form_FormInicial03ConsultasEstado.ComandoLimpiarEstadoHPS_Click` (`:340`).
   - **SubFormCentral ya es `FormInicial02ConsultasPrincipal`** (`:341-354`): reentra solo si `lblHPSEstadoIrregular`, `lblHPSAPuntoDeCaducar` o `lblHPSCaducados` están pintados como «seleccionado» (`m_ObjEntorno.EtiquetaLinkBackColorSeleccionado`, `:342-344`); en ese caso reaplica `EstablecerControlesformlateral`, reinicia la lista de consultas y limpia el estado HPS activo (`:346-353`).
4. En error del helper, se rellena `p_Error` silenciosamente (`:473-477`); el handler propaga con `Err.Raise 1000` (`:277-279`).

### Clic en un contador «Irregular / Caducado / Pendiente»

Seis handlers (`lblHPSAPuntoDeCaducar_Click`, `:110-124`; `lblHPSCaducados_Click`, `:126-140`; `lblRequiereCurso_Click`, `:174-188`; `lblRequiereCursoPtePrimeraConvocatoria_Click`, `:208-222`; `lblRequiereCursoPteSegundaConvocatoria_Click`, `:224-238`; `lblRequiereCursoPteCorreoJefeSeguridad_Click`, `:192-206`): todos llaman a `SeleccionarEnlace Me.lblXXX, m_Error`, elevan 1000 si falla, y propagan con `MsgBox` en el `errores:` (mismo patrón, bloques de errores en `:118-123, 134-139, 150-155, 182-187, 200-205, 216-221, 232-237`).

Dentro del helper (`:360-423`) reenvían a `lblIndicadores_Click` (no a `SeleccionarEnlace` con la etiqueta de indicadores), reasignan `m_Frm = m_Frm.SubFormCentral.Form` y disparan `Form_FormInicial09Indicadores.Comando*_Click` si el botón está `Enabled`: lblHPSAPuntoDeCaducar → `:360-366`, lblHPSCaducados → `:367-373`, lblRequiereCurso → `:396-402`, lblRequiereCursoPtePrimeraConvocatoria → `:403-409`, lblRequiereCursoPteSegundaConvocatoria → `:410-416`, lblRequiereCursoPteCorreoJefeSeguridad → `:417-423`. `lblHPSEstadoIrregular` (`:374-402`) y `lblSolicitados` (`:462-471`) son las dos excepciones que se describen aparte.

### Clic en «Irregular» (`lblHPSEstadoIrregular_Click`, `:142-156` → `SeleccionarEnlace`, `:374-402`)

Único caso que monta un filtro en lugar de un comando de indicadores. Si `SubFormCentral` no es ya `FormInicial02ConsultasPrincipal`, lo asigna (`:377-380`); aplica estilo (`:381-385`); fuerza `ListaConsultas.Selected(0) = True` y dispara `ListaConsultas_Click` (`:388-390`); abre el `SubFormParametrosConsulta` del central y pone `ComboEstadoHPS.Value = "Irregular"`, tras lo cual llama a `Form_FormInicial03ConsultasEstado.ComboEstadoHPS_AfterUpdate` (`:391-393`).

### Clic en «Indicadores / Configuración / SICA / Históricos»

Handlers: `lblConfiguracion_Click` `:92-106`, `lblIndicadores_Click` `:158-172`, `lblUsuariosSICA_Click` `:481-494`, `lblUsuariosHistoricos_Click` `:256-270`. Cada uno invoca `SeleccionarEnlace Me.lblXXX, m_Error` con la rama `ElseIf p_lbl.Name = "lblXXX"` que asigna `m_Frm.SubFormCentral.SourceObject` al central correspondiente y reaplica `EstablecerControlesformlateral` si ha habido recambio: `lblConfiguracion` → `:424-432`, `lblIndicadores` → `:434-443`, `lblUsuariosSICA` → `:444-452`, `lblUsuariosHistoricos` → `:453-461`.

### Clic en «Pasar a Lectura / Escritura» (`ComandoCambiarEstado_Click`, `:7-21`)

`DoCmd.OpenForm "FormInicial"` y `DoCmd.Close acForm, Me.Name, acSaveNo` con un único `MsgBox` crítico en caso de error (`:13-19`). Es el atajo documentado en el formulario para reabrir la aplicación como `FormInicial` (sin sufijo).

### `SeleccionarEnlace` (función privada, `:290-477`)

Cuatro secciones:

1. **`p_lbl Is Nothing`** (`:302-311`): recorre `Me.Controls` (`:303`), marca como `NoSeleccionado` (`:306`) todos los `Label` con `;` en el `Tag`. Es la ruta de «reset».
2. **`p_lbl` no nulo, primera parte** (`:312-324`): guarda `m_nombreControlSeleccionado` (`:312`), pone esa etiqueta a `Seleccionado` (`:313`) y reescribe el resto a `NoSeleccionado` con `Nz(ctl.Tag, ";")` (`:318-319`).
3. **`Select Case` por nombre de etiqueta** (`:329-468`): cada `If/ElseIf` corresponde a un conmutador de pestaña (líneas de ramas resumidas arriba). `lblSolicitados` (`:462-467`) reenvía a `lblIndicadores_Click` (`:463`) y dispara `ComandoEnSolicitud_Click` (`:466`); `lblIndicadores`, `lblConfiguracion`, `lblUsuariosSICA`, `lblUsuariosHistoricos` solo cambian `SourceObject`; `lblHPS*` y `lblRequiereCurso*` reenvían a `lblIndicadores_Click` y disparan `Comando*_Click` del formulario 09.
4. **Errores** (`:473-477`): asigna `p_Error` con descripción si no es 1000; **no** envía correo ni muestra `MsgBox` (`:474-476`). La notificación corresponde a los llamadores.

### Errores

Los handlers siguen el patrón del repositorio: para errores no 1000, `m_Error = "Al … se ha producido el error: " & Err.Description`, `CorreoAlAdministrador m_Error` y `MsgBox` crítico guardando el retorno en `pregunta` (bloques de error en `:14-19, 67-72, 100-105, 118-123, 134-139, 150-155, 166-171, 182-187, 200-205, 216-221, 232-237, 248-253, 264-269, 281-286, 489-494`; `FG.bas:67`). Solo `Form_Load` cierra el formulario tras el error; los `Click` no lo cierran (`:67-72`).

## Reglas, validaciones y cálculos

- **Idempotencia por `SourceObject`.** Cada rama que abre un central compara el `SourceObject` actual antes de reasignar (`:330, 377, 425, 435, 445, 454`); si ya coincide, reaplica estilo si hace falta. Las ramas que disparan comandos de indicadores (`lblHPS*`, `lblRequiereCurso*`, `lblSolicitados`) no tienen ese guard — comprueban `Enabled` del botón del central después de reasignar (`:362, 369, 384, 391, 398, 405, 469`).
- **Reset por defecto.** `Form_Load` llama a `SeleccionarEnlace` sin argumentos para devolver todas las etiquetas a `NoSeleccionado` (`:60`); después `lblUsuariosHPS_Click` selecciona la pestaña de inicio (`:65`). `lblUsuariosHPS` se reinicia a `Seleccionado` por una segunda ruta dentro del helper cuando el central ya coincide (`:341-354`).
- **`Tag` como canal de estado.** La función muta `ctl.Tag` directamente; cualquier handler que lea `Tag` debe asumir `Elegibles;<estado>` separado por `;` (`:306, :313, :319`).
- **Sin captura de errores propia.** `SeleccionarEnlace` no eleva la excepción: rellena `p_Error` y devuelve vacío (`:473-477`). Los handlers comprueban `m_Error` después de la llamada y propagan con `Err.Raise 1000` (`:39-41, 44-46, 61-63, 96-98, 114-116, 130-132, 146-148, 162-164, 178-180, 196-198, 212-214, 228-230, 244-246, 260-262, 277-279, 485-487`).
- **Estilo leído del entorno.** Colores de `Detalle.BackColor` y de cada `Label` se leen de `m_ObjEntorno.{CabeceraBackColor, EtiquetaBorderColor, EtiquetaLinkForeColor, EtiquetaLinkBackColorSeleccionado, EtiquetaForeColorSeleccionado}` en `EstablecerControlesformlateral` (`FG.bas:967-999`); los `BackColor` por defecto del `.form.txt` se sobrescriben en cada `Click` que abre un central.
- **Contadores se recalculan en cada foco.** `Form_Activate` y `m_Indicadores_LifecycleMoveCompleted` vuelven a llamar a `EstablecerContadores` con `On Error Resume Next` (`:77-80, :84-87`); `Form_Load` lo llama con propagación de error (`:38-41`).
- **`ComandoCambiarEstado` refleja `Application.TempVars("SoloLectura")`** (`:47-57`); si vale `"Sí"`, el botón se deshabilita, lo que es coherente con un sistema de solo-lectura global.
- **Las etiquetas navegables saltan directo a `Form_FormInicialXX` con nombre literal** (`:331, 378, 426, 436, 446, 455`); cada cambio de `SourceObject` arrastra un literal duplicado en su `SubFormParametrosConsulta` (análogo al par Usuario/UsuarioHistorico).

## Lecturas y escrituras

| Operación | Origen | Tipo | Evidencia |
|---|---|---|---|
| Versión, ubicación y modo | `m_ObjEntorno.{VersionAplicacion, DatosEn, ModoEscritura}` | propiedades del `Entorno` | `:34-36` |
| Recalculo de contadores | `m_ObjEntorno.Col{UsuariosEstadoIrregular, UsuariosAPuntoCaducar, UsuariosCaducados, UsuariosHPSRequierenCurso, RequiereCursoFalta*, UsuariosEnSolicitud}` | propiedades `Collection` | `FG.bas:1457-1509` |
| Semáforo del usuario | `ImagenSemaforo.Picture`, `m_ObjDatosLocal.URLImagenSemaforo` | propiedad de la imagen, helper del módulo | `:83, :84-87`, `FG.bas:2865-2904` |
| Texto del botón de modo | `Application.TempVars("SoloLectura")` | variable temporal global | `:47-57` |
| Bus de indicadores | `modIndicadores.m_Indicadores` (`clsIndicadoresBus`) | singleton `WithEvents` | `:4, :42, :83-88`, `modIndicadores.bas:8` |
| Cambio de pestaña central | `Me.Parent.SubFormCentral.SourceObject` | propiedad del subform contenedor | `:331, 378, 426, 436, 446, 455` |
| Estilo del lateral | `EstablecerControlesformlateral` | helper de `FG.bas` que pinta la paleta | `:333, 346, 381, 428, 438, 448, 457`; `FG.bas:963-1004` |

Sin acceso directo a tablas; las colecciones en memoria son la única fuente de los contadores.

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Lo que ya se ve: el lateral guarda estado en `Tag` (`Elegibles;<estado>`) en lugar de una propiedad booleana de Access, lo que complica el mapeo a CSS/`aria-pressed`; los seis conmutadores de pestaña escriben literales en duro sobre `SourceObject` (`FormInicial02ConsultasPrincipal`, `FormInicial05Configuracion`, `FormInicial07UsuariosSICA`, `FormInicial08UsuariosHistoricos`, `FormInicial09Indicadores`), replicados a su vez en `Form_FormInicial06BusquedaGlobal.FiltrarUsuarioSICA` (`:326-327`) y `Form_FormInicial06BusquedaGlobal.FiltrarUsuarioHistorico` (`:244-247`); el singleton `modIndicadores.m_Indicadores` cruza formularios por una variable global pública, lo que en web se sustituye por un canal reactivo (SSE / bus tipado); el botón `ComandoCambiarEstado` reabre `FormInicial` y se cierra (`:7-21`), un patrón sin equivalente directo que requerirá un componente de logout o de conmutación de entorno.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Form_Load` llama a `EstablecerContadores` (`:38`) **antes** de `SeleccionarEnlace , m_Error` (`:60`), por lo que los contadores se pintan con los `Tag` aún por resetear. `EstablecerControlesformlateral` (helper del módulo) reescribe el estilo tras cada `Click`, así que el orden solo afecta a la captura inicial (a verificar).
- `m_Indicadores_LifecycleMoveCompleted` (`:83-88`) está suscrito con `On Error Resume Next` y `On Error GoTo 0` sin propagar el error: cualquier fallo de `EstablecerContadores` queda silenciado, igual que en `Form_Activate` (`:76-81`).
- La rama de reentrada de `lblUsuariosHPS` (`:341-354`) solo se dispara cuando uno de los tres `lblHPS*` está pintado como seleccionado (`BackColor = m_ObjEntorno.EtiquetaLinkBackColorSeleccionado`); si en una sesión determinada están los tres deseleccionados, hacer clic en `lblUsuariosHPS` con el central ya en `FormInicial02ConsultasPrincipal` no hace nada visible.
- Las seis ramas que disparan comandos de indicadores (`lblHPSAPuntoDeCaducar`, `lblHPSCaducados`, los tres `lblRequiereCurso*`, `lblSolicitados`) reenvían primero a `lblIndicadores_Click` (`:362, 369, 398, 405, 412, 419, 463`) y luego disparan el `Comando*_Click` del formulario 09; las dos llamadas hacen que `EstablecerControlesformlateral` se aplique dos veces seguidas. La idempotencia por `SourceObject` atenúa el efecto visual pero el estilo del lateral se recalcula (a verificar contra el binario).
- `SeleccionarEnlace` no notifica al administrador: se limita a poblar `p_Error` (`:473-477`). Si un `Click` comprueba `m_Error` tras llamar al helper y eleva 1000, el `If Err.Number <> 1000` filtra el reenvío a correo en el handler (`FG.bas:1001`).
- El `Tag` de `lblRequiereCursoPte*` codifica contadores de subconsultas que solo el `EstablisharContadores` actualiza (`FG.bas:1481-1505`); el `.form.txt` los fija a literales numéricos (`.form.txt:1886, 1915, 1944`) que quedan como código muerto mientras el lateral se carga.
- Las `Caption` por defecto de los contadores en el `.form.txt` están stale (`.form.txt:1518, 1546, 1602, 1727, 1830`); son semillas de diseño que `EstablecerContadores` sobrescribe en cada activación.
- La imagen `ImagenSemaforo` apunta a una ruta UNC `\\datoste\aplicaciones_dys\…\Semaforo_Rojo.png` (`.form.txt:1806`) que no es portable.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: reemplazar `Tag` por `aria-pressed`; sustituir el conmutador por un router declarativo (`/inicial/usuarios-hps`, `/inicial/indicadores`, …) con su `SourceObject` resuelto por una tabla; mover la paleta a CSS variables leídas del entorno; sustituir el singleton `modIndicadores` por un canal reactivo tipado; un endpoint para archivos de semáforo que evite la dependencia con la ruta UNC.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al entrar, `lblUsuariosHPS` queda `Seleccionado` y el `SubFormCentral` se monta en `FormInicial02ConsultasPrincipal` con la lista de consultas en su primera fila.
2. Pulsar `lblIndicadores` cambia el `SubFormCentral` a `FormInicial09Indicadores`; el comando `ComandoEnSolicitud_Click` queda disponible si hay usuarios en solicitud.
3. Pulsar `lblHPSEstadoIrregular` con central distinto a `FormInicial02ConsultasPrincipal` monta el central, abre `SubFormParametrosConsulta` y fija `ComboEstadoHPS.Value = "Irregular"`.
4. Pulsar `lblRequiereCursoPtePrimeraConvocatoria` con el central en `FormInicial09Indicadores` dispara `ComandoPendientesConvocatoria1_Click` solo si su `Enabled` es `True`.
5. `EstablecerContadores` se ejecuta al menos una vez por `Form_Load`, una por cada `Form_Activate` y una por cada evento `LifecycleMoveCompleted` del bus.
6. Un error no interno se traduce en un `MsgBox` crítico; `Form_Load` cierra el formulario, los `Click` no.
7. `ComandoCambiarEstado` reabre `FormInicial` y cierra el lateral sin propagar el estado.

## Pendientes

- Procedencia exacta de cada `m_ObjEntorno.Col*` (vaciados en `FG.bas:1457-1505`); su lectura desde dónde se rellena, no desde el lateral.
- Si `modIndicadores.m_Indicadores` se destruye al cerrar el lateral; hoy `Set m_Indicadores = Nothing` no se ejecuta (`:42` sin balance) (a verificar en runtime).
- `EstablecerSemaforo` también pinta sobre `ImagenSemaforo` el icono del usuario activo desde `Form_FormInicial03ConsultasDatos.Form_Current` (`:14`); convendría confirmar que el lateral lo asume como «resultado de búsqueda» y no como «usuario activo».
