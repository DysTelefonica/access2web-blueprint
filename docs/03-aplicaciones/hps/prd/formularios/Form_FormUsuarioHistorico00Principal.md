# Form_FormUsuarioHistorico00Principal

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormUsuarioHistorico00Principal.cls` (98 líneas) y `.form.txt` (1116 líneas) · **PRD:** [README](../README.md)

Contenedor principal de la ficha de un usuario **histórico** (un HPS dado de baja que se conserva como histórico). Es un formulario **no enlazado** que compone la pantalla con dos subformularios: el lateral de navegación (`FormUsuarioHistorico01Lateral`) y el central, cuyo `SourceObject` cambia según la pestaña elegida. Su única acción de negocio es **pasar el usuario histórico a Actual**, que mueve ficha y anexos. Las citas son `Form_FormUsuarioHistorico00Principal.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Navegación de HPS | Se abre tras elegir un usuario en el histórico; el objeto activo es un global, no un argumento | `m_ObjUsuarioHistoricoActivo` es `Public … As UsuarioHistorico` (`VariablesEntorno.bas:227`) |
| `Form_FormUsuarioHistorico01Lateral` | Se instala como `SubFormLateral` si no lo está ya | `:78-80` |
| `Form_FormUsuarioHistorico01Lateral` | Cambia `Parent.Form.SubFormCentral.SourceObject` según la pestaña | `Form_FormUsuarioHistorico01Lateral.cls:47, 63` |
| `Form_FormUsuarioHistorico02Edicion`, `Form_FormUsuarioHistorico02Anexos` | Destino del subformulario central | `.form.txt:890` |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Ajusta tamaño, compone el título, monta los subformularios y deja el formulario en solo lectura | `:52-97` |
| (formulario) | — | `OnUnload` implícito | No hay `Form_Unload`; el estado global sobrevive al cierre | — |
| `SubFormLateral` | subformulario | — | Aloja la navegación; se fuerza a `FormUsuarioHistorico01Lateral` | `:78-80`, `.form.txt:850` |
| `SubFormCentral` | subformulario | — | Lo sustituye el lateral; su fuente es pública para los hijos | `.form.txt:890` |
| `ComandoPasarAActual` | botón | `Click` → `ComandoPasarAActual_Click` | Confirma y mueve la ficha y los anexos a Actual | `:15-49`, `.form.txt:986` |
| `cmdSalir` | botón | `Click` → `CmdSalir_Click` | Cierra el formulario sin guardar | `:8-12`, `.form.txt:942` |
| `lblTitulo` | etiqueta | — | Recibe `HISTÓRICO DETALLE DE <NOMBRE>` | `:70-73`, `.form.txt:813` |
| `Etiqueta1`, `Etiqueta3`, `EncabezadoDelFormulario`, `PieDelFormulario`, `Detalle` | etiquetas, marco y secciones | — | Estructura visual; sin lógica | `.form.txt:869, 909, 796, 928, 835` |

## Flujos

### Apertura (`Form_Load`, `:52-97`)

1. `AjustarTamaño Me` dimensiona la ventana (`:61`; `FG.bas:901`).
2. Pone `AllowEdits = False` para bloquear la edición mientras monta (`:63`).
3. Si no hay usuario histórico activo, aborta con «No se puede determinar el usuario activo» (`:64-68`).
4. Compone el título con `UCase(NombreCompleto)` y lo propaga a la etiqueta y a la barra de la ventana (`:70-73`).
5. `EstablecerControlesDatos Me, m_Error` aplica la paleta de colores (`:75-78`; `FG.bas:1005`).
6. Fuerza `SubFormLateral.SourceObject = "FormUsuarioHistorico01Lateral"` si no coincide (`:78-80`).
7. Decide si `ComandoPasarAActual` queda habilitado según `Application.TempVars("SoloLectura")` (`:81-85`).
8. `CopiarEstiloBoton` sobre los dos botones (`:86-87`; `FG.bas:2435`).
9. Restaura `AllowEdits = True` (`:88`), también en el manejador de error (`:95`).
10. En error: correo al administrador, `MsgBox` crítico y cierre sin guardar (`:89-97`).

### Pasar a Actual (`ComandoPasarAActual_Click`, `:15-49`)

1. Pide confirmación con botón por defecto **No** (`:20`). Si la respuesta no es «Sí», sale (`:21-23`).
2. Crea un `UsuarioServicio` (`:25`).
3. Activa el reloj de arena «puede tardar con muchos anexos / observación» (`:26-27`).
4. `PasarDeHistoricoAActual m_ObjUsuarioHistoricoActivo.ID, m_Error` (`:28`). Ese método delega en un coordinador transaccional que mueve datos y anexos de raíz (`:44`; `UsuarioServicio.cls:53-70`).
5. Si hay error, apaga el reloj y relanza 1000 (`:29-32`).
6. Apaga el reloj (`:33`) y, si no hubo error, notifica el movimiento a los indicadores con la dirección `"h2c"` (`:35-37`; `modIndicadores.bas:11-15`).
7. Muestra «El usuario ha sido pasado a Actual correctamente.» y cierra (`:39-41`).
8. En error: apaga el reloj, compone el mensaje, correo al administrador y `MsgBox` (`:43-49`).

### Salir (`CmdSalir_Click`, `:8-12`)

`On Error Resume Next` y `DoCmd.Close` sin guardar (`:9, 11`): si el cierre falla, el error se descarta.

### Errores

`Form_Load` y `ComandoPasarAActual_Click` envían `CorreoAlAdministrador` salvo en errores internos 1000 (`:44-48`, `:89-94`). El retorno de cada `MsgBox` va al global `pregunta` (`VariablesEntorno.bas:220`).

## Reglas, validaciones y cálculos

- **Confirmación con default negativo.** `vbDefaultButton2` sobre un `vbYesNo` sitúa el foco en «No» (`:20`): la operación destructiva no se dispara con Intro por descuido.
- **La irreversibilidad está en el servicio, no aquí.** El formulario solo confirma y delega; toda la lógica de mover ficha y anexos vive en `UsuarioLifecycleTransactionCoordinator.PasarHistoricoAActual`, que recibe la base de pruebas, las raíces de adjuntos de Actual e Histórico y el escenario de fallo (`:28, 44`; `UsuarioServicio.cls:58`).
- **Notificación best-effort.** `Indicadores_NotificarMovimiento` envuelve su cuerpo en `On Error Resume Next` y termina con `On Error GoTo 0` (`:36`; `modIndicadores.bas:12-14`): un fallo de indicadores no puede tumbar un movimiento ya realizado.
- **Solo lectura materializada.** `Me.AllowEdits = False` bloquea la edición del propio contenedor mientras se monta, y el subformulario central es de otro formulario (`:63, 88, 95`).
- **Identidad por objeto global.** El usuario sobre el que se opera es `m_ObjUsuarioHistoricoActivo`, no un identificador recibido como argumento (`:28, 36, 66, 71`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Pasar usuario histórico a Actual | ficha y anexos, vía `UsuarioLifecycleTransactionCoordinator` | transacción en el servicio, fuera de este formulario | `:28, 44`, `UsuarioServicio.cls:53-70` |
| Notificar movimiento a indicadores | `m_Indicadores.RaiseLifecycleMoveCompleted` | evento en memoria | `:36`, `modIndicadores.bas:11-15` |
| Título de la ventana | `m_ObjUsuarioHistoricoActivo.NombreCompleto` | propiedad del objeto global | `:70-73` |
| Marca «SoloLectura» | `Application.TempVars("SoloLectura")` | variable temporal de la sesión | `:81` |

## Coexistencia

Pendiente de `../coexistencia.md`. La escritura no pasa por este formulario sino por un coordinador transaccional, lo que ayuda: la regla de negocio no está en la capa de presentación. Quedan dos ataduras a Access sin equivalente directo: `Application.TempVars` como memoria compartida entre formularios (`:81`) y el subformulario cuyo `SourceObject` intercambia otro formulario (`:78-80`). El reloj de arena (`:27, 30, 33, 44`) no tiene traducción; su equivalente web es un estado de carga en la vista.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **El botón «Pasar a Actual» está siempre deshabilitado.** Las dos ramas del `If` sobre `Application.TempVars("SoloLectura")` asignan `Enabled = False` (`:81-85`). La rama de solo lectura es correcta; la rama de lectura-escritura debía poner `True`. Tal como está, la capacidad no es alcanzable desde la interfaz (a verificar contra el binario, donde el `.form.txt` podría fijar `Enabled`).
- `CmdSalir_Click` usa `On Error Resume Next` (`:9`), de modo que un cierre fallido no deja rastro.
- Tras mover la ficha no se refresca `m_ObjUsuarioHistoricoActivo` ni se limpia el estado global: el formulario se cierra, pero el objeto sigue apuntando al histórico recién pasado (`:41`).
- `Indicadores_NotificarMovimiento` se invoca con la cadena mágica `"h2c"` (`:36`) sin constante que la documente; el valor debe coincidir con lo que el consumidor de indicadores espera.
- `pregunta` es un `Long` global compartido (`VariablesEntorno.bas:220`) cuyo valor se descarta en `:49` y `:96`.
- `m_Error` es una variable de módulo del formulario, pero los helpers called con `ByRef` la sobrescriben desde fuera (`:28, 44`), lo que complica el rastro de quién produjo el error.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: pasar a Actual como operación idempotente de un endpoint transaccional en lugar de unFormulario de confirmación; estado de solo lectura derivado de permisos reales en vez de `TempVars`; y una acción de reversión o «deshacer» dado que el movimiento mueve anexos.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con la marca de solo lectura activa, el botón «Pasar a Actual» no está disponible y el resto de la ficha es consultable.
2. Con permisos de edición, el botón está disponible y pide confirmación con «No» como respuesta por defecto.
3. Confirmar mueve la ficha y sus anexos a Actual de forma atómica, o no mueve nada y comunica el error.
4. Tras un movimiento correcto, el histórico deja de ofrecerse para volver a pasar y los indicadores reciben la notificación.
5. «Salir» cierra sin guardar y no deja el formulario en estado de edición bloqueado.

## Pendientes

- Si el `.form.txt` del binario fija `ComandoPasarAActual.Enabled` en el diseño, lo que podría explicar que el botón esté deshabilitado en la práctica y no por este código (`.form.txt:986`).
- Contenido de `UsuarioLifecycleTransactionCoordinator.PasarHistoricoAActual`, en `classes/`, que es donde vive la regla real del movimiento.
- Qué menús o macros del binario abren este formulario: no hay `DoCmd.OpenForm` en `src`.
- Si `Application.TempVars("SoloLectura")` la fija otro formulario y con qué condición.