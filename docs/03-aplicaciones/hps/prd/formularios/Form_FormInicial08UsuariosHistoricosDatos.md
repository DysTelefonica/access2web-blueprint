# Form_FormInicial08UsuariosHistoricosDatos

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial08UsuariosHistoricosDatos.cls` (47 líneas) y `.form.txt` (1062 líneas) · **PRD:** [README](../README.md)

Hoja-datos de los usuarios históricos: **no filtra y no escribe**. `Form_FormInicial08UsuariosHistoricos` la monta como `SubFormResultados` (`SourceObject ="Form.FormInicial08UsuariosHistoricosDatos"`, `.form.txt:896-897`) y la reutiliza para recargar y para exportar a Excel (`.form.txt:1462, 1520, 1547-1548, 1587`). Se muestra como datasheet (`DefaultView =2`, `.form.txt:15`; `ViewsAllowed =2`, `.form.txt:16`; `AllowFormView =0`, `.form.txt:129`) con veinte columnas enlazadas a `TbUsuariosHistoricosLocal`. Su trabajo real es twofold: fijar el `RecordSource` local y **publicar en el global `m_ObjUsuarioHistoricoActivo` el usuario sobre el que el padre opera** (`:14`), que es lo que habilita los botones de ver y eliminar del anfitrión (`Form_FormInicial08UsuariosHistoricos.cls:15-19, 40-43`). Se registra además como oyente del bus de indicadores para recargarse cuando termina un movimiento de ciclo de vida (`:4, 34, 43-47`). Las citas son `Form_FormInicial08UsuariosHistoricosDatos.cls:línea` salvo indicación; `FG.bas` abrevía `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial08UsuariosHistoricos` | Se monta como `SubFormResultados` en diseño y se reassigna en tres puntos del `.cls` del padre | `.form.txt:896-897`, `Form_FormInicial08UsuariosHistoricos.form.txt:1520, 1547, 1587` |
| `modIndicadores` | `WithEvents` sobre el singleton `modIndicadores.m_Indicadores`; al terminar un movimiento del ciclo de vida fuerza un `Requery` | `:4, 34, 43-47`, `modules/modIndicadores.bas:8, 11-15`, `classes/clsIndicadoresBus.cls:16-20` |

No abre ningún formulario. El padre lo usa como origen de la exportación a Excel (`Form_FormInicial08UsuariosHistoricos.form.txt:1548`) y lo recorre con `Me.SubFormResultados.Form.Requery` (`.form.txt:1462`).

## Controles y eventos

`RecordSource` en diseño: `SELECT TbUsuariosHistoricosLocal.* FROM TbUsuariosHistoricosLocal; ` (.form.txt:83). `OnCurrent ="[Event Procedure]"` (.form.txt:84), `OnLoad ="[Event Procedure]"` (.form.txt:128) y `OnActivate` **no** declarado (`.form.txt`, sin coincidencias).

| Grupo | Controles | Inicio en `.form.txt` |
|---|---|---|
| Cajas de texto de identidad y contacto | `ID`, `DNI`, `Nombre`, `Apellido_1`, `Apellido_2`, `Telefono`, `Correo_e` | `.form.txt:338, 371, 405, 440, 474, 508, 542` |
| Cajas de texto de fechas y situación HPS | `F_Nacimiento`, `Motivo_HPS`, `F_Curso`, `CursoEnVigor`, `Requiere_Curso`, `F_Baja`, `FAvisoConcesion` | `.form.txt:576, 610, 644, 678, 713, 748, 782` |
| Cajas de texto de comunicación y expediente | `RequiereComunicacionConcesion`, `FechaHPSConcesionMinima`, `CodExp`, `EmpresaUsuario`, `EmpresaTramitadora`, `JuridicaContrato` | `.form.txt:816, 851, 885, 919, 953, 987` |
| Etiquetas emparejadas | `Etiqueta625` … `Etiqueta647` (veinte, con la numeración saltando de `Etiqueta637` a `Etiqueta641`), cada una con `Caption` igual al nombre del campo | `.form.txt:355, 389, 424, 458, 492, 526, 560, 594, 628, 662, 697, 732, 766, 800, 835, 869, 903, 937, 971, 1005` |

Cada caja lleva su `ControlSource` apuntando al campo homónimo (`ID` → `ID`, `.form.txt:344-345`; `EmpresaUsuario` → `EmpresaUsuario`, `.form.txt:926-927`). Los controles de cabecera del `.form.txt` (`.form.txt:189-315`: etiqueta, botón, casilla, cuadro, lista, combo y subformulario sin `Name`) no tienen lógica en este `.cls`. `Detalle` es la sección que agrupa las veinte columnas (`.form.txt:327, 329`).

## Flujos

### Apertura (`Form_Load`, `:29-35`)

1. Fija `Me.RecordSource = "SELECT * FROM TbUsuariosHistoricosLocal;"` (`:30-31`), equivalente en efecto al `RecordSource` de diseño (.form.txt:83).
2. `Me.Requery` (`:33`).
3. `Set m_Indicadores = modIndicadores.m_Indicadores` (`:34`): se suscribe al singleton para recibir `LifecycleMoveCompleted` (`modules/modIndicadores.bas:8`).

No tiene manejador de errores: cualquier fallo en estas tres líneas queda sin tratar.

### Cambio de registro (`Form_Current`, `:7-26`)

1. `m_Error = ""` y `m_ID = Nz(Me.ID, "")` (`:11-12`).
2. `Set m_ObjUsuarioHistoricoActivo = Constructor.getUsuarioHistorico(m_ID, , m_Error)` (`:14`), relanzando 1000 si `m_Error` quedó relleno (`:15-17`).
3. `errores:` anexa la descripción y avisa al administrador solo si `Err.Number <> 1000` (`:20-23`), ejecuta `DoCmd.Hourglass False` (`:24`) y muestra el `MsgBox` crítico con el `m_Error` pendiente, guardado en el global `pregunta` (`VariablesEntorno.bas:220`).

### Reactivación y movimiento de ciclo de vida (`Form_Activate`, `:37-41`; `m_Indicadores_LifecycleMoveCompleted`, `:43-47`)

`Form_Activate` solo hace `Me.Requery` bajo `On Error Resume Next … On Error GoTo 0` (`:38-40`): no filtra ni avisa, silencia cualquier fallo. El segundo es handler de `WithEvents` para `clsIndicadoresBus.LifecycleMoveCompleted(p_Direction As String, p_UserId As String)` (`classes/clsIndicadoresBus.cls:16`, lanzado en `:18-20` desde `modIndicadores.Indicadores_NotificarMovimiento`, `modules/modIndicadores.bas:11-15`): **ignora ambos parámetros** y solo hace `Me.Requery` bajo `On Error Resume Next` (`:44-46`).

## Reglas, validaciones y cálculos

- **La hoja no filtra.** No escribe en `m_ObjColParaFiltros` ni llama a `EstablecerFiltroUsuarios`: frente a los formularios del panel de consultas, este no altera el `RecordSource` de nadie más.
- **La fuente de la vista es la tabla local, la del objeto seleccionado es el backend.** La hoja muestra `TbUsuariosHistoricosLocal` (`:30-31`) mientras que `Constructor.getUsuarioHistorico` consulta `TbUsuariosHistoricos` en `getdb()` (`Constructor.bas:1214-1225`); la tabla local se puebla con un `DELETE FROM` seguido de un `INSERT INTO … SELECT` de 23 columnas desde el backend (`Mod_Sincronizacion_Historico.bas:45-46, 51-54`).
- **Selección por `ID`, con reserva para vacío.** `getUsuarioHistorico` solo construye el objeto si recibe `p_IDUsuario` o `p_DNI`; con ambos vacíos sale sin hacer nada (`Constructor.bas:1211-1213`), y si el recordset está vacío también devuelve `Nothing` (`Constructor.bas:1230-1231`). El `Me.ID` envuelto en `Nz` (`:12`) evita el error de `Null`, pero no evita ese retorno vacío.
- **Contrato del objeto publicado.** `getUsuarioHistorico` rellena todas las propiedades de `ColCamposSimplificado` con `Nz` sobre el recordset (`Constructor.bas:1235-1242`) usando la base que le llegue por `p_Db`, que aquí no se provee (`Constructor.bas:1224-1228`).
- **Guard del padre.** Los botones del anfitrión exigen que `m_ObjUsuarioHistoricoActivo` no sea `Nothing` y muestran «Se ha de seleccionar el usuario primero» si lo está (`Form_FormInicial08UsuariosHistoricos.cls:15-19, 40-43`); ese es el único contrato entre la hoja y el padre.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Origen de la hoja | `TbUsuariosHistoricosLocal` | `SELECT * FROM TbUsuariosHistoricosLocal;` asignado a `Me.RecordSource` | `:30-31` |
| Usuario seleccionado (objeto global) | `TbUsuariosHistoricos` (backend, vía `getdb()`) | `SELECT * … WHERE ID=<valor>;` | `:14`, `Constructor.bas:1215-1217, 1225` |
| Origen de la tabla local | `TbUsuariosHistoricos` unido a `TbSuministradores` y `TbExpedientes` | `DELETE FROM …` + `INSERT INTO … SELECT` de 23 columnas | `Mod_Sincronizacion_Historico.bas:45-46, 51-54` |

Sin escrituras propias: el `.cls` solo lee y reasigna el `RecordSource`. La escritura en la tabla local ocurre en el proceso de sincronización, no en este formulario.

## Coexistencia

Pendiente de `../coexistencia.md`. Se ve ya la separación de bases: la vista se sirve desde la caché local y el objeto que se publica al padre se relee del backend (`Constructor.bas:1225`), de modo que una caché desalineada muestra una fila y construye otra. La dependencia del bus de indicadores (`modules/modIndicadores.bas:8`) acopla la hoja al ciclo de vida completo, y el `DoCmd.Hourglass False` incondicional (`:24`) comparte estado de interfaz con el resto de la aplicación.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Form_Activate` (`:37-41`) existe en el `.cls` pero el `.form.txt` **no declara `OnActivate`**: el `Requery` no se ejecuta nunca por la activación del formulario. Es el mismo caso que la rama SICA del panel de indicadores, donde el padre solo comprueba un `SourceObject` que nadie asigna.
- `DoCmd.Hourglass False` (`:24`) se ejecuta en el manejador de error de `Form_Current` aunque este formulario nunca pone el reloj en `True`; puede apagar el indicador que otro formulario puso.
- `Form_Load` (`:29-35`) no tiene manejador de errores: si el `RecordSource` o el `Requery` fallan, no hay mensaje ni cierre, a diferencia de todos los demás formularios de la aplicación.
- Doble fuente para el mismo usuario: la fila visible viene de `TbUsuariosHistoricosLocal` (`:30-31`) y el objeto global de `TbUsuariosHistoricos` (`Constructor.bas:1215-1217`). Si la sincronización no ha corrido, se muestran y se publican datos distintos. Además, `getUsuarioHistorico` deja el `Recordset` abierto cuando está vacío: el `Exit Function` de `Constructor.bas:1231` está **después** del `rcdDatos.Close` (`Constructor.bas:1232-1233`), código inalcanzable.
- `m_Indicadores_LifecycleMoveCompleted` ignora `p_Direction` y `p_UserId` (`:43`): ante un movimiento que afecta a otro usuario, la hoja se recarga igualmente; y como no distingue dirección, tampoco distingue un alta de una baja. En la misma línea, el `RecordSource` de diseño (.form.txt:83) y el de `Form_Load` (`:30-31`) son duplicados del mismo `SELECT` con dos redacciones distintas, y `Me.Requery` aparece en tres sitios con tres envoltorios —sin protección (`:33`), con `On Error Resume Next` (`:39`, `:45`)— sin criterio declarado para el fallo.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: origen de datos único para la fila visible y para el objeto seleccionado; y recarga por suscripción en vez de `Requery` indiscriminado.

## Criterios de aceptación

Propuestos, no aprobados:

1. La hoja muestra el contenido completo de `TbUsuariosHistoricosLocal` como datasheet, con sus veinte columnas.
2. Al desplazarse entre filas, `m_ObjUsuarioHistoricoActivo` queda informado del usuario de la fila actual.
3. Al seleccionar la primera fila, los botones de ver y eliminar del padre quedan operativos; sin selección, el padre avisa al usuario.
4. Al terminar un movimiento del ciclo de vida, la hoja se recarga con el contenido actual de la caché local.
5. Si la fila no tiene `ID`, no se publica ningún usuario y el padre lo comunica al usuario en lugar de operar sobre `Nothing`.

## Pendientes

- Confirmar en el binario si existe un `OnActivate` no exportado a `src/`; si no lo hay, decidir si `Form_Activate` se elimina o se conecta.
- Estado de la sincronización de `TbUsuariosHistoricosLocal` en el momento de la lectura: `Mod_StartupCacheInitialization.bas:100-106` documenta que, sin regeneración al arranque, la hoja llegaba a mostrar una caché obsoleta.
- Decidir si la fila visible y el objeto global deben salir de la misma tabla, y en tal caso cuál; y el orden de refresco entre el bus de indicadores (`:43-47`) y el `Requery` manual del padre (`Form_FormInicial08UsuariosHistoricos.form.txt:1462`).