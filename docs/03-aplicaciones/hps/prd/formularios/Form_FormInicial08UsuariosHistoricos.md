# Form_FormInicial08UsuariosHistoricos

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial08UsuariosHistoricos.cls` (225 líneas) y `.form.txt` (1615 líneas) · **PRD:** [README](../README.md)

Hoja de datos del histórico de usuarios HPS dados de baja. Aloja el subformulario `FormInicial08UsuariosHistoricosDatos` en `SubFormResultados` y mantiene su `RecordSource` reescrito según el filtro de nombre (`EstablecerFiltroUsuariosHistoricos`, `:112-148`); la única acción de escritura que dispara es **eliminar definitivamente** un usuario histórico (`ComandoEliminarUsuarioHistorico_Click`, `:35-84`). Publica además una función pública (`EstablecerFiltroUsuariosHistoricos`, `:112-148`) que un consumidor externo (`Form_FormUsuario07SICAEdicion.cls:206-209`) usa como puente para pasar de la edición SICA al histórico. Las citas son `Form_FormInicial08UsuariosHistoricos.cls:línea` salvo indicación; `FG.bas` abrevía `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Menú lateral de `FormInicial01Lateral` | El lateral monta este formulario como bloque del histórico de usuarios | `Form_FormInicial01Lateral.cls:47, 63` (referenciado, no incluido) |
| `Form_FormUsuario07SICAEdicion` | Tras editar un usuario SICA, fija `Me.Nombre` y dispara `Nombre_Exit` + `ComandoVerUsuarioHistorico_Click` directamente sobre este formulario | `Form_FormUsuario07SICAEdicion.cls:206-209` |
| Subformulario `FormInicial08UsuariosHistoricosDatos` | `Me.SubFormResultados.SourceObject = "Form.FormInicial08UsuariosHistoricosDatos"` en cada apertura / filtro / exportación | `:129, 156, 196`, `.form.txt:896-897` |

Abre como destino a `FormUsuarioHistorico00Principal` (`ComandoVerUsuarioHistorico_Click`, `:10-33`), pero no a otros formularios.

## Controles y eventos

`OnLoad ="[Event Procedure]"` (`.form.txt:375`). El formulario no tiene `RecordSource` propio: lo que se filtra es `SubFormResultados`.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Aplica paleta, estiliza botones y monta `SubFormResultados` | `:177-206`, `.form.txt:375` |
| `Nombre` | caja de texto | `OnExit` → `Nombre_Exit` | Filtra el histórico por nombre al perder foco | `:210-225`, `.form.txt:1184` |
| `ComandoLimpiarNombre` | botón | `Click` → `ComandoLimpiarNombre_Click` | Pone `Nombre = Null` y vuelve a filtrar | `:86-90`, `.form.txt:1209` |
| `ComandoVerUsuarioHistorico` | botón | `Click` → `ComandoVerUsuarioHistorico_Click` | Abre `FormUsuarioHistorico00Principal` para el usuario activo | `:10-33`, `.form.txt:965` |
| `ComandoEliminarUsuarioHistorico` | botón | `Click` → `ComandoEliminarUsuarioHistorico_Click` | Borra definitivamente al usuario histórico activo | `:35-84`, `.form.txt:1011` |
| `ComandoListaEnExcel` | botón | `Click` → `ComandoListaEnExcel_Click` | Exporta el subformulario actual a Excel y abre el fichero | `:150-173`, `.form.txt:922` |
| `Comando5`, `Comando6` | botones | — | Restos del diseño binario sin handler en el `.cls` | `.form.txt:1058, 1104` |
| `SubFormResultados` | subformulario | — | Aloja `FormInicial08UsuariosHistoricosDatos` y recibe el `RecordSource` filtrado | `:129, 156, 196`, `.form.txt:896` |
| `lblNombre` | etiqueta | — | Rótulo de la caja de nombre | `.form.txt:1156` |

## Flujos

### Apertura (`Form_Load`, `:177-206`)

1. `EstablecerControlesDatos Me, m_Error` (`:183`; `FG.bas:1005`) y `Err.Raise 1000` si falla.
2. `CopiarEstiloBoton` sobre `ComandoVerUsuarioHistorico` (`:187`) y `ComandoEliminarUsuarioHistorico` (`:191`), con la misma cadena de comprobación.
3. Fija `Me.SubFormResultados.SourceObject = "FormInicial08UsuariosHistoricosDatos"` (`:196`): el subformulario queda montado aunque el filtro inicial lo deje vacío.
4. En error: `CorreoAlAdministrador` salvo 1000 y `MsgBox` (`:201-205`).

### Filtrar (`Filtrar`, `:92-110`)

1. Crea un `Scripting.Dictionary` nuevo con `CompareMode = TextCompare` y lo asigna a `m_ObjColParaFiltros` (`:95-96`), **reemplazando** cualquier colección previa.
2. Escribe `m_ObjColParaFiltros("Nombre") = Nz(Me.Nombre, "")` (`:97`).
3. Llama a `EstablecerFiltroUsuariosHistoricos m_Error` (`:99`) y eleva 1000 si vuelve con error.
4. En error: `CorreoAlAdministrador` salvo 1000 y `MsgBox` (`:106-108`).

`ComandoLimpiarNombre_Click` (`:86-90`) pone `Me.Nombre = Null` (`:87`) y llama a `Nombre_Exit False` (`:88`): limpiar la caja equivale a perder foco con texto vacío.

### Aplicar filtro sobre el subformulario (`EstablecerFiltroUsuariosHistoricos`, `:112-148`)

Función **`Public`** pensada para que un consumidor externo la invoque (`Form_FormUsuario07SICAEdicion.cls:206-209`):

1. Compone `m_SQLInicial = "SELECT * FROM TbUsuariosHistoricosLocal"` (`:126-127`).
2. Fija `SubFormResultados.SourceObject = "FormInicial08UsuariosHistoricosDatos"` (`:129`) — asegura el subformulario.
3. `Set frm = Me.Controls("SubFormResultados").Form` (`:130`) toma el subformulario.
4. `m_Nombre = Nz(Me.Nombre, "")` (`:131`).
5. Si `m_Nombre = ""` deja la `WHERE` como `";"`; en caso contrario compone `WHERE Nombre & ' ' & Apellido_1 & ' ' & Apellido_2 Like '*<texto>*';` (**concatenación simple del nombre + apellidos**, `:135`).
6. `frm.RecordSource = m_SQL` (`:141`) reasigna el origen del datasheet hijo.
7. En error, mensaje en `p_Error` (`:145`).

### Salir del filtro por la caja (`Nombre_Exit`, `:210-225`)

`EstablecerFiltroUsuariosHistoricos m_Error` (`:213`) y `Err.Raise 1000` si falla.

### Ver usuario histórico (`ComandoVerUsuarioHistorico_Click`, `:10-33`)

1. Si `m_ObjUsuarioHistoricoActivo Is Nothing`, eleva 1000 (`:14-17`).
2. Si `FormularioAbierto("FormUsuarioHistorico00Principal")`, lo cierra (`:19-21`).
3. `DoCmd.OpenForm "FormUsuarioHistorico00Principal"` (`:23`).
4. En error: `MsgBox` crítico (`:32`).

### Eliminar definitivamente (`ComandoEliminarUsuarioHistorico_Click`, `:35-84`)

1. Guarda contra `Nothing` (`:40-43`).
2. `MsgBox` con `vbDefaultButton2`, `vbYesNo`, `vbExclamation` (`:44-45`); si la respuesta no es «Sí», sale (`:47-49`).
3. `m_ObjUsuarioHistoricoActivo.Borrar m_Error` (`:53`; `classes/UsuarioHistorico.cls:731-749`); eleva 100 si falla (`:55`).
4. Si `m_ObjEntorno Is Nothing`, re-hidrata con `EVE` (`:56-58`).
5. `InicializarTablasLocalesYEntidad_Optimizado p_ForzarRegeneracionTotal:=EnumSiNo.No, p_Error:=m_Error` (`:63`; `FG.bas:5106`): regenera caches locales tras el borrado.
6. `ActualizarListaDatosLocal m_Error` (`:67`; `FG.bas:2060`).
7. `EstablecerContadores` (`:68`; `FG.bas:1440`) y `Me.SubFormResultados.Form.Requery` (`:71`).
8. En error: `MsgBox` crítico (`:82`).

### Exportar a Excel (`ComandoListaEnExcel_Click`, `:150-173`)

1. Asegura `SubFormResultados.SourceObject = "FormInicial08UsuariosHistoricosDatos"` (`:156`).
2. `m_URLExcelGenerado = ConsultaAExcel(Me.SubFormResultados.Form, , , EnumSiNo.Sí, m_Error)` (`:157`; `FG.bas:2660-2800`).
3. `Ejecutar Me.hWnd, "open", m_URLExcelGenerado, "", "", 1` (`:161`).
4. En error: `MsgBox` crítico (`:172`).

## Reglas, validaciones y cálculos

- **Borrado irreversible con default negativo.** `vbDefaultButton2` sobre `vbYesNo` (`:44`) sitúa el foco en «No»: la operación destructiva no se dispara con Intro por descuido.
- **El histórico se ata por objeto global.** `m_ObjUsuarioHistoricoActivo` (`VariablesEntorno.bas:227`) es el sujeto de las cuatro acciones: ver (`:14, 18-23`), eliminar (`:39-42, 52`), filtrar (`:86-90, 92-110, 210-225`); no se pasa identificador por argumento.
- **`EstablecerFiltroUsuariosHistoricos` es `Public` por una razón.** El consumidor `Form_FormUsuario07SICAEdicion.cls:206-209` lo usa tras editar un usuario SICA: pone nombre, dispara `Nombre_Exit` y luego `ComandoVerUsuarioHistorico_Click`. El formulario no es autocontenido.
- **Filtro estilo «contiene» sobre Nombre + Apellidos.** A diferencia del resto de filtros del bloque 03, este une `Nombre & ' ' & Apellido_1 & ' ' & Apellido_2` (`:135`) y aplica `Like '*<texto>*'`. Una búsqueda parcial tipo «Tor» puede traer tanto un «Torralba» como un «Victoriano» si la pieza aparece en el apellido compuesto.
- **`Filtrar` pisa el global `m_ObjColParaFiltros` con uno nuevo** (`:95-97`). No hay fusión con la colección previa.
- **Borrado regenera caches.** El borrado no se limita a la tabla histórica: pide resincronizar (`InicializarTablasLocalesYEntidad_Optimizado`, `:63`) y re-actualiza la lista local (`:67`).
- **Nula vista en caso de error.** `EstablecerFiltroUsuariosHistoricos` no usa `Err.Raise` en su `errores` (`:145`): quien la llama tiene que leer `p_Error` para enterarse.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Carga de la rejilla filtrada | `TbUsuariosHistoricosLocal` | `SELECT *` con `WHERE` opcional (`Like '*…*'`) | `:126-127, 135, 141` |
| Sincronización tras borrado | múltiples `Tb…Local` | `DELETE + INSERT` masivos | `:63`, `FG.bas:5106` |
| Actualización de la lista local | `TbUsuariosHistoricosLocal` | `INSERT` / `UPDATE` sobre la tabla local | `:67`, `FG.bas:2060` |
| Borrado de usuario histórico | `TbUsuariosHistoricos` (vía `UsuarioHistorico.Borrar`) | `DELETE` con borrado de anexos recursivo | `:53`, `classes/UsuarioHistorico.cls:731-749` |
| Recalculo de contadores del bloque superior | múltiples `TbDatosLocal` (vía `EstablecerContadores`) | agregados | `:68`, `FG.bas:1440-1523` |
| Exportación a Excel | (origen → `.xlsx`) | `ConsultaAExcel` | `:157` |

No hay `INSERT`/`UPDATE`/`DELETE` directos desde este `.cls` — todo pasa por el servicio `UsuarioHistorico.Borrar` o por rutinas de `FG.bas`.

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: **estado global leído y escrito** (`m_ObjUsuarioHistoricoActivo`, `VariablesEntorno.bas:227`) entra por el lateral, sale por aquí y vuelve al cerrarse el formulario hijo `FormUsuarioHistorico00Principal` (`:19-23`); `Public Function EstablecerFiltroUsuariosHistoricos` (`:112`) acoplada a `Form_FormUsuario07SICAEdicion.cls:208` por contrato informal; `Set frm = Me.Controls("SubFormResultados").Form` (`:130`) acoplado a `FormInicial08UsuariosHistoricosDatos`, hoja-datos que **no está documentada en este directorio**; concatenación SQL `Like '*" & m_Nombre & "*'` (`:135`) vulnerable a inyección — la entrada es la caja `Nombre`, validada solo por la mano del usuario.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComandoEliminarUsuarioHistorico_Click` pide confirmación y luego avanza sin esperar doble confirmación. La protección es el `vbDefaultButton2` (`:44`).
- `Me.Borrar m_Error` no eleva 1000 cuando falla; eleva 100 (`:55`). Quien capture el error debe acordarse de los dos números.
- `Filtrar` (`:95-97`) hace `Set m_ObjColParaFiltros = New Scripting.Dictionary`: si el llamador pasó con la esperanza de añadir una clave extra, se la lleva vaciada.
- `EstablecerFiltroUsuariosHistoricos` (`:112-148`) concatena `Like '*" & m_Nombre & "*'` sin escapar comillas.
- `Me.SubFormResultados.Form.Requery` en `:71` puede no encontrar el subformulario si el formulario se cierra entre `Borrar` y `Requery` por alguna ruta de error.
- `Comando5` y `Comando6` (`.form.txt:1058, 1104`) están en el binario sin handler.
- Tras eliminar, `m_ObjUsuarioHistoricoActivo` queda apuntando al usuario borrado; no hay reinicialización explícita.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtro parametrizable por más campos (`Apellido_1`, `Apellido_2` por separado); endpoint `DELETE` idempotente con `soft-delete` por defecto y borrado físico con doble confirmación; sustitución de la concatenación SQL por consulta parametrizada.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, el subformulario `FormInicial08UsuariosHistoricosDatos` queda montado y vacío de filtro.
2. Escribir un nombre y perder foco aplica `WHERE Nombre & ' ' & Apellido_1 & ' ' & Apellido_2 Like '*texto*'`.
3. `ComandoLimpiarNombre` pone `Nombre = Null` y vuelve a servir todas las filas.
4. `ComandoVerUsuarioHistorico_Click` con un usuario histórico activo cierra cualquier instancia previa de `FormUsuarioHistorico00Principal` y abre una nueva.
5. `ComandoEliminarUsuarioHistorico_Click` confirma con «No» por defecto; si la respuesta es «Sí», borra al usuario y deja la rejilla consistente tras la resincronización.
6. `ComandoListaEnExcel_Click` exporta la rejilla a `.xlsx` y abre el fichero con el visor por defecto.
7. `EstablecerFiltroUsuariosHistoricos` se puede invocar desde fuera con un valor de `Nombre` y devuelve el mismo resultado que perder foco en la caja.

## Pendientes

- Documentar `Form_FormInicial08UsuariosHistoricosDatos`, la hoja-datos que aloja `SubFormResultados`.
- Contenido de `classes/UsuarioHistorico.cls:731-749` (`Borrar`) — qué tablas auxiliares toca.
- Si `Comando5`/`Comando6` tienen `OnClick` en algún sitio (a verificar contra `.form.txt:1058, 1104`).
