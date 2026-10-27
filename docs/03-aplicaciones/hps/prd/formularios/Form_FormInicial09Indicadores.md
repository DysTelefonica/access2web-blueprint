# Form_FormInicial09Indicadores

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial09Indicadores.cls` (712 líneas) y `.form.txt` (1940 líneas) · **PRD:** [README](../README.md)

Panel de indicadores de HPS. Se monta en el `SubFormCentral` del contenedor raíz cuando el lateral pulsa `lblIndicadores` (`Form_FormInicial01Lateral.cls:434-436`) y decide qué hoja de datos muestra: su subformulario `SubFormIndicadores` arranca apuntando a `Form.FormInicial09IndicadoresDatosLocalDatos` (`.form.txt:794`) y cada uno de los siete botones de indicador delega en `EstablecerFuenteDatosIndicadores` (`FG.bas:3056`), que fija a la vez el `WHERE` y el `SourceObject`. La lógica de este `.cls` es casi toda de orquestación: el SQL, los contadores y el estilo viven en módulos. Las citas son `Form_FormInicial09Indicadores.cls:línea` salvo indicación; `FG.bas` abrevía `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial01Lateral` (lateral) | `SubFormCentral.SourceObject = "FormInicial09Indicadores"` | `Form_FormInicial01Lateral.cls:435-436` |
| `Form_FormInicial01Lateral` (lateral) | Invoca los siete handlers públicos por código, no por el `Click` del control | `Form_FormInicial01Lateral.cls:365, 372, 401, 408, 415, 422, 466` |
| `FG.bas` (`EstablecerIndicadores`, `EstablecerFuenteDatosIndicadores`, `EstablecerFuenteDatosIndicadoresConEmpresas`) | Entran por `Forms("FormInicial00Principal") → SubFormCentral` y solo actúan si el `SourceObject` encaja | `FG.bas:2921-2924`, `FG.bas:3067-3071`, `FG.bas:3158-3162` |
| `FormUsuario02Edicion` / `FormUsuario07SICAEdicion` | Sus eventos `EdicionRegistro` recomputan contadores y vacían `SubFormIndicadores.SourceObject` | `:646-665`, `:667-686` |
| `clsIndicadoresBus` (`modIndicadores.m_Indicadores`) | `LifecycleMoveCompleted` refresca contadores y rejilla | `:704-712`, `modIndicadores.bas:13` |

Abre `FormUsuario00Principal` o `FormUsuario07SICAEdicion` desde `ComandoEditarUsuario` (`:152-161`) y genera un fichero Excel con `ConsultaAExcel` (`:213`).

## Controles y eventos

`DefaultView =0` (`.form.txt:14`), `OnLoad ="[Event Procedure]"` (`.form.txt:223`) y sin `RecordSource` de formulario: los datos viven en el subformulario.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad` → `Form_Load` | Sincroniza caches, refresca contadores, arma los combos y aplica el estilo de botones | `:563-599`, `.form.txt:223` |
| `SubFormIndicadores` | subformulario | — | Hoja de datos que muestra el indicador activo | `.form.txt:793-794` |
| `cmdActualizarIndicadores` | botón | `Click` → `cmdActualizarIndicadores_Click` | Resincroniza caches locales y vuelve a pintar contadores y rejilla | `:11-52`, `.form.txt:1237-1239` |
| Siete botones de indicador: `ComandoPendientesCursoHPS`, `ComandoAPuntoDeCaducar`, `ComandoCaducados`, `ComandoEnSolicitud`, `ComandoPendientesConvocatoria1`, `ComandoPendientesConvocatoria2`, `ComandoPendientesCorreoJefeSeguridad` | botones | `Click` → handler homónimo | Fijan un `EnumTipoIndicador`, rotulan `lblFiltroAplicado` y muestran u ocultan el filtro por empresa | `:61-83, 85-107, 175-203, 231-265, 267-302, 304-338, 340-376`, `.form.txt:598, 644, 691, 738, 899, 945, 991` |
| `ComandoEditarUsuario` | botón | `Click` → `ComandoEditarUsuario_Click` | Resuelve el usuario activo y abre su formulario de edición | `:137-173`, `.form.txt:819-821` |
| `ComandoListaEnExcel` | botón | `Click` → `ComandoListaEnExcel_Click` | Exporta el subformulario y abre el fichero generado | `:205-229`, `.form.txt:866-868` |
| `ComandoCopiaAPortaPapelesFiltrado` | botón | `Click` → `ComandoCopiaAPortaPapelesFiltrado_Click` | Copia los correos de las filas filtradas al portapapeles | `:109-135`, `.form.txt:1047-1049` |
| `ComboTipoEmpresaFiltro` | combo (lista de valores) | `OnChange` → `ComboTipoEmpresaFiltro_Change` | Vacía y rellena `ComboNombreEmpresa` | `:533-561`, `.form.txt:1103-1107` |
| `ComboNombreEmpresa` | combo (lista de valores) | `AfterUpdate` → `ComboNombreEmpresa_AfterUpdate` | Aplica el filtro por empresa seleccionada | `:498-531`, `.form.txt:1177-1180` |
| `lbl1`, `lblFiltroAplicado` | etiquetas | — | Rótulo del filtro de empresa y contador del filtro aplicado | `:434-436`, `.form.txt:1140-1141, 1208-1209` |

## Flujos

### Apertura (`Form_Load`, `:563-599`)

1. `CopiarDatosAIndicadores m_Error, p_CacheAcabDeRegenerar:=True` (`:569`; `FG.bas:3006`): fuerza la sincronización de `TbDatosLocalParaIndicadores` y `TbUsuariosSICALocalParaIndicadores` desde sus tablas de origen, ignorando el `Static` de salto (`FG.bas:3011, 3022-3025`).
2. `EstablecerIndicadores m_Error` (`:574`; `FG.bas:2907`): vuelve a copiar los datos, crea `m_ObjIndicadores` (`FG.bas:2919`) y reescribe el `Caption` de los siete botones con `Replace(m_Titulo, "#", intNumero)` (`FG.bas:2933-2934, 2945-2946, 2955-2956, 2965-2966, 2976-2977, 2987-2988, 2997-2998`).
3. Carga la lista de tipos de empresa y vacía el combo de nombres (`:579-580`).
4. Deja `ComandoEditarUsuario` deshabilitado y le aplica el estilo, igual que a `cmdActualizarIndicadores` (`:581-586`).
5. Se engancha al bus de indicadores y, si la edición HPS ya está abierta, captura su referencia (`:587-590`).

### Indicador (`ComandoPendientesCursoHPS_Click`, `:340-376`; modelo común de los siete)

1. Hourglass y `VBA.DoEvents` antes y después de la llamada (`:346-348, 361-363`).
2. Limpia la última empresa seleccionada y anula los dos combos (`:351-354`).
3. `EstablecerFuenteDatosIndicadores m_EnumTipoIndicador, m_Error` (`:356`; `FG.bas:3056`) es quien compone el `SELECT`, fija `RecordSource` y hace `Requery` (`FG.bas:3076-3077, 3136-3139`).
4. `Me.lblFiltroAplicado.Caption = m_ObjEntorno.colFiltroEnIndicadordes(Me.ComandoPendientesCursoHPS.Name)` (`:367`): el texto del filtro sale de la configuración, no de una constante del código.
5. `HacerVisible EnumSiNo.Sí` (`:365`) es la excepción: solo el indicador de curso HPS deja visibles los combos de empresa; los otros seis llaman a `HacerVisible EnumSiNo.No` (`:74, 98, 191, 252, 288, 326`).
6. En error: `CorreoAlAdministrador` solo si `Err.Number <> 1000` y `MsgBox` con `m_Error` (`:370-375`).

### Filtro por empresa (`ComboTipoEmpresaFiltro_Change`, `:533-561`; `ComboNombreEmpresa_AfterUpdate`, `:498-531`)

1. Al cambiar el tipo, vacía `ComboNombreEmpresa` y su `RowSource` (`:537-542`).
2. Llama a `RellenarComboEmpresas` con `EnumSiNo.No` o `EnumSiNo.Sí` según la rama de `:543-553` (`FG.bas:1100`).
3. Al confirmar un nombre, compara contra `m_UltimoTipoEmpresa` / `m_UltimoNombreEmpresa` para no repetir el filtrado (`:514-516`).
4. `EstablecerFuenteDatosIndicadoresConEmpresas m_TipoEmpresa, m_NombreEmpresa, m_Error` (`:519`; `FG.bas:3147`) reescribe el `RecordSource` del subformulario sustituyendo su `WHERE`, y sale sin hacer nada si el `SourceObject` no es la hoja de datos local (`FG.bas:3166-3168`).

### Editar usuario (`ComandoEditarUsuario_Click`, `:137-173`)

1. `EstablecerUsuarioActivoIndicador m_Error` (`:143`; `:601-644`) pone a `Nothing` los ocho objetos globales de usuario, observación, anexo y ficha (`:608-616`).
2. Si el `SourceObject` es la hoja SICA, lee el `ID` del subformulario y llama a `getUsuarioSICA` (`:617-626`; `Constructor.bas:1254`); en caso contrario, `getUsuarioHPS` (`:633`; `Constructor.bas:743`). Con `ID` vacío la función sale en silencio y el botón avisa «Se ha de seleccionar el usuario primero» (`:147-149`).
3. Abre `FormUsuario00Principal` (`:155`) o `FormUsuario07SICAEdicion` (`:160`), cerrando antes el que estuviera abierto, y solo en el segundo caso captura `m_Frm` para recibir `EdicionRegistro` (`:161`).

### Exportar y copiar (`ComandoListaEnExcel_Click`, `:205-229`; `ComandoCopiaAPortaPapelesFiltrado_Click`, `:109-135`)

1. Excel: `ConsultaAExcel(Me.SubFormIndicadores.Form, , , EnumSiNo.Sí, m_Error)` (`:213`; `FG.bas:382`) y `Ejecutar Me.hWnd, "open", m_URLExcelGenerado` (`:215`).
2. Portapapeles: `CopiaAPortapapelesFiltrado` (`:121`; `:444-496`) clona el `Recordset`, recorre las filas filtradas y concatena `m_UsuarioHPS.Correo_e` con `;` solo si contiene `@` (`:467-481`), y aborta con «Ningún elemento filtrado» si el clon está vacío (`:459-464`).

### Reactivación (`Form_Activate`, `:688-702`; `m_Indicadores_LifecycleMoveCompleted`, `:704-712`)

Ambos recomputan `EstablecerIndicadores`; el segundo, además, hace `Requery` del subformulario si tiene `SourceObject`, y lo hace con `On Error Resume Next` (`:705-711`).

## Reglas, validaciones y cálculos

- **El constructor exige el ancla y sale en silencio si no está.** `EstablecerFuenteDatosIndicadores` toma `Set frm = Forms("FormInicial00Principal")` (`FG.bas:3067`) y hace `Exit Function` cuando `SubFormCentral.SourceObject <> "FormInicial09Indicadores"` (`FG.bas:3068-3070`); es el mismo patrón que `EstablecerFiltroUsuarios` (`Filtro.bas:21-23`). Los tres constructores de indicadores repeten la guarda (`FG.bas:2922`, `FG.bas:3068`, `FG.bas:3159`).
- **Siete ramas con `WHERE` escritos a mano**, todas sobre el mismo `m_SQLInicial` (`FG.bas:3076-3077`) y todas con la misma asignación de `SourceObject` (`FG.bas:3086-3088` y las seis siguientes):

| `EnumTipoIndicador` | `WHERE` literal | Evidencia |
|---|---|---|
| `PendientesCursoHPS` | `Requiere_Curso='Sí';` | `FG.bas:3084-3085` |
| `PtesPrimeraConvocatoriaCurso` | `((TbDatosLocalParaIndicadores.Requiere_PrimeraConvocatoriaCurso)='Sí'));` | `FG.bas:3090-3091` |
| `PtesSegundaConvocatoriaCurso` | `Requiere_SegundaConvocatoriaCurso='Sí';` | `FG.bas:3096-3097` |
| `PtesEnvioCorreoJefeSeguridadCurso` | `Requiere_CorreoJefeSeguridadCurso='Sí';` | `FG.bas:3102-3103` |
| `HPSAPuntoDeCaducar` | `HPS_NAC_ApuntoDeCaducar='Sí' OR HPS_OTAN_ApuntoDeCaducar='Sí' or HPS_ESA_ApuntoDeCaducar='Sí' OR HPS_UE_ApuntoDeCaducar='Sí';` | `FG.bas:3109-3113` |
| `HPSACaducadas` | `HPS_NAC_Caducado='Sí' OR HPS_OTAN_Caducado='Sí' or HPS_ESA_Caducado='Sí' OR HPS_UE_Caducado='Sí';` | `FG.bas:3118-3122` |
| `HPSASolicitandose` | `Not HPS_UE_F_Solicitud Is Null OR Not HPS_OTAN_F_Solicitud Is Null OR NOt HPS_ESA_F_Solicitud Is Null OR Not HPS_UE_F_Solicitud Is Null;` | `FG.bas:3127-3131` |

- **Los contadores no se calculan aquí.** `EstablecerIndicadores` instancia `m_ObjIndicadores` (`FG.bas:2919`) y lee sus propiedades `NPendientesCursoHPS`, `NHPSAPuntoDeCaducar`, `NHPSACaducadas`, `NHPSASolicitandose`, `NUsuariosPtesPrimeraConvocatoriaCurso`, `NUsuariosPtesSegundaConvocatoriaCurso` y `NUsuariosPtesEnvioCorreoJefeSeguridadCurso` (`FG.bas:2928, 2940, 2950, 2960, 2971, 2982, 2992`), con el número embebido en el texto del botón.
- **Protocolo de error uniforme.** Cada procedimiento usa `m_Error` y `Err.Raise 1000` para cortar; el `If Err.Number <> 1000` evita el correo al administrador en los cortes previstos y el `MsgBox` final muestra `m_Error` (`:45-51, 76-82, 166-172, 696-701`).
- **`HacerVisible` solo toca tres controles**: la etiqueta y los dos combos (`:434-436`). No oculta `SubFormIndicadores` ni los botones.
- **`CopiaAPortapapeles` (`:378-420`) no tiene ninguna llamada en `src/`**: solo se usa su variante `CopiaAPortapapelesFiltrado` (`:121`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Sincronización de las tablas de los indicadores | `TbDatosLocalParaIndicadores`, `TbUsuariosSICALocalParaIndicadores` desde `TbDatosLocal` y `TbUsuariosSICALocal` | `DELETE *` + `INSERT INTO … SELECT` | `FG.bas:3032-3046`, llamada en `:569` |
| Origen de la hoja de datos | `TbDatosLocalParaIndicadores` | `RecordSource` reasignado con `WHERE` literal | `FG.bas:3136-3139` |
| Contadores de los siete indicadores | `Indicador` | propiedades calculadas | `FG.bas:2919, 2928-2998` |
| Usuario activo de la fila | `TbUsuarios` (`getUsuarioHPS`), `TbUsuariosSICA` (`getUsuarioSICA`) | `SELECT *` por `ID` | `Constructor.bas:760-762`, `Constructor.bas:1271-1273`, llamadas en `:623, 633` |
| Correos para el portapapeles | `TbUsuarios` por cada `ID` de la hoja | `SELECT` en bucle sobre el `RecordsetClone` | `:466-483` |
| Exportación a Excel | `RecordSource` del subformulario | `ConsultaAExcel` | `:213`, `FG.bas:382` |
| Filtro por empresa | `RecordSource` vigente | sustitución del `WHERE` | `:519`, `FG.bas:3171-3177` |

Este `.cls` no contiene `INSERT`, `UPDATE` ni `DELETE` propios (`:1-712`).

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Ya se ve: los siete `WHERE` literales (`FG.bas:3085-3131`) son SQL en cadena con comillas simples que la web tendría que parametrizar; la cadena `Forms("FormInicial00Principal").Controls("SubFormCentral").Form.Controls("SubFormIndicadores").Form` (`FG.bas:3071-3136`) ata este panel al contenedor raíz por nombre, igual que el bloque 03 se ata a `SubFormResultados` (`Filtro.bas:24-26`); el `RecordSource` del subformulario se usa como estado de la vista (`FG.bas:3171`), algo sin equivalente directo en una ruta web; `DoCmd.Hourglass` y `VBA.DoEvents` gobiernan el bloqueo de interfaz (`:18, 346-348, 361-363`); y `Application.TempVars` más los globales `m_ObjUsuarioActivo`, `m_ObjUsuarioSICAActivo`, `m_ObjObservacionActiva`, `m_ObjAnexoUsuarioHPSActivo`, `m_ObjHPSNACActiva`, `m_ObjHPSOTANActiva`, `m_ObjHPSESAActiva`, `m_ObjHPSUEActiva` (`:608-616`) son estado compartido entre formularios.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **La rama «solicitándose» repite una columna y se deja otra fuera** (`FG.bas:3127-3131`): `HPS_UE_F_Solicitud` aparece en el primer y en el último `OR`, y `HPS_NAC_F_Solicitud` no aparece en ninguno. Las otras dos ramas equivalentes (`FG.bas:3109-3113`, `FG.bas:3118-3122`) sí enumeran las cuatro columnas de ficha, lo que apunta a un defecto de copiado y no a una decisión de negocio.
- `ComboTipoEmpresaFiltro_Change` compara `Me.ComboNombreEmpresa = "Jurídica Contratación"` (`:543`) cuando el control acaba de anularse a `Null` dos líneas antes (`:541`); la comparación solo puede ser cierta por la primera condición, así que la segunda rama nunca distingue el tipo «Jurídica Contratación».
- El rótulo `lbl1` dice «Tipo de Jurídica para filtrar» (`.form.txt:1141`) mientras el combo ofrece tres tipos (`:579`, `.form.txt:1105`).
- `Form_Activate` (`:688-702`) no está enganchado: el `.form.txt` no declara `OnActivate` (búsqueda sin resultados en el fichero). Es la misma conclusión que ya se registró para `lbl` en `Form_FormInicial00Principal`.
- `CopiaAPortapapeles` (`:378-420`) es código muerto: ningún procedimiento de `src/` la invoca, y sus cuatro mensajes de portapapeles describen reglas (primera convocatoria, segunda convocatoria, correo al jefe de seguridad) que hoy solo se aplican por código en la variante `Filtrado`.
- `HacerVisible` (`:422-425`) y `EstablecerUsuarioActivoIndicador` (`:601-603`) se declaran `As String` y nunca asignan su valor de retorno; sus errores se pierden porque quien las llama solo comprueba `m_Error`, no el retorno.
- `EstablecerUsuarioActivoIndicador` pone a `Nothing` los ocho globales (`:608-616`) y después sale sin restaurarlos si `m_ID` está vacío (`:620-622, 630-632`): un `Click` con la hoja sin fila borra el usuario activo de todo el formulario y luego muestra un aviso genérico (`:147-149`).
- `m_Frm` solo se captura en la rama SICA (`:161`); en la rama HPS queda sin asignar, así que `m_Frm_EdicionRegistro` (`:646`) nunca se dispara por este camino y el `SourceObject` no se vacía tras editar desde `FormUsuario00Principal`.
- `HacerVisible EnumSiNo.Sí` (`:365`) muestra el filtro por empresa para el indicador de curso HPS, pero `EstablecerFuenteDatosIndicadoresConEmpresas` sale sin filtrar si el `SourceObject` no es la hoja de datos local (`FG.bas:3166-3168`); el usuario ve los combos y su elección no surte efecto.
- `cmdActualizarIndicadores_Click` resincroniza caches (`:19`) y hace `Requery` del subformulario (`:33-38`), pero no vuelve a aplicar el `WHERE` del indicador activo: la rejilla puede mostrar el resultado del filtro anterior o quedar vacía si la cascada vació la tabla.
- El `MsgBox` de los caminos de error muestra `m_Error`, que está vacío cuando el corte fue un `Err.Raise 1000` sin haber escrito en él (`:51, 82, 106, 134`): el usuario ve un diálogo crítico en blanco.
- `pregunta` se usa como variable sin declarar en diecisiete líneas del archivo (`:82, 106, 126, 134, 172, 202, 228, 264, 301, 337, 375, 530, 560, 597, 664, 685, 701`) pese a `Option Explicit` (`:3`); es código heredado del padre y hay que contrastarlo contra el binario.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: sustituir los siete `WHERE` literales por un constructor de consulta parametrizado a partir del `EnumTipoIndicador` (`FG.bas:3083-3135`); exponer un único endpoint de indicadores que devuelva `{filtro, contador, filas}` en lugar de reescribir el `RecordSource` de un subformulario; convertir el bus `LifecycleMoveCompleted` en un evento de dominio que invalide un contador en caché en vez de recomputar en cada activación (`modIndicadores.bas:13`); y reemplazar la hoja SICA como variante por una columna discriminante en la consulta, lo que resolvería a la vez la rama muerta de `FormInicial09IndicadoresUsuariosSICADatos` (`:617`).

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir el panel, las tablas de indicadores quedan sincronizadas una vez y los siete botones muestran su contador actual.
2. Al pulsar un indicador, la hoja muestra solo las filas que casan con el `WHERE` de esa rama y `lblFiltroAplicado` refleja el filtro de la configuración.
3. La hoja de datos se monta vacía si el contenedor raíz no está abierto, sin error visible.
4. El filtro por empresa solo se ofrece en el indicador de curso HPS y su aplicación cambia el conjunto de filas.
5. Con una fila seleccionada, «Detalle» abre el formulario de edición del usuario que corresponde a la hoja SICA o a la local, y sin fila avisa sin abrir nada.
6. La exportación a Excel refleja exactamente el `WHERE` vigente en pantalla.
7. «Copiar correos» con filtro vacío avisa y no toca el portapapeles; con filtro, copia un correo por usuario separando con `;` y descartando los que no tienen `@`.

## Pendientes

- Confirmar contra el binario si `Form_Activate` (`:688`) está enganchado y si la vista SICA (`:617`) se monta desde algún punto fuera de `src/`.
- Confirmar si `CopiaAPortapapeles` (`:378-420`) tiene llamadas fuera del código exportado.
- Confirmar con negocio si la rama «solicitándose» debe incluir `HPS_NAC_F_Solicitud` (ver el defecto en `FG.bas:3127-3131`).
- Decidir el comportamiento esperado de «Actualizar Indicadores» respecto al filtro activo.