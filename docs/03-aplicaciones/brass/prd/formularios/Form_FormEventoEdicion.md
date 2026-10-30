# Form_FormEventoEdicion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormEventoEdicion.cls` (1308 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Edición de los datos generales de un evento (`TbEventos`): BUI, subsistema, equipo, tipo, criticidad, técnico, originador, descripción, fecha de alta, contacto, notas y fecha de inclusión en informe RAC. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): carga el registro con DAO en `EstablecerDatos` y persiste con la función pública `Edicion`, definida en el propio formulario. Si el evento está facturado o franqueado, se abre en modo de solo lectura y únicamente permite cambiar la fecha de informe RAC. Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`; `VG.bas` = `modules/Variables Globales.bas`. Las citas sin fichero son `Form_FormEventoEdicion.cls:línea`; las del `.form.txt` llevan el nombre completo. Capacidad: CAP-3 Eventos (los controles se asignan a ella salvo indicación).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEventoGestion` | Cierra el formulario si estaba abierto, fija la global `m_ObjEventoActivo` y lo abre con `IDEvento\|facturado\|franqueado` (facturado = marca `*` en la columna 0 de `ListaEventos`) | `Form_FormEventoGestion.cls:1590-1596, 1613-1618` |
| `Form_FormInformeSLA` | Igual; `strFranqueado` fijo a «Sí» (la lista solo muestra franqueados) y facturado vía `ObtenerEstadoFacturadoEvento` | `Form_FormInformeSLA.cls:238-246` |
| `Form_FormInformesParaRAC` | Deriva facturado (`FechaFacturado`) y franqueado (`FechaFinal`) del evento; escucha el evento `FechaInformeRACambiada` | `Form_FormInformesParaRAC.cls:36-51, 307-327` |

`OpenArgs` debe tener exactamente tres partes y las dos últimas valer «Sí» o «No»; si no, mensaje «No se han cargado todos los datos» y cierre (`:567-592, 622-628`). Hacia fuera: evento `FechaInformeRACambiada` (`:12`, lanzado en `:507`); abre `FormAnexos`, `FormEventoEqMedidaEdicion`, `FormEventoSLA`, `FormEquipoAlta` y `FormPlanificacionCierre` (ver tabla). Tras éxito de `ComandoEditar` se cierra (`:237`); `cmdSalir` cierra siempre (`:167`).

## Controles y eventos

Controles: 71 (71 Begin en el .form.txt) = 1 `Form` + 1 `Section` + 7 plantillas de estilo sin `Name` (`Form_FormEventoEdicion.form.txt:457-580`) + 62 controles con nombre.

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| Formulario (1) | Form | `OnOpen` → `Form_Open` | Valida `OpenArgs`, fija modo escritura/lectura, carga datos y técnicos | CAP-3 | `:557-629`; `Form_FormEventoEdicion.form.txt:78` |
| `Detalle` (1) | Section | — | Única sección | CAP-3 | `Form_FormEventoEdicion.form.txt:581` |
| Plantillas sin nombre (7) | Label, Rectangle, Image, CommandButton, OptionGroup, TextBox, ComboBox | — | Estilos por defecto, no son controles de usuario | sin asignar | `Form_FormEventoEdicion.form.txt:457-580` |
| `ComboBUI` | combo (lista `BUI`,`NODO`; `Column(0)`=BUI, `Column(1)`=nodo) | `AfterUpdate`, `OnEnter` | Al cambiar de BUI vacía subsistema y equipo, reconstruye sus listas y la del contacto; `Enter` hace `Requery` | CAP-12 | `:337-408`; `Form_FormEventoEdicion.form.txt:776-783` |
| `ComboSubsistema` | combo | `OnChange`, `OnEnter` | Habilita `ComandoNuevoIDEquipo`/`ComandoEquiposDeMedida` solo con BUI y subsistema; reconstruye `ComboIDEquipo` | CAP-12 | `:417-468`; `Form_FormEventoEdicion.form.txt:674-681` |
| `ComboIDEquipo` | combo de 3 columnas (ID, equipo, `FechaObsoleto`; solo la 2.ª visible) | `OnEnter` | `Requery`; obligatorio al guardar; valor = ID | CAP-5 | `:409-412, 514`; `Form_FormEventoEdicion.form.txt:994-1000` |
| `ComboTipoEvento`, `ComboOriginador` (2) | combos | `OnEnter` | `Requery`; obligatorios al guardar | CAP-3 / CAP-12 | `:413-416, 553-556`; `Form_FormEventoEdicion.form.txt:845-851, 923-929` |
| `ComboCriticidad` | combo (código y descripción) | — | Debe ser numérico al guardar | CAP-3 | `:1046-1049`; `Form_FormEventoEdicion.form.txt:886-890` |
| `ComboTecnico` | combo de lista de valores (alias, nombre, tipo) | — | Se rellena en `Form_Open` con `RellenarComboTecnicos`; obligatorio | CAP-9 | `:612-615`; `FG.bas:7224-7260`; `Form_FormEventoEdicion.form.txt:2311-2360` |
| `Contacto` | combo editable (sin `LimitToList`), valores distintos de `TbEventos.CONTACTO` | — | Lista por BUI tras cambiar BUI; opcional (`#` = anular) | CAP-3 | `:392, 1149-1155`; `FG.bas:6670-6687`; `Form_FormEventoEdicion.form.txt:2438-2443` |
| `DescripcionEvento`, `PMPR`, `Notas` (3) | texto | — | Descripción obligatoria; `PMPR` y `Notas` opcionales (`#` = anular) | CAP-3 | `:1058-1061, 1134-1140, 1157-1163`; `Form_FormEventoEdicion.form.txt:604, 810, 2247` |
| `FechaAlta` | texto, formato fecha corta | — | Obligatoria y debe ser fecha; interviene en equipo obsoleto e `IDEvento` | CAP-3 | `:1062-1082`; `Form_FormEventoEdicion.form.txt:955-957` |
| `FechaEnInformeRAC` | texto, fecha corta | — | Única edición permitida en solo lectura (`Locked=False`, `:897-899`) | CAP-10 | `:491-510, 525-529`; `Form_FormEventoEdicion.form.txt:2535` |
| `FechaUltimaActividad`, `HorasEvento`, `Franqueada` (3) | texto, `Tag=NOEDITABLE` | — | Fecha máxima de `TbActividades.FechaAlta` o «Sin Acts»; horas laborables + extra; Sí/No | CAP-4 | `:750-757, 832-842, 730-743`; `FG.bas:819-831`; `Form_FormEventoEdicion.form.txt:1677, 2057, 2219` |
| `CausaFin` | texto bloqueado | — | Solo se habilita y bloquea (`:908-909`); nunca recibe valor | sin asignar | `Form_FormEventoEdicion.form.txt:1054` |
| `Numero_Actividad` | texto, `ControlSource="IDActividad"` | — | No se referencia en el código | sin asignar | `Form_FormEventoEdicion.form.txt:633-634` |
| `ComandoEditar` (rótulo «Registrar») | botón | `OnClick` | `Registrar` y, si va bien, cierra el formulario | CAP-3 | `:213-246`; `Form_FormEventoEdicion.form.txt:1795-1797` |
| `cmdSalir` | botón | `OnClick` | Si hay cambios y modo escritura pregunta si guardar; refresca `ListaEventos` de `FormEventoGestion` si está abierto | CAP-3 | `:129-175`; `Form_FormEventoEdicion.form.txt:727-729` |
| `ComandoNuevoIDEquipo` («+») | botón | `OnClick` | Abre `FormEquipoAlta` con `nodo\|BUI\|subsistema` si los tres existen | CAP-5 | `:299-330`; `Form_FormEventoEdicion.form.txt:1023-1025` |
| `ComandoEquiposDeMedida` | botón | `OnClick` | Abre `FormEventoEqMedidaEdicion` y escucha `Alta`/`Eliminado` para actualizar el rótulo con el nº | CAP-5 | `:248-274, 1254-1308` |
| `ComandoSLU` («Datos SLU») | botón | `OnClick` | Abre `FormEventoSLA` con el `IDEvento` | CAP-3 | `:276-297` |
| `ComandoAdjuntar` | botón | `OnClick` | Fija `m_ObjEntidadParaAnexoActiva` y abre `FormAnexos` | CAP-11 | `:176-211` |
| `ImagenAnexo` | imagen | `OnClick` | Visible si el evento tiene anexos; abre un archivo o `ComandoAdjuntar_Click` si hay varios (`;`) | CAP-11 | `:13-36, 944-979` |
| `ComandoVerProgramacion` | botón | `OnClick` | Visible si el evento figura en `TbPlanificacion`; abre `FormPlanificacionCierre` | CAP-6 | `:848-860, 923-942` |
| `ComandoQuitarFechaInforme` («+») | botón | `OnClick` | Pone `FechaEnInformeRAC = Null` en el control | CAP-10 | `:332-335, 902-906` |
| `ComandoAyudaFormulario` | botón | `OnClick` declarado, sin handler | Sin efecto | sin asignar | `Form_FormEventoEdicion.form.txt:1442-1444` |
| `lblTitulo1`, `lblTitulo2` (2) | etiquetas | — | Título «Evento <id>» (+ «(Sólo lectura)») | CAP-3 | `:766, 794, 808-809` |
| `lblFacturado`, `lblFranqueado`, `lblInformeRAC` (3) | etiquetas | — | Visibles según facturado, franqueado y fecha de informe | CAP-8 / CAP-3 / CAP-10 | `:815-829` |
| Rótulos de campo (18) | etiquetas | — | Etiqueta244, 206-210, 213, 214, 217-219, 14, 15, 0, 13, 20, 26, 33 | CAP-3 | `Form_FormEventoEdicion.form.txt:1070-1340, 1693, 1720, 2015, 2125, 2263, 2457, 2552` |
| Asteriscos de obligatorio (9) | etiquetas «*» | — | Marca visual: Etiqueta99, 1, 2, 3, 6, 7, 8, 9, 21 | CAP-3 | `Form_FormEventoEdicion.form.txt:1486-1640, 2152` |
| `Cuadro2`, `Cuadro3` (2) | rectángulos | — | Decoración | sin asignar | `Form_FormEventoEdicion.form.txt:695, 1367` |

Suma: 1+1+7+8 combos+10 textos+9 botones+1 imagen+2 rectángulos+32 etiquetas = 71.

## Flujos

### Apertura (`Form_Open` → `EstablecerDatos`, `:557-629, 630-922`)

1. `strEscritura` = «No» si facturado o franqueado, «Sí» en otro caso (`:594-598`).
2. `Constructor.getEvento` carga el objeto y lo deja en `m_ObjEventoActivo` (`:646`); la consulta `TbEventos` ⟕ `TbEquipos` guarda los valores iniciales (`:667-691`) y rellena controles no vacíos (`:694-743`).
3. Modo escritura: colorea y desbloquea combos y textos salvo `NOEDITABLE` (`:765-773`); modo lectura: los bloquea y deshabilita `ComandoEditar`, `ComandoNuevoIDEquipo` y `ComandoEquiposDeMedida` (`:793-806`). Después se reactivan `FechaEnInformeRAC`, `ComandoEditar` y `ComandoEquiposDeMedida` también en lectura (`:897-910`).
4. Cualquier error: `MsgBox` crítico y cierre (`:622-628`).

### Guardado (`ComandoEditar_Click` → `Registrar`, `:213-246, 469-551`)

1. Lectura: sin cambio en `FechaEnInformeRAC` ⇒ «No ha habido cambios»; con fecha ⇒ `InformeRACRegistrarFecha`; vacía ⇒ `InformeRACQuitarFecha`; lanza `FechaInformeRACambiada` (`:490-510`; `Evento.cls:1442-1470`).
2. Escritura: lee los controles (`#` para vaciar `PMPR`, `Contacto`, `Notas`, `FechaEnInformeRAC`), actualiza el objeto activo y llama a `Edicion` (`:511-541`).
3. `Edicion`: valida (ver reglas), lee BUI anterior, si cambia genera `IDEvento` nuevo, edita `TbEventos` con DAO y, si hay `IDEvento` nuevo, ejecuta `ReNombrarIDEvento` (`:1022-1186`).
4. Error: texto en `MsgBox` crítico; el formulario queda abierto (`:239-245`). `cmdSalir` repite el guardado tras confirmación y, si falla, no cierra (`:133-149, 169-175`).

## Reglas, validaciones y cálculos

- **Obligatorios** (vacío o `#`): nodo, BUI, subsistema, equipo, tipo, técnico, originador, descripción; criticidad numérica; fecha de alta válida (`:1022-1064`). `Contacto`, `Notas`, `PMPR` son opcionales (`:1134-1163`).
- **Equipo obsoleto:** error si `TbEquipos.FechaObsoleto` < fecha de alta (`:1071-1081`).
- **Nuevo IDEvento solo si cambia el BUI:** `DameIDEvento` = iniciales (`TbBuiIDEvento`, o 3 primeras letras del BUI) + `aamm` + ordinal de 3 cifras = máximo `Right(IDEvento,3)` del prefijo + 1 (`:1096-1112`; `FG.bas:6205-6221, 6225-6242`).
- **Campos fijados en cada edición:** `HORAINICIALEVENTO` = «7:30:00», `TIEMPORESPUESTAEVENTO` = «00:00» y `FechaRegistroModificacion` = `Now()` (`:1146-1147, 1169`).
- **Cambios (`HaHabidoCambios`):** compara 12 valores con los iniciales; un error se interpreta como «hay cambios» (`:58-127`).
- **Cálculo de horas:** `HorasLabTotales` + `HorasExtTotales`, no numérico = 0 (`:832-842`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Carga del evento (con `RIGHT JOIN` a equipos) | `TbEventos`, `TbEquipos` | DAO, SQL concatenado con el `IDEvento` | — | `:667-671` |
| Valores por clave (`Dame`): actividades, planificación, obsolescencia, BUI | `TbActividades`, `TbPlanificacion`, `TbEquipos`, `TbEventos` | SQL concatenado; primero sin comillas y, si falla, con comillas | — | `:782, 848, 1071, 1089`; `FG.bas:300-311` |
| Última actividad | `TbActividades` | SQL concatenado | — | `FG.bas:819-823` |
| Listas de subsistema, equipo y contacto | `TbSubsistemaBui`, `TbEquipos`, `TbEventos` | `RowSource` concatenado | — | `FG.bas:1198-1200, 1214-1217, 6675-6679` |
| Ordinal del nuevo `IDEvento` | `TbEventos` | SQL con `Like 'prefijo*'` | no | `FG.bas:6225-6229` |
| Edición del evento | `TbEventos` | DAO `Edit`/`Update` (puede cambiar la clave `IDEvento`) | no | `:1117-1170` |
| Renombrado de clave | `TbActividades`, `TbEventoAnexos` | SQL `UPDATE` concatenado con `DoCmd.RunSQL` | no | `:1232-1242` |
| Fecha de informe RAC | `TbEventos` | `UPDATE` con literal `#mm/dd/yyyy#` o `Null`, `CurrentDb().Execute` | no | `Evento.cls:1450-1452, 1468-1470` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `Like '…*'` y `Right([IDEvento],3)` (dialecto Jet, `FG.bas:6225-6228`); literal de fecha `#mm/dd/yyyy#` (`Evento.cls:1450`); comparación `"Verdadero"` del booleano `Franqueado` (`:730`); `Now()` y `Nz`; PK textual modificada en DAO (`:1128`) sin transacción; `RIGHT JOIN` (`:668`). Con PostgreSQL vinculado la comparación de textos pasa a distinguir mayúsculas.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `strURLAnexo` se lee en `ImagenAnexo_Click` (`:950-960`) pero ninguna línea le asigna valor: el clic daría siempre «no es alcanzable».
- `ComandoAyudaFormulario` declara `OnClick` (`Form_FormEventoEdicion.form.txt:1444`) y no existe handler.
- `Edicion` sobrescribe `HORAINICIALEVENTO` y `TIEMPORESPUESTAEVENTO` con literales en cada guardado (`:1146-1147`), incluso si ya tenían otro valor.
- `ReNombrarIDEvento` solo actualiza `TbActividades` y `TbEventoAnexos` (`:1232-1242`); `TbPlanificacion` y otras tablas que referencien el evento no se tocan. Además `TbEventos` se actualiza antes (`:1170`) y sin transacción: un fallo posterior dejaría el evento renombrado y los hijos sin renombrar.
- El `IDEvento` solo se regenera si cambia el BUI (`:1096`); un cambio de fecha de alta a otro mes no lo recalcula.
- En modo escritura `Registrar` no lanza `FechaInformeRACambiada` (solo en `:507`), por lo que `FormInformesParaRAC` no refrescaría su lista (`Form_FormInformesParaRAC.cls:307-312`). Tampoco `ComandoEditar_Click` refresca `ListaEventos` (sí lo hace `cmdSalir`, `:157-161`).
- En solo lectura, `cmdSalir` solo guarda si `strEscritura = "Sí"` (`:133`): un cambio de fecha RAC se pierde al salir sin pulsar «Registrar».
- `ComboBUI_AfterUpdate` y `ComboSubsistema_Change` comparan con `strValorAnteriorBUI`/`strValorAnteriorSubSistema`, que no se inicializan en la carga (`:353, 432`): reelegir el mismo BUI inicial vaciaría subsistema y equipo.
- `Me.ComboSubsistema.Selected(1)` (`:380-381`): `Selected` es de `ListBox` (podría dar error 438) y el índice 1 sería el segundo elemento; con un único subsistema (`ListCount-1 = 0`) no se autoselecciona.
- `ComandoQuitarFechaInforme` solo se habilita si había fecha al abrir (`:902-906`).
- `EstablecerDatos` calcula `strEventoSinNada` y no lo usa (`:782-792`); `CausaFin` está bloqueado y sin valor.
- `Dame` activa `On Error Resume Next` (`FG.bas:303`) y no lo restablece: errores posteriores del método quedan silenciados.
- Mensaje del error de `ComandoEditar_Click` habla de «Eliminar» (`:230`); `Contacto` lleva `Tag=OBLIGATORIO` (`Form_FormEventoEdicion.form.txt:2443`) pero `Edicion` no lo exige.
- `"Verdadero"` depende del idioma de Access (`:730`); `.Fields("FECHAALTAEVENTO") = strFechaAlta` asigna texto a un campo fecha (`:1145`).
- Lista de técnicos: el valor se asigna en `EstablecerDatos` (`:714-716`) antes de recargar la lista (`:612`, `FG.bas:7238`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: operación única y transaccional que edite el evento y renombre todas sus dependencias (o clave sustituta que evite renombrar); ordinal de `IDEvento` generado por secuencia de base de datos; aviso explícito al salir con la fecha RAC modificada; recuperar el anexo desde la entidad en lugar de variable de módulo.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un evento facturado o franqueado se abre con «(Sólo lectura)» en el título y solo se puede cambiar la fecha de informe RAC (`ComandoEditar`, `FechaEnInformeRAC`).
2. Guardar sin BUI, subsistema, equipo, tipo, técnico, originador, descripción, o con fecha o criticidad inválidas, muestra el mensaje de la regla y no escribe.
3. Elegir un equipo cuya fecha de obsolescencia es anterior a la fecha de alta muestra el error de equipo obsoleto.
4. Cambiar el BUI de un evento y guardar le asigna un `IDEvento` con el prefijo y mes de la fecha de alta, y sus actividades y anexos pasan a ese identificador.
5. Cambiar el BUI vacía subsistema y equipo y recarga sus listas y la de contactos (`ComboBUI`).
6. Vaciar `PMPR`, `Contacto`, `Notas` o la fecha RAC los deja en nulo en el registro.
7. En lectura, registrar una fecha RAC nueva o quitarla actualiza `TbEventos` y notifica a `FormInformesParaRAC`.
8. Añadir o eliminar un equipo de medida actualiza el rótulo con el nuevo número (`ComandoEquiposDeMedida`).
9. Salir con cambios en modo escritura pregunta si guardar y, si falla el guardado, no cierra.

## Pendientes

- `Constructor.getEvento` y `getTecnicos` (`Constructor.bas:2242`, `Constructor.bas:4249`): significado exacto de los parámetros de `RellenarComboTecnicos` (`:612`).
- `Form_FormEventoEqMedidaEdicion`, `FormEquipoAlta`, `FormEventoSLA`, `FormAnexos` y `FormPlanificacionCierre`: argumentos y efectos (`:259, 289, 319, 200, 937`).
- Referencias dinámicas `Forms("FormEventoGestion").Controls("ListaEventos")` (`:158`) y `Forms("FormEventoEqMedidaEdicion")` (`:261`).
- Quién restringe el acceso a este formulario por perfil; el formulario no consulta permisos.
- Índices y tipos de `TbEventos` (clave `IDEvento` textual, `Franqueado` booleano) en el ERD de origen.
