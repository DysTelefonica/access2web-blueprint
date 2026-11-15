# Form_FormNCProyectoGestion

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCProyectoGestion.cls` (1827 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pantalla «GESTIÓN DE NC DE PROYECTOS» (`.form.txt:39`): buscador y listado de las no conformidades de proyecto, con filtros combinables, y punto de partida de alta, detalle, borrado/habilitación, alta de NC vinculada, correo, historial e informes. Es un formulario **no enlazado** (sin `RecordSource`): la lista `ListaFiltrados` es de tipo «Value List» que el código rellena fila a fila; el filtrado se hace **en memoria** sobre el conjunto de NC que devuelve `constructor`. Las citas son `Form_FormNCProyectoGestion.cls:línea` salvo indicación; `constructor.bas`, `NCProyectoOperaciones.cls` y `NCProyecto.cls` abrevian `modules/constructor.bas`, `classes/NCProyectoOperaciones.cls` y `classes/NCProyecto.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Submenú de NC de proyecto (`Form0BDOpcionesParteProyectos`) | Cierra la instancia abierta, abre el formulario, fija `EstadoNC = "Abiertas"` y, si procede, `ResponsableTelefonica` con el nombre del usuario, y llama a `EstadoNC_AfterUpdate` | `Form_Form0BDOpcionesParteProyectos.cls:127-136` |
| Menú de técnicos (`Form0BDTecnicos`) | `DoCmd.OpenForm` | `Form_Form0BDTecnicos.cls:42` |
| Fichas de AC, AR y acciones de proyecto | Llaman a `ActualizarLineaLista` tras guardar | `Form_FormNCProyectoAC.cls:202-203`, `Form_FormNCProyectoAR.cls:265-266`, `Form_FormNCProyectoAcciones.cls:139-140` |

Abre: `FormNCProyecto` (Alta y Detalle, `:296-299, 553-556`), `FormNCProyectoGeneralConVinculoNC` (`:45-48`), `FormCorreo` con `OpenArgs = Sí` (`:86-89`) y `FormExpedientesBusqueda` (`:251-254, 1579-1582`). Recibe la selección de expediente por el evento `m_FormExpedientes_Seleccionado` (`:10, 1318-1352`). «Salir» cierra sin guardar (`:15-18`).

## Controles y eventos

`.form.txt` declara 73 bloques `Begin`: el formulario (`:4`), 11 bloques de valores por defecto (`:407-479`), cabecera (`:491`), `Detalle` (`:601`), pie (`:4776`) y 58 controles, agrupados así:

| Control(es) | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaFiltrados` | lista de 7 columnas (`IDNoconformidad;Código;Descripción;Exp.;Estado;FApertura;FCierre`) | `OnClick`, `OnDblClick` | Selecciona la NC y activa botones; doble clic equivale a «Detalle» | `:1185-1283`, `.form.txt:615-706` |
| `ComboCodigo`, `ComboJuridica`, `ComboTipo`, `EstadoNC`, `ComboRequiereControlEficacia`, `ComboControlEficaciaRelleno`, `ComboRegistrosCerrados`, `RESPONSABLECALIDAD`, `RESPONSABLETELEFONICA` (9) | listas desplegables de filtro | `AfterUpdate` | Todas ejecutan `ActualizarLista` | `:617-805, 1424-1444, 1533-1553`, `.form.txt:857, 1452, 1682, 1913, 2629, 3040, 3293, 3520, 3784` |
| `Descripcion`, `Notas` | cuadros de texto de filtro | `OnExit` | `ActualizarLista` | `:807-829, 1356-1377`, `.form.txt:2131, 2167` |
| `PalabraClave`, `Google`, `IDExpediente` | cuadros de texto (expediente, búsqueda libre, id oculto de expediente) | `OnKeyDown` en los dos primeros | Intro o Tab: `ComandoBuscarExpediente_Click` (`PalabraClave`) o `ActualizarLista` (`Google`) | `:1173-1183, 1554-1569`, `.form.txt:4023, 4386` |
| 13 botones de limpiar (`ComandoLimpiarCodigo`, `…Exp`, `…Juridica`, `…Tipo`, `ComandoLimiparEstadoNC`, `…Descripcion`, `…Notas`, `ComandoLimpiaoRequiereControlEficacia`, `ComandoLimpiarControlEficaciaRelleno`, `ComandoLimpiarRegistrosCerrados`, `ComandoLimiparRESPONSABLECALIDAD`, `…TELEFONICA`, `…PalabraClave`) | botones | `OnClick` | Ponen el filtro a `Null` y refrescan | `:445-486, 574-613` |
| `ComandoBuscarExpediente`, `ComandoIrABusquedaExpedientes`, `ComandoActualizararExpedientes` («Contador») | botones | `OnClick` en los dos primeros; el tercero no tiene | Resolver o buscar expediente; el tercero no hace nada | `:190-273, 1570-1605`, `.form.txt:4054, 4138, 4202` |
| `ComandoAlta`, `ComandoEditar` («Detalle»), `ComandoInforme`, `ComandoNCProyectoHistorial`, `ComandoEliminar` («Eliminar»/«Habilitar»), `ComandoEnviarCorreoElectronico`, `ComandoAltaVinculadas`, `ComandoListaEnExcel`, `ComandoInformeLista`, `ComandoActualizarLista`, `cmdSalir` (11) | botones | `OnClick` | Ver Flujos | `.form.txt:1108, 1157, 1206, 2567, 4312, 4600, 4670, 4719, 4798, 538, 4845` |
| `Etiqueta86, 88, 90, 92, 93, 98, 99, 100, 101, 102, 111, 114, 122` (13), `lblTitulo1`, `lblVinculado`, `lblEstado` | etiquetas | — | Rótulos; `lblVinculado` («Vinculada con la …») y `lblEstado` (progreso) se muestran por código | `:1194, 1242-1247, 842`, `.form.txt:511, 4647, 4947` |
| Pie `PieDelFormulario` | sección | `OnDblClick` | Abre en HTML la NC seleccionada con acciones | `:1382-1422`, `.form.txt:4780` |

## Flujos

### Apertura (`Form_Open` `:833-862`; `EstablecerDatos` `:864-891`)

Ajusta tamaño, copia el título al `Caption` y oculta `lblEstado` (`:840-842`). Carga las listas desplegables (`EstablecerCombos`, `:1493-1531`; añade «Abiertas» con valor 0, `:1525`). Un técnico queda sin «Alta», «Enviar correo» ni «Historial» (`:873-882`). Cualquier error cierra el formulario (`:861`). **La lista no se rellena al abrir**: lo hace quien lo abre.

### Búsqueda y filtrado (`ActualizarLista` `:1097-1171`; `getNCsFiltrados` `:892-1046`)

1. Desactiva botones, vacía la selección y pone la cabecera de la lista (`:1111-1128`).
2. Origen de datos: si `EstadoNC` es «Abiertas», solo las abiertas (`:912-918`); si no, con `Google` informado, búsqueda por palabra clave (`:920-921`); si no, todas (con «Abiertas» como texto, las totales para abiertas) (`:923-927`).
3. Filtra en memoria: código, expediente, jurídica (contiene), tipo, descripción y notas (contiene, sensible a mayúsculas), requiere control de eficacia, control de eficacia relleno Sí/No, registros cerrados Sí/No (por fecha de cierre), jefe de proyecto, responsable de Calidad y estado (`:947-1026`). Sin ningún filtro devuelve todo (`:940-942, 1048-1071`).
4. Rellena filas con id, código, descripción y expediente (sustituyendo `;` por `:`), estado y fechas (`:1448-1490`); selecciona la primera fila y ejecuta el clic (`:1481-1484`). Con `p_ManteniendoSeleccionado` reselecciona la NC anterior (`:1113-1117, 1145-1156`).

### Selección (`ListaFiltrados_Click`, `:1185-1274`)

Carga la NC (`getNCProyecto`) y habilita «Detalle» e «Informe»; a no técnicos, «Alta NC vinculada», «Enviar correo», «Eliminar» e «Historial» (`:1229-1238`). Muestra «Vinculada con la …» si tiene NC asociada (`:1242-1247`) y rotula «Habilitar» si el estado es «BORRADA», «Eliminar» si no (`:1252-1256`).

### Acciones

- **Alta:** técnico no autorizado (`:546-549`); vacía las NC activas y abre `FormNCProyecto` (`:551-556`).
- **Detalle / doble clic:** fija `m_ObjNCProyectoActiva` y abre `FormNCProyecto` (`:294-299, 1276-1283`).
- **Eliminar / Habilitar:** no técnico (`:325-328`). «Eliminar»: confirmación con «No» por defecto, luego `InputBox` de motivo; «Habilitar»: confirmación (`:342-348`). Con «Sí» ejecuta `NCOp.Eliminar EnumSino.No` o `NCOp.Habilitar` y refresca manteniendo la selección (`:350-371`).
- **Alta NC vinculada:** guarda la NC en `m_ObjNCProyectoActivaPVinculada` y abre `FormNCProyectoGeneralConVinculoNC` (`:34-48`).
- **Correo, historial, informes y Excel:** `FormCorreo` (`:85-89`); historial en HTML (`:165-169`); informe Word de la NC (`:517`) o de toda la lista (`:405-424`); exportación a Excel de la lista filtrada (`:119-127`).
- **Actualizar (`ComandoActualizarLista`):** recalcula y **graba** fecha prevista de cierre, fecha de cierre y estado de todas las NC filtradas y refresca (`:490-495, 1792-1827`).

### Errores

Error 1000 ⇒ aviso con `m_Error`; otro ⇒ `CorreoAlAdministrador` y mensaje crítico (patrón en cada handler, p. ej. `:56-64`).

## Reglas, validaciones y cálculos

- **Rol:** técnico no puede Alta, Alta vinculada ni Eliminar/Habilitar (`:30-33, 325-328, 546-549`); la interfaz además lo deshabilita (`:873-882, 1232-1238`).
- **Estado «Abiertas»** es un valor sintético del desplegable (valor 0), no un estado de NC (`:912-918, 1525`).
- **Palabra clave:** busca en descripción, causa y raíz, acción correctiva y acción realizada con `Like '*…*'` (`constructor.bas:3144-3175`).
- **Eliminar** exige motivo y marca `Borrado = True`; **Habilitar** limpia motivo y `Borrado` (`NCProyectoOperaciones.cls:833-880, 980-1000`).

## Lecturas y escrituras

| Operación | Tabla o fuente | Tipo | Evidencia |
|---|---|---|---|
| Lista de abiertas | `TbNoConformidades` con `TbNCAccionCorrectivas`, `Borrado = False` | SQL dinámico concatenado | `constructor.bas:1874-1900` |
| Lista de totales / totales para abiertas | `TbNoConformidades` | SQL dinámico (`SELECT *`) | `constructor.bas:2001-2030, 2060-2090` |
| Palabra clave | `TbNoConformidades`, acciones correctivas y realizadas | SQL dinámico con `Like` y `*` | `constructor.bas:3144-3175` |
| Una NC / expedientes / desplegables | `getNCProyecto`, `getExpedientesBusqueda`, `EstablecerCombo*` | no verificado | `constructor.bas:422, 2397`, `Funciones Generales.bas:1739-1921, 2255` |
| Borrar / habilitar NC | `TbNoConformidades` (`Borrado`, `MotivoBorrado`) | DAO `Edit/Update`, más `ActualizarDatosNCProyecto` e invalidación de caché | `NCProyectoOperaciones.cls:860-884, 987-1010` |
| Recalculo en «Actualizar» | `TbNoConformidades` (`FPREVCIERRE`, `FECHACIERRE`, estado) | `UPDATE` concatenado, fecha `#mm/dd/yyyy#` | `NCProyecto.cls:1504-1575, 593-630` |
| Excel, Word, HTML | ficheros generados | no verificado | `Funciones Generales.bas:4367, 3978, 1551`, `HTML.bas:784` |

Sin transacciones explícitas.

## Coexistencia

Dialecto Jet/ACE a adaptar: `Like '*…*'` (`constructor.bas:3144-3175`), fechas `#mm/dd/yyyy#` (`NCProyecto.cls:1504-1575`), `SELECT DISTINCT … INNER JOIN` entre paréntesis y `Borrado = False`. El filtrado y la lista en memoria son independientes del motor. Parametrizar todo lo concatenado. Resto, pendiente de `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Con «Abiertas» y palabra clave, el código anula `EstadoNC` pero `m_EnumEstado` sigue en 0, así que se usa el origen de abiertas (`:913-918`) y `Google` se ignora: el bucle de filtrado no lo consulta (`:945-1027`).
- «Eliminar» muestra el `InputBox` del motivo aunque se haya respondido «No» (`:343-350`).
- En «Historial» el `Err.Raise` está comentado (`:161-162`): sin NC seleccionada continúa con la NC vacía.
- «Actualizar» escribe en base de datos para todas las NC filtradas sin confirmación (`:1808-1819`).
- `ComandoLimiparPalabraClave` limpia `Google` pero no `PalabraClave` (`:452-456`), y `ComandoLimpiarExp` hace lo contrario (`:593-598`).
- Los textos «Notas» y «Descripción» se comparan con `InStr` sensible a mayúsculas (`:973, 978`).
- El texto de `ComandoAltaVinculadas` cambia entre «Alta NC Vinculada» y «Alta NC vinculada a …» (`:1201, 1231`).
- `Ctl_ComandoActualizarLista_Click` está vacío (`:751-753`) y `ComandoActualizararExpedientes` no tiene evento (`.form.txt:4202`).
- `ActualizarLineaListaConDatos` y `SeleccionarNC` no tienen llamadas visibles en el formulario (`:1286, 1735`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtrado y paginación en servidor; búsqueda insensible a mayúsculas; separar el recálculo de estado del botón de refresco.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al entrar desde el submenú de proyecto, la lista muestra las NC abiertas y la primera queda seleccionada.
2. Cada filtro (código, expediente, jurídica, tipo, descripción, notas, control de eficacia, cierre, jefe de proyecto, responsable de Calidad, estado) restringe la lista y su botón de limpiar lo anula.
3. Con filtros vacíos se listan todas las NC.
4. Doble clic en una fila abre la ficha de la NC.
5. Un técnico no puede dar de alta, vincular, borrar ni habilitar.
6. Borrar exige motivo y rotula «Habilitar» tras el borrado; habilitar limpia el motivo.
7. «Actualizar» recalcula fechas y estado de las NC listadas.
8. Existen exportación de la lista a Excel e informes Word por NC y por lista.

## Pendientes

- Origen de los desplegables (`EstablecerComboCodigo`, `…Juridicas`, `…Tipo`, `…JP`, `…ResponsablesCalidad`, `…Estado`; `:1499-1521`) y SQL de `getNCProyecto` y `getExpedientesBusqueda`: no leídos.
- Qué control es el rotulado «Palabra clave» frente a «Exp/Act» (`.form.txt:3090, 4359`): no verificado.
- Quién llama a `SeleccionarNC` y `ActualizarLineaListaConDatos`.
