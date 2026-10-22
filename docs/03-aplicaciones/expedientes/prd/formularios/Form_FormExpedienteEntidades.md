# Form_FormExpedienteEntidades

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormExpedienteEntidades.cls` (1974 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña Entidades de la ficha de expediente: siete relaciones del expediente con catálogos (PECAL, CPV, lugares de ejecución, comerciales, RAC, responsables) y las anualidades presupuestarias. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); todas las listas son `Value List` rellenadas por código. A diferencia de General y Fechas, **escribe en la base de datos en el momento** cuando el expediente ya existe, y solo en el DTO global `m_ObjExpedienteDTOActivo` durante un alta. Las citas son `Form_FormExpedienteEntidades.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormExpediente` | botón de navegación `NavEntidades` con destino `FormExpedienteEntidades`; antes vuelca General/Fechas con `RellenarDTO` | `Form_FormExpediente.form.txt:830, 848`, `Form_FormExpediente.cls:470-492` |

Abre los maestros `FormPECAL`, `FormRAC`, `FormComercial`, `FormCPV` y `FormLugarEjecucion` y escucha su evento `Alta` con `WithEvents` (`:6-11`, `:1088-1231`). Abre el modal `FormExpedientesResponsables` oculto, lo inicializa con `InitializeEditor` y lo muestra (`:450-453`, `:1442-1445`). Copia textos al portapapeles con `CopiarAlPortapapeles` (`FUNCIONES UTILES.bas:5699`; uso en `:496, 530`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `PECALES` | combo 2 col. (ID oculto) | `AfterUpdate` | Con valor, lanza `ComandoAltaPECAL_Click` | `:1943-1947`, `.form.txt:1508-1514` |
| `ComandoAltaPECAL` | botón (`EJECUTIVO`) | `OnClick` | Alta vía `ExpedienteEntidadesAlta_Agregar("pecal")` | `:289-326` |
| `ListaPECALES` / `ComandoEliminarPECAL` | lista / botón (`EJECUTIVO`) | `OnClick` | Baja de la PECAL seleccionada | `:855-931` |
| `ComandoNuevaPECAL` | botón (sin etiqueta) | `OnClick` | Reabre `FormPECAL` vacío y se suscribe a `Alta` | `:1088-1115` |
| `CPVs` / `ComandoAltaCPV` | combo / botón (`EJECUTIVO`) | `AfterUpdate` / `OnClick` | Alta de CPV (lógica propia, no usa el helper) | `:1239-1243`, `:175-245` |
| `ListaCPVs` / `ComandoEliminarCPV` | lista / botón (`EJECUTIVO`) | `OnClick` | Baja del CPV | `:698-773` |
| `ComandoCopiarCPV`, `ComandoCopiarLugar` | botones (sin etiqueta) | `OnClick` | Copian la descripción (columna 1) de la fila seleccionada | `:482-548` |
| `ComandoNuevoCPV` | botón (sin etiqueta) | `OnClick` | Reabre `FormCPV` | `:1175-1202` |
| `LugaresEjecucion` / `ComandoAltaLugarEjecucion` | combo / botón (`EJECUTIVO`) | `AfterUpdate` / `OnClick` | Alta vía helper (`"lugar"`) | `:1255-1259`, `:250-287` |
| `ListaLugaresEjecucion` / `ComandoEliminarLugarEjecucion` | lista / botón (`EJECUTIVO`) | `OnClick` | Baja del lugar | `:778-853` |
| `ComandoNuevoLugar` | botón (sin etiqueta) | `OnClick` | Reabre `FormLugarEjecucion` | `:1204-1231` |
| `Comerciales` / `ComandoAltaComercial` | combo / botón (`EJECUTIVO`) | `AfterUpdate` / `OnClick` | Alta de comercial (lógica propia) | `:1233-1237`, `:102-173` |
| `ListaComerciales` / `ComandoEliminarComercial` | lista / botón (`EJECUTIVO`) | `OnClick` | Baja del comercial | `:621-696` |
| `ComandoNuevoComercial` | botón (sin etiqueta) | `OnClick` | Reabre `FormComercial` | `:1146-1173` |
| `RACs` / `ComandoAltaRACs` | combo / botón (`EJECUTIVO`) | `AfterUpdate` / `OnClick` | Alta vía helper (`"rac"`) | `:1949-1953`, `:328-365` |
| `ListaRACs` / `ComandoEliminarRACs` | lista / botón (`EJECUTIVO`) | `OnClick` | Baja del RAC | `:933-1008` |
| `ComandoNuevoRAC` | botón (sin etiqueta) | `OnClick` | Reabre `FormRAC` | `:1117-1144` |
| `Responsables` | combo 2 col., `ColumnHeads` | `AfterUpdate` (también cableado en `Form_Load`) | Decide modo `create`/`edit` y abre el modal | `:1384-1457`, `:1331`, `.form.txt:2036, 2084` |
| `ComandoAltaReponsable` («+») | botón (`EJECUTIVO`) | `OnClick` | Con combo informado reenvía a `Responsables_AfterUpdate`; vacío, despliega el combo | `:367-382` |
| `ListaResponsables` | lista 5 col., `ColumnHeads` | `OnClick` → no-op retirado | Columnas ID · Nombre · Prev · JP · Aviso | `:1247-1253, 1585-1615`, `.form.txt:2197, 2215, 2217` |
| `ComandoEditarResponsable` | botón (`EJECUTIVO`) | `OnClick` | Lee la fila guardada y abre el modal en modo `edit` | `:384-462` |
| `ComandoEliminarReponsable` | botón (`EJECUTIVO`) | `OnClick` | Baja del responsable | `:1010-1085` |
| `AÑO`, `BIIVA`, `BIIPSI`, `BIIGIC`, `BIEXENTA` | textos | — | Entrada de una anualidad | `.form.txt:2740-3049` |
| `ComandoAltaAnualidad` / `ComandoEliminarAnualidad` | botones (`EJECUTIVO`) | `OnClick` | Alta y baja de anualidad | `:16-100`, `:550-619` |
| `ListaAnualidad` | lista 2 col., `ColumnHeads` | — | Año · Presupuesto con formato `#,##0.00 €` | `:1616-1639`, `.form.txt:2692` |

Los `RowSource` de diseño del `.form.txt` contienen datos de ejemplo (p. ej. `.form.txt:605, 742, 2059, 2214, 2708`) que se sustituyen en ejecución por `EstablecerCombos` y `RellenarListas` (`:1743-1851`, `:1640-1700`).

## Flujos

### Apertura (`Form_Load`, `:1321-1361`)

1. `AllowEdits = False`; cablea en ejecución `Responsables.AfterUpdate` y `OnActivate` (`:1329-1337`).
2. `AllowEdits = True` solo si `Form_FormExpediente.m_AbiertoParaEditar` y `EsAdministrador = Sí` (`:1339-1345`).
3. `EstablecerDatos`: sin `Expediente` en el DTO, error «El formulario se ha abierto con parámetros insuficientes» (`:1708-1711`). Habilita los controles con `Tag = "EJECUTIVO"` con la misma condición del paso 2 y los deshabilita en otro caso (`:1712-1724`).
4. `EstablecerCombos` rellena los seis combos desde las cachés de `m_ObjEntorno` (`:1753-1843`); `RellenarListas` pinta las siete listas desde las colecciones del DTO (`:1730`, `:1648-1693`).

### Alta de una entidad (patrón común)

1. Combo vacío: sale sin hacer nada (p. ej. `:115-121`).
2. Valida que el ID exista en el catálogo: PECAL, RAC y lugares mediante `ExpedienteEntidadesAlta_Agregar`, que además exige ID numérico positivo (`modExpedienteEntidadesAltaHelper.bas:28-64`); comerciales y CPV con `constructor.getComercial`/`getCPV` en el propio formulario (`:122-129, 195-202`).
3. **Expediente sin `IDExpediente` (alta):** añade al diccionario del DTO si no estaba (clave = ID) y pinta la lista (`:131-143`; `modExpedienteEntidadesAltaHelper.bas:66-91`). La persistencia la hace `Registrar` al guardar ([Form_FormExpediente](Form_FormExpediente.md); `ExpedienteOperaciones.cls:402-498`).
4. **Expediente existente:** `ExpedienteOperaciones.Registrar<Entidad>` inserta la relación si no existe y, si devuelve «1», refresca la colección del DTO desde `Expediente.<Colección>` y repinta (`:145-158`; `modExpedienteEntidadesAltaHelper.bas:94-120`).
5. Selecciona la última fila y vacía el combo (`:160-161`).

### Baja de una entidad (patrón común)

1. Sin fila seleccionada, sale (p. ej. `:633-639`); recarga la entidad del catálogo y, si no existe, avisa «Seleccione un … de la lista» (`:640-647`).
2. Alta: quita la clave del DTO (lo deja en `Nothing` si queda vacío) y la fila de la lista (`:648-666`).
3. Existente: `ExpedienteOperaciones.Eliminar<Entidad>` ejecuta un `DELETE` y refresca DTO y lista (`:667-679`). **No pide confirmación** en ninguna baja.

### Anualidades (`:16-100`, `:550-619`)

1. Exige `AÑO` numérico («Se ha de incluir el año») y al menos uno de los cuatro importes numérico («Se ha de incluir alguno de los cuatro importes») (`:38-45`).
2. Alta: guarda en `ColAnualidades` con clave el año **el texto `Presupuesto`**, no el objeto (`:61-63`); un año repetido se ignora.
3. Existente: `RegistrarAnualidad` borra la fila del mismo año si existe e inserta la nueva (sustitución), con IVA 21, IPSI 8 e IGIC 7 por defecto (`ExpedienteOperaciones.cls:1251-1278`).
4. Tras el alta, `AÑO` avanza en uno y se vacían los importes (`:84-88`). La baja exige además `EsAdministrador = Sí` («Operacion no autorizada», `:569-572`).

### Responsables (`:367-462`, `:1384-1457`)

1. Elegir un usuario en `Responsables` decide el modo: con expediente guardado y fila existente en `TbExpedientesResponsables` (`constructor.getResponsableByUsuario`, `constructor.bas:4279-4319`), `edit` con sus marcas actuales; si no, `create` (`:1413-1431`).
2. Busca el nombre con `constructor.getUsuario` e ignora su error (`:1434-1439`); abre el modal oculto, `InitializeEditor`, lo muestra y vacía el combo (`:1442-1448`).
3. `ComandoEditarResponsable` exige fila seleccionada y expediente guardado («Expediente no persistido; solo edit de filas persistidas», `:402-417`), lee la fila por `(IDExpediente, IdUsuario)` con SQL concatenado (`:426-440`) y abre el modal en `edit` (`:450-453`).
4. El modal guarda con `ProcesarEdit` y repinta la lista a través de `Forms("FormExpediente").FrmDetalle.Form.RellenarListaResponsables` (`Form_FormExpedientesResponsables.cls:133, 153-159`).

### Errores

Cada handler acumula el texto en `m_Error`; los errores no controlados (distintos de 1000) envían `CorreoAlAdministrador` y todos muestran `MsgBox` crítico (p. ej. `:93-99`). `Form_Activate` solo envía correo, sin aviso (`:1377-1381`).

## Reglas, validaciones y cálculos

- **Permisos:** `AllowEdits` y los controles `EJECUTIVO` dependen de `m_AbiertoParaEditar` y `EsAdministrador` (`:1339-1345, 1712-1724`). Los botones «Nuevo …» y «Copiar …» no llevan esa etiqueta (`.form.txt:2235-2612, 3398, 3478`).
- **Duplicados:** en alta, el diccionario del DTO ignora el ID repetido (`:138, 210`); en edición, `Registrar<Entidad>` solo inserta si no existe la pareja expediente-entidad (`ExpedienteOperaciones.cls:1302-1312, 1369-1379`).
- **Presupuesto de anualidad** = BIIVA + BIIPSI + BIIGIC + BIEXENTA, cada no numérico como 0 (`classes/ExpedienteAnualidad.cls:69-80`).
- **Responsables:** el combo ofrece `m_ObjEntorno.ColUsuariosTecnicos` (`:1831`; `classes/Entorno.cls:1742-1752`); la lista muestra «No» si `EsPreventa` está vacío (`:1600-1604`). Las marcas Sí/No y la regla «al menos un JP y un aviso» las valida el modal (`modResponsablesEdit.bas:172-207`; `Form_FormExpedientesResponsables.cls:125-131`), no esta pestaña.
- **Derivados:** altas de PECAL, RAC y lugar llaman a `SincronizarDerivados`; altas y bajas de comerciales, bajas de PECAL/RAC/lugar/responsable llaman a `RegistrarExpEntidades`; la baja de PECAL recalcula `ActualizarCadenaPECAL`; la de responsable, la cadena de gestión de riesgos (`ExpedienteOperaciones.cls:1315, 1416, 1452, 1488, 2453, 2501, 2527-2529, 2554, 2579-2586`).
- **Lugares:** el `;` del texto se sustituye por `:` al pintar la lista (`:1497`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Alta de relación | `TbExpedientesPECAL`, `TbExpedientesRACS`, `TbExpedientesLugaresEjecucion`, `TbExpedientesComerciales`, `TbExpedientesCPVs` | DAO `AddNew` tras `SELECT` concatenado | No | `ExpedienteOperaciones.cls:1284-1490` |
| Alta de anualidad | `TbExpedientesAnualidades` | DAO `Delete` + `AddNew` | No | `ExpedienteOperaciones.cls:1255-1277` |
| Bajas | las seis tablas anteriores y `TbExpedientesResponsables` | SQL dinámico `DELETE` con `Execute` | No | `ExpedienteOperaciones.cls:2426-2577` |
| Último cambio | vía `RegistrarUltimoCambio` | — (no trazado) | No | `ExpedienteOperaciones.cls:2430`; `FUNCIONES UTILES.bas:189` |
| Lectura de responsable | `TbExpedientesResponsables` | SQL dinámico concatenado | — | `:426-427`; `constructor.bas:4293-4296` |
| Catálogos | cachés de `Entorno` (`PECALES`, `LugaresEjecucion`, `Comerciales`, `CPVs`, `RACs`) | no trazado aquí | — | `classes/Entorno.cls:778, 865, 956, 987, 1045` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `DELETE * FROM` (sintaxis Jet) e identificadores concatenados sin parámetros (`ExpedienteOperaciones.cls:2426-2427`); cada alta y baja es una escritura suelta sin transacción, con efectos derivados posteriores; los identificadores de las tablas de relación se calculan en VBA (`…Calculado`, p. ej. `ExpedienteOperaciones.cls:1297, 1364`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- En alta, `ColAnualidades` guarda el texto `Presupuesto` (`:62`), pero `Registrar` y `RellenarListaAnualidades` hacen `Set … = m_Col(m_ID)` esperando un `ExpedienteAnualidad` (`ExpedienteOperaciones.cls:408`; `:1627`): guardar un alta con anualidades fallaría y revertiría la transacción.
- `EstablecerDatos` llama a `RellenarListas` sin recoger el error (`:1730`): un fallo al pintar listas (como el anterior) queda oculto.
- Alta de anualidad: en un alta el año repetido se ignora (`:61`); en un expediente existente se sustituye la fila (`ExpedienteOperaciones.cls:1258-1261`). `RegistrarAnualidad` no tiene manejador de error ni `RegistrarUltimoCambio` (`ExpedienteOperaciones.cls:1238-1282`).
- El modal de responsables llama a `ProcesarEdit` con `p_db = Nothing` (`Form_FormExpedientesResponsables.cls:133`), y el helper solo escribe con base de datos inyectada (`modResponsablesEdit.bas:241, 274-278, 321`): en un expediente existente, altas y cambios de responsable quedarían solo en el DTO, y `Registrar` no persiste hijos en edición (`ExpedienteOperaciones.cls:402`).
- `g_NecesitaRefrescarResponsables` solo se lee aquí (`:1372`; declarada en `Variables Globales.bas:11`): no se ha encontrado quién la ponga a `True`, así que `Form_Activate` no repinta nunca.
- El combo `Responsables` tiene `ColumnHeads` (`.form.txt:2036`) pero `EstablecerCombos` no añade fila de cabecera (`:1830-1843`): el primer técnico se mostraría como cabecera y no sería seleccionable.
- La baja de responsable no valida la regla «al menos un JP y un aviso» (`ExpedienteOperaciones.cls:2562-2592`), que el modal sí aplica al guardar.
- CPV no actualiza derivados ni la caché de entidades ni en alta ni en baja (`ExpedienteOperaciones.cls:1351-1382, 2461-2482`), a diferencia del resto.
- Solo la baja de anualidad exige administrador en código (`:569`); el resto depende de que el botón esté deshabilitado. Los botones «Nuevo …» se pueden pulsar en solo lectura (sin `EJECUTIVO`).
- Restos: `ComandoCambiarEnvios_Click`, `ComandoCambiarJP_Click` sin control en el `.form.txt` (`:464-480`); `ListaResponsables_Click` no-op aunque el `.form.txt` mantiene `OnClick` (`:1247-1253`, `.form.txt:2217`); `CambiarLineaListaResponsables` sin llamadas (`:1954-1974`); `RellenarListas` fija una cabecera de 4 columnas que se sobrescribe con 5 (`:1653, 1593`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: confirmación antes de cada baja; un único modelo de alta (borrador en servidor) en lugar de bifurcar DTO/BD; escrituras de relación y derivados en una transacción; regla de responsables también en la baja.

## Criterios de aceptación

Propuestos, no aprobados:

1. En solo lectura, o sin ser administrador, no se puede añadir ni quitar ninguna entidad ni anualidad.
2. Añadir dos veces la misma entidad deja una sola relación, tanto en alta como en edición.
3. Añadir una anualidad sin año, o sin ningún importe, muestra el mensaje correspondiente y no escribe.
4. En un alta, las entidades y anualidades añadidas se guardan con el expediente en la misma transacción.
5. En un expediente existente, añadir o quitar una entidad se refleja en la lista y en la base de datos sin pulsar guardar.
6. Elegir un responsable ya asignado abre su edición con las marcas actuales; uno nuevo, el alta.

## Procedimientos auxiliares

Citas `Form_FormExpedienteEntidades.cls:línea`.

- **Alta rápida de un catálogo desde la pestaña** (`m_FormComercial_Alta`, `m_FormCPV_Alta`, `m_FormLugar_Alta`, `m_FormPECAL_Alta`, `m_FormRAC_Alta`, `:1261-1290`, `:1291`, `:1853`, `:1883`, `:1913`): al recibir el evento `Alta` del formulario del catálogo, cada uno invalida la caché de ese catálogo en `m_ObjEntorno`, vuelve a llamar a `EstablecerCombos`, preselecciona en el combo el elemento recién creado y pulsa por código el botón de alta de ese catálogo (`ComandoAlta*_Click`). Si el elemento activo ya no existe, solo se recarga el combo.
- **Listas del expediente** (`RellenarListaComerciales`, `RellenarListaCPVs`, `RellenarListaRACs`, `:1511-1584`): vacían el `RowSource` y añaden una fila `id;nombre` por cada elemento de la colección correspondiente del DTO activo (p. ej. `m_ObjExpedienteDTOActivo.ColComerciales`). Con colección vacía dejan la lista sin filas.

## Pendientes

- Si `Expediente.<Colección>` relee la base tras cada alta o baja o devuelve una caché (`:155, 677`).
- Contenido de `SincronizarDerivados`, `RegistrarExpEntidades` por ámbito y `RegistrarUltimoCambio` (`ExpedienteOperaciones.cls:763`; `FUNCIONES UTILES.bas:189`).
- Comportamiento real de `Selected(0)` y de la cabecera en listas con `ColumnHeads` (`:1475, 1607`).
- Origen de `getUsuariosTecnicos` y de las cachés de catálogo de `Entorno`.
- Comprobación en ejecución de los defectos de anualidades en alta y de persistencia de responsables (fichas de `Form_FormExpedientesResponsables` y del flujo de guardado).
