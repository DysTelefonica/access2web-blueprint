# Form_FormExpedientesGestion

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormExpedientesGestion.cls` (3963 líneas, 99 procedimientos) y `.form.txt` · **PRD:** [README](../README.md)

Bandeja principal de expedientes: panel de filtros en la cabecera y tres listas jerárquicas (Acuerdos Marco e individuales, Lotes, Basados) con botonera por lista para ver detalle, alta, alta de derivado, cambio de tipo, eliminación, enlace E2E y exportación. Es un formulario **no enlazado** (sin `RecordSource`; único evento de formulario `OnOpen`, `.form.txt:43`): las listas son `Value List` rellenadas por código. Las citas son `Form_FormExpedientesGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpciones` | `ComandoGestionExpedientes`: si el usuario no es técnico abre este formulario sin `OpenArgs` (modo edición); si lo es, `FormExpedientesGestionTecnica` | `Form_Form0BDOpciones.cls:109-124` |
| `Form_Form0BDOpcionesTecnicos` | Lo cierra si estaba abierto y lo abre con `OpenArgs:="SoloLectura"` | `Form_Form0BDOpcionesTecnicos.cls:66-69` |

| Abre | Desde | Evidencia |
|---|---|---|
| `FormExpediente` (detalle), con el mismo `OpenArgs` | `VerDetalle` | `:814-820` |
| `FormExpedienteAltaTipo` (alta y alta de derivado) | `Alta`, `AltaDerivado` | `:2472-2478`, `:2542-2548` |
| `FormExpedienteCambioTipo` | `CambiarTipo` | `:2671-2677` |
| `FOrmEleccionTipoConsulta` (exportación a Excel) | botones de exportación e informe unificado | `:1113-1116`, `:1207-1210` |
| `FormTareas` | imágenes de tareas | `:1823`, `:1851` |
| Ayuda (`AbrirAyuda`, `FUNCIONES UTILES.bas:3076`) | `ComandoAyuda` | `:2564` |

Escucha con `WithEvents` (`:6-8`): `Alta` de `FormExpedienteAltaTipo` (`Form_FormExpedienteAltaTipo.cls:9, 350`), `CambioRealizado` de `FormExpedienteCambioTipo` (`Form_FormExpedienteCambioTipo.cls:6, 165`) y `AltaExpediente`/`ExpedienteEditado` de `FormExpediente`. Declara `ExpSeleccionado` (`:13`), que nunca se lanza, y `ExpEliminado` (`:14`, lanzado en `:2777`), sin oyentes en `src`. Salida: `cmdSalir` cierra sin guardar (`:200-205`).

Estado compartido (globales de `Variables Globales.bas`): `m_ObjExpedienteActivo`, `m_ObjExpedienteDTOActivo` y `m_ObjExpedienteAlInicio`, que `VerDetalle` rellena antes de abrir el detalle (`:802-812`); `m_DatosEnMemoria`, que sale de `TempVars("DatosEnMemoria")` (`Variables Globales.bas:376-380`); `m_EnumPostAgedoCombo`, `m_EnumAmbitoCombo` y `m_NombreListaParaInforme` (`Variables Globales.bas:271-276`).

## Permisos

- `OpenArgs = "Sololectura"` ⇒ `m_AbiertoParaEditar = False`; cualquier otro valor ⇒ `True` (`:2910-2914`).
- Alta, alta de derivado, cambio de tipo y eliminación solo se habilitan con `m_AbiertoParaEditar` **y** `EsAdministrador = Sí` (`:2143-2155`, `:1992-2012`, `:3034-3048`, `:2348-2359`). `EsAdministrador` sale de `Usuario.EsAdministradorCalculado` y es excluyente con calidad y técnico (`Variables Globales.bas:383-406`).
- Los botones E2E exigen administrador, edición y fila seleccionada (`E2EEnlacesUIHelper.bas:5-12`; uso en `:2050-2053`, `:3030-3033`, `:2426-2429`).
- `lblUltimaModificacion` solo es visible para el administrador en edición (`:55-67`); `PintarTareas` solo se ejecuta en edición (`:2915-2920`).
- «Detalle» y la exportación quedan disponibles para cualquier perfil con fila seleccionada (`:1991`, `:2166-2167`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `PalabraClave` | texto | `KeyDown` | Con Intro lanza `ComandoPalabraClave_Click` | `:3678-3694`, `.form.txt:615` |
| `ComandoPalabraClave` | botón | `OnClick` | `Filtrar` | `:1548-1569` |
| `IDExpediente`, `CodExp`, `Estado`, `Suministrador`, `Comercial`, `JP`, `RAC`, `IDResponsableCalidad`, `IDResponsableSeguridad`, `Ambito`, `EnlazadoE2E` | combos | `AfterUpdate` | Todos llaman a `Filtrar` con reloj de arena | `:832-853`, `:1699-1813`, `:1871-1916`, `:3696-3718`, `:152-198` |
| `MarcoPecal` (Con 1 · Sin 2 · Todos 0), `MarcoSeguridad` (Confidencial 1 · Reservado 2 · Sin clas. 3 · Todos 0) | grupos de opciones | `AfterUpdate` | `Filtrar` | `:3337-3359`, `:3401-3422`, `.form.txt:1950-2026, 2086-2197` |
| `MarcoPostAgedo` (Todos 1 · No 2 · Sólo 3) | grupo de opciones | `AfterUpdate` | Fija `m_EnumPostAgedoCombo` y `Filtrar` | `:3362-3399`, `.form.txt:2769-2846` |
| `ComandoIniciar*` (10) y `ComandoAmbito` | botones «Ver Todos» | `OnClick` | `IniciarCampo <control>`: restablece ese filtro y filtra | `:1012-1102`, `:1229-1546` |
| `ComandoReiniciarParametrosBusqueda` | botón | `OnClick` | Vacía búsqueda y caché, restablece filtros y relanza `ComandoReiniciar` | `:1571-1605` |
| `ComandoReiniciar` («Volver a Buscar») | botón | `OnClick` | Invalida la caché de expedientes y fuerza `Filtrar` | `:3621-3657` |
| `VerEstado` (`Sí;No`) | combo | `AfterUpdate` | Guarda la preferencia (`RegistrarMostrarEstado`) y repinta las listas con columna Estado o fechas | `:3720-3743`, `FUNCIONES UTILES.bas:4086-4137` |
| `ListaAMEIndividual`, `ListaLotes`, `ListaBasados` | listas | `Click`, `DblClick` | Selección, carga de la jerarquía y habilitación de botonera; doble clic = «Detalle» | `:1918-2098`, `:2958-3144`, `:2276-2453`, `:2951-2956` |
| `ComandoVerDetalle{AMoC,Lote,Basado}` | botones | `OnClick` | `VerDetalle` | `:1607-1697`, `:781-830` |
| `ComandoAltaAMoC`, `ComandoAltaLote` | botones | `OnClick` | `Alta` | `:855-951`, `:2455-2486` |
| `ComandoAltaDerivado{AMoC,Lote}` | botones | `OnClick` | `AltaDerivado` | `:888-919`, `:2487-2555` |
| `ComandoCambiarTipo{AMoC,Lote,Basado}` | botones | `OnClick` | `CambiarTipo` | `:953-1011`, `:2584-2686` |
| `ComandoEliminar{AMoC,Lote,Basado}` | botones | `OnClick` | `Eliminar` | `:2691-2897` |
| `ComandoEnlaceE2E_{AM,Lote,Basado}` | botones | `OnClick` | `EnlazarDesvincular`; AM y Lote refiltran después | `:3865-3963`, `:3802-3860` |
| `ComandoExportarExcel{AMs,Lotes,Basados}`, `ComandoInformeUnificado` | botones | `OnClick` | Fijan `m_NombreListaParaInforme` (nombre de la lista o «Todos») y abren `FOrmEleccionTipoConsulta` | `:1104-1227` |
| `ImagenTareasON`, `ImagenTareasOFF`, `lblNumeroTareas` | imágenes, etiqueta | `OnClick` | Abren `FormTareas`; `PintarTareas` alterna las imágenes según el número de tareas | `:1815-1869`, `FUNCIONES UTILES.bas:4063-4075` |
| `ComandoAyuda` (rótulo «Salir» en `.form.txt`) | botón | `OnClick` | Abre la ayuda | `:2557-2582`, `.form.txt:1509` |
| `NAMC`, `NLOTES`, `NBASADOS`, `NTOTAL` | textos | — | Contadores: filas de cada lista menos la cabecera y su suma | `:3660-3676` |
| `IDSeleccionado`, `TipoElementoSeleccionado` | textos | — | ID y tipo («AM», «C», «LOTE», «BASADO») de la última selección | `:2013-2018`, `:3051-3052`, `:2344-2345` |
| `lblListaAM`, `lblListaLotes`, `lblListaBasados` | etiquetas | — | Títulos de lista, cambiados según la selección | `:3606-3619`, `:2027, 2038, 2342, 2397, 3055, 3058` |

Los `RowSource` de diseño de los combos y listas del `.form.txt` (p. ej. `.form.txt:2339, 2489, 2551, 2620, 2935, 4875, 5260`) contienen datos reales y se sobrescriben al abrir con `RellenarCombos` (`:91-150`).

## Flujos

### Apertura (`Form_Open`, `:2898-2947`)

1. `Ajustar`, `AllowEdits = False`, título desde `m_TituloFormulario` y modo de edición según `OpenArgs` (`:2907-2914`).
2. En edición, `PintarTareas` (`:2915-2920`; `FUNCIONES UTILES.bas:3974`).
3. `EstablecerDatos`: etiqueta de última modificación, `VerEstado` desde la preferencia del usuario, `RellenarCombos` y `setExpBusqueda` (`:52-89`). `RellenarCombos` carga estado, contratistas, nº de expediente, comerciales, jefes de proyecto, RAC, ID y responsables, pone ambos responsables a `0` (Todos) y `EnlazadoE2E` a «Todos» (`:94-142`).
4. `setExpBusqueda` con búsqueda vacía fija los valores por defecto: estado `0`, `MarcoPecal = 0`, `MarcoSeguridad = 0`, **`MarcoPostAgedo = 3` (Sólo)**, ámbito según `m_EnumAmbitoCombo` y el resto «Todos» (`FUNCIONES UTILES.bas:1471-1500`).
5. `Filtrar` (`:2927`). Cualquier error envía correo al administrador, muestra el mensaje y **cierra** el formulario (`:2939-2946`).

### Filtrado (`Filtrar`, `:652-780`)

1. Vacía la selección e invalida la caché de marcadores E2E (`:671-678`).
2. `getExpBusqueda` construye el DTO desde los controles (`FUNCIONES UTILES.bas:1145-1212`); si `HaHabidoCambiosBusqueda` frente a la última búsqueda devuelve No, sale sin repintar (`:685-694`; `FUNCIONES UTILES.bas:1297-1391`).
3. `ExpedienteSearch_Ejecutar` lanza cuatro búsquedas (AM, individuales, lotes, basados), fusiona AM e individuales y agrupa en `AM`/`Lotes`/`Basados` (`:715-732`; `modExpedienteSearch.bas:14-105, 141-193`). Un aviso devuelto se muestra con `MsgBox` (`:733-735`).
4. Pinta las tres listas (`:736-747`), deselecciona, actualiza contadores y botonera y habilita la exportación de las listas con filas (`:750-770`).

### Selección en las listas

- **AM o individual** (`:1918-2086`): carga `ExpedienteCompleto` con `constructor.getExpedienteCompleto` (`:1958-1975`), habilita botonera (alta derivada deshabilitada si es individual, `:1996-2000`) y sustituye Lotes y Basados por los hijos del AM (`getColLotes`/`getColBasados`, `:2027-2045`).
- **Lote** (`:2958-3135`): carga el lote, pinta su AM padre en la lista superior y sus basados en la inferior (`:3062-3105`).
- **Basado** (`:2276-2453`): pinta su lote padre y su AM padre (`:2364-2412`).
- Las tres refrescan el rótulo del botón E2E según `EstaVinculado`, en silencio si falla (`:2066-2070`, `:3762-3794`).

### Altas, cambio de tipo y eliminación

- **Alta:** AM/individual sin argumento (la línea con tipos está comentada, `:2461-2463`); lote con `OpenArgs = Lote` (`:2466`). **Alta derivada:** `<ID padre>|Lote;BasadoDeAM` desde un AM o `<ID padre>|BasadoDeLote` desde un lote (`:2531`, `:2535`). Al recibir `Alta`, recarga el AM o lote seleccionado si es el padre, rehace los combos, fuerza `Filtrar` y reselecciona (`:3148-3221`).
- **Cambio de tipo:** rechaza AM o lote con derivados («El AM tiene derivados…», «El Lote tiene derivados…», `:2656-2667`); al recibir `CambioRealizado` invalida la caché y refiltra (`:3223-3248`).
- **Eliminación:** confirma con «¿Desea realmente borrar el expediente seleccionado?» (`:2764-2768`); `ExpedienteOperaciones.Eliminar` (`:2770-2778`), recarga del combo de ID, refiltrado y reselección del padre (`:2779-2833`).
- **Edición desde el detalle:** `ExpedienteEditado` y `AltaExpediente` fuerzan `Filtrar`; el primero reselecciona AM y lote (`:3250-3333`).

### Errores

Patrón común: `m_Error` acumulado, `CorreoAlAdministrador` salvo el error interno 1000 y `MsgBox` crítico (p. ej. `:876-885`). `Filtrar`, `Refiltrar` y los `Rellenar*` reponen `AllowEdits = True` al fallar (`:774-778`, `:227-231`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Filtro por nº de expediente e ID: igualdad exacta | `ColPasanCriterioBusqueda` | `FUNCIONES UTILES.bas:913-916, 959-962` |
| Comercial, jefe de proyecto, contratista y RAC: subcadena (`InStr`) sobre cadenas concatenadas | idem | `FUNCIONES UTILES.bas:918-921, 964-973, 994-997` |
| Palabra clave: subcadena de título o nemotécnico, o igualdad con nº de expediente o ID | idem | `FUNCIONES UTILES.bas:975-982` |
| Estado: igualdad con `ESTADOCalculadoTexto`; un error de cálculo aborta toda la búsqueda | idem | `FUNCIONES UTILES.bas:923-938` |
| «Sin Clasificación» incluye expedientes sin clasificación | idem | `FUNCIONES UTILES.bas:943-950` |
| PECAL «Sí» exige `CadenaPecal` distinta de vacío y «N/A»; «No», lo contrario | idem | `FUNCIONES UTILES.bas:984-993` |
| Responsables: `0`/«Todos» no filtran | idem | `FUNCIONES UTILES.bas:999-1010` |
| Post AGEDO «No» ⇒ `POSTAGEDO = "No"`; «Sólo» ⇒ `"Sí"` | idem | `FUNCIONES UTILES.bas:1028-1040` |
| E2E «Sí»/«No» según enlace activo en `TbE2EEnlaces` | `EnlazadoE2E_PasaCriterioExpediente` | `FUNCIONES UTILES.bas:1046-1049`, `E2EEnlacesUIHelper.bas:66-101` |
| Un DTO con error previo aborta la búsqueda: sin resultados parciales | idem | `FUNCIONES UTILES.bas:887-892` |
| Estado «NoAplica» se muestra como «--» | `Rellenar*` | `:283-285` |
| Nemotécnico del lote: solo el texto anterior al primer «_» | `RellenarLote*` | `:419-423` |
| Columna «Cambio» = «Sí» si `HashActual` difiere de `HashUltimaExportacion` | `MarcadorCambioE2E` | `:39-50`, `E2EEnlacesUIHelper.bas:171-179` |
| No se cambia de tipo un AM o lote con derivados | `CambiarTipo` | `:2656-2667` |
| No se elimina con derivados, DPD, gestiones de riesgo, usuarios HPS o NC | `MotivoEliminarNoOK` | `ExpedienteOperaciones.cls:2188-2225` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Carga de expedientes | `TbExpedientes` LEFT JOIN `TbExpedientesConEntidades` | DAO, SQL dinámico; caché `ColExpedientesCompletos` | — | `constructor.bas:1244-1264`, `Entorno.cls:1422-1443` |
| Marcadores «Cambio» | `TbExpedientes` (hashes) | DAO, una consulta por repintado | — | `E2EEnlacesUIHelper.bas:171-175` |
| Filtro E2E y rótulo del botón | `TbE2EEnlaces` | DAO, una consulta por expediente | — | `E2EEnlacesService.bas:288-310`, `:3782` |
| Preferencia «Mostrar estado» | `TbConfMostrarEstado` | DAO `AddNew`/`Edit` | — | `FUNCIONES UTILES.bas:4123-4137` |
| Eliminación | `TbExpedientesConEntidades`, `TbExpedientesSuministradores`, hitos, anualidades, modificados, responsables, comerciales, CPV, lugares, PECAL, RAC, códigos de compra y `TbExpedientes` | `DELETE` dinámico | `ws.BeginTrans` / `CommitTrans`, `Rollback` en error | `ExpedienteOperaciones.cls:2089-2168` |
| Enlazar familia E2E | `INSERT` en `TbE2EEnlaces`, `UPDATE TbExpedientes.OrdinalE2E` | SQL dinámico | `ws.BeginTrans` | `E2EEnlacesService.bas:101-151` |
| Desvincular familia E2E | `UPDATE TbE2EEnlaces` (fecha, motivo «Manual desde Gestión», usuario) | SQL dinámico | `ws.BeginTrans` | `E2EEnlacesService.bas:191-200`, `:3834-3836` |

Las altas y ediciones no escriben aquí: las hacen `FormExpedienteAltaTipo` y [Form_FormExpediente](Form_FormExpediente.md).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `DELETE * FROM` e `IDExpediente` concatenado sin parámetros (`ExpedienteOperaciones.cls:2104-2148`); `Nz()` dentro del SQL de marcadores (`E2EEnlacesUIHelper.bas:173`); transacciones de espacio de trabajo DAO; el filtrado se hace en VBA sobre la colección completa en memoria, no en el motor (`FUNCIONES UTILES.bas:457-461, 764-769`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Tras eliminar un lote o un basado, `constructor.getExpediente` (devuelve `Expediente`, `constructor.bas:1139`) se asigna a variables `ExpedienteCompleto` (`:9-11`, `:2788`, `:2799`): posible error de tipos **después** del borrado confirmado. Igual en `CambiarTipo` y `Eliminar` cuando la selección está vacía (`:2647-2652`, `:2754-2759`).
- Con `DatosEnMemoria` distinto de «Sí», `Rellenar*PorTabla` declara `Expediente` (`:309, 456, 597`) pero recibe `ExpedienteCompleto` (`FUNCIONES UTILES.bas:457`), y `getColLotesPorTablas` pide `.Lotes`, que `ExpedienteCompleto` no expone (`FUNCIONES UTILES.bas:722`).
- `ListaLotes_Click` compara el padre del lote con `Column(1)` (Tipo) en vez de `Column(0)` (ID) (`:3054`): siempre reemplaza la lista de AM; rótulo «Acuerdo Mardo» (`:3058`).
- «Ver Todos» de Post AGEDO deja «Sólo» (`:1080`, `.form.txt:2807`), igual que el valor inicial (`FUNCIONES UTILES.bas:1478`); «Ver Todos» de ámbito deja «Defensa» (`:1076`).
- `ComandoIniciarEstado_Click` no muestra el error ni quita el reloj de arena (`:1305-1311`).
- `m_ObjExpBusquedaActiva` nunca se asigna (`Variables Globales.bas:269`; usos `:78, 1579, 3588`): la rama de restauración de `setExpBusqueda` es inalcanzable, y además usa `MarcoLugar`, que no existe en el `.form.txt`, y «SinClasificacion» frente a «Sin Clasificación» (`FUNCIONES UTILES.bas:1533, 1539, 1169`).
- `IniciarCampo` sin nombre (reinicio total) no tiene llamadas y omite `MarcoSeguridad` (`:1016-1061`).
- Filtro E2E: una consulta por expediente en cada una de las cuatro búsquedas (`FUNCIONES UTILES.bas:1046-1047`; `modExpedienteSearch.bas:39-59`).
- Filtros por subcadena: un nombre contenido en otro coincide (`FUNCIONES UTILES.bas:919, 965, 971, 995`).
- `ComandoEnlaceE2E_Basado_Click` no refiltra, a diferencia de AM y Lote (`:3937-3963` frente a `:3884`, `:3920`).
- `ExpedienteOperaciones.Eliminar` no borra filas de `TbE2EEnlaces` (`ExpedienteOperaciones.cls:2104-2148`).
- `Filtrar` lanza `Err.Raise 100` y el manejador sustituye el mensaje de `getExpBusqueda` (`:683`, `:776-778`).
- Restos: `TimerInterval` sin `Form_Timer` y `m_ClickPendiente`/`m_ClickItemData` sin uso (`:21-26, 2299, 2976`); mensaje «Juridica_AfterUpdate» en `Suministrador_AfterUpdate` (`:1909`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtrar en el servidor con consulta parametrizada; un único «Restablecer filtros» coherente con los valores por defecto; navegación jerárquica AM → lote → basado como vista maestro-detalle.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir en solo lectura, o sin ser administrador, deja deshabilitados alta, alta derivada, cambio de tipo, eliminación y E2E; «Detalle» sigue disponible con selección.
2. Cambiar cualquier filtro repinta las tres listas; repetir la misma búsqueda no la vuelve a ejecutar.
3. Seleccionar un AM muestra solo sus lotes y basados; un individual no ofrece alta derivada.
4. Cambiar de tipo un AM o lote con derivados muestra el aviso y no abre el cambio.
5. Eliminar pide confirmación y, si hay derivados, DPD, gestiones de riesgo, usuarios HPS o NC, muestra el motivo y no borra nada.
6. Enlazar o desvincular informa del número de expedientes de la familia afectados.
7. Los contadores muestran las filas de cada lista y su suma.

## Procedimientos auxiliares

Procedimientos que sostienen los flujos anteriores. Citas `Form_FormExpedientesGestion.cls:línea`.

| Procedimiento | Qué hace | Evidencia |
|---|---|---|
| `RellenarAM`, `RellenarBasado`, `RellenarLote` | Despachan a la variante `PorMemoria` o `PorTabla` según `m_DatosEnMemoria`; ante un error reponen `AllowEdits = True` y solo rellenan `p_Error` si el error no es el interno 1000 | `:235-254`, `:366`, `:521-540` |
| `RellenarLotePorTabla`, `RellenarBasadoPorTabla` | Mismas columnas y reglas que la variante en memoria; `RellenarLotePorTabla` filtra por `p_IdPadre` y `p_IDExpediente` (`:479-480`); `RellenarBasadoPorTabla` deja el ejército vacío si el basado no tiene (`:621-624`) | `:451-520`, `:594-651` |
| `RellenarComboIDExp` | «Todos» más los `IDExpediente` de `ColExpedientesCompletos`; sin datos en memoria invalida antes esa caché | `:3425-3455` |
| `RellenarComboComercial`, `RellenarComboRAC` | «Todos» más los nombres de `m_ObjEntorno.Comerciales` y `m_ObjEntorno.RACs` | `:3526-3578` |
| `RellenarComboResponsableCalidad`, `RellenarComboResponsableSeguridad` | `0;Todos` más `id;Nombre` de los usuarios con el rol «Calidad» o «Seguridad» (`modResponsablePorRolHelper.GetUsuariosPorRol`); la lista se despliega con todas sus filas (`ListRows = ListCount`) | `:3456-3525` |
| `Reestablecer` | Anula las cachés `ColExpedientesParaCombo`, `Comerciales`, `RACs` y `JPs` del entorno, vuelve a llamar a `RellenarCombos` y `setExpBusqueda`. `ReestablecerTitulosListas` repone los rótulos «Acuerdos Marco y los Expedientes individuales», «Lotes» y «Basados» | `:3579-3620` |
| `ListaSeleccionada` | `Verdadero` si alguna fila de la lista, sin contar la de cabecera, está seleccionada; una lista con solo cabecera da `Falso` | `:2253-2275` |
| `DeseleccionarLista` | Quita la primera fila seleccionada de la lista indicada; sin nombre, recorre AM, Lotes y Basados en ese orden | `:2101-2139` |
| `InvalidarCacheCambioE2E` | Descarta `m_MarcadoresCambioE2E`; `MarcadorCambioE2E` los recarga una sola vez al siguiente uso | `:35-50` |

## Pendientes

- Si el código embebido en `.form.txt` (`:5796` en adelante), más antiguo que el `.cls` (sin columna «Cambio» ni `EnlazadoE2E_AfterUpdate`), es el que contiene el binario.
- Si `Option Compare Database` (`:2`) hace que el `OpenArgs` «SoloLectura» (`Form_Form0BDOpcionesTecnicos.cls:69`) case con «Sololectura» (`:2910`).
- Contenido de `getExcelConsulta`/`GenerarConsultaExpedientes` (`Form_FormEleccionTipoConsulta.cls:300-330`, `FUNCIONES UTILES.bas:1769`) y de `PintarTareas` (`FUNCIONES UTILES.bas:3974`).
- Reglas de `DescubrirFamilia` en el enlace E2E (`E2EEnlacesService.bas:104`).
- Tratamiento que da `FormExpedienteAltaTipo` a un `OpenArgs` vacío (`:2475`).
