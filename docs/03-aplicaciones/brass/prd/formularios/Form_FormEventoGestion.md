# Form_FormEventoGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormEventoGestion.cls` (2892 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pantalla maestra de gestión de eventos (CAP-3) con tres listas encadenadas: eventos (`ListaEventos`), actividades del evento (`ListaActividades`) y materiales de la actividad (`ListaMateriales`). Es un formulario **no enlazado** (sin `RecordSource` en las propiedades del `.form.txt`, `Form_FormEventoGestion.form.txt:3-127`): las listas se alimentan con `RowSource` generado en VBA por `FiltrarEventos`, `FiltrarActividades` y `FiltrarMateriales` (`:1221, :1379, :1417`). Desde aquí se filtra, se da de alta/edita/copia/elimina, se franquea, se adjunta, se exporta a Excel y se envía un evento a la lista de informe del RAC. Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`; `Constructor.bas` = `modules/Constructor.bas`. Las citas son `Form_FormEventoGestion.cls:línea` salvo indicación; `.form.txt` es el del propio formulario.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Menú `Form0BDOpciones` | `lblEventos_Click` abre el formulario | `Form_Form0BDOpciones.cls:121-123` |
| Otros formularios (equipos, partes, planificaciones, materiales, informes, alta masiva) | `DoCmd.OpenForm "FormEventoGestion"` | `Form_FormEquipoGestion.cls:266`, `Form_FormParteGestion.cls:249`, `Form_FormPlanificacionesBusqueda.cls:509`, `Form_FormEventosAltaMasivaPorNS.cls:80`, `Form_FormInformesDesdeLista.cls:29-32` |
| Formularios hijos que lo refrescan | Llaman a procedimientos públicos: `ActualizarEventos`, `ListaEventos_Click`, `ListaActividades_Click`, `ListaMateriales_Click`, `ComandoActualizarTodo_Click`, `ComboIDEventoBusqueda_AfterUpdate` | `Form_FormInformesDesdeLista.cls:42`, `Form_FormActividadDeEventoPorLotes.cls:280-281`, `Form_FormAnexos.form.txt:1234-1240`, `Form_FormEventoAlta.cls:116-120` |

No recibe `OpenArgs`. Abre: `FormEventoAlta` (`:1133`), `FormEventoEdicion` (`:1618`), `FormEventoCopia` (`:1527`), `FormEventoFranqueo` (`:1988`), `FormParteGestion` (`:2029`), `FormActividadAlta` (`:1090, :1119, :1499`), `FormActividadEdicion` (`:1569`), `FormActividadDeEventoPorLotes` (`:166`), `FormEventosAltaMasivaPorNS` (`:187`), `FormMaterialEleccion` (`:1209`), `FormMaterialesGestion` (`:1659`), `FormAnexos` (`:939, :988, :1036`). Pasa datos por `OpenArgs` con `|` (`:1090, :1499, :1569, :1618, :1209`) y por las globales `m_ObjEventoActivo` (`:162, :1523, :1616, :1987`) y `m_ObjEntidadParaAnexoActiva` (`:936, :986, :1034`) (`Variables Globales.bas:244, :252`). Escucha `AltaEvento` de `FormEventoCopia` (`:14, :2844`) y `Franqueado` de `FormEventoFranqueo` (`:13, :2869`). Hacia `FormInformesDesdeLista` envía el evento (`:457-461`).

## Controles y eventos

Controles: 62 (73 Begin en el Form_FormEventoGestion.form.txt: 1 `Form`, 9 plantillas de estilo por defecto en `Form_FormEventoGestion.form.txt:224-356`, 1 `Section` en `:377` y los 62 controles). Una fila por control o grupo homogéneo.

| Control (n) | Tipo | Evento y handler | Qué hace / origen de datos / validación | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ListaEventos` (1) | lista, 10 columnas | `OnClick` → `ListaEventos_Click`; `OnDblClick` → edita si `ComandoEditarEvento` está habilitado | Columnas: 0 «F» (`*` si hay factura), 1 `IDEvento`, 2 `BUI`, 3 `SUBSISTEMA`, 4 `Equipo`, 5 alta, 6 descripción, 7 «Franqueo» Sí/No, 8 `IDEquipo`. Habilita/deshabilita botones, carga anexo y parte | CAP-3 | `:2266-2403, :791-800`; `Form_FormEventoGestion.form.txt:412-443` |
| `ListaActividades` (1) | lista, 8 columnas | `OnClick` → `ListaActividades_Click` | Columnas: 0 «M» (`*` si hay material), 1 ID, 2 técnico, 3 alta, 4 horas laborables, 5 descripción, 6 `IDFactura`. Carga anexo y materiales | CAP-4 | `:2159-2265`; `Form_FormEventoGestion.form.txt:621-647` |
| `ListaMateriales` (1) | lista, 7 columnas | `OnClick` → `ListaMateriales_Click` | Columnas: `IDMaterial`, `Material`, `PN`, `NS`, `EsReparacion`, `IDFactura`. Habilita detalle, adjuntos y borrado (no si el evento está facturado) | CAP-7 | `:802-890`; `Form_FormEventoGestion.form.txt:828-851` |
| `ComboIDEventoBusqueda` (1) | combo, 4 columnas | `AfterUpdate`; `OnEnter` hace `Requery` | Filtro por `IDEvento`; refresca eventos solo si cambia el valor | CAP-3 | `:551-580`; `Form_FormEventoGestion.form.txt:1897-1924` |
| `ComboBUIBusqueda` (1) | combo | `AfterUpdate` | Filtro `BUI`; refresca las tres listas y recarga subsistemas | CAP-12 | `:473-499`; `Form_FormEventoGestion.form.txt:1977-1999` |
| `ComboSubsistemaBusqueda`, `ComboEquipoBusqueda` (2) | combos | `AfterUpdate` | Filtros; el subsistema recarga `Equipo`; ambos dependen de la BUI | CAP-12 | `:582-606, :525-549`; `Form_FormEventoGestion.form.txt:1939-1962, Form_FormEventoGestion.form.txt:2095-2117` |
| `ComboTipoEventoBusqueda`, `ComboFranqueadoBusqueda` (2) | combos | `AfterUpdate` | Filtros por tipo de evento y por franqueo (lista de valores «Sí;No») | CAP-3 | `:608-631, :501-523`; `Form_FormEventoGestion.form.txt:3592-3615, Form_FormEventoGestion.form.txt:3922-3945` |
| `FechaInicialBusqueda`, `FechaFinalBusqueda` (2) | texto, formato fecha corta | `OnExit` | Rango sobre `FECHAALTAEVENTO`; actualiza si cambia el valor | CAP-3 | `:656-678, :633-655`; `Form_FormEventoGestion.form.txt:2159-2231` |
| `ComandoLimpiarIDEvento`, `…BUI`, `…Fechas`, `…Equipo`, `…SubSistema`, `…TipoEvento`, `…Franqueado` (7) | botones | `OnClick` | Vacían su filtro y refrescan las listas solo si había valor (`On Error Resume Next`) | CAP-3 | `:358-422, :267-357` |
| `ComandoActualizarTodo` (1) | botón | `OnClick` | Invoca seis limpiezas de filtro (no la de franqueo) | CAP-3 | `:27-51`; `Form_FormEventoGestion.form.txt:2392-2402` |
| `ComandoAltaEvento` (1) | botón | `OnClick` | Abre `FormEventoAlta` | CAP-3 | `:1131-1134` |
| `ComandoEditarEvento`, `ComandoCopiarEvento` (2) | botones | `OnClick` | Abren edición (pasa ID, facturado, franqueado) y copia | CAP-3 | `:1582-1629, :1511-1542` |
| `ComandoEliminarEvento` (1) | botón | `OnClick` | Confirma y llama a `Eliminar` | CAP-3 | `:1747-1821` |
| `ComandoFranquear` (1) | botón con rótulo dual | `OnClick` | «Franquear» abre `FormEventoFranqueo`; «Quitar Franqueo» llama a `Evento.QuitarFranqueo` | CAP-3 | `:1952-2026` |
| `ComandoAltaActividad`, `ComandoAltaActividadSinEvento`, `ComandoAltaActividadesPorLotes`, `ComandoEditarActividad`, `ComandoCopiarActividad`, `ComandoEliminarActividad` (6) | botones | `OnClick` | Alta (evento seleccionado o sin evento), alta por lotes, edición, copia y borrado | CAP-4 | `:1048-1129, :152-174, :1543-1581, :1450-1510, :1674-1746` |
| `ComandoAltaMaterial`, `ComandoDetalleMaterial`, `ComandoEliminarMaterial` (3) | botones | `OnClick` | Alta (`FormMaterialEleccion`), detalle en `FormMaterialesGestion`, borrado | CAP-7 | `:1135-1220, :1630-1673, :200-264` |
| `ComandoAdjuntosEventos`, `ComandoAdjuntosActividades`, `ComandoAdjuntosMateriales` (3) | botones «Adjuntar» | `OnClick` | Abren `FormAnexos` con la entidad en `m_ObjEntidadParaAnexoActiva` | CAP-11 | `:951-999, :891-950, :1000-1047` |
| `ImagenAnexoEvento`, `ImagenAnexoActividad`, `ImagenAnexoMaterial`, `ImagenParteAnexado` (4) | imágenes | `OnClick` | Con un solo anexo, `AbrirEnLocal`; con varios (`;`) abre `FormAnexos`; el parte se abre con `AbrirEnLocal` | CAP-11 | `:716-752, :679-715, :753-789, :2130-2158` |
| `ComandoExportarExcelEventos`, `…Actividades`, `…Materiales` (3) | botones (rótulo «Comando227») | `OnClick` | Exportan a Excel según los filtros o el `RowSource` vigente | CAP-10 | `:1866-1908, :1823-1865, :1910-1951` |
| `ComandoGestionPartes` (1) | botón | `OnClick` | Abre `FormParteGestion` | CAP-3 | `:2027-2030` |
| `ComandoParaInforme` (1) | botón (oculto si no está abierto `FormInformesDesdeLista`) | `OnClick` | Añade el evento a `m_ColEventosParaInformeDeRAC` | CAP-10 | `:424-470, :2077-2081` |
| `ComandoAltaMasivaPorNS` (1) | botón | `OnClick` | Abre `FormEventosAltaMasivaPorNS` | CAP-3 | `:176-198` |
| `cmdSalir` (1) | botón | `OnClick` | Cierra el formulario sin guardar | sin asignar | `:2126-2129` |
| `ComandoAyudaFormulario` (1) | botón | `OnClick` declarado en el `.form.txt`; no hay handler en el `.cls` | Sin comportamiento localizado | sin asignar | `Form_FormEventoGestion.form.txt:2340-2350` |
| `Cuadro2`, `Cuadro3` (2); `lblTitulo1`, `lblTitulo2`, `lblAnexo`, `lblParte` (4); `Etiqueta142`, `267`, `268`, `273`, `276`, `278`, `16`, `21` (8) | rectángulos y etiquetas | — | Decoración y rótulos; `lblAnexo` y `lblParte` se muestran según haya anexo o parte | sin asignar | `Form_FormEventoGestion.form.txt:388, Form_FormEventoGestion.form.txt:2275, Form_FormEventoGestion.form.txt:2297-2329, Form_FormEventoGestion.form.txt:4194-4221`; `:2067-2070` |

## Flujos

**Apertura (`Form_Open`, `:2031-2125`):** ajusta tamaño (`AjustarTamaño`, `FG.bas:261`), toma el título de `lblTitulo1` (`:2040`), deja habilitados solo alta de evento, gestión de partes y alta de actividad sin evento (`:2041-2059`), muestra `ComandoParaInforme` si `FormInformesDesdeLista` está abierto (`:2077-2081`), carga subsistemas y equipos (`:2082-2083`) y asigna los tres `RowSource`. Sin ningún filtro `FiltrarEventos` devuelve `IDEvento='-1'` y la lista queda vacía (`:1249-1253`). Un error muestra `MsgBox` y cierra el formulario (`:2119-2124`).

**Filtrado:** cada combo o fecha compara con el último valor (`strBUIAnterior`, etc., `:481, :509, :532, :559, :590, :616, :641, :664`) y llama a `ActualizarEventos` (`:2405-2508`), que vacía actividades y materiales, regenera la lista y, si queda un solo evento o hay uno seleccionado, dispara `ListaEventos_Click` (`:2479-2497`). `ActualizarActividades` y `ActualizarMaterial` repiten el patrón en cascada (`:2509-2624`).

**Edición, copia y alta:** exigen evento seleccionado y, según el caso, que no esté franqueado ni facturado (`:1063-1079, :1485-1488, :1701-1708`). **Eliminar evento:** confirmación (`:1784-1789`) → `Eliminar` (`:2683-2839`): rechaza si el evento figura en un parte (`:2738-2748`), elimina anexos (`:2762`) y equipos de medida (`:2767`), borra actividades una a una con `ActividadEliminar` (`:2783`), borra `TbEventosEquipoMedida` (`:2807-2810`) y el evento (`:2820-2822`). **Franqueo:** alta abre `FormEventoFranqueo`; el evento `Franqueado` fuerza la recarga (`:2869-2884`); quitar franqueo pide confirmación (`:1999-2002`).

**Errores:** patrón uniforme: `m_Error` + `Err.Raise 1000` para errores de negocio y `MsgBox` crítico (p. ej. `:45-50`). Los botones «Limpiar» usan `On Error Resume Next` y no informan de fallos (`:268, :292, :314, :337, :359, :381, :403`).

## Reglas, validaciones y cálculos

- **Facturado:** `Column(0)` de `ListaEventos` = `*` (existe `IDFactura`, `:1241`); impide alta, borrado y edición de adjuntos de hijos (`:1076-1079, :1485-1488, :1777-1780`, `:916-925`).
- **Franqueado:** `Column(7)` debe ser «Sí» o «No»; con «Sí» se bloquea alta de actividad y material, copia y borrado (`:1067-1075, :1155-1158, :2323-2333`).
- **Pertenencia a parte:** un evento incluido en `TbPartesDetalle` no se puede eliminar (`:2738-2748`).
- **Filtros:** sin fecha inicial se usa `01/01/1900` y sin final `01/01/2900` (`:1334-1347`); los textos usan comparación exacta y el resto `Like '*'` (`:1257-1332`). `GestorInforme.Registrar` recibe los campos filtrados (`:1356`; `GestorInforme.cls:76`).
- **Informe RAC:** un evento no puede añadirse dos veces a `m_ColEventosParaInformeDeRAC` (`:443-446`).
- **Habilitación de botones:** depende de la selección y del estado facturado/franqueado (`:2183-2219, :2285-2343`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de eventos | `TbEventos`, `TbEquipos`, `TbFacturaEventosInvolucrados` | SQL dinámico concatenado como `RowSource` (`IIf`, `Like '*'`, fechas `#mm/dd/yyyy#`) | — | `:1241-1245, :1357-1368` |
| Lista de actividades y materiales | `TbActividades`, `TbMaterial`, `TbFactura*Involucrad*` | SQL dinámico concatenado como `RowSource` | — | `:1389-1405, :1427-1440, :2444-2460` |
| Combos de filtro | `TbSubsistemaBui`, `TbEquipos`, `TbBUI`, `TbTipoEvento`, `TbEventos` | `RowSource` fijo en el `.form.txt`; subsistema y equipo con SQL concatenado | — | `:65-71, :116-119`; `Form_FormEventoGestion.form.txt:1919, Form_FormEventoGestion.form.txt:1960, Form_FormEventoGestion.form.txt:1997, Form_FormEventoGestion.form.txt:2115, Form_FormEventoGestion.form.txt:3613` |
| Cargar entidades | `TbEventos` y equivalentes | SQL concatenado con el ID entre comillas (`getEvento`) | — | `Constructor.bas:2242-2260`; `:449, :839, :926` |
| Exportación a Excel | consultas generadas | SQL a partir del `RowSource` troceado por `Where` | — | `:1834-1848, :1921-1934`; `FG.bas:1584, :1689`; `Constructor.bas:2467` |
| Parte firmado (no se llama desde el formulario) | `TbPartesPpal`, `TbPartesDetalle` | DAO `OpenRecordset` con SQL concatenado | — | `:2652-2656` |
| Borrado de evento | `TbPartesDetalle`, `TbActividades`, `TbEventosEquipoMedida`, `TbEventos` | Lectura con `Dame`, DAO `.Delete` y `Execute` | no se observa `BeginTrans` | `:2738, :2769-2772, :2807-2822` |
| Borrado de actividad/material | `TbActividades`, `TbMaterial`, `TbMaterialSeguimiento` | DAO y `DELETE` en `FG.bas` | no se observa `BeginTrans` | `FG.bas:5581-5739, FG.bas:4785-4896` |
| Quitar franqueo | `TbEventos` | `UPDATE` ejecutado con `getdb().Execute` | no | `Evento.cls:1426-1430` |

## Coexistencia

Adaptar a PostgreSQL vinculado: `IIf`/`IsNull`, `Like '*'`, literales `#mm/dd/yyyy#` y booleanos `True`/`False` en los filtros (`:1241, :1257, :1318, :1357`); `DELETE tabla.*` y `DELETE * FROM` (`:2807, Evento.cls:1982`); `Max`/`Dame` y recorridos DAO. Los ID de evento, actividad y material son texto (`='-1'`, `:1251, :1397, :1432`). El borrado en cascada se hace en VBA sin transacción: debe preservarse el orden y la protección por franqueo, facturación y parte. Pendiente de `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `ComandoFranquear_Click` obtiene el ID de evento de `Column(0)` (`:1964`), que es el marcador de facturado; el ID está en `Column(1)` (`:1241, :2275`). Solo afecta si `strIDEventoSeleccionado` llega vacío.
- `ListaActividades` devuelve 7 columnas (`:1389-1395`) pero `ComandoEditarActividad` y `ComandoEliminarActividad` leen `Column(7)` como «factura» (`:1552, :1684`): podría ser siempre vacío y tratar toda actividad como no facturada.
- `Eliminar` nunca incrementa `intNumeroActividadesEliminadas` (`:2783-2800`): el resumen devuelto muestra 0 actividades eliminadas.
- Borrado no atómico: anexos y equipos de medida se eliminan (`:2762, :2767`) antes de saber si las actividades se borran, y el evento se borra aunque alguna actividad falle (`:2785, :2820`).
- `ListaActividades_Click` oculta `ImagenAnexoMaterial` en lugar de `ImagenAnexoActividad` cuando no hay URL (`:2235`).
- SQL concatenado con valores de combos y del usuario sin escapar comillas (`:70-71, :1260, :1296, :2771`).
- `DameURLParteDeEvento` no tiene llamadas localizadas (`:2625-2681`): posible código muerto; el parte se obtiene con `Evento.URLAnexoParte` (`:2346`).
- Rótulos de diseño inconsistentes: botones «Comando227», «Comando36» y «Comando46» (`Form_FormEventoGestion.form.txt:1238, Form_FormEventoGestion.form.txt:3084, Form_FormEventoGestion.form.txt:2349`); `ComandoAyudaFormulario` sin handler.
- `Form_Open(Canel …)` y sin comprobación de rol en el `.cls`: cualquier usuario con acceso al menú puede eliminar y franquear (no se observa lectura de permisos en `:2031-2125`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: borrado en una transacción con borrado lógico; filtros parametrizados; estados «facturado» y «franqueado» como campos explícitos en lugar de columnas posicionales.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin filtros la lista de eventos aparece vacía; con un filtro se ordena por fecha de alta descendente (`:1249-1253, :1366`).
2. Seleccionar un evento habilita editar, copiar, franquear y adjuntar; si está franqueado o facturado no permite alta de actividad ni borrado (`ListaEventos`, `:2323-2338`).
3. Seleccionar una actividad muestra sus materiales; si el evento está facturado no se puede eliminar el material (`:843-847`).
4. Eliminar un evento incluido en un parte muestra el error y no borra nada (`:2745-2748`).
5. Eliminar un evento pide confirmación y borra actividades, materiales, anexos y equipos de medida (`:1784-1789, :2762-2822`).
6. Franquear abre el diálogo y, al terminar, el botón pasa a «Quitar Franqueo» (`:2339-2343, :2869-2884`).
7. «Añadir a Informe» solo aparece con el formulario de informes abierto y rechaza duplicados (`:2077-2081, :443-446`).
8. La exportación a Excel respeta los filtros activos (`:1875-1894`).
9. Un anexo único se abre directamente; varios abren la gestión de anexos (`:690-704`).

## Pendientes

- Contenido de `AjustarTamaño`, `EstaEnCombo`, `AbrirEnLocal`, `Dame`, `GenerarConsultasSQL`, `GenerarConsultaEventos` y `getEventosFiltrados` (`:37, :83, :696, :2738, :1849, :1894, :1886`).
- Comportamiento interno de `ActividadEliminar`, `EliminarMaterialEnActividad` y `Evento.EliminarAnexos`/`EliminarEquiposMedida` (`:2783, :245, :2762, :2767`).
- Formularios abiertos por nombre en tiempo de ejecución con `DoCmd.OpenForm` (`:166, :187, :939, :1133, :1209, :1527, :1569, :1659, :1988`) y acceso por `Forms("…").Controls("ComboIDMaterialBusqueda")` (`:1660`).
- Si `lblTitulo1` y `Caption` coinciden y su origen en `.form.txt` (`:2040`).
- `ComandoEditarMaterial_Click` figura citado en comentarios de `FG.bas:6056` pero no existe en este formulario.
