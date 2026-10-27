# Form_FormSolicitudesGestion

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormSolicitudesGestion.cls` (912 líneas) y `.form.txt` · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `fichero:línea` del commit fijado.

Listado de gestión de solicitudes: filtra por estado, tramitador, tipo, empresa, expediente, nombre, nemotecnico y DNI; da de alta, edita, elimina, reprograma correos, exporta a Excel y muestra el log de la solicitud seleccionada. Es un formulario **no enlazado** (sin `RecordSource`; la lista se puebla con `AddItem`, `.form.txt:754`) y es el único que **borra** datos de negocio. Escribe exclusivamente a través de `SolicitudOperaciones` y de `FU.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpciones` | `DoCmd.OpenForm "FormSolicitudesGestion"` solo si no estaba ya abierto | `Form_Form0BDOpciones.cls:195-196` |

Abre otros formularios: `FormSolicitudAlta` (`:489-493`), `FormSolicitudAltaTraspaso` con `OpenArgs:="Alta"` o `"Traspaso"` (`:528-536`), `FormSolicitudRenovacion` con `OpenArgs:=.TIPO` (`:543-546`) y, a través de `FU.bas`, `FormWeb` (`:701`; `FU.bas:1737-1740`).

No declara parámetros. Recibe su contexto de la global `m_ObjEntorno` (`VG.bas:190`), del usuario conectado `m_ObjUsuarioConectado` (`VG.bas:189`) y de `m_ObjSolicitudActiva` (`VG.bas:197`), que escribe en `ComandoEditar_Click` (`:525`).

**Acoplamiento entrante:** `FU.bas` escribe directamente en los controles de este formulario cuando está abierto: `lblUltimaModificacion` (`FU.bas:129-131`) y `ListaFiltrados` (`FU.bas:162-165`, `FU.bas:145-180`). El formulario expone además `Filtrar` como `Public` para que otros formularios lo refresquen (`Form_FormSolicitudesAltaFechas.cls:1142-1143`, `Form_FormSolicitudRenovacionFechas.cls:1125-1126`, `Form_FormSolicitudesAltaTraspasoDatos.cls:468-469`, `Form_FormSolicitudesTraspasoFechas.cls:722-723`, `Form_FormSolicitudesAltaRenovacionFechas.cls:357-358`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaFiltrados` | lista 5 columnas con cabecera | `OnClick`, `OnDblClick` | Selecciona la solicitud; el doble clic edita | `:813-899`, `:747-757`, `.form.txt:735-758` |
| `Estado` | combo 2 columnas | `AfterUpdate` | Filtra por estado | `:727-729`, `.form.txt:989-1013` |
| `Gestor` | combo 2 columnas | `AfterUpdate` | Filtra por tramitador | `:732-734`, `.form.txt:1056-1080` |
| `TIPO` | combo 1 columna | `AfterUpdate` | Filtra por tipo de solicitud | `:909-911`, `.form.txt:926-948` |
| `IDEmpresaUsuario` | combo 2 columnas | `AfterUpdate` | Filtra por empresa | `:736-739`, `.form.txt:1178-1202` |
| `IDExpediente` | combo 2 columnas | `AfterUpdate` | Filtra por expediente | `:741-744`, `.form.txt:1255-1279` |
| `Nombre` | texto | `AfterUpdate` | Filtra por nombre del solicitante | `:905-907`, `.form.txt:1121` |
| `DNI` | texto | `AfterUpdate` | Filtra por DNI | `:723-725`, `.form.txt:1548-1563` |
| `Nemotecnico` | texto | `AfterUpdate` | Filtra por nemotecnico | `:901-903`, `.form.txt:1512-1527` |
| `lblListaDeResultados` | etiqueta | — | Contador de resultados filtrados | `:179`, `:364`, `.form.txt:788` |
| `lblUltimaModificacion` | etiqueta | — | Última modificación registrada; solo visible para administrador | `:778-787`, `.form.txt:1783`; `FU.bas:129-131` |
| `ComandoActualizar` | botón | `OnClick` | Restablece los filtros a sus valores iniciales | `:380-424`, `.form.txt:600-602` |
| `ComandoActualizarEstado` | botón | `OnClick` | Recalcula el estado de todas las solicitudes y vuelve a filtrar | `:426-448`, `.form.txt:687-689` |
| `ComandoAlta` | botón | `OnClick` | Abre el alta de solicitud | `:478-508`, `.form.txt:810-812` |
| `ComandoEditar` | botón | `OnClick` | Abre el editor según el tipo de la selección | `:510-565`, `.form.txt:868-870` |
| `ComandoEliminar` | botón | `OnClick` | Elimina definitivamente la solicitud seleccionada | `:567-617`, `.form.txt:1312-1314` |
| `cmdReprogramarCorreos` | botón | `OnClick` | Reprograma los correos automáticos futuros de la selección | `:12-45`, `.form.txt:1596-1598` |
| `ComandoVerLogs` | botón | `OnClick` | Abre el log de la solicitud en `FormWeb` | `:687-721`, `.form.txt:1361-1363` |
| `ComandoExportarExcel` | botón | `OnClick` | Exporta las filas filtradas a un `.xlsx` y lo abre | `:619-685`, `.form.txt:1410-1412` |
| `ComandoAyuda` | botón | `OnClick` | Abre el PDF de ayuda del formulario | `:450-476`, `.form.txt:507-509` |
| `cmdSalir` | botón | `OnClick` | Cierra el formulario | `:146-148`, `.form.txt:1677-1679` |

## Flujos

### Apertura (`Form_Open`, `:760-811`)

1. `Ajustar Me` dimensiona el formulario (`:769`; `FU.bas:735-758`) y el título se toma de la etiqueta (`:773`).
2. `EstablecerCombos` puebla los cuatro combos (`:774`; `:47-144`): empresas desde `m_ObjEntorno.Suministradores` (`:64-80`), expedientes desde `m_ObjEntorno.Expedientes` (`:83-99`), tramitadores desde `m_ObjEntorno.Coltramitadores` con el usuario conectado por defecto (`:101-120`) y estados desde `m_ObjEntorno.ColEstadosParaTitulo` (`:122-134`), con `EnTramite` como valor inicial (`:134`).
3. Si el usuario es administrador muestra `lblUltimaModificacion` con `m_ObjEntorno.UltimoCambio.texto`; si no, la oculta (`:778-788`).
4. `Filtrar` carga la lista (`:789`).

### Filtrado (`Filtrar`, `:150-378`)

1. Anula la selección y deshabilita exportar, editar, eliminar y ver logs (`:164-183`); fija las cabeceras `IDSolicitud;Tipo;DNI;Nombre;Estado` (`:178`).
2. Obtiene la colección por estado desde la caché de `m_ObjEntorno` (`:190-296`). Con estado `Todos` usa `ColSolicitudes` (`:284-286`), que carga `TbSolicitudes` completa (`:2036-2053`; `Entorno.cls:1572-1590`).
3. Aplica en memoria los filtros finos (`:301-357`): tipo, tramitador, nombre (`vbTextCompare`, `:317`), empresa, expediente, nemotecnico (`:340-349`) y DNI (`:353-357`).
4. Añade una fila por solicitud (`:359`), actualiza el contador (`:364`), selecciona la primera (`:368`) y dispara `ListaFiltrados_Click`.

### Selección y acciones

- **Editar** (`:510-565`): fija `m_ObjSolicitudActiva` (`:525`) y abre el editor que corresponde al `TIPO`; si es renovación y no hay `UsuarioHPS` avisa y no abre (`:537-541`).
- **Eliminar** (`:567-617`): pide confirmación (`:584`) y llama a `SolicitudOperaciones.Eliminar` (`:591`; `SolicitudOperaciones.cls:284-341`).
- **Exportar** (`:619-685`): relee cada fila por identificador (`:642`), genera el Excel (`:656-658`; `FU.bas:2108-2150`) y lo abre con `Shell` (`:668`).
- **Logs** (`:687-721`): `FU.bas.InformeDelRegistro` genera el HTML y abre `FormWeb` (`:701`; `FU.bas:1721-1745`).
- **Reprogramar correos** (`:12-45`): pide confirmación (`:17`), delega en `SolicitudOperaciones.ReprogramarAvisosFuturos` (`:31`) y muestra el resultado (`:35-39`).

### Alternativos

- `ListaFiltrados_DblClick` (`:747-757`) delega en `ComandoEditar_Click` si el botón está habilitado.
- `ComandoActualizar_Click` (`:380-424`) pone `TIPO="Todos"`, `Gestor` = usuario conectado, empresa y expediente a `0`, estado a `0`, borra `Nombre` (`:387-393`), invalida todas las cachés de estado (`:394-408`) y vuelve a filtrar.
- `ComandoActualizarEstado_Click` (`:426-448`) ejecuta `FU.bas.ActualizarEstados` (`FU.bas:2315-2350`) antes de filtrar.

### Errores

Todos los handlers usan `Err.Raise 1000` como error interno controlado: el mensaje ya redactado se muestra en `MsgBox` y **no** se notifica por correo (`:417-424`, `:800-810`). Solo los errores con número distinto de 1000 disparan `CorreoAlAdministrador` (`:802-804`). `AvanceCerrar` se ejecuta en el camino feliz y en el de error de casi todos los botones (`:436` y `:445`, `:562`, `:601`, `:614`, `:718`).

## Reglas, validaciones y cálculos

- **Filtro por estado:** cada estado abre su propia colección de `m_ObjEntorno`, que a su vez llama a un `constructor.getSolicitudes…` distinto (`Entorno.cls:1543-1556`; `constructor.bas:668-685`, `constructor.bas:1835-1857`). El estado «Todos» no usa filtro SQL.
- **Filtro por tramitador:** la comparación se hace contra `Column(1)` de `Gestor`, es decir contra el **nombre** del tramitador, no contra su identificador (`:188`, `:310`; `constructor.bas:2054-2056`).
- **Comparación de textos:** nombre y nemotecnico usan `vbTextCompare` (`:317`, `:347`); el DNI usa el comparador predeterminado de `InStr`, sensible a mayúsculas (`:354`).
- **Nemotecnico:** si la propiedad de la solicitud está vacía, se recurre a `.Expediente.Nemotecnico` (`:342-345`).
- **Separador de la lista:** el título se concatena con `;` sin escapar (`:359`), a diferencia de `Form_FormExpedientesBusqueda` (`:235`), que sí sustituye el punto y coma.
- **Permisos:** el único control de rol del formulario es `EsAdministrador = EnumSiNo.Sí`, y solo afecta a la visibilidad de `lblUltimaModificacion` (`:778`). Quien llega al formulario lo hace por el menú de `Form_Form0BDOpciones` (`:195-196`); la comprobación de rol está en ese formulario, no aquí.
- **Borrado:** `SolicitudOperaciones.Eliminar` borra `TbCorreosEnviados`, `TbSolicitudesFechas` y `TbSolicitudes` del identificador y elimina el fichero adjunto y su carpeta (`SolicitudOperaciones.cls:300-334`). Si el anexo está abierto en otra aplicación, se aborta (`SolicitudOperaciones.cls:296-301`).
- **Reprogramación:** solo se ofrece si la configuración global y la de la solicitud tienen los correos automáticos activos (`:870-877`; `SolicitudOperaciones.cls:2440-2450`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Listado por estado | `TbSolicitudes` (`INNER`/`LEFT JOIN` con `TbSolicitudesFechas`) | SQL por concatenación | `constructor.bas:1835-1857`, `constructor.bas:668-685` |
| Listado completo | `TbSolicitudes` | SQL por concatenación | `constructor.bas:2050-2056` |
| Solicitud por identificador | `TbSolicitudes` | SQL por concatenación | `constructor.bas:492-497` |
| Empresas, expedientes, tramitadores y estados | catálogos de `m_ObjEntorno` | caché en memoria | `Entorno.cls:268`, `Entorno.cls:543`, `Entorno.cls:1601`, `Entorno.cls:1781` |
| Eliminar solicitud | `TbCorreosEnviados`, `TbSolicitudesFechas`, `TbSolicitudes` + anexo en disco | `DELETE` encadenado + E/S de fichero | `SolicitudOperaciones.cls:308-334` |
| Recalcular estados | `TbSolicitudes` | `UPDATE` por registro | `FU.bas:2335-2338` |
| Programar correos | cola de correos de la solicitud | inserción por `CorreoOperaciones.Registrar` | `SolicitudOperaciones.cls:2299-2303` |
| Exportar a Excel | volcado de `TbSolicitudes` + `TbSolicitudesFechas` a `.xlsx` temporal | E/S binaria con automatización de Excel | `FU.bas:2152-2161` |

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Ya se ve: el filtro se hace en memoria sobre colecciones traídas enteras del servidor (`constructor.bas:2036-2053`), no en SQL; todos los `constructor` arman la consulta concatenando parámetros (`constructor.bas:2054-2056`, `constructor.bas:492-497`), patrón que en PostgreSQL exige parametrización; `DoCmd.OpenForm`/`DoCmd.Close` como mecanismo de navegación (`Form_Form0BDOpciones.cls:195-196`); el estado vive en una columna de texto con dos diccionarios de traducción que deben mantenerse sincronizados (`Entorno.cls:1703-1740`, `Entorno.cls:1742-1778`); y la exportación depende de automatización de Excel y de `Shell` (`FU.bas:2161`; `Form_FormSolicitudesGestion.cls:668`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-1.** `cmdReprogramarCorreos` confirma, ejecuta y muestra «Proceso finalizado. Se han programado los correos…» (`:17`, `:31`, `:37`), pero las seis llamadas a `ProgramarCorreoFuturo` de `ReprogramarAvisosFuturos` pasan `0` como desplazamiento (`:2484`, `:2503`, `:2523`, `:2550`, `:2569`, `:2589` en `SolicitudOperaciones.cls`) y ese método sale sin hacer nada cuando el desplazamiento no es positivo (`SolicitudOperaciones.cls:2274`). A verificar si el botón es realmente un no-op.
- **H-HPS-S-2.** `PendienteRegistroEnHPS` (`VG.bas:133`, valor 11) está en `ColEstadosParaTitulo` (`Entorno.cls:1803`) y por tanto aparece en el combo, pero `Filtrar` no tiene rama para él (`:191-278`) y cae en el `Else` que aborta con «El tipo no está codificado» (`:279-280`).
- **H-HPS-S-3.** Se declara el evento `SolicitudModificada` (`:8`) y no se lanza en ningún punto del fichero: ningún contenedor puede estar suscrito a algo que no ocurre.
- **H-HPS-S-4.** Cada rama de estado invalida su caché dos veces: una dentro del `If` de cambio (`:193`) y otra incondicional justo después (`:196-197`, `:203-204`, `:269-270` y análogas); la segunda asignación anula la primera y el bloque condicional es código muerto. Además la decisión de invalidar depende solo de `m_EnumEstadoUltimo` y `m_TramitadorUltimo` (`:293`), no del resto de filtros, de modo que cambiar nombre, DNI o empresa reutiliza la colección ya cargada.
- **H-HPS-S-5.** `constructor.bas` concatena los parámetros en la consulta (`constructor.bas:2054-2056`, `constructor.bas:492-497`, `constructor.bas:2306-2308`) y el tramitador viaja como texto entre comillas simples; a verificar el comportamiento ante un valor con apóstrofo.
- **H-HPS-S-6.** `FU.bas` modifica los controles de este formulario desde fuera (`FU.bas:129-131`, `FU.bas:162-165`): el acoplamiento es del auxiliar global al formulario, no al revés.

**Defectos de interfaz heredados (a verificar):** los `RowSource` de diseño de `Gestor`, `IDEmpresaUsuario` e `IDExpediente` contienen datos reales de personas y de empresas (`.form.txt:1080`, `:1202`, `:1279`), y el rótulo de `lblUltimaModificacion` lleva un número de expediente y una fecha (`.form.txt:1784`). `ComandoAyuda` y `ComandoActualizar` tienen el rótulo interno `Comando46` (`.form.txt:508`, `:601`). El `RowSource` de diseño de la lista tiene cuatro cabeceras y la lista tiene cinco columnas (`.form.txt:742`, `:754`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: filtros push-down a SQL con paginación; estado como enumeración persistida en lugar de texto traducido; eliminar definitivamente sustituido por baja lógica; exportación asíncrona con plantilla en lugar de automatización de Excel; y el formulario como subruta del contenedor, no como ventana independiente.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir el formulario se listan las solicitudes en trámite del usuario conectado.
2. Cambiar cualquiera de los ocho filtros refiltra la lista y actualiza el contador.
3. Seleccionar una fila habilita editar, eliminar, ver logs, exportar y reprogramar, y deshabilita todo al vaciarse la lista.
4. Editar una solicitud de tipo traspaso abre el editor de traspaso y una de renovación abre el de renovación.
5. Eliminar pide confirmación y, aceptada, desaparece la fila y sus correos y fechas quedan borrados.
6. Actualizar deja los filtros en su estado inicial y vuelve a listar.
7. Solo un administrador ve la etiqueta de última modificación.
8. Exportar genera un fichero con las filas filtradas y lo abre.

## Pendientes

- Si `SolicitudOperaciones.ReprogramarAvisosFuturos` está realmente inerte por el desplazamiento `0` (H-HPS-S-1): es la hipótesis más relevante de la ficha.
- Si existe algún contenedor suscrito a `SolicitudModificada` (H-HPS-S-3).
- Contenido de `FU.bas:145-180` (`ActualizarListaSolicitudes`), que reescribe filas de `ListaFiltrados` y comparte la lista con este formulario.
- `m_TituloFormulario` (`VG.bas:196`), que `Form_Open` asigna al título (`:770`) y acto seguido sobrescribe con la etiqueta (`:773`).
- Semántica exacta de `EsAdministrador`, que este formulario solo consume para la visibilidad de una etiqueta.