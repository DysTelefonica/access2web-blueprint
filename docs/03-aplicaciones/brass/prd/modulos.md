# Módulos de Brass

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **PRD:** [README](README.md)

Inventario funcional de los módulos estándar de `src/modules` (producción). Las citas son `fichero:línea` sobre el clon fijado; solo `FG.bas` (`modules/Funciones Generales.bas`) y `VG.bas` (`modules/Variables Globales.bas`) se abrevian, el resto se cita por su nombre. Las clases de `src/classes` y los formularios se tratan en otros ficheros del PRD.

Rasgos comunes (verificados): no hay ningún `BeginTrans` en estos módulos; el patrón de error es `On Error GoTo errores` + `Err.Raise 1000` + `p_Error` por referencia (p. ej. `Constructor.bas:5-22`, `Módulo1.bas:5-44`); las funciones de `FG.bas` devuelven los fallos como cadena `"#ERR|texto"` (p. ej. `FG.bas:2262-2266`, comentario de contrato de `AltaPlanificacion`).

## Variables Globales
Ficha: `modules/Variables Globales.bas` · 583 líneas · 5 procedimientos (más declaraciones y constantes).

**Propósito.** Declaraciones Win32 (`GetPrivateProfileString`, `ShellExecute` como `Ejecutar`, `OpenProcess`…) con rama `#If Win64` (`VG.bas:5-53`), constantes de Word/Excel (`VG.bas:54-120`), variables públicas de sesión (`m_ObjEntorno`, `m_ObjUsuarioConectado`, `wks`, `IDAplicacion`: `VG.bas:211-240`) y el arranque `EVE`.

| Procedimiento | Qué hace | Cita | Datos / efectos |
|---|---|---|---|
| `EVE` | Inicializa el entorno: fija `IDAplicacion = "6"`, `TempVars("DatosEnLocal")="No"` y `TempVars("EnDesarrollo")="No"` fijos en código, instancia `Entorno`, calcula rutas remota/local, resuelve el usuario (por `VBA.Command` o por usuario de red, con alias `Local1`/`adm1`→`adm`), exige permiso de administrador o técnico, resetea colecciones globales y recorre `ColPropiedadesEntorno` | `VG.bas:264-457`, `VG.bas:282`, `VG.bas:293-296`, `VG.bas:310-322`, `VG.bas:343-350` | Lee vía `Constructor.getUsuario`; acumula errores por propiedad y los devuelve concatenados |
| `LeerIni` | Lee una clave del fichero ini de la aplicación (sección = nombre de la BD actual) | `VG.bas:461-472` | API `GetPrivateProfileString`; consumidor `Entorno.cls:655` (`VersionAplicacion`) |
| `getRutaBackendDesdeTablasVinculadas` | Recorre `TableDefs`, toma la primera tabla vinculada y extrae el valor `DATABASE=` de su cadena de conexión | `VG.bas:474-518` | Metadatos DAO; devuelve la ruta del backend vinculado |
| `getdb` | Abre el workspace 0 sobre el backend: prioriza la ruta de tablas vinculadas; si no, `DatosEnLocal` Sí/No decide ruta local o remota; si no hay valor, error «No se conoce el origen de los datos» | `VG.bas:548-581`, `VG.bas:558-570` | Devuelve `DAO.Database`; contraseña literal en el código (`VG.bas:572`) |
| `getdbLanzadera` | Igual que `getdb` pero para la BD «Lanzadera_Datos» (local o remota según `DatosEnLocal`) | `VG.bas:520-547`, contraseña literal en `VG.bas:539` | Lee/escribe tablas de aperturas (ver FG `RegistroApertura`) |

**Datos.** Sin tablas propias; abre las dos bases por DAO (`OpenDatabase`, `VG.bas:539`, `VG.bas:572`). Sin SQL; sin transacciones.
**Consumidores.** `getdb()` aparece en 27 ficheros de `src` (p. ej. `Módulo1.bas:62`); `getdbLanzadera()` en `Constructor.bas:71`, `Constructor.bas:110`, `Constructor.bas:165`; `EVE` se invoca desde `Form_FormInformesParaRAC.cls:188`.
**Capacidad.** CAP-1 Arranque, identidad y permisos (y «transversal» para `getdb`).
**Posibles defectos (a verificar).** (1) Contraseña de ambas bases literal en el código (`VG.bas:539`, `VG.bas:572`). (2) `getdb` abre una conexión nueva en cada llamada y `db`/`wks` se usan sin declaración visible en el módulo (`VG.bas:572-573`); revisar fugas de conexión. (3) `getdb` descarta el error de `getRutaBackendDesdeTablasVinculadas` (`VG.bas:556-559`). (4) Valores `TempVars` de datos locales/desarrollo y alias de usuario `adm` fijos en el código (`VG.bas:293-296`, `VG.bas:320-321`). (5) En `EVE` el bucle de propiedades deja `m_ValorPropiedad = ""` tras calcularlo (`VG.bas:430`), de modo que el valor calculado no se usa: revisar.
**Pendientes.** `m_ObjEntorno.getValorPropiedad(CStr(m_Propiedad))` es resolución dinámica por nombre (`VG.bas:387`, `VG.bas:406`): las propiedades concretas dependen de `Entorno.cls`; llamadas a `Constructor.getUsuario` con correo de `Command` (`VG.bas:313`).

## Constructor
Ficha: `modules/Constructor.bas` · 5698 líneas · 96 procedimientos públicos.

**Propósito.** Fábrica de objetos de dominio: cada `getXxx` compone un `SELECT` concatenado, abre un `Recordset` DAO y rellena el objeto o diccionario de objetos (`Constructor.bas:23-95` `getUsuario`; `Constructor.bas:2242-2282` `getEvento`). Es la capa de lectura de toda la aplicación.

| Grupo | Procedimientos (muestra) | Cita | Tablas / efectos |
|---|---|---|---|
| Usuario y permisos | `getUsuarioConectadoPorMaquina`, `getUsuario`, `getUsuarios`, `getAplicacionesPermisos` | `Constructor.bas:5-22`, `23-95`, `96-147`, `148-202` | BD Lanzadera (`getdbLanzadera().OpenRecordset`, `Constructor.bas:71`, `110`, `165`) |
| Anexos | `getAnexo`, `getAnexosDeEvento`, `getAnexosDeActividad`, `getAnexosDeMaterial`, `getAnexosDeGastos`, `getAnexosNoAlcanzablesBusqueda` | `Constructor.bas:205-245`, `455-509`, `511-566`, `568-622`, `681-736`, `246-320` | `TbAnexos` |
| Facturación | `getFacturas`, `getFactura`, `getPartes`, `getEventosParaFactura`, `getGastosParaFactura`, `getSubContratacionesParaFactura`, `getEventosPorBUIEnFactura`, `getFacturasBusqueda` | `Constructor.bas:796-848`, `2425-2466`, `850-900`, `1224-1282`, `1135-1192`, `1073-1134`, `3691-3781`, `3887-3958` | `TbFacturaPrincipal`, `TbFacturaEventosInvolucrados`, `TbFacturaActividadesInvolucradas`, `TbGastos` |
| Eventos y actividades | `getEvento`, `getEventosFiltrados`, `getActividad`, `getActividadesDeEvento`, `getListaEventosNoFacturados`, `getListaEventosEntreFechas` | `Constructor.bas:2242-2282`, `2467-2589`, `2283-2323`, `1589-1645`, `1284-1345`, `4456-4607` | `TbEventos`, `TbActividades` |
| Materiales y seguimientos | `getMaterial`, `getMaterialSeguimiento`, `getMaterialesEnFactura`, `getMaterialesEnActividad` | `Constructor.bas:1815-1854`, `1897-1938`, `3399-3457`, `3459-3515` | `TbMaterial` |
| Técnicos, precios, fiestas, libranzas | `getTecnico`, `getTecnicos`, `getTipoTecnicoPrecio(s)`, `getFiestas`, `getLibranzas`, `getActividadesTecnicoDia` | `Constructor.bas:1940-1976`, `4249-4314`, `2035-2091`, `4002-4064`, `4118-4198`, `4396-4455` | `TbTecnicos`, `TbTipoTecnicoPrecios`, `TbTecnicosAusencias` |
| Equipos de medida | `getEquipoMedida`, `getEquiposMedidaPorEvento`, `getEventoEquiposMedida`, `getEquipoMedidaCalibraciones`, `getUltimaCalibracion` | `Constructor.bas:4878-4920`, `4758-4817`, `4965-5019`, `5063-5117`, `5180-5222` | `TbEquiposMedida`, `TbEventosEquipoMedida` |
| Getronics | `getEventosGetronic`, `getActividadesGetronic`, `getActividadesGetronic1`, `HayAlguienDeGetronicsEnEvento` | `Constructor.bas:5224-5257`, `5292-5478`, `5479-5665`, `5259-5290` | consultas por tipo de técnico (pendiente de detallar) |

Sin detallar (pendiente): ~55 `getXxx` restantes de la misma plantilla (centros, subcontrataciones, seguimientos, partes de factura…), p. ej. `getCentros` `Constructor.bas:2860-2930`, `getSubContratacion` `Constructor.bas:3846-3885`.

**Datos.** Solo lectura: 96 `OpenRecordset` y 0 `Execute` en el módulo; SQL dinámico concatenado con los identificadores recibidos (p. ej. `getUsuario` compone el `WHERE` por parámetros opcionales, `Constructor.bas:23-95`); sin transacciones.
**Consumidores.** `Constructor.` aparece en 105 ficheros de `src` (clases y formularios); p. ej. `VG.bas:313` (`Constructor.getUsuario(, , , m_Command, p_Error)`), `Constructor.bas:15` (`Constructor.getUsuario(, objNetwork.UserName…)`).
**Capacidad.** transversal (CAP-1 en usuarios; CAP-4 a CAP-9 en el resto).
**Posibles defectos (a verificar).** Concatenación de identificadores en SQL sin parametrizar en todo el módulo (patrón de `Constructor.bas:23-95`); duplicidad aparente `getActividadesGetronic` / `getActividadesGetronic1` (`Constructor.bas:5292-5478` vs `5479-5665`) y `getListaEventosNoFacturados` / `…2` (`Constructor.bas:1284-1345` vs `4609-4684`), `getListaEventosEntreFechas` / `…2` (`Constructor.bas:4456-4607` vs `4686-4757`).
**Pendientes.** Los nombres de campo se resuelven por `ColCampos` de cada clase (patrón visto en `Módulo1.bas:95`); la lista real de columnas depende de las clases `src/classes`.

## Funciones Generales
Ficha: `modules/Funciones Generales.bas` · 7942 líneas · 106 procedimientos.

**Propósito.** Biblioteca de reglas de negocio y utilidades con acceso directo a datos: altas/bajas de actividades y materiales, planificación preventiva, franqueo, facturado, registro de aperturas e informes Word.

| Tema | Procedimientos (muestra) | Cita | Tablas / efectos |
|---|---|---|---|
| Utilidades de datos | `DameNumeroRegistrosPorSQL`, `Dame`, `DameID`, `DameID1`, `DameValorSiguienteAlMaximoDeUnCampo`, `ValorExistente` | `FG.bas:86-156`, `290-330`, `371-439`, `332-370`, `440-507`, `7175-7222` | Consultas genéricas; generación de ids por máximo+1 |
| Estado de evento | `EstaFranqueado`, `EstaFacturado`, `DameIDEvento`, `DameFechaFinEvento` | `FG.bas:6033-6170`, `5758-5837`, `6171-6254`, `798-839` | `TbEventos.Franqueado` (comentario de contrato `FG.bas:6033-6060`); consumida por `Evento.cls:2888`, `Form_FormEventoSLA.cls:57` |
| Planificación preventiva | `EquipoPlanificable`, `AltaPlanificacion`, `AnexarAPlanificacion`, `AltaEquipoEnPlanificacion`, `EliminarPlanificacion`, `CerrarPlanificacionSinTrabajos`, `CerrarPlanificacionConTrabajos` | `FG.bas:2094-2174`, `2238-2396`, `2397-2524`, `2781-2849`, `2915-2980`, `2982-3190`, `3191-3374` | `TbPlanificacion`, `TbPlanificacionAnexos`, `TbPlanificacionEquipos`, `TbPlanificacionRegistrada`; `DoCmd.RunSQL` en `FG.bas:2970` |
| Actividades | `ActividadAlta`, `ActividadEliminar`, `DameHorasTecnico`, `DameHorasMaximasEnActividadesEventoPorDia` | `FG.bas:5358-5580`, `5581-5756`, `509-586`, `841-886` | `TbActividades`; validaciones documentadas en el comentario `FG.bas:5380-5395` |
| Materiales y seguimientos | `MaterialDeActividadAlta`, `EdicionMaterialEnActividad`, `EliminarMaterialEnActividad`, `EliminarMaterialSinActividad`, `MaterialIntervencionAlta`, `SeguimientoCerrado` | `FG.bas:5111-5278`, `4350-4504`, `4785-4940`, `4941-5110`, `4013-4230`, `3872-3929` | `TbMaterial`, `TbMaterialSeguimiento`; `DoCmd.RunSQL` en `FG.bas:4215`, `4900`, `5069` |
| Informes y exportación | `RellenarInformeEnWord`, `ObtenerHojaActividadesDeExportacion`, `ObtenerHojaMaterialesDeExportacion`, `ObtenerHojaEventosDeExportacion`, `GenerarConsultas`, `GenerarConsultaEventos` | `FG.bas:6833-7007`, `3376`, `3510`, `3654`, `1384-1581`, `1689-1846` | Word/Excel vía constantes de VG |
| Sesión, equipo y registro | `RegistroApertura`, `RegistroCierre`, `EnOficina`, `getMaquina`, `getUsuariodeMaquina`, `getRutaAplicacionesLocal`, `EjecutarShell` | `FG.bas:7584-7652`, `7654-7679`, `7399-7417`, `7547-7564`, `7565-7582`, `7902-7942`, `7365-7398` | `RegistroApertura` inserta en `TbAplicacionesAperturas` de la Lanzadera (`FG.bas:7623-7650`); `RegistroCierre` la actualiza con `Execute` (`FG.bas:7668-7672`) |
| UI auxiliar | `RellenaCombo`, `RellenarComboTecnicos`, `FormularioAbierto`, `AbrirEnLocal`, `RecuadrarRango` | `FG.bas:1149-1191`, `7224-7260`, `171-194`, `218-260`, `931-1059` | Controles de formulario |
| Calibraciones | `DameDiasVigenciaCalibracion`, `getCalibracionPorFecha`, `RellenarListaEquiposMedida` | `FG.bas:1247-1292`, `7719-7787`, `7293-7328` | `TbEquiposMedida` y calibraciones |

Sin detallar (pendiente): ~60 procedimientos de utilidad menor (combos, ficheros, fechas, OneDrive: `FG.bas:7789-7899`).

**Datos.** 73 `OpenRecordset` y 1 `.Execute` (`FG.bas:7672`), 4 `DoCmd.RunSQL` (`FG.bas:2970`, `4215`, `4900`, `5069`); SQL concatenado; tablas más citadas: `TbMaterial`, `TbMaterialSeguimiento`, `TbActividades`, `TbEventos`, `TbPlanificacion*`; 0 `BeginTrans`, de modo que las altas multitabla (`FG.bas:5111-5278`) no son atómicas.
**Consumidores.** `Form_FormPlanificacionGestion.cls` (80 referencias, p. ej. `Form_FormPlanificacionGestion.cls:143` `DameHTML`), `Form_FormMaterialSeguimiento.cls:658` (`EstaFranqueado("MA", …)`), `Evento.cls:2888`.
**Capacidad.** transversal (CAP-6 Planificación, CAP-7 Materiales, CAP-4 Actividades, CAP-1 sesión).
**Posibles defectos (a verificar).** Ausencia de transacciones en altas/bajas encadenadas; `DoCmd.RunSQL` precedido de `DoCmd.SetWarnings False` (`FG.bas:4214-4215`): fallos de ejecución silenciados; ids por máximo+1 (`FG.bas:440-507`) con riesgo de concurrencia; contrato de error `"#ERR|…"` basado en `InStr(…,"|")` (`FG.bas:31`, `FG.bas:761`) frágil si el dato contiene `|`.
**Pendientes.** Nombres de tabla pasados como parámetro a `Dame`/`DameID` (`FG.bas:290-330`, `371-439`) y llamadas `Dame("TbPlanificacion", …)` (`FG.bas:760`): destino resuelto por cadena en tiempo de ejecución.

## HTML
Ficha: `modules/HTML.bas` · 258 líneas · 4 procedimientos.

**Propósito.** Genera en un directorio de trabajo local un fichero `.txt`, lo guarda como `.html` y lo abre con el navegador (`HTML.bas:116-135` comentario de `DameHTML`; `HTML.bas:99-104`).

| Procedimiento | Qué hace | Cita | Efectos |
|---|---|---|---|
| `BorraHTMLs` (privado) | Borra `.html`/`.htm` del directorio local que no estén abiertos | `HTML.bas:4-25` | Borrado de ficheros |
| `DameUntxtYHtml` | Busca el primer par `HTMLn.txt`/`HTMLn.html` libre (n de 1 a 50); devuelve `"txt;html"` o `"#ERR|…"` | `HTML.bas:26-73` | Lee el sistema de ficheros |
| `HTMLENTXT` | Escribe el mensaje en el `.txt`, lo renombra a `.html` y lo abre con `Ejecutar` (ShellExecute) | `HTML.bas:74-115`, `HTML.bas:99`, `HTML.bas:104` | Crea fichero y lanza navegador |
| `DameHTML` | Compone el informe HTML de una colección (equipos o partes) con título opcional | `HTML.bas:116-251` | Informe de preventivos |

**Datos.** Sin acceso a BD; solo sistema de ficheros (FSO). **Consumidores.** `Form_FormPlanificacionGestion.cls:143` (`flag = DameHTML(colPreventivoEquipos, "Sí", "EQUIPOS")`).
**Capacidad.** CAP-6 Planificación (informes de preventivos).
**Posibles defectos (a verificar).** Límite fijo de 50 pares (`HTML.bas:45`); `BorraHTMLs` borra todo `.html` del directorio de trabajo, no solo los propios (`HTML.bas:11-14`); el resultado de `BorraHTMLs` se ignora en `HTML.bas:44`.
**Pendientes.** `m_ObjEntorno.URLDirectorioLocal` es una propiedad resuelta en `Entorno` (`HTML.bas:11`, `HTML.bas:32`).

## JsonConverter
Ficha: `modules/JsonConverter.bas` · 1399 líneas · 31 procedimientos.

**Propósito.** Biblioteca de análisis y serialización JSON (parser recursivo, `ParseJson` `JsonConverter.bas:174-191`; `ConvertToJson` `JsonConverter.bas:201-457`; fechas ISO/UTC `JsonConverter.bas:895-1040`) más tres extensiones propias que vuelcan tablas/SQL a JSON (`JsonConverter.bas:1224-1396`).

| Procedimiento | Qué hace | Cita | Datos |
|---|---|---|---|
| `ParseJson` / `json_Parse*` | Analiza texto JSON a `Dictionary`/`Collection` | `JsonConverter.bas:174-191`, `463-665` | — |
| `ConvertToJson` / `json_Encode` | Serializa objetos VBA a JSON | `JsonConverter.bas:201-457`, `680-739` | — |
| `ParseIso`, `ConvertToIso`, `ParseUtc`, `ConvertToUtc` | Conversión fecha↔ISO/UTC con API de sistema | `JsonConverter.bas:953-1040`, `895-943` | — |
| `getJSonDeTabla` | Abre una tabla con `OpenRecordset`, construye el JSON y lo devuelve ya analizado | `JsonConverter.bas:1224-1316`, `1263`, `1308` | Lectura DAO, SQL `SELECT *` (`JsonConverter.bas:1259`) |
| `getJSonDeSQL` | Igual con SQL libre | `JsonConverter.bas:1317-1365`, `1338` | Lectura DAO |
| `GenerarTextoJSON`, `TextoParsedoParaJSON/Txt` | Escape y construcción de texto JSON | `JsonConverter.bas:1366-1396`, `1186-1222` | — |

**Datos.** Solo lectura DAO en `getJSonDeTabla`/`getJSonDeSQL`; sin transacciones. **Consumidores.** Las únicas referencias a estas funciones en `src` están dentro del propio módulo (`JsonConverter.bas:1308`); no se encuentran llamadas desde formularios o clases en el clon.
**Capacidad.** transversal (infraestructura; sin uso funcional trazado).
**Posibles defectos (a verificar).** Código aparentemente sin consumidores (`getJSonDeTabla`, `getJSonDeSQL`); mezcla de librería de terceros con funciones propias en un mismo módulo.
**Pendientes.** Ninguna referencia dinámica sin resolver identificada fuera del módulo.

## Módulo1
Ficha: `modules/Módulo1.bas` · 181 líneas · 4 procedimientos.

**Propósito.** Regulariza datos históricos: para los eventos-equipo de medida sin calibración asociada, busca la calibración vigente en la fecha del alta del evento y la graba (`Módulo1.bas:5-44`).

| Procedimiento | Qué hace | Cita | Datos |
|---|---|---|---|
| `RegularizarCalibracionesEqMedida` | Itera los registros devueltos y, si `CalibracionEnAltaEvento` no es nulo, actualiza la fila | `Módulo1.bas:5-44` | — |
| `RegistrarCalibracionEnNuevaTbEVEQ` (privado) | `UPDATE TbEventosEquipoMedida SET IDCalibracion=… WHERE IDEventoEquipoMedida=…` concatenado | `Módulo1.bas:45-72`, SQL en `Módulo1.bas:59-61`, ejecución `Módulo1.bas:62` | Escritura con `getdb().Execute`, sin transacción |
| `getEquiposMedidaEnEventosAntes1` (privado) | Join `TbEventosEquipoMedida` ↔ `TbEventosEquipoMedida_antes` con `IDCalibracion Is Null` | `Módulo1.bas:126-180`, SQL `Módulo1.bas:137-141` | Lectura |
| `getEquiposMedidaEnEventosAntes` (privado) | Lee toda `TbEventosEquipoMedida_antes`; sin llamadas en el módulo | `Módulo1.bas:73-124` | Lectura |

**Consumidores.** No se encuentran llamadas a `RegularizarCalibracionesEqMedida` fuera del módulo (búsqueda en `src`). **Capacidad.** CAP-5.
**Posibles defectos (a verificar).** Utilidad de migración puntual sin invocación localizada; `getEquiposMedidaEnEventosAntes` sin uso (`Módulo1.bas:73`); sin transacción entre lecturas y updates.
**Pendientes.** `EventoEquipoMedidaAntes.CalibracionEnAltaEvento` y `ColCampos` se resuelven en clase (`Módulo1.bas:24`, `Módulo1.bas:95`).

## Módulo2
Ficha: `modules/Módulo2.bas` · 448 líneas · 6 procedimientos.

**Propósito.** Contiene el código de edición de campos SLA de un evento, escrito como si fuera el módulo del formulario (`Me.`, `Form_Load`; cabecera `Módulo2.bas:5-8`); difiere en 35 líneas de `forms/Form_FormEventoSLA.cls` (450 líneas).

| Procedimiento | Qué hace | Cita | Datos |
|---|---|---|---|
| `Form_Load` | Lee `OpenArgs` (id de evento), bloquea si `EstaFranqueado("EV", id)` es «Sí» y carga datos | `Módulo2.bas:41-90` | Llama a `EstaFranqueado` |
| `CargarDatosEvento` | `SELECT TbEventos.*` por `IDEvento` concatenado y vuelca a controles | `Módulo2.bas:96-205`, SQL `Módulo2.bas:111-113` | Lectura DAO |
| `cmdGuardar_Click` | Valida con `SLAValidator.ValidarCamposSLA(p_EnFranqueo:=False…)` y llama a `GuardarSLA` | `Módulo2.bas:211-257`, validación `Módulo2.bas:222` | — |
| `GuardarSLA` | Reaplica la guarda de franqueo y actualiza `TbEventos` con `Recordset.Edit/.Update` | `Módulo2.bas:272-447`, `Módulo2.bas:343`, `Módulo2.bas:351`, `Módulo2.bas:429` | Escritura DAO, sin transacción |
| `LimpiarCamposSLAUI`, `cmdCerrar_Click` | Limpieza de controles y cierre sin guardar | `Módulo2.bas:17-35`, `263-266` | — |

**Consumidores.** Ninguno como módulo estándar (`Me.` no compila fuera de un módulo de formulario); la funcionalidad real la ejecuta `Form_FormEventoSLA.cls` (franqueo en `Form_FormEventoSLA.cls:57` y `296`; en este módulo `Módulo2.bas:56`, `Módulo2.bas:294`). **Capacidad.** CAP-3 Eventos.
**Posibles defectos (a verificar).** Copia casi idéntica del formulario en un módulo estándar (riesgo de divergencia); `GuardarSLA` usa `CurrentDb()` en vez de `getdb()` (`Módulo2.bas:343`).
**Pendientes.** Nombres de controles (`Me.chkUrgente`, `Me.cmbTipoReparacion`…) son referencias a controles del formulario (`Módulo2.bas:225-233`), a contrastar con `Form_FormEventoSLA.form.txt`.

## SLAValidator
Ficha: `modules/SLAValidator.bas` · 166 líneas · 2 procedimientos.

**Propósito.** Valida campos SLA de un evento y devuelve el primer código de error más un mensaje acumulado (`SLAValidator.bas:22-161`).

| Procedimiento | Qué hace | Cita | Datos |
|---|---|---|---|
| `NormalizarSiNo` (privado) | Normaliza «Sí/No» a `SI`/`NO` (quita tilde, mayúsculas) | `SLAValidator.bas:5-20` | — |
| `ValidarCamposSLA` | Reglas generales (fin sin inicio → `SLA-001`; orden de fechas → `SLA-007`; restablecimiento sin servicio afectado → `SLA-001`) y reglas de franqueo (`SLA-002`..`SLA-005`); error inesperado → `SLA-999` | `SLA-001/007`: `SLAValidator.bas:69-110`; franqueo: `SLAValidator.bas:112-140`; `SLA-999`: `SLAValidator.bas:153-160` | Pura, sin acceso a datos |

**Consumidores.** `Módulo2.bas:222` (con `p_EnFranqueo:=False`). No se localiza ninguna llamada con `p_EnFranqueo:=True`, por lo que las reglas `SLA-002`..`SLA-005` (`SLAValidator.bas:112-140`) no tienen consumidor trazado. **Capacidad.** CAP-3 Eventos (SLA).
**Posibles defectos (a verificar).** Rama de franqueo sin invocación localizada; el parámetro `p_Urgente` se copia pero no se valida (`SLAValidator.bas:30`, `SLAValidator.bas:63`).
**Pendientes.** Ninguna.

## SLAReportService
Ficha: `modules/SLAReportService.bas` · 629 líneas · 6 procedimientos.

**Propósito.** Servicio de datos y cálculo del informe SLA de eventos franqueados (`SLAReportService.bas:1-13`).

| Procedimiento | Qué hace | Cita | Datos |
|---|---|---|---|
| `GetTRSSObjetivo` | Objetivo en días del TRSS: criticidad 1→1, 3→5, 5→15, otra→0 | `SLAReportService.bas:39-46` | — |
| `GetTRSSDiasTranscurridos` (privado) | Días entre recepción y restablecimiento; -1 si fechas inválidas; conmutador `TRSS_USA_DIAS_LABORABLES = False` | `SLAReportService.bas:62-75`, `SLAReportService.bas:20` | — |
| `ConstruirSQLEventosFranqueados` | Valida fechas y compone el `SELECT` de `TbEventos` con `Franqueado = True` y rango por `FECHAALTAEVENTO` | `SLAReportService.bas:93-162`, filtro `SLAReportService.bas:149-151` | SQL concatenado con fechas formateadas |
| `ObtenerEventosFranqueados` | Abre el `Recordset` con `CurrentDb().OpenRecordset` | `SLAReportService.bas:178-205`, `SLAReportService.bas:197` | Lectura |
| `CalcularCumplimientoSLA` | Calcula cumplimiento TRES/TRCM/TRSS/global sobre los campos del recordset (con `On Error Resume Next`) | `SLAReportService.bas:220-304` | — |
| `ObtenerDetalleCumplimientoEvento` | Localiza un evento por `Bookmark` y devuelve un `Dictionary` con el detalle | `SLAReportService.bas:326-627` | — |

**Consumidores.** `SLAHTMLService.bas:776`, `SLAHTMLService.bas:949`, `SLAHTMLService.bas:1160`; `Form_FormInformeSLA.cls` (`ConstruirSQLEventosFranqueados`, formulario `Form_FormInformeSLA.form.txt:1052`). **Capacidad.** CAP-10 Informes.
**Posibles defectos (a verificar).** `On Error Resume Next` en `CalcularCumplimientoSLA` (`SLAReportService.bas:235`) puede ocultar fallos; TRSS en días corridos pese al conmutador a laborables (`SLAReportService.bas:20`); `CurrentDb()` en vez de `getdb()` (`SLAReportService.bas:197`).
**Pendientes.** Nombres de campos leídos por `p_Fields("…")` desde el recordset (`SLAReportService.bas:220-304`).

## SLAHTMLService
Ficha: `modules/SLAHTMLService.bas` · 1163 líneas · 11 procedimientos.

**Propósito.** Genera el informe HTML corporativo de SLA con CSS Mística, pestañas, gráficos de donut en SVG y panel de detalle por evento (`SLAHTMLService.bas:738-1146`).

| Procedimiento | Qué hace | Cita | Efectos |
|---|---|---|---|
| `GenerarInformeHTML_Corporate` | Orquesta el informe: construye SQL con `SLAReportService`, itera eventos y ensambla HTML | `SLAHTMLService.bas:738-1146`, `SLAHTMLService.bas:776` | Devuelve el HTML (cadena) |
| `GenerarEventDetailPanel` | Panel de detalle por evento | `SLAHTMLService.bas:455-732` | — |
| `GenerarTablaEventosHTML`, `GenerarTabReglasSLA`, `GenerarDonutChartsCard`, `GenerarHeaderHTML` | Secciones del informe | `SLAHTMLService.bas:349-447`, `202-265`, `276-339`, `180-193` | — |
| `GetCSS_MISTICA`, `f_DonutSVG`, `f_EscapeHTML`, `f_DonutLegendItem` | Estilos, SVG y escape | `SLAHTMLService.bas:16-83`, `93-127`, `137-148`, `157-172` | — |
| `CalcularCumplimiento` (privado) | Delegación en `SLAReportService.CalcularCumplimientoSLA` | `SLAHTMLService.bas:1153-1161` | — |

**Datos.** Lectura mediante `SLAReportService` (sin acceso directo a tablas: sin `Tb…` en el módulo). **Consumidores.** `Form_FormInformeSLA.cls:656` (`SLAHTMLService.GenerarInformeHTML_Corporate(strFechaInicio, strFechaFin, strTextoError)`). **Capacidad.** CAP-10 Informes.
**Posibles defectos (a verificar).** Función de 409 líneas (`SLAHTMLService.bas:738-1146`); HTML construido por concatenación con escape solo donde se llama a `f_EscapeHTML` (`SLAHTMLService.bas:137-148`).
**Pendientes.** Claves leídas del diccionario de detalle (`p_dictDetalle.Exists`, `SLAHTMLService.bas:469-714`) definidas en `ObtenerDetalleCumplimientoEvento`.

## modMigracionesSLA
Ficha: `modules/modMigracionesSLA.bas` · 199 líneas · 5 procedimientos.

**Propósito.** Migración idempotente que añade a `TbEventos` los 12 campos SLA si no existen (`modMigracionesSLA.bas:19-97`, lista en `modMigracionesSLA.bas:103-120`).

| Procedimiento | Qué hace | Cita | Efectos |
|---|---|---|---|
| `MigrarCamposSLA` | Comprueba la tabla y crea cada campo con `CreateField`/`Fields.Append` | `modMigracionesSLA.bas:19-97`, `modMigracionesSLA.bas:64-70` | DDL DAO sobre `TbEventos` |
| `GetCamposSLA` (privado) | Define campos: fechas, booleanos y `TipoReparacion` (texto 50) | `modMigracionesSLA.bas:103-120` | — |
| `TablaExiste`, `CampoExisteEnTabla` | Comprobación por captura de error | `modMigracionesSLA.bas:126-133`, `139-146` | — |
| `VerificarCamposSLA` | Lista campos faltantes | `modMigracionesSLA.bas:153-195` | Lectura de metadatos |

**Consumidores.** No se localizan llamadas fuera del módulo (búsqueda en `src`). **Capacidad.** CAP-3 Eventos (esquema SLA).
**Posibles defectos (a verificar).** Migración manual sin invocación localizada; declara `rs` sin uso (`modMigracionesSLA.bas:155`); `MigrarCamposSLA` no maneja `Resume` y el booleano resultante no distingue «creado» de «ya existía».
**Pendientes.** Ninguna.

## No aplicables
| Módulo | Motivo |
|---|---|
| `Pruebas.bas` | Módulo de pruebas |
| `PruebasSLA.bas` | Módulo de pruebas |
| `modTestingSLA.bas` | Módulo de pruebas |
| `Test_SLAReportService.bas` | Módulo de pruebas |
| `Test_SLAValidator.bas` | Módulo de pruebas |
