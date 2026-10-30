# Clases de Brass (1 de 2)

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **PRD:** [README](README.md)

Inventario funcional de 14 clases de `src/classes` (Actividad, Anexo, Entorno, Equipo, EquipoMedida, EquipoMedidaCalibracion, Evento, EventoBusqueda, EventoEquipoMedida, EventoEquipoMedidaAntes, Factura, FacturaPerfiles, Fiesta, Gasto). Las rutas son relativas a `src/classes/` salvo abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`, `VG.bas` = `modules/Variables Globales.bas`, `Constructor.bas` = `modules/Constructor.bas`. No hay ningún `BeginTrans` en el código: ninguna operación de estas clases es transaccional. Patrón común: propiedades de entidad públicas, `ColCampos`/`getPropiedad`/`SetPropiedad` por nombre (hidratación desde `Constructor.bas`) y errores por parámetro `p_Error` con `Err.Raise 1000`.

## Actividad
`Actividad.cls` · 623 líneas · 15 procedimientos
**Propósito:** entidad de una actividad de técnico sobre un evento; carga perezosa de evento, factura, técnico, materiales y anexos (`Actividad.cls:224`, `:253`, `:316`, `:192`, `:84`).

| Procedimiento | Qué hace | Cita | Datos/efectos |
|---|---|---|---|
| `IDActividadCalculada` | Calcula el id con alias de técnico y fecha de alta | `Actividad.cls:375` | `DameIDActividad` (FG) |
| `ColAnexos` / `URLAnexos` | Anexos de la actividad; concatena URLs con `;` | `Actividad.cls:84-96`, `:46-74` | `Constructor.getAnexosDeActividad` |
| `AnexosCerrados` | Sí si ningún anexo existente está abierto | `Actividad.cls:117` | `FicheroAbierto` |
| `ColMateriales` | Materiales de la actividad | `Actividad.cls:192-207` | `Constructor.getMaterialesEnActividad` |
| `FechaFacturada` | Fecha de la factura asociada | `Actividad.cls:282-305` | lectura vía `Factura` |
| `EliminarAnexos` | Borra todos los anexos si están cerrados | `Actividad.cls:393-433` | `Anexo.Eliminar` |
| `getPropiedad`/`SetPropiedad` | Acceso por nombre a 15 campos | `Actividad.cls:475`, `:549` | memoria |

**Datos:** sin SQL propio; todo vía `Constructor.bas`. **Consumidores:** `Form_FormActividadDeEventoPorLotes.cls:249`, `Form_FormEventosAltaMasivaPorNS.cls:370` (`New Actividad`). **Capacidad:** CAP-4.
**Posibles defectos (a verificar):** `FechaFacturada` entra en la rama `Factura Is Nothing` y desreferencia `Me.Factura.FechaFactura` (`Actividad.cls:293-298`): la condición parece invertida. Mensajes de error dicen «Actividade» (`Actividad.cls:540`).
**Pendientes:** ninguno.

## Anexo
`Anexo.cls` · 1083 líneas · 31 procedimientos
**Propósito:** documento adjunto a evento, actividad, material, seguimiento de material, subcontratación o gasto, con copia física del fichero (`Anexo.cls:44-68`, `:690-753`).

| Procedimiento | Qué hace | Cita | Datos/efectos |
|---|---|---|---|
| `TipoAnexo` | Deduce el tipo por el primer ID informado (orden evento, actividad, material, seguimiento, gasto, subcontratación) | `Anexo.cls:51-68` | memoria |
| `URLAnexo` | Directorio de gastos si hay `IDGasto`, si no el de documentación, más `NombreArchivo` | `Anexo.cls:136-162` | `Entorno` |
| `GetURLAnexoCalculado` | Nombre `<id>_<ordinal 00>.<ext>`; ordinal = nº de anexos de la entidad + 1 | `Anexo.cls:384-431` | `TbAnexos` (SQL concatenado) |
| `MotivoNoOK` | Valida fichero local cerrado, tipo, entidad, destino no abierto, título obligatorio y no repetido | `Anexo.cls:607-685` | lectura |
| `Registrar` | Inserta en `TbAnexos` y después copia el fichero | `Anexo.cls:690-753` | DAO `AddNew` + `fso.CopyFile` |
| `RegistrarAutorizacionMaterial` | Crea o reutiliza el anexo «Autorización Materiales» de un material | `Anexo.cls:757-814` | `TbAnexos`, copia |
| `Resubir` | Sustituye el fichero; actualiza `NombreArchivo` si cambia la extensión | `Anexo.cls:818-852` | `UPDATE TbAnexos` |
| `Eliminar` | Borra la fila y después el fichero | `Anexo.cls:912-936` | `DELETE TbAnexos` |
| `TituloRepetido` | Compara título dentro de la misma entidad; reintenta con ID numérico | `Anexo.cls:856-907` | lectura |

**Datos:** `TbAnexos` por DAO y SQL concatenado (`Anexo.cls:580-586`, `:880-891`); sistema de ficheros por `fso`. **Consumidores:** `Form_FormAnexos.cls:57` (`Registrar`), `:290` (`Resubir`), `Gasto.cls:274`, `Actividad.cls:424` (`Eliminar`). **Capacidad:** CAP-11.
**Posibles defectos (a verificar):** `Registrar` usa `Me.IDAnexoCalculado` (`Anexo.cls:718`) en vez del `IDAnexo` ya calculado en `MotivoNoOK` (`:650`); el ordinal se basa en un recuento, por lo que tras borrar un anexo puede repetirse un nombre (`:413-427`); la fila se inserta antes de copiar el fichero (`:717-745`), sin transacción; `Eliminar` borra la fila antes del fichero (`:927-931`).
**Pendientes:** `getPropiedad(m_NombreCampoID)` resuelve campos por nombre (`Anexo.cls:405`, `:876`).

## Entorno
`Entorno.cls` · 1207 líneas · 37 procedimientos
**Propósito:** objeto de entorno (`m_ObjEntorno`) con rutas, modo local/remoto y colecciones de catálogo con caché perezosa (`VG.bas:298`, `Entorno.cls:107-163`).

| Grupo | Qué hace | Cita | Datos/efectos |
|---|---|---|---|
| `DatosEnLocal` y `URLDirectorio*` (17 propiedades) | Rutas según `TempVars("DatosEnLocal")`; comprueban existencia de carpeta con `fso` | `Entorno.cls:107-163`, `:288-314` | sistema de ficheros |
| `CSS` / `URLArchivoCSS` | Lee el CSS completo desde fichero | `Entorno.cls:221-251` | lectura |
| `URLAchivoIni` / `VersionAplicacion` | Fichero `.ini` junto al proyecto o en recursos; versión leída con `LeerIni` | `Entorno.cls:614-656` | lectura |
| `DiasParaAvisoDePlanificaciones` | Valor fijo «15» | `Entorno.cls:733` | constante |
| `ColTipoAnexos` | Tabla fija 1-6 (Evento…Gasto) | `Entorno.cls:757-764` | constante |
| `ColPerfilesActivos`, `ColTecnicosActivos`, `ColEquiposMedida(Activos)`, `Centros` | Cargan catálogos | `Entorno.cls:677`, `:703`, `:787`, `:822`, `:858` | `Constructor.get*` |
| `UsuarioConectadoMaquina` | Usuario de la máquina | `Entorno.cls:61` | `Constructor.getUsuarioConectadoPorMaquina` |
| `getValorPropiedad` | Despacho por nombre de ~26 propiedades | `Entorno.cls:929-1156` | memoria |
| `RellenarColesConCampos` | Carga nombres de campos de informe | `Entorno.cls:1166-1176` | lee `TbAuxNombreCamposParaInforme` (`CurrentDb`) |

**Datos:** solo lectura, sin SQL concatenado. **Consumidores:** `VG.bas:298`, `Form_FormInformesDesdeLista.cls:235`, `Form_FormInformesPorFactura.cls:262`, `Anexo.cls:152`. **Capacidad:** CAP-2 (y transversal).
**Posibles defectos (a verificar):** `URLDirectorioAnexos` se calcula como `<aplicación>ANEXOS\` (`Entorno.cls:298`) pero `Anexo.URLAnexo` usa otros directorios (`Anexo.cls:152-154`); `RellenarColesConCampos` usa `CurrentDb` y no `getdb` (`Entorno.cls:1176`).
**Pendientes:** nombres de propiedad en `getValorPropiedad` (`Entorno.cls:929`) resueltos por cadena.

## Equipo
`Equipo.cls` · 189 líneas · 5 procedimientos
**Propósito:** equipo (nodo/BUI/subsistema) de la estructura; activo si no tiene `FechaObsoleto` (`Equipo.cls:48-59`).

| Procedimiento | Qué hace | Cita | Datos/efectos |
|---|---|---|---|
| `IDEquipoCalculado` | Siguiente id de `TbEquipo` | `Equipo.cls:29-35` | `DameID` |
| `EquipoActivo` | Sí si `FechaObsoleto` no es fecha | `Equipo.cls:48-59` | memoria |
| `getPropiedad`/`SetPropiedad` | Acceso por nombre | `Equipo.cls:101`, `:145` | memoria |

**Datos:** sin altas ni bajas en la clase. **Consumidores:** `Constructor.bas:2348` (`getEquipo`), `Form_FormEquiposMedidaGestion.cls:103`, `:142`. **Capacidad:** CAP-12.
**Posibles defectos (a verificar):** ninguno.
**Pendientes:** ninguno.

## EquipoMedida
`EquipoMedida.cls` · 671 líneas · 16 procedimientos
**Propósito:** equipo de medida con estado de calibración derivado de la última calibración (`EquipoMedida.cls:68-97`).

| Procedimiento | Qué hace | Cita | Datos/efectos |
|---|---|---|---|
| `EstadoCalibracion` | «Fuera de servicio» si hay `FechaFinServicio`; «Sin Calibraciones»; «Calibrado» si fin ≥ hoy; si no «Fuera de Calibración» | `EquipoMedida.cls:78-93` | `Constructor.getUltimaCalibracion` |
| `Registrar` | Valida nombre, NS, PN, fecha inicio, fechas coherentes; alta o edición | `EquipoMedida.cls:294-449` | DAO sobre `TbEquiposMedida` |
| `NSRepetido` | Busca el NS en `TbEquiposMedida` | `EquipoMedida.cls:503-532` | SQL concatenado |
| `Eliminar` | Impide borrar si figura en algún evento | `EquipoMedida.cls:469-494` | `DELETE TbEquiposMedida` |
| `Eventos` / `Calibraciones` | Eventos y calibraciones del equipo | `EquipoMedida.cls:136`, `:208` | `Constructor` |
| `getEquipoMedidaEnActividad` | Lee `TbEventosEquipoMedida` por id o evento+equipo | `EquipoMedida.cls:243-284` | lectura |

**Datos:** `TbEquiposMedida`, `TbEventosEquipoMedida` (`EquipoMedida.cls:260-265`). Tras editar con cambio de `FechaFinServicio` recarga `ColEquiposMedidaActivos` (`:452-454`). **Consumidores:** `Form_FormEquipoMedida.cls:145` (`Registrar`). **Capacidad:** CAP-5.
**Posibles defectos (a verificar):** el NS repetido solo se comprueba en edición y si cambia el NS (`EquipoMedida.cls:339-341`), no en alta; la condición de `getEquipoMedidaEnActividad` (`:256`) exige id aunque el SQL admite evento+equipo.
**Pendientes:** ninguno.

## EquipoMedidaCalibracion
`EquipoMedidaCalibracion.cls` · 299 líneas · 7 procedimientos
**Propósito:** calibración de un equipo de medida (`EquipoMedidaCalibracion.cls:78-167`).

| Procedimiento | Qué hace | Cita | Datos/efectos |
|---|---|---|---|
| `Registrar` | Exige `ParaAlta`, fechas obligatorias y fin ≥ inicio; fija la empresa calibradora a un literal del código; alta o edición | `EquipoMedidaCalibracion.cls:78-159` | DAO `TbEquiposMedidaCalibraciones` |
| `Eliminar` | Borra por `IDCalibracion` | `EquipoMedidaCalibracion.cls:169-183` | `DELETE` concatenado |
| `IDCalibracionCalculado` | Siguiente id | `EquipoMedidaCalibracion.cls:28-34` | `DameID` |

**Datos:** `TbEquiposMedidaCalibraciones`. **Consumidores:** `Form_FormEquipoMedidaCalibracion.cls:39`, `Constructor.bas:5044`. **Capacidad:** CAP-5.
**Posibles defectos (a verificar):** `EmpresaCalibradora` se sobrescribe siempre con un literal (`EquipoMedidaCalibracion.cls:123`), ignorando el valor recibido; `Eliminar` no comprueba uso en eventos.
**Pendientes:** ninguno.

## Evento
`Evento.cls` · 3015 líneas · 60 procedimientos (≈38 sin detallar: propiedades de carga perezosa, marcadas «pendiente»)
**Propósito:** entidad central (avería/incidencia sobre un equipo) con alta/edición, franqueo (cierre), equipos de medida, actividades por lotes y validación SLA (`Evento.cls:1477-1767`, `:1255-1313`).

| Grupo / procedimiento | Qué hace | Cita | Datos/efectos |
|---|---|---|---|
| `MotivoNoOK` | Edición: debe existir, no facturado ni franqueado; obligatorios NODO, BUI, SubSistema, Equipo, TIPOEVENTO, CRITICIDAD, ALIASTECNICO, ORIGINADOR, Descripcion; técnico registrado, no de baja y con precio/hora activo | `Evento.cls:1477-1612` | lectura |
| `Registrar` | Alta (`AddNew`) o edición (`Edit`) de `TbEventos`; `TIEMPORESPUESTAEVENTO` vacío pasa a «0» | `Evento.cls:1617-1767` | DAO `TbEventos` |
| `Modificable` | No modificable si facturado o con fecha final | `Evento.cls:428-462` | memoria |
| `MotivosNoFranqueable` | Exige ≥1 actividad y materiales con recepción | `Evento.cls:1172-1216` | lectura |
| `Franquear` | Valida, valida SLA, exige causa de fin; fecha final = mayor `FechaAlta` de actividades (`FechaFinalCalculada`, `:214-256`); hora fija; `UPDATE TbEventos` | `Evento.cls:1255-1313` | `UPDATE` concatenado |
| `QuitarFranqueo` | No permite si está en factura o no está franqueado; anula fecha, hora y causa | `Evento.cls:1220-1250`, `:1409-1440` | `UPDATE` concatenado |
| `ValidarSLA` / `MotivosNoFranqueableSLA` | Delegan en `SLAValidator.ValidarCamposSLA`; exige «urgente» informado | `Evento.cls:1367-1405`, `:1315-1350` | memoria |
| `CalcularTRES` / `CumpleTRES` | Horas entre recepción y contacto; objetivo por criticidad (1-5) | `Evento.cls:2739-2765`, `:2782-2816` | memoria |
| `GuardarSLA` | Actualiza campos SLA de `TbEventos` si no está franqueado | `Evento.cls:2876-3002` | DAO (`CurrentDb`) |
| `InformeRACRegistrarFecha` / `QuitarFecha` | Fija o anula `FechaEnInformeRAC` | `Evento.cls:1442-1475` | `UPDATE` concatenado |
| `RegistrarEquipoMedida` | Bloquea si facturado o franqueado; asocia equipo y calibración vigente | `Evento.cls:1859-1925` | lee `TbEventosEquipoMedida` |
| `EliminarEquipoMedida(s)` | Borra relaciones evento-equipo | `Evento.cls:1929-1995` | `DELETE` concatenado |
| `GenerarActividadesPorLotes` | Vacía `TbActividadesLotes`, la rellena y vuelca a `TbActividades` | `Evento.cls:2044-2130` | `DELETE`, DAO, `INSERT … SELECT` |
| `getColFechasHoras`, `getHorasSinRealizarTecnicoEnDia`, `getHorasMaximasEnActividadesEventoPorDia` | Cálculo de horas por día y técnico | `Evento.cls:2306`, `:2263`, `:1998` | lectura de `TbActividades` (`:2006`) |
| `EliminarAnexos` | Borra anexos si están cerrados | `Evento.cls:2380` | `Anexo.Eliminar` |
| `ColCampos`/`getPropiedad`/`SetPropiedad` | Acceso por nombre | `Evento.cls:2424`, `:2473`, `:2601` | memoria |
| Resto (`TieneActividades`, `ColMateriales`, `Factura`, `Parte`, `HorasLabTotales`, …) | Cálculos y carga perezosa | `Evento.cls:96-1114` | pendiente |

**Datos:** `TbEventos`, `TbEventosEquipoMedida`, `TbActividades`, `TbActividadesLotes`; todo por DAO o SQL concatenado con `getdb`/`CurrentDb`; ninguna transacción. **Consumidores:** `Form_FormEventoFranqueo.cls:21`, `Form_FormEventoGestion.cls:2006` (`QuitarFranqueo`), `Form_FormActividadDeEventoPorLotes.cls:271`, `Form_FormEventoAlta.cls:782`. **Capacidad:** CAP-3 (y CAP-4/CAP-5 en lotes y equipos de medida).
**Posibles defectos (a verificar):** `ValidacionEliminar` está comentada entera y hace referencia a «Plan de Contingencia» (`Evento.cls:1771-1800`); la hora de franqueo es un literal y la variable `m_HoraFin` no se usa en el SQL (`:1297-1300`); `RellenarTbActividadesLotes` vacía una tabla compartida sin distinguir usuarios (`:2061`); `Franquear` actualiza BD y luego el objeto sin transacción; objetivos de `CumpleTRES` fijos en el código (`:2801-2813`).
**Pendientes:** `SLAValidator.ValidarCamposSLA` (`Evento.cls:1385`) y `EstaFranqueado` (`:2888`) son módulos externos a la clase, no analizados aquí.

## EventoBusqueda
`EventoBusqueda.cls` · 21 líneas · 0 procedimientos
**Propósito:** contenedor de criterios de búsqueda de eventos (ID, BUI, subsistema, equipo, tipo, fechas, franqueado) (`EventoBusqueda.cls:13-20`).
**Datos:** ninguno. **Consumidores:** `Form_FormEventoGestion.cls:1869-1886` (se pasa a `getEventosFiltrados`). **Capacidad:** CAP-3.
**Posibles defectos (a verificar):** ninguno. **Pendientes:** ninguno.

## EventoEquipoMedida
`EventoEquipoMedida.cls` · 200 líneas · 8 procedimientos
**Propósito:** relación evento-equipo de medida con la calibración aplicada (`EventoEquipoMedida.cls:14-17`, `:46-61`).
Procedimientos: `Calibracion` y `EquipoMedida` (carga perezosa por `Constructor`, `:56`, `:89`), `IDEventoEquipoMedidaCalculado` (`:27-34`), `getPropiedad`/`SetPropiedad` (`:140`, `:171`).
**Datos:** sin SQL propio. **Consumidores:** `Constructor.bas:4947`, `EquipoMedida.cls:275`. **Capacidad:** CAP-5.
**Posibles defectos (a verificar):** `DameID("TbEquiposEquipoMedida", …)` (`EventoEquipoMedida.cls:34`) apunta a un nombre distinto de `TbEventosEquipoMedida` usado en otras clases.
**Pendientes:** ninguno.

## EventoEquipoMedidaAntes
`EventoEquipoMedidaAntes.cls` · 185 líneas · 7 procedimientos
**Propósito:** fila del histórico previo `TbEventosEquipoMedida_antes`, usada para migrar la calibración vigente en el alta del evento (`EventoEquipoMedidaAntes.cls:76-91`; `Módulo1.bas:74-90`).
`CalibracionEnAltaEvento` obtiene la calibración por fecha con `getCalibracionPorFecha` (`EventoEquipoMedidaAntes.cls:87`).
**Datos:** lectura de `TbEventosEquipoMedida_antes` en `Módulo1.bas`. **Consumidores:** `Módulo1.bas:94`, `:150`, `:24`. **Capacidad:** CAP-5 (migración).
**Posibles defectos (a verificar):** ninguno. **Pendientes:** `getCalibracionPorFecha` no resuelta en esta clase.

## Factura
`Factura.cls` · 6827 líneas · 98 procedimientos (≈45 sin detallar: propiedades `Precio*`/`col*`, hojas de Excel; «pendiente»)
**Propósito:** factura de un periodo: calcula importes, valida lo facturable, graba cabecera y tablas «involucradas» y genera salidas (Excel, Word, partes) (`Factura.cls:2006-2406`, `:3311-3384`).

| Grupo / procedimiento | Qué hace | Cita | Datos/efectos |
|---|---|---|---|
| `TipoImpositivoCalculado`, `RecargoCalculado` | Devuelven valores fijos del código | `Factura.cls:184-194` | constantes |
| `PrecioHorasCalculado` y `Precio*Calculado/PorFacturar` | Suman horas × precio por perfil, dietas y gastos, suministro, subcontrataciones, recargo, base e impuestos | `Factura.cls:200-262`, `:1862-2005` | memoria |
| `col*` / `col*PorFacturar` | Colecciones de eventos, actividades, materiales, gastos, subcontrataciones y partes | `Factura.cls:1049-1660` | `Constructor`/`SQLEventos` |
| `ColPerfilesEnFactura` / `getColTipos` | Agrupa horas por perfil y precio | `Factura.cls:1632`, `:1662-1805` | crea `FacturaPerfiles` (`:1805`) |
| `ValidacionFactura` | Exige fecha final, nuevo `IDFactura`, recargo, tipo impositivo; verifica que evento, equipo, técnico y actividad de cada elemento se resuelvan | `Factura.cls:2006-2402` | lectura |
| `EliminarFacturasAnteriores` | Borra las tablas hijas y la cabecera de la factura | `Factura.cls:2407-2450` | `DELETE` concatenado |
| `GrabarFactura` | Exige algo que facturar; borra previas y graba | `Factura.cls:3311-3384` | orquestación |
| `GrabarTablasRelacionadas` | Graba cabecera, eventos, actividades, materiales, gastos, subcontrataciones y perfiles; si falla, borra la cabecera | `Factura.cls:3205-3255` | DAO sobre 7 tablas `TbFactura*` |
| `Eliminar` | Impide si el fichero está abierto; borra hijas, cabecera y fichero | `Factura.cls:3390-3443` | `DELETE` concatenado |
| `ObtenerLI`, `GenerarExcelEventosLI/MaterialesLI/SubContratacionesLI` | Libro de listado interno | `Factura.cls:3453`, `:3590`, `:3686`, `:3836` | Excel, `SaveAs` (`:3563`) |
| `GenerarPartesDeFactura`, `EliminarPartes` | Partes por evento de la factura | `Factura.cls:4013`, `:5120` | ficheros |
| `InformeRAC` y `Rellenar*` | Informe RAC en Excel | `Factura.cls:4078-4208` | `SaveAs` (`:4208`) |
| `GenerarWordsInformeAveriasEnCarpeta` | Informes de averías en Word | `Factura.cls:5040-5076` | Word por automatización |
| `SalidaAExcel`, `CreaHoja*` | Hojas de eventos, actividades, materiales, factura y detalle | `Factura.cls:5155-6516` | Excel, `SaveAs` (`:5239`) |
| `ColCampos`/`getPropiedad`/`SetPropiedad` | Acceso por nombre | `Factura.cls:6599`, `:6641`, `:6731` | memoria |

**Datos:** `TbFacturaPrincipal`, `TbFacturaEventosInvolucrados`, `TbFacturaActividadesInvolucradas`, `TbFacturaMaterialesInvolucrados`, `TbFacturaGastosInvolucrados`, `TbFacturaSubcontratacionesInvolucradas`, `TbFacturaPrincipalPerfiles` (`Factura.cls:2464-3282`); DAO y SQL concatenado; sin transacciones. **Consumidores:** `Form_FormFacturaAlta.cls:80-84` (`ValidacionFactura`, `GrabarFactura`), `:98` (`SalidaAExcel`), `Form_FormFacturaResultado.cls:211` (`ObtenerLI`), `Form_FormFacturacionGestion.cls:212` (`Eliminar`). **Capacidad:** CAP-8 (y CAP-10 por informes).
**Posibles defectos (a verificar):** la recuperación de `GrabarTablasRelacionadas` solo borra la cabecera (`Factura.cls:3250-3252`), dejando posibles filas hijas; `GrabarFactura` borra la factura anterior antes de grabar la nueva (`:3368-3373`) sin transacción; `Eliminar` borra el fichero tras la BD (`:3436-3440`).
**Pendientes:** `SQLEventos` compone SQL dinámico (`Factura.cls:1384`); los ≈45 procedimientos de cálculo parcial y hojas no se han detallado.

## FacturaPerfiles
`FacturaPerfiles.cls` · 324 líneas · 9 procedimientos
**Propósito:** línea de factura por perfil técnico: importes = horas × precio, laborables y extra (`FacturaPerfiles.cls:61-135`).
Procedimientos: `ImporteHorasLab` (`:61-76`), `ImporteHorasExt` (`:87-102`), `Importe` suma ambos (`:113-128`), `PrecioHoraLabCalculado` (`:138-162`), `PrecioHoraExtCalculado` (`:172-196`), `TIPOTECNICO` (`:34-50`).
**Datos:** sin SQL; precios vía `TIPOTECNICO`. **Consumidores:** `Factura.cls:1805`, `Constructor.bas:997`. **Capacidad:** CAP-8.
**Posibles defectos (a verificar):** `PrecioHoraExtCalculado` lee `PrecioHoraLabActual` (`FacturaPerfiles.cls:191`), no un precio de horas extra.
**Pendientes:** ninguno.

## Fiesta
`Fiesta.cls` · 202 líneas · 6 procedimientos
**Propósito:** día festivo de técnicos (`TbTecnicosFiestas`) (`Fiesta.cls:68-99`).

| Procedimiento | Qué hace | Cita | Datos/efectos |
|---|---|---|---|
| `MotivoNoOK` | Exige fecha y observaciones; rechaza fecha ya registrada | `Fiesta.cls:30-51` | recuento con fecha concatenada |
| `Registrar` | Valida y añade fila | `Fiesta.cls:68-99` | DAO `AddNew` (`CurrentDb`) |
| `Borrar` | Borra por fecha | `Fiesta.cls:110-123` | `DELETE` concatenado |

**Datos:** `TbTecnicosFiestas`. **Consumidores:** `Form_FormTecnicoFiestas.cls:19`; consulta de festivos en `FG.bas:6729`. **Capacidad:** CAP-9.
**Posibles defectos (a verificar):** mezcla `getdb()` (`Fiesta.cls:41`) y `CurrentDb()` (`:91`, `:123`).
**Pendientes:** ninguno.

## Gasto
`Gasto.cls` · 433 líneas · 12 procedimientos
**Propósito:** gasto de un técnico, con anexos y estado de facturación (`Gasto.cls:211-241`, `:288-317`).

| Procedimiento | Qué hace | Cita | Datos/efectos |
|---|---|---|---|
| `AnexosCerrados` | Sí si ningún anexo está abierto | `Gasto.cls:36-72` | `FicheroAbierto` |
| `EliminarAnexos` | Borra anexos si todos cerrados | `Gasto.cls:243-286` | `Anexo.Eliminar` |
| `Eliminar` | Borra anexos; rechaza si facturado; `DELETE TbGastos` | `Gasto.cls:288-309` | `DELETE` concatenado |
| `Factura` / `Tecnico` | Carga perezosa | `Gasto.cls:155-173`, `:183-200` | `Constructor.getFacturaDeGasto`, `getTecnico` |

**Datos:** `TbGastos`. **Consumidores:** `Form_FormGastoGestion.cls:670`. **Capacidad:** CAP-8.
**Posibles defectos (a verificar):** `FechaFacturado` tiene la condición invertida (`Gasto.cls:221-226`); `Err.r1000` en `Tecnico` (`Gasto.cls:198`) parece un error tipográfico; `Eliminar` borra anexos antes de comprobar si está facturado (`:298-305`).
**Pendientes:** ninguno.

## No aplicables
Ninguna de las 14 clases es de prueba.
