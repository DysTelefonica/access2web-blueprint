# Clases de Brass (2 de 2)

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **PRD:** [README](README.md)

Alcance: inventario funcional de 14 clases de `src/classes` (GestorInforme, Informe, LIbranza, Material, MaterialSeguimiento, Parte, Subcontratacion, Tecnico, TecnicoHoras, TipoDia, TipoTecnico, TipoTecnicoPrecio, Usuario, UsuarioAplicacionPermisos). Abreviaturas de ruta: `Constructor.bas` = `modules/Constructor.bas`; `VG.bas` = `modules/Variables Globales.bas`; `FG.bas` = `modules/Funciones Generales.bas`; los formularios se citan como `Form_X.cls` (`forms/`). Las clases son entidades con propiedades públicas, `ColCampos`/`getPropiedad`/`SetPropiedad` (reflexión manual por nombre de campo) y errores devueltos por `p_Error` con `Err.Raise 1000` como señal interna. No hay ninguna `BeginTrans` en estas clases. Todo el SQL es concatenado (sin parámetros) y se ejecuta con DAO sobre `getdb()` o `CurrentDb()`, salvo usuarios y permisos, que van contra `getdbLanzadera()`. Las consultas de la fábrica `Constructor.*` quedan fuera de este documento salvo cita puntual.

## GestorInforme
`classes/GestorInforme.cls` · 281 líneas · 8 procedimientos públicos/privados.

**Propósito.** Registra en `Tb0FiltroGestion` una fila por cada filtrado de eventos con los campos de filtro usados, el usuario de red y un indicador «Visto» inicial «No» (`GestorInforme.cls:76-129`).

| Procedimiento | Qué hace | Cita | Datos / efectos |
|---|---|---|---|
| `IDCalculado` | Obtiene el ID con `DameID("Tb0FiltroGestion","ID")` | `GestorInforme.cls:26-43` | Lectura vía helper |
| `Registrar(p_ColDatos)` | Alta: ID, `Now()`, `Datos` (claves unidas con `\|`), `Usuario` = usuario conectado, `Visto`="No" | `GestorInforme.cls:76-129` | `AddNew` DAO en `CurrentDb()` |
| `ColADatos` (privada) | Concatena las claves del diccionario separadas por `\|` | `GestorInforme.cls:131-160` | — |
| `Borrar` | `DELETE` por `ID` sin delimitadores de fecha | `GestorInforme.cls:162-183` | SQL dinámico concatenado |
| `UsuarioObj` | Resuelve el `Usuario` por `Constructor.getUsuario` | `GestorInforme.cls:45-74` | Lectura en Lanzadera |
| `ColCampos`/`getPropiedad`/`SetPropiedad` | Reflexión de 5 campos (ID, Fecha, Datos, Usuario, Visto) | `GestorInforme.cls:185-281` | — |

**Datos.** Escribe `Tb0FiltroGestion` (DAO `AddNew` y `DELETE` concatenado). Sin transacciones.
**Consumidores.** `Form_FormEventoGestion.cls:1240` (`New GestorInforme`) y `Form_FormEventoGestion.cls:1356` (`Registrar m_ColCampos`) dentro de `FiltrarEventos` (`Form_FormEventoGestion.cls:1221`).
**Capacidad.** CAP-3 Eventos (a confirmar con CAP-10).
**Posibles defectos (a verificar).**
- `Borrar` usa `m_SQL` sin `Dim` local; resuelve a la global `VG.bas:203` (`GestorInforme.cls:162-183`).
- `Borrar` exige `IsDate(Me.ID)` y compone `WHERE ID=<valor>` sin `#`, lo que fallaría con un ID de fecha (`GestorInforme.cls:162-183`).
- Los manejadores de `IDCalculado` y `UsuarioObj` terminan con `m_Error = Me.Error`, que sobrescribe el error propio (`GestorInforme.cls:26-43`, `45-74`).
- `Registrar` devuelve cadena vacía siempre; el error solo sale por `p_Error`; el consumidor no pasa `p_Error` (`Form_FormEventoGestion.cls:1356`).
**Pendientes.** Qué consumidor lee `Visto` y quién lo pone a «Sí»: sin referencia localizada (búsqueda de `Tb0FiltroGestion` en `src` sin más lecturas).

## Informe
`classes/Informe.cls` · 987 líneas · 10 procedimientos (1 público de entrada, `PonerBordeFinal` público).

**Propósito.** Genera con Excel un libro de «Informe de partes de trabajo» a partir de una colección de eventos: una hoja por evento copiada de una plantilla, actividades, materiales y, opcionalmente, hoja resumen (`Informe.cls:23-182`).

| Procedimiento | Qué hace | Cita | Datos / efectos |
|---|---|---|---|
| `EventosRAC` | Orquesta: abre Excel, copia plantilla por evento, rellena, borde final, resumen, `SaveAs` | `Informe.cls:23-182` | Fichero local; Excel por automatización |
| `AbrirNuevoExcel` | Valida carpeta local y plantilla (`Entorno`), borra el informe previo si existe y no está abierto, abre plantilla | `Informe.cls:301-359` | `FileSystemObject`, `DeleteFile` |
| `RellenarDatosEvento` | Vuelca campos del evento según mapa de nombres de `Entorno` | `Informe.cls:361-497` | Lee propiedades de `Evento` |
| `RellenarActividades` / `GenerarFilasDeActividades` | Filas por actividad vía `getPropiedad` | `Informe.cls:529-692` | Lee `Actividad` |
| `RellenarMateriales` / `GenerarFilasDeMateriales` | Filas por material vía `getPropiedad` | `Informe.cls:693-870` | Lee `Material` |
| `GeneraHojaResumenInformeRAC`, `PonerEnlaceAResumen` | Hoja resumen por evento y enlaces internos | `Informe.cls:184-300`, `498-528` | — |
| `PonerBordeFinal` | Formato final | `Informe.cls:871-987` | Excel |

**Datos.** No accede a tablas directamente: lee objetos `Evento`, `Actividad` y `Material` ya construidos. El nombre del libro es literal y fijo y se guarda en el directorio local de `Entorno` (`Informe.cls:320-322`).
**Consumidores.** `Form_FormInformesParaRAC.cls:115-116`, `Form_FormInformesDesdeLista.cls:163-164` (ambos `New Informe` + `EventosRAC … True`). `Form_Form0BDInformes.cls` referencia ambos formularios (apertura no trazada; pendiente).
**Capacidad.** CAP-10 Informes.
**Posibles defectos (a verificar).**
- En la hoja resumen, la descripción se condiciona a `BUI <> ""` y no a `Descripcion` (`Informe.cls:247-248`).
- La comprobación de duplicados usa `Exists(m_IDMaterial)` pero la clave que se inserta es `IDMaterial|IDActividad` (`Informe.cls:113-114`): la guarda nunca detiene.
- El informe fijo en el mismo nombre de fichero impide dos informes simultáneos; si hay uno abierto aborta (`Informe.cls:323-328`).
**Pendientes.** Mapas de columnas: `m_ColParaEnumNombreCamposEnInforme` y `m_ColParaNombreCamposInformeTabla` se rellenan en `Entorno.cls:1174-1184` y `VG.bas:219-220` (contenido no inventariado aquí). Plantilla de informe: ruta en `Entorno` (no se transcribe).

## LIbranza
`classes/LIbranza.cls` · 398 líneas · 9 procedimientos.

**Propósito.** Entidad «ausencia/libranza» de un técnico en un día (tabla `TbTecnicosAusencias`) con validación, alta, edición y borrado (`LIbranza.cls:81-290`).

| Procedimiento | Qué hace | Cita | Datos / efectos |
|---|---|---|---|
| `IDAusenciaCalculada` | ID con `DameID("TbTecnicosAusencias","IDAusencia")` | `LIbranza.cls:27-45` | Lectura |
| `MotivoNoOK` | Valida alta/edición; alias, motivo y fecha obligatorios; rechaza fecha repetida por técnico | `LIbranza.cls:81-178` | `SELECT Count` concatenado (`LIbranza.cls:147-150`) |
| `Registrar` | Alta (`AddNew`) o edición (`Edit`); fija `horas` a 7,5 | `LIbranza.cls:180-262` | DAO sobre `CurrentDb()` (`LIbranza.cls:224`) |
| `Borrar` | `DELETE` por `IDAusencia` | `LIbranza.cls:264-290` | SQL concatenado |
| `Tecnico` | Resuelve el `Tecnico` | `LIbranza.cls:47-79` | — |

**Datos.** Lee y escribe `TbTecnicosAusencias`; DAO + SQL concatenado; sin transacciones.
**Consumidores.** `Form_FormTecnicoAusencias.cls:26` (`New LIbranza`), `Form_FormTecnicoAusenciasGestion.cls:104` (`getLibranza`), `Tecnico.cls:459` (`AltaAusencias`), `Constructor.bas:4171` (`getLibranzas`).
**Capacidad.** CAP-9 Técnicos, ausencias, fiestas y precios.
**Posibles defectos (a verificar).**
- La edición localiza el registro por el alias y fecha **nuevos** (`LIbranza.cls:219-222`), no por `IDAusencia` ni por los valores iniciales; si se cambia alias o fecha, la selección puede quedar vacía.
- `horas` se sobrescribe siempre con 7,5 (`LIbranza.cls:224`), ignorando cualquier valor recibido.
- `MotivoNoOK` cierra `rcdDatos` tras un posible `Nothing` (`LIbranza.cls:147-170`).
- `getLibranzas` tiene dos ramas con la misma condición, y la tercera genera SQL sin `AND` (`Constructor.bas:4135-4160`).
**Pendientes.** Ninguna referencia dinámica sin resolver.

## Material
`classes/Material.cls` · 714 líneas · 18 propiedades/funciones públicas.

**Propósito.** Entidad «material/reparación» de una actividad con navegación a actividad, equipo, factura, seguimientos y anexos y estados derivados; no tiene métodos de persistencia (`Material.cls:48-523`).

| Procedimiento | Qué hace | Cita | Datos / efectos |
|---|---|---|---|
| `ColAnexos` | Anexos vía `Constructor.getAnexosDeMaterial` | `Material.cls:48-74` | Lectura |
| `AnexosCerrados` | `No` si algún fichero de anexo existe y está abierto, propio o de seguimientos | `Material.cls:81-156` | `FileSystemObject`, `FicheroAbierto` |
| `Entregado` | `Sí` si `FechaEntrega` es fecha | `Material.cls:157-180` | — |
| `IDMaterialCalculado` | `DameIDMaterial(IDActividad)` | `Material.cls:181-198` | Lectura (`FG.bas:6588`) |
| `Actividad`, `Equipo`, `Factura`, `colSeguimientos` | Resolución perezosa | `Material.cls:199-252`, `287-315`, `348-394` | Lecturas vía `Constructor` |
| `FechaFacturado` | `FechaFactura` de la factura | `Material.cls:316-347` | — |
| `SeguimientoFinalizado` | `Sí` si algún seguimiento tiene `UltimaIntervencion = "Sí"` | `Material.cls:395-440` | — |
| `URLAnexos`, `EliminarAnexos` | Une URLs con `;`; elimina anexos si todos están cerrados | `Material.cls:441-523` | Borra ficheros vía `Anexo.Eliminar` |

**Datos.** Sin SQL propio; la lectura es de `Constructor.getMaterial*` (`Material.cls:60,212,239,302`). Las escrituras a `TbMaterial` están fuera de la clase (pendiente: `Form_FormEventoCopia.cls:423` escribe `TbMaterialSeguimiento`; `FG.bas:4896`, `5065` borran seguimientos).
**Consumidores.** `Informe.cls:827` (`getPropiedad`), `MaterialSeguimiento.cls:167` (`getMaterial`).
**Capacidad.** CAP-7 Materiales, reparaciones y repuestos.
**Posibles defectos (a verificar).**
- `TieneSeguimientos` está declarada `As String` y devuelve un `EnumSino` (`Material.cls:253-286`).
- `URLAnexos` no reinicia `m_sURLAnexos`: una segunda llamada vuelve a concatenar (`Material.cls:441-478`); igual en `MaterialSeguimiento.cls:115-152` y `Subcontratacion.cls:112-149`.
- El mensaje de error de `AnexosCerrados` dice «Parte.AnexosCerrados» (`Material.cls:151-154`).
**Pendientes.** Persistencia de `TbMaterial`: sin escritura localizada en estas clases.

## MaterialSeguimiento
`classes/MaterialSeguimiento.cls` · 378 líneas · 9 procedimientos.

**Propósito.** Entidad «seguimiento de reparación» de un material (intervenciones con destino, motivo, resultado, anexo) sin persistencia propia (`MaterialSeguimiento.cls:153-226`).

| Procedimiento | Qué hace | Cita | Datos / efectos |
|---|---|---|---|
| `ColAnexos` | Anexos vía `getAnexosDeMaterialSeguimiento` | `MaterialSeguimiento.cls:38-64` | Lectura |
| `AnexosCerrados` | Comprueba si algún anexo está abierto | `MaterialSeguimiento.cls:71-114` | `fso`, `FicheroAbierto` |
| `Material` | Material padre vía `getMaterial` | `MaterialSeguimiento.cls:153-180` | Lectura |
| `EliminarAnexos` | Elimina los anexos si están cerrados | `MaterialSeguimiento.cls:182-225` | Ficheros |

**Datos.** `TbMaterialSeguimiento` se lee en `Constructor.getMaterialSeguimiento` (`Constructor.bas:1897-1936`); escritura fuera de la clase (`Form_FormEventoCopia.cls:423`).
**Consumidores.** `Form_FormMaterialSeguimiento.cls:58`, `Form_FormMaterialSeguimiento.cls:720`, `Form_FormAnexos.cls:537`, `Constructor.bas:2396` (`New MaterialSeguimiento`).
**Capacidad.** CAP-7.
**Posibles defectos (a verificar).**
- `fso` se declara sin `New` y no hay `Class_Initialize` que lo cree (`MaterialSeguimiento.cls:13`), pero se usa en `AnexosCerrados` (`MaterialSeguimiento.cls:95`): provocaría error 91 si hay anexos.
- `IDSeguimientoCalculado` es un campo público, no una propiedad calculada (`MaterialSeguimiento.cls:27`); su cálculo no está en la clase.
**Pendientes.** Cálculo de `IDSeguimiento`: sin referencia localizada en la clase.

## Parte
`classes/Parte.cls` · 645 líneas · 15 procedimientos.

**Propósito.** «Parte» de una factura y un centro (BUI): cabecera `TbPartesPpal` y detalle `TbPartesDetalle` con los eventos incluidos; genera el Word para firma y anexa el parte firmado (`Parte.cls:167-459`).

| Procedimiento | Qué hace | Cita | Datos / efectos |
|---|---|---|---|
| `Alta(p_ObjColEventos, p_IDFactura)` | Borra partes previos de la factura y centro, calcula fechas mín./máx., inserta cabecera y una fila de detalle por evento; si falla el detalle, borra la cabecera | `Parte.cls:290-418` | `DELETE` concatenados (`Parte.cls:319-326`), `AddNew` DAO (`Parte.cls:341-394`) |
| `Eliminar` | Borra cabecera, detalle, el fichero del parte y pone `TbEventos.IDParte` a Null | `Parte.cls:420-459` | SQL concatenado, `DeleteFile` |
| `AnexarParte(p_URLLocal)` | Copia el parte firmado al directorio de partes firmados con sufijo `_PA` y actualiza `NombreArchivoAdjunto` | `Parte.cls:167-220` | `CopyFile`, `UPDATE` (`Parte.cls:208-212`) |
| `GeneraWordParaFirma` | Obtiene eventos de factura y centro, abre plantilla Word y la rellena | `Parte.cls:222-288` | Word por automatización |
| `IDParteCalculado`/`getIDParte` | ID = prefijo literal + 2 primeros dígitos del año + mes + ordinal (máx. + 1) | `Parte.cls:461-514` | `SELECT Max(Right(...))` concatenado (`Parte.cls:492-495`) |
| `URLParte`, `colEventos`, `Factura`, `AnexoCerrado` | Navegación y estado del anexo | `Parte.cls:33-165` | Lecturas |

**Datos.** Escribe `TbPartesPpal`, `TbPartesDetalle`; actualiza `TbEventos`. DAO y SQL concatenado; sin transacciones: la compensación del alta es un `DELETE` manual en el manejador (`Parte.cls:408-415`).
**Consumidores.** `Factura.cls:4051-4057` (`New Parte` + `.Alta`), `Factura.cls:5142` (`Eliminar`), `Form_FormParteGestion.cls:59` (`AnexarParte`), `Form_FormParteGestion.cls:212` (`GeneraWordParaFirma`), `Constructor.bas:1796`.
**Capacidad.** CAP-8 Facturación, gastos y subcontratación.
**Posibles defectos (a verificar).**
- `getIDParte` usa `Left(Year(Date),2)`, que da los dos primeros dígitos del año y no los dos últimos; el ordinal es `Right(...,2)`, por lo que el prefijo no distingue años y solo caben 99 partes por mes (`Parte.cls:491-505`).
- `Alta` borra partes previos de la factura y centro antes de comprobar que el nuevo se pueda insertar (`Parte.cls:319-326`).
- El manejador de `GeneraWordParaFirma` evalúa `appWord.Documents(...)` aunque `appWord` pueda ser `Nothing` (`Parte.cls:280-286`).
- `Eliminar` y `AnexarParte` desactivan avisos con `DoCmd.SetWarnings` (`Parte.cls:210-212`, `433-449`).
**Pendientes.** `getFechaInicialYFinalDeColeEventos` y `RellenarInformeEnWord` son helpers externos (`Parte.cls:254`, `266`), no inventariados aquí.

## Subcontratacion
`classes/Subcontratacion.cls` · 374 líneas · 8 procedimientos.

**Propósito.** Entidad «subcontratación» (empresa, fechas, descripción, importe) con anexos y factura asociada, sin persistencia propia (`Subcontratacion.cls:34-258`).

| Procedimiento | Qué hace | Cita | Datos / efectos |
|---|---|---|---|
| `ColAnexos` | Anexos vía `getAnexosDeSubContratacion` | `Subcontratacion.cls:34-60` | Lectura |
| `AnexosCerrados` | `No` si algún fichero de anexo está abierto | `Subcontratacion.cls:67-111` | `FileSystemObject`, `FicheroAbierto` |
| `Factura`, `FechaFacturado` | Factura vía `getFacturaDeSubContratacion` y su fecha | `Subcontratacion.cls:150-214` | Lectura |
| `URLAnexos`, `EliminarAnexos` | Une URLs con `;`; borra anexos si están todos cerrados | `Subcontratacion.cls:112-149`, `215-258` | Ficheros |

**Datos.** `TbSubcontrataciones` se lee en `Constructor.getSubContratacion` (`Constructor.bas:3868`); la escritura no está en la clase (pendiente).
**Consumidores.** `Form_FormSubContratacionEdicion.cls:113`, `Form_FormSubContratacionGestion.cls:140`, `Constructor.bas:1107`, `Constructor.bas:3606`.
**Capacidad.** CAP-8.
**Posibles defectos (a verificar).** `URLAnexos` acumula en llamadas repetidas (`Subcontratacion.cls:112-149`); el mensaje de error de `URLAnexos` cita «subContratacion» (`Subcontratacion.cls:145`).
**Pendientes.** Alta/edición/borrado de `TbSubcontrataciones`: ejecutados desde formularios (`Form_FormSubContratacionAlta.cls`), no inventariados aquí.

## Tecnico
`classes/Tecnico.cls` · 703 líneas · 15 procedimientos.

**Propósito.** Entidad «técnico» (`TbTecnicos`) con validación, alta/edición/baja y alta de ausencias por rango; calcula día laborable (`Tecnico.cls:133-576`).

| Procedimiento | Qué hace | Cita | Datos / efectos |
|---|---|---|---|
| `MotivoNoOK` | Obligatorios alias, tipo, nombre y fecha de alta; alias único; baja ≥ alta; el tipo debe existir | `Tecnico.cls:133-216` | `ValorExistente` sobre `TbTecnicos` y `TbTipoTecnico` |
| `Registrar` | Alta (`AddNew`) o edición (`Edit`) por el alias inicial | `Tecnico.cls:218-328` | DAO; SQL en `Tecnico.cls:256` |
| `ValidacionEliminar` / `Borrar` | Impide borrar si hay eventos, actividades o gastos del técnico | `Tecnico.cls:330-413` | `SELECT Count` en `TbEventos`, `TbActividades`, `TbGastos` (`Tecnico.cls:344`, `358`, `371`); `DELETE` (`Tecnico.cls:403`) |
| `AltaAusencias` | Recorre el rango de fechas, omite fines de semana, fiestas y libranzas existentes y registra una `LIbranza` por día | `Tecnico.cls:415-507` | Escribe `TbTecnicosAusencias` vía `LIbranza.Registrar` |
| `EsDiaLaborable` | No laborable sábado/domingo, libranza del técnico o fiesta | `Tecnico.cls:509-576` | `TbTecnicosAusencias` (`Tecnico.cls:534-537`), `TbTecnicosFiestas` (`Tecnico.cls:552-554`) |
| `TipoTecnicoPrecio`, `ConPrecioActivo`, `ColLibranzas` | Precio vigente y ausencias del técnico | `Tecnico.cls:31-131` | Lecturas vía `Constructor` |

**Datos.** Lee/escribe `TbTecnicos`; lee `TbTipoTecnico`, `TbEventos`, `TbActividades`, `TbGastos`, `TbTecnicosAusencias`, `TbTecnicosFiestas`; DAO y SQL concatenado; sin transacciones (un fallo a mitad de `AltaAusencias` deja días ya registrados, `Tecnico.cls:476-492`).
**Consumidores.** `LIbranza.cls:47-79` (`Tecnico`), `Constructor.bas:1980-1985` (`getTipoTecnicoPrecio` por alias, usado en `Tecnico.cls:46`); formularios de técnicos no trazados uno a uno (pendiente).
**Capacidad.** CAP-9.
**Posibles defectos (a verificar).**
- `ColLibranzas` llama a `getLibranzas(Me.Alias, m_Error)`, pero la firma es `(p_Año, p_Alias, p_Error)` (`Tecnico.cls:112`, `Constructor.bas:4118-4122`): el alias iría como año.
- `AltaAusencias` decide alta/edición por `p_ObjAusenciaAlInicio` para todos los días del rango (`Tecnico.cls:458-466`).
**Pendientes.** Ninguna referencia dinámica sin resolver.

## TecnicoHoras
`classes/TecnicoHoras.cls` · 528 líneas · 14 procedimientos.

**Propósito.** Resumen de horas de un técnico en un día: tipo de día, libranza, festivo, actividades y horas laborables/extras (`TecnicoHoras.cls:36-382`).

| Procedimiento | Qué hace | Cita | Datos / efectos |
|---|---|---|---|
| `DiaDeLaSemana` | Nombre del día con `WeekdayName(Weekday(fecha - 1))` | `TecnicoHoras.cls:36-56` | — |
| `TipoDia` | «No Laborable» si sábado/domingo o festivo; si no «Laborable» | `TecnicoHoras.cls:58-114` | — |
| `EsLibranza` / `EsFestivo` | Existencia en `TbTecnicosAusencias` / `TbTecnicosFiestas` por fecha | `TecnicoHoras.cls:384-417`, `419-450` | `SELECT` concatenado |
| `ColActividades` | Actividades del técnico en la fecha | `TecnicoHoras.cls:176-198` | `Constructor.getActividadesTecnicoDia` (`Constructor.bas:4396`) |
| `Hlab`, `HExt` | Suma de `HorasLaborables`/`HorasExtras` | `TecnicoHoras.cls:206-288` | — |
| `ActividadesTextoLab/Ext` | IDs de actividad con horas > 0 unidos con `;` | `TecnicoHoras.cls:290-382` | — |

**Datos.** Lee `TbTecnicosAusencias`, `TbTecnicosFiestas`; sin escrituras.
**Consumidores.** `Form_FormTecnicoConsultaHoras.cls:149` (`New TecnicoHoras`) y `Form_FormTecnicoConsultaHoras.cls:245` (`.TipoDia`).
**Capacidad.** CAP-9.
**Posibles defectos (a verificar).**
- `Weekday(CDate(Fecha) - 1)` desplaza un día el nombre salvo que la semana empiece en lunes por configuración (`TecnicoHoras.cls:45`).
- `ObservacionesFestivo = Festivo` asigna un `EnumSino`, no un texto (`TecnicoHoras.cls:90`).
- Con libranza, `TipoDia` sigue siendo «Laborable» y solo se anota «Libranza» en observaciones (`TecnicoHoras.cls:95-106`).
- `TipoDia`, `Libranza` y `Festivo` cachean su resultado y no se recalculan si cambian `Fecha` o `Alias` (`TecnicoHoras.cls:58-60`, `122-125`, `152-155`).
**Pendientes.** Ninguna.

## TipoDia
`classes/TipoDia.cls` · 186 líneas · 5 procedimientos.

**Propósito.** Variante de `TecnicoHoras.TipoDia` que devuelve el motivo de festivo o de libranza como texto (`TipoDia.cls:80-134`, `138-186`).

| Procedimiento | Qué hace | Cita | Datos / efectos |
|---|---|---|---|
| `TipoDia` | Fin de semana o festivo ⇒ «No Laborable»; si no «Laborable» y anota la libranza | `TipoDia.cls:80-134` | — |
| `EsDiaFestivo` | Devuelve `Observaciones` de `TbTecnicosFiestas` | `TipoDia.cls:138-161` | `SELECT *` concatenado |
| `EsDiaLibranza` | Devuelve `Motivo` de `TbTecnicosAusencias` | `TipoDia.cls:163-186` | `SELECT *` concatenado |

**Datos.** Lee `TbTecnicosFiestas` y `TbTecnicosAusencias`.
**Consumidores.** Ninguno localizado (sin `New TipoDia` ni `As TipoDia` en `src`).
**Capacidad.** CAP-9.
**Posibles defectos (a verificar).** Duplica la lógica de `TecnicoHoras` y parece código sin uso; mismo desplazamiento de día (`TipoDia.cls:68`).
**Pendientes.** Confirmar con el índice si hay uso dinámico.

## TipoTecnico
`classes/TipoTecnico.cls` · 526 líneas · 13 procedimientos.

**Propósito.** Catálogo de tipos de técnico (`TbTipoTecnico`, un único campo `TIPO`) y precio vigente por tipo (`TipoTecnico.cls:117-462`).

| Procedimiento | Qué hace | Cita | Datos / efectos |
|---|---|---|---|
| `PrecioHoraLabActual`/`PrecioHoraExtraActual` | Precio más reciente por `IDTipoTecnicoFecha DESC` | `TipoTecnico.cls:117-212` | Lee `TbTipoTecnicoPrecios` |
| `MotivoNoOK` | `TIPO` obligatorio y único; renombrar prohibido si está en facturas | `TipoTecnico.cls:214-274` | `ValorExistente` |
| `TipoEnAlgunaFactura` | Cuenta en `TbFacturaEventosInvolucrados` y `TbFacturaActividadesInvolucradas` | `TipoTecnico.cls:276-324` | `SELECT Count` |
| `Registrar` | Alta o edición; al renombrar actualiza `TbTecnicos.TIPO` en cascada | `TipoTecnico.cls:326-399`, `385-387` | `AddNew`/`Edit` + `UPDATE` |
| `ValidacionEliminar`/`Borrar` | No borra si está en facturas o lo usa algún técnico | `TipoTecnico.cls:402-462` | `DELETE` concatenado |
| `ColTecnicos*` | Técnicos del tipo (todos/activos/no activos) | `TipoTecnico.cls:27-115` | Lectura |

**Datos.** Escribe `TbTipoTecnico` y `TbTecnicos` (cascada); lee `TbTipoTecnicoPrecios` y las tablas de facturas; sin transacciones: tras el `Edit` hay un `UPDATE` aparte (`TipoTecnico.cls:382-387`).
**Consumidores.** `TipoTecnicoPrecio.cls:68` (`TIPOTECNICOObj`), `TipoTecnicoPrecio.cls:774` (`Borrar`).
**Capacidad.** CAP-9.
**Posibles defectos (a verificar).**
- El renombrado no actualiza `TbTipoTecnicoPrecios.TIPOTECNICO` (`TipoTecnico.cls:382-387`).
- `ValidacionEliminar` no consulta `TbTipoTecnicoPrecios` (`TipoTecnico.cls:402-440`).
- Los mensajes de error citan «Tecnico.MotivoNoOK»/«Tecnico.getPrecioHora» (`TipoTecnico.cls:210`, `271`).
**Pendientes.** Formularios consumidores directos: no trazados.

## TipoTecnicoPrecio
`classes/TipoTecnicoPrecio.cls` · 892 líneas · 25 procedimientos.

**Propósito.** Tarifa histórica por tipo de técnico (`TbTipoTecnicoPrecios`): cada fila tiene fechas, precio laborable y precio extra; solo una abierta por tipo (`TipoTecnicoPrecio.cls:287-785`).

| Procedimiento | Qué hace | Cita | Datos / efectos |
|---|---|---|---|
| `MotivoNoOK` | Tipo y precios numéricos obligatorios; sin fecha final en alta; no edita un tipo cerrado; no permite cerrar si hay elementos por facturar | `TipoTecnicoPrecio.cls:287-347` | Lecturas |
| `TipoEnAlgunTecnicoPorFacturar` | Cuenta actividades y eventos de técnicos del tipo sin facturar | `TipoTecnicoPrecio.cls:349-400` | `SELECT Count` |
| `Registrar` | Cierra el precio activo con fecha de hoy, inserta/edita la fila; si falla, lo reabre | `TipoTecnicoPrecio.cls:403-519`, `436-439`, `511-514` | `UPDATE` + DAO |
| `Cerrar` / `Abrir` | Pone o quita `FechaFinal` | `TipoTecnicoPrecio.cls:521-574` | `UPDATE` |
| `ValidacionEliminar` / `Borrar` | Solo la fila más reciente, sin técnicos con ese perfil; si era la única borra el tipo | `TipoTecnicoPrecio.cls:575-785` | `DELETE`; `TIPOTECNICOObj.Borrar` (`TipoTecnicoPrecio.cls:774`) |
| `TipoActivo`, `Anterior`, `Siguiente`, `Ultimo`, `EsUltima`, `ColTecnicos` | Navegación por la historia | `TipoTecnicoPrecio.cls:34-285` | Lecturas |

**Datos.** Lee y escribe `TbTipoTecnicoPrecios`; lee `TbTecnicos`, `TbActividades`, `TbEventos` y tablas de facturación; sin transacciones: la compensación es un `UPDATE` en el manejador (`TipoTecnicoPrecio.cls:510-515`).
**Consumidores.** `Tecnico.cls:46` (`getTipoTecnicoPrecio`), `TipoTecnicoPrecio.cls:261-286` (`ColTecnicos` vía `getTecnicosPorTipo`). `TipoTecnico.cls:173-212` lee la misma tabla con SQL propio, sin usar la clase.
**Capacidad.** CAP-9.
**Posibles defectos (a verificar).**
- `ValidacionEliminar` pasa `Me.TipoActivo` (un `EnumSino`) donde las funciones esperan el tipo (`TipoTecnicoPrecio.cls:620`, `630`).
- `Cerrar` compara `IDTipoTecnicoFecha` entre comillas y el resto de sentencias sin ellas (`TipoTecnicoPrecio.cls:540` frente a `566`).
- El mensaje de error de `Borrar` cita «Tecnico.Borrar» (`TipoTecnicoPrecio.cls:783`).
**Pendientes.** `TIPOTECNICOObj` se declara `Database.TIPOTECNICO` (`TipoTecnicoPrecio.cls:68`): calificación por nombre de proyecto; resolución a `TipoTecnico` sin confirmar.

## Usuario
`classes/Usuario.cls` · 412 líneas · 6 procedimientos.

**Propósito.** Entidad «usuario de la lanzadera» (`TbUsuariosAplicaciones`) con sus permisos por aplicación; sin persistencia propia (`Usuario.cls:50-118`).

| Procedimiento | Qué hace | Cita | Datos / efectos |
|---|---|---|---|
| `EsAdministradorCalculado` | «Sí»/«No» del campo `EsAdministrador` a `EnumSino` | `Usuario.cls:50-68` | — |
| `ColAplicacionesPermisos` | Permisos del usuario por correo | `Usuario.cls:70-90` | `Constructor.getAplicacionesPermisos` (`Constructor.bas:148-195`, lectura en `Constructor.bas:165`) |
| `Permisos` | Permisos de la aplicación activa (`IDAplicacion` global, `VG.bas:239`) | `Usuario.cls:92-118` | — |
| `ColCampos`, `getPropiedad`, `SetPropiedad` | Reflexión de 26 campos públicos (`Usuario.cls:13-38`), incluidos campos de contraseña y caducidad | `Usuario.cls:120-412` | — |

**Datos.** Solo lectura de la base de la lanzadera (`getdbLanzadera`, `VG.bas:520`); `getUsuario` consulta `TbUsuariosAplicaciones` (`Constructor.bas:71`). La clase mantiene en memoria campos de contraseña y marcas de caducidad; no se transcribe ningún valor.
**Consumidores.** `VG.bas:335-345` (comprobación de acceso), `Form_Form0BDOpciones.cls:62`, `GestorInforme.cls:76-129`.
**Capacidad.** CAP-1 Arranque, identidad y permisos.
**Posibles defectos (a verificar).** `Permisos` depende de la variable global `IDAplicacion`; si está vacía, devuelve `Nothing` (`Usuario.cls:92-118`); el arranque lo trata como «sin permiso» (`VG.bas:335-342`). Los campos de contraseña son cadenas públicas en la clase (`Usuario.cls:13-38`).
**Pendientes.** Escritura de usuarios: sin código en esta clase.

## UsuarioAplicacionPermisos
`classes/UsuarioAplicacionPermisos.cls` · 402 líneas · 20 procedimientos.

**Propósito.** Permisos de un usuario sobre una aplicación: siete indicadores (administrador, calidad, economía, secretaría, técnico, sin acceso, avisos de calidad) en forma de texto «Sí»/«No» y su versión `EnumSino` (`UsuarioAplicacionPermisos.cls:68-250`).

| Procedimiento | Qué hace | Cita | Datos / efectos |
|---|---|---|---|
| `Es…Calculado` Get/Let (7 pares) | Convierte «Sí»/«No» a `EnumSino`; el `Let` fija el valor en memoria | `UsuarioAplicacionPermisos.cls:68-250` | — |
| `Usuario` | Referencia al usuario propietario | `UsuarioAplicacionPermisos.cls:39-66` | — |
| `ColCampos`, `getPropiedad`, `SetPropiedad` | Reflexión de 9 campos | `UsuarioAplicacionPermisos.cls:253-402` | — |

**Datos.** Se carga desde `TbUsuariosAplicacionesPermisos` (`Constructor.bas:165`), solo lectura.
**Consumidores.** `VG.bas:343-344` (`EsUsuarioAdministradorCalculado`/`EsUsuarioTecnicoCalculado` deciden el acceso), `Constructor.bas:186`.
**Capacidad.** CAP-1.
**Posibles defectos (a verificar).** El `Let` de cada indicador no actualiza el texto de origen, por lo que `getPropiedad` y la versión calculada pueden divergir (`UsuarioAplicacionPermisos.cls:88-90`).
**Pendientes.** Uso de los indicadores calidad, economía, secretaría y avisos en los formularios: no trazado.
