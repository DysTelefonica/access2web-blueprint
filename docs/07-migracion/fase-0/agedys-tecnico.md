# Fase 0: AGEDYS Técnico (`GestionTecnicaProyectos`)

Épica `docs/07-migracion/epica-coexistencia-postgresql.md` (fase 0). Ficha de inventario de la mitad técnica de AGEDYS; la mitad del gestor económico va en su propia ficha, y las dos acaba siendo una sola aplicación web, APGEDYS, con vista de técnico o de gestor económico según el usuario que entre. No hay issue propia: la revisión se hace con el catálogo conjunto de accionables, como las fichas de HPS, Condor, Expedientes, Gestión de Riesgos, No Conformidades y Registros.

Base de análisis: rama `main`, SHA `350896d`, repositorio `DysTelefonica/AGEDYS`, copia local `C:\00repos\codigo\00_AGEDYS`.

Fuentes: ERD del front y del backend con `scripts/dysflow-erd/extract-erd.ps1` (solo lectura, sobre copia temporal), volcado del SQL de las consultas guardadas y de los contenedores de Access con DAO en solo lectura sobre copia, trazado del código con CodeGraph (`codegraph query`, índice local de 150 ficheros y 10 249 nodos) y contraste con `grep` sobre `src/`. CodeGraph estuvo operativo: no hubo fallback. Los vínculos del front apuntan a `\\datoste`, inalcanzable desde este entorno, de modo que el esquema de las 127 tablas vinculadas del front no se pudo leer; el de sus destinos sí, a través de los vínculos resueltos del backend.

## Base de análisis

| Hecho | Evidencia |
|---|---|
| La rama base es `main` | `git branch -avv` en `C:\00repos\codigo\00_AGEDYS`: `main` en `350896d`, sincronizado con `origin/main`. No existe ninguna rama RC: las otras dos son `docs/codebase-guide-alan` y `fix/agedys-lblROPVC-contador`, esta última ya fusionada (`abc6f91`) |
| No hay evidencia de UAT ni de ciclo de release en el repositorio | `main` no tiene carpeta `docs/`. `README.md:1` es solo `# AGEDYS`. La única mención de UAT es `AGENTS.md:27` («átomos TDD espejo de escenarios UAT»), que describe una práctica de trabajo, no un estado de un ciclo. No hay roadmap, ni registro de release, ni documento de firma |
| La versión que declara el repositorio no es la que declara el servidor | `GestionTecnicaProyectos.ini:2` (`VersionAplicacion=2026-001`). La lanzadera lee esa clave (`00_LANZADERA/src/classes/Aplicacion.cls:1133`) y la muestra como «Ver. Servidor» (`00_LANZADERA/src/modules/Funciones Genreales.bas:2326`). El propietario declara «AGEDYS TÉCNICO, Ver. Servidor 2026-002»: el repositorio va una versión por detrás del artefacto desplegado |
| La rama `docs/codebase-guide-alan` no sirve como base | Está dos commits por delante de su remoto (`chore: wip snapshot before removing local worktree`, 2026-10-25) y su `src/` difiere de `main` en 90 ficheros (`+1399/-87`). El conjunto de tablas que nombra el código es idéntico en las dos ramas, comprobado con `git grep` sobre ambos árboles |
| El binario del front tiene un formulario que el export no contiene | El contenedor `Forms` del `.accdb` tiene 45 formularios y `src/forms/` tiene 44 export. El sobrante es `Copia de FormGestionProyectos`, un duplicado de `FormGestionProyectos` que no está versionado |

Consecuencia: el inventario de abajo describe `main` en `350896d`, y no el binario desplegado. Las diferencias con la versión 2026-002 no son un error del inventario: no son medibles desde aquí.

## Estructura del front

| Hecho | Valor |
|---|---|
| Tablas del front | 137: 10 propias y 127 vinculadas |
| Tablas propias del front | `Copia de TbRequisitosSTG`, `Tabla1`, `TbAuxExp`, `TbAuxExpPrado`, `TbAuxNPedido`, `TbAuxTomas`, `TbHerramientaDocAyuda`, `TbProductosObsoletos`, `TbSTGParametrosParaFormulario`, `TbSuministradoresListaFavoritos` |
| Destino de los 127 vínculos | `\\datoste\aplicaciones_dys\Aplicaciones PpD\`: 101 a `Proyectos\AGEDYS_DATOS.accdb`, 11 a `EXPEDIENTES\Expedientes_datos.accdb`, 4 a `AGEDO\AGEDO20_Datos.accdb`, 4 a `0Lanzadera\Lanzadera_Datos.accdb`, 4 a `REGISTROS\Registro_Ent_Salida_Datos.accdb`, 1 a `REGISTROS\Registro_Datos.accdb`, 1 a `00Recursos\Correos_datos.accdb`, 1 a `GESTION RIESGOS\Gestion_Riesgos_Datos.accdb` |
| Vínculos con nombre distinto del origen | 7: seis con sufijo numérico (`TbAuxExp1`, `TbBloqueos1`, `TbHerramientaDocAyuda1`, `TbProductosObsoletos1`, `TbSTGParametrosParaFormulario1`, `TbSuministradoresListaFavoritos1`), de los que cinco homóniman tablas locales del front, y `TbProyectosGestionRiesgos`, que es `Gestion_Riesgos_Datos.TbProyectos` renombrada |
| Formularios con `RecordSource` | 0 |
| Controles con `ControlSource` | 0 |
| Controles con `RowSourceType` | 14: 2 de tipo `Table/Query` y 12 de tipo `Value List` |
| Asignaciones de `RowSource` en tiempo de ejecución | 43, ninguna con SQL literal: 39 listas de valores y 4 que toman una variable |
| Consultas guardadas del front | 13 (`Consulta1` a `Consulta12` y `DpsEnELLimbo`); ocho con literales de depuración incrustados: `Consulta1` y `Consulta4` (`'a'`), `Consulta3` (`'24dpd-032'`), `Consulta5` (`'jvm'`), `Consulta6` (`'24DPD-273'`), `Consulta9` (`421`), `Consulta11` (`SELECT;`) y `Consulta12` (`Year(FREGISTRO) = 2024`) |
| Módulos | 62, los 62 del `src/` (16 en `src/modules/` y 46 en `src/classes/`) |
| Macros y `AutoExec` | Ninguno. El inventario de contenedores del `.accdb` no incluye un contenedor de macros y en todo `src/` no aparece `AutoExec` ni `AutoOpen` |

La regla «salvo Brass, los formularios son desconectados» **se confirma por medición en AGEDYS Técnico**: los 44 formularios exportados no tienen `RecordSource` de formulario ni ningún `ControlSource`, y lo único que tiene origen de datos son 14 listas, de las que solo 2 llevan SQL en diseño. El resto de las listas se rellena en tiempo de ejecución con SQL ejecutado sobre `CurrentDb()` o `getdb()` y volcado con `AddItem`.

## Aperturas directas

Ocho funciones de apertura por código, todas en `src/modules/Variables Globales.bas`, con la credencial literal en las ocho:

| Apertura | Código | Ruta | Observación |
|---|---|---|---|
| Backend propio | `src/modules/Variables Globales.bas:467` (`getdb`), apertura en `:485` | `proyectos\AGEDYS_DATOS.accdb` (`:475`, `:477`) | Es la que usa la mayor parte del código: 150 apariciones de `getdb` en `src/`. De las 87 llamadas a un abridor que hay en `src/modules/constructor.bas`, 79 son a `getdb()` y 8 a `getdbLanzadera()` |
| Lanzadera | `src/modules/Variables Globales.bas:439` (`getdbLanzadera`), apertura en `:458` | `0Lanzadera\Lanzadera_Datos.accdb` (`:449`) | La usan `src/classes/UsuarioAplicacionPermisos.cls:263` (directamente sobre `TableDefs`), `src/modules/constructor.bas:1252, 1294, 1329, 1384, 1436, 1488, 3648, 3723` y `src/modules/Funciones Generales.bas:8420, 8425, 8469` |
| Correos | `src/modules/Variables Globales.bas:519` (`getdbCorreo`), apertura en `:543` | `00Recursos\Correos_datos.accdb` (`:530`) | `src/classes/Correo.cls:47`, `src/classes/CorreoOperaciones.cls:499`, `src/modules/Funciones Generales.bas:1366, 1369, 2552`, `src/modules/MCorreo.bas:81, 87` |
| Registro de entradas y salidas | `src/modules/Variables Globales.bas:552` (`getdbSalida`), apertura en `:577` | `REGISTROS\Registro_Ent_Salida_Datos.accdb` (`:561`) | `src/classes/Salida.cls:88, 190` y `src/modules/Funciones Generales.bas:2414` |
| AGEDO | `src/modules/Variables Globales.bas:620` (`getdbAgedo`), apertura en `:643` | `AGEDO\AGEDO20_Datos.accdb` (`:629`) | `src/modules/constructor.bas:106`, que lee `TbDocumentos` de AGEDO |
| Apertura de correo | `src/modules/Variables Globales.bas:495` (`getdbCorreoApertura`), apertura en `:510` | `00Recursos\Correo_Apertura.accdb` (`:505`) | **Sin llamadores**: solo su definición y su manejador de error |
| Registro antiguo | `src/modules/Variables Globales.bas:587` (`getdbSalidaAntigua`), apertura en `:611` | `REGISTROS\Registro_Datos.accdb` (`:597`) | **Sin llamadores** |
| Expedientes | `src/modules/Variables Globales.bas:891` (`getdbExpediente`), apertura en `:911` | `EXPEDIENTES\Expedientes_datos.accdb` (`:902`) | **Sin llamadores**: el front llega a Expedientes por vínculo, no por código |

Tres de las ocho aperturas no tienen ningún llamador. Las otras cinco sí, y todas resuelven contra el share de producción: `Application.TempVars("DatosEnLocal")` está fijado a `"No"` en `src/modules/Variables Globales.bas:672` (la línea que lo fijaba a `"Sí"` está comentada justo encima), de modo que el desarrollo apunta siempre a `\\datoste`.

Dos de los cinco abridores vivos tienen la condición local/remoto invertida respecto de los otros tres: `getdbSalida` (`:558-563`) y `getdbAgedo` (`:626-631`) eligen la ruta **local** cuando `DatosEnLocal = "No"` y la remota cuando vale `"Sí"`, al contrario que `getdb`, `getdbLanzadera` y `getdbCorreo`, que eligen la remota en ese mismo caso. El tercero con la inversión, `getdbSalidaAntigua` (`:593-599`), es uno de los que no tienen llamadores.

## Tablas de otros backends que consume el front

| Backend | Tablas | Cómo llega el front |
|---|---|---|
| `AGEDYS_DATOS` | 101 tablas vinculadas | Vínculo del front y `getdb()` por código. Es su backend |
| `Expedientes_datos` | `TbExpedientes`, `TbExpedientesAnualidades`, `TbExpedientesResponsables`, `TbExpedientesJuridicas`, `TbExpedientesPECAL`, `TbExpedientesRACS`, `TbEstados`, `TbGradosClasificacion`, `TbJuridicas`, `TbPECAL`, `TbRACS` | Vínculo del front, además del que tiene el propio backend |
| `Lanzadera_Datos` | `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` | Vínculo del front, más `getdbLanzadera()` por código. `TbAplicacionesAperturas` solo por código (`src/modules/Funciones Generales.bas:8420, 8424, 8465`) |
| `Correos_datos` | `TbCorreosEnviados` | Vínculo del front y `getdbCorreo()` |
| `AGEDO20_Datos` | `TbDocumentos`, `TbDocumentosID`, `TbDocExpProceso`, `TbAgedoCorreosEnviados` | Vínculo del front y `getdbAgedo()` |
| `Registro_Ent_Salida_Datos` | `TbEntradas_RES`, `TbSalidas_RES`, `TbSalidas_Indexacion`, `TbSalidas_RES_Indexacion` | Vínculo del front y `getdbSalida()` |
| `Registro_Datos` | `Salidas` | Vínculo del front |
| `Gestion_Riesgos_Datos` | `TbProyectosGestionRiesgos` (es `TbProyectos`) | Vínculo del front, renombrado. Solo lo usa la consulta guardada `Consulta8` |

El front no consulta solo su backend: la regla R-1 se refuta aquí igual que en las demás aplicaciones del ecosistema.

## Backend `AGEDYS_DATOS`

El fichero vive en el repositorio del front (`AGEDYS_DATOS.accdb`, 77 MB, 2026-09-30) y también en `C:\00repos\datos\AGEDYS_DATOS.accdb` (83 MB, 2026-10-23). **No son la misma copia**: tienen los mismos 102 nombres de tabla y los mismos 35 vínculos, pero 16 tablas tienen recuento distinto. Los recuentos de este inventario son los de la copia del repositorio.

| Hecho | Valor |
|---|---|
| Tablas | 102 propias y 35 vinculadas |
| Destino de los 35 vínculos | `Expedientes_datos` (23), `Lanzadera_Datos` (4), `AGEDO20_Datos` (3), `Registro_Ent_Salida_Datos` (2), `Registro_Datos` (1), `Correos_datos` (1) y **una autovinculación a `AGEDYS_DATOS.accdb`** (`TbSolicitudesOfertasPrevias1` → `TbSolicitudesOfertasPrevias`) |
| Relaciones | 4, de las que dos son del catálogo de navegación de Access (`MSysNavPane*`). Quedan **2 relaciones reales** entre tablas de negocio, para 102 tablas |
| Consultas guardadas | 5 (`Consulta1` a `Consulta5`), tres de ellas de tipo tabla (`dbQueryTable`) y dos con SQL de depuración |
| Tablas sin clave primaria | 11: `TbAuxExp`, `TbAuxPrado`, `TbAvalesDatosGenerales`, `TbAvalesDatosGeneralesNuevos`, `TbDpDCartaSolicitudOfertaDatosFijos`, `TbDpDCartaSolicitudOfertaPlantillasSegunReqBasicos`, `TbFacturasPrincipal`, `TbLog`, `TbProductos`, `TbProductosObsoletos`, `TbProductos_antes` |
| Campos multivalor, de datos adjuntos u objeto OLE | **Ninguno**. No hay ningún campo `IsComplex` ni de tipo 11 en las 102 tablas propias |
| Generador de identificadores | `DameID1` (`src/modules/Funciones Generales.bas:1261`) calcula `Max(campo) + 1` sobre la tabla que se le pase, con `getdb()` por defecto (`:1279-1280`). Es el generador que usan el front y el resto del ecosistema |

El backend no es solo de esta aplicación: la épica lo declara backend de AGEDYS, AGEDYS ECO y AGEDYS TEC (`docs/07-migracion/epica-coexistencia-postgresql.md:67`), y el front de AGEDYS ECO abre el mismo fichero por código (`00_AGEDYS_ECO/src/modules/Variables Globales.bas:475, 477`).

## Veredicto de propiedad

### Se quedan en el backend porque el VBA del front las necesita

40 tablas propias del backend con SQL en el código del front: `TbAdjudicacionesEnvioAlSuministrador`, `TbAdjudicacionesOfertasCartas`, `TbBloqueos`, `TbCestasAnexosOrdinaria`, `TbConsultasDetalle`, `TbDPDCartaSolicitudOfertaDatos`, `TbDPDComiteOfertas`, `TbDPDOfertaSuministrador`, `TbDpDAsuntoDefensa`, `TbDpDCartaSolicitudOfertaDatosFijos`, `TbDpDCartaSolicitudOfertaPlantillasSegunReqBasicos`, `TbDpDDocAnexada`, `TbDpDInformeCondicionamiento`, `TbDpDPrepCartaOfAcepDetalle`, `TbFacturasCliente`, `TbFacturasDetalle`, `TbFacturasPrincipal`, `TbGTV`, `TbGTVContratos`, `TbGTVContratosExpedientes`, `TbGTV_Temp`, `TbLog`, `TbNPedido`, `TbOrdinariasTipoAnexos`, `TbProductos`, `TbProyectos`, `TbRequisitosSTG`, `TbSolicitudesEnvioAlSuministrador`, `TbSolicitudesOfertasPrevias`, `TbSuministradoresNotasReferencia`, `TbSuministradoresPuntuacionConDetalle`, `TbSuministradoresSAP`, `TbTemp_PedidosEstado`, `TbTipoDocumento`, `TbTmp_DPDs`, `TbTmp_Facturas`, `TbVisadoFacturas_Nueva` y `TbVisadosGenerales`.

Tres más las nombra el front sin escribir SQL contra ellas, pero se leen por el vínculo del front: `TbConsultasPrincipal` (el `RowSource` de diseño de `ListaCategoria`), `TbHerramientaDocAyuda` y `TbSTGParametrosParaFormulario`. **Las dos últimas las lee la copia local del front, no la del backend**: `Dame` abre con `CurrentDb()` (`src/modules/Funciones Generales.bas:1474-1476`) y `getColParametrosSTG` también (`src/modules/STGmdl.bas:1103-1104`).

### Tablas mal ubicadas: el front técnico no las usa y las usa otro backend

| Tabla | Filas | La usa | Por qué está mal ubicada |
|---|---|---|---|
| `TbExpedientesAnualidades` | 223 | El backend de Expedientes, como `TbExpedientesAnualidades1` | El front técnico la nombra, pero la lee por su propio vínculo a `Expedientes_datos`, no desde `AGEDYS_DATOS`. La copia del backend no la lee nadie del núcleo |
| `TbResponsablesExpedientes` | 996 | Los backends de Expedientes y de Gestión de Riesgos | Cero referencias en el `src/` del front, verificado con CodeGraph. Es la tabla que las dos fichas de esos backends proponen aislar (`expedientes.md` E5, `gestion-riesgos.md` R4) |
| `TbComunicados` | 0 | El backend de Expedientes | Cero referencias en el `src/` del front |

### Tablas que no usa ninguno de los dos fronts

52 tablas propias del backend sin ninguna referencia en el `src/` de los dos fronts de AGEDYS, verificado con CodeGraph y con búsqueda exacta sobre `src/`:

- **Copia y históricos**: `TbExpedientes_1` (463 filas, 50 campos: copia de un `TbExpedientes` de esquema anterior), `TbNPedidoAntes` (1499), `TbNPedidosHistorica` (1), `TbProyectosHistoricos` (2047), `TbProductos_antes` (424), `TbVisadoFacturas` (1307, anterior a `TbVisadoFacturas_Nueva`).
- **Auxiliares y de trabajo**: `TbAuxExp` (40), `TbAuxExpedientesUsuario` (0), `TbAuxGTV` (40), `TbAuxIDFactura` (4923), `TbAuxPilar` (56), `TbAuxPilar1` (18), `TbAuxPrado` (468).
- **Avales**: `TbAvalesAnexo`, `TbAvalesCancelacion`, `TbAvalesDatosGenerales`, `TbAvalesDatosGeneralesNuevos`, `TbAvalesFacturasClienteVinculo`, `TbAvalesInicio`, `TbAvalesInicioNuevo`, `TbAvalesObtencionAvalYDeposito`, `TbAvalesSolicitudSIADE`.
- **Económicas y de expediente**: `TbDatosEconomicosExpedientes`, `TbExpedientesAutorizaciones`, `TbExpedientesDocRequeridos`, `TbExpedientesProrrogas`, `TbFacturaInformeMatFalsicado`, `TbFacturasDetalleSinPedido`, `TbTiposImpositivos`, `TbTiposPedido`.
- **Configuración y catálogos**: `TbAmbitoConsultas`, `TbCSS`, `TbCestasTiposDocumentos`, `TbComunicadosUsuario`, `TbConsultasOpciones`, `TbCriteriosSuministrador`, `TbDPDPlantillasParaCartaSolicitudDeOFerta`, `TbDireccionesSuministradores`, `TbNuevosProductos`, `TbProductosObsoletos`, `TbProductosParaBusqueda`, `TbProyectosDocumentosParaPedidoOrdinario`, `TbSuministradoresDireccion`, `TbSuministradoresListaFavoritos`, `TbSuministradoresPuntuacion`, `TbTareasTecnicasExplicaciones`, `TbTempDPDHTML`, `TbUsuariosEnvioTareas`, `TbCorreoDiferidoNotas`.
- **Conexiones**: `TbConexiones` (24), `TbConexionesAGEDYSECO` (9), `TbConexionesAGEDYSTEC` (15).
- **Otros**: `TbDpDPrepCartaOfAcepPrincipal` (860).

Dos advertencias sobre esta lista. La primera: `TbFacturasDetalleNoDYS` (513 filas) **sí** se usa, pero no desde el VBA: es el destino del SQL guardado de la consulta 21 de `TbConsultasDetalle`. Un rastreo solo de código la habría dado por huérfana. La segunda: ninguna de estas 52 se puede dar por retirada sin el grafo global de vínculos de la tarea T6 y el contraste con la ficha de AGEDYS ECO.

### Tablas del front

| Tabla | Filas | Uso |
|---|---|---|
| `Tabla1` | 22 | Tabla de trabajo: `src/modules/Módulo1.bas:48-50` |
| `TbAuxExpPrado` | 0 | Tabla de trabajo: se llena con un `INSERT INTO … SELECT` sobre `TbProyectos`, `TbNPedido`, `TbFacturasDetalle`, `TbGTV` y `TbExpedientes` (`src/modules/Funciones Generales.bas:16-31`, `src/forms/Form_FormConsultas.cls:426-469`) |
| `TbAuxTomas` | 62 | Tabla de trabajo: `src/modules/Módulo1.bas:424, 593-595` |
| `TbAuxNPedido` | 35 | **Sin ninguna referencia** en `src/` |
| `Copia de TbRequisitosSTG` | 139 | Copia con los mismos 73 campos que `TbRequisitosSTG`, que sí se usa. **Sin ninguna referencia** |
| `TbAuxExp` | 0 | Homónima de la del backend (40 filas, con un campo más: `FECHAADJUDICACION`). Sin referencias |
| `TbHerramientaDocAyuda` | 33 | Homónima de la del backend (10 filas). La lee el front desde la copia local |
| `TbSTGParametrosParaFormulario` | 15 | Homónima de la del backend (15 filas). La lee el front desde la copia local |
| `TbProductosObsoletos` | 25 | Homónima de la del backend (25 filas). Sin referencias |
| `TbSuministradoresListaFavoritos` | 0 | Homónima de la del backend (0 filas). Sin referencias |

Access renombró el vínculo cuando la tabla local ya ocupaba el nombre. Cinco de las seis colisiones se ven todavía en el front; `TbBloqueos1` ya no tiene tabla local que lo justifique.

### Referencias a tablas o campos que no existen

| Referencia | Dónde | Qué pasa |
|---|---|---|
| `TbGTVContratos1` | `src/classes/GTVContrato.cls:81` (`DameID`) | La tabla no existe; la del backend es `TbGTVContratos` |
| `TbResponsablesExpedientes` | Consulta guardada `Consulta5` del front | El front no la vincula: la tabla vive en `AGEDYS_DATOS` y el front no tiene ese vínculo |
| `TbExpedientes1` | Consulta guardada `Consulta5` del backend | La tabla no existe en `AGEDYS_DATOS`; la copia local se llama `TbExpedientes_1` y el vínculo a Expedientes se llama `TbExpedientes` |
| `TbAuxExpEstado` | SQL guardado de la consulta 16 de `TbConsultasDetalle` | La tabla no existe en ningún lado |
| `TbGTVContratos.IDExpediente` | SQL guardado de la consulta 17 de `TbConsultasDetalle` | El campo no existe: la unión de expediente se hace por `TbGTVContratosExpedientes` |
| `TbExpedientes.Generico` | Consulta guardada `Consulta5` del front | El campo se llama `AGEDYSGenerico` |
| `TbExpedientes.ResponsableCalidad` | Consulta guardada `Consulta6` del front | El campo se llama `IDResponsableCalidad` |

Las siete fallan en runtime. Un SQL que nombra una tabla o un campo inexistente no es un problema de PostgreSQL: ya está roto en el legacy.

## Propuestas para después del descubrimiento global

Ninguna se aplica: la puerta de descubrimiento global sigue cerrada. Los identificadores de esta ficha van con prefijo `A` para no colisionar con los de AGEDYS ECO; el catálogo conjunto (`docs/07-migracion/fase-0/catalogo-accionables.md`) tiene ya `A1`–`A7` de la Lanzadera y tendrá que desambiguar las dos series al incorporarlas.

| Id | Propuesta | Grupo (fase 1) | Comprobación previa | Reversión |
|---|---|---|---|---|
| A1 | Declarar clave primaria en las 11 tablas del backend que no la tienen: `TbAuxExp`, `TbAuxPrado`, `TbAvalesDatosGenerales`, `TbAvalesDatosGeneralesNuevos`, `TbDpDCartaSolicitudOfertaDatosFijos`, `TbDpDCartaSolicitudOfertaPlantillasSegunReqBasicos`, `TbFacturasPrincipal`, `TbLog`, `TbProductos`, `TbProductosObsoletos`, `TbProductos_antes`. | (a) | Contar duplicados por clave natural en cada una sobre una copia de producción: `TbFacturasPrincipal` (2604 filas) admite `NDPD`+`NPedido`; `TbProductos` (408) admite `Cod`; `TbLog` (19 568) necesita una clave de fila, no de negocio. Cinco de las once (`TbDpDCartaSolicitudOfertaDatosFijos`, `TbDpDCartaSolicitudOfertaPlantillasSegunReqBasicos`, `TbFacturasPrincipal`, `TbLog` y `TbProductos`) sí tienen referencias en el `src/` del front, y en ellas la clave natural no altera ninguna referencia; las otras seis no las nombra nadie. | Quitar la clave. |
| A2 | Retirar la credencial literal de las ocho cadenas de apertura del front (`src/modules/Variables Globales.bas:458, 485, 510, 543, 577, 611, 643, 911`) y leerla de la configuración. | (a) | Es la misma credencial que abre todos los binarios del ecosistema. Comprobar que ninguna otra copia del literal queda en el front ni en `00_AGEDYS_ECO`. | Restaurar la línea original. |
| A3 | Dejar una sola copia de las cinco tablas que el front y el backend tienen con el mismo nombre: `TbAuxExp`, `TbHerramientaDocAyuda`, `TbProductosObsoletos`, `TbSTGParametrosParaFormulario` y `TbSuministradoresListaFavoritos`. | (a) | Comparar contenido y recuento de las dos copias. `TbHerramientaDocAyuda` y `TbSTGParametrosParaFormulario` las lee el front desde la copia local (`Dame` y `getColParametrosSTG` abren con `CurrentDb()`), así que retirarlas exige cambiar antes la lectura a `getdb()`. Para las otras tres, el front no las usa. | Restaurar la copia local. |
| A4 | Poner en cuarentena `Copia de TbRequisitosSTG` del front: 139 filas con los mismos 73 campos que `TbRequisitosSTG`, que sí se usa a través del vínculo. | (a) | Confirmar que ninguna consulta guardada ni ningún script externo la lee. El código la nombra cero veces (CodeGraph y búsqueda exacta). | Devolver la tabla al front. |
| A5 | Poner en cuarentena `TbAuxNPedido` del front: 35 filas y cero referencias. | (a) | Confirmar que no la lee ningún otro front ni la web. El nombre sugiere una auxiliar de pedidos; ninguna búsqueda la encuentra en `src/`. | Devolver la tabla al front. |
| A6 | Corregir la referencia a `TbGTVContratos1` en `src/classes/GTVContrato.cls:81`: la tabla que existe es `TbGTVContratos`. | (a) | Reproducir en runtime el alta de un GTV: hoy `DameID` falla con «objeto no encontrado» y el identificador del contrato no se calcula. | Volver al nombre literal. |
| A7 | Sustituir el generador `Max + 1` (`DameID1`, `src/modules/Funciones Generales.bas:1261-1296`) por identidad en las tablas que escribe el front. | (a) | Enumerar las tablas y campos que pasan por `DameID1` y `DameID`, medir el hueco real de la secuencia en cada una y comprobar que no hay dos escritores concurrentes. Coordinar con las propuestas de los demás backends, que atacan el mismo generador. | Volver al `Max + 1`. |
| A8 | Declarar en el backend las relaciones que el front ya exige por SQL: el backend tiene 2 relaciones reales para 102 tablas. | (a) | Partir de los pares que el front une en sus SQL (`TbProyectos`/`TbNPedido` por `CODPPD`, `TbFacturasDetalle`/`TbVisadoFacturas_Nueva` por `NFactura` e `IDFactura`, `TbProyectos`/`TbExpedientes` por `IDExpediente`) y contar los huérfanos de cada par antes de declarar nada. Sin la limpieza previa, la integridad referencial no se puede declarar. | Quitar las relaciones. |
| A9 | Retirar o reparar las 13 consultas guardadas de depuración: ocho del front y las cinco del backend. | (a) | El front: `Consulta1` (actualiza `TbProyectos` donde `CODPROYECTOS = 'a'`), `Consulta3` (`'24dpd-032'`), `Consulta4` (actualiza `TbDPDCartaSolicitudOfertaDatos` donde `NDPD = 'a'`), `Consulta5` (`'jvm'`), `Consulta6` (`'24DPD-273'`), `Consulta9` (`IDExpediente = 421`), `Consulta11` (`SELECT;`) y `Consulta12` (`Year(FREGISTRO) = 2024`). El backend: `Consulta1` (un `UPDATE` guardado como consulta de selección), `Consulta2` (`SELECT;`), `Consulta3` (`NDOCUMENTO = '5113566585'`), `Consulta4` y `Consulta5`. Confirmar antes que ninguna las invoca desde la interfaz. | Restaurar la consulta. |
| A10 | Retirar del front los 67 vínculos que su código no usa (de los 127). | (b) | El recuento sale del cruce de los vínculos del ERD del front con las referencias del `src/`. Antes de retirar, cruzar con el grafo global de vínculos (tarea T6) y con las consultas guardadas del front, que usan `TbProyectosGestionRiesgos` (`Consulta8`) y `TbFacturasDetalleNoDYS` (SQL guardado de la consulta 21 de `TbConsultasDetalle`) sin nombrarlas en el VBA. Entre los 67 hay 59 a `AGEDYS_DATOS`, que son en su mayoría tablas que el front de AGEDYS ECO sí usa. | Restaurar el vínculo. |
| A11 | Declarar `Tabla1`, `TbAuxExpPrado` y `TbAuxTomas` como tablas de trabajo temporales del front, y revisar el formulario duplicado `Copia de FormGestionProyectos`, que está en el binario y no en el export. | (b) | Comprobar que las tres se vacían y se rellenan dentro de la misma sesión (`TbAuxExpPrado` se borra y se vuelve a llenar en `GenerarConsultaPrado`). Para el formulario duplicado, comprobar en la aplicación si el usuario llega a abrirlo. | Declararlas como tablas del backend. |
| A12 | Retirar los tres abridores sin llamadores: `getdbCorreoApertura` (`src/modules/Variables Globales.bas:495`), `getdbSalidaAntigua` (`:587`) y `getdbExpediente` (`:891`). | (b) | Confirmar con búsqueda de texto y con CodeGraph que no tienen llamadores, ni en `src/` ni por evento. El segundo y el tercero apuntan a backends a los que el front ya llega por vínculo. | Restaurar las funciones. |
| A13 | Corregir la condición local/remoto invertida de `getdbSalida` (`:558-563`), `getdbSalidaAntigua` (`:593-599`) y `getdbAgedo` (`:626-631`). | (b) | Con `DatosEnLocal = "No"` los tres eligen la ruta local y con `"Sí"` la remota, al contrario que `getdb`, `getdbLanzadera`, `getdbCorreo` y `getdbExpediente`. Comprobar qué ruta ha estado usando cada uno en producción antes de tocarlo. | Restaurar las condiciones. |
| A14 | Parametrizar y portar las 23 consultas guardadas de `TbConsultasDetalle`, que son SQL en texto que el front ejecuta en runtime. | (b) | El front las carga con `SELECT … FROM TbConsultasDetalle` (`src/forms/Form_FormConsultas.cls:63-65`), crea un `QueryDef` temporal y exporta a Excel (`src/modules/Funciones Generales.bas:69-76`). El SQL guardado usa parámetros de Access entre corchetes (`[INTRODUZCA EL Año]`, `[Introduzca la Primera Fecha]`) y Functions de Access (`IIf`, `Year`). Mapear las 13 tablas que tocan y reescribir las 23 con parámetros reales, no con el texto del marcador. El destino de una de ellas, `TbAuxExpEstado`, no existe. | Restituir el SQL guardado original. |
| A15 | Registrar el dueño y reubicar las tres tablas que solo consume otro backend: `TbExpedientesAnualidades`, `TbResponsablesExpedientes` y `TbComunicados`. | (b) | Requiere el grafo global de vínculos (tarea T6) y el acuerdo del dueño de `Expedientes_datos`. `TbResponsablesExpedientes` ya aparece como tabla ajena en `expedientes.md` E5 y `gestion-riesgos.md` R4: la decisión es transversal y no puede cerrarse en esta ficha. | Dejar la tabla donde está y documentarlo. |
| A16 | Normalizar la nomenclatura de las siete tablas de histórico y auxiliares del backend: `TbExpedientes_1`, `TbNPedidoAntes`, `TbNPedidosHistorica`, `TbProyectosHistoricos`, `TbProductos_antes`, `TbVisadoFacturas`, `TbGTV_Temp`. | (c) | Renombrar cambia el nombre, no el dato: exige actualizar el SQL guardado `Consulta5` del backend, que hoy nombra `TbExpedientes1`. Aplazar al apagado. | Renombrar de vuelta. |

Lo que **no** se toca ahora: los 127 vínculos del front, las 102 tablas propias del backend y las 10 del front, y el contenido de las 23 consultas guardadas de `TbConsultasDetalle`.

## Anexo: controles con origen de datos

### En diseño

Ningún formulario tiene `RecordSource` y ningún control tiene `ControlSource`. Los 14 controles con `RowSourceType` son listas, y solo dos llevan SQL:

| Formulario | Control | `RowSourceType` | SQL |
|---|---|---|---|
| `Form_FormConsultas` | `ListaConsultas` (`src/forms/Form_FormConsultas.form.txt:900-903`) | `Table/Query` | `SELECT TbConsultasDetalle.IDConsulta, TbConsultasDetalle.IDCategoria, TbConsultasDetalle.NombreConsulta AS Consulta, TbConsultasDetalle.DescripcionConsulta AS Descripción FROM TbConsultasDetalle WHERE (((TbConsultasDetalle.IDCategoria)=…));` |
| `Form_FormConsultas` | `ListaCategoria` (`src/forms/Form_FormConsultas.form.txt:1009-1012`) | `Table/Query` | `SELECT TbConsultasPrincipal.IDCategoria, TbConsultasPrincipal.NombreCategoria AS Categoría, TbConsultasPrincipal.DescripcionCategoria AS Descripción FROM TbConsultasPrincipal ORDER BY TbConsultasPrincipal.NombreCategoria;` |

Los otros 12 son listas de valores literales: `ListaDocumentos` (`src/forms/Form_FormAnexarDocumentacion.form.txt:583`), `ListaContratos` (`src/forms/Form_FormDPDAltaSeleccionContrato.form.txt:445`), `ListaContratos` (`src/forms/Form_FormDPDAltaSeleccionTipo.form.txt:642`), `ListaContratos` (`src/forms/Form_FormDPDEdicionCambioNContratoGTV.form.txt:885`), `ListaFiltrados` (`src/forms/Form_FormFacturasBusqueda.form.txt:907`), `ListaFiltrados` (`src/forms/Form_FormGestionProyectos.form.txt:1039`), `ListaFiltrados` (`src/forms/Form_FormGrupoArticulosBusqueda.form.txt:865`), `ListaSuministradores` (`src/forms/Form_FormRecepcionOferta.form.txt:1279`), `ListaSuministradores` (`src/forms/Form_FormSolicitudOferta.form.txt:2598`), `ListaFiltrados` (`src/forms/Form_FormSuministradoresBusqueda.form.txt:863`), `ListaFiltrados` (`src/forms/Form_FormVisadoDocumental.form.txt:957`) y `CampoBloqueado` (`src/forms/Form_Formulario1.form.txt:375`).

Un caso anómalo: `ListaFiltrados` de `Form_FormGrupoArticulosBusqueda` declara `RowSourceType ="Value List"` (`src/forms/Form_FormGrupoArticulosBusqueda.form.txt:865`) y a la vez lleva un `RowSource` con SQL sobre `TbProductos` y `TbProyectos` (`:866`). El código lo sustituye por una lista de valores en tiempo de ejecución (`src/forms/Form_FormGrupoArticulosBusqueda.cls:190, 362`), de modo que el SQL de diseño no llega a ejecutarse.

### En tiempo de ejecución

43 asignaciones a `RowSource` en `src/` (ninguna a `RecordSource` ni a `ControlSource`): 39 con lista de valores literal y 4 con una variable. Ninguna lleva SQL literal en el lado derecho. Las cuatro que sí acaban en SQL:

| Asignación | SQL | Tablas |
|---|---|---|
| `src/forms/Form_FormConsultas.cls:371` y `:379` | `SELECT TbConsultasDetalle.IDConsulta, TbConsultasDetalle.IDCategoria, TbConsultasDetalle.NombreConsulta AS Consulta, TbConsultasDetalle.DescripcionConsulta AS Descripción FROM TbConsultasDetalle WHERE (((TbConsultasDetalle.IDCategoria)=…));` | `TbConsultasDetalle` |
| `src/modules/Funciones Generales.bas:5969-5997` (además fija `RowSourceType = "Tabla/Consulta"`) | `SELECT DISTINCT TbExpedientes.Expediente, TbExpedientes.Titulo, TbExpedientes.IdExpediente FROM ((TbExpedientes LEFT JOIN TbExpedientesResponsables …) LEFT JOIN TbUsuariosAplicaciones …) LEFT JOIN TbProyectos … WHERE …` | `TbExpedientes`, `TbExpedientesResponsables`, `TbUsuariosAplicaciones`, `TbProyectos` |
| `src/forms/Form_FormVisadoDocumental.cls:233` | Lista de nombres de usuario de calidad construida en `src/classes/Entorno.cls:2172-2183` | ninguna |

Lo que de verdad ata la interfaz al esquema es el SQL que se ejecuta con `OpenRecordset` y se vuelca en las listas con `AddItem`, que no deja rastro en las propiedades del formulario: el front tiene 68 tablas nombradas en el código, de las que 40 son propias del backend.

## Qué queda en la parte técnica de APGEDYS y qué comparte con la parte ECO

AGEDYS Técnico y AGEDYS ECO son hoy dos fronts sobre **el mismo backend**, `AGEDYS_DATOS`, con las mismas 102 tablas. La decisión del propietario de que acaben siendo una sola aplicación web, APGEDYS, convierte el reparto de esas 102 tablas en la frontera entre las dos vistas.

**Reparto medido**, cruzando el nombre de las tablas que aparece en el `src/` de cada front con las 102 tablas propias del backend:

| Grupo | Tablas | Nº |
|---|---|---|
| Las usan los dos fronts | 31 | `TbAdjudicacionesEnvioAlSuministrador`, `TbAdjudicacionesOfertasCartas`, `TbCestasAnexosOrdinaria`, `TbConsultasDetalle`, `TbConsultasPrincipal`, `TbDPDCartaSolicitudOfertaDatos`, `TbDPDComiteOfertas`, `TbDPDOfertaSuministrador`, `TbDpDAsuntoDefensa`, `TbDpDCartaSolicitudOfertaDatosFijos`, `TbDpDCartaSolicitudOfertaPlantillasSegunReqBasicos`, `TbDpDDocAnexada`, `TbDpDInformeCondicionamiento`, `TbExpedientesAnualidades`, `TbExpedientesCodigoCompras`, `TbFacturasCliente`, `TbFacturasDetalle`, `TbFacturasPrincipal`, `TbGTV`, `TbGTVContratos`, `TbGTVContratosExpedientes`, `TbNPedido`, `TbOrdinariasTipoAnexos`, `TbProductos`, `TbProyectos`, `TbSolicitudesEnvioAlSuministrador`, `TbSolicitudesOfertasPrevias`, `TbSuministradoresPuntuacionConDetalle`, `TbSuministradoresSAP`, `TbVisadoFacturas_Nueva`, `TbVisadosGenerales` |
| Solo las usa el front técnico | 12 | `TbBloqueos`, `TbDpDPrepCartaOfAcepDetalle`, `TbGTV_Temp`, `TbHerramientaDocAyuda`, `TbLog`, `TbRequisitosSTG`, `TbSTGParametrosParaFormulario`, `TbSuministradoresNotasReferencia`, `TbTemp_PedidosEstado`, `TbTipoDocumento`, `TbTmp_DPDs`, `TbTmp_Facturas` |
| Solo las usa el front de AGEDYS ECO | 4 | `TbAuxExpedientesVivos`, `TbCalendarioProvisiones`, `TbCalendarioProvisionesDetalle`, `TbFacturasDetalleNoDYS` |
| No las usa ninguno de los dos | 55 | las 52 de la lista del veredicto de propiedad, más `TbComunicados`, `TbDpDPrepCartaOfAcepPrincipal` y `TbResponsablesExpedientes` |

**Lo que queda en la parte técnica.** Las 12 tablas que solo usa el front técnico, más la parte de las 31 compartidas que la vista de técnico necesita, más los 26 vínculos del front técnico que no apuntan a su propio backend: 11 a `Expedientes_datos`, 4 a `AGEDO20_Datos`, 4 a `Lanzadera_Datos`, 5 a los dos backends de Registro, 1 a `Correos_datos` y 1 a `Gestion_Riesgos_Datos`. De esos 26, el código usa 18; los 8 que no usa son `TbAgedoCorreosEnviados` y `TbDocExpProceso` (AGEDO), `TbAplicaciones` y `TbConexionesRegistro` (Lanzadera), `TbEntradas_RES`, `TbSalidas_Indexacion` y `TbSalidas_RES_Indexacion` (Registro) y `TbProyectosGestionRiesgos` (Gestión de Riesgos), que solo usa la consulta guardada `Consulta8` del front.

**Lo que comparte con la parte ECO.** El fichero `AGEDYS_DATOS` completo y las 31 tablas comunes. El catálogo de consultas guardadas `TbConsultasDetalle` vive en ese fichero: el front técnico carga y ejecuta su SQL en runtime (23 filas de SQL guardado) y el front de AGEDYS ECO abre el mismo fichero, de modo que hereda el mismo catálogo. Las dos familias `TbConexionesAGEDYSTEC` (15 filas) y `TbConexionesAGEDYSECO` (9 filas) ya separan por rol, y ninguna de las dos la escribe el código: las escribe la lanzadera.

**Tres fricciones que APGEDYS tiene que resolver.**

1. El mismo modelo de datos está duplicado en dos ficheros: `AGEDYS_DATOS` en el repositorio del front técnico y `AGEDYS_DATOS` en el de AGEDYS ECO. No son copias: tienen los mismos 102 nombres de tabla y los mismos 35 vínculos, pero 16 tablas tienen distinto recuento. Antes de unificar hay que decidir cuál de los dos ficheros es el de referencia.
2. La frontera por nombre de tabla no coincide con la frontera física: el front técnico lee `TbExpedientesAnualidades` por su vínculo a `Expedientes_datos` y no desde `AGEDYS_DATOS`, y el de AGEDYS ECO sí la usa desde `AGEDYS_DATOS`. En APGEDYS, el mismo nombre tendrá que resolverse a una sola tabla.
3. El SQL guardado en `TbConsultasDetalle` es SQL de aplicación, no de consulta: el front lo carga de la tabla y lo ejecuta en runtime. En APGEDYS deja de ser un dato y pasa a ser código, con sus parámetros ([`INTRODUZCA EL Año`], [`Introduzca la Primera Fecha`]) convertidos en entradas de la interfaz.

**Qué no se ha medido aquí.** El `src/` de AGEDYS ECO se ha recorrido solo para el recuento de nombres de tabla que da la tabla de reparto, no para un inventario de sus aperturas, sus controles conectados ni sus tablas propias. La ficha de AGEDYS ECO es la que responde de esa mitad.

## Lo que no se pudo verificar

1. El esquema de las 127 tablas vinculadas del front: los destinos están en `\\datoste`, que no es alcanzable desde este entorno. Se leyó la definición del vínculo, no los datos de origen. El de sus destinos sí se leyó, a través de los vínculos resueltos del backend contra `C:\00repos\datos`.
2. El estado de UAT: no hay evidencia en el repositorio. La única señal es la versión declarada, `2026-001` en el repositorio frente a `2026-002` en el servidor según el propietario.
3. El estado del `src/` desplegado: el binario tiene 45 formularios y 62 módulos; el export tiene 44 formularios y los 62 módulos. El formulario que falta es `Copia de FormGestionProyectos`. No se ha medido si el binario desplegado difiere del del repositorio.
4. La paridad entre el binario y el export de cada formulario: no se ha contrastado formulario a formulario con `verify_code`. Solo se ha medido el número de contenedores.
5. Los recuentos de filas: son los de la copia del repositorio (`AGEDYS_DATOS.accdb`, 2026-09-30), no los de `C:\00repos\datos\AGEDYS_DATOS.accdb` (2026-10-23), que difieren en 16 tablas.
6. El uso de las 55 tablas que ningún front nombra por código: no se ha comprobado si las leen la web, los informes, las tareas programadas o procesos externos. Es la comprobación que exige el protocolo de huérfanas.
7. Los vínculos del front que no usa su código: no se ha comprobado si los abre otro proceso con la base del front.
8. La tabla `TbAuxPrado` (468 filas) aparece como destino en la columna `TbAux` de `TbConsultasDetalle`, pero esa columna no la lee ningún código. No se ha averiguado quién la escribe.
