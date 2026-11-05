# Clasificación de las tablas

Paso 5 de la fase 0: cada tabla de los backends de la [copia de referencia](copia-referencia.md) recibe una de las clases del [protocolo de clasificación](../epica-coexistencia-postgresql.md#clasificación-de-tablas), según qué aplicaciones la usan de verdad. Las tablas propias de los fronts se clasifican aparte.

## Método

- **Uso.** Se busca el nombre de cada tabla en el VBA de cada aplicación, sin líneas comentadas ni módulos de test, en los orígenes de datos de formularios e informes (`RecordSource`, `RowSource`, `ControlSource`) y en el SQL de las consultas guardadas de su front y de su backend. Seguridad elige la tabla de cada tipo de registro desde los datos (`Tb0TipoRegistros.NombreTabla`), que también cuentan como uso. El código de formularios e informes (sección `CodeBehindForm`) se lee también sin comentarios, y una tabla que la aplicación no alcanza por su front ni por su backend solo cuenta si el nombre aparece como tabla: destino de SQL (`FROM`, `JOIN`, `INTO`, `UPDATE`, `TABLE`) o nombre exacto entre comillas o corchetes. Así no cuentan como uso las palabras corrientes que coinciden con un nombre de tabla (`Nombres`, `Derivados`, `Contratos`, `Lotes`) en comentarios, mensajes o propiedades de clases (#1714).
- **Qué tabla es.** Un nombre se resuelve en el contexto de la aplicación: tabla propia o vínculo de su front, tabla propia o vínculo de su backend y, si no está en ninguno, la tabla propia de otro backend que abre por código. Los nombres que podían referirse a dos tablas se resolvieron leyendo la conexión que abre el código (`CurrentDb`, `getdb` o `getdbXxx`); esas resoluciones están en la tabla final.
- **Tareas** sigue en uso en Access (decisión 14) y su código no se analiza: los vínculos de su backend cuentan como uso de un consumidor permanente. **AGEDO ya no se usa** (decisión 27) y sus vínculos no cuentan. **Contratos** no migra pero sigue vivo: su código sí se analiza.
- **Base de cada aplicación**: la RC1 cuando existe y no está integrada en `main`; si no, `main`.

| Aplicación | Rama | Commit |
|---|---|---|
| Lanzadera | `main` | `235b871` |
| Expedientes | `release/2026-008.2.rc1` | `e71d9e7` |
| Condor | `release/2026-007.rc1` | `6520705` |
| Seguridad | `release/2025-001.rc1` | `d02684e` |
| HPS | `main` | `dc194f4` |
| Solicitudes HPS | `main` | `cfa1937` |
| Gestión de Riesgos | `release/2026-008.2.rc1` | `3ee3394` |
| No Conformidades | `release/2026-009.rc1` | `f281b4e` |
| Registros | `main` | `44f36d3` |
| Brass | `2026-004.rc1` | `cb27951` |
| AGEDYS Técnico | `main` | `350896d` |
| AGEDYS ECO | `main` | `096478f` |
| Contratos | `main` | `effb968` |

El front de AGEDYS ECO no está versionado en su rama: sus tablas se leyeron del binario del checkout de `main`.

Se reproduce con `scripts/dysflow-erd/collect-usage.ps1` y `scripts/dysflow-erd/classify-tables.py`.

## Resumen por backend

| Backend | Propia | Propia y compartida | Mal ubicada | Mal ubicada con dueño por decidir | Huérfana | Consumida por el alcance | Sin consumidores del alcance |
|---|---:|---:|---:|---:|---:|---:|---:|
| `AGEDYS_DATOS` | 37 | 8 | · | · | 57 | · | · |
| `AGEDO20_Datos` | · | · | · | · | · | 2 | 64 |
| `Gestion_Riesgos_Datos` | 29 | 5 | · | · | 30 | · | · |
| `Gestion_Brass_Gestion_Datos` | 50 | · | · | · | 7 | · | · |
| `Expedientes_datos` | 23 | 15 | · | · | 9 | · | · |
| `NoConformidades_Datos` | 19 | 2 | · | · | 15 | · | · |
| `Seguridad_datos` | 26 | · | · | · | 10 | · | · |
| `Lanzadera_Datos` | 9 | 4 | 1 | · | 21 | · | · |
| `GestionContratos_datos` | 16 | 3 | · | · | 1 | · | · |
| `HPST` | 10 | 4 | 1 | · | 3 | · | · |
| `condor_datos` | 14 | · | · | · | 1 | · | · |
| `Registro_Ent_Salida_Datos` | 5 | 1 | · | · | 5 | · | · |
| `Solicitudes_HPS_datos` | 8 | 1 | · | · | 2 | · | · |
| `Registro_Datos` | 1 | 1 | · | · | 3 | · | · |
| `SICA_datos` | · | · | · | · | · | · | 4 |
| `Tareas_datos1` | · | · | · | · | · | · | 4 |
| `Control_Cambios_datos` | · | · | · | · | · | 3 | · |
| `Correos_datos` | · | · | · | · | · | 1 | · |
| **Total** | **247** | **44** | **2** | **0** | **164** | **6** | **72** |

## Tablas mal ubicadas

| Backend | Tabla | Filas | Clase | La usan |
|---|---|---:|---|---|
| `HPST` | `TbUsuariosEntidades` | 352 | Mal ubicada | Solicitudes HPS |
| `Lanzadera_Datos` | `TbUsuariosAplicacionesTareas` | 7 | Mal ubicada | Tareas |

## Tablas huérfanas

Ninguna aplicación del alcance, ni AGEDO ni Tareas, las usa. Entre paréntesis, las filas en la copia de referencia. Antes de ponerlas en cuarentena se aplica el [protocolo de huérfanas](../epica-coexistencia-postgresql.md#protocolo-de-huérfanas): este análisis no ve consumidores fuera de Access (Excel, informes o scripts programados).

- **`AGEDYS_DATOS`** (57): `TbAmbitoConsultas` (2), `TbAuxExp` (40), `TbAuxExpedientesUsuario` (0), `TbAuxExpedientesVivos` (40), `TbAuxGTV` (40), `TbAuxIDFactura` (4923), `TbAuxPilar` (56), `TbAuxPilar1` (18), `TbAuxPrado` (468), `TbAvalesAnexo` (32), `TbAvalesCancelacion` (187), `TbAvalesDatosGenerales` (1), `TbAvalesDatosGeneralesNuevos` (1), `TbAvalesFacturasClienteVinculo` (6), `TbAvalesInicio` (185), `TbAvalesInicioNuevo` (56), `TbAvalesObtencionAvalYDeposito` (132), `TbAvalesSolicitudSIADE` (151), `TbCestasTiposDocumentos` (0), `TbComunicados` (0), `TbComunicadosUsuario` (0), `TbConexiones` (24), `TbConexionesAGEDYSECO` (9), `TbConexionesAGEDYSTEC` (15), `TbConsultasOpciones` (6), `TbCorreoDiferidoNotas` (117), `TbCriteriosSuministrador` (6), `TbCSS` (1), `TbDatosEconomicosExpedientes` (91), `TbDireccionesSuministradores` (1), `TbDPDPlantillasParaCartaSolicitudDeOFerta` (9), `TbDpDPrepCartaOfAcepPrincipal` (860), `TbExpedientes_1` (463), `TbExpedientesAutorizaciones` (0), `TbExpedientesDocRequeridos` (158), `TbExpedientesProrrogas` (20), `TbFacturaInformeMatFalsicado` (0), `TbFacturasDetalleSinPedido` (2), `TbHerramientaDocAyuda` (10), `TbNPedidoAntes` (1499), `TbNPedidosHistorica` (1), `TbNuevosProductos` (407), `TbProductos_antes` (424), `TbProductosObsoletos` (25), `TbProductosParaBusqueda` (63), `TbProyectosDocumentosParaPedidoOrdinario` (0), `TbProyectosHistoricos` (2047), `TbSTGParametrosParaFormulario` (15), `TbSuministradoresDireccion` (24), `TbSuministradoresListaFavoritos` (0), `TbSuministradoresPuntuacion` (2450), `TbTareasTecnicasExplicaciones` (8), `TbTempDPDHTML` (193), `TbTiposImpositivos` (8), `TbTiposPedido` (3), `TbUsuariosEnvioTareas` (21), `TbVisadoFacturas` (1307).
- **`Gestion_Riesgos_Datos`** (30): `TbAnexosAntigua` (2), `TbAuxSumExp` (6), `TbFelix` (47), `TbIDAccionPlanContingencia` (251), `TbIDAccionPlanMitigacion` (1074), `TbIDAnexos` (107), `TbIDBibliotecaRiesgos` (51), `TbIDCorreosEnviados` (264), `TbIDDocumentos` (27), `TbIDEdiciones` (209), `TbIDLog` (1077), `TbIDPlanContingencia` (154), `TbIDPlanMitigacion` (640), `TbIDProyectos` (71), `TbIDRiesgos` (889), `TbIDRiesgosAIntegrar` (58), `TbIDTareas` (53), `TbLog` (5817), `TbOrigenesRiesgosDetalles` (5), `TbProyectosResponsablesCalidad` (39), `TbProyectosSuministradores` (2), `TbProyectosTipo` (0), `TbRiesgosPorTipoProyecto` (0), `TbTablas` (41), `TbTareas` (36), `TbValoresPosiblesContingencia` (4), `TbValoresPosiblesEstadoRiesgo` (7), `TbValoresPosiblesMitigacion` (4), `TbValoresPosiblesPlazoCalidadCosteVulnerabilidad` (5), `TbValoresPosiblesValoracion` (4).
- **`Lanzadera_Datos`** (21): `Tb0HerramientaDocAyuda` (4), `TbAplicacionesEdiciones` (1262), `TbAplicacionesEdicionesCambios` (1364), `TbAplicacionesEstados` (6), `TbAplicacionesParametros` (60), `TbAplicacionesPerfiles` (34), `TbAplicacionesVideos` (0), `TbConexiones` (24), `TbCuestionarioPreguntas` (0), `TbCuestionarios` (0), `TbCuestionaroRespuestas` (0), `TbDetalleVersiones` (28), `TbParametros` (53), `TbPermisos` (205), `TbTablasAVincular` (411), `TbUbicaciones` (8), `TbUsuarioAplicacionesSolicitud` (0), `TbUsuarioConfiguracion` (14), `tbUsuarios` (156), `TbUsuariosTareasDiarias` (8), `TbVideosCuestionario` (0).
- **`NoConformidades_Datos`** (15): `TbAnexos` (228), `TbAuditoriaLog` (0), `TbCacheIndicadoresProyectoDetalle` (0), `TbCacheIndicadoresProyectoHeader` (0), `TbCacheListadoNC` (0), `TbConexiones` (0), `TbConfiguracion` (1), `TbConsultasPorFechas` (8), `TbNCARAvisos` (432), `TbNCInformacionRAC` (2), `TbNCInformacionRACAnexos` (1), `TbNoConformidadesIngresoPorLotesDetalle` (166), `TbNoConformidadesIngresoPorLotesPrincipal` (1), `TbNoConformidadesIngresoPorLotesTemporal` (166), `TbTareasExplicaciones` (10).
- **`Seguridad_datos`** (10): `Tb0AdjuntosIndexaciones` (1), `Tb0Conexiones` (3), `Tb0RegistrosPendientes` (2), `Tb0TipoRegistrosDescripcion` (21), `Tb0TipoRegistroUso` (0), `Tb0Vinculo` (0), `TbAuditorias` (0), `TbBuzonTEOficio` (0), `TbCarta` (0), `TbRecibosMP` (0).
- **`Expedientes_datos`** (9): `Copia de TbExpedientes` (0), `Copia de TbExpedientesConEntidades` (365), `E2E_export_20260723_142133_lista` (89), `TbAusExpPostAGEDO` (203), `TbAuxEstadosMartina` (143), `TbAuxNemotecnico` (0), `TbExpAGEDYS1` (116), `TbExpedientesCadenaContratacion` (0), `TbJefaturas` (3).
- **`Gestion_Brass_Gestion_Datos`** (7): `TbAuxManteniminetosPreventivosCalendario` (93), `TbCausaFin` (4), `TbGuiaConciliaciones` (0), `TbRepuestosTipo` (3), `TbResultadoVerificacion` (3), `TbSistema` (2), `TbTipoAsistencia` (4).
- **`Registro_Ent_Salida_Datos`** (5): `TbConexiones` (12), `TbEntradaPersonasContacto_RES` (0), `TbEntradas_indexacion` (0), `TbEntradas_RES_Indexacion` (2), `TbSalidas_RES_Indexacion` (1063).
- **`HPST`** (3): `TbAuxCursos` (90), `TbHPSEquivalencia` (3), `TbJuridicasContratacion` (3).
- **`Registro_Datos`** (3): `DClases` (9), `Nombres` (57), `TbConfiguracion` (1).
- **`Solicitudes_HPS_datos`** (2): `Copia de TbExpedientes` (70), `TbLogs` (0).
- **`condor_datos`** (1): `tbHistorialRechazos` (0).
- **`GestionContratos_datos`** (1): `TbParaMarioloExpedientes` (67).

## Tablas de los fronts sin uso en su código

- **AGEDYS ECO**: `TbAuxExpPrado`, `TbAuxParaProvisiones`, `TbAuxPedidosGustavo`.
- **AGEDYS Técnico**: `Copia de TbRequisitosSTG`, `TbAuxExp`, `TbAuxNPedido`, `TbProductosObsoletos`, `TbSuministradoresListaFavoritos`.
- **Brass**: `TbAux`.
- **Condor**: `tbSandboxLog`.
- **Expedientes**: `TbAuxEstadosMartina`.
- **HPS**: `Copia de TbUsuariosEntidades`, `TbCursosLocal`, `TbDatosLocalParaIndicadores`, `tblInfo`, `tblSettings`, `TbUsuariosSICALocalParaIndicadores`.
- **No Conformidades**: `TbTipologiaAux`.
- **Seguridad**: `Tb0AuxOrdenarColeccion`, `TbAuxURLNoAlcanzables`.
- **Solicitudes HPS**: `TbCorreosEnviadosParaPruebas`.

Las otras 29 tablas propias de los fronts las usa su código.

## Nombres ambiguos resueltos leyendo el código

Estos nombres podían referirse a dos tablas. Se resolvieron mirando qué conexión abre el código (`CurrentDb` es el front; `getdb`, el backend propio; `getdbXxx`, otro backend).

| Aplicación | Nombre | Resolución |
|---|---|---|
| Lanzadera | `tbherramientadocayuda` | getdb() sobre Lanzadera_Datos, donde no existe: ayuda rota (Lanzadera#24) |
| Registros | `tbherramientadocayuda` | Dame() usa CurrentDb y el front no tiene la tabla: ayuda rota |
| Seguridad | `tbcorreosenviados` | abre con getdbCorreo |
| Gestión de Riesgos | `entradas` | palabra de un texto del formulario, no una tabla |
| Gestión de Riesgos | `salidas` | palabra de un texto del formulario, no una tabla |
| Expedientes | `tbexpedientes_1` | alias en un SQL (TbExpedientes AS TbExpedientes_1) |
| AGEDYS Técnico | `tbexpedientesanualidades` | CurrentDb: vínculo del front a Expedientes_datos |
| AGEDYS Técnico | `tbherramientadocayuda` | Dame() usa CurrentDb: la copia del front |
| AGEDYS Técnico | `tbstgparametrosparaformulario` | CurrentDb: la copia del front |
| Registros | `tbusuarios` | RowSource con CurrentDb y el front no vincula la tabla: combo roto (REG-G5) |
| AGEDYS ECO | `tbauxexpedientesvivos` | CurrentDb: la copia del front |

## Para decidir por el propietario

1. **`TbExpedientesCodigoCompras` tiene dos copias**, una en `AGEDYS_DATOS` y otra en `Expedientes_datos`. **Resuelto (2026-11-05, decisión 25):** las dos están vacías y la funcionalidad nunca se ha usado; se retira (issues [AGEDYS_ECO#11](https://github.com/DysTelefonica/AGEDYS_ECO/issues/11), [Expedientes#278](https://github.com/DysTelefonica/Expedientes/issues/278) y [AGEDYS#4](https://github.com/DysTelefonica/AGEDYS/issues/4)).
2. **`TbExpedientesAnualidades`**: AGEDYS Técnico ya usa la de Expedientes por su vínculo; AGEDYS ECO sigue usando la copia de `AGEDYS_DATOS`. Con la decisión 21, ECO tiene que pasar a la de Expedientes.
3. **`TbCorreosEnviados`**: además de la cola de `Correos_datos`, Contratos y Solicitudes HPS tienen cada una su propia tabla local con ese nombre.
   - **Contratos, resuelto (2026-11-05, decisión 26):** la clase `Correo` que escribe en la tabla no se instancia en ningún sitio, la tabla tiene 204 filas de 2023-2024 (184 sin enviar) y ningún proceso de Automatización abre el backend. Se retira ([CONTRATOS#2](https://github.com/DysTelefonica/CONTRATOS/issues/2)).
   - **Solicitudes HPS, abierto:** el propietario elige la cola común, pero la local no es una copia sino un diseño distinto, y falta decidir qué pasa con lo que solo ese diseño hace.
     - **Emisor propio en producción.** `EnviarCorreosHPS.vbs` lee la tabla local (`FechaEnvio IS NULL`), envía por SMTP y marca `FechaEnvio`; `script.vbs` de Automatización lo lanza en cada ciclo (paso 9/9). El informe diario `HPS_SOLICITUDES.vbs` también la lee.
     - **Estructura distinta.** Solo la local tiene `FechaOrdenEnvio`, `Programado`, `Intentos`, `FechaProceso`, `NombrePlantilla`, `VersionPlantilla`, `CadenaRecursos`, `IDSolicitud`, `Accion`, `DesencadenadoPor` y `TipoCorreo`: plantillas, envíos programados, reintentos y recordatorios.
     - **Datos.** La local tiene 9 correos de prueba (diciembre de 2025 a febrero de 2026), entre ellos los únicos recordatorios. Los avisos reales siguen saliendo por la cola común: 212 hasta mayo de 2026, ninguno recordatorio.
     - **Lo que falta decidir.** Quedarse con la cola común sin recordatorios, o ampliar la cola común con las columnas del diseño nuevo.
4. **`Nombres` de `Registro_Datos`**. **Resuelto (2026-11-05):** no la usa nadie. Los usos que se le atribuían eran la palabra «nombres» en comentarios, en un parámetro y en un mensaje; corregido el clasificador (#1714), es huérfana. Es una lista antigua de iniciales de personas (`Cod`, `Nombre`) con filas de prueba (`x`, `zzz`), y va a cuarentena con las demás huérfanas (`CLS-H1`).
5. **`TbUsuariosEntidades` de `HPST`**. **Resuelto (2026-11-05):** es una copia aplanada de los usuarios y las HPS que HPS mantenía como caché y nadie leía. HPS dejó de escribirla en la versión 2026-005 ([HPS#33](https://github.com/DysTelefonica/HPS/issues/33)); sus datos están congelados desde entonces. Solicitudes HPS aún la lee desde objetos del binario no exportados (su consulta `Solicitudes_alta` la cruza con `TbSolicitudes`) y debe pasar a los datos reales de HPS ([HPS_SOLICITUDES#1](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/1)); sus dos escrituras (`ActualizaDatosLocalEmpresas`, `AsociarUsuarioHPSConSolicitud`) no tienen llamadores. Después la tabla se retira (`CLS-M1`).
6. **`tbUsuarios` de `Lanzadera_Datos`**. **Resuelto (2026-11-05, decisión 27):** solo la vinculaba AGEDO, que ya no se usa; pasa a huérfana y va a cuarentena (`CLS-H1`).
7. **Tablas de la Lanzadera sin uso en el código actual pese a tener datos**: `TbAplicacionesEdiciones` (1 262 filas), `TbAplicacionesEdicionesCambios` (1 364), `TbPermisos` (205), `TbParametros` (53) y `TbTablasAVincular` (411), entre otras. Ningún script encontrado las escribe. ¿Se usan fuera de Access o son restos de versiones anteriores?

## Límites

- Un nombre construido en tiempo de ejecución (concatenado o leído de datos) no se detecta, salvo el catálogo de Seguridad.
- No se ven consumidores fuera de Access: Excel, informes ni scripts programados.
- Los módulos de test no cuentan como uso.
- Las filas de los fronts son las del front publicado; en producción cada usuario tiene su copia.
