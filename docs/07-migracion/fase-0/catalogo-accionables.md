# Fase 0: catálogo único de accionables

> **Issue:** #1581, tarea T7 · **Épica:** [coexistencia progresiva](../epica-coexistencia-postgresql.md), fase 0 · **Fecha:** 2026-11-04

## Propósito

Este documento consolida en un solo catálogo revisable las propuestas de acción que las ocho fichas de aplicación de la fase 0 y el inventario de aperturas directas han dejado catalogadas. No ejecuta nada: cataloga, ordena por decisión y señala qué falta para decidir.

Cada fila remite a su ficha y conserva su identificador original. El texto íntegro de la propuesta, su comprobación previa y su reversión siguen en la ficha; aquí solo se ordena, se cualifica y se cruza.

## Fuentes

| Ficha | Aplicación | Backend | Identificadores | Propuestas |
|---|---|---|---|---|
| [lanzadera-datos.md](lanzadera-datos.md) | Lanzadera | `Lanzadera_Datos` | A1–A7 | 7 |
| [aperturas-directas-lanzadera.md](aperturas-directas-lanzadera.md) | 11 fronts | varios | P1–P4 | 4 |
| [brass.md](brass.md) | Brass | `Gestion_Brass_Gestion_Datos` | B1–B6 | 6 |
| [hps.md](hps.md) | HPS | `HPST` | H1–H12 | 12 |
| [condor.md](condor.md) | Condor | `condor_datos` | C1–C9 | 9 |
| [expedientes.md](expedientes.md) | Expedientes | `Expedientes_datos` | E1–E11 | 11 |
| [gestion-riesgos.md](gestion-riesgos.md) | Gestión de Riesgos | `Gestion_Riesgos_Datos` | R1–R10 | 10 |
| [no-conformidades.md](no-conformidades.md) | No Conformidades | `NoConformidades_Datos` | N1–N11 | 11 |
| [registros.md](registros.md) | Registros | `Registro_Ent_Salida_Datos` | G1–G14 | 14 |
| **Total** | | | | **84** |

El [grafo de dependencias](grafo-dependencias.md) no aporta propuestas: es la fuente de verdad sobre qué tablas son compartidas y cuántos backends las consumen, y se usa en la sección de conflictos.

Las fichas de HPS, Condor, Expedientes, Gestión de Riesgos, No Conformidades y Registros no citan issue propia; su revisión se hace con este catálogo conjunto, como se decidió en #1576. La de la Lanzadera viene de #1560, la de aperturas directas de #1570 y la de Brass de #1574.

## Reglas que gobiernan el catálogo

Estas reglas vienen del propietario y condicionan cómo se lee la tabla. No son criterios de este documento.

- **Nada se aplica en el legacy hasta que el propietario revise el conjunto.** La fase 0 cataloga; la puerta de descubrimiento global sigue cerrada.
- **Una tabla debe existir en el backend solo si el VBA de su front propietario la necesita.** Si solo la usa otro front, está mal ubicada y se reasigna; si nadie la usa, es candidata a retirada. Ningún borrado automático.
- **Los cambios de esquema durante la coexistencia son solo aditivos.** Este catálogo no aplica nada, pero deja registrado dónde una propuesta choca con esa regla.
- **Los grupos de la fase 1** son: (a) imprescindible antes de migrar, (b) recomendable si es barato, (c) aplazar a la refactorización final tras el apagado. La clasificación por grupo es la de cada ficha, según la [tabla de clasificación de la épica](../epica-coexistencia-postgresql.md#fase-1-saneamiento-en-access).

## Cómo se lee el catálogo

| Columna | Qué significa |
|---|---|
| **Id** | Identificador normalizado: prefijo de la ficha más número. No hay identificadores nuevos ni renumeraciones. |
| **Propuesta** | Qué se propone, en una línea. El detalle, la comprobación previa y la reversión están en la ficha. |
| **Aplicación** | Backend al que pertenece el cambio, no la ficha de la que sale. Una propuesta que actúa sobre una tabla de otra aplicación lo dice. |
| **Tablas afectadas** | Tablas sobre las que actúa el cambio. «—» cuando no toca ninguna tabla (credenciales, código, consultas). |
| **Bloqueada** | «Sí» cuando la decisión no puede cerrarse con la evidencia disponible hoy: falta una prueba que hoy no se puede hacer. «No» cuando la comprobación previa está definida y es ejecutable con el material actual. La sección de pruebas da el detalle. |
| **Ficha** | Fichero e identificador original de la propuesta. |

El orden es por decisión, no por aplicación: primero el grupo (a), después el (b), después el (c). Dentro de cada grupo van primero las propuestas que tocan tablas compartidas, ordenadas por número de backends consumidores, y después las que solo tocan tablas propias. El grafo da el criterio: una tabla compartida obliga a coordinar a todos sus consumidores y su cambio de esquema afecta a varias aplicaciones a la vez.

## Grupo (a): imprescindible antes de migrar

54 propuestas. 12 tocan tablas compartidas y 42 solo tablas propias.

### (a.1) Propuestas que tocan tablas compartidas

Ordenadas por número de backends consumidores de la tabla que tocan.

| Id | Propuesta | Aplicación | Tablas afectadas | Bloqueada | Ficha |
|---|---|---|---|---|---|
| C5 | Declarar clave primaria en `TbCorreosEnviados` | `Correos_datos` (no catalogado) | `TbCorreosEnviados` | Sí | [condor.md C5](condor.md#propuestas-para-después-del-descubrimiento-global) |
| N3 | Declarar clave primaria en `TbCorreosEnviados`, la única tabla viva sin clave que escribe No Conformidades | `Correos_datos` (no catalogado) | `TbCorreosEnviados` | Sí | [no-conformidades.md N3](no-conformidades.md#propuestas-para-después-del-descubrimiento-global) |
| R2 | Declarar clave primaria en las 5 tablas de Riesgos sin ella; una de ellas es de otro backend | `Gestion_Riesgos_Datos` y `Correos_datos` | `ProyectoRespCalidad`, `TbAuxProyectosRiesgos`, `TbCachePublicabilidadEdicion`, `TbRiesgosValoracion`, `TbCorreosEnviados` | Sí | [gestion-riesgos.md R2](gestion-riesgos.md#propuestas-para-después-del-descubrimiento-global) |
| E2 | Declarar clave primaria en las 6 tablas de Expedientes sin ella; tres pertenecen a otros backends | `Expedientes_datos`, `Correos_datos`, `AGEDYS_DATOS` | `TbCorreosEnviados`, `TbFacturasPrincipal`, `TbExpedientes1`, `TbExpAgedys`, `TbExpAGEDYS1`, `E2E_export_20260723_142133_lista` | Sí | [expedientes.md E2](expedientes.md#propuestas-para-después-del-descubrimiento-global) |
| N8 | Registrar el dueño real de `TbRiesgosNC` y de `TbCorreosEnviados`: las escribe No Conformidades | `Gestion_Riesgos_Datos` y `Correos_datos` | `TbRiesgosNC`, `TbCorreosEnviados` | Sí | [no-conformidades.md N8](no-conformidades.md#propuestas-para-después-del-descubrimiento-global) |
| E3 | Resolver las tres escrituras del front de Expedientes en tablas de otro backend | `Expedientes_datos` y los dueños de las tres tablas | `TbProyectos_Riesgos`, `TbNoConformidades`, `TbCorreosEnviados` | Sí | [expedientes.md E3](expedientes.md#propuestas-para-después-del-descubrimiento-global) |
| N4 | Sustituir el generador `MAX + 1` por identidad | `NoConformidades_Datos` y los dueños ajenos | 16 tablas propias, entre ellas `TbCorreosEnviados` y `TbRiesgosNC` | Sí | [no-conformidades.md N4](no-conformidades.md#propuestas-para-después-del-descubrimiento-global) |
| R6 | Sustituir el generador `Max + 1` por identidad | `Gestion_Riesgos_Datos` y 4 dueños ajenos | 18 tablas propias, más `TbCorreosEnviados`, `TbDocumentos`, `TbExpedientesResponsables`, `TbNoConformidades` | Sí | [gestion-riesgos.md R6](gestion-riesgos.md#propuestas-para-después-del-descubrimiento-global) |
| G2 | Sustituir los 5 generadores `Max + 1` del front | `Registro_Ent_Salida_Datos` y los dueños ajenos | `TbEntradas_RES`, `TbSalidas_RES`, `TbLog`, `TbCorreosEnviados`, `TbAplicacionesAperturas` | Sí | [registros.md G2](registros.md#propuestas-para-después-del-descubrimiento-global) |
| E8 | Sustituir `DameID` (máximo más uno) por identidad | `Expedientes_datos` | 27 pares tabla-campo, incluido `TbCorreosEnviados` | Sí | [expedientes.md E8](expedientes.md#propuestas-para-después-del-descubrimiento-global) |
| P1 | Sustituir el `MAX+1` por una fuente de identificador atómica, coordinando a todos los escritores de aperturas | `Lanzadera_Datos` (el mecanismo vive en 5 fronts) | `TbAplicacionesAperturas` | Sí | [aperturas-directas-lanzadera.md P1](aperturas-directas-lanzadera.md#acciones-de-preparación-derivadas) |
| A5 | Borrar los registros huérfanos de `TbConexionUltimaAppAbierta` y declarar la relación con integridad referencial | `Lanzadera_Datos` | `TbConexionUltimaAppAbierta`, `TbAplicaciones` | Sí | [lanzadera-datos.md A5](lanzadera-datos.md#catálogo-de-propuestas-para-después-del-descubrimiento-global) |

### (a.2) Propuestas sobre tablas propias

| Id | Propuesta | Aplicación | Tablas afectadas | Bloqueada | Ficha |
|---|---|---|---|---|---|
| A1 | Renombrar `Tb0HerramientaDocAyuda` a `TbHerramientaDocAyuda`, o retirar la llamada si la ayuda ya no se usa | `Lanzadera_Datos` | `Tb0HerramientaDocAyuda` | Sí | [lanzadera-datos.md A1](lanzadera-datos.md#catálogo-de-propuestas-para-después-del-descubrimiento-global) |
| A2 | Poner en cuarentena las 6 tablas vacías y sin referencias | `Lanzadera_Datos` | `TbAplicacionesVideos`, `TbCuestionarioPreguntas`, `TbCuestionarios`, `TbCuestionaroRespuestas`, `TbUsuarioAplicacionesSolicitud`, `TbVideosCuestionario` | Sí | [lanzadera-datos.md A2](lanzadera-datos.md#catálogo-de-propuestas-para-después-del-descubrimiento-global) |
| A3 | Poner en cuarentena `TbTablasAVincular`, o añadirle clave primaria si la usa un front sin repositorio | `Lanzadera_Datos` | `TbTablasAVincular` | Sí | [lanzadera-datos.md A3](lanzadera-datos.md#catálogo-de-propuestas-para-después-del-descubrimiento-global) |
| A4 | Poner en cuarentena las 15 candidatas restantes con datos, empezando por las de uso más improbable | `Lanzadera_Datos` | `TbAplicacionesEdiciones`, `TbAplicacionesEdicionesCambios`, `TbAplicacionesEstados`, `TbAplicacionesParametros`, `TbAplicacionesPerfiles`, `TbConexiones`, `TbDetalleVersiones`, `TbParametros`, `TbPermisos`, `TbUbicaciones`, `TbUsuarioConfiguracion`, `tbUsuarios`, `TbUsuariosTareasDiarias`, `TbVideosVisionados`, `TbUsuariosCorreosEnvio` | Sí | [lanzadera-datos.md A4](lanzadera-datos.md#catálogo-de-propuestas-para-después-del-descubrimiento-global) |
| P2 | Extraer las credenciales del código y definir cómo se proporcionan en Citrix | 11 fronts | — | Sí | [aperturas-directas-lanzadera.md P2](aperturas-directas-lanzadera.md#acciones-de-preparación-derivadas) |
| B1 | Poner en cuarentena las 7 tablas del backend sin uso en el front | `Gestion_Brass_Gestion_Datos` | `TbAuxManteniminetosPreventivosCalendario`, `TbCausaFin`, `TbGuiaConciliaciones`, `TbRepuestosTipo`, `TbResultadoVerificacion`, `TbSistema`, `TbTipoAsistencia` | Sí | [brass.md B1](brass.md#propuestas-para-después-del-descubrimiento-global) |
| B2 | Mover al backend las 4 tablas que viven en el front, o declararlas temporales del front | front de Brass y `Gestion_Brass_Gestion_Datos` | `TbActividadesLotes`, `TbAux`, `TbAuxEventosParaInforme`, `TbAuxNombreCamposParaInforme` | Sí | [brass.md B2](brass.md#propuestas-para-después-del-descubrimiento-global) |
| B3 | Declarar clave primaria en las 5 tablas auxiliares sin ella | `Gestion_Brass_Gestion_Datos` | `TbAuxActividad`, `TbAuxManteniminetosPreventivosCalendario`, `TbAuxMateriales`, `TbAuxPlanificacion`, `TbBuiIDEvento` | Sí | [brass.md B3](brass.md#propuestas-para-después-del-descubrimiento-global) |
| B4 | Retirar la credencial literal del código del front | front de Brass | — | Sí | [brass.md B4](brass.md#propuestas-para-después-del-descubrimiento-global) |
| B5 | Revisar las referencias a tablas inexistentes: o se recuperan en el backend o se elimina la referencia | front de Brass | `TbDocumentos`, `TbIDEquipos`, `TbEquiposEquipoMedida`, `TbSellados`, `TbPlanificacionAñadida` | Sí | [brass.md B5](brass.md#propuestas-para-después-del-descubrimiento-global) |
| B6 | Evitar que el front escriba en el share de producción desde desarrollo | front de Brass | 57 vínculos del front | Sí | [brass.md B6](brass.md#propuestas-para-después-del-descubrimiento-global) |
| H1 | Poner en cuarentena las 3 tablas propias del backend sin ninguna referencia en el front | `HPST` | `TbAuxCursos`, `TbHPSEquivalencia`, `TbJuridicasContratacion` | Sí | [hps.md H1](hps.md#propuestas-para-después-del-descubrimiento-global) |
| H2 | Retirar `TbUsuariosEntidades` del backend y su copia del front | `HPST` y front de HPS | `TbUsuariosEntidades`, `Copia de TbUsuariosEntidades` | Sí | [hps.md H2](hps.md#propuestas-para-después-del-descubrimiento-global) |
| H3 | Retirar del front las 2 cachés de indicadores retiradas en #32 | front de HPS | `TbDatosLocalParaIndicadores`, `TbUsuariosSICALocalParaIndicadores` | No | [hps.md H3](hps.md#propuestas-para-después-del-descubrimiento-global) |
| H6 | Retirar del front los 5 vínculos que apuntan a la ruta de producción | front de HPS | `TbExpedientes`, `TbSuministradores`, `TbHPS`, `TbUsuariosEntidades`, `TbUsuariosHistoricos` | Sí | [hps.md H6](hps.md#propuestas-para-después-del-descubrimiento-global) |
| H9 | Retirar la credencial literal de las cinco cadenas de conexión | front de HPS | — | Sí | [hps.md H9](hps.md#propuestas-para-después-del-descubrimiento-global) |
| H10 | Sustituir los generadores `Max + 1` por autonumérico, empezando por `DameID` y por el identificador de texto de SICA | `HPST` | `TbUsuarios`, `TbUsuariosHistoricos`, `TbObservaciones`, `TbObservacionesHistoricas`, `TbAnexosUsuariosHPS`, `TbAnexosUsuariosHistoricos`, `TbAnexosUsuariosSICA`, `TbUsuarioAnexos` | Sí | [hps.md H10](hps.md#propuestas-para-después-del-descubrimiento-global) |
| C1 | Poner en cuarentena `tbHistorialRechazos` | `condor_datos` | `tbHistorialRechazos` | Sí | [condor.md C1](condor.md#propuestas-para-después-del-descubrimiento-global) |
| C2 | Declarar `TbConfiguracionBackends` como configuración propia del front, exenta de la regla «las tablas viven en el backend» | front de Condor | `TbConfiguracionBackends` | Sí | [condor.md C2](condor.md#propuestas-para-después-del-descubrimiento-global) |
| C4 | Retirar los cuatro abridores externos sin llamadores | front de Condor | — | Sí | [condor.md C4](condor.md#propuestas-para-después-del-descubrimiento-global) |
| C6 | Sustituir los 9 generadores `Max + 1` por autonumérico o identidad | `condor_datos` | `tbSolicitudes`, `tbDatosPCSUB`, `tbDatosPC`, `tbDatosCDCASUB`, `tbDatosCDCA`, `tbValidacionRevision`, `tbLogErrores`, `tbRechazos`, `tbAdjuntos`, `tbLogCambios`, `tbMapeoCampos`, `tbLogEstados`, `tbEstados`, `tbTransiciones` | Sí | [condor.md C6](condor.md#propuestas-para-después-del-descubrimiento-global) |
| E1 | Aislar el vínculo `TbExpedientes1`, que apunta al AGEDYS de producción desde el backend de desarrollo | `Expedientes_datos` y `AGEDYS_DATOS` | `TbExpedientes1` | Sí | [expedientes.md E1](expedientes.md#propuestas-para-después-del-descubrimiento-global) |
| E4 | Obtener el SQL de las once consultas guardadas del backend y mapear sus tablas | `Expedientes_datos` | — (es una lectura) | Sí | [expedientes.md E4](expedientes.md#propuestas-para-después-del-descubrimiento-global) |
| E5 | Poner en cuarentena las 25 tablas que el VBA del front no usa | `Expedientes_datos` (11 propias) y los dueños de las 14 vinculadas | Propias: `Copia de TbExpedientes`, `Copia de TbExpedientesConEntidades`, `E2E_export_20260723_142133_lista`, `ListaPrevia`, `TbAusExpPostAGEDO`, `TbAuxEstadosMartina`, `TbAuxNemotecnico`, `TbExpAGEDYS1`, `TbExpedientes_antes`, `TbExpedientesCadenaContratacion`, `TbJefaturas`. Vinculadas: `TbComunicados`, `TbDpDInformeCondicionamiento`, `TbExpedientesAnualidades1`, `TbFacturasDetalle`, `TbFacturasPrincipal`, `TbGestionRiesgos`, `TbGTVContratos`, `TbGTVContratosExpedientes`, `TbNPedido`, `TbResponsablesExpedientes`, `TbSolicitudesOfertasPrevias`, `TbSuministradoresSAP`, `TbVisadoFacturas_Nueva`, `TbVisadosGenerales` | Sí | [expedientes.md E5](expedientes.md#propuestas-para-después-del-descubrimiento-global) |
| E6 | Retirar del front los cuatro vínculos que su código no usa | front de Expedientes | `Copia de TbExpedientes`, `Copia de TbExpedientesConEntidades`, `ListaPrevia`, `TbAusExpPostAGEDO` | Sí | [expedientes.md E6](expedientes.md#propuestas-para-después-del-descubrimiento-global) |
| E7 | Reconciliar `TbExpedientesDatosAGEDYS`, que el código cita y no está en el ERD del backend | `Expedientes_datos` y `AGEDYS_DATOS` | `TbExpedientesDatosAGEDYS`, `TbExpAgedys`, `TbExpAGEDYS1` | Sí | [expedientes.md E7](expedientes.md#propuestas-para-después-del-descubrimiento-global) |
| R1 | Sacar la contraseña del backend de la tabla del front | front de Gestión de Riesgos | `TbConfiguracionBackends` | Sí | [gestion-riesgos.md R1](gestion-riesgos.md#propuestas-para-después-del-descubrimiento-global) |
| R3 | Poner en cuarentena las 30 tablas propias sin uso en el front | `Gestion_Riesgos_Datos` | Lookups de identificador (14): `TbIDAccionPlanContingencia`, `TbIDAccionPlanMitigacion`, `TbIDAnexos`, `TbIDBibliotecaRiesgos`, `TbIDCorreosEnviados`, `TbIDDocumentos`, `TbIDEdiciones`, `TbIDLog`, `TbIDPlanContingencia`, `TbIDPlanMitigacion`, `TbIDProyectos`, `TbIDRiesgos`, `TbIDRiesgosAIntegrar`, `TbIDTareas`. Catálogos (5): `TbValoresPosiblesContingencia`, `TbValoresPosiblesEstadoRiesgo`, `TbValoresPosiblesMitigacion`, `TbValoresPosiblesPlazoCalidadCosteVulnerabilidad`, `TbValoresPosiblesValoracion`. Resto (11): `TbAnexosAntigua`, `TbAux`, `TbAuxSumExp`, `TbFelix`, `TbLog`, `TbOrigenesRiesgosDetalles`, `TbProyectosResponsablesCalidad`, `TbProyectosSuministradores`, `TbProyectosTipo`, `TbRiesgosPorTipoProyecto`, `TbTablas`. Aparte: `ProyectoRespCalidad` y `TbTareas` | Sí | [gestion-riesgos.md R3](gestion-riesgos.md#propuestas-para-después-del-descubrimiento-global) |
| R4 | Retirar del backend los 5 vínculos que el front no usa | `Gestion_Riesgos_Datos` y los dueños de las 5 tablas | `TbAplicaciones`, `TbDocumentosID`, `TbDPD`, `TbResponsablesExpedientes`, `TbSuministradoresSAP` | Sí | [gestion-riesgos.md R4](gestion-riesgos.md#propuestas-para-después-del-descubrimiento-global) |
| R5 | Resolver las 13 referencias a tablas inexistentes o mal nombradas: o la tabla se vincula o la referencia se elimina | `Gestion_Riesgos_Datos` y los backends que sí tienen esas tablas | `TbAplicacionesVersiones`, `TbVersionCambios`, `TbRACS`, `TbExpedientesRACS`, `TbProyectosJefesDeProyecto`, `TbOrganosContratacion`, `TbJuridicas`, `TbTipologia`, `TbCambioDocumentos`, `TbExpedientesConEntidades`, `TbIndicadoresRiesgos`, `TbExpedientes1`, `TbSuministradorSAP` | Sí | [gestion-riesgos.md R5](gestion-riesgos.md#propuestas-para-después-del-descubrimiento-global) |
| R8 | Fijar la semántica de comparación de texto antes de migrar | `Gestion_Riesgos_Datos` | — (afecta a las consultas de 357 módulos) | Sí | [gestion-riesgos.md R8](gestion-riesgos.md#propuestas-para-después-del-descubrimiento-global) |
| R9 | Decidir el nivel de aislamiento de las 571 aperturas de recordset antes de que el backend pase a ODBC | `Gestion_Riesgos_Datos` | — | No | [gestion-riesgos.md R9](gestion-riesgos.md#propuestas-para-después-del-descubrimiento-global) |
| N1 | Poner en cuarentena las 15 tablas propias sin referencias del backend | `NoConformidades_Datos` | `TbAnexos`, `TbNCARAvisos`, `TbNoConformidadesIngresoPorLotesTemporal`, `TbNoConformidadesIngresoPorLotesDetalle`, `TbNoConformidadesIngresoPorLotesPrincipal`, `TbTareasExplicaciones`, `TbConsultasPorFechas`, `TbConfiguracion`, `TbNCInformacionRAC`, `TbNCInformacionRACAnexos`, `TbAuditoriaLog`, `TbCacheIndicadoresProyectoHeader`, `TbCacheIndicadoresProyectoDetalle`, `TbCacheListadoNC`, `TbConexiones` | Sí | [no-conformidades.md N1](no-conformidades.md#propuestas-para-después-del-descubrimiento-global) |
| N2 | Ejecutar la migración de anexos en producción y retirar después `TbAnexosAuditoria` y `TbAnexosNCAuditorias` | `NoConformidades_Datos` | `TbAnexosAuditoria`, `TbAnexosNCAuditorias`, `TbDocumentosAuditorias` | Sí | [no-conformidades.md N2](no-conformidades.md#propuestas-para-después-del-descubrimiento-global) |
| N5 | Revisar las referencias a tablas que no existen y las dos consultas del front que las citan | front de No Conformidades | `TbACParaLista`, `TbARParaLista`, `TbNCDocumentosAux` | Sí | [no-conformidades.md N5](no-conformidades.md#propuestas-para-después-del-descubrimiento-global) |
| N6 | Retirar `NCProyectoOperaciones.ActualizarRiesgosNC`, que escribe una columna que no existe y no tiene llamadores | `NoConformidades_Datos` | `TbRiesgosNC` | Sí | [no-conformidades.md N6](no-conformidades.md#propuestas-para-después-del-descubrimiento-global) |
| N7 | Poner en cuarentena `TbTipologiaAux` del front, copia antigua de `TbTipologia` | front de No Conformidades | `TbTipologiaAux` | No | [no-conformidades.md N7](no-conformidades.md#propuestas-para-después-del-descubrimiento-global) |
| G1 | Poner en cuarentena las 5 tablas del backend propio que no usa ninguna aplicación | `Registro_Ent_Salida_Datos` | `TbConexiones`, `TbEntradaPersonasContacto_RES`, `TbEntradas_indexacion`, `TbEntradas_RES_Indexacion`, `TbSalidas_Indexacion` | Sí | [registros.md G1](registros.md#propuestas-para-después-del-descubrimiento-global) |
| G3 | Declarar clave primaria en las tablas del backend propio que no la tengan | `Registro_Ent_Salida_Datos` | 11 tablas del backend propio | Sí | [registros.md G3](registros.md#propuestas-para-después-del-descubrimiento-global) |
| G4 | Reponer los dos vínculos que el catálogo declara y el front no tiene, y corregir las columnas `Año` | front de Registros y `Registro_Datos` | `Entradas`, `Salidas`, `Entradas.Año`, `Salidas.Año` | Sí | [registros.md G4](registros.md#propuestas-para-después-del-descubrimiento-global) |
| G5 | Resolver las referencias que no tienen tabla | front de Registros | `TbSalidas`, `tbUsuarios` | Sí | [registros.md G5](registros.md#propuestas-para-después-del-descubrimiento-global) |
| G6 | Decidir el destino de `Registro_Datos.accdb` y de sus 5 tablas, que son la aplicación «RegistroAntiguo» | `Registro_Datos` | `Entradas`, `Salidas`, `Nombres`, `DClases`, `TbConfiguracion` | Sí | [registros.md G6](registros.md#propuestas-para-después-del-descubrimiento-global) |

## Grupo (b): recomendable si es barato

26 propuestas. 4 tocan tablas compartidas y 22 solo tablas propias.

### (b.1) Propuestas que tocan tablas compartidas

| Id | Propuesta | Aplicación | Tablas afectadas | Bloqueada | Ficha |
|---|---|---|---|---|---|
| E9 | Declarar la integridad referencial que falta hacia `TbUsuariosAplicaciones` | `Expedientes_datos` hacia `Lanzadera_Datos` | `TbExpedientes`, `TbUsuariosAplicaciones` | Sí | [expedientes.md E9](expedientes.md#propuestas-para-después-del-descubrimiento-global) |
| G12 | Corregir los `RowSource` de `TbExpedientes` y `TbProyectos` y apuntar al backend dueño | front de Registros | `TbExpedientes`, `TbProyectos` | Sí | [registros.md G12](registros.md#propuestas-para-después-del-descubrimiento-global) |
| H7 | Añadir a `HPST` el vínculo a `TbExpedientesSuministradores`, o mover esa lectura al backend | `HPST` | `TbExpedientesSuministradores` | No | [hps.md H7](hps.md#propuestas-para-después-del-descubrimiento-global) |
| H8 | Retirar de `HPST` los 2 vínculos que ninguna aplicación usa | `HPST` y los dueños de las 2 tablas | `TbExpedientesLugaresEjecucion`, `TbSolicitudesFechas` | Sí | [hps.md H8](hps.md#propuestas-para-después-del-descubrimiento-global) |

### (b.2) Propuestas sobre tablas propias

| Id | Propuesta | Aplicación | Tablas afectadas | Bloqueada | Ficha |
|---|---|---|---|---|---|
| A6 | Declarar con integridad referencial las relaciones que ya cumplen los datos, sin borrado en cascada | `Lanzadera_Datos` | `TbAplicacionesAperturas`, `TbVideos`, `TbVideosCategorias`, `TbCategorias` | No | [lanzadera-datos.md A6](lanzadera-datos.md#catálogo-de-propuestas-para-después-del-descubrimiento-global) |
| A7 | Borrar las consultas de prueba `Consulta1` a `Consulta4` del backend | `Lanzadera_Datos` | `Consulta1`, `Consulta2`, `Consulta3`, `Consulta4` | No | [lanzadera-datos.md A7](lanzadera-datos.md#catálogo-de-propuestas-para-después-del-descubrimiento-global) |
| P3 | Revisar la vida de las conexiones antes de ODBC | 11 fronts | — | Sí | [aperturas-directas-lanzadera.md P3](aperturas-directas-lanzadera.md#acciones-de-preparación-derivadas) |
| P4 | Decidir qué hacer con `getEsUsuarioAutorizado` de Contratos antes de conservarla o retirarla | front de Contratos | `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` | Sí | [aperturas-directas-lanzadera.md P4](aperturas-directas-lanzadera.md#acciones-de-preparación-derivadas) |
| H4 | Retirar del front las 4 tablas locales sin ninguna referencia en el repositorio | front de HPS | `Copia de TbUsuariosEntidades`, `TbCursosLocal`, `tblInfo`, `tblSettings` | Sí | [hps.md H4](hps.md#propuestas-para-después-del-descubrimiento-global) |
| H5 | Fijar por escrito el contrato de las 4 cachés vivas: origen único en el backend y regenerador declarado | front de HPS | `TbDatosLocal`, `TbUsuariosSICALocal`, `TbUsuariosHistoricosLocal`, `TbSuministradoresLocal` | Sí | [hps.md H5](hps.md#propuestas-para-después-del-descubrimiento-global) |
| H11 | Eliminar las ramas muertas de `DameIDAnexo` | front de HPS | `TbAnexosHistoricos`, `TbUsuarioHistoricosAnexos` | Sí | [hps.md H11](hps.md#propuestas-para-después-del-descubrimiento-global) |
| H12 | Revisar las referencias a `TbObservacionesLocal` y `TbObservacionesHistoricasLocal` | front de HPS | `TbObservacionesLocal`, `TbObservacionesHistoricasLocal` | Sí | [hps.md H12](hps.md#propuestas-para-después-del-descubrimiento-global) |
| C3 | Poner en cuarentena `tbSandboxLog` en el front | front de Condor | `tbSandboxLog` | Sí | [condor.md C3](condor.md#propuestas-para-después-del-descubrimiento-global) |
| C7 | Declarar las cuatro relaciones que faltan en el backend | `condor_datos` | `tbSolicitudes`, `tbDatosPCSUB`, `tbRechazos`, `tbLogEstados`, `tbValidacionRevision` | Sí | [condor.md C7](condor.md#propuestas-para-después-del-descubrimiento-global) |
| C8 | Reescribir el filtro de usuarios de los dos combos de log sin `MSysObjects` | front de Condor | `tbLogCambios`, `tbLogErrores` | No | [condor.md C8](condor.md#propuestas-para-después-del-descubrimiento-global) |
| C9 | Decidir el destino de la contraseña del backend | front de Condor | `TbConfiguracionBackends` | Sí | [condor.md C9](condor.md#propuestas-para-después-del-descubrimiento-global) |
| E10 | Decidir el destino de `TbAuxEstadosMartina`, duplicada en el front y en el backend | front de Expedientes y `Expedientes_datos` | `TbAuxEstadosMartina` | Sí | [expedientes.md E10](expedientes.md#propuestas-para-después-del-descubrimiento-global) |
| E11 | Revisar el `RowSource` residual de `Form_FormE2ECola` y las rutinas dormidas del instalador | front de Expedientes | `TbExpedientes`, `TbExpAgedys`, `TbExpedientes1` | Sí | [expedientes.md E11](expedientes.md#propuestas-para-después-del-descubrimiento-global) |
| R7 | Decidir el tratamiento de las 35 consultas guardadas: 4 de nombre ad hoc a cuarentena, 31 de negocio a retirada tras confirmar consumidores | `Gestion_Riesgos_Datos` | 35 consultas guardadas | Sí | [gestion-riesgos.md R7](gestion-riesgos.md#propuestas-para-después-del-descubrimiento-global) |
| R10 | Sustituir los 45 literales de fecha por parámetros o literales compatibles con ambos motores | `Gestion_Riesgos_Datos` | 8 ficheros de código | No | [gestion-riesgos.md R10](gestion-riesgos.md#propuestas-para-después-del-descubrimiento-global) |
| N9 | Retirar o documentar como consultas de análisis las 4 consultas guardadas del backend | `NoConformidades_Datos` | `ColumnaFalta`, `Consulta1`, `NC_Estudio_Replanificacion_2026`, `Replanificaciones_2026_Max_Fecha_Fin_Prevista` | Sí | [no-conformidades.md N9](no-conformidades.md#propuestas-para-después-del-descubrimiento-global) |
| G7 | Quitar el vínculo del front a `TbSalidas_RES_Indexacion`, que usan otras aplicaciones | front de Registros | `TbSalidas_RES_Indexacion` | Sí | [registros.md G7](registros.md#propuestas-para-después-del-descubrimiento-global) |
| G8 | Aislar el desarrollo de la producción: los vínculos apuntan a `\\datoste` y `getdb()` resuelve siempre la ruta remota | front de Registros | 13 vínculos y 3 abridores | Sí | [registros.md G8](registros.md#propuestas-para-después-del-descubrimiento-global) |
| G9 | Normalizar las rutas: cinco grafías distintas para cuatro backends | front de Registros | 13 vínculos | No | [registros.md G9](registros.md#propuestas-para-después-del-descubrimiento-global) |
| G10 | Retirar la credencial literal de las tres cadenas de apertura | front de Registros | — | Sí | [registros.md G10](registros.md#propuestas-para-después-del-descubrimiento-global) |
| G13 | Retirar la dirección de correo nominal fija escrita en el código | front de Registros | — | Sí | [registros.md G13](registros.md#propuestas-para-después-del-descubrimiento-global) |

## Grupo (c): aplazar a la refactorización final tras el apagado

4 propuestas.

| Id | Propuesta | Aplicación | Tablas afectadas | Bloqueada | Ficha |
|---|---|---|---|---|---|
| G14 | Renombrar `Tb*_RES` a una nomenclatura estable | `Registro_Ent_Salida_Datos` | `TbEntradas_RES`, `TbSalidas_RES` | No | [registros.md G14](registros.md#propuestas-para-después-del-descubrimiento-global) |
| N10 | Decidir cómo se reproduce la comparación sin distinguir mayúsculas | `NoConformidades_Datos` | 180 módulos de `src` | No | [no-conformidades.md N10](no-conformidades.md#propuestas-para-después-del-descubrimiento-global) |
| N11 | Decidir dónde vive `TbConfiguracionBackends`, con la ruta y la clave del backend en claro | front de No Conformidades | `TbConfiguracionBackends` | No | [no-conformidades.md N11](no-conformidades.md#propuestas-para-después-del-descubrimiento-global) |
| G11 | Mover las 3 consultas guardadas del front al backend y parametrizar el año y el SSID | front de Registros hacia `Registro_Ent_Salida_Datos` | `Consulta1`, `Consulta2`, `Consulta3` | No | [registros.md G11](registros.md#propuestas-para-después-del-descubrimiento-global) |

## Conflictos y dependencias entre propuestas

### Conflictos

| Id | Conflicto | Propuestas implicadas | Qué pasa si no se resuelve |
|---|---|---|---|
| X-01 | `TbCorreosEnviados` recibe cuatro propuestas de clave primaria desde cuatro aplicaciones, y vive en `Correos_datos.accdb`, un backend que no está en el catálogo | C5, N3, R2, E2 | Cuatro declaraciones de clave sobre la misma tabla, coordinadas por escrito y con la tabla en un backend sin dueño catalogado |
| X-02 | El dueño de `TbCorreosEnviados` y de `TbRiesgosNC` está sin registrar, y las dos las escribe una aplicación que no es su dueña | N8, N3, C5, E3, N4 | El esquema PostgreSQL de destino se crea sin dueño, que es exactamente lo que la regla de propiedad prohíbe |
| X-03 | Cuarentena y clave primaria sobre la misma tabla | B1 y B3 (`TbAuxManteniminetosPreventivosCalendario`); R3 y R2 (`ProyectoRespCalidad`, `TbAuxProyectosRiesgos`, `TbCachePublicabilidadEdicion`, `TbRiesgosValoracion`); E5 y E2 (`TbExpAgedys`, `TbExpAGEDYS1`, `E2E_export_20260723_142133_lista`); G1 y G3, sin delimitación posible | Una tabla renombrada a `zzBAJA_` y con clave declarada a la vez; el solapamiento de G1 y G3 no se puede acotar porque los índices del backend propio no son legibles |
| X-04 | H2 y H4 retiran la misma tabla del front | H2, H4 (`Copia de TbUsuariosEntidades`) | Dos retiradas para la misma tabla, con dos comprobaciones previas distintas |
| X-05 | G4 repone vínculos hacia un backend que G6 propone retirar | G4, G6 | Un vínculo restaurado hacia un archivo declarado histórico |
| X-06 | Siete propuestas cambian el mismo mecanismo, el generador `Max + 1`, y varias escriben en tablas compartidas | P1, C6, E8, H10, N4, R6, G2 | Cambios parciales incompatibles entre el front y la web durante la coexistencia, y generadores que dejan de generar el mismo identificador en tablas compartidas |
| X-07 | La misma decisión de semántica de texto está en dos grupos y en dos aplicaciones | R8 (grupo a), N10 (grupo c) | Se decide dos veces y se aplica en dos momentos distintos: la web y el legacy discreparán sobre el mismo dato |
| X-08 | Siete propuestas tratan las credenciales, con tres destinos distintos y tres grupos | P2, B4, H9, G10 (literales en el código); R1 (grupo a), C9 (grupo b), N11 (grupo c) | El mismo secreto sigue en el código de cuatro aplicaciones mientras su destino se decide tres veces |
| X-09 | R5 propone resolver referencias a tablas inexistentes, pero el grafo muestra que siete de esas tablas existen en otro backend | R5 (`TbRACS`, `TbExpedientesRACS`, `TbExpedientesConEntidades`, `TbJuridicas`, `TbOrganosContratacion`, `TbTipologia`, `TbSuministradoresSAP`) | Se elimina una referencia que solo necesitaba un vínculo, y se pierde el acceso a datos que existen |
| X-10 | E5 pone en cuarentena 14 tablas que no son de Expedientes, sino vínculos a tablas de AGEDYS y de Gestión de Riesgos | E5, E2, E3, R3 | Renombrar un vínculo en Access no retira la tabla de origen: la propuesta toca el enlace, no el dato, y no produce el efecto que su enunciado promete |
| X-11 | Dos tablas distintas comparten el nombre `TbProyectos`, en backends distintos, y el front de Registros la consume desde `AGEDYS_DATOS` | G12, R3, R4 y la ficha de No Conformidades | En el DDL de PostgreSQL, dos tablas homónimas de backends distintos se homonimian por esquema sin que nadie lo decida |
| X-12 | `TbHerramientaDocAyuda` existe con ese nombre en Gestión de Riesgos y en No Conformidades, y la propuesta A1 lo crea en la Lanzadera | A1 y las fichas de Gestión de Riesgos y de No Conformidades | Tres tablas homónimas en tres esquemas; el destino en PostgreSQL necesita un nombre desambiguado, y la decisión no está tomada |
| X-13 | La regla «los cambios de esquema durante la coexistencia son solo aditivos» choca con el renombrado de cuarentena | A2, A3, A4, B1, C1, C3, E5, G1, H1, H2, H3, H4, N1, N7 | Cuarentenas que renombran tablas en el legacy durante la coexistencia, sin que el propietario haya declarado que la regla se refiere solo al esquema de destino |
| X-14 | El inventario de aperturas mide `main`, y la ficha de Condor mide su RC1, donde los cuatro abridores externos no tienen llamadores | P1, P2, P3, P4, C4 | La superficie de adaptación se dimensiona sobre un código que en Condor ya no existe |

### Dependencias

| Id | Dependencia | Orden exigido |
|---|---|---|
| D-01 | N8 registra el dueño de `TbCorreosEnviados` y `TbRiesgosNC` | N8 antes de C5, N3, R2, E2, E3, N4 y R6: ninguna declaración de clave ni de identidad se ejecuta sobre una tabla sin dueño |
| D-02 | E4 obtiene el SQL de las once consultas guardadas de Expedientes | E4 antes de E5: sin el mapa de tablas no se puede afirmar que ninguna consulta usa una de las 25 candidatas |
| D-03 | E7 decide el nombre real de la tabla de AGEDYS | E7 antes de E2: la clave se declara sobre la tabla que resulte ser `TbExpAgedys` o `TbExpAGEDYS1`, no sobre las dos |
| D-04 | A5 limpia los huérfanos de `TbConexionUltimaAppAbierta` | A5 antes de A6: la integridad referencial se declara sobre datos ya limpios |
| D-05 | C6 cambia la generación de identificadores de las tablas de detalle | C6 antes de C7: las cuatro relaciones que faltan apuntan a `idSolicitud` |
| D-06 | N6 retira el procedimiento que escribe una columna inexistente | N6 antes de N4: la tabla `TbRiesgosNC` solo se migra a identidad con su esquema estable |
| D-07 | G4 restaura los vínculos y comprueba en producción si el front desplegado los tenía | G4 antes de G6: la decisión sobre `Registro_Datos.accdb` necesita saber si el frente vivo depende de él |
| D-08 | El grafo global de vínculos (tarea T6) cruza todos los backends | T6 antes de A3, B1, E5, E6, H2, H8, N1, R3, R4, G1, G6 y C1: ninguna cuarentena se aplica sin saber quién más consume la tabla |
| D-09 | P1 coordina a los cinco fronts que escriben aperturas | P1 se coordina con G2, que usa `DameID1` sobre `TbAplicacionesAperturas`, y con A5 y A6, que tocan la misma tabla. Ninguna de las dos partes cambia el esquema de esa tabla sin la otra |
| D-10 | R4 retira vínculos que apuntan a tablas de AGEDYS y AGEDO | R4 después del acuerdo con el dueño de AGEDYS, que está fuera del alcance de la migración y cuyos vínculos deben seguir funcionando |
| D-11 | E9 declara la integridad referencial hacia `TbUsuariosAplicaciones` | E9 después de fijar el tipo de `TbUsuariosAplicaciones.Id`, que la ficha de la Lanzadera deja como `Integer` de 16 bits sin cambiar |

## Pruebas que bloquean

Una propuesta está bloqueada cuando la decisión no puede cerrarse con la evidencia disponible hoy. De las 84 propuestas, 71 lo están. El bloqueo dominante no es el análisis de código, que está hecho, sino tres cosas: la evidencia de producción, el grafo global de vínculos y el SQL de las consultas guardadas, que no se ha podido leer con los permisos actuales.

Una misma propuesta puede quedar bloqueada por más de una prueba.

| Prueba | Propuestas que desbloquea | Prueba concreta |
|---|---|---|
| Exportar el SQL de las consultas guardadas | E4, H1, H4, R3, R7, N9, C1, A7 | Abrir una sesión de Access con permiso de lectura sobre `MSysQueries` y exportar las 64 consultas de backend (11 de Expedientes, 8 de HPS, 35 de Riesgos, 4 de No Conformidades, 2 de Condor, 4 de la Lanzadera) y la consulta del front de HPS. Es la prueba que más propuestas desbloquea por sesión |
| Recuento de filas y de huérfanos sobre una copia reciente de producción | A2, A3, A4, A5, B1, B3, C1, C6, E2, E5, E8, G1, G2, H1, H10, N1, N2, N7, R2, R6 | Repetir sobre una copia de producción los recuentos que hoy son de desarrollo: `SELECT COUNT(*)` por tabla en cuarentena, la consulta de duplicados de `IDBBDD`/`IDAplicacion`/`NombreTabla` de A3, y las consultas de huérfanos de A5 y A6 |
| Grafo global de vínculos de todos los backends (tarea T6) | A3, B1, E5, E6, H1, H2, H8, N1, R3, R4, G1, G6, C1 | Cruzar el grafo de vínculos de los backends del [grafo de dependencias](grafo-dependencias.md) con la búsqueda del nombre de la tabla en el `src` de los demás fronts, incluido el SQL dinámico |
| Reproducir en runtime los recorridos afectados | A1, B2, B5, C3, C4, E7, E11, G4, G5, G12, H3, H5, H12, N4, N5, N6, P1, P4, R5 | Ejecutar en la aplicación el recorrido que la propuesta nombra: abrir la ayuda, cerrar una planificación, buscar registros antiguos, ejecutar `RellenarTbProyectoParaLista`, guardar una NC con riesgos asociados, guardar una versión, y ensayar dos inserciones simultáneas para medir el generador de identificadores |
| Extraer los índices del backend de Registros | G3 | El volcado del que se dispone (`docs/ERD-auxiliar/`) no incluye índices y la interfaz de solo lectura no los expone. Hace falta una extracción que los incluya, para las 11 tablas del backend propio |
| Confirmar el entorno de cada copia y los permisos de escritura en el share | B6, E1, G8, H6 | Declarar, para cada copia, si es de desarrollo o de producción, y qué permisos de escritura tiene en `\\datoste`. Ninguna de las cuatro propuestas se puede cerrar sin ese dato |
| Acuerdo con los dueños ajenos | C2, C5, N3, N8, R2, R4, P2, B4, H9, G10, C9, H2, G13, E3, E10, E11 | Acordar con el dueño de Correos, de AGEDYS, de AGEDO, de Tareas, de HPS Solicitudes, de AGEDYS ECO y de AGEDYS TEC el registro de propiedad, la declaración de clave y el destino de las credenciales |
| Consumidores fuera de Access | A4, E5, E6, N1, R3, R7, G1, G6 | Revisar Excel, informes y tareas programadas que puedan leer las tablas en cuarentena. Es lo que exige el [protocolo de huérfanas](../epica-coexistencia-postgresql.md#protocolo-de-huérfanas) |
| Aislar el vínculo a producción y decidir su entorno destino | E1, y con él E2 | `TbExpedientes1` apunta a `\\datoste` desde el backend de desarrollo. Reorientarlo y comparar el conjunto de expedientes que devuelve la unión con `TbProyectos_Riesgos` antes y después del cambio |
| Fijar la versión de referencia de cada front | P1, P2, P3, P4 | El inventario de aperturas mide `main` de cada repositorio; la ficha de Condor mide su RC1. Confirmar qué rama va a producción en cada aplicación antes de dimensionar la adaptación |

## Qué necesita decidir el propietario

Las decisiones van en orden: primero las de criterio, que fijan la regla con la que se pueden decidir las propuestas; después las de aplicación, que dependen de esas reglas.

| Id | Decisión | Recomendación y fundamento |
|---|---|---|
| P-01 | Criterio general de una tabla que el front propietario no usa pero otro sí: ¿se reasigna o se retira? | Reasignar al dueño real y mantener el vínculo solo si el dueño lo pide; retirar solo tras el protocolo de huérfanas. Fundamento: R-2 no se confirma de forma general en el grafo, y las fichas H8, R4 y A3 ya apuntan en esa dirección |
| P-02 | Dueño de `TbCorreosEnviados` y `TbRiesgosNC`, y del backend `Correos_datos` que no está catalogado | Delegar en el dueño de Correos la declaración de clave y registrar su propiedad; no aprobar cuatro declaraciones en paralelo. Fundamento: seis backends consumen `TbCorreosEnviados` y ninguno es su dueño |
| P-03 | `TbUsuariosAplicacionesTareas`, que vive en la Lanzadera y solo la usa `Tareas_datos1`, fuera de alcance | Reasignarla a Tareas y documentarlo. No hay propuesta numerada para ella: la decisión se cierra aquí y se incorpora después como propuesta |
| P-04 | Aplicación de la cuarentena y la retirada a las 87 candidatas de 14 propuestas | No renombrar en el legacy hasta que la regla de cambios aditivos quede aclarada (X-13) y exista el grafo global. Fundamento: el protocolo de huérfanas exige la evidencia de no uso antes del renombrado, y esa evidencia falta para casi todas |
| P-05 | Riesgo de los vínculos a la ruta de producción `\\datoste` | Aislar antes del primer corte los vínculos de los fronts (H6, G8, B6) y el vínculo `TbExpedientes1` (E1); aceptar el riesgo solo con AGEDYS y AGEDO, que están fuera de la migración y cuyos vínculos deben seguir funcionando |
| P-06 | Destino único de las credenciales y orden de migración de los 11 fronts con literal | Un solo destino, aplicado por el mismo mecanismo que ya existe (`TbConfiguracionBackends`) y, a medio plazo, el gestor de secretos del servidor. Fundamento: la misma credencial abre todos los binarios del ecosistema (P2) |
| P-07 | `TbConfiguracionBackends`: exención declarada de la regla «las tablas viven en el backend» | Declarar la exención de forma transversal, con su motivo: contiene la ruta y la clave del backend, y moverla al backend es circular (C2). Aplicarla a las cuatro copias: Condor, Expedientes, Gestión de Riesgos y No Conformidades |
| P-08 | Semántica de comparación de texto (`Option Compare Database`) | Una sola decisión, en el grupo (a), porque afecta a las consultas que se migran. Fundamento: 356 de 357 módulos en Riesgos y 178 de 180 en No Conformidades; PostgreSQL no hereda esa comparación |
| P-09 | Identidad en lugar de `Max + 1`: ¿se genera en PostgreSQL desde el corte o se mantiene el generador actual? | Mantener el `Max + 1` en el legacy hasta el apagado y migrar a identidad, sin cambiar el legacy, coordinando las siete propuestas. Fundamento: el cambio en el legacy es un cambio de esquema no aditivo y el identificador se calcula antes de insertar la fila padre en al menos un caso (E8) |
| P-10 | Adaptación de las 11 aperturas directas: alternativa A (vínculo ODBC en el backend) o B (conexión ODBC centralizada) | B para la lectura de identidad, que es el caso de los 11 fronts y no depende del esquema; A solo para los que escriben `TbAplicacionesAperturas` y no la vinculan (Contratos y Registros, alcanzables, más AGEDYS, AGEDYS ECO y Brass si se reactivan sus llamadas). Decidir antes de migrar `Lanzadera_Datos`. Fundamento: DAO no admite recordsets de tipo tabla, así que B hay que probarla front por front, y la alternativa A exige crear vínculos en cinco backends que hoy no los tienen, coste que no está en ninguna propuesta numerada |
| P-11 | Consultas guardadas: retirada de las 31 de negocio de Riesgos y de las 4 de nombre ad hoc | Exportar primero el SQL; retirar solo las de nombre ad hoc; decidir las 31 de negocio por consumidor externo. Fundamento: ninguna se invoca desde el front, pero el SQL no se ha podido leer |
| P-12 | Tablas locales del front: política única para caché, tabla de trabajo y residuo | Tres clases declaradas: caché con contrato y regenerador (H5), tabla de trabajo temporal del front (B2) y residuo retirado (H3, H4, N7, C3). Fundamento: `TbAuxPriorizacion` de Gestión de Riesgos no tiene ninguna propuesta y el patrón está repartido en seis aplicaciones |
| P-13 | `Registro_Datos.accdb` y la aplicación «RegistroAntiguo» | Declararla archivo histórico de solo lectura. Fundamento: cero referencias vivas desde el front y tres vínculos entrantes que hay que confirmar con AGEDYS |
| P-14 | Alcance de la fase 1: catálogo entero o por aplicación | Aprobar solo lo desbloqueable, en una sola tanda previa. Fundamento: 71 de las 84 propuestas dependen de tres pruebas que se hacen una vez: la evidencia de producción, el grafo global de vínculos y el SQL de las consultas guardadas |

## Resumen por aplicación

| Aplicación | Backend | Propuestas | (a) | (b) | (c) | Bloqueadas | Toca tablas de otra aplicación |
|---|---|---|---|---|---|---|---|
| Lanzadera | `Lanzadera_Datos` | 11 | 7 | 4 | 0 | 9 | 3 |
| Brass | `Gestion_Brass_Gestion_Datos` | 6 | 6 | 0 | 0 | 6 | 0 |
| HPS | `HPST` | 12 | 6 | 6 | 0 | 10 | 3 |
| Condor | `condor_datos` | 9 | 5 | 4 | 0 | 8 | 1 |
| Expedientes | `Expedientes_datos` | 11 | 8 | 3 | 0 | 11 | 6 |
| Gestión de Riesgos | `Gestion_Riesgos_Datos` | 10 | 8 | 2 | 0 | 8 | 3 |
| No Conformidades | `NoConformidades_Datos` | 11 | 8 | 1 | 2 | 8 | 3 |
| Registros | `Registro_Ent_Salida_Datos` | 14 | 6 | 6 | 2 | 11 | 4 |
| **Total** | | **84** | **54** | **26** | **4** | **71** | **23** |

La columna «toca tablas de otra aplicación» cuenta propuestas cuyo cambio se aplica sobre una tabla propiedad de otra aplicación: una declaración de clave sobre `TbCorreosEnviados` cuenta en las cuatro aplicaciones que la proponen, y una cuarentena de tablas vinculadas cuenta donde el vínculo se retira, no donde el dato vive.

R7 está en el grupo (b) para sus 4 consultas de nombre ad hoc y en el grupo (c) para las 31 de negocio. En esta tabla cuenta una sola vez, en (b).

## Recuento de propuestas incorporadas

El catálogo incorpora **84 propuestas** de las ocho fichas de aplicación y del inventario de aperturas directas: A1–A7 (7), P1–P4 (4), B1–B6 (6), H1–H12 (12), C1–C9 (9), E1–E11 (11), R1–R10 (10), N1–N11 (11) y G1–G14 (14). La suma es 84, y el catálogo contiene 84 filas.

Por grupo: 54 en (a), 26 en (b) y 4 en (c), con R7 repartida entre (b) y (c) y contada una vez en (b).

## Lo que este catálogo no cubre

- **Huecos sin propuesta numerada.** Cuatro hallazgos tienen dueño claro y propuesta ausente: `TbUsuariosAplicacionesTareas` (mal ubicada, solo la usa `Tareas_datos1`), los vínculos que la alternativa A de las aperturas directas exige crear en cinco backends, los 69 huérfanos de `TbPermisos` contra `TbAplicaciones`, y `TbAuxPriorizacion` del front de Gestión de Riesgos. No se han inventado identificadores para ellos: se recogen aquí para que el propietario decida si entran en el catálogo.
- **Backends fuera del catálogo.** `Correos_datos`, `AGEDYS_DATOS`, `AGEDO20_Datos`, `Solicitudes_HPS_datos` y `Tareas_datos1` son dueños de tablas que este catálogo propone tocar y no tienen ficha de fase 0. Sus propuestas se quedan sin ficha: no se pueden decidir mientras tanto.
- **Aplicaciones sin ficha.** AGEDYS, AGEDYS ECO, Contratos, Seguridad, Solicitudes HPS, Tareas y Correos aparecen en el grafo y en el inventario de aperturas, pero no tienen inventario de fase 0 propio.
- **La web.** Ninguna ficha ha inventariado los accesos de la web a estos backends. El grafo solo cubre los fronts de Access.
